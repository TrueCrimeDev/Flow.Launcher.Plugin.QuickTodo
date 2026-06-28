using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Flow.Launcher.Plugin.QuickTodo.Models;

namespace Flow.Launcher.Plugin.QuickTodo.Services;

public class OutlookTaskScriptClient
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    private readonly string _scriptPath;
    private readonly Action<string, string>? _logWarn;
    private readonly Action<string, string>? _logInfo;

    public OutlookTaskScriptClient(
        string? scriptPath = null,
        Action<string, string>? logWarn = null,
        Action<string, string>? logInfo = null)
    {
        _scriptPath = scriptPath ?? Path.Combine(AppContext.BaseDirectory, "Scripts", "QuickTodo.OutlookTasks.ps1");
        _logWarn = logWarn;
        _logInfo = logInfo;
    }

    public TodoItem Add(string title, Priority priority, string category, DateTime? dueDate,
        Recurrence recurrence = Recurrence.None)
    {
        var args = NewBaseArguments("add");
        args.Add(Param("Subject", title));
        args.Add("-Importance");
        args.Add(ToOutlookImportance(priority));
        args.Add(Param("Categories", category));

        if (dueDate.HasValue)
        {
            args.Add("-DueDate");
            args.Add(dueDate.Value.ToString("yyyy-MM-dd"));
        }

        if (recurrence != Recurrence.None)
        {
            args.Add("-Recurrence");
            args.Add(recurrence.ToString());
        }

        var output = Run(args);
        var record = JsonSerializer.Deserialize<OutlookTaskRecord>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook task script returned no task record.");

        return ConvertFromOutlookRecord(record);
    }

    public OutlookDiagnostics Diagnose()
    {
        var args = NewBaseArguments("diag");
        var output = Run(args);
        return JsonSerializer.Deserialize<OutlookDiagnostics>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook diagnostics returned no data.");
    }

    /// <summary>
    /// Drives Outlook's own Instant Search box for <paramref name="query"/> and brings
    /// Outlook to the front. No results are returned to Flow Launcher — the user lands
    /// inside Outlook looking at live results.
    /// </summary>
    public OutlookSearchResult Search(string query, string scope = "AllFolders")
    {
        if (string.IsNullOrWhiteSpace(query))
            throw new ArgumentException("Search query must not be empty.", nameof(query));

        var args = NewBaseArguments("search");
        args.Add(Param("Query", query));
        args.Add("-Scope");
        args.Add(scope);

        // `search` can cold-start Outlook (open the app, an Explorer, run Instant Search),
        // which easily exceeds the 15s budget the quick task ops use. It runs off the UI thread
        // (QueryHandler backgrounds it), so a longer budget can't freeze Flow Launcher.
        var output = Run(args, timeoutMs: 60_000);
        return JsonSerializer.Deserialize<OutlookSearchResult>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook search returned no result.");
    }

    public List<TodoItem> List(bool includeCompleted = false)
    {
        var args = NewBaseArguments("list");
        if (includeCompleted)
        {
            args.Add("-IncludeCompleted");
        }

        var output = Run(args);
        if (string.IsNullOrWhiteSpace(output))
        {
            return new List<TodoItem>();
        }

        var records = JsonSerializer.Deserialize<List<OutlookTaskRecord>>(output, JsonOptions);
        if (records != null)
        {
            return records.Select(ConvertFromOutlookRecord).ToList();
        }

        var single = JsonSerializer.Deserialize<OutlookTaskRecord>(output, JsonOptions);
        return single == null ? new List<TodoItem>() : new List<TodoItem> { ConvertFromOutlookRecord(single) };
    }

    public TodoItem SetComplete(TodoItem task, bool complete)
    {
        var args = NewBaseArguments(complete ? "complete" : "incomplete");
        AddTaskIdentityArguments(args, task);

        var output = Run(args);
        var record = JsonSerializer.Deserialize<OutlookTaskRecord>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook task script returned no task record.");

        return ConvertFromOutlookRecord(record);
    }

    public TodoItem Rename(TodoItem task, string subject)
    {
        var args = NewBaseArguments("rename");
        AddTaskIdentityArguments(args, task);
        args.Add(Param("Subject", subject));

        var output = Run(args);
        var record = JsonSerializer.Deserialize<OutlookTaskRecord>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook task script returned no task record.");

        return ConvertFromOutlookRecord(record);
    }

    public void Delete(TodoItem task)
    {
        var args = NewBaseArguments("delete");
        AddTaskIdentityArguments(args, task);
        Run(args);
    }

    private List<string> NewBaseArguments(string command)
    {
        return new List<string>
        {
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            _scriptPath,
            command,
            "-AsJson"
        };
    }

    private static void AddTaskIdentityArguments(List<string> args, TodoItem task)
    {
        if (string.IsNullOrWhiteSpace(task.OutlookEntryId))
        {
            throw new InvalidOperationException("This task does not have an Outlook EntryId.");
        }

        args.Add("-EntryId");
        args.Add(task.OutlookEntryId);

        if (!string.IsNullOrWhiteSpace(task.OutlookStoreId))
        {
            args.Add("-StoreId");
            args.Add(task.OutlookStoreId);
        }
    }

    private string Run(List<string> arguments, int timeoutMs = 15_000)
    {
        if (!File.Exists(_scriptPath))
        {
            throw new FileNotFoundException("Outlook task bridge script was not found.", _scriptPath);
        }

        using var process = new Process();
        process.StartInfo.FileName = "powershell.exe";
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;
        process.StartInfo.CreateNoWindow = true;

        foreach (var arg in arguments)
        {
            process.StartInfo.ArgumentList.Add(arg);
        }

        var command = "powershell.exe " + string.Join(" ", arguments.Select(QuoteIfNeeded));
        _logInfo?.Invoke(nameof(OutlookTaskScriptClient), $"COM bridge invoke: {command}");

        var stopwatch = Stopwatch.StartNew();
        process.Start();

        // Read both pipes asynchronously BEFORE the timed wait. A synchronous ReadToEnd() returns
        // only when the child closes stdout, so if the bridge hangs with its pipes open (Outlook
        // security prompt, MAPI profile dialog, a blocked COM call) it would block here forever and
        // never reach WaitForExit — defeating the timeout. Async reads also prevent the classic
        // two-pipe deadlock where the child fills stderr while we block draining stdout.
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();

        if (!process.WaitForExit(timeoutMs))
        {
            try { process.Kill(entireProcessTree: true); }
            catch { /* best effort */ }
            _logWarn?.Invoke(nameof(OutlookTaskScriptClient),
                $"COM bridge timed out after {timeoutMs / 1000}s: {command}");
            throw new TimeoutException($"Outlook task script timed out after {timeoutMs / 1000} seconds.");
        }

        // The process has exited; block until the async readers have drained to EOF so the captured
        // stdout isn't truncated (which would break JSON deserialization).
        process.WaitForExit();
        var stdout = stdoutTask.Result;
        var stderr = stderrTask.Result;

        stopwatch.Stop();
        _logInfo?.Invoke(nameof(OutlookTaskScriptClient),
            $"COM bridge exit={process.ExitCode} in {stopwatch.ElapsedMilliseconds}ms, " +
            $"stdout={stdout.Length} chars, stderr={stderr.Length} chars");

        // stderr can be present even on success (e.g. best-effort recurrence warnings).
        if (!string.IsNullOrWhiteSpace(stderr))
        {
            _logWarn?.Invoke(nameof(OutlookTaskScriptClient), $"COM bridge stderr: {stderr.Trim()}");
        }

        if (process.ExitCode != 0)
        {
            // On failure, dump the FULL stdout+stderr to the log (the success path only
            // logs char counts), so a COM error is fully diagnosable from the FL log
            // without having to reproduce it.
            _logWarn?.Invoke(nameof(OutlookTaskScriptClient),
                $"COM bridge FAILED exit={process.ExitCode}: {command}\n" +
                $"--- stdout ---\n{stdout.Trim()}\n--- stderr ---\n{stderr.Trim()}");

            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                ? $"Outlook task script failed with exit code {process.ExitCode}."
                : stderr.Trim());
        }

        return stdout.Trim();
    }

    private static string QuoteIfNeeded(string arg)
        => arg.Contains(' ') ? $"\"{arg}\"" : arg;

    // PowerShell's parameter binder treats a bare argv token that looks like "-Name" as a parameter
    // name, so a free-text value beginning with '-' (e.g. a search term that prefix-matches a real
    // param) would mis-bind and fail. The colon form "-Name:value" forces everything after the colon
    // to bind as the literal value, even with a leading dash.
    private static string Param(string name, string value) => $"-{name}:{value}";

    private static TodoItem ConvertFromOutlookRecord(OutlookTaskRecord record)
    {
        var dueDate = DateTime.TryParse(record.DueDate, out var parsedDueDate)
            ? parsedDueDate.Date
            : (DateTime?)null;

        var priority = record.Importance?.Equals("High", StringComparison.OrdinalIgnoreCase) == true
            ? Priority.High
            : record.Importance?.Equals("Low", StringComparison.OrdinalIgnoreCase) == true
                ? Priority.Low
                : Priority.Medium;

        var category = SplitCategories(record.Categories).FirstOrDefault() ?? "Outlook";

        return new TodoItem
        {
            Id = CreateStableId(record.EntryId ?? record.Subject ?? Guid.NewGuid().ToString()),
            Title = record.Subject ?? string.Empty,
            Priority = priority,
            Category = category,
            DueDate = dueDate,
            IsCompleted = record.Complete,
            CompletedAt = record.Complete ? DateTime.Now : null,
            OutlookEntryId = record.EntryId,
            OutlookStoreId = record.StoreId
        };
    }

    private static IEnumerable<string> SplitCategories(string? categories)
    {
        if (string.IsNullOrWhiteSpace(categories))
        {
            yield break;
        }

        foreach (var category in categories.Split(','))
        {
            var clean = category.Trim();
            if (clean.Length > 0)
            {
                yield return clean;
            }
        }
    }

    private static Guid CreateStableId(string value)
    {
        var bytes = MD5.HashData(Encoding.UTF8.GetBytes("outlook:" + value));
        bytes[6] = (byte)((bytes[6] & 0x0f) | 0x30);
        bytes[8] = (byte)((bytes[8] & 0x3f) | 0x80);
        return new Guid(bytes);
    }

    private static string ToOutlookImportance(Priority priority) => priority switch
    {
        Priority.High => "High",
        Priority.Low => "Low",
        _ => "Normal"
    };

    private sealed class OutlookTaskRecord
    {
        public string? EntryId { get; set; }
        public string? StoreId { get; set; }
        public string? Subject { get; set; }
        public string? DueDate { get; set; }
        public bool Complete { get; set; }
        public string? Importance { get; set; }
        public string? Categories { get; set; }
        // No Body: the bridge no longer returns it (reading TaskItem.Body could hang the COM call).
    }
}

/// <summary>Result of probing the Outlook COM connector, one entry per step.</summary>
public sealed class OutlookDiagnostics
{
    public bool Ok { get; set; }
    public string? BindMethod { get; set; }
    public string? OutlookVersion { get; set; }
    public string? ProfileName { get; set; }
    public string? DefaultStore { get; set; }
    public string? TasksFolderName { get; set; }
    public int? TaskCount { get; set; }
    public int? IncompleteTaskCount { get; set; }
    public string? InboxFolderName { get; set; }
    public int? InboxItemCount { get; set; }
    public int? AccountCount { get; set; }
    public int? StoreCount { get; set; }
    public bool? ExplorerAvailable { get; set; }
    public string? ComApartmentState { get; set; }
    public string? PowerShellVersion { get; set; }
    public bool Is64BitProcess { get; set; }
    public List<OutlookDiagnosticStep> Steps { get; set; } = new();
    public string? Error { get; set; }
}

/// <summary>Outcome of an Outlook Instant Search driven through the COM bridge.</summary>
public sealed class OutlookSearchResult
{
    public string? Query { get; set; }
    public string? Scope { get; set; }
    public bool Ok { get; set; }

    /// <summary>Non-fatal note from the bridge (e.g. it couldn't switch to the Inbox, so the search
    /// ran in the current view). Null on a clean search.</summary>
    public string? Warning { get; set; }
}

public sealed class OutlookDiagnosticStep
{
    public string Name { get; set; } = string.Empty;
    public bool Ok { get; set; }
    public string? Detail { get; set; }
}
