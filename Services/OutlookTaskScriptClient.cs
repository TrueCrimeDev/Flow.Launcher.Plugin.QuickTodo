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
        args.AddRange(new[] { "-Subject", title, "-Importance", ToOutlookImportance(priority), "-Categories", category });

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
        args.Add("-Subject");
        args.Add(subject);

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

    private string Run(List<string> arguments)
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

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!process.WaitForExit(15_000))
        {
            try { process.Kill(entireProcessTree: true); }
            catch { /* best effort */ }
            _logWarn?.Invoke(nameof(OutlookTaskScriptClient),
                $"COM bridge timed out after 15s: {command}");
            throw new TimeoutException("Outlook task script timed out after 15 seconds.");
        }

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
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                ? $"Outlook task script failed with exit code {process.ExitCode}."
                : stderr.Trim());
        }

        return stdout.Trim();
    }

    private static string QuoteIfNeeded(string arg)
        => arg.Contains(' ') ? $"\"{arg}\"" : arg;

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
        public string? Body { get; set; }
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
    public string? PowerShellVersion { get; set; }
    public bool Is64BitProcess { get; set; }
    public List<OutlookDiagnosticStep> Steps { get; set; } = new();
    public string? Error { get; set; }
}

public sealed class OutlookDiagnosticStep
{
    public string Name { get; set; } = string.Empty;
    public bool Ok { get; set; }
    public string? Detail { get; set; }
}
