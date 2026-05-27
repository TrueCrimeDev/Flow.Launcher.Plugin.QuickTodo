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

    public OutlookTaskScriptClient(string? scriptPath = null, Action<string, string>? logWarn = null)
    {
        _scriptPath = scriptPath ?? Path.Combine(AppContext.BaseDirectory, "Scripts", "QuickTodo.OutlookTasks.ps1");
        _logWarn = logWarn;
    }

    public TodoItem Add(string title, Priority priority, string category, DateTime? dueDate)
    {
        var args = NewBaseArguments("add");
        args.AddRange(new[] { "-Subject", title, "-Importance", ToOutlookImportance(priority), "-Categories", category });

        if (dueDate.HasValue)
        {
            args.Add("-DueDate");
            args.Add(dueDate.Value.ToString("yyyy-MM-dd"));
        }

        var output = Run(args);
        var record = JsonSerializer.Deserialize<OutlookTaskRecord>(output, JsonOptions)
            ?? throw new InvalidOperationException("Outlook task script returned no task record.");

        return ConvertFromOutlookRecord(record);
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

        process.Start();

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();

        if (!process.WaitForExit(15_000))
        {
            try { process.Kill(entireProcessTree: true); }
            catch { /* best effort */ }
            throw new TimeoutException("Outlook task script timed out after 15 seconds.");
        }

        if (process.ExitCode != 0)
        {
            _logWarn?.Invoke(nameof(OutlookTaskScriptClient), stderr);
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                ? $"Outlook task script failed with exit code {process.ExitCode}."
                : stderr.Trim());
        }

        return stdout.Trim();
    }

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
