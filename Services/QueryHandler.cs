using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Flow.Launcher.Plugin;
using Flow.Launcher.Plugin.QuickTodo.Models;

namespace Flow.Launcher.Plugin.QuickTodo.Services;

public class QueryHandler
{
    /// <summary>
    /// Dedicated action keyword that routes straight to Outlook mode.
    /// Registered automatically in <see cref="Main.InitAsync"/>.
    /// </summary>
    public const string OutlookActionKeyword = "tdo";

    private readonly TodoStore _store;
    private readonly PluginInitContext _context;
    private readonly OutlookTaskScriptClient? _outlookTasks;

    private static readonly Regex PriorityRegex = new(
        @"!(h|high|m|medium|l|low)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex CategoryRegex = new(
        @"@(\S+)", RegexOptions.Compiled);

    private static readonly Regex DateRegex = new(
        @"#(\S+)", RegexOptions.Compiled);

    // Time-of-day formats accepted after a date, e.g. "#tomorrow@1430", "#daily@9am".
    private static readonly Regex TimeAmPmRegex = new(
        @"^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TimeColonRegex = new(
        @"^(\d{1,2}):(\d{2})$", RegexOptions.Compiled);
    private static readonly Regex TimeCompactRegex = new(
        @"^\d{3,4}$", RegexOptions.Compiled);
    private static readonly Regex TimeHourRegex = new(
        @"^\d{1,2}$", RegexOptions.Compiled);

    public QueryHandler(TodoStore store, PluginInitContext context, OutlookTaskScriptClient? outlookTasks = null)
    {
        _store = store;
        _context = context;
        _outlookTasks = outlookTasks;
    }

    public List<Result> Handle(Query query)
    {
        var search = query.Search.Trim();

        // The dedicated "tdo" keyword routes every query straight to Outlook mode.
        if (string.Equals(query.ActionKeyword, OutlookActionKeyword, StringComparison.OrdinalIgnoreCase))
            return BuildOutlookResults(search, OutlookActionKeyword);

        if (string.IsNullOrEmpty(search))
            return BuildHomeResults();

        var firstWord = query.FirstSearch.ToLowerInvariant();

        return firstWord switch
        {
            "list" => BuildListResults(query.SecondToEndSearch),
            "cat" => BuildCategoryResults(query.SecondToEndSearch),
            "edit" => BuildEditResults(query.SecondToEndSearch),
            "outlook" => BuildOutlookResults(query.SecondToEndSearch, "td outlook"),
            _ => BuildAddResults(search)
        };
    }

    // --- HOME MODE ---

    private List<Result> BuildHomeResults()
    {
        var results = new List<Result>();

        results.Add(new Result
        {
            Title = "Add a task...",
            SubTitle = "Type after 'td' to create a new task",
            IcoPath = "Images\\todo.png",
            Score = 10000,
            AutoCompleteText = "td ",
            Action = _ => false
        });

        var overdue = _store.GetOverdue();
        foreach (var task in overdue.OrderByDescending(t => t.Priority))
        {
            results.Add(TaskToResult(task, FormatSubTitle(task), 5000));
        }

        var dueToday = _store.GetDueToday();
        foreach (var task in dueToday.OrderByDescending(t => t.Priority))
        {
            results.Add(TaskToResult(task, FormatSubTitle(task), 4000));
        }

        var incomplete = _store.GetAll()
            .Where(t => !t.IsCompleted && (t.DueDate == null || t.DueDate.Value.Date > DateTime.Now.Date))
            .OrderByDescending(t => t.Priority)
            .Take(10);

        foreach (var task in incomplete)
        {
            results.Add(TaskToResult(task, FormatSubTitle(task), PriorityScore(task.Priority)));
        }

        return results;
    }

    // --- ADD MODE ---

    private List<Result> BuildAddResults(string input)
    {
        var parsed = ParseModifiers(input);
        var results = new List<Result>();

        var subParts = new List<string> { parsed.Priority.ToString(), parsed.Category };
        if (parsed.Recurrence != Recurrence.None)
            subParts.Add(RecurrenceLabel(parsed.Recurrence));
        if (parsed.DueDate.HasValue)
            subParts.Add($"Due: {FormatDuePreview(parsed.DueDate.Value, parsed.HasDueTime)}");
        if (parsed.CategoryWarning != null)
            subParts.Add(parsed.CategoryWarning);

        var subTitle = string.Join(" | ", subParts);

        // Check for exact title match — offer update instead of duplicate add
        var exactMatch = _store.GetAll()
            .FirstOrDefault(t => t.Title.Equals(parsed.Title, StringComparison.OrdinalIgnoreCase));

        if (exactMatch != null)
        {
            results.Add(new Result
            {
                Title = $"Update: {exactMatch.Title}",
                SubTitle = subTitle,
                IcoPath = PriorityIcon(parsed.Priority),
                Score = 5001,
                Action = _ =>
                {
                    if (parsed.DueDate.HasValue)
                        _store.SetDueDate(exactMatch.Id, parsed.DueDate, parsed.HasDueTime);
                    if (parsed.Priority != exactMatch.Priority)
                        _store.SetPriority(exactMatch.Id, parsed.Priority);
                    if (!parsed.Category.Equals(exactMatch.Category, StringComparison.OrdinalIgnoreCase))
                        _store.SetCategory(exactMatch.Id, parsed.Category);
                    if (parsed.Recurrence != exactMatch.Recurrence)
                        _store.SetRecurrence(exactMatch.Id, parsed.Recurrence);
                    _context.API.ShowMsg("QuickTodo", $"Updated: {exactMatch.Title}");
                    return true;
                }
            });
        }

        results.Add(new Result
        {
            Title = $"Add: {parsed.Title}",
            SubTitle = subTitle,
            IcoPath = PriorityIcon(parsed.Priority),
            Score = 5000,
            Action = _ =>
            {
                _store.Add(parsed.Title, parsed.Priority, parsed.Category, parsed.DueDate,
                    parsed.Recurrence, parsed.HasDueTime);
                _context.API.ShowMsg("QuickTodo", $"Added: {parsed.Title}");
                return true;
            }
        });

        var existing = _store.GetAll()
            .Where(t => t.Title.Contains(parsed.Title, StringComparison.OrdinalIgnoreCase)
                        && !t.Title.Equals(parsed.Title, StringComparison.OrdinalIgnoreCase))
            .Take(5);

        foreach (var task in existing)
        {
            results.Add(TaskToResult(task, FormatSubTitle(task), 100));
        }

        return results;
    }

    // --- LIST MODE ---

    private List<Result> BuildListResults(string filterText)
    {
        var tasks = _store.GetAll();
        var filter = filterText.Trim().ToLowerInvariant();

        Priority? priorityFilter = null;
        string? categoryFilter = null;
        bool? overdueOnly = null;
        bool? doneOnly = null;
        var searchTerms = filter;

        var pm = PriorityRegex.Match(searchTerms);
        if (pm.Success)
        {
            priorityFilter = ParsePriority(pm.Groups[1].Value);
            searchTerms = PriorityRegex.Replace(searchTerms, "").Trim();
        }

        var cm = CategoryRegex.Match(searchTerms);
        if (cm.Success)
        {
            categoryFilter = cm.Groups[1].Value;
            searchTerms = CategoryRegex.Replace(searchTerms, "").Trim();
        }

        if (searchTerms == "overdue") { overdueOnly = true; searchTerms = ""; }
        else if (searchTerms == "done") { doneOnly = true; searchTerms = ""; }

        IEnumerable<TodoItem> filtered = tasks;

        if (priorityFilter.HasValue)
            filtered = filtered.Where(t => t.Priority == priorityFilter.Value);
        if (categoryFilter != null)
            filtered = filtered.Where(t => t.Category.Equals(categoryFilter, StringComparison.OrdinalIgnoreCase));
        if (overdueOnly == true)
            filtered = filtered.Where(t => !t.IsCompleted && t.DueDate.HasValue && t.DueDate.Value.Date < DateTime.Now.Date);
        if (doneOnly == true)
            filtered = filtered.Where(t => t.IsCompleted);
        if (!string.IsNullOrEmpty(searchTerms))
            filtered = filtered.Where(t => t.Title.Contains(searchTerms, StringComparison.OrdinalIgnoreCase));

        var sorted = filtered
            .OrderBy(t => t.IsCompleted)
            .ThenByDescending(t => t.Priority)
            .ThenBy(t => t.DueDate ?? DateTime.MaxValue);

        var results = new List<Result>();
        int score = 1000;
        foreach (var task in sorted)
        {
            results.Add(TaskToResult(task, FormatSubTitle(task), score--));
        }

        if (results.Count == 0)
        {
            results.Add(new Result
            {
                Title = "No tasks found",
                SubTitle = string.IsNullOrEmpty(filter) ? "Add tasks with 'td <task name>'" : $"No matches for '{filter}'",
                IcoPath = "Images\\todo.png"
            });
        }

        return results;
    }

    // --- CATEGORY MODE ---

    private List<Result> BuildCategoryResults(string args)
    {
        var parts = args.Trim().Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
        var subCommand = parts.Length > 0 ? parts[0].ToLowerInvariant() : "";
        var catName = parts.Length > 1 ? parts[1].Trim() : "";

        if (subCommand == "add" && !string.IsNullOrEmpty(catName))
        {
            return new List<Result>
            {
                new()
                {
                    Title = $"Add category: {catName}",
                    SubTitle = "Press Enter to add",
                    IcoPath = "Images\\todo.png",
                    Score = 1000,
                    Action = _ =>
                    {
                        if (_store.AddCategory(catName))
                            _context.API.ShowMsg("QuickTodo", $"Category '{catName}' added");
                        else
                            _context.API.ShowMsg("QuickTodo", $"Category '{catName}' already exists");
                        return true;
                    }
                }
            };
        }

        if (subCommand == "remove" && !string.IsNullOrEmpty(catName))
        {
            return new List<Result>
            {
                new()
                {
                    Title = $"Remove category: {catName}",
                    SubTitle = "Press Enter to remove (fails if tasks use this category)",
                    IcoPath = "Images\\todo.png",
                    Score = 1000,
                    Action = _ =>
                    {
                        if (_store.RemoveCategory(catName))
                            _context.API.ShowMsg("QuickTodo", $"Category '{catName}' removed");
                        else
                            _context.API.ShowMsg("QuickTodo", $"Cannot remove '{catName}' — in use or not found");
                        return true;
                    }
                }
            };
        }

        var categories = _store.GetCategories();
        var results = new List<Result>();
        int score = 1000;
        foreach (var cat in categories)
        {
            var taskCount = _store.GetAll().Count(t => t.Category.Equals(cat, StringComparison.OrdinalIgnoreCase));
            results.Add(new Result
            {
                Title = cat,
                SubTitle = $"{taskCount} task(s)",
                IcoPath = "Images\\todo.png",
                Score = score--,
                AutoCompleteText = $"td list @{cat}",
                Action = _ =>
                {
                    _context.API.ChangeQuery($"td list @{cat}");
                    return false;
                }
            });
        }

        return results;
    }

    // --- EDIT MODE ---

    private List<Result> BuildEditResults(string args)
    {
        var input = args.Trim();
        var firstToken = input.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";

        // Save mode: "edit <id> <new title + modifiers>" targets a specific task.
        if (Guid.TryParse(firstToken, out var id) && _store.GetById(id) is { } target)
        {
            var rest = input.Length > firstToken.Length ? input[firstToken.Length..].Trim() : "";
            return BuildEditSaveResults(target, rest);
        }

        // List mode: pick a task to edit.
        return BuildEditListResults(input);
    }

    private List<Result> BuildEditListResults(string filter)
    {
        var tasks = _store.GetAll()
            .Where(t => string.IsNullOrEmpty(filter)
                        || t.Title.Contains(filter, StringComparison.OrdinalIgnoreCase))
            .OrderBy(t => t.IsCompleted)
            .ThenByDescending(t => t.Priority)
            .Take(15)
            .ToList();

        if (tasks.Count == 0)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "No tasks to edit",
                    SubTitle = string.IsNullOrEmpty(filter)
                        ? "Add tasks with 'td <task>'"
                        : $"No matches for '{filter}'",
                    IcoPath = "Images\\todo.png"
                }
            };
        }

        var results = new List<Result>();
        var score = 1000;
        foreach (var task in tasks)
        {
            results.Add(new Result
            {
                Title = $"Edit: {task.Title}",
                SubTitle = $"{FormatSubTitle(task)} | Enter to change title/modifiers",
                IcoPath = PriorityIcon(task.Priority),
                Score = score--,
                AutoCompleteText = $"td edit {task.Id} {task.Title}",
                Action = _ =>
                {
                    _context.API.ChangeQuery($"td edit {task.Id} {task.Title}");
                    return false;
                }
            });
        }

        return results;
    }

    private List<Result> BuildEditSaveResults(TodoItem target, string rest)
    {
        if (string.IsNullOrWhiteSpace(rest))
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Type the new title and modifiers",
                    SubTitle = $"Editing: {target.Title}",
                    IcoPath = "Images\\todo.png",
                    Score = 5000,
                    Action = _ => false
                }
            };
        }

        var parsed = ParseModifiers(rest);
        var newTitle = string.IsNullOrWhiteSpace(parsed.Title) ? target.Title : parsed.Title;

        // Only override fields the user actually typed; otherwise keep the task's
        // current values (ParseModifiers fills unspecified fields with defaults).
        var priority = PriorityRegex.IsMatch(rest) ? parsed.Priority : target.Priority;
        var category = CategoryRegex.IsMatch(rest) ? parsed.Category : target.Category;
        var dateGiven = DateRegex.IsMatch(rest);
        var dueDate = dateGiven ? parsed.DueDate : target.DueDate;
        var hasDueTime = dateGiven ? parsed.HasDueTime : target.HasDueTime;
        var recurrence = dateGiven ? parsed.Recurrence : target.Recurrence;

        var subParts = new List<string> { priority.ToString(), category };
        if (recurrence != Recurrence.None)
            subParts.Add(RecurrenceLabel(recurrence));
        if (dueDate.HasValue)
            subParts.Add($"Due: {FormatDuePreview(dueDate.Value, hasDueTime)}");
        if (parsed.CategoryWarning != null && CategoryRegex.IsMatch(rest))
            subParts.Add(parsed.CategoryWarning);

        return new List<Result>
        {
            new()
            {
                Title = $"Save: {newTitle}",
                SubTitle = string.Join(" | ", subParts),
                IcoPath = PriorityIcon(priority),
                Score = 5000,
                Action = _ =>
                {
                    _store.SetTitle(target.Id, newTitle);
                    _store.SetPriority(target.Id, priority);
                    _store.SetCategory(target.Id, category);
                    _store.SetDueDate(target.Id, dueDate, hasDueTime);
                    _store.SetRecurrence(target.Id, recurrence);
                    _context.API.ShowMsg("QuickTodo", $"Updated: {newTitle}");
                    return true;
                }
            }
        };
    }

    // --- OUTLOOK MODE ---

    private List<Result> BuildOutlookResults(string args, string prefix)
    {
        if (_outlookTasks == null)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Outlook task bridge is not available",
                    SubTitle = "The plugin was not initialized with an Outlook task client",
                    IcoPath = "Images\\todo-high.png",
                    Score = 1000
                }
            };
        }

        var input = args.Trim();
        if (string.IsNullOrEmpty(input))
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Add Outlook task...",
                    SubTitle = $"Type: {prefix} <task> #tomorrow !high @Work",
                    IcoPath = "Images\\todo.png",
                    Score = 1000,
                    AutoCompleteText = $"{prefix} ",
                    Action = _ => false
                },
                new()
                {
                    Title = "List Outlook tasks",
                    SubTitle = $"Type: {prefix} list",
                    IcoPath = "Images\\todo.png",
                    Score = 900,
                    AutoCompleteText = $"{prefix} list",
                    Action = _ =>
                    {
                        _context.API.ChangeQuery($"{prefix} list");
                        return false;
                    }
                },
                new()
                {
                    Title = "Diagnose Outlook connection",
                    SubTitle = $"Type: {prefix} diag — probe the COM bridge step by step",
                    IcoPath = "Images\\todo.png",
                    Score = 800,
                    AutoCompleteText = $"{prefix} diag",
                    Action = _ =>
                    {
                        _context.API.ChangeQuery($"{prefix} diag");
                        return false;
                    }
                }
            };
        }

        var parts = input.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
        var subCommand = parts.Length > 0 ? parts[0].ToLowerInvariant() : "";

        if (subCommand == "list")
        {
            return BuildOutlookListResults();
        }

        if (subCommand == "diag")
        {
            return BuildOutlookDiagResults();
        }

        return BuildOutlookAddResults(input);
    }

    private List<Result> BuildOutlookAddResults(string input)
    {
        var parsed = ParseModifiers(input, validateCategory: false);
        var subParts = new List<string> { parsed.Priority.ToString(), parsed.Category };
        if (parsed.Recurrence != Recurrence.None)
            subParts.Add(RecurrenceLabel(parsed.Recurrence));
        if (parsed.DueDate.HasValue)
            subParts.Add($"Due: {parsed.DueDate.Value:yyyy-MM-dd}"); // Outlook tasks store date only

        return new List<Result>
        {
            new()
            {
                Title = $"Add to Outlook: {parsed.Title}",
                SubTitle = string.Join(" | ", subParts),
                IcoPath = PriorityIcon(parsed.Priority),
                Score = 5000,
                Action = _ =>
                {
                    try
                    {
                        _outlookTasks!.Add(parsed.Title, parsed.Priority, parsed.Category,
                            parsed.DueDate, parsed.Recurrence);
                        _context.API.ShowMsg("QuickTodo Outlook", $"Added: {parsed.Title}");
                    }
                    catch (Exception ex)
                    {
                        _context.API.ShowMsg("QuickTodo Outlook", ex.Message);
                    }
                    return true;
                }
            }
        };
    }

    private List<Result> BuildOutlookListResults()
    {
        try
        {
            var tasks = _outlookTasks!.List()
                .OrderBy(t => t.DueDate ?? DateTime.MaxValue)
                .ThenByDescending(t => t.Priority)
                .ToList();

            if (tasks.Count == 0)
            {
                return new List<Result>
                {
                    new()
                    {
                        Title = "No incomplete Outlook tasks found",
                        SubTitle = "Add one with: td outlook <task>",
                        IcoPath = "Images\\todo.png",
                        Score = 1000
                    }
                };
            }

            var results = new List<Result>();
            var score = 1000;
            foreach (var task in tasks)
            {
                results.Add(OutlookTaskToResult(task, score--));
            }

            return results;
        }
        catch (Exception ex)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Unable to read Outlook tasks",
                    SubTitle = ex.Message,
                    IcoPath = "Images\\todo-high.png",
                    Score = 1000
                }
            };
        }
    }

    // --- OUTLOOK DIAGNOSTICS ---

    private List<Result> BuildOutlookDiagResults()
    {
        if (_outlookTasks == null)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Outlook task bridge is not available",
                    SubTitle = "The plugin was not initialized with an Outlook task client",
                    IcoPath = "Images\\todo-high.png",
                    Score = 1000
                }
            };
        }

        OutlookDiagnostics diag;
        try
        {
            diag = _outlookTasks.Diagnose();
        }
        catch (Exception ex)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Outlook diagnostics failed to run",
                    SubTitle = $"{ex.Message} — Enter to copy details",
                    IcoPath = "Images\\todo-high.png",
                    Score = 5000,
                    Action = _ =>
                    {
                        _context.API.CopyToClipboard(ex.ToString());
                        return true;
                    }
                }
            };
        }

        var results = new List<Result>();
        var score = 5000;

        results.Add(new Result
        {
            Title = diag.Ok ? "Outlook connection OK" : "Outlook connection FAILED",
            SubTitle = BuildDiagSummary(diag),
            IcoPath = diag.Ok ? "Images\\todo-done.png" : "Images\\todo-high.png",
            Score = score--,
            Action = _ =>
            {
                var json = JsonSerializer.Serialize(diag, new JsonSerializerOptions { WriteIndented = true });
                _context.API.CopyToClipboard(json);
                return true;
            }
        });

        foreach (var step in diag.Steps)
        {
            var mark = step.Ok ? "✓" : "✗";
            results.Add(new Result
            {
                Title = $"{mark} {step.Name}",
                SubTitle = string.IsNullOrWhiteSpace(step.Detail) ? "" : step.Detail,
                IcoPath = step.Ok ? "Images\\todo-done.png" : "Images\\todo-high.png",
                Score = score--
            });
        }

        if (!string.IsNullOrWhiteSpace(diag.Error))
        {
            results.Add(new Result
            {
                Title = "Error detail",
                SubTitle = diag.Error,
                IcoPath = "Images\\todo-high.png",
                Score = score--
            });
        }

        return results;
    }

    private static string BuildDiagSummary(OutlookDiagnostics d)
    {
        var parts = new List<string>();
        if (d.BindMethod != null) parts.Add($"bind: {d.BindMethod}");
        if (d.OutlookVersion != null) parts.Add($"Outlook {d.OutlookVersion}");
        if (d.ProfileName != null) parts.Add($"profile: {d.ProfileName}");
        if (d.TasksFolderName != null) parts.Add($"folder: {d.TasksFolderName}");
        if (d.TaskCount.HasValue) parts.Add($"{d.IncompleteTaskCount}/{d.TaskCount} open");
        parts.Add($"PS {d.PowerShellVersion}");
        parts.Add(d.Is64BitProcess ? "x64" : "x86");
        return string.Join(" | ", parts) + " — Enter to copy JSON";
    }

    // --- HELPERS ---

    public record ParsedInput(
        string Title,
        Priority Priority,
        string Category,
        DateTime? DueDate,
        Recurrence Recurrence = Recurrence.None,
        bool HasDueTime = false,
        string? CategoryWarning = null);

    public ParsedInput ParseModifiers(string input, bool validateCategory = true)
    {
        var priority = Priority.Medium;
        string category = "Personal";
        DateTime? dueDate = null;
        var recurrence = Recurrence.None;
        var hasDueTime = false;
        var remaining = input;

        var pm = PriorityRegex.Match(remaining);
        if (pm.Success)
        {
            priority = ParsePriority(pm.Groups[1].Value);
            remaining = PriorityRegex.Replace(remaining, "").Trim();
        }

        // Parse the date/recurrence token before the category so a "#date@time"
        // suffix is consumed before the category matcher could grab the "@time".
        var dm = DateRegex.Match(remaining);
        if (dm.Success)
        {
            var token = ParseDateToken(dm.Groups[1].Value);
            dueDate = token.DueDate;
            recurrence = token.Recurrence;
            hasDueTime = token.HasTime;
            remaining = DateRegex.Replace(remaining, "").Trim();
        }

        string? categoryWarning = null;
        var cm = CategoryRegex.Match(remaining);
        if (cm.Success)
        {
            var requested = cm.Groups[1].Value;
            var categories = _store.GetCategories();
            var match = categories.FirstOrDefault(c => c.Equals(requested, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                category = match;
            }
            else if (!validateCategory)
            {
                category = requested;
            }
            else
            {
                categoryWarning = $"Unknown category '@{requested}', using Personal";
            }
            remaining = CategoryRegex.Replace(remaining, "").Trim();
        }

        return new ParsedInput(remaining, priority, category, dueDate, recurrence, hasDueTime, categoryWarning);
    }

    private record DateToken(DateTime? DueDate, Recurrence Recurrence, bool HasTime);

    // Parses a "#" token into a due date, optional recurrence, and optional time-of-day.
    // Examples: "tomorrow", "2024-05-30", "daily", "every-monday", "tomorrow@1430", "daily@9am".
    private static DateToken ParseDateToken(string token)
    {
        var datePart = token;
        string? timePart = null;
        var at = token.IndexOf('@');
        if (at >= 0)
        {
            datePart = token[..at];
            timePart = token[(at + 1)..];
        }

        var recurrence = Recurrence.None;
        DateTime? date;

        switch (datePart.ToLowerInvariant())
        {
            case "daily": recurrence = Recurrence.Daily; date = DateTime.Today; break;
            case "weekly": recurrence = Recurrence.Weekly; date = DateTime.Today; break;
            case "monthly": recurrence = Recurrence.Monthly; date = DateTime.Today; break;
            case "yearly": recurrence = Recurrence.Yearly; date = DateTime.Today; break;
            default:
                var lower = datePart.ToLowerInvariant();
                if (lower.StartsWith("every-")
                    && Enum.TryParse<DayOfWeek>(lower["every-".Length..], ignoreCase: true, out var dow))
                {
                    recurrence = Recurrence.Weekly;
                    date = NextWeekday(dow);
                }
                else
                {
                    date = ParseDate(datePart);
                }
                break;
        }

        var hasTime = false;
        if (date.HasValue && !string.IsNullOrEmpty(timePart))
        {
            var time = ParseTime(timePart);
            if (time.HasValue)
            {
                date = date.Value.Date + time.Value;
                hasTime = true;
            }
        }

        return new DateToken(date, recurrence, hasTime);
    }

    private static DateTime NextWeekday(DayOfWeek dow)
    {
        var today = DateTime.Today;
        var daysUntil = ((int)dow - (int)today.DayOfWeek + 7) % 7;
        if (daysUntil == 0) daysUntil = 7; // "next" that weekday, never today
        return today.AddDays(daysUntil);
    }

    private static TimeSpan? ParseTime(string token)
    {
        token = token.Trim();
        if (token.Length == 0) return null;

        var ampm = TimeAmPmRegex.Match(token);
        if (ampm.Success)
        {
            var h = int.Parse(ampm.Groups[1].Value);
            var m = ampm.Groups[2].Success ? int.Parse(ampm.Groups[2].Value) : 0;
            if (h is < 1 or > 12 || m > 59) return null;
            var isPm = ampm.Groups[3].Value.Equals("pm", StringComparison.OrdinalIgnoreCase);
            if (isPm && h != 12) h += 12;
            if (!isPm && h == 12) h = 0;
            return new TimeSpan(h, m, 0);
        }

        var colon = TimeColonRegex.Match(token);
        if (colon.Success)
        {
            var h = int.Parse(colon.Groups[1].Value);
            var m = int.Parse(colon.Groups[2].Value);
            return h > 23 || m > 59 ? null : new TimeSpan(h, m, 0);
        }

        if (TimeCompactRegex.IsMatch(token)) // HHmm, e.g. 1430 or 0900
        {
            var padded = token.PadLeft(4, '0');
            var h = int.Parse(padded[..2]);
            var m = int.Parse(padded[2..]);
            return h > 23 || m > 59 ? null : new TimeSpan(h, m, 0);
        }

        if (TimeHourRegex.IsMatch(token)) // bare hour, e.g. 9 -> 09:00
        {
            var h = int.Parse(token);
            return h > 23 ? null : new TimeSpan(h, 0, 0);
        }

        return null;
    }

    private static string RecurrenceLabel(Recurrence r) => r switch
    {
        Recurrence.Daily => "Repeats daily",
        Recurrence.Weekly => "Repeats weekly",
        Recurrence.Monthly => "Repeats monthly",
        Recurrence.Yearly => "Repeats yearly",
        _ => ""
    };

    private static string FormatDuePreview(DateTime due, bool hasTime)
        => due.ToString(hasTime ? "yyyy-MM-dd HH:mm" : "yyyy-MM-dd");

    private static Priority ParsePriority(string token) => token.ToLowerInvariant() switch
    {
        "h" or "high" => Priority.High,
        "m" or "medium" => Priority.Medium,
        "l" or "low" => Priority.Low,
        _ => Priority.Medium
    };

    private static DateTime? ParseDate(string token)
    {
        var lower = token.ToLowerInvariant();

        if (lower == "today") return DateTime.Today;
        if (lower == "tomorrow") return DateTime.Today.AddDays(1);

        if (Enum.TryParse<DayOfWeek>(token, ignoreCase: true, out var dow))
            return NextWeekday(dow);

        if (DateTime.TryParseExact(token, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d1))
            return d1;

        if (DateTime.TryParseExact(token, "MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d2))
            return new DateTime(DateTime.Now.Year, d2.Month, d2.Day);

        return null;
    }

    private Result TaskToResult(TodoItem task, string subTitle, int score)
    {
        var prefix = task.IsCompleted ? "\u2713 " : "";
        return new Result
        {
            Title = $"{prefix}{task.Title}",
            SubTitle = subTitle,
            IcoPath = task.IsCompleted ? "Images\\todo-done.png" : PriorityIcon(task.Priority),
            Score = score,
            ContextData = task,
            Action = _ =>
            {
                _store.ToggleComplete(task.Id);
                _context.API.ReQuery();
                return false;
            }
        };
    }

    private Result OutlookTaskToResult(TodoItem task, int score)
    {
        return new Result
        {
            Title = task.Title,
            SubTitle = FormatSubTitle(task),
            IcoPath = PriorityIcon(task.Priority),
            Score = score,
            ContextData = task,
            Action = _ =>
            {
                try
                {
                    _outlookTasks!.SetComplete(task, complete: true);
                    _context.API.ShowMsg("QuickTodo Outlook", $"Completed: {task.Title}");
                    _context.API.ReQuery();
                }
                catch (Exception ex)
                {
                    _context.API.ShowMsg("QuickTodo Outlook", ex.Message);
                }
                return false;
            }
        };
    }

    private static string FormatSubTitle(TodoItem task)
    {
        if (task.IsCompleted)
            return $"Completed: {task.CompletedAt:yyyy-MM-dd} | {task.Category}";

        var parts = new List<string> { task.Priority.ToString(), task.Category };

        if (task.Recurrence != Recurrence.None)
            parts.Add(RecurrenceLabel(task.Recurrence));

        if (task.DueDate.HasValue)
            parts.Add(FormatDue(task));

        return string.Join(" | ", parts);
    }

    private static string FormatDue(TodoItem task)
    {
        var due = task.DueDate!.Value;
        var now = DateTime.Now;

        if (task.HasDueTime)
        {
            var time = $" {due:HH:mm}";
            if (due < now)
            {
                var ago = now - due;
                if (ago.TotalDays >= 1) return $"OVERDUE by {(int)ago.TotalDays} day(s)";
                if (ago.TotalHours >= 1) return $"OVERDUE by {(int)ago.TotalHours}h";
                return $"OVERDUE by {Math.Max(1, (int)ago.TotalMinutes)}m";
            }
            if (due.Date == now.Date) return $"Due TODAY{time}";
            if (due.Date == now.Date.AddDays(1)) return $"Due tomorrow{time}";
            return $"Due: {due:yyyy-MM-dd}{time}";
        }

        var days = (due.Date - now.Date).Days;
        if (days < 0) return $"OVERDUE by {-days} day(s)";
        if (days == 0) return "Due TODAY";
        if (days == 1) return "Due tomorrow";
        return $"Due: {due:yyyy-MM-dd}";
    }

    internal static string PriorityIcon(Priority p) => p switch
    {
        Priority.High => "Images\\todo-high.png",
        Priority.Medium => "Images\\todo-medium.png",
        Priority.Low => "Images\\todo-low.png",
        _ => "Images\\todo.png"
    };

    private static int PriorityScore(Priority p) => p switch
    {
        Priority.High => 300,
        Priority.Medium => 200,
        Priority.Low => 100,
        _ => 0
    };
}
