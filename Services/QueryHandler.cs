using System.Globalization;
using System.Text.RegularExpressions;
using Flow.Launcher.Plugin;
using Flow.Launcher.Plugin.QuickTodo.Models;

namespace Flow.Launcher.Plugin.QuickTodo.Services;

public class QueryHandler
{
    private readonly TodoStore _store;
    private readonly PluginInitContext _context;

    private static readonly Regex PriorityRegex = new(
        @"!(h|high|m|medium|l|low)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex CategoryRegex = new(
        @"@(\S+)", RegexOptions.Compiled);

    private static readonly Regex DateRegex = new(
        @"#(\S+)", RegexOptions.Compiled);

    public QueryHandler(TodoStore store, PluginInitContext context)
    {
        _store = store;
        _context = context;
    }

    public List<Result> Handle(Query query)
    {
        var search = query.Search.Trim();

        if (string.IsNullOrEmpty(search))
            return BuildHomeResults();

        var firstWord = query.FirstSearch.ToLowerInvariant();

        return firstWord switch
        {
            "list" => BuildListResults(query.SecondToEndSearch),
            "cat" => BuildCategoryResults(query.SecondToEndSearch),
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
            var days = (DateTime.Now.Date - task.DueDate!.Value.Date).Days;
            results.Add(TaskToResult(task, $"OVERDUE by {days} day(s) | {task.Priority} | {task.Category}", 5000));
        }

        var dueToday = _store.GetDueToday();
        foreach (var task in dueToday.OrderByDescending(t => t.Priority))
        {
            results.Add(TaskToResult(task, $"Due TODAY | {task.Priority} | {task.Category}", 4000));
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

        var subParts = new List<string> { parsed.Priority.ToString() };
        subParts.Add(parsed.Category);
        if (parsed.DueDate.HasValue)
            subParts.Add($"Due: {parsed.DueDate.Value:yyyy-MM-dd}");
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
                        _store.SetDueDate(exactMatch.Id, parsed.DueDate);
                    if (parsed.Priority != exactMatch.Priority)
                        _store.SetPriority(exactMatch.Id, parsed.Priority);
                    if (!parsed.Category.Equals(exactMatch.Category, StringComparison.OrdinalIgnoreCase))
                        _store.SetCategory(exactMatch.Id, parsed.Category);
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
                _store.Add(parsed.Title, parsed.Priority, parsed.Category, parsed.DueDate);
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

    // --- HELPERS ---

    public record ParsedInput(string Title, Priority Priority, string Category, DateTime? DueDate, string? CategoryWarning = null);

    public ParsedInput ParseModifiers(string input)
    {
        var priority = Priority.Medium;
        string category = "Personal";
        DateTime? dueDate = null;
        var remaining = input;

        var pm = PriorityRegex.Match(remaining);
        if (pm.Success)
        {
            priority = ParsePriority(pm.Groups[1].Value);
            remaining = PriorityRegex.Replace(remaining, "").Trim();
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
            else
            {
                categoryWarning = $"Unknown category '@{requested}', using Personal";
            }
            remaining = CategoryRegex.Replace(remaining, "").Trim();
        }

        var dm = DateRegex.Match(remaining);
        if (dm.Success)
        {
            dueDate = ParseDate(dm.Groups[1].Value);
            remaining = DateRegex.Replace(remaining, "").Trim();
        }

        return new ParsedInput(remaining, priority, category, dueDate, categoryWarning);
    }

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
        {
            var today = DateTime.Today;
            var daysUntil = ((int)dow - (int)today.DayOfWeek + 7) % 7;
            if (daysUntil == 0) daysUntil = 7;
            return today.AddDays(daysUntil);
        }

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

    private static string FormatSubTitle(TodoItem task)
    {
        if (task.IsCompleted)
            return $"Completed: {task.CompletedAt:yyyy-MM-dd} | {task.Category}";

        var parts = new List<string> { task.Priority.ToString(), task.Category };

        if (task.DueDate.HasValue)
        {
            var days = (task.DueDate.Value.Date - DateTime.Now.Date).Days;
            if (days < 0)
                parts.Add($"OVERDUE by {-days} day(s)");
            else if (days == 0)
                parts.Add("Due TODAY");
            else
                parts.Add($"Due: {task.DueDate.Value:yyyy-MM-dd}");
        }

        return string.Join(" | ", parts);
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
