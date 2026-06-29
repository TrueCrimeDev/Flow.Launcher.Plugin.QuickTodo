using System.Globalization;
using System.Text;
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

    /// <summary>
    /// Dedicated action keyword for Outlook search: <c>os &lt;term&gt;</c> drives Outlook's
    /// Instant Search over mail. Registered automatically in <see cref="Main.InitAsync"/>.
    /// </summary>
    public const string SearchActionKeyword = "os";

    /// <summary>Outlook-styled icon (solid blue notebook) shown throughout tdo / Outlook mode.</summary>
    public const string OutlookIcon = "Images\\tdo.png";

    /// <summary>Per-task check icons for the tdo list: filled check when complete, hollow dot when not.</summary>
    public const string CheckedIcon = "Images\\tdo-checked.png";
    public const string UncheckedIcon = "Images\\tdo-unchecked.png";

    private readonly TodoStore _store;
    private readonly PluginInitContext _context;
    private readonly IconLoader _iconLoader;
    private readonly OutlookTaskScriptClient? _outlookTasks;

    // Outlook reads go through a (possibly slow / 30s-on-failure) COM bridge. To keep the
    // query thread from blocking — which freezes FL and makes "tdo" fall through to global
    // results — the list is fetched on a background thread and cached. A query returns the
    // cache (or a "Loading…" row) immediately; when the fetch finishes we ReQuery so the
    // normal render path picks up the fresh data. Both success and error are cached for a
    // short TTL so a closed Outlook doesn't trigger a tight 30s-retry loop.
    private static readonly long OutlookCacheTtlMs = (long)TimeSpan.FromSeconds(5).TotalMilliseconds;
    private readonly object _outlookCacheLock = new();
    private List<TodoItem>? _outlookCacheTasks;
    private string? _outlookCacheError;
    // Environment.TickCount64 (monotonic) at the last publish; null = never fetched / invalidated.
    // Wall-clock (DateTime.Now) would freeze the cache if the clock stepped backwards (DST/NTP).
    private long? _outlookCacheAtTicks;
    private bool _outlookFetching;
    // Bumped on every invalidation so an in-flight fetch can tell its snapshot is stale and refuse
    // to publish a pre-mutation list over the invalidation.
    private int _outlookCacheGeneration;

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

    public QueryHandler(TodoStore store, PluginInitContext context, IconLoader iconLoader, OutlookTaskScriptClient? outlookTasks = null)
    {
        _store = store;
        _context = context;
        _iconLoader = iconLoader;
        _outlookTasks = outlookTasks;
    }

    public List<Result> Handle(Query query)
    {
        var search = query.Search.Trim();

        // The dedicated "tdo" keyword routes every query straight to Outlook mode.
        if (string.Equals(query.ActionKeyword, OutlookActionKeyword, StringComparison.OrdinalIgnoreCase))
            return BuildOutlookResults(search, OutlookActionKeyword);

        // The dedicated "os" keyword routes straight to Outlook search.
        if (string.Equals(query.ActionKeyword, SearchActionKeyword, StringComparison.OrdinalIgnoreCase))
            return BuildOutlookSearchResults(search);

        if (string.IsNullOrEmpty(search))
            return BuildHomeResults();

        var firstWord = query.FirstSearch.ToLowerInvariant();

        return firstWord switch
        {
            "help" or "?" => BuildHelpResults(),
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
            IcoPath = "Images\\td.png",
            Score = 10000,
            AutoCompleteText = "td ",
            Preview = MarkdownPreview(BuildListMarkdown()),
            Action = _ => false
        });

        results.Add(new Result
        {
            Title = "Help — all commands",
            SubTitle = "Type 'td help' for the full command list",
            IcoPath = "Images\\td.png",
            Score = 50,
            AutoCompleteText = "td help",
            Preview = MarkdownPreview(BuildHelpMarkdown()),
            Action = _ =>
            {
                _context.API.ChangeQuery("td help");
                return false;
            }
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

    // --- HELP MODE ---

    // `td help` (or `td ?`) lists every command as navigable rows, each carrying the full
    // markdown cheat-sheet in the preview pane. Selecting a row jumps into that mode.
    private List<Result> BuildHelpResults()
    {
        var cheatSheet = MarkdownPreview(BuildHelpMarkdown());
        var score = 1000;

        Result Row(string title, string subTitle, string jumpTo, string icon) => new()
        {
            Title = title,
            SubTitle = subTitle,
            IcoPath = icon,
            Score = score--,
            AutoCompleteText = jumpTo,
            Preview = cheatSheet,
            Action = _ =>
            {
                _context.API.ChangeQuery(jumpTo);
                return false;
            }
        };

        return new List<Result>
        {
            Row("Add a task", "td <task> !h @Work #tomorrow@9am", "td ", "Images\\td.png"),
            Row("List & filter tasks", "td list [text] @Cat !h overdue|done", "td list ", "Images\\td.png"),
            Row("Categories", "td cat · td cat add <name> · td cat remove <name>", "td cat ", "Images\\td.png"),
            Row("Edit a task", "td edit [text] — pick a task, then change it", "td edit ", "Images\\td.png"),
            Row("Outlook tasks", "tdo <task> · tdo list · tdo diag", "tdo ", OutlookIcon),
            Row("Search Outlook mail", "os <term> · os diag", "os ", OutlookIcon),
        };
    }

    private static string BuildHelpMarkdown() => """
        ## QuickTodo — Commands

        ### Add a task
        `td <task>` — create a task. Combine modifiers:
        - `!h` `!m` `!l` — priority (high / medium / low)
        - `@Category` — assign a category
        - `#today` `#tomorrow` `#friday` `#2026-07-01` `#07-01` — due date
        - `#daily` `#weekly` `#monthly` `#yearly` `#every-monday` — recurring
        - `@14:30` `@9am` `@9` after a date — time (e.g. `#tomorrow@9am`)

        *Example: `td Email Sarah !h @Work #tomorrow@9am`*

        ### List & filter
        `td list` · `td list <text>` · `td list @Cat` · `td list !h` · `td list overdue` · `td list done`

        ### Categories
        `td cat` · `td cat add <name>` · `td cat remove <name>`

        ### Edit
        `td edit` · `td edit <text>` — pick a task, then change its title + modifiers

        ### Outlook tasks — `tdo`
        `tdo <task>` · `tdo list` · `tdo diag`

        ### Search Outlook mail — `os`
        `os <term>` · `os diag`
        """;

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
                Preview = MarkdownPreview(BuildAddMarkdown("Update task", parsed)),
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
            Preview = MarkdownPreview(BuildAddMarkdown("New task", parsed)),
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
                IcoPath = "Images\\td.png",
                Preview = MarkdownPreview(BuildListMarkdown())
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
                    IcoPath = "Images\\td.png",
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
                    IcoPath = "Images\\td.png",
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
                IcoPath = "Images\\td.png",
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
                    IcoPath = "Images\\td.png"
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
                    IcoPath = "Images\\td.png",
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
                    IcoPath = "Images\\td-high.png",
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
                    IcoPath = OutlookIcon,
                    Score = 1000,
                    AutoCompleteText = $"{prefix} ",
                    Action = _ => false
                },
                new()
                {
                    Title = "List Outlook tasks",
                    SubTitle = $"Type: {prefix} list",
                    IcoPath = OutlookIcon,
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
                    IcoPath = OutlookIcon,
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

    // --- OUTLOOK SEARCH MODE (os) ---

    private List<Result> BuildOutlookSearchResults(string args)
    {
        if (_outlookTasks == null)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Outlook bridge is not available",
                    SubTitle = "The plugin was not initialized with an Outlook client",
                    IcoPath = "Images\\td-high.png",
                    Score = 1000
                }
            };
        }

        var term = args.Trim();
        if (string.IsNullOrEmpty(term))
        {
            return new List<Result>
            {
                new()
                {
                    Title = "Search Outlook email…",
                    SubTitle = "Type a word after 'os' to search your Outlook mail (supports from:, subject:)",
                    IcoPath = OutlookIcon,
                    Score = 1000,
                    AutoCompleteText = "os ",
                    Action = _ => false
                },
                new()
                {
                    Title = "Diagnose Outlook connection",
                    SubTitle = "Type: os diag — probe the COM bridge step by step",
                    IcoPath = OutlookIcon,
                    Score = 900,
                    AutoCompleteText = "os diag",
                    Action = _ =>
                    {
                        _context.API.ChangeQuery("os diag");
                        return false;
                    }
                }
            };
        }

        // "os diag" shares the same step-by-step report as "tdo diag".
        if (term.Equals("diag", StringComparison.OrdinalIgnoreCase))
            return BuildOutlookDiagResults();

        return new List<Result>
        {
            new()
            {
                Title = $"Search Outlook for \"{term}\"",
                SubTitle = "Enter to run Outlook's search over your mail (all folders)",
                IcoPath = OutlookIcon,
                Score = 5000,
                Action = _ =>
                {
                    // Run the COM bridge off the UI thread so Flow Launcher doesn't freeze while
                    // Outlook (cold-)starts and runs the search. Errors and non-fatal warnings still
                    // surface via ShowMsg, which works after the launcher has closed.
                    Task.Run(() =>
                    {
                        try
                        {
                            var result = _outlookTasks!.Search(term);
                            if (!string.IsNullOrWhiteSpace(result?.Warning))
                                _context.API.ShowMsg("QuickTodo Outlook", result!.Warning!);
                        }
                        catch (TimeoutException)
                        {
                            _context.API.ShowMsg("QuickTodo Outlook",
                                "Outlook is taking a while to start — it should open with your search shortly.");
                        }
                        catch (Exception ex)
                        {
                            _context.API.ShowMsg("QuickTodo Outlook", ex.Message);
                        }
                    });
                    return true; // close the launcher so Outlook takes focus
                }
            }
        };
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
                IcoPath = OutlookIcon,
                Score = 5000,
                Action = _ =>
                {
                    try
                    {
                        _outlookTasks!.Add(parsed.Title, parsed.Priority, parsed.Category,
                            parsed.DueDate, parsed.Recurrence);
                        InvalidateOutlookCache();
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
        lock (_outlookCacheLock)
        {
            var hasData = _outlookCacheTasks != null || _outlookCacheError != null;
            var cacheIsFresh = hasData
                               && _outlookCacheAtTicks is long at
                               && (Environment.TickCount64 - at) < OutlookCacheTtlMs;

            if (cacheIsFresh)
            {
                return _outlookCacheTasks != null
                    ? RenderOutlookList(_outlookCacheTasks)
                    : OutlookErrorResults(_outlookCacheError!);
            }

            if (!_outlookFetching)
            {
                _outlookFetching = true;
                StartOutlookFetch();
            }

            // Stale-while-revalidate: if a prior successful list is still held, render it while the
            // background refetch runs rather than blanking to a "Loading…" row on every keystroke
            // once the TTL lapses. The fetch's ReQuery swaps in fresh data when it lands.
            if (_outlookCacheTasks != null)
                return RenderOutlookList(_outlookCacheTasks);
        }

        return new List<Result>
        {
            new()
            {
                Title = "Loading Outlook tasks…",
                SubTitle = "Querying Outlook via the COM bridge",
                IcoPath = OutlookIcon,
                Score = 1000,
                Action = _ => false
            }
        };
    }

    // Fetches the Outlook list off the query thread, caches the outcome, then ReQueries so
    // BuildOutlookListResults re-runs and renders the cached data through the normal path.
    // Caller holds _outlookCacheLock and has set _outlookFetching = true.
    private void StartOutlookFetch()
    {
        var gen = _outlookCacheGeneration;
        Task.Run(() =>
        {
            List<TodoItem>? tasks = null;
            string? error = null;
            try
            {
                tasks = _outlookTasks!.List(includeCompleted: true)
                    .OrderBy(t => t.IsCompleted)                    // incomplete first, completed sink to the bottom
                    .ThenBy(t => t.DueDate ?? DateTime.MaxValue)
                    .ThenByDescending(t => t.Priority)
                    .ToList();
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            lock (_outlookCacheLock)
            {
                // A mutation (add/complete/delete) invalidated the cache while this read was in
                // flight — its snapshot predates the change, so discard it and refetch instead of
                // publishing stale data over the invalidation. _outlookFetching stays true.
                if (gen != _outlookCacheGeneration)
                {
                    StartOutlookFetch();
                    return;
                }

                _outlookCacheTasks = tasks;
                _outlookCacheError = error;
                _outlookCacheAtTicks = Environment.TickCount64;
                _outlookFetching = false;
            }

            _context.API.ReQuery();
        });
    }

    // Drops the cached list so the next `tdo list` refetches — call after a mutation
    // (add / complete) so the change shows up immediately rather than after the TTL.
    private void InvalidateOutlookCache()
    {
        lock (_outlookCacheLock)
        {
            _outlookCacheAtTicks = null;
            _outlookCacheTasks = null;
            _outlookCacheError = null;
            _outlookCacheGeneration++;   // fence any in-flight fetch from republishing a pre-mutation snapshot
        }
    }

    private List<Result> RenderOutlookList(List<TodoItem> tasks)
    {
        if (tasks.Count == 0)
        {
            return new List<Result>
            {
                new()
                {
                    Title = "No Outlook tasks found",
                    SubTitle = "Add one with: tdo <task>",
                    IcoPath = OutlookIcon,
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

    private static List<Result> OutlookErrorResults(string message)
    {
        return new List<Result>
        {
            new()
            {
                Title = "Unable to read Outlook tasks",
                SubTitle = message,
                IcoPath = "Images\\td-high.png",
                Score = 1000
            }
        };
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
                    IcoPath = "Images\\td-high.png",
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
                    IcoPath = "Images\\td-high.png",
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
            IcoPath = diag.Ok ? "Images\\td-done.png" : "Images\\td-high.png",
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
                IcoPath = step.Ok ? "Images\\td-done.png" : "Images\\td-high.png",
                Score = score--
            });
        }

        if (!string.IsNullOrWhiteSpace(diag.Error))
        {
            results.Add(new Result
            {
                Title = "Error detail",
                SubTitle = diag.Error,
                IcoPath = "Images\\td-high.png",
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
        if (d.InboxItemCount.HasValue) parts.Add($"inbox: {d.InboxItemCount}");
        if (d.AccountCount.HasValue) parts.Add($"accounts: {d.AccountCount}");
        if (d.StoreCount.HasValue) parts.Add($"stores: {d.StoreCount}");
        if (d.ExplorerAvailable.HasValue) parts.Add($"explorer: {(d.ExplorerAvailable.Value ? "yes" : "no")}");
        if (d.ComApartmentState != null) parts.Add(d.ComApartmentState);
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

    // --- MARKDOWN PREVIEW ---

    /// <summary>Cap on completed tasks shown per category in the preview pane.</summary>
    private const int MaxDonePerCategory = 5;

    // PreviewInfo.ContentType (and the PreviewContentType enum) is a fork-only addition for the
    // markdown preview pane; the published Plugin API (NuGet 5.2.0) lacks it. Resolved by reflection
    // against the running assembly so it compiles against NuGet yet still drives markdown on the fork.
    private static readonly System.Reflection.PropertyInfo? ContentTypeProp =
        typeof(Result.PreviewInfo).GetProperty("ContentType");

    /// <summary>
    /// Preview pane payload rendered as markdown. On Flow Launcher builds with markdown preview
    /// support the pane auto-opens for these results; older builds fall back to showing the
    /// description as plain text.
    /// </summary>
    private static Result.PreviewInfo MarkdownPreview(string markdown)
    {
        var info = new Result.PreviewInfo { Description = markdown };

        if (ContentTypeProp != null)
        {
            try
            {
                ContentTypeProp.SetValue(info, Enum.Parse(ContentTypeProp.PropertyType, "Markdown"));
            }
            catch { /* enum shape differs on this host — leave as plain text */ }
        }

        return info;
    }

    /// <summary>
    /// Renders the whole local list as a markdown dashboard grouped by category.
    /// <paramref name="highlightId"/> bolds the task the selected row refers to.
    /// </summary>
    private string BuildListMarkdown(Guid? highlightId = null)
    {
        var tasks = _store.GetAll();
        var open = tasks.Count(t => !t.IsCompleted);
        var done = tasks.Count - open;

        var sb = new StringBuilder();
        sb.AppendLine($"## Tasks — {open} open · {done} done");

        if (tasks.Count == 0)
        {
            sb.AppendLine();
            sb.AppendLine("*No tasks yet — type after `td` to add one.*");
            return sb.ToString();
        }

        foreach (var group in tasks
                     .GroupBy(t => t.Category, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
        {
            sb.AppendLine();
            sb.AppendLine($"### {EscapeMarkdown(group.Key)}");

            var ordered = group
                .OrderBy(t => t.IsCompleted)
                .ThenByDescending(t => t.Priority)
                .ThenBy(t => t.DueDate ?? DateTime.MaxValue);

            var doneShown = 0;
            var doneHidden = 0;
            foreach (var task in ordered)
            {
                if (task.IsCompleted && ++doneShown > MaxDonePerCategory)
                {
                    doneHidden++;
                    continue;
                }
                AppendTaskLine(sb, task, task.Id == highlightId);
            }

            if (doneHidden > 0)
                sb.AppendLine($"*…and {doneHidden} more done*  ");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Add/Update preview: the task as it will be saved, then the current list for context.
    /// </summary>
    private string BuildAddMarkdown(string heading, ParsedInput parsed)
    {
        var meta = new List<string> { parsed.Priority.ToString(), parsed.Category };
        if (parsed.DueDate.HasValue)
            meta.Add($"due {FormatDuePreview(parsed.DueDate.Value, parsed.HasDueTime)}");
        if (parsed.Recurrence != Recurrence.None)
            meta.Add(RecurrenceLabel(parsed.Recurrence).ToLowerInvariant());

        var sb = new StringBuilder();
        sb.AppendLine($"## {heading}");
        sb.AppendLine();
        sb.AppendLine($"☐ **{EscapeMarkdown(parsed.Title)}** — *{string.Join(" · ", meta)}*");

        if (parsed.CategoryWarning != null)
        {
            sb.AppendLine();
            sb.AppendLine($"> ⚠ {EscapeMarkdown(parsed.CategoryWarning)}");
        }

        sb.AppendLine();
        sb.Append(BuildListMarkdown());
        return sb.ToString();
    }

    // Task lines are plain paragraph lines ending in a markdown hard break (two trailing
    // spaces) — list syntax would add MdXaml's own bullet next to the ☐/☑ glyph.
    private static void AppendTaskLine(StringBuilder sb, TodoItem task, bool highlight)
    {
        var title = EscapeMarkdown(task.Title);
        if (task.IsCompleted)
        {
            sb.AppendLine($"☑ ~~{title}~~  ");
            return;
        }

        var meta = new List<string> { task.Priority.ToString() };
        if (task.DueDate.HasValue)
            meta.Add(FormatDue(task));
        if (task.Recurrence != Recurrence.None)
            meta.Add(RecurrenceLabel(task.Recurrence).ToLowerInvariant());

        sb.AppendLine(highlight
            ? $"☐ **{title}** — *{string.Join(" · ", meta)}* ◀  "
            : $"☐ {title} — *{string.Join(" · ", meta)}*  ");
    }

    /// <summary>Backslash-escapes characters MdXaml would treat as markup inside titles.</summary>
    private static string EscapeMarkdown(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (c is '\\' or '*' or '_' or '`' or '#' or '[' or ']' or '~' or '>' or '|')
                sb.Append('\\');
            sb.Append(c);
        }
        return sb.ToString();
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
        // Completed tasks get the blue check (matches the tdo list); incomplete keep their
        // priority-coloured icon so priority stays readable at a glance.
        var result = new Result
        {
            Title = task.Title,
            SubTitle = subTitle,
            IcoPath = task.IsCompleted ? CheckedIcon : PriorityIcon(task.Priority),
            Score = score,
            ContextData = task,
            Preview = MarkdownPreview(BuildListMarkdown(task.Id)),
        };

        result.Action = _ =>
        {
            _store.ToggleComplete(task.Id);

            // Swap just this row's icon in place — no ReQuery — so the list does not rebuild
            // and the preview pane does not flash. We set Result.Icon (not IcoPath) because the
            // shell thumbnail provider is broken on this machine; ReloadResultImage then reloads
            // this single row from the new icon source.
            var completed = _store.GetById(task.Id)?.IsCompleted ?? task.IsCompleted;
            var image = _iconLoader.LoadIcon(completed ? CheckedIcon : PriorityIcon(task.Priority));
            if (image != null)
            {
                result.Icon = () => image;
                result.IcoPath = string.Empty;
            }

            TryReloadResultImage(result);
            return false;
        };

        return result;
    }

    // ReloadResultImage refreshes a single row's icon in place. It exists on the dev/fork build
    // of Flow Launcher but not in the published Plugin API (NuGet 5.2.0), so we resolve it off the
    // live API's concrete type by reflection: present on the fork → refresh in place; absent on
    // standard FL → harmless no-op. Resolved once and cached.
    private System.Reflection.MethodInfo? _reloadResultImage;
    private bool _reloadResultImageResolved;

    private void TryReloadResultImage(Result result)
    {
        if (!_reloadResultImageResolved)
        {
            _reloadResultImage = _context.API.GetType()
                .GetMethod("ReloadResultImage", new[] { typeof(Result) });
            _reloadResultImageResolved = true;
        }

        _reloadResultImage?.Invoke(_context.API, new object[] { result });
    }

    private Result OutlookTaskToResult(TodoItem task, int score)
    {
        return new Result
        {
            Title = task.Title,
            SubTitle = FormatSubTitle(task),
            IcoPath = task.IsCompleted ? CheckedIcon : UncheckedIcon,
            Score = score,
            ContextData = task,
            Action = _ =>
            {
                var nowComplete = !task.IsCompleted;
                try
                {
                    _outlookTasks!.SetComplete(task, complete: nowComplete);
                    InvalidateOutlookCache();
                    _context.API.ShowMsg("QuickTodo Outlook",
                        nowComplete ? $"Completed: {task.Title}" : $"Marked incomplete: {task.Title}");
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

    // Context menu (Shift+Enter) for an Outlook task: toggle done state + delete. Outlook
    // tasks aren't in the local store, so Main routes them here instead of the local menu.
    internal List<Result> BuildOutlookContextMenus(TodoItem task)
    {
        var willComplete = !task.IsCompleted;

        return new List<Result>
        {
            new()
            {
                Title = willComplete ? "Mark Complete" : "Mark Incomplete",
                SubTitle = willComplete
                    ? "Check it off (stays in the list until deleted)"
                    : "Move it back to your active tasks",
                IcoPath = willComplete ? CheckedIcon : UncheckedIcon,
                Action = _ =>
                {
                    try
                    {
                        _outlookTasks!.SetComplete(task, complete: willComplete);
                        InvalidateOutlookCache();
                        _context.API.ShowMsg("QuickTodo Outlook",
                            willComplete ? $"Completed: {task.Title}" : $"Marked incomplete: {task.Title}");
                        _context.API.ReQuery();
                    }
                    catch (Exception ex)
                    {
                        _context.API.ShowMsg("QuickTodo Outlook", ex.Message);
                    }
                    return false;
                }
            },
            new()
            {
                Title = "Delete from Outlook",
                SubTitle = task.Title,
                IcoPath = "Images\\td-high.png",
                Action = _ =>
                {
                    try
                    {
                        _outlookTasks!.Delete(task);
                        InvalidateOutlookCache();
                        _context.API.ShowMsg("QuickTodo Outlook", $"Deleted: {task.Title}");
                        _context.API.ReQuery();
                    }
                    catch (Exception ex)
                    {
                        _context.API.ShowMsg("QuickTodo Outlook", ex.Message);
                    }
                    return false;
                }
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
        Priority.High => "Images\\td-high.png",
        Priority.Medium => "Images\\td-medium.png",
        Priority.Low => "Images\\td-low.png",
        _ => "Images\\td.png"
    };

    private static int PriorityScore(Priority p) => p switch
    {
        Priority.High => 300,
        Priority.Medium => 200,
        Priority.Low => 100,
        _ => 0
    };
}
