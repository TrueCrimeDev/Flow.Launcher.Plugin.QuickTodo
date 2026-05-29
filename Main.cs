using System.IO;
using System.Windows.Controls;
using Flow.Launcher.Plugin;
using Flow.Launcher.Plugin.QuickTodo.Models;
using Flow.Launcher.Plugin.QuickTodo.Services;
using Flow.Launcher.Plugin.QuickTodo.Settings;
using Microsoft.Toolkit.Uwp.Notifications;

namespace Flow.Launcher.Plugin.QuickTodo;

public class Main : IAsyncPlugin, IContextMenu, ISettingProvider, IDisposable
{
    private PluginInitContext _context = null!;
    private TodoStore _store = null!;
    private OutlookTaskScriptClient _outlookTasks = null!;
    private QueryHandler _queryHandler = null!;
    private ReminderService _reminderService = null!;
    private QuickTodoSettings _settings = null!;
    private SettingsViewModel _settingsViewModel = null!;
    private OnActivated? _toastHandler;

    public Task InitAsync(PluginInitContext context)
    {
        _context = context;

        RegisterOutlookActionKeyword(context);

        _settings = context.API.LoadSettingJsonStorage<QuickTodoSettings>();

        _store = new TodoStore(
            logWarn: (cls, msg) => context.API.LogWarn(cls, msg));
        _store.Load();

        _outlookTasks = new OutlookTaskScriptClient(
            Path.Combine(AppContext.BaseDirectory, "Scripts", "QuickTodo.OutlookTasks.ps1"),
            logWarn: (cls, msg) => context.API.LogWarn(cls, msg),
            logInfo: (cls, msg) => context.API.LogInfo(cls, msg));

        _queryHandler = new QueryHandler(_store, context, _outlookTasks);

        _reminderService = new ReminderService(
            _store,
            () => _settings.ReminderIntervalMinutes,
            () => _settings.SnoozeDurationMinutes,
            () => _settings.NotificationSoundEnabled);
        _reminderService.Start();

        _toastHandler = args => _reminderService.HandleToastAction(args.Argument);
        ToastNotificationManagerCompat.OnActivated += _toastHandler;

        _settingsViewModel = new SettingsViewModel(_settings, _store);

        return Task.CompletedTask;
    }

    // Registers the dedicated "tdo" keyword so it goes straight to Outlook mode,
    // alongside the default "td" keyword from plugin.json. Idempotent across restarts,
    // and skips registration if another plugin already owns the keyword.
    private static void RegisterOutlookActionKeyword(PluginInitContext context)
    {
        var meta = context.CurrentPluginMetadata;
        if (meta.ActionKeywords.Contains(QueryHandler.OutlookActionKeyword))
            return;
        if (context.API.ActionKeywordAssigned(QueryHandler.OutlookActionKeyword))
            return;

        context.API.AddActionKeyword(meta.ID, QueryHandler.OutlookActionKeyword);
    }

    public Task<List<Result>> QueryAsync(Query query, CancellationToken token)
    {
        return Task.FromResult(_queryHandler.Handle(query));
    }

    public List<Result> LoadContextMenus(Result selectedResult)
    {
        if (selectedResult.ContextData is not TodoItem contextTask)
            return new List<Result>();

        // Re-fetch from store to get fresh state. A null local lookup means this is
        // an Outlook-backed task (not in the local store), so local-only actions are gated.
        var localTask = _store.GetById(contextTask.Id);
        var task = localTask ?? contextTask;

        var results = new List<Result>();

        // Toggle complete
        results.Add(new Result
        {
            Title = task.IsCompleted ? "Mark Incomplete" : "Mark Complete",
            IcoPath = "Images\\todo-done.png",
            Action = _ =>
            {
                _store.ToggleComplete(task.Id);
                _context.API.ReQuery();
                return false;
            }
        });

        // Priority options
        foreach (var p in Enum.GetValues<Priority>())
        {
            var prefix = task.Priority == p ? "\u2713 " : "";
            results.Add(new Result
            {
                Title = $"{prefix}Priority: {p}",
                IcoPath = QueryHandler.PriorityIcon(p),
                Action = _ =>
                {
                    _store.SetPriority(task.Id, p);
                    _context.API.ReQuery();
                    return false;
                }
            });
        }

        // Category options
        foreach (var cat in _store.GetCategories())
        {
            var prefix = task.Category.Equals(cat, StringComparison.OrdinalIgnoreCase) ? "\u2713 " : "";
            results.Add(new Result
            {
                Title = $"{prefix}Category: {cat}",
                IcoPath = "Images\\todo.png",
                Action = _ =>
                {
                    _store.SetCategory(task.Id, cat);
                    _context.API.ReQuery();
                    return false;
                }
            });
        }

        // Set due date
        results.Add(new Result
        {
            Title = "Set Due Date",
            SubTitle = "Type a date after #",
            IcoPath = "Images\\todo.png",
            Action = _ =>
            {
                _context.API.ChangeQuery($"td {task.Title} #");
                return false;
            }
        });

        // Edit (title + all modifiers) — local tasks only
        if (localTask != null)
        {
            results.Add(new Result
            {
                Title = "Edit Task",
                SubTitle = "Change title, priority, category, due date, or recurrence",
                IcoPath = "Images\\todo.png",
                Action = _ =>
                {
                    _context.API.ChangeQuery($"td edit {task.Id} {task.Title}");
                    return false;
                }
            });
        }

        // Delete
        results.Add(new Result
        {
            Title = "Delete Task",
            SubTitle = task.Title,
            IcoPath = "Images\\todo-high.png",
            Action = _ =>
            {
                _store.Delete(task.Id);
                _context.API.ReQuery();
                return false;
            }
        });

        return results;
    }

    public Control CreateSettingPanel()
    {
        return new QuickTodoSettingsPanel(_settingsViewModel);
    }

    public void Dispose()
    {
        _reminderService?.Dispose();
        if (_toastHandler != null)
            ToastNotificationManagerCompat.OnActivated -= _toastHandler;
    }
}
