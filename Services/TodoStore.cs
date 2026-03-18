using System.IO;
using System.Text.Json;
using Flow.Launcher.Plugin.QuickTodo.Models;

namespace Flow.Launcher.Plugin.QuickTodo.Services;

public class TodoStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly string _filePath;
    private readonly string _backupPath;
    private readonly object _lock = new();
    private TodoData _data = new();
    private Action<string, string>? _logWarn;

    public TodoStore(string? dataDir = null, Action<string, string>? logWarn = null)
    {
        var dir = dataDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "QuickTodo");
        Directory.CreateDirectory(dir);
        _filePath = Path.Combine(dir, "quicktodo.json");
        _backupPath = Path.Combine(dir, "quicktodo.backup.json");
        _logWarn = logWarn;
    }

    public void Load()
    {
        lock (_lock)
        {
            _data = TryLoadFile(_filePath)
                    ?? TryLoadFile(_backupPath)
                    ?? new TodoData();
        }
    }

    private TodoData? TryLoadFile(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<TodoData>(json, JsonOptions);
        }
        catch (Exception ex)
        {
            _logWarn?.Invoke(nameof(TodoStore), $"Failed to load {path}: {ex.Message}");
            return null;
        }
    }

    private void Save()
    {
        // Caller must hold _lock
        try
        {
            if (File.Exists(_filePath))
                File.Copy(_filePath, _backupPath, overwrite: true);
        }
        catch { /* best-effort backup */ }

        var json = JsonSerializer.Serialize(_data, JsonOptions);
        File.WriteAllText(_filePath, json);
    }

    // --- CRUD ---

    public List<TodoItem> GetAll()
    {
        lock (_lock) return _data.Tasks.ToList();
    }

    public TodoItem? GetById(Guid id)
    {
        lock (_lock) return _data.Tasks.FirstOrDefault(t => t.Id == id);
    }

    public TodoItem Add(string title, Priority priority, string category, DateTime? dueDate)
    {
        lock (_lock)
        {
            var item = new TodoItem
            {
                Title = title,
                Priority = priority,
                Category = category,
                DueDate = dueDate
            };
            _data.Tasks.Add(item);
            Save();
            return item;
        }
    }

    public void Update(TodoItem item)
    {
        lock (_lock)
        {
            var idx = _data.Tasks.FindIndex(t => t.Id == item.Id);
            if (idx >= 0)
            {
                _data.Tasks[idx] = item;
                Save();
            }
        }
    }

    public void Delete(Guid id)
    {
        lock (_lock)
        {
            _data.Tasks.RemoveAll(t => t.Id == id);
            Save();
        }
    }

    public void ToggleComplete(Guid id)
    {
        lock (_lock)
        {
            var item = _data.Tasks.FirstOrDefault(t => t.Id == id);
            if (item == null) return;
            item.IsCompleted = !item.IsCompleted;
            item.CompletedAt = item.IsCompleted ? DateTime.Now : null;
            Save();
        }
    }

    public void SetPriority(Guid id, Priority priority)
    {
        lock (_lock)
        {
            var item = _data.Tasks.FirstOrDefault(t => t.Id == id);
            if (item == null) return;
            item.Priority = priority;
            Save();
        }
    }

    public void SetCategory(Guid id, string category)
    {
        lock (_lock)
        {
            var item = _data.Tasks.FirstOrDefault(t => t.Id == id);
            if (item == null) return;
            item.Category = category;
            Save();
        }
    }

    public void SetDueDate(Guid id, DateTime? dueDate)
    {
        lock (_lock)
        {
            var item = _data.Tasks.FirstOrDefault(t => t.Id == id);
            if (item == null) return;
            item.DueDate = dueDate;
            Save();
        }
    }

    public void SetSnoozedUntil(Guid id, DateTime? until)
    {
        lock (_lock)
        {
            var item = _data.Tasks.FirstOrDefault(t => t.Id == id);
            if (item == null) return;
            item.SnoozedUntil = until;
            Save();
        }
    }

    // --- Categories ---

    public List<string> GetCategories()
    {
        lock (_lock) return _data.Categories.ToList();
    }

    public bool AddCategory(string name)
    {
        lock (_lock)
        {
            if (_data.Categories.Any(c => c.Equals(name, StringComparison.OrdinalIgnoreCase)))
                return false;
            _data.Categories.Add(name);
            Save();
            return true;
        }
    }

    public bool RemoveCategory(string name)
    {
        lock (_lock)
        {
            if (_data.Tasks.Any(t => t.Category.Equals(name, StringComparison.OrdinalIgnoreCase)))
                return false; // category in use
            var removed = _data.Categories.RemoveAll(
                c => c.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (removed > 0) Save();
            return removed > 0;
        }
    }

    // --- Queries ---

    public List<TodoItem> GetOverdue()
    {
        lock (_lock)
        {
            var now = DateTime.Now;
            return _data.Tasks
                .Where(t => !t.IsCompleted
                    && t.DueDate.HasValue
                    && t.DueDate.Value.Date < now.Date
                    && (!t.SnoozedUntil.HasValue || t.SnoozedUntil.Value < now))
                .ToList();
        }
    }

    public List<TodoItem> GetDueToday()
    {
        lock (_lock)
        {
            return _data.Tasks
                .Where(t => !t.IsCompleted
                    && t.DueDate.HasValue
                    && t.DueDate.Value.Date == DateTime.Now.Date)
                .ToList();
        }
    }
}
