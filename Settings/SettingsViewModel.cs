using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Flow.Launcher.Plugin.QuickTodo.Services;

namespace Flow.Launcher.Plugin.QuickTodo.Settings;

public class SettingsViewModel : INotifyPropertyChanged
{
    private readonly QuickTodoSettings _settings;
    private readonly TodoStore _store;
    private string _newCategory = string.Empty;

    public SettingsViewModel(QuickTodoSettings settings, TodoStore store)
    {
        _settings = settings;
        _store = store;
        Categories = new ObservableCollection<string>(_store.GetCategories());
    }

    public int ReminderIntervalMinutes
    {
        get => _settings.ReminderIntervalMinutes;
        set
        {
            if (value >= 5)
            {
                _settings.ReminderIntervalMinutes = value;
                OnPropertyChanged();
            }
        }
    }

    public int SnoozeDurationMinutes
    {
        get => _settings.SnoozeDurationMinutes;
        set
        {
            if (value >= 1)
            {
                _settings.SnoozeDurationMinutes = value;
                OnPropertyChanged();
            }
        }
    }

    public bool NotificationSoundEnabled
    {
        get => _settings.NotificationSoundEnabled;
        set
        {
            _settings.NotificationSoundEnabled = value;
            OnPropertyChanged();
        }
    }

    public ObservableCollection<string> Categories { get; }

    public string NewCategory
    {
        get => _newCategory;
        set { _newCategory = value; OnPropertyChanged(); }
    }

    public void AddCategory()
    {
        if (string.IsNullOrWhiteSpace(NewCategory)) return;
        if (_store.AddCategory(NewCategory))
        {
            Categories.Add(NewCategory);
            NewCategory = string.Empty;
        }
    }

    public void RemoveCategory(string name)
    {
        if (_store.RemoveCategory(name))
            Categories.Remove(name);
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
