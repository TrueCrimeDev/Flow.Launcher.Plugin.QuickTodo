namespace Flow.Launcher.Plugin.QuickTodo.Settings;

public class QuickTodoSettings
{
    public int ReminderIntervalMinutes { get; set; } = 60;
    public int SnoozeDurationMinutes { get; set; } = 10;
    public bool NotificationSoundEnabled { get; set; } = true;
}
