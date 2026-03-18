namespace Flow.Launcher.Plugin.QuickTodo.Models;

public class TodoData
{
    public List<TodoItem> Tasks { get; set; } = new();
    public List<string> Categories { get; set; } = new() { "Work", "Personal", "Errands" };
}
