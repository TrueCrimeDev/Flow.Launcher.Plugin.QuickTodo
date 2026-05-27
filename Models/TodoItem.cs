using System.Text.Json.Serialization;

namespace Flow.Launcher.Plugin.QuickTodo.Models;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum Priority
{
    Low,
    Medium,
    High
}

public class TodoItem
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = string.Empty;
    public Priority Priority { get; set; } = Priority.Medium;
    public string Category { get; set; } = "Personal";
    public DateTime? DueDate { get; set; }
    public bool IsCompleted { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime? CompletedAt { get; set; }
    public DateTime? SnoozedUntil { get; set; }
    public string? OutlookEntryId { get; set; }
    public string? OutlookStoreId { get; set; }
}
