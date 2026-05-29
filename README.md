# QuickTodo for Flow Launcher

QuickTodo is a Flow Launcher plugin for creating and reviewing lightweight tasks with priorities, categories, due dates, reminders, and optional Outlook Tasks support.

## Commands

- `td <task>` adds a local QuickTodo task.
- `td <task> !high @Work #tomorrow` adds a local task with priority, category, and due date.
- `td list` lists local tasks.
- `td edit` lists tasks to edit; pick one to prefill `td edit <id> <title>`, then change the title or any modifier and press Enter to save. Also available from a task's context menu ("Edit Task").
- `td outlook <task> !low @Work #tomorrow` creates a real Outlook task through desktop Outlook's COM object model.
- `td outlook list` lists incomplete Outlook tasks. Press Enter on a result to mark it complete.
- `td outlook diag` (or `tdo diag`) probes the Outlook COM connector step by step (bind, MAPI namespace, profile, default Tasks folder, task counts) and reports where it fails. Press Enter on the summary row to copy the full diagnostics JSON. Each bridge invocation is also logged to the Flow Launcher log.
- `tdo <task> !low @Work #tomorrow` is a shortcut for `td outlook <task>` — the `tdo` keyword goes straight to Outlook mode.
- `tdo list` lists incomplete Outlook tasks (same as `td outlook list`).
- `td cat add <name>` adds a local category.
- `td cat remove <name>` removes an unused local category.

## Date and priority modifiers

- Priorities: `!low`, `!medium`, `!high` or `!l`, `!m`, `!h`.
- Dates: `#today`, `#tomorrow`, `#monday`, `#yyyy-MM-dd`, or `#MM-dd`.
- Times: append `@<time>` to a date, e.g. `#tomorrow@1430`, `#friday@9am`, `#today@17:00`. Accepts `HHmm`, `H:mm`, and `h[:mm]am/pm`. A timed task only counts as overdue once its time passes, and its reminder fires at that time.
- Recurrence: `#daily`, `#weekly`, `#monthly`, `#yearly`, or `#every-monday` … `#every-sunday`. For local tasks, completing a recurring task rolls its due date forward to the next occurrence instead of marking it done. For Outlook tasks (`td outlook` / `tdo`), recurrence is applied as a native Outlook `RecurrencePattern` so Outlook regenerates the task itself. Combine with a time, e.g. `#daily@9am` (times apply to local tasks only; Outlook tasks store the date).
- Categories: `@Work`, `@Personal`, `@Errands`, or custom local categories.

## Outlook support

The Outlook path uses desktop Outlook automation only:

- `Outlook.Application`
- `GetNamespace("MAPI")`
- `CreateItem(3)` for tasks
- `GetDefaultFolder(13)` for the default Tasks folder
- `Save()`, `Delete()`, and the task object model properties for mutations

Outlook sync is handled by Outlook itself. If your default Tasks folder is backed by Exchange or Microsoft 365, saved tasks should sync to the mailbox and Microsoft To Do.

## Manual install

Download the release zip, then in Flow Launcher run:

```text
pm install <path-or-url-to-zip>
```

After install or update, restart Flow Launcher if the plugin was already loaded.
