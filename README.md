# QuickTodo for Flow Launcher

QuickTodo is a Flow Launcher plugin for creating and reviewing lightweight tasks with priorities, categories, due dates, reminders, and optional Outlook Tasks support.

## Commands

- `td <task>` adds a local QuickTodo task.
- `td <task> !high @Work #tomorrow` adds a local task with priority, category, and due date.
- `td list` lists local tasks.
- `td outlook <task> !low @Work #tomorrow` creates a real Outlook task through desktop Outlook's COM object model.
- `td outlook list` lists incomplete Outlook tasks. Press Enter on a result to mark it complete.
- `td cat add <name>` adds a local category.
- `td cat remove <name>` removes an unused local category.

## Date and priority modifiers

- Priorities: `!low`, `!medium`, `!high` or `!l`, `!m`, `!h`.
- Dates: `#today`, `#tomorrow`, `#monday`, `#yyyy-MM-dd`, or `#MM-dd`.
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
