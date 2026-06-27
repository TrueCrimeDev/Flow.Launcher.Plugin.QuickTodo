# Outlook Search (`os`) + Extended COM Debugging — Design

Date: 2026-06-26
Status: Approved

## Goal

Two additions to the QuickTodo Flow Launcher plugin:

1. **`os` Outlook search** — a dedicated action keyword. `os <term>` drives Outlook's
   own Instant Search over mail and brings Outlook to the front showing live results.
2. **Extended COM debugging** — more probe steps and reported fields in the existing
   `diag` report, plus full stdout/stderr capture in the bridge client on failure.

## 1. `os` Outlook search

### Behavior
- Trigger: dedicated Flow Launcher action keyword `os` (registered at runtime, like
  the existing `tdo`; collision-safe).
- While typing: no COM call. A single instant result renders:
  `🔍 Search Outlook for "<term>"`.
- Empty `os`: a hint row explaining usage (+ an `os diag` row).
- On Enter: shell the bridge `search` command, which activates Outlook and runs its
  Instant Search over **mail**, scope = all mail folders.
- Outlook search operators (`from:`, `subject:`, etc.) pass through verbatim.
- Latency note: the COM round-trip (powershell startup ~0.5–1.5 s) happens only on
  Enter, never per keystroke, so the picker stays responsive.

### PowerShell bridge — new `search` command
`Scripts/QuickTodo.OutlookTasks.ps1`:
- Add `search` to the `[ValidateSet(...)]` for `$Command`.
- Add `-Query [string]` and `-Scope [ValidateSet('CurrentFolder','AllFolders','Subfolders','AllOutlookItems')]='AllFolders'` params.
- New `Invoke-OutlookSearch`:
  - Bind Outlook (`Get-OutlookApplication`).
  - Get Inbox: `$ns.GetDefaultFolder(6)` (`olFolderInbox`).
  - Reuse `ActiveExplorer()` if present, set `CurrentFolder = $inbox` to force mail
    context; otherwise `$inbox.GetExplorer()` + `Display()`.
  - `$explorer.Activate()` then `$explorer.Search($Query, $scopeValue)`
    (`olSearchScopeAllFolders = 1`).
  - Return `{ Query, Scope, Ok }`.
- Wire `search` into the command `switch` in `Invoke-QuickTodoOutlookTaskCommand`.
  It binds Outlook itself before the `switch`, so `search` fits the normal path
  (it does not need the pre-bind special-case that `diag` uses).

### `OutlookTaskScriptClient.cs`
- New `OutlookSearchResult Search(string query)` (mirrors `Diagnose()`): builds args
  `search -Query <term> -AsJson`, runs, deserializes `{ Query, Scope, Ok }`.
- New `OutlookSearchResult` record type.

### `QueryHandler.cs`
- `public const string SearchActionKeyword = "os";`
- In `Handle`: if `query.ActionKeyword == "os"` → `BuildOutlookSearchResults(search)`.
- `BuildOutlookSearchResults(term)`:
  - Null client → "Outlook bridge not available" row (reuse existing pattern).
  - Empty term → hint row (`AutoCompleteText = "os "`) + an `os diag` row.
  - `diag` subcommand → reuse `BuildOutlookDiagResults()`.
  - Else → single result `Search Outlook for "<term>"`; `Action` calls
    `_outlookTasks.Search(term)` inside try/catch, `ShowMsg` on success/failure,
    returns `true` (closes the launcher so Outlook takes focus).

## 2. Extended COM debugging

### `Invoke-OutlookDiagnostics` (PowerShell) — new probe steps + fields
Add steps (each independent, never throws — appended to the report):
- `Get default Inbox folder (6)` — search prerequisite; sets `InboxFolderName`,
  `InboxItemCount`.
- `ActiveExplorer available` — sets `ExplorerAvailable` (true/false); confirms the
  search UI can be driven.
- `Enumerate accounts` — sets `AccountCount` (+ names in detail).
- `Enumerate stores` — sets `StoreCount` (+ names in detail).
- Capture `ComApartmentState` = `[System.Threading.Thread]::CurrentThread.GetApartmentState()`.

New fields on the report object (all nullable): `InboxFolderName`, `InboxItemCount`,
`AccountCount`, `StoreCount`, `ExplorerAvailable`, `ComApartmentState`.

### `OutlookDiagnostics` (C#)
Add matching properties: `InboxFolderName`, `InboxItemCount` (`int?`),
`AccountCount` (`int?`), `StoreCount` (`int?`), `ExplorerAvailable` (`bool?`),
`ComApartmentState` (`string?`).

### `BuildOutlookDiagResults` / `BuildDiagSummary` (C#)
The step rows already render generically from `diag.Steps`, so new steps show with no
change. Extend `BuildDiagSummary` to surface the high-value new fields
(e.g. `inbox: <n>`, `accounts: <n>`, `explorer: yes/no`, apartment state).

### `OutlookTaskScriptClient.Run` — full failure capture
On non-zero exit or thrown exception, log the **full** stdout + stderr via `_logWarn`
(today only char counts are logged on success). Keeps the success path quiet; makes a
real failure fully diagnosable from the Flow Launcher log.

## 3. Registration & manifest

- `Main.cs`: generalize `RegisterOutlookActionKeyword` to register a set of runtime
  keywords (`tdo`, `os`) with the same collision-safe guards
  (`meta.ActionKeywords.Contains` / `API.ActionKeywordAssigned`).
- `plugin.json`: bump `1.4.0 → 1.5.0`; mention `os` Outlook search in the description.

## Out of scope (YAGNI)
- Rendering email results inline in Flow Launcher.
- Searching tasks/calendar/contacts via `os` (mail only for now; `-Scope` param exists
  in the bridge for a future toggle but the UI always passes `AllFolders`).
- A separate search-only diagnostics report (the shared `diag` covers both paths).

## Test plan
- Build (Windows `dotnet.exe`, Release) → deploy to the fork's
  `Output/Debug/Plugins/...` → restart `Output/Debug/Flow.Launcher.exe`.
- Verify: FL process up; no plugin-load errors in `Output/Debug/UserData/Logs/`;
  `os foo` shows the search row; `os diag` (or `tdo diag`) shows the new probe steps
  and fields. Live Enter→Outlook search verified by the user (Outlook COM is
  unavailable in the WSL/sandbox env).
