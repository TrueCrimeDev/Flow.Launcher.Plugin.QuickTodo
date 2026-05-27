$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $failures.Add("$Name missing: $Path")
    }
}

function Assert-SourceContains {
    param(
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $Pattern,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        $failures.Add("$Name source missing: $Path")
        return
    }

    $source = Get-Content -LiteralPath $Path -Raw
    if ($source -notmatch $Pattern) {
        $failures.Add("$Name did not match pattern: $Pattern")
    }
}

$clientPath = Join-Path $repoRoot 'Services\OutlookTaskScriptClient.cs'
$queryHandlerPath = Join-Path $repoRoot 'Services\QueryHandler.cs'
$mainPath = Join-Path $repoRoot 'Main.cs'
$todoItemPath = Join-Path $repoRoot 'Models\TodoItem.cs'

Assert-FileExists $clientPath 'Outlook task script client'
Assert-SourceContains $clientPath 'QuickTodo\.OutlookTasks\.ps1' 'Client points at the COM bridge script'
Assert-SourceContains $clientPath 'powershell\.exe' 'Client launches PowerShell'
Assert-SourceContains $clientPath '"-File"' 'Client passes script via -File'
Assert-SourceContains $clientPath 'ConvertFromOutlookRecord' 'Client maps Outlook JSON into TodoItem records'

Assert-SourceContains $queryHandlerPath '"outlook"\s*=>\s*BuildOutlookResults' 'Flow command includes td outlook route'
Assert-SourceContains $queryHandlerPath 'Add to Outlook:' 'Flow add result writes to Outlook'
Assert-SourceContains $queryHandlerPath 'List Outlook tasks' 'Flow exposes Outlook task list'

Assert-SourceContains $mainPath 'new OutlookTaskScriptClient' 'Plugin constructs Outlook client'
Assert-SourceContains $mainPath 'Scripts.*QuickTodo\.OutlookTasks\.ps1' 'Plugin uses built script path'

Assert-SourceContains $todoItemPath 'OutlookEntryId' 'TodoItem carries Outlook EntryId for mutations'
Assert-SourceContains $todoItemPath 'OutlookStoreId' 'TodoItem carries Outlook StoreId for mutations'

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Host 'QuickTodo.FlowOutlook.Tests.ps1: all tests passed'
