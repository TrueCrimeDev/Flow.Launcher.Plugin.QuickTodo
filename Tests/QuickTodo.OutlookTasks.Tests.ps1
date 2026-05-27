$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$scriptPath = Join-Path $PSScriptRoot '..\Scripts\QuickTodo.OutlookTasks.ps1'
. $scriptPath

$failures = [System.Collections.Generic.List[string]]::new()

function Assert-Equal {
    param(
        [Parameter(Mandatory = $true)] [object] $Expected,
        [Parameter(Mandatory = $true)] [object] $Actual,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    if ($Expected -ne $Actual) {
        $failures.Add("$Name expected '$Expected' but got '$Actual'")
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)] [bool] $Condition,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    if (-not $Condition) {
        $failures.Add("$Name expected true")
    }
}

function Assert-Null {
    param(
        [object] $Actual,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    if ($null -ne $Actual) {
        $failures.Add("$Name expected null but got '$Actual'")
    }
}

Assert-Equal 0 (ConvertTo-OutlookImportance -Importance Low) 'Low importance maps to olImportanceLow'
Assert-Equal 1 (ConvertTo-OutlookImportance -Importance Normal) 'Normal importance maps to olImportanceNormal'
Assert-Equal 1 (ConvertTo-OutlookImportance -Importance Medium) 'Medium importance maps to olImportanceNormal'
Assert-Equal 2 (ConvertTo-OutlookImportance -Importance High) 'High importance maps to olImportanceHigh'

Assert-Equal 'Work, Calls' (ConvertTo-OutlookCategoryString -Categories @(' Work ', '', 'Calls')) 'Categories trim blanks and join for Outlook'
Assert-Null (ConvertTo-OutlookCategoryString -Categories @('', ' ')) 'Blank categories return null'

$task = [pscustomobject]@{
    EntryID = 'entry-1'
    Parent = [pscustomobject]@{ StoreID = 'store-1' }
    Subject = 'Call Alex'
    DueDate = [datetime]'2026-06-03'
    Complete = $false
    Importance = 2
    Categories = 'Work, Calls'
    Body = 'Discuss schedule'
}

$record = ConvertFrom-OutlookTaskItem -Task $task
Assert-Equal 'entry-1' $record.EntryId 'Task record includes EntryId'
Assert-Equal 'store-1' $record.StoreId 'Task record includes StoreId'
Assert-Equal 'Call Alex' $record.Subject 'Task record includes Subject'
Assert-Equal '2026-06-03' $record.DueDate 'Task record formats DueDate as yyyy-MM-dd'
Assert-Equal $false $record.Complete 'Task record includes Complete'
Assert-Equal 'High' $record.Importance 'Task record converts Outlook importance'
Assert-Equal 'Work, Calls' $record.Categories 'Task record includes Categories'
Assert-Equal 'Discuss schedule' $record.Body 'Task record includes Body'

$undatedTask = [pscustomobject]@{
    EntryID = 'entry-2'
    Parent = [pscustomobject]@{ StoreID = 'store-2' }
    Subject = 'No due date'
    DueDate = [datetime]'4501-01-01'
    Complete = $false
    Importance = 1
    Categories = ''
    Body = ''
}

$undatedRecord = ConvertFrom-OutlookTaskItem -Task $undatedTask
Assert-Null $undatedRecord.DueDate 'Outlook no-date sentinel is normalized to null'
Assert-Equal 'Normal' $undatedRecord.Importance 'Normal Outlook importance is named'

$source = Get-Content -LiteralPath $scriptPath -Raw
Assert-True ($source -match 'Outlook\.Application') 'Script binds Outlook.Application'
Assert-True ($source -match 'GetNamespace\(''MAPI''\)') 'Script gets the MAPI namespace'
Assert-True ($source -match 'CreateItem\(3\)') 'Script creates olTaskItem tasks'
Assert-True ($source -match 'GetDefaultFolder\(13\)') 'Script reads olFolderTasks'
Assert-True ($source -match '\.Restrict\(''\[Complete\] = false''\)') 'Script restricts incomplete tasks'
Assert-True ($source -notmatch 'Graph|Redemption|PropertyAccessor') 'Script avoids non-object-model task paths'

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Host 'QuickTodo.OutlookTasks.Tests.ps1: all tests passed'
