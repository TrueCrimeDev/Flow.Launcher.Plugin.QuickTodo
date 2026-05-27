[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('add', 'list', 'complete', 'incomplete', 'rename', 'delete')]
    [string] $Command = 'list',

    [string] $Subject,
    [string] $Body,
    [datetime] $DueDate,

    [ValidateSet('Low', 'Normal', 'Medium', 'High')]
    [string] $Importance = 'Normal',

    [string[]] $Categories,
    [string] $EntryId,
    [string] $StoreId,
    [switch] $IncludeCompleted,
    [switch] $AsJson
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-OutlookApplication {
    $progId = 'Outlook.Application'

    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId)
    }
    catch {
        try {
            return New-Object -ComObject $progId
        }
        catch {
            throw "Unable to bind $progId. Confirm desktop Outlook is installed and a MAPI profile is configured. $($_.Exception.Message)"
        }
    }
}

function Get-OutlookNamespace {
    param(
        [Parameter(Mandatory = $true)] [object] $Application
    )

    return $Application.GetNamespace('MAPI')
}

function ConvertTo-OutlookImportance {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Low', 'Normal', 'Medium', 'High')]
        [string] $Importance
    )

    switch ($Importance) {
        'Low' { return 0 }
        'High' { return 2 }
        default { return 1 }
    }
}

function ConvertFrom-OutlookImportance {
    param(
        [Parameter(Mandatory = $true)] [int] $Importance
    )

    switch ($Importance) {
        0 { return 'Low' }
        2 { return 'High' }
        default { return 'Normal' }
    }
}

function ConvertTo-OutlookCategoryString {
    param(
        [AllowNull()] [string[]] $Categories
    )

    if ($null -eq $Categories) {
        return $null
    }

    $clean = @(
        $Categories |
            ForEach-Object { if ($null -eq $_) { '' } else { $_.Trim() } } |
            Where-Object { $_.Length -gt 0 }
    )

    if ($clean.Count -eq 0) {
        return $null
    }

    return ($clean -join ', ')
}

function ConvertFrom-OutlookTaskItem {
    param(
        [Parameter(Mandatory = $true)] [object] $Task
    )

    $dueDate = $Task.DueDate
    $normalizedDueDate = $null
    if ($null -ne $dueDate -and $dueDate -is [datetime] -and $dueDate.Year -lt 4500) {
        $normalizedDueDate = $dueDate.ToString('yyyy-MM-dd')
    }

    $storeId = $null
    try {
        if ($null -ne $Task.Parent) {
            $storeId = $Task.Parent.StoreID
        }
    }
    catch {
        $storeId = $null
    }

    [pscustomobject]@{
        EntryId = $Task.EntryID
        StoreId = $storeId
        Subject = $Task.Subject
        DueDate = $normalizedDueDate
        Complete = [bool] $Task.Complete
        Importance = ConvertFrom-OutlookImportance -Importance ([int] $Task.Importance)
        Categories = $Task.Categories
        Body = $Task.Body
    }
}

function New-OutlookTask {
    param(
        [Parameter(Mandatory = $true)] [object] $Application,
        [Parameter(Mandatory = $true)] [string] $Subject,
        [AllowNull()] [object] $DueDate = $null,
        [AllowNull()] [string] $Body = $null,
        [ValidateSet('Low', 'Normal', 'Medium', 'High')] [string] $Importance = 'Normal',
        [AllowNull()] [string[]] $Categories = $null
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) {
        throw 'Subject is required when adding an Outlook task.'
    }

    $task = $Application.CreateItem(3)
    $task.Subject = $Subject.Trim()

    if ($null -ne $DueDate) {
        $task.DueDate = ([datetime] $DueDate).Date
    }

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $task.Body = $Body
    }

    $task.Importance = ConvertTo-OutlookImportance -Importance $Importance

    $categoryString = ConvertTo-OutlookCategoryString -Categories $Categories
    if ($null -ne $categoryString) {
        $task.Categories = $categoryString
    }

    $task.Save()
    return ConvertFrom-OutlookTaskItem -Task $task
}

function Get-OutlookTasks {
    param(
        [Parameter(Mandatory = $true)] [object] $Namespace,
        [switch] $IncludeCompleted
    )

    $folder = $Namespace.GetDefaultFolder(13)
    $items = $folder.Items
    $items.Sort('[DueDate]')

    if (-not $IncludeCompleted) {
        $items = $items.Restrict('[Complete] = false')
    }

    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $items) {
        $records.Add((ConvertFrom-OutlookTaskItem -Task $item))
    }

    return $records
}

function Get-OutlookTaskById {
    param(
        [Parameter(Mandatory = $true)] [object] $Namespace,
        [Parameter(Mandatory = $true)] [string] $EntryId,
        [AllowNull()] [string] $StoreId = $null
    )

    if ([string]::IsNullOrWhiteSpace($EntryId)) {
        throw 'EntryId is required for this Outlook task command.'
    }

    if ([string]::IsNullOrWhiteSpace($StoreId)) {
        return $Namespace.GetItemFromID($EntryId)
    }

    return $Namespace.GetItemFromID($EntryId, $StoreId)
}

function Set-OutlookTaskComplete {
    param(
        [Parameter(Mandatory = $true)] [object] $Task,
        [Parameter(Mandatory = $true)] [bool] $Complete
    )

    $Task.Complete = $Complete
    $Task.Save()
    return ConvertFrom-OutlookTaskItem -Task $Task
}

function Rename-OutlookTask {
    param(
        [Parameter(Mandatory = $true)] [object] $Task,
        [Parameter(Mandatory = $true)] [string] $Subject
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) {
        throw 'Subject is required when renaming an Outlook task.'
    }

    $Task.Subject = $Subject.Trim()
    $Task.Save()
    return ConvertFrom-OutlookTaskItem -Task $Task
}

function Remove-OutlookTask {
    param(
        [Parameter(Mandatory = $true)] [object] $Task,
        [Parameter(Mandatory = $true)] [string] $EntryId,
        [AllowNull()] [string] $StoreId = $null
    )

    $Task.Delete()

    [pscustomobject]@{
        EntryId = $EntryId
        StoreId = $StoreId
        Deleted = $true
    }
}

function Write-QuickTodoOutput {
    param(
        [AllowNull()] [object] $Value,
        [switch] $AsJson
    )

    if ($AsJson) {
        $Value | ConvertTo-Json -Depth 6
        return
    }

    $Value
}

function Invoke-QuickTodoOutlookTaskCommand {
    [CmdletBinding()]
    param(
        [ValidateSet('add', 'list', 'complete', 'incomplete', 'rename', 'delete')]
        [string] $Command = 'list',
        [string] $Subject,
        [string] $Body,
        [datetime] $DueDate,
        [ValidateSet('Low', 'Normal', 'Medium', 'High')]
        [string] $Importance = 'Normal',
        [string[]] $Categories,
        [string] $EntryId,
        [string] $StoreId,
        [switch] $IncludeCompleted,
        [switch] $AsJson
    )

    $application = Get-OutlookApplication
    $namespace = Get-OutlookNamespace -Application $application

    switch ($Command) {
        'add' {
            $addParams = @{
                Application = $application
                Subject = $Subject
                Importance = $Importance
            }

            if ($PSBoundParameters.ContainsKey('Body')) {
                $addParams.Body = $Body
            }
            if ($PSBoundParameters.ContainsKey('DueDate')) {
                $addParams.DueDate = $DueDate
            }
            if ($PSBoundParameters.ContainsKey('Categories')) {
                $addParams.Categories = $Categories
            }

            Write-QuickTodoOutput -Value (New-OutlookTask @addParams) -AsJson:$AsJson
        }
        'list' {
            Write-QuickTodoOutput -Value (Get-OutlookTasks -Namespace $namespace -IncludeCompleted:$IncludeCompleted) -AsJson:$AsJson
        }
        'complete' {
            $task = Get-OutlookTaskById -Namespace $namespace -EntryId $EntryId -StoreId $StoreId
            Write-QuickTodoOutput -Value (Set-OutlookTaskComplete -Task $task -Complete $true) -AsJson:$AsJson
        }
        'incomplete' {
            $task = Get-OutlookTaskById -Namespace $namespace -EntryId $EntryId -StoreId $StoreId
            Write-QuickTodoOutput -Value (Set-OutlookTaskComplete -Task $task -Complete $false) -AsJson:$AsJson
        }
        'rename' {
            $task = Get-OutlookTaskById -Namespace $namespace -EntryId $EntryId -StoreId $StoreId
            Write-QuickTodoOutput -Value (Rename-OutlookTask -Task $task -Subject $Subject) -AsJson:$AsJson
        }
        'delete' {
            $task = Get-OutlookTaskById -Namespace $namespace -EntryId $EntryId -StoreId $StoreId
            Write-QuickTodoOutput -Value (Remove-OutlookTask -Task $task -EntryId $EntryId -StoreId $StoreId) -AsJson:$AsJson
        }
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-QuickTodoOutlookTaskCommand @PSBoundParameters
}
