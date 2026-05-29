[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('add', 'list', 'complete', 'incomplete', 'rename', 'delete', 'diag')]
    [string] $Command = 'list',

    [string] $Subject,
    [string] $Body,
    [datetime] $DueDate,

    [ValidateSet('Low', 'Normal', 'Medium', 'High')]
    [string] $Importance = 'Normal',

    [string[]] $Categories,
    [ValidateSet('None', 'Daily', 'Weekly', 'Monthly', 'Yearly')]
    [string] $Recurrence = 'None',
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

function ConvertTo-OutlookRecurrenceType {
    param(
        [ValidateSet('None', 'Daily', 'Weekly', 'Monthly', 'Yearly')]
        [string] $Recurrence = 'None'
    )

    # olRecurrenceType: Daily=0, Weekly=1, Monthly=2, Yearly=5
    switch ($Recurrence) {
        'Daily' { return 0 }
        'Weekly' { return 1 }
        'Monthly' { return 2 }
        'Yearly' { return 5 }
        default { return $null }
    }
}

function Set-OutlookTaskRecurrence {
    param(
        [Parameter(Mandatory = $true)] [object] $Task,
        [Parameter(Mandatory = $true)] [int] $RecurrenceType,
        [Parameter(Mandatory = $true)] [datetime] $StartDate
    )

    $pattern = $Task.GetRecurrencePattern()
    $pattern.RecurrenceType = $RecurrenceType
    $pattern.Interval = 1

    # Weekly (1) needs a day-of-week mask. olDaysOfWeek is a bit flag where
    # Sunday=1, Monday=2, ... Saturday=64, i.e. 2^(.NET DayOfWeek index).
    if ($RecurrenceType -eq 1) {
        $pattern.DayOfWeekMask = [int][Math]::Pow(2, [int] $StartDate.DayOfWeek)
    }

    $pattern.PatternStartDate = $StartDate.Date
    $Task.Save()
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
        [AllowNull()] [string[]] $Categories = $null,
        [ValidateSet('None', 'Daily', 'Weekly', 'Monthly', 'Yearly')] [string] $Recurrence = 'None'
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

    # Apply recurrence after the initial save. Best-effort: if Outlook rejects the
    # pattern, the task is still created as a one-off rather than failing the add.
    $recurrenceType = ConvertTo-OutlookRecurrenceType -Recurrence $Recurrence
    if ($null -ne $recurrenceType) {
        try {
            $startDate = if ($null -ne $DueDate) { [datetime] $DueDate } else { Get-Date }
            Set-OutlookTaskRecurrence -Task $task -RecurrenceType $recurrenceType -StartDate $startDate
        }
        catch {
            Write-Error "Recurrence not applied: $($_.Exception.Message)" -ErrorAction Continue
        }
    }

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

function Invoke-OutlookDiagnostics {
    # Probes each COM step independently so a failure point is pinpointed rather
    # than collapsing into one opaque error. Never throws: returns a report object.
    $steps = [System.Collections.Generic.List[object]]::new()

    function Add-DiagStep {
        param([string] $Name, [bool] $Ok, [string] $Detail)
        $steps.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    }

    $report = [ordered]@{
        Ok                  = $false
        BindMethod          = $null
        OutlookVersion      = $null
        ProfileName         = $null
        DefaultStore        = $null
        TasksFolderName     = $null
        TaskCount           = $null
        IncompleteTaskCount = $null
        PowerShellVersion   = $PSVersionTable.PSVersion.ToString()
        Is64BitProcess      = [Environment]::Is64BitProcess
        Steps               = $steps
        Error               = $null
    }

    $application = $null
    try {
        try {
            $application = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
            $report.BindMethod = 'GetActiveObject (running instance)'
        }
        catch {
            $application = New-Object -ComObject 'Outlook.Application'
            $report.BindMethod = 'New-Object (started instance)'
        }
        Add-DiagStep -Name 'Bind Outlook.Application' -Ok $true -Detail $report.BindMethod
    }
    catch {
        Add-DiagStep -Name 'Bind Outlook.Application' -Ok $false -Detail $_.Exception.Message
        $report.Error = $_.Exception.Message
        return [pscustomobject]$report
    }

    try {
        $report.OutlookVersion = $application.Version
        Add-DiagStep -Name 'Read Outlook.Version' -Ok $true -Detail $application.Version
    }
    catch {
        Add-DiagStep -Name 'Read Outlook.Version' -Ok $false -Detail $_.Exception.Message
    }

    $namespace = $null
    try {
        $namespace = $application.GetNamespace('MAPI')
        Add-DiagStep -Name 'Get MAPI namespace' -Ok $true -Detail 'ok'
    }
    catch {
        Add-DiagStep -Name 'Get MAPI namespace' -Ok $false -Detail $_.Exception.Message
        $report.Error = $_.Exception.Message
        return [pscustomobject]$report
    }

    try {
        $report.ProfileName = $namespace.CurrentProfileName
        Add-DiagStep -Name 'Read current profile' -Ok $true -Detail $namespace.CurrentProfileName
    }
    catch {
        Add-DiagStep -Name 'Read current profile' -Ok $false -Detail $_.Exception.Message
    }

    $folder = $null
    try {
        $folder = $namespace.GetDefaultFolder(13)
        $report.TasksFolderName = $folder.Name
        try { $report.DefaultStore = $folder.Store.DisplayName } catch { }
        Add-DiagStep -Name 'Get default Tasks folder (13)' -Ok $true -Detail $folder.Name
    }
    catch {
        Add-DiagStep -Name 'Get default Tasks folder (13)' -Ok $false -Detail $_.Exception.Message
        $report.Error = $_.Exception.Message
        return [pscustomobject]$report
    }

    try {
        $items = $folder.Items
        $report.TaskCount = [int] $items.Count
        $incomplete = $items.Restrict('[Complete] = false')
        $report.IncompleteTaskCount = [int] $incomplete.Count
        Add-DiagStep -Name 'Enumerate tasks' -Ok $true -Detail "$($report.TaskCount) total, $($report.IncompleteTaskCount) incomplete"
    }
    catch {
        Add-DiagStep -Name 'Enumerate tasks' -Ok $false -Detail $_.Exception.Message
    }

    $report.Ok = $true
    return [pscustomobject]$report
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
        [ValidateSet('add', 'list', 'complete', 'incomplete', 'rename', 'delete', 'diag')]
        [string] $Command = 'list',
        [string] $Subject,
        [string] $Body,
        [datetime] $DueDate,
        [ValidateSet('Low', 'Normal', 'Medium', 'High')]
        [string] $Importance = 'Normal',
        [string[]] $Categories,
        [ValidateSet('None', 'Daily', 'Weekly', 'Monthly', 'Yearly')]
        [string] $Recurrence = 'None',
        [string] $EntryId,
        [string] $StoreId,
        [switch] $IncludeCompleted,
        [switch] $AsJson
    )

    # Diagnostics binds Outlook itself and reports failures as data, so it must run
    # before the unconditional bind below (which would otherwise throw first).
    if ($Command -eq 'diag') {
        Write-QuickTodoOutput -Value (Invoke-OutlookDiagnostics) -AsJson:$AsJson
        return
    }

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
            if ($PSBoundParameters.ContainsKey('Recurrence')) {
                $addParams.Recurrence = $Recurrence
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
