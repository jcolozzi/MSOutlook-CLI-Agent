# ── TaskOps.ps1 ─────────────────────────────────────────────────────────────
# Task operations for OutlookPOSH
# ────────────────────────────────────────────────────────────────────────────

function Get-OutlookTask {
<#
.SYNOPSIS
    Lists tasks from the default Tasks folder.
.PARAMETER Filter
    Optional DASL filter string to restrict results.
.PARAMETER MaxItems
    Maximum number of tasks to return. Default 50.
.PARAMETER IncludeCompleted
    Include completed tasks in results. By default only incomplete tasks are returned.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookTask -MaxItems 10
.EXAMPLE
    Get-OutlookTask -IncludeCompleted
#>
    [CmdletBinding()]
    param(
        [string]$Filter,
        [int]$MaxItems = 50,
        [switch]$IncludeCompleted,
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    # olFolderTasks = 13
    $folder = $ns.GetDefaultFolder(13)
    $items  = $folder.Items

    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $items = $items.Restrict($Filter)
    } elseif (-not $IncludeCompleted) {
        $items = $items.Restrict("[Complete] = False")
    }

    $results = [System.Collections.Generic.List[hashtable]]::new()
    $count   = 0

    foreach ($t in $items) {
        if ($count -ge $MaxItems) { break }
        try {
            $results.Add(@{
                entryID         = ConvertTo-OutlookSafeValue $t.EntryID
                subject         = ConvertTo-OutlookSafeValue $t.Subject
                status          = ConvertTo-OutlookSafeValue $t.Status
                percentComplete = ConvertTo-OutlookSafeValue $t.PercentComplete
                startDate       = ConvertTo-OutlookSafeValue $t.StartDate
                dueDate         = ConvertTo-OutlookSafeValue $t.DueDate
                importance      = ConvertTo-OutlookSafeValue $t.Importance
                categories      = ConvertTo-OutlookSafeValue $t.Categories
                owner           = ConvertTo-OutlookSafeValue $t.Owner
                isRecurring     = ConvertTo-OutlookSafeValue $t.IsRecurring
            })
            $count++
        } catch { }
    }

    Format-OutlookOutput -Data $results -AsJson:$AsJson
}

function Get-OutlookTaskItem {
<#
.SYNOPSIS
    Returns full details for a single task by EntryID.
.PARAMETER EntryID
    The EntryID of the task item.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookTaskItem -EntryID '00000...'
#>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Get-OutlookTaskItem: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $t   = $ns.GetItemFromID($EntryID)

    $result = @{
        entryID         = ConvertTo-OutlookSafeValue $t.EntryID
        subject         = ConvertTo-OutlookSafeValue $t.Subject
        body            = ConvertTo-OutlookSafeValue $t.Body
        status          = ConvertTo-OutlookSafeValue $t.Status
        percentComplete = ConvertTo-OutlookSafeValue $t.PercentComplete
        startDate       = ConvertTo-OutlookSafeValue $t.StartDate
        dueDate         = ConvertTo-OutlookSafeValue $t.DueDate
        dateCompleted   = ConvertTo-OutlookSafeValue $t.DateCompleted
        importance      = ConvertTo-OutlookSafeValue $t.Importance
        sensitivity     = ConvertTo-OutlookSafeValue $t.Sensitivity
        categories      = ConvertTo-OutlookSafeValue $t.Categories
        owner           = ConvertTo-OutlookSafeValue $t.Owner
        isRecurring     = ConvertTo-OutlookSafeValue $t.IsRecurring
        reminderSet     = ConvertTo-OutlookSafeValue $t.ReminderSet
        reminderTime    = ConvertTo-OutlookSafeValue $t.ReminderTime
        totalWork       = ConvertTo-OutlookSafeValue $t.TotalWork
        actualWork      = ConvertTo-OutlookSafeValue $t.ActualWork
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookTask {
<#
.SYNOPSIS
    Creates a new task.
.PARAMETER Subject
    Task subject line.
.PARAMETER DueDate
    Due date for the task.
.PARAMETER StartDate
    Start date for the task.
.PARAMETER Body
    Task body / notes.
.PARAMETER Importance
    Importance level (Low, Normal, High).
.PARAMETER Categories
    Comma-separated categories string.
.PARAMETER ReminderDate
    Date/time for the reminder.
.PARAMETER PercentComplete
    Percent complete (0-100).
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    New-OutlookTask -Subject 'Review PR' -DueDate '2026-05-15'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Subject,
        [datetime]$DueDate,
        [datetime]$StartDate,
        [string]$Body,
        [string]$Importance,
        [string]$Categories,
        [datetime]$ReminderDate,
        [int]$PercentComplete = -1,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) {
        throw "New-OutlookTask: -Subject is required."
    }

    if (-not $PSCmdlet.ShouldProcess($Subject, 'Create task')) { return }

    $app = Connect-OutlookSession

    # olTaskItem = 3
    $t = $app.CreateItem(3)
    $t.Subject = $Subject

    if ($PSBoundParameters.ContainsKey('DueDate'))   { $t.DueDate   = $DueDate }
    if ($PSBoundParameters.ContainsKey('StartDate')) { $t.StartDate = $StartDate }
    if (-not [string]::IsNullOrWhiteSpace($Body))    { $t.Body      = $Body }

    if (-not [string]::IsNullOrWhiteSpace($Importance)) {
        $t.Importance = Resolve-EnumValue -Map $script:OL_IMPORTANCE -Key $Importance -EnumName 'OlImportance'
    }

    if (-not [string]::IsNullOrWhiteSpace($Categories)) { $t.Categories = $Categories }

    if ($PSBoundParameters.ContainsKey('ReminderDate')) {
        $t.ReminderSet  = $true
        $t.ReminderTime = $ReminderDate
    }

    if ($PercentComplete -ge 0 -and $PercentComplete -le 100) {
        $t.PercentComplete = $PercentComplete
    }

    $t.Save()

    $result = @{
        entryID = ConvertTo-OutlookSafeValue $t.EntryID
        subject = ConvertTo-OutlookSafeValue $t.Subject
        status  = 'Created'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Set-OutlookTask {
<#
.SYNOPSIS
    Updates an existing task. Only provided (non-null) fields are changed.
.PARAMETER EntryID
    The EntryID of the task to update.
.PARAMETER Subject
    New subject.
.PARAMETER DueDate
    New due date.
.PARAMETER StartDate
    New start date.
.PARAMETER Body
    New body text.
.PARAMETER Status
    New status (NotStarted, InProgress, Complete, WaitingOnSomeone, Deferred).
.PARAMETER PercentComplete
    New percent complete (0-100).
.PARAMETER Importance
    New importance level (Low, Normal, High).
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Set-OutlookTask -EntryID '00000...' -PercentComplete 50 -Status 'InProgress'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$Subject,
        [datetime]$DueDate,
        [datetime]$StartDate,
        [string]$Body,
        [string]$Status,
        [int]$PercentComplete = -1,
        [string]$Importance,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Set-OutlookTask: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $t   = $ns.GetItemFromID($EntryID)

    if (-not $PSCmdlet.ShouldProcess($t.Subject, 'Update task')) { return }

    if (-not [string]::IsNullOrWhiteSpace($Subject))    { $t.Subject    = $Subject }
    if ($PSBoundParameters.ContainsKey('DueDate'))       { $t.DueDate    = $DueDate }
    if ($PSBoundParameters.ContainsKey('StartDate'))     { $t.StartDate  = $StartDate }
    if (-not [string]::IsNullOrWhiteSpace($Body))        { $t.Body       = $Body }

    if (-not [string]::IsNullOrWhiteSpace($Status)) {
        $t.Status = Resolve-EnumValue -Map $script:OL_TASK_STATUS -Key $Status -EnumName 'OlTaskStatus'
    }

    if ($PercentComplete -ge 0 -and $PercentComplete -le 100) {
        $t.PercentComplete = $PercentComplete
    }

    if (-not [string]::IsNullOrWhiteSpace($Importance)) {
        $t.Importance = Resolve-EnumValue -Map $script:OL_IMPORTANCE -Key $Importance -EnumName 'OlImportance'
    }

    $t.Save()

    $result = @{
        entryID = ConvertTo-OutlookSafeValue $t.EntryID
        subject = ConvertTo-OutlookSafeValue $t.Subject
        status  = 'Updated'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Complete-OutlookTask {
<#
.SYNOPSIS
    Marks a task as complete.
.PARAMETER EntryID
    The EntryID of the task to complete.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Complete-OutlookTask -EntryID '00000...'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Complete-OutlookTask: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $t   = $ns.GetItemFromID($EntryID)

    if (-not $PSCmdlet.ShouldProcess($t.Subject, 'Complete task')) { return }

    # olTaskComplete = 2
    $t.Status          = 2
    $t.PercentComplete = 100
    $t.DateCompleted   = (Get-Date)
    $t.MarkComplete()
    $t.Save()

    $result = @{
        entryID = ConvertTo-OutlookSafeValue $t.EntryID
        subject = ConvertTo-OutlookSafeValue $t.Subject
        status  = 'Completed'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookTask {
<#
.SYNOPSIS
    Deletes a task by EntryID.
.PARAMETER EntryID
    The EntryID of the task to delete.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Remove-OutlookTask -EntryID '00000...'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Remove-OutlookTask: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $t   = $ns.GetItemFromID($EntryID)

    $subj = $t.Subject
    if (-not $PSCmdlet.ShouldProcess($subj, 'Delete task')) { return }

    $t.Delete()

    $result = @{ subject = $subj; status = 'Deleted' }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Set-OutlookTaskReminder {
<#
.SYNOPSIS
    Sets or clears a reminder on a task.
.PARAMETER EntryID
    The EntryID of the task.
.PARAMETER ReminderDate
    Date/time for the reminder.
.PARAMETER Clear
    Clear the existing reminder instead of setting one.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Set-OutlookTaskReminder -EntryID '00000...' -ReminderDate '2026-05-20 09:00'
.EXAMPLE
    Set-OutlookTaskReminder -EntryID '00000...' -Clear
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [datetime]$ReminderDate,
        [switch]$Clear,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Set-OutlookTaskReminder: -EntryID is required."
    }

    if (-not $Clear -and -not $PSBoundParameters.ContainsKey('ReminderDate')) {
        throw "Set-OutlookTaskReminder: -ReminderDate is required unless -Clear is specified."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $t   = $ns.GetItemFromID($EntryID)

    $action = if ($Clear) { 'Clear reminder' } else { "Set reminder to $ReminderDate" }
    if (-not $PSCmdlet.ShouldProcess($t.Subject, $action)) { return }

    if ($Clear) {
        $t.ReminderSet = $false
    } else {
        $t.ReminderSet  = $true
        $t.ReminderTime = $ReminderDate
    }

    $t.Save()

    $result = @{
        entryID     = ConvertTo-OutlookSafeValue $t.EntryID
        subject     = ConvertTo-OutlookSafeValue $t.Subject
        reminderSet = ConvertTo-OutlookSafeValue $t.ReminderSet
        status      = if ($Clear) { 'ReminderCleared' } else { 'ReminderSet' }
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
