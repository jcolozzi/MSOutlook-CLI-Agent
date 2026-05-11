# ── Calendar Operations ──────────────────────────────────────────────────────

function Get-OutlookAppointment {
    <#
    .SYNOPSIS
        Lists appointments/meetings in a date range from the calendar.
    .PARAMETER Start
        Start of the date range.
    .PARAMETER End
        End of the date range.
    .PARAMETER MaxItems
        Maximum items to return (default 50).
    .PARAMETER FolderPath
        Full path to a calendar folder. Defaults to the default Calendar.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookAppointment -Start '2026-05-01' -End '2026-05-31'
    .EXAMPLE
        Get-OutlookAppointment -Start (Get-Date) -End (Get-Date).AddDays(7) -AsJson
    #>
    [CmdletBinding()]
    param(
        [datetime]$Start,
        [datetime]$End,
        [int]$MaxItems = 50,
        [string]$FolderPath,
        [switch]$AsJson
    )

    if ($Start -eq [datetime]::MinValue) { throw 'Get-OutlookAppointment: -Start is required.' }
    if ($End   -eq [datetime]::MinValue) { throw 'Get-OutlookAppointment: -End is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = $ns.GetDefaultFolder(9)  # olFolderCalendar
    }

    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort('[Start]')

    $startStr = $Start.ToString('MM/dd/yyyy HH:mm')
    $endStr   = $End.ToString('MM/dd/yyyy HH:mm')
    $filter   = "[Start] >= '$startStr' AND [End] <= '$endStr'"
    $filtered = $items.Restrict($filter)

    $list  = @()
    $count = 0
    foreach ($item in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += @{
                entryID       = ConvertTo-OutlookSafeValue $item.EntryID
                subject       = ConvertTo-OutlookSafeValue $item.Subject
                start         = ConvertTo-OutlookSafeValue $item.Start
                end           = ConvertTo-OutlookSafeValue $item.End
                duration      = ConvertTo-OutlookSafeValue $item.Duration
                location      = ConvertTo-OutlookSafeValue $item.Location
                organizer     = ConvertTo-OutlookSafeValue $item.Organizer
                busyStatus    = ConvertTo-OutlookSafeValue $item.BusyStatus
                isRecurring   = ConvertTo-OutlookSafeValue $item.IsRecurring
                allDayEvent   = ConvertTo-OutlookSafeValue $item.AllDayEvent
                meetingStatus = ConvertTo-OutlookSafeValue $item.MeetingStatus
            }
            $count++
        } catch {
            Write-Verbose "Skipped calendar item: $_"
        }
    }

    if ($AsJson) {
        return ($list | ForEach-Object { [PSCustomObject]$_ }) | ConvertTo-Json -Depth 10 -Compress
    }
    return $list | ForEach-Object { [PSCustomObject]$_ }
}

function Get-OutlookAppointmentItem {
    <#
    .SYNOPSIS
        Returns full details for a single appointment/meeting by EntryID.
    .PARAMETER EntryID
        The EntryID of the appointment.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookAppointmentItem -EntryID '000000...'
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookAppointmentItem: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $body = ConvertTo-OutlookSafeValue $item.Body
    if ($body -is [string] -and $body.Length -gt 5000) { $body = $body.Substring(0, 5000) }

    $result = @{
        entryID               = ConvertTo-OutlookSafeValue $item.EntryID
        subject               = ConvertTo-OutlookSafeValue $item.Subject
        body                  = $body
        start                 = ConvertTo-OutlookSafeValue $item.Start
        end                   = ConvertTo-OutlookSafeValue $item.End
        duration              = ConvertTo-OutlookSafeValue $item.Duration
        location              = ConvertTo-OutlookSafeValue $item.Location
        organizer             = ConvertTo-OutlookSafeValue $item.Organizer
        requiredAttendees     = ConvertTo-OutlookSafeValue $item.RequiredAttendees
        optionalAttendees     = ConvertTo-OutlookSafeValue $item.OptionalAttendees
        resources             = ConvertTo-OutlookSafeValue $item.Resources
        busyStatus            = ConvertTo-OutlookSafeValue $item.BusyStatus
        meetingStatus         = ConvertTo-OutlookSafeValue $item.MeetingStatus
        responseStatus        = ConvertTo-OutlookSafeValue $item.ResponseStatus
        isRecurring           = ConvertTo-OutlookSafeValue $item.IsRecurring
        allDayEvent           = ConvertTo-OutlookSafeValue $item.AllDayEvent
        categories            = ConvertTo-OutlookSafeValue $item.Categories
        importance            = ConvertTo-OutlookSafeValue $item.Importance
        sensitivity           = ConvertTo-OutlookSafeValue $item.Sensitivity
        reminderSet           = ConvertTo-OutlookSafeValue $item.ReminderSet
        reminderMinutesBefore = ConvertTo-OutlookSafeValue $item.ReminderMinutesBeforeStart
        isOnlineMeeting       = ConvertTo-OutlookSafeValue $item.IsOnlineMeeting
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookAppointment {
    <#
    .SYNOPSIS
        Creates a calendar appointment (no attendees).
    .PARAMETER Subject
        Appointment subject.
    .PARAMETER Start
        Start date/time.
    .PARAMETER End
        End date/time.
    .PARAMETER Location
        Location string.
    .PARAMETER Body
        Body text.
    .PARAMETER BusyStatus
        Busy status: free, tentative, busy, outOfOffice, workingElsewhere.
    .PARAMETER Sensitivity
        Sensitivity: normal, personal, private, confidential.
    .PARAMETER ReminderMinutes
        Reminder minutes before start.
    .PARAMETER AllDayEvent
        Mark as all-day event.
    .PARAMETER Categories
        Categories (semicolon-delimited string).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookAppointment -Subject 'Lunch' -Start '2026-05-12 12:00' -End '2026-05-12 13:00'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Subject,
        [datetime]$Start,
        [datetime]$End,
        [string]$Location,
        [string]$Body,
        [string]$BusyStatus,
        [string]$Sensitivity,
        [int]$ReminderMinutes = -1,
        [switch]$AllDayEvent,
        [string]$Categories,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject))    { throw 'New-OutlookAppointment: -Subject is required.' }
    if ($Start -eq [datetime]::MinValue)           { throw 'New-OutlookAppointment: -Start is required.' }
    if ($End   -eq [datetime]::MinValue)           { throw 'New-OutlookAppointment: -End is required.' }

    $app = Connect-OutlookSession

    if ($PSCmdlet.ShouldProcess($Subject, 'Create appointment')) {
        $appt = $app.CreateItem(1)  # olAppointmentItem
        $appt.Subject = $Subject
        $appt.Start   = $Start
        $appt.End     = $End

        if (-not [string]::IsNullOrWhiteSpace($Location)) { $appt.Location = $Location }
        if (-not [string]::IsNullOrWhiteSpace($Body))     { $appt.Body     = $Body }

        if (-not [string]::IsNullOrWhiteSpace($BusyStatus)) {
            $appt.BusyStatus = Resolve-EnumValue -Map $script:OL_BUSY_STATUS -Key $BusyStatus -EnumName 'BusyStatus'
        }
        if (-not [string]::IsNullOrWhiteSpace($Sensitivity)) {
            $appt.Sensitivity = Resolve-EnumValue -Map $script:OL_SENSITIVITY -Key $Sensitivity -EnumName 'Sensitivity'
        }
        if ($ReminderMinutes -ge 0) {
            $appt.ReminderSet = $true
            $appt.ReminderMinutesBeforeStart = $ReminderMinutes
        }
        if ($AllDayEvent) { $appt.AllDayEvent = $true }
        if (-not [string]::IsNullOrWhiteSpace($Categories)) { $appt.Categories = $Categories }

        $appt.Save()

        $result = @{
            status  = 'created'
            entryID = ConvertTo-OutlookSafeValue $appt.EntryID
            subject = ConvertTo-OutlookSafeValue $appt.Subject
            start   = ConvertTo-OutlookSafeValue $appt.Start
            end     = ConvertTo-OutlookSafeValue $appt.End
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function New-OutlookMeeting {
    <#
    .SYNOPSIS
        Creates a meeting with attendees and sends invitations.
    .PARAMETER Subject
        Meeting subject.
    .PARAMETER Start
        Start date/time.
    .PARAMETER End
        End date/time.
    .PARAMETER Location
        Location string.
    .PARAMETER Body
        Body text.
    .PARAMETER RequiredAttendees
        Required attendees, semicolon-delimited email addresses.
    .PARAMETER OptionalAttendees
        Optional attendees, semicolon-delimited email addresses.
    .PARAMETER BusyStatus
        Busy status: free, tentative, busy, outOfOffice, workingElsewhere.
    .PARAMETER ReminderMinutes
        Reminder minutes before start.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookMeeting -Subject 'Sprint Review' -Start '2026-05-13 10:00' -End '2026-05-13 11:00' -RequiredAttendees 'alice@example.com;bob@example.com'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Subject,
        [datetime]$Start,
        [datetime]$End,
        [string]$Location,
        [string]$Body,
        [string]$RequiredAttendees,
        [string]$OptionalAttendees,
        [string]$BusyStatus,
        [int]$ReminderMinutes = -1,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject))           { throw 'New-OutlookMeeting: -Subject is required.' }
    if ($Start -eq [datetime]::MinValue)                  { throw 'New-OutlookMeeting: -Start is required.' }
    if ($End   -eq [datetime]::MinValue)                  { throw 'New-OutlookMeeting: -End is required.' }
    if ([string]::IsNullOrWhiteSpace($RequiredAttendees)) { throw 'New-OutlookMeeting: -RequiredAttendees is required.' }

    $app = Connect-OutlookSession

    if ($PSCmdlet.ShouldProcess($Subject, 'Create meeting and send invitations')) {
        $meeting = $app.CreateItem(1)  # olAppointmentItem
        $meeting.Subject       = $Subject
        $meeting.Start         = $Start
        $meeting.End           = $End
        $meeting.MeetingStatus = 1     # olMeeting

        if (-not [string]::IsNullOrWhiteSpace($Location)) { $meeting.Location = $Location }
        if (-not [string]::IsNullOrWhiteSpace($Body))     { $meeting.Body     = $Body }

        if (-not [string]::IsNullOrWhiteSpace($BusyStatus)) {
            $meeting.BusyStatus = Resolve-EnumValue -Map $script:OL_BUSY_STATUS -Key $BusyStatus -EnumName 'BusyStatus'
        }
        if ($ReminderMinutes -ge 0) {
            $meeting.ReminderSet = $true
            $meeting.ReminderMinutesBeforeStart = $ReminderMinutes
        }

        # Add required attendees
        foreach ($addr in ($RequiredAttendees -split ';')) {
            $addr = $addr.Trim()
            if ($addr -ne '') {
                $recip = $meeting.Recipients.Add($addr)
                $recip.Type = 1  # olRequired
            }
        }

        # Add optional attendees
        if (-not [string]::IsNullOrWhiteSpace($OptionalAttendees)) {
            foreach ($addr in ($OptionalAttendees -split ';')) {
                $addr = $addr.Trim()
                if ($addr -ne '') {
                    $recip = $meeting.Recipients.Add($addr)
                    $recip.Type = 2  # olOptional
                }
            }
        }

        $meeting.Recipients.ResolveAll() | Out-Null
        $meeting.Send()

        $result = @{
            status  = 'meeting_sent'
            subject = ConvertTo-OutlookSafeValue $Subject
            start   = ConvertTo-OutlookSafeValue $Start
            end     = ConvertTo-OutlookSafeValue $End
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Set-OutlookAppointment {
    <#
    .SYNOPSIS
        Updates fields on an existing appointment.
    .PARAMETER EntryID
        The EntryID of the appointment to update.
    .PARAMETER Subject
        New subject.
    .PARAMETER Start
        New start date/time.
    .PARAMETER End
        New end date/time.
    .PARAMETER Location
        New location.
    .PARAMETER Body
        New body text.
    .PARAMETER BusyStatus
        New busy status.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookAppointment -EntryID '000000...' -Location 'Room 201'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$Subject,
        [datetime]$Start,
        [datetime]$End,
        [string]$Location,
        [string]$Body,
        [string]$BusyStatus,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Set-OutlookAppointment: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($PSCmdlet.ShouldProcess($item.Subject, 'Update appointment')) {
        if (-not [string]::IsNullOrWhiteSpace($Subject))  { $item.Subject  = $Subject }
        if ($Start -ne [datetime]::MinValue)              { $item.Start    = $Start }
        if ($End   -ne [datetime]::MinValue)              { $item.End      = $End }
        if (-not [string]::IsNullOrWhiteSpace($Location)) { $item.Location = $Location }
        if (-not [string]::IsNullOrWhiteSpace($Body))     { $item.Body     = $Body }

        if (-not [string]::IsNullOrWhiteSpace($BusyStatus)) {
            $item.BusyStatus = Resolve-EnumValue -Map $script:OL_BUSY_STATUS -Key $BusyStatus -EnumName 'BusyStatus'
        }

        $item.Save()

        $result = @{
            status  = 'updated'
            entryID = ConvertTo-OutlookSafeValue $item.EntryID
            subject = ConvertTo-OutlookSafeValue $item.Subject
            start   = ConvertTo-OutlookSafeValue $item.Start
            end     = ConvertTo-OutlookSafeValue $item.End
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Remove-OutlookAppointment {
    <#
    .SYNOPSIS
        Deletes an appointment or meeting.
    .PARAMETER EntryID
        The EntryID of the appointment to delete.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Remove-OutlookAppointment -EntryID '000000...'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Remove-OutlookAppointment: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)
    $subj = $item.Subject

    if ($PSCmdlet.ShouldProcess($subj, 'Delete appointment')) {
        $item.Delete()

        $result = @{ status = 'deleted'; subject = $subj }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Send-OutlookMeetingResponse {
    <#
    .SYNOPSIS
        Responds to a meeting invitation (accept, decline, or tentative).
    .PARAMETER EntryID
        The EntryID of the meeting item.
    .PARAMETER Response
        Response type: accept, decline, tentative.
    .PARAMETER SendResponse
        Whether to send the response to the organizer (default $true).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Send-OutlookMeetingResponse -EntryID '000000...' -Response 'accept'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,

        [ValidateSet('accept','decline','tentative')]
        [string]$Response,

        [bool]$SendResponse = $true,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))  { throw 'Send-OutlookMeetingResponse: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($Response)) { throw 'Send-OutlookMeetingResponse: -Response is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $responseMap = @{ accept = 3; decline = 4; tentative = 2 }
    $responseVal = $responseMap[$Response]

    if ($PSCmdlet.ShouldProcess($item.Subject, "Respond '$Response'")) {
        $responseItem = $item.Respond($responseVal, $null, $SendResponse)

        # If SendResponse is $false, Respond returns a MeetingItem that needs saving
        if ($null -ne $responseItem -and -not $SendResponse) {
            $responseItem.Save()
        }

        $result = @{
            status   = 'responded'
            response = $Response
            subject  = ConvertTo-OutlookSafeValue $item.Subject
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Get-OutlookFreeBusy {
    <#
    .SYNOPSIS
        Returns the free/busy string for a recipient.
    .PARAMETER RecipientName
        Display name or email address of the recipient.
    .PARAMETER Start
        Start date for the free/busy query.
    .PARAMETER Duration
        Interval in minutes for each free/busy slot (default 60).
    .PARAMETER Days
        Number of days to query (default 30).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookFreeBusy -RecipientName 'alice@example.com' -Start (Get-Date)
    #>
    [CmdletBinding()]
    param(
        [string]$RecipientName,
        [datetime]$Start,
        [int]$Duration = 60,
        [int]$Days = 30,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($RecipientName)) { throw 'Get-OutlookFreeBusy: -RecipientName is required.' }
    if ($Start -eq [datetime]::MinValue)              { throw 'Get-OutlookFreeBusy: -Start is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $recip = $ns.CreateRecipient($RecipientName)
    $recip.Resolve() | Out-Null

    if (-not $recip.Resolved) {
        throw "Get-OutlookFreeBusy: could not resolve recipient '$RecipientName'."
    }

    $fbString = $recip.FreeBusy($Start, $Duration, $true)

    # Trim to requested days
    $slotsPerDay = [math]::Ceiling(1440 / $Duration)
    $totalSlots  = $slotsPerDay * $Days
    if ($fbString.Length -gt $totalSlots) {
        $fbString = $fbString.Substring(0, $totalSlots)
    }

    $result = @{
        recipientName = $RecipientName
        start         = ConvertTo-OutlookSafeValue $Start
        duration      = $Duration
        days          = $Days
        freeBusy      = $fbString
        legend        = '0=free, 1=tentative, 2=busy, 3=outOfOffice, 4=workingElsewhere'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookRecurrence {
    <#
    .SYNOPSIS
        Returns the recurrence pattern of an appointment.
    .PARAMETER EntryID
        The EntryID of the recurring appointment.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookRecurrence -EntryID '000000...'
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookRecurrence: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if (-not $item.IsRecurring) {
        throw 'Get-OutlookRecurrence: item is not recurring.'
    }

    $pattern = $item.GetRecurrencePattern()

    $result = @{
        recurrenceType = ConvertTo-OutlookSafeValue $pattern.RecurrenceType
        interval       = ConvertTo-OutlookSafeValue $pattern.Interval
        dayOfWeekMask  = ConvertTo-OutlookSafeValue $pattern.DayOfWeekMask
        dayOfMonth     = ConvertTo-OutlookSafeValue $pattern.DayOfMonth
        monthOfYear    = ConvertTo-OutlookSafeValue $pattern.MonthOfYear
        patternStart   = ConvertTo-OutlookSafeValue $pattern.PatternStartDate
        patternEnd     = ConvertTo-OutlookSafeValue $pattern.PatternEndDate
        noEndDate      = ConvertTo-OutlookSafeValue $pattern.NoEndDate
        occurrences    = ConvertTo-OutlookSafeValue $pattern.Occurrences
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Set-OutlookRecurrence {
    <#
    .SYNOPSIS
        Sets or updates the recurrence pattern on an appointment.
    .PARAMETER EntryID
        The EntryID of the appointment.
    .PARAMETER RecurrenceType
        Type: daily, weekly, monthly, monthNth, yearly, yearNth.
    .PARAMETER Interval
        Recurrence interval (e.g. every 2 weeks).
    .PARAMETER DayOfWeekMask
        Bitmask for days of the week (sum of: Sun=1, Mon=2, Tue=4, Wed=8, Thu=16, Fri=32, Sat=64).
    .PARAMETER PatternEndDate
        End date for the recurrence pattern.
    .PARAMETER Occurrences
        Number of occurrences (alternative to PatternEndDate).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookRecurrence -EntryID '000000...' -RecurrenceType 'weekly' -Interval 1 -DayOfWeekMask 42
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$RecurrenceType,
        [int]$Interval = -1,
        [int]$DayOfWeekMask = -1,
        [datetime]$PatternEndDate,
        [int]$Occurrences = -1,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))        { throw 'Set-OutlookRecurrence: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($RecurrenceType)) { throw 'Set-OutlookRecurrence: -RecurrenceType is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($PSCmdlet.ShouldProcess($item.Subject, "Set recurrence '$RecurrenceType'")) {
        $pattern = $item.GetRecurrencePattern()
        $pattern.RecurrenceType = Resolve-EnumValue -Map $script:OL_RECURRENCE_TYPE -Key $RecurrenceType -EnumName 'RecurrenceType'

        if ($Interval -ge 1)       { $pattern.Interval      = $Interval }
        if ($DayOfWeekMask -ge 0)  { $pattern.DayOfWeekMask = $DayOfWeekMask }

        if ($PatternEndDate -ne [datetime]::MinValue) {
            $pattern.PatternEndDate = $PatternEndDate
            $pattern.NoEndDate      = $false
        } elseif ($Occurrences -ge 1) {
            $pattern.Occurrences = $Occurrences
            $pattern.NoEndDate   = $false
        }

        $item.Save()

        $result = @{
            status         = 'recurrence_set'
            entryID        = ConvertTo-OutlookSafeValue $item.EntryID
            recurrenceType = $RecurrenceType
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}
