# Tests/CalendarOps.Tests.ps1
# Structural tests for CalendarOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookAppointment ───────────────────────────────────────────────────

Describe 'Get-OutlookAppointment' {
    BeforeAll { $cmd = Get-Command Get-OutlookAppointment }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -End parameter' { $cmd.Parameters.Keys | Should -Contain 'End' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookAppointmentItem ───────────────────────────────────────────────

Describe 'Get-OutlookAppointmentItem' {
    BeforeAll { $cmd = Get-Command Get-OutlookAppointmentItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookAppointment ───────────────────────────────────────────────────

Describe 'New-OutlookAppointment' {
    BeforeAll { $cmd = Get-Command New-OutlookAppointment }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -End parameter' { $cmd.Parameters.Keys | Should -Contain 'End' }
    It 'has -Location parameter' { $cmd.Parameters.Keys | Should -Contain 'Location' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -BusyStatus parameter' { $cmd.Parameters.Keys | Should -Contain 'BusyStatus' }
    It 'has -Sensitivity parameter' { $cmd.Parameters.Keys | Should -Contain 'Sensitivity' }
    It 'has -ReminderMinutes parameter' { $cmd.Parameters.Keys | Should -Contain 'ReminderMinutes' }
    It 'has -AllDayEvent parameter' { $cmd.Parameters.Keys | Should -Contain 'AllDayEvent' }
    It 'has -Categories parameter' { $cmd.Parameters.Keys | Should -Contain 'Categories' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── New-OutlookMeeting ───────────────────────────────────────────────────────

Describe 'New-OutlookMeeting' {
    BeforeAll { $cmd = Get-Command New-OutlookMeeting }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -End parameter' { $cmd.Parameters.Keys | Should -Contain 'End' }
    It 'has -Location parameter' { $cmd.Parameters.Keys | Should -Contain 'Location' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -RequiredAttendees parameter' { $cmd.Parameters.Keys | Should -Contain 'RequiredAttendees' }
    It 'has -OptionalAttendees parameter' { $cmd.Parameters.Keys | Should -Contain 'OptionalAttendees' }
    It 'has -BusyStatus parameter' { $cmd.Parameters.Keys | Should -Contain 'BusyStatus' }
    It 'has -ReminderMinutes parameter' { $cmd.Parameters.Keys | Should -Contain 'ReminderMinutes' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookAppointment ───────────────────────────────────────────────────

Describe 'Set-OutlookAppointment' {
    BeforeAll { $cmd = Get-Command Set-OutlookAppointment }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -End parameter' { $cmd.Parameters.Keys | Should -Contain 'End' }
    It 'has -Location parameter' { $cmd.Parameters.Keys | Should -Contain 'Location' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -BusyStatus parameter' { $cmd.Parameters.Keys | Should -Contain 'BusyStatus' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookAppointment ────────────────────────────────────────────────

Describe 'Remove-OutlookAppointment' {
    BeforeAll { $cmd = Get-Command Remove-OutlookAppointment }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Send-OutlookMeetingResponse ──────────────────────────────────────────────

Describe 'Send-OutlookMeetingResponse' {
    BeforeAll { $cmd = Get-Command Send-OutlookMeetingResponse }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Response parameter' { $cmd.Parameters.Keys | Should -Contain 'Response' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'Response has ValidateSet' {
        $attr = $cmd.Parameters['Response'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'accept'
        $attr.ValidValues | Should -Contain 'decline'
        $attr.ValidValues | Should -Contain 'tentative'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Get-OutlookFreeBusy ──────────────────────────────────────────────────────

Describe 'Get-OutlookFreeBusy' {
    BeforeAll { $cmd = Get-Command Get-OutlookFreeBusy }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -RecipientName parameter' { $cmd.Parameters.Keys | Should -Contain 'RecipientName' }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -Duration parameter' { $cmd.Parameters.Keys | Should -Contain 'Duration' }
    It 'has -Days parameter' { $cmd.Parameters.Keys | Should -Contain 'Days' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookRecurrence ────────────────────────────────────────────────────

Describe 'Get-OutlookRecurrence' {
    BeforeAll { $cmd = Get-Command Get-OutlookRecurrence }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Set-OutlookRecurrence ────────────────────────────────────────────────────

Describe 'Set-OutlookRecurrence' {
    BeforeAll { $cmd = Get-Command Set-OutlookRecurrence }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -RecurrenceType parameter' { $cmd.Parameters.Keys | Should -Contain 'RecurrenceType' }
    It 'has -Interval parameter' { $cmd.Parameters.Keys | Should -Contain 'Interval' }
    It 'has -DayOfWeekMask parameter' { $cmd.Parameters.Keys | Should -Contain 'DayOfWeekMask' }
    It 'has -PatternEndDate parameter' { $cmd.Parameters.Keys | Should -Contain 'PatternEndDate' }
    It 'has -Occurrences parameter' { $cmd.Parameters.Keys | Should -Contain 'Occurrences' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
