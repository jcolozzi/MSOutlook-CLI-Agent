# Tests/TaskOps.Tests.ps1
# Structural tests for TaskOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookTask ──────────────────────────────────────────────────────────

Describe 'Get-OutlookTask' {
    BeforeAll { $cmd = Get-Command Get-OutlookTask }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Filter parameter' { $cmd.Parameters.Keys | Should -Contain 'Filter' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -IncludeCompleted parameter' { $cmd.Parameters.Keys | Should -Contain 'IncludeCompleted' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookTaskItem ─────────────────────────────────────────────────────

Describe 'Get-OutlookTaskItem' {
    BeforeAll { $cmd = Get-Command Get-OutlookTaskItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookTask ──────────────────────────────────────────────────────────

Describe 'New-OutlookTask' {
    BeforeAll { $cmd = Get-Command New-OutlookTask }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -DueDate parameter' { $cmd.Parameters.Keys | Should -Contain 'DueDate' }
    It 'has -StartDate parameter' { $cmd.Parameters.Keys | Should -Contain 'StartDate' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -Importance parameter' { $cmd.Parameters.Keys | Should -Contain 'Importance' }
    It 'has -Categories parameter' { $cmd.Parameters.Keys | Should -Contain 'Categories' }
    It 'has -ReminderDate parameter' { $cmd.Parameters.Keys | Should -Contain 'ReminderDate' }
    It 'has -PercentComplete parameter' { $cmd.Parameters.Keys | Should -Contain 'PercentComplete' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookTask ──────────────────────────────────────────────────────────

Describe 'Set-OutlookTask' {
    BeforeAll { $cmd = Get-Command Set-OutlookTask }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -DueDate parameter' { $cmd.Parameters.Keys | Should -Contain 'DueDate' }
    It 'has -StartDate parameter' { $cmd.Parameters.Keys | Should -Contain 'StartDate' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -Status parameter' { $cmd.Parameters.Keys | Should -Contain 'Status' }
    It 'has -PercentComplete parameter' { $cmd.Parameters.Keys | Should -Contain 'PercentComplete' }
    It 'has -Importance parameter' { $cmd.Parameters.Keys | Should -Contain 'Importance' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Complete-OutlookTask ─────────────────────────────────────────────────────

Describe 'Complete-OutlookTask' {
    BeforeAll { $cmd = Get-Command Complete-OutlookTask }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookTask ──────────────────────────────────────────────────────

Describe 'Remove-OutlookTask' {
    BeforeAll { $cmd = Get-Command Remove-OutlookTask }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookTaskReminder ──────────────────────────────────────────────────

Describe 'Set-OutlookTaskReminder' {
    BeforeAll { $cmd = Get-Command Set-OutlookTaskReminder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -ReminderDate parameter' { $cmd.Parameters.Keys | Should -Contain 'ReminderDate' }
    It 'has -Clear parameter' { $cmd.Parameters.Keys | Should -Contain 'Clear' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
