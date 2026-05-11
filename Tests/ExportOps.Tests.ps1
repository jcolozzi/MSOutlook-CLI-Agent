# Tests/ExportOps.Tests.ps1
# Structural tests for ExportOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Export-OutlookItem ───────────────────────────────────────────────────────

Describe 'Export-OutlookItem' {
    BeforeAll { $cmd = Get-Command Export-OutlookItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -DestinationPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationPath' }
    It 'has -Format parameter' { $cmd.Parameters.Keys | Should -Contain 'Format' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'Format has ValidateSet' {
        $attr = $cmd.Parameters['Format'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'msg'
        $attr.ValidValues | Should -Contain 'html'
        $attr.ValidValues | Should -Contain 'txt'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Export-OutlookCalendar ───────────────────────────────────────────────────

Describe 'Export-OutlookCalendar' {
    BeforeAll { $cmd = Get-Command Export-OutlookCalendar }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Start parameter' { $cmd.Parameters.Keys | Should -Contain 'Start' }
    It 'has -End parameter' { $cmd.Parameters.Keys | Should -Contain 'End' }
    It 'has -DestinationPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Export-OutlookContacts ───────────────────────────────────────────────────

Describe 'Export-OutlookContacts' {
    BeforeAll { $cmd = Get-Command Export-OutlookContacts }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -DestinationPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationPath' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Export-OutlookFolderItems ────────────────────────────────────────────────

Describe 'Export-OutlookFolderItems' {
    BeforeAll { $cmd = Get-Command Export-OutlookFolderItems }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -DestinationPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationPath' }
    It 'has -Format parameter' { $cmd.Parameters.Keys | Should -Contain 'Format' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'Format has ValidateSet' {
        $attr = $cmd.Parameters['Format'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'msg'
        $attr.ValidValues | Should -Contain 'html'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
