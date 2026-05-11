# Tests/FolderOps.Tests.ps1
# Structural tests for FolderOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookDefaultFolder ─────────────────────────────────────────────────

Describe 'Get-OutlookDefaultFolder' {
    BeforeAll { $cmd = Get-Command Get-OutlookDefaultFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'FolderType has ValidateSet' {
        $attr = $cmd.Parameters['FolderType'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'inbox'
        $attr.ValidValues | Should -Contain 'calendar'
        $attr.ValidValues | Should -Contain 'contacts'
    }
}

# ── Get-OutlookFolder ────────────────────────────────────────────────────────

Describe 'Get-OutlookFolder' {
    BeforeAll { $cmd = Get-Command Get-OutlookFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookFolderList ────────────────────────────────────────────────────

Describe 'Get-OutlookFolderList' {
    BeforeAll { $cmd = Get-Command Get-OutlookFolderList }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookFolderInfo ────────────────────────────────────────────────────

Describe 'Get-OutlookFolderInfo' {
    BeforeAll { $cmd = Get-Command Get-OutlookFolderInfo }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookFolder ────────────────────────────────────────────────────────

Describe 'New-OutlookFolder' {
    BeforeAll { $cmd = Get-Command New-OutlookFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -ParentFolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'ParentFolderType' }
    It 'has -ParentFolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'ParentFolderPath' }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Rename-OutlookFolder ─────────────────────────────────────────────────────

Describe 'Rename-OutlookFolder' {
    BeforeAll { $cmd = Get-Command Rename-OutlookFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -NewName parameter' { $cmd.Parameters.Keys | Should -Contain 'NewName' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookFolder ─────────────────────────────────────────────────────

Describe 'Remove-OutlookFolder' {
    BeforeAll { $cmd = Get-Command Remove-OutlookFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Move-OutlookFolder ───────────────────────────────────────────────────────

Describe 'Move-OutlookFolder' {
    BeforeAll { $cmd = Get-Command Move-OutlookFolder }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -DestinationPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
