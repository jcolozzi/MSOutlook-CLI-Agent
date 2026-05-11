# Tests/CategoryOps.Tests.ps1
# Structural tests for CategoryOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookCategory ─────────────────────────────────────────────────────

Describe 'Get-OutlookCategory' {
    BeforeAll { $cmd = Get-Command Get-OutlookCategory }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookCategory ─────────────────────────────────────────────────────

Describe 'New-OutlookCategory' {
    BeforeAll { $cmd = Get-Command New-OutlookCategory }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -Color parameter' { $cmd.Parameters.Keys | Should -Contain 'Color' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookCategory ──────────────────────────────────────────────────

Describe 'Remove-OutlookCategory' {
    BeforeAll { $cmd = Get-Command Remove-OutlookCategory }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookItemCategory ─────────────────────────────────────────────────

Describe 'Set-OutlookItemCategory' {
    BeforeAll { $cmd = Get-Command Set-OutlookItemCategory }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Categories parameter' { $cmd.Parameters.Keys | Should -Contain 'Categories' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Get-OutlookCategoryColor ────────────────────────────────────────────────

Describe 'Get-OutlookCategoryColor' {
    BeforeAll { $cmd = Get-Command Get-OutlookCategoryColor }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}
