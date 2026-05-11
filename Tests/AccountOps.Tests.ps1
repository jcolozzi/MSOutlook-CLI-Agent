# Tests/AccountOps.Tests.ps1
# Structural tests for AccountOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookAccount ───────────────────────────────────────────────────────

Describe 'Get-OutlookAccount' {
    BeforeAll { $cmd = Get-Command Get-OutlookAccount }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookDefaultAccount ────────────────────────────────────────────────

Describe 'Get-OutlookDefaultAccount' {
    BeforeAll { $cmd = Get-Command Get-OutlookDefaultAccount }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookStore ─────────────────────────────────────────────────────────

Describe 'Get-OutlookStore' {
    BeforeAll { $cmd = Get-Command Get-OutlookStore }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookCurrentUser ───────────────────────────────────────────────────

Describe 'Get-OutlookCurrentUser' {
    BeforeAll { $cmd = Get-Command Get-OutlookCurrentUser }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}
