# Tests/SearchOps.Tests.ps1
# Structural tests for SearchOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Find-OutlookItem ─────────────────────────────────────────────────────────

Describe 'Find-OutlookItem' {
    BeforeAll { $cmd = Get-Command Find-OutlookItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -Filter parameter' { $cmd.Parameters.Keys | Should -Contain 'Filter' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Find-OutlookMailBySubject ────────────────────────────────────────────────

Describe 'Find-OutlookMailBySubject' {
    BeforeAll { $cmd = Get-Command Find-OutlookMailBySubject }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Find-OutlookMailBySender ─────────────────────────────────────────────────

Describe 'Find-OutlookMailBySender' {
    BeforeAll { $cmd = Get-Command Find-OutlookMailBySender }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -SenderEmail parameter' { $cmd.Parameters.Keys | Should -Contain 'SenderEmail' }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Find-OutlookMailByDate ───────────────────────────────────────────────────

Describe 'Find-OutlookMailByDate' {
    BeforeAll { $cmd = Get-Command Find-OutlookMailByDate }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -After parameter' { $cmd.Parameters.Keys | Should -Contain 'After' }
    It 'has -Before parameter' { $cmd.Parameters.Keys | Should -Contain 'Before' }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}
