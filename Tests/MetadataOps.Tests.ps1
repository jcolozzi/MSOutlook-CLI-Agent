# Tests/MetadataOps.Tests.ps1
# Structural tests for MetadataOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookItemProperty ─────────────────────────────────────────────────

Describe 'Get-OutlookItemProperty' {
    BeforeAll { $cmd = Get-Command Get-OutlookItemProperty }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Set-OutlookItemUserProperty ──────────────────────────────────────────────

Describe 'Set-OutlookItemUserProperty' {
    BeforeAll { $cmd = Get-Command Set-OutlookItemUserProperty }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -PropertyName parameter' { $cmd.Parameters.Keys | Should -Contain 'PropertyName' }
    It 'has -PropertyValue parameter' { $cmd.Parameters.Keys | Should -Contain 'PropertyValue' }
    It 'has -PropertyType parameter' { $cmd.Parameters.Keys | Should -Contain 'PropertyType' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'PropertyType has ValidateSet' {
        $attr = $cmd.Parameters['PropertyType'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'text'
        $attr.ValidValues | Should -Contain 'number'
        $attr.ValidValues | Should -Contain 'yesNo'
        $attr.ValidValues | Should -Contain 'dateTime'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Get-OutlookConversation ─────────────────────────────────────────────────

Describe 'Get-OutlookConversation' {
    BeforeAll { $cmd = Get-Command Get-OutlookConversation }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookMessageHeader ────────────────────────────────────────────────

Describe 'Get-OutlookMessageHeader' {
    BeforeAll { $cmd = Get-Command Get-OutlookMessageHeader }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}
