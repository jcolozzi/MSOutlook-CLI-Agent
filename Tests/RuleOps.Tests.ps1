# Tests/RuleOps.Tests.ps1
# Structural tests for RuleOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookRule ──────────────────────────────────────────────────────────

Describe 'Get-OutlookRule' {
    BeforeAll { $cmd = Get-Command Get-OutlookRule }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookRule ──────────────────────────────────────────────────────────

Describe 'New-OutlookRule' {
    BeforeAll { $cmd = Get-Command New-OutlookRule }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -MoveToFolder parameter' { $cmd.Parameters.Keys | Should -Contain 'MoveToFolder' }
    It 'has -FromAddress parameter' { $cmd.Parameters.Keys | Should -Contain 'FromAddress' }
    It 'has -SubjectContains parameter' { $cmd.Parameters.Keys | Should -Contain 'SubjectContains' }
    It 'has -Enabled parameter' { $cmd.Parameters.Keys | Should -Contain 'Enabled' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookRule ───────────────────────────────────────────────────────

Describe 'Remove-OutlookRule' {
    BeforeAll { $cmd = Get-Command Remove-OutlookRule }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookRuleEnabled ──────────────────────────────────────────────────

Describe 'Set-OutlookRuleEnabled' {
    BeforeAll { $cmd = Get-Command Set-OutlookRuleEnabled }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -Enabled parameter' { $cmd.Parameters.Keys | Should -Contain 'Enabled' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Invoke-OutlookRules ─────────────────────────────────────────────────────

Describe 'Invoke-OutlookRules' {
    BeforeAll { $cmd = Get-Command Invoke-OutlookRules }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -RuleName parameter' { $cmd.Parameters.Keys | Should -Contain 'RuleName' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
