# Tests/ContactOps.Tests.ps1
# Structural tests for ContactOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookContact ───────────────────────────────────────────────────────

Describe 'Get-OutlookContact' {
    BeforeAll { $cmd = Get-Command Get-OutlookContact }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -Filter parameter' { $cmd.Parameters.Keys | Should -Contain 'Filter' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookContactItem ──────────────────────────────────────────────────

Describe 'Get-OutlookContactItem' {
    BeforeAll { $cmd = Get-Command Get-OutlookContactItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookContact ──────────────────────────────────────────────────────

Describe 'New-OutlookContact' {
    BeforeAll { $cmd = Get-Command New-OutlookContact }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FirstName parameter' { $cmd.Parameters.Keys | Should -Contain 'FirstName' }
    It 'has -LastName parameter' { $cmd.Parameters.Keys | Should -Contain 'LastName' }
    It 'has -Email parameter' { $cmd.Parameters.Keys | Should -Contain 'Email' }
    It 'has -CompanyName parameter' { $cmd.Parameters.Keys | Should -Contain 'CompanyName' }
    It 'has -JobTitle parameter' { $cmd.Parameters.Keys | Should -Contain 'JobTitle' }
    It 'has -BusinessPhone parameter' { $cmd.Parameters.Keys | Should -Contain 'BusinessPhone' }
    It 'has -MobilePhone parameter' { $cmd.Parameters.Keys | Should -Contain 'MobilePhone' }
    It 'has -HomePhone parameter' { $cmd.Parameters.Keys | Should -Contain 'HomePhone' }
    It 'has -Categories parameter' { $cmd.Parameters.Keys | Should -Contain 'Categories' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookContact ──────────────────────────────────────────────────────

Describe 'Set-OutlookContact' {
    BeforeAll { $cmd = Get-Command Set-OutlookContact }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -FirstName parameter' { $cmd.Parameters.Keys | Should -Contain 'FirstName' }
    It 'has -LastName parameter' { $cmd.Parameters.Keys | Should -Contain 'LastName' }
    It 'has -Email parameter' { $cmd.Parameters.Keys | Should -Contain 'Email' }
    It 'has -CompanyName parameter' { $cmd.Parameters.Keys | Should -Contain 'CompanyName' }
    It 'has -JobTitle parameter' { $cmd.Parameters.Keys | Should -Contain 'JobTitle' }
    It 'has -BusinessPhone parameter' { $cmd.Parameters.Keys | Should -Contain 'BusinessPhone' }
    It 'has -MobilePhone parameter' { $cmd.Parameters.Keys | Should -Contain 'MobilePhone' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Remove-OutlookContact ───────────────────────────────────────────────────

Describe 'Remove-OutlookContact' {
    BeforeAll { $cmd = Get-Command Remove-OutlookContact }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Find-OutlookContact ─────────────────────────────────────────────────────

Describe 'Find-OutlookContact' {
    BeforeAll { $cmd = Get-Command Find-OutlookContact }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -Email parameter' { $cmd.Parameters.Keys | Should -Contain 'Email' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookDistributionList ──────────────────────────────────────────────

Describe 'Get-OutlookDistributionList' {
    BeforeAll { $cmd = Get-Command Get-OutlookDistributionList }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookDistributionList ──────────────────────────────────────────────

Describe 'New-OutlookDistributionList' {
    BeforeAll { $cmd = Get-Command New-OutlookDistributionList }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Name parameter' { $cmd.Parameters.Keys | Should -Contain 'Name' }
    It 'has -Members parameter' { $cmd.Parameters.Keys | Should -Contain 'Members' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
