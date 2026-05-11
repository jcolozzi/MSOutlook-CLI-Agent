# Tests/MailOps.Tests.ps1
# Structural tests for MailOps functions (no COM required)

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

# ── Get-OutlookMail ──────────────────────────────────────────────────────────

Describe 'Get-OutlookMail' {
    BeforeAll { $cmd = Get-Command Get-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -FolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderType' }
    It 'has -FolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'FolderPath' }
    It 'has -Filter parameter' { $cmd.Parameters.Keys | Should -Contain 'Filter' }
    It 'has -MaxItems parameter' { $cmd.Parameters.Keys | Should -Contain 'MaxItems' }
    It 'has -UnreadOnly parameter' { $cmd.Parameters.Keys | Should -Contain 'UnreadOnly' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Get-OutlookMailItem ──────────────────────────────────────────────────────

Describe 'Get-OutlookMailItem' {
    BeforeAll { $cmd = Get-Command Get-OutlookMailItem }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── New-OutlookMailDraft ─────────────────────────────────────────────────────

Describe 'New-OutlookMailDraft' {
    BeforeAll { $cmd = Get-Command New-OutlookMailDraft }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -To parameter' { $cmd.Parameters.Keys | Should -Contain 'To' }
    It 'has -CC parameter' { $cmd.Parameters.Keys | Should -Contain 'CC' }
    It 'has -BCC parameter' { $cmd.Parameters.Keys | Should -Contain 'BCC' }
    It 'has -BodyFormat parameter' { $cmd.Parameters.Keys | Should -Contain 'BodyFormat' }
    It 'has -Importance parameter' { $cmd.Parameters.Keys | Should -Contain 'Importance' }
    It 'has -Sensitivity parameter' { $cmd.Parameters.Keys | Should -Contain 'Sensitivity' }
    It 'has -Attachments parameter' { $cmd.Parameters.Keys | Should -Contain 'Attachments' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Send-OutlookMail ─────────────────────────────────────────────────────────

Describe 'Send-OutlookMail' {
    BeforeAll { $cmd = Get-Command Send-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -Subject parameter' { $cmd.Parameters.Keys | Should -Contain 'Subject' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -To parameter' { $cmd.Parameters.Keys | Should -Contain 'To' }
    It 'has -CC parameter' { $cmd.Parameters.Keys | Should -Contain 'CC' }
    It 'has -BCC parameter' { $cmd.Parameters.Keys | Should -Contain 'BCC' }
    It 'has -BodyFormat parameter' { $cmd.Parameters.Keys | Should -Contain 'BodyFormat' }
    It 'has -Attachments parameter' { $cmd.Parameters.Keys | Should -Contain 'Attachments' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Send-OutlookMailDraft ────────────────────────────────────────────────────

Describe 'Send-OutlookMailDraft' {
    BeforeAll { $cmd = Get-Command Send-OutlookMailDraft }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Reply-OutlookMail ────────────────────────────────────────────────────────

Describe 'Reply-OutlookMail' {
    BeforeAll { $cmd = Get-Command Reply-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -Send parameter' { $cmd.Parameters.Keys | Should -Contain 'Send' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Reply-OutlookMailAll ─────────────────────────────────────────────────────

Describe 'Reply-OutlookMailAll' {
    BeforeAll { $cmd = Get-Command Reply-OutlookMailAll }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -Send parameter' { $cmd.Parameters.Keys | Should -Contain 'Send' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Forward-OutlookMail ──────────────────────────────────────────────────────

Describe 'Forward-OutlookMail' {
    BeforeAll { $cmd = Get-Command Forward-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -To parameter' { $cmd.Parameters.Keys | Should -Contain 'To' }
    It 'has -Body parameter' { $cmd.Parameters.Keys | Should -Contain 'Body' }
    It 'has -Send parameter' { $cmd.Parameters.Keys | Should -Contain 'Send' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Move-OutlookMail ─────────────────────────────────────────────────────────

Describe 'Move-OutlookMail' {
    BeforeAll { $cmd = Get-Command Move-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -DestinationFolderType parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationFolderType' }
    It 'has -DestinationFolderPath parameter' { $cmd.Parameters.Keys | Should -Contain 'DestinationFolderPath' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Copy-OutlookMail ─────────────────────────────────────────────────────────

Describe 'Copy-OutlookMail' {
    BeforeAll { $cmd = Get-Command Copy-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }
}

# ── Remove-OutlookMail ───────────────────────────────────────────────────────

Describe 'Remove-OutlookMail' {
    BeforeAll { $cmd = Get-Command Remove-OutlookMail }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookMailRead ──────────────────────────────────────────────────────

Describe 'Set-OutlookMailRead' {
    BeforeAll { $cmd = Get-Command Set-OutlookMailRead }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Read parameter' { $cmd.Parameters.Keys | Should -Contain 'Read' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookMailFlag ──────────────────────────────────────────────────────

Describe 'Set-OutlookMailFlag' {
    BeforeAll { $cmd = Get-Command Set-OutlookMailFlag }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -FlagStatus parameter' { $cmd.Parameters.Keys | Should -Contain 'FlagStatus' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'FlagStatus has ValidateSet' {
        $attr = $cmd.Parameters['FlagStatus'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'noFlag'
        $attr.ValidValues | Should -Contain 'flagComplete'
        $attr.ValidValues | Should -Contain 'flagMarked'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}

# ── Set-OutlookMailImportance ────────────────────────────────────────────────

Describe 'Set-OutlookMailImportance' {
    BeforeAll { $cmd = Get-Command Set-OutlookMailImportance }

    It 'is exported' { $cmd | Should -Not -BeNullOrEmpty }
    It 'has CmdletBinding' { $cmd.CmdletBinding | Should -BeTrue }
    It 'has -EntryID parameter' { $cmd.Parameters.Keys | Should -Contain 'EntryID' }
    It 'has -Importance parameter' { $cmd.Parameters.Keys | Should -Contain 'Importance' }
    It 'has -AsJson parameter' { $cmd.Parameters.Keys | Should -Contain 'AsJson' }

    It 'Importance has ValidateSet' {
        $attr = $cmd.Parameters['Importance'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] }
        $attr | Should -Not -BeNullOrEmpty
        $attr.ValidValues | Should -Contain 'low'
        $attr.ValidValues | Should -Contain 'normal'
        $attr.ValidValues | Should -Contain 'high'
    }

    It 'supports ShouldProcess' {
        $attr = $cmd.ScriptBlock.Attributes | Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
        $attr.SupportsShouldProcess | Should -BeTrue
    }
}
