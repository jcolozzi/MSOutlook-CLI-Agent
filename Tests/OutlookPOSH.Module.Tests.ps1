# Tests/OutlookPOSH.Module.Tests.ps1
# Pester 5+ tests — module loading, function exports, file structure

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]param()

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\OutlookPOSH\OutlookPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
}

Describe 'OutlookPOSH Module' {

    Context 'Module manifest' {
        It 'Manifest file exists' {
            Test-Path $modulePath | Should -BeTrue
        }

        It 'Manifest is valid' {
            { Test-ModuleManifest -Path $modulePath -ErrorAction Stop } | Should -Not -Throw
        }

        It 'Manifest has correct RootModule' {
            $manifest = Test-ModuleManifest -Path $modulePath
            $manifest.RootModule | Should -Be 'OutlookPOSH.psm1'
        }
    }

    Context 'Module loads' {
        BeforeAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }

        AfterAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It 'Module is loaded' {
            Get-Module OutlookPOSH | Should -Not -BeNullOrEmpty
        }

        It 'Module version is 1.0.0' {
            (Get-Module OutlookPOSH).Version.ToString() | Should -Be '1.0.0'
        }

        It 'Exports exactly 81 public functions' {
            $exported = (Get-Module OutlookPOSH).ExportedFunctions.Keys
            $exported.Count | Should -Be 81
        }
    }

    Context 'Expected function exports' {
        BeforeAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module OutlookPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Exports <_>" -ForEach @(
            # ApplicationOps (3)
            'Get-OutlookApplicationInfo', 'Get-OutlookTip', 'Close-OutlookSession',
            # AccountOps (4)
            'Get-OutlookAccount', 'Get-OutlookDefaultAccount', 'Get-OutlookStore',
            'Get-OutlookCurrentUser',
            # FolderOps (8)
            'Get-OutlookDefaultFolder', 'Get-OutlookFolder', 'Get-OutlookFolderList',
            'Get-OutlookFolderInfo', 'New-OutlookFolder', 'Rename-OutlookFolder',
            'Remove-OutlookFolder', 'Move-OutlookFolder',
            # MailOps (14)
            'Get-OutlookMail', 'Get-OutlookMailItem', 'New-OutlookMailDraft',
            'Send-OutlookMail', 'Send-OutlookMailDraft', 'Reply-OutlookMail',
            'Reply-OutlookMailAll', 'Forward-OutlookMail', 'Move-OutlookMail',
            'Copy-OutlookMail', 'Remove-OutlookMail', 'Set-OutlookMailRead',
            'Set-OutlookMailFlag', 'Set-OutlookMailImportance',
            # CalendarOps (10)
            'Get-OutlookAppointment', 'Get-OutlookAppointmentItem',
            'New-OutlookAppointment', 'New-OutlookMeeting', 'Set-OutlookAppointment',
            'Remove-OutlookAppointment', 'Send-OutlookMeetingResponse',
            'Get-OutlookFreeBusy', 'Get-OutlookRecurrence', 'Set-OutlookRecurrence',
            # ContactOps (8)
            'Get-OutlookContact', 'Get-OutlookContactItem', 'New-OutlookContact',
            'Set-OutlookContact', 'Remove-OutlookContact', 'Find-OutlookContact',
            'Get-OutlookDistributionList', 'New-OutlookDistributionList',
            # TaskOps (7)
            'Get-OutlookTask', 'Get-OutlookTaskItem', 'New-OutlookTask',
            'Set-OutlookTask', 'Complete-OutlookTask', 'Remove-OutlookTask',
            'Set-OutlookTaskReminder',
            # AttachmentOps (5)
            'Get-OutlookAttachment', 'Save-OutlookAttachment',
            'Save-OutlookAllAttachments', 'Add-OutlookAttachment',
            'Remove-OutlookAttachment',
            # CategoryOps (5)
            'Get-OutlookCategory', 'New-OutlookCategory', 'Remove-OutlookCategory',
            'Set-OutlookItemCategory', 'Get-OutlookCategoryColor',
            # SearchOps (4)
            'Find-OutlookItem', 'Find-OutlookMailBySubject',
            'Find-OutlookMailBySender', 'Find-OutlookMailByDate',
            # ExportOps (4)
            'Export-OutlookItem', 'Export-OutlookCalendar',
            'Export-OutlookContacts', 'Export-OutlookFolderItems',
            # RuleOps (5)
            'Get-OutlookRule', 'New-OutlookRule', 'Remove-OutlookRule',
            'Set-OutlookRuleEnabled', 'Invoke-OutlookRules',
            # MetadataOps (4)
            'Get-OutlookItemProperty', 'Set-OutlookItemUserProperty',
            'Get-OutlookConversation', 'Get-OutlookMessageHeader'
        ) {
            $script:exported | Should -Contain $_
        }
    }

    Context 'No private functions leaked' {
        BeforeAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
            $script:exported = (Get-Module OutlookPOSH).ExportedFunctions.Keys
        }

        AfterAll {
            Get-Module OutlookPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It "Does NOT export private function <_>" -ForEach @(
            'Test-OutlookAlive',
            'Connect-OutlookSession',
            'Disconnect-OutlookSession',
            'ConvertTo-OutlookSafeValue',
            'Format-OutlookOutput',
            'Resolve-EnumValue',
            'Build-DASLFilter'
        ) {
            $script:exported | Should -Not -Contain $_
        }
    }

    Context 'File structure' {
        It 'Private/ folder exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Private') | Should -BeTrue
        }

        It 'Public/ folder exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public') | Should -BeTrue
        }

        It 'Private/Session.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Private\Session.ps1') | Should -BeTrue
        }

        It 'Private/Utilities.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Private\Utilities.ps1') | Should -BeTrue
        }

        It 'Public/ApplicationOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\ApplicationOps.ps1') | Should -BeTrue
        }

        It 'Public/AccountOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\AccountOps.ps1') | Should -BeTrue
        }

        It 'Public/FolderOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\FolderOps.ps1') | Should -BeTrue
        }

        It 'Public/MailOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\MailOps.ps1') | Should -BeTrue
        }

        It 'Public/CalendarOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\CalendarOps.ps1') | Should -BeTrue
        }

        It 'Public/ContactOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\ContactOps.ps1') | Should -BeTrue
        }

        It 'Public/TaskOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\TaskOps.ps1') | Should -BeTrue
        }

        It 'Public/AttachmentOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\AttachmentOps.ps1') | Should -BeTrue
        }

        It 'Public/CategoryOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\CategoryOps.ps1') | Should -BeTrue
        }

        It 'Public/SearchOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\SearchOps.ps1') | Should -BeTrue
        }

        It 'Public/ExportOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\ExportOps.ps1') | Should -BeTrue
        }

        It 'Public/RuleOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\RuleOps.ps1') | Should -BeTrue
        }

        It 'Public/MetadataOps.ps1 exists' {
            Test-Path (Join-Path $PSScriptRoot '..\OutlookPOSH\Public\MetadataOps.ps1') | Should -BeTrue
        }
    }
}
