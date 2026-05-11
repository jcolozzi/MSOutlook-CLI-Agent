@{
    RootModule        = 'OutlookPOSH.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b7e3c9a1-4f2d-4e8b-9a16-d5c7f0e2b834'
    Author            = 'Colozzi'
    CompanyName       = 'VA'
    Copyright         = '(c) 2025-2026. All rights reserved.'
    Description       = 'PowerShell COM automation for Microsoft Outlook – mail, calendar, contacts, tasks, attachments, categories, search, export, rules, and metadata operations.'
    PowerShellVersion = '5.1'
    CompatiblePSEditions = @('Desktop','Core')
    FunctionsToExport = @(
        # ── ApplicationOps (3) ──
        'Get-OutlookApplicationInfo'
        'Get-OutlookTip'
        'Close-OutlookSession'

        # ── AccountOps (4) ──
        'Get-OutlookAccount'
        'Get-OutlookDefaultAccount'
        'Get-OutlookStore'
        'Get-OutlookCurrentUser'

        # ── FolderOps (8) ──
        'Get-OutlookDefaultFolder'
        'Get-OutlookFolder'
        'Get-OutlookFolderList'
        'Get-OutlookFolderInfo'
        'New-OutlookFolder'
        'Rename-OutlookFolder'
        'Remove-OutlookFolder'
        'Move-OutlookFolder'

        # ── MailOps (14) ──
        'Get-OutlookMail'
        'Get-OutlookMailItem'
        'New-OutlookMailDraft'
        'Send-OutlookMail'
        'Send-OutlookMailDraft'
        'Reply-OutlookMail'
        'Reply-OutlookMailAll'
        'Forward-OutlookMail'
        'Move-OutlookMail'
        'Copy-OutlookMail'
        'Remove-OutlookMail'
        'Set-OutlookMailRead'
        'Set-OutlookMailFlag'
        'Set-OutlookMailImportance'

        # ── CalendarOps (10) ──
        'Get-OutlookAppointment'
        'Get-OutlookAppointmentItem'
        'New-OutlookAppointment'
        'New-OutlookMeeting'
        'Set-OutlookAppointment'
        'Remove-OutlookAppointment'
        'Send-OutlookMeetingResponse'
        'Get-OutlookFreeBusy'
        'Get-OutlookRecurrence'
        'Set-OutlookRecurrence'

        # ── ContactOps (8) ──
        'Get-OutlookContact'
        'Get-OutlookContactItem'
        'New-OutlookContact'
        'Set-OutlookContact'
        'Remove-OutlookContact'
        'Find-OutlookContact'
        'Get-OutlookDistributionList'
        'New-OutlookDistributionList'

        # ── TaskOps (7) ──
        'Get-OutlookTask'
        'Get-OutlookTaskItem'
        'New-OutlookTask'
        'Set-OutlookTask'
        'Complete-OutlookTask'
        'Remove-OutlookTask'
        'Set-OutlookTaskReminder'

        # ── AttachmentOps (5) ──
        'Get-OutlookAttachment'
        'Save-OutlookAttachment'
        'Save-OutlookAllAttachments'
        'Add-OutlookAttachment'
        'Remove-OutlookAttachment'

        # ── CategoryOps (5) ──
        'Get-OutlookCategory'
        'New-OutlookCategory'
        'Remove-OutlookCategory'
        'Set-OutlookItemCategory'
        'Get-OutlookCategoryColor'

        # ── SearchOps (4) ──
        'Find-OutlookItem'
        'Find-OutlookMailBySubject'
        'Find-OutlookMailBySender'
        'Find-OutlookMailByDate'

        # ── ExportOps (4) ──
        'Export-OutlookItem'
        'Export-OutlookCalendar'
        'Export-OutlookContacts'
        'Export-OutlookFolderItems'

        # ── RuleOps (5) ──
        'Get-OutlookRule'
        'New-OutlookRule'
        'Remove-OutlookRule'
        'Set-OutlookRuleEnabled'
        'Invoke-OutlookRules'

        # ── MetadataOps (4) ──
        'Get-OutlookItemProperty'
        'Set-OutlookItemUserProperty'
        'Get-OutlookConversation'
        'Get-OutlookMessageHeader'
    )
    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
    PrivateData = @{
        PSData = @{
            Tags       = @('Outlook','COM','Automation','Office','Email','Calendar')
            ProjectUri = ''
        }
    }
}
