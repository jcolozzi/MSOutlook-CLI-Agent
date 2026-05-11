<#
.SYNOPSIS
    OutlookPOSH – PowerShell COM automation for Microsoft Outlook.

.DESCRIPTION
    Provides 81 cmdlet-style functions to drive Outlook via its COM object model.
    Covers mail, calendar, contacts, tasks, attachments, categories, search,
    export, rules, and metadata operations.

    Outlook is a COM singleton — New-Object -ComObject Outlook.Application always
    attaches to the running instance or launches a new one.  The module NEVER
    calls Application.Quit so the user's Outlook session is never disrupted.

.EXAMPLE
    Import-Module .\OutlookPOSH\OutlookPOSH.psd1
    Get-OutlookMail -Folder Inbox -Count 10
#>

# ── Outlook Enum Constants ───────────────────────────────────────────────────

$script:OL_ITEM_TYPE = @{ mail=0; appointment=1; contact=2; task=3; journal=4; note=5; post=6; distList=7 }

$script:OL_DEFAULT_FOLDER = @{
    deletedItems=3; outbox=4; sentMail=5; inbox=6; calendar=9; contacts=10
    journal=11; notes=12; tasks=13; drafts=16; conflicts=19; syncIssues=20
    localFailures=21; serverFailures=22; junk=23; rssFeeds=25; toDo=28
    managedEmail=29; suggestedContacts=30
}

$script:OL_IMPORTANCE = @{ low=0; normal=1; high=2 }
$script:OL_SENSITIVITY = @{ normal=0; personal=1; private=2; confidential=3 }
$script:OL_BODY_FORMAT = @{ unspecified=0; plain=1; html=2; richText=3 }
$script:OL_BUSY_STATUS = @{ free=0; tentative=1; busy=2; outOfOffice=3; workingElsewhere=4 }
$script:OL_MEETING_STATUS = @{ nonMeeting=0; meeting=1; received=3; canceled=5; receivedAndCanceled=7 }
$script:OL_RESPONSE_STATUS = @{ none=0; organized=1; tentative=2; accepted=3; declined=4; notResponded=5 }
$script:OL_MAIL_RECIPIENT_TYPE = @{ originator=0; to=1; cc=2; bcc=3 }
$script:OL_MEETING_RECIPIENT_TYPE = @{ organizer=0; required=1; optional=2; resource=3 }
$script:OL_TASK_STATUS = @{ notStarted=0; inProgress=1; complete=2; waiting=3; deferred=4 }
$script:OL_SAVE_AS_TYPE = @{ txt=0; rtf=1; template=2; msg=3; doc=4; html=5; vCard=6; vCal=7; iCal=8; msgUnicode=9; mhtml=10 }
$script:OL_ATTACHMENT_TYPE = @{ byValue=1; byReference=4; embeddedItem=5; ole=6 }
$script:OL_FLAG_STATUS = @{ noFlag=0; flagComplete=1; flagMarked=2 }
$script:OL_RECURRENCE_TYPE = @{ daily=0; weekly=1; monthly=2; monthNth=3; yearly=5; yearNth=6 }

$script:OL_CATEGORY_COLOR = @{
    none=0; red=1; orange=2; peach=3; yellow=4; green=5; teal=6; olive=7
    blue=8; purple=9; maroon=10; steel=11; darkSteel=12; gray=13; darkGray=14
    black=15; darkRed=16; darkOrange=17; darkPeach=18; darkYellow=19
    darkGreen=20; darkTeal=21; darkOlive=22; darkBlue=23; darkPurple=24; darkMaroon=25
}

$script:OL_OBJECT_CLASS = @{
    mailItem=43; appointmentItem=26; contactItem=40; taskItem=48
    journalItem=42; noteItem=44; postItem=45; distListItem=69; meetingItem=53; reportItem=46
}

# ── Session State ────────────────────────────────────────────────────────────

$script:OutlookSession = @{
    App       = $null
    Namespace = $null
}

# ── Dot-source Private & Public functions ────────────────────────────────────

foreach ($file in Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1") { . $file.FullName }
foreach ($file in Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1")  { . $file.FullName }

# ── Module Cleanup ───────────────────────────────────────────────────────────
# Release COM references on module removal — NEVER call .Quit() as that would
# close the user's running Outlook instance.

$ExecutionContext.SessionState.Module.OnRemove += {
    if ($null -ne $script:OutlookSession.App) {
        try {
            if ($null -ne $script:OutlookSession.Namespace) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:OutlookSession.Namespace) | Out-Null
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:OutlookSession.App) | Out-Null
        } catch { }
        $script:OutlookSession.App = $null
        $script:OutlookSession.Namespace = $null
    }
}
