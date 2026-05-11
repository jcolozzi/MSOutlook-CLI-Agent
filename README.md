# MSOutlook-CLI-Agent

> **Automate Microsoft Outlook from plain English inside VS Code — no MCP server, no Python, no extra processes.**

![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue?logo=windows)
![PowerShell: 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![VS Code](https://img.shields.io/badge/VS%20Code-GitHub%20Copilot%20Chat-blueviolet?logo=visual-studio-code)
![Functions: 81](https://img.shields.io/badge/functions-81-brightgreen)
![Module Version](https://img.shields.io/badge/version-1.0.0-orange)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

## What is this?

**MSOutlook-CLI-Agent** is a VS Code agent (powered by GitHub Copilot Chat) that lets you talk to Microsoft Outlook in plain language. You describe what you want and the agent translates it into PowerShell commands that automate your mail, calendar, contacts, and tasks live via COM — no manual VBA editing required.

```text
You:   "Send an email to user@domain.com with subject 'Hello' and search my inbox for the quarterly report"
Agent: → Send-OutlookMail → Find-OutlookMailBySubject → confirms success
```

The **OutlookPOSH** module (included) provides **81 public functions** across 13 categories covering mail, calendar, contacts, tasks, attachments, categories, search, export, rules, accounts, folders, and metadata.

> **Note:** Unlike the file-based agents (Access, Word, Excel, PowerPoint), Outlook is session-based. There is no "open file" step — the module connects to the running Outlook instance and never calls `.Quit()`.

## How it works

```text
VS Code Copilot Chat (agent mode)
        │
        ▼
  outlook-dev agent (.md instructions)
        │  describes which PowerShell command to run
        ▼
  OutlookPOSH module  (imported in the VS Code terminal)
        │  COM calls via Outlook Object Model
        ▼
  Microsoft Outlook (running instance)
```

- **No separate server** — the module runs directly in the VS Code integrated terminal.
- **No Python / Node** — pure PowerShell 5.1+ on Windows.
- **Full COM access** — everything you can do from VBA, you can do from the agent.
- **-WhatIf / -Confirm** — all state-changing functions support PowerShell's standard risk-mitigation flags.
- **Pester tests** — 13 test files cover every public command group.

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 (COM automation is Windows-only) |
| **Microsoft Outlook** | Outlook 2016, 2019, 2021, or Microsoft 365 (desktop) |
| **PowerShell** | 5.1 (Windows PowerShell) **or** PowerShell 7+ |
| **VS Code** | Latest stable, with the **GitHub Copilot Chat** extension |
| **Copilot** | An active GitHub Copilot subscription |

### Trust Center Configuration (Object Model Guard)

Outlook's **Object Model Guard** displays security prompts when external programs access mail, contacts, or the address book. To suppress these prompts for uninterrupted automation:

1. Open **Outlook → File → Options → Trust Center → Trust Center Settings**
2. Select **Programmatic Access**
3. Choose **"Never warn me about suspicious activity (not recommended)"**

**Alternative approaches:**

- **Antivirus auto-suppress** — If a recognized, up-to-date antivirus program is installed and registered with Windows Security Center, Outlook automatically suppresses Object Model Guard prompts.
- **Group Policy** — Enforce the setting organization-wide via registry:
  ```
  HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Security
  Value: ObjectModelGuard (DWORD) = 2
  ```
  `2` = Never warn; `1` = Warn if AV is inactive; `0` = Always warn.

> **Note:** Without this configuration, Outlook will pop up a dialog every time the module accesses protected properties, blocking automation.

## Setup

### 1 — Clone the repo

```powershell
git clone https://github.com/jcolozzi/MSOutlook-CLI-Agent.git
```

### 2 — Install the agent instructions

Choose **one** of the following:

#### Option A — User-level (available in every workspace)

Copy the `.agent.md` file from the repo root to:
```
C:\Users\%USERNAME%\AppData\Roaming\Code\User\prompts\
```

#### Option B — Workspace-level (scoped to this project)

Copy the `.agent.md` file into a `.github\agents\` folder in your workspace root. VS Code automatically detects any `.md` files in that folder as custom agents.

> [!NOTE]
> VS Code detects any `.md` files in the `.github/agents/` folder of your workspace as custom agents.

### 3 — Update the module path inside the agent file

Open the `.agent.md` file and replace the placeholder path with the actual path to `OutlookPOSH.psd1` on your machine:

```powershell
# Before
Import-Module "C:\path\to\OutlookPOSH\OutlookPOSH.psd1"

# After (example)
Import-Module "C:\Projects\MSOutlook-agent\OutlookPOSH\OutlookPOSH.psd1"
```

### 4 — Select the agent and start prompting

In VS Code Copilot Chat, click the agent picker and choose **outlook-dev**. Outlook must be running — then start describing what you want.

## Usage examples

| Prompt | Functions called |
|---|---|
| "Show my Outlook version and account info" | `Get-OutlookApplicationInfo` |
| "Read the 10 most recent inbox messages" | `Get-OutlookMail` |
| "Send an email to user@domain.com with subject 'Hello'" | `Send-OutlookMail` |
| "Get calendar appointments for the next 7 days" | `Get-OutlookAppointment` |
| "Create a task 'Review docs' due in 3 days" | `New-OutlookTask` |
| "Search my inbox for 'quarterly report'" | `Find-OutlookMailBySubject` |
| "Save all attachments from the latest email to C:\Downloads" | `Save-OutlookAllAttachments` |
| "Create a meeting with user@domain.com tomorrow at 2pm" | `New-OutlookMeeting` |
| "Export my contacts to CSV" | `Export-OutlookContacts` |
| "Create a rule to move emails from boss@company.com to VIP folder" | `New-OutlookRule` |
| "List all my email accounts" | `Get-OutlookAccount` |
| "Mark all unread emails in Inbox as read" | `Set-OutlookMailRead` |

## Project structure

```text
MSOutlook-agent/
├── OutlookPOSH/                 # PowerShell module (the engine)
│   ├── OutlookPOSH.psd1        # Module manifest (v1.0.0, PS 5.1+, Desktop + Core)
│   ├── OutlookPOSH.psm1        # Module loader
│   ├── Public/                  # 13 files — one per command category
│   │   ├── ApplicationOps.ps1
│   │   ├── AccountOps.ps1
│   │   ├── FolderOps.ps1
│   │   ├── MailOps.ps1
│   │   └── ...
│   └── Private/                 # Internal helpers (COM session, error formatting, etc.)
├── Tests/                       # Pester test suite — 13 test files
│   ├── OutlookPOSH.Module.Tests.ps1
│   ├── MailOps.Tests.ps1
│   └── ...
├── .agent.md                    # Agent instructions (the Copilot Chat prompt)
└── README.md
```

## Running the tests

```powershell
# From the repo root
Invoke-Pester .\Tests\ -Output Detailed
```

> Requires [Pester](https://github.com/pester/Pester) 5.x: `Install-Module Pester -MinimumVersion 5.0 -Force`

## Function reference

<details>
<summary><strong>View all 81 public functions</strong></summary>

| Category | Functions |
|---|---|
| **Application** | `Get-OutlookApplicationInfo`, `Get-OutlookTip`, `Close-OutlookSession` |
| **Account** | `Get-OutlookAccount`, `Get-OutlookDefaultAccount`, `Get-OutlookStore`, `Get-OutlookCurrentUser` |
| **Folder** | `Get-OutlookDefaultFolder`, `Get-OutlookFolder`, `Get-OutlookFolderList`, `Get-OutlookFolderInfo`, `New-OutlookFolder`, `Rename-OutlookFolder`, `Remove-OutlookFolder`, `Move-OutlookFolder` |
| **Mail** | `Get-OutlookMail`, `Get-OutlookMailItem`, `New-OutlookMailDraft`, `Send-OutlookMail`, `Send-OutlookMailDraft`, `Reply-OutlookMail`, `Reply-OutlookMailAll`, `Forward-OutlookMail`, `Move-OutlookMail`, `Copy-OutlookMail`, `Remove-OutlookMail`, `Set-OutlookMailRead`, `Set-OutlookMailFlag`, `Set-OutlookMailImportance` |
| **Calendar** | `Get-OutlookAppointment`, `Get-OutlookAppointmentItem`, `New-OutlookAppointment`, `New-OutlookMeeting`, `Set-OutlookAppointment`, `Remove-OutlookAppointment`, `Send-OutlookMeetingResponse`, `Get-OutlookFreeBusy`, `Get-OutlookRecurrence`, `Set-OutlookRecurrence` |
| **Contact** | `Get-OutlookContact`, `Get-OutlookContactItem`, `New-OutlookContact`, `Set-OutlookContact`, `Remove-OutlookContact`, `Find-OutlookContact`, `Get-OutlookDistributionList`, `New-OutlookDistributionList` |
| **Task** | `Get-OutlookTask`, `Get-OutlookTaskItem`, `New-OutlookTask`, `Set-OutlookTask`, `Complete-OutlookTask`, `Remove-OutlookTask`, `Set-OutlookTaskReminder` |
| **Attachment** | `Get-OutlookAttachment`, `Save-OutlookAttachment`, `Save-OutlookAllAttachments`, `Add-OutlookAttachment`, `Remove-OutlookAttachment` |
| **Category** | `Get-OutlookCategory`, `New-OutlookCategory`, `Remove-OutlookCategory`, `Set-OutlookItemCategory`, `Get-OutlookCategoryColor` |
| **Search** | `Find-OutlookItem`, `Find-OutlookMailBySubject`, `Find-OutlookMailBySender`, `Find-OutlookMailByDate` |
| **Export** | `Export-OutlookItem`, `Export-OutlookCalendar`, `Export-OutlookContacts`, `Export-OutlookFolderItems` |
| **Rule** | `Get-OutlookRule`, `New-OutlookRule`, `Remove-OutlookRule`, `Set-OutlookRuleEnabled`, `Invoke-OutlookRules` |
| **Metadata** | `Get-OutlookItemProperty`, `Set-OutlookItemUserProperty`, `Get-OutlookConversation`, `Get-OutlookMessageHeader` |

</details>

All state-changing functions support `-WhatIf` and `-Confirm` via PowerShell's standard `ShouldProcess` mechanism.

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you would like to change. Please include or update Pester tests for any new or modified functions.

## Credits

- PowerShell port and VS Code agent integration: OutlookPOSH

## License

[MIT](LICENSE) © 2026 OutlookPOSH
