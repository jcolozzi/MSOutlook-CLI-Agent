---
description: "Use when the user mentions Outlook, OutlookPOSH, PowerShell COM automation for Outlook, or asks to read/send/manage email, calendar, contacts, or tasks."
---
# OutlookPOSH Module

When working with Microsoft Outlook automation, import the PowerShell module:

```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSOutlook-agent\OutlookPOSH\OutlookPOSH.psd1" -Force
```

This module provides 81 PowerShell functions for full Outlook automation via COM across 13 categories: Application, Account, Folder, Mail, Calendar, Contact, Task, Attachment, Category, Search, Export, Rule, and Metadata. Use `-AsJson` on any function for structured output. The `@outlook-dev` agent has the complete function reference.

## Error Handling

If the module fails to import or a function fails, follow these steps in order:

### Step 1: Verify Module Path
Confirm the path exists and is syntactically correct. Verify the current user has accessible permissions.
- **If this fails:** Proceed to Step 2.

### Step 2: Check File Exists
Verify OutlookPOSH.psd1 file exists at the specified location.
- **If file is missing:** Check the path with your system administrator.
- **If file exists:** Proceed to Step 3.

### Step 3: Verify Dependencies & Permissions
Check for missing dependencies and verify sufficient permissions are granted. The module requires PowerShell 5.1+ and a running Outlook instance.
- **If dependencies are missing:** Install required components.
- **If permissions are insufficient:** Request elevated access.
- **If both are OK:** Proceed to Step 4.

### Step 4: Check Execution Policy
Ensure PowerShell execution policy allows module loading.
- **If policy blocks execution:** Run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
- **If policy is acceptable:** Proceed to Step 5.

### Step 5: Verify Trust Center / Object Model Guard
Outlook's Object Model Guard blocks programmatic access to protected properties (email addresses, message bodies, attachments, address book) unless Trust Center is configured.
- **Configure:** File → Options → Trust Center → Trust Center Settings → Programmatic Access → set to "Never warn me about suspicious activity"
- **Group Policy alternative:** In managed environments, use Group Policy to suppress Object Model Guard prompts
- **If prompts persist:** Ensure antivirus software is detected by Outlook (Outlook trusts programmatic access when a recognized AV is active)
- **If issue persists:** Consult the system administrator or refer to the OutlookPOSH documentation for advanced troubleshooting.
