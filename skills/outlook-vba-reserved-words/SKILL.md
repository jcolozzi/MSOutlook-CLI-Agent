---
name: outlook-vba-reserved-words
description: >-
  Detect and avoid Microsoft Outlook / VBA reserved words in identifiers.
license: CC-BY-4.0
---

# Outlook/VBA Reserved Words – Naming Safety Skill

## Purpose
Help developers avoid case-insensitive identifier collisions in Outlook VBA and PowerShell scripts that interact with the Outlook object model.

### Collision Categories

**VBA keywords & operators**  
`Dim`, `If`, `Select`, `Function`, `ByVal`, `And`, `Or`, `Not`, `Set`, `New`, `With`, `End`, `Sub`, `Property`, etc.

**Built-in VBA function names**  
`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, `Format`, `CStr`, `CInt`, `CLng`, `CBool`, etc.

**Outlook object model names**  
`MailItem`, `AppointmentItem`, `ContactItem`, `TaskItem`, `JournalItem`, `NoteItem`, `PostItem`, `MeetingItem`, `Folder`, `Folders`, `NameSpace`, `Application`, `Inspector`, `Explorer`, `Attachment`, `Attachments`, `Recipient`, `Recipients`, `AddressEntry`, `AddressList`, `Store`, `Stores`, `Account`, `Accounts`, `Category`, `Categories`, `Rule`, `Rules`, `RuleAction`, `RuleCondition`, `Items`, `Item`, `Selection`, `Conversation`, `PropertyAccessor`, `UserProperty`, `UserProperties`, `OlItemType`, `OlDefaultFolders`, `OlImportance`, `OlFlagStatus`, `OlTaskStatus`, `OlMeetingResponse`, `OlRecurrenceType`, `OlSensitivity`, `OlBodyFormat`.

**DASL property names used in filters**  
`Subject`, `Body`, `SenderEmailAddress`, `ReceivedTime`, `SentOn`, `Importance`, `FlagStatus`, `HasAttachments`, `Categories`, `ConversationTopic`, `ConversationIndex`.

**Special characters/symbols**  
Spaces, `'`, `"`, `.`, `!`, `?`, `*`, `+`, `-`, `=`, `<`, `>`, `#`, `%`, `$`, `&`, `@`, `\`, `/`, `^`, `~`, `{}`, `[]`, `()`.

> Outlook VBA and the COM object model treat names **case-insensitively**; reusing these terms as identifiers leads to compile or runtime errors, or silent shadowing of built-in members.

## Procedure

> **Priority:** Complete Steps 1–4 in order. Step 5 is optional and should be considered only when explicitly requested.

### Step 1: Identify Reserved Words (Priority)
1. **Scan identifiers** in the current context (variables, procedure names, module names, form control names, and embedded references to Outlook objects).
2. **Flag exact case-insensitive matches** to reserved words and report each occurrence with the specific line number and file name.
3. **Partial-match guidance:** Identifiers that *contain* a reserved word as a complete substring (e.g., `MailItemRef`, `FolderPath`) should be noted as potential concerns, but only exact or very obvious collisions require immediate action. Use judgment based on context.

### Step 2: Suggest Safe Replacements
4. **Propose descriptive, context-specific names**:
   - `Item` → `olmMailItem` or `currentItem`
   - `Subject` → `msgSubject` / `apptSubject`
   - `Body` → `msgBody` / `htmlBody`
   - `Name` → `contactName` / `folderName`
   - `Folder` → `fldInbox` / `targetFolder`
   - `Start` → `apptStart` / `dtmStart`
   - `End` → `apptEnd` / `dtmEnd`
   - `Attachment` → `attFile` / `currentAttachment`
   - `Recipients` → `msgRecipients` / `toList`
   - `Categories` → `itemCategories` / `catList`

### Step 3: Apply Naming Conventions (After renaming)
5. **Ensure correct formatting**:
   - Use CamelCase (no spaces or special characters).
   - Apply Outlook-specific type/object prefixes for clarity:
     - `olm` — mail item (e.g., `olmReport`, `olmDraft`)
     - `ola` — appointment item (e.g., `olaStandup`, `olaMeeting`)
     - `olc` — contact item (e.g., `olcJaneDoe`, `olcVendor`)
     - `olt` — task item (e.g., `oltDeliverable`, `oltFollowUp`)
     - `fld` — folder (e.g., `fldInbox`, `fldArchive`)
     - `att` — attachment (e.g., `attReport`, `attInvoice`)
     - `cat` — category (e.g., `catUrgent`, `catProject`)
     - `rul` — rule (e.g., `rulMoveReports`, `rulFlagUrgent`)
   - Standard VBA type prefixes also apply: `str`, `lng`, `dtm`, `bln`, `obj`, `col`, etc.

### Step 4: Handle Existing Objects
6. **If renaming variables or objects** already in use, prefer renaming over relying on disambiguation through qualification (e.g., `Outlook.MailItem` vs. a variable named `MailItem`). While qualification works, renaming avoids subtle shadowing bugs.

### Step 5: Repository Analysis (Optional)
7. **(Optional) Run a repository scan** when requested to list offenders and propose bulk renames across all VBA modules and scripts.

## Common Offenders (teach by example)

- **Outlook object names used as variables:** `Item`, `Folder`, `Attachment`, `Recipient`, `Inspector`, `Explorer`, `Store`, `Account`, `Rule`, `Category`
- **Property names used as variables:** `Subject`, `Body`, `Name`, `Start`, `End`, `Duration`, `Location`, `Importance`, `Sender`, `Recipients`, `Attachments`, `Categories`
- **VBA keywords:** `Date`, `Time`, `Now`, `Value`, `Text`, `Type`, `Class`, `Index`
- **Enumeration names:** `OlItemType`, `OlDefaultFolders`, `OlImportance`, `OlFlagStatus`

## Naming Recommendations

- ✅ Prefer descriptive, specific names: `msgSubject`, `apptStartDate`, `contactEmail`, `taskDueDate`.
- ✅ Use CamelCase; apply type/object prefixes (`olm`, `ola`, `olc`, `olt`, `fld`, `att`, `str`, `dtm`) for clarity.
- ✅ For loop variables iterating Outlook collections, use prefixed names: `For Each olmCurrent In fldInbox.Items`.
- ❌ Avoid reserved words and Outlook object model names as any identifier.
- ❌ Avoid generic names like `Item`, `Folder`, `Data` when they shadow Outlook COM objects.
- ❌ Never name a variable `MailItem`, `AppointmentItem`, `ContactItem`, or `TaskItem` — these are Outlook class names.

## Example Prompts (to trigger this skill)

- "Scan this module and flag any **Outlook/VBA reserved-word** variable names."
- "Suggest safe replacements for variables named `Item` and `Subject` in my mail processing macro."
- "Check my Outlook VBA for identifiers that collide with the Outlook object model."
- "Rename `Folder` and `Attachment` variables to follow Outlook naming conventions."

## References

- [Microsoft: Outlook object model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)
- [Microsoft: VBA language reference – reserved words](https://learn.microsoft.com/en-us/office/vba/language/reference/keywords)
- [Microsoft: Filtering items using DASL syntax](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items)
- [Microsoft: Outlook Object Model Guard](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/security/security-behavior-of-the-outlook-object-model)
- Allen Browne: Problem names & reserved words in Access/VBA (also applies to Outlook VBA)
