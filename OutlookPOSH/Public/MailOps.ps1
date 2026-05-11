# ── Mail Operations ──────────────────────────────────────────────────────────

function Get-OutlookMail {
    <#
    .SYNOPSIS
        Lists mail items from a folder with optional filtering.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER Filter
        DASL or Jet filter string for Restrict().
    .PARAMETER MaxItems
        Maximum items to return (default 50).
    .PARAMETER UnreadOnly
        Only return unread items.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookMail
    .EXAMPLE
        Get-OutlookMail -UnreadOnly -MaxItems 10
    .EXAMPLE
        Get-OutlookMail -Filter "[SenderName] = 'Alice'" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$FolderType = 'inbox',
        [string]$FolderPath,
        [string]$Filter,
        [int]$MaxItems = 50,
        [switch]$UnreadOnly,
        [switch]$AsJson
    )

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $items = $folder.Items
    $items.Sort('[ReceivedTime]', $true)

    if ($UnreadOnly) {
        $items = $items.Restrict('[UnRead] = True')
    }

    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $items = $items.Restrict($Filter)
    }

    $list = @()
    $count = 0
    foreach ($item in $items) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += @{
                entryID        = ConvertTo-OutlookSafeValue $item.EntryID
                subject        = ConvertTo-OutlookSafeValue $item.Subject
                senderName     = ConvertTo-OutlookSafeValue $item.SenderName
                senderEmail    = ConvertTo-OutlookSafeValue $item.SenderEmailAddress
                receivedTime   = ConvertTo-OutlookSafeValue $item.ReceivedTime
                importance     = ConvertTo-OutlookSafeValue $item.Importance
                unRead         = ConvertTo-OutlookSafeValue $item.UnRead
                hasAttachments = ConvertTo-OutlookSafeValue $item.Attachments.Count -gt 0
                size           = ConvertTo-OutlookSafeValue $item.Size
                categories     = ConvertTo-OutlookSafeValue $item.Categories
            }
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    if ($AsJson) {
        return ($list | ForEach-Object { [PSCustomObject]$_ }) | ConvertTo-Json -Depth 10 -Compress
    }
    return $list | ForEach-Object { [PSCustomObject]$_ }
}

function Get-OutlookMailItem {
    <#
    .SYNOPSIS
        Returns full details for a single mail item by EntryID.
    .PARAMETER EntryID
        The EntryID of the mail item.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookMailItem -EntryID '000000...'
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookMailItem: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $body     = ConvertTo-OutlookSafeValue $item.Body
    $htmlBody = ConvertTo-OutlookSafeValue $item.HTMLBody
    if ($body -is [string] -and $body.Length -gt 5000)         { $body     = $body.Substring(0, 5000) }
    if ($htmlBody -is [string] -and $htmlBody.Length -gt 10000) { $htmlBody = $htmlBody.Substring(0, 10000) }

    $result = @{
        entryID         = ConvertTo-OutlookSafeValue $item.EntryID
        subject         = ConvertTo-OutlookSafeValue $item.Subject
        body            = $body
        htmlBody        = $htmlBody
        senderName      = ConvertTo-OutlookSafeValue $item.SenderName
        senderEmail     = ConvertTo-OutlookSafeValue $item.SenderEmailAddress
        to              = ConvertTo-OutlookSafeValue $item.To
        cc              = ConvertTo-OutlookSafeValue $item.CC
        bcc             = ConvertTo-OutlookSafeValue $item.BCC
        receivedTime    = ConvertTo-OutlookSafeValue $item.ReceivedTime
        sentOn          = ConvertTo-OutlookSafeValue $item.SentOn
        importance      = ConvertTo-OutlookSafeValue $item.Importance
        sensitivity     = ConvertTo-OutlookSafeValue $item.Sensitivity
        unRead          = ConvertTo-OutlookSafeValue $item.UnRead
        hasAttachments  = $item.Attachments.Count -gt 0
        attachmentCount = ConvertTo-OutlookSafeValue $item.Attachments.Count
        categories      = ConvertTo-OutlookSafeValue $item.Categories
        conversationID  = ConvertTo-OutlookSafeValue $item.ConversationID
        size            = ConvertTo-OutlookSafeValue $item.Size
        bodyFormat      = ConvertTo-OutlookSafeValue $item.BodyFormat
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookMailDraft {
    <#
    .SYNOPSIS
        Creates an unsent mail draft.
    .PARAMETER Subject
        Mail subject.
    .PARAMETER Body
        Mail body text.
    .PARAMETER To
        Recipient(s), semicolon-delimited.
    .PARAMETER CC
        CC recipient(s), semicolon-delimited.
    .PARAMETER BCC
        BCC recipient(s), semicolon-delimited.
    .PARAMETER BodyFormat
        Body format: plain, html, richText (default html).
    .PARAMETER Importance
        Importance: low, normal, high.
    .PARAMETER Sensitivity
        Sensitivity: normal, personal, private, confidential.
    .PARAMETER Attachments
        Array of file paths to attach.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookMailDraft -Subject 'Hello' -Body '<p>Hi</p>' -To 'user@example.com'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Subject,
        [string]$Body,
        [string]$To,
        [string]$CC,
        [string]$BCC,
        [string]$BodyFormat = 'html',
        [string]$Importance,
        [string]$Sensitivity,
        [string[]]$Attachments,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) { throw 'New-OutlookMailDraft: -Subject is required.' }
    if ([string]::IsNullOrWhiteSpace($To))      { throw 'New-OutlookMailDraft: -To is required.' }

    $app = Connect-OutlookSession

    if ($PSCmdlet.ShouldProcess($Subject, 'Create mail draft')) {
        $mail = $app.CreateItem(0)  # olMailItem
        $mail.Subject = $Subject
        $mail.To      = $To

        $fmtVal = Resolve-EnumValue -Map $script:OL_BODY_FORMAT -Key $BodyFormat -EnumName 'BodyFormat'
        $mail.BodyFormat = $fmtVal

        if ($fmtVal -eq 2) { $mail.HTMLBody = $Body } else { $mail.Body = $Body }

        if (-not [string]::IsNullOrWhiteSpace($CC))  { $mail.CC  = $CC }
        if (-not [string]::IsNullOrWhiteSpace($BCC)) { $mail.BCC = $BCC }

        if (-not [string]::IsNullOrWhiteSpace($Importance)) {
            $mail.Importance = Resolve-EnumValue -Map $script:OL_IMPORTANCE -Key $Importance -EnumName 'Importance'
        }
        if (-not [string]::IsNullOrWhiteSpace($Sensitivity)) {
            $mail.Sensitivity = Resolve-EnumValue -Map $script:OL_SENSITIVITY -Key $Sensitivity -EnumName 'Sensitivity'
        }

        if ($Attachments) {
            foreach ($path in $Attachments) {
                if (-not (Test-Path $path)) { throw "New-OutlookMailDraft: attachment not found: $path" }
                $mail.Attachments.Add($path) | Out-Null
            }
        }

        $mail.Save()

        $result = @{
            status  = 'draft_created'
            entryID = ConvertTo-OutlookSafeValue $mail.EntryID
            subject = ConvertTo-OutlookSafeValue $mail.Subject
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Send-OutlookMail {
    <#
    .SYNOPSIS
        Creates and sends a mail message immediately.
    .PARAMETER Subject
        Mail subject.
    .PARAMETER Body
        Mail body text.
    .PARAMETER To
        Recipient(s), semicolon-delimited.
    .PARAMETER CC
        CC recipient(s), semicolon-delimited.
    .PARAMETER BCC
        BCC recipient(s), semicolon-delimited.
    .PARAMETER BodyFormat
        Body format: plain, html, richText (default html).
    .PARAMETER Importance
        Importance: low, normal, high.
    .PARAMETER Sensitivity
        Sensitivity: normal, personal, private, confidential.
    .PARAMETER Attachments
        Array of file paths to attach.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Send-OutlookMail -Subject 'Report' -Body 'See attached.' -To 'boss@example.com'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Subject,
        [string]$Body,
        [string]$To,
        [string]$CC,
        [string]$BCC,
        [string]$BodyFormat = 'html',
        [string]$Importance,
        [string]$Sensitivity,
        [string[]]$Attachments,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) { throw 'Send-OutlookMail: -Subject is required.' }
    if ([string]::IsNullOrWhiteSpace($To))      { throw 'Send-OutlookMail: -To is required.' }

    $app = Connect-OutlookSession

    if ($PSCmdlet.ShouldProcess("$Subject → $To", 'Send mail')) {
        $mail = $app.CreateItem(0)
        $mail.Subject = $Subject
        $mail.To      = $To

        $fmtVal = Resolve-EnumValue -Map $script:OL_BODY_FORMAT -Key $BodyFormat -EnumName 'BodyFormat'
        $mail.BodyFormat = $fmtVal

        if ($fmtVal -eq 2) { $mail.HTMLBody = $Body } else { $mail.Body = $Body }

        if (-not [string]::IsNullOrWhiteSpace($CC))  { $mail.CC  = $CC }
        if (-not [string]::IsNullOrWhiteSpace($BCC)) { $mail.BCC = $BCC }

        if (-not [string]::IsNullOrWhiteSpace($Importance)) {
            $mail.Importance = Resolve-EnumValue -Map $script:OL_IMPORTANCE -Key $Importance -EnumName 'Importance'
        }
        if (-not [string]::IsNullOrWhiteSpace($Sensitivity)) {
            $mail.Sensitivity = Resolve-EnumValue -Map $script:OL_SENSITIVITY -Key $Sensitivity -EnumName 'Sensitivity'
        }

        if ($Attachments) {
            foreach ($path in $Attachments) {
                if (-not (Test-Path $path)) { throw "Send-OutlookMail: attachment not found: $path" }
                $mail.Attachments.Add($path) | Out-Null
            }
        }

        $mail.Send()

        $result = @{ status = 'sent'; subject = $Subject; to = $To }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Send-OutlookMailDraft {
    <#
    .SYNOPSIS
        Sends an existing draft mail item by EntryID.
    .PARAMETER EntryID
        The EntryID of the draft to send.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Send-OutlookMailDraft -EntryID '000000...'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Send-OutlookMailDraft: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($PSCmdlet.ShouldProcess($item.Subject, 'Send draft')) {
        $item.Send()

        $result = @{ status = 'sent'; subject = ConvertTo-OutlookSafeValue $item.Subject }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Reply-OutlookMail {
    <#
    .SYNOPSIS
        Replies to a mail item.
    .PARAMETER EntryID
        The EntryID of the mail to reply to.
    .PARAMETER Body
        Reply body text (prepended above the original).
    .PARAMETER Send
        Send the reply immediately instead of saving as draft.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Reply-OutlookMail -EntryID '000000...' -Body 'Thanks!' -Send
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$Body,
        [switch]$Send,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Reply-OutlookMail: -EntryID is required.' }

    $app   = Connect-OutlookSession
    $ns    = $script:OutlookSession.Namespace
    $item  = $ns.GetItemFromID($EntryID)
    $reply = $item.Reply()

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $reply.HTMLBody = $Body + $reply.HTMLBody
    }

    if ($PSCmdlet.ShouldProcess($item.Subject, 'Reply')) {
        if ($Send) { $reply.Send() } else { $reply.Save() }

        $action = if ($Send) { 'sent' } else { 'draft_created' }
        $result = @{ status = $action; subject = ConvertTo-OutlookSafeValue $reply.Subject }
        if (-not $Send) { $result.entryID = ConvertTo-OutlookSafeValue $reply.EntryID }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Reply-OutlookMailAll {
    <#
    .SYNOPSIS
        Replies to all recipients of a mail item.
    .PARAMETER EntryID
        The EntryID of the mail to reply to.
    .PARAMETER Body
        Reply body text (prepended above the original).
    .PARAMETER Send
        Send the reply immediately instead of saving as draft.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Reply-OutlookMailAll -EntryID '000000...' -Body 'Agreed.' -Send
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$Body,
        [switch]$Send,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Reply-OutlookMailAll: -EntryID is required.' }

    $app   = Connect-OutlookSession
    $ns    = $script:OutlookSession.Namespace
    $item  = $ns.GetItemFromID($EntryID)
    $reply = $item.ReplyAll()

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $reply.HTMLBody = $Body + $reply.HTMLBody
    }

    if ($PSCmdlet.ShouldProcess($item.Subject, 'Reply All')) {
        if ($Send) { $reply.Send() } else { $reply.Save() }

        $action = if ($Send) { 'sent' } else { 'draft_created' }
        $result = @{ status = $action; subject = ConvertTo-OutlookSafeValue $reply.Subject }
        if (-not $Send) { $result.entryID = ConvertTo-OutlookSafeValue $reply.EntryID }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Forward-OutlookMail {
    <#
    .SYNOPSIS
        Forwards a mail item to specified recipients.
    .PARAMETER EntryID
        The EntryID of the mail to forward.
    .PARAMETER To
        Recipient(s), semicolon-delimited.
    .PARAMETER Body
        Additional body text (prepended above the original).
    .PARAMETER Send
        Send the forward immediately instead of saving as draft.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Forward-OutlookMail -EntryID '000000...' -To 'user@example.com' -Send
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$To,
        [string]$Body,
        [switch]$Send,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Forward-OutlookMail: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($To))      { throw 'Forward-OutlookMail: -To is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)
    $fwd  = $item.Forward()
    $fwd.To = $To

    if (-not [string]::IsNullOrWhiteSpace($Body)) {
        $fwd.HTMLBody = $Body + $fwd.HTMLBody
    }

    if ($PSCmdlet.ShouldProcess("$($item.Subject) → $To", 'Forward')) {
        if ($Send) { $fwd.Send() } else { $fwd.Save() }

        $action = if ($Send) { 'sent' } else { 'draft_created' }
        $result = @{ status = $action; subject = ConvertTo-OutlookSafeValue $fwd.Subject }
        if (-not $Send) { $result.entryID = ConvertTo-OutlookSafeValue $fwd.EntryID }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Move-OutlookMail {
    <#
    .SYNOPSIS
        Moves a mail item to another folder.
    .PARAMETER EntryID
        The EntryID of the mail to move.
    .PARAMETER DestinationFolderType
        Default folder type of the destination.
    .PARAMETER DestinationFolderPath
        Full path of the destination folder. Takes precedence over -DestinationFolderType.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Move-OutlookMail -EntryID '000000...' -DestinationFolderType 'deletedItems'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$DestinationFolderType,
        [string]$DestinationFolderPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Move-OutlookMail: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($DestinationFolderType) -and [string]::IsNullOrWhiteSpace($DestinationFolderPath)) {
        throw 'Move-OutlookMail: supply either -DestinationFolderType or -DestinationFolderPath.'
    }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if (-not [string]::IsNullOrWhiteSpace($DestinationFolderPath)) {
        $dest = Resolve-OutlookFolder -FolderPath $DestinationFolderPath
    } else {
        $dest = Resolve-OutlookFolder -FolderType $DestinationFolderType
    }

    if ($PSCmdlet.ShouldProcess($item.Subject, "Move to '$($dest.FolderPath)'")) {
        $moved = $item.Move($dest)

        $result = @{
            status  = 'moved'
            entryID = ConvertTo-OutlookSafeValue $moved.EntryID
            folder  = ConvertTo-OutlookSafeValue $dest.FolderPath
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Copy-OutlookMail {
    <#
    .SYNOPSIS
        Creates a copy of a mail item in the same folder.
    .PARAMETER EntryID
        The EntryID of the mail to copy.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Copy-OutlookMail -EntryID '000000...'
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Copy-OutlookMail: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)
    $copy = $item.Copy()

    $result = @{
        status  = 'copied'
        entryID = ConvertTo-OutlookSafeValue $copy.EntryID
        subject = ConvertTo-OutlookSafeValue $copy.Subject
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookMail {
    <#
    .SYNOPSIS
        Deletes a mail item (moves to Deleted Items).
    .PARAMETER EntryID
        The EntryID of the mail to delete.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Remove-OutlookMail -EntryID '000000...'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Remove-OutlookMail: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)
    $subj = $item.Subject

    if ($PSCmdlet.ShouldProcess($subj, 'Delete mail')) {
        $item.Delete()

        $result = @{ status = 'deleted'; subject = $subj }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Set-OutlookMailRead {
    <#
    .SYNOPSIS
        Marks a mail item as read or unread.
    .PARAMETER EntryID
        The EntryID of the mail.
    .PARAMETER Read
        $true to mark read, $false to mark unread.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookMailRead -EntryID '000000...' -Read $true
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [bool]$Read,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Set-OutlookMailRead: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $state = if ($Read) { 'read' } else { 'unread' }

    if ($PSCmdlet.ShouldProcess($item.Subject, "Mark as $state")) {
        $item.UnRead = -not $Read
        $item.Save()

        $result = @{
            status  = $state
            entryID = ConvertTo-OutlookSafeValue $item.EntryID
            subject = ConvertTo-OutlookSafeValue $item.Subject
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Set-OutlookMailFlag {
    <#
    .SYNOPSIS
        Sets the flag status on a mail item.
    .PARAMETER EntryID
        The EntryID of the mail.
    .PARAMETER FlagStatus
        Flag status: noFlag, flagComplete, flagMarked.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookMailFlag -EntryID '000000...' -FlagStatus 'flagMarked'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,

        [ValidateSet('noFlag','flagComplete','flagMarked')]
        [string]$FlagStatus,

        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))    { throw 'Set-OutlookMailFlag: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($FlagStatus)) { throw 'Set-OutlookMailFlag: -FlagStatus is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($PSCmdlet.ShouldProcess($item.Subject, "Set flag to '$FlagStatus'")) {
        $item.FlagStatus = Resolve-EnumValue -Map $script:OL_FLAG_STATUS -Key $FlagStatus -EnumName 'FlagStatus'
        $item.Save()

        $result = @{
            status     = 'flagged'
            flagStatus = $FlagStatus
            entryID    = ConvertTo-OutlookSafeValue $item.EntryID
            subject    = ConvertTo-OutlookSafeValue $item.Subject
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Set-OutlookMailImportance {
    <#
    .SYNOPSIS
        Sets the importance level on a mail item.
    .PARAMETER EntryID
        The EntryID of the mail.
    .PARAMETER Importance
        Importance level: low, normal, high.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookMailImportance -EntryID '000000...' -Importance 'high'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,

        [ValidateSet('low','normal','high')]
        [string]$Importance,

        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))    { throw 'Set-OutlookMailImportance: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($Importance)) { throw 'Set-OutlookMailImportance: -Importance is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($PSCmdlet.ShouldProcess($item.Subject, "Set importance to '$Importance'")) {
        $item.Importance = Resolve-EnumValue -Map $script:OL_IMPORTANCE -Key $Importance -EnumName 'Importance'
        $item.Save()

        $result = @{
            status     = 'updated'
            importance = $Importance
            entryID    = ConvertTo-OutlookSafeValue $item.EntryID
            subject    = ConvertTo-OutlookSafeValue $item.Subject
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}
