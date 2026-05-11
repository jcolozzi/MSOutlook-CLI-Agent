# ── AttachmentOps.ps1 ───────────────────────────────────────────────────────
# Attachment operations for OutlookPOSH
# ────────────────────────────────────────────────────────────────────────────

function Get-OutlookAttachment {
<#
.SYNOPSIS
    Lists attachments on an Outlook item.
.PARAMETER EntryID
    The EntryID of the item whose attachments to list.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookAttachment -EntryID '00000...'
#>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Get-OutlookAttachment: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    # Build reverse lookup for attachment type friendly names
    $typeReverse = @{}
    foreach ($key in $script:OL_ATTACHMENT_TYPE.Keys) {
        $typeReverse[$script:OL_ATTACHMENT_TYPE[$key]] = $key
    }

    $results = [System.Collections.Generic.List[hashtable]]::new()

    for ($i = 1; $i -le $item.Attachments.Count; $i++) {
        $att      = $item.Attachments.Item($i)
        $typeVal  = $att.Type
        $typeName = if ($typeReverse.ContainsKey($typeVal)) { $typeReverse[$typeVal] } else { "Unknown($typeVal)" }

        $results.Add(@{
            index       = $i
            fileName    = ConvertTo-OutlookSafeValue $att.FileName
            displayName = ConvertTo-OutlookSafeValue $att.DisplayName
            size        = ConvertTo-OutlookSafeValue $att.Size
            type        = $typeName
        })
    }

    Format-OutlookOutput -Data $results -AsJson:$AsJson
}

function Save-OutlookAttachment {
<#
.SYNOPSIS
    Saves a single attachment from an Outlook item to disk.
.PARAMETER EntryID
    The EntryID of the item containing the attachment.
.PARAMETER AttachmentIndex
    1-based index of the attachment to save.
.PARAMETER DestinationPath
    Folder path where the file will be saved.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Save-OutlookAttachment -EntryID '00000...' -AttachmentIndex 1 -DestinationPath 'C:\Downloads'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [int]$AttachmentIndex,
        [string]$DestinationPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Save-OutlookAttachment: -EntryID is required."
    }
    if ($AttachmentIndex -lt 1) {
        throw "Save-OutlookAttachment: -AttachmentIndex must be >= 1."
    }
    if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
        throw "Save-OutlookAttachment: -DestinationPath is required."
    }
    if (-not (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
        throw "Save-OutlookAttachment: destination folder does not exist: $DestinationPath"
    }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($AttachmentIndex -gt $item.Attachments.Count) {
        throw "Save-OutlookAttachment: AttachmentIndex $AttachmentIndex exceeds attachment count ($($item.Attachments.Count))."
    }

    $att      = $item.Attachments.Item($AttachmentIndex)
    $fullPath = Join-Path -Path $DestinationPath -ChildPath $att.FileName

    if (-not $PSCmdlet.ShouldProcess($fullPath, 'Save attachment')) { return }

    $att.SaveAsFile($fullPath)

    $result = @{
        fileName  = ConvertTo-OutlookSafeValue $att.FileName
        savedPath = $fullPath
        status    = 'Saved'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Save-OutlookAllAttachments {
<#
.SYNOPSIS
    Saves all attachments from an Outlook item to a folder.
.PARAMETER EntryID
    The EntryID of the item whose attachments to save.
.PARAMETER DestinationPath
    Folder path where files will be saved.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Save-OutlookAllAttachments -EntryID '00000...' -DestinationPath 'C:\Downloads'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$DestinationPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Save-OutlookAllAttachments: -EntryID is required."
    }
    if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
        throw "Save-OutlookAllAttachments: -DestinationPath is required."
    }
    if (-not (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
        throw "Save-OutlookAllAttachments: destination folder does not exist: $DestinationPath"
    }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($item.Attachments.Count -eq 0) {
        $result = @{ savedPaths = @(); status = 'NoAttachments' }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
        return
    }

    if (-not $PSCmdlet.ShouldProcess("$($item.Attachments.Count) attachment(s)", 'Save all attachments')) { return }

    $savedPaths = [System.Collections.Generic.List[string]]::new()

    for ($i = 1; $i -le $item.Attachments.Count; $i++) {
        $att      = $item.Attachments.Item($i)
        $fullPath = Join-Path -Path $DestinationPath -ChildPath $att.FileName
        $att.SaveAsFile($fullPath)
        $savedPaths.Add($fullPath)
    }

    $result = @{
        savedPaths = $savedPaths
        count      = $savedPaths.Count
        status     = 'Saved'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Add-OutlookAttachment {
<#
.SYNOPSIS
    Adds a file attachment to an Outlook item (typically a draft).
.PARAMETER EntryID
    The EntryID of the item to attach the file to.
.PARAMETER FilePath
    Full path to the file to attach.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Add-OutlookAttachment -EntryID '00000...' -FilePath 'C:\Docs\report.pdf'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$FilePath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Add-OutlookAttachment: -EntryID is required."
    }
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        throw "Add-OutlookAttachment: -FilePath is required."
    }
    if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        throw "Add-OutlookAttachment: file does not exist: $FilePath"
    }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    if (-not $PSCmdlet.ShouldProcess($fileName, 'Add attachment')) { return }

    $item.Attachments.Add($FilePath) | Out-Null
    $item.Save()

    $result = @{
        fileName = $fileName
        status   = 'Attached'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookAttachment {
<#
.SYNOPSIS
    Removes an attachment from an Outlook item by index.
.PARAMETER EntryID
    The EntryID of the item containing the attachment.
.PARAMETER AttachmentIndex
    1-based index of the attachment to remove.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Remove-OutlookAttachment -EntryID '00000...' -AttachmentIndex 2
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [int]$AttachmentIndex,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Remove-OutlookAttachment: -EntryID is required."
    }
    if ($AttachmentIndex -lt 1) {
        throw "Remove-OutlookAttachment: -AttachmentIndex must be >= 1."
    }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    if ($AttachmentIndex -gt $item.Attachments.Count) {
        throw "Remove-OutlookAttachment: AttachmentIndex $AttachmentIndex exceeds attachment count ($($item.Attachments.Count))."
    }

    $att      = $item.Attachments.Item($AttachmentIndex)
    $fileName = $att.FileName

    if (-not $PSCmdlet.ShouldProcess($fileName, 'Remove attachment')) { return }

    $item.Attachments.Remove($AttachmentIndex)
    $item.Save()

    $result = @{
        fileName = $fileName
        status   = 'Removed'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
