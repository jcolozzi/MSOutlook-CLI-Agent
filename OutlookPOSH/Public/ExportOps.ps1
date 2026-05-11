# ── Export Operations ────────────────────────────────────────────────────────

function Export-OutlookItem {
    <#
    .SYNOPSIS
        Saves a single Outlook item to a file in the specified format.
    .PARAMETER EntryID
        The EntryID of the item to export.
    .PARAMETER DestinationPath
        Full file path for the exported item.
    .PARAMETER Format
        Export format (default 'msg').
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Export-OutlookItem -EntryID '000000...' -DestinationPath 'C:\Export\mail.msg'
    .EXAMPLE
        Export-OutlookItem -EntryID '000000...' -DestinationPath 'C:\Export\mail.html' -Format html
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$DestinationPath,
        [ValidateSet('txt','rtf','msg','msgUnicode','html','mhtml','template','iCal','vCal','vCard')]
        [string]$Format = 'msg',
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))         { throw 'Export-OutlookItem: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($DestinationPath))  { throw 'Export-OutlookItem: -DestinationPath is required.' }

    $destDir = Split-Path -Path $DestinationPath -Parent
    if (-not (Test-Path -LiteralPath $destDir)) { throw "Export-OutlookItem: destination directory '$destDir' does not exist." }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $formatValue = Resolve-EnumValue -Map $script:OL_SAVE_AS_TYPE -Key $Format -EnumName 'SaveAsType'

    $subj = ConvertTo-OutlookSafeValue $item.Subject
    if (-not $PSCmdlet.ShouldProcess("$subj -> $DestinationPath", "Export item as $Format")) { return }

    $item.SaveAs($DestinationPath, $formatValue)

    $result = @{
        entryID         = ConvertTo-OutlookSafeValue $item.EntryID
        subject         = $subj
        format          = $Format
        destinationPath = $DestinationPath
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Export-OutlookCalendar {
    <#
    .SYNOPSIS
        Exports calendar appointments in a date range to ICS files.
    .PARAMETER Start
        Start of the date range.
    .PARAMETER End
        End of the date range.
    .PARAMETER DestinationPath
        Folder path where ICS files will be saved.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Export-OutlookCalendar -Start '2026-05-01' -End '2026-05-31' -DestinationPath 'C:\Export\Cal'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [datetime]$Start,
        [datetime]$End,
        [string]$DestinationPath,
        [switch]$AsJson
    )

    if ($Start -eq [datetime]::MinValue)                 { throw 'Export-OutlookCalendar: -Start is required.' }
    if ($End   -eq [datetime]::MinValue)                 { throw 'Export-OutlookCalendar: -End is required.' }
    if ([string]::IsNullOrWhiteSpace($DestinationPath))  { throw 'Export-OutlookCalendar: -DestinationPath is required.' }
    if (-not (Test-Path -LiteralPath $DestinationPath))  { throw "Export-OutlookCalendar: destination folder '$DestinationPath' does not exist." }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $folder = $ns.GetDefaultFolder(9)  # olFolderCalendar

    $items = $folder.Items
    $items.IncludeRecurrences = $true
    $items.Sort('[Start]')

    $startStr = $Start.ToString('MM/dd/yyyy HH:mm')
    $endStr   = $End.ToString('MM/dd/yyyy HH:mm')
    $filter   = "[Start] >= '$startStr' AND [End] <= '$endStr'"
    $filtered = $items.Restrict($filter)

    if (-not $PSCmdlet.ShouldProcess("$($Start.ToString('yyyy-MM-dd')) to $($End.ToString('yyyy-MM-dd'))", 'Export calendar items to ICS')) { return }

    $exported = @()
    $count    = 0
    $iCalType = $script:OL_SAVE_AS_TYPE['iCal']  # 8

    foreach ($item in $filtered) {
        try {
            $safeName = ($item.Subject -replace '[\\/:*?"<>|]', '_')
            if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "appointment_$count" }
            $filePath = Join-Path $DestinationPath "$safeName.ics"

            # Avoid overwriting — append index if needed
            if (Test-Path -LiteralPath $filePath) {
                $filePath = Join-Path $DestinationPath "${safeName}_$count.ics"
            }

            $item.SaveAs($filePath, $iCalType)
            $exported += $filePath
            $count++
        } catch {
            Write-Verbose "Skipped calendar item: $_"
        }
    }

    $result = @{
        exportedCount   = $count
        destinationPath = $DestinationPath
        files           = $exported
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Export-OutlookContacts {
    <#
    .SYNOPSIS
        Exports contacts to VCF (vCard) files.
    .PARAMETER DestinationPath
        Folder path where VCF files will be saved.
    .PARAMETER MaxItems
        Maximum contacts to export (default 100).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Export-OutlookContacts -DestinationPath 'C:\Export\Contacts'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$DestinationPath,
        [int]$MaxItems = 100,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($DestinationPath)) { throw 'Export-OutlookContacts: -DestinationPath is required.' }
    if (-not (Test-Path -LiteralPath $DestinationPath)) { throw "Export-OutlookContacts: destination folder '$DestinationPath' does not exist." }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $folder = $ns.GetDefaultFolder(10)  # olFolderContacts

    if (-not $PSCmdlet.ShouldProcess($DestinationPath, 'Export contacts to VCF')) { return }

    $vCardType = $script:OL_SAVE_AS_TYPE['vCard']  # 6
    $exported  = @()
    $count     = 0

    foreach ($item in $folder.Items) {
        if ($count -ge $MaxItems) { break }
        try {
            # Only export ContactItem (class 40)
            if ($item.Class -ne 40) { continue }

            $safeName = ($item.FullName -replace '[\\/:*?"<>|]', '_')
            if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "contact_$count" }
            $filePath = Join-Path $DestinationPath "$safeName.vcf"

            if (Test-Path -LiteralPath $filePath) {
                $filePath = Join-Path $DestinationPath "${safeName}_$count.vcf"
            }

            $item.SaveAs($filePath, $vCardType)
            $exported += $filePath
            $count++
        } catch {
            Write-Verbose "Skipped contact: $_"
        }
    }

    $result = @{
        exportedCount   = $count
        destinationPath = $DestinationPath
        files           = $exported
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Export-OutlookFolderItems {
    <#
    .SYNOPSIS
        Batch exports items from a folder to individual files.
    .PARAMETER FolderType
        Default folder type (e.g. inbox, sentMail).
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER DestinationPath
        Folder path where files will be saved.
    .PARAMETER Format
        Export format (default 'msg').
    .PARAMETER MaxItems
        Maximum items to export (default 100).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Export-OutlookFolderItems -FolderType inbox -DestinationPath 'C:\Export\Inbox'
    .EXAMPLE
        Export-OutlookFolderItems -FolderPath '\\Mailbox\Archive' -DestinationPath 'C:\Export' -Format html -MaxItems 50
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$FolderType,
        [string]$FolderPath,
        [string]$DestinationPath,
        [ValidateSet('txt','rtf','msg','msgUnicode','html','mhtml','template','iCal','vCal','vCard')]
        [string]$Format = 'msg',
        [int]$MaxItems = 100,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($DestinationPath)) { throw 'Export-OutlookFolderItems: -DestinationPath is required.' }
    if (-not (Test-Path -LiteralPath $DestinationPath)) { throw "Export-OutlookFolderItems: destination folder '$DestinationPath' does not exist." }
    if ([string]::IsNullOrWhiteSpace($FolderType) -and [string]::IsNullOrWhiteSpace($FolderPath)) {
        throw 'Export-OutlookFolderItems: supply either -FolderType or -FolderPath.'
    }

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $formatValue = Resolve-EnumValue -Map $script:OL_SAVE_AS_TYPE -Key $Format -EnumName 'SaveAsType'
    $folderName  = $folder.Name

    if (-not $PSCmdlet.ShouldProcess("$folderName (up to $MaxItems items)", "Export as $Format to $DestinationPath")) { return }

    $exported = @()
    $count    = 0

    foreach ($item in $folder.Items) {
        if ($count -ge $MaxItems) { break }
        try {
            $subj = try { $item.Subject } catch { "item_$count" }
            $safeName = ($subj -replace '[\\/:*?"<>|]', '_')
            if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "item_$count" }
            $ext      = if ($Format -eq 'msgUnicode') { 'msg' } else { $Format }
            $filePath = Join-Path $DestinationPath "$safeName.$ext"

            if (Test-Path -LiteralPath $filePath) {
                $filePath = Join-Path $DestinationPath "${safeName}_$count.$ext"
            }

            $item.SaveAs($filePath, $formatValue)
            $exported += $filePath
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    $result = @{
        folder          = $folderName
        exportedCount   = $count
        format          = $Format
        destinationPath = $DestinationPath
        files           = $exported
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
