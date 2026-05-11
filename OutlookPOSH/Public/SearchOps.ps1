# ── Search Operations ────────────────────────────────────────────────────────

function Find-OutlookItem {
    <#
    .SYNOPSIS
        Searches a folder using a DASL/Jet filter string.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER Filter
        DASL or Jet filter string for Restrict().
    .PARAMETER MaxItems
        Maximum items to return (default 25).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Find-OutlookItem -Filter "[SenderName] = 'Alice'"
    .EXAMPLE
        Find-OutlookItem -FolderType calendar -Filter "[Subject] LIKE '%Review%'" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$FolderType = 'inbox',
        [string]$FolderPath,
        [string]$Filter,
        [int]$MaxItems = 25,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Filter)) { throw 'Find-OutlookItem: -Filter is required.' }

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $items    = $folder.Items
    $filtered = $items.Restrict($Filter)

    $list  = @()
    $count = 0
    foreach ($item in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $entry = @{
                entryID = ConvertTo-OutlookSafeValue $item.EntryID
                subject = ConvertTo-OutlookSafeValue $item.Subject
                class   = ConvertTo-OutlookSafeValue $item.Class
            }
            # Add type-appropriate date/sender fields
            try { $entry.receivedTime = ConvertTo-OutlookSafeValue $item.ReceivedTime } catch { }
            try { $entry.start        = ConvertTo-OutlookSafeValue $item.Start }        catch { }
            try { $entry.senderName   = ConvertTo-OutlookSafeValue $item.SenderName }   catch { }
            try { $entry.organizer    = ConvertTo-OutlookSafeValue $item.Organizer }    catch { }

            $list += [PSCustomObject]$entry
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Find-OutlookMailBySubject {
    <#
    .SYNOPSIS
        Finds mail items whose subject contains the given text.
    .PARAMETER Subject
        Text to search for in the subject line.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER MaxItems
        Maximum items to return (default 25).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Find-OutlookMailBySubject -Subject 'Budget'
    .EXAMPLE
        Find-OutlookMailBySubject -Subject 'Meeting' -MaxItems 10 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$Subject,
        [string]$FolderType = 'inbox',
        [int]$MaxItems = 25,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Subject)) { throw 'Find-OutlookMailBySubject: -Subject is required.' }

    $app = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderType $FolderType

    $filter   = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%$Subject%'"
    $items    = $folder.Items
    $filtered = $items.Restrict($filter)

    $list  = @()
    $count = 0
    foreach ($item in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += [PSCustomObject]@{
                entryID      = ConvertTo-OutlookSafeValue $item.EntryID
                subject      = ConvertTo-OutlookSafeValue $item.Subject
                senderName   = ConvertTo-OutlookSafeValue $item.SenderName
                receivedTime = ConvertTo-OutlookSafeValue $item.ReceivedTime
            }
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Find-OutlookMailBySender {
    <#
    .SYNOPSIS
        Finds mail items from a specific sender email address.
    .PARAMETER SenderEmail
        The sender email address to filter by.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER MaxItems
        Maximum items to return (default 25).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Find-OutlookMailBySender -SenderEmail 'alice@example.com'
    #>
    [CmdletBinding()]
    param(
        [string]$SenderEmail,
        [string]$FolderType = 'inbox',
        [int]$MaxItems = 25,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($SenderEmail)) { throw 'Find-OutlookMailBySender: -SenderEmail is required.' }

    $app = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderType $FolderType

    $filter   = "[SenderEmailAddress] = '$SenderEmail'"
    $items    = $folder.Items
    $filtered = $items.Restrict($filter)

    $list  = @()
    $count = 0
    foreach ($item in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += [PSCustomObject]@{
                entryID      = ConvertTo-OutlookSafeValue $item.EntryID
                subject      = ConvertTo-OutlookSafeValue $item.Subject
                senderName   = ConvertTo-OutlookSafeValue $item.SenderName
                receivedTime = ConvertTo-OutlookSafeValue $item.ReceivedTime
            }
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Find-OutlookMailByDate {
    <#
    .SYNOPSIS
        Finds mail items received within a date range.
    .PARAMETER After
        Start of the date range (inclusive).
    .PARAMETER Before
        End of the date range (inclusive).
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER MaxItems
        Maximum items to return (default 25).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Find-OutlookMailByDate -After '2026-05-01' -Before '2026-05-10'
    #>
    [CmdletBinding()]
    param(
        [datetime]$After,
        [datetime]$Before,
        [string]$FolderType = 'inbox',
        [int]$MaxItems = 25,
        [switch]$AsJson
    )

    if ($After  -eq [datetime]::MinValue) { throw 'Find-OutlookMailByDate: -After is required.' }
    if ($Before -eq [datetime]::MinValue) { throw 'Find-OutlookMailByDate: -Before is required.' }

    $app = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderType $FolderType

    $afterStr  = $After.ToString('MM/dd/yyyy')
    $beforeStr = $Before.ToString('MM/dd/yyyy')
    $filter    = "[ReceivedTime] >= '$afterStr' AND [ReceivedTime] <= '$beforeStr'"
    $items     = $folder.Items
    $items.Sort('[ReceivedTime]', $true)
    $filtered  = $items.Restrict($filter)

    $list  = @()
    $count = 0
    foreach ($item in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += [PSCustomObject]@{
                entryID      = ConvertTo-OutlookSafeValue $item.EntryID
                subject      = ConvertTo-OutlookSafeValue $item.Subject
                senderName   = ConvertTo-OutlookSafeValue $item.SenderName
                receivedTime = ConvertTo-OutlookSafeValue $item.ReceivedTime
            }
            $count++
        } catch {
            Write-Verbose "Skipped item: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}
