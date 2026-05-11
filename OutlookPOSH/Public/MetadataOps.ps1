# ── Metadata Operations ──────────────────────────────────────────────────────

function Get-OutlookItemProperty {
    <#
    .SYNOPSIS
        Gets standard and custom (UserProperties) metadata for an Outlook item.
    .PARAMETER EntryID
        The EntryID of the item.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookItemProperty -EntryID '000000...'
    .EXAMPLE
        Get-OutlookItemProperty -EntryID '000000...' -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookItemProperty: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $userProps = @()
    try {
        for ($i = 1; $i -le $item.UserProperties.Count; $i++) {
            $up = $item.UserProperties.Item($i)
            $userProps += @{
                name  = ConvertTo-OutlookSafeValue $up.Name
                value = ConvertTo-OutlookSafeValue $up.Value
                type  = ConvertTo-OutlookSafeValue $up.Type
            }
        }
    } catch {
        Write-Verbose "Could not read UserProperties: $_"
    }

    $result = @{
        entryID              = ConvertTo-OutlookSafeValue $item.EntryID
        class                = ConvertTo-OutlookSafeValue $item.Class
        messageClass         = ConvertTo-OutlookSafeValue $item.MessageClass
        creationTime         = ConvertTo-OutlookSafeValue $item.CreationTime
        lastModificationTime = ConvertTo-OutlookSafeValue $item.LastModificationTime
        size                 = ConvertTo-OutlookSafeValue $item.Size
        userProperties       = $userProps
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Set-OutlookItemUserProperty {
    <#
    .SYNOPSIS
        Sets a custom user property on an Outlook item.
    .PARAMETER EntryID
        The EntryID of the item.
    .PARAMETER PropertyName
        Name of the user property.
    .PARAMETER PropertyValue
        Value to assign.
    .PARAMETER PropertyType
        Type of the property (default 'text').
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookItemUserProperty -EntryID '000000...' -PropertyName 'ProjectCode' -PropertyValue 'ALPHA'
    .EXAMPLE
        Set-OutlookItemUserProperty -EntryID '000000...' -PropertyName 'Priority' -PropertyValue 5 -PropertyType number
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$PropertyName,
        $PropertyValue,
        [ValidateSet('text','number','yesNo','dateTime')]
        [string]$PropertyType = 'text',
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))      { throw 'Set-OutlookItemUserProperty: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($PropertyName))  { throw 'Set-OutlookItemUserProperty: -PropertyName is required.' }
    if ($null -eq $PropertyValue)                     { throw 'Set-OutlookItemUserProperty: -PropertyValue is required.' }

    # Map friendly type to OlUserPropertyType
    $typeMap = @{ text=1; number=3; dateTime=5; yesNo=6 }
    $olType  = $typeMap[$PropertyType]

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $subj = ConvertTo-OutlookSafeValue $item.Subject
    if (-not $PSCmdlet.ShouldProcess("$subj [$PropertyName]", "Set user property to '$PropertyValue'")) { return }

    $prop = $item.UserProperties.Add($PropertyName, $olType, $true)
    $prop.Value = $PropertyValue
    $item.Save()

    $result = @{
        entryID       = ConvertTo-OutlookSafeValue $item.EntryID
        subject       = $subj
        propertyName  = $PropertyName
        propertyValue = ConvertTo-OutlookSafeValue $prop.Value
        propertyType  = $PropertyType
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookConversation {
    <#
    .SYNOPSIS
        Gets all items in the same conversation thread as the specified item.
    .PARAMETER EntryID
        The EntryID of a mail item in the conversation.
    .PARAMETER MaxItems
        Maximum conversation items to return (default 20).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookConversation -EntryID '000000...'
    .EXAMPLE
        Get-OutlookConversation -EntryID '000000...' -MaxItems 10 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [int]$MaxItems = 20,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookConversation: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $convID = $item.ConversationID
    if ([string]::IsNullOrWhiteSpace($convID)) {
        throw 'Get-OutlookConversation: item has no ConversationID.'
    }

    # Search the parent folder for items with the same ConversationID
    $folder   = $item.Parent
    $filter   = "[ConversationID] = '$convID'"
    $items    = $folder.Items
    $items.Sort('[ReceivedTime]', $false)
    $filtered = $items.Restrict($filter)

    $list  = @()
    $count = 0
    foreach ($conv in $filtered) {
        if ($count -ge $MaxItems) { break }
        try {
            $list += [PSCustomObject]@{
                entryID      = ConvertTo-OutlookSafeValue $conv.EntryID
                subject      = ConvertTo-OutlookSafeValue $conv.Subject
                senderName   = ConvertTo-OutlookSafeValue $conv.SenderName
                receivedTime = ConvertTo-OutlookSafeValue $conv.ReceivedTime
            }
            $count++
        } catch {
            Write-Verbose "Skipped conversation item: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Get-OutlookMessageHeader {
    <#
    .SYNOPSIS
        Gets the internet message headers for a mail item.
    .PARAMETER EntryID
        The EntryID of the mail item.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookMessageHeader -EntryID '000000...'
    .EXAMPLE
        Get-OutlookMessageHeader -EntryID '000000...' -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) { throw 'Get-OutlookMessageHeader: -EntryID is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $headerPropTag = 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'
    $headers = $null
    try {
        $headers = $item.PropertyAccessor.GetProperty($headerPropTag)
    } catch {
        Write-Verbose "Could not retrieve message headers: $_"
    }

    $result = @{
        entryID = ConvertTo-OutlookSafeValue $item.EntryID
        subject = ConvertTo-OutlookSafeValue $item.Subject
        headers = ConvertTo-OutlookSafeValue $headers
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
