# ── Folder Operations ────────────────────────────────────────────────────────

# ── Local helper ─────────────────────────────────────────────────────────────

function Resolve-OutlookFolder {
    <#
    .SYNOPSIS
        Resolves a Folder COM object from -FolderType or -FolderPath.
    #>
    param(
        [string]$FolderType,
        [string]$FolderPath
    )

    $ns = $script:OutlookSession.Namespace

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        # Navigate hierarchy by splitting on backslash
        $parts = $FolderPath.Trim('\').Split('\') | Where-Object { $_ -ne '' }
        if ($parts.Count -eq 0) { throw 'Resolve-OutlookFolder: -FolderPath is empty after parsing.' }

        # First part is the store / root folder name
        $folder = $null
        foreach ($store in $ns.Folders) {
            if ($store.Name -eq $parts[0]) {
                $folder = $store
                break
            }
        }
        if ($null -eq $folder) {
            throw "Resolve-OutlookFolder: root folder '$($parts[0])' not found."
        }

        for ($i = 1; $i -lt $parts.Count; $i++) {
            $child = $null
            foreach ($sub in $folder.Folders) {
                if ($sub.Name -eq $parts[$i]) {
                    $child = $sub
                    break
                }
            }
            if ($null -eq $child) {
                throw "Resolve-OutlookFolder: subfolder '$($parts[$i])' not found under '$($folder.FolderPath)'."
            }
            $folder = $child
        }
        return $folder
    }

    if (-not [string]::IsNullOrWhiteSpace($FolderType)) {
        $enumVal = Resolve-EnumValue -Map $script:OL_DEFAULT_FOLDER -Key $FolderType -EnumName 'FolderType'
        return $ns.GetDefaultFolder($enumVal)
    }

    throw 'Resolve-OutlookFolder: supply either -FolderType or -FolderPath.'
}

# ── Public Functions ─────────────────────────────────────────────────────────

function Get-OutlookDefaultFolder {
    <#
    .SYNOPSIS
        Returns metadata for a default Outlook folder (Inbox, Calendar, etc.).
    .PARAMETER FolderType
        Default folder type name (e.g. inbox, calendar, contacts).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookDefaultFolder -FolderType inbox
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('deletedItems','outbox','sentMail','inbox','calendar','contacts',
                     'journal','notes','tasks','drafts','conflicts','syncIssues',
                     'localFailures','serverFailures','junk','rssFeeds','toDo',
                     'managedEmail','suggestedContacts')]
        [string]$FolderType,

        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderType)) { throw 'Get-OutlookDefaultFolder: -FolderType is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $enumVal = Resolve-EnumValue -Map $script:OL_DEFAULT_FOLDER -Key $FolderType -EnumName 'FolderType'
    $folder  = $ns.GetDefaultFolder($enumVal)

    $result = @{
        name            = ConvertTo-OutlookSafeValue $folder.Name
        folderPath      = ConvertTo-OutlookSafeValue $folder.FolderPath
        itemCount       = ConvertTo-OutlookSafeValue $folder.Items.Count
        unreadCount     = ConvertTo-OutlookSafeValue $folder.UnReadItemCount
        defaultItemType = ConvertTo-OutlookSafeValue $folder.DefaultItemType
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookFolder {
    <#
    .SYNOPSIS
        Returns metadata for a folder identified by its full path.
    .PARAMETER FolderPath
        Full folder path (e.g. "\\Mailbox - User\Inbox\Subfolder").
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookFolder -FolderPath '\\Mailbox - User\Inbox\Projects'
    #>
    [CmdletBinding()]
    param(
        [string]$FolderPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { throw 'Get-OutlookFolder: -FolderPath is required.' }

    $app    = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderPath $FolderPath

    $result = @{
        name            = ConvertTo-OutlookSafeValue $folder.Name
        folderPath      = ConvertTo-OutlookSafeValue $folder.FolderPath
        itemCount       = ConvertTo-OutlookSafeValue $folder.Items.Count
        unreadCount     = ConvertTo-OutlookSafeValue $folder.UnReadItemCount
        defaultItemType = ConvertTo-OutlookSafeValue $folder.DefaultItemType
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookFolderList {
    <#
    .SYNOPSIS
        Lists subfolders of a specified folder.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookFolderList
    .EXAMPLE
        Get-OutlookFolderList -FolderPath '\\Mailbox - User\Inbox'
    #>
    [CmdletBinding()]
    param(
        [string]$FolderType = 'inbox',
        [string]$FolderPath,
        [switch]$AsJson
    )

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $list = @()
    foreach ($sub in $folder.Folders) {
        $list += @{
            name           = ConvertTo-OutlookSafeValue $sub.Name
            folderPath     = ConvertTo-OutlookSafeValue $sub.FolderPath
            itemCount      = ConvertTo-OutlookSafeValue $sub.Items.Count
            subfolderCount = ConvertTo-OutlookSafeValue $sub.Folders.Count
        }
    }

    if ($AsJson) {
        return ($list | ForEach-Object { [PSCustomObject]$_ }) | ConvertTo-Json -Depth 10 -Compress
    }
    return $list | ForEach-Object { [PSCustomObject]$_ }
}

function Get-OutlookFolderInfo {
    <#
    .SYNOPSIS
        Returns detailed metadata for a folder including StoreID and EntryID.
    .PARAMETER FolderType
        Default folder type name.
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookFolderInfo -FolderType inbox
    .EXAMPLE
        Get-OutlookFolderInfo -FolderPath '\\Mailbox - User\Inbox'
    #>
    [CmdletBinding()]
    param(
        [string]$FolderType,
        [string]$FolderPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderType) -and [string]::IsNullOrWhiteSpace($FolderPath)) {
        throw 'Get-OutlookFolderInfo: supply either -FolderType or -FolderPath.'
    }

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $result = @{
        name            = ConvertTo-OutlookSafeValue $folder.Name
        folderPath      = ConvertTo-OutlookSafeValue $folder.FolderPath
        itemCount       = ConvertTo-OutlookSafeValue $folder.Items.Count
        unreadCount     = ConvertTo-OutlookSafeValue $folder.UnReadItemCount
        defaultItemType = ConvertTo-OutlookSafeValue $folder.DefaultItemType
        storeID         = ConvertTo-OutlookSafeValue $folder.StoreID
        entryID         = ConvertTo-OutlookSafeValue $folder.EntryID
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookFolder {
    <#
    .SYNOPSIS
        Creates a new subfolder under a specified parent folder.
    .PARAMETER ParentFolderType
        Default folder type of the parent.
    .PARAMETER ParentFolderPath
        Full path of the parent folder. Takes precedence over -ParentFolderType.
    .PARAMETER Name
        Name for the new subfolder.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookFolder -ParentFolderType inbox -Name 'Projects'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$ParentFolderType,
        [string]$ParentFolderPath,
        [string]$Name,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'New-OutlookFolder: -Name is required.' }
    if ([string]::IsNullOrWhiteSpace($ParentFolderType) -and [string]::IsNullOrWhiteSpace($ParentFolderPath)) {
        throw 'New-OutlookFolder: supply either -ParentFolderType or -ParentFolderPath.'
    }

    $app = Connect-OutlookSession

    if (-not [string]::IsNullOrWhiteSpace($ParentFolderPath)) {
        $parent = Resolve-OutlookFolder -FolderPath $ParentFolderPath
    } else {
        $parent = Resolve-OutlookFolder -FolderType $ParentFolderType
    }

    if ($PSCmdlet.ShouldProcess("$($parent.FolderPath)\$Name", 'Create folder')) {
        $newFolder = $parent.Folders.Add($Name)

        $result = @{
            name       = ConvertTo-OutlookSafeValue $newFolder.Name
            folderPath = ConvertTo-OutlookSafeValue $newFolder.FolderPath
            entryID    = ConvertTo-OutlookSafeValue $newFolder.EntryID
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Rename-OutlookFolder {
    <#
    .SYNOPSIS
        Renames an Outlook folder.
    .PARAMETER FolderPath
        Full path of the folder to rename.
    .PARAMETER NewName
        New name for the folder.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Rename-OutlookFolder -FolderPath '\\Mailbox - User\Inbox\Old' -NewName 'New'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$FolderPath,
        [string]$NewName,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { throw 'Rename-OutlookFolder: -FolderPath is required.' }
    if ([string]::IsNullOrWhiteSpace($NewName))    { throw 'Rename-OutlookFolder: -NewName is required.' }

    $app    = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderPath $FolderPath

    if ($PSCmdlet.ShouldProcess($FolderPath, "Rename to '$NewName'")) {
        $folder.Name = $NewName

        $result = @{
            name       = ConvertTo-OutlookSafeValue $folder.Name
            folderPath = ConvertTo-OutlookSafeValue $folder.FolderPath
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Remove-OutlookFolder {
    <#
    .SYNOPSIS
        Deletes an Outlook folder and its contents.
    .PARAMETER FolderPath
        Full path of the folder to delete.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Remove-OutlookFolder -FolderPath '\\Mailbox - User\Inbox\Old'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [string]$FolderPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) { throw 'Remove-OutlookFolder: -FolderPath is required.' }

    $app    = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    $name   = $folder.Name

    if ($PSCmdlet.ShouldProcess($FolderPath, 'Delete folder')) {
        $folder.Delete()

        $result = @{ status = 'deleted'; folder = $name }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}

function Move-OutlookFolder {
    <#
    .SYNOPSIS
        Moves an Outlook folder to a new parent folder.
    .PARAMETER FolderPath
        Full path of the folder to move.
    .PARAMETER DestinationPath
        Full path of the destination parent folder.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Move-OutlookFolder -FolderPath '\\Mailbox\Inbox\Old' -DestinationPath '\\Mailbox\Archive'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$FolderPath,
        [string]$DestinationPath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath))      { throw 'Move-OutlookFolder: -FolderPath is required.' }
    if ([string]::IsNullOrWhiteSpace($DestinationPath))  { throw 'Move-OutlookFolder: -DestinationPath is required.' }

    $app    = Connect-OutlookSession
    $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    $dest   = Resolve-OutlookFolder -FolderPath $DestinationPath

    if ($PSCmdlet.ShouldProcess($FolderPath, "Move to '$DestinationPath'")) {
        $folder.MoveTo($dest)

        $result = @{
            name       = ConvertTo-OutlookSafeValue $folder.Name
            folderPath = ConvertTo-OutlookSafeValue $folder.FolderPath
        }

        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}
