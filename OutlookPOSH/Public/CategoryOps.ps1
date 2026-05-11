# ── Category Operations ──────────────────────────────────────────────────────

function Get-OutlookCategory {
    <#
    .SYNOPSIS
        Lists all master categories defined in Outlook.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookCategory
    .EXAMPLE
        Get-OutlookCategory -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    # Build reverse lookup: int → friendly name
    $colorReverse = @{}
    foreach ($kv in $script:OL_CATEGORY_COLOR.GetEnumerator()) {
        $colorReverse[$kv.Value] = $kv.Key
    }

    $list = @()
    foreach ($cat in $ns.Categories) {
        try {
            $colorVal  = ConvertTo-OutlookSafeValue $cat.Color
            $colorName = if ($colorReverse.ContainsKey([int]$colorVal)) { $colorReverse[[int]$colorVal] } else { "unknown($colorVal)" }
            $list += [PSCustomObject]@{
                name        = ConvertTo-OutlookSafeValue $cat.Name
                color       = $colorName
                colorValue  = $colorVal
                shortcutKey = ConvertTo-OutlookSafeValue $cat.ShortcutKey
            }
        } catch {
            Write-Verbose "Skipped category: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function New-OutlookCategory {
    <#
    .SYNOPSIS
        Creates a new master category.
    .PARAMETER Name
        Name of the category to create.
    .PARAMETER Color
        Color name (e.g. red, blue, green). Resolved via OL_CATEGORY_COLOR.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookCategory -Name 'Project Alpha' -Color blue
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name,
        [string]$Color = 'none',
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'New-OutlookCategory: -Name is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $colorValue = Resolve-EnumValue -Map $script:OL_CATEGORY_COLOR -Key $Color -EnumName 'CategoryColor'

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create category')) { return }

    $cat = $ns.Categories.Add($Name, $colorValue)

    $result = @{
        name       = ConvertTo-OutlookSafeValue $cat.Name
        color      = $Color
        colorValue = $colorValue
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookCategory {
    <#
    .SYNOPSIS
        Deletes a master category by name.
    .PARAMETER Name
        Name of the category to remove.
    .EXAMPLE
        Remove-OutlookCategory -Name 'Project Alpha'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'Remove-OutlookCategory: -Name is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    # Find the category index (1-based)
    $found = $false
    for ($i = 1; $i -le $ns.Categories.Count; $i++) {
        if ($ns.Categories.Item($i).Name -eq $Name) {
            if (-not $PSCmdlet.ShouldProcess($Name, 'Remove category')) { return }
            $ns.Categories.Remove($i)
            $found = $true
            break
        }
    }

    if (-not $found) { throw "Remove-OutlookCategory: category '$Name' not found." }
    Write-Verbose "Removed category '$Name'."
}

function Set-OutlookItemCategory {
    <#
    .SYNOPSIS
        Sets the categories on an Outlook item.
    .PARAMETER EntryID
        The EntryID of the item.
    .PARAMETER Categories
        Comma-separated category names (e.g. 'Red,Blue').
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookItemCategory -EntryID '000000...' -Categories 'Project Alpha, Urgent'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$Categories,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID))    { throw 'Set-OutlookItemCategory: -EntryID is required.' }
    if ([string]::IsNullOrWhiteSpace($Categories))  { throw 'Set-OutlookItemCategory: -Categories is required.' }

    $app  = Connect-OutlookSession
    $ns   = $script:OutlookSession.Namespace
    $item = $ns.GetItemFromID($EntryID)

    $subj = ConvertTo-OutlookSafeValue $item.Subject
    if (-not $PSCmdlet.ShouldProcess($subj, "Set categories to '$Categories'")) { return }

    $item.Categories = $Categories
    $item.Save()

    $result = @{
        entryID    = ConvertTo-OutlookSafeValue $item.EntryID
        subject    = $subj
        categories = ConvertTo-OutlookSafeValue $item.Categories
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookCategoryColor {
    <#
    .SYNOPSIS
        Lists all available category color names and their numeric values.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookCategoryColor
    .EXAMPLE
        Get-OutlookCategoryColor -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $list = @()
    foreach ($kv in ($script:OL_CATEGORY_COLOR.GetEnumerator() | Sort-Object Value)) {
        $list += [PSCustomObject]@{
            colorName  = $kv.Key
            colorValue = $kv.Value
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}
