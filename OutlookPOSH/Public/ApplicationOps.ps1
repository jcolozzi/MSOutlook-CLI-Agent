# ── Application Operations ───────────────────────────────────────────────────

function Get-OutlookApplicationInfo {
    <#
    .SYNOPSIS
        Returns version, build, current user, and profile information for the running Outlook instance.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookApplicationInfo
    .EXAMPLE
        Get-OutlookApplicationInfo -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $currentUser = $ns.CurrentUser
    $result = @{
        name               = ConvertTo-OutlookSafeValue $app.Name
        version            = ConvertTo-OutlookSafeValue $app.Version
        build              = ConvertTo-OutlookSafeValue $app.Build
        defaultProfileName = ConvertTo-OutlookSafeValue $app.DefaultProfileName
        currentUserName    = ConvertTo-OutlookSafeValue $currentUser.Name
        currentUserAddress = ConvertTo-OutlookSafeValue $currentUser.Address
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookTip {
    <#
    .SYNOPSIS
        Returns a random tip about using OutlookPOSH effectively.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookTip
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $tips = @(
        'Use -AsJson on any command to get machine-readable JSON output for piping to other tools.'
        'Use EntryIDs to target specific items — they are stable identifiers across sessions.'
        'For calendar queries, always supply -Start and -End to scope the date range; Outlook recurrence expansion requires it.'
        'Configure Trust Center > Programmatic Access to avoid Outlook security prompts when using COM automation.'
        'Use DASL filter strings with Get-OutlookMail -Filter for fast server-side searches instead of client-side iteration.'
        'Call Close-OutlookSession when finished to release COM references without closing Outlook.'
        'Semicolons delimit multiple recipients: -To "alice@example.com;bob@example.com".'
        'Use Get-OutlookFolderList to discover subfolder names and paths before navigating into them.'
        'Set -MaxItems to control how many items are returned — the default is 50 to keep output manageable.'
        'Mark messages read/unread in bulk by piping EntryIDs to Set-OutlookMailRead.'
    )

    $tip = $tips | Get-Random

    $result = @{ tip = $tip }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Close-OutlookSession {
    <#
    .SYNOPSIS
        Releases COM references to Outlook without closing the application.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Close-OutlookSession
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [switch]$AsJson
    )

    if ($PSCmdlet.ShouldProcess('Outlook COM session', 'Disconnect')) {
        Disconnect-OutlookSession

        $result = @{ status = 'disconnected'; message = 'COM references released. Outlook remains running.' }
        Format-OutlookOutput -Data $result -AsJson:$AsJson
    }
}
