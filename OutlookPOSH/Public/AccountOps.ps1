# ── Account Operations ───────────────────────────────────────────────────────

function Get-OutlookAccount {
    <#
    .SYNOPSIS
        Lists all configured Outlook accounts.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookAccount
    .EXAMPLE
        Get-OutlookAccount -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $list = @()
    foreach ($acct in $ns.Accounts) {
        try {
            $list += [PSCustomObject]@{
                displayName = ConvertTo-OutlookSafeValue $acct.DisplayName
                smtpAddress = ConvertTo-OutlookSafeValue $acct.SmtpAddress
                accountType = ConvertTo-OutlookSafeValue $acct.AccountType
                userName    = ConvertTo-OutlookSafeValue $acct.UserName
            }
        } catch {
            Write-Verbose "Skipped account: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Get-OutlookDefaultAccount {
    <#
    .SYNOPSIS
        Gets the default sending account (first account).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookDefaultAccount
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession

    $acct = $app.Session.Accounts.Item(1)

    $result = @{
        displayName = ConvertTo-OutlookSafeValue $acct.DisplayName
        smtpAddress = ConvertTo-OutlookSafeValue $acct.SmtpAddress
        accountType = ConvertTo-OutlookSafeValue $acct.AccountType
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Get-OutlookStore {
    <#
    .SYNOPSIS
        Lists all data stores (PST, OST, Exchange).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookStore
    .EXAMPLE
        Get-OutlookStore -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $list = @()
    foreach ($store in $ns.Stores) {
        try {
            $list += [PSCustomObject]@{
                displayName     = ConvertTo-OutlookSafeValue $store.DisplayName
                filePath        = ConvertTo-OutlookSafeValue $store.FilePath
                storeKind       = ConvertTo-OutlookSafeValue $store.ExchangeStoreType
                isDataFileStore = ConvertTo-OutlookSafeValue $store.IsDataFileStore
                storeID         = ConvertTo-OutlookSafeValue $store.StoreID
            }
        } catch {
            Write-Verbose "Skipped store: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function Get-OutlookCurrentUser {
    <#
    .SYNOPSIS
        Gets the current MAPI user profile.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookCurrentUser
    .EXAMPLE
        Get-OutlookCurrentUser -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $user = $ns.CurrentUser

    $result = @{
        name    = ConvertTo-OutlookSafeValue $user.Name
        address = ConvertTo-OutlookSafeValue $user.Address
        entryID = ConvertTo-OutlookSafeValue $user.EntryID
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
