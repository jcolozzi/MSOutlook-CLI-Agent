# ── ContactOps.ps1 ──────────────────────────────────────────────────────────
# Contact and Distribution List operations for OutlookPOSH
# ────────────────────────────────────────────────────────────────────────────

function Get-OutlookContact {
<#
.SYNOPSIS
    Lists contacts from the default Contacts folder.
.PARAMETER FolderPath
    Optional folder path. If omitted, uses the default Contacts folder.
.PARAMETER Filter
    Optional DASL filter string to restrict results.
.PARAMETER MaxItems
    Maximum number of contacts to return. Default 50.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookContact -MaxItems 10
.EXAMPLE
    Get-OutlookContact -Filter "[CompanyName] = 'Contoso'"
#>
    [CmdletBinding()]
    param(
        [string]$FolderPath,
        [string]$Filter,
        [int]$MaxItems = 50,
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    # olFolderContacts = 10
    $folder = $ns.GetDefaultFolder(10)

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = $ns.GetFolderFromPath($FolderPath)
    }

    $items = $folder.Items
    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $items = $items.Restrict($Filter)
    }

    $results = [System.Collections.Generic.List[hashtable]]::new()
    $count   = 0

    foreach ($c in $items) {
        if ($count -ge $MaxItems) { break }
        try {
            if ($c.Class -ne 40) { continue }   # olContact = 40
            $results.Add(@{
                entryID       = ConvertTo-OutlookSafeValue $c.EntryID
                fullName      = ConvertTo-OutlookSafeValue $c.FullName
                firstName     = ConvertTo-OutlookSafeValue $c.FirstName
                lastName      = ConvertTo-OutlookSafeValue $c.LastName
                companyName   = ConvertTo-OutlookSafeValue $c.CompanyName
                jobTitle      = ConvertTo-OutlookSafeValue $c.JobTitle
                email1        = ConvertTo-OutlookSafeValue $c.Email1Address
                email2        = ConvertTo-OutlookSafeValue $c.Email2Address
                email3        = ConvertTo-OutlookSafeValue $c.Email3Address
                businessPhone = ConvertTo-OutlookSafeValue $c.BusinessTelephoneNumber
                mobilePhone   = ConvertTo-OutlookSafeValue $c.MobileTelephoneNumber
                homePhone     = ConvertTo-OutlookSafeValue $c.HomeTelephoneNumber
                categories    = ConvertTo-OutlookSafeValue $c.Categories
            })
            $count++
        } catch { }
    }

    Format-OutlookOutput -Data $results -AsJson:$AsJson
}

function Get-OutlookContactItem {
<#
.SYNOPSIS
    Returns full details for a single contact by EntryID.
.PARAMETER EntryID
    The EntryID of the contact item.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookContactItem -EntryID '00000000AB...'
#>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Get-OutlookContactItem: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $c   = $ns.GetItemFromID($EntryID)

    $result = @{
        entryID             = ConvertTo-OutlookSafeValue $c.EntryID
        fullName            = ConvertTo-OutlookSafeValue $c.FullName
        firstName           = ConvertTo-OutlookSafeValue $c.FirstName
        lastName            = ConvertTo-OutlookSafeValue $c.LastName
        middleName          = ConvertTo-OutlookSafeValue $c.MiddleName
        companyName         = ConvertTo-OutlookSafeValue $c.CompanyName
        department          = ConvertTo-OutlookSafeValue $c.Department
        jobTitle            = ConvertTo-OutlookSafeValue $c.JobTitle
        fileAs              = ConvertTo-OutlookSafeValue $c.FileAs
        email1Address       = ConvertTo-OutlookSafeValue $c.Email1Address
        email1DisplayName   = ConvertTo-OutlookSafeValue $c.Email1DisplayName
        email2Address       = ConvertTo-OutlookSafeValue $c.Email2Address
        email3Address       = ConvertTo-OutlookSafeValue $c.Email3Address
        businessPhone       = ConvertTo-OutlookSafeValue $c.BusinessTelephoneNumber
        homePhone           = ConvertTo-OutlookSafeValue $c.HomeTelephoneNumber
        mobilePhone         = ConvertTo-OutlookSafeValue $c.MobileTelephoneNumber
        businessFax         = ConvertTo-OutlookSafeValue $c.BusinessFaxNumber
        businessAddress     = @{
            street     = ConvertTo-OutlookSafeValue $c.BusinessAddressStreet
            city       = ConvertTo-OutlookSafeValue $c.BusinessAddressCity
            state      = ConvertTo-OutlookSafeValue $c.BusinessAddressState
            postalCode = ConvertTo-OutlookSafeValue $c.BusinessAddressPostalCode
            country    = ConvertTo-OutlookSafeValue $c.BusinessAddressCountry
        }
        homeAddress         = ConvertTo-OutlookSafeValue $c.HomeAddress
        mailingAddress      = ConvertTo-OutlookSafeValue $c.MailingAddress
        birthday            = ConvertTo-OutlookSafeValue $c.Birthday
        categories          = ConvertTo-OutlookSafeValue $c.Categories
        importance          = ConvertTo-OutlookSafeValue $c.Importance
        body                = ConvertTo-OutlookSafeValue $c.Body
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookContact {
<#
.SYNOPSIS
    Creates a new contact in the default Contacts folder.
.PARAMETER FirstName
    Contact first name.
.PARAMETER LastName
    Contact last name.
.PARAMETER Email
    Primary email address.
.PARAMETER CompanyName
    Company / organisation name.
.PARAMETER JobTitle
    Job title.
.PARAMETER BusinessPhone
    Business telephone number.
.PARAMETER MobilePhone
    Mobile telephone number.
.PARAMETER HomePhone
    Home telephone number.
.PARAMETER Categories
    Comma-separated categories string.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    New-OutlookContact -FirstName 'Jane' -LastName 'Doe' -Email 'jane@example.com'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Email,
        [string]$CompanyName,
        [string]$JobTitle,
        [string]$BusinessPhone,
        [string]$MobilePhone,
        [string]$HomePhone,
        [string]$Categories,
        [switch]$AsJson
    )

    $displayName = ($FirstName, $LastName | Where-Object { $_ }) -join ' '
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        throw "New-OutlookContact: at least -FirstName or -LastName is required."
    }

    if (-not $PSCmdlet.ShouldProcess($displayName, 'Create contact')) { return }

    $app = Connect-OutlookSession
    # olContactItem = 2
    $c = $app.CreateItem(2)

    if (-not [string]::IsNullOrWhiteSpace($FirstName))     { $c.FirstName                = $FirstName }
    if (-not [string]::IsNullOrWhiteSpace($LastName))      { $c.LastName                 = $LastName }
    if (-not [string]::IsNullOrWhiteSpace($Email))         { $c.Email1Address            = $Email }
    if (-not [string]::IsNullOrWhiteSpace($CompanyName))   { $c.CompanyName              = $CompanyName }
    if (-not [string]::IsNullOrWhiteSpace($JobTitle))      { $c.JobTitle                 = $JobTitle }
    if (-not [string]::IsNullOrWhiteSpace($BusinessPhone)) { $c.BusinessTelephoneNumber  = $BusinessPhone }
    if (-not [string]::IsNullOrWhiteSpace($MobilePhone))   { $c.MobileTelephoneNumber    = $MobilePhone }
    if (-not [string]::IsNullOrWhiteSpace($HomePhone))     { $c.HomeTelephoneNumber      = $HomePhone }
    if (-not [string]::IsNullOrWhiteSpace($Categories))    { $c.Categories               = $Categories }

    $c.Save()

    $result = @{
        entryID  = ConvertTo-OutlookSafeValue $c.EntryID
        fullName = ConvertTo-OutlookSafeValue $c.FullName
        status   = 'Created'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Set-OutlookContact {
<#
.SYNOPSIS
    Updates an existing contact. Only provided (non-empty) fields are changed.
.PARAMETER EntryID
    The EntryID of the contact to update.
.PARAMETER FirstName
    New first name.
.PARAMETER LastName
    New last name.
.PARAMETER Email
    New primary email address.
.PARAMETER CompanyName
    New company name.
.PARAMETER JobTitle
    New job title.
.PARAMETER BusinessPhone
    New business phone.
.PARAMETER MobilePhone
    New mobile phone.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Set-OutlookContact -EntryID '00000...' -JobTitle 'Senior Dev'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [string]$FirstName,
        [string]$LastName,
        [string]$Email,
        [string]$CompanyName,
        [string]$JobTitle,
        [string]$BusinessPhone,
        [string]$MobilePhone,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Set-OutlookContact: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $c   = $ns.GetItemFromID($EntryID)

    if (-not $PSCmdlet.ShouldProcess($c.FullName, 'Update contact')) { return }

    if (-not [string]::IsNullOrWhiteSpace($FirstName))     { $c.FirstName                = $FirstName }
    if (-not [string]::IsNullOrWhiteSpace($LastName))      { $c.LastName                 = $LastName }
    if (-not [string]::IsNullOrWhiteSpace($Email))         { $c.Email1Address            = $Email }
    if (-not [string]::IsNullOrWhiteSpace($CompanyName))   { $c.CompanyName              = $CompanyName }
    if (-not [string]::IsNullOrWhiteSpace($JobTitle))      { $c.JobTitle                 = $JobTitle }
    if (-not [string]::IsNullOrWhiteSpace($BusinessPhone)) { $c.BusinessTelephoneNumber  = $BusinessPhone }
    if (-not [string]::IsNullOrWhiteSpace($MobilePhone))   { $c.MobileTelephoneNumber    = $MobilePhone }

    $c.Save()

    $result = @{
        entryID  = ConvertTo-OutlookSafeValue $c.EntryID
        fullName = ConvertTo-OutlookSafeValue $c.FullName
        status   = 'Updated'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookContact {
<#
.SYNOPSIS
    Deletes a contact by EntryID.
.PARAMETER EntryID
    The EntryID of the contact to delete.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Remove-OutlookContact -EntryID '00000...'
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Remove-OutlookContact: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $c   = $ns.GetItemFromID($EntryID)

    $name = $c.FullName
    if (-not $PSCmdlet.ShouldProcess($name, 'Delete contact')) { return }

    $c.Delete()

    $result = @{ fullName = $name; status = 'Deleted' }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Find-OutlookContact {
<#
.SYNOPSIS
    Convenience search for contacts by name or email using Jet filter.
.PARAMETER Name
    Searches the FullName field (contains match).
.PARAMETER Email
    Searches the Email1Address field (contains match).
.PARAMETER MaxItems
    Maximum results to return. Default 20.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Find-OutlookContact -Name 'Smith' -MaxItems 5
.EXAMPLE
    Find-OutlookContact -Email 'contoso.com'
#>
    [CmdletBinding()]
    param(
        [string]$Name,
        [string]$Email,
        [int]$MaxItems = 20,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name) -and [string]::IsNullOrWhiteSpace($Email)) {
        throw "Find-OutlookContact: at least -Name or -Email is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $folder = $ns.GetDefaultFolder(10)
    $items  = $folder.Items

    # Build Jet-style filter
    $filters = @()
    if (-not [string]::IsNullOrWhiteSpace($Name))  {
        $filters += "[FullName] ci_phrasematch '$($Name.Replace("'","''"))'"
    }
    if (-not [string]::IsNullOrWhiteSpace($Email)) {
        $filters += "[Email1Address] ci_phrasematch '$($Email.Replace("'","''"))'"
    }

    $filterStr = $filters -join ' AND '
    $items = $items.Restrict($filterStr)

    $results = [System.Collections.Generic.List[hashtable]]::new()
    $count   = 0

    foreach ($c in $items) {
        if ($count -ge $MaxItems) { break }
        try {
            if ($c.Class -ne 40) { continue }
            $results.Add(@{
                entryID       = ConvertTo-OutlookSafeValue $c.EntryID
                fullName      = ConvertTo-OutlookSafeValue $c.FullName
                firstName     = ConvertTo-OutlookSafeValue $c.FirstName
                lastName      = ConvertTo-OutlookSafeValue $c.LastName
                companyName   = ConvertTo-OutlookSafeValue $c.CompanyName
                jobTitle      = ConvertTo-OutlookSafeValue $c.JobTitle
                email1        = ConvertTo-OutlookSafeValue $c.Email1Address
                email2        = ConvertTo-OutlookSafeValue $c.Email2Address
                email3        = ConvertTo-OutlookSafeValue $c.Email3Address
                businessPhone = ConvertTo-OutlookSafeValue $c.BusinessTelephoneNumber
                mobilePhone   = ConvertTo-OutlookSafeValue $c.MobileTelephoneNumber
                homePhone     = ConvertTo-OutlookSafeValue $c.HomeTelephoneNumber
                categories    = ConvertTo-OutlookSafeValue $c.Categories
            })
            $count++
        } catch { }
    }

    Format-OutlookOutput -Data $results -AsJson:$AsJson
}

function Get-OutlookDistributionList {
<#
.SYNOPSIS
    Returns distribution list info including members.
.PARAMETER EntryID
    The EntryID of the distribution list.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    Get-OutlookDistributionList -EntryID '00000...'
#>
    [CmdletBinding()]
    param(
        [string]$EntryID,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($EntryID)) {
        throw "Get-OutlookDistributionList: -EntryID is required."
    }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace
    $dl  = $ns.GetItemFromID($EntryID)

    $members = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = 1; $i -le $dl.MemberCount; $i++) {
        $member = $dl.GetMember($i)
        $members.Add(@{
            name    = ConvertTo-OutlookSafeValue $member.Name
            address = ConvertTo-OutlookSafeValue $member.Address
        })
    }

    $result = @{
        dlName      = ConvertTo-OutlookSafeValue $dl.DLName
        memberCount = $dl.MemberCount
        members     = $members
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function New-OutlookDistributionList {
<#
.SYNOPSIS
    Creates a new distribution list with the specified members.
.PARAMETER Name
    Display name for the distribution list.
.PARAMETER Members
    Array of email addresses to add as members.
.PARAMETER AsJson
    Return output as a JSON string.
.EXAMPLE
    New-OutlookDistributionList -Name 'Team Alpha' -Members @('a@example.com','b@example.com')
#>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name,
        [string[]]$Members,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw "New-OutlookDistributionList: -Name is required."
    }

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create distribution list')) { return }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    # olDistributionListItem = 7
    $dl = $app.CreateItem(7)
    $dl.DLName = $Name

    if ($Members -and $Members.Count -gt 0) {
        foreach ($addr in $Members) {
            $recip = $ns.CreateRecipient($addr)
            $recip.Resolve() | Out-Null
            $dl.AddMember($recip)
        }
    }

    $dl.Save()

    $result = @{
        entryID     = ConvertTo-OutlookSafeValue $dl.EntryID
        dlName      = ConvertTo-OutlookSafeValue $dl.DLName
        memberCount = $dl.MemberCount
        status      = 'Created'
    }

    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
