# ── Rule Operations ──────────────────────────────────────────────────────────

function Get-OutlookRule {
    <#
    .SYNOPSIS
        Lists all Outlook mail rules.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Get-OutlookRule
    .EXAMPLE
        Get-OutlookRule -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $rules = $ns.DefaultStore.GetRules()

    $list = @()
    foreach ($rule in $rules) {
        try {
            $list += [PSCustomObject]@{
                name           = ConvertTo-OutlookSafeValue $rule.Name
                enabled        = ConvertTo-OutlookSafeValue $rule.Enabled
                isLocalRule    = ConvertTo-OutlookSafeValue $rule.IsLocalRule
                ruleType       = ConvertTo-OutlookSafeValue $rule.RuleType
                executionOrder = ConvertTo-OutlookSafeValue $rule.ExecutionOrder
            }
        } catch {
            Write-Verbose "Skipped rule: $_"
        }
    }

    if ($AsJson) {
        return $list | ConvertTo-Json -Depth 10 -Compress
    }
    return $list
}

function New-OutlookRule {
    <#
    .SYNOPSIS
        Creates a simple mail receive rule.
    .PARAMETER Name
        Name for the new rule.
    .PARAMETER MoveToFolder
        Full folder path to move matching items to.
    .PARAMETER FromAddress
        Sender email address to match.
    .PARAMETER SubjectContains
        Text that must appear in the subject.
    .PARAMETER Enabled
        Whether the rule is enabled (default $true).
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        New-OutlookRule -Name 'Archive newsletters' -FromAddress 'news@example.com' -MoveToFolder '\\Mailbox\Archive'
    .EXAMPLE
        New-OutlookRule -Name 'Flag budget' -SubjectContains 'Budget' -Enabled $true
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name,
        [string]$MoveToFolder,
        [string]$FromAddress,
        [string]$SubjectContains,
        [bool]$Enabled = $true,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'New-OutlookRule: -Name is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $rules = $ns.DefaultStore.GetRules()

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create rule')) { return }

    $rule = $rules.Create($Name, 0)  # 0 = olRuleReceive
    $rule.Enabled = $Enabled

    # ── Conditions ───────────────────────────────────────────────────────
    if (-not [string]::IsNullOrWhiteSpace($FromAddress)) {
        $fromCondition = $rule.Conditions.From
        $fromCondition.Enabled = $true
        $fromCondition.Recipients.Add($FromAddress)
        $fromCondition.Recipients.ResolveAll()
    }

    if (-not [string]::IsNullOrWhiteSpace($SubjectContains)) {
        $subjCondition = $rule.Conditions.Subject
        $subjCondition.Enabled = $true
        $subjCondition.Text    = @($SubjectContains)
    }

    # ── Actions ──────────────────────────────────────────────────────────
    if (-not [string]::IsNullOrWhiteSpace($MoveToFolder)) {
        $targetFolder = Resolve-OutlookFolder -FolderPath $MoveToFolder
        $moveAction   = $rule.Actions.MoveToFolder
        $moveAction.Enabled = $true
        $moveAction.Folder  = $targetFolder
    }

    $rules.Save()

    $result = @{
        name    = ConvertTo-OutlookSafeValue $rule.Name
        enabled = ConvertTo-OutlookSafeValue $rule.Enabled
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Remove-OutlookRule {
    <#
    .SYNOPSIS
        Deletes a mail rule by name.
    .PARAMETER Name
        Name of the rule to remove.
    .EXAMPLE
        Remove-OutlookRule -Name 'Archive newsletters'
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'Remove-OutlookRule: -Name is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $rules = $ns.DefaultStore.GetRules()

    $found = $false
    for ($i = 1; $i -le $rules.Count; $i++) {
        if ($rules.Item($i).Name -eq $Name) {
            if (-not $PSCmdlet.ShouldProcess($Name, 'Remove rule')) { return }
            $rules.Remove($i)
            $found = $true
            break
        }
    }

    if (-not $found) { throw "Remove-OutlookRule: rule '$Name' not found." }

    $rules.Save()
    Write-Verbose "Removed rule '$Name'."
}

function Set-OutlookRuleEnabled {
    <#
    .SYNOPSIS
        Enables or disables an existing mail rule.
    .PARAMETER Name
        Name of the rule.
    .PARAMETER Enabled
        $true to enable, $false to disable.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Set-OutlookRuleEnabled -Name 'Archive newsletters' -Enabled $false
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$Name,
        [bool]$Enabled,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { throw 'Set-OutlookRuleEnabled: -Name is required.' }

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    $rules = $ns.DefaultStore.GetRules()

    $target = $null
    foreach ($rule in $rules) {
        if ($rule.Name -eq $Name) {
            $target = $rule
            break
        }
    }

    if ($null -eq $target) { throw "Set-OutlookRuleEnabled: rule '$Name' not found." }

    $action = if ($Enabled) { 'Enable' } else { 'Disable' }
    if (-not $PSCmdlet.ShouldProcess($Name, "$action rule")) { return }

    $target.Enabled = $Enabled
    $rules.Save()

    $result = @{
        name    = ConvertTo-OutlookSafeValue $target.Name
        enabled = ConvertTo-OutlookSafeValue $target.Enabled
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}

function Invoke-OutlookRules {
    <#
    .SYNOPSIS
        Executes rules on a folder. Optionally executes only a specific rule.
    .PARAMETER FolderType
        Default folder type (defaults to 'inbox').
    .PARAMETER FolderPath
        Full folder path. Takes precedence over -FolderType.
    .PARAMETER RuleName
        Optional: execute only the named rule. If omitted, all enabled rules are executed.
    .PARAMETER AsJson
        Emit output as a compressed JSON string instead of a PSCustomObject.
    .EXAMPLE
        Invoke-OutlookRules
    .EXAMPLE
        Invoke-OutlookRules -RuleName 'Archive newsletters' -FolderType inbox
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [string]$FolderType = 'inbox',
        [string]$FolderPath,
        [string]$RuleName,
        [switch]$AsJson
    )

    $app = Connect-OutlookSession
    $ns  = $script:OutlookSession.Namespace

    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $folder = Resolve-OutlookFolder -FolderPath $FolderPath
    } else {
        $folder = Resolve-OutlookFolder -FolderType $FolderType
    }

    $rules      = $ns.DefaultStore.GetRules()
    $folderName = $folder.Name
    $executed   = @()

    if (-not [string]::IsNullOrWhiteSpace($RuleName)) {
        # Execute specific rule
        $target = $null
        foreach ($rule in $rules) {
            if ($rule.Name -eq $RuleName) {
                $target = $rule
                break
            }
        }
        if ($null -eq $target) { throw "Invoke-OutlookRules: rule '$RuleName' not found." }

        if (-not $PSCmdlet.ShouldProcess("'$RuleName' on $folderName", 'Execute rule')) { return }
        $target.Execute($false, $folder)
        $executed += $RuleName
    } else {
        # Execute all enabled rules
        if (-not $PSCmdlet.ShouldProcess("All enabled rules on $folderName", 'Execute rules')) { return }
        foreach ($rule in $rules) {
            if ($rule.Enabled) {
                try {
                    $rule.Execute($false, $folder)
                    $executed += $rule.Name
                } catch {
                    Write-Verbose "Failed to execute rule '$($rule.Name)': $_"
                }
            }
        }
    }

    $result = @{
        folder        = $folderName
        executedRules = $executed
        executedCount = $executed.Count
    }
    Format-OutlookOutput -Data $result -AsJson:$AsJson
}
