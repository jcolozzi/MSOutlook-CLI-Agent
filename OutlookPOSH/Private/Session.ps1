# ── Private Session Management ───────────────────────────────────────────────

function Test-OutlookAlive {
    <#
    .SYNOPSIS
        Returns $true if the cached Outlook COM instance is still responsive.
    #>
    if ($null -eq $script:OutlookSession.App) { return $false }
    try {
        $null = $script:OutlookSession.App.Version
        return $true
    } catch {
        return $false
    }
}

function Connect-OutlookSession {
    <#
    .SYNOPSIS
        Attaches to the running Outlook instance or launches a new one.
        Returns the Application COM object.
    #>
    if (Test-OutlookAlive) {
        Write-Verbose 'OutlookPOSH: reusing existing Outlook session.'
        return $script:OutlookSession.App
    }

    Write-Verbose 'OutlookPOSH: creating Outlook.Application COM object...'
    $app = New-Object -ComObject Outlook.Application

    $script:OutlookSession.App       = $app
    $script:OutlookSession.Namespace = $app.GetNamespace('MAPI')

    Write-Verbose "OutlookPOSH: connected to Outlook $($app.Version)."
    return $app
}

function Disconnect-OutlookSession {
    <#
    .SYNOPSIS
        Releases COM references without calling Quit.
    #>
    if ($null -ne $script:OutlookSession.Namespace) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:OutlookSession.Namespace) | Out-Null
        } catch { }
        $script:OutlookSession.Namespace = $null
        Write-Verbose 'OutlookPOSH: released Namespace COM reference.'
    }

    if ($null -ne $script:OutlookSession.App) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:OutlookSession.App) | Out-Null
        } catch { }
        $script:OutlookSession.App = $null
        Write-Verbose 'OutlookPOSH: released Application COM reference.'
    }
}
