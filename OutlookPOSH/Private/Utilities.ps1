# ── Private Utility Functions ────────────────────────────────────────────────

function ConvertTo-OutlookSafeValue {
    <#
    .SYNOPSIS
        Converts COM values to JSON-safe types.
    #>
    param($Value)

    if ($null -eq $Value)                        { return $null }
    if ($Value -is [System.DBNull])              { return $null }
    if ($Value -is [datetime])                   { return $Value.ToString('o') }
    if ($Value -is [decimal])                    { return [double]$Value }
    if ($Value -is [byte[]])                     { return "<binary $($Value.Length) bytes>" }
    return $Value
}

function Format-OutlookOutput {
    <#
    .SYNOPSIS
        Wraps a hashtable as PSCustomObject, optionally serialising to JSON.
    #>
    param(
        [hashtable]$Data,
        [switch]$AsJson
    )

    $obj = [PSCustomObject]$Data
    if ($AsJson) {
        return $obj | ConvertTo-Json -Depth 10 -Compress
    }
    return $obj
}

function Resolve-EnumValue {
    <#
    .SYNOPSIS
        Maps a friendly name to its integer enum value, or passes through a numeric string.
    #>
    param(
        [hashtable]$Map,
        [string]$Key,
        [string]$EnumName
    )

    # Numeric passthrough
    if ($Key -match '^\d+$') { return [int]$Key }

    $lower = $Key.ToLower()
    if ($Map.ContainsKey($lower)) { return $Map[$lower] }

    $valid = ($Map.Keys | Sort-Object) -join ', '
    throw "Invalid ${EnumName} value '${Key}'. Valid values: ${valid}"
}

function Build-DASLFilter {
    <#
    .SYNOPSIS
        Constructs Outlook DASL / Jet filter strings for Restrict() or Find().
    #>
    param(
        [string]$Property,
        $Value,
        [string]$Operator = 'equals',
        [string[]]$And,
        [string[]]$Or
    )

    $clause = switch ($Operator) {
        'equals'   { _Build-SingleClause -Property $Property -Value $Value -Op '=' }
        'contains' {
            if ($Value -is [string]) {
                "@SQL=`"urn:schemas:httpmail:$Property`" LIKE '%$($Value.Replace("'","''"))%'"
            } else {
                _Build-SingleClause -Property $Property -Value $Value -Op '='
            }
        }
        'gt'  { _Build-SingleClause -Property $Property -Value $Value -Op '>' }
        'lt'  { _Build-SingleClause -Property $Property -Value $Value -Op '<' }
        'gte' { _Build-SingleClause -Property $Property -Value $Value -Op '>=' }
        'lte' { _Build-SingleClause -Property $Property -Value $Value -Op '<=' }
        default { throw "Unknown operator '$Operator'. Use: equals, contains, gt, lt, gte, lte" }
    }

    if ($And) {
        foreach ($extra in $And) { $clause = "$clause AND $extra" }
    }
    if ($Or) {
        foreach ($extra in $Or) { $clause = "($clause) OR ($extra)" }
    }

    return $clause
}

# ── Helper for Build-DASLFilter ──────────────────────────────────────────────

function _Build-SingleClause {
    param(
        [string]$Property,
        $Value,
        [string]$Op
    )

    if ($Value -is [datetime]) {
        $formatted = $Value.ToString('MM/dd/yyyy HH:mm')
        return "[${Property}] ${Op} '${formatted}'"
    }
    if ($Value -is [bool]) {
        $boolStr = if ($Value) { 'True' } else { 'False' }
        return "[${Property}] ${Op} ${boolStr}"
    }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
        return "[${Property}] ${Op} ${Value}"
    }
    # String — wrap in single quotes
    $escaped = "$Value".Replace("'", "''")
    return "[${Property}] ${Op} '${escaped}'"
}
