# Autodoc v1.0
function Read-AutodocJson {
    <#
        Reads a JSON file the same way in Windows PowerShell 5.1 and PowerShell 7.

        - Encoding: the BOM is detected (UTF-8, or UTF-16 from older 5.1 exports);
          without a BOM the file is read as UTF-8, so å/ä/ö survive in 5.1.
        - Arrays: in 5.1, "Get-Content | ConvertFrom-Json" passes a JSON array on
          as ONE object. Here an array is always returned element by element,
          so @(Read-AutodocJson ...) gives the rows in both versions.
        - -AsHashtable: objects become ordered hashtables (ConvertFrom-Json
          -AsHashtable does not exist in 5.1). Key order is kept.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$AsHashtable
    )

    $full = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $text = [System.IO.File]::ReadAllText($full, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    $parsed = ConvertFrom-Json -InputObject $text
    if ($AsHashtable) { $parsed = ConvertTo-AutodocHashtable -InputObject $parsed }
    if ($parsed -is [System.Array]) { foreach ($item in $parsed) { $item } }
    else { $parsed }
}

function ConvertTo-AutodocHashtable {
    <#
        Turns the objects from ConvertFrom-Json into ordered hashtables,
        recursively (arrays stay arrays). Used instead of -AsHashtable so the
        same code runs in Windows PowerShell 5.1.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-AutodocHashtable -InputObject $p.Value }
        return $h
    }
    if ($InputObject -is [System.Array]) {
        $list = @(foreach ($item in $InputObject) { , (ConvertTo-AutodocHashtable -InputObject $item) })
        return , $list
    }
    return $InputObject
}

function ConvertTo-AutodocJsonSafe {
    <#
        Windows PowerShell 5.1 writes dates as "\/Date(1700000000000)\/" in
        ConvertTo-Json. This copies the data with every date turned into ISO 8601
        text (the format PowerShell 7 writes), so the JSON files look the same
        whichever version made them. In PowerShell 7 the data is returned as is.
    #>
    [CmdletBinding()]
    param([AllowNull()]$InputObject, [int]$Depth = 15)

    if ($PSVersionTable.PSEdition -ne 'Desktop') { return , $InputObject }
    $iso = 'yyyy-MM-ddTHH:mm:ss.FFFFFFFK'
    $convert = $null
    $convert = {
        param($v, [int]$d)
        if ($null -eq $v) { return $null }
        if ($v -is [datetime])       { return $v.ToString($iso, [System.Globalization.CultureInfo]::InvariantCulture) }
        if ($v -is [datetimeoffset]) { return $v.ToString('yyyy-MM-ddTHH:mm:ss.FFFFFFFzzz', [System.Globalization.CultureInfo]::InvariantCulture) }
        if ($d -le 0 -or $v -is [string] -or $v -is [ValueType]) { return $v }
        if ($v -is [System.Collections.IDictionary]) {
            $h = [ordered]@{}
            foreach ($k in $v.Keys) { $h[$k] = & $convert $v[$k] ($d - 1) }
            return $h
        }
        if ($v -is [System.Collections.IEnumerable]) {
            return , @(foreach ($i in $v) { , (& $convert $i ($d - 1)) })
        }
        $props = @($v.PSObject.Properties | Where-Object { $_.MemberType -match 'Property' })
        if (-not $props.Count) { return $v }
        $o = [ordered]@{}
        foreach ($p in $props) {
            try { $o[$p.Name] = & $convert $p.Value ($d - 1) } catch { $o[$p.Name] = $null }
        }
        return [pscustomobject]$o
    }
    return , (& $convert $InputObject $Depth)
}
