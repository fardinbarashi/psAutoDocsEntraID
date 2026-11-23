# Autodoc v1.0
function Save-SecurityImprovementData {
    <#
        Writes the result of Get-SecurityImprovementData to a JSON file. This
        file is what the Word, Excel and map files are built from, so the three
        always show the same proposals and can be rebuilt later from the file
        (Build-EntraSecurityReport -DataJson). The texts stay in
        Entra\Config\SecurityImprovements.json and are not copied in.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string]$SourceFolder
    )

    $doc = [ordered]@{
        Generated    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        SourceFolder = $SourceFolder
        Language     = $Data.Language
        Tenant       = $Data.Tenant
        AreaOrder    = @($Data.AreaOrder)
        SevOrder     = @($Data.SevOrder)
        Improvements = @($Data.Improvements)
        Hunt         = $Data.Hunt
    }
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $doc | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
    return $Path
}

function Import-SecurityImprovementData {
    <#
        Reads a JSON file written by Save-SecurityImprovementData and returns the
        same shape as Get-SecurityImprovementData (texts re-attached from
        Entra\Config\SecurityImprovements.json). -Language overrides the
        language stored in the file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Language
    )

    $doc = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))   # Functions\Reports\Security -> Entra
    $cfg = Get-Content -LiteralPath (Join-Path $entraRoot 'Config\SecurityImprovements.json') -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
    $lang = if ($Language) { $(if ($Language -eq 'en') { 'en' } else { 'sv' }) } else { "$($doc.Language)" }

    # another language than the file: take the rule texts from the config again
    # (texts edited by hand in the JSON are kept when the language is the same)
    $improvements = @($doc.Improvements | Where-Object { $null -ne $_ })
    if ($lang -ne "$($doc.Language)") {
        foreach ($row in $improvements) {
            $rule = $cfg.Rules["$($row.Id)"]
            if (-not $rule) { continue }
            $pick = { param($e) if ($e -and "$($e[$lang])") { "$($e[$lang])" } else { "$($e['en'])" } }
            $row.AreaName       = & $pick $cfg.Areas["$($rule.Area)"]
            $row.Title          = & $pick $rule.Title
            $row.Finding        = (& $pick $rule.Finding) -f $row.Count, "$($row.Extra)"
            $row.Recommendation = & $pick $rule.Recommendation
        }
    }

    [pscustomobject]@{
        Language     = $lang
        Config       = $cfg
        Tenant       = $doc.Tenant
        SourceFolder = "$($doc.SourceFolder)"
        Improvements = $improvements
        Hunt         = [pscustomobject]@{
            Ran         = [bool]$doc.Hunt.Ran
            Findings    = @($doc.Hunt.Findings | Where-Object { $null -ne $_ })
            Triage      = @($doc.Hunt.Triage | Where-Object { $null -ne $_ })
            Credentials = @($doc.Hunt.Credentials | Where-Object { $null -ne $_ })
            Report      = @($doc.Hunt.Report)
        }
        AreaOrder    = @($doc.AreaOrder)
        SevOrder     = @($doc.SevOrder)
    }
}
