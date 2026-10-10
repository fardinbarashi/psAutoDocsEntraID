# Autodoc v1.0
function Get-MapIconSet {
    <#
        Indexes the icon folder once. Returns @{ Files = <FileInfo[]>; ByName = @{} }
        where ByName is keyed on the file name in upper case.

        PNG and SVG are not interchangeable downstream:
          PNG  works in both the .vsdx and the .svg map
          SVG  works in the .svg map only - Visio stores pictures as ForeignData
               records, which carry raster data
        Microsoft's pack ships both, so keeping the PNG gets icons into Visio too.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder)

    $files = @()
    if (Test-Path $Folder) {
        $files = @(Get-ChildItem -Path $Folder -File -Recurse -ErrorAction SilentlyContinue |
                   Where-Object { $_.Extension -in '.png', '.svg' })
    }
    $byName = @{}      # PNG beats SVG (for Visio, which needs raster)
    $bySvg  = @{}      # SVG beats PNG (for SVG maps, which want vector)
    foreach ($f in $files) {
        $key = [IO.Path]::GetFileNameWithoutExtension($f.Name).ToUpperInvariant()
        # PNG-preferred index
        if (-not ($byName.ContainsKey($key) -and $byName[$key].Extension -eq '.png')) { $byName[$key] = $f }
        # SVG-preferred index
        if (-not ($bySvg.ContainsKey($key) -and $bySvg[$key].Extension -eq '.svg')) { $bySvg[$key] = $f }
    }
    @{ Files = $files; ByName = $byName; BySvg = $bySvg }
}

function Get-IconMap {
    <#
        Loads the SKU-to-file-name table from Entra\Config\IconMap.psd1.
        Kept as data rather than code so the mapping can be edited without
        touching any function.
    #>
    [CmdletBinding()]
    param([string]$Path)

    if (-not $Path) {
        $entraRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
        $Path = Join-Path $entraRoot 'Config\IconMap.psd1'
    }
    if (-not (Test-Path $Path)) { return @{} }
    try { $d = Import-PowerShellDataFile $Path; if ($d.Map) { return $d.Map } }
    catch { Write-Host "  Could not read icon map: $($_.Exception.Message)" -ForegroundColor Yellow }
    @{}
}

function Resolve-MapIconFile {
    <#
        Finds the icon file for one SKU. Three passes, first hit wins:
          1. exact file name          ENTERPRISEPACK.png
          2. the icon map             ENTERPRISEPACK -> 'Azure-Active-Directory'
          3. SKU up to first underscore  SPE_E3 -> SPE
        Returns the FileInfo, or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$IconSet,
        [Parameter(Mandatory)][string]$Sku,
        [hashtable]$IconMap = @{},
        [switch]$PreferSvg
    )

    # pick the name index: SVG-preferred for SVG maps, PNG-preferred otherwise
    $index = if ($PreferSvg -and $IconSet.BySvg) { $IconSet.BySvg } else { $IconSet.ByName }

    $key = $Sku.ToUpperInvariant()
    if ($index.ContainsKey($key)) { return $index[$key] }

    $fragment = $null
    foreach ($k in $IconMap.Keys) { if ($k.ToUpperInvariant() -eq $key) { $fragment = $IconMap[$k]; break } }
    if ($fragment) {
        $hits = @($IconSet.Files | Where-Object { $_.BaseName -like "*$fragment*" })
        $wantExt = if ($PreferSvg) { '.svg' } else { '.png' }
        $pick = @($hits | Where-Object { $_.Extension -eq $wantExt })[0]
        if (-not $pick) { $pick = $hits[0] }
        if ($pick) { return $pick }
    }

    $stem = ($key -split '_')[0]
    if ($stem -and $index.ContainsKey($stem)) { return $index[$stem] }
    $null
}

function Get-MapIcon {
    <# Returns @{ Path; Base64; Mime; IsPng } for a SKU, or $null. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$IconSet,
        [Parameter(Mandatory)][string]$Sku,
        [hashtable]$IconMap = @{},
        [switch]$PreferSvg
    )

    $file = Resolve-MapIconFile -IconSet $IconSet -Sku $Sku -IconMap $IconMap -PreferSvg:$PreferSvg
    if (-not $file) { return $null }

    $bytes = [IO.File]::ReadAllBytes($file.FullName)
    @{
        Path   = $file.FullName
        Base64 = [Convert]::ToBase64String($bytes)
        Mime   = $(if ($file.Extension -eq '.png') { 'image/png' } else { 'image/svg+xml' })
        IsPng  = ($file.Extension -eq '.png')
    }
}
