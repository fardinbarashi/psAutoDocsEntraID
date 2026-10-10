# Autodoc v1.0
function Show-IconMapStatus {
    <#
        Reports how the tenant's SKUs line up with the icons on disk: which SKUs
        have an icon and how it was found, which have none, and which icon files
        nothing points at.

        Run this after dropping icons into files\cache\Azure icons\Icons to see what is left
        to map, then add the missing entries to Entra\Config\IconMap.psd1.

            Show-IconMapStatus
            Show-IconMapStatus -ShowUnused
    #>
    [CmdletBinding()]
    param(
        [string]$ExportsRoot,
        [string]$SourceFolder,
        [string]$IconFolder,
        [switch]$ShowUnused
    )

    $entraRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
    if (-not $ExportsRoot) { $ExportsRoot = Join-Path $entraRoot 'files\exports' }
    if (-not $IconFolder)  { $IconFolder  = Get-MapIconFolder -BuilderRoot $PSScriptRoot }

    if (-not $SourceFolder) {
        if (-not (Test-Path $ExportsRoot)) { Write-Host "No exports at $ExportsRoot - run a collection first." -ForegroundColor Yellow; return }
        $SourceFolder = (Get-ChildItem -Path $ExportsRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
    }

    $skus = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    if (-not $skus) { Write-Host "No licence data in $SourceFolder" -ForegroundColor Yellow; return }

    $iconSet = Get-MapIconSet -Folder $IconFolder
    $iconMap = Get-IconMap

    Write-Host ""
    Write-Host "Icon folder : $IconFolder" -ForegroundColor Cyan
    Write-Host "Icon files  : $($iconSet.Files.Count)   ($(@($iconSet.Files | Where-Object Extension -eq '.png').Count) PNG, $(@($iconSet.Files | Where-Object Extension -eq '.svg').Count) SVG)" -ForegroundColor Cyan
    Write-Host "Map entries : $($iconMap.Count)   (Entra\Config\IconMap.psd1)" -ForegroundColor Cyan
    Write-Host ""

    $used = @{}
    $rows = foreach ($s in $skus) {
        $sku  = [string]$s.skuPartNumber
        $file = Resolve-MapIconFile -IconSet $iconSet -Sku $sku -IconMap $iconMap
        $how  = 'none'
        if ($file) {
            $used[$file.FullName] = $true
            $key = $sku.ToUpperInvariant()
            $how = if ($iconSet.ByName.ContainsKey($key) -and $iconSet.ByName[$key].FullName -eq $file.FullName) { 'file name' }
                   elseif ($iconMap.Keys | Where-Object { $_.ToUpperInvariant() -eq $key }) { 'icon map' }
                   else { 'sku stem' }
        }
        [pscustomobject]@{
            Sku    = $sku
            Icon   = $(if ($file) { $file.Name } else { '' })
            Via    = $how
            Visio  = $(if ($file -and $file.Extension -eq '.png') { 'yes' } elseif ($file) { 'svg only' } else { '' })
        }
    }

    $have = @($rows | Where-Object { $_.Icon })
    $miss = @($rows | Where-Object { -not $_.Icon })

    Write-Host "MATCHED ($($have.Count) of $($rows.Count))" -ForegroundColor Green
    $have | Sort-Object Sku | ForEach-Object { Write-Host ("   {0,-34} {1,-9} {2,-9} {3}" -f $_.Sku, $_.Via, $_.Visio, $_.Icon) }

    if ($miss) {
        Write-Host ""
        Write-Host "NO ICON ($($miss.Count))  - add these to Entra\Config\IconMap.psd1" -ForegroundColor Yellow
        $miss | Sort-Object Sku | ForEach-Object { Write-Host ("   '{0}' = ''" -f $_.Sku) -ForegroundColor DarkGray }
    }

    $unused = @($iconSet.Files | Where-Object { -not $used.ContainsKey($_.FullName) })
    Write-Host ""
    Write-Host "UNUSED ICON FILES: $($unused.Count)" -ForegroundColor Cyan
    if ($ShowUnused -and $unused) {
        $unused | Sort-Object Name | Select-Object -First 60 | ForEach-Object { Write-Host ("   " + $_.Name) -ForegroundColor DarkGray }
        if ($unused.Count -gt 60) { Write-Host ("   ... and $($unused.Count - 60) more") -ForegroundColor DarkGray }
    }
    elseif ($unused) { Write-Host "   run with -ShowUnused to list them" -ForegroundColor DarkGray }
    Write-Host ""
}
