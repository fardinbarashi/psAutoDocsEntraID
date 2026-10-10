# Autodoc v1.0
function Build-LicenceMapVisio {
    <#
        Builds ONLY the Visio (.vsdx) licence map. Nothing here touches the SVG
        renderer, so changes are safe to make without affecting the picture.

        Visio note: per-line bold inside one shape does not survive (cp/pp runs
        collapse the text), so blocks that need a bold heading are drawn as a
        background rectangle plus separate uniformly-formatted text shapes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 6
    )

    $data = Get-LicenceMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No licence data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $purchased = $data.Purchased; $viral = $data.Viral

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet     = Get-MapIconSet -Folder $iconFolder
    $iconMap     = Get-IconMap
    $tenantIcon  = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap
    $genericIcon = Get-MapIcon -IconSet $iconSet -Sku 'SKU' -IconMap $iconMap

    # ---------- layout ----------
    $nodeW = 2.55; $nodeH = 0.85; $gapX = 0.28; $gapY = 0.65
    $tenantH = 1.25; $bandH = 0.6; $marginTop = 0.5; $marginBottom = 0.6; $legendH = 1.25
    $cols  = [math]::Min($PerRow, [math]::Max($purchased.Count, 1))
    $pageW = [math]::Max(11, [math]::Max(5.6, $cols * ($nodeW + $gapX) + $gapX + 1))
    $rowsP = if ($purchased) { [math]::Ceiling($purchased.Count / $cols) } else { 0 }
    $rowsV = if ($viral)     { [math]::Ceiling($viral.Count / $cols) }     else { 0 }
    $vSpan = $marginTop + $tenantH + $gapY + $legendH + $gapY
    if ($purchased) { $vSpan += $bandH + $gapY + $rowsP * $nodeH + ($rowsP - 1) * $gapY + $gapY }
    if ($viral)     { $vSpan += $bandH + $gapY + $rowsV * $nodeH + ($rowsV - 1) * $gapY }
    $vSpan += $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @(); $links = @()
    $topY  = $pageH - $marginTop - $tenantH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant: background + bold name + normal details (split for Visio)
    if ($tenant) {
        $shapes += @{ Id='tenant'; Kind='Rectangle'; Text=''
                      X=$pageW/2; Y=$topY; W=5.2; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }
        $shapes += @{ Id='tenant-name'; Kind='Rectangle'; Text=$tenant.TenantName; Bold=$true
                      X=$pageW/2 + 0.55; Y=($topY + $tenantH/2 - 0.28); W=4.0; H=0.3
                      Fill='none'; Line='none'; FontSize=11; LinesTop=$true }
        $detail = "$($tenant.PrimaryDomain)`nTenant ID: $($tenant.TenantId)`nGenerated $today"
        $shapes += @{ Id='tenant-detail'; Kind='Rectangle'; Text=$detail
                      X=$pageW/2 + 0.55; Y=($topY - 0.12); W=4.0; H=0.72
                      Fill='none'; Line='none'; FontSize=11; LinesTop=$true }
    } else {
        $shapes += @{ Id='tenant'; Kind='Rectangle'; Text='Tenant'; Bold=$true
                      X=$pageW/2; Y=$topY; W=5.2; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }
    }

    # colour legend: dashed box + per-line bold key + normal rest
    $legendY = $topY - $tenantH / 2 - $gapY - $legendH / 2
    $legendRows = @(
        @{ Key = 'Numbers: '; Rest = 'Used / Total  (percent used)' }
        @{ Key = 'Red: ';     Rest = '10% or fewer licences free' }
        @{ Key = 'Yellow: ';  Rest = '11-30% of licences free' }
        @{ Key = 'Green: ';   Rest = '31-100% of licences free' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Text=''; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=5.2; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }
    $lgLineH = 0.24; $lgTop = $legendY + $legendH/2 - 0.20; $lgLeft = $pageW/2 - 5.2/2 + 0.18
    for ($li = 0; $li -lt $legendRows.Count; $li++) {
        $ly = $lgTop - $li * $lgLineH
        $keyW = (Measure-MapTextWidth $legendRows[$li].Key 9) + 0.05
        $shapes += @{ Id="lg-k$li"; Kind='Rectangle'; Text=$legendRows[$li].Key; Bold=$true
                      X=($lgLeft + $keyW/2); Y=$ly; W=$keyW; H=$lgLineH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
        $shapes += @{ Id="lg-r$li"; Kind='Rectangle'; Text=$legendRows[$li].Rest
                      X=($lgLeft + $keyW + 4.6/2); Y=$ly; W=4.6; H=$lgLineH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
    }

    function Add-BandV([string]$BandId, [string]$BandText, [double]$BandY, [string]$BandFill) {
        @{ Id=$BandId; Kind='Rectangle'; Text=$BandText; Bold=$true; CenterText=$true; TopInset=0.14
           X=$pageW/2; Y=$BandY; W=2.6; H=0.6; Fill=$BandFill; Line='#666666'; FontSize=11 }
    }

    $y = $legendY - $legendH / 2 - $gapY - $bandH / 2
    if ($purchased) {
        $shapes += (Add-BandV 'band-p' "Purchased  ($($purchased.Count))" $y '#DCE7F3')
        $y -= $bandH / 2 + $gapY + $nodeH / 2
        $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
        for ($i = 0; $i -lt $purchased.Count; $i++) {
            $p = $purchased[$i]; $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            $shapes += @{ Id="p$i"; Kind='Rectangle'; Text="$($p.Sku)`n$($p.Consumed)/$($p.Enabled)  ($($p.Util)%)"; CenterText=$true; TopInset=0.12
                          X=$x; Y=($y - $row * ($nodeH + $gapY)); W=$nodeW; H=$nodeH
                          Fill=(Get-LicenceFreeFill $p.PercentFree); Line='#555555'; FontSize=9 }
        }
        $y -= ($rowsP - 1) * ($nodeH + $gapY) + $nodeH / 2 + $gapY + $bandH / 2
    }
    if ($viral) {
        $shapes += (Add-BandV 'band-v' "Free / viral  ($($viral.Count))" $y '#EEEEEE')
        $y -= $bandH / 2 + $gapY + $nodeH / 2
        $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
        for ($i = 0; $i -lt $viral.Count; $i++) {
            $v = $viral[$i]; $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            $shapes += @{ Id="v$i"; Kind='Rectangle'; Text="$($v.Sku)`n$($v.Consumed) used"; CenterText=$true; TopInset=0.12
                          X=$x; Y=($y - $row * ($nodeH + $gapY)); W=$nodeW; H=$nodeH
                          Fill='#F2F2F2'; Line='#999999'; FontSize=9 }
        }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $visioDir = Join-Path $SourceFolder 'visio'
    if (-not (Test-Path $visioDir)) { New-Item -Path $visioDir -ItemType Directory -Force | Out-Null }
    $vsdx = Join-Path $visioDir "LicenceMap $stamp.vsdx"
    New-VisioDocument -Path $vsdx -Shape $shapes -Connector $links `
        -PageWidth $pageW -PageHeight $pageH -PageName 'Licences' `
        -Title $(if ($tenant) { "Licence map - $($tenant.TenantName)" } else { 'Licence map' }) | Out-Null

    Write-Host "  Visio: 1 tenant, $($purchased.Count) purchased, $($viral.Count) free/viral" -ForegroundColor DarkGray
    return $vsdx
}
