# Autodoc v1.0
function Build-OrgMapVisio {
    <#
        Builds ONLY the Visio (.vsdx) organisation map. Independent of the SVG
        builder. Workplace cards use the split-shape approach (background box +
        bold header shape + normal units shape) so bold survives in Visio.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$MaxUnitsPerWorkplace = 0,
        [int]$PerRow = 6
    )

    $data = Get-OrgMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No user data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $tree = $data.Tree; $workplaces = $data.Workplaces; $userCount = $data.UserCount

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $wpIcon     = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap

    $gapX = 0.3; $gapY = 0.4
    $tenantH = 1.25; $marginTop = 0.5; $marginBottom = 0.6
    $cols = [math]::Max(1, [math]::Min($PerRow, 4))
    $lineH = 0.185; $cardPad = 0.12
    $pageW = 13.0
    $nodeW = ($pageW - 2 - ($cols - 1) * $gapX) / $cols
    $innerW = $nodeW - 2 * $cardPad

    # card content + measured height
    $cards = @()
    foreach ($wp in $workplaces) {
        $units = @($tree[$wp].GetEnumerator() | Sort-Object Value -Descending)
        $total = ($tree[$wp].Values | Measure-Object -Sum).Sum
        if ($MaxUnitsPerWorkplace -gt 0) { $shown = @($units | Select-Object -First $MaxUnitsPerWorkplace); $hidden = $units.Count - $shown.Count }
        else { $shown = $units; $hidden = 0 }
        $lines = @($shown | ForEach-Object { "$($_.Key)  ($($_.Value))" })
        if ($hidden -gt 0) { $lines += "+$hidden more units" }
        $headText = "$wp  -  $total users"
        $visualLines = [math]::Max(1, [math]::Ceiling((Measure-MapTextWidth $headText 9) * 1.05 / $innerW))
        foreach ($ln in $lines) { $visualLines += [math]::Max(1, [math]::Ceiling((Measure-MapTextWidth $ln 9) / $innerW)) }
        $cardH = $cardPad + $visualLines * $lineH + $cardPad
        $cards += [pscustomobject]@{ Name = $wp; Total = $total; Lines = $lines; Height = $cardH }
    }

    # masonry placement
    $colHeights = @(0) * $cols; $placement = @()
    foreach ($card in $cards) {
        $c = 0; for ($k = 1; $k -lt $cols; $k++) { if ($colHeights[$k] -lt $colHeights[$c]) { $c = $k } }
        $placement += [pscustomobject]@{ Col = $c; TopOffset = $colHeights[$c] }
        $colHeights[$c] += $card.Height + $gapY
    }
    $tallestCol = ($colHeights | Measure-Object -Maximum).Maximum
    $orgLegendH = 1.15
    $vSpan = $marginTop + $tenantH + $gapY + $orgLegendH + $gapY + 0.6 + $tallestCol + $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @(); $links = @()
    $topY = $pageH - $marginTop - $tenantH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant: split for Visio
    if ($tenant) {
        $shapes += @{ Id='tenant'; Kind='Rectangle'; Text=''
                      X=$pageW/2; Y=$topY; W=5.2; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$wpIcon }
        $shapes += @{ Id='tenant-name'; Kind='Rectangle'; Text=$tenant.TenantName; Bold=$true
                      X=$pageW/2 + 0.55; Y=($topY + $tenantH/2 - 0.28); W=4.0; H=0.3
                      Fill='none'; Line='none'; FontSize=11; LinesTop=$true }
        $detail = "$($tenant.PrimaryDomain)`nUsers: $userCount   Workplaces: $($workplaces.Count)`nGenerated $today"
        $shapes += @{ Id='tenant-detail'; Kind='Rectangle'; Text=$detail
                      X=$pageW/2 + 0.55; Y=($topY - 0.12); W=4.0; H=0.72
                      Fill='none'; Line='none'; FontSize=11; LinesTop=$true }
    } else {
        $shapes += @{ Id='tenant'; Kind='Rectangle'; Text='Tenant'; Bold=$true
                      X=$pageW/2; Y=$topY; W=5.2; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$wpIcon }
    }

    # legend: split for Visio
    $orgLegendY = $topY - $tenantH / 2 - $gapY - $orgLegendH / 2
    $orgLegW = 6.0
    $shapes += @{ Id='orglegend'; Kind='Rectangle'; Text=''; Dashed=$true
                  X=$pageW/2; Y=$orgLegendY; W=$orgLegW; H=$orgLegendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }
    $olLeft = $pageW/2 - $orgLegW/2 + 0.18; $olTop = $orgLegendY + $orgLegendH/2 - 0.20; $olH = 0.24
    $shapes += @{ Id='ol-0'; Kind='Rectangle'; Text='How to read this map'; Bold=$true
                  X=($olLeft + 2.5/2); Y=$olTop; W=2.5; H=$olH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
    $shapes += @{ Id='ol-1'; Kind='Rectangle'; Text='Users grouped by their CompanyName and Department fields'
                  X=($olLeft + 5.5/2); Y=($olTop - $olH); W=5.5; H=$olH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
    $olRows = @(
        @{ Key = 'Block = '; Rest = 'CompanyName (total users)' }
        @{ Key = 'Line = ';  Rest = 'Department (users)' }
    )
    for ($oi = 0; $oi -lt $olRows.Count; $oi++) {
        $oy = $olTop - ($oi + 2) * $olH
        $kw = (Measure-MapTextWidth $olRows[$oi].Key 9) + 0.05
        $shapes += @{ Id="ol-k$oi"; Kind='Rectangle'; Text=$olRows[$oi].Key; Bold=$true
                      X=($olLeft + $kw/2); Y=$oy; W=$kw; H=$olH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
        $shapes += @{ Id="ol-r$oi"; Kind='Rectangle'; Text=$olRows[$oi].Rest
                      X=($olLeft + $kw + 3.5/2); Y=$oy; W=3.5; H=$olH; Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
    }

    # cards: background + bold header + normal units
    $palette = @('#DCE7F3','#E7F0DC','#F3E7DC','#EDDCF3','#DCF3EF','#F3DCE4','#F0F3DC','#DCE9F3')
    $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
    $gridTop = $orgLegendY - $orgLegendH / 2 - $gapY - 0.6
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]; $pl = $placement[$i]; $cardH = $card.Height
        $cx = $left + $pl.Col * ($nodeW + $gapX) + $nodeW / 2
        $cy = $gridTop - $pl.TopOffset - $cardH / 2
        $id = "wp$i"
        $shapes += @{ Id=$id; Kind='Rectangle'; Text=''
                      X=$cx; Y=$cy; W=$nodeW; H=$cardH; Fill=$palette[$i % $palette.Count]; Line='#888888'; FontSize=9 }
        $hdrH = 0.26
        $shapes += @{ Id="$id-h"; Kind='Rectangle'; Text="$($card.Name)  -  $($card.Total) users"; Bold=$true
                      X=$cx; Y=($cy + $cardH/2 - $hdrH/2); W=($nodeW - 0.12); H=$hdrH
                      Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
        if ($card.Lines.Count -gt 0) {
            $unitsH = $cardH - $hdrH
            $shapes += @{ Id="$id-u"; Kind='Rectangle'; Text=($card.Lines -join "`n")
                          X=$cx; Y=($cy + $cardH/2 - $hdrH - $unitsH/2); W=($nodeW - 0.12); H=$unitsH
                          Fill='none'; Line='none'; FontSize=9; LinesTop=$true }
        }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $visioDir = Join-Path $SourceFolder 'visio'
    if (-not (Test-Path $visioDir)) { New-Item -Path $visioDir -ItemType Directory -Force | Out-Null }
    $vsdx = Join-Path $visioDir "OrgMap $stamp.vsdx"
    New-VisioDocument -Path $vsdx -Shape $shapes -Connector $links `
        -PageWidth $pageW -PageHeight $pageH -PageName 'Organisation' `
        -Title $(if ($tenant) { "Organisation map - $($tenant.TenantName)" } else { 'Organisation map' }) | Out-Null

    Write-Host "  Visio: 1 tenant, $($workplaces.Count) workplaces" -ForegroundColor DarkGray
    return $vsdx
}
