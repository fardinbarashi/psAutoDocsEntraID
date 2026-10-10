# Autodoc v1.0
function Build-OrgMapSvg {
    <#
        Builds the SVG organisation map.
        Cards use the rich Lines model (bold header + normal unit lines in one
        shape) since SVG renders per-line bold correctly.
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
    $wpIcon     = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    $gapX = 0.3; $gapY = 0.4
    $tenantH = 1.25; $marginTop = 0.5; $marginBottom = 0.6
    $cols = [math]::Max(1, [math]::Min($PerRow, 4))
    $lineH = 0.185; $cardPad = 0.12
    $pageW = 13.0

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'UserInformation' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    $nodeW = ($pageW - 2 - ($cols - 1) * $gapX) / $cols
    $innerW = $nodeW - 2 * $cardPad

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

    $colHeights = @(0) * $cols; $placement = @()
    foreach ($card in $cards) {
        $c = 0; for ($k = 1; $k -lt $cols; $k++) { if ($colHeights[$k] -lt $colHeights[$c]) { $c = $k } }
        $placement += [pscustomobject]@{ Col = $c; TopOffset = $colHeights[$c] }
        $colHeights[$c] += $card.Height + $gapY
    }
    $tallestCol = ($colHeights | Measure-Object -Maximum).Maximum
    $orgLegendH = 1.15
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $orgLegendH + $gapY + 0.6 + $tallestCol + $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @(); $links = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant: one shape with per-line bold
    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Users: $userCount   Workplaces: $($workplaces.Count)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$wpIcon }

    # legend: one dashed shape with bold heading + bold-prefix rows
    $orgLegendY = $dsY - $dsH / 2 - $gapY - $orgLegendH / 2
    $orgLegendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ Text = 'Users grouped by their CompanyName and Department fields'; Bold = $false; Align = 'start' }
        @{ BoldPrefix = 'Block = '; Text = 'CompanyName (total users)'; Align = 'start' }
        @{ BoldPrefix = 'Line = ';  Text = 'Department (users)';        Align = 'start' }
    )
    $shapes += @{ Id='orglegend'; Kind='Rectangle'; Lines=$orgLegendLines; LinesTop=$true; Dashed=$true; TopInset=0.14
                  X=$pageW/2; Y=$orgLegendY; W=6.0; H=$orgLegendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    # cards: one shape each, bold header line + normal unit lines
    $palette = @('#DCE7F3','#E7F0DC','#F3E7DC','#EDDCF3','#DCF3EF','#F3DCE4','#F0F3DC','#DCE9F3')
    $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
    $gridTop = $orgLegendY - $orgLegendH / 2 - $gapY - 0.6
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]; $pl = $placement[$i]; $cardH = $card.Height
        $cx = $left + $pl.Col * ($nodeW + $gapX) + $nodeW / 2
        $cy = $gridTop - $pl.TopOffset - $cardH / 2
        $cardLines = @(@{ Text = "$($card.Name)  -  $($card.Total) users"; Bold = $true; Align = 'start' })
        foreach ($ln in $card.Lines) { $cardLines += @{ Text = $ln; Bold = $false; Align = 'start' } }
        $shapes += @{ Id="wp$i"; Kind='Rectangle'; Lines=$cardLines; LinesTop=$true
                      X=$cx; Y=$cy; W=$nodeW; H=$cardH; Fill=$palette[$i % $palette.Count]; Line='#888888'; FontSize=9 }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Organization Department Officelocation chart $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector $links -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: 1 tenant, $($workplaces.Count) workplaces" -ForegroundColor DarkGray
    return $svg
}
