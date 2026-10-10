# Autodoc v1.0
function Build-GroupDeptListSvg {
    <#
        View 1 of group -> department: one card per group with its departments
        listed inside (name + member count), like the organisation map. Best for
        detail and works at any group size.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 4,
        [int]$MaxDeptsPerCard = 0
    )

    $data = Get-GroupDeptMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No group-member data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $tree = $data.Tree; $groups = $data.Groups

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    $gapX = 0.3; $gapY = 0.4
    $tenantH = 1.25; $legendH = 0.75; $marginTop = 0.5; $marginBottom = 0.6
    $cols = [math]::Max(1, [math]::Min($PerRow, 4))
    $lineH = 0.185; $cardPad = 0.12
    $pageW = 13.0

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'GroupMembers' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    $nodeW = ($pageW - 2 - ($cols - 1) * $gapX) / $cols
    $innerW = $nodeW - 2 * $cardPad

    $cards = @()
    foreach ($grp in $groups) {
        $depts = @($tree[$grp].GetEnumerator() | Sort-Object Value -Descending)
        $total = ($tree[$grp].Values | Measure-Object -Sum).Sum
        if ($MaxDeptsPerCard -gt 0 -and $depts.Count -gt $MaxDeptsPerCard) {
            $shown = @($depts | Select-Object -First $MaxDeptsPerCard); $hidden = $depts.Count - $shown.Count
        } else { $shown = $depts; $hidden = 0 }
        $lines = @($shown | ForEach-Object { "$($_.Key)  ($($_.Value))" })
        if ($hidden -gt 0) { $lines += "+$hidden more departments" }
        $statText = "$total members | $($depts.Count) departments"
        # wrap every line to the card width so long group / department names
        # never spill outside the box
        $avail = $innerW * 0.97
        $cardLines = @()
        foreach ($p in @(Split-MapTextToWidth -Text "$grp" -AvailIn $avail -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $true; Align = 'start' } }
        foreach ($p in @(Split-MapTextToWidth -Text $statText -AvailIn $avail -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $false; Align = 'start' } }
        foreach ($ln in $lines) { foreach ($p in @(Split-MapTextToWidth -Text $ln -AvailIn $avail -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $false; Align = 'start' } } }
        $cardH = $cardPad + $cardLines.Count * $lineH + $cardPad + 0.12
        $cards += [pscustomobject]@{ Group = $grp; CardLines = $cardLines; Height = $cardH }
    }

    $colHeights = @(0) * $cols; $placement = @()
    foreach ($card in $cards) {
        $c = 0; for ($k = 1; $k -lt $cols; $k++) { if ($colHeights[$k] -lt $colHeights[$c]) { $c = $k } }
        $placement += [pscustomobject]@{ Col = $c; TopOffset = $colHeights[$c] }
        $colHeights[$c] += $card.Height + $gapY
    }
    $tallestCol = ($colHeights | Measure-Object -Maximum).Maximum
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + 0.6 + $tallestCol + $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Groups: $($groups.Count)   Memberships: $($data.TotalMembers)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Block = '; Text = 'a group (total members)'; Align = 'start' }
        @{ BoldPrefix = 'Line = ';  Text = 'a department in that group (members)'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; LinesTop=$true; Dashed=$true; TopInset=0.14
                  X=$pageW/2; Y=$legendY; W=6.0; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    $palette = @('#DCE7F3','#E7F0DC','#F3E7DC','#EDDCF3','#DCF3EF','#F3DCE4','#F0F3DC','#DCE9F3')
    $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
    $gridTop = $legendY - $legendH / 2 - $gapY - 0.6
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]; $pl = $placement[$i]; $cardH = $card.Height
        $cx = $left + $pl.Col * ($nodeW + $gapX) + $nodeW / 2
        $cy = $gridTop - $pl.TopOffset - $cardH / 2
        $shapes += @{ Id="grp$i"; Kind='Rectangle'; Lines=$card.CardLines; LinesTop=$true
                      X=$cx; Y=$cy; W=$nodeW; H=$cardH; Fill=$palette[$i % $palette.Count]; Line='#888888'; FontSize=9 }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Groups - members by department (list) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null
    Write-Host "  SVG (list): $($groups.Count) groups" -ForegroundColor DarkGray
    return $svg
}
