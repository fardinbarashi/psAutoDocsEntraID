# Autodoc v1.0
function Build-GroupOwnerMapSvg {
    <#
        SVG map of group ownership: one card per owner, listing the groups they
        own. Colour strength = number of groups owned (darker = more), so owners
        responsible for many groups stand out. Members across those groups are
        shown as a secondary stat.

        Data: EntraGroups.Owners (an owner is a UPN/email; a group can have
        several owners). Groups with no owner do not appear here - the Groups map
        already flags ownerless groups with a red edge.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 4,
        [int]$MaxGroupsPerCard = 12
    )

    $data = Get-GroupOwnerMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No group data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $owners = $data.Owners
    if (-not $owners -or $owners.Count -eq 0) { Write-Host "No owned groups in this export." -ForegroundColor Yellow; return }

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    $ownerIcon  = Get-MapIcon -IconSet $iconSet -Sku 'users' -IconMap $iconMap -PreferSvg

    # colour ramp by number of groups owned (light -> dark Azure blue)
    $maxOwned = ($owners | Measure-Object -Property { $_.Groups.Count } -Maximum).Maximum
    if (-not $maxOwned) { $maxOwned = 1 }
    function Get-HeatFill([int]$n, [int]$max) {
        $t = [math]::Min(1.0, $n / [double]$max)
        $r = [int](220 + ($t * (15  - 220)))
        $g = [int](231 + ($t * (76  - 231)))
        $b = [int](243 + ($t * (129 - 243)))
        '#{0:X2}{1:X2}{2:X2}' -f $r, $g, $b
    }

    $gapX = 0.3; $gapY = 0.4
    $tenantH = 1.25; $marginTop = 0.5; $marginBottom = 0.6
    $legendH = (Get-LegendHeight -LineCount 3)
    $cols = [math]::Max(1, [math]::Min($PerRow, 4))
    $lineH = 0.185; $cardPad = 0.14
    $pageW = 13.0

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'EntraGroups' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    $nodeW = ($pageW - 2 - ($cols - 1) * $gapX) / $cols
    $innerW = $nodeW - 2 * $cardPad

    $cards = @()
    foreach ($own in $owners) {
        $groups = @($own.Groups | Sort-Object -Unique)
        if ($groups.Count -gt $MaxGroupsPerCard) {
            $shown = @($groups | Select-Object -First $MaxGroupsPerCard)
            $hidden = $groups.Count - $shown.Count
        } else { $shown = $groups; $hidden = 0 }
        $lines = @($shown)
        if ($hidden -gt 0) { $lines += "+$hidden more groups" }

        $grpWord = if ($groups.Count -eq 1) { 'group' } else { 'groups' }
        $memWord = if ($own.Members -eq 1) { 'member' } else { 'members' }
        $statText = "owns $($groups.Count) $grpWord | $($own.Members) $memWord"
        # wrap each line to the width right of the icon so long group names never spill
        $availW = ($nodeW - $nodeW * 0.34 - 2 * $cardPad) * 0.97
        $cardLines = @()
        foreach ($p in @(Split-MapTextToWidth -Text "$($own.Owner)" -AvailIn $availW -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $true; Align = 'start' } }
        foreach ($p in @(Split-MapTextToWidth -Text $statText -AvailIn $availW -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $false; Align = 'start' } }
        foreach ($ln in $lines) { foreach ($p in @(Split-MapTextToWidth -Text "$ln" -AvailIn $availW -FontSize 10)) { $cardLines += @{ Text = $p; Bold = $false; Align = 'start' } } }
        $cardH = $cardPad + $cardLines.Count * $lineH + $cardPad + 0.12

        $cards += [pscustomobject]@{ CardLines = $cardLines; Height = $cardH; Fill = (Get-HeatFill $groups.Count $maxOwned) }
    }

    # masonry placement
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
            @{ Text = "$($tenant.TenantName)";                 Bold = $true;  Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)";              Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Owners: $($owners.Count)   Owned groups: $($data.OwnedGroupCount)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today";                      Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Card = '; Text = 'an owner and the groups they own'; Align = 'start' }
        @{ BoldPrefix = 'Darker = '; Text = 'owns more groups'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=6.0; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
    $gridTop = $legendY - $legendH / 2 - $gapY - 0.6
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]; $pl = $placement[$i]; $cardH = $card.Height
        $cx = $left + $pl.Col * ($nodeW + $gapX) + $nodeW / 2
        $cy = $gridTop - $pl.TopOffset - $cardH / 2

        $shapes += @{ Id="own$i"; Kind='Rectangle'; Lines=$card.CardLines; LinesTop=$true; Icon=$ownerIcon
                      X=$cx; Y=$cy; W=$nodeW; H=$cardH; Fill=$card.Fill; Line='#888888'; FontSize=9 }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Group owners $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($owners.Count) owners, $($data.OwnedGroupCount) owned groups" -ForegroundColor DarkGray
    return $svg
}
