# Autodoc v1.0
function Build-GroupsMapSvg {
    <#
        SVG map of the tenant's groups: tenant at the top, a dashed legend, then
        one band per group category (Security, M365) with the groups as blocks.

        Encoding (agreed with the user):
          Fill   = source:  blue  = cloud-native,  grey = synced from on-prem AD
          Border = health:  red border if the group has no owner OR is empty
          Text   = name, member count, and a short "no owner"/"empty" note

        Scale guard: a real tenant has thousands of groups. Every category is
        shown, but only the largest groups per category are drawn individually;
        the rest collapse into a single "+N more groups" block.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 6,
        [int]$MaxPerCategory = 0     # 0 = show every group
    )

    $data = Get-GroupsMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No group data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $cats = $data.Categories

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    $iconSecurity = Get-MapIcon -IconSet $iconSet -Sku 'groups' -IconMap $iconMap -PreferSvg
    $iconM365     = Get-MapIcon -IconSet $iconSet -Sku 'groups' -IconMap $iconMap -PreferSvg

    # colours
    $cloudFill = '#0F6CBD'   # Azure blue = cloud-native
    $syncFill  = '#B8BFC7'   # steel grey = synced from on-prem
    $warnLine  = '#C81E1E'   # red border = no owner or empty
    $okLine    = '#555555'

    # ---------- layout ----------
    $nodeW = 2.55; $nodeH = 0.85; $gapX = 0.28; $gapY = 0.65
    $tenantH = 1.25; $bandH = 0.6; $marginTop = 0.5; $marginBottom = 0.6; $legendH = (Get-LegendHeight -LineCount 4)
    $cols = $PerRow

    # pre-compute blocks per category (with the +N more collapse)
    $catBlocks = @()
    foreach ($c in $cats) {
        $items = @($c.Group | Sort-Object Members -Descending)
        if ($MaxPerCategory -gt 0 -and $items.Count -gt $MaxPerCategory) {
            $shown = @($items | Select-Object -First $MaxPerCategory)
            $hidden = $items.Count - $shown.Count
        } else { $shown = $items; $hidden = 0 }
        $catBlocks += [pscustomobject]@{ Name = $c.Name; Shown = $shown; Hidden = $hidden; Count = $items.Count }
    }

    $pageW = [math]::Max(11, $cols * ($nodeW + $gapX) + $gapX + 1)

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'EntraGroups' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    # height: tenant + legend + for each category [band + its rows]
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY
    foreach ($cb in $catBlocks) {
        $n = $cb.Shown.Count + $(if ($cb.Hidden -gt 0) { 1 } else { 0 })
        $rows = [math]::Max(1, [math]::Ceiling($n / $cols))
        $vSpan += $bandH + $gapY + $rows * $nodeH + ($rows - 1) * $gapY + $gapY
    }
    $vSpan += $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @(); $links = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant (SVG: per-line bold via Lines)
    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)";     Bold = $true;  Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)";  Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Groups: $($data.Total)";    Bold = $false; Align = 'start' }
            @{ Text = "Generated $today";          Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    # legend
    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ BoldPrefix = 'Blue: ';     Text = 'cloud-native group';              Align = 'start' }
        @{ BoldPrefix = 'Grey: ';     Text = 'synced from on-prem AD';          Align = 'start' }
        @{ BoldPrefix = 'Red edge: '; Text = 'no owner or no members';          Align = 'start' }
        @{ Text = 'Text shows members and mail status (mail-enabled / no mail)'; Bold = $false; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=5.6; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    $y = $legendY - $legendH / 2 - $gapY - $bandH / 2
    foreach ($cb in $catBlocks) {
        # band per category
        $bandFill = if ($cb.Name -eq 'Security') { '#DCE7F3' } elseif ($cb.Name -eq 'M365') { '#E7F0DC' } else { '#F0F0F0' }
        $shapes += @{ Id="band-$($cb.Name)"; Kind='Rectangle'; Lines=@(@{ Text="$($cb.Name)  ($($cb.Count))"; Bold=$true; Align='middle' }); TopInset=0.14
                      X=$pageW/2; Y=$y; W=2.8; H=$bandH; Fill=$bandFill; Line='#666666'; FontSize=11 }
        $y -= $bandH / 2 + $gapY + $nodeH / 2

        $cells = @($cb.Shown)
        $total = $cells.Count + $(if ($cb.Hidden -gt 0) { 1 } else { 0 })
        $gridW = $cols * $nodeW + ($cols - 1) * $gapX
        $left  = ($pageW - $gridW) / 2

        for ($i = 0; $i -lt $total; $i++) {
            $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            $cy = $y - $row * ($nodeH + $gapY)

            if ($i -lt $cells.Count) {
                $grp = $cells[$i]
                $fill = if ($grp.Synced) { $syncFill } else { $cloudFill }
                $line = if (-not $grp.HasOwner -or $grp.Empty) { $warnLine } else { $okLine }
                # note line: flag the problems
                $notes = @()
                if (-not $grp.HasOwner) { $notes += 'no owner' }
                if ($grp.Empty)         { $notes += 'empty' }
                $noteText = if ($notes.Count) { "  [" + ($notes -join ', ') + "]" } else { '' }
                $mailText = if ($grp.MailOn) { 'mail-enabled' } else { 'no mail' }
                $blockIcon = if ($cb.Name -eq 'Security') { $iconSecurity } elseif ($cb.Name -eq 'M365') { $iconM365 } else { $null }
                $shapes += @{ Id="g$row`_$col"; Kind='Rectangle'; CenterText=$true; Icon=$blockIcon
                              Text="$($grp.Name)`n$($grp.Members) members  -  $mailText$noteText"
                              X=$x; Y=$cy; W=$nodeW; H=$nodeH; Fill=$fill; Line=$line; FontSize=9 }
            } else {
                $shapes += @{ Id="gmore$row`_$col"; Kind='Rectangle'; CenterText=$true
                              Text="+$($cb.Hidden) more groups"
                              X=$x; Y=$cy; W=$nodeW; H=$nodeH; Fill='#F2F2F2'; Line='#999999'; FontSize=9 }
            }
        }
        $rows = [math]::Max(1, [math]::Ceiling($total / $cols))
        $y -= ($rows - 1) * ($nodeH + $gapY) + $nodeH / 2 + $gapY + $bandH / 2
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Groups (all) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector $links -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: 1 tenant, $($data.Total) groups in $($catBlocks.Count) categories" -ForegroundColor DarkGray
    return $svg
}
