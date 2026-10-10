# Autodoc v1.0
function Build-UsersMapSvg {
    <#
        A single users overview with several sections:
          - a summary card: totals, UserType, UserCategory, enabled/disabled,
            no-mailbox count, and multi-method auth count
          - a band per breakdown: Domains, Job titles, Authentication methods
        Each breakdown lists its values with the number of users.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 4
    )

    $data = Get-UsersMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No user data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    $userIcon   = Get-MapIcon -IconSet $iconSet -Sku 'users' -IconMap $iconMap -PreferSvg

    $nodeW = 2.7; $nodeH = 0.95; $gapX = 0.28; $gapY = 0.5
    $tenantH = 1.35; $bandH = 0.6; $marginTop = 0.5; $marginBottom = 0.6
    $legendH = (Get-LegendHeight -LineCount 2)
    $cols = $PerRow

    # the sections to render as bands (name -> list of {Name,Count})
    $sections = @(
        @{ Label = 'User types';             Items = $data.Types }
        @{ Label = 'User categories';        Items = $data.Categories }
        @{ Label = 'Users per domain';       Items = $data.Domains }
        @{ Label = 'Authentication methods'; Items = $data.Auth }
        @{ Label = 'Job titles';             Items = $data.Titles }
    )

    $pageW = [math]::Max(11, $cols * ($nodeW + $gapX) + $gapX + 1)

    # shared "Data source" block (JSON path + field glossary) shown at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'UserInformation' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height

    # summary card height (matches the 4 lines actually drawn below)
    $summaryLines = 4
    $summaryH = 0.16 + $summaryLines * 0.2 + 0.06

    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + $summaryH + $gapY
    foreach ($s in $sections) {
        $rows = [math]::Max(1, [math]::Ceiling(@($s.Items).Count / $cols))
        $vSpan += $bandH + $gapY + $rows * $nodeH + ($rows - 1) * $gapY + $gapY
    }
    $vSpan += $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Users: $($data.Total)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Each block = '; Text = 'a value and the number of users it applies to'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=6.2; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    # summary card
    $summaryY = $legendY - $legendH / 2 - $gapY - $summaryH / 2
    $summaryLinesArr = @(
        @{ Text = 'Account summary'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Enabled: '; Text = "$($data.Enabled)    Disabled: $($data.Disabled)"; Align = 'start' }
        @{ BoldPrefix = 'Without mailbox: '; Text = "$($data.NoMail)"; Align = 'start' }
        @{ BoldPrefix = 'Use more than one auth method: '; Text = "$($data.MultiAuth)"; Align = 'start' }
    )
    $shapes += @{ Id='summary'; Kind='Rectangle'; Lines=$summaryLinesArr; LinesTop=$true; TopInset=0.12
                  X=$pageW/2; Y=$summaryY; W=6.2; H=$summaryH; Fill='#EEF2F7'; Line='#888888'; FontSize=9 }

    # one colour per band; value blocks below use a lighter shade of the same
    # hue so they visibly belong to their band
    $palette = @(
        @{ Band = '#F3D7E0'; Block = '#FBEDF2' }   # rose (User types)
        @{ Band = '#E7F0DC'; Block = '#F3F8EC' }   # green
        @{ Band = '#EADCF0'; Block = '#F6EEFA' }   # purple
        @{ Band = '#F6E6CC'; Block = '#FCF4E6' }   # amber
        @{ Band = '#D8EBE9'; Block = '#EEF7F6' }   # teal
    )

    $y = $summaryY - $summaryH / 2 - $gapY - $bandH / 2
    for ($si = 0; $si -lt $sections.Count; $si++) {
        $s = $sections[$si]
        $pal = $palette[$si % $palette.Count]
        $items = @($s.Items)
        $shapes += @{ Id="band-$($s.Label)"; Kind='Rectangle'; Lines=@(@{ Text="$($s.Label)  ($($items.Count))"; Bold=$true; Align='middle' }); TopInset=0.14
                      X=$pageW/2; Y=$y; W=4.0; H=$bandH; Fill=$pal.Band; Line='#666666'; FontSize=11 }
        $y -= $bandH / 2 + $gapY + $nodeH / 2

        $gridW = $cols * $nodeW + ($cols - 1) * $gapX
        $left  = ($pageW - $gridW) / 2
        # usable text width inside a value block (right of the user icon)
        $iconSide = [math]::Min($nodeH * 96 * 0.6, $nodeW * 96 * 0.34)
        $blockAvailIn = ($nodeW * 96 * 0.95 - $iconSide - 12) / 96 * 0.97
        for ($i = 0; $i -lt $items.Count; $i++) {
            $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            $cy = $y - $row * ($nodeH + $gapY)
            $it = $items[$i]
            $w = if ($it.Count -eq 1) { 'user' } else { 'users' }
            # left-aligned, bold value name (wrapped, max 2 lines), count below,
            # in the band's lighter shade; larger text than before for readability
            $nameLines = @(Split-MapTextToWidth -Text "$($it.Name)" -AvailIn $blockAvailIn -FontSize 10)
            if ($nameLines.Count -gt 2) { $nameLines = @($nameLines[0], ($nameLines[1].TrimEnd() + '…')) }
            $blockLines = @()
            foreach ($nl in $nameLines) { $blockLines += @{ Text = $nl; Bold = $true; Align = 'start' } }
            $blockLines += @{ Text = "$($it.Count) $w"; Bold = $false; Align = 'start' }
            $shapes += @{ Id="u$row`_$col`_$($s.Label)"; Kind='Rectangle'; Lines=$blockLines; LinesTop=$true; TopInset=0.10; Icon=$userIcon
                          X=$x; Y=$cy; W=$nodeW; H=$nodeH; Fill=$pal.Block; Line='#888888'; FontSize=10 }
        }
        $rows = [math]::Max(1, [math]::Ceiling($items.Count / $cols))
        $y -= ($rows - 1) * ($nodeH + $gapY) + $nodeH / 2 + $gapY + $bandH / 2
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Users (types, domains, auth, titles) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($data.Total) users ($($data.Enabled) enabled / $($data.Disabled) disabled)" -ForegroundColor DarkGray
    return $svg
}
