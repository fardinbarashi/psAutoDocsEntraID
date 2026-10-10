# Autodoc v1.0
function Build-DomainMapSvg {
    <#
        SVG map of the tenant's verified domains, banded by Type (Managed,
        Federated, None). Each domain is a block; the default domain gets a
        distinct fill and a [default] tag. Colour by type keeps the bands visually
        distinct.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$PerRow = 4
    )

    $data = Get-DomainMapData -SourceFolder $SourceFolder
    if (-not $data) { Write-Host "No domain data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant; $types = $data.Types

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    # Per-band service icons, resolved by file name via the SVG-preferred index:
    #   Managed   -> 10222-icon-service-Entra-Domain-Services.svg
    #   Federated -> 10064-icon-service-DNS-Zones.svg
    $managedIcon   = Get-MapIcon -IconSet $iconSet -Sku '10222-icon-service-Entra-Domain-Services' -IconMap $iconMap -PreferSvg
    $federatedIcon = Get-MapIcon -IconSet $iconSet -Sku '10064-icon-service-DNS-Zones' -IconMap $iconMap -PreferSvg

    # type -> band fill
    function Get-TypeFill([string]$t) {
        switch ($t) {
            'Managed'   { return '#DCE7F3' }
            'Federated' { return '#E7F0DC' }
            'None'      { return '#F0F0F0' }
            default     { return '#F0F0F0' }
        }
    }
    # type -> domain-block fill: a lighter shade of the band colour, so every
    # block in a row is coloured by its type. The default domain gets the darker
    # band colour so it still stands out within its row.
    function Get-TypeBlockFill([string]$t, [bool]$isDefault) {
        switch ($t) {
            'Managed'   { if ($isDefault) { return '#DCE7F3' } else { return '#EAF1FA' } }
            'Federated' { if ($isDefault) { return '#E7F0DC' } else { return '#F1F6E9' } }
            'None'      { if ($isDefault) { return '#E4E4E4' } else { return '#F7F7F7' } }
            default     { if ($isDefault) { return '#E4E4E4' } else { return '#F7F7F7' } }
        }
    }

    $nodeW = 2.7; $nodeH = 0.85; $gapX = 0.28; $gapY = 0.55
    $tenantH = 1.35; $bandH = 0.6; $marginTop = 0.5; $marginBottom = 0.6
    $legendH = (Get-LegendHeight -LineCount 3)
    $cols = $PerRow

    $pageW = [math]::Max(11, $cols * ($nodeW + $gapX) + $gapX + 1)

    # shared "Data source" block (JSON path + field glossary) shown at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'TenantInformationDomain' -WidthIn $dsW -FontSize 10 -GridFields -GridMaxColumns 3
    $dsH = $dsBlock.Height
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY
    foreach ($ty in $types) {
        $rows = [math]::Max(1, [math]::Ceiling(@($ty.Group).Count / $cols))
        $vSpan += $bandH + $gapY + $rows * $nodeH + ($rows - 1) * $gapY + $gapY
    }
    $vSpan += $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10; Columns=$dsBlock.Columns; GridFrom=$dsBlock.GridFrom
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Domains: $($data.Total), Default: $($data.Default)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ Text = 'How to read this map'; Bold = $true; Align = 'start' }
        @{ BoldPrefix = 'Bands & blocks = '; Text = 'coloured by domain type (Managed, Federated, None)'; Align = 'start' }
        @{ BoldPrefix = '[default] tag = '; Text = 'the default domain'; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=6.0; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    $y = $legendY - $legendH / 2 - $gapY - $bandH / 2
    foreach ($ty in $types) {
        $items = @($ty.Group)
        $bandIcon = switch ($ty.Name) { 'Managed' { $managedIcon } 'Federated' { $federatedIcon } default { $null } }
        $shapes += @{ Id="band-$($ty.Name)"; Kind='Rectangle'; Lines=@(@{ Text="$($ty.Name)  ($($items.Count))"; Bold=$true; Align='middle' }); TopInset=0.14; Icon=$bandIcon
                      X=$pageW/2; Y=$y; W=3.0; H=$bandH; Fill=(Get-TypeFill $ty.Name); Line='#666666'; FontSize=11 }
        $y -= $bandH / 2 + $gapY + $nodeH / 2

        $gridW = $cols * $nodeW + ($cols - 1) * $gapX
        $left  = ($pageW - $gridW) / 2
        for ($i = 0; $i -lt $items.Count; $i++) {
            $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            $cy = $y - $row * ($nodeH + $gapY)
            $dom = $items[$i]
            $tag = if ($dom.IsDefault) { '  [default]' } else { '' }
            $defaultText = if ($dom.IsDefault) { 'Yes' } else { 'No' }
            # all parameters in the block: domain (bold), type, default flag
            $blockLines = @(
                @{ Text = "$($dom.Name)$tag"; Bold = $true; Align = 'start' }
                @{ BoldPrefix = 'Type: ';    Text = $dom.Type; Align = 'start' }
                @{ BoldPrefix = 'Default: '; Text = $defaultText; Align = 'start' }
            )
            $shapes += @{ Id="d$row`_$col"; Kind='Rectangle'; Lines=$blockLines; LinesTop=$true; TopInset=0.10
                          X=$x; Y=$cy; W=$nodeW; H=$nodeH; Fill=(Get-TypeBlockFill $dom.Type $dom.IsDefault); Line='#888888'; FontSize=9 }
        }
        $rows = [math]::Max(1, [math]::Ceiling($items.Count / $cols))
        $y -= ($rows - 1) * ($nodeH + $gapY) + $nodeH / 2 + $gapY + $bandH / 2
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Domains $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($data.Total) domains ($(@($types).Count) types)" -ForegroundColor DarkGray
    return $svg
}
