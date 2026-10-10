# Autodoc v1.0
function Build-LicenceMapSvg {
    <#
        Builds ONLY the SVG licence map. Independent of the Visio builder, so the
        picture the user approved is never disturbed by Visio-side changes.

        SVG can do per-line bold within one shape, so the tenant and legend use
        the rich Lines model (no split shapes needed). Icons appear here.
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
    $tenantIcon  = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    $genericIcon = Get-MapIcon -IconSet $iconSet -Sku 'SKU' -IconMap $iconMap -PreferSvg

    $nodeW = 2.55; $nodeH = 0.85; $gapX = 0.28; $gapY = 0.65
    $tenantH = 1.25; $bandH = 0.6; $marginTop = 0.5; $marginBottom = 0.6; $legendH = (Get-LegendHeight -LineCount 4)
    $cols  = [math]::Min($PerRow, [math]::Max($purchased.Count, 1))
    $pageW = [math]::Max(11, [math]::Max(5.6, $cols * ($nodeW + $gapX) + $gapX + 1))

    # shared "Data source" block (JSON path only) at the very top
    $dsW = [math]::Max(6.0, $pageW - 1)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'LicensesInformation' -WidthIn $dsW -FontSize 10 -PathOnly
    $dsH = $dsBlock.Height
    $rowsP = if ($purchased) { [math]::Ceiling($purchased.Count / $cols) } else { 0 }
    $rowsV = if ($viral)     { [math]::Ceiling($viral.Count / $cols) }     else { 0 }
    $vSpan = $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY
    if ($purchased) { $vSpan += $bandH + $gapY + $rowsP * $nodeH + ($rowsP - 1) * $gapY + $gapY }
    if ($viral)     { $vSpan += $bandH + $gapY + $rowsV * $nodeH + ($rowsV - 1) * $gapY }
    $vSpan += $marginBottom
    $pageH = [math]::Max(8.5, $vSpan)

    $shapes = @(); $links = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    # tenant: one shape, per-line bold via Lines (SVG handles it)
    $allSkus = @($purchased) + @($viral)
    $enabledCount = @($allSkus | Where-Object { $_.Status -eq 'Enabled' }).Count
    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)";          Bold = $true;  Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)";       Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Licences: $enabledCount of $($allSkus.Count) Enabled"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today";               Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    # colour legend: one dashed shape, bold keyword prefix per line
    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ BoldPrefix = 'Numbers: '; Text = 'Used / Total  (percent used)'; Align = 'start' }
        @{ BoldPrefix = 'Red: ';     Text = '10% or fewer licences free';   Align = 'start' }
        @{ BoldPrefix = 'Yellow: ';  Text = '11-30% of licences free';      Align = 'start' }
        @{ BoldPrefix = 'Green: ';   Text = '31-100% of licences free';     Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=5.2; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    function Add-BandS([string]$BandId, [string]$BandText, [double]$BandY, [string]$BandFill) {
        @{ Id=$BandId; Kind='Rectangle'; Lines=@(@{ Text=$BandText; Bold=$true; Align='middle' }); TopInset=0.14
           X=$pageW/2; Y=$BandY; W=2.6; H=0.6; Fill=$BandFill; Line='#666666'; FontSize=11 }
    }

    $y = $legendY - $legendH / 2 - $gapY - $bandH / 2
    if ($purchased) {
        $shapes += (Add-BandS 'band-p' "Purchased  ($($purchased.Count))" $y '#DCE7F3')
        $y -= $bandH / 2 + $gapY + $nodeH / 2
        $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
        for ($i = 0; $i -lt $purchased.Count; $i++) {
            $p = $purchased[$i]; $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            # icon present -> text stays left of the icon (SVG renderer handles this)
            $shapes += @{ Id="p$i"; Kind='Rectangle'; Text="$($p.Sku)`n$($p.Consumed)/$($p.Enabled)  ($($p.Util)%)"
                          X=$x; Y=($y - $row * ($nodeH + $gapY)); W=$nodeW; H=$nodeH
                          Fill=(Get-LicenceFreeFill $p.PercentFree); Line='#555555'; FontSize=9; Icon=$genericIcon }
        }
        $y -= ($rowsP - 1) * ($nodeH + $gapY) + $nodeH / 2 + $gapY + $bandH / 2
    }
    if ($viral) {
        $shapes += (Add-BandS 'band-v' "Free / viral  ($($viral.Count))" $y '#EEEEEE')
        $y -= $bandH / 2 + $gapY + $nodeH / 2
        $gridW = $cols * $nodeW + ($cols - 1) * $gapX; $left = ($pageW - $gridW) / 2
        for ($i = 0; $i -lt $viral.Count; $i++) {
            $v = $viral[$i]; $row = [math]::Floor($i / $cols); $col = $i % $cols
            $x = $left + $col * ($nodeW + $gapX) + $nodeW / 2
            # no icon -> centre the text
            $shapes += @{ Id="v$i"; Kind='Rectangle'; Text="$($v.Sku)`n$($v.Consumed) used"; CenterText=$true
                          X=$x; Y=($y - $row * ($nodeH + $gapY)); W=$nodeW; H=$nodeH
                          Fill='#F2F2F2'; Line='#999999'; FontSize=9 }
        }
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Licenses (SKUs) $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector $links -PageWidth $pageW -PageHeight $pageH | Out-Null

    $matched = @($shapes | Where-Object { $_.Icon }).Count
    Write-Host "  SVG: 1 tenant, $($purchased.Count) purchased, $($viral.Count) free/viral, $matched icons" -ForegroundColor DarkGray
    return $svg
}
