# Autodoc v1.0
function Build-LimitsMapSvg {
    <#
        Builds the "Service limits & recommendations" SVG: one horizontal bar
        per tracked metric, showing how the tenant's current count sits against
        the documented Microsoft Entra limit. The track is the limit (100 %),
        the coloured fill is current usage, so a short bar means plenty of
        headroom and a long/red bar means the limit is close.

        Limits come from Config/EntraServiceLimits.json (from Microsoft Learn);
        current values are counted from the exported JSON by Get-EntraLimitStatus.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $rows = Get-EntraLimitStatus -SourceFolder $SourceFolder
    if (-not $rows -or @($rows).Count -eq 0) { Write-Host "No limit data for this export." -ForegroundColor Yellow; return }
    $rows = @($rows)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $today  = Get-Date -Format 'yyyy-MM-dd'

    # header icon (falls back to no icon if the file isn't in the icon set)
    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $limitsIcon = Get-MapIcon -IconSet $iconSet -Sku 'service-limits-recommendations' -IconMap $iconMap -PreferSvg

    # ---- layout (inches) ----
    $marginX = 0.5
    $labelX0 = 0.5;  $labelW = 3.5
    $trackX0 = 4.2;  $trackW = 6.8
    $valueX0 = $trackX0 + $trackW + 0.2; $valueW = 2.2
    $pageW   = $valueX0 + $valueW + 0.3
    $barH    = 0.32; $rowStep = 0.64
    $minBar  = 0.05          # smallest visible sliver so tiny usage still shows

    $n = $rows.Count
    # vertical plan from the top
    $mTop = 0.4; $tenantH = 1.4; $gapNote = 0.2; $noteH = 0.68; $gap1 = 0.45; $titleH = 0.55; $gap2 = 0.2
    $rowsSpan = $n * $rowStep
    $gap3 = 0.35; $legendH = 0.35; $mBottom = 0.4
    $pageH = $mTop + $tenantH + $gapNote + $noteH + $gap1 + $titleH + $gap2 + $rowsSpan + $gap3 + $legendH + $mBottom

    function TopY([double]$offsetFromTop) { return $pageH - $offsetFromTop }

    $shapes = @()

    # ---- tenant / title box ----
    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)";                 Bold = $true;  Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)";              Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "Service limits & recommendations";      Bold = $true;  Align = 'start' }
            @{ Text = "Generated $today";                      Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Service limits & recommendations'; Bold = $true; Align = 'start' }) }
    $tenantCY = TopY ($mTop + $tenantH / 2)
    $shapes += @{ Id = 'tenant'; Kind = 'Rectangle'; Lines = $tenantLines;
                  X = $pageW / 2; Y = $tenantCY; W = $pageW - 2 * $marginX; H = $tenantH;
                  Fill = '#0F6CBD'; Line = '#0B4C87'; FontSize = 11; Icon = $limitsIcon }

    # ---- chart title ----
    $chartTop = $mTop + $tenantH + $gapNote + $noteH + $gap1
    $titleCY  = TopY ($chartTop + $titleH / 2)
    $shapes += @{ Id = 'charttitle'; Kind = 'Rectangle'; Fill = 'none'; Line = 'none';
                  Lines = @(
                      @{ Text = 'Current usage vs documented Microsoft Entra limit'; Bold = $true; Align = 'start' }
                  );
                  X = $labelX0 + ($pageW - 2 * $marginX) / 2; Y = $titleCY; W = $pageW - 2 * $marginX; H = $titleH; FontSize = 10 }

    # ---- threshold guide lines (60 % amber, 85 % red) spanning the rows ----
    $rowsTop    = $chartTop + $titleH + $gap2
    $rowsBottom = $rowsTop + $rowsSpan
    $guideCY    = TopY (($rowsTop + $rowsBottom) / 2)
    $guideH     = $rowsSpan
    foreach ($g in @(@{ P = 0.60; C = '#F9A825' }, @{ P = 0.85; C = '#C62828' })) {
        $gx = $trackX0 + $trackW * $g.P
        $shapes += @{ Id = "guide$($g.P)"; Kind = 'Rectangle'; X = $gx; Y = $guideCY; W = 0.012; H = $guideH;
                      Fill = $g.C; Line = 'none' }
    }

    # ---- one row per metric ----
    for ($i = 0; $i -lt $n; $i++) {
        $r = $rows[$i]
        $rowCY = TopY ($rowsTop + $rowStep * $i + $rowStep / 2)

        # metric label (name + type), left aligned
        $shapes += @{ Id = "lbl$i"; Kind = 'Rectangle'; Fill = 'none'; Line = 'none';
                      Lines = @(
                          @{ Text = "$($r.Metric)"; Bold = $true;  Align = 'start' }
                          @{ Text = "$($r.Type)";   Bold = $false; Align = 'start' }
                      );
                      X = $labelX0 + $labelW / 2; Y = $rowCY; W = $labelW; H = $barH + 0.18; FontSize = 9 }

        # track (the limit = 100 %)
        $shapes += @{ Id = "trk$i"; Kind = 'Rectangle'; X = $trackX0 + $trackW / 2; Y = $rowCY; W = $trackW; H = $barH;
                      Fill = '#ECECEC'; Line = '#D6D6D6'; StrokeWidth = '0.8' }

        # fill (current usage), min sliver so tiny values remain visible
        $frac   = [math]::Min([double]$r.PercentUsed / 100, 1)
        $barLen = [math]::Max($minBar, $trackW * $frac)
        $shapes += @{ Id = "bar$i"; Kind = 'Rectangle'; X = $trackX0 + $barLen / 2; Y = $rowCY; W = $barLen; H = $barH;
                      Fill = $r.Color; Line = 'none' }

        # value label (current / limit, percent + status)
        $curTxt = '{0:N0}' -f $r.Current
        $limTxt = '{0:N0}' -f $r.Limit
        $shapes += @{ Id = "val$i"; Kind = 'Rectangle'; Fill = 'none'; Line = 'none';
                      Lines = @(
                          @{ Text = "$curTxt / $limTxt";               Bold = $true;  Align = 'start' }
                          @{ Text = "$($r.PercentUsed)% · $($r.Status)"; Bold = $false; Align = 'start' }
                      );
                      X = $valueX0 + $valueW / 2; Y = $rowCY; W = $valueW; H = $barH + 0.18; FontSize = 9 }
    }

    # ---- legend ----
    $legendCY = TopY ($rowsBottom + $gap3 + $legendH / 2)
    $legendLine = @(
        @{ Text = 'Fill = current usage. Track = documented limit.  '; Bold = $false }
        @{ Text = 'OK'; Bold = $true }, @{ Text = ' < 60%   ' }
        @{ Text = 'Watch'; Bold = $true }, @{ Text = ' 60-85%   ' }
        @{ Text = 'Near/Over'; Bold = $true }, @{ Text = ' >= 85% (amber/red guides).' }
    )
    $legendText = ($legendLine | ForEach-Object { $_.Text }) -join ''
    $shapes += @{ Id = 'legend'; Kind = 'Rectangle'; Fill = 'none'; Line = 'none';
                  Lines = @(@{ Text = $legendText; Bold = $false; Align = 'start' });
                  X = $labelX0 + ($pageW - 2 * $marginX) / 2; Y = $legendCY; W = $pageW - 2 * $marginX; H = $legendH; FontSize = 8.5 }

    # ---- source note (top, under the header) ----
    $noteCY = TopY ($mTop + $tenantH + $gapNote + $noteH / 2)
    $shapes += @{ Id = 'note'; Kind = 'Rectangle'; Fill = '#F7F7F7'; Line = '#CCCCCC';
                  Lines = @(
                      @{ BoldPrefix = 'Limits: '; Text = 'Microsoft Entra service limits and restrictions (Microsoft Learn), via Config\EntraServiceLimits.json.'; Align = 'start' }
                      @{ Text = 'Current values counted from the exported JSON for this tenant.'; Bold = $false; Align = 'start' }
                  );
                  X = $labelX0 + ($pageW - 2 * $marginX) / 2; Y = $noteCY; W = $pageW - 2 * $marginX; H = $noteH; FontSize = 10 }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - Service limits & recommendations $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -Connector @() -PageWidth $pageW -PageHeight $pageH | Out-Null

    $flagged = @($rows | Where-Object { $_.Status -ne 'OK' }).Count
    Write-Host "  SVG: $n metrics vs Entra limits, $flagged above 60%." -ForegroundColor DarkGray
    return $svg
}
