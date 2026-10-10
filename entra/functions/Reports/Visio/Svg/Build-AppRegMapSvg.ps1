# Autodoc v1.0
function Build-AppRegMapSvg {
    <#
        SVG map of app registrations as detail cards in a masonry layout. Each
        card shows the app name in bold, then every field the export carries
        (IDs, owner, secrets/certs and their expiry, credential health, SSO,
        audience, API permissions and names, sign-in activity). Card fill still
        signals credential health (red expired / amber expiring / green healthy),
        and a red border flags an over-permissioned app (many API permissions or
        Graph access). Cards vary in height, so they are packed into the shortest
        column for a tight fit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [int]$Columns = 3,
        [int]$ApiWarnThreshold = 10      # >= this many API permissions -> red edge
    )

    $data = Get-AppRegMapData -SourceFolder $SourceFolder
    $full = Get-AppFullData   -SourceFolder $SourceFolder
    if (-not $data -or -not $full) { Write-Host "No app registration data in this export." -ForegroundColor Yellow; return }
    $tenant = $data.Tenant

    # map health + over-priv from the trimmed data, keyed by app name
    $meta = @{}
    foreach ($aud in $data.Audiences) { foreach ($a in $aud.Group) { $meta[$a.Name] = $a } }

    $iconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $tenantIcon = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg
    # App-registration cards use the dedicated service icon file
    # (files\cache\Azure icons\Svg\10232-icon-service-App-Registrations.svg),
    # resolved by its file name via the SVG-preferred icon index.

    function Get-HealthFill([string]$h) {
        switch ($h) {
            'expired'  { return '#F8696B' }
            'expiring' { return '#FFEB84' }
            'healthy'  { return '#C8E6C9' }
            default    { return '#FFFFFF' }
        }
    }

    # Wrap a value to lines that fit a given width (inches) at a font size,
    # breaking on natural separators (space ; , / _ @ -) and hard-splitting any
    # single token that is by itself wider than a line. The first line can
    # reserve room for a bold field prefix (e.g. "Owner: ") via -FirstReserveIn.
    # Always returns at least one line.
    function Split-TextToWidth {
        param([string]$Text, [double]$AvailIn, [double]$FontSize, [double]$FirstReserveIn = 0)
        $Text = "$Text"
        $out = @()
        $cur = ''
        $firstBudget = $AvailIn - $FirstReserveIn
        $tokens = [regex]::Split($Text, '(?<=[\s;,/_@-])')
        foreach ($tok in $tokens) {
            $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
            while ((Measure-MapTextWidth $tok $FontSize) -gt $budget -and $tok.Length -gt 1) {
                $budget = if ($out.Count -eq 0 -and -not $cur) { $firstBudget } else { $AvailIn }
                $n = $tok.Length
                while ($n -gt 1 -and (Measure-MapTextWidth ($cur + $tok.Substring(0, $n)) $FontSize) -gt $budget) { $n-- }
                $out += ($cur + $tok.Substring(0, $n)).TrimEnd()
                $cur = ''
                $tok = $tok.Substring($n)
            }
            $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
            if ($cur -and (Measure-MapTextWidth ($cur + $tok) $FontSize) -gt $budget) {
                $out += $cur.TrimEnd(); $cur = $tok
            }
            else {
                $cur += $tok
            }
        }
        if ($cur) { $out += $cur.TrimEnd() }
        if ($out.Count -eq 0) { $out = @('') }
        return $out
    }

    # geometry
    $marginTop = 0.5; $marginBottom = 0.6; $tenantH = 1.25; $gapY = 0.6
    $legendH = (Get-LegendHeight -LineCount 4)
    $cardW = 3.4; $cardGapX = 0.35; $cardGapY = 0.35
    $lineH = 0.19; $titleH = 0.26; $cardPad = 0.14
    $cols = [math]::Max(1, $Columns)

    # Usable text width inside a card (inches) for wrapping. Text is drawn to
    # the RIGHT of the icon (see Export-MapAsSvg): the icon occupies W*0.34 at
    # W*0.05 from the left, with 6px padding on each side. Wrap to 97% of that
    # as a cushion for bold text and glyph-measurement rounding.
    $cardAvailIn = ($cardW - $cardW * 0.05 - $cardW * 0.34 - 2 * (6 / 96)) * 0.97

    $marginSide = 0.6
    $pageW = [math]::Max(11, $cols * $cardW + ($cols - 1) * $cardGapX + 2 * $marginSide)

    # shared "Data source" block (JSON path + field glossary) shown at the very top
    $dsW = [math]::Max(6.0, $pageW - 2 * $marginSide)
    $dsBlock = Get-MapDataSourceBlock -SourceFolder $SourceFolder -BaseName 'AppRegEnterpriseApps' -WidthIn $dsW -FontSize 10 -LineH $lineH -Pad $cardPad -GridFields
    $dsH = $dsBlock.Height

    # build each card's wrapped lines + height first. Long values (owner lists,
    # permission names, ...) are wrapped to the card's usable width so nothing
    # spills outside the block; the card then grows to fit the extra lines.
    $cards = @()
    foreach ($app in $full.Apps) {
        $detail = @(Get-AppDetailLines -App $app -View 'appreg')
        $m = $meta["$($app.AppName)"]
        $health = if ($m) { $m.Health } else { 'healthy' }
        $apiCount = [int]("$($app.ApiPermissionCount)".Trim() -replace '[^\d]', '')
        $usesGraph = ("$($app.UsesGraphPermissions)".Trim().ToLower() -eq 'yes')
        $overPriv = ($apiCount -ge $ApiWarnThreshold) -or $usesGraph
        $edge = if ($overPriv) { '#C81E1E' } else { '#B8BFC7' }

        # title (bold), wrapped if long
        $lines = @()
        foreach ($t in @(Split-TextToWidth -Text $app.AppName -AvailIn $cardAvailIn -FontSize 10)) {
            $lines += @{ Text = $t; Bold = $true; Align = 'start' }
        }
        # each "Field: value": bold field prefix stays on the first line, the
        # value wraps under it; the first line reserves room for the prefix.
        foreach ($d in $detail) {
            $i = $d.IndexOf(':')
            if ($i -gt 0) {
                $prefix  = $d.Substring(0, $i + 1) + ' '
                $value   = $d.Substring($i + 1).Trim()
                $reserve = Measure-MapTextWidth $prefix 8
                $parts   = @(Split-TextToWidth -Text $value -AvailIn $cardAvailIn -FontSize 10 -FirstReserveIn $reserve)
                $lines += @{ BoldPrefix = $prefix; Text = $parts[0]; Align = 'start' }
                for ($k = 1; $k -lt $parts.Count; $k++) { $lines += @{ Text = $parts[$k]; Bold = $false; Align = 'start' } }
            }
            else {
                foreach ($t in @(Split-TextToWidth -Text $d -AvailIn $cardAvailIn -FontSize 10)) {
                    $lines += @{ Text = $t; Bold = $false; Align = 'start' }
                }
            }
        }

        # height: first line billed at title height, the rest at line height
        $h = $cardPad * 2 + $titleH + ($lines.Count - 1) * $lineH
        $cards += [pscustomobject]@{
            Name = $app.AppName; Lines = $lines; Height = $h
            Fill = (Get-HealthFill $health); Edge = $edge
        }
    }

    # column masonry: place each card in the currently shortest column
    $gridW = $cols * $cardW + ($cols - 1) * $cardGapX
    $left  = ($pageW - $gridW) / 2
    $colX  = for ($c = 0; $c -lt $cols; $c++) { $left + $c * ($cardW + $cardGapX) + $cardW / 2 }
    $colBottom = @(0..($cols - 1) | ForEach-Object { 0.0 })   # consumed height per column

    # first pass to compute total height (place virtually)
    $tmpBottom = @(0..($cols - 1) | ForEach-Object { 0.0 })
    foreach ($card in $cards) {
        $ci = 0; for ($c = 1; $c -lt $cols; $c++) { if ($tmpBottom[$c] -lt $tmpBottom[$ci]) { $ci = $c } }
        $tmpBottom[$ci] += $card.Height + $cardGapY
    }
    $colBlockH = ($tmpBottom | Measure-Object -Maximum).Maximum
    if (-not $colBlockH) { $colBlockH = 1 }

    $pageH = [math]::Max(8.5, $marginTop + $dsH + $gapY + $tenantH + $gapY + $legendH + $gapY + $colBlockH + $marginBottom)

    $shapes = @()
    $topY   = $pageH - $marginTop - $tenantH / 2
    $dsY  = $topY - $tenantH / 2 - $gapY - $dsH / 2
    $today = Get-Date -Format 'yyyy-MM-dd'

    $shapes += @{ Id='datasource'; Kind='Rectangle'; Lines=$dsBlock.Lines; LinesTop=$true; TopInset=0.10; Columns=$dsBlock.Columns; GridFrom=$dsBlock.GridFrom
                  X=$pageW/2; Y=$dsY; W=$dsW; H=$dsH; Fill='#F7F7F7'; Line='#888888'; FontSize=10 }

    $expiredN = 0; $expiringN = 0
    foreach ($aud in $data.Audiences) { foreach ($a in $aud.Group) { if ($a.Health -eq 'expired') { $expiredN++ } elseif ($a.Health -eq 'expiring') { $expiringN++ } } }

    $tenantLines = if ($tenant) {
        @(
            @{ Text = "$($tenant.TenantName)"; Bold = $true; Align = 'start' }
            @{ Text = "$($tenant.PrimaryDomain)"; Bold = $false; Align = 'start' }
            @{ Text = (Get-TenantHybridLabel -Tenant $tenant); Bold = $false; Align = 'start' }
            @{ Text = "App registrations: $($data.Total)   (expired $expiredN / expiring $expiringN)"; Bold = $false; Align = 'start' }
            @{ Text = "Generated $today"; Bold = $false; Align = 'start' }
        )
    } else { @(@{ Text = 'Tenant'; Bold = $true; Align = 'start' }) }
    $shapes += @{ Id='tenant'; Kind='Rectangle'; Lines=$tenantLines
                  X=$pageW/2; Y=$topY; W=$dsW; H=$tenantH; Fill='#0F6CBD'; Line='#0B4C87'; FontSize=11; Icon=$tenantIcon }

    $legendY = $dsY - $dsH / 2 - $gapY - $legendH / 2
    $legendLines = @(
        @{ BoldPrefix = 'Red: ';   Text = 'a credential has expired';           Align = 'start' }
        @{ BoldPrefix = 'Amber: '; Text = 'a credential expires within 30 days'; Align = 'start' }
        @{ BoldPrefix = 'Green: '; Text = 'credentials healthy';                 Align = 'start' }
        @{ BoldPrefix = 'Red edge: '; Text = "$ApiWarnThreshold+ API permissions or Graph access"; Align = 'start' }
    )
    $shapes += @{ Id='legend'; Kind='Rectangle'; Lines=$legendLines; Dashed=$true
                  X=$pageW/2; Y=$legendY; W=6.0; H=$legendH; Fill='#FFFFFF'; Line='#888888'; FontSize=9 }

    # place the cards (real pass)
    $cardsTopY = $legendY - $legendH / 2 - $gapY
    $colTop = @(0..($cols - 1) | ForEach-Object { $cardsTopY })
    $idx = 0
    foreach ($card in $cards) {
        $ci = 0; for ($c = 1; $c -lt $cols; $c++) { if ($colTop[$c] -gt $colTop[$ci]) { $ci = $c } }
        $cy = $colTop[$ci] - $card.Height / 2
        $shapes += @{ Id="app$idx"; Kind='Rectangle'; Lines=$card.Lines; LinesTop=$true; TopInset=0.12
                      X=$colX[$ci]; Y=$cy; W=$cardW; H=$card.Height; Fill=$card.Fill; Line=$card.Edge; FontSize=8 }
        $colTop[$ci] = $colTop[$ci] - $card.Height - $cardGapY
        $idx++
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) { New-Item -Path $svgDir -ItemType Directory -Force | Out-Null }
    $svg = Join-Path $svgDir "Entra ID - App registrations $stamp.svg"
    Export-MapAsSvg -Path $svg -Shape $shapes -PageWidth $pageW -PageHeight $pageH | Out-Null

    Write-Host "  SVG: $($data.Total) app registrations, detail cards (expired $expiredN / expiring $expiringN)" -ForegroundColor DarkGray
    return $svg
}
