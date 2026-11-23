# Autodoc v1.0
function Build-EntraSecurityMapSvg {
    <#
        "Security and improvements" map: a header box with the tenant, then one
        card per area. The card colour is the highest severity in the area; the
        card lists the count per severity and the proposals in that area.
        Input is the object from Get-SecurityImprovementData, so the map shows
        exactly what the Word and Excel files show.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path
    )

    $cfg  = $Data.Config
    $lang = $Data.Language
    function TX($k) { $e = $cfg.Text[$k]; if ($e -and "$($e[$lang])") { "$($e[$lang])" } elseif ($e) { "$($e['en'])" } else { $k } }
    function AreaName($k) { $e = $cfg.Areas[$k]; if ($e -and "$($e[$lang])") { "$($e[$lang])" } else { $k } }

    $pale   = @{ CRITICAL = '#F8D7DA'; HIGH = '#F3E5F5'; MEDIUM = '#FFF3CD'; INFO = '#E8F0FE'; OK = '#E6F4EA' }
    $strong = @{ CRITICAL = '#C62828'; HIGH = '#8E24AA'; MEDIUM = '#B26A00'; INFO = '#607080'; OK = '#2E7D32' }

    $iconFolder = Get-MapIconFolder -BuilderRoot (Join-Path (Split-Path $PSScriptRoot) 'Svg')
    $iconSet    = Get-MapIconSet -Folder $iconFolder
    $iconMap    = Get-IconMap
    $hubIcon    = Get-MapIcon -IconSet $iconSet -Sku 'Azure-Active-Directory' -IconMap $iconMap -PreferSvg

    # ---- one card per area, in config order ----
    $cards = @(foreach ($area in $Data.AreaOrder) {
        $items = @($Data.Improvements | Where-Object { $_.Area -eq $area })
        $worst = 'OK'
        foreach ($s in $Data.SevOrder) { if (@($items | Where-Object Severity -eq $s).Count) { $worst = $s; break } }
        $counts = (@($Data.SevOrder | ForEach-Object { "$(TX "Sev_$_") $(@($items | Where-Object Severity -eq $_).Count)" })) -join '  ·  '
        $lines = @(
            @{ Text = (AreaName $area); Bold = $true; Align = 'start' }
            @{ Text = $(if ($items.Count) { $counts } else { TX 'Sev_OK' }); Bold = $false; Align = 'start' }
        )
        foreach ($it in ($items | Select-Object -First 6)) {
            $t = "$($it.Title)"
            if ($t.Length -gt 46) { $t = $t.Substring(0, 45) + '…' }
            $lines += @{ BoldPrefix = "$(TX "Sev_$($it.Severity)"):"; Text = ([char]0x00A0 + $t); Align = 'start' }
        }
        if ($items.Count -gt 6) { $lines += @{ Text = "+ $($items.Count - 6)"; Bold = $false; Align = 'start' } }
        [pscustomobject]@{ Area = $area; Worst = $worst; Lines = $lines }
    })

    # ---- layout (inches, bottom-left origin) ----
    $cols = 3; $cardW = 4.3; $gapX = 0.3; $gapY = 0.3; $marginX = 0.5
    $lineH = 0.19
    $maxLines = (@($cards | ForEach-Object { $_.Lines.Count }) | Measure-Object -Maximum).Maximum
    $cardH = [math]::Max(1.2, 0.35 + $maxLines * $lineH)
    $rowsN = [math]::Ceiling($cards.Count / $cols)
    $pageW = 2 * $marginX + $cols * $cardW + ($cols - 1) * $gapX
    $mTop = 0.4; $headH = 1.4; $gap1 = 0.4; $legendH = 0.4; $mBottom = 0.4
    $pageH = $mTop + $headH + $gap1 + $rowsN * $cardH + ($rowsN - 1) * $gapY + 0.3 + $legendH + $mBottom
    function TopY([double]$o) { $pageH - $o }

    $shapes = @()
    $t = $Data.Tenant
    $hubLines = @(
        @{ Text = $(if ($t) { "$($t.TenantName)" } else { 'Microsoft Entra ID' }); Bold = $true; Align = 'start' }
        @{ Text = $(if ($t) { "$($t.PrimaryDomain)" } else { '' }); Bold = $false; Align = 'start' }
        @{ Text = (TX 'Map_Title'); Bold = $true; Align = 'start' }
        @{ Text = ("$(TX 'Total'): $(@($Data.Improvements).Count)   " + ((@($Data.SevOrder | ForEach-Object { "$(TX "Sev_$_") $(@($Data.Improvements | Where-Object Severity -eq $_).Count)" })) -join '  ·  ')); Bold = $false; Align = 'start' }
        @{ Text = "$(Get-Date -Format 'yyyy-MM-dd')"; Bold = $false; Align = 'start' }
    )
    $shapes += @{ Id = 'hub'; Kind = 'Rectangle'; Lines = $hubLines; X = $pageW / 2; Y = (TopY ($mTop + $headH / 2)); W = $pageW - 2 * $marginX; H = $headH
                  Fill = '#E8F0FE'; Line = '#0F6CBD'; StrokeWidth = '2'; FontSize = 12; Icon = $hubIcon }

    for ($i = 0; $i -lt $cards.Count; $i++) {
        $c = $cards[$i]
        $col = $i % $cols; $row = [math]::Floor($i / $cols)
        $x = $marginX + $col * ($cardW + $gapX) + $cardW / 2
        $yTop = $mTop + $headH + $gap1 + $row * ($cardH + $gapY)
        $shapes += @{ Id = "area$i"; Kind = 'Rectangle'; Lines = $c.Lines; LinesTop = $true; TopInset = 0.02
                      X = $x; Y = (TopY ($yTop + $cardH / 2)); W = $cardW; H = $cardH
                      Fill = $pale[$c.Worst]; Line = $strong[$c.Worst]; StrokeWidth = '2.5'; FontSize = 10 }
    }

    # legend: the rule, then one swatch + label per level
    $legendY = TopY ($pageH - $mBottom - $legendH / 2)
    $shapes += @{ Id = 'legend'; Kind = 'Rectangle'; Fill = 'none'; Line = 'none'; Lines = @(@{ Text = (TX 'Map_Legend'); Align = 'start' })
                  X = $marginX + 2.0; Y = $legendY; W = 4.0; H = $legendH; FontSize = 9 }
    $sw = 0.22; $kx = $marginX + 4.3
    foreach ($lvl in 'CRITICAL', 'HIGH', 'MEDIUM', 'INFO', 'OK') {
        $shapes += @{ Id = "key$lvl"; Kind = 'Rectangle'; X = $kx + $sw / 2; Y = $legendY; W = $sw; H = $sw
                      Fill = $pale[$lvl]; Line = $strong[$lvl]; StrokeWidth = '2' }
        $label = TX "Sev_$lvl"
        $lw = 0.25 + 0.075 * $label.Length
        $shapes += @{ Id = "keyl$lvl"; Kind = 'Rectangle'; Fill = 'none'; Line = 'none'; Lines = @(@{ Text = $label; Align = 'start' })
                      X = $kx + $sw + 0.05 + $lw / 2; Y = $legendY; W = $lw; H = $legendH; FontSize = 9 }
        $kx += $sw + 0.1 + $lw + 0.2
    }

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Export-MapAsSvg -Path $Path -Shape $shapes -Connector @() -PageWidth $pageW -PageHeight $pageH | Out-Null
    return $Path
}
