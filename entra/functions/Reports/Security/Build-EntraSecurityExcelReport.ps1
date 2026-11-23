# Autodoc v1.0
function Build-EntraSecurityExcelReport {
    <#
        "Security and improvements" workbook. Same content as the Word file,
        with every affected object listed:
          Summary          cards + proposals per area and severity (chart)
          Proposals        one row per proposal: finding, recommendation, reference
          Affected objects one row per object behind each proposal
          Hunt ...         workload identity hunt findings, triage, credentials
                           (only when the hunt has been run for the export)
        Input is the object from Get-SecurityImprovementData.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][string]$Path,
        [string]$SourceFolder
    )

    $cfg  = $Data.Config
    $lang = $Data.Language
    function TX($k) { $e = $cfg.Text[$k]; if ($e -and "$($e[$lang])") { "$($e[$lang])" } elseif ($e) { "$($e['en'])" } else { $k } }
    function SevName($s) { TX "Sev_$s" }
    $fill = @{ CRITICAL = 'F8D7DA'; HIGH = 'F3E5F5'; MEDIUM = 'FFF3CD'; INFO = 'E8F0FE' }

    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (Test-Path $Path) { Remove-Item -LiteralPath $Path -Force }

    $impr = @($Data.Improvements)
    $sheetSum  = TX 'Sheet_Summary'
    $sheetProp = TX 'Sheet_Proposals'
    $sheetObj  = TX 'Sheet_Objects'

    # ---- summary: area x severity table + stacked chart ----
    $areaRows = @(foreach ($area in $Data.AreaOrder) {
        $items = @($impr | Where-Object Area -eq $area)
        $row = [ordered]@{ (TX 'Area') = (@($cfg.Areas[$area][$lang], $cfg.Areas[$area]['en'], $area) | Where-Object { $_ } | Select-Object -First 1) }
        foreach ($s in $Data.SevOrder) { $row[(SevName $s)] = @($items | Where-Object Severity -eq $s).Count }
        $row[(TX 'Total')] = $items.Count
        [pscustomobject]$row
    })
    $startRow = 8
    $n = $areaRows.Count
    $yRanges = @(for ($i = 0; $i -lt $Data.SevOrder.Count; $i++) {
        $L = ConvertTo-ExcelColumnLetter -Index (2 + $i)
        "'$sheetSum'!`$$L`$$($startRow + 1):`$$L`$$($startRow + $n)"
    })
    $chart = New-ExcelChartDefinition -Title (TX 'Chart_Areas') -ChartType ColumnStacked `
        -XRange "'$sheetSum'!`$A`$$($startRow + 1):`$A`$$($startRow + $n)" -YRange $yRanges `
        -SeriesHeader @($Data.SevOrder | ForEach-Object { SevName $_ }) -Row ($startRow - 1) -Column 8 -Width 760 -Height 420
    $areaRows | Export-Excel -Path $Path -WorksheetName $sheetSum -StartRow $startRow -TableName 'tblAreas' -TableStyle Medium2 -ExcelChartDefinition $chart

    # ---- proposals ----
    $propRows = @($impr | ForEach-Object {
        [pscustomobject][ordered]@{
            (TX 'Severity')       = (SevName $_.Severity)
            (TX 'Area')           = $_.AreaName
            (TX 'Proposal')       = $_.Title
            (TX 'Count')          = $_.Count
            (TX 'Finding')        = $_.Finding
            (TX 'Recommendation') = $_.Recommendation
            (TX 'Reference')      = $_.Reference
        }
    })
    if ($propRows.Count) { $propRows | Export-Excel -Path $Path -WorksheetName $sheetProp -TableName 'tblProposals' -TableStyle Medium2 -FreezeTopRow }

    # ---- affected objects ----
    $objRows = @(foreach ($p in $impr) {
        foreach ($o in @($p.Objects)) {
            [pscustomobject][ordered]@{
                (TX 'Severity') = (SevName $p.Severity)
                (TX 'Area')     = $p.AreaName
                (TX 'Proposal') = $p.Title
                (TX 'Object')   = "$($o.Name)"
                (TX 'Detail')   = "$($o.Detail)"
            }
        }
    })
    if ($objRows.Count) { $objRows | Export-Excel -Path $Path -WorksheetName $sheetObj -TableName 'tblObjects' -TableStyle Medium2 -FreezeTopRow }

    # ---- workload identity hunt ----
    $hunt = $Data.Hunt
    $huntSheets = @()
    if ($hunt.Ran) {
        foreach ($h in @(
            @{ Rows = $hunt.Findings;    Sheet = (TX 'Sheet_HuntFindings'); Table = 'tblHuntFindings' }
            @{ Rows = $hunt.Triage;      Sheet = (TX 'Sheet_HuntTriage');   Table = 'tblHuntTriage' }
            @{ Rows = $hunt.Credentials; Sheet = (TX 'Sheet_HuntCreds');    Table = 'tblHuntCredentials' })) {
            if (@($h.Rows).Count) {
                @($h.Rows) | Export-Excel -Path $Path -WorksheetName $h.Sheet -TableName $h.Table -TableStyle Medium2 -FreezeTopRow
                $huntSheets += $h.Sheet
            }
        }
    }

    # ---- formatting pass ----
    $xl = Open-ExcelPackage -Path $Path
    $ws = $xl.Workbook.Worksheets[$sheetSum]
    $huntTop = @($hunt.Findings | Where-Object { $_.Severity -in 'CRITICAL', 'HIGH' }).Count
    $cards = @(@{ Label = (TX 'Total'); Value = $impr.Count })
    foreach ($s in $Data.SevOrder) { $cards += @{ Label = (SevName $s); Value = @($impr | Where-Object Severity -eq $s).Count } }
    $cards += @{ Label = (TX 'HuntFindings'); Value = $(if ($hunt.Ran) { @($hunt.Findings).Count } else { '-' }) }
    Add-KpiCards -Worksheet $ws -Title (TX 'FileName') -Cards $cards
    $ws.Cells['A6'].Value = ((TX 'Excel_Note') -f $(if ($SourceFolder) { Split-Path $SourceFolder -Leaf } else { '' }))
    $ws.Cells['A6'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
    $ws.Column(1).Width = 32
    # chart series in the severity colours (same as the map and Word)
    $barColor = @{ CRITICAL = '#C62828'; HIGH = '#8E24AA'; MEDIUM = '#F0B44C'; INFO = '#7A8C9E' }
    foreach ($drawing in $ws.Drawings) {
        try {
            for ($i = 0; $i -lt $drawing.Series.Count -and $i -lt $Data.SevOrder.Count; $i++) {
                $drawing.Series[$i].Fill.Color = [System.Drawing.ColorTranslator]::FromHtml($barColor[$Data.SevOrder[$i]])
            }
            try { $drawing.Overlap = 100 } catch { }
        } catch { }
    }

    # severity colour on every sheet that has a severity column
    foreach ($sheet in @($sheetProp, $sheetObj) + $huntSheets) {
        $w = $xl.Workbook.Worksheets[$sheet]
        if (-not $w -or -not $w.Dimension) { continue }
        $sevCol = 0
        for ($c = 1; $c -le $w.Dimension.End.Column; $c++) { if ("$($w.Cells[1, $c].Value)" -in (TX 'Severity'), 'Severity') { $sevCol = $c; break } }
        if ($sevCol) {
            $byName = @{}; foreach ($s in $Data.SevOrder) { $byName[(SevName $s)] = $s; $byName[$s] = $s }
            for ($r = 2; $r -le $w.Dimension.End.Row; $r++) {
                $key = $byName["$($w.Cells[$r, $sevCol].Value)"]
                if ($key) {
                    $cell = $w.Cells[$r, $sevCol]
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($fill[$key])"))
                    $cell.Style.Font.Bold = $true
                }
            }
        }
        # readable widths: long text columns wrap, short ones fit
        for ($c = 1; $c -le $w.Dimension.End.Column; $c++) {
            $max = 0
            for ($r = 1; $r -le [math]::Min($w.Dimension.End.Row, 400); $r++) { $len = "$($w.Cells[$r, $c].Value)".Length; if ($len -gt $max) { $max = $len } }
            $w.Column($c).Width = [math]::Min([math]::Max($max + 2, 10), 70)
            if ($max -gt 70) { $w.Column($c).Style.WrapText = $true }
        }
        $w.Cells[1, 1, $w.Dimension.End.Row, $w.Dimension.End.Column].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    }
    $xl.Workbook.Worksheets.MoveToStart($sheetSum)
    $xl.Workbook.View.ActiveTab = 0
    Close-ExcelPackage $xl
    return $Path
}
