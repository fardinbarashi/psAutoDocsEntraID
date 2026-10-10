# Autodoc v1.0
function Add-KpiSheetLimits {
    <#
        "Limits & recommendations" sheet: the tenant's current counts against the
        documented Microsoft Entra service limits (Config/EntraServiceLimits.json,
        from Microsoft Learn). A bar chart shows percent-of-limit used per metric,
        a table lists the detail, and the Status cells are colour-coded.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [string]$SheetName = '13 Limits & recommendations',
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $rows = Get-EntraLimitStatus -SourceFolder $SourceFolder
    if (-not $rows -or @($rows).Count -eq 0) {
        Write-Host "  13 Limits & recommendations skipped - no data." -ForegroundColor DarkYellow
        return $false
    }
    $rows = @($rows)

    $srcCol    = 1
    $detailCol = 12   # column L, beside the chart

    # --- chart source table (left): metric + percent of limit used ---
    $chartTbl = @($rows | ForEach-Object {
        [pscustomobject]@{
            'Metric'         = $_.Metric
            '% of limit'     = $_.PercentUsed
        }
    })

    # --- detail table (right, beside the chart) ---
    $detail = @($rows | ForEach-Object {
        [pscustomobject]@{
            'Area'    = $_.Area
            'Metric'  = $_.Metric
            'Current' = $_.Current
            'Limit'   = $_.Limit
            '% used'  = $_.PercentUsed
            'Status'  = $_.Status
            'Type'    = $_.Type
        }
    })

    # chart anchored top-left; its source table sits below it
    $chart = New-KpiChart -Title '% of documented limit used' -ChartType BarClustered `
                -LabelColumn (ConvertTo-ExcelColumnLetter -Index $srcCol) `
                -ValueColumn (ConvertTo-ExcelColumnLetter -Index ($srcCol + 1)) `
                -FirstRow 30 -RowCount $chartTbl.Count `
                -Row 5 -Column 0 -Width 620 -Height 380 -SeriesHeader '% of limit'

    # source table for the chart (headers on row 29, data from 30)
    $chartTbl | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow 29 -StartColumn $srcCol -ExcelChartDefinition $chart

    # detail table beside the charts
    $detailRow = 6
    $detail | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $detailRow -StartColumn $detailCol -TableName 'tblLimits' -BoldTopRow

    # --- cards, titles, coloured status cells ---
    $highest    = ($rows | Measure-Object PercentUsed -Maximum).Maximum
    $approaching = @($rows | Where-Object { $_.PercentUsed -ge 60 }).Count
    $over        = @($rows | Where-Object { $_.Status -eq 'Over' }).Count

    $xl = Open-ExcelPackage -Path $ReportPath
    $ws = $xl.Workbook.Worksheets[$SheetName]
    Add-KpiCards -Worksheet $ws -Title 'SERVICE LIMITS & RECOMMENDATIONS' -Cards @(
        @{ Label = 'Limits tracked';      Value = $rows.Count }
        @{ Label = 'Highest usage';        Value = ("{0}%" -f $highest) }
        @{ Label = 'Approaching (>=60%)';  Value = $approaching }
        @{ Label = 'At or over limit';     Value = $over }
    )

    # label above the chart source table + above the detail table
    $srcL = ConvertTo-ExcelColumnLetter -Index $srcCol
    $ws.Cells["${srcL}28"].Value = 'Percent of limit used'
    $ws.Cells["${srcL}28"].Style.Font.Bold = $true
    $detL = ConvertTo-ExcelColumnLetter -Index $detailCol
    $ws.Cells["$detL$($detailRow-1)"].Value = "Limits vs current ($($detail.Count) metrics)"
    $ws.Cells["$detL$($detailRow-1)"].Style.Font.Bold = $true

    # colour the Status column in the detail table (Status is the 6th column)
    $statusCol = $detailCol + 5
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $cell = $ws.Cells[($detailRow + 1 + $i), $statusCol]
        $hex  = $rows[$i].Color.TrimStart('#')
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($rows[$i].Color))
    }

    # source note under the cards
    $ws.Cells['A2'].Value = 'Limits: Microsoft Entra service limits and restrictions (Microsoft Learn) via Config\EntraServiceLimits.json. Current values counted from this export.'
    $ws.Cells['A2'].Style.Font.Size = 9
    $ws.Cells['A2'].Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml('#808080'))
    Close-ExcelPackage $xl

    Set-KpiSheetLayout -Path $ReportPath -WorksheetName $SheetName

    # the shared layout pins columns 1-12 to width 11; the table's first column
    # (Area, in L) needs a little more room so "Conditional Access" fits.
    $xl2 = Open-ExcelPackage -Path $ReportPath
    $ws2 = $xl2.Workbook.Worksheets[$SheetName]
    $ws2.Column($detailCol).Width = 19
    Close-ExcelPackage $xl2
    return $true
}
