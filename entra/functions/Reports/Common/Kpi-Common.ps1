# Autodoc v1.0
function ConvertTo-ExcelColumnLetter {
    <# 1 -> A, 28 -> AB. Used to place KPI cards and chart source tables. #>
    param([Parameter(Mandatory)][int]$Index)
    $s = ''
    while ($Index -gt 0) { $m = ($Index - 1) % 26; $s = [char](65 + $m) + $s; $Index = [int](($Index - $m) / 26) }
    $s
}

function Get-KpiPercent {
    <# Rounded percentage, safe when the total is zero. #>
    param($Part, $Whole)
    if ($Whole) { [math]::Round(100 * $Part / $Whole, 1) } else { 0 }
}

function Get-KpiLabel {
    <#
        Builds a chart category label carrying its own count and share, e.g.
        "Enabled  28 (66.7 %)". Numbers are baked into the label instead of
        using chart data labels, which keeps the workbook XML valid.
    #>
    param([string]$Text, [int]$Count, [int]$Total)
    "$Text  $Count ($(Get-KpiPercent $Count $Total) %)"
}

function Get-ShortLabel {
    <# Truncates long names so chart legends stay readable. #>
    param([string]$Text, [int]$Max = 34)
    if ($Text.Length -gt $Max) { $Text.Substring(0, $Max - 1) + '…' } else { $Text }
}

function Add-KpiCards {
    <#
        Writes a sheet title plus a row of KPI cards (grey label above a large
        bold value), spaced two columns apart.
    #>
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$Cards,
        [int]$LabelRow = 3
    )

    $Worksheet.Cells['A1'].Value = $Title
    $Worksheet.Cells['A1'].Style.Font.Size = 16
    $Worksheet.Cells['A1'].Style.Font.Bold = $true

    $col = 1
    foreach ($card in $Cards) {
        $L = ConvertTo-ExcelColumnLetter -Index $col
        $Worksheet.Cells["${L}$LabelRow"].Value = $card.Label
        $Worksheet.Cells["${L}$LabelRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
        $Worksheet.Cells["${L}$($LabelRow + 1)"].Value = $card.Value
        $Worksheet.Cells["${L}$($LabelRow + 1)"].Style.Font.Size = 20
        $Worksheet.Cells["${L}$($LabelRow + 1)"].Style.Font.Bold = $true
        $Worksheet.Column($col).Width = 22
        $col += 2
    }
}

function ConvertTo-SafeDateTime {
    <#
        Parses a date string and returns $null when it can't be read.

        Avoid [datetime]::TryParse with [ref]$var here: when the variable was
        initialised to $null it is untyped, so .NET can't match the
        TryParse(string, ref DateTime) overload and PowerShell throws
        "Cannot find an overload for TryParse and the argument count: 2".
    #>
    [CmdletBinding()]
    param($Value)

    if (-not $Value) { return $null }
    try   { return [datetime]::Parse("$Value", [System.Globalization.CultureInfo]::InvariantCulture) }
    catch {
        try   { return [datetime]"$Value" }
        catch { return $null }
    }
}

function Get-AppPermissionInfo {
    <#
        Returns @{ Text = <readable permissions>; Count = <n> } for an app.

        Newer exports carry ApiPermissionNames / ApiPermissionCount, resolved at
        collection time from the resource service principals. Exports taken
        before that change only have the raw ApiPermissions string, so this
        falls back to it and counts the entries - the report stays usable
        against older data instead of showing blanks.
    #>
    [CmdletBinding()]
    param($App)

    $text = "$($App.ApiPermissionNames)"
    if ($text) {
        $count = 0
        if ($App.ApiPermissionCount -ne $null -and "$($App.ApiPermissionCount)" -ne '') {
            [void][int]::TryParse("$($App.ApiPermissionCount)", [ref]$count)
        }
        else { $count = @($text -split ',' | Where-Object { $_.Trim() }).Count }
        return @{ Text = $text; Count = $count }
    }

    $raw = "$($App.ApiPermissions)"
    if (-not $raw) { return @{ Text = ''; Count = 0 } }
    @{ Text = "$raw  (GUIDs - re-run the collection to get names)"; Count = @($raw -split ',' | Where-Object { $_.Trim() }).Count }
}

function New-KpiChart {
    <#
        Builds one chart definition. Category labels already carry their count
        and share, so the numbers read straight off the chart.

        Bar and column charts are left as a single series; per-bar colouring is
        applied afterwards by Set-KpiChartColors, which flips the chart's
        VaryColors flag. That is a plain boolean on the chart element and
        round-trips cleanly - unlike the data-label objects an earlier version
        wrote by hand, which produced a workbook Excel refused to open.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$ChartType,
        [Parameter(Mandatory)][string]$LabelColumn,
        [Parameter(Mandatory)][string]$ValueColumn,
        [Parameter(Mandatory)][int]$FirstRow,
        [Parameter(Mandatory)][int]$RowCount,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][int]$Column,
        [int]$Width = 520,
        [int]$Height = 340,
        [string]$SeriesHeader = 'Value'
    )

    $last = $FirstRow + $RowCount - 1
    $def = @{
        Title = $Title; ChartType = $ChartType
        XRange = "$LabelColumn${FirstRow}:$LabelColumn$last"
        YRange = "$ValueColumn${FirstRow}:$ValueColumn$last"
        Width = $Width; Height = $Height; Row = $Row; Column = $Column
        TitleBold = $true; TitleSize = 13
    }
    if ($ChartType -in 'Pie', 'Doughnut') { $def['ShowPercent'] = $true }
    else                                  { $def['SeriesHeader'] = $SeriesHeader }
    New-ExcelChartDefinition @def
}

function New-KpiChartGrid {
    <#
        Places charts two per row and returns each one's anchor.

        Heights vary - a 30-row bar chart is twice as tall as a doughnut - so
        each grid row advances by the tallest chart in it rather than by a fixed
        step. A fixed step is what let the RBAC role chart grow down into the
        chart below it.

        $Heights is the pixel height of every chart, in order. Returns an array
        of @{ Row; Column } plus the first free row after the grid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int[]]$Heights,
        [int]$StartRow = 6,
        [int]$SecondColumn = 11,
        [int]$RowHeightPx = 20,
        [int]$GapRows = 3
    )

    $anchors = @()
    $row = $StartRow
    for ($i = 0; $i -lt $Heights.Count; $i += 2) {
        $anchors += @{ Row = $row; Column = 1 }
        if ($i + 1 -lt $Heights.Count) { $anchors += @{ Row = $row; Column = $SecondColumn } }

        $tallest = $Heights[$i]
        if ($i + 1 -lt $Heights.Count -and $Heights[$i + 1] -gt $tallest) { $tallest = $Heights[$i + 1] }
        $row += [math]::Ceiling($tallest / $RowHeightPx) + $GapRows
    }
    @{ Anchors = $anchors; NextRow = $row }
}

function Set-KpiChartColors {
    <#
        Turns on VaryColors for bar and column charts so each bar gets its own
        colour. Pie and doughnut charts already vary by default.
        Run after every export to the sheet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorksheetName
    )

    try {
        $xl = Open-ExcelPackage -Path $Path
        $ws = $xl.Workbook.Worksheets[$WorksheetName]
        foreach ($drawing in $ws.Drawings) {
            try { if ($drawing.Series.Count -eq 1) { $drawing.VaryColors = $true } } catch { }
        }
        Close-ExcelPackage $xl
    }
    catch { Write-Host "  (colour pass skipped for $WorksheetName)" -ForegroundColor DarkGray }
}

function Add-KpiDataSource {
    <#
        Writes a consistent header onto a KPI sheet:
          A1  "<TITLE> - OVERVIEW"           (only when -Title is given)
          A<Row> "Data source (JSON): <path(s)>"  (single cell, overflows right)
        The JSON path(s) are resolved from <SourceFolder>\json for each -BaseName.
        Kept on rows 1-2 (above the charts/table) so nothing hides it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorksheetName,
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string[]]$BaseName,
        [string]$Title,
        [int]$Row = 2
    )
    try {
        if (-not $SourceFolder) { return }
        $jsonDir = Join-Path $SourceFolder 'rawDataJson'
        $paths = foreach ($b in $BaseName) {
            $f = Get-ChildItem -Path $jsonDir -Filter "$b*.json" -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($f) { $f.FullName } else { "$b (file not found)" }
        }
        $xl = Open-ExcelPackage -Path $Path
        $ws = $xl.Workbook.Worksheets[$WorksheetName]
        if ($ws) {
            if ($Title) {
                $ws.Cells['A1'].Value = "$Title - OVERVIEW"
                $ws.Cells['A1'].Style.Font.Bold = $true
                $ws.Cells['A1'].Style.Font.Size = 14
            }
            $shortPaths = $paths | ForEach-Object { $_ -replace '^.*?(\\Entra\\)', '$1' }
            $ws.Cells[$Row, 1].Value = 'Data source (JSON): ' + ($shortPaths -join '   |   ')
            $ws.Cells[$Row, 1].Style.Font.Bold = $false
            $ws.Cells[$Row, 1].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
        }
        Close-ExcelPackage $xl
    }
    catch { Write-Host "  (data-source note skipped for $WorksheetName : $($_.Exception.Message))" -ForegroundColor DarkGray }
}

function Set-KpiSheetLayout {
    <#
        Final layout pass for a KPI sheet. Run this AFTER every export to the
        sheet, never before.

        Why it exists: charts are anchored by column INDEX, but Export-Excel
        -AutoSize sizes columns to their content. A detail table holding a
        200-character licence list widens column A to thousands of pixels, and a
        chart anchored at "column 6" then lands far off screen. Fixing the width
        of the first few columns makes chart placement deterministic, and
        capping the rest stops long text from stretching the sheet sideways.

        It also replaces -AutoSize on these sheets: autofitting a range that
        includes a hidden column raises "You cannot call a method on a
        null-valued expression", which is where that warning came from.

        Note on hiding: this function no longer hides anything. Excel does not
        plot cells in hidden columns, so hiding a chart's category column
        silently strips every label off the chart - which is exactly what
        happened to the office chart. Extra columns stay visible instead.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorksheetName,
        [int]$ChartZoneColumns = 12,
        [double]$ChartZoneWidth = 11,
        [double]$MinWidth = 10,
        [double]$MaxWidth = 55
    )

    try {
        $xl = Open-ExcelPackage -Path $Path
        $ws = $xl.Workbook.Worksheets[$WorksheetName]

        if ($ws -and $ws.Dimension) {
            # Fit to content but never wider than MaxWidth
            try { $ws.Cells[$ws.Dimension.Address].AutoFitColumns($MinWidth, $MaxWidth) } catch { }

            # The chart zone keeps a fixed width so anchors are predictable
            for ($i = 1; $i -le $ChartZoneColumns; $i++) { $ws.Column($i).Width = $ChartZoneWidth }
        }

        Close-ExcelPackage $xl
    }
    catch { Write-Host "  (layout pass skipped for $WorksheetName : $($_.Exception.Message))" -ForegroundColor DarkGray }
}

function Measure-MapTextWidth {
    <#
        Rendered width of a string in inches at 9pt, summed from a per-character
        table MEASURED by rendering each glyph. Keyed on character code so upper
        and lower case are distinct (PowerShell hash keys are case-insensitive).
        Unknown characters use the average. Scales linearly with font size.
        Replaces the old single-average guess that mis-sized cards.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [double]$FontSize = 9
    )
    if (-not $script:MapGlyphWidth) {
        $script:MapGlyphWidth = @{
            32 = 0.0312
            65 = 0.0625
            66 = 0.0625
            67 = 0.0625
            68 = 0.0729
            69 = 0.0625
            70 = 0.0521
            71 = 0.0729
            72 = 0.0729
            73 = 0.0312
            74 = 0.0312
            75 = 0.0625
            76 = 0.0521
            77 = 0.0833
            78 = 0.0729
            79 = 0.0729
            80 = 0.0521
            81 = 0.0729
            82 = 0.0625
            83 = 0.0625
            84 = 0.0625
            85 = 0.0729
            86 = 0.0625
            87 = 0.0938
            88 = 0.0625
            89 = 0.0625
            90 = 0.0625
            97 = 0.0625
            98 = 0.0625
            99 = 0.0521
            100 = 0.0625
            101 = 0.0625
            102 = 0.0312
            103 = 0.0625
            104 = 0.0625
            105 = 0.0312
            106 = 0.0312
            107 = 0.0521
            108 = 0.0312
            109 = 0.0938
            110 = 0.0625
            111 = 0.0625
            112 = 0.0625
            113 = 0.0625
            114 = 0.0417
            115 = 0.0521
            116 = 0.0417
            117 = 0.0625
            118 = 0.0521
            119 = 0.0729
            120 = 0.0521
            121 = 0.0521
            122 = 0.0521
            48 = 0.0625
            49 = 0.0625
            50 = 0.0625
            51 = 0.0625
            52 = 0.0625
            53 = 0.0625
            54 = 0.0625
            55 = 0.0625
            56 = 0.0625
            57 = 0.0625
            40 = 0.0417
            41 = 0.0417
            91 = 0.0417
            93 = 0.0417
            45 = 0.0312
            95 = 0.0521
            46 = 0.0312
            47 = 0.0312
            58 = 0.0312
            44 = 0.0312
            59 = 0.0312
            43 = 0.0833
            229 = 0.0625
            228 = 0.0625
            246 = 0.0625
            197 = 0.0625
            196 = 0.0625
            214 = 0.0729
            233 = 0.0625
            232 = 0.0625
            252 = 0.0625
        }
        $script:MapGlyphAvg = ($script:MapGlyphWidth.Values | Measure-Object -Average).Average
    }
    $w = 0.0
    foreach ($ch in $Text.ToCharArray()) {
        $cw = $script:MapGlyphWidth[[int][char]$ch]
        if ($null -eq $cw) { $cw = $script:MapGlyphAvg }
        $w += $cw
    }
    $w * ($FontSize / 9.0)
}
