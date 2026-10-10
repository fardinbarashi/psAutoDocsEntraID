# Autodoc v1.0
function Add-KpiSheetConditionalAccess {
    <#
        Conditional Access overview: policy states, targeting, MFA usage and
        exclusions.

        Note on "policies with no hits": the image this sheet is modelled on
        shows unused policies, which requires CA sign-in insights from the
        reporting API. That data is not collected, so this sheet reports
        exclusion counts instead - the closest thing we can prove from what we
        actually export.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)]$Data,
        [string]$SheetName = '09 Conditional Access',
        [string]$SourceFolder
    )

    # The CA collector writes a flat table to CSV but a structured object to
    # JSON (ExportedAt + three arrays). Accept whichever shape we were handed.
    if ($Data -and @($Data)[0].PSObject.Properties.Name -contains 'ConditionalAccessPolicies') {
        $policies = @(@($Data)[0].ConditionalAccessPolicies)
    }
    else {
        $policies = @($Data | Where-Object { $_.RecordType -eq 'Policy' })
    }
    if (-not $policies) {
        Write-Host "  09 Conditional Access skipped - no policies in the export." -ForegroundColor DarkYellow
        return $false
    }

    $total    = $policies.Count
    $enabled  = @($policies | Where-Object { "$($_.State)" -eq 'enabled' }).Count
    $report   = @($policies | Where-Object { "$($_.State)" -match 'report' }).Count
    $disabled = @($policies | Where-Object { "$($_.State)" -eq 'disabled' }).Count

    # --- per state ---
    # Source tables keep real columns - name, count, share - with the combined
    # chart label in a separate (hidden) column, so the block reads as data.
    $tblState = @(
        [pscustomobject]@{ 'State' = 'Enabled';     'Policies' = $enabled;  'Share %' = (Get-KpiPercent $enabled $total);  'Chart label' = (Get-KpiLabel 'Enabled' $enabled $total) }
        [pscustomobject]@{ 'State' = 'Report-only'; 'Policies' = $report;   'Share %' = (Get-KpiPercent $report $total);   'Chart label' = (Get-KpiLabel 'Report-only' $report $total) }
        [pscustomobject]@{ 'State' = 'Disabled';    'Policies' = $disabled; 'Share %' = (Get-KpiPercent $disabled $total); 'Chart label' = (Get-KpiLabel 'Disabled' $disabled $total) }
    ) | Where-Object { $_.Policies -gt 0 }

    # --- what each policy targets (a policy can appear in several buckets) ---
    $condUsers  = @($policies | Where-Object { $_.IncludeUsersResolved  -or $_.IncludeGroupsResolved -or $_.IncludeRolesResolved }).Count
    $condApps   = @($policies | Where-Object { $_.IncludeApplicationsResolved }).Count
    $condDevice = @($policies | Where-Object { $_.DeviceFilterRule -or $_.IncludePlatforms }).Count
    $condPlace  = @($policies | Where-Object { $_.IncludeLocationsResolved }).Count
    $tblType = @(
        @{ N='Users'; V=$condUsers }, @{ N='Apps'; V=$condApps }, @{ N='Devices'; V=$condDevice }, @{ N='Locations'; V=$condPlace }
    ) | Where-Object { $_.V -gt 0 } | ForEach-Object {
        [pscustomobject]@{ 'Condition' = $_.N; 'Policies' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- most-targeted groups ---
    $targetCount = @{}
    foreach ($p in $policies) {
        foreach ($grp in @("$($p.IncludeGroupsResolved)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
            if (-not $targetCount.ContainsKey($grp)) { $targetCount[$grp] = 0 }
            $targetCount[$grp]++
        }
    }
    $tblTarget = @($targetCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
        ForEach-Object { [pscustomobject]@{ 'Target group' = $_.Key; 'Policies' = $_.Value; 'Share %' = (Get-KpiPercent $_.Value $total); 'Chart label' = (Get-KpiLabel (Get-ShortLabel $_.Key 26) $_.Value $total) } })

    # --- MFA required ---
    $mfaYes = @($policies | Where-Object { "$($_.BuiltInControls)" -match 'mfa' -or $_.AuthenticationStrengthResolved }).Count
    $tblMfa = @(
        [pscustomobject]@{ 'MFA' = 'MFA required'; 'Policies' = $mfaYes; 'Share %' = (Get-KpiPercent $mfaYes $total); 'Chart label' = (Get-KpiLabel 'MFA required' $mfaYes $total) }
        [pscustomobject]@{ 'MFA' = 'Not required'; 'Policies' = ($total - $mfaYes); 'Share %' = (Get-KpiPercent ($total-$mfaYes) $total); 'Chart label' = (Get-KpiLabel 'Not required' ($total-$mfaYes) $total) }
    )

    # --- exclusions ---
    $exUsers = @($policies | Where-Object { $_.ExcludeUsersResolved }).Count
    $exGrps  = @($policies | Where-Object { $_.ExcludeGroupsResolved }).Count
    $exRoles = @($policies | Where-Object { $_.ExcludeRolesResolved }).Count
    $exApps  = @($policies | Where-Object { $_.ExcludeApplicationsResolved }).Count
    $exLocs  = @($policies | Where-Object { $_.ExcludeLocationsResolved }).Count
    $withExclusions = @($policies | Where-Object {
        $_.ExcludeUsersResolved -or $_.ExcludeGroupsResolved -or $_.ExcludeRolesResolved -or
        $_.ExcludeApplicationsResolved -or $_.ExcludeLocationsResolved }).Count

    # The policy list sits beside the charts (right), and the small summary
    # tables behind the charts sit below the chart grid on the left.
    # Charts read the label column (4th) for categories, the count column (2nd) for values.
    $r = 6
    $srcCol = 1
    $detailCol = 22
    $strips = @(
        @{ Table = $tblState; Names = 'State'; Title = 'Policies per state';           Type = 'Pie' }
        @{ Table = $tblTarget; Names = 'Target group'; Title = 'Policies per target group';    Type = 'BarClustered' }
        @{ Table = $tblMfa; Names = 'MFA'; Title = 'MFA required in policy';       Type = 'Doughnut' }
        @{ Table = $tblType; Names = 'Condition'; Title = 'Conditions used';              Type = 'ColumnClustered' }
    )
    # Only strips that actually have data take part in the layout.
    $live = @($strips | Where-Object { $_.Table })
    # A bar chart with many categories is far taller than a doughnut, so the
    # grid advances by real heights instead of a fixed step - a fixed step is
    # what let the tallest chart grow down into the one beneath it.
    $heights = @($live | ForEach-Object {
        if ($_.Type -eq 'BarClustered') { 340 } else { 340 }
    })
    $grid = New-KpiChartGrid -Heights $heights

    $charts = @()
    $srcRow = $grid.NextRow        # summary tables start below the charts
    for ($idx = 0; $idx -lt $live.Count; $idx++) {
        $strip  = $live[$idx]
        $n      = @($strip.Table).Count
        $L      = @(0,1,2,3) | ForEach-Object { ConvertTo-ExcelColumnLetter -Index ($srcCol + $_) }
        $anchor = $grid.Anchors[$idx]
        $strip['Row'] = $srcRow
        $srcRow += $n + 4
        $charts += New-KpiChart -Title $strip.Title -ChartType $strip.Type `
                        -LabelColumn $L[3] -ValueColumn $L[1] -FirstRow ($strip.Row + 2) -RowCount $n `
                        -Height $heights[$idx] -Row $anchor.Row -Column $anchor.Column `
                        -SeriesHeader 'Policies'
    }

    # Charts ride along with the final export, so they must attach to the last
    # strip that is actually written - not $strips[-1], which may be empty.
    for ($i = 0; $i -lt $live.Count; $i++) {
        $strip = $live[$i]
        if ($i -eq $live.Count - 1) {
            $strip.Table | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow ($strip.Row + 1) -StartColumn $srcCol -ExcelChartDefinition $charts
        }
        else {
            $strip.Table | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow ($strip.Row + 1) -StartColumn $srcCol
        }
    }

    # --- full policy list, below every chart ---
    $detailRow = $r        # policy list sits beside the charts
    $detail = @($policies | Sort-Object DisplayName | ForEach-Object {
        [pscustomobject]@{
            'Policy name'  = $_.DisplayName
            'State'        = $_.State
            'MFA required' = $(if ("$($_.BuiltInControls)" -match 'mfa' -or $_.AuthenticationStrengthResolved) { 'Yes' } else { 'No' })
            'Grant'        = $_.BuiltInControls
            'Include'      = "$($_.IncludeUsersResolved) $($_.IncludeGroupsResolved)".Trim()
            'Exclude'      = "$($_.ExcludeUsersResolved) $($_.ExcludeGroupsResolved)".Trim()
        }
    })
    $detail | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $detailRow -StartColumn $detailCol -TableName 'tblCaPolicies' -BoldTopRow

    $xl = Open-ExcelPackage -Path $ReportPath
    $ws = $xl.Workbook.Worksheets[$SheetName]
    Add-KpiCards -Worksheet $ws -Title 'CONDITIONAL ACCESS - OVERVIEW' -Cards @(
        @{ Label = 'Total policies';  Value = $total }
        @{ Label = 'Enabled';         Value = $enabled }
        @{ Label = 'Report-only';     Value = $report }
        @{ Label = 'Disabled';        Value = $disabled }
        @{ Label = 'With exclusions'; Value = $withExclusions }
    )
    $ws.Cells["V$($detailRow-1)"].Value = "All Conditional Access policies ($($detail.Count))"
    $ws.Cells["V$($detailRow-1)"].Style.Font.Bold = $true
    $srcL = ConvertTo-ExcelColumnLetter -Index $srcCol
    foreach ($strip in $live) {
        $ws.Cells["$srcL$($strip.Row)"].Value = $strip.Title
        $ws.Cells["$srcL$($strip.Row)"].Style.Font.Bold = $true
    }
    Close-ExcelPackage $xl

    # Widths last: this makes the chart anchors land predictably and hides the
    # label columns (4th of each strip).
    Set-KpiChartColors -Path $ReportPath -WorksheetName $SheetName
    Set-KpiSheetLayout -Path $ReportPath -WorksheetName $SheetName
    # Policy name / MFA required / Include / Exclude hold long text; the layout
    # pass caps columns at 55, so widen these afterwards.
    if ($detail) {
        try {
            $xlW = Open-ExcelPackage -Path $ReportPath
            $wsW = $xlW.Workbook.Worksheets[$SheetName]
            foreach ($w in @(
                    @{ Col = $detailCol;     Prop = 'Policy name' }
                    @{ Col = $detailCol + 2; Prop = 'MFA required' }
                    @{ Col = $detailCol + 4; Prop = 'Include' }
                    @{ Col = $detailCol + 5; Prop = 'Exclude' }
                )) {
                $len = ($detail | ForEach-Object { "$($_.($w.Prop))".Length } | Measure-Object -Maximum).Maximum
                $wsW.Column($w.Col).Width = [math]::Min([math]::Max($len + 2, 18), 255)
            }
            Close-ExcelPackage $xlW
        } catch { }
    }
    Add-KpiDataSource -Path $ReportPath -WorksheetName $SheetName -SourceFolder $SourceFolder -BaseName 'ConditionalAccess' -Row 2
    return $true
}
