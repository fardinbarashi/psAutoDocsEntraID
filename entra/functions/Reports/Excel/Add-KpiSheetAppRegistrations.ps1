# Autodoc v1.0
function Add-KpiSheetAppRegistrations {
    <#
        App registration overview: ownership, tenancy, credential expiry,
        API permissions and sign-in recency.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)]$Data,
        [string]$SheetName = '10 App registrations',
        [string]$SourceFolder
    )

    $apps = @($Data)
    if (-not $apps) { return $false }
    $total = $apps.Count

    # --- tenancy ---
    $single = @($apps | Where-Object { "$($_.SignInAudience)" -eq 'AzureADMyOrg' }).Count
    $multi  = $total - $single
    # Real columns - name, count, share - with the chart label kept separate
    $tblTenancy = @(
        @{ N='Single tenant'; V=$single }, @{ N='Multi tenant'; V=$multi }
    ) | ForEach-Object { [pscustomobject]@{ 'Tenancy' = $_.N; 'Apps' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- owners ---
    $noOwner = @($apps | Where-Object { "$($_.OwnerInfo)" -eq 'MissingOwners' -or -not $_.OwnerInfo }).Count
    $ownerCount = @{}
    foreach ($a in $apps) {
        if ("$($a.OwnerInfo)" -eq 'MissingOwners') { continue }
        foreach ($o in @("$($a.OwnerInfo)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
            if (-not $ownerCount.ContainsKey($o)) { $ownerCount[$o] = 0 }
            $ownerCount[$o]++
        }
    }
    $tblOwner = @($ownerCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
        ForEach-Object { [pscustomobject]@{ 'Owner' = $_.Key; 'Apps' = $_.Value; 'Share %' = (Get-KpiPercent $_.Value $total); 'Chart label' = (Get-KpiLabel (Get-ShortLabel $_.Key 26) $_.Value $total) } })

    # --- credential expiry bands ---
    $now = Get-Date
    $bandOf = {
        param($app)
        if ("$($app.CredentialExpired)" -eq 'Yes') { return 'Expired' }
        $d = ConvertTo-SafeDateTime -Value $app.NearestCredentialExpiry
        if (-not $d) { return 'No credentials' }
        $days = ($d - $now).TotalDays
        if     ($days -lt 7)  { '< 7 days' }
        elseif ($days -lt 30) { '7-30 days' }
        elseif ($days -lt 90) { '30-90 days' }
        else                  { '90+ days' }
    }
    $expiryOrder = 'Expired','< 7 days','7-30 days','30-90 days','90+ days','No credentials'
    $appBand = @{}
    foreach ($a in $apps) { $appBand[$a.AppId] = & $bandOf $a }
    $tblExpiry = @(foreach ($b in $expiryOrder) {
        $cnt = @($appBand.Values | Where-Object { $_ -eq $b }).Count
        if ($cnt -gt 0) { [pscustomobject]@{ 'Expiry' = $b; 'Apps' = $cnt; 'Share %' = (Get-KpiPercent $cnt $total); 'Chart label' = (Get-KpiLabel $b $cnt $total) } }
    })
    $expiringSoon = @($apps | Where-Object { "$($_.CredentialExpiresWithin30Days)" -eq 'Yes' }).Count

    # --- API permissions requested ---
    $permCountOf = @{}
    $permTextOf  = @{}
    $permBandOf  = @{}
    foreach ($a in $apps) {
        $info = Get-AppPermissionInfo -App $a
        $n = $info.Count
        $permCountOf[$a.AppId] = $n
        $permTextOf[$a.AppId]  = $info.Text
        $permBandOf[$a.AppId]  = if     ($n -eq 0)  { '0' }
                                 elseif ($n -le 5)  { '1-5' }
                                 elseif ($n -le 20) { '6-20' }
                                 elseif ($n -le 50) { '21-50' }
                                 else               { '50+' }
    }
    $permBand = @($permBandOf.Values)
    $tblPerm = @(foreach ($b in '0','1-5','6-20','21-50','50+') {
        $cnt = @($permBand | Where-Object { $_ -eq $b }).Count
        if ($cnt -gt 0) { [pscustomobject]@{ 'Permissions' = $b; 'Apps' = $cnt; 'Share %' = (Get-KpiPercent $cnt $total); 'Chart label' = (Get-KpiLabel $b $cnt $total) } }
    })

    # --- sign-in recency (7 / 8-30 / 31-90 / 90+) ---
    $recencyOrder = 'Last 7 days','8-30 days','31-90 days','90+ days','Never / no data'
    $recency = @(foreach ($a in $apps) { Get-SignInRecencyBand -LastSignIn $a.LastSignInDateTime })
    $tblRecency = @(foreach ($b in $recencyOrder) {
        $cnt = @($recency | Where-Object { $_ -eq $b }).Count
        if ($cnt -gt 0) { [pscustomobject]@{ 'Last sign-in' = $b; 'Apps' = $cnt; 'Share %' = (Get-KpiPercent $cnt $total); 'Chart label' = (Get-KpiLabel $b $cnt $total) } }
    })

    # Charts (grid) on the left; the two app tables on the right (column V);
    # the small summary tables the charts read from go below the charts (column A).
    $r = 6
    $srcCol    = 1     # chart source (summary) tables -> column A, below the charts
    $detailCol = 22    # the two app tables -> column V, beside the charts
    $strips = @(
        @{ Table = $tblOwner;   Names = 'Owner';        Title = 'App registrations per owner (top 10)'; Type = 'BarClustered' }
        @{ Table = $tblTenancy; Names = 'Tenancy';      Title = 'Authentication type';                  Type = 'Doughnut' }
        @{ Table = $tblExpiry;  Names = 'Expiry';       Title = 'Credential expiry';                     Type = 'ColumnClustered' }
        @{ Table = $tblPerm;    Names = 'Permissions';  Title = 'API permissions requested';             Type = 'ColumnClustered' }
        @{ Table = $tblRecency; Names = 'Last sign-in'; Title = 'Last sign-in';                          Type = 'ColumnClustered' }
    )
    $live = @($strips | Where-Object { $_.Table })
    $heights = @($live | ForEach-Object { 340 })
    $grid = New-KpiChartGrid -Heights $heights

    $charts = @()
    $srcRow = 67       # chart-source/summary tables start at row 67 (below the charts)
    for ($idx = 0; $idx -lt $live.Count; $idx++) {
        $strip  = $live[$idx]
        $n      = @($strip.Table).Count
        $L      = @(0, 1, 2, 3) | ForEach-Object { ConvertTo-ExcelColumnLetter -Index ($srcCol + $_) }
        $anchor = $grid.Anchors[$idx]
        $strip['Row'] = $srcRow
        $srcRow += $n + 4
        $charts += New-KpiChart -Title $strip.Title -ChartType $strip.Type `
            -LabelColumn $L[3] -ValueColumn $L[1] -FirstRow ($strip.Row + 2) -RowCount $n `
            -Height $heights[$idx] -Row $anchor.Row -Column $anchor.Column -SeriesHeader 'Apps'
    }
    for ($i = 0; $i -lt $live.Count; $i++) {
        $strip = $live[$i]
        if ($i -eq $live.Count - 1) {
            $strip.Table | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow ($strip.Row + 1) -StartColumn $srcCol -ExcelChartDefinition $charts
        }
        else {
            $strip.Table | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow ($strip.Row + 1) -StartColumn $srcCol
        }
    }

    # --- apps with risk findings (beside the charts, top-right) ---
    $riskRow = 7
    $risky = @(foreach ($a in $apps) {
        $f = Get-AppRiskFinding -App $a
        if ($f.Risk) {
            [pscustomobject]@{
                'App name'     = $a.AppName
                'Reason'       = $f.Reason
                'Owner'        = $(if ("$($a.OwnerInfo)" -eq 'MissingOwners') { '(none)' } else { $a.OwnerInfo })
                'Last sign-in' = $a.LastSignInDateTime
            }
        }
    })
    $risky = @($risky | Sort-Object 'App name')
    if ($risky) {
        $risky | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $riskRow -StartColumn $detailCol -TableName 'tblAppRisk' -BoldTopRow
    }

    # --- every app registration (beside the charts, below the findings) ---
    $detailRow = [math]::Max(65, $riskRow + $risky.Count + 5)
    $detail = @($apps | Sort-Object AppName | ForEach-Object {
        [pscustomobject]@{
            'App name'          = $_.AppName
            'Credential expiry' = $appBand[$_.AppId]
            'Nearest expiry'    = $_.NearestCredentialExpiry
            'Last sign-in'      = $_.LastSignInDateTime
            'Owner'             = $(if ("$($_.OwnerInfo)" -eq 'MissingOwners') { '(none)' } else { $_.OwnerInfo })
            'API permissions'   = $permTextOf[$_.AppId]
            'Permission count'  = $permCountOf[$_.AppId]
        }
    })
    if ($detail) {
        $detail | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $detailRow -StartColumn $detailCol -TableName 'tblAppDetail' -BoldTopRow
    }

    $xl = Open-ExcelPackage -Path $ReportPath
    $ws = $xl.Workbook.Worksheets[$SheetName]
    Add-KpiCards -Worksheet $ws -Title 'APP REGISTRATIONS - OVERVIEW' -Cards @(
        @{ Label = 'Total app registrations'; Value = $total }
        @{ Label = 'Single tenant';           Value = $single }
        @{ Label = 'Multi tenant';            Value = $multi }
        @{ Label = 'Without owner';           Value = $noOwner }
        @{ Label = 'Expires < 30 days';       Value = $expiringSoon }
        @{ Label = 'Apps with risk';      Value = $risky.Count }
    )
    $detailL = ConvertTo-ExcelColumnLetter -Index $detailCol
    $ws.Cells["$detailL$($riskRow-1)"].Value = "Apps with risk ($($risky.Count))"
    $ws.Cells["$detailL$($riskRow-1)"].Style.Font.Bold = $true
    $ws.Cells["$detailL$($detailRow-1)"].Value = "All app registrations ($(@($detail).Count))"
    $ws.Cells["$detailL$($detailRow-1)"].Style.Font.Bold = $true
    $srcL = ConvertTo-ExcelColumnLetter -Index $srcCol
    foreach ($strip in $live) {
        $ws.Cells["$srcL$($strip.Row)"].Value = $strip.Title
        $ws.Cells["$srcL$($strip.Row)"].Style.Font.Bold = $true
    }
    Close-ExcelPackage $xl
    Set-KpiChartColors -Path $ReportPath -WorksheetName $SheetName
    Set-KpiSheetLayout -Path $ReportPath -WorksheetName $SheetName
    # widen the two app tables (column V onward) so nothing has to be dragged out
    try {
        $xlW = Open-ExcelPackage -Path $ReportPath
        $wsW = $xlW.Workbook.Worksheets[$SheetName]
        $firstR = $riskRow - 1
        $lastR  = $detailRow + @($detail).Count + 1
        for ($c = $detailCol; $c -le ($detailCol + 6); $c++) {
            $maxLen = 0
            for ($rr = $firstR; $rr -le $lastR; $rr++) {
                if ($rr -eq ($riskRow - 1) -or $rr -eq ($detailRow - 1)) { continue }  # skip the long title rows
                $len = "$($wsW.Cells[$rr, $c].Value)".Length
                if ($len -gt $maxLen) { $maxLen = $len }
            }
            if ($maxLen -gt 0) { $wsW.Column($c).Width = [math]::Min([math]::Max($maxLen + 2, 10), 90) }
        }
        Close-ExcelPackage $xlW
    } catch { }
    Add-KpiDataSource -Path $ReportPath -WorksheetName $SheetName -SourceFolder $SourceFolder -BaseName 'AppRegEnterpriseApps' -Row 2
    return $true
}
