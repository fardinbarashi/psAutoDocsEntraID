# Autodoc v1.0
function Add-KpiSheetEnterpriseApps {
    <#
        Enterprise application overview: SSO configuration, protocol mix,
        Graph permission usage and sign-in recency.

        SCOPE LIMITATION - read this before trusting the totals.
        The collector builds its app list from Get-MgApplication (app
        registrations) and joins each one to its service principal. Gallery
        and pre-integrated enterprise apps that have a service principal but
        no app registration in this tenant (Salesforce, Workday, ServiceNow
        and similar) are therefore NOT included, and neither are sign-in
        counts - only the date of the most recent sign-in.

        Covering those properly means collecting all service principals as
        their own dataset; until then this sheet describes the applications
        that are registered here.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)]$Data,
        $Groups,
        [string]$SheetName = '11 Enterprise apps',
        [string]$SourceFolder
    )

    $apps = @($Data | Where-Object { "$($_.ServicePrincipalExists)" -eq 'Yes' })
    if (-not $apps) { return $false }
    $total = $apps.Count

    # --- SSO protocol mix ---
    $ssoGroups = $apps | Group-Object SSOType | Sort-Object Count -Descending
    $tblSso = @($ssoGroups | ForEach-Object {
        $name = if ($_.Name) { $_.Name } else { 'None' }
        [pscustomobject]@{ 'SSO type' = $name; 'Apps' = $_.Count; 'Share %' = (Get-KpiPercent $_.Count $total); 'Chart label' = (Get-KpiLabel $name $_.Count $total) }
    })
    $ssoConfigured = @($apps | Where-Object { "$($_.SSOConfigured)" -eq 'Yes' }).Count

    # --- Graph permissions ---
    $graphYes = @($apps | Where-Object { "$($_.UsesGraphPermissions)" -eq 'Yes' }).Count
    $tblGraph = @(
        @{ N='Uses Graph API'; V=$graphYes }, @{ N='No Graph API'; V=($total - $graphYes) }
    ) | ForEach-Object { [pscustomobject]@{ 'Graph' = $_.N; 'Apps' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- sign-in recency (7 / 8-30 / 31-90 / 90+) ---
    $recencyOrder = 'Last 7 days','8-30 days','31-90 days','90+ days','Never / no data'
    $recency = @(foreach ($a in $apps) { Get-SignInRecencyBand -LastSignIn $a.LastSignInDateTime })
    $tblRecency = @(foreach ($b in $recencyOrder) {
        $cnt = @($recency | Where-Object { $_ -eq $b }).Count
        if ($cnt -gt 0) { [pscustomobject]@{ 'Last sign-in' = $b; 'Apps' = $cnt; 'Share %' = (Get-KpiPercent $cnt $total); 'Chart label' = (Get-KpiLabel $b $cnt $total) } }
    })
    $usedRecently = @($apps | Where-Object { "$($_.UsedWithin30Days)" -eq 'Yes' }).Count

    # --- how apps are reached: group assignment vs none recorded ---
    # The group collector records each group's app-role assignments, so we can
    # tell which apps are granted through a group. Per-USER direct assignments
    # (servicePrincipal appRoleAssignedTo) are not collected, so the second
    # bucket means "no group assignment found", not "assigned directly".
    $groupAssignedNames = @{}
    foreach ($g in @($Groups)) {
        foreach ($appName in @("$($g.Applications)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
            $groupAssignedNames[$appName] = $true
        }
    }
    $viaGroup = @($apps | Where-Object { $groupAssignedNames.ContainsKey("$($_.AppName)") }).Count
    $tblAssign = @(
        @{ N='Assigned via group'; V=$viaGroup }, @{ N='No group assignment'; V=($total - $viaGroup) }
    ) | ForEach-Object { [pscustomobject]@{ 'Assignment' = $_.N; 'Apps' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- public client (a risk signal: legacy / device-code flows) ---
    $publicYes = @($apps | Where-Object { "$($_.PublicClient)" -eq 'Yes' }).Count
    $tblClient = @(
        @{ N='Public client'; V=$publicYes }, @{ N='Confidential client'; V=($total - $publicYes) }
    ) | ForEach-Object { [pscustomobject]@{ 'Client' = $_.N; 'Apps' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # Charts (grid) on the left; the two app tables on the right (column V);
    # the small summary tables the charts read from go below the charts (column A).
    $r = 6
    $srcCol    = 1     # chart source (summary) tables -> column A, below the charts
    $detailCol = 22    # the two app tables -> column V, beside the charts
    $strips = @(
        @{ Table = $tblSso;     Names = 'SSO type';     Title = 'Authentication protocol (SSO type)'; Type = 'Doughnut' }
        @{ Table = $tblRecency; Names = 'Last sign-in'; Title = 'Last sign-in';                        Type = 'ColumnClustered' }
        @{ Table = $tblGraph;   Names = 'Graph';        Title = 'Microsoft Graph permissions';         Type = 'Doughnut' }
        @{ Table = $tblClient;  Names = 'Client';       Title = 'Client type';                         Type = 'Doughnut' }
        @{ Table = $tblAssign;  Names = 'Assignment';   Title = 'App assignment through groups';        Type = 'Doughnut' }
    )
    $live = @($strips | Where-Object { $_.Table })
    $heights = @($live | ForEach-Object { 340 })
    $grid = New-KpiChartGrid -Heights $heights

    $charts = @()
    $srcRow = 77       # chart-source/summary tables start at row 77 (below the charts)
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
                'Last sign-in' = $a.LastSignInDateTime
                'SSO type'     = $a.SSOType
                'Reason'       = $f.Reason
            }
        }
    })
    $risky = @($risky | Sort-Object 'App name')
    if ($risky) {
        $risky | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $riskRow -StartColumn $detailCol -TableName 'tblEntAppRisk' -BoldTopRow
    }

    # --- every enterprise app (beside the charts, below the risk list) ---
    $detailRow = [math]::Max(61, $riskRow + @($risky).Count + 3)
    $detail = @($apps | Sort-Object AppName | ForEach-Object {
        $permInfo  = Get-AppPermissionInfo -App $_
        [pscustomobject]@{
            'App name'         = $_.AppName
            'Last sign-in'     = $_.LastSignInDateTime
            'Permission count' = $permInfo.Count
            'SSO type'         = $(if ($_.SSOType) { $_.SSOType } else { 'None' })
            'SSO configured'   = $_.SSOConfigured
            'API permissions'  = $permInfo.Text
        }
    })
    if ($detail) {
        $detail | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $detailRow -StartColumn $detailCol -TableName 'tblEnterpriseApps' -BoldTopRow
    }

    $xl = Open-ExcelPackage -Path $ReportPath
    $ws = $xl.Workbook.Worksheets[$SheetName]
    Add-KpiCards -Worksheet $ws -Title 'ENTERPRISE APPS - OVERVIEW' -Cards @(
        @{ Label = 'Apps with service principal'; Value = $total }
        @{ Label = 'SSO configured';              Value = $ssoConfigured }
        @{ Label = 'Without SSO';                 Value = ($total - $ssoConfigured) }
        @{ Label = 'Used last 30 days';           Value = $usedRecently }
        @{ Label = 'Uses Graph API';              Value = $graphYes }
        @{ Label = 'Apps with risk';          Value = $risky.Count }
    )
    $ws.Cells['A5'].Value = 'Scope: applications registered in this tenant. Gallery apps without a local app registration are not included.'
    $ws.Cells['A5'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
    $detailL = ConvertTo-ExcelColumnLetter -Index $detailCol
    $ws.Cells["$detailL$($detailRow-1)"].Value = "All enterprise apps ($($detail.Count))"
    $ws.Cells["$detailL$($detailRow-1)"].Style.Font.Bold = $true
    $ws.Cells["$detailL$($riskRow-1)"].Value = "Apps with risk ($($risky.Count))"
    $ws.Cells["$detailL$($riskRow-1)"].Style.Font.Bold = $true
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
        for ($c = $detailCol; $c -le ($detailCol + 5); $c++) {
            $maxLen = 0
            for ($rr = $firstR; $rr -le $lastR; $rr++) {
                if ($rr -eq ($detailRow - 1) -or $rr -eq ($riskRow - 1)) { continue }  # skip the long title rows
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
