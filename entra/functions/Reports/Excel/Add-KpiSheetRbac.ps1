# Autodoc v1.0
function Add-KpiSheetRbac {
    <#
        RBAC overview: role assignments, privileged roles, principals holding
        several roles, and permanent versus time-bound assignments.

        Two deviations from the layout this sheet is modelled on:
          * "Assignments per scope" (subscription / resource group) belongs to
            Azure Resource Manager RBAC. We collect Entra *directory* roles,
            which have no such scope, so the sheet shows the principal type
            split (user / group / service principal) instead.
          * "Inactive assignments" needs sign-in telemetry we don't collect;
            the active-versus-eligible split is shown in its place.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportPath,
        [Parameter(Mandatory)]$Data,
        [string]$SheetName = '12 RBAC',
        [string]$SourceFolder
    )

    $assignments = @($Data | Where-Object { $_.RecordType -like 'RoleAssignment*' })
    if (-not $assignments) { return $false }
    $total = $assignments.Count

    # Well-known high-impact Entra directory roles
    $privilegedRoles = @(
        'Global Administrator','Privileged Role Administrator','Privileged Authentication Administrator',
        'Security Administrator','Conditional Access Administrator','Application Administrator',
        'Cloud Application Administrator','User Administrator','Exchange Administrator',
        'SharePoint Administrator','Intune Administrator','Hybrid Identity Administrator',
        'Domain Name Administrator','Directory Synchronization Accounts','Partner Tier2 Support'
    )

    $distinctRoles = @($assignments | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count
    $privAssign    = @($assignments | Where-Object { $privilegedRoles -contains "$($_.RoleDefinitionName)" })
    $timeBound     = @($assignments | Where-Object { $_.EndDateTime }).Count
    $permanent     = $total - $timeBound

    # --- assignments per role ---
    $tblRole = @($assignments | Group-Object RoleDefinitionName | Sort-Object Count -Descending | Select-Object -First 30 |
        ForEach-Object { [pscustomobject]@{ 'Role' = $_.Name; 'Assignments' = $_.Count; 'Share %' = (Get-KpiPercent $_.Count $total); 'Chart label' = (Get-KpiLabel (Get-ShortLabel $_.Name 28) $_.Count $total) } })

    # --- principal type ---
    $tblPrincipal = @($assignments | Group-Object PrincipalType | Sort-Object Count -Descending |
        ForEach-Object {
            $name = if ($_.Name) { $_.Name } else { 'Unknown' }
            [pscustomobject]@{ 'Principal type' = $name; 'Assignments' = $_.Count; 'Share %' = (Get-KpiPercent $_.Count $total); 'Chart label' = (Get-KpiLabel $name $_.Count $total) }
        })

    # --- permanent vs time-bound ---
    $tblDuration = @(
        @{ N='Permanent'; V=$permanent }, @{ N='Time-bound'; V=$timeBound }
    ) | ForEach-Object { [pscustomobject]@{ 'Duration' = $_.N; 'Assignments' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- active vs eligible ---
    $active   = @($assignments | Where-Object { $_.RecordType -eq 'RoleAssignmentActive' }).Count
    $eligible = @($assignments | Where-Object { $_.RecordType -eq 'RoleAssignmentEligible' }).Count
    $tblKind = @(
        @{ N='Active'; V=$active }, @{ N='Eligible'; V=$eligible }
    ) | ForEach-Object { [pscustomobject]@{ 'Assignment' = $_.N; 'Assignments' = $_.V; 'Share %' = (Get-KpiPercent $_.V $total); 'Chart label' = (Get-KpiLabel $_.N $_.V $total) } }

    # --- principals holding several roles ---
    $multiRole = @($assignments | Group-Object PrincipalId |
        Where-Object { @($_.Group | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count -gt 1 })

    # Charts (grid) on the left; the detail table on the right (column V);
    # the small summary tables the charts read from go below the charts (column A).
    $r = 6
    $srcCol    = 1     # chart source (summary) tables -> column A, below the charts
    $detailCol = 22    # detail table -> column V, beside the charts
    $strips = @(
        @{ Table = $tblRole; Names = 'Role'; Title = 'Assignments per role';           Type = 'BarClustered' }
        @{ Table = $tblPrincipal; Names = 'Principal type'; Title = 'Assignments per principal type'; Type = 'Doughnut' }
        @{ Table = $tblDuration; Names = 'Duration'; Title = 'Permanent vs time-bound';        Type = 'Doughnut' }
        @{ Table = $tblKind; Names = 'Assignment'; Title = 'Active vs eligible';             Type = 'Doughnut' }
    )
    # Only strips that actually have data take part in the layout.
    $live = @($strips | Where-Object { $_.Table })
    # A bar chart with many categories is far taller than a doughnut, so the
    # grid advances by real heights instead of a fixed step - a fixed step is
    # what let the tallest chart grow down into the one beneath it.
    $heights = @($live | ForEach-Object {
        if ($_.Type -eq 'BarClustered') { [math]::Max(340, 20 * @($_.Table).Count + 140) } else { 340 }
    })
    $grid = New-KpiChartGrid -Heights $heights

    $charts = @()
    $srcRow = 77       # chart-source/summary tables start at row 77 (below the charts)
    for ($idx = 0; $idx -lt $live.Count; $idx++) {
        $strip  = $live[$idx]
        $n      = @($strip.Table).Count
        $L      = @(0,1,2,3) | ForEach-Object { ConvertTo-ExcelColumnLetter -Index ($srcCol + $_) }
        $anchor = $grid.Anchors[$idx]
        $aRow = $anchor.Row; $aCol = $anchor.Column; $aH = $heights[$idx]
        # stack these two doughnuts under the principal-type chart in column L
        # (row is 0-based, so 26 -> cell L27 and 42 -> cell L43), same size as the rest
        if     ($strip.Title -eq 'Permanent vs time-bound') { $aRow = 24; $aCol = 11; $aH = 340 }
        elseif ($strip.Title -eq 'Active vs eligible')       { $aRow = 42; $aCol = 11; $aH = 340 }
        $strip['Row'] = $srcRow
        $srcRow += $n + 4
        $charts += New-KpiChart -Title $strip.Title -ChartType $strip.Type `
                        -LabelColumn $L[3] -ValueColumn $L[1] -FirstRow ($strip.Row + 2) -RowCount $n `
                        -Height $aH -Row $aRow -Column $aCol `
                        -SeriesHeader 'Assignments'
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
    # --- privileged role assignments in full (beside the charts, top-right) ---
    $detailRow = 7
    $detail = @($privAssign | Sort-Object RoleDefinitionName, PrincipalDisplayName | ForEach-Object {
        [pscustomobject]@{
            'Role'           = $_.RoleDefinitionName
            'Principal'      = $_.PrincipalDisplayName
            'Type'           = $_.PrincipalType
            'Assignment'     = $_.AssignmentType
            'Activated (JIT)'= $_.JITActivated
            'Ends'           = $(if ($_.EndDateTime) { $_.EndDateTime } else { 'Permanent' })
        }
    })
    if ($detail) {
        $detail | Export-Excel -Path $ReportPath -WorksheetName $SheetName -StartRow $detailRow -StartColumn $detailCol -TableName 'tblPrivilegedRoles' -BoldTopRow
    }

    $xl = Open-ExcelPackage -Path $ReportPath
    $ws = $xl.Workbook.Worksheets[$SheetName]
    Add-KpiCards -Worksheet $ws -Title 'RBAC - OVERVIEW' -Cards @(
        @{ Label = 'Distinct roles';         Value = $distinctRoles }
        @{ Label = 'Role assignments';       Value = $total }
        @{ Label = 'Privileged assignments'; Value = $privAssign.Count }
        @{ Label = 'Principals > 1 role';    Value = $multiRole.Count }
        @{ Label = 'Time-bound';             Value = $timeBound }
    )
    $detailL = ConvertTo-ExcelColumnLetter -Index $detailCol
    $ws.Cells["$detailL$($detailRow-1)"].Value = "Privileged role assignments ($(@($detail).Count))"
    $ws.Cells["$detailL$($detailRow-1)"].Style.Font.Bold = $true
    $srcL = ConvertTo-ExcelColumnLetter -Index $srcCol
    foreach ($strip in $live) {
        $ws.Cells["$srcL$($strip.Row)"].Value = $strip.Title
        $ws.Cells["$srcL$($strip.Row)"].Style.Font.Bold = $true
    }
    Close-ExcelPackage $xl
    Set-KpiChartColors -Path $ReportPath -WorksheetName $SheetName
    Set-KpiSheetLayout -Path $ReportPath -WorksheetName $SheetName
    # widen the detail table (column V onward) so nothing has to be dragged out
    if ($detail) {
        try {
            $xlW = Open-ExcelPackage -Path $ReportPath
            $wsW = $xlW.Workbook.Worksheets[$SheetName]
            $firstR = $detailRow
            $lastR  = $detailRow + @($detail).Count + 1
            for ($c = $detailCol; $c -le ($detailCol + 5); $c++) {
                $maxLen = 0
                for ($rr = $firstR; $rr -le $lastR; $rr++) {
                    $len = "$($wsW.Cells[$rr, $c].Value)".Length
                    if ($len -gt $maxLen) { $maxLen = $len }
                }
                if ($maxLen -gt 0) { $wsW.Column($c).Width = [math]::Min([math]::Max($maxLen + 2, 10), 90) }
            }
            Close-ExcelPackage $xlW
        } catch { }
    }
    Add-KpiDataSource -Path $ReportPath -WorksheetName $SheetName -SourceFolder $SourceFolder -BaseName 'RBAC' -Row 2
    return $true
}
