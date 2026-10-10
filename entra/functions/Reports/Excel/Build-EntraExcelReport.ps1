# Autodoc v1.0
function Build-EntraExcelReport {
    <#
        STEP 2 - Builds "KPI <date>_<time>.xlsx": a licence-focused workbook with
        a dashboard and one sheet per analysis, each with its own chart.

        Reads the JSON (CSV fallback) from the most recent export folder unless
        -SourceFolder is given. No Graph connection needed.

        Sheets:
          Dashboard              KPI cards + overview charts
          01 Licenses            full SKU inventory + utilisation % (colour coded)
          02 Users per licence   how many users hold each licence
          03 Assigned vs Free    purchased SKUs, consumed vs unused
          04 Per department      top departments x top licences (matrix)
          05 Top 10 SKU          most-consumed SKUs
          06 Per office          licences per OfficeLocation
          07 Multi-licence users users holding 1, 2, 3, 4+ licences
          08 Groups overview     group types, owners, sizes, empty and ownerless groups
          09 Conditional Access  policy states, targeting, MFA and exclusions
          10 App registrations   owners, tenancy, credential expiry, permissions
          11 Enterprise apps     SSO protocols, Graph usage, sign-in recency
          12 RBAC                role assignments, privileged roles, duration

        On visible numbers: chart data labels are NOT set through the EPPlus
        DataLabel API - writing those elements after the fact produces XML that
        Excel rejects ("workbook repaired"). Instead every chart's category
        label has the value baked into the text, so the numbers are always on
        screen without hovering, and the file stays schema-clean.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportsRoot,
        [string]$SourceFolder,
        [int]$TopDepartments = 10,
        [int]$TopLicenses    = 5
    )

    try { Import-Module ImportExcel -ErrorAction Stop }
    catch { Write-Host "ImportExcel is required for reports: $($_.Exception.Message)" -ForegroundColor Red; return }

    # ---------- resolve source ----------
    if (-not $SourceFolder) {
        if (-not (Test-Path $ExportsRoot)) { Write-Host "No exports found at $ExportsRoot - run a collection first." -ForegroundColor Yellow; return }
        $SourceFolder = (Get-ChildItem -Path $ExportsRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1).FullName
    }
    if (-not $SourceFolder -or -not (Test-Path $SourceFolder)) { Write-Host "No export folder to build a report from." -ForegroundColor Yellow; return }

    Write-Host "Building KPI workbook from: $SourceFolder" -ForegroundColor Cyan

    # Timestamp includes time so repeated runs on the same day don't overwrite.
    # The workbook goes in its own "excelkpi" sub-folder of the export.
    $kpiDir = Join-Path $SourceFolder 'Report\excelkpi'
    if (-not (Test-Path $kpiDir)) { New-Item -Path $kpiDir -ItemType Directory -Force | Out-Null }
    $reportPath = Join-Path $kpiDir ("KPI " + (Get-Date -Format 'yyyy-MM-dd_HH.mm.ss') + ".xlsx")
    if (Test-Path $reportPath) { Remove-Item $reportPath -Force }

    # ---------- load data ----------
    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation')  | Select-Object -First 1
    $skus   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    $users  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    $groups = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    $ca     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess')
    $appReg = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    $rbac   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC')

    # CA is written as a structured object in JSON and a flat table in CSV
    $caPolicyCount = 0
    if ($ca) {
        $caPolicyCount = if (@($ca)[0].PSObject.Properties.Name -contains 'ConditionalAccessPolicies') {
            @(@($ca)[0].ConditionalAccessPolicies).Count
        } else {
            @($ca | Where-Object { $_.RecordType -eq 'Policy' }).Count
        }
    }

    if (-not $skus -and -not $users -and -not $groups) { Write-Host "Neither licence nor user data found - nothing to report." -ForegroundColor Yellow; return }

    # ---------- shared ----------
    function ColLetter([int]$i) {   # 1 -> A
        $s = ''
        while ($i -gt 0) { $m = ($i - 1) % 26; $s = [char](65 + $m) + $s; $i = [int](($i - $m) / 26) }
        $s
    }
    function Pct($part, $whole) { if ($whole) { [math]::Round(100 * $part / $whole, 1) } else { 0 } }
    function Short([string]$t, [int]$max = 34) { if ($t.Length -gt $max) { $t.Substring(0, $max - 1) + '…' } else { $t } }

    # ============================================================
    # AGGREGATES
    # ============================================================
    $skuRows = foreach ($s in $skus) {
        $enabled  = [int]$s.enabledUnits
        $consumed = [int]$s.consumedUnits
        [pscustomobject]@{
            Sku             = $s.skuPartNumber
            Category        = Get-LicenseSkuCategory -SkuPartNumber $s.skuPartNumber -EnabledUnits $enabled
            Enabled         = $enabled
            Consumed        = $consumed
            Free            = $enabled - $consumed
            'Utilisation %' = Pct $consumed $enabled
            Status          = $s.capabilityStatus
        }
    }
    $skuRows   = @($skuRows | Sort-Object Category, @{e='Consumed';d=$true})
    $purchased = @($skuRows | Where-Object Category -eq 'Purchased')

    $licUserCount = @{}
    $multi = [ordered]@{ '0 licences'=0; '1 licence'=0; '2 licences'=0; '3 licences'=0; '4+ licences'=0 }
    $deptLic = @{}
    $officeCount = @{}
    $officeLic = @{}          # office -> licence -> users
    $licensedUsers = @()      # every user holding at least one licence

    foreach ($u in $users) {
        $names = @()
        if ($u.AssignedLicenses) { $names = @($u.AssignedLicenses -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        $n = $names.Count

        switch ($n) {
            0       { $multi['0 licences']++ }
            1       { $multi['1 licence']++ }
            2       { $multi['2 licences']++ }
            3       { $multi['3 licences']++ }
            default { $multi['4+ licences']++ }
        }

        if ($n -gt 0) {
            $dept = if ($u.Department)     { [string]$u.Department }     else { '(none)' }
            $off  = if ($u.OfficeLocation) { [string]$u.OfficeLocation } else { '(none)' }
            if (-not $officeCount.ContainsKey($off)) { $officeCount[$off] = 0 }
            $officeCount[$off] += $n
            if (-not $officeLic.ContainsKey($off)) { $officeLic[$off] = @{} }

            $licensedUsers += [pscustomobject]@{
                'User'              = $u.DisplayName
                'UserPrincipalName' = $u.UserPrincipalName
                'Licences held'     = $n
                'Band'              = $(switch ($n) { 1 { '1 licence' } 2 { '2 licences' } 3 { '3 licences' } default { '4+ licences' } })
                'Status'            = $u.AccountStatus
                'Department'        = $dept
                'Office'            = $off
                'Licences'          = ($names -join '; ')   # last: widened after layout
            }

            foreach ($nm in $names) {
                if (-not $officeLic[$off].ContainsKey($nm)) { $officeLic[$off][$nm] = 0 }
                $officeLic[$off][$nm]++
                if (-not $licUserCount.ContainsKey($nm)) { $licUserCount[$nm] = 0 }
                $licUserCount[$nm]++
                if (-not $deptLic.ContainsKey($dept)) { $deptLic[$dept] = @{} }
                if (-not $deptLic[$dept].ContainsKey($nm)) { $deptLic[$dept][$nm] = 0 }
                $deptLic[$dept][$nm]++
            }
        }
    }

    $licSorted     = @($licUserCount.GetEnumerator() | Sort-Object Value -Descending)
    $totalAssigned = ($licSorted | Measure-Object -Property Value -Sum).Sum
    $usersPerLicence = @($licSorted | ForEach-Object {
        [pscustomobject]@{
            Licence         = $_.Key
            Users           = $_.Value
            'Share %'       = Pct $_.Value $totalAssigned
            'Chart label'   = "$(Short $_.Key) ($($_.Value))"
        }
    })

    # --- group classification (used by both the Dashboard and sheet 08) ---
    $gInfo = @()
    if ($groups) {
        $gInfo = @(foreach ($grp in $groups) {
            $types   = [string]$grp.GroupTypes
            $unified = $types -match 'Unified'
            $secure  = "$($grp.SecurityEnabled)" -eq 'True'
            $mail    = "$($grp.MailEnabled)"     -eq 'True'
            $members = 0; [void][int]::TryParse("$($grp.MemberCount)", [ref]$members)
            $owners  = @()
            if ($grp.Owners) { $owners = @("$($grp.Owners)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

            [pscustomobject]@{
                Id        = [string]$grp.Id
                Name      = [string]$grp.DisplayName
                Kind      = $(if     ($unified)            { 'Microsoft 365' }
                              elseif ($secure -and $mail)  { 'Mail-enabled security' }
                              elseif ($secure)             { 'Security' }
                              elseif ($mail)               { 'Distribution' }
                              else                         { 'Other' })
                Members   = $members
                Owners    = $owners
                OwnerText = ($owners -join '; ')
                IsDynamic = [bool]$grp.DynamicGroupMembershipRule
                Created   = $(if ($grp.CreatedDateTime) { ("$($grp.CreatedDateTime)" -replace 'T', ' ') -replace 'Z','' } else { '' })
            }
        })
    }
    $grpTotal    = $gInfo.Count
    $grpSecurity = @($gInfo | Where-Object Kind -eq 'Security').Count
    $grpM365     = @($gInfo | Where-Object Kind -eq 'Microsoft 365').Count

    Write-Host "  Data: $($skuRows.Count) SKUs ($($purchased.Count) purchased), $($users.Count) users, $($usersPerLicence.Count) licence types, $grpTotal groups, $caPolicyCount CA policies, $($appReg.Count) apps, $($rbac.Count) role records." -ForegroundColor DarkGray

    # ============================================================
    # DASHBOARD
    # ============================================================
    $dashTop = @($usersPerLicence | Select-Object -First 10 |
                 ForEach-Object { [pscustomobject]@{ 'Licence' = $_.'Chart label'; 'Users' = $_.Users } })

    $dashAssigned = @($purchased | Sort-Object Free -Descending | Select-Object -First 12 |
                      ForEach-Object { [pscustomobject]@{
                          'SKU'      = "$($_.Sku)  $($_.Consumed)/$($_.Free)"
                          'Assigned' = $_.Consumed
                          'Unused'   = $_.Free } })

    $dashMulti = @($multi.GetEnumerator() |
                   ForEach-Object { [pscustomobject]@{ 'Holding' = "$($_.Key) ($($_.Value))"; 'Users' = $_.Value } })

    $dashCat = @(
        [pscustomobject]@{ SkuCategory = 'Purchased';  SkuCount = @($skuRows | Where-Object Category -eq 'Purchased').Count }
        [pscustomobject]@{ SkuCategory = 'Free/Viral'; SkuCount = @($skuRows | Where-Object Category -eq 'Free/Viral').Count }
    )

    # Summary tables sit under the heading on the left (column A), stacked one
    # under the other with a blank gap, each under its own bold title.
    $srcRow = 58
    $dashBlocks = @(
        @{ Table = $dashTop;      Title = 'Users per licence (top 10)' }
        @{ Table = $dashAssigned; Title = 'Assigned vs unused' }
        @{ Table = $dashMulti;    Title = 'Users by number of licences' }
        @{ Table = $dashCat;      Title = 'SKU categories' }
    )
    $rowNow = $srcRow
    foreach ($b in $dashBlocks) {
        if (-not $b.Table) { continue }
        $b['Row'] = $rowNow                       # heading row; table starts one below
        $rowNow += @($b.Table).Count + 4
    }
    $dashTopRow      = @($dashBlocks | Where-Object { $_.Title -eq 'Users per licence (top 10)' })[0].Row
    $dashAssignedRow = @($dashBlocks | Where-Object { $_.Title -eq 'Assigned vs unused' })[0].Row
    $dashMultiRow    = @($dashBlocks | Where-Object { $_.Title -eq 'Users by number of licences' })[0].Row
    if ($dashTop)      { $dashTop      | Export-Excel -Path $reportPath -WorksheetName 'Dashboard' -StartRow ($dashTopRow + 1)      -StartColumn 1 }
    if ($dashAssigned) { $dashAssigned | Export-Excel -Path $reportPath -WorksheetName 'Dashboard' -StartRow ($dashAssignedRow + 1) -StartColumn 1 }
    if ($dashMulti)    { $dashMulti    | Export-Excel -Path $reportPath -WorksheetName 'Dashboard' -StartRow ($dashMultiRow + 1)    -StartColumn 1 }

    $charts = @()
    if ($dashTop) {
        $n = $dashTop.Count
        $charts += New-ExcelChartDefinition -Title 'Users per licence (top 10)' -ChartType ColumnClustered `
                    -XRange "A$($dashTopRow+2):A$($dashTopRow+1+$n)" -YRange "B$($dashTopRow+2):B$($dashTopRow+1+$n)" `
                    -SeriesHeader 'Users' -Width 520 -Height 340 -Row 10 -Column 1
        $charts += New-ExcelChartDefinition -Title 'Licence share' -ChartType Pie `
                    -XRange "A$($dashTopRow+2):A$($dashTopRow+1+$n)" -YRange "B$($dashTopRow+2):B$($dashTopRow+1+$n)" `
                    -ShowPercent -Width 520 -Height 340 -Row 10 -Column 11
    }
    if ($dashAssigned) {
        $n = $dashAssigned.Count
        $charts += New-ExcelChartDefinition -Title 'Assigned vs unused  (label shows assigned/unused)' -ChartType BarStacked `
                    -XRange "A$($dashAssignedRow+2):A$($dashAssignedRow+1+$n)" -YRange @("B$($dashAssignedRow+2):B$($dashAssignedRow+1+$n)", "C$($dashAssignedRow+2):C$($dashAssignedRow+1+$n)") `
                    -SeriesHeader 'Assigned','Unused' -Width 520 -Height 340 -Row 32 -Column 1
    }
    if ($dashMulti) {
        $n = $dashMulti.Count
        $charts += New-ExcelChartDefinition -Title 'Users by number of licences' -ChartType Doughnut `
                    -XRange "A$($dashMultiRow+2):A$($dashMultiRow+1+$n)" -YRange "B$($dashMultiRow+2):B$($dashMultiRow+1+$n)" `
                    -ShowPercent -Width 520 -Height 340 -Row 32 -Column 11
    }
    $dashCatRow = @($dashBlocks | Where-Object { $_.Title -eq 'SKU categories' })[0].Row
    $dashCat | Export-Excel -Path $reportPath -WorksheetName 'Dashboard' -StartRow ($dashCatRow + 1) -StartColumn 1 -ExcelChartDefinition $charts

    # ---- title + KPI cards ----
    $purEnabled  = ($purchased | Measure-Object Enabled  -Sum).Sum
    $purConsumed = ($purchased | Measure-Object Consumed -Sum).Sum
    $purFree     = $purEnabled - $purConsumed

    $xl = Open-ExcelPackage -Path $reportPath
    $ws = $xl.Workbook.Worksheets['Dashboard']

    $tenantName = $(if ($tenant) { $tenant.TenantName } else { 'Unknown tenant' })
    $tenantId   = $(if ($tenant) { $tenant.TenantId }   else { 'n/a' })
    $ws.Cells['A1'].Value = "KPI - $tenantName    Tenant ID : $tenantId"
    $ws.Cells['A1'].Style.Font.Size = 16
    $ws.Cells['A1'].Style.Font.Bold = $true

    # Full source path, not just the folder name
    $srcShort = "$SourceFolder" -replace '^.*?\\(Entra\\)', '$1'
    $ws.Cells['A2'].Value = "Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')   |   Source: $srcShort"
    $ws.Cells['A2'].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)

    # Counts for the application and role cards
    $appRegTotal = @($appReg).Count
    $appWithSp   = @($appReg | Where-Object { "$($_.ServicePrincipalExists)" -eq 'Yes' }).Count
    $roleAssign  = @($rbac | Where-Object { $_.RecordType -like 'RoleAssignment*' })
    $roleDistinct= @($roleAssign | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count
    $privRoleList = @(
        'Global Administrator','Privileged Role Administrator','Privileged Authentication Administrator',
        'Security Administrator','Conditional Access Administrator','Application Administrator',
        'Cloud Application Administrator','User Administrator','Exchange Administrator',
        'SharePoint Administrator','Intune Administrator','Hybrid Identity Administrator',
        'Domain Name Administrator','Directory Synchronization Accounts','Partner Tier2 Support'
    )
    $privAssignCount = @($roleAssign | Where-Object { $privRoleList -contains "$($_.RoleDefinitionName)" }).Count

    # Thirteen cards don't fit on one line, so they run over two rows
    $kpiRow1 = @(
        @{ Label = 'Users';                Value = $(if ($tenant) { $tenant.Users }   else { $users.Count }) }
        @{ Label = 'Devices';              Value = $(if ($tenant) { $tenant.Devices } else { 'n/a' }) }
        @{ Label = 'Assigned licences';    Value = $totalAssigned }
        @{ Label = 'Purchased SKUs';       Value = $purchased.Count }
        @{ Label = 'Unused units';         Value = $purFree }
        @{ Label = 'Total groups';         Value = $grpTotal }
        @{ Label = 'Security groups';      Value = $grpSecurity }
    )
    $kpiRow2 = @(
        @{ Label = 'Microsoft 365 groups';    Value = $grpM365 }
        @{ Label = 'Total app registrations'; Value = $appRegTotal }
        @{ Label = 'Apps with service principal'; Value = $appWithSp }
        @{ Label = 'Distinct roles';          Value = $roleDistinct }
        @{ Label = 'Role assignments';        Value = $roleAssign.Count }
        @{ Label = 'Privileged assignments';  Value = $privAssignCount }
    )
    foreach ($block in @(@{ Cards = $kpiRow1; Row = 4 }, @{ Cards = $kpiRow2; Row = 7 })) {
        $col = 1
        foreach ($k in $block.Cards) {
            $L = ColLetter $col
            $ws.Cells["${L}$($block.Row)"].Value = $k.Label
            $ws.Cells["${L}$($block.Row)"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
            $ws.Cells["${L}$($block.Row + 1)"].Value = $k.Value
            $ws.Cells["${L}$($block.Row + 1)"].Style.Font.Size = 20
            $ws.Cells["${L}$($block.Row + 1)"].Style.Font.Bold = $true
            $ws.Column($col).Width = 18
            $col += 2
        }
    }
    $ws.Cells["A$($srcRow-1)"].Value = 'Summary tables - the numbers behind the charts above'
    $ws.Cells["A$($srcRow-1)"].Style.Font.Bold = $true
    foreach ($b in $dashBlocks) {
        if (-not $b.Table) { continue }
        $ws.Cells["A$($b.Row)"].Value = $b.Title
        $ws.Cells["A$($b.Row)"].Style.Font.Bold = $true
    }
    Close-ExcelPackage $xl
    Set-KpiChartColors -Path $reportPath -WorksheetName 'Dashboard'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName 'Dashboard' -ChartZoneColumns 13 -ChartZoneWidth 18
    Write-Host "  Dashboard added." -ForegroundColor DarkGreen

    # ============================================================
    # 01  LICENSES - inventory with a colour-coded utilisation column
    # ============================================================
    if ($skuRows) {
        $n = $skuRows.Count
        $c = New-ExcelChartDefinition -Title 'Utilisation % by SKU' -ChartType BarClustered `
                -XRange "M5:M$($n+4)" -YRange "R5:R$($n+4)" -SeriesHeader 'Utilisation %' `
                -Width 520 -Height (25 * $n + 120) -Row 3 -Column 1
        $xl = $skuRows | Export-Excel -Path $reportPath -WorksheetName '01 Licenses' -TableName 'tblLicenses' `
                -StartColumn 13 -StartRow 4 -BoldTopRow -FreezeTopRow -ExcelChartDefinition $c -PassThru
        $ws = $xl.Workbook.Worksheets['01 Licenses']

        # Colours are applied as plain cell fills rather than conditional
        # formatting: the bands are then exact (so the legend can state them)
        # and no extra XML rules are written into the file.
        $bands = @(
            @{ Min =  0; Max = 49.9;  Rgb = @(248,105,107); Text = 'Under 50 % - most units unused' }
            @{ Min = 50; Max = 79.9;  Rgb = @(255,235,132); Text = '50-79 % - roughly half to three quarters in use' }
            @{ Min = 80; Max = 100;   Rgb = @( 99,190,123); Text = '80-100 % - close to fully used' }
        )
        for ($i = 0; $i -lt $n; $i++) {
            $val  = [double]$skuRows[$i].'Utilisation %'
            $band = $bands | Where-Object { $val -ge $_.Min -and $val -le $_.Max } | Select-Object -First 1
            if ($band) {
                $cell = $ws.Cells["R$($i + 5)"]
                $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb($band.Rgb[0], $band.Rgb[1], $band.Rgb[2]))
            }
        }

        # Legend so the colours mean something without guessing
        $lr = $n + 7
        $ws.Cells["M$lr"].Value = 'What the colours in "Utilisation %" mean'
        $ws.Cells["M$lr"].Style.Font.Bold = $true
        $row = $lr + 1
        foreach ($band in $bands) {
            $ws.Cells["M$row"].Value = "$($band.Min)-$([math]::Ceiling($band.Max)) %"
            $ws.Cells["M$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $ws.Cells["M$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb($band.Rgb[0], $band.Rgb[1], $band.Rgb[2]))
            $ws.Cells["N$row"].Value = $band.Text
            $row++
        }
        Close-ExcelPackage $xl
        Set-KpiChartColors -Path $reportPath -WorksheetName '01 Licenses'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '01 Licenses'
        Add-KpiDataSource -Path $reportPath -WorksheetName '01 Licenses' -SourceFolder $SourceFolder -BaseName 'LicensesInformation' -Title 'LICENSES'
        Write-Host "  01 Licenses added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 02  USERS PER LICENCE
    # ============================================================
    if ($usersPerLicence) {
        $top = [math]::Min($usersPerLicence.Count, 15)
        $c1 = New-ExcelChartDefinition -Title 'Users per licence' -ChartType ColumnClustered `
                -XRange "P5:P$($top+4)" -YRange "N5:N$($top+4)" -SeriesHeader 'Users' `
                -Width 520 -Height 340 -Row 3 -Column 1
        $c2 = New-ExcelChartDefinition -Title 'Licence share' -ChartType Pie `
                -XRange "P5:P$($top+4)" -YRange "N5:N$($top+4)" -ShowPercent `
                -Width 520 -Height 340 -Row 22 -Column 1
        $usersPerLicence | Export-Excel -Path $reportPath -WorksheetName '02 Users per licence' `
            -StartColumn 13 -StartRow 4 -TableName 'tblUsersPerLicence' -BoldTopRow -FreezeTopRow -ExcelChartDefinition @($c1, $c2)
        Set-KpiChartColors -Path $reportPath -WorksheetName '02 Users per licence'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '02 Users per licence'
        Add-KpiDataSource -Path $reportPath -WorksheetName '02 Users per licence' -SourceFolder $SourceFolder -BaseName 'UserInformation', 'LicensesInformation' -Title 'USERS PER LICENCE'
        Write-Host "  02 Users per licence added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 03  ASSIGNED vs FREE  (purchased SKUs only)
    # ============================================================
    if ($purchased) {
        $rows3 = @($purchased | Sort-Object Enabled -Descending | ForEach-Object {
            [pscustomobject]@{
                Sku             = $_.Sku
                Consumed        = $_.Consumed
                Free            = $_.Free
                Enabled         = $_.Enabled
                'Utilisation %' = $_.'Utilisation %'
                'Chart label'   = "$($_.Sku)  $($_.Consumed)/$($_.Free)"
            }
        })
        $n = $rows3.Count
        $c = New-ExcelChartDefinition -Title 'Assigned vs unused  (label shows assigned/unused)' -ChartType BarStacked `
                -XRange "R5:R$($n+4)" -YRange @("N5:N$($n+4)", "O5:O$($n+4)") -SeriesHeader 'Assigned','Unused' `
                -Width 520 -Height (24 * $n + 120) -Row 3 -Column 1
        $rows3 | Export-Excel -Path $reportPath -WorksheetName '03 Assigned vs Free' `
            -StartColumn 13 -StartRow 4 `
            -TableName 'tblAssignedFree' -BoldTopRow -FreezeTopRow -ExcelChartDefinition $c
        Set-KpiChartColors -Path $reportPath -WorksheetName '03 Assigned vs Free'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '03 Assigned vs Free'
        Add-KpiDataSource -Path $reportPath -WorksheetName '03 Assigned vs Free' -SourceFolder $SourceFolder -BaseName 'LicensesInformation' -Title 'ASSIGNED VS FREE'
        Write-Host "  03 Assigned vs Free added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 04  PER DEPARTMENT
    # ============================================================
    if ($deptLic.Count -gt 0) {
        $topLicNames = @($usersPerLicence | Select-Object -First $TopLicenses -ExpandProperty Licence)
        $deptTotals  = $deptLic.GetEnumerator() | ForEach-Object {
            [pscustomobject]@{ Dept = $_.Key; Total = ($_.Value.Values | Measure-Object -Sum).Sum } }
        $topDepts = @($deptTotals | Sort-Object Total -Descending | Select-Object -First $TopDepartments -ExpandProperty Dept)

        $headers = @(foreach ($l in $topLicNames) { Short $l 31 })
        $matrix = foreach ($d in $topDepts) {
            $o = [ordered]@{ Department = $d }
            for ($i = 0; $i -lt $topLicNames.Count; $i++) {
                $l = $topLicNames[$i]
                $o[$headers[$i]] = $(if ($deptLic[$d].ContainsKey($l)) { $deptLic[$d][$l] } else { 0 })
            }
            [pscustomobject]$o
        }
        $matrix = @($matrix); $n = $matrix.Count
        $yRanges = @(); for ($i = 14; $i -le (13 + $topLicNames.Count); $i++) { $L = ColLetter $i; $yRanges += "${L}5:${L}$($n+4)" }

        # Stacked vertically: matrix at the top left, chart directly beneath it
        $c = New-ExcelChartDefinition -Title 'Licences per department' -ChartType ColumnClustered `
                -XRange "M5:M$($n+4)" -YRange $yRanges -SeriesHeader $headers `
                -Width 620 -Height 420 -Row 3 -Column 1
        $matrix | Export-Excel -Path $reportPath -WorksheetName '04 Per department' `
            -StartColumn 13 -StartRow 4 -TableName 'tblPerDepartment' -BoldTopRow -FreezeTopRow -ExcelChartDefinition $c
        Set-KpiChartColors -Path $reportPath -WorksheetName '04 Per department'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '04 Per department'
        Add-KpiDataSource -Path $reportPath -WorksheetName '04 Per department' -SourceFolder $SourceFolder -BaseName 'UserInformation' -Title 'PER DEPARTMENT'
        Write-Host "  04 Per department added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 05  TOP 10 SKU
    # ============================================================
    if ($purchased) {
        $rows5 = @($purchased | Sort-Object Consumed -Descending | Select-Object -First 10 | ForEach-Object {
            [pscustomobject]@{
                Sku             = $_.Sku
                Consumed        = $_.Consumed
                Free            = $_.Free
                'Utilisation %' = $_.'Utilisation %'
                'Chart label'   = "$($_.Sku) ($($_.Consumed))"
            }
        })
        $n = $rows5.Count
        $c = New-KpiChart -Title 'Top 10 SKUs by assigned units' -ChartType BarClustered `
                -LabelColumn 'Q' -ValueColumn 'N' -FirstRow 5 -RowCount $n `
                -Row 3 -Column 1 -Height 400 -SeriesHeader 'Assigned'
        $rows5 | Export-Excel -Path $reportPath -WorksheetName '05 Top 10 SKU' `
            -StartColumn 13 -StartRow 4 -TableName 'tblTop10Sku' -BoldTopRow -FreezeTopRow -ExcelChartDefinition $c
        Set-KpiChartColors -Path $reportPath -WorksheetName '05 Top 10 SKU'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '05 Top 10 SKU'
        Add-KpiDataSource -Path $reportPath -WorksheetName '05 Top 10 SKU' -SourceFolder $SourceFolder -BaseName 'LicensesInformation' -Title 'TOP 10 SKU'
        Write-Host "  05 Top 10 SKU added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 06  PER OFFICE LOCATION
    # ============================================================
    if ($officeCount.Count -gt 0) {
        # The table lists every office. The chart shows the largest 40: a bar
        # per office would run to thousands of pixels and render as an
        # unreadable smear, so the full list lives in the table below it.
        $all6 = @($officeCount.GetEnumerator() | Sort-Object Value -Descending |
                  ForEach-Object { [pscustomobject]@{
                      Office        = $_.Key
                      Licences      = $_.Value
                      'Share %'     = Pct $_.Value $totalAssigned
                      'Chart label' = "$(Short $_.Key 28) ($($_.Value))" } })
        $chartRows = [math]::Min($all6.Count, 40)

        $c = New-ExcelChartDefinition -Title "Licences per office location (largest $chartRows of $($all6.Count))" `
                -ChartType BarClustered -TitleBold -TitleSize 13 `
                -XRange "P5:P$($chartRows+4)" -YRange "N5:N$($chartRows+4)" -SeriesHeader 'Licences' `
                -Width 520 -Height (18 * $chartRows + 140) -Row 3 -Column 1
        $all6 | Export-Excel -Path $reportPath -WorksheetName '06 Per office' `
            -StartColumn 13 -StartRow 4 -TableName 'tblPerOffice' -BoldTopRow -FreezeTopRow -ExcelChartDefinition $c

        # Which licences are used at each office - filter the Office column
        $officeDetail = @(foreach ($off in ($officeLic.Keys | Sort-Object)) {
            foreach ($lic in ($officeLic[$off].Keys | Sort-Object { -$officeLic[$off][$_] })) {
                [pscustomobject]@{
                    'Office'   = $off
                    'Licence'  = $lic
                    'Users'    = $officeLic[$off][$lic]
                    'Share of office %' = Pct $officeLic[$off][$lic] $officeCount[$off]
                }
            }
        })
        $detailRow6 = $all6.Count + 7
        if ($officeDetail) {
            $officeDetail | Export-Excel -Path $reportPath -WorksheetName '06 Per office' `
                -StartRow $detailRow6 -StartColumn 13 -TableName 'tblOfficeLicences' -BoldTopRow
            $xl6 = Open-ExcelPackage -Path $reportPath
            $ws6 = $xl6.Workbook.Worksheets['06 Per office']
            $ws6.Cells["M$($detailRow6-1)"].Value = "Licences used per office location ($($officeDetail.Count) rows - filter the Office column)"
            $ws6.Cells["M$($detailRow6-1)"].Style.Font.Bold = $true
            Close-ExcelPackage $xl6
        }
        Set-KpiChartColors -Path $reportPath -WorksheetName '06 Per office'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '06 Per office'
        Add-KpiDataSource -Path $reportPath -WorksheetName '06 Per office' -SourceFolder $SourceFolder -BaseName 'UserInformation' -Title 'PER OFFICE'
        Write-Host "  06 Per office added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 07  MULTI-LICENCE USERS  (column + doughnut)
    # ============================================================
    if ($users) {
        $totalUsers = ($multi.Values | Measure-Object -Sum).Sum
        $rows7 = @($multi.GetEnumerator() | ForEach-Object {
            [pscustomobject]@{
                Holding       = $_.Key
                Users         = $_.Value
                'Share %'     = Pct $_.Value $totalUsers
                'Chart label' = "$($_.Key) ($($_.Value))"
            }
        })
        $n = $rows7.Count
        # Vertical stack: summary table, then the two charts, then the user list
        $c1 = New-KpiChart -Title 'Users by number of licences' -ChartType ColumnClustered `
                -LabelColumn 'P' -ValueColumn 'N' -FirstRow 4 -RowCount $n `
                -Row 3 -Column 1 -SeriesHeader 'Users'
        $c2 = New-ExcelChartDefinition -Title 'Share of users' -ChartType Doughnut `
                -XRange "P4:P$($n+3)" -YRange "N4:N$($n+3)" -ShowPercent `
                -Width 520 -Height 340 -Row 22 -Column 1 -TitleBold -TitleSize 13
        $rows7 | Export-Excel -Path $reportPath -WorksheetName '07 Multi-licence users' `
            -StartColumn 13 -StartRow 3 -TableName 'tblMultiLicence' -BoldTopRow -FreezeTopRow -ExcelChartDefinition @($c1, $c2)

        # Every licensed user with the band they fall into, so each slice of the
        # chart can be traced to the actual people. Users with no licence are
        # not listed - there is nothing to show for them.
        $userDetail = @($licensedUsers | Sort-Object 'Licences held' -Descending)
        $detailRow7 = 11
        if ($userDetail) {
            $userDetail | Export-Excel -Path $reportPath -WorksheetName '07 Multi-licence users' `
                -StartRow $detailRow7 -StartColumn 13 -TableName 'tblLicensedUsers' -BoldTopRow
            $xl7 = Open-ExcelPackage -Path $reportPath
            $ws7 = $xl7.Workbook.Worksheets['07 Multi-licence users']
            $ws7.Cells["M$($detailRow7-1)"].Value = "Licensed users and what they hold ($($userDetail.Count) users - filter the Band column)"
            $ws7.Cells["M$($detailRow7-1)"].Style.Font.Bold = $true
            Close-ExcelPackage $xl7
        }

        Set-KpiChartColors -Path $reportPath -WorksheetName '07 Multi-licence users'
        Set-KpiSheetLayout -Path $reportPath -WorksheetName '07 Multi-licence users'
        # Licences is the last column and holds a long list; the layout pass caps
        # every column at 55, so widen this one afterwards to show it in full.
        if ($userDetail) {
            try {
                $maxLic = ($userDetail | ForEach-Object { "$($_.Licences)".Length } | Measure-Object -Maximum).Maximum
                $xlW = Open-ExcelPackage -Path $reportPath
                $wsW = $xlW.Workbook.Worksheets['07 Multi-licence users']
                $wsW.Column(20).Width = [math]::Min([math]::Max($maxLic + 2, 40), 255)   # column T = Licences
                Close-ExcelPackage $xlW
            } catch { }
        }
        Add-KpiDataSource -Path $reportPath -WorksheetName '07 Multi-licence users' -SourceFolder $SourceFolder -BaseName 'UserInformation' -Title 'MULTI-LICENCE USERS'
        Write-Host "  07 Multi-licence users added." -ForegroundColor DarkGreen
    }


    # ============================================================
    # 08  GROUPS OVERVIEW
    # ============================================================
    if ($gInfo) {
        $noOwner   = @($gInfo | Where-Object { $_.Owners.Count -eq 0 })
        $emptyGrps = @($gInfo | Where-Object { $_.Members -eq 0 })
        $dynCount  = @($gInfo | Where-Object IsDynamic).Count
        $totalGrp  = $gInfo.Count

        # Labels carry both the count and the share, so every slice shows its
        # numbers on screen without hovering and without touching chart XML.
        function GLabel([string]$text, [int]$count) { "$text  $count ($(Pct $count $totalGrp) %)" }
        function GRow([string]$name, [int]$count, [string]$col) {
            # Real columns first, combined chart label last (hidden later)
            [pscustomobject]@{ $col = $name; 'Groups' = $count; 'Share %' = (Pct $count $totalGrp); 'Chart label' = (GLabel $name $count) }
        }

        $kindOrder = 'Security','Microsoft 365','Distribution','Mail-enabled security','Other'
        $tblKind = @(foreach ($k in $kindOrder) {
            $cnt = @($gInfo | Where-Object Kind -eq $k).Count
            if ($cnt -gt 0) { GRow $k $cnt 'Group type' }
        })

        $ownerCount = @{}
        foreach ($item in $gInfo) {
            foreach ($o in $item.Owners) {
                if (-not $ownerCount.ContainsKey($o)) { $ownerCount[$o] = 0 }
                $ownerCount[$o]++
            }
        }
        $tblOwner = @($ownerCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
                      ForEach-Object { [pscustomobject]@{ 'Owner' = $_.Key; 'Groups' = $_.Value; 'Share %' = (Pct $_.Value $totalGrp); 'Chart label' = (GLabel (Short $_.Key 26) $_.Value) } })

        $bandDefs = @(
            @{ Label = '0 (empty)'; Test = { param($m) $m -eq 0 } }
            @{ Label = '1-10';      Test = { param($m) $m -ge 1   -and $m -le 10 } }
            @{ Label = '11-100';    Test = { param($m) $m -ge 11  -and $m -le 100 } }
            @{ Label = '101-500';   Test = { param($m) $m -ge 101 -and $m -le 500 } }
            @{ Label = '500+';      Test = { param($m) $m -gt 500 } }
        )
        $tblSize = @(foreach ($b in $bandDefs) {
            $cnt = @($gInfo | Where-Object { & $b.Test $_.Members }).Count
            GRow $b.Label $cnt 'Size'
        })

        $tblDyn = @(
            GRow 'Dynamic' $dynCount ('Membership')
            GRow 'Static'  ($totalGrp - $dynCount) ('Membership')
        )

        # --- groups by the owner's department / office location ---
        # Owners are stored as UPNs, so look each one up in the user export.
        # A group counts once per distinct department (or office) among its owners.
        $userByUpn = @{}
        foreach ($u in $users) {
            $upnKey = "$($u.UserPrincipalName)".Trim().ToLower()
            if ($upnKey -and -not $userByUpn.ContainsKey($upnKey)) { $userByUpn[$upnKey] = $u }
        }
        $ownerInfo = {
            param($ownerList)
            $depts = @{}; $offs = @{}
            foreach ($o in $ownerList) {
                $usr = $userByUpn["$o".Trim().ToLower()]
                if (-not $usr) { continue }
                $depts[$(if ($usr.Department)     { [string]$usr.Department }     else { '(none)' })] = $true
                $offs[ $(if ($usr.OfficeLocation) { [string]$usr.OfficeLocation } else { '(none)' })] = $true
            }
            [pscustomobject]@{ Departments = @($depts.Keys); Offices = @($offs.Keys) }
        }
        $deptGrpCount = @{}; $offGrpCount = @{}
        foreach ($item in $gInfo) {
            $oi = & $ownerInfo $item.Owners
            foreach ($d in $oi.Departments) { if (-not $deptGrpCount.ContainsKey($d)) { $deptGrpCount[$d] = 0 }; $deptGrpCount[$d]++ }
            foreach ($o in $oi.Offices)     { if (-not $offGrpCount.ContainsKey($o))  { $offGrpCount[$o]  = 0 }; $offGrpCount[$o]++ }
        }
        $deptGrpTotal = ($deptGrpCount.Values | Measure-Object -Sum).Sum
        $offGrpTotal  = ($offGrpCount.Values  | Measure-Object -Sum).Sum
        $tblGrpDept = @($deptGrpCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
            ForEach-Object { [pscustomobject]@{ 'Department' = $_.Key; 'Groups' = $_.Value
                                                'Share %' = (Pct $_.Value $deptGrpTotal)
                                                'Chart label' = "$(Short $_.Key 26)  $($_.Value)" } })
        $tblGrpOffice = @($offGrpCount.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 |
            ForEach-Object { [pscustomobject]@{ 'Office location' = $_.Key; 'Groups' = $_.Value
                                                'Share %' = (Pct $_.Value $offGrpTotal)
                                                'Chart label' = "$(Short $_.Key 26)  $($_.Value)" } })

        # The group list sits beside the charts (right), and the small summary
        # tables behind the charts sit below the chart grid on the left.
        $gSrc = 6
        $gSrcCol = 1
        $gDetailCol = 22
        $sheet08 = '08 Groups overview'
        $gStrips = @(
            @{ Table = $tblKind; Names = 'Group type'; Title = 'Group types';              Type = 'Pie' }
            @{ Table = $tblOwner; Names = 'Owner'; Title = 'Groups per owner (top 10)'; Type = 'BarClustered' }
            @{ Table = $tblSize; Names = 'Size'; Title = 'Members per group (size)'; Type = 'Doughnut' }
            @{ Table = $tblDyn; Names = 'Membership'; Title = 'Dynamic vs static groups'; Type = 'Doughnut' }
            @{ Table = $tblGrpDept;   Names = 'Department';      Title = 'Groups owned per department (top 10)';      Type = 'BarClustered' }
            @{ Table = $tblGrpOffice; Names = 'Office location'; Title = 'Groups owned per office location (top 10)'; Type = 'BarClustered' }
        )
        $gLive = @($gStrips | Where-Object { $_.Table })
        $gHeights = @($gLive | ForEach-Object { 340 })
        $gGrid = New-KpiChartGrid -Heights $gHeights -StartRow $gSrc

        $gCharts = @()
        $gSrcRow = $gGrid.NextRow        # summary tables start below the charts
        for ($gIdx = 0; $gIdx -lt $gLive.Count; $gIdx++) {
            $strip  = $gLive[$gIdx]
            $n      = @($strip.Table).Count
            $L      = @(0,1,2,3) | ForEach-Object { ConvertTo-ExcelColumnLetter -Index ($gSrcCol + $_) }
            $anchor = $gGrid.Anchors[$gIdx]
            $strip['Row'] = $gSrcRow
            $gSrcRow += $n + 4
            $gCharts += New-KpiChart -Title $strip.Title -ChartType $strip.Type `
                            -LabelColumn $L[3] -ValueColumn $L[1] -FirstRow ($strip.Row + 2) -RowCount $n `
                            -Row $anchor.Row -Column $anchor.Column -SeriesHeader 'Groups'
        }
        for ($i = 0; $i -lt $gLive.Count; $i++) {
            $strip = $gLive[$i]
            if ($i -eq $gLive.Count - 1) {
                $strip.Table | Export-Excel -Path $reportPath -WorksheetName $sheet08 -StartRow ($strip.Row + 1) -StartColumn $gSrcCol -ExcelChartDefinition $gCharts
            }
            else {
                $strip.Table | Export-Excel -Path $reportPath -WorksheetName $sheet08 -StartRow ($strip.Row + 1) -StartColumn $gSrcCol
            }
        }

        # --- ALL groups, enriched with member Department / Office / Manager ---
        # Department, OfficeLocation and Manager live per member in GroupMembers,
        # so aggregate them per group. A group can span hundreds of departments/
        # offices, so show the most common few (by frequency) + "(+N more)".
        $groupMembers = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'GroupMembers')
        $memAgg = @{}
        foreach ($m in $groupMembers) {
            $gid = "$($m.GroupId)"
            if (-not $gid) { continue }
            if (-not $memAgg.ContainsKey($gid)) { $memAgg[$gid] = @{ Dept = @{}; Off = @{}; Mgr = @{} } }
            $a = $memAgg[$gid]
            $d = "$($m.Department)".Trim();      if ($d) { $a.Dept[$d] = 1 + ($a.Dept[$d]) }
            $o = "$($m.OfficeLocation)".Trim();  if ($o) { $a.Off[$o]  = 1 + ($a.Off[$o]) }
            $g = "$($m.ManagerDisplayName)".Trim(); if (-not $g) { $g = "$($m.Manager)".Trim() }
            if ($g) { $a.Mgr[$g] = 1 + ($a.Mgr[$g]) }
        }
        $topN = {
            param($counts, $max = 4)
            if (-not $counts -or $counts.Count -eq 0) { return '(none)' }
            $ordered = @($counts.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { $_.Key })
            $shown = @($ordered | Select-Object -First $max)
            $text = $shown -join '; '
            if ($ordered.Count -gt $max) { $text += "  (+$($ordered.Count - $max) more)" }
            $text
        }

        # welcome-email flag + a raw lookup for the many EntraGroups fields
        $welcomeById = @{}
        foreach ($w in @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'GroupWelcomeEmail')) {
            if ($w.GroupId) { $welcomeById["$($w.GroupId)"] = "$($w.WelcomeEmailEnabled)" }
        }
        $rawById = @{}
        foreach ($grp in $groups) { if ($grp.Id) { $rawById["$($grp.Id)"] = $grp } }

        $emptyRow = $gSrc        # group list sits beside the charts
        $emptyAll = @($gInfo | Sort-Object Name | ForEach-Object {
            $agg = $memAgg["$($_.Id)"]
            $raw = $rawById["$($_.Id)"]
            $yn  = { param($v) if ("$v" -eq 'True') { 'Yes' } elseif ("$v" -eq 'False') { 'No' } else { '' } }
            [pscustomobject]@{
                'Group name'       = $_.Name
                'Type'             = $_.Kind
                'Category'         = "$($raw.GroupCategory)"
                'Members'          = $_.Members
                'Status'           = $(if ($_.Members -eq 0) { 'Empty' } else { 'Has members' })
                'Department'       = $(if ($agg) { & $topN $agg.Dept } else { '(none)' })
                'Office location'  = $(if ($agg) { & $topN $agg.Off } else { '(none)' })
                'Source'           = "$($raw.Source)"
                'Visibility'       = "$($raw.Visibility)"
                'Dynamic'          = $(if ($_.IsDynamic) { 'Yes' } else { 'No' })
                'Dynamic rule'     = "$($raw.DynamicGroupMembershipRule)"
                'Mail enabled'     = (& $yn $raw.MailEnabled)
                'Security enabled' = (& $yn $raw.SecurityEnabled)
                'Mail nickname'    = "$($raw.MailNickname)"
                'Assignable to role' = (& $yn $raw.IsAssignableToRole)
                'On-prem synced'   = (& $yn $raw.OnPremisesSyncEnabled)
                'Licenses'         = "$($raw.Licenses)"
                'Roles & admins'   = "$($raw.RolesAndAdministrators)"
                'Applications'     = "$($raw.Applications)"
                'Welcome email'    = $(if ($welcomeById.ContainsKey("$($_.Id)")) { $welcomeById["$($_.Id)"] } else { "$($raw.WelcomeEmailEnabled)" })
                'Description'      = "$($raw.Description)"
                'Created'          = $_.Created
            }
        })
        if ($emptyAll) {
            $emptyAll | Export-Excel -Path $reportPath -WorksheetName $sheet08 `
                -StartRow $emptyRow -StartColumn $gDetailCol -TableName 'tblAllGroups' -BoldTopRow
        }

        # --- title + KPI cards ---
        $xl = Open-ExcelPackage -Path $reportPath
        $ws = $xl.Workbook.Worksheets[$sheet08]

        $ws.Cells['A1'].Value = 'GROUPS - OVERVIEW'
        $ws.Cells['A1'].Style.Font.Size = 16
        $ws.Cells['A1'].Style.Font.Bold = $true

        $gBasic = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroupsBasicInfo') | Select-Object -First 1
        $gKpis = @(
            @{ Label = 'Total groups';      Value = $(if ($gBasic) { $gBasic.TotalGroups } else { $totalGrp }) }
            @{ Label = 'Security';          Value = $(if ($gBasic) { $gBasic.SecurityGroups } else { @($gInfo | Where-Object Kind -eq 'Security').Count }) }
            @{ Label = 'Microsoft 365';     Value = $(if ($gBasic) { $gBasic.M365Groups } else { @($gInfo | Where-Object Kind -eq 'Microsoft 365').Count }) }
            @{ Label = 'Distribution';      Value = @($gInfo | Where-Object Kind -eq 'Distribution').Count }
            @{ Label = 'Dynamic';           Value = $(if ($gBasic) { $gBasic.DynamicGroups } else { @($gInfo | Where-Object IsDynamic).Count }) }
            @{ Label = 'Cloud';             Value = $(if ($gBasic) { $gBasic.CloudGroups } else { '' }) }
            @{ Label = 'On-premises';       Value = $(if ($gBasic) { $gBasic.OnPremisesGroups } else { '' }) }
            @{ Label = 'Without owner';     Value = $noOwner.Count }
            @{ Label = 'Empty (0 members)'; Value = $emptyGrps.Count }
        )
        $col = 1
        foreach ($k in $gKpis) {
            $L = ColLetter $col
            $ws.Cells["${L}3"].Value = $k.Label
            $ws.Cells["${L}3"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
            $ws.Cells["${L}4"].Value = $k.Value
            $ws.Cells["${L}4"].Style.Font.Size = 20
            $ws.Cells["${L}4"].Style.Font.Bold = $true
            $ws.Column($col).Width = 16
            $col += 2
        }

        $ws.Cells["V$($emptyRow-1)"].Value = "All groups ($($emptyAll.Count))"
        $ws.Cells["V$($emptyRow-1)"].Style.Font.Bold = $true
        $gSrcL = ConvertTo-ExcelColumnLetter -Index $gSrcCol
        foreach ($strip in $gLive) {
            $ws.Cells["$gSrcL$($strip.Row)"].Value = $strip.Title
            $ws.Cells["$gSrcL$($strip.Row)"].Style.Font.Bold = $true
        }
        Close-ExcelPackage $xl
        Set-KpiChartColors -Path $reportPath -WorksheetName $sheet08
        Set-KpiSheetLayout -Path $reportPath -WorksheetName $sheet08
        # The group table sits in the fixed-width chart zone, so the layout pass
        # leaves its columns narrow. Widen each one to fit its own content
        # (header + values) so nothing has to be dragged out by hand.
        if ($emptyAll) {
            try {
                $xlG = Open-ExcelPackage -Path $reportPath
                $wsG = $xlG.Workbook.Worksheets[$sheet08]
                $props08 = @($emptyAll[0].PSObject.Properties.Name)
                for ($ci = 0; $ci -lt $props08.Count; $ci++) {
                    $prop = $props08[$ci]
                    $len = ($emptyAll | ForEach-Object { "$($_.$prop)".Length } | Measure-Object -Maximum).Maximum
                    $len = [math]::Max($len, "$prop".Length)
                    $wsG.Column($gDetailCol + $ci).Width = [math]::Min([math]::Max($len + 2, 10), 80)
                }
                Close-ExcelPackage $xlG
            } catch { }
        }
        Add-KpiDataSource -Path $reportPath -WorksheetName $sheet08 -SourceFolder $SourceFolder -BaseName 'EntraGroups', 'GroupMembers', 'GroupWelcomeEmail', 'EntraGroupsBasicInfo' -Row 2
        Write-Host "  08 Groups overview added." -ForegroundColor DarkGreen
    }

    # ============================================================
    # 09-12  CONDITIONAL ACCESS / APP REG / ENTERPRISE APPS / RBAC
    # Each builder is a separate function and returns if its dataset
    # was not collected, so a partial collection still produces a workbook.
    # ============================================================
    $sheetJobs = @(
        @{ Name = '09 Conditional Access'; Run = { Add-KpiSheetConditionalAccess -ReportPath $reportPath -Data $ca -SourceFolder $SourceFolder } ;    When = [bool]$ca }
        @{ Name = '10 App registrations';  Run = { Add-KpiSheetAppRegistrations  -ReportPath $reportPath -Data $appReg -SourceFolder $SourceFolder }; When = [bool]$appReg }
        @{ Name = '11 Enterprise apps';    Run = { Add-KpiSheetEnterpriseApps    -ReportPath $reportPath -Data $appReg -Groups $groups -SourceFolder $SourceFolder }; When = [bool]$appReg }
        @{ Name = '12 RBAC';               Run = { Add-KpiSheetRbac              -ReportPath $reportPath -Data $rbac -SourceFolder $SourceFolder };   When = [bool]$rbac }
        @{ Name = '13 Limits & recommendations'; Run = { Add-KpiSheetLimits -ReportPath $reportPath -SourceFolder $SourceFolder }; When = $true }
    )
    foreach ($job in $sheetJobs) {
        if (-not $job.When) { continue }
        try {
            $built = & $job.Run
            if ($built) { Write-Host "  $($job.Name) added." -ForegroundColor DarkGreen }
        }
        catch {
            # One bad sheet must not cost the whole workbook
            Write-Host "  $($job.Name) FAILED: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "Done. KPI workbook created:" -ForegroundColor Green
    Write-Host $reportPath -ForegroundColor Green
    return $reportPath
}
