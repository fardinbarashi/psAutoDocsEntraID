# Autodoc v1.0
function Export-TenantSummary {
    <#
        Section 1.12 - Builds the consolidated TenantSummary: a multi-sheet Excel
        workbook (Dashboard + DataSets index + one sheet per dataset), plus a CSV
        index and a structured JSON. Reads every dataset from the $Collected
        dictionary populated by the earlier collectors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [Parameter(Mandatory)]$Collected
    )

    $Section = 'Section 1.12 : Consolidated TenantSummary'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        $fileDate = $Paths.FileDate

        # Pull datasets from the shared collection (missing = $null)
        $tenantData        = $Collected['TenantData']
        $domainData        = $Collected['DomainData']
        $licenseInventory  = $Collected['LicenseInventory']
        $userData          = $Collected['UserData']
        $groupBasicInfo    = $Collected['GroupBasicInfo']
        $entraGroups       = $Collected['EntraGroups']
        $groupMembers      = $Collected['GroupMembers']
        $groupWelcomeEmail = $Collected['GroupWelcomeEmail']
        $appRegEnterprise  = $Collected['AppRegEnterprise']
        $conditionalAccess = $Collected['ConditionalAccessCsv']
        $rbac              = $Collected['Rbac']
        $passwordReset     = $Collected['PasswordReset']

        $summaryDataSets = @(
            [pscustomobject]@{ Name = "TenantInformation";       WorksheetName = "TenantInformation"; Data = $tenantData
                CsvPath = (Join-Path $Paths.Csv "TenantInformation $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "TenantInformation $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "TenantInformation $fileDate.xlsx")
                Description = "Tenant metadata and high-level object counts." }
            [pscustomobject]@{ Name = "TenantInformationDomain"; WorksheetName = "TenantDomains"; Data = $domainData
                CsvPath = (Join-Path $Paths.Csv "TenantInformationDomain $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "TenantInformationDomain $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "TenantInformationDomain $fileDate.xlsx")
                Description = "Verified tenant domains." }
            [pscustomobject]@{ Name = "LicensesInformation";     WorksheetName = "LicensesInformation"; Data = $licenseInventory
                CsvPath = (Join-Path $Paths.Csv "LicensesInformation $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "LicensesInformation $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "LicensesInformation $fileDate.xlsx")
                Description = "Subscribed SKU inventory, consumed units and free units." }
            [pscustomobject]@{ Name = "UserInformation";         WorksheetName = "UserInformation"; Data = $userData
                CsvPath = (Join-Path $Paths.Csv "UserInformation $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "UserInformation $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "UserInformation $fileDate.xlsx")
                Description = "Users, licensing, account status, manager and authentication method information." }
            [pscustomobject]@{ Name = "EntraGroupsBasicInfo";    WorksheetName = "GroupsBasicInfo"; Data = $groupBasicInfo
                CsvPath = (Join-Path $Paths.Csv "EntraGroupsBasicInfo $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "EntraGroupsBasicInfo $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "EntraGroupsBasicInfo $fileDate.xlsx")
                Description = "Group count summary by group category and source." }
            [pscustomobject]@{ Name = "EntraGroups";             WorksheetName = "EntraGroups"; Data = $entraGroups
                CsvPath = (Join-Path $Paths.Csv "EntraGroups $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "EntraGroups $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "EntraGroups $fileDate.xlsx")
                Description = "Detailed group inventory, owners, roles, applications, licenses and welcome-email status." }
            [pscustomobject]@{ Name = "GroupMembers";            WorksheetName = "GroupMembers"; Data = $groupMembers
                CsvPath = (Join-Path $Paths.Csv "GroupMembers $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "GroupMembers $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "GroupMembers $fileDate.xlsx")
                Description = "Expanded group membership with selected user attributes." }
            [pscustomobject]@{ Name = "GroupWelcomeEmail";       WorksheetName = "GroupWelcomeEmail"; Data = $groupWelcomeEmail
                CsvPath = (Join-Path $Paths.Csv "GroupWelcomeEmail $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "GroupWelcomeEmail $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "GroupWelcomeEmail $fileDate.xlsx")
                Description = "Welcome-email status for Microsoft 365 groups." }
            [pscustomobject]@{ Name = "AppRegEnterpriseApps";    WorksheetName = "AppRegEnterpriseApps"; Data = $appRegEnterprise
                CsvPath = (Join-Path $Paths.Csv "AppRegEnterpriseApps $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "AppRegEnterpriseApps $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "AppRegEnterpriseApps $fileDate.xlsx")
                Description = "Application registrations and enterprise app information, owners, SSO, credentials and sign-in activity." }
            [pscustomobject]@{ Name = "ConditionalAccess";       WorksheetName = "ConditionalAccess"; Data = $conditionalAccess
                CsvPath = (Join-Path $Paths.Csv "ConditionalAccess $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "ConditionalAccess $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "ConditionalAccess $fileDate.xlsx")
                Description = "Conditional Access policies, named locations and authentication strength policies." }
            [pscustomobject]@{ Name = "RBAC";                    WorksheetName = "RBAC"; Data = $rbac
                CsvPath = (Join-Path $Paths.Csv "RBAC $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "RBAC $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "RBAC $fileDate.xlsx")
                Description = "Active and eligible role assignments plus role-assignable groups." }
            [pscustomobject]@{ Name = "PasswordReset";           WorksheetName = "PasswordReset"; Data = $passwordReset
                CsvPath = (Join-Path $Paths.Csv "PasswordReset $fileDate.csv"); JsonPath = (Join-Path $Paths.Json "PasswordReset $fileDate.json"); ExcelPath = (Join-Path $Paths.Excel "PasswordReset $fileDate.xlsx")
                Description = "Self-service password reset configuration." }
        )

        $summaryDashboard = @(
            [pscustomobject]@{ Category = "Tenant";             Metric = "Tenant name";                         Value = $tenantData.TenantName }
            [pscustomobject]@{ Category = "Tenant";             Metric = "Tenant ID";                           Value = $tenantData.TenantId }
            [pscustomobject]@{ Category = "Tenant";             Metric = "Primary domain";                      Value = $tenantData.PrimaryDomain }
            [pscustomobject]@{ Category = "Tenant";             Metric = "Hybrid status";                       Value = $tenantData.HybridStatus }
            [pscustomobject]@{ Category = "Users";              Metric = "Total users";                         Value = (Get-DataRowCount -Data $userData) }
            [pscustomobject]@{ Category = "Users";              Metric = "Enabled users";                       Value = @($userData | Where-Object { $_.AccountEnabled -eq $true }).Count }
            [pscustomobject]@{ Category = "Users";              Metric = "Disabled users";                      Value = @($userData | Where-Object { $_.AccountEnabled -eq $false }).Count }
            [pscustomobject]@{ Category = "Users";              Metric = "Guest users";                         Value = @($userData | Where-Object { $_.UserType -eq "Guest" }).Count }
            [pscustomobject]@{ Category = "Groups";             Metric = "Total groups";                        Value = (Get-DataRowCount -Data $entraGroups) }
            [pscustomobject]@{ Category = "Groups";             Metric = "Dynamic groups";                      Value = @($entraGroups | Where-Object { $_.GroupCategory -eq "Dynamic" }).Count }
            [pscustomobject]@{ Category = "Groups";             Metric = "M365 groups";                         Value = @($entraGroups | Where-Object { $_.GroupCategory -eq "M365" }).Count }
            [pscustomobject]@{ Category = "Groups";             Metric = "Security groups";                     Value = @($entraGroups | Where-Object { $_.GroupCategory -eq "Security" }).Count }
            [pscustomobject]@{ Category = "Licenses";           Metric = "Enabled license units";               Value = (Get-MeasureSum -Data $licenseInventory -PropertyName "enabledUnits") }
            [pscustomobject]@{ Category = "Licenses";           Metric = "Consumed license units";              Value = (Get-MeasureSum -Data $licenseInventory -PropertyName "consumedUnits") }
            [pscustomobject]@{ Category = "Licenses";           Metric = "Free license units";                  Value = (Get-MeasureSum -Data $licenseInventory -PropertyName "freeUnits") }
            [pscustomobject]@{ Category = "Applications";       Metric = "Applications";                        Value = (Get-DataRowCount -Data $appRegEnterprise) }
            [pscustomobject]@{ Category = "Applications";       Metric = "Applications without owners";         Value = @($appRegEnterprise | Where-Object { $_.OwnerInfo -eq "MissingOwners" }).Count }
            [pscustomobject]@{ Category = "Applications";       Metric = "Expired credentials";                 Value = @($appRegEnterprise | Where-Object { $_.CredentialExpired -eq "Yes" }).Count }
            [pscustomobject]@{ Category = "Applications";       Metric = "Credentials expiring within 30 days"; Value = @($appRegEnterprise | Where-Object { $_.CredentialExpiresWithin30Days -eq "Yes" }).Count }
            [pscustomobject]@{ Category = "Conditional Access"; Metric = "Conditional Access records";          Value = (Get-DataRowCount -Data $conditionalAccess) }
            [pscustomobject]@{ Category = "RBAC";               Metric = "RBAC records";                        Value = (Get-DataRowCount -Data $rbac) }
            [pscustomobject]@{ Category = "Password Reset";     Metric = "SSPR enabled scope";                  Value = $passwordReset.SsprEnabledScope }
            [pscustomobject]@{ Category = "Export";             Metric = "Export generated at";                 Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
            [pscustomobject]@{ Category = "Export";             Metric = "Export folder";                       Value = $Paths.Export }
        )

        $summaryIndexData = foreach ($dataSet in $summaryDataSets) {
            $rowCount = Get-DataRowCount -Data $dataSet.Data
            [pscustomobject]@{
                DataSet       = $dataSet.Name
                WorksheetName = (Get-SafeExcelWorksheetName -Name $dataSet.WorksheetName)
                RowCount      = $rowCount
                Status        = if ($rowCount -gt 0) { "OK" } else { "No data or section failed" }
                CsvPath       = $dataSet.CsvPath
                JsonPath      = $dataSet.JsonPath
                ExcelPath     = $dataSet.ExcelPath
                Description   = $dataSet.Description
            }
        }

        if (Test-Path -Path $Paths.SummaryExcel) { Remove-Item -Path $Paths.SummaryExcel -Force }

        Write-Host "Exporting consolidated TenantSummary workbook..." -ForegroundColor Yellow

        $summaryDashboard | Export-Excel -Path $Paths.SummaryExcel -WorksheetName "Dashboard" -AutoSize -TableName "tbl_Dashboard" -BoldTopRow -FreezeTopRow
        $summaryIndexData | Export-Excel -Path $Paths.SummaryExcel -WorksheetName "DataSets"  -AutoSize -TableName "tbl_DataSets"  -BoldTopRow -FreezeTopRow

        foreach ($dataSet in $summaryDataSets) {
            $worksheetName = Get-SafeExcelWorksheetName -Name $dataSet.WorksheetName
            $tableName     = Get-SafeExcelTableName     -Name $dataSet.WorksheetName
            $rowCount      = Get-DataRowCount           -Data $dataSet.Data

            if ($rowCount -gt 0) {
                $exportRows = @($dataSet.Data)
            }
            else {
                $exportRows = [pscustomobject]@{
                    DataSet     = $dataSet.Name
                    Info        = "No data was collected, or the section failed before this variable was populated."
                    GeneratedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
            }

            Write-Host "-> Adding worksheet: $worksheetName ($rowCount rows)" -ForegroundColor Yellow
            $exportRows | Export-Excel -Path $Paths.SummaryExcel -WorksheetName $worksheetName -AutoSize -TableName $tableName -BoldTopRow -FreezeTopRow
        }

        $summaryIndexData | Export-Csv -Path $Paths.SummaryCsv -NoTypeInformation -Encoding UTF8 -Force

        [pscustomobject]@{
            GeneratedAt      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            TenantName       = $tenantData.TenantName
            TenantId         = $tenantData.TenantId
            ExportFolder     = $Paths.Export
            SummaryExcelPath = $Paths.SummaryExcel
            DataSetSummary   = $summaryIndexData
            DashboardSummary = $summaryDashboard
        } | ConvertTo-Json -Depth 20 | Out-File -Path $Paths.SummaryJson -Encoding UTF8 -Force

        Write-Host "Done. Consolidated TenantSummary workbook created:" -ForegroundColor DarkGreen
        Write-Host $Paths.SummaryExcel -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
