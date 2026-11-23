# Autodoc v1.0
function Write-ExportReadme {
    <#
        Writes a README.md at the root of an export folder describing what every
        sub-folder holds and, for each JSON file, its purpose and the fields
        (parameters) each record carries. Field meanings come from the shared
        map glossary; fields not listed there are shown by name only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportFolder
    )
    if (-not (Test-Path $ExportFolder)) { return }

    $jsonDir  = Join-Path $ExportFolder 'rawDataJson'
    $glossary = Get-MapFieldGlossary

    # one-line purpose per JSON area (base name without the trailing date)
    $purpose = @{
        'TenantInformation'       = 'Tenant identity: name, tenant/organisation IDs, primary domain and directory-sync status.'
        'TenantInformationDomain' = 'Verified domains in the tenant, with each domain''s type and whether it is the default.'
        'UserInformation'         = 'Every user account: type (member/guest), status, department, office, licences and authentication methods.'
        'LicensesInformation'     = 'Licence SKUs (subscriptions): enabled / consumed / free units and the resulting utilisation.'
        'EntraGroups'             = 'All groups: type, category, owners, member count, dynamic membership rule and created date.'
        'EntraGroupsBasicInfo'    = 'Aggregate group counts for the tenant (totals per type and per source).'
        'GroupMembers'            = 'One row per group membership - the member together with their department, office location and manager.'
        'GroupWelcomeEmail'       = 'Per-group welcome-email setting (whether a welcome mail is sent to new members).'
        'ConditionalAccess'       = 'Conditional Access policies: state, grant controls (MFA, ...), and targeted / excluded users, groups and roles.'
        'AppRegEnterpriseApps'    = 'App registrations joined to their enterprise app (service principal): ownership, credentials, API permissions and sign-in recency.'
        'EnterpriseApps'          = 'Enterprise applications (service principals): single sign-on, Microsoft Graph usage and client type.'
        'RBAC'                    = 'Directory role assignments (active and eligible) with the principal that holds each role.'
        'PasswordReset'           = 'Self-service password reset (SSPR) configuration for the tenant.'
        'WorkloadIdentityHunt'    = 'Workload identity hunt findings, sorted by severity: IoC sign-ins, account triage verdicts, secret failures, token sweeps, credentials and posture.'
        'HuntAccountTriage'       = 'Workload identity hunt: one row per account with sign-ins from IoC addresses - verdict, error codes and action.'
        'HuntCredentials'         = 'Workload identity hunt: every client secret, certificate and federated credential on apps and service principals.'
        'HuntAuditChanges'        = 'Workload identity hunt: audit log changes to app credentials, owners and consents.'
        'HuntIocSpSignIns'        = 'Workload identity hunt: service principal sign-ins from the IoC addresses.'
        'HuntIocUserSignIns'      = 'Workload identity hunt: user and managed identity sign-ins from the IoC addresses.'
        'HuntSecretFailures'      = 'Workload identity hunt: invalid / expired secret errors (AADSTS7000215 / 7000222) from unknown addresses.'
        'HuntSpSignIns'           = 'Workload identity hunt raw data (only with SaveRawSignIns): every service principal sign-in in the window.'
        'HuntTriageSignIns'       = 'Workload identity hunt raw data (only with SaveRawSignIns): every sign-in of the triaged accounts.'
    }

    # short description of each output sub-folder
    $folders = @(
        @{ N = 'rawDataJson';          D = 'Raw data pulled from Microsoft Graph, one file per area (each file is documented below).' }
        @{ N = 'rawDataCsv';           D = 'The same data as CSV - one file per area, opens directly in Excel or a text editor.' }
        @{ N = 'rawDataExcel';         D = 'Per-area Excel workbooks (one sheet of raw data each).' }
        @{ N = 'WorkloadIdentityHunt'; D = 'Workload identity hunt (only after a hunt run on the Tools tab): one CSV per data set and a self-contained HTML report with search, filters and CSV export.' }
        @{ N = 'rawDataTenantSummary'; D = 'A short, high-level summary of the tenant.' }
        @{ N = 'Report\excelkpi';      D = 'The KPI workbook "KPI <stamp>.xlsx": dashboard, charts and one analysed sheet per area.' }
        @{ N = 'Report\svg';           D = 'Vector maps of the tenant (open in a web browser). That folder has its own README.' }
        @{ N = 'Report\pdf';           D = 'PDF copies of every SVG map, gathered in one place for sharing or printing.' }
        @{ N = 'Report\security';     D = 'Security and improvements: a JSON file with the proposals, and the Word document, Excel workbook and map (SVG + PDF) built from it, all with the same name.' }
        @{ N = 'Report\word';          D = 'The Word system-documentation report (.docx).' }
    )

    $sb = [System.Text.StringBuilder]::new()
    $add = { param($t) [void]$sb.AppendLine($t) }

    & $add "# Autodoc export - $(Split-Path $ExportFolder -Leaf)"
    & $add ''
    & $add "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    try {
        $t = @(Get-AutodocExportData -ExportFolder $ExportFolder -BaseName 'TenantInformation') | Select-Object -First 1
        if ($t) { & $add "Tenant: $($t.TenantName)  ($($t.PrimaryDomain))" }
    } catch { }
    & $add ''
    & $add 'This folder is a single export - one point-in-time pull of the tenant. Everything here (CSV, Excel, maps) is built from the raw JSON in the `json\` sub-folder, so the JSON is the source of truth.'
    & $add ''

    & $add '## Folders'
    & $add ''
    foreach ($f in $folders) {
        if (Test-Path (Join-Path $ExportFolder $f.N)) { & $add "- **$($f.N)\** - $($f.D)" }
    }
    & $add ''

    & $add '## JSON files (json\)'
    & $add ''
    & $add 'Each file is an export of one area. The fields listed under each file are the columns every record carries; a dash means the field name is self-explanatory.'
    & $add ''

    if (Test-Path $jsonDir) {
        $files = @(Get-ChildItem -Path $jsonDir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name)
        $seen = @{}
        foreach ($file in $files) {
            $base = ($file.BaseName -replace '\s*\d{4}-\d{2}-\d{2}.*$', '').Trim()
            if (-not $base -or $seen.ContainsKey($base)) { continue }
            $seen[$base] = $true

            & $add "### $base"
            if ($purpose.ContainsKey($base)) { & $add $purpose[$base] }
            & $add ''

            try {
                $rows = @(Get-AutodocExportData -ExportFolder $ExportFolder -BaseName $base)
                if ($rows.Count) {
                    $first = $rows[0]
                    # some exports wrap the real records inside an array-valued
                    # property (e.g. Conditional Access) - descend into the
                    # largest such array so we list the record fields, not the wrapper
                    $arrProps = @($first.PSObject.Properties | Where-Object {
                            $_.Value -is [System.Collections.IEnumerable] -and $_.Value -isnot [string] -and
                            @($_.Value).Count -gt 0 -and @($_.Value)[0].PSObject
                        })
                    if ($rows.Count -eq 1 -and $arrProps.Count) {
                        $biggest = $arrProps | Sort-Object { @($_.Value).Count } -Descending | Select-Object -First 1
                        $first = @($biggest.Value)[0]
                    }
                    foreach ($fld in $first.PSObject.Properties.Name) {
                        $mean = if ($glossary.ContainsKey($fld)) { $glossary[$fld] } else { '-' }
                        & $add "- ``$fld`` : $mean"
                    }
                }
                else {
                    & $add '(no records in this export)'
                }
            }
            catch { & $add '(could not read fields)' }
            & $add ''
        }
    }

    $path = Join-Path $ExportFolder 'README.md'
    $sb.ToString() | Out-File -FilePath $path -Encoding utf8
    Write-Host "  README written: $path" -ForegroundColor DarkGreen
    return $path
}
