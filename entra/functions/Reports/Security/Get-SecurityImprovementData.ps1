# Autodoc v1.0
function Get-SecurityImprovementData {
    <#
        Builds everything the "Security and improvements" report shows (Word,
        Excel and map), so the three files always agree.

        - Improvements: rules from Entra\Config\SecurityImprovements.json checked
          against the export. A rule is reported only when it triggers. Each row:
          Id, Area, AreaName, Severity, Title, Finding, Recommendation, Count,
          Reference, plus Objects (the affected objects, for the Excel detail).
        - Hunt: the workload identity hunt results when the export has them
          (JSON datasets WorkloadIdentityHunt, HuntAccountTriage, HuntCredentials).

        Everything is read from the export's JSON files; nothing is fetched from
        Graph here. Build-EntraSecurityReport saves this result as JSON and the
        Word, Excel and map files are built from that JSON file.

        Language: 'sv' or 'en' (texts come from the JSON, English as fallback).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [string]$Language = 'sv',
        [int]$InactiveDays = 90,
        [int]$GlobalAdminMax = 5
    )

    $lang = if ("$Language" -eq 'en') { 'en' } else { 'sv' }
    $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))   # Functions\Reports\Security -> Entra
    $cfg = Get-Content -LiteralPath (Join-Path $entraRoot 'Config\SecurityImprovements.json') -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable

    function Pick($entry) { if ($entry -and "$($entry[$lang])") { "$($entry[$lang])" } elseif ($entry) { "$($entry['en'])" } else { '' } }
    function Share([double]$part, [double]$whole) {
        if ($whole -le 0) { return '0 %' }
        $s = '{0:0.#}' -f (100 * $part / $whole)
        if ($lang -eq 'sv') { $s = $s -replace '\.', ',' }
        "$s %"
    }
    function IsYes($v) { $v -eq $true -or "$v" -match '^(Yes|True)$' }

    $privRoles = @(
        'Global Administrator','Privileged Role Administrator','Privileged Authentication Administrator',
        'Security Administrator','Conditional Access Administrator','Application Administrator',
        'Cloud Application Administrator','User Administrator','Exchange Administrator',
        'SharePoint Administrator','Intune Administrator','Hybrid Identity Administrator',
        'Domain Name Administrator','Authentication Administrator','Authentication Policy Administrator'
    )

    # ---- export data ----
    $tenant   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $users    = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    $groups   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    $apps     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    $eapps    = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps')
    $rbac     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC')
    $skus     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    $sspr     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'PasswordReset') | Select-Object -First 1
    $caRoot   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess') | Select-Object -First 1
    $policies = @()
    if ($caRoot) {
        $policies = if ($caRoot.PSObject.Properties.Name -contains 'ConditionalAccessPolicies') { @($caRoot.ConditionalAccessPolicies) }
                    else { @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess' | Where-Object { $_.RecordType -eq 'Policy' }) }
    }

    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Rule {
        param([string]$Id, [int]$Count, $Extra = '', $Objects = @(), [string]$Severity)
        $r = $cfg.Rules[$Id]
        if (-not $r) { return }
        $sev = if ($Severity) { $Severity } else { "$($r.Severity)" }
        $rows.Add([pscustomobject]@{
            Id             = $Id
            Area           = "$($r.Area)"
            AreaName       = (Pick $cfg.Areas["$($r.Area)"])
            Severity       = $sev
            Title          = (Pick $r.Title)
            Finding        = ((Pick $r.Finding) -f $Count, $Extra)
            Recommendation = (Pick $r.Recommendation)
            Count          = $Count
            Extra          = "$Extra"
            Reference      = "$($r.Reference)"
            Objects        = @($Objects)
        })
    }

    # ---- users ----
    if ($users.Count) {
        $enabledMembers = @($users | Where-Object { (IsYes $_.AccountEnabled) -and "$($_.UserType)" -ne 'Guest' })
        $pwdOnly = @($enabledMembers | Where-Object {
            $m = @("$($_.AuthenticationMethods)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'Password' })
            $m.Count -eq 0
        })
        if ($pwdOnly.Count) {
            $sev = if ($pwdOnly.Count / [math]::Max(1, $enabledMembers.Count) -ge 0.05) { 'HIGH' } else { 'MEDIUM' }
            Add-Rule 'MFA_PASSWORD_ONLY' $pwdOnly.Count (Share $pwdOnly.Count $enabledMembers.Count) `
                @($pwdOnly | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = $_.UserPrincipalName } }) -Severity $sev
        }
        $disLic = @($users | Where-Object { -not (IsYes $_.AccountEnabled) -and [int]"0$($_.AssignedLicenseCount)" -gt 0 })
        if ($disLic.Count) {
            Add-Rule 'USERS_DISABLED_LICENSED' $disLic.Count '' @($disLic | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = "$($_.UserPrincipalName) | $($_.AssignedLicenses)" } })
        }
        $guests = @($users | Where-Object { "$($_.UserType)" -eq 'Guest' })
        if ($guests.Count) {
            Add-Rule 'GUESTS_REVIEW' $guests.Count (Share $guests.Count $users.Count) @($guests | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = $_.UserPrincipalName } })
        }
    }

    # ---- privileged access ----
    $active   = @($rbac | Where-Object { $_.RecordType -eq 'RoleAssignmentActive' })
    $eligible = @($rbac | Where-Object { $_.RecordType -eq 'RoleAssignmentEligible' })
    if ($active.Count -or $eligible.Count) {
        $ga = @($active | Where-Object { "$($_.RoleDefinitionName)" -eq 'Global Administrator' })
        $gaObjects = @($ga | ForEach-Object { [pscustomobject]@{ Name = $_.PrincipalDisplayName; Detail = "$($_.PrincipalType)" } })
        if ($ga.Count -ge $GlobalAdminMax) { Add-Rule 'GA_TOO_MANY' $ga.Count $GlobalAdminMax $gaObjects }
        elseif ($ga.Count -lt 2)          { Add-Rule 'GA_TOO_FEW'  $ga.Count '' $gaObjects }

        $privActive = @($active | Where-Object { $privRoles -contains "$($_.RoleDefinitionName)" -and "$($_.PrincipalType)" -ne 'ServicePrincipal' })
        $privElig   = @($eligible | Where-Object { $privRoles -contains "$($_.RoleDefinitionName)" })
        if ($privActive.Count -gt $privElig.Count) {
            Add-Rule 'PIM_NOT_USED' $privActive.Count $privElig.Count @($privActive | ForEach-Object { [pscustomobject]@{ Name = $_.PrincipalDisplayName; Detail = $_.RoleDefinitionName } })
        }
        $spPriv = @($active | Where-Object { $privRoles -contains "$($_.RoleDefinitionName)" -and "$($_.PrincipalType)" -eq 'ServicePrincipal' })
        if ($spPriv.Count) {
            Add-Rule 'PRIV_SERVICE_PRINCIPAL' $spPriv.Count '' @($spPriv | ForEach-Object { [pscustomobject]@{ Name = $_.PrincipalDisplayName; Detail = $_.RoleDefinitionName } })
        }
    }

    # ---- conditional access ----
    if ($policies.Count) {
        $enabled = @($policies | Where-Object { "$($_.State)" -eq 'enabled' })
        $polObj  = { param($list) @($list | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = "$($_.State)" } }) }
        $mfaAll = @($enabled | Where-Object {
            "$($_.IncludeUsersRaw)" -match '(^|;)All($|;)' -and ("$($_.BuiltInControls)" -match 'mfa' -or "$($_.AuthenticationStrengthId)")
        })
        if (-not $mfaAll.Count) { Add-Rule 'CA_NO_MFA_ALL' 0 }
        $legacy = @($enabled | Where-Object { "$($_.BuiltInControls)" -match 'block' -and "$($_.ClientAppTypes)" -match 'exchangeActiveSync|other' })
        if (-not $legacy.Count) { Add-Rule 'CA_LEGACY_AUTH' 0 }
        $risk = @($enabled | Where-Object { "$($_.SignInRiskLevels)" -or "$($_.UserRiskLevels)" })
        if (-not $risk.Count) { Add-Rule 'CA_RISK_POLICIES' 0 }
        $ro = @($policies | Where-Object { "$($_.State)" -match 'report' })
        if ($ro.Count) { Add-Rule 'CA_REPORT_ONLY' $ro.Count '' (& $polObj $ro) }
        $off = @($policies | Where-Object { "$($_.State)" -eq 'disabled' })
        if ($off.Count) { Add-Rule 'CA_DISABLED' $off.Count '' (& $polObj $off) }
        $excl = @($enabled | Where-Object { $_.ExcludeUsersResolved -or $_.ExcludeGroupsResolved -or $_.ExcludeRolesResolved -or $_.ExcludeApplicationsResolved -or $_.ExcludeLocationsResolved })
        if ($excl.Count) { Add-Rule 'CA_EXCLUSIONS' $excl.Count '' (& $polObj $excl) }
    }

    # ---- app registrations ----
    if ($apps.Count) {
        $appObj = { param($list, $what) @($list | ForEach-Object { [pscustomobject]@{ Name = $_.AppName; Detail = (& $what $_) } }) }
        $high = @($apps | Where-Object { (Get-AppRiskFinding -App $_).Risk -eq 'High' })
        if ($high.Count) { Add-Rule 'APP_HIGH_RISK' $high.Count '' (& $appObj $high { param($a) (Get-AppRiskFinding -App $a).Reason }) }
        $noOwner = @($apps | Where-Object { "$($_.OwnerInfo)" -eq 'MissingOwners' -or -not $_.OwnerInfo })
        if ($noOwner.Count) { Add-Rule 'APP_NO_OWNER' $noOwner.Count '' (& $appObj $noOwner { param($a) "$($a.AppId)" }) }
        $expired = @($apps | Where-Object { "$($_.CredentialExpired)" -eq 'Yes' })
        if ($expired.Count) { Add-Rule 'APP_CRED_EXPIRED' $expired.Count '' (& $appObj $expired { param($a) "$($a.NearestCredentialExpiry)" }) }
        $soon = @($apps | Where-Object { "$($_.CredentialExpiresWithin30Days)" -eq 'Yes' })
        if ($soon.Count) { Add-Rule 'APP_CRED_EXPIRING' $soon.Count '' (& $appObj $soon { param($a) "$($a.NearestCredentialExpiry)" }) }
        $secrets = @($apps | Where-Object { "$($_.HasSecrets)" -eq 'Yes' })
        if ($secrets.Count) { Add-Rule 'APP_SECRETS' $secrets.Count '' (& $appObj $secrets { param($a) "$($a.SecretCount) secret(s)" }) }
    }

    # ---- enterprise apps ----
    if ($eapps.Count) {
        $own = @($eapps | Where-Object { "$($_.AppOrigin)" -eq 'This tenant' })
        $cut = (Get-Date).AddDays(-$InactiveDays)
        $inactive = @(foreach ($e in $own) {
            if ("$($e.HasSecrets)" -ne 'Yes' -and "$($e.HasCertificates)" -ne 'Yes') { continue }
            if ("$($e.CredentialExpired)" -eq 'Yes') { continue }
            $last = ConvertTo-SafeDateTime -Value $e.LastSignInDateTime
            if (-not $last -or $last -lt $cut) { $e }
        })
        if ($inactive.Count) {
            Add-Rule 'EAPP_INACTIVE' $inactive.Count $InactiveDays @($inactive | ForEach-Object { [pscustomobject]@{ Name = $_.AppName; Detail = "$($_.LastSignInDateTime)" } })
        }
        $open = @($own | Where-Object { (IsYes $_.SSOConfigured) -and -not (IsYes $_.AppRoleAssignmentRequired) })
        if ($open.Count) { Add-Rule 'EAPP_ASSIGNMENT_NOT_REQUIRED' $open.Count '' @($open | ForEach-Object { [pscustomobject]@{ Name = $_.AppName; Detail = "$($_.SSOType)" } }) }
    }

    # ---- password reset ----
    if ($sspr) {
        $scope = "$($sspr.SsprEnabledScope)"
        $isAll = $scope -eq 'All' -or ($scope -eq 'Selected' -and "$($sspr.SsprSelectedGroupId)" -eq 'all_users')
        if (-not $isAll) { Add-Rule 'SSPR_NOT_ALL' 0 $(if ($scope) { $scope } else { 'None' }) }
        $methods = @("$($sspr.MethodsAvailableToUsers)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($methods.Count -lt 2) { Add-Rule 'SSPR_FEW_METHODS' $methods.Count ($methods -join ', ') }
        if (IsYes $sspr.SecurityQuestionsEnabled) { Add-Rule 'SSPR_SECURITY_QUESTIONS' 1 }
    }

    # ---- groups ----
    if ($groups.Count) {
        $cloudNoOwner = @($groups | Where-Object { "$($_.Source)" -eq 'Cloud' -and -not "$($_.Owners)".Trim() })
        if ($cloudNoOwner.Count) { Add-Rule 'GROUPS_NO_OWNER' $cloudNoOwner.Count '' @($cloudNoOwner | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = "$($_.GroupCategory)" } }) }
        $empty = @($groups | Where-Object { "$($_.MemberCount)" -match '^\d+$' -and [int]$_.MemberCount -eq 0 -and -not $_.DynamicGroupMembershipRule })
        if ($empty.Count) { Add-Rule 'GROUPS_EMPTY' $empty.Count '' @($empty | ForEach-Object { [pscustomobject]@{ Name = $_.DisplayName; Detail = "$($_.Source)" } }) }
    }

    # ---- licences ----
    if ($skus.Count) {
        $purchased = @($skus | Where-Object { (Get-LicenseSkuCategory -SkuPartNumber $_.skuPartNumber -EnabledUnits ([int]$_.enabledUnits)) -eq 'Purchased' })
        $free = [int](($purchased | Measure-Object enabledUnits -Sum).Sum) - [int](($purchased | Measure-Object consumedUnits -Sum).Sum)
        if ($free -gt 0) {
            Add-Rule 'LICENSE_UNUSED' $free '' @($purchased | Where-Object { [int]$_.enabledUnits -gt [int]$_.consumedUnits } |
                ForEach-Object { [pscustomobject]@{ Name = $_.skuPartNumber; Detail = "$([int]$_.enabledUnits - [int]$_.consumedUnits) of $($_.enabledUnits) unused" } })
        }
    }

    # ---- workload identity hunt ----
    # Hunt results come from the JSON datasets the collector saved in rawDataJson
    # (CSV fallback through Get-AutodocExportData). Exports made before the hunt
    # saved JSON still work: then the hunt folder's own CSV files are read.
    $huntDir = Join-Path $SourceFolder 'WorkloadIdentityHunt'
    $hasSet = { param($base)
        foreach ($sub in 'rawDataJson', 'rawDataCsv') {
            if (Get-ChildItem -Path (Join-Path $SourceFolder $sub) -Filter "$base *" -ErrorAction SilentlyContinue | Select-Object -First 1) { return $true }
        }
        $false
    }
    $readSet = { param($base, $csvName)
        if (& $hasSet $base) { return @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName $base | Where-Object { $null -ne $_ }) }
        $p = Join-Path $huntDir $csvName
        if (Test-Path $p) { return @(Import-Csv -LiteralPath $p) }
        @()
    }
    $hunt = [ordered]@{
        Ran         = ((& $hasSet 'WorkloadIdentityHunt') -or (Test-Path (Join-Path $huntDir 'FINDINGS.csv')))
        Findings    = @(& $readSet 'WorkloadIdentityHunt' 'FINDINGS.csv')
        Triage      = @(& $readSet 'HuntAccountTriage' 'account_triage_summary.csv')
        Credentials = @(& $readSet 'HuntCredentials' 'credential_inventory.csv')
        Report      = @(Get-ChildItem -Path $huntDir -Filter '*.html' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1 | ForEach-Object { $_.FullName })
    }
    if (-not $hunt.Ran) {
        Add-Rule 'HUNT_NOT_RUN' 0
    } else {
        $top = @($hunt.Findings | Where-Object { $_.Severity -in 'CRITICAL', 'HIGH' })
        if ($top.Count) {
            $crit = @($top | Where-Object Severity -eq 'CRITICAL').Count
            $what = (@($top | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join '; ')
            Add-Rule 'HUNT_FINDINGS' $top.Count $what @($top | ForEach-Object { [pscustomobject]@{ Name = "$($_.Severity) - $($_.Category)"; Detail = "$($_.AppName) $($_.IP)".Trim() } }) -Severity $(if ($crit) { 'CRITICAL' } else { 'HIGH' })
        }
    }

    $sevOrder = @{ CRITICAL = 0; HIGH = 1; MEDIUM = 2; INFO = 3 }
    $sorted = @($rows | Sort-Object @{ Expression = { $sevOrder["$($_.Severity)"] } }, AreaName, Title)

    [pscustomobject]@{
        Language     = $lang
        Config       = $cfg
        Tenant       = $tenant
        Improvements = $sorted
        Hunt         = $hunt
        AreaOrder    = @($cfg.Areas.Keys)
        SevOrder     = @('CRITICAL', 'HIGH', 'MEDIUM', 'INFO')
    }
}
