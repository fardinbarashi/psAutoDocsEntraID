# Autodoc v1.0
function Get-WordTemplateValues {
    <#
        Builds the flat {placeholder} -> value table used to fill the Word
        template text files. Every number comes from the exported JSON (via the
        shared map-data plus a little extra counting); the template files
        supply all the wording around them, so the prose can be edited freely
        without touching code.

        Returns an ordered hashtable of placeholder name -> string value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $v = [ordered]@{}
    $set = { param($k, $val) $v[$k] = if ($null -eq $val) { '' } else { "$val" } }

    # ---- gather each area (all optional) ----
    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $usr  = Get-UsersMapData             -SourceFolder $SourceFolder
    $grpB = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroupsBasicInfo') | Select-Object -First 1
    $grp  = Get-GroupsMapData            -SourceFolder $SourceFolder
    $lic  = Get-LicenceMapData           -SourceFolder $SourceFolder
    $ca   = Get-ConditionalAccessMapData -SourceFolder $SourceFolder
    $areg = Get-AppRegMapData            -SourceFolder $SourceFolder
    $appRaw = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    $eapp = Get-EnterpriseAppMapData     -SourceFolder $SourceFolder
    $rbac = Get-RbacMapData              -SourceFolder $SourceFolder
    $rbacRaw = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC' | Where-Object { $_.RecordType -like 'RoleAssignment*' })
    $dept = Get-DepartmentMapData        -SourceFolder $SourceFolder
    $dom  = Get-DomainMapData            -SourceFolder $SourceFolder
    $sspr = Get-PasswordResetData        -SourceFolder $SourceFolder

    # ---- tenant / intro ----
    & $set 'TenantName'            $(if ($tenant) { $tenant.TenantName } else { 'the tenant' })
    & $set 'PrimaryDomain'         $(if ($tenant) { $tenant.PrimaryDomain } else { '' })
    & $set 'Environment'           ("$(Get-TenantHybridLabel -Tenant $tenant)" -replace '^Environment:\s*', '')
    & $set 'GeneratedDate'         (Get-Date -Format 'yyyy-MM-dd')

    # ---- users ----
    if ($usr) {
        & $set 'TotalUsers'          $usr.Total
        $mem = 0; $gue = 0
        foreach ($t in $usr.Types) { if ("$($t.Name)" -match 'Guest') { $gue = $t.Count } elseif ("$($t.Name)" -match 'Member') { $mem = $t.Count } }
        & $set 'Members'             $mem
        & $set 'Guests'              $gue
        & $set 'EnabledUsers'        $usr.Enabled
        & $set 'DisabledUsers'       $usr.Disabled
        & $set 'UsersWithoutMailbox' $usr.NoMail
        & $set 'UsersMultiAuth'      $usr.MultiAuth
    }

    # ---- groups (authoritative counts from BasicInfo) ----
    if ($grpB) {
        & $set 'TotalGroups'      $grpB.TotalGroups
        & $set 'SecurityGroups'   $grpB.SecurityGroups
        & $set 'M365Groups'       $grpB.M365Groups
        & $set 'DynamicGroups'    $grpB.DynamicGroups
        & $set 'CloudGroups'      $grpB.CloudGroups
        & $set 'OnPremisesGroups' $grpB.OnPremisesGroups
    }
    elseif ($grp) { & $set 'TotalGroups' $grp.Total }

    # ---- licences ----
    if ($lic) {
        $pc = @($lic.Purchased).Count; $vc = @($lic.Viral).Count
        & $set 'PurchasedSkus'   $pc
        & $set 'ViralSkus'       $vc
        & $set 'TotalLicenceSkus' ($pc + $vc)
    }

    # ---- conditional access ----
    if ($ca) {
        & $set 'TotalPolicies' $ca.Total
        $st = @($ca.Policies | ForEach-Object { "$($_.State)".ToLower() })
        & $set 'EnabledPolicies'    (@($st | Where-Object { $_ -eq 'enabled' }).Count)
        & $set 'ReportOnlyPolicies' (@($st | Where-Object { $_ -match 'report' }).Count)
        & $set 'DisabledPolicies'   (@($st | Where-Object { $_ -eq 'disabled' }).Count)
    }

    # ---- app registrations ----
    if ($areg) {
        & $set 'TotalAppRegistrations' $areg.Total
        $single = @($appRaw | Where-Object { "$($_.SignInAudience)" -eq 'AzureADMyOrg' }).Count
        & $set 'SingleTenantApps' $single
        & $set 'MultiTenantApps'  ($areg.Total - $single)
        & $set 'AppsWithoutOwner' (@($appRaw | Where-Object { "$($_.OwnerInfo)" -eq 'MissingOwners' -or -not $_.OwnerInfo }).Count)
        & $set 'AppsExpiringSoon' (@($appRaw | Where-Object { "$($_.CredentialExpiresWithin30Days)" -eq 'Yes' }).Count)
        & $set 'AppsWithRisk'     (@($appRaw | Where-Object { (Get-AppRiskFinding -App $_).Risk }).Count)
    }

    # ---- enterprise apps ----
    if ($eapp) {
        & $set 'TotalEnterpriseApps' $eapp.Total
        & $set 'SsoConfigured'       $eapp.SsoConfigured
        & $set 'WithoutSso'          ($eapp.Total - $eapp.SsoConfigured)
    }

    # ---- RBAC ----
    if ($rbac) {
        & $set 'DistinctRoles'      (@($rbac.Roles).Count)
        & $set 'TotalAssignments'   $rbac.Total
        & $set 'ActiveAssignments'  $rbac.ActiveN
        & $set 'EligibleAssignments' $rbac.EligibleN
        $privRoles = @('Global Administrator', 'Privileged Role Administrator', 'Privileged Authentication Administrator',
            'Security Administrator', 'Conditional Access Administrator', 'Application Administrator',
            'Cloud Application Administrator', 'User Administrator', 'Exchange Administrator',
            'SharePoint Administrator', 'Intune Administrator', 'Hybrid Identity Administrator',
            'Domain Name Administrator', 'Directory Synchronization Accounts', 'Partner Tier2 Support')
        & $set 'PrivilegedAssignments' (@($rbacRaw | Where-Object { $privRoles -contains "$($_.RoleDefinitionName)" }).Count)
        $multi = @($rbacRaw | Group-Object PrincipalId | Where-Object { @($_.Group | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count -gt 1 })
        & $set 'PrincipalsWithMultipleRoles' $multi.Count
    }

    # ---- departments / offices ----
    if ($dept) {
        & $set 'DepartmentCount' (@($dept.Departments).Count)
        & $set 'OfficeCount'     (@($dept.Offices).Count)
    }

    # ---- domains ----
    if ($dom) {
        & $set 'TotalDomains'  $dom.Total
        & $set 'DefaultDomain' $dom.Default
    }

    # ---- password reset (SSPR) ----
    if ($sspr -and $sspr.Config) {
        & $set 'SsprScope' $sspr.Config.SsprEnabledScope
    }

    $v
}
