# Autodoc v1.0
function Get-KpiCardData {
    <#
        Returns the KPI cards shown at the top of the Excel KPI sheets, computed
        with the same rules as the sheets themselves, so other reports (Word)
        show exactly the same figures.

        Output: ordered dictionary  Area -> @( @{ Label; Value }, ... )
          Dashboard, Groups, ConditionalAccess, AppRegistrations,
          EnterpriseApps, Rbac
        An area is left out when its dataset is missing from the export.

        Keep in step with:
          Build-EntraExcelReport (Dashboard cards, 08 Groups overview cards)
          Add-KpiSheetConditionalAccess, Add-KpiSheetAppRegistrations,
          Add-KpiSheetEnterpriseApps, Add-KpiSheetRbac
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $skus   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    $users  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    $groups = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    $ca     = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess')
    $appReg = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    $rbac   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC')

    $privRoleList = @(
        'Global Administrator','Privileged Role Administrator','Privileged Authentication Administrator',
        'Security Administrator','Conditional Access Administrator','Application Administrator',
        'Cloud Application Administrator','User Administrator','Exchange Administrator',
        'SharePoint Administrator','Intune Administrator','Hybrid Identity Administrator',
        'Domain Name Administrator','Directory Synchronization Accounts','Partner Tier2 Support'
    )

    # ---- licences: purchased SKUs and per-user assignments ----
    $purchased = @(foreach ($s in $skus) {
        $enabled = [int]$s.enabledUnits
        if ((Get-LicenseSkuCategory -SkuPartNumber $s.skuPartNumber -EnabledUnits $enabled) -eq 'Purchased') {
            [pscustomobject]@{ Enabled = $enabled; Consumed = [int]$s.consumedUnits }
        }
    })
    $totalAssigned = 0
    foreach ($u in $users) {
        if ($u.AssignedLicenses) { $totalAssigned += @($u.AssignedLicenses -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }).Count }
    }
    $purFree = [int](($purchased | Measure-Object Enabled -Sum).Sum) - [int](($purchased | Measure-Object Consumed -Sum).Sum)

    # ---- groups: same classification as the KPI workbook ----
    $gInfo = @(foreach ($grp in $groups) {
        $unified = [string]$grp.GroupTypes -match 'Unified'
        $secure  = "$($grp.SecurityEnabled)" -eq 'True'
        $mail    = "$($grp.MailEnabled)"     -eq 'True'
        $members = 0; [void][int]::TryParse("$($grp.MemberCount)", [ref]$members)
        $owners  = @()
        if ($grp.Owners) { $owners = @("$($grp.Owners)" -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        [pscustomobject]@{
            Kind      = $(if ($unified) { 'Microsoft 365' } elseif ($secure -and $mail) { 'Mail-enabled security' } elseif ($secure) { 'Security' } elseif ($mail) { 'Distribution' } else { 'Other' })
            Members   = $members
            OwnerCount = $owners.Count
            IsDynamic = [bool]$grp.DynamicGroupMembershipRule
        }
    })

    # ---- RBAC ----
    $roleAssign = @($rbac | Where-Object { $_.RecordType -like 'RoleAssignment*' })

    $out = [ordered]@{}

    $out['Dashboard'] = @(
        @{ Label = 'Users';                       Value = $(if ($tenant) { $tenant.Users }   else { $users.Count }) }
        @{ Label = 'Devices';                     Value = $(if ($tenant) { $tenant.Devices } else { 'n/a' }) }
        @{ Label = 'Assigned licences';           Value = $totalAssigned }
        @{ Label = 'Purchased SKUs';              Value = $purchased.Count }
        @{ Label = 'Unused units';                Value = $purFree }
        @{ Label = 'Total groups';                Value = $gInfo.Count }
        @{ Label = 'Security groups';             Value = @($gInfo | Where-Object Kind -eq 'Security').Count }
        @{ Label = 'Microsoft 365 groups';        Value = @($gInfo | Where-Object Kind -eq 'Microsoft 365').Count }
        @{ Label = 'Total app registrations';     Value = @($appReg).Count }
        @{ Label = 'Apps with service principal'; Value = @($appReg | Where-Object { "$($_.ServicePrincipalExists)" -eq 'Yes' }).Count }
        @{ Label = 'Distinct roles';              Value = @($roleAssign | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count }
        @{ Label = 'Role assignments';            Value = $roleAssign.Count }
        @{ Label = 'Privileged assignments';      Value = @($roleAssign | Where-Object { $privRoleList -contains "$($_.RoleDefinitionName)" }).Count }
    )

    if ($gInfo.Count) {
        $gBasic = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroupsBasicInfo') | Select-Object -First 1
        $out['Groups'] = @(
            @{ Label = 'Total groups';      Value = $(if ($gBasic) { $gBasic.TotalGroups } else { $gInfo.Count }) }
            @{ Label = 'Security';          Value = $(if ($gBasic) { $gBasic.SecurityGroups } else { @($gInfo | Where-Object Kind -eq 'Security').Count }) }
            @{ Label = 'Microsoft 365';     Value = $(if ($gBasic) { $gBasic.M365Groups } else { @($gInfo | Where-Object Kind -eq 'Microsoft 365').Count }) }
            @{ Label = 'Distribution';      Value = @($gInfo | Where-Object Kind -eq 'Distribution').Count }
            @{ Label = 'Dynamic';           Value = $(if ($gBasic) { $gBasic.DynamicGroups } else { @($gInfo | Where-Object IsDynamic).Count }) }
            @{ Label = 'Cloud';             Value = $(if ($gBasic) { $gBasic.CloudGroups } else { '' }) }
            @{ Label = 'On-premises';       Value = $(if ($gBasic) { $gBasic.OnPremisesGroups } else { '' }) }
            @{ Label = 'Without owner';     Value = @($gInfo | Where-Object { $_.OwnerCount -eq 0 }).Count }
            @{ Label = 'Empty (0 members)'; Value = @($gInfo | Where-Object { $_.Members -eq 0 }).Count }
        )
    }

    if ($ca.Count) {
        $policies = if (@($ca)[0].PSObject.Properties.Name -contains 'ConditionalAccessPolicies') { @(@($ca)[0].ConditionalAccessPolicies) } else { @($ca | Where-Object { $_.RecordType -eq 'Policy' }) }
        if ($policies.Count) {
            $out['ConditionalAccess'] = @(
                @{ Label = 'Total policies';  Value = $policies.Count }
                @{ Label = 'Enabled';         Value = @($policies | Where-Object { "$($_.State)" -eq 'enabled' }).Count }
                @{ Label = 'Report-only';     Value = @($policies | Where-Object { "$($_.State)" -match 'report' }).Count }
                @{ Label = 'Disabled';        Value = @($policies | Where-Object { "$($_.State)" -eq 'disabled' }).Count }
                @{ Label = 'With exclusions'; Value = @($policies | Where-Object {
                            $_.ExcludeUsersResolved -or $_.ExcludeGroupsResolved -or $_.ExcludeRolesResolved -or
                            $_.ExcludeApplicationsResolved -or $_.ExcludeLocationsResolved }).Count }
            )
        }
    }

    if ($appReg.Count) {
        $single = @($appReg | Where-Object { "$($_.SignInAudience)" -eq 'AzureADMyOrg' }).Count
        $out['AppRegistrations'] = @(
            @{ Label = 'Total app registrations'; Value = $appReg.Count }
            @{ Label = 'Single tenant';           Value = $single }
            @{ Label = 'Multi tenant';            Value = ($appReg.Count - $single) }
            @{ Label = 'Without owner';           Value = @($appReg | Where-Object { "$($_.OwnerInfo)" -eq 'MissingOwners' -or -not $_.OwnerInfo }).Count }
            @{ Label = 'Expires < 30 days';       Value = @($appReg | Where-Object { "$($_.CredentialExpiresWithin30Days)" -eq 'Yes' }).Count }
            @{ Label = 'Apps with risk';          Value = @($appReg | Where-Object { (Get-AppRiskFinding -App $_).Risk }).Count }
        )

        $sp = @($appReg | Where-Object { "$($_.ServicePrincipalExists)" -eq 'Yes' })
        if ($sp.Count) {
            $sso = @($sp | Where-Object { "$($_.SSOConfigured)" -eq 'Yes' }).Count
            $out['EnterpriseApps'] = @(
                @{ Label = 'Apps with service principal'; Value = $sp.Count }
                @{ Label = 'SSO configured';              Value = $sso }
                @{ Label = 'Without SSO';                 Value = ($sp.Count - $sso) }
                @{ Label = 'Used last 30 days';           Value = @($sp | Where-Object { "$($_.UsedWithin30Days)" -eq 'Yes' }).Count }
                @{ Label = 'Uses Graph API';              Value = @($sp | Where-Object { "$($_.UsesGraphPermissions)" -eq 'Yes' }).Count }
                @{ Label = 'Apps with risk';              Value = @($sp | Where-Object { (Get-AppRiskFinding -App $_).Risk }).Count }
            )
        }
    }

    if ($roleAssign.Count) {
        $out['Rbac'] = @(
            @{ Label = 'Distinct roles';         Value = @($roleAssign | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count }
            @{ Label = 'Role assignments';       Value = $roleAssign.Count }
            @{ Label = 'Privileged assignments'; Value = @($roleAssign | Where-Object { $privRoleList -contains "$($_.RoleDefinitionName)" }).Count }
            @{ Label = 'Principals > 1 role';    Value = @($roleAssign | Group-Object PrincipalId | Where-Object { @($_.Group | Select-Object -ExpandProperty RoleDefinitionName -Unique).Count -gt 1 }).Count }
            @{ Label = 'Time-bound';             Value = @($roleAssign | Where-Object { $_.EndDateTime }).Count }
        )
    }

    return $out
}
