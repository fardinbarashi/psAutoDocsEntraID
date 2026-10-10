# Autodoc v1.0
# ---------------------------------------------------------------------------
# Shared DATA for the maps. These only read and shape the exported
# data - they do NOT build any drawing model. Each SVG map builder makes its
# own shapes from this data, so a change to one map never affects another.
# ---------------------------------------------------------------------------

function Get-LicenceMapData {
    <# Returns the purchased + free/viral SKU nodes, grouped and ordered. #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $skus   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    if (-not $skus) { return $null }

    $nodes = foreach ($s in $skus) {
        $enabled  = [int]$s.enabledUnits
        $consumed = [int]$s.consumedUnits
        [pscustomobject]@{
            Sku      = [string]$s.skuPartNumber
            Category = Get-LicenseSkuCategory -SkuPartNumber $s.skuPartNumber -EnabledUnits $enabled
            Enabled  = $enabled
            Consumed = $consumed
            Util     = if ($enabled) { [math]::Round(100 * $consumed / $enabled) } else { 0 }
            PercentFree = if ($enabled) { [math]::Round(100 * ($enabled - $consumed) / $enabled) } else { 0 }
            Status   = if ("$($s.capabilityStatus)".Trim()) { "$($s.capabilityStatus)".Trim() } else { 'Unknown' }
        }
    }
    $purchased = @($nodes | Where-Object Category -eq 'Purchased' | Sort-Object Consumed -Descending)
    $viral     = @($nodes | Where-Object Category -eq 'Free/Viral' | Sort-Object Consumed -Descending)

    # group purchased SKUs by category so related ones sit together
    $catPath = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) 'Config\LicenseCategory.psd1'
    $groups = @()
    if (Test-Path $catPath) { try { $groups = @((Import-PowerShellDataFile $catPath).Groups) } catch { } }
    if ($groups.Count) {
        $purchased = @($purchased |
            Sort-Object @{ e = { Get-SkuGroupIndexShared -Sku $_.Sku -Groups $groups } },
                        @{ e = { -$_.Consumed } })
    }

    [pscustomobject]@{ Tenant = $tenant; Purchased = $purchased; Viral = $viral }
}

function Get-SkuGroupIndexShared {
    param([string]$Sku, $Groups)
    $u = $Sku.ToUpperInvariant()
    for ($g = 0; $g -lt $Groups.Count; $g++) {
        foreach ($x in @($Groups[$g].Skus)) { if ($x.ToUpperInvariant() -eq $u) { return $g } }
    }
    for ($g = 0; $g -lt $Groups.Count; $g++) {
        foreach ($m in @($Groups[$g].Match)) { if ($u -like "*$($m.ToUpperInvariant())*") { return $g } }
    }
    return $Groups.Count
}

function Get-OrgMapData {
    <# Returns the workplace -> unit -> count tree, ordered largest first. #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $users  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    if (-not $users) { return $null }

    $tree = @{}
    foreach ($u in $users) {
        $comp = if ("$($u.CompanyName)".Trim()) { "$($u.CompanyName)".Trim() } else { '(unknown workplace)' }
        $dept = if ("$($u.Department)".Trim())  { "$($u.Department)".Trim() }  else { '(no unit)' }
        if (-not $tree.ContainsKey($comp)) { $tree[$comp] = @{} }
        if (-not $tree[$comp].ContainsKey($dept)) { $tree[$comp][$dept] = 0 }
        $tree[$comp][$dept]++
    }
    $workplaces = @($tree.Keys | Sort-Object { -($tree[$_].Values | Measure-Object -Sum).Sum })

    [pscustomobject]@{ Tenant = $tenant; Tree = $tree; Workplaces = $workplaces; UserCount = $users.Count }
}

function Get-LicenceFreeFill {
    <# Red <=10% free, Yellow 11-30%, Green 31-100% free. #>
    param([int]$PercentFree)
    if ($PercentFree -le 10) { return '#F8696B' }
    if ($PercentFree -le 30) { return '#FFEB84' }
    return '#63BE7B'
}

function Get-GroupsMapData {
    <#
        Returns the groups split by category (Security / M365), each carrying the
        fields the map needs: member count, whether it has an owner, and whether
        it is cloud-native or synced from on-prem AD.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $groups = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    if (-not $groups) { return $null }

    $nodes = foreach ($g in $groups) {
        $members  = 0; [void][int]::TryParse("$($g.MemberCount)".Trim(), [ref]$members)
        $hasOwner = [bool]("$($g.Owners)".Trim())
        $synced   = ("$($g.OnPremisesSyncEnabled)".Trim().ToLower() -eq 'true') -or ("$($g.Source)".Trim() -eq 'Windows Server')
        [pscustomobject]@{
            Name     = [string]$g.DisplayName
            Category = if ("$($g.GroupCategory)".Trim()) { "$($g.GroupCategory)".Trim() } else { 'Other' }
            Members  = $members
            HasOwner = $hasOwner
            Synced   = $synced          # true = on-prem synced, false = cloud-native
            Empty    = ($members -eq 0)
            MailOn   = ("$($g.MailEnabled)".Trim().ToLower() -eq 'true')
        }
    }

    # Security first, then M365, then anything else; largest membership first
    $order = @{ 'Security' = 0; 'M365' = 1 }
    $byCat = $nodes | Group-Object Category | Sort-Object @{ e = { if ($order.ContainsKey($_.Name)) { $order[$_.Name] } else { 9 } } }, Name

    [pscustomobject]@{ Tenant = $tenant; Categories = $byCat; Total = $nodes.Count }
}

function Get-GroupDeptMapData {
    <#
        From GroupMembers, builds group -> department -> count. Used by the three
        group/department views (list card, stacked-bar card, matrix).
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $members = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'GroupMembers')
    if (-not $members) { return $null }

    $tree = @{}          # group -> @{ dept -> count }
    $deptTotals = @{}    # dept -> total (for matrix ordering / colour keys)
    foreach ($m in $members) {
        $grp  = "$($m.GroupDisplayName)".Trim(); if (-not $grp) { $grp = '(unknown group)' }
        $dept = "$($m.Department)".Trim();       if (-not $dept) { $dept = '(no department)' }
        if (-not $tree.ContainsKey($grp)) { $tree[$grp] = @{} }
        if (-not $tree[$grp].ContainsKey($dept)) { $tree[$grp][$dept] = 0 }
        $tree[$grp][$dept]++
        if (-not $deptTotals.ContainsKey($dept)) { $deptTotals[$dept] = 0 }
        $deptTotals[$dept]++
    }

    $groups = @($tree.Keys | Sort-Object { -($tree[$_].Values | Measure-Object -Sum).Sum })
    $depts  = @($deptTotals.Keys | Sort-Object { -$deptTotals[$_] })

    [pscustomobject]@{ Tenant = $tenant; Tree = $tree; Groups = $groups; Departments = $depts; TotalMembers = $members.Count }
}

function Get-LegendHeight {
    <#
        Consistent dashed legend height: a fixed top and bottom padding plus one
        line-height per line, so every legend has the SAME breathing room above
        and below its text regardless of how many lines it has. Pair this with a
        vertically-centred legend (do NOT set LinesTop) for even spacing.
    #>
    param([int]$LineCount, [double]$LineH = 0.24, [double]$Pad = 0.22)
    return ($Pad * 2) + ($LineCount * $LineH)
}

function Get-GroupOwnerMapData {
    <#
        From EntraGroups.Owners, aggregates by owner: which groups each owner
        owns, and the total members across those groups. An owner value is a
        UPN/email; a group can have several owners (semicolon-separated).
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $groups = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    if (-not $groups) { return $null }

    $byOwner = @{}
    foreach ($g in $groups) {
        $owners = "$($g.Owners)".Trim()
        if (-not $owners) { continue }
        $members = 0; [void][int]::TryParse("$($g.MemberCount)".Trim(), [ref]$members)
        foreach ($o in ($owners -split ';')) {
            $owner = $o.Trim()
            if (-not $owner) { continue }
            if (-not $byOwner.ContainsKey($owner)) {
                $byOwner[$owner] = [pscustomobject]@{ Owner = $owner; Groups = @(); Members = 0 }
            }
            $byOwner[$owner].Groups   += [string]$g.DisplayName
            $byOwner[$owner].Members  += $members
        }
    }

    $list = @($byOwner.Values | Sort-Object { $_.Groups.Count } -Descending)
    [pscustomobject]@{ Tenant = $tenant; Owners = $list; OwnedGroupCount = @($groups | Where-Object { "$($_.Owners)".Trim() }).Count; TotalGroups = $groups.Count }
}

function Get-ConditionalAccessMapData {
    <#
        Reads the ConditionalAccess export and returns just the policy records
        (RecordType 'Policy'), shaped for the map: name, state, and the resolved
        who / apps / conditions / grant-control fields. Named locations and auth
        strengths in the same file are skipped here (they are reference rows).
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $records = Expand-CaRecords (Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess')
    if (-not $records) { return $null }

    $policies = @($records | Where-Object { "$($_.RecordType)" -eq 'Policy' })
    if (-not $policies) { return $null }

    $nodes = foreach ($p in $policies) {
        $joinOrAll = {
            param($v)
            $t = "$v".Trim()
            if (-not $t) { return '' }
            return $t
        }
        [pscustomobject]@{
            Name        = [string]$p.DisplayName
            State       = "$($p.State)".Trim()
            IncludeUsers = & $joinOrAll $p.IncludeUsersResolved
            ExcludeUsers = & $joinOrAll $p.ExcludeUsersResolved
            IncludeGroups = & $joinOrAll $p.IncludeGroupsResolved
            ExcludeGroups = & $joinOrAll $p.ExcludeGroupsResolved
            IncludeRoles = & $joinOrAll $p.IncludeRolesResolved
            IncludeApps  = & $joinOrAll $p.IncludeApplicationsResolved
            ExcludeApps  = & $joinOrAll $p.ExcludeApplicationsResolved
            Locations    = & $joinOrAll $p.IncludeLocationsResolved
            Platforms    = & $joinOrAll $p.IncludePlatforms
            DeviceFilter = & $joinOrAll $p.DeviceFilterRule
            Grant        = & $joinOrAll $p.BuiltInControls
            AuthStrength = & $joinOrAll $p.AuthenticationStrengthResolved
            ExcludeRoles      = & $joinOrAll $p.ExcludeRolesResolved
            ExcludeLocations  = & $joinOrAll $p.ExcludeLocationsResolved
            ExcludePlatforms  = & $joinOrAll $p.ExcludePlatforms
            UserActions       = & $joinOrAll $p.IncludeUserActionsRaw
            ClientAppTypes    = & $joinOrAll $p.ClientAppTypes
            GrantOperator     = & $joinOrAll $p.GrantOperator
            SignInRisk        = & $joinOrAll $p.SignInRiskLevels
            UserRisk          = & $joinOrAll $p.UserRiskLevels
            SpRisk            = & $joinOrAll $p.ServicePrincipalRiskLevels
            SignInFreq        = ("$($p.SignInFrequencyValue) $($p.SignInFrequencyType)").Trim()
            PersistentBrowser = & $joinOrAll $p.PersistentBrowserMode
            AppEnforced       = & $joinOrAll $p.AppEnforcedRestrictionsEnabled
            CloudAppSecurity  = & $joinOrAll $p.CloudAppSecurityType
            TermsOfUse        = & $joinOrAll $p.TermsOfUse
        }
    }
    # enabled first, then report-only, then disabled; then by name
    $stateOrder = @{ 'enabled' = 0; 'enabledForReportingButNotEnforced' = 1; 'disabled' = 2 }
    $nodes = @($nodes | Sort-Object @{ e = { $o = $stateOrder["$($_.State)"]; if ($null -ne $o) { $o } else { 9 } } }, Name)

    [pscustomobject]@{ Tenant = $tenant; Policies = $nodes; Total = $nodes.Count }
}

function Get-AppRegMapData {
    <#
        Reads the AppRegEnterpriseApps export and shapes each app registration for
        the map: name, owner, SignInAudience (used for banding), credential health
        (expired / expiring / healthy), API-permission count, public-client flag
        and last sign-in. Same file also feeds the enterprise-apps map.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $apps   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    if (-not $apps) { return $null }

    $nodes = foreach ($a in $apps) {
        $expired  = ("$($a.CredentialExpired)".Trim() -eq 'Yes')
        $expiring = ("$($a.CredentialExpiresWithin30Days)".Trim() -eq 'Yes')
        $health   = if ($expired) { 'expired' } elseif ($expiring) { 'expiring' } else { 'healthy' }
        $apiCount = 0; [void][int]::TryParse("$($a.ApiPermissionCount)".Trim(), [ref]$apiCount)
        [pscustomobject]@{
            Name          = [string]$a.AppName
            Owner         = "$($a.OwnerInfo)".Trim()
            Audience      = if ("$($a.SignInAudience)".Trim()) { "$($a.SignInAudience)".Trim() } else { 'Unknown' }
            Health        = $health
            ApiCount      = $apiCount
            UsesGraph     = ("$($a.UsesGraphPermissions)".Trim() -eq 'Yes')
            PublicClient  = ("$($a.PublicClient)".Trim() -eq 'Yes')
            NearestExpiry = "$($a.NearestCredentialExpiry)".Trim()
            LastSignIn    = "$($a.LastSignInDateTime)".Trim()
        }
    }

    # friendlier audience labels + band order
    $audienceLabel = {
        param($a)
        switch ("$a") {
            'AzureADMyOrg'                     { 'This tenant only' }
            'AzureADMultipleOrgs'              { 'Any Microsoft tenant' }
            'AzureADandPersonalMicrosoftAccount' { 'Any tenant + personal' }
            'PersonalMicrosoftAccount'         { 'Personal accounts' }
            default                            { "$a" }
        }
    }
    foreach ($n in $nodes) { $n | Add-Member -NotePropertyName AudienceLabel -NotePropertyValue (& $audienceLabel $n.Audience) -Force }

    $order = @{ 'This tenant only' = 0; 'Any Microsoft tenant' = 1; 'Any tenant + personal' = 2; 'Personal accounts' = 3 }
    $byAud = $nodes | Group-Object AudienceLabel |
        Sort-Object @{ e = { $o = $order[$_.Name]; if ($null -ne $o) { $o } else { 9 } } }, Name

    [pscustomobject]@{ Tenant = $tenant; Audiences = $byAud; Total = $nodes.Count }
}

function Get-EnterpriseAppMapData {
    <#
        From the same AppRegEnterpriseApps export, shapes each app for the
        enterprise-app view: name, SSO type (for banding), whether SSO is
        configured (for colour), whether a service principal exists, last
        sign-in and Graph usage.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $apps   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps')
    if (-not $apps) { $apps = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps') }  # fallback for older exports
    if (-not $apps) { return $null }

    $nodes = foreach ($a in $apps) {
        [pscustomobject]@{
            Name        = [string]$a.AppName
            SSOType     = if ("$($a.SSOType)".Trim()) { "$($a.SSOType)".Trim() } else { 'None' }
            SSOOn       = ("$($a.SSOConfigured)".Trim() -eq 'Yes')
            SpExists    = ("$($a.ServicePrincipalExists)".Trim() -eq 'Yes')
            UsesGraph   = ("$($a.UsesGraphPermissions)".Trim() -eq 'Yes')
            LastSignIn  = "$($a.LastSignInDateTime)".Trim()
            Owner       = "$($a.OwnerInfo)".Trim()
        }
    }

    # band order: real SSO types first, then None/OAuth-API last
    $order = @{ 'SAML' = 0; 'OIDC' = 1; 'Password SSO' = 2; 'Federated' = 3; 'OAuth/API' = 4; 'Not Supported' = 5; 'None' = 6 }
    $bySso = $nodes | Group-Object SSOType |
        Sort-Object @{ e = { $o = $order[$_.Name]; if ($null -ne $o) { $o } else { 8 } } }, Name

    $ssoOnN = @($nodes | Where-Object SSOOn).Count
    [pscustomobject]@{ Tenant = $tenant; SsoTypes = $bySso; Total = $nodes.Count; SsoConfigured = $ssoOnN }
}

function Get-MapIconFolder {
    <#
        Single source of truth for where map icons live. This file sits in
        ...\Reports\Svg\Common\, so four parents up is the Entra root; the
        icons are runtime resources under files\cache. Change this one function
        to move the icon folder for every map at once.

        -BuilderRoot is accepted for compatibility; the location is resolved
        from this file, so it works from any caller.
    #>
    param([string]$BuilderRoot)
    $entraRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
    return Join-Path $entraRoot 'files\cache\Azure icons\Svg'
}

function Get-DepartmentMapData {
    <#
        From GroupMembers, counts distinct people per Department and per
        OfficeLocation. Used by the overview tree's 'Departments' and 'Office
        locations' areas. Counts unique users (by UPN) so a person in several
        groups isn't counted twice.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $members = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'GroupMembers')
    if (-not $members) { return $null }

    $deptUsers   = @{}   # dept   -> set of UPNs
    $officeUsers = @{}   # office -> set of UPNs
    foreach ($m in $members) {
        $upn = "$($m.UserPrincipalName)".Trim(); if (-not $upn) { $upn = "$($m.DisplayName)".Trim() }
        $dept   = "$($m.Department)".Trim();     if (-not $dept)   { $dept   = '(no department)' }
        $office = "$($m.OfficeLocation)".Trim(); if (-not $office) { $office = '(no office)' }
        if (-not $deptUsers.ContainsKey($dept))     { $deptUsers[$dept]     = New-Object System.Collections.Generic.HashSet[string] }
        if (-not $officeUsers.ContainsKey($office)) { $officeUsers[$office] = New-Object System.Collections.Generic.HashSet[string] }
        [void]$deptUsers[$dept].Add($upn)
        [void]$officeUsers[$office].Add($upn)
    }

    $departments = foreach ($k in $deptUsers.Keys) {
        [pscustomobject]@{ Name = $k; Members = $deptUsers[$k].Count }
    }
    $offices = foreach ($k in $officeUsers.Keys) {
        [pscustomobject]@{ Name = $k; Members = $officeUsers[$k].Count }
    }

    [pscustomobject]@{
        Tenant      = $tenant
        Departments = @($departments | Sort-Object Members -Descending)
        Offices     = @($offices     | Sort-Object Members -Descending)
    }
}

function Get-TenantHybridLabel {
    <#
        Turns the tenant's HybridStatus into a short environment label for the
        tenant block: "Cloud", "Onprem", or "Cloud + Onprem". Handles both the
        current collector wording and older exports:
          active sync            -> Cloud + Onprem
          previously synced      -> Onprem
          never synced / cloud   -> Cloud
        Returns $null if the field is missing so callers can skip the line.
    #>
    param($Tenant)
    if (-not $Tenant) { return $null }
    $hs = "$($Tenant.HybridStatus)".Trim()
    if (-not $hs) { return "Environment: Cloud" }

    # new wording is already the final value
    if ($hs -eq 'Cloud + Onprem') { return "Environment: Cloud + Onprem" }
    if ($hs -eq 'Onprem')         { return "Environment: Onprem" }
    if ($hs -eq 'Cloud')          { return "Environment: Cloud" }

    # older wording: map from the descriptive text
    if ($hs -match '(?i)active sync')       { return "Environment: Cloud + Onprem" }
    if ($hs -match '(?i)previously synced') { return "Environment: Onprem" }
    if ($hs -match '(?i)never synced')      { return "Environment: Cloud" }
    if ($hs -match '^(?i)yes')              { return "Environment: Cloud + Onprem" }
    return "Environment: Cloud"
}

function Get-DomainMapData {
    <#
        Reads TenantInformationDomain and groups the verified domains by Type
        (Managed / Federated / None), flagging which one is the default. Feeds the
        domain map and the overview tree.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $domains = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformationDomain')
    if (-not $domains) { return $null }

    $nodes = foreach ($d in $domains) {
        [pscustomobject]@{
            Name      = [string]$d.Domain
            Type      = if ("$($d.Type)".Trim()) { "$($d.Type)".Trim() } else { 'None' }
            IsDefault = ("$($d.IsDefault)".Trim().ToLower() -eq 'true')
        }
    }

    $order = @{ 'Managed' = 0; 'Federated' = 1; 'None' = 2 }
    $byType = $nodes | Group-Object Type |
        Sort-Object @{ e = { $o = $order[$_.Name]; if ($null -ne $o) { $o } else { 9 } } }, Name

    [pscustomobject]@{ Tenant = $tenant; Types = $byType; Total = $nodes.Count; Default = (@($nodes | Where-Object IsDefault | Select-Object -First 1).Name) }
}

function Get-RbacMapData {
    <#
        Reads the RBAC export and organises active + eligible role assignments by
        role. Each principal (user or group) that holds a role is captured with
        its type and whether the assignment is Active or Eligible (PIM).
        RoleAssignableGroup rows are reference data and skipped here.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $records = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC')
    if (-not $records) { return $null }

    $assignments = @($records | Where-Object { "$($_.RecordType)" -in 'RoleAssignmentActive','RoleAssignmentEligible' })
    if (-not $assignments) { return $null }

    $nodes = foreach ($a in $assignments) {
        [pscustomobject]@{
            Role      = if ("$($a.RoleDefinitionName)".Trim()) { "$($a.RoleDefinitionName)".Trim() } else { '(unknown role)' }
            Principal = if ("$($a.PrincipalDisplayName)".Trim()) { "$($a.PrincipalDisplayName)".Trim() } else { '(unknown)' }
            PType     = if ("$($a.PrincipalType)".Trim()) { "$($a.PrincipalType)".Trim() } else { 'Unknown' }
            Kind      = if ("$($a.RecordType)" -eq 'RoleAssignmentEligible') { 'Eligible' } else { 'Active' }
        }
    }

    # group by role, biggest first
    $byRole = $nodes | Group-Object Role | Sort-Object { -$_.Count }, Name
    # distinct principals (for the matrix columns)
    $principals = @($nodes | Select-Object -ExpandProperty Principal -Unique | Sort-Object)

    [pscustomobject]@{
        Tenant     = $tenant
        Roles      = $byRole
        Principals = $principals
        Total      = $nodes.Count
        ActiveN    = @($nodes | Where-Object Kind -eq 'Active').Count
        EligibleN  = @($nodes | Where-Object Kind -eq 'Eligible').Count
    }
}

function Get-UsersMapData {
    <#
        Reads UserInformation and produces the summary breakdowns for the users
        map: counts by UserType and UserCategory, enabled vs disabled accounts,
        users per domain (from the UPN), how many lack a mailbox, job-title
        distribution, and authentication-method usage (including how many use
        more than one method).
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $users  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    if (-not $users) { return $null }

    $val = { param($x,$k) "$($x.$k)".Trim() }

    $typeCounts = @{}; $catCounts = @{}; $domainCounts = @{}; $titleCounts = @{}; $authCounts = @{}
    $enabled = 0; $disabled = 0; $noMail = 0; $multiAuth = 0

    foreach ($u in $users) {
        $t = & $val $u 'UserType';     if (-not $t) { $t = '(unknown)' }
        $c = & $val $u 'UserCategory'; if (-not $c) { $c = '(unknown)' }
        $typeCounts[$t] = 1 + $typeCounts[$t]
        $catCounts[$c]  = 1 + $catCounts[$c]

        $ae = (& $val $u 'AccountEnabled').ToLower()
        if ($ae -eq 'true') { $enabled++ } elseif ($ae -eq 'false') { $disabled++ }

        if (-not (& $val $u 'Mail')) { $noMail++ }

        $upn = & $val $u 'UserPrincipalName'
        if ($upn -match '@(.+)$') { $dom = $Matches[1].ToLower(); $domainCounts[$dom] = 1 + $domainCounts[$dom] }

        $jt = & $val $u 'JobTitle'; if (-not $jt) { $jt = '(no title)' }
        $titleCounts[$jt] = 1 + $titleCounts[$jt]

        $methods = @((& $val $u 'AuthenticationMethods') -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        foreach ($m in $methods) { $authCounts[$m] = 1 + $authCounts[$m] }
        if ($methods.Count -gt 1) { $multiAuth++ }
    }

    $toSorted = { param($h) @($h.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { [pscustomobject]@{ Name = $_.Key; Count = $_.Value } }) }

    [pscustomobject]@{
        Tenant     = $tenant
        Total      = $users.Count
        Types      = & $toSorted $typeCounts
        Categories = & $toSorted $catCounts
        Domains    = & $toSorted $domainCounts
        Titles     = & $toSorted $titleCounts
        Auth       = & $toSorted $authCounts
        Enabled    = $enabled
        Disabled   = $disabled
        NoMail     = $noMail
        MultiAuth  = $multiAuth
    }
}

function Get-AppFullData {
    <#
        Returns the full app records (all fields) split into app registrations and
        enterprise apps for the overview tree. Both come from the same
        AppRegEnterpriseApps export; every column is passed through so the tree
        can show all values. The dedicated map builders keep their trimmed shape.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $apps = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    if (-not $apps) { return $null }
    # same rows feed both views; the tree labels them differently
    [pscustomobject]@{ Apps = $apps; Total = $apps.Count }
}

function Get-MapFieldGlossary {
    # field name -> short meaning, shared by the "Data source" info block on the
    # maps. Fields not listed here are shown by name only.
    @{
        AppName                       = 'Display name of the application'
        AppId                         = 'Application (client) ID'
        ObjectId                      = 'Directory object ID of the app registration'
        ServicePrincipalObjectId      = 'Object ID of the enterprise app (service principal)'
        ServicePrincipalExists        = 'Whether a service principal exists for the app'
        ServicePrincipalType          = 'SP kind: Application, ManagedIdentity, Legacy, SocialIdp'
        AppOrigin                     = 'Home of the app: this tenant vs external/Microsoft'
        AppOwnerOrganizationId        = 'Tenant ID of the organisation that owns/published the app'
        IsMicrosoftApp                = 'Whether the app is a Microsoft first-party application'
        HasAppRegistration            = 'Whether an app registration exists for this SP'
        AccountEnabled                = 'Whether the account/app is enabled'
        AccountStatus                 = 'Account status text (Enabled / Disabled)'
        OwnerInfo                     = 'Owners, or MissingOwners when none set'
        SignInAudience                = 'Who can sign in (single tenant, multi-tenant, personal)'
        PublicClient                  = 'Whether the app is a public (native) client'
        SSOConfigured                 = 'Whether single sign-on is configured'
        SSOType                       = 'Single sign-on mode (SAML, OIDC, Password, ...)'
        AppRoleAssignmentRequired     = 'Whether user assignment is required to sign in'
        Homepage                      = 'App home page URL'
        Tags                          = 'Service principal tags'
        HasSecrets                    = 'Whether the app has client secrets'
        HasCertificates               = 'Whether the app has certificates'
        SecretCount                   = 'Number of client secrets'
        CertificateCount              = 'Number of certificates'
        SecretExpiryDates             = 'Expiry date(s) of client secrets'
        CertificateExpiryDates        = 'Expiry date(s) of certificates'
        NearestSecretExpiry           = 'Soonest secret expiry'
        NearestCertificateExpiry      = 'Soonest certificate expiry'
        NearestCredentialExpiry       = 'Soonest of any credential expiry'
        CredentialExpired             = 'Whether any credential has already expired'
        CredentialExpiresWithin30Days = 'Whether a credential expires within 30 days'
        UsesGraphPermissions          = 'Whether the app holds Microsoft Graph permissions'
        ApiPermissions                = 'Raw API permission IDs granted'
        ApiPermissionNames            = 'Readable API permissions granted'
        ApiPermissionCount            = 'Number of API permissions'
        LastSignInDateTime            = 'Most recent sign-in seen for the app'
        LastDelegatedClientSignIn     = 'Most recent delegated (user) sign-in'
        LastApplicationAuthSignIn     = 'Most recent app-only sign-in'
        UsedWithin30Days              = 'Whether the app was used in the last 30 days'
        DisplayName                   = 'Display name'
        State                         = 'Policy state (enabled, disabled, report-only)'
        Domain                        = 'Domain name'
        Type                          = 'Domain type (Managed, Federated, None)'
        IsDefault                     = 'Whether this is the default domain'
        Role                          = 'Directory role name'
        RoleName                      = 'Directory role name'
        UserPrincipalName             = 'User principal name (sign-in name)'
        AssignmentType                = 'Role assignment type (active / eligible)'
        MemberCount                   = 'Number of members'
        OwnerCount                    = 'Number of owners'
        GroupType                     = 'Group type (security, Microsoft 365, ...)'
        Department                    = 'Department'
        OnPremisesSyncEnabled         = 'Whether the object is synced from on-prem AD'
        RecordType                    = 'Internal record type marker'
        PolicyId                      = 'Conditional Access policy ID'
        IncludeUsersRaw               = 'Targeted users (raw object IDs)'
        IncludeUsersResolved          = 'Targeted users (resolved names)'
        ExcludeUsersRaw               = 'Excluded users (raw object IDs)'
        ExcludeUsersResolved          = 'Excluded users (resolved names)'
        IncludeGroupsRaw              = 'Targeted groups (raw object IDs)'
        IncludeGroupsResolved         = 'Targeted groups (resolved names)'
        ExcludeGroupsRaw              = 'Excluded groups (raw object IDs)'
        ExcludeGroupsResolved         = 'Excluded groups (resolved names)'
        IncludeRolesRaw               = 'Targeted directory roles (raw IDs)'
        IncludeRolesResolved          = 'Targeted directory roles (resolved names)'
        ExcludeRolesRaw               = 'Excluded directory roles (raw IDs)'
        ExcludeRolesResolved          = 'Excluded directory roles (resolved names)'
        IncludeApplicationsRaw        = 'Targeted apps (raw IDs, or All)'
        IncludeApplicationsResolved   = 'Targeted apps (resolved names)'
        ExcludeApplicationsRaw        = 'Excluded apps (raw IDs)'
        ExcludeApplicationsResolved   = 'Excluded apps (resolved names)'
        IncludeUserActionsRaw         = 'Targeted user actions (e.g. register security info)'
        IncludeLocationsRaw           = 'Targeted locations (raw IDs, or All)'
        IncludeLocationsResolved      = 'Targeted locations (resolved names)'
        ExcludeLocationsRaw           = 'Excluded locations (raw IDs)'
        ExcludeLocationsResolved      = 'Excluded locations (resolved names)'
        ClientAppTypes                = 'Client app types the policy applies to'
        IncludePlatforms              = 'Targeted device platforms'
        ExcludePlatforms              = 'Excluded device platforms'
        SignInRiskLevels              = 'Sign-in risk levels that trigger the policy'
        UserRiskLevels                = 'User risk levels that trigger the policy'
        ServicePrincipalRiskLevels    = 'Workload-identity risk levels that trigger the policy'
        DeviceFilterMode              = 'Device filter mode (include / exclude)'
        DeviceFilterRule              = 'Device filter rule expression'
        GrantOperator                 = 'How grant controls combine (AND / OR)'
        BuiltInControls               = 'Required grant controls (MFA, compliant device, ...)'
        CustomAuthenticationFactors   = 'Custom authentication factors required'
        TermsOfUse                    = 'Terms of use that must be accepted'
        AuthenticationStrengthId      = 'Authentication strength ID'
        AuthenticationStrengthResolved= 'Authentication strength (resolved name)'
        SignInFrequencyType           = 'Sign-in frequency unit (hours / days)'
        SignInFrequencyValue          = 'Sign-in frequency value'
        PersistentBrowserMode         = 'Persistent browser session control'
        AppEnforcedRestrictionsEnabled= 'Whether app-enforced restrictions are on'
        CloudAppSecurityType          = 'Conditional Access App Control (MCAS) mode'
        NamedLocationId               = 'Named location ID'
        NamedLocationType             = 'Named location type (IP / country)'
        CountriesAndRegions           = 'Countries/regions for the named location'
        IncludeUnknownCountriesAndRegions = 'Whether unknown countries are included'
        IpRanges                      = 'IP ranges for the named location'
        IsTrusted                     = 'Whether the location is marked trusted'
        PolicyType                    = 'Policy type (builtIn / custom)'
        RequirementsSatisfied         = 'Requirements the auth strength satisfies'
        AllowedCombinations           = 'Allowed authentication method combinations'
        CreatedDateTime               = 'When the object was created'
        ModifiedDateTime              = 'When the object was last modified'
        ExportedAt                    = 'When this export was generated'
        ConditionalAccessPolicies     = 'Array of Conditional Access policies'
        ConditionalAccessNamedLocations = 'Array of named locations'
        ConditionalAccessAuthStrengths= 'Array of authentication strengths'
        TenantId                      = 'Directory (tenant) ID'
        TenantName                    = 'Tenant display name'
        PrimaryDomain                 = 'Primary (default) domain'
        CountryCode                   = 'Tenant country code'
        SkuPartNumber                 = 'License SKU part number'
        SkuId                         = 'License SKU ID'
        ConsumedUnits                 = 'Licenses assigned'
        PrepaidEnabled                = 'Prepaid (available) license units'
        ServicePlans                  = 'Service plans included in the SKU'
        AssignedLicenses              = 'Licenses assigned to the user'
        Manager                       = 'The user''s manager'
        MfaMethods                    = 'Registered MFA methods'
        MfaCapable                    = 'Whether the user can perform MFA'
        Members                       = 'Group members'
        Owners                        = 'Group owners'
        MembershipType                = 'Membership type (assigned / dynamic)'
        MembershipRule                = 'Dynamic membership rule'
        IsAssignableToRole            = 'Whether the group can hold directory roles'
        OnPremisesSyncEnabledGroup    = 'Whether the group is synced from on-prem AD'
        RoleDisplayName               = 'Directory role name'
        PrincipalDisplayName          = 'Assigned principal (user/group/SP) name'
        PrincipalType                 = 'Assigned principal type'
        DirectoryScope                = 'Scope of the role assignment'
        SsprEnabledScope              = 'Who self-service password reset is enabled for'
        MethodsAvailableToUsers       = 'Authentication methods available for reset'
        SecurityQuestionsEnabled      = 'Whether security questions are enabled'

        # --- groups (EntraGroups) ---
        Id                            = 'Directory object ID'
        Description                   = 'Free-text description'
        Source                        = 'Where the object is mastered (Cloud or on-premises AD)'
        GroupCategory                 = 'Group category (Security, Microsoft 365, ...)'
        GroupTypes                    = 'Underlying group type flags (e.g. Unified, DynamicMembership)'
        MailEnabled                   = 'Whether the group has a mailbox / can receive mail'
        SecurityEnabled               = 'Whether the group can be used to grant access (security group)'
        MailNickname                  = 'Mail alias (nickname) of the group'
        Visibility                    = 'Group visibility (Public, Private, HiddenMembership)'
        DynamicGroupMembershipRule    = 'Rule that decides membership for a dynamic group'
        DynamicMembershipRuleProcessingState = 'Whether the dynamic membership rule is running or paused'
        ResourceProvisioningOptions   = 'Extra provisioning options (e.g. Team)'
        Licenses                      = 'Licences assigned to the group (group-based licensing)'
        RolesAndAdministrators        = 'Directory roles assigned to the group'
        AdministrativeUnits           = 'Administrative units the group belongs to'
        GroupMemberships              = 'Groups this group is a member of (nested membership)'
        Applications                  = 'Applications assigned to the group'
        WelcomeEmailEnabled           = 'Whether a welcome email is sent to new members'

        # --- group totals (EntraGroupsBasicInfo) ---
        TotalGroups                   = 'Total number of groups in the tenant'
        M365Groups                    = 'Number of Microsoft 365 groups'
        SecurityGroups                = 'Number of security groups'
        DynamicGroups                 = 'Number of dynamic-membership groups'
        CloudGroups                   = 'Number of cloud-only groups'
        OnPremisesGroups              = 'Number of groups synced from on-premises AD'

        # --- group members (GroupMembers) ---
        GroupId                       = 'Directory object ID of the group'
        GroupDisplayName              = 'Display name of the group'
        UserId                        = 'Directory object ID of the member'
        OfficeLocation                = 'Office location'
        ManagerDisplayName            = 'Display name of the manager'

        # --- licences (LicensesInformation) ---
        enabledUnits                  = 'Number of licence units purchased / enabled'
        freeUnits                     = 'Number of unused (free) licence units'
        capabilityStatus              = 'Subscription status (Enabled, Warning, Suspended)'

        # --- self-service password reset (PasswordReset) ---
        SsprSelectedGroupId           = 'Group scoped for self-service password reset'
        AdministratorSsprEnabled      = 'Whether SSPR is enabled for administrators'
        AdministratorMethodsRequired  = 'Number of methods admins must register for SSPR'
        AdministratorMethodsAvailable = 'Authentication methods available to admins for SSPR'

        # --- role assignments (RBAC) ---
        RoleDefinitionName            = 'Directory role name'
        RoleDefinitionId              = 'Directory role definition ID'
        PrincipalId                   = 'Object ID of the principal holding the role'
        MemberType                    = 'How the role is held (Direct, Group, Inherited)'
        Status                        = 'Status of the assignment / record'
        StartDateTime                 = 'When the assignment starts (time-bound roles)'
        EndDateTime                   = 'When the assignment ends (blank = permanent)'
        JITActivated                  = 'Whether an eligible role is currently activated (just-in-time)'

        # --- tenant (TenantInformation) ---
        NotificationLanguage          = 'Default notification language for the tenant'
        TechnicalContact              = 'Technical contact address(es)'
        GlobalPrivacyContact          = 'Privacy contact for the tenant'
        PrivacyStatementUrl           = 'URL of the tenant privacy statement'
        Users                         = 'Total number of users in the tenant'
        Groups                        = 'Total number of groups in the tenant'
        AppRegistrations              = 'Total number of app registrations'
        Devices                       = 'Total number of registered devices'
        HybridStatus                  = 'Directory hybrid status (cloud-only vs synced with on-prem AD)'

        # --- users (UserInformation) ---
        UserType                      = 'Account type (Member or Guest)'
        UserCategory                  = 'Derived user category (e.g. internal / external)'
        Mail                          = 'Primary email address'
        AssignedPlans                 = 'Service plans assigned to the user'
        AssignedLicenseCount          = 'Number of licences assigned to the user'
        AssignedPlanCount             = 'Number of service plans assigned to the user'
        MobilePhone                   = 'Mobile phone number'
        JobTitle                      = 'Job title'
        CompanyName                   = 'Company name'
        ManagerId                     = 'Object ID of the user''s manager'
        AuthenticationMethods         = 'Registered authentication methods (MFA, ...)'
        AuthenticationMethodCount     = 'Number of registered authentication methods'
    }
}

function Split-MapTextToWidth {
    # Shared word-wrap used by the "Data source" info block. Breaks on natural
    # separators and hard-splits any single token wider than a line. The first
    # line can reserve room for a bold prefix. Always returns >= 1 line.
    param([string]$Text, [double]$AvailIn, [double]$FontSize, [double]$FirstReserveIn = 0)
    $Text = "$Text"
    $out = @()
    $cur = ''
    $firstBudget = $AvailIn - $FirstReserveIn
    $tokens = [regex]::Split($Text, '(?<=[\s;,/_@-])')
    foreach ($tok in $tokens) {
        $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
        while ((Measure-MapTextWidth $tok $FontSize) -gt $budget -and $tok.Length -gt 1) {
            $budget = if ($out.Count -eq 0 -and -not $cur) { $firstBudget } else { $AvailIn }
            $n = $tok.Length
            while ($n -gt 1 -and (Measure-MapTextWidth ($cur + $tok.Substring(0, $n)) $FontSize) -gt $budget) { $n-- }
            $out += ($cur + $tok.Substring(0, $n)).TrimEnd()
            $cur = ''
            $tok = $tok.Substring($n)
        }
        $budget = if ($out.Count -eq 0) { $firstBudget } else { $AvailIn }
        if ($cur -and (Measure-MapTextWidth ($cur + $tok) $FontSize) -gt $budget) {
            $out += $cur.TrimEnd(); $cur = $tok
        }
        else { $cur += $tok }
    }
    if ($cur) { $out += $cur.TrimEnd() }
    if ($out.Count -eq 0) { $out = @('') }
    return $out
}

function Get-MapDataSourceBlock {
    <#
        Builds the shared "Data source" info block shown at the top of a map:
        the path to the JSON file the map was built from, and every field in that
        JSON with a short meaning. Returns @{ Lines; Height; Width; Path } ready
        to drop in as a Rectangle shape.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [Parameter(Mandatory)][string]$BaseName,
        [double]$WidthIn = 8.0,
        [double]$FontSize = 8,
        [double]$LineH = 0.19,
        [double]$Pad = 0.14,
        [switch]$PathOnly,
        [switch]$GridFields,
        [int]$GridMaxColumns = 10
    )

    # resolve the actual JSON file the map reads (newest match), else show pattern
    $jsonDir = Join-Path $SourceFolder 'rawDataJson'
    $path = Join-Path $jsonDir "$BaseName*.json"
    $fields = @()
    try {
        $file = Get-ChildItem -Path $jsonDir -Filter "$BaseName*.json" -ErrorAction Stop |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($file) {
            $path = $file.FullName
            if (-not $PathOnly) {
                $rows = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName $BaseName)
            if ($rows.Count) {
                $first = $rows[0]
                # Wrapper-style exports (e.g. Conditional Access) hold the real
                # records inside an array-valued property; descend into the
                # largest such array so we list the record fields, not the wrapper.
                $arrProps = @($first.PSObject.Properties | Where-Object {
                    $_.Value -is [System.Collections.IEnumerable] -and $_.Value -isnot [string] -and @($_.Value).Count -gt 0 -and @($_.Value)[0].PSObject
                })
                if ($rows.Count -eq 1 -and $arrProps.Count) {
                    $biggest = $arrProps | Sort-Object { @($_.Value).Count } -Descending | Select-Object -First 1
                    $first = @($biggest.Value)[0]
                }
                $fields = @($first.PSObject.Properties.Name)
            }
            }
        }
    } catch { }

    $glossary = Get-MapFieldGlossary
    $availIn  = $WidthIn - 2 * $Pad - (6 / 96)   # minus padding; block has no icon

    $lines = @()
    $lines += @{ Text = 'Data source'; Bold = $true; Align = 'start' }

    # JSON path (bold prefix, wrapped) - trimmed to start at \Entra\ so the long
    # absolute prefix (C:\Users\...\Current Projects\...) is dropped
    $reserve = Measure-MapTextWidth 'JSON: ' $FontSize
    $shortPath = $path -replace '^.*?(\\Entra\\)', '$1'
    $pathParts = @(Split-MapTextToWidth -Text $shortPath -AvailIn $availIn -FontSize $FontSize -FirstReserveIn $reserve)
    $lines += @{ BoldPrefix = 'JSON: '; Text = $pathParts[0]; Align = 'start' }
    for ($k = 1; $k -lt $pathParts.Count; $k++) { $lines += @{ Text = $pathParts[$k]; Bold = $false; Align = 'start' } }

    $gridFrom = 0
    $columns = 1
    if ($fields.Count -and -not $PathOnly) {
        $hdr = if ($GridFields) { 'Fields in this JSON:' } else { 'Fields in this JSON and what they mean:' }
        $lines += @{ Text = $hdr; Bold = $true; Align = 'start' }
        $gridFrom = $lines.Count
        foreach ($f in $fields) {
            if ($GridFields) {
                # compact grid: field name only (descriptive on its own)
                $lines += @{ Text = $f; Bold = $true; Align = 'start' }
                continue
            }
            $meaning = if ($glossary.ContainsKey($f)) { $glossary[$f] } else { '' }
            $meaning = ("$meaning" -replace '\s*\([^)]*\)', '').Trim()   # drop parenthetical text
            if ($meaning) {
                $prefix  = "$f - "
                $reserve = Measure-MapTextWidth $prefix $FontSize
                $parts   = @(Split-MapTextToWidth -Text $meaning -AvailIn $availIn -FontSize $FontSize -FirstReserveIn $reserve)
                $lines += @{ BoldPrefix = $prefix; Text = $parts[0]; Align = 'start' }
                for ($k = 1; $k -lt $parts.Count; $k++) { $lines += @{ Text = $parts[$k]; Bold = $false; Align = 'start' } }
            }
            else {
                $lines += @{ Text = $f; Bold = $true; Align = 'start' }
            }
        }

        if ($GridFields) {
            # pick a column count so the widest field name fits
            $maxW = 0.5
            for ($j = $gridFrom; $j -lt $lines.Count; $j++) {
                $wv = Measure-MapTextWidth "$($lines[$j].Text)" $FontSize
                if ($wv -gt $maxW) { $maxW = $wv }
            }
            $columns = [math]::Max(1, [math]::Floor($availIn / ($maxW + 0.25)))
            $columns = [math]::Min($columns, $GridMaxColumns)
        }
    }

    if ($columns -gt 1 -and $gridFrom -gt 0) {
        $n = $lines.Count - $gridFrom
        $rowsPerCol = [math]::Ceiling($n / $columns)
        $height = $Pad * 2 + ($gridFrom + $rowsPerCol) * $LineH
    }
    else {
        $height = $Pad * 2 + $lines.Count * $LineH
    }
    [pscustomobject]@{ Lines = $lines; Height = $height; Width = $WidthIn; Path = $path; Columns = $columns; GridFrom = $gridFrom }
}

function Get-EnterpriseAppFullData {
    <#
        Returns the full enterprise-app (service principal) records for the
        enterprise-app map. Prefers the dedicated 'EnterpriseApps' export (all
        service principals, incl. gallery apps and managed identities); falls
        back to the combined 'AppRegEnterpriseApps' export so older exports still
        render.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $apps = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps')
    if (-not $apps) { $apps = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps') }
    if (-not $apps) { return $null }
    [pscustomobject]@{ Apps = $apps; Total = $apps.Count }
}

function Get-AppDetailLines {
    <#
        Turns one app record into an ordered list of "Field: value" lines for the
        tree, skipping empty/null values. Field set differs by view: 'appreg'
        emphasises credentials & permissions, 'enterprise' emphasises SSO &
        service principal. Shared identity fields are shown in both.
    #>
    param([Parameter(Mandatory)]$App, [ValidateSet('appreg','enterprise')]$View = 'appreg')

    $spec = if ($View -eq 'appreg') {
        @(
            @('App ID','AppId'), @('Object ID','ObjectId'), @('Owner','OwnerInfo'),
            @('Sign-in audience','SignInAudience'), @('Public client','PublicClient'),
            @('Secrets','SecretCount'), @('Certificates','CertificateCount'),
            @('Secret expiry','SecretExpiryDates'), @('Cert expiry','CertificateExpiryDates'),
            @('Nearest cred expiry','NearestCredentialExpiry'), @('Credential expired','CredentialExpired'),
            @('Expires within 30d','CredentialExpiresWithin30Days'), @('API permissions','ApiPermissionCount'),
            @('Uses Graph','UsesGraphPermissions'), @('Permission names','ApiPermissionNames'),
            @('Last sign-in','LastSignInDateTime'), @('Used within 30d','UsedWithin30Days')
        )
    } else {
        @(
            @('App ID','AppId'), @('Service principal ID','ServicePrincipalObjectId'),
            @('SP type','ServicePrincipalType'), @('Origin','AppOrigin'),
            @('Has app registration','HasAppRegistration'), @('Account enabled','AccountEnabled'),
            @('Owner','OwnerInfo'), @('SSO configured','SSOConfigured'), @('SSO type','SSOType'),
            @('Assignment required','AppRoleAssignmentRequired'), @('Homepage','Homepage'),
            @('Tags','Tags'), @('Sign-in audience','SignInAudience'),
            @('Secrets','SecretCount'), @('Certificates','CertificateCount'),
            @('Secret expiry','SecretExpiryDates'), @('Cert expiry','CertificateExpiryDates'),
            @('Nearest cred expiry','NearestCredentialExpiry'), @('Credential expired','CredentialExpired'),
            @('Expires within 30d','CredentialExpiresWithin30Days'),
            @('Last sign-in','LastSignInDateTime'), @('Last delegated sign-in','LastDelegatedClientSignIn'),
            @('Last app-auth sign-in','LastApplicationAuthSignIn'), @('Used within 30d','UsedWithin30Days')
        )
    }

    $lines = @()
    foreach ($pair in $spec) {
        $value = "$($App.($pair[1]))".Trim()
        if ($value -and $value -ne 'null') { $lines += "$($pair[0]): $value" }
    }
    return $lines
}

function Get-CaFullData {
    <#
        Returns the full CA policy records (all fields) for the overview tree.
        Only RecordType 'Policy' rows; named locations and auth strengths are
        reference rows and skipped.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)
    $records = Expand-CaRecords (Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess')
    if (-not $records) { return $null }
    $policies = @($records | Where-Object { "$($_.RecordType)" -eq 'Policy' })
    if (-not $policies) { return $null }
    [pscustomobject]@{ Policies = $policies; Total = $policies.Count }
}

function Get-CaWhoLines {
    <#
        Turns a Conditional Access "WHO" target into readable rich lines: each
        user on its own bold name line followed by a non-bold "UPN: ..." line
        (split from the resolved "Name <upn>" value). Roles and groups are listed
        too. Entries without a <upn> (e.g. "All users") are shown as plain lines.
    #>
    param(
        [string]$IncludeUsers, [string]$ExcludeUsers, [string]$Roles, [string]$Groups,
        [double]$AvailIn = 2.0, [double]$FontSize = 9
    )
    $out = [System.Collections.Generic.List[object]]::new()
    $out.Add(@{ Text = 'WHO:'; Bold = $true; Align = 'start' })

    foreach ($sec in @(
            [pscustomobject]@{ Head = ''; Val = $IncludeUsers },
            [pscustomobject]@{ Head = 'Excluded users:'; Val = $ExcludeUsers }
        )) {
        $val = "$($sec.Val)".Trim()
        if (-not $val -or $val -eq '-') { continue }
        if ($sec.Head) { $out.Add(@{ Text = $sec.Head; Bold = $true; Align = 'start' }) }
        foreach ($entry in ($val -split ';')) {
            $e = $entry.Trim()
            if (-not $e -or $e -eq '-') { continue }
            if ($e -match '^(.*?)\s*<([^>]+)>\s*$') {
                $nm = $Matches[1].Trim(); $upn = $Matches[2].Trim()
                foreach ($t in @(Split-MapTextToWidth -Text $nm -AvailIn $AvailIn -FontSize $FontSize)) { $out.Add(@{ Text = $t; Bold = $true; Align = 'start' }) }
                foreach ($t in @(Split-MapTextToWidth -Text "UPN: $upn" -AvailIn $AvailIn -FontSize $FontSize)) { $out.Add(@{ Text = $t; Bold = $false; Align = 'start' }) }
            }
            else {
                foreach ($t in @(Split-MapTextToWidth -Text $e -AvailIn $AvailIn -FontSize $FontSize)) { $out.Add(@{ Text = $t; Bold = $true; Align = 'start' }) }
            }
        }
    }
    if ($Roles  -and "$Roles".Trim()  -ne '-') { foreach ($t in @(Split-MapTextToWidth -Text "Roles: $Roles"  -AvailIn $AvailIn -FontSize $FontSize)) { $out.Add(@{ Text = $t; Bold = $false; Align = 'start' }) } }
    if ($Groups -and "$Groups".Trim() -ne '-') { foreach ($t in @(Split-MapTextToWidth -Text "Groups: $Groups" -AvailIn $AvailIn -FontSize $FontSize)) { $out.Add(@{ Text = $t; Bold = $false; Align = 'start' }) } }
    $out.ToArray()
}

function Get-CaDetailLines {
    <#
        Turns one CA policy record into an ordered list of "Field: value" lines
        for the tree, skipping empty/null values so a simple policy stays short.
        Covers who/where/conditions/grant/session - every meaningful field the
        export carries.
    #>
    param([Parameter(Mandatory)]$Policy)

    $spec = @(
        @('State','State'),
        @('Include users','IncludeUsersResolved'),
        @('Exclude users','ExcludeUsersResolved'),
        @('Include groups','IncludeGroupsResolved'),
        @('Exclude groups','ExcludeGroupsResolved'),
        @('Include roles','IncludeRolesResolved'),
        @('Exclude roles','ExcludeRolesResolved'),
        @('Include apps','IncludeApplicationsResolved'),
        @('Exclude apps','ExcludeApplicationsResolved'),
        @('User actions','IncludeUserActionsRaw'),
        @('Include locations','IncludeLocationsResolved'),
        @('Exclude locations','ExcludeLocationsResolved'),
        @('Client app types','ClientAppTypes'),
        @('Include platforms','IncludePlatforms'),
        @('Exclude platforms','ExcludePlatforms'),
        @('Sign-in risk','SignInRiskLevels'),
        @('User risk','UserRiskLevels'),
        @('SP risk','ServicePrincipalRiskLevels'),
        @('Device filter mode','DeviceFilterMode'),
        @('Device filter rule','DeviceFilterRule'),
        @('Grant operator','GrantOperator'),
        @('Grant controls','BuiltInControls'),
        @('Custom auth factors','CustomAuthenticationFactors'),
        @('Terms of use','TermsOfUse'),
        @('Auth strength','AuthenticationStrengthResolved'),
        @('Sign-in frequency','SignInFrequencyValue'),
        @('Sign-in freq type','SignInFrequencyType'),
        @('Persistent browser','PersistentBrowserMode'),
        @('App-enforced restrictions','AppEnforcedRestrictionsEnabled'),
        @('Cloud app security','CloudAppSecurityType'),
        @('Named location','NamedLocationId'),
        @('Named location type','NamedLocationType'),
        @('Countries/regions','CountriesAndRegions'),
        @('IP ranges','IpRanges'),
        @('Created','CreatedDateTime'),
        @('Modified','ModifiedDateTime')
    )

    $lines = @()
    foreach ($pair in $spec) {
        $value = "$($Policy.($pair[1]))".Trim()
        if ($value -and $value -ne 'null') { $lines += "$($pair[0]): $value" }
    }
    return $lines
}

function Get-PasswordResetData {
    <#
        Reads the PasswordReset (SSPR) export - a single configuration record -
        and returns it with the tenant. SSPR is tenant-wide config, so there is
        just one row.
    #>
    param([Parameter(Mandatory)][string]$SourceFolder)

    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $sspr   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'PasswordReset') | Select-Object -First 1
    if (-not $sspr) { return $null }
    [pscustomobject]@{ Tenant = $tenant; Config = $sspr }
}

function Get-PasswordResetLines {
    <#
        Ordered "Field: value" lines for the SSPR configuration, skipping empty
        values. Used by both the SSPR card and the overview tree.
    #>
    param([Parameter(Mandatory)]$Config)

    $spec = @(
        @('User SSPR scope','SsprEnabledScope'),
        @('Selected group','SsprSelectedGroupId'),
        @('Methods available to users','MethodsAvailableToUsers'),
        @('Security questions enabled','SecurityQuestionsEnabled'),
        @('Admin SSPR enabled','AdministratorSsprEnabled'),
        @('Admin methods required','AdministratorMethodsRequired'),
        @('Admin methods available','AdministratorMethodsAvailable')
    )
    $lines = @()
    foreach ($pair in $spec) {
        $value = "$($Config.($pair[1]))".Trim()
        if ($value -and $value -ne 'null') { $lines += "$($pair[0]): $value" }
    }
    return $lines
}

function Expand-CaRecords {
    <#
        The ConditionalAccess JSON is saved as a wrapper object
        { ExportedAt, ConditionalAccessPolicies: [ ... ] } rather than a flat
        array. This unwraps it: if the loaded data is that wrapper, return the
        inner policy array; otherwise return the records as-is (e.g. the flat CSV
        fallback). Always returns an array.
    #>
    param($Records)
    $records = @($Records)
    if ($records.Count -eq 1 -and $records[0].PSObject.Properties.Name -contains 'ConditionalAccessPolicies') {
        return @($records[0].ConditionalAccessPolicies)
    }
    return $records
}
