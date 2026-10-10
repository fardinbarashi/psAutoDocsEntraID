# Autodoc v1.0
function Export-ConditionalAccessData {
    <#
        Section 1.5 - Conditional Access policies, named locations and
        authentication strength policies, with IDs resolved to display names.
        Requires Policy.Read.All (+ User/Group/Application.Read.All for lookups).

        CSV/Excel receive one flat combined table; JSON receives a structured
        object with the three record sets separated.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths
    )

    $Section = 'Section 1.5 : Conditional Access'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        $specialUsers = @{
            "All"                   = "All Users"
            "None"                  = "None"
            "GuestsOrExternalUsers" = "Guests or External Users"
        }
        $specialApps = @{
            "All"                   = "All Cloud Apps"
            "None"                  = "None"
            "Office365"             = "Office 365"
            "MicrosoftAdminPortals" = "Microsoft Admin Portals"
        }
        $specialLocations = @{
            "All"        = "All Locations"
            "AllTrusted" = "All Trusted Locations"
        }

        Write-Host "Build lookup tables..." -ForegroundColor Yellow

        $userLookup = @{}
        foreach ($user in (Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$user.Id)) {
                $userLookup[[string]$user.Id] = if ($user.UserPrincipalName) { "$($user.DisplayName) <$($user.UserPrincipalName)>" } else { $user.DisplayName }
            }
        }

        $groupLookup = @{}
        foreach ($group in (Get-MgGroup -All -Property "Id,DisplayName")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$group.Id)) { $groupLookup[[string]$group.Id] = $group.DisplayName }
        }

        $roleLookup = @{}
        foreach ($role in (Get-MgDirectoryRoleTemplate -All)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$role.Id)) { $roleLookup[[string]$role.Id] = $role.DisplayName }
        }

        $appLookup = @{}
        foreach ($app in (Get-MgApplication -All -Property "Id,AppId,DisplayName")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$app.AppId)) { $appLookup[[string]$app.AppId] = $app.DisplayName }
            if (-not [string]::IsNullOrWhiteSpace([string]$app.Id))    { $appLookup[[string]$app.Id]    = $app.DisplayName }
        }

        $servicePrincipalLookup = @{}
        foreach ($sp in (Get-MgServicePrincipal -All -Property "Id,AppId,DisplayName")) {
            if (-not [string]::IsNullOrWhiteSpace([string]$sp.AppId)) { $servicePrincipalLookup[[string]$sp.AppId] = $sp.DisplayName }
            if (-not [string]::IsNullOrWhiteSpace([string]$sp.Id))    { $servicePrincipalLookup[[string]$sp.Id]    = $sp.DisplayName }
        }

        $combinedAppLookup = @{}
        foreach ($key in $appLookup.Keys)              { $combinedAppLookup[$key] = $appLookup[$key] }
        foreach ($key in $servicePrincipalLookup.Keys) { $combinedAppLookup[$key] = $servicePrincipalLookup[$key] }

        $namedLocationLookup = @{}
        $namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All
        foreach ($namedLocation in $namedLocations) {
            if (-not [string]::IsNullOrWhiteSpace([string]$namedLocation.Id)) { $namedLocationLookup[[string]$namedLocation.Id] = $namedLocation.DisplayName }
        }

        $authStrengthLookup   = @{}
        $authStrengthPolicies = @()
        try {
            $authStrengthPolicies = Get-MgPolicyAuthenticationStrengthPolicy -All
            foreach ($authStrength in $authStrengthPolicies) {
                if (-not [string]::IsNullOrWhiteSpace([string]$authStrength.Id)) { $authStrengthLookup[[string]$authStrength.Id] = $authStrength.DisplayName }
            }
        }
        catch {
            Write-Host "Could not read Authentication Strength policies: $($_.Exception.Message)" -ForegroundColor Yellow
            $authStrengthPolicies = @()
        }

        Write-Host "Get Conditional Access policies..." -ForegroundColor Yellow
        $caPolicies = Get-MgIdentityConditionalAccessPolicy -All

        $conditionalAccessPolicyData = foreach ($policy in $caPolicies) {
            Write-Host "Process CA Policy: $($policy.DisplayName)" -ForegroundColor Cyan

            $includeUsersRaw  = if ($policy.Conditions.Users -and $policy.Conditions.Users.IncludeUsers)  { $policy.Conditions.Users.IncludeUsers  -join ";" } else { $null }
            $excludeUsersRaw  = if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeUsers)  { $policy.Conditions.Users.ExcludeUsers  -join ";" } else { $null }
            $includeGroupsRaw = if ($policy.Conditions.Users -and $policy.Conditions.Users.IncludeGroups) { $policy.Conditions.Users.IncludeGroups -join ";" } else { $null }
            $excludeGroupsRaw = if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeGroups) { $policy.Conditions.Users.ExcludeGroups -join ";" } else { $null }
            $includeRolesRaw  = if ($policy.Conditions.Users -and $policy.Conditions.Users.IncludeRoles)  { $policy.Conditions.Users.IncludeRoles  -join ";" } else { $null }
            $excludeRolesRaw  = if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeRoles)  { $policy.Conditions.Users.ExcludeRoles  -join ";" } else { $null }

            $includeApplicationsRaw = if ($policy.Conditions.Applications -and $policy.Conditions.Applications.IncludeApplications) { $policy.Conditions.Applications.IncludeApplications -join ";" } else { $null }
            $excludeApplicationsRaw = if ($policy.Conditions.Applications -and $policy.Conditions.Applications.ExcludeApplications) { $policy.Conditions.Applications.ExcludeApplications -join ";" } else { $null }
            $includeUserActionsRaw  = if ($policy.Conditions.Applications -and $policy.Conditions.Applications.IncludeUserActions)  { $policy.Conditions.Applications.IncludeUserActions  -join ";" } else { $null }

            $includeLocationsRaw = if ($policy.Conditions.Locations -and $policy.Conditions.Locations.IncludeLocations) { $policy.Conditions.Locations.IncludeLocations -join ";" } else { $null }
            $excludeLocationsRaw = if ($policy.Conditions.Locations -and $policy.Conditions.Locations.ExcludeLocations) { $policy.Conditions.Locations.ExcludeLocations -join ";" } else { $null }

            $includeUsersResolved  = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.IncludeUsers  -Lookup $userLookup  -SpecialMap $specialUsers
            $excludeUsersResolved  = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.ExcludeUsers  -Lookup $userLookup  -SpecialMap $specialUsers
            $includeGroupsResolved = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.IncludeGroups -Lookup $groupLookup
            $excludeGroupsResolved = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.ExcludeGroups -Lookup $groupLookup
            $includeRolesResolved  = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.IncludeRoles  -Lookup $roleLookup
            $excludeRolesResolved  = Resolve-CaIdsToNames -Ids $policy.Conditions.Users.ExcludeRoles  -Lookup $roleLookup

            $includeApplicationsResolved = Resolve-CaIdsToNames -Ids $policy.Conditions.Applications.IncludeApplications -Lookup $combinedAppLookup   -SpecialMap $specialApps
            $excludeApplicationsResolved = Resolve-CaIdsToNames -Ids $policy.Conditions.Applications.ExcludeApplications -Lookup $combinedAppLookup   -SpecialMap $specialApps
            $includeLocationsResolved    = Resolve-CaIdsToNames -Ids $policy.Conditions.Locations.IncludeLocations       -Lookup $namedLocationLookup -SpecialMap $specialLocations
            $excludeLocationsResolved    = Resolve-CaIdsToNames -Ids $policy.Conditions.Locations.ExcludeLocations       -Lookup $namedLocationLookup -SpecialMap $specialLocations

            $clientAppTypes             = if ($policy.Conditions.ClientAppTypes) { $policy.Conditions.ClientAppTypes -join ";" } else { $null }
            $includePlatforms           = if ($policy.Conditions.Platforms -and $policy.Conditions.Platforms.IncludePlatforms) { $policy.Conditions.Platforms.IncludePlatforms -join ";" } else { $null }
            $excludePlatforms           = if ($policy.Conditions.Platforms -and $policy.Conditions.Platforms.ExcludePlatforms) { $policy.Conditions.Platforms.ExcludePlatforms -join ";" } else { $null }
            $signInRiskLevels           = if ($policy.Conditions.SignInRiskLevels) { $policy.Conditions.SignInRiskLevels -join ";" } else { $null }
            $userRiskLevels             = if ($policy.Conditions.UserRiskLevels)   { $policy.Conditions.UserRiskLevels   -join ";" } else { $null }
            $servicePrincipalRiskLevels = if ($policy.Conditions.ServicePrincipalRiskLevels) { $policy.Conditions.ServicePrincipalRiskLevels -join ";" } else { $null }

            $deviceFilterMode = $null
            $deviceFilterRule = $null
            if ($policy.Conditions.Devices -and $policy.Conditions.Devices.DeviceFilter) {
                $deviceFilterMode = $policy.Conditions.Devices.DeviceFilter.Mode
                $deviceFilterRule = $policy.Conditions.Devices.DeviceFilter.Rule
            }

            $grantOperator     = if ($policy.GrantControls -and $policy.GrantControls.Operator)                    { $policy.GrantControls.Operator } else { $null }
            $builtInControls   = if ($policy.GrantControls -and $policy.GrantControls.BuiltInControls)             { $policy.GrantControls.BuiltInControls -join ";" } else { $null }
            $customAuthFactors = if ($policy.GrantControls -and $policy.GrantControls.CustomAuthenticationFactors) { $policy.GrantControls.CustomAuthenticationFactors -join ";" } else { $null }
            $termsOfUse        = if ($policy.GrantControls -and $policy.GrantControls.TermsOfUse)                  { $policy.GrantControls.TermsOfUse -join ";" } else { $null }

            $authenticationStrengthId       = $null
            $authenticationStrengthResolved = $null
            if ($policy.GrantControls -and $policy.GrantControls.AuthenticationStrength) {
                $authenticationStrengthId = [string]$policy.GrantControls.AuthenticationStrength.Id
                if (-not [string]::IsNullOrWhiteSpace($authenticationStrengthId) -and $authStrengthLookup.ContainsKey($authenticationStrengthId)) { $authenticationStrengthResolved = $authStrengthLookup[$authenticationStrengthId] }
                else { $authenticationStrengthResolved = $authenticationStrengthId }
            }

            $signInFrequencyType  = $null
            $signInFrequencyValue = $null
            if ($policy.SessionControls -and $policy.SessionControls.SignInFrequency) {
                $signInFrequencyType  = $policy.SessionControls.SignInFrequency.Type
                $signInFrequencyValue = $policy.SessionControls.SignInFrequency.Value
            }

            $persistentBrowserMode = $null
            if ($policy.SessionControls -and $policy.SessionControls.PersistentBrowser) { $persistentBrowserMode = $policy.SessionControls.PersistentBrowser.Mode }

            $appEnforcedRestrictionsEnabled = $null
            if ($policy.SessionControls -and $policy.SessionControls.ApplicationEnforcedRestrictions) { $appEnforcedRestrictionsEnabled = $policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled }

            $cloudAppSecurityType = $null
            if ($policy.SessionControls -and $policy.SessionControls.CloudAppSecurity) { $cloudAppSecurityType = $policy.SessionControls.CloudAppSecurity.CloudAppSecurityType }

            [pscustomobject]@{
                RecordType                        = "Policy"
                PolicyId                          = $policy.Id
                DisplayName                       = $policy.DisplayName
                State                             = $policy.State
                IncludeUsersRaw                   = $includeUsersRaw
                IncludeUsersResolved              = $includeUsersResolved
                ExcludeUsersRaw                   = $excludeUsersRaw
                ExcludeUsersResolved              = $excludeUsersResolved
                IncludeGroupsRaw                  = $includeGroupsRaw
                IncludeGroupsResolved             = $includeGroupsResolved
                ExcludeGroupsRaw                  = $excludeGroupsRaw
                ExcludeGroupsResolved             = $excludeGroupsResolved
                IncludeRolesRaw                   = $includeRolesRaw
                IncludeRolesResolved              = $includeRolesResolved
                ExcludeRolesRaw                   = $excludeRolesRaw
                ExcludeRolesResolved              = $excludeRolesResolved
                IncludeApplicationsRaw            = $includeApplicationsRaw
                IncludeApplicationsResolved       = $includeApplicationsResolved
                ExcludeApplicationsRaw            = $excludeApplicationsRaw
                ExcludeApplicationsResolved       = $excludeApplicationsResolved
                IncludeUserActionsRaw             = $includeUserActionsRaw
                IncludeLocationsRaw               = $includeLocationsRaw
                IncludeLocationsResolved          = $includeLocationsResolved
                ExcludeLocationsRaw               = $excludeLocationsRaw
                ExcludeLocationsResolved          = $excludeLocationsResolved
                ClientAppTypes                    = $clientAppTypes
                IncludePlatforms                  = $includePlatforms
                ExcludePlatforms                  = $excludePlatforms
                SignInRiskLevels                  = $signInRiskLevels
                UserRiskLevels                    = $userRiskLevels
                ServicePrincipalRiskLevels        = $servicePrincipalRiskLevels
                DeviceFilterMode                  = $deviceFilterMode
                DeviceFilterRule                  = $deviceFilterRule
                GrantOperator                     = $grantOperator
                BuiltInControls                   = $builtInControls
                CustomAuthenticationFactors       = $customAuthFactors
                TermsOfUse                        = $termsOfUse
                AuthenticationStrengthId          = $authenticationStrengthId
                AuthenticationStrengthResolved    = $authenticationStrengthResolved
                SignInFrequencyType               = $signInFrequencyType
                SignInFrequencyValue              = $signInFrequencyValue
                PersistentBrowserMode             = $persistentBrowserMode
                AppEnforcedRestrictionsEnabled    = $appEnforcedRestrictionsEnabled
                CloudAppSecurityType              = $cloudAppSecurityType
                NamedLocationId                   = $null
                NamedLocationType                 = $null
                CountriesAndRegions               = $null
                IncludeUnknownCountriesAndRegions = $null
                IpRanges                          = $null
                IsTrusted                         = $null
                PolicyType                        = $null
                RequirementsSatisfied             = $null
                AllowedCombinations               = $null
                CreatedDateTime                   = $null
                ModifiedDateTime                  = $null
            }
        }

        Write-Host "Process Named Locations..." -ForegroundColor Yellow
        $conditionalAccessNamedLocationData = foreach ($namedLocation in $namedLocations) {
            $odataType      = $namedLocation.AdditionalProperties.'@odata.type'
            $countries      = $null
            $includeUnknown = $null
            $ipRanges       = $null
            $isTrusted      = $null

            if ($namedLocation.AdditionalProperties.countriesAndRegions)                        { $countries      = ($namedLocation.AdditionalProperties.countriesAndRegions) -join ";" }
            if ($null -ne $namedLocation.AdditionalProperties.includeUnknownCountriesAndRegions) { $includeUnknown = $namedLocation.AdditionalProperties.includeUnknownCountriesAndRegions }
            if ($namedLocation.AdditionalProperties.ipRanges)                                   { $ipRanges       = ($namedLocation.AdditionalProperties.ipRanges | ForEach-Object { $_.cidrAddress }) -join ";" }
            if ($null -ne $namedLocation.AdditionalProperties.isTrusted)                        { $isTrusted      = $namedLocation.AdditionalProperties.isTrusted }

            [pscustomobject]@{
                RecordType                        = "NamedLocation"
                PolicyId                          = $null
                DisplayName                       = $namedLocation.DisplayName
                State                             = $null
                IncludeUsersRaw                   = $null; IncludeUsersResolved = $null; ExcludeUsersRaw = $null; ExcludeUsersResolved = $null
                IncludeGroupsRaw                  = $null; IncludeGroupsResolved = $null; ExcludeGroupsRaw = $null; ExcludeGroupsResolved = $null
                IncludeRolesRaw                   = $null; IncludeRolesResolved = $null; ExcludeRolesRaw = $null; ExcludeRolesResolved = $null
                IncludeApplicationsRaw            = $null; IncludeApplicationsResolved = $null; ExcludeApplicationsRaw = $null; ExcludeApplicationsResolved = $null
                IncludeUserActionsRaw             = $null
                IncludeLocationsRaw               = $null; IncludeLocationsResolved = $null; ExcludeLocationsRaw = $null; ExcludeLocationsResolved = $null
                ClientAppTypes                    = $null; IncludePlatforms = $null; ExcludePlatforms = $null
                SignInRiskLevels                  = $null; UserRiskLevels = $null; ServicePrincipalRiskLevels = $null
                DeviceFilterMode                  = $null; DeviceFilterRule = $null
                GrantOperator                     = $null; BuiltInControls = $null; CustomAuthenticationFactors = $null; TermsOfUse = $null
                AuthenticationStrengthId          = $null; AuthenticationStrengthResolved = $null
                SignInFrequencyType               = $null; SignInFrequencyValue = $null; PersistentBrowserMode = $null
                AppEnforcedRestrictionsEnabled    = $null; CloudAppSecurityType = $null
                NamedLocationId                   = $namedLocation.Id
                NamedLocationType                 = $odataType
                CountriesAndRegions               = $countries
                IncludeUnknownCountriesAndRegions = $includeUnknown
                IpRanges                          = $ipRanges
                IsTrusted                         = $isTrusted
                PolicyType                        = $null; RequirementsSatisfied = $null; AllowedCombinations = $null
                CreatedDateTime                   = $null; ModifiedDateTime = $null
            }
        }

        Write-Host "Process Authentication Strength Policies..." -ForegroundColor Yellow
        $conditionalAccessAuthStrengthData = foreach ($authStrength in $authStrengthPolicies) {
            [pscustomobject]@{
                RecordType                        = "AuthenticationStrength"
                PolicyId                          = $null
                DisplayName                       = $authStrength.DisplayName
                State                             = $null
                IncludeUsersRaw                   = $null; IncludeUsersResolved = $null; ExcludeUsersRaw = $null; ExcludeUsersResolved = $null
                IncludeGroupsRaw                  = $null; IncludeGroupsResolved = $null; ExcludeGroupsRaw = $null; ExcludeGroupsResolved = $null
                IncludeRolesRaw                   = $null; IncludeRolesResolved = $null; ExcludeRolesRaw = $null; ExcludeRolesResolved = $null
                IncludeApplicationsRaw            = $null; IncludeApplicationsResolved = $null; ExcludeApplicationsRaw = $null; ExcludeApplicationsResolved = $null
                IncludeUserActionsRaw             = $null
                IncludeLocationsRaw               = $null; IncludeLocationsResolved = $null; ExcludeLocationsRaw = $null; ExcludeLocationsResolved = $null
                ClientAppTypes                    = $null; IncludePlatforms = $null; ExcludePlatforms = $null
                SignInRiskLevels                  = $null; UserRiskLevels = $null; ServicePrincipalRiskLevels = $null
                DeviceFilterMode                  = $null; DeviceFilterRule = $null
                GrantOperator                     = $null; BuiltInControls = $null; CustomAuthenticationFactors = $null; TermsOfUse = $null
                AuthenticationStrengthId          = $authStrength.Id
                AuthenticationStrengthResolved    = $authStrength.DisplayName
                SignInFrequencyType               = $null; SignInFrequencyValue = $null; PersistentBrowserMode = $null
                AppEnforcedRestrictionsEnabled    = $null; CloudAppSecurityType = $null
                NamedLocationId                   = $null; NamedLocationType = $null; CountriesAndRegions = $null
                IncludeUnknownCountriesAndRegions = $null; IpRanges = $null; IsTrusted = $null
                PolicyType                        = $authStrength.PolicyType
                RequirementsSatisfied             = $authStrength.RequirementsSatisfied
                AllowedCombinations               = if ($authStrength.AllowedCombinations) { $authStrength.AllowedCombinations -join ";" } else { $null }
                CreatedDateTime                   = $authStrength.CreatedDateTime
                ModifiedDateTime                  = $authStrength.ModifiedDateTime
            }
        }

        $conditionalAccessCsv  = @()
        $conditionalAccessCsv += $conditionalAccessPolicyData
        $conditionalAccessCsv += $conditionalAccessNamedLocationData
        $conditionalAccessCsv += $conditionalAccessAuthStrengthData

        $conditionalAccessExport = [pscustomobject]@{
            ExportedAt                      = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ConditionalAccessPolicies       = $conditionalAccessPolicyData
            ConditionalAccessNamedLocations = $conditionalAccessNamedLocationData
            ConditionalAccessAuthStrengths  = $conditionalAccessAuthStrengthData
        }

        Write-Host "Export Conditional Access data to files..."
        Export-InventoryData -Data $conditionalAccessCsv -BaseName 'ConditionalAccess' -JsonData $conditionalAccessExport -Paths $Paths

        Write-Host "Done. CA records: $(@($conditionalAccessCsv).Count) (policies, named locations, auth strengths)." -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ ConditionalAccessCsv = $conditionalAccessCsv }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
