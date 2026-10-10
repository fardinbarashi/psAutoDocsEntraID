# Autodoc v1.0
function Export-EnterpriseAppData {
    <#
        Section 1.5 - Enterprise applications (service principals) in their own
        right: EVERY service principal in the tenant, including gallery apps and
        managed identities that have no app registration. Captures type, origin,
        owners, SSO configuration, assignment requirement, credentials and
        sign-in activity. Written to its own 'EnterpriseApps' export so it is
        independent of the app-registration collector.

        Requires Application.Read.All (and AuditLog.Read.All for sign-in activity).

        NOTE: this collector performs live Microsoft Graph calls and cannot be
        run offline; verify the run against a real tenant.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [int]$UsageDays = 30
    )

    $Section = 'Section 1.5 : Enterprise Application (service principal) data'
    try {
        Write-Host "Start $Section"

        Write-Host "Loading service principal sign-in activity (requires AuditLog.Read.All)..." -ForegroundColor Yellow
        $spSignInActivityLookup = Get-ServicePrincipalSignInActivityLookup
        $cutoffDate = (Get-Date).AddDays(-$UsageDays)
        Write-Host "Cutoff for 'used within $UsageDays days': $cutoffDate" -ForegroundColor DarkGray

        $tenantId = $null
        try { $tenantId = (Get-MgContext).TenantId } catch { $tenantId = $null }

        # Well-known Microsoft tenants that own first-party service principals.
        # Used to flag Microsoft apps so the map can filter them out (the portal's
        # "Enterprise applications" view hides these by default).
        $msTenantIds = @(
            'f8cdef31-a31e-4b4a-93e4-5f571e91255a',
            '72f988bf-86f1-41af-91ab-2d7cd011db47',
            'cdc5aeea-15c5-4db6-b079-fcadd2505dc2'
        )

        # AppIds that have an app registration in THIS tenant, so we can flag
        # service principals that exist without a matching app registration.
        Write-Host "Indexing app registrations (to flag SP-only apps)..." -ForegroundColor Yellow
        $appRegAppIds = @{}
        try {
            foreach ($a in (Get-MgApplication -All -Property @("Id","AppId"))) {
                if ($a.AppId) { $appRegAppIds[$a.AppId] = $true }
            }
        } catch { Write-Host "Could not list app registrations: $($_.Exception.Message)" -ForegroundColor Yellow }

        Write-Host "Get Enterprise Applications (all service principals)..." -ForegroundColor Yellow
        $servicePrincipals = Get-MgServicePrincipal -All -Property @(
            "Id","AppId","DisplayName","ServicePrincipalType","AccountEnabled",
            "AppOwnerOrganizationId","Tags","PreferredSingleSignOnMode",
            "AppRoleAssignmentRequired","Homepage","SignInAudience",
            "KeyCredentials","PasswordCredentials"
        )

        $enterpriseAppData = @()

        foreach ($sp in $servicePrincipals) {
            Write-Host "Process SP: $($sp.DisplayName)" -ForegroundColor Cyan

            $owners = @()
            try { $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $sp.Id -All }
            catch { $owners = @() }
            $ownersList = (
                $owners | ForEach-Object {
                    if     ($_.AdditionalProperties.userPrincipalName) { $_.AdditionalProperties.userPrincipalName }
                    elseif ($_.AdditionalProperties.displayName)       { $_.AdditionalProperties.displayName }
                    else   { $_.Id }
                }
            ) -join ";"
            $ownerInfo = if ([string]::IsNullOrWhiteSpace($ownersList)) { "MissingOwners" } else { $ownersList }

            $secretExpiryDates   = @()
            $nearestSecretExpiry = $null
            if ($sp.PasswordCredentials) {
                $secretExpiryDates   = $sp.PasswordCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { ([datetime]$_.EndDateTime).ToString("yyyy-MM-dd") }
                $nearestSecretExpiry = $sp.PasswordCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { [datetime]$_.EndDateTime } | Sort-Object | Select-Object -First 1
            }

            $certificateExpiryDates   = @()
            $nearestCertificateExpiry = $null
            if ($sp.KeyCredentials) {
                $certificateExpiryDates   = $sp.KeyCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { ([datetime]$_.EndDateTime).ToString("yyyy-MM-dd") }
                $nearestCertificateExpiry = $sp.KeyCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { [datetime]$_.EndDateTime } | Sort-Object | Select-Object -First 1
            }

            $credentialDates = @()
            if ($sp.PasswordCredentials) { $credentialDates += $sp.PasswordCredentials.EndDateTime }
            if ($sp.KeyCredentials)      { $credentialDates += $sp.KeyCredentials.EndDateTime }
            $nearestExpiry = $credentialDates | Where-Object { $_ } | ForEach-Object { [datetime]$_ } | Sort-Object | Select-Object -First 1

            $spTodayDate = Get-Date
            $expiryLimit = $spTodayDate.AddDays(30)
            $credentialExpired     = $false
            $credentialExpiresSoon = $false
            foreach ($credentialDate in $credentialDates) {
                if (-not $credentialDate) { continue }
                $credentialDateTime = [datetime]$credentialDate
                if     ($credentialDateTime -lt $spTodayDate) { $credentialExpired = $true }
                elseif ($credentialDateTime -lt $expiryLimit) { $credentialExpiresSoon = $true }
            }

            $ssoType = "None"
            switch ($sp.PreferredSingleSignOnMode) {
                "saml"         { $ssoType = "SAML" }
                "oidc"         { $ssoType = "OIDC" }
                "password"     { $ssoType = "Password SSO" }
                "external"     { $ssoType = "Federated" }
                "notSupported" { $ssoType = "Not Supported" }
                default        { $ssoType = if ($sp.PreferredSingleSignOnMode) { "OAuth/API" } else { "None" } }
            }

            # Sign-in activity / used within N days
            $lastSignIn   = $null
            $usedRecently = "Unknown"
            $signInInfo   = $null
            if ($sp.AppId -and $spSignInActivityLookup.ContainsKey($sp.AppId)) {
                $signInInfo = $spSignInActivityLookup[$sp.AppId]
                $allDates = @(
                    $signInInfo.LastSignInActivity,
                    $signInInfo.LastDelegatedClientSignInActivity,
                    $signInInfo.LastDelegatedResourceSignInActivity,
                    $signInInfo.LastApplicationAuthClientSignInActivity,
                    $signInInfo.LastApplicationAuthResourceSignInActivity
                ) | Where-Object { $_ } | ForEach-Object { [datetime]$_ }

                if ($allDates) {
                    $lastSignIn   = ($allDates | Sort-Object -Descending | Select-Object -First 1)
                    $usedRecently = if ($lastSignIn -ge $cutoffDate) { "Yes" } else { "No" }
                }
                else { $usedRecently = "No" }
            }
            else { $usedRecently = "No (no activity data)" }

            $appOrigin = if ($tenantId -and "$($sp.AppOwnerOrganizationId)" -eq "$tenantId") { "This tenant" }
                         elseif ($sp.AppOwnerOrganizationId) { "External / Microsoft" }
                         else { "Unknown" }

            $enterpriseAppData += [pscustomobject]@{
                AppName                       = $sp.DisplayName
                AppId                         = $sp.AppId
                ServicePrincipalObjectId      = $sp.Id
                ServicePrincipalType          = $sp.ServicePrincipalType
                AppOrigin                     = $appOrigin
                AppOwnerOrganizationId        = "$($sp.AppOwnerOrganizationId)"
                IsMicrosoftApp                = if ($sp.AppOwnerOrganizationId -and ($msTenantIds -contains "$($sp.AppOwnerOrganizationId)".ToLower())) { 'Yes' } else { 'No' }
                HasAppRegistration            = if ($sp.AppId -and $appRegAppIds.ContainsKey($sp.AppId)) { "Yes" } else { "No" }
                AccountEnabled                = if ($sp.AccountEnabled) { "Yes" } else { "No" }
                OwnerInfo                     = $ownerInfo
                SSOConfigured                 = if ($sp.PreferredSingleSignOnMode) { "Yes" } else { "No" }
                SSOType                       = $ssoType
                AppRoleAssignmentRequired     = if ($sp.AppRoleAssignmentRequired) { "Yes" } else { "No" }
                Homepage                      = $sp.Homepage
                Tags                          = ($sp.Tags -join ";")
                SignInAudience                = $sp.SignInAudience
                HasSecrets                    = if (@($sp.PasswordCredentials).Count -gt 0) { "Yes" } else { "No" }
                HasCertificates               = if (@($sp.KeyCredentials).Count -gt 0) { "Yes" } else { "No" }
                SecretCount                   = @($sp.PasswordCredentials).Count
                CertificateCount              = @($sp.KeyCredentials).Count
                SecretExpiryDates             = ($secretExpiryDates -join ";")
                CertificateExpiryDates        = ($certificateExpiryDates -join ";")
                NearestSecretExpiry           = if ($nearestSecretExpiry)      { $nearestSecretExpiry.ToString("yyyy-MM-dd") }      else { $null }
                NearestCertificateExpiry      = if ($nearestCertificateExpiry) { $nearestCertificateExpiry.ToString("yyyy-MM-dd") } else { $null }
                NearestCredentialExpiry       = if ($nearestExpiry)            { $nearestExpiry.ToString("yyyy-MM-dd") }            else { $null }
                CredentialExpired             = if ($credentialExpired)     { "Yes" } else { "No" }
                CredentialExpiresWithin30Days = if ($credentialExpiresSoon) { "Yes" } else { "No" }
                LastSignInDateTime            = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                UsedWithin30Days              = $usedRecently
                LastDelegatedClientSignIn     = $signInInfo.LastDelegatedClientSignInActivity
                LastApplicationAuthSignIn     = $signInInfo.LastApplicationAuthClientSignInActivity
            }
        }

        Write-Host "Done. Enterprise apps collected: $(@($enterpriseAppData).Count)" -ForegroundColor Green

        Write-Host "Export EnterpriseApps data to files..."
        Export-InventoryData -Data $enterpriseAppData -BaseName 'EnterpriseApps' -WorksheetName 'EnterpriseApps' -Paths $Paths

        Write-Host "End $Section" -ForegroundColor Green

        return @{ EnterpriseApps = $enterpriseAppData }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
