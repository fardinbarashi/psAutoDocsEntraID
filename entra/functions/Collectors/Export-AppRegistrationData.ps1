# Autodoc v1.0
function Export-AppRegistrationData {
    <#
        Section 1.4 - App registrations + matching enterprise apps (service
        principals): owners, secret/certificate expiry, API permissions, SSO
        type and sign-in activity ("used within N days").
        Requires Application.Read.All and AuditLog.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [int]$UsageDays = 30
    )

    $Section = 'Section 1.4 : App Registration and Enterprise Application data'
    try {
        Write-Host "Start $Section"

        Write-Host "Loading service principal sign-in activity (requires AuditLog.Read.All)..." -ForegroundColor Yellow
        $spSignInActivityLookup = Get-ServicePrincipalSignInActivityLookup
        $cutoffDate = (Get-Date).AddDays(-$UsageDays)
        Write-Host "Cutoff for 'used within $UsageDays days': $cutoffDate" -ForegroundColor DarkGray

        Write-Host "Get App-Registrations..." -ForegroundColor Yellow
        $getAppRegData = Get-MgApplication -All -Property @(
            "Id","AppId","DisplayName","PasswordCredentials","KeyCredentials",
            "RequiredResourceAccess","SignInAudience","IsFallbackPublicClient"
        )

        Write-Host "Get Enterprise Applications..." -ForegroundColor Yellow
        # AppRoles and Oauth2PermissionScopes are what let us turn permission
        # GUIDs into names like "User.Read.All" further down.
        $servicePrincipals = Get-MgServicePrincipal -All -Property @(
            "Id","AppId","DisplayName","PreferredSingleSignOnMode","AppRoles","Oauth2PermissionScopes"
        )

        Write-Host "Building API permission name lookup..." -ForegroundColor Yellow
        $permissionLookup = Get-ApiPermissionLookup -ServicePrincipals $servicePrincipals

        $servicePrincipalsLookup = @{}
        foreach ($sp in $servicePrincipals) {
            if ($sp.AppId) { $servicePrincipalsLookup[$sp.AppId] = $sp }
        }

        $appRegistrationData = @()

        foreach ($app in $getAppRegData) {
            Write-Host "Process App: $($app.DisplayName)" -ForegroundColor Cyan

            $servicePrincipal = $null
            if ($app.AppId -and $servicePrincipalsLookup.ContainsKey($app.AppId)) {
                $servicePrincipal = $servicePrincipalsLookup[$app.AppId]
            }

            $owners = @()
            try { $owners = Get-MgApplicationOwner -ApplicationId $app.Id -All }
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
            if ($app.PasswordCredentials) {
                $secretExpiryDates   = $app.PasswordCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { ([datetime]$_.EndDateTime).ToString("yyyy-MM-dd") }
                $nearestSecretExpiry = $app.PasswordCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { [datetime]$_.EndDateTime } | Sort-Object | Select-Object -First 1
            }

            $certificateExpiryDates   = @()
            $nearestCertificateExpiry = $null
            if ($app.KeyCredentials) {
                $certificateExpiryDates   = $app.KeyCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { ([datetime]$_.EndDateTime).ToString("yyyy-MM-dd") }
                $nearestCertificateExpiry = $app.KeyCredentials | Where-Object { $_.EndDateTime } | ForEach-Object { [datetime]$_.EndDateTime } | Sort-Object | Select-Object -First 1
            }

            $credentialDates = @()
            if ($app.PasswordCredentials) { $credentialDates += $app.PasswordCredentials.EndDateTime }
            if ($app.KeyCredentials)      { $credentialDates += $app.KeyCredentials.EndDateTime }
            $nearestExpiry = $credentialDates | Where-Object { $_ } | ForEach-Object { [datetime]$_ } | Sort-Object | Select-Object -First 1

            $appRegTodayDate = Get-Date
            $expiryLimit     = $appRegTodayDate.AddDays(30)
            $credentialExpired     = $false
            $credentialExpiresSoon = $false
            foreach ($credentialDate in $credentialDates) {
                if (-not $credentialDate) { continue }
                $credentialDateTime = [datetime]$credentialDate
                if     ($credentialDateTime -lt $appRegTodayDate) { $credentialExpired = $true }
                elseif ($credentialDateTime -lt $expiryLimit)     { $credentialExpiresSoon = $true }
            }

            $usesGraphPermissions = $false
            if ($app.RequiredResourceAccess.ResourceAppId -contains "00000003-0000-0000-c000-000000000000") { $usesGraphPermissions = $true }

            $resolvedPermissions = Resolve-ApiPermissionName -RequiredResourceAccess $app.RequiredResourceAccess -Lookup $permissionLookup

            $apiPermissions = (
                $app.RequiredResourceAccess | ForEach-Object {
                    $resourceAppId = $_.ResourceAppId
                    $permissionIds = ($_.ResourceAccess | ForEach-Object { "$($_.Type):$($_.Id)" }) -join ","
                    "$resourceAppId => $permissionIds"
                }
            ) -join "; "

            $ssoType = "None"
            if ($servicePrincipal) {
                switch ($servicePrincipal.PreferredSingleSignOnMode) {
                    "saml"         { $ssoType = "SAML" }
                    "oidc"         { $ssoType = "OIDC" }
                    "password"     { $ssoType = "Password SSO" }
                    "external"     { $ssoType = "Federated" }
                    "notSupported" { $ssoType = "Not Supported" }
                    default        { $ssoType = "OAuth/API" }
                }
            }

            # Sign-in activity / used within N days
            $lastSignIn   = $null
            $usedRecently = "Unknown"
            $signInInfo   = $null
            if ($app.AppId -and $spSignInActivityLookup.ContainsKey($app.AppId)) {
                $signInInfo = $spSignInActivityLookup[$app.AppId]
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

            $appRegistrationData += [pscustomobject]@{
                AppName                       = $app.DisplayName
                AppId                         = $app.AppId
                ObjectId                      = $app.Id
                ServicePrincipalObjectId      = if ($servicePrincipal) { $servicePrincipal.Id } else { $null }
                OwnerInfo                     = $ownerInfo
                HasSecrets                    = if (@($app.PasswordCredentials).Count -gt 0) { "Yes" } else { "No" }
                HasCertificates               = if (@($app.KeyCredentials).Count -gt 0) { "Yes" } else { "No" }
                SecretCount                   = @($app.PasswordCredentials).Count
                CertificateCount              = @($app.KeyCredentials).Count
                SecretExpiryDates             = ($secretExpiryDates -join ";")
                CertificateExpiryDates        = ($certificateExpiryDates -join ";")
                NearestSecretExpiry           = if ($nearestSecretExpiry)      { $nearestSecretExpiry.ToString("yyyy-MM-dd") }      else { $null }
                NearestCertificateExpiry      = if ($nearestCertificateExpiry) { $nearestCertificateExpiry.ToString("yyyy-MM-dd") } else { $null }
                NearestCredentialExpiry       = if ($nearestExpiry)            { $nearestExpiry.ToString("yyyy-MM-dd") }            else { $null }
                CredentialExpired             = if ($credentialExpired)        { "Yes" } else { "No" }
                CredentialExpiresWithin30Days = if ($credentialExpiresSoon)    { "Yes" } else { "No" }
                ServicePrincipalExists        = if ($servicePrincipal)         { "Yes" } else { "No" }
                SSOConfigured                 = if ($servicePrincipal -and $servicePrincipal.PreferredSingleSignOnMode) { "Yes" } else { "No" }
                SSOType                       = $ssoType
                PublicClient                  = if ($app.IsFallbackPublicClient) { "Yes" } else { "No" }
                SignInAudience                = $app.SignInAudience
                UsesGraphPermissions          = if ($usesGraphPermissions) { "Yes" } else { "No" }
                ApiPermissions                = $apiPermissions
                ApiPermissionNames            = $resolvedPermissions.Text
                ApiPermissionCount            = $resolvedPermissions.Count
                LastSignInDateTime            = if ($lastSignIn) { $lastSignIn.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
                UsedWithin30Days              = $usedRecently
                LastDelegatedClientSignIn     = $signInInfo.LastDelegatedClientSignInActivity
                LastApplicationAuthSignIn     = $signInInfo.LastApplicationAuthClientSignInActivity
            }
        }

        Write-Host "Done. Applications collected: $(@($appRegistrationData).Count)" -ForegroundColor Green

        Write-Host "Export AppReg and EnterpriseApps data to files..."
        Export-InventoryData -Data $appRegistrationData -BaseName 'AppRegEnterpriseApps' -WorksheetName 'appReg-EnterpriseApps' -Paths $Paths

        Write-Host "End $Section" -ForegroundColor Green

        return @{ AppRegEnterprise = $appRegistrationData }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
