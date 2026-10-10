# Autodoc v1.0
function Export-UserData {
    <#
        Section 1.2.3 - Users + assigned licenses/plans, manager and auth methods.
        Requires User.Read.All and UserAuthenticationMethod.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [hashtable]$LicenseRef,
        [Parameter(Mandatory)][hashtable]$GraphRef
    )

    $Section = 'Section 1.2.3 : User Data'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        # Fallback so license/plan resolution still works if the CSV reference
        # was not loaded (Graph lookups and GUID fallback still apply).
        if (-not $LicenseRef) {
            $LicenseRef = @{ SkuByGuid = @{}; SkuByStringId = @{}; PlanById = @{}; PlanByName = @{} }
        }

        Write-Host "Get User data... this can take a while for large tenants, please wait." -ForegroundColor Yellow

        $userProperties = @(
            'id','displayName','userPrincipalName','userType','mail','accountEnabled',
            'assignedLicenses','assignedPlans','onPremisesSyncEnabled','mobilePhone',
            'jobTitle','department','officeLocation','companyName'
        ) -join ','

        $users = Get-MgUser -All -Property $userProperties -ExpandProperty "manager(`$select=id,userPrincipalName,displayName)"
        if ($users -isnot [System.Collections.IEnumerable] -or $users -is [string]) { $users = @($users) }

        $userData = foreach ($userObject in $users) {
            Write-Host "Process user: $($userObject.UserPrincipalName)"

            $managerInfo = Get-ManagerInfo -User $userObject

            $licenseNames = @()
            foreach ($lic in $userObject.AssignedLicenses) {
                $resolved = Resolve-LicenseName -AssignedLicense $lic -LicenseRef $LicenseRef -GraphRef $GraphRef
                if ($resolved) { $licenseNames += $resolved }
            }

            $planNames = @()
            foreach ($plan in $userObject.AssignedPlans) {
                $resolved = Resolve-PlanName -AssignedPlan $plan -LicenseRef $LicenseRef -GraphRef $GraphRef -ExcludeRetired
                if ($resolved) { $planNames += $resolved }
            }

            $licenseNames = $licenseNames | Sort-Object -Unique
            $planNames    = $planNames    | Sort-Object -Unique

            $authInfo = Get-UserAuthenticationMethodInfo -UserId $userObject.Id

            [pscustomobject]@{
                Id                        = $userObject.Id
                DisplayName               = $userObject.DisplayName
                UserPrincipalName         = $userObject.UserPrincipalName
                UserType                  = $userObject.UserType
                UserCategory              = if ($userObject.UserType -eq 'Guest') { 'Guest' } else { 'Member' }
                Mail                      = $userObject.Mail
                AccountEnabled            = $userObject.AccountEnabled
                AccountStatus             = if ($userObject.AccountEnabled) { 'Enabled' } else { 'Disabled' }
                AssignedLicenses          = $licenseNames -join '; '
                AssignedPlans             = $planNames -join '; '
                AssignedLicenseCount      = $licenseNames.Count
                AssignedPlanCount         = $planNames.Count
                OnPremisesSyncEnabled     = $userObject.OnPremisesSyncEnabled
                MobilePhone               = $userObject.MobilePhone
                JobTitle                  = $userObject.JobTitle
                Department                = $userObject.Department
                CompanyName               = $userObject.CompanyName
                OfficeLocation            = $userObject.OfficeLocation
                Manager                   = $managerInfo.ManagerUserPrincipalName
                ManagerDisplayName        = $managerInfo.ManagerDisplayName
                ManagerId                 = $managerInfo.ManagerId
                AuthenticationMethods     = $authInfo.AuthenticationMethods
                AuthenticationMethodCount = $authInfo.AuthenticationMethodCount
            }
        }

        Write-Host ""
        Write-Host "Export User Information to files..."
        Export-InventoryData -Data $userData -BaseName 'UserInformation' -Paths $Paths

        Write-Host "Done. User Information collected: $(@($userData).Count)" -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ UserData = $userData }   # reused by Export-GroupData and the summary
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
