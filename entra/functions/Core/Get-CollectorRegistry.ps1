# Autodoc v1.0
function Get-CollectorRegistry {
    <#
        Central manifest of collectable datasets, in run order. One entry per
        dataset, so the picker shows one checkbox per dataset. Each entry has:
          Key         - dataset id / export BaseName, used for selection
          Name        - display name (checkbox label)
          Owner       - collector id; entries sharing an Owner come from one
                        collector run (see Invoke-CollectorRun), so a collector
                        that produces several datasets runs only once even when
                        several of its datasets are selected
          Description - short explanation
          Slow        - $true if it can take a while on large tenants
          Invoke      - scriptblock { param($c) ... } that runs the owning
                        collector, using the run context $c (Paths, Settings,
                        LicenseRef, GraphRef, Collected) and returning its
                        dataset hashtable.
        The registry is the single place to add/remove/reorder datasets.
    #>
    @(
        [pscustomobject]@{
            Key = 'TenantInformation'; Name = 'Tenant information'; Owner = 'Tenant'; Slow = $false
            Description = 'Tenant metadata and object counts.'
            Invoke = { param($c) Export-TenantInformation -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'TenantInformationDomain'; Name = 'Verified domains'; Owner = 'Tenant'; Slow = $false
            Description = 'Verified custom domains and their type (Managed / Federated / None).'
            Invoke = { param($c) Export-TenantInformation -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'LicensesInformation'; Name = 'Licenses'; Owner = 'License'; Slow = $false
            Description = 'Subscribed SKU inventory.'
            Invoke = { param($c) Export-LicenseData -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'UserInformation'; Name = 'User information'; Owner = 'User'; Slow = $true
            Description = 'Users, licenses, manager and MFA methods.'
            Invoke = { param($c) Export-UserData -Paths $c.Paths -LicenseRef $c.LicenseRef -GraphRef $c.GraphRef }
        }
        [pscustomobject]@{
            Key = 'EntraGroups'; Name = 'Groups'; Owner = 'Group'; Slow = $true
            Description = 'Groups with owners, roles and settings.'
            Invoke = { param($c) Export-GroupData -Paths $c.Paths -LicenseRef $c.LicenseRef -GraphRef $c.GraphRef -UserData $c.Collected['UserData'] }
        }
        [pscustomobject]@{
            Key = 'EntraGroupsBasicInfo'; Name = 'Groups (basic info)'; Owner = 'Group'; Slow = $true
            Description = 'Condensed group list (name, type, counts).'
            Invoke = { param($c) Export-GroupData -Paths $c.Paths -LicenseRef $c.LicenseRef -GraphRef $c.GraphRef -UserData $c.Collected['UserData'] }
        }
        [pscustomobject]@{
            Key = 'GroupMembers'; Name = 'Group members'; Owner = 'Group'; Slow = $true
            Description = 'Members per group (with department and office).'
            Invoke = { param($c) Export-GroupData -Paths $c.Paths -LicenseRef $c.LicenseRef -GraphRef $c.GraphRef -UserData $c.Collected['UserData'] }
        }
        [pscustomobject]@{
            Key = 'GroupWelcomeEmail'; Name = 'Group welcome-email'; Owner = 'Group'; Slow = $true
            Description = 'Welcome-email (hidden-membership) status per group.'
            Invoke = { param($c) Export-GroupData -Paths $c.Paths -LicenseRef $c.LicenseRef -GraphRef $c.GraphRef -UserData $c.Collected['UserData'] }
        }
        [pscustomobject]@{
            Key = 'AppRegEnterpriseApps'; Name = 'App registrations'; Owner = 'AppReg'; Slow = $true
            Description = 'App registrations: owners, credentials, API permissions and sign-in activity.'
            Invoke = { param($c) Export-AppRegistrationData -Paths $c.Paths -UsageDays $c.Settings.AppUsageDays }
        }
        [pscustomobject]@{
            Key = 'EnterpriseApps'; Name = 'Enterprise apps'; Owner = 'EnterpriseApp'; Slow = $true
            Description = 'All service principals: type, origin, owners, SSO, credentials and sign-in activity.'
            Invoke = { param($c) Export-EnterpriseAppData -Paths $c.Paths -UsageDays $c.Settings.AppUsageDays }
        }
        [pscustomobject]@{
            Key = 'ConditionalAccess'; Name = 'Conditional Access'; Owner = 'CA'; Slow = $false
            Description = 'CA policies, named locations and authentication strengths.'
            Invoke = { param($c) Export-ConditionalAccessData -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'RBAC'; Name = 'RBAC / PIM'; Owner = 'RBAC'; Slow = $false
            Description = 'Active and eligible role assignments plus role-assignable groups.'
            Invoke = { param($c) Export-RbacData -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'PasswordReset'; Name = 'Password reset (SSPR)'; Owner = 'SSPR'; Slow = $false
            Description = 'Self-service password reset configuration.'
            Invoke = { param($c) Export-PasswordResetData -Paths $c.Paths }
        }
        [pscustomobject]@{
            Key = 'Summary'; Name = 'Consolidated summary'; Owner = 'Summary'; Slow = $false
            Description = 'TenantSummary workbook + index. For creating excel, visio and doc files.'
            Invoke = { param($c) Export-TenantSummary -Paths $c.Paths -Collected $c.Collected; $null }
        }
    )
}
