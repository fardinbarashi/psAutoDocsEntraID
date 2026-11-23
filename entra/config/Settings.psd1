@{
    # Target Microsoft Graph module version
    TargetGraphVersion = '2.33.0'

    # App registration expiry warning limit (days)
    AppRegExpiryLimit = 30

    # "Used within X days" window for app / enterprise-app sign-in activity
    AppUsageDays = 30

    # Modules imported by the script
    RequiredModules = @(
        'ImportExcel'
        'Microsoft.Graph.Identity.DirectoryManagement'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Applications'
        'Microsoft.Graph.Identity.SignIns'
        'Microsoft.Graph.Identity.Governance'
        'Microsoft.Graph.Security'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Reports'
        'Microsoft.Graph.Beta.Reports'
    )

    # Graph scopes requested at connect
    Scopes = @(
        'AccessReview.Read.All'
        'Application.Read.All'
        'AuditLog.Read.All'
        'DelegatedPermissionGrant.Read.All'
        'Device.Read.All'
        'Directory.Read.All'
        'Domain.Read.All'
        'Group.Read.All'
        'Organization.Read.All'
        'Policy.Read.All'
        'Reports.Read.All'
        'RoleManagement.Read.Directory'
        'SecurityEvents.Read.All'
        'User.Read.All'
        'UserAuthenticationMethod.Read.All'
    )

    # Workload identity hunt (Tools tab / -Mode Hunt): its own module
    # list and scopes, so a normal collection never asks for IdentityRiskyUser.
    HuntModules = @(
        'ImportExcel'
        'Microsoft.Graph.Authentication'
    )
    HuntScopes = @(
        'AuditLog.Read.All'
        'Directory.Read.All'
        'Application.Read.All'
        'Policy.Read.All'
        'IdentityRiskyUser.Read.All'
    )
}
