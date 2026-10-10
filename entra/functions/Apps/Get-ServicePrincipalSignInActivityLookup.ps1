# Autodoc v1.0
function Get-ServicePrincipalSignInActivityLookup {
    <#
        Builds an AppId -> sign-in activity lookup from the beta
        servicePrincipalSignInActivities report. Requires AuditLog.Read.All.
        Returns an empty hashtable if the data can't be loaded.
    #>
    [CmdletBinding()]
    param()

    $lookup = @{}
    try {
        $uri = "https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities"
        do {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($item in $resp.value) {
                if ($item.appId) {
                    $lookup[$item.appId] = [pscustomobject]@{
                        AppId                                     = $item.appId
                        LastSignInActivity                        = $item.lastSignInActivity.lastSignInDateTime
                        LastDelegatedClientSignInActivity         = $item.delegatedClientSignInActivity.lastSignInDateTime
                        LastDelegatedResourceSignInActivity       = $item.delegatedResourceSignInActivity.lastSignInDateTime
                        LastApplicationAuthClientSignInActivity   = $item.applicationAuthenticationClientSignInActivity.lastSignInDateTime
                        LastApplicationAuthResourceSignInActivity = $item.applicationAuthenticationResourceSignInActivity.lastSignInDateTime
                    }
                }
            }
            $uri = $resp.'@odata.nextLink'
        } while ($uri)
    }
    catch {
        Write-Host "Could not load servicePrincipalSignInActivities (requires AuditLog.Read.All): $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return $lookup
}
