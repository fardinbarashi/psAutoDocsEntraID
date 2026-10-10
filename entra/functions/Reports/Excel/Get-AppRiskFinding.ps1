# Autodoc v1.0
function Get-AppRiskFinding {
    <#
        Scores one application against the risk signals we actually collect and
        returns @{ Risk = 'High'|'Medium'|''; Reason = '...' }.

        WHAT THIS CANNOT SEE. The layout this table is modelled on lists
        "SHA-1 certificate" and "old redirect URIs" as reasons. Neither is
        collected: the app export keeps credential *dates* but not the signing
        algorithm, and reply URLs aren't gathered at all. Rather than invent
        those findings, the rules below use signals that can be proven from the
        exported data. Adding the other two means extending
        Export-AppRegistrationData with KeyCredentials' algorithm and the
        application's replyUrls.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$App,
        [int]$InactiveDays = 180
    )

    $reasons = @()
    $high    = $false

    if ("$($App.CredentialExpired)" -eq 'Yes') {
        $reasons += 'Expired credentials'; $high = $true
    }

    $ownerless = ("$($App.OwnerInfo)" -eq 'MissingOwners' -or -not $App.OwnerInfo)
    if ($ownerless -and "$($App.UsesGraphPermissions)" -eq 'Yes') {
        $reasons += 'No owner and holds Graph permissions'; $high = $true
    }
    elseif ($ownerless) {
        $reasons += 'No owner'
    }

    # Inactivity only counts as a finding when the app can still authenticate
    $hasCreds = ("$($App.HasSecrets)" -eq 'Yes' -or "$($App.HasCertificates)" -eq 'Yes')
    $lastSeen = ConvertTo-SafeDateTime -Value $App.LastSignInDateTime
    if ($lastSeen) {
        $days = [int]((Get-Date) - $lastSeen).TotalDays
        if ($days -gt $InactiveDays) {
            $reasons += "Inactive > $InactiveDays days"
            if ($hasCreds) { $high = $true }
        }
    }
    elseif ($hasCreds) {
        $reasons += 'Never signed in but has credentials'; $high = $true
    }

    if ("$($App.CredentialExpiresWithin30Days)" -eq 'Yes') { $reasons += 'Credentials expire within 30 days' }
    if ("$($App.PublicClient)" -eq 'Yes')                  { $reasons += 'Public client (legacy auth flows)' }
    if ("$($App.SignInAudience)" -and "$($App.SignInAudience)" -ne 'AzureADMyOrg') { $reasons += 'Multi-tenant exposure' }

    $permCount = (Get-AppPermissionInfo -App $App).Count
    if ($permCount -gt 20) { $reasons += "$permCount API permissions requested" }

    if (-not $reasons) { return @{ Risk = ''; Reason = '' } }
    @{
        Risk   = $(if ($high) { 'High' } else { 'Medium' })
        Reason = ($reasons -join '; ')
    }
}

function Get-SignInRecencyBand {
    <# Four-band sign-in recency, matching the 7 / 8-30 / 31-90 / 90+ split. #>
    [CmdletBinding()]
    param($LastSignIn)

    $d = ConvertTo-SafeDateTime -Value $LastSignIn
    if (-not $d) { return 'Never / no data' }
    $days = ((Get-Date) - $d).TotalDays
    if     ($days -le 7)  { 'Last 7 days' }
    elseif ($days -le 30) { '8-30 days' }
    elseif ($days -le 90) { '31-90 days' }
    else                  { '90+ days' }
}
