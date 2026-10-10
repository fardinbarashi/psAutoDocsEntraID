# Autodoc v1.0
function Export-TenantInformation {
    <#
        Section 1.2.1 - Tenant info + directory object counts + verified domains.
        Requires Organization.Read.All and Directory.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths
    )

    $Section = 'Section 1.2.1 : TenantInformation'
    try {
        Write-Host "Start $Section"

        Write-Host "Get Organization objects..."
        $org = Get-MgOrganization -Property "id,displayName,verifiedDomains,onPremisesSyncEnabled,preferredLanguage,country,countryLetterCode,technicalNotificationMails,privacyProfile"
        if (-not $org) { throw "No organization data returned from Microsoft Graph." }

        Write-Host "Counting directory objects..."
        $null = Get-MgUser        -Top 1 -ConsistencyLevel eventual -CountVariable userCount
        $null = Get-MgGroup       -Top 1 -ConsistencyLevel eventual -CountVariable groupCount
        $null = Get-MgDevice      -Top 1 -ConsistencyLevel eventual -CountVariable deviceCount
        $null = Get-MgApplication -Top 1 -ConsistencyLevel eventual -CountVariable applicationCount

        $primaryDomain = ($org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1).Name
        if (-not $primaryDomain) { $primaryDomain = ($org.VerifiedDomains | Select-Object -First 1).Name }

        $tenantData = [pscustomobject]@{
            TenantId             = $org.Id
            TenantName           = $org.DisplayName
            PrimaryDomain        = $primaryDomain
            CountryCode          = $org.CountryLetterCode
            NotificationLanguage = $org.PreferredLanguage
            TechnicalContact     = ($org.TechnicalNotificationMails -join "; ")
            GlobalPrivacyContact = $org.PrivacyProfile.ContactEmail
            PrivacyStatementUrl  = $org.PrivacyProfile.StatementUrl
            Users                = $userCount
            Groups               = $groupCount
            AppRegistrations     = $applicationCount
            Devices              = $deviceCount
            HybridStatus         = if ($org.OnPremisesSyncEnabled -eq $true) { "Cloud + Onprem" }
                                   elseif ($org.OnPremisesSyncEnabled -eq $false) { "Onprem" }
                                   else { "Cloud" }
        }

        $domainData = $org.VerifiedDomains | ForEach-Object {
            [pscustomobject]@{
                Domain    = $_.Name
                Type      = $_.Type
                IsDefault = $_.IsDefault
            }
        }

        Write-Host "Exporting TenantInformation to files..." -ForegroundColor Yellow
        Export-InventoryData -Data $tenantData -BaseName 'TenantInformation'       -Paths $Paths
        Export-InventoryData -Data $domainData -BaseName 'TenantInformationDomain' -Paths $Paths

        Write-Host "Done. TenantInformation collected: 1 Tenant, $($domainData.Count) Domains." -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ TenantData = $tenantData; DomainData = $domainData }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
