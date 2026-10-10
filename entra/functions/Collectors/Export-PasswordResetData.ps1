# Autodoc v1.0
function Export-PasswordResetData {
    <#
        Section 1.7.1 - Self-service password reset (SSPR) configuration:
        enabled scope, available methods, security questions and admin SSPR.
        Requires Policy.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths
    )

    $Section = 'Section 1.7.1 : Security Settings - Password Reset'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        Write-Host "Getting authentication methods policy..."
        $authMethodsPolicy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/authenticationMethodsPolicy"
        Write-Host "Getting SSPR / authorization policy settings..."
        $authorizationPolicy = Get-MgPolicyAuthorizationPolicy

        $methodConfigTable = @{}
        foreach ($config in $authMethodsPolicy.authenticationMethodConfigurations) { $methodConfigTable[$config.id] = $config }

        $scopeSourceConfig = $null
        if     ($methodConfigTable.ContainsKey("sms"))   { $scopeSourceConfig = $methodConfigTable["sms"] }
        elseif ($methodConfigTable.ContainsKey("voice")) { $scopeSourceConfig = $methodConfigTable["voice"] }
        elseif ($methodConfigTable.ContainsKey("email")) { $scopeSourceConfig = $methodConfigTable["email"] }
        $scopeInfo = if ($scopeSourceConfig) { Resolve-IncludeTarget -Targets $scopeSourceConfig.includeTargets } else { [pscustomobject]@{ Scope = "Unknown"; GroupId = $null } }

        $enabledMethods = New-Object System.Collections.Generic.List[string]
        if ($methodConfigTable.ContainsKey("microsoftAuthenticator") -and $methodConfigTable["microsoftAuthenticator"].state -eq "enabled") { $enabledMethods.Add("Mobile app notification / Mobile app code") }
        if ($methodConfigTable.ContainsKey("email")             -and $methodConfigTable["email"].state             -eq "enabled") { $enabledMethods.Add("Email") }
        if ($methodConfigTable.ContainsKey("sms")               -and $methodConfigTable["sms"].state               -eq "enabled") { $enabledMethods.Add("Mobile phone") }
        if ($methodConfigTable.ContainsKey("voice")             -and $methodConfigTable["voice"].state             -eq "enabled") { $enabledMethods.Add("Office phone / Voice call") }
        if ($methodConfigTable.ContainsKey("securityQuestions") -and $methodConfigTable["securityQuestions"].state -eq "enabled") { $enabledMethods.Add("Security questions") }

        $securityQuestionsEnabled = $enabledMethods -contains "Security questions"
        $adminSsprEnabled         = $authorizationPolicy.allowedToUseSSPR

        $passwordResetData = [pscustomobject]@{
            SsprEnabledScope              = $scopeInfo.Scope
            SsprSelectedGroupId           = $scopeInfo.GroupId
            MethodsAvailableToUsers       = ($enabledMethods -join "; ")
            SecurityQuestionsEnabled      = $securityQuestionsEnabled
            AdministratorSsprEnabled      = $adminSsprEnabled
            AdministratorMethodsRequired  = if ($adminSsprEnabled -eq $true) { 2 } else { $null }
            AdministratorMethodsAvailable = if ($adminSsprEnabled -eq $true) { "Email; Mobile phone (SMS only); Mobile phone; Office phone; Mobile app code; Mobile app notification" } else { $null }
        }

        Write-Host "Export Password Reset data to files..."
        Export-InventoryData -Data $passwordResetData -BaseName 'PasswordReset' -Paths $Paths

        Write-Host "Done. Password Reset settings collected." -ForegroundColor DarkGreen
        Write-Host "End $Section" -ForegroundColor Green

        return @{ PasswordReset = $passwordResetData }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
