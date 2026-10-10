# Autodoc v1.0
function Get-GroupWelcomeEmailEnabled {
    <#
        For M365 (Unified) groups, reads autoSubscribeNewMembers from the beta
        endpoint. Returns $true/$false, or $null for non-Unified groups / errors.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [AllowNull()][AllowEmptyCollection()][string[]]$GroupTypes = @()
    )

    # Groups with no groupTypes (plain security groups) arrive as $null - treat as non-Unified
    if (-not $GroupTypes -or $GroupTypes -notcontains 'Unified') { return $null }

    try {
        $beta = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/groups/$GroupId" -ErrorAction Stop
        if ($null -ne $beta.autoSubscribeNewMembers) { return [bool]$beta.autoSubscribeNewMembers }
        return $null
    }
    catch { return $null }
}
