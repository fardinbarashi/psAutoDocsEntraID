# Autodoc v1.0
function Update-EntraServiceLimits {
    <#
        (Re)creates the Entra service-limits config. It starts from a built-in
        baseline (so the file can always be produced, even offline) and then
        tries to refresh each documented number straight from Microsoft's public
        Entra docs on GitHub. Any limit it can confidently parse overrides the
        baseline; anything it can't parse keeps the baseline value.

        Called whenever an export is produced, so the recommendations shipped
        next to the data are always current. Writes the JSON to -OutputPath and
        returns that path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$SourceMarkdown = 'https://raw.githubusercontent.com/MicrosoftDocs/entra-docs/main/docs/includes/entra-service-limits-include.md',
        [string]$SourceUrl = 'https://learn.microsoft.com/en-us/entra/identity/users/directory-service-limits-restrictions',
        [int]$TimeoutSec = 30
    )

    # Built-in baseline. 'parse' is the pattern used to refresh 'limit' from the
    # live doc; metrics without a pattern (e.g. the recommended user ceiling,
    # which lives on a different page) always keep the baseline value.
    $baseline = @(
        [ordered]@{ key = 'DirectoryObjects'; area = 'Directory'; label = 'Directory objects (total)'; limit = 300000; type = 'Quota (verified-domain default)'; note = 'Default quota with at least one verified domain (50,000 without). Raise via Microsoft Support.'; parse = 'extended to ([\d,]+) Microsoft Entra resources' }
        [ordered]@{ key = 'Users'; area = 'Users'; label = 'Users'; limit = 1000000; type = 'Recommended max per tenant'; note = 'Microsoft recommends a single tenant holds no more than ~1,000,000 users (~3,000,000 objects).'; parse = $null }
        [ordered]@{ key = 'ConditionalAccessPolicies'; area = 'Conditional Access'; label = 'Conditional Access policies'; limit = 240; type = 'Hard limit'; note = 'Maximum 240 policies per tenant across all policy states.'; parse = 'maximum of ([\d,]+) policies can be created' }
        [ordered]@{ key = 'ManagedDomains'; area = 'Domains'; label = 'Managed domains'; limit = 5000; type = 'Hard limit'; note = 'No more than 5,000 managed domain names can be added.'; parse = 'no more than ([\d,]+) managed domain' }
        [ordered]@{ key = 'FederatedDomains'; area = 'Domains'; label = 'Federated domains'; limit = 300; type = 'Recommended (max 2,500)'; note = 'Federation with on-premises AD is recommended to stay at or below 300 domains; hard max is 2,500.'; parse = 'limiting to ([\d,]+) domain names' }
        [ordered]@{ key = 'LicenseSubscriptions'; area = 'Licenses'; label = 'License subscriptions'; limit = 300; type = 'Hard limit'; note = 'Limit of 300 license-based subscriptions (such as Microsoft 365 subscriptions) per tenant.'; parse = 'Limit of ([\d,]+) \[?license-based subscriptions' }
        [ordered]@{ key = 'DynamicGroups'; area = 'Groups'; label = 'Dynamic groups'; limit = 15000; type = 'Hard limit'; note = 'Max 15,000 dynamic groups and dynamic administrative units combined per tenant.'; parse = 'maximum of ([\d,]+) dynamic groups' }
        [ordered]@{ key = 'RoleAssignableGroups'; area = 'Groups'; label = 'Role-assignable groups'; limit = 500; type = 'Hard limit'; note = 'Max 500 role-assignable groups can be created in a single tenant.'; parse = 'maximum of ([\d,]+) \[?role-assignable groups' }
    )

    # --- try to refresh from the live doc ---
    $md = $null
    try {
        $md = (Invoke-WebRequest -Uri $SourceMarkdown -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop).Content
    }
    catch {
        Write-Host "  Service limits: live source unavailable, using built-in baseline ($($_.Exception.Message))." -ForegroundColor DarkYellow
    }

    $refreshed = 0
    $metrics = foreach ($m in $baseline) {
        $limit = [int]$m.limit
        if ($md -and $m.parse) {
            $rx = [regex]::Match($md, $m.parse, 'IgnoreCase')
            if ($rx.Success) {
                $parsed = [int]($rx.Groups[1].Value -replace ',', '')
                if ($parsed -gt 0) { $limit = $parsed; $refreshed++ }
            }
        }
        [ordered]@{ key = $m.key; area = $m.area; label = $m.label; limit = $limit; type = $m.type; note = $m.note }
    }

    $out = [ordered]@{
        source         = 'Microsoft Entra service limits and restrictions'
        sourceUrl      = $SourceUrl
        sourceMarkdown = $SourceMarkdown
        retrieved      = (Get-Date -Format 'yyyy-MM-dd')
        limitsSource   = if ($md) { "live ($refreshed of $(@($baseline | Where-Object parse).Count) refreshed from Microsoft docs)" } else { 'built-in baseline (offline)' }
        note           = 'Recreated on each export. Numbers refreshed from the Microsoft Entra docs when reachable; edit here to pin values or adjust to your own governance thresholds.'
        metrics        = @($metrics)
    }

    $dir = Split-Path $OutputPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $out | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Host "  Service limits config written: $OutputPath [$($out.limitsSource)]" -ForegroundColor DarkGreen
    return $OutputPath
}
