# Autodoc v1.0 - test double for Microsoft.Graph.Authentication used by Test-WorkloadIdentityHunt.ps1
# Test double for Microsoft.Graph.Authentication. Generates Graph-like data, no network.
# $env:MOCK_DATES = 'string' (like Windows PowerShell 5.1) or 'datetime' (like PowerShell 7)
# $env:MOCK_SCENARIO = 'full' | 'empty'
$script:T0       = [datetime]::UtcNow.AddDays(-2)
$script:Epoch    = [datetime]::UtcNow.Date.AddDays(-40)
$script:GoneLeft = 3
$script:Requests = 0

function Connect-MgGraph { [CmdletBinding()] param([string[]]$Scopes, [switch]$NoWelcome, [string]$TenantId) }
function Disconnect-MgGraph { [CmdletBinding()] param() }
function Get-MgContext {
    [pscustomobject]@{
        TenantId = '11111111-2222-3333-4444-555555555555'
        Account  = 'admin@test.onmicrosoft.com'
        Scopes   = @('AuditLog.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Policy.Read.All', 'IdentityRiskyUser.Read.All')
    }
}

function ConvertTo-MockDate([datetime]$d) {
    if ($env:MOCK_DATES -eq 'datetime') { return [datetime]::SpecifyKind($d, 'Utc') }
    return $d.ToString('yyyy-MM-ddTHH:mm:ssZ')
}

function New-SpRec([datetime]$t, [string]$app, [string]$appId, [string]$ip, [int]$code, [string]$res, [string]$resId, [string]$key) {
    [pscustomobject]@{
        id = "sp-$($t.Ticks)-$appId-$code-$res"; createdDateTime = (ConvertTo-MockDate $t)
        appDisplayName = $app; appId = $appId; servicePrincipalId = "spid-$appId"; ipAddress = $ip
        resourceDisplayName = $res; resourceId = $resId
        status = [pscustomobject]@{ errorCode = $code; failureReason = $(if ($code) { "AADSTS$code" } else { $null }) }
        clientCredentialType = 'clientSecret'; servicePrincipalCredentialKeyId = $key; servicePrincipalCredentialThumbprint = $null
        location = [pscustomobject]@{ countryOrRegion = 'US'; city = 'Ashburn' }
    }
}

function New-UserRec([datetime]$t, [string]$upn, [string]$ip, [int]$code, [string]$app = 'Microsoft Azure CLI') {
    [pscustomobject]@{
        id = "u-$($t.Ticks)-$upn-$code"; createdDateTime = (ConvertTo-MockDate $t); userPrincipalName = $upn
        appDisplayName = $app; appId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'; ipAddress = $ip; resourceDisplayName = 'Windows Azure Service Management API'
        status = [pscustomobject]@{ errorCode = $code; failureReason = $(if ($code) { "Error $code" } else { $null }) }
        correlationId = [guid]::NewGuid().ToString(); clientAppUsed = 'Mobile Apps and Desktop clients'
        conditionalAccessStatus = 'notApplied'; mfaDetail = $null; riskLevelDuringSignIn = 'none'
        location = [pscustomobject]@{ countryOrRegion = 'NL'; city = 'Amsterdam' }
    }
}

function Get-SpRecords([datetime]$from, [datetime]$to) {
    $out = New-Object System.Collections.Generic.List[object]
    if ($env:MOCK_SCENARIO -eq 'empty') { return , $out }
    # background: one sign-in every 180 s
    $k = [math]::Ceiling(($from - $script:Epoch).TotalSeconds / 180)
    while ($true) {
        $t = $script:Epoch.AddSeconds($k * 180)
        if ($t -ge $to) { break }
        $n = $k % 7
        [void]$out.Add((New-SpRec $t "App-$n" "app-$n" "10.0.0.$($k % 5)" 0 'Microsoft Graph' '00000003-0000-0000-c000-000000000000' "key-$n"))
        $k++
    }
    $s = $script:T0
    $special = @(
        (New-SpRec $s.AddMinutes(1) 'Backup-Automation' 'app-backup' '199.19.77.84' 7000215 'Azure Key Vault' 'cfa8b339-82a2-471a-a3c9-0fc0be7a4093' ''),
        (New-SpRec $s.AddMinutes(2) 'Backup-Automation' 'app-backup' '199.19.77.84' 7000215 'Azure Key Vault' 'cfa8b339-82a2-471a-a3c9-0fc0be7a4093' ''),
        (New-SpRec $s.AddMinutes(3) 'Backup-Automation' 'app-backup' '199.19.77.84' 0 'Azure Key Vault' 'cfa8b339-82a2-471a-a3c9-0fc0be7a4093' 'key-stolen'),
        (New-SpRec $s.AddMinutes(5) 'Old-App' 'app-old' '192.0.2.99' 7000222 'Microsoft Graph' '00000003-0000-0000-c000-000000000000' ''),
        (New-SpRec $s.AddMinutes(10) 'Sweeper' 'app-sweep' '203.0.113.50' 0 'Windows Azure Service Management API' '797f4846-ba00-4fd7-ba43-dac1f8f63013' 'k1'),
        (New-SpRec $s.AddMinutes(11) 'Sweeper' 'app-sweep' '203.0.113.50' 0 'Azure Key Vault' 'cfa8b339-82a2-471a-a3c9-0fc0be7a4093' 'k1'),
        (New-SpRec $s.AddMinutes(12) 'Sweeper' 'app-sweep' '203.0.113.50' 0 'Azure Storage' 'e406a681-f3d4-42a8-90b6-c2b029497af1' 'k1'),
        (New-SpRec $s.AddMinutes(13) 'Sweeper' 'app-sweep' '203.0.113.50' 0 'Microsoft Graph' '00000003-0000-0000-c000-000000000000' 'k1'),
        (New-SpRec $s.AddMinutes(20) 'Single-Hit' 'app-single' '65.87.223.26' 0 'Microsoft Graph' '00000003-0000-0000-c000-000000000000' 'k9')
    )
    foreach ($r in $special) {
        $rt = [datetime]::Parse(([string]$(if ($r.createdDateTime -is [datetime]) { $r.createdDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { $r.createdDateTime })), [cultureinfo]::InvariantCulture, 'AdjustToUniversal')
        if ($rt -ge $from -and $rt -lt $to) { [void]$out.Add($r) }
    }
    return , $out
}

function Get-UserRecords([string]$upn) {
    $s = $script:T0
    $ioc = '199.19.77.18'
    $all = @(
        (New-UserRec $s.AddMinutes(1) 'alice@test.se' $ioc 50126), (New-UserRec $s.AddMinutes(2) 'alice@test.se' $ioc 50126), (New-UserRec $s.AddMinutes(3) 'alice@test.se' $ioc 50126),
        (New-UserRec $s.AddDays(-5) 'alice@test.se' '10.1.1.1' 0 'Office 365'),
        (New-UserRec $s.AddMinutes(4) 'Bob@Test.se' $ioc 50126), (New-UserRec $s.AddMinutes(5) 'Bob@Test.se' $ioc 50074),
        (New-UserRec $s.AddDays(-6) 'Bob@Test.se' '10.1.1.2' 0 'Office 365'),
        (New-UserRec $s.AddMinutes(90) 'Bob@Test.se' '185.220.101.5' 0 'Office 365'),
        (New-UserRec $s.AddMinutes(6) 'carol@test.se' $ioc 0),
        (New-UserRec $s.AddMinutes(7) 'dave@test.se' $ioc 50053),
        (New-UserRec $s.AddMinutes(8) 'erin@test.se' $ioc 53003),
        (New-UserRec $s.AddMinutes(9) "o'neil@test.se" $ioc 99999),
        (New-UserRec $s.AddDays(-3) 'zed@test.se' '10.1.1.9' 0 'Office 365')
    )
    if ($upn) { return @($all | Where-Object { $_.userPrincipalName -eq $upn }) }
    return $all
}

function Get-Page($items, [string]$uri, [int]$size) {
    $offset = 0
    if ($uri -match 'skiptoken=(\d+)') { $offset = [int]$Matches[1] }
    $tmp = New-Object System.Collections.Generic.List[object]
    foreach ($it in $items) { [void]$tmp.Add($it) }
    $arr = $tmp.ToArray()
    $end = [math]::Min($arr.Count, $offset + $size)
    $page = @()
    if ($end -gt $offset) { $page = [object[]]$arr[$offset..($end - 1)] }
    $resp = [pscustomobject]@{ value = $page }
    if ($end -lt $arr.Count) {
        $base = $uri -replace '&\$skiptoken=\d+', ''
        $resp | Add-Member -NotePropertyName '@odata.nextLink' -NotePropertyValue ($base + '&$skiptoken=' + $end)
    }
    return $resp
}

function Invoke-MgGraphRequest {
    [CmdletBinding()] param([string]$Method, [string]$Uri, [string]$OutputType)
    $script:Requests++
    $u = [uri]::UnescapeDataString($Uri)
    $empty = [pscustomobject]@{ value = @() }
    $isEmpty = ($env:MOCK_SCENARIO -eq 'empty')

    if ($u -match '/organization') {
        return [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'Test: Tenant/AB åäö'; verifiedDomains = @([pscustomobject]@{ name = 'test.onmicrosoft.com'; isDefault = $true }) }) }
    }
    if ($u -match '/auditLogs/signIns') {
        $from = [datetime]::MinValue; $to = [datetime]::MaxValue
        if ($u -match 'createdDateTime ge (\S+)') { $from = [datetime]::Parse($Matches[1], [cultureinfo]::InvariantCulture, 'AdjustToUniversal') }
        if ($u -match 'createdDateTime lt (\S+)') { $to = [datetime]::Parse($Matches[1], [cultureinfo]::InvariantCulture, 'AdjustToUniversal') }
        if ($u -match "'servicePrincipal'") {
            # simulate an expired paging token on page 2 of long slices, three times
            if ($u -match 'skiptoken' -and $script:GoneLeft -gt 0 -and ($to - $from).TotalHours -ge 20 -and $env:MOCK_GONE -eq '1') {
                $script:GoneLeft--
                throw 'Response status code does not indicate success: Gone (Gone).'
            }
            return (Get-Page (Get-SpRecords $from $to) $Uri 300)
        }
        if ($isEmpty) { return $empty }
        if ($u -match "ipAddress eq '([^']+)'") {
            $ip = $Matches[1]
            if ($u -match "'interactiveUser'") { return (Get-Page @(Get-UserRecords $null | Where-Object { $_.ipAddress -eq $ip }) $Uri 5) }
            if ($u -match "'managedIdentity'" -and $ip -eq '65.87.223.26') {
                $mi = New-UserRec $script:T0.AddMinutes(30) $null $ip 0 'mi-app'
                return [pscustomobject]@{ value = @($mi) }
            }
            return $empty
        }
        if ($u -match "userPrincipalName eq '((?:[^']|'')+)'") {
            $upn = $Matches[1].Replace("''", "'")
            if ($u -match "'interactiveUser'") { return (Get-Page @(Get-UserRecords $upn) $Uri 5) }
            return $empty
        }
        return $empty
    }
    if ($u -match '/auditLogs/directoryAudits') {
        if ($isEmpty) { return $empty }
        if ($u -match "'UserManagement'") {
            return [pscustomobject]@{ value = @(
                [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(60)); activityDisplayName = 'User registered security info'; result = 'success'
                    targetResources = @([pscustomobject]@{ userPrincipalName = 'bob@test.se'; displayName = $null; id = 'u2' })
                    initiatedBy = [pscustomobject]@{ user = [pscustomobject]@{ userPrincipalName = 'bob@test.se'; ipAddress = '185.220.101.5' }; app = $null } },
                [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(61)); activityDisplayName = 'Update user'; result = 'success'
                    targetResources = @([pscustomobject]@{ userPrincipalName = 'x@test.se'; displayName = $null; id = 'u3' })
                    initiatedBy = [pscustomobject]@{ user = $null; app = [pscustomobject]@{ displayName = 'Sync' } } }
            ) }
        }
        return [pscustomobject]@{ value = @(
            [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(70)); activityDisplayName = "Update application $([char]0x2013) Certificates and secrets management "; result = 'success'
                targetResources = @([pscustomobject]@{ displayName = 'Backup-Automation'; id = 'obj-backup' })
                initiatedBy = [pscustomobject]@{ user = [pscustomobject]@{ userPrincipalName = 'carol@test.se'; ipAddress = '199.19.77.84' }; app = $null } },
            [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(71)); activityDisplayName = 'Add service principal credentials'; result = 'success'
                targetResources = @([pscustomobject]@{ displayName = 'Sweeper'; id = 'obj-sweep' })
                initiatedBy = [pscustomobject]@{ user = $null; app = [pscustomobject]@{ displayName = 'Rogue App' } } },
            [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(72)); activityDisplayName = 'Add owner to application'; result = 'success'
                targetResources = @([pscustomobject]@{ displayName = 'App-1'; id = 'obj-1' })
                initiatedBy = [pscustomobject]@{ user = [pscustomobject]@{ userPrincipalName = 'Bob@Test.se'; ipAddress = '10.1.1.2' }; app = $null } },
            [pscustomobject]@{ activityDateTime = (ConvertTo-MockDate $script:T0.AddMinutes(73)); activityDisplayName = 'Update application'; result = 'success'
                targetResources = @(); initiatedBy = [pscustomobject]@{ user = $null; app = $null } }
        ) }
    }
    if ($u -match '/identityProtection/riskyUsers') {
        if ($u -match 'carol') { return [pscustomobject]@{ value = @([pscustomobject]@{ riskLevel = 'high'; riskState = 'atRisk'; riskLastUpdatedDateTime = (ConvertTo-MockDate $script:T0) }) } }
        return $empty
    }
    if ($u -match '/reports/servicePrincipalSignInActivities') {
        if ($isEmpty) { return $empty }
        return [pscustomobject]@{ value = @(
            [pscustomobject]@{ appId = 'app-backup'; lastSignInActivity = [pscustomobject]@{ lastSignInDateTime = (ConvertTo-MockDate $script:T0) } },
            [pscustomobject]@{ appId = 'app-old'; lastSignInActivity = [pscustomobject]@{ lastSignInDateTime = (ConvertTo-MockDate $script:T0.AddDays(-200)) } },
            [pscustomobject]@{ appId = 'app-none'; lastSignInActivity = $null }
        ) }
    }
    if ($u -match '/applications/([^/]+)/federatedIdentityCredentials') {
        if ($Matches[1] -eq 'obj-fed') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'fic1'; name = 'github'; issuer = 'https://token.actions.githubusercontent.com'; subject = 'repo:x/y:ref:refs/heads/main' }) } }
        if ($Matches[1] -eq 'obj-err') { throw 'Response status code does not indicate success: Forbidden (Forbidden).' }
        return $empty
    }
    if ($u -match '/applications/([^/]+)/owners') {
        if ($Matches[1] -eq 'obj-backup') { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'o1'; displayName = 'Carol'; userPrincipalName = 'carol@test.se' }, [pscustomobject]@{ id = 'o2'; displayName = 'Owner SP'; userPrincipalName = $null }) } }
        return $empty
    }
    if ($u -match '/v1.0/applications\?') {
        if ($isEmpty) { return $empty }
        $apps = New-Object System.Collections.Generic.List[object]
        [void]$apps.Add([pscustomobject]@{ id = 'obj-backup'; appId = 'app-backup'; displayName = 'Backup-Automation'; signInAudience = 'AzureADMyOrg'
            passwordCredentials = @(
                [pscustomobject]@{ keyId = 'k-new'; displayName = 'added by attacker'; startDateTime = (ConvertTo-MockDate $script:T0); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(700)); hint = 'abc' },
                [pscustomobject]@{ keyId = 'k-exp'; displayName = 'old'; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-900)); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(-200)); hint = 'def' })
            keyCredentials = @([pscustomobject]@{ keyId = 'c1'; displayName = 'CN=cert'; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-100)); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(265)); customKeyIdentifier = 'AABB' }) })
        [void]$apps.Add([pscustomobject]@{ id = 'obj-old'; appId = 'app-old'; displayName = 'Old-App "quoted" <b>'; signInAudience = 'AzureADMultipleOrgs'
            passwordCredentials = @([pscustomobject]@{ keyId = 'k-old'; displayName = $null; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-400)); endDateTime = $null; hint = $null }); keyCredentials = @() })
        [void]$apps.Add([pscustomobject]@{ id = 'obj-fed'; appId = 'app-fed'; displayName = 'Fed-App'; signInAudience = 'AzureADMyOrg'; passwordCredentials = @(); keyCredentials = @() })
        [void]$apps.Add([pscustomobject]@{ id = 'obj-err'; appId = 'app-err'; displayName = 'Err-App'; signInAudience = 'AzureADMyOrg'; passwordCredentials = $null; keyCredentials = $null })
        for ($i = 0; $i -lt 130; $i++) {
            [void]$apps.Add([pscustomobject]@{ id = "obj-$i"; appId = "app-x$i"; displayName = "Bulk-$i"; signInAudience = 'AzureADMyOrg'
                passwordCredentials = @([pscustomobject]@{ keyId = "kb-$i"; displayName = 's'; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-50)); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(300)); hint = 'x' }); keyCredentials = @() })
        }
        return (Get-Page $apps.ToArray() $Uri 50)
    }
    if ($u -match '/v1.0/servicePrincipals\?') {
        if ($isEmpty) { return $empty }
        return [pscustomobject]@{ value = @(
            [pscustomobject]@{ id = 'sp1'; appId = 'app-sweep'; displayName = 'Sweeper'; appOwnerOrganizationId = '99999999-0000-0000-0000-000000000000'; accountEnabled = $true; servicePrincipalType = 'Application'
                passwordCredentials = @([pscustomobject]@{ keyId = 'spk1'; displayName = 'hidden'; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-3)); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(300)) })
                keyCredentials = @([pscustomobject]@{ keyId = 'spc1'; displayName = 'saml'; usage = 'Verify'; startDateTime = (ConvertTo-MockDate $script:T0.AddDays(-3)); endDateTime = (ConvertTo-MockDate $script:T0.AddDays(300)) }, [pscustomobject]@{ keyId = 'spc2'; displayName = 'saml'; usage = 'Sign'; startDateTime = $null; endDateTime = $null }) },
            [pscustomobject]@{ id = 'sp2'; appId = 'app-0'; displayName = 'App-0'; appOwnerOrganizationId = '11111111-2222-3333-4444-555555555555'; accountEnabled = $true; servicePrincipalType = 'Application'; passwordCredentials = @(); keyCredentials = @() }
        ) }
    }
    if ($u -match '/conditionalAccess/namedLocations') {
        if ($isEmpty) { return $empty }
        return [pscustomobject]@{ value = @(
            [pscustomobject]@{ id = 'loc1'; displayName = 'Blocked IoC'; ipRanges = @([pscustomobject]@{ cidrAddress = '199.19.77.84/32' }, [pscustomobject]@{ cidrAddress = '2001:db8::/32' }) },
            [pscustomobject]@{ id = 'loc2'; displayName = 'Countries'; countriesAndRegions = @('SE') }
        ) }
    }
    if ($u -match '/conditionalAccess/policies') {
        if ($isEmpty) { return $empty }
        return [pscustomobject]@{ value = @(
            [pscustomobject]@{ displayName = 'Block IoC users'; state = 'enabled'; grantControls = [pscustomobject]@{ builtInControls = @('block') }
                conditions = [pscustomobject]@{ users = [pscustomobject]@{ includeUsers = @('All') }; clientApplications = $null; locations = [pscustomobject]@{ includeLocations = @('loc1') } } },
            [pscustomobject]@{ displayName = 'MFA all'; state = 'enabled'; grantControls = [pscustomobject]@{ builtInControls = @('mfa') }
                conditions = [pscustomobject]@{ users = [pscustomobject]@{ includeUsers = @('All') }; clientApplications = $null; locations = $null } },
            [pscustomobject]@{ displayName = 'Session only'; state = 'enabledForReportingButNotEnforced'; grantControls = $null
                conditions = [pscustomobject]@{ users = [pscustomobject]@{ includeUsers = @('All') }; clientApplications = $null; locations = $null } }
        ) }
    }
    if ($u -match '/policies/defaultAppManagementPolicy') {
        return [pscustomobject]@{ isEnabled = $false; applicationRestrictions = [pscustomobject]@{ passwordCredentials = @() } }
    }
    throw "MOCK: unhandled URI $u"
}
Export-ModuleMember -Function Connect-MgGraph, Get-MgContext, Disconnect-MgGraph, Invoke-MgGraphRequest
