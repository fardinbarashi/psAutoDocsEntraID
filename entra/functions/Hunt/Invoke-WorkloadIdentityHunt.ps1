# Autodoc v1.0
function Invoke-WorkloadIdentityHunt {
    <#
        Workload identity hunt. Looks for signs of a stolen or abused app
        credential, triages every user account that was hit from the IoC
        addresses, and inventories every client secret, certificate and
        federated credential in the tenant.

        READ-ONLY: only GET requests are sent to Microsoft Graph. The caller owns
        the Graph connection (Autodoc connects before the collectors run), so
        this function neither connects nor disconnects.

        Steps:
           1. Service principal sign-ins, matched against the IoC IPs
              (HighConfidenceIPs = high confidence, the rest = weaker link)
           2. Interactive, non-interactive and managed identity sign-ins from the IoC IPs
           3. Per-account triage for every user hit in step 2 (plus -TriageUsers):
                - all interactive + non-interactive sign-ins in the window
                - error code classification:
                    50126                          wrong password      -> spray, password not confirmed
                    50074/50076/500121/50079/...   MFA stage           -> PASSWORD WAS CORRECT
                    53003 / 53000-53002            blocked by CA       -> PASSWORD WAS CORRECT
                    50053                          account locked      -> active attack
                    0                              success             -> COMPROMISED if from IoC IP
                - verdict: COMPROMISED / PASSWORD_KNOWN / ACTIVE_ATTACK / SPRAY / NO_CLEAR_SIGNAL
                - successful sign-ins from NEW IPs around the attack time
                - security info (MFA method) changes on the account
                - Identity Protection risk level
           4. AADSTS7000215 (invalid secret) / AADSTS7000222 (expired secret) from unknown IPs
           5. Failure followed by success from the same IP for the same app -> CRITICAL
           6. One IP hitting >= SpThreshold different service principals
           7. One app getting tokens for Resource Manager, Key Vault, Storage AND Graph in a short window
           8. Inventory of all secrets / certificates / federated credentials, owners, last sign-in
           9. Audit log: credential additions / owner changes on apps and service principals
          10. Posture: CA block for the IoC IP, app management policy blocking new secrets

        Output (all under -OutputFolder):
          - one CSV per data set, FINDINGS.csv sorted by severity (the raw sign-in
            rows only with -SaveRawSignIns)
          - a self-contained HTML report (New-HuntHtmlReport) unless -NoHtml

        Sign-in logs are read in time slices (-SliceHours). If Graph expires the
        paging token ("410 Gone"), the slice is read again or split in two.

        Settings normally come from Entra\Config\WorkloadIdentityHunt.json
        through Export-WorkloadIdentityHunt.

        Required Graph scopes:
          AuditLog.Read.All, Directory.Read.All, Application.Read.All, Policy.Read.All,
          IdentityRiskyUser.Read.All

        KQL for longer retention (Log Analytics / Sentinel):
          let iocs = dynamic(["199.19.77.84","199.19.77.18","65.87.223.26"]);
          union AADServicePrincipalSignInLogs, AADNonInteractiveUserSignInLogs, SigninLogs, AADManagedIdentitySignInLogs
          | where TimeGenerated > ago(365d)
          | where IPAddress in (iocs) or ResultType in ("7000215","7000222")
          | project TimeGenerated, Type, AppId, AppDisplayName, ServicePrincipalName, UserPrincipalName,
                    IPAddress, ResultType, ResourceDisplayName, ServicePrincipalCredentialKeyId
          | order by TimeGenerated asc

        Returns a hashtable: Findings, Triage, TriageSignIns, Credentials, Audit,
        IocSpSignIns, IocUserSignIns, SecretFailures, SpSignIns (only with
        -SaveRawSignIns) and ReportPath (the HTML report, or $null).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputFolder,
        [int]$DaysBack = 30,
        [string[]]$IocIPs = @('199.19.77.84', '199.19.77.18', '65.87.223.26'),
        [string[]]$HighConfidenceIPs = @('199.19.77.84'),
        # Own known egress IPs / CIDRs (excluded from the "unknown IP" checks)
        [string[]]$KnownIPs = @(),
        # Extra accounts (UPNs) to triage in step 3 even without an IoC hit
        [string[]]$TriageUsers = @(),
        [int]$SpThreshold = 3,
        [int]$TokenWindowMinutes = 60,
        [int]$MaxSignIns = 250000,
        [int]$MaxUserSignIns = 20000,
        [int]$NewCredentialDays = 30,
        [int]$UnusedDays = 90,
        # Time slice (hours) used when reading the sign-in logs
        [int]$SliceHours = 24,
        [switch]$SkipInventory,
        [switch]$SkipFederatedCredentials,
        # Hide DEBUG lines (Graph paging, per-app details)
        [switch]$Quiet,
        # Folder for the HTML report. Default: -OutputFolder
        [string]$ReportFolder,
        # Max rows per table in the HTML report (all rows are always in the CSV files)
        [int]$HtmlMaxRows = 20000,
        [switch]$NoHtml,
        # Also write the raw sign-in rows (all service principal sign-ins and every
        # sign-in of the triaged accounts) to CSV. Off by default: large and personal.
        [switch]$SaveRawSignIns
    )

    $ErrorActionPreference = 'Stop'
    $runStamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    if (-not $ReportFolder) { $ReportFolder = $OutputFolder }
    if (@($HighConfidenceIPs).Count -eq 0) { $HighConfidenceIPs = @($IocIPs | Select-Object -First 1) }

    # ---------------------------------------------------------------- logging and shared state
    $state = @{ RunStart = (Get-Date); StepStart = (Get-Date); StepNo = 0; StepTotal = 10; StepFinds = 0 }

    $LevelColor = @{
        INFO = 'White'; OK = 'Green'; WARN = 'Yellow'; ERROR = 'Red'
        DEBUG = 'DarkGray'; STEP = 'Cyan'; FIND = 'White'
    }
    $SevColor = @{ CRITICAL = 'Red'; HIGH = 'Magenta'; MEDIUM = 'Yellow'; INFO = 'Gray' }

    # Used instead of "$x = if (...) {...}" expressions (avoids a Windows PowerShell 5.1
    # compiler bug that throws "Argument types do not match")
    function Select-HuntValue {
        param($Condition, $IfTrue, $IfFalse)
        if ($Condition) { return $IfTrue }
        return $IfFalse
    }

    function Write-HuntLog {
        param([string]$Message, [string]$Level = 'INFO', [string]$Color)
        if ($Level -eq 'DEBUG' -and $Quiet) { return }
        if (-not $Color) {
            $Color = 'White'
            if ($LevelColor.ContainsKey($Level)) { $Color = $LevelColor[$Level] }
        }
        $ts   = (Get-Date).ToString('HH:mm:ss')
        $line = "[{0}] {1,-5} {2}" -f $ts, $Level, $Message
        Write-Host $line -ForegroundColor $Color
    }

    function Start-HuntStep([string]$Title) {
        $state.StepNo = $state.StepNo + 1
        $state.StepStart = Get-Date
        $state.StepFinds = 0
        Write-Host ''
        Write-HuntLog ("===== [{0}/{1}] {2} =====" -f $state.StepNo, $state.StepTotal, $Title) -Level STEP
    }

    function Stop-HuntStep([string]$Summary) {
        $sec = [math]::Round(((Get-Date) - $state.StepStart).TotalSeconds, 1)
        $lvl = 'OK'
        if ($state.StepFinds -gt 0) { $lvl = 'WARN' }
        Write-HuntLog ("{0} | findings: {1} | {2} s" -f $Summary, $state.StepFinds, $sec) -Level $lvl
    }

    # Timestamps in messages: always yyyy-MM-dd HH:mm:ss (string interpolation would give MM/dd/yyyy)
    function Format-HuntTime($Value) {
        if ($null -eq $Value) { return '' }
        if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
        return [string]$Value
    }

    function Format-HuntDuration([timespan]$t) { '{0:00}:{1:00}:{2:00}' -f [math]::Floor($t.TotalHours), $t.Minutes, $t.Seconds }

    $findings = New-Object System.Collections.Generic.List[object]

    function Add-HuntFinding {
        param(
            [string]$Severity, [string]$Category, [string]$Detail,
            [string]$AppName, [string]$AppId, [string]$IP, $Time
        )
        $obj = New-Object PSObject -Property ([ordered]@{
            Severity = $Severity; Category = $Category; AppName = $AppName; AppId = $AppId
            IP = $IP; Time = $Time; Detail = $Detail
        })
        [void]$findings.Add($obj)
        $state.StepFinds = $state.StepFinds + 1
        $c = 'Gray'
        if ($SevColor.ContainsKey($Severity)) { $c = $SevColor[$Severity] }
        $who = (@($AppName, $IP) | Where-Object { $_ }) -join ' @ '
        $suffix = ''
        if ($who) { $suffix = " | $who" }
        $lineLevel = 'FIND'
        if ($Severity -eq 'INFO') { $lineLevel = 'DEBUG' }
        Write-HuntLog ("  >> {0,-8} {1}{2}" -f $Severity, $Category, $suffix) -Level $lineLevel -Color $c
        if ($Severity -eq 'CRITICAL' -or $Severity -eq 'HIGH') { Write-HuntLog "     $Detail" -Level FIND -Color $c }
    }

    function Invoke-HuntGraphGetAll {
        param([Parameter(Mandatory)][string]$Uri, [int]$Max = 0, [string]$Label = 'Graph')
        $out  = New-Object System.Collections.Generic.List[object]
        $next = $Uri
        $page = 0
        $restarts = 0
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        Write-HuntLog "$Label : GET $([uri]::UnescapeDataString($Uri))" -Level DEBUG
        while ($next) {
            $attempt = 0
            $resp = $null
            $restart = $false
            while ($true) {
                try {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
                    break
                } catch {
                    $attempt++
                    $msg = $_.Exception.Message
                    if ($msg -match 'Gone|Skip token') {
                        # Paging token expired. Read again from the first page once, then let the
                        # caller decide (Get-HuntSignInSlice splits the time slice).
                        if ($restarts -lt 1 -and $page -gt 0) {
                            $restarts++
                            Write-HuntLog "$Label : paging token expired after page $page - reading again from the first page" -Level WARN
                            $restart = $true
                            break
                        }
                        Write-HuntLog "$Label : paging token expired after page $page (Gone)" -Level WARN
                        throw
                    }
                    if ($attempt -le 6 -and $msg -match '429|TooManyRequests|503|504|ServiceUnavailable|GatewayTimeout|timed out') {
                        $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                        Write-HuntLog "$Label : throttled/transient error (attempt $attempt/6), waiting $wait s" -Level WARN
                        Start-Sleep -Seconds $wait
                    } else {
                        Write-HuntLog "$Label : request failed: $msg" -Level ERROR
                        throw
                    }
                }
            }
            if ($restart) {
                $out.Clear()
                $next = $Uri
                $page = 0
                continue
            }
            $page++
            $n = 0
            if ($resp.value) { foreach ($v in $resp.value) { [void]$out.Add($v); $n++ } }
            Write-HuntLog ("{0} : page {1} +{2} (total {3}, {4:N1} s)" -f $Label, $page, $n, $out.Count, $sw.Elapsed.TotalSeconds) -Level DEBUG
            if ($Max -gt 0 -and $out.Count -ge $Max) {
                Write-HuntLog "$Label : stopped after $Max records." -Level WARN
                break
            }
            $next = $resp.'@odata.nextLink'
        }
        return , $out
    }

    function New-HuntFilterUri([string]$Base, [string]$Filter, [string]$Extra = '') {
        return $Base + '?$filter=' + [uri]::EscapeDataString($Filter) + $Extra
    }

    # Reads sign-in logs in time slices. The sign-in API expires its paging token if paging takes
    # too long ("410 Gone - Skip token has expired"). When that happens the slice is read again;
    # if it fails again it is split in two smaller slices. Records are de-duplicated on id.
    function Get-HuntSignInSlice {
        param([string]$Filter, [datetime]$From, [datetime]$To, [string]$Label, [int]$Max = 0, [int]$Depth = 0)
        $f = $From.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $t = $To.ToString('yyyy-MM-ddTHH:mm:ssZ')
        $full = "createdDateTime ge $f and createdDateTime lt $t"
        if ($Filter) { $full = "$full and $Filter" }
        $uri = New-HuntFilterUri 'https://graph.microsoft.com/beta/auditLogs/signIns' $full '&$top=1000'
        $sliceLabel = "$Label [$($From.ToString('MM-dd HH:mm'))-$($To.ToString('MM-dd HH:mm'))]"
        $attempt = 0
        while ($true) {
            $attempt++
            try {
                $result = Invoke-HuntGraphGetAll -Uri $uri -Max $Max -Label $sliceLabel
                return , $result
            } catch {
                $msg = $_.Exception.Message
                if ($msg -notmatch 'Gone|Skip token') { throw }
                $minutes = ($To - $From).TotalMinutes
                if ($minutes -gt 30 -and $Depth -lt 8) {
                    $mid = $From.AddMinutes([math]::Floor($minutes / 2))
                    Write-HuntLog "$Label : splitting slice $f -> $t in two" -Level WARN
                    $out = New-Object System.Collections.Generic.List[object]
                    $firstHalf = Get-HuntSignInSlice -Filter $Filter -From $From -To $mid -Label $Label -Max $Max -Depth ($Depth + 1)
                    foreach ($x in $firstHalf) { [void]$out.Add($x) }
                    $secondHalf = Get-HuntSignInSlice -Filter $Filter -From $mid -To $To -Label $Label -Max $Max -Depth ($Depth + 1)
                    foreach ($x in $secondHalf) { [void]$out.Add($x) }
                    return , $out
                }
                if ($attempt -lt 3) {
                    Write-HuntLog "$Label : slice $f -> $t still failing, waiting 10 s (attempt $attempt/3)" -Level WARN
                    Start-Sleep -Seconds 10
                    continue
                }
                throw
            }
        }
    }

    function Invoke-HuntSignInQuery {
        param(
            [string]$Filter,
            [Parameter(Mandatory)][datetime]$From,
            [Parameter(Mandatory)][datetime]$To,
            [string]$Label = 'Sign-ins',
            [int]$Max = 0,
            [int]$Hours = 24
        )
        $out  = New-Object System.Collections.Generic.List[object]
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        if ($Hours -lt 1) { $Hours = 1 }
        $slices = [math]::Max(1, [math]::Ceiling(($To - $From).TotalHours / $Hours))
        $start = $From
        $n = 0
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($start -lt $To) {
            $n++
            $end = $start.AddHours($Hours)
            if ($end -gt $To) { $end = $To }
            $part = Get-HuntSignInSlice -Filter $Filter -From $start -To $end -Label $Label -Max $Max
            $added = 0
            foreach ($x in $part) {
                if ($x.id -and -not $seen.Add([string]$x.id)) { continue }
                [void]$out.Add($x); $added++
            }
            if ($slices -gt 1) {
                Write-HuntLog ("{0} : slice {1}/{2} {3} -> {4}: +{5} (total {6}, {7:N0} s)" -f $Label, $n, $slices,
                    $start.ToString('yyyy-MM-dd HH:mm'), $end.ToString('yyyy-MM-dd HH:mm'), $added, $out.Count, $sw.Elapsed.TotalSeconds)
            }
            if ($Max -gt 0 -and $out.Count -ge $Max) {
                Write-HuntLog "$Label : stopped after $Max records. Narrow -DaysBack or raise the limit." -Level WARN
                break
            }
            $start = $end
        }
        return , $out
    }

    function Test-HuntIpInRange([string]$Ip, [string]$Range) {
        if ([string]::IsNullOrWhiteSpace($Ip)) { return $false }
        if ($Range -notmatch '/') { return ($Ip -eq $Range) }
        $parts  = $Range.Split('/')
        $net    = [System.Net.IPAddress]::Parse($parts[0])
        $prefix = [int]$parts[1]
        $addr   = $null
        if (-not [System.Net.IPAddress]::TryParse($Ip, [ref]$addr)) { return $false }
        if ($addr.AddressFamily -ne $net.AddressFamily) { return $false }
        $a = $addr.GetAddressBytes(); $n = $net.GetAddressBytes()
        $full = [int][math]::Floor($prefix / 8); $rem = $prefix % 8
        for ($i = 0; $i -lt $full; $i++) { if ($a[$i] -ne $n[$i]) { return $false } }
        if ($rem -gt 0) {
            $mask = [int]((0xFF -shl (8 - $rem)) -band 0xFF)
            if (([int]$a[$full] -band $mask) -ne ([int]$n[$full] -band $mask)) { return $false }
        }
        return $true
    }

    function Test-HuntKnownIp([string]$Ip) {
        foreach ($r in $KnownIPs) { if (Test-HuntIpInRange $Ip $r) { return $true } }
        return $false
    }

    # Graph returns timestamps as ISO strings (PowerShell 5.1) or as UTC DateTime (PowerShell 7).
    # Always returns local time, or $null when the value is missing (never throws).
    function ConvertTo-HuntLocalDate($Value) {
        if ($null -eq $Value) { return $null }
        if ($Value -is [datetime]) {
            if ($Value.Kind -eq [System.DateTimeKind]::Utc) { return $Value.ToLocalTime() }
            return $Value
        }
        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $parsed = [datetime]::MinValue
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $style = [System.Globalization.DateTimeStyles]::AssumeUniversal
        if ([datetime]::TryParse($text, $culture, $style, [ref]$parsed)) { return $parsed }
        return $null
    }

    # User sign-in error codes -> category | meaning
    #   SUCCESS  = signed in
    #   PWD_OK   = failed AFTER the password was validated (MFA stage etc.) -> password is correct
    #   CA_BLOCK = blocked by Conditional Access (evaluated after primary auth) -> password is correct
    #   LOCKED   = smart lockout -> active attack
    #   SPRAY    = wrong password
    #   OTHER    = anything else
    $UserErrMap = @{
        0      = 'SUCCESS|Successful sign-in'
        50126  = 'SPRAY|Wrong username or password'
        50053  = 'LOCKED|Account locked (smart lockout) or sign-in from malicious IP blocked'
        50074  = 'PWD_OK|Strong authentication (MFA) required - password was correct'
        50076  = 'PWD_OK|MFA required (location/policy) - password was correct'
        500121 = 'PWD_OK|MFA failed or denied - password was correct'
        50079  = 'PWD_OK|MFA registration required - password was correct'
        50158  = 'PWD_OK|External security challenge not satisfied - password was correct'
        50055  = 'PWD_OK|Password expired - password was correct'
        50140  = 'PWD_OK|Keep me signed in interrupt - password was correct'
        50097  = 'PWD_OK|Device authentication required - password was correct'
        53003  = 'CA_BLOCK|Blocked by Conditional Access - password was correct'
        53000  = 'CA_BLOCK|CA: device not compliant - password was correct'
        53001  = 'CA_BLOCK|CA: device not domain joined - password was correct'
        53002  = 'CA_BLOCK|CA: app not approved - password was correct'
        50057  = 'OTHER|Account disabled'
        50034  = 'OTHER|User does not exist'
        50089  = 'OTHER|Flow token expired'
        50173  = 'OTHER|Grant expired (password changed/revoked)'
        70043  = 'OTHER|Refresh token expired (CA sign-in frequency)'
        700082 = 'OTHER|Refresh token expired (inactivity)'
    }
    $CatColor = @{ SUCCESS = 'Red'; PWD_OK = 'Red'; CA_BLOCK = 'Red'; LOCKED = 'Magenta'; SPRAY = 'Yellow'; OTHER = 'Gray' }

    function Get-HuntUserErrInfo([int]$Code, [string]$Reason) {
        if ($UserErrMap.ContainsKey($Code)) {
            $p = $UserErrMap[$Code].Split('|', 2)
            return New-Object PSObject -Property @{ Cat = $p[0]; Text = $p[1] }
        }
        return New-Object PSObject -Property @{ Cat = 'OTHER'; Text = "Other: $Reason" }
    }

    # Resource categories for the "token sweep" check
    $resById = @{
        '797f4846-ba00-4fd7-ba43-dac1f8f63013' = 'ARM'
        'cfa8b339-82a2-471a-a3c9-0fc0be7a4093' = 'KeyVault'
        'e406a681-f3d4-42a8-90b6-c2b029497af1' = 'Storage'
        '00000003-0000-0000-c000-000000000000' = 'Graph'
    }
    $resByName = @{
        'Windows Azure Service Management API' = 'ARM'
        'Azure Resource Manager'               = 'ARM'
        'Azure Key Vault'                      = 'KeyVault'
        'Azure Storage'                        = 'Storage'
        'Microsoft Graph'                      = 'Graph'
    }
    function Get-HuntResCategory($row) {
        if ($row.ResourceId -and $resById.ContainsKey([string]$row.ResourceId)) { return $resById[[string]$row.ResourceId] }
        if ($row.Resource -and $resByName.ContainsKey([string]$row.Resource))   { return $resByName[[string]$row.Resource] }
        return $null
    }

    function Export-HuntCsv($Data, [string]$Name) {
        $path = Join-Path $OutputFolder $Name
        # .ToArray() for generic lists: @(<List>) can throw "Argument types do not match" in PowerShell 5.1
        $rows = @()
        if ($Data -is [System.Collections.Generic.List[object]]) { $rows = $Data.ToArray() }
        elseif ($null -ne $Data) { $rows = @($Data | ForEach-Object { $_ }) }
        if ($rows.Count -gt 0) {
            # Dates are written as yyyy-MM-dd HH:mm:ss regardless of the regional settings
            $outRows = New-Object System.Collections.Generic.List[object]
            foreach ($row in $rows) {
                $copy = [ordered]@{}
                foreach ($prop in $row.PSObject.Properties) {
                    if ($prop.Value -is [datetime]) { $copy[$prop.Name] = $prop.Value.ToString('yyyy-MM-dd HH:mm:ss') }
                    else { $copy[$prop.Name] = $prop.Value }
                }
                [void]$outRows.Add((New-Object PSObject -Property $copy))
            }
            $outRows.ToArray() | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
        }
        else { Set-Content -Path $path -Value '' -Encoding UTF8 }
        Write-HuntLog "Wrote $Name ($($rows.Count) rows)" -Level DEBUG
    }
    $now      = (Get-Date).ToUniversalTime()
    $nowLocal = $now.ToLocalTime()
    $sinceDt  = $now.AddDays(-$DaysBack)
    $since    = $sinceDt.ToString('yyyy-MM-ddTHH:mm:ssZ')

    # ================================================================ MAIN
    # Collections the steps fill; declared up front so a failure part-way still returns what was found
    $triageSummary = New-Object System.Collections.Generic.List[object]
    $triageSignins = New-Object System.Collections.Generic.List[object]
    $credRows = @()
    $audit    = @()
    $sorted   = @()
    $spRows   = @()
    $iocSp    = @()
    $iocOther = @()
    $secretFails = @()
    $htmlPath = $null
    try {

    Write-HuntLog "Workload identity hunt started (READ-ONLY)" -Level STEP
    Write-HuntLog "PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition)) on $([Environment]::MachineName) as $([Environment]::UserName)"
    Write-HuntLog "Parameters: DaysBack=$DaysBack, SpThreshold=$SpThreshold, TokenWindow=${TokenWindowMinutes}min, MaxSignIns=$MaxSignIns, MaxUserSignIns=$MaxUserSignIns"
    Write-HuntLog "IoC IPs: $($IocIPs -join ', ')  (high confidence: $($HighConfidenceIPs -join ', '))"
    if ($KnownIPs.Count -gt 0) { Write-HuntLog "Known/excluded IPs: $($KnownIPs -join ', ')" }
    else { Write-HuntLog "No KnownIPs given - own egress IPs may show up as findings in the IP checks" -Level WARN }
    if ($TriageUsers.Count -gt 0) { Write-HuntLog "Extra accounts to triage: $($TriageUsers -join ', ')" }
    Write-HuntLog "Output folder: $OutputFolder"

    $ctx = Get-MgContext
    if ($null -eq $ctx) { throw "Not connected to Microsoft Graph. Connect first (Autodoc does this before the collectors run)." }
    Write-HuntLog "Using Graph connection: tenant $($ctx.TenantId) as $($ctx.Account)" -Level OK
    $tenantName   = [string]$ctx.TenantId
    $tenantDomain = ''
    try {
        $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains' -OutputType PSObject
        $orgObj = $org.value | Select-Object -First 1
        if ($orgObj.displayName) { $tenantName = [string]$orgObj.displayName }
        $def = @($orgObj.verifiedDomains | Where-Object { $_.isDefault })
        if ($def.Count -gt 0) { $tenantDomain = [string]$def[0].name }
        Write-HuntLog "Tenant name: $tenantName ($tenantDomain)" -Level OK
    } catch { Write-HuntLog "Could not read tenant name, using tenant id: $($_.Exception.Message)" -Level WARN }
    $scopes  = 'AuditLog.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Policy.Read.All', 'IdentityRiskyUser.Read.All'
    $missing = @($scopes | Where-Object { @($ctx.Scopes) -notcontains $_ })
    if ($ctx.Scopes -and $missing.Count -gt 0) { Write-HuntLog "Token is missing scopes: $($missing -join ', ') - some checks may fail" -Level WARN }

    $now      = (Get-Date).ToUniversalTime()
    $nowLocal = $now.ToLocalTime()
    $sinceDt  = $now.AddDays(-$DaysBack)
    $since    = $sinceDt.ToString('yyyy-MM-ddTHH:mm:ssZ')
    Write-HuntLog "Time window: $since -> $($now.ToString('yyyy-MM-ddTHH:mm:ssZ')) (UTC)"

    Write-HuntLog "Time window: $since -> $($now.ToString('yyyy-MM-ddTHH:mm:ssZ')) (UTC)"

    # ---------------------------------------------------------------- 1. service principal sign-ins
    Start-HuntStep "Service principal sign-ins (incl. IoC IP match)"
    Write-HuntLog "Fetching all service principal sign-ins since $since - this can take a while in large tenants ..."
    Write-HuntLog "Reading in slices of $SliceHours h (use -SliceHours to change)"
    $spRaw = Invoke-HuntSignInQuery -Filter "signInEventTypes/any(t: t eq 'servicePrincipal')" `
        -From $sinceDt -To $now -Hours $SliceHours -Max $MaxSignIns -Label 'SP sign-ins'

    Write-HuntLog "Normalising $($spRaw.Count) records ..."
    $spList = New-Object System.Collections.Generic.List[object]
    $oldest = $null
    $newest = $null
    $noTime = 0
    $seenInLogs = @{}   # appId -> latest successful sign-in found in the log
    foreach ($s in $spRaw) {
        $when = ConvertTo-HuntLocalDate $s.createdDateTime
        if ($null -eq $when) { $noTime++; continue }
        if ($null -eq $oldest -or $when -lt $oldest) { $oldest = $when }
        if ($null -eq $newest -or $when -gt $newest) { $newest = $when }
        $appKey = [string]$s.appId
        if ($appKey -and ([int]$s.status.errorCode -eq 0)) {
            if (-not $seenInLogs.ContainsKey($appKey) -or $when -gt $seenInLogs[$appKey]) { $seenInLogs[$appKey] = $when }
        }
        [void]$spList.Add((New-Object PSObject -Property ([ordered]@{
            Time           = $when
            AppName        = [string]$s.appDisplayName
            AppId          = [string]$s.appId
            SpId           = [string]$s.servicePrincipalId
            IP             = [string]$s.ipAddress
            Resource       = [string]$s.resourceDisplayName
            ResourceId     = [string]$s.resourceId
            ErrorCode      = [int]$s.status.errorCode
            FailureReason  = [string]$s.status.failureReason
            CredType       = [string]$s.clientCredentialType
            CredKeyId      = [string]$s.servicePrincipalCredentialKeyId
            CredThumbprint = [string]$s.servicePrincipalCredentialThumbprint
            Country        = [string]$s.location.countryOrRegion
            City           = [string]$s.location.city
            RequestId      = [string]$s.id
        })))
    }
    $spRows = $spList.ToArray()
    if ($SaveRawSignIns) { Export-HuntCsv $spRows 'sp_signins_all.csv' }

    if ($noTime -gt 0) { Write-HuntLog "$noTime records without a timestamp were skipped" -Level WARN }
    if ($spRows.Count -gt 0) {
        $failedCount = @($spRows | Where-Object { $_.ErrorCode -ne 0 }).Count
        $appCount    = @($spRows | ForEach-Object { $_.AppId } | Sort-Object -Unique).Count
        $ipCount     = @($spRows | ForEach-Object { $_.IP } | Sort-Object -Unique).Count
        Write-HuntLog ("Sign-ins: {0} | failed: {1} | apps: {2} | IPs: {3} | range: {4} -> {5}" -f $spRows.Count, $failedCount, $appCount, $ipCount, (Format-HuntTime $oldest), (Format-HuntTime $newest))
        if ($oldest -gt $sinceDt.ToLocalTime().AddDays(1)) {
            Write-HuntLog "Oldest record is newer than the requested window - retention is probably shorter than $DaysBack days" -Level WARN
        }
        Write-HuntLog "Top 5 apps by sign-in volume:" -Level DEBUG
        foreach ($tg in @($spRows | Group-Object AppName | Sort-Object Count -Descending | Select-Object -First 5)) {
            Write-HuntLog ("   {0,7}  {1}" -f $tg.Count, $tg.Name) -Level DEBUG
        }
    } else {
        Write-HuntLog "No service principal sign-ins returned. Check licence (P1/P2) and AuditLog.Read.All." -Level WARN
    }

    Write-HuntLog "Matching against IoC IPs ..."
    $iocSp = @($spRows | Where-Object { $IocIPs -contains $_.IP })
    foreach ($ip in $IocIPs) {
        $n = @($iocSp | Where-Object { $_.IP -eq $ip }).Count
        $lvl = Select-HuntValue ($n -gt 0) 'WARN' 'INFO'
        Write-HuntLog ("   {0,-16} {1} service principal sign-ins" -f $ip, $n) -Level $lvl
    }
    foreach ($g in @($iocSp | Group-Object AppId, IP)) {
        $f   = @($g.Group | Sort-Object Time)
        $ip  = $f[0].IP
        $sev = Select-HuntValue ($HighConfidenceIPs -contains $ip) 'CRITICAL' 'HIGH'
        $ok  = @($f | Where-Object { $_.ErrorCode -eq 0 }).Count
        $keys = (@($f | ForEach-Object { $_.CredKeyId; $_.CredThumbprint }) | Where-Object { $_ } | Sort-Object -Unique) -join ';'
        $resources = (@($f | ForEach-Object { $_.Resource }) | Where-Object { $_ } | Sort-Object -Unique) -join ', '
        Add-HuntFinding $sev 'IoC IP - service principal sign-in' `
            "$($f.Count) sign-ins ($ok successful), $(Format-HuntTime $f[0].Time) -> $(Format-HuntTime $f[-1].Time). Resources: $resources. Credential used: $keys" `
            $f[0].AppName $f[0].AppId $ip $f[0].Time
    }
    Export-HuntCsv $iocSp 'ioc_sp_signins.csv'
    Stop-HuntStep "$($spRows.Count) sign-ins analysed, $($iocSp.Count) from IoC IPs"

    # ---------------------------------------------------------------- 2. other sign-in types from IoC IPs
    Start-HuntStep "Interactive / non-interactive / managed identity sign-ins from IoC IPs"
    $iocOtherList = New-Object System.Collections.Generic.List[object]
    foreach ($ip in $IocIPs) {
        foreach ($type in @('interactiveUser', 'nonInteractiveUser', 'managedIdentity')) {
            try {
                # Invoke-HuntSignInQuery returns the list itself - do not wrap it in @(), that would
                # give an array with the list as its only element
                $res = Invoke-HuntSignInQuery -Filter "ipAddress eq '$ip' and signInEventTypes/any(t: t eq '$type')" `
                    -From $sinceDt -To $now -Hours ($DaysBack * 24) -Label "$type @ $ip"
                $lvl = Select-HuntValue ($res.Count -gt 0) 'WARN' 'INFO'
                Write-HuntLog ("   {0,-16} {1,-19} {2}" -f $ip, $type, $res.Count) -Level $lvl
                foreach ($s in $res) {
                    $when = ConvertTo-HuntLocalDate $s.createdDateTime
                    if ($null -eq $when) { continue }
                    [void]$iocOtherList.Add((New-Object PSObject -Property ([ordered]@{
                        Time = $when; EventType = $type; User = [string]$s.userPrincipalName
                        AppName = [string]$s.appDisplayName; AppId = [string]$s.appId; IP = [string]$s.ipAddress
                        Resource = [string]$s.resourceDisplayName; ErrorCode = [int]$s.status.errorCode
                        FailureReason = [string]$s.status.failureReason; CorrelationId = [string]$s.correlationId
                        Country = [string]$s.location.countryOrRegion; City = [string]$s.location.city
                    })))
                }
            } catch { Write-HuntLog "Query failed for $ip / $type : $($_.Exception.Message)" -Level ERROR }
        }
    }
    $iocOther = $iocOtherList.ToArray()
    foreach ($g in @($iocOther | Group-Object EventType, User, AppId, IP)) {
        $f   = @($g.Group | Sort-Object Time)
        $sev = Select-HuntValue ($HighConfidenceIPs -contains $f[0].IP) 'CRITICAL' 'HIGH'
        $ok  = @($f | Where-Object { $_.ErrorCode -eq 0 }).Count
        $reasons = (@($f | Where-Object { $_.ErrorCode -ne 0 } | ForEach-Object { "$($_.ErrorCode)" }) | Sort-Object -Unique) -join ','
        if (-not $reasons) { $reasons = 'none' }
        $who = '(no user - managed identity)'
        $hint = ''
        if ($f[0].User) { $who = "User $($f[0].User)"; $hint = ' See account triage (step 3) for verdict.' }
        Add-HuntFinding $sev "IoC IP - $($f[0].EventType) sign-in" `
            "${who}: $($f.Count) sign-ins ($ok successful), $(Format-HuntTime $f[0].Time) -> $(Format-HuntTime $f[-1].Time). Error codes: $reasons.$hint" `
            $f[0].AppName $f[0].AppId $f[0].IP $f[0].Time
    }
    Export-HuntCsv $iocOther 'ioc_user_mi_signins.csv'
    $hitUsers = @($iocOther | ForEach-Object { $_.User } | Where-Object { $_ } | Sort-Object -Unique)
    Stop-HuntStep "$($iocOther.Count) user/MI sign-ins from IoC IPs ($($hitUsers.Count) users)"

    # ---------------------------------------------------------------- 3. per-account triage
    Start-HuntStep "Per-account triage (all sign-ins + error codes for every account with a hit)"
    $triageUpns  = New-Object System.Collections.Generic.List[string]   # as written in the log / by the user
    $triageLower = New-Object System.Collections.Generic.List[string]   # lower case, for comparisons
    foreach ($candidate in ($hitUsers + $TriageUsers)) {
        if (-not $candidate) { continue }
        $candidateLower = ([string]$candidate).ToLower()
        if ($triageLower.Contains($candidateLower)) { continue }
        [void]$triageLower.Add($candidateLower)
        [void]$triageUpns.Add([string]$candidate)
    }

    $triageSummary = New-Object System.Collections.Generic.List[object]
    $triageSignins = New-Object System.Collections.Generic.List[object]

    if ($triageUpns.Count -eq 0) {
        Write-HuntLog "No accounts to triage (no user hits from IoC IPs and no -TriageUsers)" -Level OK
    } else {
        Write-HuntLog "Accounts to triage: $($triageUpns.Count)"

        # Security info (MFA method) changes - fetched once, matched per user
        $secInfo = @()
        try {
            Write-HuntLog "Fetching security info / authentication method changes from the audit log ..."
            $uri = New-HuntFilterUri 'https://graph.microsoft.com/v1.0/auditLogs/directoryAudits' `
                "activityDateTime ge $since and category eq 'UserManagement'" '&$top=999'
            $umRaw = Invoke-HuntGraphGetAll $uri -Label 'UserManagement audits'
            $siList = New-Object System.Collections.Generic.List[object]
            foreach ($a in $umRaw) {
                if ([string]$a.activityDisplayName -notmatch 'security info|Strong Authentication|authentication method') { continue }
                $tgt = (@($a.targetResources | ForEach-Object { $_.userPrincipalName }) | Where-Object { $_ }) -join ';'
                [void]$siList.Add((New-Object PSObject -Property ([ordered]@{
                    Time = (ConvertTo-HuntLocalDate $a.activityDateTime); Activity = [string]$a.activityDisplayName; Result = [string]$a.result
                    Target = $tgt; By = [string]$a.initiatedBy.user.userPrincipalName; IP = [string]$a.initiatedBy.user.ipAddress
                })))
            }
            $secInfo = $siList.ToArray()
            Write-HuntLog "$($secInfo.Count) security info changes in the window" -Level DEBUG
        } catch { Write-HuntLog "Could not read security info changes: $($_.Exception.Message)" -Level WARN }

        $u = 0
        foreach ($upn in $triageUpns) {
            $u++
            Write-Host ''
            Write-HuntLog ("--- Account {0}/{1}: {2} ---" -f $u, $triageUpns.Count, $upn) -Level STEP
            $upnEsc = $upn.Replace("'", "''")
            $rowList = New-Object System.Collections.Generic.List[object]

            foreach ($type in @('interactiveUser', 'nonInteractiveUser')) {
                try {
                    $res = Invoke-HuntSignInQuery -Filter "userPrincipalName eq '$upnEsc' and signInEventTypes/any(t: t eq '$type')" `
                        -From $sinceDt -To $now -Hours ($DaysBack * 24) -Max $MaxUserSignIns -Label "$type $upn"
                    Write-HuntLog ("   {0,-19} {1} sign-ins in the window" -f $type, $res.Count)
                    foreach ($s in $res) {
                        $when = ConvertTo-HuntLocalDate $s.createdDateTime
                        if ($null -eq $when) { continue }
                        $code = [int]$s.status.errorCode
                        $info = Get-HuntUserErrInfo $code ([string]$s.status.failureReason)
                        $sip  = [string]$s.ipAddress
                        [void]$rowList.Add((New-Object PSObject -Property ([ordered]@{
                            Time          = $when
                            User          = $upn
                            EventType     = $type
                            IP            = $sip
                            IsIocIP       = [bool]($IocIPs -contains $sip)
                            KnownIP       = [bool](Test-HuntKnownIp $sip)
                            Country       = [string]$s.location.countryOrRegion
                            City          = [string]$s.location.city
                            AppName       = [string]$s.appDisplayName
                            ClientApp     = [string]$s.clientAppUsed
                            ErrorCode     = $code
                            Category      = $info.Cat
                            Meaning       = $info.Text
                            FailureReason = [string]$s.status.failureReason
                            CAStatus      = [string]$s.conditionalAccessStatus
                            MfaMethod     = [string]$s.mfaDetail.authMethod
                            RiskDuring    = [string]$s.riskLevelDuringSignIn
                            CorrelationId = [string]$s.correlationId
                        })))
                    }
                } catch { Write-HuntLog "   Query failed for $type : $($_.Exception.Message)" -Level ERROR }
            }

            $all = @($rowList | Sort-Object Time)
            foreach ($r in $all) { [void]$triageSignins.Add($r) }
            $ioc = @($all | Where-Object { $_.IsIocIP })

            # Error code breakdown - IoC IPs
            Write-HuntLog "   From IoC IPs: $($ioc.Count) sign-ins"
            foreach ($cg in @($ioc | Group-Object ErrorCode | Sort-Object Count -Descending)) {
                $first = $cg.Group[0]
                $codeColor = 'Gray'
                if ($CatColor.ContainsKey($first.Category)) { $codeColor = $CatColor[$first.Category] }
                $ipsHere = (@($cg.Group | ForEach-Object { $_.IP }) | Sort-Object -Unique) -join ', '
                Write-HuntLog ("     {0,6} x {1,-7} {2}  [{3}]" -f $cg.Count, $cg.Name, $first.Meaning, $ipsHere) -Level FIND -Color $codeColor
            }

            # Error code breakdown - all other IPs (context)
            $other = @($all | Where-Object { -not $_.IsIocIP })
            Write-HuntLog "   From other IPs: $($other.Count) sign-ins, $(@($other | ForEach-Object { $_.IP } | Sort-Object -Unique).Count) distinct IPs" -Level DEBUG
            foreach ($cg in @($other | Group-Object ErrorCode | Sort-Object Count -Descending | Select-Object -First 8)) {
                Write-HuntLog ("     {0,6} x {1,-7} {2}" -f $cg.Count, $cg.Name, $cg.Group[0].Meaning) -Level DEBUG
            }

            # Verdict (based on IoC IP sign-ins)
            $cats = @($ioc | ForEach-Object { $_.Category } | Sort-Object -Unique)
            $verdict = 'NO_CLEAR_SIGNAL'; $sev = 'INFO'
            $action  = 'No IoC sign-ins with a clear signal. Review the sign-ins manually.'
            if ($cats -contains 'SPRAY') {
                $verdict = 'SPRAY'; $sev = 'MEDIUM'
                $action  = 'Wrong password only - password not confirmed. Monitor the account, make sure MFA is enforced, block the IP in CA.'
            }
            if ($cats -contains 'LOCKED') {
                $verdict = 'ACTIVE_ATTACK'; $sev = 'HIGH'
                $action  = 'Smart lockout triggered - active attack. Consider a password reset, block the IP in CA, watch for further attempts.'
            }
            if ($cats -contains 'PWD_OK' -or $cats -contains 'CA_BLOCK') {
                $verdict = 'PASSWORD_KNOWN'; $sev = 'CRITICAL'
                $action  = "Attacker has the correct password (blocked only by MFA/CA). Reset password now, run: Revoke-MgUserSignInSession -UserId $upn ; check registered MFA methods and ask the user whether they approved any prompts."
            }
            if ($cats -contains 'SUCCESS') {
                $verdict = 'COMPROMISED'; $sev = 'CRITICAL'
                $action  = "SUCCESSFUL sign-in from IoC IP. Reset password, run: Revoke-MgUserSignInSession -UserId $upn ; remove unknown MFA methods, check inbox rules/forwarding, app consents and all activity after the first success."
            }

            $firstIoc = $null; $lastIoc = $null
            if ($ioc.Count -gt 0) { $firstIoc = $ioc[0].Time; $lastIoc = $ioc[-1].Time }

            # Successful sign-ins from NEW IPs around the attack time
            $newIpSuccess = @()
            if ($ioc.Count -gt 0) {
                $wStart = $firstIoc.AddHours(-24)
                $wEnd   = $lastIoc.AddHours(72)
                $seenBefore = @($all | Where-Object { $_.Time -lt $wStart } | ForEach-Object { $_.IP } | Sort-Object -Unique)
                $newIpSuccess = @($all | Where-Object {
                    $_.Category -eq 'SUCCESS' -and -not $_.IsIocIP -and -not $_.KnownIP -and
                    $_.Time -ge $wStart -and $_.Time -le $wEnd -and ($seenBefore -notcontains $_.IP)
                })
                if ($newIpSuccess.Count -gt 0) {
                    foreach ($ng in @($newIpSuccess | Group-Object IP)) {
                        $n0 = $ng.Group[0]
                        Write-HuntLog ("   NEW IP with success near attack: {0} ({1} {2}) x{3}, app: {4}" -f $ng.Name, $n0.Country, $n0.City, $ng.Count, $n0.AppName) -Level WARN
                    }
                    $nsev = Select-HuntValue ($verdict -eq 'PASSWORD_KNOWN' -or $verdict -eq 'COMPROMISED') 'CRITICAL' 'HIGH'
                    $newIps = (@($newIpSuccess | ForEach-Object { "$($_.IP) ($($_.Country))" }) | Sort-Object -Unique) -join ', '
                    Add-HuntFinding $nsev 'Account: successful sign-in from new IP around attack time' `
                        "$upn signed in successfully from IPs not seen before the attack window: $newIps. Verify with the user - the attacker may have switched IP. (Note: 'not seen before' is limited to the $DaysBack-day window.)" `
                        $upn '' $newIps $newIpSuccess[0].Time
                } else {
                    Write-HuntLog "   No successful sign-ins from new IPs around the attack time" -Level OK
                }
            }

            # Security info changes on this account
            $upnLower = $upn.ToLower()
            $si = @($secInfo | Where-Object { $_.Target -and ($_.Target.ToLower().Contains($upnLower)) })
            if ($si.Count -gt 0) {
                foreach ($x in $si) { Write-HuntLog ("   Security info: {0}  {1} (by {2}, IP {3})" -f (Format-HuntTime $x.Time), $x.Activity, $x.By, $x.IP) -Level WARN }
                $ssev = Select-HuntValue ($verdict -eq 'PASSWORD_KNOWN' -or $verdict -eq 'COMPROMISED') 'CRITICAL' 'MEDIUM'
                Add-HuntFinding $ssev 'Account: security info (MFA method) changed in window' `
                    "$upn had $($si.Count) security info change(s), latest $(Format-HuntTime $si[-1].Time). If this happened after the attack, an attacker may have registered their own MFA method." `
                    $upn '' '' $si[-1].Time
            } else {
                Write-HuntLog "   No security info changes in the window" -Level OK
            }

            # Identity Protection risk
            $riskText = 'n/a'
            try {
                $ru = Invoke-HuntGraphGetAll (New-HuntFilterUri 'https://graph.microsoft.com/v1.0/identityProtection/riskyUsers' "userPrincipalName eq '$upnEsc'") -Label "Risk $upn"
                if ($ru.Count -gt 0) {
                    $riskText = "$($ru[0].riskLevel) / $($ru[0].riskState)"
                    $rlvl = Select-HuntValue ($ru[0].riskLevel -eq 'high' -or $ru[0].riskLevel -eq 'medium') 'WARN' 'INFO'
                    Write-HuntLog "   Identity Protection: risk $riskText (updated $(Format-HuntTime (ConvertTo-HuntLocalDate $ru[0].riskLastUpdatedDateTime)))" -Level $rlvl
                } else {
                    $riskText = 'not risky'
                    Write-HuntLog "   Identity Protection: not flagged as risky" -Level DEBUG
                }
            } catch { Write-HuntLog "   Could not read risky user state: $($_.Exception.Message)" -Level DEBUG }

            $codeSummary = (@($ioc | Group-Object ErrorCode | ForEach-Object { "$($_.Name)x$($_.Count)" })) -join ' '
            $vColor = 'Gray'
            if ($SevColor.ContainsKey($sev)) { $vColor = $SevColor[$sev] }
            Write-HuntLog ("   VERDICT: {0}  ({1})" -f $verdict, $codeSummary) -Level FIND -Color $vColor
            Write-HuntLog "   ACTION:  $action" -Level FIND -Color $vColor

            Add-HuntFinding $sev "Account triage: $verdict" `
                "$upn. IoC sign-ins: $($ioc.Count) [$codeSummary] $(Format-HuntTime $firstIoc) -> $(Format-HuntTime $lastIoc). New-IP successes: $($newIpSuccess.Count). Security info changes: $($si.Count). Risk: $riskText. ACTION: $action" `
                $upn '' ((@($ioc | ForEach-Object { $_.IP }) | Sort-Object -Unique) -join ';') $firstIoc

            [void]$triageSummary.Add((New-Object PSObject -Property ([ordered]@{
                User                 = $upn
                Verdict              = $verdict
                Severity             = $sev
                IocSignIns           = $ioc.Count
                IocSuccess           = @($ioc | Where-Object { $_.Category -eq 'SUCCESS' }).Count
                IocPasswordCorrect   = @($ioc | Where-Object { $_.Category -eq 'PWD_OK' -or $_.Category -eq 'CA_BLOCK' }).Count
                IocWrongPassword     = @($ioc | Where-Object { $_.Category -eq 'SPRAY' }).Count
                IocLocked            = @($ioc | Where-Object { $_.Category -eq 'LOCKED' }).Count
                IocErrorCodes        = $codeSummary
                IocIPs               = (@($ioc | ForEach-Object { $_.IP }) | Sort-Object -Unique) -join ';'
                IocApps              = (@($ioc | ForEach-Object { $_.AppName }) | Sort-Object -Unique) -join ';'
                FirstIoc             = $firstIoc
                LastIoc              = $lastIoc
                NewIpSuccesses       = $newIpSuccess.Count
                SecurityInfoChanges  = $si.Count
                IdentityProtection   = $riskText
                TotalSignInsInWindow = $all.Count
                Action               = $action
            })))
        }
    }
    Export-HuntCsv $triageSummary 'account_triage_summary.csv'
    if ($SaveRawSignIns) { Export-HuntCsv $triageSignins 'account_triage_signins.csv' }
    $compromised = @($triageSummary | Where-Object { $_.Verdict -eq 'COMPROMISED' -or $_.Verdict -eq 'PASSWORD_KNOWN' }).Count
    Stop-HuntStep "$($triageSummary.Count) accounts triaged, $compromised with known password or compromise"

    # ---------------------------------------------------------------- 4. invalid / expired secret
    Start-HuntStep "AADSTS7000215 / AADSTS7000222 from unknown IPs"
    $allSecretFails = @($spRows | Where-Object { $_.ErrorCode -eq 7000215 -or $_.ErrorCode -eq 7000222 })
    $secretFails    = @($allSecretFails | Where-Object { -not (Test-HuntKnownIp $_.IP) })
    Write-HuntLog ("Total 7000215: {0} | 7000222: {1} | excluded as known IP: {2} | remaining: {3}" -f `
        @($allSecretFails | Where-Object { $_.ErrorCode -eq 7000215 }).Count,
        @($allSecretFails | Where-Object { $_.ErrorCode -eq 7000222 }).Count,
        ($allSecretFails.Count - $secretFails.Count), $secretFails.Count)
    foreach ($sg in @($secretFails | Group-Object AppName | Sort-Object Count -Descending)) {
        $ips = (@($sg.Group | ForEach-Object { $_.IP }) | Sort-Object -Unique) -join ', '
        Write-HuntLog ("   {0,6}  {1}  (IPs: {2})" -f $sg.Count, $sg.Name, $ips) -Level DEBUG
    }
    Export-HuntCsv $secretFails 'secret_failures.csv'
    $failApps = @($secretFails | ForEach-Object { $_.AppId } | Sort-Object -Unique).Count
    Stop-HuntStep "$($secretFails.Count) secret failures from unknown IPs across $failApps apps"

    # ---------------------------------------------------------------- 5. failure -> success
    Start-HuntStep "Correlating secret failure -> success from the same IP"
    $successIndex = @{}
    foreach ($r in $spRows) {
        if ($r.ErrorCode -ne 0) { continue }
        $k = "$($r.AppId)|$($r.IP)"
        if (-not $successIndex.ContainsKey($k)) { $successIndex[$k] = New-Object System.Collections.Generic.List[object] }
        [void]$successIndex[$k].Add($r)
    }
    Write-HuntLog "Indexed successful sign-ins: $($successIndex.Count) app+IP combinations" -Level DEBUG

    $groups = @($secretFails | Group-Object AppId, IP)
    Write-HuntLog "Checking $($groups.Count) app+IP combinations with secret failures ..."
    $critCount = 0
    foreach ($g in $groups) {
        $f     = @($g.Group | Sort-Object Time)
        $codes = (@($f | ForEach-Object { $_.ErrorCode }) | Sort-Object -Unique) -join ','
        $k     = "$($f[0].AppId)|$($f[0].IP)"
        $firstFail = $f[0].Time
        $later = @()
        if ($successIndex.ContainsKey($k)) { $later = @($successIndex[$k] | Where-Object { $_.Time -gt $firstFail } | Sort-Object Time) }
        Write-HuntLog ("   {0} @ {1}: {2} failures, {3} later successes" -f $f[0].AppName, $f[0].IP, $f.Count, $later.Count) -Level DEBUG
        if ($later.Count -gt 0) {
            $critCount++
            $res = (@($later | ForEach-Object { $_.Resource }) | Where-Object { $_ } | Sort-Object -Unique) -join ', '
            Add-HuntFinding 'CRITICAL' 'Secret failure followed by SUCCESS (same app + IP)' `
                "$($f.Count) failures ($codes) from $(Format-HuntTime $firstFail), first success $(Format-HuntTime $later[0].Time) using credential $($later[0].CredKeyId)$($later[0].CredThumbprint). Resources after: $res" `
                $f[0].AppName $f[0].AppId $f[0].IP $firstFail
        } else {
            Add-HuntFinding 'HIGH' 'Invalid/expired secret from unknown IP' `
                "$($f.Count) failures ($codes) $(Format-HuntTime $firstFail) -> $(Format-HuntTime $f[-1].Time), $($f[0].Country) $($f[0].City). Someone may hold an old credential for this app." `
                $f[0].AppName $f[0].AppId $f[0].IP $firstFail
        }
    }
    Stop-HuntStep "$critCount failure->success sequences"

    # ---------------------------------------------------------------- 6. one IP -> many service principals
    Start-HuntStep "One IP hitting >= $SpThreshold service principals"
    $ipGroups = @($spRows | Where-Object { $_.IP -and -not (Test-HuntKnownIp $_.IP) } | Group-Object IP)
    Write-HuntLog "Evaluating $($ipGroups.Count) unknown source IPs ..."
    foreach ($grp in $ipGroups) {
        $apps = @($grp.Group | ForEach-Object { $_.AppId } | Sort-Object -Unique)
        if ($apps.Count -ge $SpThreshold) {
            $names  = (@($grp.Group | ForEach-Object { $_.AppName }) | Sort-Object -Unique) -join ', '
            $first  = @($grp.Group | Sort-Object Time)[0]
            $failed = @($grp.Group | Where-Object { $_.ErrorCode -ne 0 }).Count
            Add-HuntFinding 'MEDIUM' "One IP -> $($apps.Count) service principals" `
                "Apps: $names. $($grp.Count) sign-ins, $failed failed. Verify this is your own automation/Azure egress, otherwise add to -KnownIPs or investigate." `
                '' '' $grp.Name $first.Time
            Write-HuntLog "     Apps: $names" -Level DEBUG
        }
    }
    Stop-HuntStep "$($ipGroups.Count) IPs evaluated"

    # ---------------------------------------------------------------- 7. token sweep ARM + KV + Storage + Graph
    Start-HuntStep "One app grabbing ARM + Key Vault + Storage + Graph within $TokenWindowMinutes min"
    $sweepList = New-Object System.Collections.Generic.List[object]
    foreach ($r in $spRows) {
        if ($r.ErrorCode -ne 0) { continue }
        $cat = Get-HuntResCategory $r
        if ($cat) {
            [void]$sweepList.Add((New-Object PSObject -Property ([ordered]@{
                Time = $r.Time; AppId = $r.AppId; AppName = $r.AppName; IP = $r.IP; Cat = $cat
            })))
        }
    }
    $sweepRows = $sweepList.ToArray()
    foreach ($cat in @('ARM', 'KeyVault', 'Storage', 'Graph')) {
        $cnt = @($sweepRows | Where-Object { $_.Cat -eq $cat }).Count
        Write-HuntLog ("   {0,-9} {1} successful token requests" -f $cat, $cnt) -Level DEBUG
    }
    $sweepGroups = @($sweepRows | Group-Object AppId)
    Write-HuntLog "Evaluating $($sweepGroups.Count) apps with tokens for at least one of the four resources ..."
    foreach ($g in $sweepGroups) {
        $list = @($g.Group | Sort-Object Time)
        $cats = @($list | ForEach-Object { $_.Cat } | Sort-Object -Unique)
        if ($cats.Count -lt 4) { continue }
        Write-HuntLog "   $($list[0].AppName) has tokens for all four over the whole period - checking ${TokenWindowMinutes}-min windows" -Level DEBUG
        $counts = @{}
        $start = 0
        for ($end = 0; $end -lt $list.Count; $end++) {
            $c = $list[$end].Cat
            if ($counts.ContainsKey($c)) { $counts[$c] = $counts[$c] + 1 } else { $counts[$c] = 1 }
            while (($list[$end].Time - $list[$start].Time).TotalMinutes -gt $TokenWindowMinutes) {
                $sc = $list[$start].Cat
                $counts[$sc] = $counts[$sc] - 1
                if ($counts[$sc] -le 0) { $counts.Remove($sc) }
                $start++
            }
            if ($counts.Count -ge 4) {
                $win = $list[$start..$end]
                $winIps = (@($win | ForEach-Object { $_.IP }) | Sort-Object -Unique)
                Add-HuntFinding 'HIGH' 'Token sweep: ARM + Key Vault + Storage + Graph' `
                    "All four resources between $(Format-HuntTime $win[0].Time) and $(Format-HuntTime $win[-1].Time). IPs: $($winIps -join ', ')" `
                    $list[0].AppName $g.Name ($winIps -join ';') $win[0].Time
                break
            }
        }
    }
    Stop-HuntStep "$($sweepGroups.Count) apps evaluated"

    # ---------------------------------------------------------------- 8. credential inventory
    $credList = New-Object System.Collections.Generic.List[object]
    $credRows = @()
    Start-HuntStep "Credential inventory (apps + service principals)"
    if (-not $SkipInventory) {
        $lastSignIn = @{}
        try {
            Write-HuntLog "Fetching last sign-in per service principal ..."
            foreach ($a in (Invoke-HuntGraphGetAll 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?$top=1000' -Label 'SP sign-in activity')) {
                $lastSeen = ConvertTo-HuntLocalDate $a.lastSignInActivity.lastSignInDateTime
                if ($lastSeen) { $lastSignIn[[string]$a.appId] = $lastSeen }
            }
            Write-HuntLog "Last sign-in known for $($lastSignIn.Count) apps" -Level OK
        } catch { Write-HuntLog "Could not read servicePrincipalSignInActivities: $($_.Exception.Message)" -Level WARN }
        # The sign-in log read in step 1 can be newer than the report - use the latest of the two
        foreach ($appKey in $seenInLogs.Keys) {
            if (-not $lastSignIn.ContainsKey($appKey) -or $seenInLogs[$appKey] -gt $lastSignIn[$appKey]) { $lastSignIn[$appKey] = $seenInLogs[$appKey] }
        }

        Write-HuntLog "Fetching app registrations ..."
        $apps = Invoke-HuntGraphGetAll 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName,signInAudience,createdDateTime,passwordCredentials,keyCredentials&$top=999' -Label 'Applications'
        $what = 'credentials and owners'
        if (-not $SkipFederatedCredentials) { $what = 'credentials, federated credentials and owners' }
        Write-HuntLog "$($apps.Count) app registrations. Reading $what per app ..."
        if (-not $SkipFederatedCredentials -and $apps.Count -gt 500) {
            Write-HuntLog "Many apps - federated credential lookups need one call per app. Use -SkipFederatedCredentials to speed up." -Level WARN
        }
        $i = 0
        $appsWithCreds = 0
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        foreach ($app in $apps) {
            $i++
            Write-Progress -Activity 'Applications' -Status "$i / $($apps.Count): $($app.displayName)" -PercentComplete ([int](100 * $i / [math]::Max(1, $apps.Count)))
            if ($i % 100 -eq 0) {
                $rate = $i / [math]::Max(0.1, $sw.Elapsed.TotalSeconds)
                $eta  = [timespan]::FromSeconds(($apps.Count - $i) / [math]::Max(0.01, $rate))
                Write-HuntLog ("   progress: {0}/{1} apps, {2} with credentials, ETA {3}" -f $i, $apps.Count, $appsWithCreds, (Format-HuntDuration $eta))
            }
            $creds = New-Object System.Collections.Generic.List[object]
            foreach ($p in $app.passwordCredentials) { if ($p) { [void]$creds.Add(@{ Type = 'ClientSecret'; KeyId = $p.keyId; Name = $p.displayName; Start = $p.startDateTime; End = $p.endDateTime; Extra = $p.hint }) } }
            foreach ($k in $app.keyCredentials) { if ($k) { [void]$creds.Add(@{ Type = 'Certificate';  KeyId = $k.keyId; Name = $k.displayName; Start = $k.startDateTime; End = $k.endDateTime; Extra = $k.customKeyIdentifier }) } }
            if (-not $SkipFederatedCredentials) {
                try {
                    foreach ($fic in (Invoke-HuntGraphGetAll "https://graph.microsoft.com/v1.0/applications/$($app.id)/federatedIdentityCredentials" -Label "FIC $($app.displayName)")) {
                        [void]$creds.Add(@{ Type = 'FederatedCredential'; KeyId = $fic.id; Name = $fic.name; Start = $null; End = $null; Extra = "$($fic.issuer) | $($fic.subject)" })
                    }
                } catch { Write-HuntLog "Federated credential read failed for $($app.displayName)" -Level WARN }
            }
            if ($creds.Count -eq 0) { continue }
            $appsWithCreds++

            $owners = ''
            try {
                $o = Invoke-HuntGraphGetAll "https://graph.microsoft.com/v1.0/applications/$($app.id)/owners?`$select=id,displayName,userPrincipalName" -Label "Owners $($app.displayName)"
                $owners = (@($o | ForEach-Object { Select-HuntValue $_.userPrincipalName $_.userPrincipalName $_.displayName })) -join ';'
            } catch { Write-HuntLog "Owner read failed for $($app.displayName)" -Level WARN }

            $ownerText = Select-HuntValue $owners $owners '(none)'
            Write-HuntLog ("   {0}: {1} secret(s), {2} cert(s), {3} federated | owners: {4}" -f $app.displayName,
                @($creds | Where-Object { $_.Type -eq 'ClientSecret' }).Count,
                @($creds | Where-Object { $_.Type -eq 'Certificate' }).Count,
                @($creds | Where-Object { $_.Type -eq 'FederatedCredential' }).Count,
                $ownerText) -Level DEBUG

            $appLast = $null
            if ($lastSignIn.ContainsKey([string]$app.appId)) { $appLast = $lastSignIn[[string]$app.appId] }
            foreach ($c in $creds) {
                $end   = ConvertTo-HuntLocalDate $c.End
                $start = ConvertTo-HuntLocalDate $c.Start
                $expired = $false; if ($end -and $end -lt $nowLocal) { $expired = $true }
                $recent  = $false; if ($start -and $start -gt $nowLocal.AddDays(-$NewCredentialDays)) { $recent = $true }
                [void]$credList.Add((New-Object PSObject -Property ([ordered]@{
                    Object          = 'Application'
                    AppName         = [string]$app.displayName
                    AppId           = [string]$app.appId
                    ObjectId        = [string]$app.id
                    SignInAudience  = [string]$app.signInAudience
                    CredentialType  = $c.Type
                    KeyId           = [string]$c.KeyId
                    CredentialName  = [string]$c.Name
                    Start           = $start
                    End             = $end
                    Expired         = $expired
                    AddedRecently   = $recent
                    Owners          = $owners
                    LastSignIn      = $appLast
                    Details         = [string]$c.Extra
                })))
            }
        }
        Write-Progress -Activity 'Applications' -Completed
        Write-HuntLog "App registrations done: $appsWithCreds of $($apps.Count) have credentials ($([math]::Round($sw.Elapsed.TotalSeconds,1)) s)" -Level OK

        Write-HuntLog "Fetching service principals (looking for credentials placed directly on SP objects) ..."
        $sps = Invoke-HuntGraphGetAll 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId,accountEnabled,servicePrincipalType,passwordCredentials,keyCredentials&$top=999' -Label 'Service principals'
        Write-HuntLog "$($sps.Count) service principals"
        $spWithCreds = 0
        foreach ($sp in $sps) {
            $list = New-Object System.Collections.Generic.List[object]
            foreach ($p in $sp.passwordCredentials) { if ($p) { [void]$list.Add(@{ Type = 'ClientSecret'; KeyId = $p.keyId; Name = $p.displayName; Start = $p.startDateTime; End = $p.endDateTime }) } }
            foreach ($k in $sp.keyCredentials) { if ($k -and $k.usage -eq 'Verify') { [void]$list.Add(@{ Type = 'Certificate'; KeyId = $k.keyId; Name = $k.displayName; Start = $k.startDateTime; End = $k.endDateTime }) } }
            if ($list.Count -eq 0) { continue }
            $spWithCreds++
            Write-HuntLog ("   SP {0}: {1} credential(s) on the SP object" -f $sp.displayName, $list.Count) -Level DEBUG

            $foreign = $false
            if ($sp.appOwnerOrganizationId -and ([string]$sp.appOwnerOrganizationId -ne [string]$ctx.TenantId)) { $foreign = $true }
            $audience = 'This tenant'
            if ($foreign) { $audience = "Multi-tenant (owner $($sp.appOwnerOrganizationId))" }
            $spLast = $null
            if ($lastSignIn.ContainsKey([string]$sp.appId)) { $spLast = $lastSignIn[[string]$sp.appId] }

            foreach ($c in $list) {
                $end   = ConvertTo-HuntLocalDate $c.End
                $start = ConvertTo-HuntLocalDate $c.Start
                $expired = $false; if ($end -and $end -lt $nowLocal) { $expired = $true }
                $recent  = $false; if ($start -and $start -gt $nowLocal.AddDays(-$NewCredentialDays)) { $recent = $true }
                [void]$credList.Add((New-Object PSObject -Property ([ordered]@{
                    Object = 'ServicePrincipal'; AppName = [string]$sp.displayName; AppId = [string]$sp.appId; ObjectId = [string]$sp.id
                    SignInAudience = $audience
                    CredentialType = $c.Type; KeyId = [string]$c.KeyId; CredentialName = [string]$c.Name; Start = $start; End = $end
                    Expired = $expired; AddedRecently = $recent
                    Owners = ''; LastSignIn = $spLast; Details = "SP type: $($sp.servicePrincipalType), enabled: $($sp.accountEnabled)"
                })))
                if ($c.Type -eq 'ClientSecret' -and -not $expired) {
                    $sev = 'MEDIUM'
                    if ($foreign -or $recent) { $sev = 'HIGH' }
                    Add-HuntFinding $sev 'Secret on service principal object' `
                        "Secret '$($c.Name)' (KeyId $($c.KeyId)) added $(Format-HuntTime $start) is on the SP, not the app registration - not visible under App registrations. Verify it is expected." `
                        $sp.displayName $sp.appId '' $start
                }
            }
        }
        $spLvl = Select-HuntValue ($spWithCreds -gt 0) 'WARN' 'OK'
        Write-HuntLog "$spWithCreds service principals have credentials on the SP object itself" -Level $spLvl

        $credRows = $credList.ToArray()
        Export-HuntCsv ($credRows | Sort-Object Object, AppName) 'credential_inventory.csv'

        foreach ($r in $credRows) {
            if ($r.AddedRecently -and $r.Object -eq 'Application') {
                Add-HuntFinding 'MEDIUM' "Credential added last $NewCredentialDays days" `
                    "$($r.CredentialType) '$($r.CredentialName)' KeyId $($r.KeyId), owners: $($r.Owners)" $r.AppName $r.AppId '' $r.Start
            }
        }
        $unusedCutoff = $nowLocal.AddDays(-$UnusedDays)
        $unused = @($credRows | Where-Object { $_.CredentialType -ne 'FederatedCredential' -and -not $_.Expired -and (-not $_.LastSignIn -or $_.LastSignIn -lt $unusedCutoff) })
        foreach ($ug in @($unused | Group-Object AppId)) {
            $r = $ug.Group[0]
            Add-HuntFinding 'INFO' "Valid credential, no sign-in for $UnusedDays+ days" `
                "$($ug.Count) active credential(s). Last sign-in: $(Format-HuntTime $r.LastSignIn). Candidate for removal." $r.AppName $r.AppId '' $r.LastSignIn
        }
        Stop-HuntStep "$($credRows.Count) credentials inventoried"
    } else {
        Write-HuntLog "Inventory skipped (-SkipInventory)" -Level WARN
        Stop-HuntStep "skipped"
    }

    # ---------------------------------------------------------------- 9. audit log: credential & owner changes
    Start-HuntStep "Audit log: credential / owner / consent changes"
    $auditList = New-Object System.Collections.Generic.List[object]
    try {
        $uri = New-HuntFilterUri 'https://graph.microsoft.com/v1.0/auditLogs/directoryAudits' `
            "activityDateTime ge $since and category eq 'ApplicationManagement'" '&$top=999'
        $auditRaw = Invoke-HuntGraphGetAll $uri -Label 'Directory audits'
        Write-HuntLog "$($auditRaw.Count) ApplicationManagement events, filtering for credential/owner/consent changes ..."
        $pattern = 'Certificates and secrets|service principal credentials|federated identity credential|Add owner to (application|service principal)|Consent to application|Add app role assignment|Add delegated permission'
        foreach ($a in $auditRaw) {
            if ([string]$a.activityDisplayName -notmatch $pattern) { continue }
            $byUser = [string]$a.initiatedBy.user.userPrincipalName
            $byApp  = [string]$a.initiatedBy.app.displayName
            $ip     = [string]$a.initiatedBy.user.ipAddress
            $row = New-Object PSObject -Property ([ordered]@{
                Time = (ConvertTo-HuntLocalDate $a.activityDateTime); Activity = [string]$a.activityDisplayName; Result = [string]$a.result
                InitiatedByUser = $byUser; InitiatedByApp = $byApp; IP = $ip
                Target   = (@($a.targetResources | ForEach-Object { $_.displayName }) | Where-Object { $_ }) -join ';'
                TargetId = (@($a.targetResources | ForEach-Object { $_.id }) | Where-Object { $_ }) -join ';'
            })
            [void]$auditList.Add($row)
            $by = "app:$byApp"
            if ($byUser) { $by = $byUser }
            Write-HuntLog ("   {0}  {1} | {2} | by {3}" -f (Format-HuntTime $row.Time), $row.Activity, $row.Target, $by) -Level DEBUG
            if ($ip -and ($IocIPs -contains $ip)) {
                Add-HuntFinding 'CRITICAL' 'Audit change from IoC IP' "$($row.Activity) on $($row.Target) by $byUser" $row.Target '' $ip $row.Time
            } elseif ($byApp -and -not $byUser -and $row.Activity -match 'secret|credential|federated') {
                Add-HuntFinding 'MEDIUM' 'Credential added by an application (not a user)' `
                    "$($row.Activity) on $($row.Target) by app '$byApp'. Apps adding credentials to apps is a common persistence technique." $row.Target '' '' $row.Time
            } elseif ($byUser -and $triageLower.Contains($byUser.ToLower())) {
                Add-HuntFinding 'HIGH' 'App change made by a triaged (attacked) account' `
                    "$($row.Activity) on $($row.Target) by $byUser from IP $ip. This account was hit from an IoC IP - verify the change." $row.Target '' $ip $row.Time
            }
        }
        foreach ($ag in @($auditList | Group-Object Activity | Sort-Object Count -Descending)) {
            Write-HuntLog ("   {0,5}  {1}" -f $ag.Count, $ag.Name)
        }
    } catch { Write-HuntLog "Audit log read failed: $($_.Exception.Message)" -Level ERROR }
    $audit = $auditList.ToArray()
    Export-HuntCsv ($audit | Sort-Object Time) 'audit_credential_changes.csv'
    Stop-HuntStep "$($audit.Count) relevant audit events"

    # ---------------------------------------------------------------- 10. posture
    Start-HuntStep "Posture: Conditional Access and app management policy"
    try {
        $locs = Invoke-HuntGraphGetAll 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations' -Label 'Named locations'
        Write-HuntLog "$($locs.Count) named locations"
        $iocLocs = New-Object System.Collections.Generic.List[object]
        foreach ($l in $locs) {
            foreach ($cidr in @($l.ipRanges | ForEach-Object { $_.cidrAddress })) {
                if ($cidr -and (Test-HuntIpInRange $HighConfidenceIPs[0] $cidr)) { [void]$iocLocs.Add($l); break }
            }
        }
        $iocLocIds = @($iocLocs | ForEach-Object { [string]$_.id })
        if ($iocLocs.Count -gt 0) { Write-HuntLog "Named location(s) containing $($HighConfidenceIPs[0]): $((@($iocLocs | ForEach-Object { $_.displayName })) -join ', ')" -Level OK }

        $policies = Invoke-HuntGraphGetAll 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Label 'CA policies'
        Write-HuntLog ("{0} CA policies ({1} enabled, {2} report-only, {3} targeting workload identities)" -f $policies.Count,
            @($policies | Where-Object { $_.state -eq 'enabled' }).Count,
            @($policies | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' }).Count,
            @($policies | Where-Object { $_.conditions.clientApplications.includeServicePrincipals }).Count)

        $wiBlock   = New-Object System.Collections.Generic.List[object]
        $userBlock = New-Object System.Collections.Generic.List[object]
        foreach ($p in $policies) {
            if ($p.state -ne 'enabled') { continue }
            if ($p.grantControls.builtInControls -notcontains 'block') { continue }
            $hitsLoc = $false
            foreach ($li in $p.conditions.locations.includeLocations) { if ($iocLocIds -contains [string]$li) { $hitsLoc = $true } }
            if (-not $hitsLoc) { continue }
            if ($p.conditions.clientApplications.includeServicePrincipals) { [void]$wiBlock.Add($p) }
            if ($p.conditions.users.includeUsers) { [void]$userBlock.Add($p) }
        }
        if ($iocLocIds.Count -eq 0) {
            Add-HuntFinding 'MEDIUM' 'Posture: IoC IP not in any named location' "Create a named location with $($HighConfidenceIPs[0])/32 (the address, not the range)." '' '' $HighConfidenceIPs[0] $null
        }
        if ($wiBlock.Count -eq 0) {
            Add-HuntFinding 'MEDIUM' 'Posture: no enabled workload identity CA block for IoC IP' "Needs Workload ID Premium. Assign service principals directly (group assignment is not enforced). Covers single-tenant apps registered in this tenant only." '' '' $HighConfidenceIPs[0] $null
        } else {
            Write-HuntLog "Workload identity block policy: $((@($wiBlock | ForEach-Object { $_.displayName })) -join '; ')" -Level OK
        }
        if ($userBlock.Count -eq 0) {
            Add-HuntFinding 'MEDIUM' 'Posture: no enabled user CA block for IoC IP' "Add the named location to a user policy to block replayed sessions." '' '' $HighConfidenceIPs[0] $null
        } else {
            Write-HuntLog "User block policy: $((@($userBlock | ForEach-Object { $_.displayName })) -join '; ')" -Level OK
        }
    } catch { Write-HuntLog "Conditional Access read failed: $($_.Exception.Message)" -Level ERROR }

    try {
        Write-HuntLog "Reading default app management policy ..."
        $amp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/defaultAppManagementPolicy' -OutputType PSObject
        Write-HuntLog "Default app management policy enabled: $($amp.isEnabled)" -Level DEBUG
        $blocksSecrets = $false
        if ($amp.isEnabled) {
            foreach ($pr in $amp.applicationRestrictions.passwordCredentials) {
                if ($pr -and $pr.restrictionType -eq 'passwordAddition' -and $pr.state -ne 'disabled') { $blocksSecrets = $true }
            }
        }
        if (-not $blocksSecrets) {
            Add-HuntFinding 'MEDIUM' 'Posture: new client secrets are not blocked' 'Enable a tenant app management policy with passwordAddition restriction once apps have moved to managed identity / workload identity federation / certificates.' '' '' '' $null
        } else {
            Write-HuntLog "New client secrets are blocked by the default app management policy" -Level OK
        }
    } catch { Write-HuntLog "App management policy read failed: $($_.Exception.Message)" -Level ERROR }
    Stop-HuntStep "posture checked"

    # ---------------------------------------------------------------- output
    $sevOrder = @{ CRITICAL = 0; HIGH = 1; MEDIUM = 2; INFO = 3 }
    $sorted = @($findings | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } }, Time)
    Export-HuntCsv $sorted 'FINDINGS.csv'

    Write-Host ''
    Write-HuntLog "================ SUMMARY ================" -Level STEP
    Write-HuntLog "Window: last $DaysBack days (since $since)"
    Write-HuntLog "Service principal sign-ins analysed: $($spRows.Count)"

    if ($triageSummary.Count -gt 0) {
        Write-HuntLog "Account triage:"
        foreach ($t in @($triageSummary | Sort-Object @{ Expression = { $sevOrder[$_.Severity] } })) {
            $tc = 'Gray'
            if ($SevColor.ContainsKey($t.Severity)) { $tc = $SevColor[$t.Severity] }
            Write-HuntLog ("   {0,-16} {1}  [{2}]" -f $t.Verdict, $t.User, $t.IocErrorCodes) -Color $tc
        }
    }

    if (-not $SkipInventory) {
        $secrets = @($credRows | Where-Object { $_.CredentialType -eq 'ClientSecret' -and -not $_.Expired })
        $secretApps = @($secrets | ForEach-Object { $_.AppId } | Sort-Object -Unique).Count
        Write-HuntLog "Active client secrets: $($secrets.Count) on $secretApps apps"
        Write-HuntLog "  without owner:        $(@($secrets | Where-Object { $_.Object -eq 'Application' -and -not $_.Owners }).Count)"
        Write-HuntLog "  on SP objects:        $(@($secrets | Where-Object { $_.Object -eq 'ServicePrincipal' }).Count)"
        Write-HuntLog "Active certificates:   $(@($credRows | Where-Object { $_.CredentialType -eq 'Certificate' -and -not $_.Expired }).Count)"
        Write-HuntLog "Federated credentials: $(@($credRows | Where-Object { $_.CredentialType -eq 'FederatedCredential' }).Count)"
    }
    foreach ($s in @('CRITICAL', 'HIGH', 'MEDIUM', 'INFO')) {
        $n = @($findings | Where-Object { $_.Severity -eq $s }).Count
        Write-HuntLog ("{0,-9} {1}" -f $s, $n) -Level FIND -Color $SevColor[$s]
    }
    $top = @($sorted | Where-Object { $_.Severity -eq 'CRITICAL' -or $_.Severity -eq 'HIGH' })
    if ($top.Count -gt 0) {
        $top | Format-Table Severity, Category, AppName, IP, Time -AutoSize -Wrap | Out-Host
    }

    if (@($findings | Where-Object { $_.Severity -eq 'CRITICAL' }).Count -gt 0) {
        Write-Host @"
CRITICAL findings.
  Accounts (COMPROMISED / PASSWORD_KNOWN):
  - Reset password, Revoke-MgUserSignInSession, remove unknown MFA methods
  - Check inbox rules/forwarding, app consents, activity after first success
  Apps / service principals:
  - Remove ALL secrets, certificates and federated credentials (incl. ones you did not add)
  - Disable the service principal if needed - issued tokens stay valid until they expire
  - Rotate everything in any Key Vault the app could read and credentials in files it accessed
  - MSPs: check Partner Center / GDAP and notify affected customers
"@ -ForegroundColor Red
    }
    Write-HuntLog "Files in $OutputFolder :"
    foreach ($fi in @(Get-ChildItem $OutputFolder -File)) { Write-HuntLog ("   {0,-36} {1,10:N0} bytes" -f $fi.Name, $fi.Length) }
    Write-HuntLog "Total run time: $(Format-HuntDuration ((Get-Date) - $state.RunStart))" -Level OK

    # ---------------------------------------------------------------- HTML report
    if (-not $NoHtml) {
        Write-Host ''
        Write-HuntLog "===== HTML report =====" -Level STEP
        try {
            New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null
            $safeName = ($tenantName -replace '[\\/:*?"<>|]', '_').Trim()
            $htmlPath = Join-Path $ReportFolder ("WorkloadIdentityHunt " + $safeName + " " + $runStamp + ".html")
            $htmlSw   = [System.Diagnostics.Stopwatch]::StartNew()

            $activeSecrets = @($credRows | Where-Object { $_.CredentialType -eq 'ClientSecret' -and -not $_.Expired })
            $meta = [ordered]@{
                title = "Workload identity hunt - $tenantName"
                stamp = $runStamp
                info  = [ordered]@{
                    'Tenant'     = $tenantName
                    'Domain'     = $tenantDomain
                    'Tenant ID'  = [string]$ctx.TenantId
                    'Run by'     = [string]$ctx.Account
                    'Window'     = "$DaysBack days (from $since UTC)"
                    'Generated'  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    'IoC IP'     = ($IocIPs -join ', ')
                    'Run time'   = (Format-HuntDuration ((Get-Date) - $state.RunStart))
                    'PowerShell' = "$($PSVersionTable.PSVersion)"
                }
                stats = [ordered]@{
                    'SP sign-ins analysed'      = $spRows.Count
                    'Accounts triaged'          = $triageSummary.Count
                    'Active client secrets'     = $activeSecrets.Count
                    'Secrets without owner'     = @($activeSecrets | Where-Object { $_.Object -eq 'Application' -and -not $_.Owners }).Count
                    'Secrets on SP objects'     = @($activeSecrets | Where-Object { $_.Object -eq 'ServicePrincipal' }).Count
                }
            }

            Write-HuntLog "Building report sections ..."
            $sections = @(
                (New-HuntReportSection -Id 'findings' -Title 'Findings' -DefaultSort 'Severity' `
                    -Description 'All findings, sorted by severity. Click the cards above to filter.' `
                    -Rows $sorted -Columns @('Severity', 'Category', 'AppName', 'IP', 'Time', 'Detail', 'AppId')),
                (New-HuntReportSection -Id 'triage' -Title 'Account triage' -DefaultSort 'Verdict' `
                    -Description 'One row per account with a hit: verdict, error codes from IoC IPs and action.' `
                    -Rows $triageSummary.ToArray()),
                (New-HuntReportSection -Id 'triage_signins' -Title 'Account sign-ins' -DefaultSort 'Time' -MaxRows $HtmlMaxRows `
                    -Description 'All sign-ins for the triaged accounts, with the error code classified (Category/Meaning).' `
                    -Rows $triageSignins.ToArray() -Columns @('Time', 'User', 'EventType', 'IP', 'IsIocIP', 'KnownIP', 'Country', 'City', 'AppName', 'ClientApp', 'ErrorCode', 'Category', 'Meaning', 'CAStatus', 'MfaMethod', 'RiskDuring', 'FailureReason', 'CorrelationId')),
                (New-HuntReportSection -Id 'ioc_sp' -Title 'SP from IoC IP' -DefaultSort 'Time' -MaxRows $HtmlMaxRows `
                    -Description 'Service principal sign-ins from the IoC addresses.' -Rows $iocSp),
                (New-HuntReportSection -Id 'ioc_users' -Title 'Users/MI from IoC IP' -DefaultSort 'Time' -MaxRows $HtmlMaxRows `
                    -Description 'Interactive, non-interactive and managed identity sign-ins from the IoC addresses.' -Rows $iocOther),
                (New-HuntReportSection -Id 'secret_failures' -Title 'Secret failures' -DefaultSort 'Time' -MaxRows $HtmlMaxRows `
                    -Description 'AADSTS7000215 / 7000222 from unknown IP addresses.' -Rows $secretFails),
                (New-HuntReportSection -Id 'credentials' -Title 'Credentials' -DefaultSort 'AppName' -MaxRows $HtmlMaxRows `
                    -Description 'All secrets, certificates and federated credentials on apps and service principals.' -Rows $credRows),
                (New-HuntReportSection -Id 'audit' -Title 'Audit' -DefaultSort 'Time' -MaxRows $HtmlMaxRows `
                    -Description 'Changes to credentials, owners and consents on apps.' -Rows $audit)
            )
            foreach ($sec in $sections) {
                Write-HuntLog ("   {0,-16} {1} rows{2}" -f $sec.Id, $sec.Total, (Select-HuntValue $sec.Truncated " (truncated to $HtmlMaxRows)" '')) -Level DEBUG
            }
            New-HuntHtmlReport -Path $htmlPath -Meta $meta -Sections $sections | Out-Null
            $size = (Get-Item $htmlPath).Length
            Write-HuntLog ("HTML report: {0} ({1:N0} KB, {2:N1} s)" -f $htmlPath, ($size / 1KB), $htmlSw.Elapsed.TotalSeconds) -Level OK
        } catch {
            $htmlPath = $null
            Write-HuntLog "HTML report failed: $($_.Exception.Message)" -Level ERROR
            if ($_.InvocationInfo) { Write-HuntLog "   at $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -Level ERROR }
        }
    }



    }
    catch {
        $err = $_
        Write-Host ''
        Write-HuntLog "HUNT FAILED: $($err.Exception.Message)" -Level ERROR
        Write-HuntLog "Exception type: $($err.Exception.GetType().Name)" -Level ERROR
        if ($err.InvocationInfo) {
            $failLine = [string]$err.InvocationInfo.Line
            Write-HuntLog "At line $($err.InvocationInfo.ScriptLineNumber): $($failLine.Trim())" -Level ERROR
        }
        if ($err.ScriptStackTrace) {
            Write-HuntLog "Stack trace:" -Level ERROR
            foreach ($l in ($err.ScriptStackTrace -split "`n")) { Write-HuntLog "   $l" -Level ERROR }
        }
        if ($findings.Count -gt 0) {
            try {
                $findings.ToArray() | Export-Csv (Join-Path $OutputFolder 'FINDINGS_partial.csv') -NoTypeInformation -Encoding UTF8
                Write-HuntLog "Findings so far saved to FINDINGS_partial.csv ($($findings.Count) rows)" -Level WARN
            } catch { }
            if (@($sorted).Count -eq 0) { $sorted = $findings.ToArray() }
        }
        Write-HuntLog "Run time before failure: $(Format-HuntDuration ((Get-Date) - $state.RunStart))" -Level ERROR
    }

    return @{
        Findings      = @($sorted)
        Triage        = $triageSummary.ToArray()
        TriageSignIns = $triageSignins.ToArray()
        Credentials   = @($credRows)
        Audit         = @($audit)
        IocSpSignIns  = @($iocSp)
        IocUserSignIns = @($iocOther)
        SecretFailures = @($secretFails)
        SpSignIns     = $(if ($SaveRawSignIns) { @($spRows) } else { @() })
        ReportPath    = $htmlPath
    }
}
