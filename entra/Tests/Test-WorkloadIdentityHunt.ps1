# Autodoc v1.0
<#
    Self-test for the workload identity hunt. Needs no tenant and no sign-in.

    Every case runs in a child PowerShell process against a test double of
    Microsoft.Graph.Authentication (Tests\modules) that generates Graph-like
    data: 14 000+ service principal sign-ins, IoC hits, every account verdict,
    an expired paging token ("410 Gone"), empty results and single rows.
    Nothing is sent to Microsoft Graph and nothing outside Tests\out is written.

        pwsh -NoProfile -File .\Entra\Tests\Test-WorkloadIdentityHunt.ps1

    Result: PASS/FAIL per case. Output is kept in Tests\out for inspection.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$here      = $PSScriptRoot
$entraRoot = Split-Path $here -Parent
$outDir    = Join-Path $here 'out'
$engine    = (Get-Process -Id $PID).Path
if ([System.IO.Path]::GetFileNameWithoutExtension([string]$engine) -ne 'pwsh') {
    $found = Get-Command pwsh -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $found) { throw "Could not find pwsh to run the test cases" }
    $engine = $found.Source
}
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$fullCategories = @(
    'IoC IP - service principal sign-in',
    'Secret failure followed by SUCCESS (same app + IP)',
    'Invalid/expired secret from unknown IP',
    'Token sweep: ARM + Key Vault + Storage + Graph',
    'Account triage: COMPROMISED',
    'Account triage: PASSWORD_KNOWN',
    'Account triage: ACTIVE_ATTACK',
    'Account triage: SPRAY',
    'Account: security info (MFA method) changed in window',
    'Account: successful sign-in from new IP around attack time',
    'Audit change from IoC IP',
    'Secret on service principal object'
)

# Args is PowerShell source for a hashtable splatted onto Invoke-WorkloadIdentityHunt.
# Mode 'Collector' runs the Autodoc collector (settings file + dataset export) instead.
$cases = @(
    @{ Name = 'full-string-gone'; Dates = 'string';   Scenario = 'full';  Gone = '1'; Html = $true;  Mode = 'Engine'; Expect = $fullCategories
       Args = "@{ TriageUsers = @('zed@test.se'); KnownIPs = @('10.0.0.0/8'); Quiet = `$true }" },
    @{ Name = 'full-datetime';    Dates = 'datetime'; Scenario = 'full';  Gone = '0'; Html = $true;  Mode = 'Engine'; Expect = ($fullCategories + 'One IP -> 7 service principals')
       Args = "@{}" },
    @{ Name = 'empty-tenant';     Dates = 'string';   Scenario = 'empty'; Gone = '0'; Html = $true;  Mode = 'Engine'; Expect = @()
       Args = "@{ Quiet = `$true }" },
    @{ Name = 'skip-inventory';   Dates = 'string';   Scenario = 'full';  Gone = '0'; Html = $false; Mode = 'Engine'; Expect = @('Account triage: COMPROMISED')
       Args = "@{ SkipInventory = `$true; NoHtml = `$true; Quiet = `$true; DaysBack = 5; SliceHours = 6 }" },
    @{ Name = 'raw-sign-ins';     Dates = 'string';   Scenario = 'full';  Gone = '0'; Html = $false; Mode = 'Engine'; Expect = @('Account triage: COMPROMISED')
       Args = "@{ SaveRawSignIns = `$true; NoHtml = `$true; Quiet = `$true; SkipInventory = `$true }"; RawCsv = $true },
    @{ Name = 'collector';        Dates = 'string';   Scenario = 'full';  Gone = '0'; Html = $true;  Mode = 'Collector'; Expect = @('Account triage: COMPROMISED', 'Secret on service principal object')
       Args = "@{}" }
)

$runner = @'
param($EntraRoot, $CaseDir, $Mode, $ArgsText)
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Get-ChildItem (Join-Path $EntraRoot 'Functions\Hunt') -Filter '*.ps1' | ForEach-Object { . $_.FullName }
. (Join-Path $EntraRoot 'Functions\Core\Read-AutodocJson.ps1')
Start-Transcript -Path (Join-Path $CaseDir 'transcript.log') -Force | Out-Null
try {
    if ($Mode -eq 'Collector') {
        . (Join-Path $EntraRoot 'Functions\Core\Export-InventoryData.ps1')
        . (Join-Path $EntraRoot 'Functions\Collectors\Export-WorkloadIdentityHunt.ps1')
        $paths = @{ Root = $EntraRoot; Export = $CaseDir; FileDate = 'test'; Formats = @('Csv', 'Json')
                    Csv = (Join-Path $CaseDir 'rawDataCsv'); Json = (Join-Path $CaseDir 'rawDataJson'); Excel = (Join-Path $CaseDir 'rawDataExcel') }
        foreach ($f in @($paths.Csv, $paths.Json)) { New-Item -ItemType Directory -Path $f -Force | Out-Null }
        $r = Export-WorkloadIdentityHunt -Paths $paths
        foreach ($set in 'WorkloadIdentityHunt', 'HuntAccountTriage', 'HuntCredentials', 'HuntAuditChanges', 'HuntIocSpSignIns', 'HuntIocUserSignIns', 'HuntSecretFailures') {
            if (-not (Test-Path (Join-Path $paths.Json "$set test.json"))) { Write-Host "CHECK FAILED: dataset JSON missing: $set" }
        }
        foreach ($raw in 'HuntSpSignIns test.json', 'HuntTriageSignIns test.json') {
            if (Test-Path (Join-Path $paths.Json $raw)) { Write-Host "CHECK FAILED: raw sign-ins saved without SaveRawSignIns: $raw" }
        }
        if (Test-Path (Join-Path $CaseDir 'WorkloadIdentityHunt\sp_signins_all.csv')) { Write-Host 'CHECK FAILED: raw sign-in CSV written without SaveRawSignIns' }
    }
    else {
        $splat = Invoke-Expression $ArgsText
        $r = Invoke-WorkloadIdentityHunt -OutputFolder (Join-Path $CaseDir 'WorkloadIdentityHunt') @splat
    }
    Write-Host "HUNT RETURNED: $(@($r.Findings).Count) findings"
}
finally { Stop-Transcript | Out-Null }
'@
$runnerPath = Join-Path $outDir 'run-case.ps1'
Set-Content -Path $runnerPath -Value $runner -Encoding UTF8

$savedPath = $env:PSModulePath
$failed = 0
Write-Host "Engine: $engine (PowerShell $($PSVersionTable.PSVersion))"
try {
    $env:PSModulePath = (Join-Path $here 'modules') + [System.IO.Path]::PathSeparator + $savedPath
    foreach ($case in $cases) {
        $caseDir = Join-Path $outDir $case.Name
        $huntDir = Join-Path $caseDir 'WorkloadIdentityHunt'
        $logFile = Join-Path $caseDir 'transcript.log'
        New-Item -ItemType Directory -Path $caseDir -Force | Out-Null
        $env:MOCK_DATES    = $case.Dates
        $env:MOCK_SCENARIO = $case.Scenario
        $env:MOCK_GONE     = $case.Gone

        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        & $engine -NoProfile -File $runnerPath -EntraRoot $entraRoot -CaseDir $caseDir -Mode $case.Mode -ArgsText $case.Args | Out-Null
        $timer.Stop()

        $problems = New-Object System.Collections.Generic.List[string]
        if (-not (Test-Path $logFile)) {
            [void]$problems.Add('no transcript written')
        } else {
            $log = Get-Content -Path $logFile
            foreach ($line in $log) {
                if ($line -match 'Federated credential read failed for Err-App|FIC Err-App') { continue }   # provoked on purpose
                if ($line -match 'HUNT FAILED|HTML report failed|Query failed|read failed|Could not read|MOCK: unhandled|CHECK FAILED|\] ERROR ') {
                    [void]$problems.Add("log: $line")
                }
            }
            if (-not ($log -match 'Total run time')) { [void]$problems.Add('hunt did not reach the end') }
            if (-not ($log -match 'HUNT RETURNED')) { [void]$problems.Add('no result returned') }
        }
        $findingsFile = Join-Path $huntDir 'FINDINGS.csv'
        if (-not (Test-Path $findingsFile)) {
            [void]$problems.Add('FINDINGS.csv missing')
        } else {
            $categories = New-Object System.Collections.Generic.List[string]
            foreach ($row in (Import-Csv -Path $findingsFile)) { [void]$categories.Add([string]$row.Category) }
            foreach ($wanted in $case.Expect) {
                if (-not $categories.Contains($wanted)) { [void]$problems.Add("expected finding missing: $wanted") }
            }
        }
        $rawCsv = Join-Path $huntDir 'sp_signins_all.csv'
        if ($case.RawCsv -and -not (Test-Path $rawCsv)) { [void]$problems.Add('SaveRawSignIns did not write sp_signins_all.csv') }
        if (-not $case.RawCsv -and (Test-Path $rawCsv)) { [void]$problems.Add('sp_signins_all.csv written without SaveRawSignIns') }
        $htmlCount = 0
        if (Test-Path $huntDir) { $htmlCount = (Get-ChildItem -Path $huntDir -Filter '*.html' | Measure-Object).Count }
        if ($case.Html -and $htmlCount -ne 1) { [void]$problems.Add("expected 1 HTML report, found $htmlCount") }
        if (-not $case.Html -and $htmlCount -ne 0) { [void]$problems.Add("expected no HTML report, found $htmlCount") }

        if ($problems.Count -eq 0) {
            Write-Host ("PASS  {0,-18} {1,5:N0} s" -f $case.Name, $timer.Elapsed.TotalSeconds) -ForegroundColor Green
        } else {
            $failed++
            Write-Host ("FAIL  {0,-18} {1,5:N0} s" -f $case.Name, $timer.Elapsed.TotalSeconds) -ForegroundColor Red
            foreach ($p in $problems) { Write-Host "      $p" -ForegroundColor Red }
        }
    }
}
finally {
    $env:PSModulePath = $savedPath
    Remove-Item Env:\MOCK_DATES, Env:\MOCK_SCENARIO, Env:\MOCK_GONE -ErrorAction SilentlyContinue
}

Write-Host ''
if ($failed -eq 0) { Write-Host "ALL $($cases.Count) CASES PASSED. Output: $outDir" -ForegroundColor Green }
else { Write-Host "$failed of $($cases.Count) CASES FAILED. Output: $outDir" -ForegroundColor Red }
