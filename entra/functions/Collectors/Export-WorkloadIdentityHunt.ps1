# Autodoc v1.0
function Export-WorkloadIdentityHunt {
    <#
        Workload identity hunt as Autodoc datasets. Reads its settings from
        Entra\Config\WorkloadIdentityHunt.json (IoC IPs, known IPs, extra
        accounts to triage, thresholds), runs Invoke-WorkloadIdentityHunt and
        writes:
          rawDataJson / rawDataCsv / rawDataExcel
            WorkloadIdentityHunt   findings, sorted by severity
            HuntAccountTriage      one row per triaged account (verdict, action)
            HuntCredentials        secrets, certificates, federated credentials
            HuntAuditChanges       credential / owner / consent changes
            HuntIocSpSignIns       service principal sign-ins from IoC IPs
            HuntIocUserSignIns     user / managed identity sign-ins from IoC IPs
            HuntSecretFailures     AADSTS7000215 / 7000222 from unknown IPs
            HuntSpSignIns,         raw sign-in rows - only when SaveRawSignIns
            HuntTriageSignIns      is true in the settings (JSON + CSV, no Excel)
          <export>\WorkloadIdentityHunt\   the hunt's own CSV files + HTML report

        The report step ("Security and improvements") reads these JSON files;
        it never calls Graph itself.
        Read-only. Requires AuditLog.Read.All, Directory.Read.All,
        Application.Read.All, Policy.Read.All and IdentityRiskyUser.Read.All.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Paths,
        [string]$ConfigPath
    )

    $Section = 'Workload identity hunt'
    try {
        Write-Host "Start $Section" -ForegroundColor Yellow

        if (-not $ConfigPath) { $ConfigPath = Join-Path $Paths.Root 'Config\WorkloadIdentityHunt.json' }
        $huntArgs = @{}
        if (Test-Path $ConfigPath) {
            $cfg = Read-AutodocJson -Path $ConfigPath
            $known = (Get-Command Invoke-WorkloadIdentityHunt).Parameters
            foreach ($p in $cfg.PSObject.Properties) {
                if (-not $known.ContainsKey($p.Name) -or $p.Name -in 'OutputFolder','ReportFolder') { continue }
                if ($known[$p.Name].ParameterType -eq [switch]) {
                    if ($p.Value -eq $true) { $huntArgs[$p.Name] = $true }
                }
                elseif ($known[$p.Name].ParameterType.IsArray) { $huntArgs[$p.Name] = @($p.Value | Where-Object { "$_" }) }
                else { $huntArgs[$p.Name] = $p.Value }
            }
            Write-Host "Settings: $ConfigPath"
        }
        else {
            Write-Host "Settings file not found ($ConfigPath) - using the built-in defaults." -ForegroundColor DarkYellow
        }

        $huntFolder = Join-Path $Paths.Export 'WorkloadIdentityHunt'
        $result = Invoke-WorkloadIdentityHunt -OutputFolder $huntFolder @huntArgs

        # The hunt writes several datasets under one registry key, so the
        # per-dataset selection filter is lifted for these exports.
        $huntPaths = $Paths.Clone()
        $huntPaths.SelectedOutputs = $null
        $formats = if ($Paths.Formats) { @($Paths.Formats) } else { @('Csv','Json','Excel') }
        $saveSet = {
            param($Rows, [string]$BaseName, [switch]$NoExcel)
            # dates as yyyy-MM-dd HH:mm:ss text, same as the hunt's CSV files
            $rows = @(foreach ($r in @($Rows | Where-Object { $null -ne $_ })) {
                $o = [ordered]@{}
                foreach ($pr in $r.PSObject.Properties) { $o[$pr.Name] = if ($pr.Value -is [datetime]) { $pr.Value.ToString('yyyy-MM-dd HH:mm:ss') } else { $pr.Value } }
                [pscustomobject]$o
            })
            $p = $huntPaths.Clone()
            $p.Formats = if ($NoExcel) { @($formats | Where-Object { $_ -ne 'Excel' }) } else { $formats }
            if ($rows.Count) { Export-InventoryData -Data $rows -BaseName $BaseName -Paths $p -JsonDepth 6 }
            elseif ($p.Formats -contains 'Json') {
                # an empty result is still a result: "[]" tells the report step the hunt ran
                Set-Content -LiteralPath (Join-Path $p.Json "$BaseName $($p.FileDate).json") -Value '[]' -Encoding UTF8
            }
            Write-Host ("  {0,-22} {1} rows" -f $BaseName, $rows.Count)
        }

        Write-Host "Export workload identity hunt datasets to files..."
        & $saveSet $result.Findings       'WorkloadIdentityHunt'
        & $saveSet $result.Triage         'HuntAccountTriage'
        & $saveSet $result.Credentials    'HuntCredentials'
        & $saveSet $result.Audit          'HuntAuditChanges'
        & $saveSet $result.IocSpSignIns   'HuntIocSpSignIns'
        & $saveSet $result.IocUserSignIns 'HuntIocUserSignIns'
        & $saveSet $result.SecretFailures 'HuntSecretFailures'
        if ($huntArgs['SaveRawSignIns']) {
            & $saveSet $result.SpSignIns     'HuntSpSignIns' -NoExcel
            & $saveSet $result.TriageSignIns 'HuntTriageSignIns' -NoExcel
        }

        $findings = @($result.Findings)
        Write-Host "Done. $($findings.Count) findings. Output: $huntFolder" -ForegroundColor DarkGreen
        if ($result.ReportPath) { Write-Host "HTML report: $($result.ReportPath)" -ForegroundColor DarkGreen }
        Write-Host "End $Section" -ForegroundColor Green

        return @{ WorkloadIdentityHunt = $findings; WorkloadIdentityCredentials = @($result.Credentials) }
    }
    catch {
        Write-Host "ERROR on $Section" -ForegroundColor Red
        Write-Host "ERROR:" $_.Exception.Message
    }
}
