# Autodoc v1.0
function Invoke-AutodocHunt {
    <#
        Runs the workload identity hunt on its own (Tools tab, or
        .\Autodoc.ps1 -Mode Hunt). It is kept out of the normal collection:
        it reads sign-in logs, needs IdentityRiskyUser.Read.All and is run on
        suspicion or as a recurring check, not every time the tenant is documented.

        - Settings: Entra\Config\WorkloadIdentityHunt.json. -HuntSettings (from
          the tab) is written back to that file first, so the GUI and headless
          runs use the same values.
        - Target: the hunt's JSON datasets go into -TargetFolder (an existing
          export), so "Security and improvements" can combine them with the
          documentation data. -NewFolder, or no export at all, creates a new
          export folder instead.
        - -BuildReport builds "Security and improvements" from that export
          straight after the hunt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EntraRoot,
        [Parameter(Mandatory)][hashtable]$Settings,
        [string]$TargetFolder,
        [switch]$NewFolder,
        [hashtable]$HuntSettings,
        [switch]$BuildReport,
        [string]$Language = 'sv',
        [string]$ConsultantLogo,
        [string]$CustomerLogo
    )

    $configPath = Join-Path $EntraRoot 'Config\WorkloadIdentityHunt.json'
    if ($HuntSettings -and $HuntSettings.Count) { Save-HuntSettings -Path $configPath -Values $HuntSettings }

    $useExisting = (-not $NewFolder) -and $TargetFolder -and (Test-Path $TargetFolder)
    $Paths = if ($useExisting) { New-AutodocPaths -Root $EntraRoot -ExportFolder $TargetFolder } else { New-AutodocPaths -Root $EntraRoot }
    $Paths.Formats = @('Csv','Json','Excel')
    $Paths.TranscriptFile = Join-Path $Paths.Logs "AutodocHunt - $($Paths.RunStamp).txt"

    Start-Transcript -Path $Paths.TranscriptFile -Force | Out-Null
    try {
        Write-Host "Workload identity hunt (read-only)" -ForegroundColor Cyan
        Write-Host "Export folder: $($Paths.Export)$(if (-not $useExisting) { '  (new)' })" -ForegroundColor Cyan

        $huntModules = if ($Settings.HuntModules) { @($Settings.HuntModules) } else { @('ImportExcel', 'Microsoft.Graph.Authentication') }
        Initialize-Modules -RequiredModules $huntModules -TargetGraphVersion ([version]$Settings.TargetGraphVersion) | Out-Null
        Connect-EntraGraph -Scopes @($Settings.HuntScopes) | Out-Null

        Export-WorkloadIdentityHunt -Paths $Paths -ConfigPath $configPath | Out-Null
        Write-ExportReadme -ExportFolder $Paths.Export | Out-Null

        if ($BuildReport) {
            Write-Host ""
            $rx = @{ SourceFolder = $Paths.Export; Language = $Language }
            if ($ConsultantLogo) { $rx['ConsultantLogo'] = $ConsultantLogo }
            if ($CustomerLogo)   { $rx['CustomerLogo'] = $CustomerLogo }
            Build-EntraSecurityReport @rx | Out-Null
        }

        Write-Host ""
        Write-Host "Hunt finished. Output: $($Paths.Export)" -ForegroundColor Green
        return $Paths.Export
    }
    catch {
        Write-Host "Hunt failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
    finally {
        Write-Host "Log: $($Paths.TranscriptFile)" -ForegroundColor DarkGray
        Stop-Transcript | Out-Null
    }
}
