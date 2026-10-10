# Autodoc v1.0
function Invoke-AutodocCollect {
    <#
        STEP 1 - Runs a full collection: module check, Graph connect, reference
        data, then the selected collectors. Extracted from the launcher so the
        GUI can call it directly and stay open while it runs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EntraRoot,
        [Parameter(Mandatory)][hashtable]$Settings,
        [Parameter(Mandatory)]$Registry,
        [Parameter(Mandatory)][string[]]$Selected,
        [string[]]$Formats = @('Csv','Json','Excel')
    )

    if (-not $Selected -or $Selected.Count -eq 0) { Write-Host "No sections selected." -ForegroundColor Yellow; return }
    if (-not $Formats  -or $Formats.Count  -eq 0) { Write-Host "No file formats selected."   -ForegroundColor Yellow; return }

    Write-Host "Sections: $($Selected -join ', ')" -ForegroundColor Cyan
    Write-Host "Formats : $($Formats  -join ', ')" -ForegroundColor Cyan

    $Paths = New-AutodocPaths -Root $EntraRoot
    $Paths.Formats = @($Formats)
    Start-Transcript -Path $Paths.TranscriptFile -Force | Out-Null

    try {
        Initialize-Modules -RequiredModules $Settings.RequiredModules `
                           -TargetGraphVersion ([version]$Settings.TargetGraphVersion)

        Connect-EntraGraph -Scopes $Settings.Scopes | Out-Null

        Get-MsLearnConfig -ConfigPath $Paths.MsLearnConfig -CacheFolder $Paths.Cache

        try {
            $licenseRef = Get-LicenseReferenceData -Path $Paths.LicenseReferenceFile
        }
        catch {
            Write-Host "License reference not loaded - name resolution will be limited." -ForegroundColor Yellow
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow
            $licenseRef = $null
        }
        $graphRef = Get-GraphSkuData

        $context = @{
            Paths      = $Paths
            Settings   = $Settings
            LicenseRef = $licenseRef
            GraphRef   = $graphRef
            Collected  = [ordered]@{}
        }
        Invoke-CollectorRun -Registry $Registry -Selected $Selected -Context $context | Out-Null

        # Keep the service-limits config current (in Entra\Config), live from the
        # Microsoft docs when reachable. A refresh failure must not stop a collect.
        try {
            $limitsConfig = Join-Path (Split-Path (Split-Path $PSScriptRoot)) 'Config/EntraServiceLimits.json'
            Update-EntraServiceLimits -OutputPath $limitsConfig | Out-Null
        }
        catch { Write-Host "  Service limits: refresh skipped ($($_.Exception.Message))." -ForegroundColor DarkYellow }

        Write-ExportReadme -ExportFolder $Paths.Export | Out-Null   # document the export

        Write-Host ""
        Write-Host "Collection finished. Output: $($Paths.Export)" -ForegroundColor Green
    }
    finally {
        Stop-Transcript | Out-Null
    }
}
