# Autodoc v1.0
function Get-MsLearnConfig {
    <#
        Reads AzureDocs.json and downloads each data source into the cache
        folder (e.g. the Microsoft licensing reference CSV).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ConfigPath,
        [Parameter(Mandatory)][string]$CacheFolder
    )

    if (-not (Test-Path -Path $ConfigPath)) {
        Write-Host "MS Learn config not found: $ConfigPath" -ForegroundColor Red
        return
    }

    $config = Get-Content $ConfigPath | ConvertFrom-Json

    foreach ($source in $config.dataSources) {
        $url = if ($source.downloadUrl) { $source.downloadUrl } else { $source.url }
        if (-not $url -or -not $source.fileName) {
            Write-Host "Skipping invalid entry: $($source.name)" -ForegroundColor Red
            continue
        }

        $outFile = Join-Path $CacheFolder $source.fileName
        Write-Host "`nDownloading [$($source.name)]..." -ForegroundColor Cyan
        Write-Host "URL: $url" -ForegroundColor DarkGray

        try {
            Invoke-WebRequest -Uri $url -OutFile $outFile -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
            Write-Host "--> Saved to: $outFile" -ForegroundColor DarkGreen
        }
        catch {
            Write-Host "FAILED [$($source.name)]: $($_.Exception.Message)" -ForegroundColor Red
            continue
        }
    }
}
