# Autodoc v1.0
function Initialize-Modules {
    <#
        Checks, installs (with prompt for Graph version pinning) and imports
        every module the script needs. Behaviour matches the original loop.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$RequiredModules,
        [Parameter(Mandatory)][version]$TargetGraphVersion
    )

    Write-Host "-> Start checking required modules..."

    foreach ($module in $RequiredModules) {
        Write-Host "`nChecking module: $module" -ForegroundColor DarkGray

        $installedModules = @(Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending)
        $isGraphModule    = $module -like 'Microsoft.Graph*'

        if ($isGraphModule) {
            $hasCorrectVersion = $installedModules | Where-Object { $_.Version -eq $TargetGraphVersion }
            if (-not $hasCorrectVersion) {
                if ($installedModules.Count -gt 0) {
                    Write-Host "---> Microsoft Graph module found, but not version $TargetGraphVersion." -ForegroundColor Yellow
                    Write-Host "     Installed version(s): $($installedModules.Version -join ', ')" -ForegroundColor Yellow
                }
                else { Write-Host "---> Microsoft Graph module not found." -ForegroundColor Red }

                $answer = Read-Host "Do you want to uninstall existing versions and install $module version $TargetGraphVersion? (Y/N)"
                if ($answer -match '^(Y|y|J|j)$') {
                    Write-Host "---> Installing correct version for $module..." -ForegroundColor Yellow
                    try {
                        if (Get-Module -Name $module) { Remove-Module -Name $module -Force -ErrorAction SilentlyContinue }
                        if ($installedModules.Count -gt 0) {
                            Get-InstalledModule -Name $module -AllVersions -ErrorAction SilentlyContinue |
                                Uninstall-Module -Force -ErrorAction SilentlyContinue
                        }
                        Install-Module -Name $module -RequiredVersion $TargetGraphVersion -Scope CurrentUser -Force -AllowClobber
                        Write-Host "----> $module version $TargetGraphVersion installed." -ForegroundColor DarkGreen
                    }
                    catch {
                        Write-Host "----> Failed to install $module version $TargetGraphVersion." -ForegroundColor Red
                        Write-Host "     Error: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "     Continuing script..." -ForegroundColor Yellow
                    }
                }
                else { Write-Host "---> Skipping installation of correct version. Continuing script..." -ForegroundColor Yellow }
            }
            else { Write-Host "---> Correct Microsoft Graph module version found: $TargetGraphVersion" -ForegroundColor Green }
        }
        else {
            if (-not $installedModules) {
                Write-Host "---> Module not found! Installing..." -ForegroundColor Red
                try { Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber }
                catch {
                    Write-Host "----> Failed to install $module." -ForegroundColor Red
                    Write-Host "     Error: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "     Continuing script..." -ForegroundColor Yellow
                }
            }
            else { Write-Host "---> Module found." -ForegroundColor Green }
        }

        try {
            Write-Host "---> Importing module..." -ForegroundColor Yellow
            Import-Module $module -ErrorAction Stop
            Write-Host "----> Module imported." -ForegroundColor DarkGreen
        }
        catch {
            Write-Host "----> Could not import module: $module" -ForegroundColor Red
            Write-Host "     Error: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "     Continuing script..." -ForegroundColor Yellow
        }
    }

    Write-Host "`nAll module checks completed!" -ForegroundColor Green
}
