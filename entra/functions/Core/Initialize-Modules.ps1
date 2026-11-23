# Autodoc v1.0
function Initialize-Modules {
    <#
        Makes sure every module the script needs is installed, then imports them.

        1. Check: every module is checked first, without installing anything.
           Microsoft.Graph* modules must be exactly -TargetGraphVersion; other
           modules only need to be installed.
        2. Ask once: if something is missing or has the wrong version, the list is
           shown and you choose
             [A] install all
             [S] select - answer Y/N per module
             [N] install nothing and continue
        3. Install what was chosen (CurrentUser scope). A Graph module with the
           wrong version has its other versions uninstalled first.
        4. Import every module.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$RequiredModules,
        [Parameter(Mandatory)][version]$TargetGraphVersion
    )

    Write-Host "-> Checking required modules..."

    # ---- 1. check everything first ----
    $todo = [System.Collections.Generic.List[object]]::new()
    foreach ($module in $RequiredModules) {
        $installed = @(Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending)
        $isGraph   = $module -like 'Microsoft.Graph*'
        $status = if (-not $installed) { 'Missing' }
                  elseif ($isGraph -and -not ($installed | Where-Object { $_.Version -eq $TargetGraphVersion })) { 'WrongVersion' }
                  else { 'OK' }
        if ($status -eq 'OK') {
            Write-Host ("   {0,-55} OK ({1})" -f $module, $installed[0].Version) -ForegroundColor DarkGreen
            continue
        }
        $todo.Add([pscustomobject]@{
            Module    = $module
            IsGraph   = $isGraph
            Status    = $status
            Installed = (@($installed.Version | Select-Object -Unique) -join ', ')
            Target    = $(if ($isGraph) { "$TargetGraphVersion" } else { 'latest' })
        })
    }

    # ---- 2. ask once ----
    if ($todo.Count -gt 0) {
        Write-Host ""
        Write-Host "-> $($todo.Count) module(s) need to be installed:" -ForegroundColor Yellow
        foreach ($t in $todo) {
            $what = if ($t.Status -eq 'Missing') { 'missing' } else { "installed $($t.Installed), needs $($t.Target)" }
            Write-Host ("   {0,-55} {1}" -f $t.Module, $what) -ForegroundColor Yellow
        }
        Write-Host ""
        $choice = ''
        while ($choice -notmatch '^[ASN]$') {
            $choice = (Read-Host "Install? [A] All  [S] Select per module  [N] None").Trim().ToUpper()
            if ($choice -eq 'J') { $choice = 'A' }   # Swedish "ja" = all
        }

        $chosen = switch ($choice) {
            'A' { @($todo) }
            'N' { @() }
            'S' {
                @(foreach ($t in $todo) {
                    $a = ''
                    while ($a -notmatch '^[YNJ]$') { $a = (Read-Host "  Install $($t.Module) ($($t.Target))? (Y/N)").Trim().ToUpper() }
                    if ($a -in 'Y', 'J') { $t }
                })
            }
        }
        $skipped = @($todo | Where-Object { $chosen -notcontains $_ })
        foreach ($s in $skipped) { Write-Host "   Skipped: $($s.Module) - parts of the script that need it may fail." -ForegroundColor DarkYellow }

        # ---- 3. install what was chosen ----
        $i = 0
        foreach ($t in $chosen) {
            $i++
            Write-Host "`n-> [$i/$($chosen.Count)] Installing $($t.Module) ($($t.Target))..." -ForegroundColor Yellow
            try {
                if ($t.IsGraph) {
                    if (Get-Module -Name $t.Module) { Remove-Module -Name $t.Module -Force -ErrorAction SilentlyContinue }
                    if ($t.Status -eq 'WrongVersion') {
                        Get-InstalledModule -Name $t.Module -AllVersions -ErrorAction SilentlyContinue |
                            Uninstall-Module -Force -ErrorAction SilentlyContinue
                    }
                    Install-Module -Name $t.Module -RequiredVersion $TargetGraphVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                }
                else {
                    Install-Module -Name $t.Module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                }
                Write-Host "   Installed." -ForegroundColor DarkGreen
            }
            catch {
                Write-Host "   Failed to install $($t.Module): $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "   Continuing script..." -ForegroundColor Yellow
            }
        }
    }

    # ---- 4. import ----
    Write-Host ""
    Write-Host "-> Importing modules..."
    foreach ($module in $RequiredModules) {
        try {
            if ($module -like 'Microsoft.Graph*' -and (Get-Module -ListAvailable -Name $module | Where-Object { $_.Version -eq $TargetGraphVersion })) {
                Import-Module $module -RequiredVersion $TargetGraphVersion -ErrorAction Stop
            }
            else { Import-Module $module -ErrorAction Stop }
            Write-Host "   $module" -ForegroundColor DarkGreen
        }
        catch {
            Write-Host "   Could not import $module : $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   Continuing script..." -ForegroundColor Yellow
        }
    }

    Write-Host "`nAll module checks completed!" -ForegroundColor Green
}
