# Autodoc v1.0
function Get-EntraLimitStatus {
    <#
        Compares the tenant's actual counts against the documented Microsoft
        Entra service limits. The limits themselves come from an editable config
        file (Config/EntraServiceLimits.json, transcribed from Microsoft Learn);
        the current values are counted from the exported JSON.

        Returns one object per metric: Area, Metric, Current, Limit, Type,
        PercentUsed, Status (OK/Watch/Near/Over), Color, Note, SourceUrl.

        Thresholds: OK < 60 %, Watch 60-85 %, Near >= 85 %, Over above the limit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder,
        [string]$ConfigPath
    )

    if (-not $ConfigPath) {
        # The limits config lives in the cache (Entra\files\cache), refreshed
        # from Microsoft on each collect/report. Three levels up from this file
        # (Functions\Reports\Common) is the Entra root.
        $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))
        $ConfigPath = Join-Path $entraRoot 'Config\EntraServiceLimits.json'
    }
    if (-not (Test-Path $ConfigPath)) {
        # Not cached yet (e.g. a fresh clone or a standalone call) - create it
        # from the built-in baseline, refreshed live when reachable.
        try { Update-EntraServiceLimits -OutputPath $ConfigPath | Out-Null } catch { }
    }
    if (-not (Test-Path $ConfigPath)) {
        Write-Host "Service-limits config not found: $ConfigPath" -ForegroundColor Yellow
        return
    }
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

    # ---- gather the raw counts once ----
    $ti    = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $eapp  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps')
    $caObj = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess') | Select-Object -First 1
    $doms  = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformationDomain')
    $lic   = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    $gInfo = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroupsBasicInfo') | Select-Object -First 1
    $groups = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')

    $usersN = if ($ti -and $ti.Users) { [int]$ti.Users } else { @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation').Count }
    $groupsN = if ($ti -and $ti.Groups) { [int]$ti.Groups } else { $groups.Count }
    $aregN = if ($ti -and $ti.AppRegistrations) { [int]$ti.AppRegistrations } else { @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps').Count }
    $caPolicies = if ($caObj -and $caObj.PSObject.Properties.Name -contains 'ConditionalAccessPolicies') { @($caObj.ConditionalAccessPolicies) } else { @() }

    # ---- current value per metric key ----
    $current = @{
        DirectoryObjects      = $usersN + $groupsN + $aregN + $eapp.Count
        Users                 = $usersN
        ConditionalAccessPolicies = $caPolicies.Count
        ManagedDomains        = @($doms | Where-Object { "$($_.Type)" -eq 'Managed' }).Count
        FederatedDomains      = @($doms | Where-Object { "$($_.Type)" -eq 'Federated' }).Count
        LicenseSubscriptions  = $lic.Count
        DynamicGroups         = if ($gInfo) { [int]$gInfo.DynamicGroups } else { @($groups | Where-Object { "$($_.GroupCategory)" -eq 'Dynamic' }).Count }
        RoleAssignableGroups  = @($groups | Where-Object { "$($_.IsAssignableToRole)" -eq 'True' }).Count
    }

    $results = foreach ($m in $config.metrics) {
        if (-not $current.ContainsKey($m.key)) { continue }
        $cur = [int]$current[$m.key]
        $lim = [int]$m.limit
        $pct = if ($lim -gt 0) { [math]::Round($cur / $lim * 100, 1) } else { 0 }
        $status =
            if ($cur -gt $lim) { 'Over' }
            elseif ($pct -ge 85) { 'Near' }
            elseif ($pct -ge 60) { 'Watch' }
            else { 'OK' }
        $color = switch ($status) {
            'OK'    { '#2E7D32' }
            'Watch' { '#F9A825' }
            'Near'  { '#C62828' }
            'Over'  { '#B71C1C' }
        }
        [pscustomobject]@{
            Area        = $m.area
            Metric      = $m.label
            Current     = $cur
            Limit       = $lim
            Type        = $m.type
            PercentUsed = $pct
            Status      = $status
            Color       = $color
            Note        = $m.note
            SourceUrl   = $config.sourceUrl
        }
    }
    ,@($results)
}
