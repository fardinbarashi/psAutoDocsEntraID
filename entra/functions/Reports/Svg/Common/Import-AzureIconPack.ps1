function Import-AzureIconPack {
    <#
        Fills files\cache\Azure icons\Icons from Microsoft's official Azure icon pack.

        Download it yourself from
        https://learn.microsoft.com/en-us/azure/architecture/icons/
        (the page links Azure_Public_Service_Icons_V<n>.zip) and point this at
        the zip. The pack is published under Microsoft's own terms - it is not
        redistributed with Autodoc, which is why this imports rather than ships.

        The pack names files like "10221-icon-service-Azure-Active-Directory.svg",
        so they are matched to SKU part numbers through the table below and
        copied out under the SKU name the map builder looks for. Both the SVG
        and the PNG are taken when the pack has them; the maps prefer the SVG.

        Anything unmatched is left alone - drop files in by hand at any time,
        named after the SKU part number.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [string]$IconFolder,
        [hashtable]$Map,
        [switch]$ListUnmatched
    )

    if (-not (Test-Path $ZipPath)) { Write-Host "Icon pack not found: $ZipPath" -ForegroundColor Red; return }

    if (-not $IconFolder) {
        $IconFolder = Get-MapIconFolder -BuilderRoot $PSScriptRoot
    }
    New-Item -Path $IconFolder -ItemType Directory -Force | Out-Null

    # SKU part number -> a fragment of the icon file name in the pack.
    # Extend this freely; the fragment match is case-insensitive.
    if (-not $Map) {
        $Map = @{
            'ENTERPRISEPACK'                  = 'Azure-Active-Directory'
            'SPE_E3'                          = 'Azure-Active-Directory'
            'SPE_E5'                          = 'Azure-Active-Directory'
            'SPE_F1'                          = 'Azure-Active-Directory'
            'EMS'                             = 'Enterprise-Applications'
            'AAD_PREMIUM'                     = 'Azure-Active-Directory'
            'AAD_PREMIUM_P2'                  = 'Azure-Active-Directory'
            'IDENTITY_THREAT_PROTECTION'      = 'Defender'
            'DEFENDER_ENDPOINT_P1'            = 'Defender'
            'WIN10_VDA_E5'                    = 'Virtual-Desktop'
            'EXCHANGESTANDARD'                = 'Exchange'
            'POWER_BI_PRO'                    = 'Power-BI'
            'POWER_BI_STANDARD'               = 'Power-BI'
            'FLOW_FREE'                       = 'Power-Automate'
            'POWERAPPS_VIRAL'                 = 'Power-Apps'
            'POWERAPPS_DEV'                   = 'Power-Apps'
            'POWERAPPS_PER_USER'              = 'Power-Apps'
            'PROJECTPREMIUM'                  = 'Project'
            'VISIOCLIENT'                     = 'Visio'
            'VISIOONLINE_PLAN1'               = 'Visio'
            'STREAM'                          = 'Media'
            'M365EDU_A3_FACULTY'              = 'Azure-Active-Directory'
            'M365EDU_A3_STUUSEBNFT'           = 'Azure-Active-Directory'
            'STANDARDWOFFPACK_FACULTY'        = 'Azure-Active-Directory'
            'STANDARDWOFFPACK_STUDENT'        = 'Azure-Active-Directory'
            'Microsoft_365_Copilot'           = 'Cognitive-Services'
            'Microsoft_365_E3_Extra_Features' = 'Azure-Active-Directory'
        }
    }

    $stage = Join-Path ([IO.Path]::GetTempPath()) ("iconpack_" + [guid]::NewGuid().ToString('N'))
    try {
        Write-Host "Extracting icon pack..." -ForegroundColor Yellow
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $stage)

        $all = @(Get-ChildItem -Path $stage -File -Recurse | Where-Object { $_.Extension -in '.svg', '.png' })
        Write-Host "  $($all.Count) icon files in the pack" -ForegroundColor DarkGray

        $copied = 0; $missed = @()
        foreach ($sku in $Map.Keys) {
            $fragment = $Map[$sku]
            $hits = @($all | Where-Object { $_.BaseName -like "*$fragment*" })
            if (-not $hits) { $missed += "$sku (no icon matching '$fragment')"; continue }

            foreach ($ext in '.svg', '.png') {
                $pick = $hits | Where-Object { $_.Extension -eq $ext } | Sort-Object Length -Descending | Select-Object -First 1
                if ($pick) {
                    Copy-Item -Path $pick.FullName -Destination (Join-Path $IconFolder "$sku$ext") -Force
                    $copied++
                }
            }
        }

        Write-Host "Done. $copied files written to $IconFolder" -ForegroundColor Green
        if ($missed) {
            Write-Host "  $($missed.Count) SKUs had no match:" -ForegroundColor Yellow
            if ($ListUnmatched) { $missed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray } }
            else { Write-Host "    run with -ListUnmatched to see them" -ForegroundColor DarkGray }
        }
    }
    catch { Write-Host "Icon import failed: $($_.Exception.Message)" -ForegroundColor Red }
    finally { if (Test-Path $stage) { Remove-Item $stage -Recurse -Force } }
}
