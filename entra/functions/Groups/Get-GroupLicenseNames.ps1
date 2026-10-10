# Autodoc v1.0
function Get-GroupLicenseNames {
    <# Resolves a group's assigned licenses to friendly names. #>
    param(
        $Group,
        [hashtable]$LicenseRef,
        [hashtable]$GraphRef
    )

    $licenseNames = @()
    foreach ($lic in $Group.AssignedLicenses) {
        $resolved = Resolve-LicenseName -AssignedLicense $lic -LicenseRef $LicenseRef -GraphRef $GraphRef
        if ($resolved) { $licenseNames += $resolved }
    }
    return ($licenseNames | Sort-Object -Unique) -join "; "
}
