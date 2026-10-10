# Autodoc v1.0
function Get-GroupSource {
    <# Cloud vs on-prem synced, based on OnPremisesSyncEnabled. #>
    param($OnPremisesSyncEnabled)
    if ($OnPremisesSyncEnabled) { "Windows Server" } else { "Cloud" }
}
