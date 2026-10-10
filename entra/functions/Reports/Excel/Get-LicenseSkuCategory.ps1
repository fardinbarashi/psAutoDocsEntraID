# Autodoc v1.0
function Get-LicenseSkuCategory {
    <#
        Classifies a SKU as 'Purchased' or 'Free/Viral'.

        Why this matters: free and viral SKUs (Stream, Power BI free, Flow Free,
        viral PowerApps, dev/trial plans) are provisioned with enormous quotas -
        1,000,000 units is common. Charting them next to real licences makes
        every purchased SKU invisible and turns tenant-wide "utilisation" into a
        meaningless fraction of a percent. They are therefore reported separately.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SkuPartNumber,
        [int]$EnabledUnits = 0,
        [int]$FreePoolThreshold = 10000
    )

    if ($SkuPartNumber -match 'FREE|VIRAL|TRIAL|_DEV|_IW|POWER_BI_STANDARD|STREAM') { return 'Free/Viral' }
    if ($EnabledUnits -ge $FreePoolThreshold)                                       { return 'Free/Viral' }
    return 'Purchased'
}
