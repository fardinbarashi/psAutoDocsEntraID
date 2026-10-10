# Autodoc v1.0
function Get-DataRowCount {
    <# Row count that treats $null as 0 and a single object as 1. #>
    param($Data)

    if ($null -eq $Data) { return 0 }
    return @($Data).Count
}
