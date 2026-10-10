# Autodoc v1.0
function Get-MeasureSum {
    <# Sums a numeric property across a dataset, returning 0 for null/empty input. #>
    param(
        $Data,
        [Parameter(Mandatory)][string]$PropertyName
    )

    if ($null -eq $Data) { return 0 }
    $sum = ($Data | Measure-Object -Property $PropertyName -Sum -ErrorAction SilentlyContinue).Sum
    if ($null -eq $sum) { return 0 }
    return $sum
}
