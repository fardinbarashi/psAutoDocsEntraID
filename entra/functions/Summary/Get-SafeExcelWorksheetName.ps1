# Autodoc v1.0
function Get-SafeExcelWorksheetName {
    <# Sanitises a string into a valid Excel worksheet name (<=31 chars, no : \ / ? * [ ]). #>
    param([Parameter(Mandatory)][string]$Name)

    $safeName = $Name -replace '[:\\/\?\*\[\]]', '_'
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "Sheet" }
    if ($safeName.Length -gt 31) { $safeName = $safeName.Substring(0, 31) }
    return $safeName
}
