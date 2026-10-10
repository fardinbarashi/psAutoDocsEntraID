# Autodoc v1.0
function Get-SafeExcelTableName {
    <# Sanitises a string into a valid Excel table name (letters/digits/underscore, tbl_ prefix). #>
    param([Parameter(Mandatory)][string]$Name)

    $safeName = $Name -replace '[^A-Za-z0-9_]', '_'
    if ($safeName -notmatch '^[A-Za-z_]') { $safeName = "T_$safeName" }
    if ($safeName.Length -gt 200) { $safeName = $safeName.Substring(0, 200) }
    return "tbl_$safeName"
}
