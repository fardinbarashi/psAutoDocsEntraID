# Autodoc v1.0
function Get-UserAuthenticationMethodInfo {
    <# Returns the distinct registered auth methods for a user as a joined string + count. #>
    param([string]$UserId)

    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $UserId -All -ErrorAction Stop
        $methodNames = foreach ($method in $methods) {
            $odataType = $method.AdditionalProperties['@odata.type']
            Resolve-AuthenticationMethodName -ODataType $odataType
        }
        $methodNames = $methodNames | Where-Object { $_ } | Sort-Object -Unique

        return @{
            AuthenticationMethods     = $methodNames -join "; "
            AuthenticationMethodCount = $methodNames.Count
        }
    }
    catch {
        return @{
            AuthenticationMethods     = ""
            AuthenticationMethodCount = 0
        }
    }
}
