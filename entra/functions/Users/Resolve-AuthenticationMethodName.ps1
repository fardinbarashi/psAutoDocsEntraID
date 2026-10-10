# Autodoc v1.0
function Resolve-AuthenticationMethodName {
    <# Maps a Graph auth-method @odata.type to a friendly name. #>
    param([string]$ODataType)

    switch ($ODataType) {
        "#microsoft.graph.passwordAuthenticationMethod"                { "Password" }
        "#microsoft.graph.phoneAuthenticationMethod"                   { "Phone" }
        "#microsoft.graph.emailAuthenticationMethod"                   { "Email" }
        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod"  { "Microsoft Authenticator" }
        "#microsoft.graph.fido2AuthenticationMethod"                   { "FIDO2" }
        "#microsoft.graph.softwareOathAuthenticationMethod"            { "Software OATH" }
        "#microsoft.graph.temporaryAccessPassAuthenticationMethod"     { "Temporary Access Pass" }
        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { "Windows Hello for Business" }
        default                                                        { $ODataType }
    }
}
