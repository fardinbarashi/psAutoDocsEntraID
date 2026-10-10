# Autodoc v1.0
function Get-AllGroupOwners {
    <# Returns the owners of a group as simple objects (id / name / UPN / mail). #>
    param([string]$GroupId)

    $owners = Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction SilentlyContinue
    $owners | ForEach-Object {
        [pscustomobject]@{
            Id                = $_.Id
            DisplayName       = $_.AdditionalProperties.displayName
            UserPrincipalName = $_.AdditionalProperties.userPrincipalName
            Mail              = $_.AdditionalProperties.mail
        }
    }
}
