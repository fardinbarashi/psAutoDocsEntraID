# Autodoc v1.0
function Build-EntraSvgMaps {
    <#
        Builds every SVG map for an export into <export>\Report\svg and returns
        that folder. Shared by the SVG report and the HTML report so the list of
        maps lives in one place.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    # remove any earlier set of maps first, so running SVG + HTML (+ Word) together,
    # or re-running, leaves a single current set instead of duplicates
    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (Test-Path $svgDir) { Get-ChildItem -Path $svgDir -Filter '*.svg' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }

    Build-KpiDashboardSvg         -SourceFolder $SourceFolder | Out-Null   # KPI charts (Excel graphs as doughnuts)
    Build-LicenceMapSvg           -SourceFolder $SourceFolder | Out-Null   # licences
    Build-OrgMapSvg               -SourceFolder $SourceFolder | Out-Null   # organisation
    Build-GroupsMapSvg            -SourceFolder $SourceFolder | Out-Null   # groups by category + mail
    Build-GroupOwnerMapSvg        -SourceFolder $SourceFolder | Out-Null   # group ownership
    Build-GroupDeptListSvg        -SourceFolder $SourceFolder | Out-Null   # group -> department (list)
    Build-GroupDeptMatrixSvg      -SourceFolder $SourceFolder | Out-Null   # group x department (matrix)
    Build-ConditionalAccessMapSvg -SourceFolder $SourceFolder | Out-Null   # CA policies
    Build-AppRegMapSvg            -SourceFolder $SourceFolder | Out-Null   # app registrations
    Build-EnterpriseAppMapSvg     -SourceFolder $SourceFolder | Out-Null   # enterprise apps (SSO)
    Build-DomainMapSvg            -SourceFolder $SourceFolder | Out-Null   # verified domains
    Build-RbacMapSvg              -SourceFolder $SourceFolder | Out-Null   # RBAC roles (bands)
    Build-RbacMatrixSvg           -SourceFolder $SourceFolder | Out-Null   # RBAC role x principal matrix
    Build-UsersMapSvg             -SourceFolder $SourceFolder | Out-Null   # user breakdowns
    Build-PasswordResetMapSvg     -SourceFolder $SourceFolder | Out-Null   # SSPR configuration
    Build-LimitsMapSvg            -SourceFolder $SourceFolder | Out-Null   # service limits & recommendations
    Build-OverviewMapSvg          -SourceFolder $SourceFolder -Style hub  | Out-Null   # one-page overview (hub)
    Build-OverviewMapSvg          -SourceFolder $SourceFolder -Style tree | Out-Null   # full expanded tree

    return (Join-Path $SourceFolder 'Report\svg')
}
