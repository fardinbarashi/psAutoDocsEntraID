# Autodoc v1.0
function Write-SvgReadme {
    <#
        Writes a README.md into the SVG output folder describing each generated
        SVG map and what it contains. Called by the SVG report step so a fresh
        description travels with every export.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SvgFolder)

    $readme = @'
# Autodoc - SVG maps

This folder contains the SVG maps generated from the Entra ID (Azure AD) export
(the JSON files under `../json`). Each file is a self-contained diagram - open it
in a browser or import into Visio. Every map has a "Data source" box at the top
showing the path to the JSON file it was built from.

Filenames end with a timestamp (yyyy-MM-dd_HH.mm.ss) so re-runs never overwrite
earlier output.

A PDF copy of every map is written to the `../pdf` folder.

## Files and what they contain

- **App registrations** - one card per app registration: IDs, owners, sign-in
  audience, credentials (secrets/certificates) and their expiry, API permissions
  and recent sign-in activity. Card colour/edge flag expiring or expired
  credentials. Source: `AppRegEnterpriseApps`.

- **Enterprise apps** - every service principal (enterprise application): type,
  origin (own tenant vs Microsoft/gallery), owners, SSO configuration, assignment
  requirement, credentials and sign-in activity. Microsoft first-party apps are
  hidden by default. Green fill = SSO configured; red edge = service principal
  disabled. Source: `EnterpriseApps`.

- **Conditional Access policies** - one card per CA policy: WHO it targets (each
  user on its own line with a bold name + `UPN:` line), APPS, WHEN
  (locations / platforms / client app types), GRANT controls, plus risk and
  session settings and any other populated attribute. Source: `ConditionalAccess`.

- **Domains** - verified custom domains grouped by type
  (Managed / Federated / None); the default domain is tagged.
  Source: `TenantInformationDomain`.

- **Users (types, domains, auth, titles)** - the user population by user type,
  category, domain, authentication methods and job titles, with an account
  summary (enabled/disabled, without mailbox, multi-method). Source:
  `UserInformation`.

- **Groups (all)** - every group with owners, type, membership and settings.
  Source: `EntraGroups`.

- **Group owners** - owners and how many groups each person owns.
  Source: `EntraGroups`.

- **Groups - members by department (list)** - one card per group listing the
  member count per department. Source: `GroupMembers`.

- **Groups x departments - <Letter>** - ONE FILE PER STARTING LETTER. Each file
  is a matrix of every group beginning with that letter (rows) against all
  departments those groups have members in (rotated columns). Cell number =
  members of the group who are in that department; empty = 0. Source:
  `GroupMembers`.

- **RBAC roles by category** - directory role assignments grouped into six
  categories. Each role card lists the people who hold it (filled dot = active,
  hollow dot = eligible / PIM). Privileged roles are marked "P" (and a red
  border), read-only roles "R". Source: `RBAC`.

- **RBAC matrix (roles x people)** - six category matrices: people as rows, roles
  as columns; filled dot = active, hollow dot = eligible; sensitive/privileged
  roles have a red column header. Source: `RBAC`.

- **Licenses (SKUs)** - subscribed license SKUs and how many are consumed.
  Source: `LicensesInformation`.

- **Password reset (SSPR)** - the tenant self-service password reset
  configuration. Source: `PasswordReset`.

- **Service limits & recommendations** - the tenant's current counts (users,
  directory objects, Conditional Access policies, domains, licence subscriptions,
  dynamic and role-assignable groups) shown against the documented Microsoft
  Entra limits, colour-coded by how close each is to its limit. Source: several
  JSON files + `Config/EntraServiceLimits.json`.

- **Organization Department Officelocation chart** - users arranged by manager relationships.
  Source: `UserInformation`.

- **Tenant overview - hub** - a one-page summary: the tenant with a node per area
  (licences, groups, apps, Conditional Access, users, ...) and its headline
  number. Source: several JSON files.

- **Tenant overview - tree** - the whole tenant expanded as a tree: every area and
  each of its items in a column beneath it. Category boxes carry icons; item
  boxes are colour-coded per category. Source: several JSON files.
'@

    if (-not (Test-Path $SvgFolder)) { New-Item -ItemType Directory -Path $SvgFolder -Force | Out-Null }
    $path = Join-Path $SvgFolder 'README.md'
    $readme | Out-File -FilePath $path -Encoding utf8
    Write-Host "  Wrote $path" -ForegroundColor DarkGray
}
