# psAutoDocs
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/img/logo.png)
![](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/gui1.jpg)
![](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/Excel%20dashboard.jpg)
![](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/data-mock%20(Lab%20Contoso)/svg/EntraID%20-%20Hub%202026-08-02_20.01.53.svg) 


## Table of Contents
- [About](#about)
- [Notes](#notes)
- [Performance notes](#performance-notes)
- [News](#news)
- [System requirements](#system-requirements)
- [How it works](#how-it-works)
- [Project layout](#project-layout)
- [Screenshots](#screenshots)
- [Roadmap](#roadmap)

---
# Automated documentation : 

- **Microsoft Entra ID**.
```
Autodoc connects to Microsoft Graph, exports a read-only snapshot of a tenant's
configuration as JSON, and turns that snapshot into ready-to-share
documentation: 
- an Excel KPI workbook ( Excel KPI only works in microsoft excel. ) 
  dashboard plus one analysed sheet per area (users,
  groups, licenses, Conditional Access, app registrations, enterprise apps,
  RBAC) with charts and cards
  
- SVG maps
  Vector diagrams of the tenant, each self-contained (open in a
  browser): a one-page hub overview, an overview tree, and per-area maps for
  users, groups, group owners, RBAC (roles and matrix), licenses, Conditional
  Access, app registrations, enterprise apps, domains, self-service password
  reset (SSPR), and service limits.
  
- PDF ( copies of svg ) 

> All Graph access is read-only
  What gets collected: 
  Tenant information and domains, licenses (subscribed SKUs), users, groups
  (several views plus members and welcome-email settings), app registrations,
  enterprise applications, Conditional Access policies, RBAC role assignments,
  and self-service password reset (SSPR) configuration

  How it works
  Autodoc has two steps:
   1. Collect (Download Data) – connects to Microsoft Graph and pulls each
   selected dataset into a timestamped export folder under
   Entra\files\exports\, as JSON (and CSV, and per-area Excel).
   2. Report (Generate Documents) – builds the chosen formats (Excel KPI, SVG,…) 
   from the JSON in an existing export. No Graph calls are needed for this
   step, so you can regenerate documents any time without touching the tenant.

   Output structure
   Each export folder (Entra\files\exports\<timestamp>\) contains only the
   ub-folders that were produced:
   <export>\
    json\            Raw Graph data, one file per area (the source of masterdata for the script)
    csv\             Raw Graph data, The same data as json
    excel\           Raw Graph data, Per-area workbooks (raw data)
    tenantsummary\   A short high-level summary
   
    excelkpi\        The KPI workbook (dashboard, charts, analysed sheets)
    svg\             Vector maps (each map self-contained; has its own README)
    pdf\             PDF copies of the svg-maps

    README.md        Describes every folder and documents each JSON file
   
   word\            Word report

```
---
## Notes

---
## Performance notes
When downloading data
- **Users, Groups and App registrations are the heavy sections** and can take a
  while on large tenants.
- **Groups runs faster when Users is selected too**, because it reuses the user
  data already collected instead of querying each member again.
- The **Report** step never calls Graph — once an export exists you can rebuild
  documents as often as you like without re-collecting.

SVG data
- In large tenants the **Overview tree** map can get very tall; open it in 
**Firefox** for the smoothest rendering. ![mapsSvgFiles](https://github.com/fardinbarashi/psAutoDocs/tree/main/githubRepoContentDeleteIfYouWant/data-mock%20(Lab%20Contoso)/svg)

---

## News
---

## System requirements
### Runtime

| Requirement | Detail |
|-------------|--------|
| PowerShell | 7.4.0 or later (Core) |
| PowerShell modules | `ImportExcel` and the `Microsoft.Graph.*` modules below (installed automatically if missing) |

### Modules
`ImportExcel`, `Microsoft.Graph.Authentication`,
`Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Groups`,
`Microsoft.Graph.Users`, `Microsoft.Graph.Applications`,
`Microsoft.Graph.Identity.SignIns`, `Microsoft.Graph.Identity.Governance`,
`Microsoft.Graph.Security`, `Microsoft.Graph.Reports`,
`Microsoft.Graph.Beta.Reports`.

### Graph scopes (read-only)
All requested scopes are `*.Read.All`, for example `Directory.Read.All`,
`Group.Read.All`, `User.Read.All`, `Application.Read.All`, `Policy.Read.All`,
`RoleManagement.Read.Directory`, `Domain.Read.All`, `Organization.Read.All`,
`Reports.Read.All`, `AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`.


---

## Getting started
```powershell
git clone https://github.com/fardinbarashi/psAutoDocs.git
cd psAutoDocs
# Launch the interface (sign in when prompted)

```
The first run signs you in to Microsoft Graph with `Connect-MgGraph` and the
read-only scopes above. Pick the datasets on **Download Data**, run the collect,
then switch to **Generate Documents**, tick the formats you want, and click Run.

---

## Project layout
```
Autodoc.ps1                 Launcher (GUI + CLI dispatcher)
Config\CollectorPicker.xaml The WPF interface
Entra\
   Config\                  Settings, paths
   Functions\
      Core\                 Collect/report orchestration, Graph connect, registry
      Groups\ Users\ ...    Per-area collectors
      Reports\
         Excel\             KPI workbook and sheets
         Word\              Word templates + resolver (builder in progress)
         Common\            Shared helpers (KPI, limits, PDF, glossary)
   Settings\                Word templates, script settings
   files\                   Cache and exports
```
---
## Screenshots

### Excel

#### Dashboard:
high-level summary of the tenant and the export.
![Dashboard](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/Excel%20dashboard.jpg)

#### 01 Licenses — subscribed SKU inventory: 
enabled, consumed and free units per licence.
![Licenses — subscribed SKU inventory](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/Excel1.jpg)

#### 02 Users per licence:
how many users hold each licence.
![Users per licence](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/Excel2.jpg)

#### 03 Assigned vs Free:
assigned versus still-available units per SKU.
![Assigned vs Free](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel3.jpg)

#### 04 Per department:
users and licences broken down by department.
![04 Per department](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel%204.jpg)

#### 05 Top 10 SKU:
the ten most-used licences.
![Top 10 SKU](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel5.jpg)

#### 06 Per office:
users and licences broken down by office location.
![Per office](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel6.jpg)

#### 07 Multi-licence users:
users holding more than one licence.
![Multi-licence users](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel7.jpg)

#### 08 Groups overview:
group counts by type and source (security, Microsoft 365, dynamic, cloud vs on-prem).
![Groups overview](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel8.jpg)

#### 09 Conditional Access:
Conditional Access policies and their state (enabled, report-only, disabled).
![Conditional Access](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel9.jpg)

#### 10 App registrations:
app registrations with owners, credential expiry and risk findings.
![App registrations](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel10.jpg)

#### 11 Enterprise apps:
enterprise applications (service principals): single sign-on, permissions and risk.
![Enterprise apps](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel11.jpg)

#### 12 RBAC:
directory role assignments (active vs eligible, privileged roles).
![RBAC](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel12.jpg)

#### 13 Limits & recommendations:
current usage against the documented Microsoft Entra service limits, colour-coded.
![Limits & recommendations](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/excel13.jpg)


---  
### SVG

#### Tenant overview - tree:
the whole tenant expanded as a tree; every area with its items listed. (Very large in big tenants — best viewed in Firefox.)
![Tenant overview - tree](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EntraID%20-%20Overview%20tree.jpg)

#### Tenant overview - hub:
A one-page summary: the tenant in the centre with a node per area (users, groups, roles, apps, domains, and more).
![Tenant overview - hub](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EntraID%20-%20Hub.jpg)

#### Users (types, domains, auth, titles):
the user population by user type, domain, authentication methods and job titles.
![Users, types, domains, auth, titles](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Entra%20ID%20-%20Users%20(types%2C%20domains%2C%20auth%2C%20titles)%201.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Entra%20ID%20-%20Users%20(types%2C%20domains%2C%20auth%2C%20titles)%202.jpg)

#### Groups(all):
every group with its owners, type, membership and settings.
![Groups all](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Group%20All.jpg)

#### Group owners:
owners and how many groups each person owns.
![Group owners](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/GroupOwners.jpg)

#### Groups - members by department (list):
one card per group listing its members grouped by department.
![Groups - members by department list](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Group%20member%20by%20department.jpg)

#### Groups x departments - <Letter>:
one file per starting letter: a matrix of those groups against departments.
![Groups x departments - Letter](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/GroupDepartemnsLetter.jpg)

#### Organization chart:
users arranged by manager relationships.
![Organization chart](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Org%20Department%20office%20location%20chart.jpg)

#### Licenses (SKUs):
subscribed licence SKUs and how many are consumed.
![Licenses SKUs](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Lic1.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Lic2.jpg)

#### Conditional Access policies:
one card per policy, who it targets, what it grants or blocks, and its state.
![Conditional Access policies](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA1.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA2.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA3.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA4.jpg)

#### App registrations:
one card per app registration: IDs, owners, credential expiry and risk findings.
![App registrations](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%201.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%202.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%203.jpg)

#### Enterprise apps:
Enterprise application (service principal): one card per app, type, single sign-on and permissions.
![Enterprise application](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%201.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%202.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%203.jpg)

#### RBAC roles by category:
directory role assignments grouped into categories.
![RBAC roles category](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20roles%20by%20category%201.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20roles%20by%20category%202.jpg)


#### RBAC matrix (roles x people):
category matrices with people as rows and roles as columns.
![RBAC roles matrix](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%201.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%202.jpg)
![](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%202.jpg)

#### Domains:
verified custom domains grouped by type.
![Domains](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Domain%201.jpg)
![](https://github.com/fardinbarashi/psAutoDocs/blob/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Domain%202.jpg)


#### Password reset (SSPR):
the tenant's self-service password-reset configuration.
![Password reset SSPR](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/SSPR.jpg)

#### Service limits & recommendations:
current usage against the documented Microsoft Entra service limits
![Service limits recommendations](https://raw.githubusercontent.com/fardinbarashi/psAutoDocs/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Servicelimit.jpg)


---

## Roadmap



---

## Author
**Fardin Barashi** — <https://github.com/fardinbarashi/psAutoDocs>


