<p align="center">
  <img src="https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/img/logo.png" alt="psAutoDocs" width="360">
</p>

<h3 align="center">Automated documentation of Microsoft Entra ID as Excel, SVG and PDF</h3>

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-7.4%2B-5391FE?logo=powershell&logoColor=white" alt="PowerShell 7.4+">
  <img src="https://img.shields.io/badge/Microsoft%20Graph-read--only-2F7CF6" alt="Microsoft Graph read-only">
  <img src="https://img.shields.io/badge/Entra%20ID-documentation-0078D4" alt="Entra ID">
  <img src="https://img.shields.io/badge/output-Excel%20%7C%20SVG%20%7C%20PDF-2EA44F" alt="Excel | SVG | PDF">
</p>

psAutoDocs connects to Microsoft Graph, saves a **read-only** snapshot of the tenant as JSON and turns it into documentation you can share: an Excel KPI workbook, SVG maps of the tenant and PDF copies. Documents can be rebuilt from an export at any time without contacting the tenant again.

![psAutoDocs interface](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/gui1.jpg)

---

## Contents

| Get started | Output | Reference |
|---|---|---|
| [Features](#features) | [Excel KPI workbook](#excel-kpi-workbook) | [Graph scopes](#graph-scopes) |
| [How it works](#how-it-works) | [SVG maps](#svg-maps) | [Project layout](#project-layout) |
| [Requirements](#requirements) | [Export folder](#export-folder) | [Performance](#performance) |
| [Quick start](#quick-start) | | |

---

## Features

| Feature | What you get |
|---|---|
| **Read-only collection** | Tenant, domains, licences, users, groups, app registrations, enterprise apps, Conditional Access, RBAC and SSPR |
| **Excel KPI workbook** | A dashboard and one analysed sheet per area, with charts and cards |
| **SVG maps** | Hub overview, tenant tree and maps per area. Each file opens in a browser |
| **PDF** | Copies of the SVG maps |
| **Raw data** | JSON, CSV and Excel per area |
| **Rebuild offline** | Generate new documents from an existing export without Graph calls |

---

## How it works

1. **Download Data** connects to Microsoft Graph and saves each selected area to `Entra\files\exports\<timestamp>\`.
2. **Generate Documents** builds the chosen formats from the JSON in an export.

![Excel dashboard](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/Excel%20dashboard.jpg)

---

## Requirements

| Component | Requirement |
|---|---|
| **PowerShell** | 7.4.0 or later |
| **Modules** | `ImportExcel` and `Microsoft.Graph.*` (installed automatically if missing) |
| **Permissions** | Read-only Graph scopes, see [Graph scopes](#graph-scopes) |
| **Excel KPI** | Microsoft Excel to view the workbook |

---

## Quick start

**1. Download the repository**

```powershell
git clone https://github.com/fardinbarashi/psAutoDocsEntraID.git
cd psAutoDocsEntraID
```

**2. Start the interface**

```powershell
.\Autodoc.ps1
```

**3. Collect data**

Sign in when prompted, select the areas on **Download Data** and run the collection.

**4. Generate documents**

Switch to **Generate Documents**, select the formats and click **Run**.

---

## Excel KPI workbook

| Sheet | Content |
|---|---|
| Dashboard | Summary of the tenant and the export |
| 01 Licenses | Enabled, consumed and free units per licence |
| 02 Users per licence | Users holding each licence |
| 03 Assigned vs Free | Assigned and available units per SKU |
| 04 Per department | Users and licences per department |
| 05 Top 10 SKU | The most used licences |
| 06 Per office | Users and licences per office |
| 07 Multi-licence users | Users with more than one licence |
| 08 Groups overview | Groups by type and source |
| 09 Conditional Access | Policies and their state |
| 10 App registrations | Owners, credential expiry and risk findings |
| 11 Enterprise apps | Single sign-on, permissions and risk |
| 12 RBAC | Active and eligible role assignments |
| 13 Limits & recommendations | Usage against Entra service limits |

<details>
<summary><b>Show Excel screenshots</b></summary>

**01 Licenses**
![Licenses](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/Excel1.jpg)

**02 Users per licence**
![Users per licence](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/Excel2.jpg)

**03 Assigned vs Free**
![Assigned vs Free](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel3.jpg)

**04 Per department**
![Per department](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel%204.jpg)

**05 Top 10 SKU**
![Top 10 SKU](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel5.jpg)

**06 Per office**
![Per office](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel6.jpg)

**07 Multi-licence users**
![Multi-licence users](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel7.jpg)

**08 Groups overview**
![Groups overview](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel8.jpg)

**09 Conditional Access**
![Conditional Access](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel9.jpg)

**10 App registrations**
![App registrations](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel10.jpg)

**11 Enterprise apps**
![Enterprise apps](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel11.jpg)

**12 RBAC**
![RBAC](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel12.jpg)

**13 Limits & recommendations**
![Limits and recommendations](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/excel13.jpg)

</details>

---

## SVG maps

![Tenant overview hub](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EntraID%20-%20Hub.jpg)

| Map | Content |
|---|---|
| Tenant overview, hub | One page with a node per area |
| Tenant overview, tree | The whole tenant as a tree |
| Users | User types, domains, authentication methods and job titles |
| Groups | Owners, type, membership and settings |
| Group owners | Owners and how many groups each owns |
| Groups by department | Members per group grouped by department |
| Groups x departments | One matrix per starting letter |
| Organization chart | Users by manager |
| Licenses | Subscribed SKUs and consumption |
| Conditional Access | One card per policy |
| App registrations | IDs, owners, credential expiry and risk |
| Enterprise apps | Type, single sign-on and permissions |
| RBAC | Roles by category and role matrix |
| Domains | Verified domains by type |
| SSPR | Self-service password reset settings |
| Service limits | Usage against Entra service limits |

<details>
<summary><b>Show SVG screenshots</b></summary>

**Tenant overview, tree**
![Tenant overview tree](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EntraID%20-%20Overview%20tree.jpg)

**Users**
![Users 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Entra%20ID%20-%20Users%20(types%2C%20domains%2C%20auth%2C%20titles)%201.jpg)
![Users 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Entra%20ID%20-%20Users%20(types%2C%20domains%2C%20auth%2C%20titles)%202.jpg)

**Groups**
![Groups](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Group%20All.jpg)

**Group owners**
![Group owners](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/GroupOwners.jpg)

**Groups by department**
![Groups by department](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Group%20member%20by%20department.jpg)

**Groups x departments**
![Groups x departments](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/GroupDepartemnsLetter.jpg)

**Organization chart**
![Organization chart](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Org%20Department%20office%20location%20chart.jpg)

**Licenses**
![Licenses 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Lic1.jpg)
![Licenses 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Lic2.jpg)

**Conditional Access**
![Conditional Access 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA1.jpg)
![Conditional Access 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA2.jpg)
![Conditional Access 3](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA3.jpg)
![Conditional Access 4](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/CA4.jpg)

**App registrations**
![App registrations 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%201.jpg)
![App registrations 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%202.jpg)
![App registrations 3](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Appreg%203.jpg)

**Enterprise apps**
![Enterprise apps 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%201.jpg)
![Enterprise apps 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%202.jpg)
![Enterprise apps 3](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/EnterpriseApps%203.jpg)

**RBAC roles by category**
![RBAC by category 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20roles%20by%20category%201.jpg)
![RBAC by category 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20roles%20by%20category%202.jpg)

**RBAC matrix**
![RBAC matrix 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%201.jpg)
![RBAC matrix 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%202.jpg)
![RBAC matrix 3](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/RBAC%20matrix%20roles%20x%20people%203.jpg)

**Domains**
![Domains 1](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Domain%201.jpg)
![Domains 2](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Domain%202.jpg)

**Self-service password reset**
![SSPR](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/SSPR.jpg)

**Service limits**
![Service limits](https://raw.githubusercontent.com/fardinbarashi/psAutoDocsEntraID/refs/heads/main/githubRepoContentDeleteIfYouWant/screenshots/svg/Servicelimit.jpg)

</details>

Sample SVG files from a lab tenant are in [data-mock (Lab Contoso)](https://github.com/fardinbarashi/psAutoDocsEntraID/tree/main/githubRepoContentDeleteIfYouWant/data-mock%20(Lab%20Contoso)).

---

## Export folder

```text
Entra\files\exports\<timestamp>\
  json\            Raw Graph data, one file per area
  csv\             The same data as CSV
  excel\           Raw data per area
  tenantsummary\   Short summary
  excelkpi\        KPI workbook
  svg\             Vector maps
  pdf\             PDF copies of the maps
  README.md        Description of every folder and JSON file
```

---

## Reference

### Graph scopes

All scopes are read-only, for example `Directory.Read.All`, `Group.Read.All`, `User.Read.All`, `Application.Read.All`, `Policy.Read.All`, `RoleManagement.Read.Directory`, `Domain.Read.All`, `Organization.Read.All`, `Reports.Read.All`, `AuditLog.Read.All` and `UserAuthenticationMethod.Read.All`.

<details>
<summary><b>Modules</b></summary>

`ImportExcel`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Users`, `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.SignIns`, `Microsoft.Graph.Identity.Governance`, `Microsoft.Graph.Security`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Beta.Reports`

</details>

### Project layout

```text
Autodoc.ps1                 Launcher (GUI and CLI)
config\CollectorPicker.xaml WPF interface
entra\
  Config\                   Settings and paths
  Functions\                Collectors and report builders
  Settings\                 Templates and script settings
  files\                    Cache and exports
```

### Performance

| Area | Note |
|---|---|
| Users, groups, app registrations | The slowest areas in large tenants |
| Groups | Faster when Users is also selected, since member data is reused |
| Overview tree | Can be very tall in large tenants. Opens best in Firefox |

---

<p align="center">
  Made by <a href="https://www.linkedin.com/in/fardin-barashi-a56310a2/">Fardin Barashi</a> · <a href="LICENSE">License</a>
</p>
