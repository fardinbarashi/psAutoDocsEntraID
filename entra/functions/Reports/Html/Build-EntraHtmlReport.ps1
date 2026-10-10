# Autodoc v1.0
function Build-EntraHtmlReport {
    <#
        Builds one self-contained HTML report that gathers every SVG map for an
        export into a single page: a header, a contents menu, and each map inlined
        (the SVGs are already self-contained, so the file opens anywhere with no
        external files). Reads the latest map of each kind from <export>\Report\svg
        and writes to <export>\Report\html.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceFolder
    )

    $svgDir = Join-Path $SourceFolder 'Report\svg'
    if (-not (Test-Path $svgDir)) {
        Write-Host "No SVG maps found to build an HTML report - build the SVG maps first." -ForegroundColor Yellow
        return
    }

    # latest file per map (strip the trailing timestamp, keep newest)
    $stampRx = ' \d{4}-\d{2}-\d{2}_\d{2}\.\d{2}\.\d{2}$'
    $maps = Get-ChildItem -Path $svgDir -Filter '*.svg' -ErrorAction SilentlyContinue |
        Group-Object { $_.BaseName -replace $stampRx, '' } |
        ForEach-Object { $_.Group | Sort-Object LastWriteTime | Select-Object -Last 1 }
    if (-not $maps) {
        Write-Host "No SVG maps found in $svgDir." -ForegroundColor Yellow
        return
    }

    # preferred order (overview first), then anything else alphabetically
    $order = @('Dashboard', 'Hub', 'Overview tree', 'Users', 'Groups (all)', 'Group owners',
        'Groups - members by department', 'Groups x departments', 'Organization',
        'Licenses', 'Conditional Access', 'App registrations', 'Enterprise apps',
        'RBAC roles', 'RBAC matrix', 'Domains', 'Password reset', 'Service limits')
    function Get-MapTitle([string]$base) {
        ($base -replace $stampRx, '' -replace '^Entra ID - ', '' -replace '^EntraID - ', '').Trim()
    }
    # The "Groups x departments - <letter>" matrices are a very granular drill-down
    # (one huge file per starting letter) and would bloat the single HTML enormously,
    # so they are left as standalone SVGs and not inlined here.
    $maps = $maps | Where-Object { (Get-MapTitle $_.BaseName) -notlike 'Groups x departments*' }
    function Get-OrderIndex([string]$title) {
        for ($i = 0; $i -lt $order.Count; $i++) { if ($title -like "$($order[$i])*") { return $i } }
        return 999
    }
    $maps = $maps | Sort-Object @{ E = { Get-OrderIndex (Get-MapTitle $_.BaseName) } }, @{ E = { Get-MapTitle $_.BaseName } }

    # header data
    $tenant = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'TenantInformation') | Select-Object -First 1
    $tName  = if ($tenant) { "$($tenant.TenantName)" } else { 'Microsoft Entra ID' }
    $tDom   = if ($tenant) { "$($tenant.PrimaryDomain)" } else { '' }
    $tEnv   = ("$(Get-TenantHybridLabel -Tenant $tenant)" -replace '^Environment:\s*', '')
    $today  = Get-Date -Format 'yyyy-MM-dd'
    function HtmlEnc([string]$s) { [System.Net.WebUtility]::HtmlEncode("$s") }

    # small summary counts (best effort)
    # enterprise-app count consistent with the hub/enterprise map: exclude Microsoft's
    # built-in service principals (IsMicrosoftApp = Yes), which otherwise inflate the total
    $eappFull = try { Get-EnterpriseAppFullData -SourceFolder $SourceFolder } catch { $null }
    $eappN = if ($eappFull -and $eappFull.Apps) {
        @($eappFull.Apps | Where-Object { "$($_.IsMicrosoftApp)".Trim() -ne 'Yes' }).Count
    }
    else {
        @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps' | Where-Object { "$($_.IsMicrosoftApp)".Trim() -ne 'Yes' }).Count
    }
    $cards = @()
    if ($tenant) {
        if ($tenant.Users)            { $cards += @{ L = 'Users';             V = $tenant.Users } }
        if ($tenant.Groups)           { $cards += @{ L = 'Groups';            V = $tenant.Groups } }
        if ($tenant.AppRegistrations) { $cards += @{ L = 'App registrations'; V = $tenant.AppRegistrations } }
        if ($eappN)                   { $cards += @{ L = 'Enterprise apps';   V = $eappN } }
    }

    # ---- filterable data tables (licenses, groups) ----
    function New-HtmlTableInner($id, $title, $cols, $rows) {
        $t = [System.Text.StringBuilder]::new()
        $n = @($rows).Count
        [void]$t.Append("  <div class='tbltools'><input id='$id-q' oninput=""tblFilter('$id')"" placeholder='Filter $n rows...'><span class='count'><b id='$id-count'>$n</b> of $n</span>")
        [void]$t.Append("<div class='colpick'><button type='button' onclick=""colMenu('$id')"">Columns &#9662;</button><div class='colmenu' id='$id-cols'>")
        for ($ci = 0; $ci -lt $cols.Count; $ci++) { [void]$t.Append("<label><input type='checkbox' checked onchange=""colToggle('$id-t',$ci,this.checked)""> $([System.Net.WebUtility]::HtmlEncode($cols[$ci].H))</label>") }
        [void]$t.AppendLine('</div></div></div>')
        [void]$t.AppendLine("  <div class='tblwrap'><table class='data' id='$id-t'><thead><tr>")
        for ($ci = 0; $ci -lt $cols.Count; $ci++) { [void]$t.Append("<th onclick=""tblSort('$id-t',$ci,this)"">$([System.Net.WebUtility]::HtmlEncode($cols[$ci].H))</th>") }
        [void]$t.AppendLine('</tr></thead><tbody>')
        foreach ($r in $rows) {
            [void]$t.Append('<tr>')
            foreach ($col in $cols) {
                $val = $r.($col.K)
                if ($col.Num) { [void]$t.Append("<td data-v='$([double]$val)' style='text-align:right'>$('{0:N0}' -f [double]$val)</td>") }
                else { [void]$t.Append("<td>$([System.Net.WebUtility]::HtmlEncode("$val"))</td>") }
            }
            [void]$t.Append('</tr>')
        }
        [void]$t.AppendLine('</tbody></table></div>')
        return $t.ToString()
    }

    $tableSections = @()

    $lic = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'LicensesInformation')
    if ($lic.Count) {
        $licRows = foreach ($l in $lic) {
            $en = [int]$l.enabledUnits; $co = [int]$l.consumedUnits
            [pscustomobject]@{ SKU = "$($l.skuPartNumber)"; Enabled = $en; Consumed = $co; Free = [int]$l.freeUnits
                UsedPct = if ($en -gt 0) { [math]::Round(100 * $co / $en) } else { 0 }; Status = "$($l.capabilityStatus)" }
        }
        $licRows = $licRows | Sort-Object -Property @{ E = { $_.Consumed } } -Descending
        $licCols = @(@{H='SKU';K='SKU'}, @{H='Enabled';K='Enabled';Num=$true}, @{H='Consumed';K='Consumed';Num=$true}, @{H='Free';K='Free';Num=$true}, @{H='Used %';K='UsedPct';Num=$true}, @{H='Status';K='Status'})
        $tableSections += @{ Id = 'tbl-lic'; Title = 'Licenses (table)'; Inner = (New-HtmlTableInner 'tbl-lic' 'Licenses (table)' $licCols $licRows) }
    }

    $grp = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EntraGroups')
    if ($grp.Count) {
        $grpRows = foreach ($g in $grp) {
            [pscustomobject]@{
                Name       = "$($g.DisplayName)"
                Category   = "$($g.GroupCategory)"
                Membership = if (@($g.GroupTypes) -contains 'DynamicMembership') { 'Dynamic' } else { 'Assigned' }
                Source     = if ($g.OnPremisesSyncEnabled) { 'On-premises' } else { 'Cloud' }
                Members    = [int]$g.MemberCount
                Mail       = if ($g.MailEnabled) { 'Yes' } else { 'No' }
                Role       = if ($g.IsAssignableToRole) { 'Yes' } else { 'No' }
            }
        }
        $grpRows = $grpRows | Sort-Object Name
        $grpCols = @(@{H='Name';K='Name'}, @{H='Category';K='Category'}, @{H='Membership';K='Membership'}, @{H='Source';K='Source'}, @{H='Members';K='Members';Num=$true}, @{H='Mail';K='Mail'}, @{H='Role-assignable';K='Role'})
        $tableSections += @{ Id = 'tbl-grp'; Title = 'Groups (table)'; Inner = (New-HtmlTableInner 'tbl-grp' 'Groups (table)' $grpCols $grpRows) }
    }

    function YN($x) { if ($x -is [bool]) { if ($x) { 'Yes' } else { 'No' } } elseif ("$x" -match '^(true|1|yes)$') { 'Yes' } elseif ("$x" -match '^(false|0|no|none|)$') { 'No' } else { "$x" } }
    function Trunc($s, $n) { $s = "$s"; if ($s.Length -gt $n) { $s.Substring(0, $n) + [char]0x2026 } else { $s } }
    function Dt($s) { $s = "$s"; if ($s -match '^\d{4}-\d{2}-\d{2}') { $s.Substring(0, 10) } else { $s } }

    # Organization - users by department / office / manager
    $usr = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'UserInformation')
    if ($usr.Count) {
        $usrRows = foreach ($u in $usr) {
            [pscustomobject]@{ Name = "$($u.DisplayName)"; JobTitle = "$($u.JobTitle)"; Department = "$($u.Department)"
                Office = "$($u.OfficeLocation)"; Manager = "$($u.ManagerDisplayName)"; Type = "$($u.UserType)"; Enabled = (YN $u.AccountEnabled) }
        }
        $usrRows = $usrRows | Sort-Object Department, Name
        $usrCols = @(@{H='Name';K='Name'}, @{H='Job title';K='JobTitle'}, @{H='Department';K='Department'}, @{H='Office';K='Office'}, @{H='Manager';K='Manager'}, @{H='Type';K='Type'}, @{H='Enabled';K='Enabled'})
        $tableSections += @{ Id = 'tbl-org'; Title = 'Organization (users)'; Inner = (New-HtmlTableInner 'tbl-org' 'Organization - users by department, office, manager' $usrCols $usrRows) }
    }

    # App registrations
    $areg = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'AppRegEnterpriseApps')
    if ($areg.Count) {
        $aregRows = foreach ($a in $areg) {
            [pscustomobject]@{ App = "$($a.AppName)"; Owner = (Trunc $a.OwnerInfo 40); Secrets = [int]$a.SecretCount; Certs = [int]$a.CertificateCount
                NextExpiry = (Dt $a.NearestCredentialExpiry); Expired = (YN $a.CredentialExpired); Audience = "$($a.SignInAudience)"; LastSignIn = (Dt $a.LastSignInDateTime) }
        }
        $aregRows = $aregRows | Sort-Object App
        $aregCols = @(@{H='App';K='App'}, @{H='Owner';K='Owner'}, @{H='Secrets';K='Secrets';Num=$true}, @{H='Certs';K='Certs';Num=$true}, @{H='Next expiry';K='NextExpiry'}, @{H='Expired';K='Expired'}, @{H='Audience';K='Audience'}, @{H='Last sign-in';K='LastSignIn'})
        $tableSections += @{ Id = 'tbl-areg'; Title = 'App registrations (table)'; Inner = (New-HtmlTableInner 'tbl-areg' 'App registrations (table)' $aregCols $aregRows) }
    }

    # Enterprise apps
    $eapp = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'EnterpriseApps')
    if ($eapp.Count) {
        $eappRows = foreach ($e in $eapp) {
            [pscustomobject]@{ App = "$($e.AppName)"; Type = "$($e.ServicePrincipalType)"; Origin = "$($e.AppOrigin)"; Enabled = (YN $e.AccountEnabled)
                SSO = (YN $e.SSOConfigured); SSOType = "$($e.SSOType)"; AssignReq = (YN $e.AppRoleAssignmentRequired); LastSignIn = (Dt $e.LastSignInDateTime) }
        }
        $eappRows = $eappRows | Sort-Object App
        $eappCols = @(@{H='App';K='App'}, @{H='Type';K='Type'}, @{H='Origin';K='Origin'}, @{H='Enabled';K='Enabled'}, @{H='SSO';K='SSO'}, @{H='SSO type';K='SSOType'}, @{H='Assignment req.';K='AssignReq'}, @{H='Last sign-in';K='LastSignIn'})
        $tableSections += @{ Id = 'tbl-eapp'; Title = 'Enterprise apps (table)'; Inner = (New-HtmlTableInner 'tbl-eapp' 'Enterprise apps (table)' $eappCols $eappRows) }
    }

    # Conditional Access
    $caRoot = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'ConditionalAccess') | Select-Object -First 1
    $caPol = @($caRoot.ConditionalAccessPolicies)
    if ($caPol.Count) {
        $caRows = foreach ($p in $caPol) {
            [pscustomobject]@{ Policy = "$($p.DisplayName)"; State = "$($p.State)"; Grant = (Trunc ("$($p.GrantOperator) $($p.BuiltInControls)").Trim() 40)
                Applications = (Trunc $p.IncludeApplicationsResolved 40); AuthStrength = "$($p.AuthenticationStrengthResolved)"; Users = (Trunc $p.IncludeUsersResolved 40) }
        }
        $caRows = $caRows | Sort-Object Policy
        $caCols = @(@{H='Policy';K='Policy'}, @{H='State';K='State'}, @{H='Grant';K='Grant'}, @{H='Applications';K='Applications'}, @{H='Auth strength';K='AuthStrength'}, @{H='Users';K='Users'})
        $tableSections += @{ Id = 'tbl-ca'; Title = 'Conditional Access (table)'; Inner = (New-HtmlTableInner 'tbl-ca' 'Conditional Access (table)' $caCols $caRows) }
    }

    # RBAC / PIM
    $rbac = @(Get-AutodocExportData -ExportFolder $SourceFolder -BaseName 'RBAC')
    if ($rbac.Count) {
        $rbacRows = foreach ($r in $rbac) {
            [pscustomobject]@{ Role = "$($r.RoleDefinitionName)"; Principal = "$($r.PrincipalDisplayName)"; PrincipalType = "$($r.PrincipalType)"
                Assignment = "$($r.AssignmentType)"; MemberType = "$($r.MemberType)" }
        }
        $rbacRows = $rbacRows | Sort-Object Role, Principal
        $rbacCols = @(@{H='Role';K='Role'}, @{H='Principal';K='Principal'}, @{H='Principal type';K='PrincipalType'}, @{H='Assignment';K='Assignment'}, @{H='Member type';K='MemberType'})
        $tableSections += @{ Id = 'tbl-rbac'; Title = 'RBAC (table)'; Inner = (New-HtmlTableInner 'tbl-rbac' 'RBAC (table)' $rbacCols $rbacRows) }
    }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!doctype html><html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine("<title>$(HtmlEnc $tName) - Entra ID documentation</title>")
    [void]$sb.AppendLine(@'
<style>
  :root { --blue:#0F6CBD; --blue-d:#0B4C87; --ink:#1a1a1a; --muted:#5b6470; --line:#e3e6ea; --bg:#f4f6f8; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI',system-ui,sans-serif; color:var(--ink); background:var(--bg); }
  header.top { background:linear-gradient(135deg,var(--blue),var(--blue-d)); color:#fff; padding:28px 32px; }
  header.top h1 { margin:0 0 4px; font-size:24px; }
  header.top .sub { opacity:.92; font-size:14px; line-height:1.5; }
  .summary { display:flex; flex-wrap:wrap; gap:12px; margin-top:16px; }
  .summary .card { background:rgba(255,255,255,.15); border:1px solid rgba(255,255,255,.25); border-radius:8px; padding:10px 16px; min-width:120px; }
  .summary .card .v { font-size:20px; font-weight:700; }
  .summary .card .l { font-size:12px; opacity:.9; }
  .wrap { max-width:1680px; margin:0 auto; padding:24px 32px 24px; display:flex; gap:28px; align-items:flex-start; }
  nav.toc { flex:0 0 250px; width:250px; background:#fff; border:1px solid var(--line); border-radius:10px; padding:16px 18px; position:sticky; top:16px; max-height:calc(100vh - 32px); overflow:auto; }
  nav.toc h2 { margin:0 0 10px; font-size:13px; color:var(--muted); text-transform:uppercase; letter-spacing:.04em; }
  nav.toc ul { margin:0; padding:0; list-style:none; display:flex; flex-direction:column; gap:2px; }
  nav.toc a { display:block; color:var(--blue); text-decoration:none; font-size:14px; padding:6px 8px; border-radius:6px; }
  nav.toc a:hover { background:var(--bg); }
  .content { flex:1; min-width:0; }
  .tabs { display:flex; gap:4px; margin-bottom:20px; border-bottom:2px solid var(--line); }
  .tabs .tab { padding:10px 22px; border:none; background:none; font-size:15px; font-weight:600; color:var(--muted); cursor:pointer; border-bottom:3px solid transparent; margin-bottom:-2px; }
  .tabs .tab:hover { color:var(--blue); }
  .tabs .tab.active { color:var(--blue); border-bottom-color:var(--blue); }
  section.map { background:#fff; border:1px solid var(--line); border-radius:10px; margin-bottom:22px; overflow:hidden; scroll-margin-top:16px; }
  section.map > h2 { margin:0; padding:14px 20px; font-size:17px; border-bottom:1px solid var(--line); }
  section.map .canvas { max-height:85vh; overflow:auto; padding:16px 20px; background:#fff; }
  section.map .canvas svg { display:block; height:auto; max-width:100%; }
  footer { text-align:center; color:var(--muted); font-size:12px; padding:24px; }
  section.tbl { background:#fff; border:1px solid var(--line); border-radius:10px; margin-bottom:22px; overflow:hidden; scroll-margin-top:16px; }
  section.tbl > h2 { margin:0; padding:14px 20px; font-size:17px; border-bottom:1px solid var(--line); }
  .tbltools { display:flex; align-items:center; gap:12px; padding:12px 20px; border-bottom:1px solid var(--line); background:#fafbfc; }
  .tbltools input { flex:1; max-width:340px; padding:8px 12px; border:1px solid var(--line); border-radius:6px; font-size:14px; }
  .tbltools .count { color:var(--muted); font-size:13px; }
  .colpick { position:relative; }
  .colpick > button { padding:8px 12px; border:1px solid var(--line); border-radius:6px; background:#fff; cursor:pointer; font-size:13px; }
  .colpick > button:hover { background:var(--bg); }
  .colmenu { display:none; position:absolute; right:0; top:38px; background:#fff; border:1px solid var(--line); border-radius:8px; box-shadow:0 6px 20px rgba(0,0,0,.14); padding:8px; z-index:20; min-width:190px; max-height:320px; overflow:auto; }
  .colmenu label { display:block; padding:5px 8px; font-size:13px; white-space:nowrap; cursor:pointer; }
  .colmenu label:hover { background:var(--bg); border-radius:4px; }
  .colmenu input { margin-right:8px; }
  .tblwrap { max-height:70vh; overflow:auto; }
  table.data { border-collapse:collapse; width:100%; font-size:13px; }
  table.data thead th { position:sticky; top:0; background:#eef3f8; text-align:left; padding:9px 14px; border-bottom:2px solid var(--line); cursor:pointer; white-space:nowrap; user-select:none; }
  table.data thead th:hover { background:#e2ebf5; }
  table.data thead th::after { content:''; }
  table.data thead th[data-dir=asc]::after  { content:' \25B2'; font-size:10px; }
  table.data thead th[data-dir=desc]::after { content:' \25BC'; font-size:10px; }
  table.data td { padding:7px 14px; border-bottom:1px solid #eef1f4; white-space:normal; word-break:break-word; overflow-wrap:anywhere; vertical-align:top; max-width:340px; }
  table.data td[data-v] { white-space:nowrap; max-width:none; }
  table.data tbody tr:nth-child(even) { background:#fafbfc; }
  table.data tbody tr:hover { background:#eef6ff; }
</style></head><body>
'@)

    # header
    [void]$sb.AppendLine('<header class="top">')
    [void]$sb.AppendLine("  <h1>$(HtmlEnc $tName)</h1>")
    $subBits = @($tDom, "Environment: $tEnv", "Generated $today") | Where-Object { $_ -and $_ -ne 'Environment: ' }
    [void]$sb.AppendLine("  <div class='sub'>$(HtmlEnc ($subBits -join '  |  '))</div>")
    if ($cards.Count) {
        [void]$sb.AppendLine('  <div class="summary">')
        foreach ($c in $cards) { [void]$sb.AppendLine("    <div class='card'><div class='v'>$('{0:N0}' -f [int]$c.V)</div><div class='l'>$(HtmlEnc $c.L)</div></div>") }
        [void]$sb.AppendLine('  </div>')
    }
    [void]$sb.AppendLine('</header>')

    [void]$sb.AppendLine('<div class="wrap">')

    # section metadata for the maps
    $idx = 0
    $sections = foreach ($m in $maps) {
        [pscustomobject]@{ Id = 'map' + (++$idx); Title = (Get-MapTitle $m.BaseName); File = $m.FullName }
    }

    # sidebar - grouped by tab
    [void]$sb.AppendLine('<nav class="toc">')
    [void]$sb.AppendLine('  <h2>Maps</h2><ul>')
    foreach ($s in $sections) { [void]$sb.AppendLine("    <li><a href='#$($s.Id)' onclick=""showTab('maps')"">$(HtmlEnc $s.Title)</a></li>") }
    [void]$sb.AppendLine('  </ul>')
    if ($tableSections.Count) {
        [void]$sb.AppendLine('  <h2 style="margin-top:16px">Tables</h2><ul>')
        foreach ($ts in $tableSections) { [void]$sb.AppendLine("    <li><a href='#$($ts.Id)' onclick=""showTab('tables')"">$(HtmlEnc $ts.Title)</a></li>") }
        [void]$sb.AppendLine('  </ul>')
    }
    [void]$sb.AppendLine('</nav>')

    # content with tabs
    [void]$sb.AppendLine('<div class="content">')
    [void]$sb.AppendLine('  <div class="tabs">')
    [void]$sb.AppendLine("    <button class='tab active' data-tab='maps' onclick=""showTab('maps')"">Maps</button>")
    if ($tableSections.Count) { [void]$sb.AppendLine("    <button class='tab' data-tab='tables' onclick=""showTab('tables')"">Tables</button>") }
    [void]$sb.AppendLine('  </div>')

    # Maps panel
    [void]$sb.AppendLine('<div id="panel-maps">')
    foreach ($s in $sections) {
        $svg = Get-Content -Raw -LiteralPath $s.File
        $svg = $svg -replace '(?s)^\s*<\?xml.*?\?>', ''   # drop XML declaration
        [void]$sb.AppendLine("<section class='map' id='$($s.Id)'>")
        [void]$sb.AppendLine("  <h2>$(HtmlEnc $s.Title)</h2>")
        [void]$sb.AppendLine("  <div class='canvas'>$svg</div>")
        [void]$sb.AppendLine('</section>')
    }
    [void]$sb.AppendLine('</div>')

    # Tables panel
    [void]$sb.AppendLine('<div id="panel-tables" style="display:none">')
    foreach ($ts in $tableSections) {
        [void]$sb.AppendLine("<section class='tbl' id='$($ts.Id)'><h2>$(HtmlEnc $ts.Title)</h2>$($ts.Inner)</section>")
    }
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('</div>')   # /content
    [void]$sb.AppendLine('</div>')   # /wrap
    [void]$sb.AppendLine("<footer>$(HtmlEnc $tName) - $today</footer>")
    [void]$sb.AppendLine(@'
<script>
function showTab(name){
  var pm=document.getElementById('panel-maps'), pt=document.getElementById('panel-tables');
  if(pm) pm.style.display=(name==='maps')?'':'none';
  if(pt) pt.style.display=(name==='tables')?'':'none';
  document.querySelectorAll('.tabs .tab').forEach(function(b){ b.classList.toggle('active', b.getAttribute('data-tab')===name); });
}
function tblFilter(id){
  var q=document.getElementById(id+'-q').value.toLowerCase();
  var rows=document.querySelectorAll('#'+id+'-t tbody tr'), n=0;
  rows.forEach(function(r){var s=r.textContent.toLowerCase().indexOf(q)>-1; r.style.display=s?'':'none'; if(s)n++;});
  document.getElementById(id+'-count').textContent=n;
}
function colMenu(id){
  var m=document.getElementById(id+'-cols');
  var open=m.style.display==='block';
  document.querySelectorAll('.colmenu').forEach(function(x){x.style.display='none';});
  m.style.display=open?'none':'block';
}
function colToggle(tid,col,show){
  var t=document.getElementById(tid), d=show?'':'none';
  var th=t.querySelectorAll('thead th')[col]; if(th)th.style.display=d;
  t.querySelectorAll('tbody tr').forEach(function(r){ if(r.children[col]) r.children[col].style.display=d; });
}
document.addEventListener('click',function(e){
  if(!e.target.closest('.colpick')){ document.querySelectorAll('.colmenu').forEach(function(x){x.style.display='none';}); }
});
function tblSort(tid,col,th){
  var tb=document.querySelector('#'+tid+' tbody');
  var rows=[].slice.call(tb.querySelectorAll('tr'));
  var asc=th.getAttribute('data-dir')!=='asc';
  th.parentNode.querySelectorAll('th').forEach(function(x){x.removeAttribute('data-dir');});
  th.setAttribute('data-dir',asc?'asc':'desc');
  rows.sort(function(a,b){
    var x=a.children[col].getAttribute('data-v'); if(x===null)x=a.children[col].textContent;
    var y=b.children[col].getAttribute('data-v'); if(y===null)y=b.children[col].textContent;
    var nx=parseFloat(x),ny=parseFloat(y);
    var c=(!isNaN(nx)&&!isNaN(ny))?(nx-ny):(''+x).localeCompare(''+y);
    return asc?c:-c;
  });
  rows.forEach(function(r){tb.appendChild(r);});
}
</script>
'@)
    [void]$sb.AppendLine('</body></html>')

    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $htmlDir = Join-Path $SourceFolder 'Report\html'
    if (-not (Test-Path $htmlDir)) { New-Item -Path $htmlDir -ItemType Directory -Force | Out-Null }
    $out = Join-Path $htmlDir "Entra ID - Full report $stamp.html"
    Set-Content -LiteralPath $out -Value $sb.ToString() -Encoding UTF8

    Write-Host "  HTML report: $(@($sections).Count) maps -> $out" -ForegroundColor DarkGreen
    return $out
}
