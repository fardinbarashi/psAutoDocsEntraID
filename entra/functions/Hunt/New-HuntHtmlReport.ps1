# Autodoc v1.0
<#
    Self-contained HTML report (no external files or CDN) for
    Invoke-WorkloadIdentityHunt:
      $sections = @( New-HuntReportSection -Id 'findings' -Title 'Findings' -Rows $rows ... )
      New-HuntHtmlReport -Path '...\report.html' -Meta $meta -Sections $sections

    The report has:
      - summary cards (click to filter findings / accounts)
      - one tab per section
      - free text search over all columns
      - per-column filter (dropdown for columns with few distinct values, text otherwise)
      - sorting by clicking a column header
      - paging
      - export of the filtered rows to CSV (semicolon separated, UTF-8 with BOM for Excel)
      - light and dark mode following the OS setting
#>

function ConvertTo-HuntJsonString {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '""' }
    $s = $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t')
    $s = $s.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
    $s = [regex]::Replace($s, '[\x00-\x1f\u2028\u2029]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
    return '"' + $s + '"'
}

function ConvertTo-HuntCellValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    if ($Value -is [bool]) {
        if ($Value) { return 'Yes' }
        return 'No'
    }
    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($x in $Value) { [void]$parts.Add([string]$x) }
        return ($parts.ToArray() -join '; ')
    }
    return [string]$Value
}

function New-HuntReportSection {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Title,
        [string]$Description = '',
        [AllowNull()]$Rows,
        [string[]]$Columns,
        [int]$MaxRows = 0,
        [string]$DefaultSort = ''
    )
    $all   = @($Rows | Where-Object { $null -ne $_ })
    $total = $all.Count
    $truncated = $false
    if ($MaxRows -gt 0 -and $total -gt $MaxRows) {
        $all = @($all | Select-Object -First $MaxRows)
        $truncated = $true
    }
    if (-not $Columns -or $Columns.Count -eq 0) {
        $Columns = @()
        if ($all.Count -gt 0) { $Columns = @($all[0].PSObject.Properties | ForEach-Object { $_.Name }) }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('{"id":').Append((ConvertTo-HuntJsonString $Id))
    [void]$sb.Append(',"title":').Append((ConvertTo-HuntJsonString $Title))
    [void]$sb.Append(',"description":').Append((ConvertTo-HuntJsonString $Description))
    [void]$sb.Append(',"defaultSort":').Append((ConvertTo-HuntJsonString $DefaultSort))
    [void]$sb.Append(',"total":').Append($total)
    $truncText = 'false'
    if ($truncated) { $truncText = 'true' }
    [void]$sb.Append(',"truncated":').Append($truncText)

    [void]$sb.Append(',"columns":[')
    for ($c = 0; $c -lt $Columns.Count; $c++) {
        if ($c -gt 0) { [void]$sb.Append(',') }
        [void]$sb.Append((ConvertTo-HuntJsonString $Columns[$c]))
    }
    [void]$sb.Append('],"rows":[')
    for ($r = 0; $r -lt $all.Count; $r++) {
        if ($r -gt 0) { [void]$sb.Append(',') }
        [void]$sb.Append('[')
        $row = $all[$r]
        for ($c = 0; $c -lt $Columns.Count; $c++) {
            if ($c -gt 0) { [void]$sb.Append(',') }
            $val = ConvertTo-HuntCellValue ($row.($Columns[$c]))
            [void]$sb.Append((ConvertTo-HuntJsonString $val))
        }
        [void]$sb.Append(']')
    }
    [void]$sb.Append(']}')

    return New-Object PSObject -Property ([ordered]@{
        Id        = $Id
        Total     = $total
        Truncated = $truncated
        Json      = $sb.ToString()
    })
}

function ConvertTo-HuntJsonObject {
    param([System.Collections.IDictionary]$Table)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($k in $Table.Keys) {
        $v = $Table[$k]
        if ($v -is [System.Collections.IDictionary]) {
            [void]$parts.Add((ConvertTo-HuntJsonString ([string]$k)) + ':' + (ConvertTo-HuntJsonObject $v))
        } else {
            [void]$parts.Add((ConvertTo-HuntJsonString ([string]$k)) + ':' + (ConvertTo-HuntJsonString (ConvertTo-HuntCellValue $v)))
        }
    }
    return '{' + ($parts -join ',') + '}'
}

function New-HuntHtmlReport {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Meta,
        [Parameter(Mandatory)][object[]]$Sections
    )
    $sectionJson = (@($Sections | ForEach-Object { $_.Json })) -join ','
    $json = '{"meta":' + (ConvertTo-HuntJsonObject $Meta) + ',"sections":[' + $sectionJson + ']}'

    $template = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Workload identity hunt</title>
<style>
:root{--bg:#f5f6f8;--panel:#ffffff;--text:#1b1f24;--muted:#5d6670;--line:#dde1e6;--accent:#2457c5;--crit:#c62828;--high:#8e24aa;--med:#b26a00;--info:#607080;--ok:#2e7d32;--badge-text:#ffffff;--hover:#eef3fc}
@media (prefers-color-scheme:dark){:root{--bg:#121519;--panel:#1b1f25;--text:#e6e9ed;--muted:#9aa4af;--line:#2d333b;--accent:#6c9cff;--crit:#ff6b6b;--high:#d68cff;--med:#f0b44c;--info:#9aa8b5;--ok:#6fcf76;--badge-text:#111418;--hover:#232a35}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 "Segoe UI",system-ui,-apple-system,sans-serif}
header{padding:20px 24px 14px;background:var(--panel);border-bottom:1px solid var(--line)}
header h1{margin:0 0 6px;font-size:20px;font-weight:650}
.meta{display:flex;flex-wrap:wrap;gap:4px 18px;color:var(--muted);font-size:12.5px}
.meta b{color:var(--text);font-weight:600}
#cards{display:flex;flex-wrap:wrap;gap:10px;padding:16px 24px}
.card{display:flex;flex-direction:column;align-items:flex-start;min-width:128px;padding:10px 14px;border:1px solid var(--line);border-left:4px solid var(--info);border-radius:8px;background:var(--panel);color:var(--text);font:inherit;text-align:left}
button.card{cursor:pointer}
button.card:hover{border-color:var(--accent);border-left-color:inherit}
.card .num{font-size:22px;font-weight:700;line-height:1.2}
.card .lbl{font-size:12px;color:var(--muted)}
.c-crit{border-left-color:var(--crit)!important}.c-high{border-left-color:var(--high)!important}.c-med{border-left-color:var(--med)!important}.c-info{border-left-color:var(--info)!important}.c-ok{border-left-color:var(--ok)!important}.c-stat{border-left-color:var(--accent)!important}
#tabs{display:flex;flex-wrap:wrap;gap:2px;padding:0 24px;border-bottom:1px solid var(--line)}
#tabs button{border:0;border-bottom:2px solid transparent;background:none;color:var(--muted);padding:10px 12px;font:inherit;cursor:pointer;border-radius:0}
#tabs button:hover{color:var(--text)}
#tabs button.active{color:var(--text);border-bottom-color:var(--accent);font-weight:600}
.cnt{display:inline-block;min-width:22px;padding:0 6px;margin-left:4px;border-radius:10px;background:var(--line);color:var(--text);font-size:11.5px;text-align:center;font-weight:500}
main{padding:16px 24px 32px}
.sec-head h2{margin:0 0 4px;font-size:17px}
.sec-head p{margin:0 0 10px;color:var(--muted)}
.sec-head p.warn{color:var(--med)}
.toolbar{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-bottom:10px}
.toolbar input[type=search]{flex:1 1 260px;max-width:420px}
input,select,button{font:inherit;font-size:13px;color:var(--text);background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:5px 8px}
input:focus,select:focus{outline:2px solid var(--accent);outline-offset:-1px}
.toolbar button,.pager button{cursor:pointer}
.toolbar button:hover,.pager button:hover:not(:disabled){border-color:var(--accent)}
button:disabled{opacity:.45;cursor:default}
.grow{flex:1}
#count{color:var(--muted)}
.tablewrap{overflow:auto;max-height:72vh;border:1px solid var(--line);border-radius:8px;background:var(--panel)}
table{border-collapse:separate;border-spacing:0;width:100%}
th,td{padding:6px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;white-space:nowrap}
thead th{position:sticky;background:var(--panel);z-index:2}
thead tr:first-child th{top:0}
thead tr.filters th{top:33px;padding-top:4px;padding-bottom:6px}
th.sortable{cursor:pointer;user-select:none;font-weight:600}
th.sortable:hover{color:var(--accent)}
.arrow{color:var(--accent)}
tr.filters input,tr.filters select{width:100%;min-width:90px;padding:3px 6px;font-size:12.5px;font-weight:400}
tbody tr:hover td{background:var(--hover)}
td.long{white-space:normal;min-width:320px;max-width:560px;word-break:break-word}
.badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11.5px;font-weight:600;color:var(--badge-text)}
.b-crit{background:var(--crit)}.b-high{background:var(--high)}.b-med{background:var(--med)}.b-info{background:var(--info)}.b-ok{background:var(--ok)}
.pager{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-top:10px;color:var(--muted)}
.empty{color:var(--muted);text-align:center;padding:22px}
footer{padding:0 24px 24px;color:var(--muted);font-size:12px}
@media (max-width:700px){header,#cards,#tabs,main,footer{padding-left:12px;padding-right:12px}}
</style>
</head>
<body>
<header>
  <h1 id="title">Workload identity hunt</h1>
  <div class="meta" id="meta"></div>
</header>
<section id="cards"></section>
<nav id="tabs" role="tablist"></nav>
<main id="view"></main>
<footer>Read-only analysis. The full data is in the CSV files next to this report. Times are shown in the local time of the computer that ran the hunt.</footer>
<script id="report-data" type="application/json">__REPORT_DATA__</script>
<script>
(function () {
  'use strict';
  var D = JSON.parse(document.getElementById('report-data').textContent);
  var SEV = { CRITICAL: 0, HIGH: 1, MEDIUM: 2, INFO: 3 };
  var VERD = { COMPROMISED: 0, PASSWORD_KNOWN: 1, ACTIVE_ATTACK: 2, SPRAY: 3, NO_CLEAR_SIGNAL: 4 };
  var BADGE = {
    CRITICAL: 'crit', HIGH: 'high', MEDIUM: 'med', INFO: 'info',
    COMPROMISED: 'crit', PASSWORD_KNOWN: 'crit', ACTIVE_ATTACK: 'high', SPRAY: 'med', NO_CLEAR_SIGNAL: 'info',
    SUCCESS: 'ok', PWD_OK: 'crit', CA_BLOCK: 'crit', LOCKED: 'high', OTHER: 'info'
  };
  var BADGE_COLS = { Severity: 1, Verdict: 1, Category: 1 };
  var IOC_COLS = { IsIocIP: 1 };
  var LONG_COLS = { Detail: 1, Action: 1, Meaning: 1, FailureReason: 1, Details: 1, Target: 1, IocApps: 1, IocErrorCodes: 1 };
  var EMPTY = '__EMPTY__';
  var state = {};
  var current = null;
  var timer = null;

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function byId(id) { return document.getElementById(id); }
  function sectionById(id) {
    for (var i = 0; i < D.sections.length; i++) { if (D.sections[i].id === id) return D.sections[i]; }
    return null;
  }
  function rank(col, v) {
    if (col === 'Severity' && v in SEV) return SEV[v];
    if (col === 'Verdict' && v in VERD) return VERD[v];
    return null;
  }
  function cmp(a, b) {
    var na = Number(a), nb = Number(b);
    if (a !== '' && b !== '' && !isNaN(na) && !isNaN(nb)) return na - nb;
    return String(a).localeCompare(String(b), 'en');
  }
  function cmpCol(col, a, b) {
    var ra = rank(col, a), rb = rank(col, b);
    if (ra !== null || rb !== null) return (ra === null ? 99 : ra) - (rb === null ? 99 : rb);
    return cmp(a, b);
  }
  function isSelect(s, i) { return s.distinct[i] !== null && s.distinct[i].length > 1; }

  function prepare() {
    D.sections.forEach(function (s) {
      s.lc = null;
      s.distinct = s.columns.map(function (c, i) {
        var seen = {}, list = [];
        for (var k = 0; k < s.rows.length; k++) {
          var v = s.rows[k][i];
          if (!(v in seen)) { seen[v] = 1; list.push(v); if (list.length > 30) return null; }
        }
        return list.sort(function (a, b) { return cmpCol(c, a, b); });
      });
      var sortIdx = s.defaultSort ? s.columns.indexOf(s.defaultSort) : -1;
      var dir = (s.defaultSort === 'Time') ? -1 : 1;
      state[s.id] = { q: '', f: {}, sort: sortIdx, dir: dir, page: 0, size: 100, idx: [] };
    });
  }

  function renderHeader() {
    byId('title').textContent = (D.meta.title || 'Workload identity hunt');
    document.title = (D.meta.title || 'Workload identity hunt') + (D.meta.stamp ? ' (' + D.meta.stamp + ')' : '');
    var info = D.meta.info || {}, html = [];
    Object.keys(info).forEach(function (k) { html.push('<span><b>' + esc(k) + ':</b> ' + esc(info[k]) + '</span>'); });
    byId('meta').innerHTML = html.join('');
  }

  function renderCards() {
    var out = [];
    var f = sectionById('findings');
    if (f) {
      var si = f.columns.indexOf('Severity');
      ['CRITICAL', 'HIGH', 'MEDIUM', 'INFO'].forEach(function (sev) {
        var n = 0;
        f.rows.forEach(function (r) { if (r[si] === sev) n++; });
        out.push('<button class="card c-' + BADGE[sev] + '" data-sec="findings" data-col="Severity" data-val="' + sev + '" title="Show ' + sev + ' findings"><span class="num">' + n + '</span><span class="lbl">' + sev + '</span></button>');
      });
    }
    var t = sectionById('triage');
    if (t && t.rows.length) {
      var vi = t.columns.indexOf('Verdict'), counts = {};
      t.rows.forEach(function (r) { counts[r[vi]] = (counts[r[vi]] || 0) + 1; });
      Object.keys(counts).sort(function (a, b) { return cmpCol('Verdict', a, b); }).forEach(function (v) {
        out.push('<button class="card c-' + (BADGE[v] || 'info') + '" data-sec="triage" data-col="Verdict" data-val="' + esc(v) + '" title="Show accounts with ' + esc(v) + '"><span class="num">' + counts[v] + '</span><span class="lbl">Accounts: ' + esc(v) + '</span></button>');
      });
    }
    var stats = D.meta.stats || {};
    Object.keys(stats).forEach(function (k) {
      out.push('<div class="card c-stat"><span class="num">' + esc(stats[k]) + '</span><span class="lbl">' + esc(k) + '</span></div>');
    });
    var el = byId('cards');
    el.innerHTML = out.join('');
    Array.prototype.forEach.call(el.querySelectorAll('button.card'), function (b) {
      b.addEventListener('click', function () {
        var sec = sectionById(b.getAttribute('data-sec')), st = state[sec.id];
        st.q = ''; st.f = {}; st.page = 0;
        st.f[sec.columns.indexOf(b.getAttribute('data-col'))] = b.getAttribute('data-val');
        show(sec.id);
      });
    });
  }

  function renderTabs() {
    var nav = byId('tabs');
    nav.innerHTML = D.sections.map(function (s) {
      return '<button role="tab" data-id="' + esc(s.id) + '">' + esc(s.title) + '<span class="cnt">' + s.total + '</span></button>';
    }).join('');
    Array.prototype.forEach.call(nav.querySelectorAll('button'), function (b) {
      b.addEventListener('click', function () { show(b.getAttribute('data-id')); });
    });
  }

  function filterCell(s, i) {
    var cur = state[s.id].f[i];
    if (cur === undefined || cur === null) cur = '';
    if (isSelect(s, i)) {
      var opts = '<option value="">All</option>' + s.distinct[i].map(function (x) {
        var val = (x === '') ? EMPTY : x;
        return '<option value="' + esc(val) + '"' + (val === cur ? ' selected' : '') + '>' + esc(x === '' ? '(empty)' : x) + '</option>';
      }).join('');
      return '<th><select data-i="' + i + '" aria-label="Filter ' + esc(s.columns[i]) + '">' + opts + '</select></th>';
    }
    return '<th><input data-i="' + i + '" placeholder="Filter" value="' + esc(cur) + '" aria-label="Filter ' + esc(s.columns[i]) + '"></th>';
  }

  function show(id) {
    current = id;
    Array.prototype.forEach.call(document.querySelectorAll('#tabs button'), function (b) {
      var on = b.getAttribute('data-id') === id;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    var s = sectionById(id), st = state[id], v = byId('view');
    var h = '<div class="sec-head"><h2>' + esc(s.title) + '</h2><p>' + esc(s.description || '') + '</p>';
    if (s.truncated) h += '<p class="warn">Showing ' + s.rows.length + ' of ' + s.total + ' rows (limited by HtmlMaxRows). All rows are in the CSV file.</p>';
    h += '</div>';
    if (!s.columns.length) {
      v.innerHTML = h + '<div class="tablewrap"><p class="empty">No rows.</p></div>';
      return;
    }
    h += '<div class="toolbar">' +
      '<input type="search" id="q" placeholder="Search all columns..." value="' + esc(st.q) + '">' +
      '<button id="clear" type="button">Clear filters</button>' +
      '<span id="count"></span><span class="grow"></span>' +
      '<label>Rows per page <select id="size">' + [50, 100, 250, 1000].map(function (n) {
        return '<option' + (n === st.size ? ' selected' : '') + '>' + n + '</option>';
      }).join('') + '</select></label>' +
      '<button id="csv" type="button">Export selection (CSV)</button></div>';
    h += '<div class="tablewrap"><table><thead><tr>' +
      s.columns.map(function (c, i) { return '<th class="sortable" data-i="' + i + '">' + esc(c) + '<span class="arrow"></span></th>'; }).join('') +
      '</tr><tr class="filters">' + s.columns.map(function (c, i) { return filterCell(s, i); }).join('') +
      '</tr></thead><tbody></tbody></table></div><div class="pager" id="pager"></div>';
    v.innerHTML = h;

    byId('q').addEventListener('input', function (e) {
      clearTimeout(timer);
      timer = setTimeout(function () { st.q = e.target.value; st.page = 0; update(); }, 150);
    });
    byId('clear').addEventListener('click', function () { st.q = ''; st.f = {}; st.page = 0; show(current); });
    byId('size').addEventListener('change', function (e) { st.size = Number(e.target.value); st.page = 0; update(); });
    byId('csv').addEventListener('click', exportCsv);
    Array.prototype.forEach.call(v.querySelectorAll('th.sortable'), function (th) {
      th.addEventListener('click', function () {
        var i = Number(th.getAttribute('data-i'));
        if (st.sort === i) { st.dir = -st.dir; } else { st.sort = i; st.dir = 1; }
        update();
      });
    });
    Array.prototype.forEach.call(v.querySelectorAll('tr.filters select'), function (el) {
      el.addEventListener('change', function () { st.f[Number(el.getAttribute('data-i'))] = el.value; st.page = 0; update(); });
    });
    Array.prototype.forEach.call(v.querySelectorAll('tr.filters input'), function (el) {
      el.addEventListener('input', function () {
        clearTimeout(timer);
        timer = setTimeout(function () { st.f[Number(el.getAttribute('data-i'))] = el.value; st.page = 0; update(); }, 150);
      });
    });
    byId('pager').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-p]');
      if (!b || b.disabled) return;
      st.page = Number(b.getAttribute('data-p'));
      update();
    });
    update();
  }

  function filtered(s) {
    var st = state[s.id];
    if (!s.lc) s.lc = s.rows.map(function (r) { return r.join('\u0001').toLowerCase(); });
    var q = st.q.trim().toLowerCase();
    var fs = [];
    Object.keys(st.f).forEach(function (k) {
      var val = st.f[k];
      if (val === '' || val === null || val === undefined) return;
      var i = Number(k);
      fs.push({ i: i, v: val, sel: isSelect(s, i), lv: String(val).toLowerCase() });
    });
    var idx = [];
    for (var k = 0; k < s.rows.length; k++) {
      if (q && s.lc[k].indexOf(q) === -1) continue;
      var r = s.rows[k], ok = true;
      for (var j = 0; j < fs.length; j++) {
        var f = fs[j];
        if (f.sel) {
          var want = (f.v === EMPTY) ? '' : f.v;
          if (r[f.i] !== want) { ok = false; break; }
        } else if (r[f.i].toLowerCase().indexOf(f.lv) === -1) { ok = false; break; }
      }
      if (ok) idx.push(k);
    }
    if (st.sort >= 0) {
      var col = s.columns[st.sort];
      idx.sort(function (a, b) { return st.dir * cmpCol(col, s.rows[a][st.sort], s.rows[b][st.sort]) || (a - b); });
    }
    return idx;
  }

  function cell(col, v) {
    if (BADGE_COLS[col] && BADGE[v]) return '<td><span class="badge b-' + BADGE[v] + '">' + esc(v) + '</span></td>';
    if (IOC_COLS[col] && v === 'Yes') return '<td><span class="badge b-crit">Yes</span></td>';
    if (LONG_COLS[col]) return '<td class="long">' + esc(v) + '</td>';
    return '<td>' + esc(v) + '</td>';
  }

  function update() {
    var s = sectionById(current), st = state[current];
    var idx = filtered(s);
    st.idx = idx;
    var pages = Math.max(1, Math.ceil(idx.length / st.size));
    if (st.page >= pages) st.page = pages - 1;
    if (st.page < 0) st.page = 0;
    var from = st.page * st.size, slice = idx.slice(from, from + st.size);
    var body = slice.map(function (k) {
      return '<tr>' + s.rows[k].map(function (v, i) { return cell(s.columns[i], v); }).join('') + '</tr>';
    }).join('');
    if (!body) body = '<tr><td colspan="' + s.columns.length + '" class="empty">No rows match the filter.</td></tr>';
    document.querySelector('#view tbody').innerHTML = body;
    byId('count').textContent = idx.length + ' of ' + s.rows.length + ' rows';
    Array.prototype.forEach.call(document.querySelectorAll('#view th.sortable'), function (th) {
      var i = Number(th.getAttribute('data-i'));
      th.querySelector('.arrow').textContent = (i === st.sort) ? (st.dir > 0 ? ' \u25B2' : ' \u25BC') : '';
    });
    var p = st.page;
    byId('pager').innerHTML =
      '<button type="button" data-p="0"' + (p === 0 ? ' disabled' : '') + '>&laquo;</button>' +
      '<button type="button" data-p="' + (p - 1) + '"' + (p === 0 ? ' disabled' : '') + '>&lsaquo;</button>' +
      '<span>Page ' + (p + 1) + ' of ' + pages + '</span>' +
      '<button type="button" data-p="' + (p + 1) + '"' + (p >= pages - 1 ? ' disabled' : '') + '>&rsaquo;</button>' +
      '<button type="button" data-p="' + (pages - 1) + '"' + (p >= pages - 1 ? ' disabled' : '') + '>&raquo;</button>';
  }

  function exportCsv() {
    var s = sectionById(current), st = state[current];
    var q = function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; };
    var lines = [s.columns.map(q).join(';')];
    st.idx.forEach(function (k) { lines.push(s.rows[k].map(q).join(';')); });
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = s.id + '_' + (D.meta.stamp || 'export') + '.csv';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
  }

  prepare();
  renderHeader();
  renderCards();
  renderTabs();
  if (D.sections.length) show(D.sections[0].id);
})();
</script>
</body>
</html>
'@

    $html = $template.Replace('__REPORT_DATA__', $json)
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding $false))
    return $Path
}
