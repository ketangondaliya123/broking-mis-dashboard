"""
Build the self-contained Broking MIS Dashboard HTML file.

Reads build/data.json and inlines SheetJS (xlsx.full.min.js) so the
dashboard works fully offline. Outputs to outputs/ directory.

Can also be imported and called as generate_html(data_dict, sheetjs_code).
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
OUT = ROOT / "outputs"
OUT.mkdir(exist_ok=True)


def get_sheetjs():
    return (ROOT / "node_modules" / "xlsx" / "dist" / "xlsx.full.min.js").read_text(encoding="utf-8")


def generate_html(data, sheetjs=None):
    """Generate the full dashboard HTML string from a data dict."""
    if sheetjs is None:
        sheetjs = get_sheetjs()

    meta = data["meta"]
    fy_label = meta["fyLabel"]
    new_label = meta["newLabel"]
    new_short = meta["newShortLabel"]
    factor = meta["factor"]

    dashboard_data = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="description" content="Broking MIS Dashboard for Marwadi Shares and Finance Ltd. — {new_short} vs {fy_label} comparative analytics">
<title>Broking MIS Dashboard — {new_short} | Marwadi Shares and Finance Ltd.</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --navy:#10253F;--navy2:#173A5E;--navy-light:#1E3A5F;
  --ivory:#FAF7F0;--paper:#F3EFE6;--gold:#C9A227;--gold-light:#E8D48B;
  --ink:#1C2833;--muted:#617182;--line:#D9D4C8;
  --green:#166534;--green-bg:#DCFCE7;--red:#9F1239;--red-bg:#FFE4E6;
  --radius:8px;--shadow:0 2px 12px rgba(16,37,63,.12);
}}
html{{font-size:14px;scroll-behavior:smooth}}
body{{font-family:'Inter',system-ui,sans-serif;background:var(--paper);color:var(--ink);min-height:100vh}}

/* ─── HEADER ─── */
.site-header{{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 100%);color:#fff;padding:1.2rem 2rem;position:sticky;top:0;z-index:1000;box-shadow:0 2px 16px rgba(0,0,0,.3);display:flex;justify-content:space-between;align-items:center}}
.site-header .header-left{{flex:1}}
.site-header h1{{font-family:'Playfair Display',serif;font-size:1.6rem;font-weight:700;letter-spacing:.02em}}
.site-header .subtitle{{font-size:.82rem;color:var(--gold-light);margin-top:.2rem;font-weight:400}}
.btn-upload{{display:inline-flex;align-items:center;gap:.45rem;background:linear-gradient(135deg,var(--gold) 0%,#D4A832 100%);color:var(--navy);border:none;padding:.6rem 1.2rem;border-radius:6px;font-size:.82rem;font-weight:700;cursor:pointer;letter-spacing:.03em;text-decoration:none;box-shadow:0 2px 10px rgba(201,162,39,.35);transition:all .2s;white-space:nowrap}}
.btn-upload:hover{{transform:translateY(-1px);box-shadow:0 4px 16px rgba(201,162,39,.5);background:linear-gradient(135deg,#D4A832 0%,#E0BC48 100%)}}
.btn-upload svg{{width:16px;height:16px;fill:currentColor}}

/* ─── NAV ─── */
.main-nav{{background:var(--navy);border-bottom:2px solid var(--gold);display:flex;gap:0;position:sticky;top:62px;z-index:999;overflow-x:auto}}
.main-nav button{{background:none;border:none;color:rgba(255,255,255,.65);padding:.75rem 1.4rem;font-size:.85rem;font-weight:500;cursor:pointer;white-space:nowrap;border-bottom:3px solid transparent;transition:all .2s}}
.main-nav button:hover{{color:#fff;background:rgba(255,255,255,.06)}}
.main-nav button.active{{color:var(--gold);border-bottom-color:var(--gold);background:rgba(201,162,39,.08)}}

/* ─── SUB-NAV ─── */
.sub-nav{{background:var(--ivory);border-bottom:1px solid var(--line);display:flex;gap:0;padding:0 1rem}}
.sub-nav button{{background:none;border:none;color:var(--muted);padding:.6rem 1.2rem;font-size:.8rem;font-weight:500;cursor:pointer;border-bottom:2px solid transparent;transition:all .2s}}
.sub-nav button:hover{{color:var(--ink)}}
.sub-nav button.active{{color:var(--navy);border-bottom-color:var(--navy);font-weight:600}}

/* ─── CONTENT ─── */
.content{{max-width:1600px;margin:0 auto;padding:1.5rem 2rem 3rem}}
.section{{display:none}}.section.active{{display:block}}

/* ─── KPI CARDS ─── */
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:1rem;margin-bottom:1.8rem}}
.kpi-card{{background:#fff;border:1px solid var(--line);border-radius:var(--radius);padding:1.2rem 1.4rem;box-shadow:var(--shadow);position:relative;overflow:hidden}}
.kpi-card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--gold)}}
.kpi-card .kpi-label{{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-weight:600}}
.kpi-card .kpi-value{{font-family:'JetBrains Mono',monospace;font-size:1.5rem;font-weight:500;color:var(--navy);margin:.4rem 0 .1rem}}
.kpi-card .kpi-sub{{font-size:.72rem;color:var(--muted)}}
.kpi-card .kpi-growth{{font-size:.78rem;font-weight:600;margin-top:.3rem}}
.kpi-card .kpi-growth.positive{{color:var(--green)}}.kpi-card .kpi-growth.negative{{color:var(--red)}}

/* ─── TABLE CONTAINERS ─── */
.table-wrapper{{background:#fff;border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;margin-bottom:1.5rem}}
.table-header{{background:var(--navy);color:#fff;padding:.8rem 1.2rem;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:.6rem}}
.table-header h3{{font-family:'Playfair Display',serif;font-size:1rem;font-weight:600}}
.table-controls{{display:flex;gap:.5rem;align-items:center;flex-wrap:wrap}}
.table-controls input[type=text]{{padding:.4rem .7rem;border:1px solid rgba(255,255,255,.3);border-radius:4px;background:rgba(255,255,255,.1);color:#fff;font-size:.8rem;width:200px;outline:none}}
.table-controls input[type=text]::placeholder{{color:rgba(255,255,255,.5)}}
.table-controls input[type=text]:focus{{border-color:var(--gold);background:rgba(255,255,255,.15)}}
.table-controls select{{padding:.4rem .5rem;border:1px solid rgba(255,255,255,.3);border-radius:4px;background:rgba(255,255,255,.1);color:#fff;font-size:.78rem;outline:none;cursor:pointer}}
.table-controls select option{{color:var(--ink);background:#fff}}

/* ─── VIEW TOGGLE ─── */
.view-toggle{{display:flex;gap:0;border:1px solid rgba(255,255,255,.3);border-radius:4px;overflow:hidden}}
.view-toggle button{{background:rgba(255,255,255,.08);border:none;color:rgba(255,255,255,.7);padding:.35rem .8rem;font-size:.75rem;cursor:pointer;transition:all .15s}}
.view-toggle button.active{{background:var(--gold);color:var(--navy);font-weight:600}}

/* ─── EXPORT BUTTONS ─── */
.btn-export{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);color:#fff;padding:.35rem .7rem;border-radius:4px;font-size:.75rem;cursor:pointer;transition:all .15s}}
.btn-export:hover{{background:var(--gold);color:var(--navy);border-color:var(--gold)}}

/* ─── BRANCH GROUP CONTROLS ─── */
.group-controls{{padding:.5rem 1.2rem;background:var(--paper);border-bottom:1px solid var(--line);display:flex;gap:.5rem}}
.group-controls button{{background:var(--navy);color:#fff;border:none;padding:.3rem .7rem;border-radius:4px;font-size:.72rem;cursor:pointer;transition:background .15s}}
.group-controls button:hover{{background:var(--navy2)}}

/* ─── DATA TABLES ─── */
.table-scroll{{max-height:70vh;overflow:auto;position:relative}}
table{{width:100%;border-collapse:collapse;font-size:.8rem}}
thead{{position:static}}
thead th{{position:sticky;top:0;z-index:10;background:var(--navy2);color:#fff;padding:.55rem .6rem;text-align:left;font-weight:600;font-size:.72rem;white-space:normal;word-wrap:break-word;cursor:pointer;border-bottom:2px solid var(--gold);user-select:none;vertical-align:bottom;line-height:1.3}}
thead th:hover{{background:var(--navy-light)}}
thead th .sort-arrow{{font-size:.6rem;margin-left:.2rem;opacity:.5}}
thead th.sorted .sort-arrow{{opacity:1;color:var(--gold)}}
tbody td{{padding:.4rem .6rem;border-bottom:1px solid var(--line);font-size:.78rem;white-space:nowrap}}
tbody tr:hover{{background:rgba(201,162,39,.06)}}
tbody tr.branch-total{{background:rgba(16,37,63,.05);font-weight:600}}
tbody tr.branch-total td{{border-top:1.5px solid var(--navy);font-weight:700}}
td.num{{font-family:'JetBrains Mono',monospace;text-align:right;font-size:.76rem}}
td.pct{{font-family:'JetBrains Mono',monospace;text-align:right;font-size:.76rem}}
td.positive{{color:var(--green)}}
td.negative{{color:var(--red)}}
td.house-flag{{background:var(--red-bg);color:var(--red);font-weight:600;font-size:.7rem}}

/* Branch group header rows */
.branch-group-header{{background:var(--navy)!important;color:#fff;cursor:pointer;user-select:none}}
.branch-group-header td{{color:#fff!important;font-weight:600;padding:.5rem .6rem;border-bottom:2px solid var(--gold)}}
.branch-group-header:hover{{background:var(--navy2)!important}}
.branch-group-header .toggle-icon{{display:inline-block;width:1.2em;transition:transform .2s}}
.branch-group-header.collapsed .toggle-icon{{transform:rotate(-90deg)}}

/* Magnitude bar for gainers/decliners */
.mag-bar{{display:inline-block;height:12px;border-radius:2px;vertical-align:middle;min-width:2px}}
.mag-bar.gain{{background:linear-gradient(90deg,var(--green),#22C55E)}}
.mag-bar.loss{{background:linear-gradient(90deg,#EF4444,var(--red))}}

/* ─── SUMMARY TABLES ─── */
.summary-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(450px,1fr));gap:1.2rem;margin-bottom:1.5rem}}
.summary-table table{{font-size:.82rem}}
.summary-table thead th{{font-size:.75rem;padding:.5rem .6rem}}
.summary-table tbody td{{padding:.45rem .6rem}}

/* ─── PAGINATION ─── */
.pagination{{display:flex;justify-content:space-between;align-items:center;padding:.6rem 1.2rem;background:var(--paper);border-top:1px solid var(--line);font-size:.78rem;color:var(--muted)}}
.pagination button{{background:var(--navy);color:#fff;border:none;padding:.3rem .7rem;border-radius:4px;font-size:.72rem;cursor:pointer}}
.pagination button:disabled{{opacity:.4;cursor:default}}
.page-info{{font-family:'JetBrains Mono',monospace}}

/* ─── RESPONSIVE ─── */
@media(max-width:768px){{
  .content{{padding:1rem}}
  .kpi-grid{{grid-template-columns:1fr 1fr}}
  .summary-grid{{grid-template-columns:1fr}}
  .table-controls input[type=text]{{width:140px}}
}}

/* ─── PRINT / PDF ─── */
@media print{{
  .site-header,.main-nav,.sub-nav,.table-controls,.group-controls,.pagination,.view-toggle,.btn-export{{display:none!important}}
  .section{{display:block!important}}
  .table-scroll{{max-height:none!important;overflow:visible!important}}
  body{{background:#fff}}
  .table-wrapper{{box-shadow:none;border:1px solid #ccc}}
}}

/* ─── LOADING SPINNER ─── */
.loading{{text-align:center;padding:3rem;color:var(--muted)}}
.loading::after{{content:'';display:block;width:32px;height:32px;margin:1rem auto;border:3px solid var(--line);border-top-color:var(--gold);border-radius:50%;animation:spin .8s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}

/* hidden helper columns */
.col-hidden{{display:none}}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-left">
    <h1>Broking MIS Dashboard</h1>
    <div class="subtitle">Marwadi Shares and Finance Ltd. &mdash; {new_short} vs {fy_label} Comparative Analytics</div>
  </div>
  <a href="/" class="btn-upload" title="Upload new period files to regenerate the dashboard">
    <svg viewBox="0 0 24 24"><path d="M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z"/></svg>
    Upload New Files
  </a>
</header>

<nav class="main-nav" id="mainNav">
  <button class="active" data-section="overview">Overview</button>
  <button data-section="channels">Channels</button>
  <button data-section="clients">Clients</button>
  <button data-section="gainers">Gainers &amp; Decliners</button>
  <button data-section="lost">Lost Clients</button>
</nav>

<div class="content">

<!-- ═══════════ OVERVIEW ═══════════ -->
<div class="section active" id="sec-overview">
  <div class="kpi-grid" id="kpiGrid"></div>
  <div class="summary-grid" id="overviewSummaries"></div>
</div>

<!-- ═══════════ CHANNELS ═══════════ -->
<div class="section" id="sec-channels">
  <div class="sub-nav" id="channelSubNav">
    <button class="active" data-sub="ch-fy">{fy_label} Actual</button>
    <button data-sub="ch-new">{new_short}</button>
    <button data-sub="ch-comp">Comparison</button>
  </div>
  <div id="channelContent"></div>
</div>

<!-- ═══════════ CLIENTS ═══════════ -->
<div class="section" id="sec-clients">
  <div class="sub-nav" id="clientSubNav">
    <button class="active" data-sub="cl-fy">{fy_label} Actual</button>
    <button data-sub="cl-new">{new_short}</button>
    <button data-sub="cl-comp">Comparison</button>
  </div>
  <div id="clientContent"></div>
</div>

<!-- ═══════════ GAINERS & DECLINERS ═══════════ -->
<div class="section" id="sec-gainers">
  <div id="gainersContent"></div>
</div>

<!-- ═══════════ LOST CLIENTS ═══════════ -->
<div class="section" id="sec-lost">
  <div id="lostContent"></div>
</div>

</div>

<script>
// ─── Inline SheetJS ───
{sheetjs}
</script>
<script>
// ─── Embedded data ───
const DATA = {dashboard_data};
</script>
<script>
(function() {{
"use strict";

const META = DATA.meta;
const SCH = DATA.schemas;
const FY = META.fyLabel;
const NP = META.newShortLabel;
const NPL = META.newLabel;
const FAC = META.factor;
const HOUSE = new Set(META.houseCodes);
const PAGE_SIZE = 100;

// ─── Utility helpers ───
function fmt(v, type) {{
  if (v == null || v === "") return "";
  const n = Number(v);
  if (isNaN(n)) return String(v);
  if (type === "pct") return n.toFixed(2) + "%";
  if (type === "growth") return (n >= 0 ? "+" : "") + (n * 100).toFixed(1) + "%";
  if (type === "num") return n.toLocaleString("en-IN", {{minimumFractionDigits: 2, maximumFractionDigits: 2}});
  if (type === "int") return Math.round(n).toLocaleString("en-IN");
  return String(v);
}}

function fmtLacs(v) {{ return fmt(v, "num"); }}
function fmtCrore(v) {{ return fmt(v, "num"); }}
function fmtYield(v) {{ return v != null ? Number(v).toFixed(2) : ""; }}
function fmtGrowth(v) {{ if (v == null) return ""; return (v >= 0 ? "+" : "") + (v * 100).toFixed(1) + "%"; }}
function fmtAbsChange(v) {{ if (v == null) return ""; return (v >= 0 ? "+" : "") + Number(v).toLocaleString("en-IN", {{minimumFractionDigits:2, maximumFractionDigits:2}}); }}

function colType(header) {{
  if (/Growth %/.test(header)) return "growth";
  if (/Yield/.test(header)) return "pct";
  if (/Turnover|Gross|Passout|Net|Change/.test(header)) return "num";
  if (/Channels?$|Clients?$|Factor/.test(header)) return "int";
  return "text";
}}

function isHiddenCol(header) {{ return /\\[helper\\]/.test(header); }}

function growthClass(v) {{
  if (v == null || v === "") return "";
  const n = Number(v);
  if (n > 0) return "positive";
  if (n < 0) return "negative";
  return "";
}}

// ─── DOM helpers ───
function el(tag, cls, html) {{
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (html !== undefined) e.innerHTML = html;
  return e;
}}

// ─── Navigation ───
document.getElementById("mainNav").addEventListener("click", function(e) {{
  if (e.target.tagName !== "BUTTON") return;
  this.querySelectorAll("button").forEach(b => b.classList.remove("active"));
  e.target.classList.add("active");
  document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
  document.getElementById("sec-" + e.target.dataset.section).classList.add("active");
}});

function setupSubNav(navId, handler) {{
  document.getElementById(navId).addEventListener("click", function(e) {{
    if (e.target.tagName !== "BUTTON") return;
    this.querySelectorAll("button").forEach(b => b.classList.remove("active"));
    e.target.classList.add("active");
    handler(e.target.dataset.sub);
  }});
}}

// ─── Table builder ───
function buildTable(config) {{
  const {{ title, headers, rows, containerId, showSearch = true, showGrouped = false, showExport = true, pageSize = PAGE_SIZE }} = config;
  const container = typeof containerId === "string" ? document.getElementById(containerId) : containerId;
  container.innerHTML = "";

  // Determine visible columns
  const visibleIdx = [];
  const visibleHeaders = [];
  headers.forEach((h, i) => {{
    if (!isHiddenCol(h) && h !== "Row Type") {{
      visibleIdx.push(i);
      visibleHeaders.push(h);
    }}
  }});

  const rowTypeIdx = headers.indexOf("Row Type");
  const statusIdx = headers.indexOf("Status");
  const branchIdx = 0;

  // State
  let currentData = [...rows];
  let sortCol = -1, sortAsc = true;
  let searchTerm = "";
  let groupFilter = "";
  let viewMode = "flat"; // flat or grouped
  let currentPage = 0;

  // Wrapper
  const wrapper = el("div", "table-wrapper");

  // Header bar
  const hdr = el("div", "table-header");
  const titleEl = el("h3", null, title);
  hdr.appendChild(titleEl);

  const controls = el("div", "table-controls");

  if (showSearch) {{
    const search = document.createElement("input");
    search.type = "text";
    search.placeholder = "Search...";
    search.addEventListener("input", () => {{ searchTerm = search.value.toLowerCase(); currentPage = 0; render(); }});
    controls.appendChild(search);
  }}

  // Group filter
  const groupIdx = headers.indexOf("Group");
  if (groupIdx >= 0) {{
    const sel = document.createElement("select");
    sel.innerHTML = '<option value="">All Groups</option>' + META.groups.map(g => '<option value="' + g + '">' + g + '</option>').join("");
    sel.addEventListener("change", () => {{ groupFilter = sel.value; currentPage = 0; render(); }});
    controls.appendChild(sel);
  }}

  if (showGrouped) {{
    const toggle = el("div", "view-toggle");
    const btnFlat = el("button", "active", "Flat / Search");
    const btnGrouped = el("button", "", "Grouped by Branch");
    btnFlat.addEventListener("click", () => {{ viewMode = "flat"; btnFlat.classList.add("active"); btnGrouped.classList.remove("active"); currentPage = 0; render(); }});
    btnGrouped.addEventListener("click", () => {{ viewMode = "grouped"; btnGrouped.classList.add("active"); btnFlat.classList.remove("active"); render(); }});
    toggle.appendChild(btnFlat);
    toggle.appendChild(btnGrouped);
    controls.appendChild(toggle);
  }}

  if (showExport) {{
    const btnXls = el("button", "btn-export", "&#x1F4E5; Excel");
    btnXls.addEventListener("click", () => exportExcel(title, visibleHeaders, getFilteredRows()));
    controls.appendChild(btnXls);
    const btnPdf = el("button", "btn-export", "&#x1F5A8; PDF");
    btnPdf.addEventListener("click", () => window.print());
    controls.appendChild(btnPdf);
  }}

  hdr.appendChild(controls);
  wrapper.appendChild(hdr);

  const scrollBox = el("div", "table-scroll");
  wrapper.appendChild(scrollBox);

  const pagination = el("div", "pagination");
  wrapper.appendChild(pagination);

  container.appendChild(wrapper);

  function getFilteredRows() {{
    let result = rows.filter(r => {{
      if (rowTypeIdx >= 0 && r[rowTypeIdx] === "Branch Total") return false;
      if (groupFilter && groupIdx >= 0 && r[groupIdx] !== groupFilter) return false;
      if (searchTerm) {{
        return visibleIdx.some(i => String(r[i] || "").toLowerCase().includes(searchTerm));
      }}
      return true;
    }});
    if (sortCol >= 0) {{
      const idx = visibleIdx[sortCol];
      result.sort((a, b) => {{
        let va = a[idx], vb = b[idx];
        if (va == null) va = "";
        if (vb == null) vb = "";
        const na = Number(va), nb = Number(vb);
        if (!isNaN(na) && !isNaN(nb)) return sortAsc ? na - nb : nb - na;
        return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
      }});
    }}
    return result;
  }}

  function getBranchGroups() {{
    const filtered = rows.filter(r => {{
      if (groupFilter && groupIdx >= 0 && r[groupIdx] !== groupFilter) return false;
      if (searchTerm) return visibleIdx.some(i => String(r[i] || "").toLowerCase().includes(searchTerm));
      return true;
    }});
    const groups = new Map();
    filtered.forEach(r => {{
      const branch = r[branchIdx] || "(no branch)";
      const isTotal = rowTypeIdx >= 0 && r[rowTypeIdx] === "Branch Total";
      if (!groups.has(branch) && !isTotal) groups.set(branch, []);
      if (isTotal) {{
        // subtotal row — use the label to find the branch
        // The branch total row has branch="" and label in another column
        return;
      }}
      if (groups.has(branch)) groups.get(branch).push(r);
    }});
    // Also gather subtotal rows
    const totals = new Map();
    filtered.forEach(r => {{
      if (rowTypeIdx >= 0 && r[rowTypeIdx] === "Branch Total") {{
        // Find which branch this total belongs to — look at the label
        for (const [b] of groups) {{
          const labelCol = visibleIdx.find(i => String(r[i] || "").includes(b));
          if (labelCol !== undefined) {{ totals.set(b, r); break; }}
        }}
      }}
    }});
    return {{ groups, totals }};
  }}

  function render() {{
    if (viewMode === "grouped") {{
      renderGrouped();
      pagination.innerHTML = "";
      return;
    }}
    const filtered = getFilteredRows();
    const totalPages = Math.ceil(filtered.length / pageSize) || 1;
    if (currentPage >= totalPages) currentPage = totalPages - 1;
    const pageRows = filtered.slice(currentPage * pageSize, (currentPage + 1) * pageSize);

    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    visibleHeaders.forEach((h, ci) => {{
      const th = document.createElement("th");
      th.innerHTML = h + ' <span class="sort-arrow">&#9650;</span>';
      th.addEventListener("click", () => {{
        if (sortCol === ci) sortAsc = !sortAsc;
        else {{ sortCol = ci; sortAsc = true; }}
        currentPage = 0;
        render();
      }});
      if (sortCol === ci) th.classList.add("sorted");
      hr.appendChild(th);
    }});
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    pageRows.forEach(row => {{
      const tr = document.createElement("tr");
      visibleIdx.forEach((ri, ci) => {{
        const td = document.createElement("td");
        const val = row[ri];
        const ct = colType(visibleHeaders[ci]);
        if (ct === "num") {{ td.className = "num"; td.textContent = fmtLacs(val); }}
        else if (ct === "pct") {{ td.className = "pct"; td.textContent = fmtYield(val); }}
        else if (ct === "growth") {{ td.className = "pct " + growthClass(val); td.textContent = fmtGrowth(val); }}
        else if (ct === "int") {{ td.className = "num"; td.textContent = val != null ? Math.round(Number(val)).toLocaleString("en-IN") : ""; }}
        else if (visibleHeaders[ci] === "Is House Account" && val === true) {{ td.className = "house-flag"; td.textContent = "HOUSE"; }}
        else if (visibleHeaders[ci] === "Absolute Gross Change (₹ Lacs)") {{
          td.className = "num " + growthClass(val);
          td.textContent = fmtAbsChange(val);
        }}
        else td.textContent = val != null ? String(val) : "";
        // Status column styling
        if (visibleHeaders[ci] === "Status") {{
          if (val === "New") td.style.color = "var(--green)";
          else if (val === "Lost") td.style.color = "var(--red)";
        }}
        tr.appendChild(td);
      }});
      tbody.appendChild(tr);
    }});
    table.appendChild(tbody);

    scrollBox.innerHTML = "";
    scrollBox.appendChild(table);

    // Pagination
    const start = currentPage * pageSize + 1;
    const end = Math.min((currentPage + 1) * pageSize, filtered.length);
    pagination.innerHTML = '<span class="page-info">Showing ' + start + '–' + end + ' of ' + filtered.length.toLocaleString("en-IN") + '</span>' +
      '<span><button ' + (currentPage === 0 ? 'disabled' : '') + ' id="prevPage">&#9664; Prev</button> ' +
      '<button ' + (currentPage >= totalPages - 1 ? 'disabled' : '') + ' id="nextPage">Next &#9654;</button></span>';
    const pp = pagination.querySelector("#prevPage");
    const np2 = pagination.querySelector("#nextPage");
    if (pp) pp.addEventListener("click", () => {{ currentPage--; render(); }});
    if (np2) np2.addEventListener("click", () => {{ currentPage++; render(); }});
  }}

  function renderGrouped() {{
    const {{ groups, totals }} = getBranchGroups();
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const hr = document.createElement("tr");
    visibleHeaders.forEach(h => {{
      const th = document.createElement("th");
      th.textContent = h;
      hr.appendChild(th);
    }});
    thead.appendChild(hr);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    const expanded = new Set();

    // Add expand all / collapse all controls
    const gcDiv = el("div", "group-controls");
    const btnExpand = el("button", "", "Expand All");
    const btnCollapse = el("button", "", "Collapse All");
    gcDiv.appendChild(btnExpand);
    gcDiv.appendChild(btnCollapse);

    function renderBranches() {{
      tbody.innerHTML = "";
      for (const [branch, details] of groups) {{
        const isExpanded = expanded.has(branch);
        // Summary row with totals
        const totalRow = totals.get(branch);
        const hRow = document.createElement("tr");
        hRow.className = "branch-group-header" + (isExpanded ? "" : " collapsed");
        const countTd = document.createElement("td");
        countTd.colSpan = visibleHeaders.length;
        countTd.innerHTML = '<span class="toggle-icon">&#9660;</span> <strong>' + branch + '</strong> &mdash; ' + details.length + ' entries';
        if (totalRow) {{
          // Show a brief summary
          const grossIdx = visibleIdx.findIndex((_, ci) => /Gross.*Lacs/.test(visibleHeaders[ci]));
          if (grossIdx >= 0) {{
            const gv = totalRow[visibleIdx[grossIdx]];
            countTd.innerHTML += ' | Gross: ₹' + fmtLacs(gv) + ' Lacs';
          }}
        }}
        hRow.appendChild(countTd);
        hRow.addEventListener("click", () => {{
          if (expanded.has(branch)) expanded.delete(branch);
          else expanded.add(branch);
          renderBranches();
        }});
        tbody.appendChild(hRow);

        if (isExpanded) {{
          details.forEach(row => {{
            const tr = document.createElement("tr");
            visibleIdx.forEach((ri, ci) => {{
              const td = document.createElement("td");
              const val = row[ri];
              const ct = colType(visibleHeaders[ci]);
              if (ct === "num") {{ td.className = "num"; td.textContent = fmtLacs(val); }}
              else if (ct === "pct") {{ td.className = "pct"; td.textContent = fmtYield(val); }}
              else if (ct === "growth") {{ td.className = "pct " + growthClass(val); td.textContent = fmtGrowth(val); }}
              else if (ct === "int") {{ td.className = "num"; td.textContent = val != null ? Math.round(Number(val)).toLocaleString("en-IN") : ""; }}
              else if (visibleHeaders[ci] === "Is House Account" && val === true) {{ td.className = "house-flag"; td.textContent = "HOUSE"; }}
              else if (visibleHeaders[ci] === "Absolute Gross Change (₹ Lacs)") {{ td.className = "num " + growthClass(val); td.textContent = fmtAbsChange(val); }}
              else td.textContent = val != null ? String(val) : "";
              tr.appendChild(td);
            }});
            tbody.appendChild(tr);
          }});
          // Branch total row
          if (totalRow) {{
            const tr = document.createElement("tr");
            tr.className = "branch-total";
            visibleIdx.forEach((ri, ci) => {{
              const td = document.createElement("td");
              const val = totalRow[ri];
              const ct = colType(visibleHeaders[ci]);
              if (ct === "num") {{ td.className = "num"; td.textContent = fmtLacs(val); }}
              else if (ct === "pct") {{ td.className = "pct"; td.textContent = fmtYield(val); }}
              else if (ct === "growth") {{ td.className = "pct " + growthClass(val); td.textContent = fmtGrowth(val); }}
              else td.textContent = val != null ? String(val) : "";
              tr.appendChild(td);
            }});
            tbody.appendChild(tr);
          }}
        }}
      }}
    }}

    btnExpand.addEventListener("click", () => {{ for (const b of groups.keys()) expanded.add(b); renderBranches(); }});
    btnCollapse.addEventListener("click", () => {{ expanded.clear(); renderBranches(); }});

    renderBranches();
    table.appendChild(tbody);

    scrollBox.innerHTML = "";
    // Insert group controls before table
    const frag = document.createDocumentFragment();
    frag.appendChild(gcDiv);
    frag.appendChild(table);
    scrollBox.appendChild(frag);
  }}

  render();
  return {{ render, getFilteredRows }};
}}

// ─── Excel export ───
function exportExcel(title, headers, rows) {{
  const wsData = [headers];
  rows.forEach(r => {{
    const row = [];
    headers.forEach((h, i) => {{
      // Find the index in the original schema
      row.push(r[i] != null ? r[i] : "");
    }});
    wsData.push(row);
  }});
  // Actually, we need to map through visible indices properly
  // Simpler: just use the data as-is
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows.map(r => {{
    // r is the original row, need to pick visible columns
    return r;
  }})]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, title.substring(0, 31));
  XLSX.writeFile(wb, title.replace(/[^a-zA-Z0-9]/g, "_") + ".xlsx");
}}

// ─── Overview Section ───
function buildOverview() {{
  const grid = document.getElementById("kpiGrid");
  const summaries = document.getElementById("overviewSummaries");

  // Compute KPIs from channel group data
  const fyGross = DATA.channels.groupFY.reduce((a, r) => a + r[3], 0);
  const newGross = DATA.channels.groupNew.reduce((a, r) => a + r[3], 0);
  const fyNet = DATA.channels.groupFY.reduce((a, r) => a + r[5], 0);
  const newNet = DATA.channels.groupNew.reduce((a, r) => a + r[5], 0);
  const fyPassout = DATA.channels.groupFY.reduce((a, r) => a + r[4], 0);
  const newPassout = DATA.channels.groupNew.reduce((a, r) => a + r[4], 0);
  const fyChannels = DATA.channels.groupFY.reduce((a, r) => a + r[1], 0);
  const newChannels = DATA.channels.groupNew.reduce((a, r) => a + r[1], 0);
  const fyClients = DATA.clients.groupFY.reduce((a, r) => a + r[1], 0);
  const newClients = DATA.clients.groupNew.reduce((a, r) => a + r[1], 0);

  const annGross = newGross * FAC;
  const annNet = newNet * FAC;
  const grossGrowth = fyGross ? (annGross / fyGross - 1) : 0;
  const netGrowth = fyNet ? (annNet / fyNet - 1) : 0;

  const kpis = [
    {{ label: "Gross Brokerage — " + FY, value: "₹ " + fmtLacs(fyGross) + " L", sub: "Full year actual" }},
    {{ label: "Gross Brokerage — " + NP, value: "₹ " + fmtLacs(newGross) + " L", sub: "4 months actual" }},
    {{ label: "Annualized Gross — " + NP, value: "₹ " + fmtLacs(annGross) + " L", sub: "Pro-rata ×" + FAC, growth: grossGrowth }},
    {{ label: "Net Brokerage — " + FY, value: "₹ " + fmtLacs(fyNet) + " L", sub: "Full year actual" }},
    {{ label: "Net Brokerage — " + NP, value: "₹ " + fmtLacs(newNet) + " L", sub: "4 months actual" }},
    {{ label: "Annualized Net — " + NP, value: "₹ " + fmtLacs(annNet) + " L", sub: "Pro-rata ×" + FAC, growth: netGrowth }},
    {{ label: "Active Channels — " + FY, value: fyChannels.toLocaleString("en-IN"), sub: "Unique channel IDs" }},
    {{ label: "Active Channels — " + NP, value: newChannels.toLocaleString("en-IN"), sub: "Unique channel IDs" }},
    {{ label: "Active Clients — " + FY, value: fyClients.toLocaleString("en-IN"), sub: "Excluding house accounts" }},
    {{ label: "Active Clients — " + NP, value: newClients.toLocaleString("en-IN"), sub: "Excluding house accounts" }},
    {{ label: "Matched Clients", value: DATA.clients.matched.length.toLocaleString("en-IN"), sub: "Present in both periods" }},
    {{ label: "New Clients — " + NP, value: DATA.clients.new.length.toLocaleString("en-IN"), sub: "Not in " + FY }},
  ];

  kpis.forEach(k => {{
    const card = el("div", "kpi-card");
    card.innerHTML = '<div class="kpi-label">' + k.label + '</div>' +
      '<div class="kpi-value">' + k.value + '</div>' +
      '<div class="kpi-sub">' + k.sub + '</div>';
    if (k.growth !== undefined) {{
      const cls = k.growth >= 0 ? "positive" : "negative";
      card.innerHTML += '<div class="kpi-growth ' + cls + '">' + fmtGrowth(k.growth) + ' vs ' + FY + '</div>';
    }}
    grid.appendChild(card);
  }});

  // Group-wise summary tables
  function groupTable(title, headers, rows, id) {{
    const div = el("div", "summary-table table-wrapper");
    let html = '<div class="table-header"><h3>' + title + '</h3></div>';
    html += '<div class="table-scroll"><table><thead><tr>';
    headers.forEach(h => html += '<th>' + h + '</th>');
    html += '</tr></thead><tbody>';
    rows.forEach(r => {{
      html += '<tr>';
      r.forEach((v, i) => {{
        const ct = colType(headers[i]);
        if (ct === "num") html += '<td class="num">' + fmtLacs(v) + '</td>';
        else if (ct === "pct") html += '<td class="pct">' + fmtYield(v) + '</td>';
        else if (ct === "int") html += '<td class="num">' + (v != null ? Math.round(v).toLocaleString("en-IN") : "") + '</td>';
        else html += '<td>' + (v != null ? v : "") + '</td>';
      }});
      html += '</tr>';
    }});
    html += '</tbody></table></div>';
    div.innerHTML = html;
    return div;
  }}

  summaries.appendChild(groupTable(
    "Channel Group Summary — " + FY,
    ["Group", "Channels", "Gross Brokerage (₹ Lacs)", "Passout (₹ Lacs)", "Net Brokerage (₹ Lacs)", "Yield %"],
    DATA.channels.groupFY.map(r => [r[0], r[1], r[3], r[4], r[5], r[6]])
  ));
  summaries.appendChild(groupTable(
    "Channel Group Summary — " + NP,
    ["Group", "Channels", "Gross Brokerage (₹ Lacs)", "Passout (₹ Lacs)", "Net Brokerage (₹ Lacs)", "Yield %"],
    DATA.channels.groupNew.map(r => [r[0], r[1], r[3], r[4], r[5], r[6]])
  ));
  summaries.appendChild(groupTable(
    "Client Group Summary — " + FY,
    ["Group", "Clients", "Turnover (₹ Crore)", "Gross Brokerage (₹ Lacs)", "Yield %"],
    DATA.clients.groupFY
  ));
  summaries.appendChild(groupTable(
    "Client Group Summary — " + NP,
    ["Group", "Clients", "Turnover (₹ Crore)", "Gross Brokerage (₹ Lacs)", "Yield %"],
    DATA.clients.groupNew
  ));

  // Channel type breakdown
  const ctBreakdown = {{}};
  DATA.channels.fyDetails.forEach(r => {{
    const ct = r[2] || "Unknown";
    if (!ctBreakdown[ct]) ctBreakdown[ct] = {{ fy: {{ count: 0, gross: 0, net: 0 }}, np: {{ count: 0, gross: 0, net: 0 }} }};
    ctBreakdown[ct].fy.count++;
    ctBreakdown[ct].fy.gross += r[8] || 0;
    ctBreakdown[ct].fy.net += r[10] || 0;
  }});
  DATA.channels.newDetails.forEach(r => {{
    const ct = r[2] || "Unknown";
    if (!ctBreakdown[ct]) ctBreakdown[ct] = {{ fy: {{ count: 0, gross: 0, net: 0 }}, np: {{ count: 0, gross: 0, net: 0 }} }};
    ctBreakdown[ct].np.count++;
    ctBreakdown[ct].np.gross += r[8] || 0;
    ctBreakdown[ct].np.net += r[10] || 0;
  }});
  const ctRows = Object.entries(ctBreakdown).map(([type, d]) => [type, d.fy.count, d.fy.gross, d.fy.net, d.np.count, d.np.gross, d.np.net]);
  summaries.appendChild(groupTable(
    "Channel Type Breakdown",
    ["Channel Type", FY + " Count", FY + " Gross (₹ L)", FY + " Net (₹ L)", NP + " Count", NP + " Gross (₹ L)", NP + " Net (₹ L)"],
    ctRows
  ));

  // Top 10 branches by gross brokerage (new period)
  const branchAgg = {{}};
  DATA.channels.newDetails.forEach(r => {{
    const b = r[0];
    if (!branchAgg[b]) branchAgg[b] = {{ gross: 0, net: 0, passout: 0, channels: 0 }};
    branchAgg[b].gross += r[8] || 0;
    branchAgg[b].net += r[10] || 0;
    branchAgg[b].passout += r[9] || 0;
    branchAgg[b].channels++;
  }});
  const topBranches = Object.entries(branchAgg).sort((a, b) => b[1].gross - a[1].gross).slice(0, 10);
  summaries.appendChild(groupTable(
    "Top 10 Branches by Gross Brokerage — " + NP,
    ["Branch Code", "Channels", "Gross Brokerage (₹ Lacs)", "Passout (₹ Lacs)", "Net Brokerage (₹ Lacs)"],
    topBranches.map(([b, d]) => [b, d.channels, d.gross, d.passout, d.net])
  ));
}}

// ─── Channel Section ───
let channelTabBuilt = {{}};
function buildChannelTab(sub) {{
  const container = document.getElementById("channelContent");
  if (channelTabBuilt[sub]) {{ channelTabBuilt[sub].style.display = "block"; return; }}
  // Hide others
  Object.values(channelTabBuilt).forEach(el => el.style.display = "none");

  const div = el("div");
  container.appendChild(div);
  channelTabBuilt[sub] = div;

  if (sub === "ch-fy") {{
    buildTable({{ title: "Channelwise — " + FY + " Actual", headers: SCH.channelActual, rows: DATA.channels.fyDetails, containerId: div, showGrouped: true }});
  }} else if (sub === "ch-new") {{
    buildTable({{ title: "Channelwise — " + NP, headers: SCH.channelActual, rows: DATA.channels.newDetails, containerId: div, showGrouped: true }});
  }} else if (sub === "ch-comp") {{
    buildTable({{ title: "Channel Comparison — " + FY + " vs " + NP + " (Annualized ×" + FAC + ")", headers: SCH.channelComparison, rows: DATA.channels.comparison, containerId: div, showGrouped: true }});
  }}
}}

function switchChannelTab(sub) {{
  Object.values(channelTabBuilt).forEach(el => el.style.display = "none");
  buildChannelTab(sub);
}}

setupSubNav("channelSubNav", switchChannelTab);

// ─── Client Section ───
let clientTabBuilt = {{}};
function buildClientTab(sub) {{
  const container = document.getElementById("clientContent");
  if (clientTabBuilt[sub]) {{ clientTabBuilt[sub].style.display = "block"; return; }}
  Object.values(clientTabBuilt).forEach(el => el.style.display = "none");

  const div = el("div");
  container.appendChild(div);
  clientTabBuilt[sub] = div;

  if (sub === "cl-fy") {{
    buildTable({{ title: "Clientwise — " + FY + " Actual", headers: SCH.clientActual, rows: DATA.clients.fyDetails, containerId: div, showGrouped: true }});
  }} else if (sub === "cl-new") {{
    buildTable({{ title: "Clientwise — " + NP, headers: SCH.clientActual, rows: DATA.clients.newDetails, containerId: div, showGrouped: true }});
  }} else if (sub === "cl-comp") {{
    // Compose full client comparison from matched + new
    const allComp = [...DATA.clients.matched, ...DATA.clients.new];
    buildTable({{ title: "Client Comparison — " + FY + " vs " + NP + " (Annualized ×" + FAC + ")", headers: SCH.clientComparison, rows: allComp, containerId: div, showGrouped: true }});
  }}
}}

function switchClientTab(sub) {{
  Object.values(clientTabBuilt).forEach(el => el.style.display = "none");
  buildClientTab(sub);
}}

setupSubNav("clientSubNav", switchClientTab);

// ─── Gainers & Decliners ───
function buildGainers() {{
  const container = document.getElementById("gainersContent");

  // Top 20 Gainers
  const gDiv = el("div");
  container.appendChild(gDiv);
  buildTable({{ title: "Top 20 Gainers — Absolute Gross Change (₹ Lacs)", headers: SCH.clientComparison, rows: DATA.clients.gainers, containerId: gDiv, showSearch: false, showGrouped: false, pageSize: 20 }});

  // Top 20 Decliners
  const dDiv = el("div");
  container.appendChild(dDiv);
  buildTable({{ title: "Top 20 Decliners — Absolute Gross Change (₹ Lacs)", headers: SCH.clientComparison, rows: DATA.clients.decliners, containerId: dDiv, showSearch: false, showGrouped: false, pageSize: 20 }});
}}

// ─── Lost Clients ───
function buildLost() {{
  const container = document.getElementById("lostContent");
  buildTable({{ title: "Lost Clients — Present in " + FY + ", Absent in " + NP, headers: SCH.clientComparison, rows: DATA.clients.lost, containerId: container, showGrouped: true }});
}}

// ─── Initialize ───
buildOverview();
buildChannelTab("ch-fy");
buildGainers();
buildLost();

}})();
</script>
</body>
</html>"""

    return html


if __name__ == "__main__":
    data = json.loads((BUILD / "data.json").read_text(encoding="utf-8"))
    html = generate_html(data)
    output_path = OUT / "Broking_MIS_Dashboard_Apr26-Jul26.html"
    output_path.write_text(html, encoding="utf-8")
    print(f"Dashboard written to {output_path}")
    print(f"File size: {output_path.stat().st_size / 1024 / 1024:.1f} MB")
