# Système de design partagé entre tous les dashboards
DESIGN_CSS = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#F8F8F6;--surface:#FFFFFF;--border:#E8E6E0;--border-mid:#D4D2CC;
  --text:#1A1A18;--text-2:#6B6B67;--text-3:#9B9B96;
  --blue:#185FA5;--blue-l:#E6F1FB;--blue-b:#B5D4F4;
  --green:#3B6D11;--green-l:#EAF3DE;
  --red:#A32D2D;--red-l:#FCEBEB;
  --amber:#854F0B;--amber-l:#FAEEDA;
  --litt:#534AB7;--litt-l:#EEEDFE;
  --geodp:#0F6E56;--geodp-l:#E1F5EE;
}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.5}

/* ── Topbar ── */
.topbar{background:var(--surface);border-bottom:1px solid var(--border);height:48px;display:flex;align-items:center;padding:0 24px;gap:0;position:sticky;top:0;z-index:100}
.topbar-brand{font-size:13px;font-weight:500;margin-right:32px;white-space:nowrap;color:var(--text)}
.topbar-brand span{color:var(--text-2);font-weight:400}
.nav-tabs{display:flex;align-items:stretch;height:100%;gap:0;flex:1}
.nav-tab{display:flex;align-items:center;gap:6px;padding:0 14px;font-size:13px;color:var(--text-2);border-bottom:2px solid transparent;text-decoration:none;white-space:nowrap;transition:color .15s}
.nav-tab:hover{color:var(--text)}
.nav-tab.active{color:var(--text);border-bottom-color:var(--blue);font-weight:500}
.nav-tab svg{width:15px;height:15px;stroke:currentColor;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
.topbar-right{display:flex;align-items:center;gap:8px;margin-left:auto}
.badge-date{font-size:11px;color:var(--text-3);background:var(--bg);border:1px solid var(--border);padding:3px 10px;border-radius:99px;font-family:'DM Mono',monospace}
.btn-refresh{display:flex;align-items:center;gap:5px;font-size:12px;font-weight:500;padding:5px 12px;border-radius:8px;cursor:pointer;background:var(--blue-l);color:var(--blue);border:1px solid var(--blue-b);transition:opacity .15s;text-decoration:none}
.btn-refresh:hover{opacity:.8}

/* ── Layout ── */
.page{max-width:1100px;margin:0 auto;padding:28px 24px}
.page-header{margin-bottom:24px}
.page-title{font-size:20px;font-weight:500}
.page-sub{font-size:13px;color:var(--text-2);margin-top:3px}

/* ── Scope tabs ── */
.scope-wrap{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px}
.scope-tabs{display:flex;gap:3px;background:var(--border);padding:3px;border-radius:9px;width:fit-content}
.scope-tab{padding:5px 16px;font-size:12px;font-weight:500;border-radius:7px;cursor:pointer;color:var(--text-2);border:none;background:transparent;transition:all .15s}
.scope-tab.active{background:var(--surface);color:var(--text);box-shadow:0 1px 2px rgba(0,0,0,.06)}

/* ── KPI cards ── */
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:24px}
.kpi{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 18px}
.kpi-label{font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.05em;color:var(--text-3);margin-bottom:8px}
.kpi-value{font-size:22px;font-weight:500;font-family:'DM Mono',monospace}
.kpi-value.green{color:var(--green)}
.kpi-value.red{color:var(--red)}
.kpi-value.amber{color:var(--amber)}
.kpi-trend{display:inline-flex;align-items:center;gap:3px;font-size:11px;margin-top:6px;padding:2px 8px;border-radius:99px;font-weight:500}
.kpi-trend.up{background:var(--green-l);color:var(--green)}
.kpi-trend.down{background:var(--red-l);color:var(--red)}
.kpi-trend.flat{background:var(--bg);color:var(--text-3);border:1px solid var(--border)}

/* ── Cards ── */
.card{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:16px}
.card-head{padding:12px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between}
.card-title{font-size:13px;font-weight:500}
.card-sub{font-size:12px;color:var(--text-2)}
.card-body{padding:16px}

/* ── Section label ── */
.section-label{font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.06em;color:var(--text-3);margin-bottom:12px}

/* ── Bar chart ── */
.bar-row{display:flex;align-items:center;gap:10px;margin-bottom:10px;font-size:12px}
.bar-label{width:110px;color:var(--text-2);flex-shrink:0;text-align:right}
.bar-track{flex:1;height:8px;background:var(--bg);border-radius:4px;overflow:hidden;border:1px solid var(--border)}
.bar-fill{height:100%;border-radius:4px;transition:width .6s ease}
.bar-value{width:80px;text-align:right;font-family:'DM Mono',monospace;font-size:11px;color:var(--text)}

/* ── Table ── */
.data-table{width:100%;border-collapse:collapse;font-size:12px}
.data-table th{padding:8px 12px;text-align:left;font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.04em;color:var(--text-3);border-bottom:1px solid var(--border)}
.data-table th.r,.data-table td.r{text-align:right}
.data-table td{padding:9px 12px;border-bottom:1px solid var(--bg);font-size:12px;color:var(--text)}
.data-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:var(--bg)}
.pill{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:500}
.pill-red{background:var(--red-l);color:var(--red)}
.pill-amber{background:var(--amber-l);color:var(--amber)}
.pill-green{background:var(--green-l);color:var(--green)}
.pill-blue{background:var(--blue-l);color:var(--blue)}
.pill-litt{background:var(--litt-l);color:var(--litt)}
.pill-geodp{background:var(--geodp-l);color:var(--geodp)}

/* ── Grid ── */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px}

/* ── Progress bar ── */
.prog-wrap{margin-top:12px}
.prog-meta{display:flex;justify-content:space-between;font-size:11px;color:var(--text-2);margin-bottom:5px}
.prog-track{height:10px;background:var(--bg);border-radius:5px;overflow:hidden;border:1px solid var(--border)}
.prog-fill{height:100%;border-radius:5px;transition:width .6s ease}

@media(max-width:700px){
  .g2{grid-template-columns:1fr}
  .topbar{padding:0 12px}
  .nav-tab span{display:none}
  .page{padding:16px}
}
</style>
"""

NAV_HTML = """
<nav class="topbar">
  <div class="topbar-brand">PS France <span>/ Tableau de bord</span></div>
  <div class="nav-tabs">
    <a class="nav-tab {a_home}" href="index.html">
      <svg viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
      <span>Accueil</span>
    </a>
    <a class="nav-tab {a_obj}" href="objectifs.html">
      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>
      <span>Objectifs</span>
    </a>
    <a class="nav-tab {a_global}" href="rapport_global.html">
      <svg viewBox="0 0 24 24"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>
      <span>Rapport global</span>
    </a>
    <a class="nav-tab {a_backlog}" href="backlog.html">
      <svg viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
      <span>Backlog</span>
    </a>
    <a class="nav-tab {a_charge}" href="charge.html">
      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
      <span>Charge</span>
    </a>
  </div>
  <div class="topbar-right">
    <span class="badge-date">Mis à jour le {today}</span>
    <a href="index.html#actualiser" class="btn-refresh">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>
      Actualiser
    </a>
  </div>
</nav>
"""

def nav(active, today):
    tabs = {"home":"","obj":"","global":"","backlog":"","charge":""}
    tabs[active] = "active"
    return NAV_HTML.format(
        a_home=tabs["home"], a_obj=tabs["obj"], a_global=tabs["global"],
        a_backlog=tabs["backlog"], a_charge=tabs["charge"], today=today
    )

def fmt_eur(v):
    if v is None or v == 0: return "—"
    return f"{round(v):,} €".replace(",", "\u202f")

def fmt_h(v):
    if v is None or v == 0: return "—"
    return f"{v:.1f} h"

def fmt_pct(v):
    if v is None: return "—"
    return f"{v:.1f} %"

def trend_pill(pct):
    if pct is None: return '<span class="kpi-trend flat">—</span>'
    sign = "+" if pct >= 0 else ""
    arrow = "▲" if pct >= 0 else "▼"
    cls = "up" if pct >= 0 else "down"
    return f'<span class="kpi-trend {cls}">{arrow} {sign}{pct:.1f}% vs mois préc.</span>'
