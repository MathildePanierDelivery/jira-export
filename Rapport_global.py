"""
Rapport Global — Dashboard HTML
---------------------------------
Lit directement le fichier Excel mensuel (jira_export_*.xlsx),
met à jour l'historique cumulé, et génère un dashboard HTML interactif.

Usage :
  python Rapport_global.py jira_export_2026-03-06.xlsx
  python Rapport_global.py 2026-02_data_mois.xlsx
  python Rapport_global.py                              # cherche le plus récent
"""

import os, sys, glob, json, re
from datetime import datetime
import pandas as pd

HISTORIQUE_FILE   = "_historique.xlsx"
RAPPORT_HTML_FILE = "rapport_global.html"
MOIS_FR = ["Janvier","Février","Mars","Avril","Mai","Juin",
           "Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
DATA_KEYS = ["ca", "ca_deal", "commandes", "charge", "backlog", "anciennete"]

# Objectif annuel de CA
OBJECTIF_ANNUEL = 1_316_750.00

# ══════════════════════════════════════════
# 1. TROUVER LE FICHIER
# ══════════════════════════════════════════
if len(sys.argv) > 1:
    src = sys.argv[1]
else:
    candidates = sorted(glob.glob("jira_export_*.xlsx") + glob.glob("*data_mois*.xlsx"))
    if not candidates:
        print("❌ Aucun fichier trouvé. Passez-le en argument.")
        sys.exit(1)
    src = candidates[-1]
if not os.path.exists(src):
    print(f"❌ {src} introuvable."); sys.exit(1)
print(f"📥 {src}")

xls = pd.ExcelFile(src)
all_sheets = xls.sheet_names

# Détecter mois/année
raw0 = pd.read_excel(src, sheet_name=0, header=None)
title = str(raw0.iloc[0, 0]) if not raw0.empty else ""
mois_label = next((m for m in MOIS_FR if m.lower() in title.lower()), "Janvier")
ym = re.search(r"20\d{2}", title)
annee_val = int(ym.group()) if ym else datetime.now().year
print(f"   → {mois_label} {annee_val}")

def find_hdr(raw, marker):
    for i, row in raw.iterrows():
        if marker in [str(v).strip() for v in row.values if pd.notna(v)]:
            return i
    return None

def read_with_hdr(filepath, sheet, marker):
    try: raw = pd.read_excel(filepath, sheet_name=sheet, header=None)
    except: return pd.DataFrame()
    hr = find_hdr(raw, marker)
    if hr is None: return pd.DataFrame()
    df = pd.read_excel(filepath, sheet_name=sheet, header=hr)
    df = df.dropna(how="all")
    df = df[~df.iloc[:,0].astype(str).str.contains("TOTAL|RÉSUMÉ", na=False)]
    return df

def safe_float(v):
    s = str(v).replace("€","").replace(",","").replace("\xa0","").replace("+","").replace("%","").strip()
    if not s or s == "—" or s == "nan": return 0
    try: return float(s)
    except: return 0

# ══════════════════════════════════════════
# 2. EXTRACTION
# ══════════════════════════════════════════

# ── CA ──
df_ca = pd.DataFrame()
try:
    if "Historique CA" in all_sheets:
        df_h = read_with_hdr(src, "Historique CA", "Mois")
        if not df_h.empty and mois_label in df_h["Mois"].values:
            r = df_h[df_h["Mois"]==mois_label].iloc[0]
            row = {"Mois":mois_label, "Année":annee_val,
                   "CA réalisé (€)": safe_float(r.get("CA réalisé (€)",0)),
                   "CA LITTERALIS (€)": safe_float(r.get("CA LITT. (€)",0)),
                   "CA GEODP (€)": safe_float(r.get("CA GEODP (€)",0)),
                   "Objectif global (€)": safe_float(r.get("Objectif (€)",0)),
                   "CA N-1 global (€)": safe_float(r.get("CA N-1 (€)",0))}
            df_ca = pd.DataFrame([row])
    if "Performance CA du mois" in all_sheets and not df_ca.empty:
        raw = pd.read_excel(src, sheet_name="Performance CA du mois", header=None)
        for i, row in raw.iterrows():
            vals = [v for v in row.values if pd.notna(v)]
            if len(vals) >= 4:
                sol = str(vals[0]).strip()
                if sol in ("LITTERALIS","GEODP") and isinstance(vals[1], (int,float)):
                    obj_val = safe_float(vals[2])
                    n1_val = safe_float(vals[3]) if len(vals) > 3 else 0
                    if obj_val > 1000:
                        df_ca[f"Objectif {sol} (€)"] = obj_val
                        df_ca[f"CA N-1 {sol} (€)"] = n1_val
except Exception as e:
    print(f"  ⚠️ CA: {e}")
print(f"  CA: {len(df_ca)} lignes")

# ── CA par deal ──
df_deal = pd.DataFrame()
try:
    if "Performance CA du mois" in all_sheets:
        raw = pd.read_excel(src, sheet_name="Performance CA du mois", header=None)
        rows = []
        for i, row in raw.iterrows():
            vals = [v for v in row.values if pd.notna(v)]
            if len(vals) >= 2:
                deal = str(vals[0]).strip()
                if any(k in deal for k in ["Nouveau","Vente","Migration"]):
                    ca = safe_float(vals[1])
                    if ca > 0:
                        sol = "GLOBAL"
                        for j in range(i-1, max(i-10,0), -1):
                            c = str(raw.iloc[j,0]) if pd.notna(raw.iloc[j,0]) else ""
                            if "GEODP" in c.upper(): sol="GEODP"; break
                            elif "LITT" in c.upper() or "ittéralis" in c: sol="LITTERALIS"; break
                            elif "Global" in c: sol="GLOBAL"; break
                        rows.append({"Mois":mois_label,"Année":annee_val,"Type de deal":deal,"Solution":sol,"CA (€)":ca})
        if rows: df_deal = pd.DataFrame(rows)
except Exception as e:
    print(f"  ⚠️ CA deal: {e}")
print(f"  CA deal: {len(df_deal)} lignes")

# ── Commandes ──
df_cmd = pd.DataFrame()
try:
    if "Commandes du mois" in all_sheets:
        df = read_with_hdr(src, "Commandes du mois", "Date de création")
        if not df.empty:
            df.insert(0,"Année",annee_val); df.insert(0,"Mois",mois_label)
            df_cmd = df
except Exception as e:
    print(f"  ⚠️ Cmd: {e}")
print(f"  Commandes: {len(df_cmd)} lignes")

# ── Charge ──
df_charge = pd.DataFrame()
try:
    ch = {"Mois":mois_label,"Année":annee_val}
    for sn in all_sheets:
        if "Charge Littéralis" in sn or "Charge GEODP" in sn:
            raw = pd.read_excel(src, sheet_name=sn, header=None)
            pfx = "Litt" if "Littéralis" in sn else "GEODP"
            for i, row in raw.iterrows():
                vals = [v for v in row.values if pd.notna(v)]
                if len(vals)>=5 and str(vals[0]).strip() in ["Littéralis","GEODP"]:
                    ch[f"Productif {pfx}"] = safe_float(vals[1])
                    ch[f"Support {pfx}"] = safe_float(vals[4])
            for i, row in raw.iterrows():
                vals = [str(v) for v in row.values if pd.notna(v)]
                if any("dont Rework" in v for v in vals):
                    for j in range(i+1, min(i+15, len(raw))):
                        rvals = [v for v in raw.iloc[j].values if pd.notna(v)]
                        if len(rvals) >= 3 and mois_label in str(rvals[0]):
                            ch[f"Rework {pfx}"] = safe_float(rvals[2])
                            break
    if "Charge globale" in all_sheets:
        raw = pd.read_excel(src, sheet_name="Charge globale", header=None)
        for i, row in raw.iterrows():
            vals = [v for v in row.values if pd.notna(v)]
            if len(vals)>=2:
                label = str(vals[0]).strip()
                if label == "Rework":     ch["Rework Total"] = safe_float(vals[1])
                elif label == "Interne":  ch["Interne"] = safe_float(vals[1])
                elif label == "Support":  ch["Support Total"] = safe_float(vals[1])
                elif label == "Productif": ch["Productif Total"] = safe_float(vals[1])
    for k in ["Productif Litt","Productif GEODP","Support Litt","Support GEODP",
              "Rework Litt","Rework GEODP","Rework Total","Interne",
              "Support Total","Productif Total"]:
        ch.setdefault(k,0)
    df_charge = pd.DataFrame([ch])
except Exception as e:
    print(f"  ⚠️ Charge: {e}")
print(f"  Charge: {len(df_charge)} lignes")

# ── Backlog ──
df_bl = pd.DataFrame()
try:
    bl = {"Mois":mois_label,"Année":annee_val}
    for sn in all_sheets:
        if "Backlog" not in sn: continue
        sn_up = sn.upper()
        if "TOTAL" in sn_up: key = "TOTAL"
        elif "LITT" in sn_up: key = "LITT"
        elif "GEODP" in sn_up: key = "GEODP"
        else: continue
        raw = pd.read_excel(src, sheet_name=sn, header=None)
        for i, row in raw.iterrows():
            vals = [v for v in row.values if pd.notna(v)]
            if not vals: continue
            lab = str(vals[0]).strip()
            if lab=="Total backlog" and len(vals)>=2: bl[f"total_backlog_{key}"] = safe_float(vals[1])
            elif lab=="Total bloqué" and len(vals)>=2: bl[f"total_bloque_{key}"] = safe_float(vals[1])
            elif lab=="Total mobilisable" and len(vals)>=2: bl[f"total_mobilisable_{key}"] = safe_float(vals[1])
            elif lab=="Epic ouverts" and len(vals)>=2: bl[f"nb_epics_{key}"] = int(safe_float(vals[1]))
            elif lab=="Projets bloqués" and len(vals)>=2: bl[f"nb_bloques_{key}"] = int(safe_float(vals[1]))
            elif lab=="Projets" and f"nb_epics_{key}" in bl and len(vals)>=2: bl[f"nb_projets_{key}"] = int(safe_float(vals[1]))
            elif lab=="Rework" and f"nb_epics_{key}" in bl and len(vals)>=2: bl[f"nb_rework_{key}"] = int(safe_float(vals[1]))
    if len(bl)>2: df_bl = pd.DataFrame([bl])
except Exception as e:
    print(f"  ⚠️ Backlog: {e}")
print(f"  Backlog: {len(df_bl)} lignes")

# ── Ancienneté du CA ──
EPIC_DATES_FILE = "_epic_dates.xlsx"

df_anc = pd.DataFrame()
try:
    if "Suivi de production" in all_sheets:
        df_sp = read_with_hdr(src, "Suivi de production", "Epic")
        if not df_sp.empty and "Montant déclaré ce mois" in df_sp.columns:
            df_with_ca = df_sp[df_sp["Montant déclaré ce mois"].apply(safe_float) > 0].copy()
            has_anc_cols = "Ancienneté (mois)" in df_sp.columns and "Date commande" in df_sp.columns

            # Charger le référentiel _epic_dates.xlsx (généré par fetch_epic_dates.py)
            epic_lookup = {}
            if os.path.exists(EPIC_DATES_FILE):
                df_dates = pd.read_excel(EPIC_DATES_FILE)
                epic_lookup = dict(zip(df_dates["Epic"].astype(str), df_dates["Date commande"].astype(str)))
                print(f"  📁 Référentiel {EPIC_DATES_FILE} chargé ({len(epic_lookup)} Epics)")
            elif not has_anc_cols:
                print(f"  ⚠️ {EPIC_DATES_FILE} introuvable — lancez d'abord : python fetch_epic_dates.py")

            if not df_with_ca.empty:
                mois_ref = datetime(annee_val, MOIS_FR.index(mois_label) + 1, 1)
                rows_anc = []
                for _, r in df_with_ca.iterrows():
                    epic_key = str(r.get("Epic",""))

                    if has_anc_cols and pd.notna(r.get("Date commande")) and str(r.get("Date commande","")).strip():
                        date_cmd = str(r["Date commande"])
                        anc = int(safe_float(r.get("Ancienneté (mois)", 0)))
                    elif epic_key in epic_lookup:
                        date_cmd = epic_lookup[epic_key]
                        try:
                            dt = pd.to_datetime(date_cmd, dayfirst=True)
                            anc = (mois_ref.year - dt.year) * 12 + (mois_ref.month - dt.month)
                        except:
                            anc = 0
                    else:
                        date_cmd = ""
                        anc = 0

                    rows_anc.append({
                        "Mois": mois_label, "Année": annee_val,
                        "Epic": epic_key,
                        "Nom": str(r.get("Epic Nom",""))[:55],
                        "Solution": str(r.get("Solution","")),
                        "Catégorie": str(r.get("Catégorie","")),
                        "Type de deal": str(r.get("Type de deal","")),
                        "Date commande": date_cmd,
                        "Ancienneté": anc,
                        "CA": round(safe_float(r.get("Montant déclaré ce mois",0)), 2),
                    })
                df_anc = pd.DataFrame(rows_anc)
except Exception as e:
    print(f"  ⚠️ Ancienneté: {e}")
print(f"  Ancienneté CA: {len(df_anc)} lignes")

# ══════════════════════════════════════════
# 3. HISTORIQUE
# ══════════════════════════════════════════
def upsert(old, new):
    if new.empty: return old
    m,a = new.iloc[0].get("Mois",""), new.iloc[0].get("Année",0)
    if old.empty: return new.copy()
    return pd.concat([old[~((old["Mois"]==m)&(old["Année"]==a))], new], ignore_index=True)

hist = {}
if os.path.exists(HISTORIQUE_FILE):
    for k in DATA_KEYS:
        try: hist[k] = pd.read_excel(HISTORIQUE_FILE, sheet_name=k)
        except: hist[k] = pd.DataFrame()
else:
    hist = {k: pd.DataFrame() for k in DATA_KEYS}

new = {"ca":df_ca,"ca_deal":df_deal,"commandes":df_cmd,"charge":df_charge,"backlog":df_bl,"anciennete":df_anc}
for k in DATA_KEYS:
    hist[k] = upsert(hist[k], new[k])

with pd.ExcelWriter(HISTORIQUE_FILE, engine="openpyxl") as w:
    for k,df in hist.items():
        if not df.empty: df.to_excel(w, sheet_name=k, index=False)
print(f"✅ Historique sauvegardé")

# ══════════════════════════════════════════
# 4. HTML
# ══════════════════════════════════════════
def to_json(df):
    if df.empty: return []
    return df.fillna("").to_dict(orient="records")

dj = {k: to_json(df) for k,df in hist.items()}
periods = sorted(set(
    (str(r["Mois"]), int(float(str(r["Année"]))))
    for recs in dj.values() for r in recs if r.get("Mois") and r.get("Année")
), key=lambda x: (x[1], MOIS_FR.index(x[0]) if x[0] in MOIS_FR else 0))

HTML = r"""<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Rapport Global — COORDIN</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
:root{--b9:#1F3864;--b7:#2E5F9E;--b6:#2E75B6;--b1:#DEEAF1;--b0:#EBF3FB;--gn:#375623;--gnb:#E2EFDA;--rd:#C00000;--rdb:#FCE4D6;--or:#C55A11;--g0:#F8F9FA;--g2:#E5E7EB;--g4:#9CA3AF;--g6:#595959;--g8:#1F2937;--sh:0 1px 3px rgba(0,0,0,.08);--r:8px}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'DM Sans',system-ui,sans-serif;background:var(--g0);color:var(--g8);line-height:1.5}
.hdr{background:linear-gradient(135deg,var(--b9),var(--b7));color:#fff;padding:24px 32px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px}
.hdr h1{font-size:22px;font-weight:700}.hdr .sub{font-size:13px;opacity:.75;margin-top:2px}
.sel{display:flex;align-items:center;gap:10px;background:rgba(255,255,255,.15);padding:8px 16px;border-radius:6px}
.sel label{font-size:13px;font-weight:600}
.sel select{padding:6px 12px;border:1px solid rgba(255,255,255,.3);border-radius:4px;background:rgba(255,255,255,.2);color:#fff;font-family:inherit;font-size:14px;font-weight:600;cursor:pointer}
.sel select option{color:var(--g8);background:#fff}
.tabs-nav{display:flex;gap:0;background:rgba(255,255,255,.1);border-radius:6px;overflow:hidden}
.tab-btn{padding:8px 20px;font-family:inherit;font-size:13px;font-weight:600;color:rgba(255,255,255,.7);background:transparent;border:none;cursor:pointer;transition:all .2s;white-space:nowrap}
.tab-btn:hover{color:#fff;background:rgba(255,255,255,.15)}
.tab-btn.active{color:#fff;background:rgba(255,255,255,.25)}
.tab-panel{display:none}.tab-panel.active{display:block}
.cnt{max-width:1280px;margin:0 auto;padding:24px}
.section{margin-bottom:24px}
.section-title{font-size:15px;font-weight:700;color:var(--b9);margin-bottom:12px;padding-left:2px;letter-spacing:-.2px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
.cd{background:#fff;border-radius:var(--r);box-shadow:var(--sh);overflow:hidden}
.ch{background:var(--b9);color:#fff;padding:10px 16px;font-size:13px;font-weight:600}
.ch.sub{background:var(--b6)}
.cb{padding:14px 16px;overflow-x:auto}
.kr{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:24px}
.kp{background:#fff;border-radius:var(--r);box-shadow:var(--sh);padding:14px 16px;text-align:center}
.kl{font-size:10px;color:var(--g4);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.kv{font-size:24px;font-weight:700;color:var(--b9);margin-top:3px}
.kv.gn{color:var(--gn)}.kv.rd{color:var(--rd)}.kv.or{color:var(--or)}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:var(--b1);color:var(--b9);font-weight:600;text-align:left;padding:7px 10px;border-bottom:2px solid var(--b6);white-space:nowrap}
th.r,td.r{text-align:right}
td{padding:6px 10px;border-bottom:1px solid var(--g2)}
tr:nth-child(even) td{background:var(--b0)}
tr:hover td{background:var(--b1)}
.tg{display:inline-block;padding:2px 7px;border-radius:3px;font-size:11px;font-weight:600}
.tg-g{background:var(--gnb);color:var(--gn)}.tg-r{background:var(--rdb);color:var(--rd)}
.es{text-align:center;padding:24px;color:var(--g4);font-style:italic;font-size:13px}
.tr-total td{border-top:2px solid var(--b6);font-weight:700;background:var(--b1) !important}
.accordion{background:#fff;border-radius:var(--r);box-shadow:var(--sh);overflow:hidden}
.accordion-toggle{display:flex;align-items:center;justify-content:space-between;padding:12px 18px;cursor:pointer;user-select:none;background:var(--b9);color:#fff;font-size:14px;font-weight:600}
.accordion-toggle:hover{background:var(--b7)}
.accordion-toggle .arrow{transition:transform .2s;font-size:12px}
.accordion-toggle.open .arrow{transform:rotate(180deg)}
.accordion-body{display:none;padding:16px 18px;overflow-x:auto}
.accordion-body.open{display:block}
.badge{background:rgba(255,255,255,.2);padding:2px 10px;border-radius:10px;font-size:12px;margin-left:8px}
.bl-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}
.bl-kpi{background:#fff;border-radius:var(--r);box-shadow:var(--sh);padding:18px 20px;text-align:center}
.bl-kpi .bl-val{font-size:28px;font-weight:700;margin:4px 0 2px}
.bl-kpi .bl-lbl{font-size:10px;color:var(--g4);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.bl-kpi .bl-sub{font-size:11px;color:var(--g6)}
.bl-kpi.green .bl-val{color:var(--gn)}
.bl-kpi.red .bl-val{color:var(--rd)}
.bl-kpi.blue .bl-val{color:var(--b9)}
.prop-bar{display:flex;height:28px;border-radius:6px;overflow:hidden;margin:8px 0}
.prop-bar div{display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;color:#fff;min-width:40px}
.prop-legend{display:flex;gap:16px;margin-top:6px;font-size:11px;color:var(--g6)}
.prop-legend span{display:inline-flex;align-items:center;gap:4px}
.prop-legend .dot{width:10px;height:10px;border-radius:2px;display:inline-block}
.bl-detail{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.bl-metric{padding:10px 0;border-bottom:1px solid var(--g2);display:flex;justify-content:space-between;align-items:center}
.bl-metric:last-child{border-bottom:none}
.bl-metric .lbl{font-size:12px;color:var(--g6)}
.bl-metric .val{font-size:14px;font-weight:600;color:var(--g8)}
.bl-split{display:flex;gap:6px;align-items:center;margin-top:10px}
.bl-split-bar{height:8px;border-radius:4px;flex:1}
.bl-split-label{font-size:10px;color:var(--g6);white-space:nowrap}
/* ── Évolution ── */
.annee-sel-bar{display:flex;align-items:center;gap:12px;margin-bottom:20px;background:#fff;padding:12px 18px;border-radius:var(--r);box-shadow:var(--sh)}
.annee-sel-bar label{font-size:13px;font-weight:600;color:var(--b9)}
.annee-sel-bar select{padding:6px 12px;border:1px solid var(--g2);border-radius:4px;font-family:inherit;font-size:13px;font-weight:600;color:var(--g8);cursor:pointer}
.gauge-wrap{background:#fff;border-radius:var(--r);box-shadow:var(--sh);padding:24px 28px;margin-bottom:24px}
.gauge-title{font-size:14px;font-weight:700;color:var(--b9);margin-bottom:16px}
.gauge-track{height:32px;background:var(--g2);border-radius:16px;overflow:hidden;position:relative}
.gauge-fill{height:100%;border-radius:16px;transition:width .8s ease;display:flex;align-items:center;justify-content:flex-end;padding-right:12px;min-width:48px}
.gauge-pct{font-size:13px;font-weight:700;color:#fff}
.gauge-labels{display:flex;justify-content:space-between;margin-top:8px;font-size:11px;color:var(--g6)}
.gauge-meta{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-top:20px}
.gauge-kpi{text-align:center;padding:14px;background:var(--g0);border-radius:8px}
.gauge-kpi .gk-lbl{font-size:10px;color:var(--g4);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.gauge-kpi .gk-val{font-size:20px;font-weight:700;color:var(--b9);margin-top:4px}
.gauge-kpi .gk-sub{font-size:11px;color:var(--g6);margin-top:2px}
.chart-card{background:#fff;border-radius:var(--r);box-shadow:var(--sh);overflow:hidden;margin-bottom:20px}
.chart-head{background:var(--b9);color:#fff;padding:10px 18px;font-size:13px;font-weight:600}
.chart-body{padding:18px;position:relative}
.chart-body canvas{max-height:300px}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
.recap-table-wrap{background:#fff;border-radius:var(--r);box-shadow:var(--sh);overflow:hidden}
.recap-table-head{background:var(--b9);color:#fff;padding:10px 18px;font-size:13px;font-weight:600}
@media print{body{background:#fff}.cd,.accordion{break-inside:avoid;box-shadow:none;border:1px solid var(--g2)}.accordion-body{display:block !important}.hdr{print-color-adjust:exact;-webkit-print-color-adjust:exact}.tab-panel{display:block !important}}
.anc-btn{padding:6px 14px;border-radius:4px;border:1px solid var(--g2);background:transparent;color:var(--g6);font-family:inherit;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s}
.anc-btn:hover{border-color:var(--ac,var(--b6));color:var(--ac,var(--b6))}
.anc-btn.active{background:var(--ac,var(--b6));color:#fff;border-color:var(--ac,var(--b6))}
@media(max-width:900px){.g2,.g3,.chart-grid,.gauge-meta,.bl-kpis{grid-template-columns:1fr}.kr{grid-template-columns:repeat(2,1fr)}.tabs-nav{flex-wrap:wrap}}
</style></head><body>

<div class="hdr">
<div><h1>Rapport Global — COORDIN</h1><div class="sub">Tableau de bord mensuel interactif</div></div>
<div style="display:flex;flex-direction:column;gap:10px;align-items:flex-end">
  <div class="tabs-nav">
    <button class="tab-btn active" onclick="switchTab('mensuel',this)">&#x1F4CA; Vue mensuelle</button>
    <button class="tab-btn" onclick="switchTab('evolution',this)">&#x1F4C8; Evolution annuelle</button>
    <button class="tab-btn" onclick="switchTab('anciennete',this)">&#x1F551; Ancienneté CA</button>
  </div>
  <div class="sel" id="month-sel-wrap"><label>Mois</label><select id="sm"></select><label>Année</label><select id="sa"></select></div>
</div>
</div>

<!-- ═══════════ ONGLET MENSUEL ═══════════ -->
<div id="tab-mensuel" class="tab-panel active">
<div class="cnt">
<div class="kr" id="kpi"></div>

<div class="section">
<div class="section-title">Chiffre d'affaires</div>
<div class="cd"><div class="ch">CA réalisé vs Objectif vs N-1</div><div class="cb"><table id="tca"></table></div></div>
</div>

<div class="section">
<div class="section-title">Répartition du CA par type de deal</div>
<div class="g3" id="deal-grid">
<div class="cd"><div class="ch sub">Global</div><div class="cb"><table id="tdl-g"></table></div></div>
<div class="cd"><div class="ch sub">Littéralis</div><div class="cb"><table id="tdl-l"></table></div></div>
<div class="cd"><div class="ch sub">GEODP</div><div class="cb"><table id="tdl-d"></table></div></div>
</div>
</div>

<div class="section">
<div class="section-title">Charge équipe</div>
<div class="cd" style="margin-bottom:16px"><div class="ch">Synthèse globale</div><div class="cb"><table id="tch"></table></div></div>
<div class="g2">
<div class="cd"><div class="ch sub">Détail Littéralis</div><div class="cb"><table id="tch-l"></table></div></div>
<div class="cd"><div class="ch sub">Détail GEODP</div><div class="cb"><table id="tch-d"></table></div></div>
</div>
</div>

<div class="section">
<div class="section-title">Backlog projets</div>
<div class="bl-kpis" id="bl-kpis"></div>
<div class="cd" style="margin:16px 0"><div class="cb" id="bl-bar"></div></div>
<div class="g2">
<div class="cd"><div class="ch sub">Littéralis</div><div class="cb" id="bl-litt"></div></div>
<div class="cd"><div class="ch sub">GEODP</div><div class="cb" id="bl-geodp"></div></div>
</div>
</div>

<div class="section">
<div class="accordion">
<div class="accordion-toggle" id="cmd-toggle" onclick="toggleCmd()">
<span>Commandes reçues <span class="badge" id="cmd-count">0</span></span>
<span class="arrow">&#9660;</span>
</div>
<div class="accordion-body" id="cmd-body"><table id="tcm"></table></div>
</div>
</div>
</div>
</div>

<!-- ═══════════ ONGLET ÉVOLUTION ═══════════ -->
<div id="tab-evolution" class="tab-panel">
<div class="cnt">

<div class="annee-sel-bar">
  <label>Année analysée :</label>
  <select id="ea"></select>
  <span style="font-size:12px;color:var(--g4);margin-left:auto">Objectif annuel global&nbsp;: <b style="color:var(--b9)">1 316 750 €</b></span>
</div>

<div class="gauge-wrap">
  <div class="gauge-title">Avancement vers l'objectif annuel — CA cumulé</div>
  <div class="gauge-track"><div class="gauge-fill" id="gauge-fill" style="width:0%;background:var(--b6)"><span class="gauge-pct" id="gauge-pct">0%</span></div></div>
  <div class="gauge-labels"><span>0 €</span><span>1 316 750 €</span></div>
  <div class="gauge-meta" id="gauge-meta"></div>
</div>

<div class="chart-card">
  <div class="chart-head">CA mensuel réalisé vs objectif mensuel vs N-1</div>
  <div class="chart-body"><canvas id="chart-ca-mensuel"></canvas></div>
</div>

<div class="chart-card">
  <div class="chart-head">CA cumulé vs objectif annuel cumulé</div>
  <div class="chart-body"><canvas id="chart-ca-cumul"></canvas></div>
</div>

<div class="chart-grid">
  <div class="chart-card" style="margin-bottom:0">
    <div class="chart-head">Heures mensuelles (Productif + Support + Interne)</div>
    <div class="chart-body"><canvas id="chart-heures"></canvas></div>
  </div>
  <div class="chart-card" style="margin-bottom:0">
    <div class="chart-head">Rework mensuel (heures)</div>
    <div class="chart-body"><canvas id="chart-rework"></canvas></div>
  </div>
</div>

<div class="recap-table-wrap" style="margin-top:20px">
  <div class="recap-table-head">Récapitulatif mensuel de l'année</div>
  <div style="overflow-x:auto;padding:4px"><table id="recap-table"></table></div>
</div>

</div>
</div>

<!-- ═══════════ ONGLET ANCIENNETÉ CA ═══════════ -->
<div id="tab-anciennete" class="tab-panel">
<div class="cnt">

<div class="annee-sel-bar">
  <label>Mois / Année :</label>
  <select id="anc-sm"></select>
  <select id="anc-sa"></select>
  <div style="margin-left:12px;display:flex;gap:6px" id="anc-sol-btns">
    <button class="anc-btn active" onclick="ancFilter('all',this)">Tout</button>
    <button class="anc-btn" onclick="ancFilter('LITTERALIS',this)" style="--ac:#7c3aed">Littéralis</button>
    <button class="anc-btn" onclick="ancFilter('GEODP',this)" style="--ac:#059669">GEODP</button>
  </div>
</div>

<div class="kr" id="anc-kpis"></div>

<div class="chart-card">
  <div class="chart-head">Répartition du CA déclaré par ancienneté de commande</div>
  <div class="chart-body"><canvas id="chart-anc-bar"></canvas></div>
</div>

<div class="chart-card">
  <div class="chart-head">Chaque projet : CA vs ancienneté</div>
  <div class="chart-body"><canvas id="chart-anc-scatter"></canvas></div>
</div>

<div class="recap-table-wrap">
  <div class="recap-table-head">Synthèse par tranche d'ancienneté</div>
  <div style="overflow-x:auto;padding:4px"><table id="anc-synth"></table></div>
</div>

<div class="cd" style="margin-top:20px">
  <div class="ch">Détail des projets avec CA déclaré</div>
  <div class="cb"><table id="anc-detail"></table></div>
</div>

</div>
</div>

<script>
var D=__DATA__,P=__PERIODS__,CUR=__CUR__,ML=__ML__;
var OBJ_ANNUEL=__OBJ_ANNUEL__;
var G=function(o,k){return o[k]||o[k.replace("\u20ac","\u20ac")]||0};
function fmt(v,t){if(v===""||v===null||v===undefined||isNaN(v))return"—";var n=+v;if(t==="eur")return n.toLocaleString("fr-FR",{minimumFractionDigits:0,maximumFractionDigits:0})+" €";if(t==="eur2")return n.toLocaleString("fr-FR",{minimumFractionDigits:2,maximumFractionDigits:2})+" €";if(t==="h")return n.toFixed(1)+" h";if(t==="num")return n.toLocaleString("fr-FR");return""+v}
function fl(k,m,a){return(D[k]||[]).filter(function(r){return r.Mois===m&&""+Math.round(+(r["Année"]||r.Annee||0))===""+a})}
function fla(k,a){return(D[k]||[]).filter(function(r){return""+Math.round(+(r["Année"]||r.Annee||0))===""+a})}
function pc(v,r){return r>0?v/r:null}
function ev(v,r){return r>0?(v-r)/r:null}
function tg(v,t){if(v===null)return"—";var n=+v,s;if(t==="pct"){s=(n*100).toFixed(1)+"%";return n>=1?'<span class="tg tg-g">'+s+"</span>":n<.8?'<span class="tg tg-r">'+s+"</span>":s}s=(n>=0?"+":"")+(n*100).toFixed(1)+"%";return n>=0?'<span class="tg tg-g">'+s+"</span>":'<span class="tg tg-r">'+s+"</span>"}

function dealTable(deals,sol){var dd=deals.filter(function(d){return d.Solution===sol});if(!dd.length)return'<div class="es">Aucune donnée</div>';var tot=dd.reduce(function(s,d){return s+Number(G(d,"CA (\u20ac)"))},0);return'<tr><th>Type de deal</th><th class="r">CA</th><th class="r">%</th></tr>'+dd.map(function(d){var v=Number(G(d,"CA (\u20ac)"));return'<tr><td>'+(d["Type de deal"]||"")+'</td><td class="r">'+fmt(v,"eur")+'</td><td class="r">'+(tot>0?(v/tot*100).toFixed(0)+"%":"")+'</td></tr>'}).join("")+'<tr class="tr-total"><td>Total</td><td class="r">'+fmt(tot,"eur")+'</td><td class="r">100%</td></tr>'}
function chDetail(chr,pfx){var rows=[["Productif","Productif "+pfx],["dont Rework","Rework "+pfx],["Support","Support "+pfx]];return'<tr><th>Indicateur</th><th class="r">Heures</th></tr>'+rows.map(function(r){return'<tr><td>'+r[0]+'</td><td class="r">'+fmt(chr[r[1]],"h")+'</td></tr>'}).join("")}
function toggleCmd(){var t=document.getElementById("cmd-toggle"),b=document.getElementById("cmd-body");var open=b.classList.toggle("open");t.classList.toggle("open",open)}

// ── Onglets ──
function switchTab(name,btn){
  document.querySelectorAll(".tab-panel").forEach(function(p){p.classList.remove("active")});
  document.querySelectorAll(".tab-btn").forEach(function(b){b.classList.remove("active")});
  document.getElementById("tab-"+name).classList.add("active");
  btn.classList.add("active");
  document.getElementById("month-sel-wrap").style.display=name==="mensuel"?"flex":"none";
  if(name==="evolution") renderEvolution();
  if(name==="anciennete") renderAnciennete();
}

// ── Render mensuel ──
function render(){
var m=document.getElementById("sm").value,a=document.getElementById("sa").value;
var ca=fl("ca",m,a),cr=ca[0]||{},re=+G(cr,"CA réalisé (\u20ac)"),ob=+G(cr,"Objectif global (\u20ac)"),n1=+G(cr,"CA N-1 global (\u20ac)"),at=pc(re,ob),ev1=ev(re,n1);
var ch=fl("charge",m,a),chr=ch[0]||{};
var ht=Number(chr["Productif Total"]||0)+Number(chr["Support Total"]||0)+Number(chr["Interne"]||0);
var cmd=fl("commandes",m,a);
document.getElementById("kpi").innerHTML='<div class="kp"><div class="kl">CA réalisé</div><div class="kv">'+fmt(re,"eur")+'</div></div><div class="kp"><div class="kl">Objectif</div><div class="kv">'+fmt(ob,"eur")+'</div></div><div class="kp"><div class="kl">% Atteinte</div><div class="kv '+(at!==null&&at>=1?"gn":at!==null&&at<.8?"rd":"or")+'">'+(at!==null?(at*100).toFixed(1)+"%":"—")+'</div></div><div class="kp"><div class="kl">Évol. N-1</div><div class="kv '+(ev1!==null&&ev1>=0?"gn":"rd")+'">'+(ev1!==null?(ev1>=0?"+":"")+(ev1*100).toFixed(1)+"%":"—")+'</div></div><div class="kp"><div class="kl">Heures loguées</div><div class="kv">'+fmt(ht,"h")+'</div></div><div class="kp"><div class="kl">Commandes</div><div class="kv">'+cmd.length+'</div></div>';
var cL=+G(cr,"CA LITTERALIS (\u20ac)"),cG=+G(cr,"CA GEODP (\u20ac)"),oL=+G(cr,"Objectif LITTERALIS (\u20ac)"),oG=+G(cr,"Objectif GEODP (\u20ac)"),nL=+G(cr,"CA N-1 LITTERALIS (\u20ac)"),nG=+G(cr,"CA N-1 GEODP (\u20ac)");
var ls=[["GLOBAL",re,ob,n1],["LITTERALIS",cL,oL,nL],["GEODP",cG,oG,nG]];
document.getElementById("tca").innerHTML='<tr><th>Périmètre</th><th class="r">Réalisé</th><th class="r">Objectif</th><th class="r">% Att.</th><th class="r">N-1</th><th class="r">Évol.</th><th class="r">Écart</th></tr>'+ls.map(function(l){return'<tr><td><b>'+l[0]+'</b></td><td class="r">'+fmt(l[1],"eur")+'</td><td class="r">'+fmt(l[2],"eur")+'</td><td class="r">'+tg(pc(l[1],l[2]),"pct")+'</td><td class="r">'+fmt(l[3],"eur")+'</td><td class="r">'+tg(ev(l[1],l[3]),"evol")+'</td><td class="r">'+fmt(l[1]-l[2],"eur")+'</td></tr>'}).join("");
var dl=fl("ca_deal",m,a);
document.getElementById("tdl-g").innerHTML=dealTable(dl,"GLOBAL");
document.getElementById("tdl-l").innerHTML=dealTable(dl,"LITTERALIS");
document.getElementById("tdl-d").innerHTML=dealTable(dl,"GEODP");
var chl=[["Productif",chr["Productif Litt"],chr["Productif GEODP"]],["dont Rework",chr["Rework Litt"],chr["Rework GEODP"]],["Support",chr["Support Litt"],chr["Support GEODP"]],["Interne",chr["Interne"]||0,""]];
document.getElementById("tch").innerHTML='<tr><th>Indicateur</th><th class="r">Litt.</th><th class="r">GEODP</th><th class="r">Total</th></tr>'+chl.map(function(l){var a2=+(l[1]||0),b2=+(l[2]||0);var tot=l[0]==="Interne"?a2:l[0]==="Support"?Number(chr["Support Total"]||a2+b2):a2+b2;return'<tr><td><b>'+l[0]+'</b></td><td class="r">'+fmt(a2,"h")+'</td><td class="r">'+(l[0]==="Interne"?"":fmt(b2,"h"))+'</td><td class="r"><b>'+fmt(tot,"h")+'</b></td></tr>'}).join("")+'<tr class="tr-total"><td>TOTAL</td><td class="r"></td><td class="r"></td><td class="r">'+fmt(ht,"h")+'</td></tr>';
document.getElementById("tch-l").innerHTML=chDetail(chr,"Litt");
document.getElementById("tch-d").innerHTML=chDetail(chr,"GEODP");
var bl=fl("backlog",m,a),b=bl[0]||{},hb=Object.keys(b).length>2;
if(hb){
var bkT=+b.total_backlog_TOTAL||0,bkB=+b.total_bloque_TOTAL||0,bkM=+b.total_mobilisable_TOTAL||0;
var bkPctB=bkT>0?(bkB/bkT*100).toFixed(0):0,bkPctM=bkT>0?(bkM/bkT*100).toFixed(0):0;
document.getElementById("bl-kpis").innerHTML='<div class="bl-kpi blue"><div class="bl-lbl">Backlog total</div><div class="bl-val">'+fmt(bkT,"eur")+'</div><div class="bl-sub">'+(+b.nb_epics_TOTAL||0)+' epics ouverts</div></div><div class="bl-kpi red"><div class="bl-lbl">Bloqué</div><div class="bl-val">'+fmt(bkB,"eur")+'</div><div class="bl-sub">'+bkPctB+'% du backlog — '+(+b.nb_bloques_TOTAL||0)+' projets</div></div><div class="bl-kpi green"><div class="bl-lbl">Mobilisable</div><div class="bl-val">'+fmt(bkM,"eur")+'</div><div class="bl-sub">'+bkPctM+'% du backlog</div></div>';
var pB=bkT>0?bkB/bkT*100:0,pM=bkT>0?bkM/bkT*100:0;
document.getElementById("bl-bar").innerHTML='<div style="font-size:12px;font-weight:600;color:var(--b9);margin-bottom:6px">Répartition du backlog</div><div class="prop-bar"><div style="width:'+pB+'%;background:#C55A11">'+pB.toFixed(0)+'%</div><div style="width:'+pM+'%;background:#375623">'+pM.toFixed(0)+'%</div></div><div class="prop-legend"><span><span class="dot" style="background:#C55A11"></span>Bloqué</span><span><span class="dot" style="background:#375623"></span>Mobilisable</span></div>';
function blSolCard(sfx){var tot=+b["total_backlog_"+sfx]||0,blq=+b["total_bloque_"+sfx]||0,mob=+b["total_mobilisable_"+sfx]||0,ne=+b["nb_epics_"+sfx]||0,np=+b["nb_projets_"+sfx]||0,nr=+b["nb_rework_"+sfx]||0,nb2=+b["nb_bloques_"+sfx]||0;var pctOfTotal=bkT>0?(tot/bkT*100).toFixed(0):"0";return'<div style="font-size:11px;color:var(--g4);margin-bottom:8px">'+pctOfTotal+'% du backlog total</div><div class="bl-metric"><span class="lbl">Backlog</span><span class="val">'+fmt(tot,"eur")+'</span></div><div class="bl-metric"><span class="lbl">Bloqué</span><span class="val" style="color:var(--rd)">'+fmt(blq,"eur")+'</span></div><div class="bl-metric"><span class="lbl">Mobilisable</span><span class="val" style="color:var(--gn)">'+fmt(mob,"eur")+'</span></div>'+(function(){var pMob=tot>0?(mob/tot*100):0,pBlq=tot>0?(blq/tot*100):0;return'<div style="margin:10px 0;font-size:11px;font-weight:600;color:var(--b9)">Répartition financière</div><div class="prop-bar" style="height:22px;margin:4px 0">'+(pMob>5?'<div style="width:'+pMob+'%;background:#375623;font-size:10px">'+pMob.toFixed(0)+'%</div>':pMob>0?'<div style="width:'+pMob+'%;background:#375623"></div>':'')+(pBlq>5?'<div style="width:'+pBlq+'%;background:#C55A11;font-size:10px">'+pBlq.toFixed(0)+'%</div>':pBlq>0?'<div style="width:'+pBlq+'%;background:#C55A11"></div>':'')+'</div><div class="prop-legend"><span><span class="dot" style="background:#375623"></span>Mobilisable '+pMob.toFixed(0)+'%</span><span><span class="dot" style="background:#C55A11"></span>Bloqué '+pBlq.toFixed(0)+'%</span></div>'})()+'<div style="border-top:1px solid var(--g2);margin:10px 0;padding-top:10px"><div class="bl-metric"><span class="lbl">Epics ouverts</span><span class="val">'+ne+'</span></div><div class="bl-metric"><span class="lbl">Projets</span><span class="val">'+np+'</span></div><div class="bl-metric" style="padding-left:16px"><span class="lbl" style="font-style:italic">dont bloqués</span><span class="val" style="color:var(--rd)">'+nb2+'</span></div><div class="bl-metric"><span class="lbl">Rework</span><span class="val" style="color:var(--or)">'+nr+'</span></div></div><div style="margin-top:10px;font-size:11px;font-weight:600;color:var(--b9)">Répartition des epics</div>'+(function(){var nFree=Math.max(0,np-nb2),pFree=ne>0?(nFree/ne*100):0,pRwk=ne>0?(nr/ne*100):0,pNb=ne>0?(nb2/ne*100):0;return'<div class="prop-bar" style="height:22px;margin:6px 0 4px">'+(pFree>5?'<div style="width:'+pFree+'%;background:#2E75B6;font-size:10px">'+nFree+'</div>':pFree>0?'<div style="width:'+pFree+'%;background:#2E75B6"></div>':'')+(pRwk>5?'<div style="width:'+pRwk+'%;background:#C55A11;font-size:10px">'+nr+'</div>':pRwk>0?'<div style="width:'+pRwk+'%;background:#C55A11"></div>':'')+(pNb>5?'<div style="width:'+pNb+'%;background:#C00000;font-size:10px">'+nb2+'</div>':pNb>0?'<div style="width:'+pNb+'%;background:#C00000"></div>':'')+'</div><div class="prop-legend"><span><span class="dot" style="background:#2E75B6"></span>En cours '+nFree+'</span><span><span class="dot" style="background:#C55A11"></span>Rework '+nr+'</span><span><span class="dot" style="background:#C00000"></span>Bloqués '+nb2+'</span></div>'})()}
document.getElementById("bl-litt").innerHTML=blSolCard("LITT");
document.getElementById("bl-geodp").innerHTML=blSolCard("GEODP");
}else{document.getElementById("bl-kpis").innerHTML='';document.getElementById("bl-bar").innerHTML='<div class="es">Aucune donnée backlog</div>';document.getElementById("bl-litt").innerHTML='';document.getElementById("bl-geodp").innerHTML='';}
document.getElementById("cmd-count").textContent=cmd.length;
var cc=[["Date de création"],["Clé Epic"],["Nom / Résumé"],["Client"],["Solution"],["Type de deal"],["Statut"],["Montant total prestations (\u20ac)"],["CA reconnu ce mois (\u20ac)"]];
var hh=["Date","Clé","Nom","Client","Sol.","Type","Statut","Montant","CA rec."];
document.getElementById("tcm").innerHTML=cmd.length?'<tr>'+hh.map(function(h,i){return'<th'+(i>=7?' class="r"':'')+'>'+h+'</th>'}).join("")+'</tr>'+cmd.map(function(r){return'<tr>'+cc.map(function(c,i){var v=r[c[0]]||"";return'<td'+(i>=7?' class="r"':'')+'>'+(i>=7?fmt(v,"eur2"):i===2?(""+v).substring(0,55):""+v)+'</td>'}).join("")+'</tr>'}).join(""):'<div class="es">Aucune commande</div>';
}

// ── Render évolution ──
var _charts={};
function destroyChart(id){if(_charts[id]){_charts[id].destroy();delete _charts[id];}}
var _evoRendered=false;

function renderEvolution(){
var a=document.getElementById("ea").value;
if(!a)return;
var caData=fla("ca",a).sort(function(x,y){return ML.indexOf(x.Mois)-ML.indexOf(y.Mois)});
var chData=fla("charge",a).sort(function(x,y){return ML.indexOf(x.Mois)-ML.indexOf(y.Mois)});
var labels=caData.map(function(r){return r.Mois.substring(0,4)});
var realises=caData.map(function(r){return +G(r,"CA réalisé (\u20ac)")||0});
var objectifs=caData.map(function(r){return +G(r,"Objectif global (\u20ac)")||0});
var n1vals=caData.map(function(r){return +G(r,"CA N-1 global (\u20ac)")||0});
var objMensuel=OBJ_ANNUEL/12;
var cumRealise=0,cumulRealises=[];
for(var i=0;i<realises.length;i++){cumRealise+=realises[i];cumulRealises.push(Math.round(cumRealise));}
var cumulObjectifs=[],cumulLabels=ML.map(function(m){return m.substring(0,4)});
for(var j=1;j<=12;j++){cumulObjectifs.push(Math.round(objMensuel*j));}
var hProd=chData.map(function(r){return Number(r["Productif Total"]||0)});
var hSupp=chData.map(function(r){return Number(r["Support Total"]||0)});
var hInt=chData.map(function(r){return Number(r["Interne"]||0)});
var hRw=chData.map(function(r){return Number(r["Rework Total"]||0)});
var labelsH=chData.map(function(r){return r.Mois.substring(0,4)});
var totalRealise=realises.reduce(function(s,v){return s+v},0);
var pct=Math.min(100,(totalRealise/OBJ_ANNUEL)*100);
var fillColor=pct>=100?"#375623":pct>=80?"#2E75B6":pct>=60?"#C55A11":"#C00000";
var fill=document.getElementById("gauge-fill");
fill.style.width=pct.toFixed(1)+"%";fill.style.background=fillColor;
document.getElementById("gauge-pct").textContent=pct.toFixed(1)+"%";
var restant=Math.max(0,OBJ_ANNUEL-totalRealise),moisRestants=12-realises.length;
var moy=realises.length>0?Math.round(totalRealise/realises.length):0;
var proj=moy*12;
document.getElementById("gauge-meta").innerHTML=
'<div class="gauge-kpi"><div class="gk-lbl">CA cumulé</div><div class="gk-val" style="color:'+fillColor+'">'+fmt(totalRealise,"eur")+'</div><div class="gk-sub">sur '+fmt(OBJ_ANNUEL,"eur")+'</div></div>'+
'<div class="gauge-kpi"><div class="gk-lbl">Restant</div><div class="gk-val" style="color:var(--rd)">'+fmt(restant,"eur")+'</div><div class="gk-sub">'+moisRestants+' mois restants</div></div>'+
'<div class="gauge-kpi"><div class="gk-lbl">Moy. mensuelle</div><div class="gk-val" style="color:var(--b9)">'+(realises.length>0?fmt(moy,"eur"):"—")+'</div><div class="gk-sub">CA moyen réalisé</div></div>'+
'<div class="gauge-kpi"><div class="gk-lbl">Projection fin d\'année</div><div class="gk-val" style="color:'+(proj>=OBJ_ANNUEL?"var(--gn)":"var(--or)")+'">'+fmt(proj,"eur")+'</div><div class="gk-sub">'+(proj>=OBJ_ANNUEL?"\u2705 Objectif atteignable":"\u26a0\ufe0f Sous l\'objectif")+'</div></div>';
var toFixed2=function(v){return v.toLocaleString("fr-FR")+" \u20ac"};
destroyChart("ca-m");
var ctx1=document.getElementById("chart-ca-mensuel").getContext("2d");
_charts["ca-m"]=new Chart(ctx1,{type:"bar",data:{labels:labels,datasets:[
{label:"CA réalisé",data:realises,backgroundColor:"rgba(46,117,182,0.85)",borderRadius:4,order:2},
{label:"Objectif mensuel",data:objectifs.length?objectifs:new Array(labels.length).fill(Math.round(objMensuel)),borderColor:"#C55A11",backgroundColor:"transparent",borderWidth:2,borderDash:[6,3],type:"line",pointRadius:4,pointBackgroundColor:"#C55A11",fill:false,order:1},
{label:"CA N-1",data:n1vals,borderColor:"rgba(149,179,215,0.9)",backgroundColor:"transparent",borderWidth:2,type:"line",pointRadius:3,pointBackgroundColor:"rgba(149,179,215,0.9)",fill:false,order:0}
]},options:{responsive:true,plugins:{legend:{position:"top",labels:{font:{size:11}}},tooltip:{callbacks:{label:function(c){return c.dataset.label+": "+c.raw.toLocaleString("fr-FR")+" \u20ac"}}}},scales:{y:{ticks:{callback:function(v){return(v/1000).toFixed(0)+"k \u20ac"}},grid:{color:"rgba(0,0,0,.05)"}},x:{grid:{display:false}}}}});
destroyChart("ca-c");
var ctx2=document.getElementById("chart-ca-cumul").getContext("2d");
var cumulRealisesPlot=cumulRealises.concat(new Array(12-cumulRealises.length).fill(null));
_charts["ca-c"]=new Chart(ctx2,{type:"line",data:{labels:cumulLabels,datasets:[
{label:"CA cumulé réalisé",data:cumulRealisesPlot,borderColor:"#2E75B6",backgroundColor:"rgba(46,117,182,0.1)",borderWidth:2.5,pointRadius:5,pointBackgroundColor:"#2E75B6",fill:true,tension:0.3},
{label:"Objectif cumulé (linéaire)",data:cumulObjectifs,borderColor:"#C55A11",backgroundColor:"transparent",borderWidth:2,borderDash:[6,3],pointRadius:3,pointBackgroundColor:"#C55A11",fill:false,tension:0},
{label:"Objectif annuel",data:new Array(12).fill(OBJ_ANNUEL),borderColor:"rgba(192,0,0,0.35)",backgroundColor:"transparent",borderWidth:1.5,borderDash:[2,5],pointRadius:0,fill:false}
]},options:{responsive:true,plugins:{legend:{position:"top",labels:{font:{size:11}}},tooltip:{callbacks:{label:function(c){return c.raw!==null?c.dataset.label+": "+c.raw.toLocaleString("fr-FR")+" \u20ac":""}}}},scales:{y:{ticks:{callback:function(v){return(v/1000).toFixed(0)+"k \u20ac"}},grid:{color:"rgba(0,0,0,.05)"}},x:{grid:{display:false}}}}});
destroyChart("h");
var ctx3=document.getElementById("chart-heures").getContext("2d");
_charts["h"]=new Chart(ctx3,{type:"bar",data:{labels:labelsH,datasets:[
{label:"Productif",data:hProd,backgroundColor:"rgba(46,117,182,0.85)",borderRadius:3,stack:"h"},
{label:"Support",data:hSupp,backgroundColor:"rgba(46,95,158,0.7)",borderRadius:3,stack:"h"},
{label:"Interne",data:hInt,backgroundColor:"rgba(149,179,215,0.7)",borderRadius:3,stack:"h"}
]},options:{responsive:true,plugins:{legend:{position:"top",labels:{font:{size:11}}},tooltip:{callbacks:{label:function(c){return c.dataset.label+": "+c.raw.toFixed(1)+" h"}}}},scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,ticks:{callback:function(v){return v+" h"}},grid:{color:"rgba(0,0,0,.05)"}}}}});
destroyChart("rw");
var ctx4=document.getElementById("chart-rework").getContext("2d");
_charts["rw"]=new Chart(ctx4,{type:"bar",data:{labels:labelsH,datasets:[
{label:"Rework (h)",data:hRw,backgroundColor:"rgba(197,90,17,0.8)",borderRadius:4}
]},options:{responsive:true,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){return "Rework: "+c.raw.toFixed(1)+" h"}}}},scales:{x:{grid:{display:false}},y:{ticks:{callback:function(v){return v+" h"}},grid:{color:"rgba(0,0,0,.05)"}}}}});
var allMois=caData.map(function(r,i){var re2=realises[i]||0,ob2=objectifs[i]||Math.round(objMensuel),n12=n1vals[i]||0,ch2=chData[i]||{},ht2=Number(ch2["Productif Total"]||0)+Number(ch2["Support Total"]||0)+Number(ch2["Interne"]||0);return{mois:r.Mois,re:re2,ob:ob2,n1:n12,ht:ht2,rw:Number(ch2["Rework Total"]||0)};});
var totRe=allMois.reduce(function(s,r){return s+r.re},0),totHt=allMois.reduce(function(s,r){return s+r.ht},0),totRw=allMois.reduce(function(s,r){return s+r.rw},0);
document.getElementById("recap-table").innerHTML=
'<tr><th>Mois</th><th class="r">CA réalisé</th><th class="r">Obj. mensuel</th><th class="r">% Att.</th><th class="r">Évol. N-1</th><th class="r">H loguées</th><th class="r">Rework</th></tr>'+
allMois.map(function(r){return'<tr><td><b>'+r.mois+'</b></td><td class="r">'+fmt(r.re,"eur")+'</td><td class="r">'+fmt(r.ob,"eur")+'</td><td class="r">'+tg(pc(r.re,r.ob),"pct")+'</td><td class="r">'+tg(ev(r.re,r.n1),"evol")+'</td><td class="r">'+fmt(r.ht,"h")+'</td><td class="r" style="color:var(--or)">'+fmt(r.rw,"h")+'</td></tr>'}).join("")+
(allMois.length?'<tr class="tr-total"><td>TOTAL</td><td class="r">'+fmt(totRe,"eur")+'</td><td class="r">'+fmt(OBJ_ANNUEL,"eur")+'</td><td class="r">'+tg(pc(totRe,OBJ_ANNUEL),"pct")+'</td><td class="r">—</td><td class="r">'+fmt(totHt,"h")+'</td><td class="r" style="color:var(--or)">'+fmt(totRw,"h")+'</td></tr>':"");
}

(function(){
var sm=document.getElementById("sm"),sa=document.getElementById("sa"),an=[];
P.forEach(function(p){if(an.indexOf(p.annee)===-1)an.push(p.annee)});an.sort();
ML.forEach(function(m){var o=document.createElement("option");o.value=m;o.textContent=m;if(m===CUR.mois)o.selected=true;sm.appendChild(o)});
an.forEach(function(a){var o=document.createElement("option");o.value=a;o.textContent=a;if(a===CUR.annee)o.selected=true;sa.appendChild(o)});
sm.addEventListener("change",render);sa.addEventListener("change",render);
var ea=document.getElementById("ea");
an.forEach(function(a){var o=document.createElement("option");o.value=a;o.textContent=a;if(a===CUR.annee)o.selected=true;ea.appendChild(o)});
ea.addEventListener("change",renderEvolution);
// Ancienneté selectors
var ancSm=document.getElementById("anc-sm"),ancSa=document.getElementById("anc-sa");
ML.forEach(function(m){var o=document.createElement("option");o.value=m;o.textContent=m;if(m===CUR.mois)o.selected=true;ancSm.appendChild(o)});
an.forEach(function(a){var o=document.createElement("option");o.value=a;o.textContent=a;if(a===CUR.annee)o.selected=true;ancSa.appendChild(o)});
ancSm.addEventListener("change",renderAnciennete);ancSa.addEventListener("change",renderAnciennete);
render();
})();

// ── Ancienneté CA ──
var _ancSolFilter="all";
function ancFilter(sol,btn){_ancSolFilter=sol;document.querySelectorAll(".anc-btn").forEach(function(b){b.classList.remove("active")});btn.classList.add("active");renderAnciennete()}
var BUCKET_ORDER=["M+0","M+1","M+2","M+3","M+4\u21926","M+7\u219212","M+13+"];
var BUCKET_COLORS=["#10b981","#3b82f6","#6366f1","#8b5cf6","#f59e0b","#f97316","#ef4444"];
function ancBucket(v){if(v<=0)return"M+0";if(v<=1)return"M+1";if(v<=2)return"M+2";if(v<=3)return"M+3";if(v<=6)return"M+4\u21926";if(v<=12)return"M+7\u219212";return"M+13+"}

function renderAnciennete(){
var m=document.getElementById("anc-sm").value,a=document.getElementById("anc-sa").value;
var raw=fl("anciennete",m,a);
if(_ancSolFilter!=="all")raw=raw.filter(function(r){return r.Solution===_ancSolFilter});
if(!raw.length){document.getElementById("anc-kpis").innerHTML='<div class="es" style="grid-column:1/-1">Aucune donnée d\'ancienneté pour ce mois.<br>Assurez-vous que l\'export contient les colonnes "Date commande" et "Ancienneté (mois)".</div>';
["chart-anc-bar","chart-anc-scatter"].forEach(function(id){destroyChart(id)});
document.getElementById("anc-synth").innerHTML="";document.getElementById("anc-detail").innerHTML="";return}

var totalCA=raw.reduce(function(s,r){return s+(+r.CA||0)},0);
var avgAnc=totalCA>0?raw.reduce(function(s,r){return s+(+r.CA||0)*(+r["Anciennet\u00e9"]||0)},0)/totalCA:0;
var caRecent=raw.filter(function(r){return(+r["Anciennet\u00e9"]||0)<=3}).reduce(function(s,r){return s+(+r.CA||0)},0);
var caOld=raw.filter(function(r){return(+r["Anciennet\u00e9"]||0)>12}).reduce(function(s,r){return s+(+r.CA||0)},0);
var pctRecent=totalCA>0?(caRecent/totalCA*100).toFixed(0):"0";
var pctOld=totalCA>0?(caOld/totalCA*100).toFixed(0):"0";

document.getElementById("anc-kpis").innerHTML=
'<div class="kp"><div class="kl">CA total déclaré</div><div class="kv">'+fmt(totalCA,"eur")+'</div></div>'+
'<div class="kp"><div class="kl">Projets avec CA</div><div class="kv">'+raw.length+'</div></div>'+
'<div class="kp"><div class="kl">Âge moyen pondéré</div><div class="kv">'+avgAnc.toFixed(1)+' mois</div></div>'+
'<div class="kp"><div class="kl">CA \u2264 3 mois</div><div class="kv gn">'+pctRecent+'%</div></div>'+
'<div class="kp"><div class="kl">CA > 12 mois</div><div class="kv rd">'+pctOld+'%</div></div>';

// Buckets
var bk={};BUCKET_ORDER.forEach(function(b){bk[b]={ca:0,count:0,projets:[]}});
raw.forEach(function(r){var b=ancBucket(+r["Anciennet\u00e9"]||0);bk[b].ca+=(+r.CA||0);bk[b].count++;bk[b].projets.push(r)});
var barLabels=BUCKET_ORDER,barCA=BUCKET_ORDER.map(function(b){return Math.round(bk[b].ca)});

// Bar chart
destroyChart("anc-bar");
var ctx=document.getElementById("chart-anc-bar").getContext("2d");
_charts["anc-bar"]=new Chart(ctx,{type:"bar",data:{labels:barLabels,datasets:[{
label:"CA déclaré",data:barCA,backgroundColor:BUCKET_COLORS,borderRadius:5,maxBarThickness:55
}]},options:{responsive:true,plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){var b=BUCKET_ORDER[c.dataIndex],d=bk[b];return[fmt(d.ca,"eur"),d.count+" projet"+(d.count>1?"s":""),totalCA>0?(d.ca/totalCA*100).toFixed(1)+"% du CA":""]}}}},scales:{y:{ticks:{callback:function(v){return(v/1000).toFixed(0)+"k \u20ac"}},grid:{color:"rgba(0,0,0,.05)"}},x:{grid:{display:false}}}}});

// Scatter chart
destroyChart("anc-scatter");
var scatterData=raw.map(function(r){return{x:+r["Anciennet\u00e9"]||0,y:+r.CA||0,nom:r.Nom,epic:r.Epic,sol:r.Solution}});
var scLitt=scatterData.filter(function(d){return d.sol==="LITTERALIS"});
var scGeod=scatterData.filter(function(d){return d.sol!=="LITTERALIS"});
var ctx2=document.getElementById("chart-anc-scatter").getContext("2d");
_charts["anc-scatter"]=new Chart(ctx2,{type:"scatter",data:{datasets:[
{label:"GEODP",data:scGeod,backgroundColor:"rgba(16,185,129,0.7)",pointRadius:function(c){return Math.max(4,Math.min(18,Math.sqrt(c.raw.y)/4))},pointHoverRadius:10},
{label:"Littéralis",data:scLitt,backgroundColor:"rgba(124,58,237,0.7)",pointRadius:function(c){return Math.max(4,Math.min(18,Math.sqrt(c.raw.y)/4))},pointHoverRadius:10}
]},options:{responsive:true,plugins:{legend:{position:"top",labels:{font:{size:11}}},tooltip:{callbacks:{title:function(cs){var d=cs[0].raw;return d.nom},label:function(c){var d=c.raw;return[d.epic+" \u00b7 "+d.sol,"CA: "+fmt(d.y,"eur"),"Anciennet\u00e9: "+d.x+" mois"]}}}},scales:{x:{title:{display:true,text:"Anciennet\u00e9 (mois)",font:{size:11}},grid:{color:"rgba(0,0,0,.05)"}},y:{title:{display:true,text:"CA d\u00e9clar\u00e9 (\u20ac)",font:{size:11}},ticks:{callback:function(v){return(v/1000).toFixed(0)+"k \u20ac"}},grid:{color:"rgba(0,0,0,.05)"}}}}});

// Synth table
var cum=0;
document.getElementById("anc-synth").innerHTML='<tr><th>Tranche</th><th class="r">Projets</th><th class="r">CA déclaré</th><th class="r">% du total</th><th class="r">Cumulé</th></tr>'+
BUCKET_ORDER.map(function(b,i){var d=bk[b];cum+=d.ca;var pct=totalCA>0?(d.ca/totalCA*100).toFixed(1):"0";var cpct=totalCA>0?(cum/totalCA*100).toFixed(1):"0";
return'<tr><td><span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:'+BUCKET_COLORS[i]+';margin-right:6px;vertical-align:middle"></span><b>'+b+'</b></td><td class="r">'+d.count+'</td><td class="r">'+fmt(d.ca,"eur")+'</td><td class="r">'+pct+'%</td><td class="r">'+cpct+'%</td></tr>'}).join("")+
'<tr class="tr-total"><td>TOTAL</td><td class="r">'+raw.length+'</td><td class="r">'+fmt(totalCA,"eur")+'</td><td class="r">100%</td><td class="r"></td></tr>';

// Detail table
var sorted=raw.slice().sort(function(a,b){return(+b.CA||0)-(+a.CA||0)});
document.getElementById("anc-detail").innerHTML='<tr><th>Epic</th><th>Nom</th><th>Solution</th><th>Date commande</th><th class="r">Ancienneté</th><th class="r">CA déclaré</th></tr>'+
sorted.map(function(r){var anc=+r["Anciennet\u00e9"]||0;var ancColor=anc>12?"var(--rd)":anc>6?"var(--or)":"inherit";
return'<tr><td>'+r.Epic+'</td><td>'+r.Nom+'</td><td>'+r.Solution+'</td><td>'+r["Date commande"]+'</td><td class="r" style="font-weight:700;color:'+ancColor+'">'+anc+' mois</td><td class="r"><b>'+fmt(+r.CA,"eur")+'</b></td></tr>'}).join("")}

</script></body></html>"""

HTML = HTML.replace("__DATA__", json.dumps(dj, ensure_ascii=False, default=str))
HTML = HTML.replace("__PERIODS__", json.dumps([{"mois":m,"annee":a} for m,a in periods], ensure_ascii=False))
HTML = HTML.replace("__CUR__", json.dumps({"mois":mois_label,"annee":annee_val}, ensure_ascii=False))
HTML = HTML.replace("__ML__", json.dumps(MOIS_FR, ensure_ascii=False))
HTML = HTML.replace("__OBJ_ANNUEL__", str(OBJECTIF_ANNUEL))

with open(RAPPORT_HTML_FILE, "w", encoding="utf-8") as f:
    f.write(HTML)
print(f"\n✅ {RAPPORT_HTML_FILE} ({len(HTML)//1024} Ko)")
print(f"   Ouvrir dans Chrome/Edge — Ctrl+P pour PDF")