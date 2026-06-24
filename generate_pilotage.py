"""
generate_pilotage.py
====================
Génère le dashboard de pilotage annuel (pilotage.html) : vue TENDANCES
multi-mois, à partir de _historique.xlsx.

Tendances :
  1. CA mensuel (réalisé global / Littéralis / GEODP, + objectif, + PY)
  2. Commandes reçues vs CA déclaré (+ rapport mensuel et cumulé)
  3. Charge mensuelle (productif / support / interne)
  4. Backlog début de mois (évolution)
  5. Clôtures par mois (projets + rework)

+ tableau récapitulatif annuel.

Usage : python generate_pilotage.py
Sortie : pilotage.html
"""

import json
import pandas as pd

HISTORIQUE_FILE = "_historique.xlsx"
OUTPUT_HTML     = "pilotage.html"
TEMPLATE_FILE   = "_pilotage_template.html"

MOIS_ORDRE = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
              "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _ordonner(df):
    """Trie un df par ordre chronologique de mois."""
    df = df[df["Mois"].notna()].copy()
    df["_o"] = df["Mois"].map(lambda m: MOIS_ORDRE.index(m) if m in MOIS_ORDRE else 99)
    return df.sort_values("_o")


def build_data():
    xls = pd.ExcelFile(HISTORIQUE_FILE)
    df_ca   = _ordonner(pd.read_excel(xls, "ca"))       if "ca" in xls.sheet_names else pd.DataFrame()
    df_chg  = _ordonner(pd.read_excel(xls, "charge"))   if "charge" in xls.sheet_names else pd.DataFrame()
    df_bl   = _ordonner(pd.read_excel(xls, "backlog"))  if "backlog" in xls.sheet_names else pd.DataFrame()
    df_clot = _ordonner(pd.read_excel(xls, "clotures")) if "clotures" in xls.sheet_names else pd.DataFrame()
    df_cmd  = pd.read_excel(xls, "commandes")           if "commandes" in xls.sheet_names else pd.DataFrame()

    mois_labels = list(df_ca["Mois"]) if not df_ca.empty else []

    # 1. CA mensuel
    ca = {
        "global": [_num(v) for v in df_ca.get("CA réalisé (€)", [])],
        "litt":   [_num(v) for v in df_ca.get("CA LITTERALIS (€)", [])],
        "geodp":  [_num(v) for v in df_ca.get("CA GEODP (€)", [])],
        "obj":    [_num(v) for v in df_ca.get("Objectif global (€)", [])],
        "n1":     [_num(v) for v in df_ca.get("CA N-1 global (€)", [])],
    }

    # 2. Commandes par mois (somme des montants)
    cmd_par_mois = {}
    if not df_cmd.empty:
        col_m = "Montant total prestations (€)"
        if col_m in df_cmd.columns:
            g = df_cmd[df_cmd["Mois"].notna()].groupby("Mois")[col_m].sum()
            cmd_par_mois = {m: _num(v) for m, v in g.items()}
    commandes = [cmd_par_mois.get(m, 0) for m in mois_labels]

    # Rapport commandes / CA déclaré (mensuel et cumulé)
    rapport = []
    cum_cmd, cum_ca, rapport_cum = 0, 0, []
    for i, m in enumerate(mois_labels):
        c = commandes[i]
        d = ca["global"][i] if i < len(ca["global"]) else 0
        rapport.append(round(c / d, 2) if d > 0 else 0)
        cum_cmd += c
        cum_ca  += d
        rapport_cum.append(round(cum_cmd / cum_ca, 2) if cum_ca > 0 else 0)

    # 3. Charge mensuelle
    charge = {
        "prod":    [_num(a)+_num(b) for a, b in zip(df_chg.get("Productif Litt", []), df_chg.get("Productif GEODP", []))] if not df_chg.empty else [],
        "support": [_num(a)+_num(b) for a, b in zip(df_chg.get("Support Litt", []), df_chg.get("Support GEODP", []))] if not df_chg.empty else [],
        "interne": [_num(v) for v in df_chg.get("Interne", [])] if not df_chg.empty else [],
    }

    # 4. Backlog début de mois (sinon courant)
    if not df_bl.empty:
        if "backlog_debut_mois_TOTAL" in df_bl.columns and df_bl["backlog_debut_mois_TOTAL"].notna().any():
            backlog = [_num(v) for v in df_bl["backlog_debut_mois_TOTAL"]]
        else:
            backlog = [_num(v) for v in df_bl.get("total_backlog_TOTAL", [])]
    else:
        backlog = []

    # 5. Clôtures
    clot = {
        "projets": [int(_num(v)) for v in df_clot.get("Projets clôturés", [])] if not df_clot.empty else [],
        "rework":  [int(_num(v)) for v in df_clot.get("dont Rework", [])] if not df_clot.empty else [],
    }
    clot_mois = list(df_clot["Mois"]) if not df_clot.empty else []

    return {
        "mois": mois_labels,
        "ca": ca,
        "commandes": commandes,
        "rapport": rapport,
        "rapport_cum": rapport_cum,
        "charge": charge,
        "backlog": backlog,
        "clotures": clot,
        "clotures_mois": clot_mois,
        "annee": int(df_ca["Année"].iloc[0]) if not df_ca.empty else 2026,
    }


def render():
    data = build_data()
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()
    return template.replace("__DATA__", json.dumps(data, ensure_ascii=False))


if __name__ == "__main__":
    print("📈 Génération du pilotage annuel...")
    html = render()
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ {OUTPUT_HTML}")
