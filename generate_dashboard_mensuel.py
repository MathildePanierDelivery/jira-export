"""
generate_dashboard_mensuel.py
=============================
Génère un tableau de bord mensuel HTML interactif (dashboard_mensuel.html)
à partir de _historique.xlsx. Sélecteur mois / année, données embarquées
en JSON, fichier autonome (hors-ligne, imprimable en PDF).

Version B : lit l'historique existant (CA, charge, commandes, backlog).

Usage : python generate_dashboard_mensuel.py
Sortie : dashboard_mensuel.html
"""

import os
import json
import pandas as pd

HISTORIQUE_FILE = "_historique.xlsx"
OUTPUT_HTML     = "dashboard_mensuel.html"
TEMPLATE_FILE   = "_dashboard_template.html"

MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def build_data():
    if not os.path.exists(HISTORIQUE_FILE):
        raise FileNotFoundError(f"{HISTORIQUE_FILE} introuvable.")

    xls = pd.ExcelFile(HISTORIQUE_FILE)
    df_ca     = pd.read_excel(xls, "ca")        if "ca" in xls.sheet_names else pd.DataFrame()
    df_charge = pd.read_excel(xls, "charge")    if "charge" in xls.sheet_names else pd.DataFrame()
    df_cmd    = pd.read_excel(xls, "commandes") if "commandes" in xls.sheet_names else pd.DataFrame()
    df_bl     = pd.read_excel(xls, "backlog")   if "backlog" in xls.sheet_names else pd.DataFrame()

    data = {}

    def slot(annee, mois):
        a = str(int(annee))
        data.setdefault(a, {})
        data[a].setdefault(mois, {})
        return data[a][mois]

    for _, r in df_ca.iterrows():
        mois, annee = r.get("Mois"), r.get("Année")
        if pd.isna(mois) or pd.isna(annee):
            continue
        slot(annee, mois)["ca"] = {
            "global": _num(r.get("CA réalisé (€)")),
            "litt":   _num(r.get("CA LITTERALIS (€)")),
            "geodp":  _num(r.get("CA GEODP (€)")),
            "obj_global": _num(r.get("Objectif global (€)")),
            "obj_litt":   _num(r.get("Objectif LITTERALIS (€)")),
            "obj_geodp":  _num(r.get("Objectif GEODP (€)")),
            "py_global": _num(r.get("CA PY global (€)")),
            "py_litt":   _num(r.get("CA PY LITTERALIS (€)")),
            "py_geodp":  _num(r.get("CA PY GEODP (€)")),
        }

    for _, r in df_charge.iterrows():
        mois, annee = r.get("Mois"), r.get("Année")
        if pd.isna(mois) or pd.isna(annee):
            continue
        pl, pg = _num(r.get("Productif Litt")), _num(r.get("Productif GEODP"))
        sl, sg = _num(r.get("Support Litt")), _num(r.get("Support GEODP"))
        rl, rg = _num(r.get("Rework Litt")), _num(r.get("Rework GEODP"))
        slot(annee, mois)["charge"] = {
            "prod_litt": pl, "prod_geodp": pg,
            "sup_litt": sl, "sup_geodp": sg,
            "rew_litt": rl, "rew_geodp": rg,
            "interne": _num(r.get("Interne")),
            "prod_total": _num(r.get("Productif Total")) or (pl + pg),
            "sup_total":  _num(r.get("Support Total")) or (sl + sg),
            "rew_total":  _num(r.get("Rework Total")) or (rl + rg),
        }

    col_montant = None
    if not df_cmd.empty:
        col_montant = next((c for c in df_cmd.columns if "montant" in str(c).lower()), None)
    if not df_cmd.empty:
        for (annee, mois), grp in df_cmd.groupby(["Année", "Mois"]):
            if pd.isna(mois) or pd.isna(annee):
                continue
            s = slot(annee, mois)
            if col_montant:
                litt  = grp[grp["Solution"] == "LITTERALIS"][col_montant].apply(_num).sum()
                geodp = grp[grp["Solution"] == "GEODP"][col_montant].apply(_num).sum()
                s["commandes"] = {"global": grp[col_montant].apply(_num).sum(),
                                  "litt": litt, "geodp": geodp, "nb": len(grp)}
            else:
                s["commandes"] = {"global": 0, "litt": 0, "geodp": 0, "nb": len(grp)}

    for _, r in df_bl.iterrows():
        mois, annee = r.get("Mois"), r.get("Année")
        if pd.isna(mois) or pd.isna(annee):
            continue
        slot(annee, mois)["backlog"] = {
            "total":       _num(r.get("total_backlog_TOTAL")),
            "bloque":      _num(r.get("total_bloque_TOTAL")),
            "mobilisable": _num(r.get("total_mobilisable_TOTAL")),
            "nb_projets":  int(_num(r.get("nb_projets_TOTAL"))),
            "nb_rework":   int(_num(r.get("nb_rework_TOTAL"))),
        }

    return data


def render_html(data):
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()

    annees = sorted(data.keys())
    derniere_annee = annees[-1] if annees else "2026"
    mois_dispo = [m for m in MOIS_FR if m in data.get(derniere_annee, {})]
    dernier_mois = mois_dispo[-1] if mois_dispo else "Janvier"

    return (template
            .replace("__DATA__", json.dumps(data, ensure_ascii=False))
            .replace("__MOIS__", json.dumps(MOIS_FR, ensure_ascii=False))
            .replace("__DEF_ANNEE__", derniere_annee)
            .replace("__DEF_MOIS__", dernier_mois))


if __name__ == "__main__":
    print("📊 Lecture de l'historique...")
    data = build_data()
    n = sum(len(v) for v in data.values())
    print(f"   {n} mois sur {len(data)} année(s)")
    print("🎨 Génération du HTML...")
    html = render_html(data)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ {OUTPUT_HTML}")
