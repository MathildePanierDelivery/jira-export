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
from horodatage import maj_texte

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
    df_clot   = pd.read_excel(xls, "clotures")  if "clotures" in xls.sheet_names else pd.DataFrame()
    df_anc    = pd.read_excel(xls, "anciennete") if "anciennete" in xls.sheet_names else pd.DataFrame()
    df_tnv    = pd.read_excel(xls, "temps_non_valorise") if "temps_non_valorise" in xls.sheet_names else pd.DataFrame()

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
        prod_total = _num(r.get("Productif Total")) or (pl + pg)
        sup_total  = _num(r.get("Support Total")) or (sl + sg)
        interne    = _num(r.get("Interne"))
        # Heures saisies : colonne dédiée si présente (mois courant),
        # sinon total des heures du mois (mois clos) = prod + support + interne.
        h_saisies = _num(r.get("Heures saisies à date"))
        if h_saisies == 0:
            h_saisies = prod_total + sup_total + interne
        slot(annee, mois)["charge"] = {
            "prod_litt": pl, "prod_geodp": pg,
            "sup_litt": sl, "sup_geodp": sg,
            "rew_litt": rl, "rew_geodp": rg,
            "proj_litt":    _num(r.get("Projet Litt")),
            "proj_geodp":   _num(r.get("Projet GEODP")),
            "gratuit_litt": _num(r.get("Gratuit Litt")),
            "gratuit_geodp":_num(r.get("Gratuit GEODP")),
            "interne": interne,
            "prod_total": prod_total,
            "sup_total":  sup_total,
            "rew_total":  _num(r.get("Rework Total")) or (rl + rg),
            "heures_saisies":     h_saisies,
            "capacite_attendue":  _num(r.get("Capacité attendue à date")),
            "capacite_totale":    _num(r.get("Capacité totale du mois")),
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

    for _, r in df_clot.iterrows():
        mois, annee = r.get("Mois"), r.get("Année")
        if pd.isna(mois) or pd.isna(annee):
            continue
        slot(annee, mois)["clotures"] = {
            "projets": int(_num(r.get("Projets clôturés"))),
            "rework":  int(_num(r.get("dont Rework"))),
        }

    # Ancienneté du CA : répartition du CA du mois par tranche d'âge, PAR SOLUTION
    def _bucket(anc):
        if anc <= 0:  return "M+0"
        if anc == 1:  return "M+1"
        if anc == 2:  return "M+2"
        if anc == 3:  return "M+3"
        if anc <= 6:  return "M+4→6"
        if anc <= 12: return "M+7→12"
        return "M+13+"
    BUCKETS = ["M+0", "M+1", "M+2", "M+3", "M+4→6", "M+7→12", "M+13+"]

    def _norm_sol(s):
        s = str(s).lower()
        if "geodp" in s: return "GEODP"
        if "litt" in s or "liess" in s or "sherpa" in s: return "LITTERALIS"
        return "Autre"

    if not df_anc.empty:
        for (annee, mois), grp in df_anc.groupby(["Année", "Mois"]):
            if pd.isna(mois) or pd.isna(annee):
                continue
            # tranches par solution : CA + nombre de BDC
            par_sol = {"LITTERALIS": {b: 0.0 for b in BUCKETS},
                       "GEODP":      {b: 0.0 for b in BUCKETS}}
            nb_sol  = {"LITTERALIS": {b: 0 for b in BUCKETS},
                       "GEODP":      {b: 0 for b in BUCKETS}}
            for _, r in grp.iterrows():
                ca = _num(r.get("CA"))
                if ca <= 0:
                    continue
                sol = _norm_sol(r.get("Solution"))
                if sol not in par_sol:
                    continue
                b = _bucket(int(_num(r.get("Ancienneté"))))
                par_sol[sol][b] += ca
                nb_sol[sol][b]  += 1

            def _pct_vieux(tr):
                t = sum(tr.values())
                return round((tr["M+7→12"] + tr["M+13+"]) / t * 100, 1) if t > 0 else 0

            total_global = sum(sum(par_sol[s].values()) for s in par_sol)
            slot(annee, mois)["anciennete"] = {
                "litt":  {"tranches": par_sol["LITTERALIS"], "nb": nb_sol["LITTERALIS"],
                          "total": sum(par_sol["LITTERALIS"].values()),
                          "pct_vieux": _pct_vieux(par_sol["LITTERALIS"])},
                "geodp": {"tranches": par_sol["GEODP"], "nb": nb_sol["GEODP"],
                          "total": sum(par_sol["GEODP"].values()),
                          "pct_vieux": _pct_vieux(par_sol["GEODP"])},
                "total": total_global,
                "ordre": BUCKETS,
            }

    # ── Temps non valorisé (par mois × solution) ──
    for _, r in df_tnv.iterrows():
        mois, annee = r.get("Mois"), r.get("Année")
        if pd.isna(mois) or pd.isna(annee):
            continue
        sol = r.get("Solution")
        if sol not in ("LITTERALIS", "GEODP"):
            continue
        def _n(col):
            v = r.get(col)
            try:
                return round(float(v), 2) if not pd.isna(v) else 0.0
            except (ValueError, TypeError):
                return 0.0
        s = slot(annee, mois).setdefault("tnv", {})
        s[sol] = {
            "rework": _n("Rework"), "gratuit": _n("Gratuit"),
            "projet_sans_ca": _n("Projet sans CA"),
            "sans_ca_bloque": _n("dont Bloqué"), "sans_ca_depasse": _n("dont Dépassé"),
            "sans_ca_reste": _n("dont Reste"),
            "avec_ca_bloque": _n("Projet avec CA - Bloqué"),
            "avec_ca_depasse": _n("Projet avec CA - Dépassé"),
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
            .replace("__DEF_MOIS__", dernier_mois)
            .replace("__MAJ__", maj_texte()))


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
