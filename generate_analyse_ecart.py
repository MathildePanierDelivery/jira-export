"""
generate_analyse_ecart.py
=========================
Génère l'onglet "Analyse de l'écart de CA" (analyse_ecart.html).

Répond à : pourquoi l'écart entre l'objectif et le CA réalisé ce mois ?
Décompose les heures "à 0 €" par cause, les valorise avec un taux horaire de
référence, et distingue le récupérable (jalons non atteints) du définitif.

Taux horaire de référence = moyenne, sur les 3 mois précédents, de
(CA réalisé du mois ÷ heures productives totales du mois).

Usage : python generate_analyse_ecart.py
Sortie : analyse_ecart.html
"""

import json
import pandas as pd
from horodatage import maj_texte

HISTORIQUE_FILE = "_historique.xlsx"
OUTPUT_HTML     = "analyse_ecart.html"
TEMPLATE_FILE   = "_analyse_ecart_template.html"

MOIS_ORDRE = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
              "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _mois_precedents(mois, annee, n=3):
    """Retourne la liste des (mois, année) des n mois précédant (mois, annee)."""
    idx = MOIS_ORDRE.index(mois)
    res = []
    a, i = annee, idx
    for _ in range(n):
        i -= 1
        if i < 0:
            i = 11
            a -= 1
        res.append((MOIS_ORDRE[i], a))
    return res


def build_data():
    xls = pd.ExcelFile(HISTORIQUE_FILE)
    df_ca     = pd.read_excel(xls, "ca")                  if "ca" in xls.sheet_names else pd.DataFrame()
    df_charge = pd.read_excel(xls, "charge")             if "charge" in xls.sheet_names else pd.DataFrame()
    df_tnv    = pd.read_excel(xls, "temps_non_valorise") if "temps_non_valorise" in xls.sheet_names else pd.DataFrame()

    if df_ca.empty or df_tnv.empty:
        return {"disponible": False}

    # Index CA et productif par (mois, année)
    ca_par_mois = {}
    for _, r in df_ca.iterrows():
        m, a = r.get("Mois"), r.get("Année")
        if pd.isna(m) or pd.isna(a):
            continue
        ca_par_mois[(m, int(a))] = {
            "realise": _num(r.get("CA réalisé (€)")),
            "obj": _num(r.get("Objectif global (€)")),
        }
    prod_par_mois = {}
    saisie_par_mois = {}
    for _, r in df_charge.iterrows():
        m, a = r.get("Mois"), r.get("Année")
        if pd.isna(m) or pd.isna(a):
            continue
        prod_par_mois[(m, int(a))] = _num(r.get("Productif Total"))
        saisie_par_mois[(m, int(a))] = {
            "saisies": _num(r.get("Heures saisies à date")),
            "attendues": _num(r.get("Capacité attendue à date")),
            "totale": _num(r.get("Capacité totale du mois")),
        }

    # Mois analysables = ceux présents dans l'onglet temps_non_valorise
    mois_dispo = []
    for _, r in df_tnv.iterrows():
        m, a = r.get("Mois"), r.get("Année")
        if not pd.isna(m) and not pd.isna(a):
            mois_dispo.append((m, int(a)))
    if not mois_dispo:
        return {"disponible": False}
    mois_dispo = sorted(set(mois_dispo), key=lambda x: (x[1], MOIS_ORDRE.index(x[0])))

    LABELS = {
        "jalons":  "Jalons non atteints",
        "depasse": "Chiffrage dépassé / sous-estimation",
        "rework":  "Rework",
        "bloque":  "Tickets bloqués",
        "gratuit": "Prestations offertes",
    }
    RECUPERABLE = {"jalons": True, "depasse": False, "rework": False,
                   "bloque": False, "gratuit": False}

    def _analyse_mois(mois, annee):
        # Taux horaire de référence : moyenne 3 mois précédents (CA / productif)
        taux_detail, taux_vals = [], []
        for (m, a) in _mois_precedents(mois, annee, 3):
            ca_m = ca_par_mois.get((m, a), {}).get("realise", 0)
            prod_m = prod_par_mois.get((m, a), 0)
            if prod_m > 0:
                t = ca_m / prod_m
                taux_vals.append(t)
                taux_detail.append({"mois": m, "annee": a, "ca": ca_m, "prod": prod_m, "taux": round(t, 2)})
        taux_ref = round(sum(taux_vals) / len(taux_vals), 2) if taux_vals else 0.0

        # Heures par cause (somme des 2 solutions)
        causes = {"jalons": 0.0, "depasse": 0.0, "rework": 0.0, "bloque": 0.0, "gratuit": 0.0}
        for _, r in df_tnv.iterrows():
            if r.get("Mois") != mois or int(_num(r.get("Année"))) != annee:
                continue
            causes["rework"]  += _num(r.get("Rework"))
            causes["gratuit"] += _num(r.get("Gratuit"))
            causes["jalons"]  += _num(r.get("dont Reste"))
            causes["bloque"]  += _num(r.get("dont Bloqué")) + _num(r.get("Projet avec CA - Bloqué"))
            causes["depasse"] += _num(r.get("dont Dépassé")) + _num(r.get("Projet avec CA - Dépassé"))

        lignes = []
        for k in ["jalons", "depasse", "rework", "bloque", "gratuit"]:
            h = round(causes[k], 2)
            lignes.append({"cle": k, "label": LABELS[k], "heures": h,
                           "montant": round(h * taux_ref, 0), "recuperable": RECUPERABLE[k]})

        total_h = round(sum(causes.values()), 2)
        ca_mois = ca_par_mois.get((mois, annee), {})
        # Complétude de saisie (heures saisies vs attendues à date)
        sais = saisie_par_mois.get((mois, annee), {})
        h_saisies = sais.get("saisies", 0)
        h_attendues = sais.get("attendues", 0)
        pct_saisie = round(h_saisies / h_attendues * 100) if h_attendues > 0 else None
        return {
            "mois": mois, "annee": annee,
            "objectif": round(ca_mois.get("obj", 0), 0),
            "realise": round(ca_mois.get("realise", 0), 0),
            "ecart_reel": round(ca_mois.get("obj", 0) - ca_mois.get("realise", 0), 0),
            "taux_ref": taux_ref, "taux_detail": taux_detail, "lignes": lignes,
            "total_heures": total_h, "total_montant": round(total_h * taux_ref, 0),
            "recuperable_montant": round(sum(l["montant"] for l in lignes if l["recuperable"]), 0),
            "definitif_montant": round(sum(l["montant"] for l in lignes if not l["recuperable"]), 0),
            "h_saisies": round(h_saisies, 1), "h_attendues": round(h_attendues, 1),
            "pct_saisie": pct_saisie,
        }

    # Calculer tous les mois, indexés par "Mois AAAA"
    analyses = {}
    cles_ordonnees = []
    for (m, a) in mois_dispo:
        cle = f"{m} {a}"
        analyses[cle] = _analyse_mois(m, a)
        cles_ordonnees.append(cle)

    return {
        "disponible": True,
        "analyses": analyses,
        "mois_ordonnes": cles_ordonnees,
        "defaut": cles_ordonnees[-1],   # le plus récent
    }


def render():
    data = build_data()
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()
    return (template
            .replace("__DATA__", json.dumps(data, ensure_ascii=False))
            .replace("__MAJ__", maj_texte()))


if __name__ == "__main__":
    print("📊 Génération de l'analyse d'écart de CA...")
    html = render()
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ {OUTPUT_HTML}")
