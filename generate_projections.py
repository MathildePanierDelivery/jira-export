"""
generate_projections.py
=======================
Génère l'onglet d'analyse du backlog (projections.html) à partir de la photo
courante stockée dans _historique.xlsx (onglet backlog_facturable).

Analyses (distinction Littéralis / GEODP) :
  1. Composantes du backlog facturable : mobilisable vs bloqué (montant + nb)
  2. Rappel du modèle de jalonnement (4 phases)
  3. Répartition des montants par phase

Usage : python generate_projections.py
Sortie : projections.html
"""

import json
import pandas as pd

HISTORIQUE_FILE = "_historique.xlsx"
OUTPUT_HTML     = "projections.html"
TEMPLATE_FILE   = "_projections_template.html"

PHASES = ["Lancement & prérequis", "Mise en service",
          "Paramétrage & recette", "PV signé"]
BLOCAGES = ["Aucun", "Blocage client", "Blocage produit", "Blocage commerce", "Autre"]


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def build_data():
    xls = pd.ExcelFile(HISTORIQUE_FILE)
    if "backlog_facturable" not in xls.sheet_names:
        return {"disponible": False}

    df = pd.read_excel(xls, "backlog_facturable")
    solutions = {}
    photo = ""
    for _, r in df.iterrows():
        sol = r["Solution"]
        photo = str(r.get("Date photo", ""))
        solutions[sol] = {
            "total": _num(r.get("Total")),
            "mobilisable": _num(r.get("Mobilisable")),
            "bloque": _num(r.get("Bloqué")),
            "nb": int(_num(r.get("Nb"))),
            "nb_mobilisable": int(_num(r.get("Nb mobilisable"))),
            "nb_bloque": int(_num(r.get("Nb bloqué"))),
            "temps_total": _num(r.get("Temps total")),
            "temps_mobilisable": _num(r.get("Temps mobilisable")),
            "temps_bloque": _num(r.get("Temps bloqué")),
            "par_phase": {p: _num(r.get(f"Phase: {p}")) for p in PHASES},
            "nb_par_phase": {p: int(_num(r.get(f"NbPhase: {p}"))) for p in PHASES},
            "par_blocage": {b: _num(r.get(f"Bloc: {b}")) for b in BLOCAGES},
            "nb_par_blocage": {b: int(_num(r.get(f"NbBloc: {b}"))) for b in BLOCAGES},
        }

    # Total global (somme des deux solutions)
    litt = solutions.get("LITTERALIS", {})
    geodp = solutions.get("GEODP", {})

    return {
        "disponible": True,
        "photo": photo,
        "litt": litt,
        "geodp": geodp,
        "phases": PHASES,
        "blocages": BLOCAGES,
    }


def render():
    data = build_data()
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()
    return template.replace("__DATA__", json.dumps(data, ensure_ascii=False))


if __name__ == "__main__":
    print("📊 Génération de l'analyse backlog...")
    html = render()
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ {OUTPUT_HTML}")
