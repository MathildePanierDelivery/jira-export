"""
generate_objectifs.py
=====================
Génère le dashboard "Suivi des objectifs d'équipe" (objectifs.html) à partir de :
  - objectifs.xlsx  → onglet ObjectifsEquipe (cibles + paliers, révisables)
  - _historique.xlsx → réalisé cumulé (CA annuel, clôtures rework cumulées)

Trois objectifs :
  1. CA non récurrent  : CA cumulé annuel vs paliers (déclencheur / palier 2 / target)
  2. Réduction rework  : cumul des projets rework clôturés depuis janvier vs paliers
  3. CSAT              : lien vers le fichier SharePoint

Usage : python generate_objectifs.py
Sortie : objectifs.html
"""

import os
import json
from horodatage import maj_texte
import pandas as pd

HISTORIQUE_FILE = "_historique.xlsx"
OBJECTIFS_FILE  = "objectifs.xlsx"
OBJ_ONGLET      = "ObjectifsEquipe"
OUTPUT_HTML     = "objectifs.html"
TEMPLATE_FILE   = "_objectifs_template.html"

MOIS_ORDRE = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
              "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def lire_objectifs():
    """Lit l'onglet ObjectifsEquipe → dict structuré."""
    wb = pd.ExcelFile(OBJECTIFS_FILE)
    df = pd.read_excel(wb, OBJ_ONGLET, header=None)

    def trouver_section(cle):
        for i, row in df.iterrows():
            if isinstance(row[0], str) and cle in row[0].upper():
                return i
        return None

    obj = {"ca": {"paliers": []}, "rework": {"paliers": [], "contexte": ""}, "csat": {"lien": ""}}

    # CA non récurrent
    i = trouver_section("CA NON RÉCURRENT")
    if i is not None:
        for r in range(i + 2, i + 5):  # 3 paliers
            nom, seuil, cible = df.iloc[r, 0], df.iloc[r, 1], df.iloc[r, 2]
            if pd.notna(cible):
                obj["ca"]["paliers"].append({"nom": str(nom), "seuil": _num(seuil), "cible": _num(cible)})

    # Rework
    i = trouver_section("REWORK")
    if i is not None:
        for r in range(i + 2, i + 5):
            nom, seuil, cible = df.iloc[r, 0], df.iloc[r, 1], df.iloc[r, 2]
            if pd.notna(cible) and isinstance(nom, str) and "Contexte" not in nom:
                obj["rework"]["paliers"].append({"nom": str(nom), "seuil": _num(seuil), "cible": int(_num(cible))})
        # contexte
        for r in range(i, min(i + 7, len(df))):
            if isinstance(df.iloc[r, 0], str) and "Contexte" in df.iloc[r, 0]:
                obj["rework"]["contexte"] = str(df.iloc[r, 2])

    # CSAT
    i = trouver_section("CSAT")
    if i is not None:
        for r in range(i, min(i + 3, len(df))):
            if isinstance(df.iloc[r, 0], str) and "ien" in df.iloc[r, 0]:  # "Lien"
                obj["csat"]["lien"] = str(df.iloc[r, 1])

    return obj


def lire_realise():
    """Lit l'historique → CA cumulé annuel + clôtures rework cumulées."""
    xls = pd.ExcelFile(HISTORIQUE_FILE)
    res = {"ca_cumule": 0.0, "ca_par_mois": [], "rework_cumule": 0, "rework_par_mois": []}

    # CA cumulé (onglet ca)
    if "ca" in xls.sheet_names:
        df = pd.read_excel(xls, "ca")
        df = df[df["Mois"].notna()]
        df["_ordre"] = df["Mois"].map(lambda m: MOIS_ORDRE.index(m) if m in MOIS_ORDRE else 99)
        df = df.sort_values("_ordre")
        cumul = 0.0
        for _, r in df.iterrows():
            v = _num(r.get("CA réalisé (€)"))
            cumul += v
            res["ca_par_mois"].append({"mois": r["Mois"], "valeur": v, "cumul": cumul})
        res["ca_cumule"] = cumul

    # Clôtures rework cumulées (onglet clotures)
    if "clotures" in xls.sheet_names:
        df = pd.read_excel(xls, "clotures")
        df = df[df["Mois"].notna()]
        df["_ordre"] = df["Mois"].map(lambda m: MOIS_ORDRE.index(m) if m in MOIS_ORDRE else 99)
        df = df.sort_values("_ordre")
        cumul = 0
        for _, r in df.iterrows():
            v = int(_num(r.get("dont Rework")))
            cumul += v
            res["rework_par_mois"].append({"mois": r["Mois"], "valeur": v, "cumul": cumul})
        res["rework_cumule"] = cumul

    return res


def render():
    obj = lire_objectifs()
    realise = lire_realise()

    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()

    payload = {"objectifs": obj, "realise": realise}
    return template.replace("__PAYLOAD__", json.dumps(payload, ensure_ascii=False)).replace("__MAJ__", maj_texte())


if __name__ == "__main__":
    print("🎯 Lecture des objectifs d'équipe...")
    html = render()
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ {OUTPUT_HTML}")
