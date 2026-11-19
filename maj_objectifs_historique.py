"""
maj_objectifs_historique.py
===========================
Met à jour les objectifs des mois DÉJÀ figés dans _historique.xlsx (onglet ca)
avec la règle "dernière estimation" (LE2 > LE1 > Budget), valeur par valeur.

À lancer une seule fois pour corriger le passé après le changement de règle.
Ne touche qu'aux 3 colonnes Objectif ; le reste de l'historique est préservé.

Usage : python maj_objectifs_historique.py
"""

import openpyxl
import pandas as pd

HISTORIQUE_FILE = "_historique.xlsx"
OBJECTIFS_FILE  = "objectifs.xlsx"
PRIORITE = ["LE3", "LE2", "LE1", "Budget"]


def _lire_onglet(filepath, onglet, mois):
    try:
        df = pd.read_excel(filepath, sheet_name=onglet, header=3)
        ligne = df[df["Mois"] == mois]
        if ligne.empty:
            return None
        row = ligne.iloc[0]
        cg  = next((c for c in df.columns if "global" in c.lower()), None)
        cl  = next((c for c in df.columns if "litt" in c.lower()), None)
        cge = next((c for c in df.columns if "geodp" in c.lower()), None)
        def v(c):
            if not c: return 0.0
            try:
                x = float(row[c]); return x if x == x else 0.0
            except (ValueError, TypeError):
                return 0.0
        return v(cg), v(cl), v(cge)
    except Exception:
        return None


def objectif_dernier(mois):
    """(global, litt, geodp) = dernière estimation disponible pour le mois."""
    lectures = {}
    for o in PRIORITE:
        r = _lire_onglet(OBJECTIFS_FILE, o, mois)
        if r is not None:
            lectures[o] = r
    def choisir(idx):
        for o in PRIORITE:
            if o in lectures and lectures[o][idx] > 0:
                return lectures[o][idx]
        return 0.0
    return choisir(0), choisir(1), choisir(2)


def main():
    wb = openpyxl.load_workbook(HISTORIQUE_FILE)
    ws = wb["ca"]
    h = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    i_mois = h.index("Mois") + 1
    i_g  = h.index("Objectif global (€)") + 1
    i_l  = h.index("Objectif LITTERALIS (€)") + 1
    i_ge = h.index("Objectif GEODP (€)") + 1

    modifs = 0
    for r in range(2, ws.max_row + 1):
        mois = ws.cell(r, i_mois).value
        if not mois:
            continue
        g, l, ge = objectif_dernier(mois)
        anciens = (ws.cell(r, i_g).value, ws.cell(r, i_l).value, ws.cell(r, i_ge).value)
        ws.cell(r, i_g).value  = round(g, 2)
        ws.cell(r, i_l).value  = round(l, 2)
        ws.cell(r, i_ge).value = round(ge, 2)
        if anciens != (round(g, 2), round(l, 2), round(ge, 2)):
            print(f"  {mois}: {anciens[0]} → {round(g)}€ (global)")
            modifs += 1

    wb.save(HISTORIQUE_FILE)
    print(f"\n✅ {modifs} mois mis à jour dans {HISTORIQUE_FILE}")


if __name__ == "__main__":
    main()
