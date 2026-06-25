"""
archiver_mois.py
================
Archive l'export Excel du mois qui vient d'être figé dans archives_mensuelles/,
sous un nom triable : AAAA-MM_Mois.xlsx (ex : 2026-06_Juin.xlsx).

À lancer juste après le figement du mois précédent (action figer_precedent).
Le mois figé est le mois PRÉCÉDENT par rapport à aujourd'hui.

Usage : python archiver_mois.py
"""

import glob
import os
import shutil
from datetime import date

ARCHIVE_DIR = "archives_mensuelles"
MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


def mois_precedent(d=None):
    """Retourne (année, numéro_mois, nom_mois) du mois précédent."""
    d = d or date.today()
    if d.month == 1:
        return d.year - 1, 12, MOIS_FR[11]
    return d.year, d.month - 1, MOIS_FR[d.month - 2]


def main():
    # Trouver l'export du jour (le plus récent jira_export_*.xlsx)
    exports = sorted(glob.glob("jira_export_*.xlsx"))
    if not exports:
        print("⚠️ Aucun export jira_export_*.xlsx trouvé — rien à archiver.")
        return

    source = exports[-1]   # le plus récent
    annee, num, nom = mois_precedent()

    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    cible = os.path.join(ARCHIVE_DIR, f"{annee}-{num:02d}_{nom}.xlsx")

    shutil.copy2(source, cible)
    print(f"📁 Mois figé archivé : {source} → {cible}")


if __name__ == "__main__":
    main()
