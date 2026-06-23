"""
orchestration.py
================
Détermine ce que le pipeline doit faire selon la date du jour.
Écrit des "flags" dans la sortie GitHub Actions ($GITHUB_OUTPUT) que le
workflow lit pour déclencher (ou non) les étapes conditionnelles.

Logique :
  - CHAQUE JOUR        : run du mois courant (toujours)
  - 1er JOUR DU MOIS   : EN PLUS → figer le mois précédent (mode precedent)
                         + capturer le prévisionnel du mois qui commence

Usage : python orchestration.py
Sortie : variables is_premier_jour=true/false dans $GITHUB_OUTPUT
"""

import os
from datetime import date


def est_premier_jour_du_mois(d=None):
    d = d or date.today()
    return d.day == 1


def main():
    d = date.today()
    premier = est_premier_jour_du_mois(d)

    print(f"📅 Date du jour : {d.strftime('%d/%m/%Y')}")
    print(f"   Premier jour du mois : {'OUI' if premier else 'non'}")
    if premier:
        print("   → Le pipeline va AUSSI figer le mois précédent et capturer le prévisionnel.")
    else:
        print("   → Run quotidien standard (mois courant).")

    # Écrire les flags pour le workflow
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"is_premier_jour={'true' if premier else 'false'}\n")


if __name__ == "__main__":
    main()
