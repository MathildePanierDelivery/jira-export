"""
orchestration.py
================
Gère le VERROU DE BASCULE MENSUELLE et décide si le run doit traiter le mois
courant.

Principe :
  - Le 1er du mois (calendaire), le verrou se POSE automatiquement : le mois
    courant se met en pause (les données de l'ancien mois sont encore dans Jira,
    on ne veut pas les traiter comme "courant").
  - Pendant que le verrou est posé, les runs automatiques SAUTENT le mois courant.
  - L'utilisateur fait sa clôture (ajustements Jira → figement du mois précédent
    via un run manuel → remise à zéro Jira), puis LÈVE le verrou manuellement.
  - Filet de sécurité : si le verrou traîne plus de DELAI_LEVEE_AUTO jours, il
    est levé automatiquement (évite un mois courant gelé en cas d'oubli).

Le verrou est un fichier versionné dans le dépôt (bascule_en_cours.flag) pour
persister d'un run à l'autre. Il contient la date de pose (ISO).

Usage :
  python orchestration.py                 → run normal (pose/lève le verrou auto, décide)
  python orchestration.py --lever         → lève le verrou manuellement (fin de bascule)

Sorties dans $GITHUB_OUTPUT :
  bascule_en_cours = true|false   (le mois courant doit-il être sauté ?)
  verrou_pose      = true|false   (le verrou vient-il d'être posé à ce run ?)
"""

import os
import sys
from datetime import date, datetime

VERROU_FILE = "bascule_en_cours.flag"
DELAI_LEVEE_AUTO = 10   # jours avant levée automatique de sécurité


def _lire_date_pose():
    """Retourne la date de pose du verrou, ou None s'il n'existe pas."""
    if not os.path.exists(VERROU_FILE):
        return None
    try:
        with open(VERROU_FILE, encoding="utf-8") as f:
            txt = f.read().strip()
        return datetime.fromisoformat(txt).date()
    except Exception:
        # Fichier présent mais illisible : on considère le verrou posé aujourd'hui
        return date.today()


def poser_verrou(d=None):
    d = d or date.today()
    with open(VERROU_FILE, "w", encoding="utf-8") as f:
        f.write(d.isoformat())
    print(f"🔒 Verrou de bascule POSÉ ({d.isoformat()}) — mois courant en pause.")


def lever_verrou(motif=""):
    if os.path.exists(VERROU_FILE):
        os.remove(VERROU_FILE)
        print(f"🔓 Verrou de bascule LEVÉ{(' — ' + motif) if motif else ''}.")
        return True
    print("   (pas de verrou à lever)")
    return False


def _ecrire_sortie(bascule_en_cours, verrou_pose=False):
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"bascule_en_cours={'true' if bascule_en_cours else 'false'}\n")
            f.write(f"verrou_pose={'true' if verrou_pose else 'false'}\n")


def main():
    # Mode "lever manuellement" (fin de bascule, déclenché par l'utilisateur)
    if "--lever" in sys.argv:
        lever_verrou("levée manuelle (fin de bascule)")
        _ecrire_sortie(bascule_en_cours=False)
        return

    d = date.today()
    date_pose = _lire_date_pose()
    verrou_pose_ce_run = False

    if date_pose is None:
        # Pas de verrou actuellement. Faut-il le poser ? Oui si on est le 1er du mois.
        if d.day == 1:
            poser_verrou(d)
            date_pose = d
            verrou_pose_ce_run = True
            print("   → 1er du mois : entrée en bascule mensuelle.")
        else:
            print(f"📅 {d.isoformat()} : pas de bascule en cours, run normal (mois courant).")
    else:
        # Verrou présent. Levée auto de sécurité si trop ancien ?
        age = (d - date_pose).days
        print(f"🔒 Bascule en cours depuis {date_pose.isoformat()} ({age} j).")
        if age >= DELAI_LEVEE_AUTO:
            lever_verrou(f"levée AUTO de sécurité après {age} jours")
            date_pose = None

    bascule = date_pose is not None
    if bascule:
        print("   → Mois courant EN PAUSE (verrou actif). En attente de la fin de bascule.")
    _ecrire_sortie(bascule_en_cours=bascule, verrou_pose=verrou_pose_ce_run)


if __name__ == "__main__":
    main()
