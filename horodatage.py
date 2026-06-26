"""
horodatage.py
=============
Fournit l'horodatage "dernière mise à jour" en heure de Paris, pour affichage
dans l'en-tête des pages HTML.

Gère automatiquement l'heure d'été / d'hiver (Europe/Paris).
"""

from datetime import datetime, timezone, timedelta

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin",
           "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


def _maintenant_paris():
    """Retourne l'heure actuelle en fuseau Europe/Paris."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Paris"))
    except Exception:
        # Repli si zoneinfo indisponible : UTC+2 (été) approximatif.
        # On détecte grossièrement l'heure d'été (avr-oct = +2, sinon +1).
        now_utc = datetime.now(timezone.utc)
        offset = 2 if 3 <= now_utc.month <= 10 else 1
        return now_utc + timedelta(hours=offset)


def maj_texte():
    """Texte court : 'Mise à jour le 26/06/2026 à 12:03'."""
    d = _maintenant_paris()
    return f"Mise à jour le {d.strftime('%d/%m/%Y')} à {d.strftime('%H:%M')}"


def ecrire_fichier_maj(chemin="derniere_maj.txt"):
    """Écrit l'horodatage dans un fichier (lu par l'index statique)."""
    with open(chemin, "w", encoding="utf-8") as f:
        f.write(maj_texte())


if __name__ == "__main__":
    print(maj_texte())
