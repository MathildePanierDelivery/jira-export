"""
controle_coherence_ca.py
========================
Contrôle de cohérence du CA déclaré dans Jira, à partir de jira_cache.pkl.
Repère deux types d'anomalies (bugs Jira fréquents) :
  1. Dépassement   : total reconnu > montant de la commande
  2. Avancement    : total reconnu ≠ (% avancement × montant commande), tolérance 1€

N'échoue JAMAIS le pipeline. Écrit un rapport :
  - dans la console (logs du run)
  - dans le résumé du run GitHub (variable d'env GITHUB_STEP_SUMMARY)

Champs Jira utilisés (déjà dans le cache) :
  customfield_20701 = montant commande
  customfield_21883 = % avancement
  customfield_26980 = montant reconnu avant 2026
  customfield_22043 = montant reconnu (cumul)   [fallback]
"""

import os
import pickle

CACHE_FILE = "jira_cache.pkl"
TOLERANCE  = 1.0  # euros

CF_COMMANDE   = "customfield_20701"
CF_AVANCEMENT = "customfield_21883"
CF_AVANT_2026 = "customfield_26980"
CF_RECONNU    = "customfield_22043"
CF_CA_MOIS    = "customfield_22998"


def _f(raw):
    """Convertit une valeur de champ en float (gère dict option / None)."""
    if isinstance(raw, dict):
        raw = raw.get("value") or raw.get("name")
    try:
        return float(raw) if raw is not None else None
    except (TypeError, ValueError):
        return None


def charger_cache():
    if not os.path.exists(CACHE_FILE):
        raise FileNotFoundError(f"{CACHE_FILE} introuvable.")
    with open(CACHE_FILE, "rb") as f:
        return pickle.load(f)


def controler(cache):
    """Parcourt les tickets CA et repère les anomalies. Retourne deux listes."""
    issues = cache["issues_by_key"]
    depassements = []
    incoherences = []

    for key, info in issues.items():
        if info.get("issuetype") != "COORDIN : Suivi CA":
            continue
        raw = info["raw_fields"]

        commande   = _f(raw.get(CF_COMMANDE))
        avancement = _f(raw.get(CF_AVANCEMENT))
        avant_2026 = _f(raw.get(CF_AVANT_2026)) or 0
        reconnu    = _f(raw.get(CF_RECONNU))

        # Total reconnu cumulé : champ dédié, sinon avant 2026 + CA du mois
        if reconnu is None:
            ca_mois = _f(raw.get(CF_CA_MOIS)) or 0
            reconnu = avant_2026 + ca_mois

        client = raw.get("customfield_10070")
        if isinstance(client, dict):
            client = client.get("displayName") or client.get("name") or ""
        projet = raw.get("customfield_23608") or ""
        libelle = f"{key} — {client or projet or info.get('summary','')}".strip()

        # 1. Dépassement : reconnu > commande
        if commande is not None and reconnu is not None and reconnu > commande + TOLERANCE:
            depassements.append({
                "ticket": key, "libelle": libelle,
                "reconnu": reconnu, "commande": commande,
                "ecart": reconnu - commande,
            })

        # 2. Incohérence avancement : reconnu ≠ % avancement × commande
        if commande is not None and avancement is not None and reconnu is not None:
            attendu = avancement / 100 * commande
            diff = reconnu - attendu
            if abs(diff) > TOLERANCE:
                incoherences.append({
                    "ticket": key, "libelle": libelle,
                    "reconnu": reconnu, "attendu": attendu,
                    "avancement": avancement, "commande": commande,
                    "ecart": diff,
                })

    return depassements, incoherences


def _euro(v):
    return f"{v:,.2f} €".replace(",", " ")


def ecrire_rapport(cache, depassements, incoherences):
    mois = cache.get("mois_label", "")
    lignes = []
    lignes.append(f"# Contrôle de cohérence CA — {mois}")
    lignes.append("")

    if not depassements and not incoherences:
        lignes.append("✅ **Aucune anomalie détectée.** Les montants déclarés sont cohérents.")
    else:
        n = len(depassements) + len(incoherences)
        lignes.append(f"⚠️ **{n} anomalie(s) détectée(s)** — à vérifier dans Jira.")
        lignes.append("")

        if depassements:
            lignes.append(f"## Dépassements ({len(depassements)})")
            lignes.append("Total reconnu supérieur au montant de la commande.")
            lignes.append("")
            lignes.append("| Ticket | Projet / Client | Reconnu | Commande | Écart |")
            lignes.append("|--------|-----------------|--------:|---------:|------:|")
            for d in sorted(depassements, key=lambda x: -x["ecart"]):
                lignes.append(f"| {d['ticket']} | {d['libelle'].split('—')[-1].strip()} "
                              f"| {_euro(d['reconnu'])} | {_euro(d['commande'])} | +{_euro(d['ecart'])} |")
            lignes.append("")

        if incoherences:
            lignes.append(f"## Écarts d'avancement ({len(incoherences)})")
            lignes.append("Le CA reconnu ne correspond pas à l'avancement déclaré.")
            lignes.append("")
            lignes.append("| Ticket | Projet / Client | Avanc. | Reconnu | Attendu | Écart |")
            lignes.append("|--------|-----------------|-------:|--------:|--------:|------:|")
            for d in sorted(incoherences, key=lambda x: -abs(x["ecart"])):
                lignes.append(f"| {d['ticket']} | {d['libelle'].split('—')[-1].strip()} "
                              f"| {d['avancement']:.0f}% | {_euro(d['reconnu'])} "
                              f"| {_euro(d['attendu'])} | {d['ecart']:+,.2f} € |".replace(",", " "))
            lignes.append("")

    rapport = "\n".join(lignes)

    # Console
    print(rapport)

    # Résumé du run GitHub
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(rapport + "\n")


if __name__ == "__main__":
    try:
        cache = charger_cache()
        dep, inc = controler(cache)
        ecrire_rapport(cache, dep, inc)
    except Exception as e:
        # Ne jamais faire échouer le pipeline pour un souci de contrôle
        print(f"⚠️ Contrôle de cohérence non effectué : {e}")
