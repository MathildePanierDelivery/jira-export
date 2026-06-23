"""
collect_jira_worklogs.py  —  VERSION A (Tempo)
==============================================
Collecteur unique des données Jira + Tempo.

Source des WORKLOGS  : API Tempo (api.eu.tempo.io) — donne le VRAI auteur,
                       contrairement à l'API Jira native qui renvoie le bot
                       "Timesheets by Tempo" sur les saisies faites via Tempo.
Source des ISSUES    : API Jira (métadonnées brutes : solution, catégorie,
                       client, montants CA, etc.)

Le cache produit deux étages :
  - worklogs : liste filtrée sur l'équipe (issue_key déjà résolu, nom, heures, date)
  - issues   : dict {key -> {id, key, summary, issuetype, parent, raw_fields}}
               raw_fields = TOUS les champs Jira bruts (les scripts transforment eux-mêmes)

Usage :
    python collect_jira_worklogs.py [--mois precedent|courant] [--refresh]
"""

import os
import sys
import pickle
import argparse
import requests
from datetime import datetime, date
from collections import defaultdict

from jira import JIRA

# ═══════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════
JIRA_URL    = os.environ.get("JIRA_URL", "https://sogelink.atlassian.net")
JIRA_EMAIL  = os.environ.get("JIRA_EMAIL", "mathilde.panier@sogelink.com")
# Accepte les deux noms : JIRA_TOKEN (poste local) ou JIRA_API_TOKEN (secret GitHub)
JIRA_TOKEN  = os.environ.get("JIRA_API_TOKEN") or os.environ.get("JIRA_TOKEN", "")
TEMPO_TOKEN = os.environ.get("TEMPO_TOKEN", "")

TEMPO_BASE = "https://api.eu.tempo.io/4"

# Mapping accountId -> nom (référence : extract_clockwork_times.py)
# C'est CE mapping qui définit le périmètre "équipe" : seuls ces 8 comptes
# sont conservés dans les worklogs. Tout le reste (autres équipes) est ignoré.
ACCOUNT_TO_COLLAB = {
    '63fe26c90a4a47fb8d213c54'                    : 'Quentin Bordillon',
    '712020:bb1c3fc2-7106-4c8b-9955-2e4917706722' : 'Fabien Reutenauer',
    '63fe26c7f00d095406f2590c'                    : 'Bérénice Bossard',
    '712020:f59016f5-1528-45f5-a182-d91eb6f69cad' : 'Flavie Bardin',
    '712020:e3996f09-39b9-4957-8f83-d925d9543c98' : 'Maxime Pontonnier',
    '712020:30146e03-1c8d-42f4-a105-c6987f9b2e2c' : 'Marine Masingarbe',
    '712020:ab30e040-78a7-40e1-aef4-ff6781645c8e' : 'Duncan Hamelin',
    '712020:a40869a4-1333-4240-ac8a-670ed2314e0b' : 'Remy Vincent',
}

MOIS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
           "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Champs Jira réellement lus par les 4 scripts (au lieu de *all → 10x plus rapide).
# Standard + tous les customfields repérés dans Export_worklog_V3, Tps_collab,
# extract_clockwork_times, maj_reconnaissance_ca, + temps bloqué/dépassé.
JIRA_FIELDS = [
    # --- standard ---
    "summary", "status", "parent", "issuetype", "created", "resolutiondate",
    "assignee", "issuelinks",
    # --- métier / CA ---
    "customfield_10070",   # Client
    "customfield_10220",   # Montant Matériel
    "customfield_10221",   # BDC matériel
    "customfield_20514",   # Solution cible
    "customfield_20701",   # Montant commande (priorité)
    "customfield_21883",   # (maj_reconnaissance_ca)
    "customfield_22042",   # Montant restant à reconnaître
    "customfield_22043",   # Montant reconnu
    "customfield_22108",   # Deal Type
    "customfield_22998",   # CA du mois en cours
    "customfield_23103",   # Type d'Epic (Rework ?)
    "customfield_23595",   # Niveau de confiance (prévisionnel pondéré)
    "customfield_23608",
    "customfield_23610",   # Revenu planifié
    "customfield_23955",
    "customfield_24269",
    "customfield_24437",
    "customfield_24438",   # N° Order Salesforce / BDC
    "customfield_24529",
    "customfield_24564",   # Montant commande (fallback)
    "customfield_26148",   # Temps passé bloqué
    "customfield_26606",   # Temps passé dépassé
    "customfield_26980",   # (maj_reconnaissance_ca)
]


# ═══════════════════════════════════════════════════════════════════
# UTILITAIRES DATE
# ═══════════════════════════════════════════════════════════════════
def compute_month_bounds(month_mode):
    """Retourne (month_start, month_end, mois_label) — month_end exclusif."""
    today = datetime.now()
    if month_mode == "precedent":
        target = today.replace(day=1)
        if target.month > 1:
            target = target.replace(month=target.month - 1)
        else:
            target = target.replace(year=target.year - 1, month=12)
    else:  # courant
        target = today.replace(day=1)

    month_start = target.date()
    if month_start.month < 12:
        month_end = month_start.replace(month=month_start.month + 1)
    else:
        month_end = month_start.replace(year=month_start.year + 1, month=1)

    mois_label = MOIS_FR[month_start.month - 1]
    return month_start, month_end, mois_label


# ═══════════════════════════════════════════════════════════════════
# ÉTAGE 1 — WORKLOGS DEPUIS TEMPO
# ═══════════════════════════════════════════════════════════════════
def fetch_tempo_worklogs(month_start, month_end):
    """
    Récupère TOUS les worklogs Tempo entre month_start et month_end-1 (inclus),
    puis ne garde que ceux des comptes de l'équipe (ACCOUNT_TO_COLLAB).

    Tempo borne déjà par 'from'/'to' (to INCLUSIF), donc on passe
    to = dernier jour du mois = month_end - 1 jour.
    """
    from datetime import timedelta
    to_inclusive = month_end - timedelta(days=1)

    headers = {"Authorization": f"Bearer {TEMPO_TOKEN}"}
    url = f"{TEMPO_BASE}/worklogs"
    params = {
        "from":  month_start.isoformat(),
        "to":    to_inclusive.isoformat(),
        "limit": 100,
    }

    raw = []          # tous les worklogs Tempo (équipe seulement), avec issue.id
    total_seen = 0
    page = 0

    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()

        results = data.get("results", [])
        total_seen += len(results)
        page += 1

        for wl in results:
            acc = wl["author"]["accountId"]
            nom = ACCOUNT_TO_COLLAB.get(acc)
            if not nom:
                continue  # hors équipe → ignoré

            raw.append({
                "issue_id": str(wl["issue"]["id"]),   # ⚠ ID numérique, pas la clé
                "author_id": acc,
                "author_name": nom,
                "hours": wl["timeSpentSeconds"] / 3600.0,
                "date": date.fromisoformat(wl["startDate"]),
            })

        print(f"   page {page:>3} | {total_seen} worklogs vus | "
              f"{len(raw)} équipe retenus", end="\r")

        # Pagination : 'next' est une URL complète, ou absente en fin
        url = data.get("metadata", {}).get("next")
        params = None  # l'URL next contient déjà from/to/offset/limit

    print()
    return raw


# ═══════════════════════════════════════════════════════════════════
# ÉTAGE 2 — ISSUES DEPUIS JIRA (métadonnées brutes)
# ═══════════════════════════════════════════════════════════════════
def fetch_clotures_du_mois(jira, month_start):
    """
    Récupère les clés des tickets passés à "Terminé" DURANT le mois,
    via le mécanisme changelog de Jira (status changed TO ... AFTER ...).
    Indispensable car resolutiondate est souvent vide.
    Retourne (hw_clotures, epics_clotures) : deux ensembles de clés.
    """
    debut = month_start.isoformat()

    def _keys(jql):
        try:
            return {i.key for i in jira.search_issues(jql, maxResults=False, fields="key")}
        except Exception as e:
            print(f"   ⚠ JQL clôtures échouée : {e}")
            return set()

    # Tickets Matériel GEODP clôturés ce mois (onglet Hardware)
    hw = _keys(
        f'project = PDMDEP AND issuetype = "Matériel GEODP" '
        f'AND status changed TO "Terminé" AFTER "{debut}"'
    )
    # Épics PDMDEP clôturées ce mois (onglet Projets clôturés)
    ep = _keys(
        f'project = PDMDEP AND issuetype = Epic '
        f'AND status changed TO "Terminé" AFTER "{debut}"'
    )
    print(f"   → {len(hw)} Matériel GEODP clôturés, {len(ep)} épics clôturées ce mois")
    return hw, ep


def fetch_jira_issues(jira):
    """
    Récupère toutes les issues PDMDEP + PSC avec TOUS leurs champs bruts.
    Sert à : (a) la jointure issue.id -> issue.key
             (b) fournir aux scripts les métadonnées (solution, CA, client...).
    """
    print("   Requête JQL project in (PDMDEP, PSC)...")
    issues = jira.search_issues(
        'project in (PDMDEP, PSC)',
        maxResults=False,
        fields=",".join(JIRA_FIELDS),   # champs ciblés → bien plus rapide que *all
        expand="",
    )
    print(f"   {len(issues)} issues récupérées")

    issues_by_key = {}
    id_to_key = {}

    for iss in issues:
        key = iss.key
        iid = str(iss.id)
        id_to_key[iid] = key

        parent = getattr(iss.fields, "parent", None)
        parent_key = parent.key if parent else None

        itype = getattr(iss.fields, "issuetype", None)
        itype_name = itype.name if itype else None

        issues_by_key[key] = {
            "id": iid,
            "key": key,
            "summary": getattr(iss.fields, "summary", ""),
            "issuetype": itype_name,
            "parent_key": parent_key,
            "raw_fields": iss.raw.get("fields", {}),   # dict brut complet
        }

    return issues_by_key, id_to_key


# ═══════════════════════════════════════════════════════════════════
# COLLECTE PRINCIPALE
# ═══════════════════════════════════════════════════════════════════
def collect_all(month_mode):
    # --- garde-fous tokens ---
    missing = [n for n, v in [("JIRA_TOKEN", JIRA_TOKEN),
                              ("TEMPO_TOKEN", TEMPO_TOKEN)] if not v]
    if missing:
        print(f"❌ Variable(s) manquante(s) : {', '.join(missing)}")
        sys.exit(1)

    month_start, month_end, mois_label = compute_month_bounds(month_mode)
    print(f"📅 Mode : {month_mode} | Mois : {mois_label} {month_start.year}")
    print(f"   Période : {month_start} → {month_end} (exclus)\n")

    # --- Étage 1 : worklogs Tempo ---
    print("⏱️  Worklogs depuis Tempo...")
    tempo_raw = fetch_tempo_worklogs(month_start, month_end)
    print(f"   ✅ {len(tempo_raw)} worklogs équipe\n")

    # --- Étage 2 : issues Jira ---
    print("📥 Issues depuis Jira...")
    jira = JIRA(server=JIRA_URL, basic_auth=(JIRA_EMAIL, JIRA_TOKEN))
    issues_by_key, id_to_key = fetch_jira_issues(jira)

    # --- Clôtures du mois (via changelog) ---
    print("🔒 Clôtures du mois...")
    hw_clotures, epics_clotures = fetch_clotures_du_mois(jira, month_start)
    print()

    # --- Jointure issue.id -> issue.key sur les worklogs ---
    worklogs = []
    orphelins = 0
    for wl in tempo_raw:
        key = id_to_key.get(wl["issue_id"])
        if not key:
            orphelins += 1
            continue
        worklogs.append({
            "issue_key": key,
            "author_id": wl["author_id"],
            "author_name": wl["author_name"],
            "hours": wl["hours"],
            "date": wl["date"],
        })

    if orphelins:
        print(f"   ⚠ {orphelins} worklogs sur des tickets hors PDMDEP/PSC (ignorés)")

    # --- Index parent -> enfants (utile aux scripts) ---
    children_by_parent = defaultdict(list)
    for key, info in issues_by_key.items():
        if info["parent_key"]:
            children_by_parent[info["parent_key"]].append(key)

    cache = {
        "month_start": month_start,
        "month_end": month_end,
        "mois_label": mois_label,
        "worklogs": worklogs,
        "issues_by_key": issues_by_key,
        "id_to_key": id_to_key,
        "children_by_parent": dict(children_by_parent),
        "hw_clotures": hw_clotures,
        "epics_clotures": epics_clotures,
        "account_to_collab": ACCOUNT_TO_COLLAB,
        "generated_at": datetime.now(),
    }
    return cache


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Collecteur unique Jira + Tempo")
    parser.add_argument("--mois", choices=["precedent", "courant"],
                        default="precedent")
    parser.add_argument("--refresh", action="store_true",
                        help="Force le rechargement même si un cache existe")
    args = parser.parse_args()

    cache_file = "jira_cache.pkl"

    if os.path.exists(cache_file) and not args.refresh:
        print(f"✅ Cache existant : {cache_file} (--refresh pour recharger)")
        with open(cache_file, "rb") as f:
            cache = pickle.load(f)
    else:
        cache = collect_all(args.mois)
        with open(cache_file, "wb") as f:
            pickle.dump(cache, f)
        print(f"\n💾 Cache sauvegardé : {cache_file}")

    # --- Résumé + contrôle par collaborateur ---
    from collections import Counter
    print(f"\n📊 Résumé — {cache['mois_label']} {cache['month_start'].year}")
    print(f"   Worklogs équipe : {len(cache['worklogs'])}")
    print(f"   Issues          : {len(cache['issues_by_key'])}")

    par_nom = Counter(wl["author_name"] for wl in cache["worklogs"])
    heures  = defaultdict(float)
    for wl in cache["worklogs"]:
        heures[wl["author_name"]] += wl["hours"]

    print("\n   Par collaborateur :")
    for nom, n in par_nom.most_common():
        print(f"     {nom:<22} {n:>4} logs | {heures[nom]:>7.1f}h")

    total_h = sum(heures.values())
    print(f"\n   TOTAL : {len(cache['worklogs'])} logs | {total_h:.1f}h")
