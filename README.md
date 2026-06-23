# Pilotage de production — PS France

Pipeline automatisé **Jira + Tempo → Excel → Historique → Dashboard HTML** pour le suivi mensuel de la production (CA, charge, backlog, clôtures, prévisionnel) des solutions Littéralis et GEODP.

Dépôt : `MathildePanierDelivery/jira-export` · Publication : GitHub Pages

---

## Architecture

```
collect_jira_worklogs.py   →  jira_cache.pkl              (cache, NON publié)
        │   Jira API (issues + clôtures) · Tempo API (worklogs)
        ▼
controle_coherence_ca.py   →  rapport d'alertes           (résumé du run GitHub)
        │
export_mensuel.py          →  jira_export_AAAA-MM-JJ.xlsx (8 onglets)
        │                  →  _historique.xlsx            (accumulation du mois)
        ▼
generate_dashboard_mensuel.py  →  dashboard_mensuel.html  (lit _historique.xlsx)

generate_previsionnel.py   →  previsionnel_2026.xlsx      (snapshot 1er du mois)
```

---

## Les scripts

| Script | Rôle | Sorties |
|--------|------|---------|
| `collect_jira_worklogs.py` | Collecte Jira + Tempo en un passage (`--mois courant\|precedent`) | `jira_cache.pkl` |
| `controle_coherence_ca.py` | Détecte les incohérences CA (bugs Jira) → résumé du run | rapport |
| `export_mensuel.py` | Export Excel 8 onglets **et** accumulation historique | `jira_export_*.xlsx` |
| `maj_historique.py` | Module appelé par l'export : écrit la ligne du mois | — |
| `generate_dashboard_mensuel.py` | Dashboard web depuis l'historique | `dashboard_mensuel.html` |
| `generate_previsionnel.py` | Prévisionnel pondéré du mois (par CP, par solution) | `previsionnel_2026.xlsx` |
| `orchestration.py` | Détermine les actions selon la date (1er du mois ?) | flags workflow |

### Pourquoi un collecteur séparé ?

Les worklogs sont saisis via **Tempo**, pas l'API Jira native (qui renvoie un compte technique au lieu du vrai auteur). Le collecteur interroge l'API Tempo (`api.eu.tempo.io`), filtre sur l'équipe (8 collaborateurs via `accountId`) et joint à Jira. Une seule collecte alimente tout.

---

## L'export Excel (8 onglets)

1. **Tableau de bord** — synthèse : CA vs objectif, charge à date, projets clôturés / ouverts
2. **Suivi de production** — CA par épic (Hardware exclu)
3. **Suivi Hardware** — tickets Matériel GEODP clôturés dans le mois
4. **Capacité productive** — totaux d'équipe (aucune donnée nominative)
5. **Charge** — heures par solution et par type, % occupation
6. **Backlog** — restant à reconnaître, mobilisable vs bloqué
7. **Commandes du mois** — épics créées dans le mois
8. **Projets clôturés** — épics terminées, rework mis en évidence

### Règles métier clés

- **Solution** : normalisée par mot-clé insensible casse/accents → LITTERALIS / GEODP
- **Épics PSC** : solution = LITTERALIS ; catégorie = Projet si ticket CA rattaché, sinon Commande sans prestation
- **Temps par solution** : seul le productif compte ; Support et Interne à part
- **Matériel GEODP** : ne consomme pas d'heures (réparties sur les services)
- **Clôtures** : via changelog Jira (`status → "Terminé"`)
- **Objectifs** : onglet lu défini par `OBJECTIFS_ONGLET` dans `export_mensuel.py` (`LE1`)
- **Prévisionnel pondéré** : Prévision (cf 23610) × taux de confiance (cf 23595) —
  Haute 0,85 · Moyenne 0,50 · Basse 0,15 · non renseigné 0

---

## Accumulation & snapshots dans l'historique

À chaque run, `export_mensuel.py` écrit (ou écrase) la ligne du **mois courant** dans
`_historique.xlsx` (onglets `ca`, `charge`, `backlog`, `commandes`, `ca_deal`, `anciennete`).

- **CA N-1** récupéré automatiquement depuis `ca_2025.xlsx`.
- **Ancienneté** : seules les épics avec du CA déclaré.
- **Backlog début de mois** : photo figée au 1er run du mois (colonnes `backlog_debut_mois_*`),
  jamais écrasée → permet de suivre la tendance mensuelle.

---

## Orchestration temporelle

`orchestration.py` détecte le 1er jour du mois et pilote le workflow :

| Quand | Actions |
|-------|---------|
| **Chaque jour** | collecte courant → contrôle CA → export → historique → dashboard |
| **1er du mois (en plus)** | fige le mois précédent (mode `precedent`) + capture le prévisionnel |

Le 1er du mois, l'historique reçoit **deux lignes** : le mois précédent figé, et le mois courant créé.

---

## Le contrôle de cohérence CA

`controle_coherence_ca.py` repère deux bugs Jira fréquents :
1. **Dépassement** : CA reconnu > montant commande
2. **Écart d'avancement** : CA reconnu ≠ (% avancement × montant commande), tolérance 1€

Rapport affiché dans le **résumé du run** (onglet Actions). N'échoue jamais le pipeline.
Lançable aussi en local : `python controle_coherence_ca.py`.

---

## Fichiers de configuration

| Fichier | Contenu | Publié ? |
|---------|---------|----------|
| `absences_2026.xlsx` | Congés / maladie, un onglet par mois | ❌ sensible (Secret) |
| `objectifs.xlsx` | Objectifs CA (Budget / LE1 / LE2) | ✅ |
| `ca_2025.xlsx` | CA réalisé 2025 (référence N-1) | ✅ |
| `_historique.xlsx` | Cumul mensuel (alimenté par le pipeline) | ✅ |
| `previsionnel_2026.xlsx` | Prévisionnel pondéré mensuel | ✅ (non sensible) |

---

## ⚠️ Données sensibles — à ne JAMAIS committer

- **`absences_2026.xlsx`** — congés / maladie nominatifs (données de santé)
- **`jira_cache.pkl`** — tous les worklogs de l'équipe

Protection : `.gitignore` + le fichier absences entre par le Secret `ABSENCES_B64`
(reconstitué le temps du run puis supprimé) ; le cache est généré puis supprimé avant
publication. L'export Excel publié ne contient **aucune** donnée RH nominative
(onglet Capacité = totaux d'équipe uniquement).

---

## Déploiement GitHub Actions

### Secrets requis

*Settings → Secrets and variables → Actions*

| Secret | Valeur |
|--------|--------|
| `JIRA_URL` | `https://sogelink.atlassian.net` |
| `JIRA_EMAIL` | `mathilde.panier@sogelink.com` |
| `JIRA_API_TOKEN` | token Jira |
| `TEMPO_TOKEN` | token Tempo |
| `ABSENCES_B64` | `absences_2026.xlsx` encodé en base64 |

### Encoder les absences (PowerShell)

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("absences_2026.xlsx")) | Set-Clipboard
```

Coller dans le Secret `ABSENCES_B64`. **À refaire chaque mois.** (Dans CMD, taper `powershell` d'abord.)

### Lancer

- **Manuel** : Actions → Pipeline mensuel → Run workflow
- **Automatique** : cron quotidien 06:00 UTC (le 1er du mois déclenche les étapes spéciales)

---

## Lancer en local

```bash
pip install -r requirements.txt

# PowerShell
$env:JIRA_URL = "https://sogelink.atlassian.net"
$env:JIRA_EMAIL = "mathilde.panier@sogelink.com"
$env:JIRA_TOKEN = "ton_token_jira"
$env:TEMPO_TOKEN = "ton_token_tempo"

python collect_jira_worklogs.py --mois courant --refresh
python export_mensuel.py
python generate_dashboard_mensuel.py
python generate_previsionnel.py        # optionnel : prévisionnel du mois
```

---

## Équipe suivie

Bérénice Bossard · Marine Masingarbe · Duncan Hamelin · Maxime Pontonnier · Quentin Bordillon · Fabien Reutenauer · Flavie Bardin (départ, 0 %) · Rémy Vincent

---

## Chantiers à venir

- Comparaison prévu / réalisé (taux de report du prévisionnel)
- Enrichissement du dashboard (détail productif/gratuit, heures à date, capacité dynamique)
- Suivi continu des temps par projet (`Temps_Clockwork`)
- Le fichier de reconnaissance CA détaillé reste généré **en local** (non publié, sensible)
