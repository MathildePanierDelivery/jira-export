# Pilotage de production — PS France

Pipeline automatisé **Jira + Tempo → Excel → Historique → Dashboard HTML** pour le suivi mensuel de la production (CA, charge, backlog, clôtures) des solutions Littéralis et GEODP.

Dépôt : `MathildePanierDelivery/jira-export` · Publication : GitHub Pages

---

## Architecture

```
collect_jira_worklogs.py   →  jira_cache.pkl              (cache, NON publié)
        │   Jira API (issues + clôtures) · Tempo API (worklogs)
        ▼
export_mensuel.py          →  jira_export_AAAA-MM-JJ.xlsx (8 onglets)
        │                  →  _historique.xlsx            (accumulation du mois)
        ▼
generate_dashboard_mensuel.py  →  dashboard_mensuel.html (lit _historique.xlsx)
```

---

## Les scripts

| Script | Rôle | Sorties |
|--------|------|---------|
| `collect_jira_worklogs.py` | Collecte Jira + Tempo en un seul passage | `jira_cache.pkl` |
| `export_mensuel.py` | Génère l'export Excel **et** accumule le mois dans l'historique | `jira_export_*.xlsx` |
| `maj_historique.py` | Module appelé par l'export : écrit la ligne du mois dans `_historique.xlsx` | — |
| `generate_dashboard_mensuel.py` | Génère le dashboard web depuis l'historique | `dashboard_mensuel.html` |

### Pourquoi un collecteur séparé ?

Les worklogs sont saisis via **Tempo**, pas l'API Jira native (qui renvoie un compte technique au lieu du vrai auteur). Le collecteur interroge l'API Tempo (`api.eu.tempo.io`), filtre sur l'équipe (8 collaborateurs via `accountId`) et joint à Jira pour les métadonnées. Une seule collecte alimente tout.

---

## L'export Excel (8 onglets)

1. **Tableau de bord** — synthèse : CA vs objectif, charge à date, projets clôturés / ouverts
2. **Suivi de production** — CA par épic (Hardware exclu)
3. **Suivi Hardware** — tickets Matériel GEODP clôturés dans le mois
4. **Capacité productive** — disponibilité de l'équipe (jours ouvrés − absences)
5. **Charge** — heures par solution et par type, % occupation coloré
6. **Backlog** — montant restant à reconnaître, mobilisable vs bloqué
7. **Commandes du mois** — épics créées dans le mois
8. **Projets clôturés** — épics terminées, rework mis en évidence

### Règles métier clés

- **Solution** : normalisée par mot-clé (insensible casse/accents) → LITTERALIS / GEODP
- **Épics PSC** (New delivery) : solution = LITTERALIS ; catégorie = Projet si ticket CA rattaché, sinon Commande sans prestation
- **Temps par solution** : seul le productif compte (Projet, Rework, Maintenance, Prestation offerte, Commande sans prestation) ; Support et Interne à part
- **Matériel GEODP** : ne consomme pas d'heures ; le temps se répartit sur les tickets services
- **Clôtures** : détectées via le changelog Jira (`status changed TO "Terminé"`), car `resolutiondate` est souvent vide
- **Objectifs** : onglet lu défini par `OBJECTIFS_ONGLET` dans `export_mensuel.py` (actuellement `LE1`)

---

## L'accumulation dans l'historique

À chaque exécution, `export_mensuel.py` écrit (ou écrase) la ligne du **mois courant** dans `_historique.xlsx`, sur 6 onglets : `ca`, `charge`, `backlog`, `commandes`, `ca_deal`, `anciennete`.

- Le **CA N-1** est récupéré automatiquement depuis `ca_2025.xlsx`.
- L'**ancienneté** ne retient que les épics avec du CA déclaré.
- Le dashboard lit ensuite cet historique → sélecteur multi-mois.

---

## Le dashboard HTML

`generate_dashboard_mensuel.py` lit `_historique.xlsx` et produit `dashboard_mensuel.html` :
fichier autonome, sélecteur mois / année, CA vs objectifs, CA vs commandes, répartition des heures (donuts), clôtures. Données embarquées en JSON (non sensibles).

`index.html` est la page d'accueil qui pointe vers le dashboard et les autres pages.

---

## Fichiers de configuration

| Fichier | Contenu | Mise à jour |
|---------|---------|-------------|
| `absences_2026.xlsx` | Congés / maladie, **un onglet par mois** | Manuelle (⚠️ sensible, voir plus bas) |
| `objectifs.xlsx` | Objectifs CA (onglets Budget / LE1 / LE2) | Selon validation direction |
| `ca_2025.xlsx` | CA réalisé 2025 (référence N-1) | Figé |
| `_historique.xlsx` | Cumul mensuel (alimenté par le pipeline) | Automatique |

---

## ⚠️ Données sensibles — à ne JAMAIS committer

Deux fichiers contiennent des données personnelles, exclus par `.gitignore` :

- **`absences_2026.xlsx`** — congés / maladie nominatifs (données de santé)
- **`jira_cache.pkl`** — tous les worklogs de l'équipe

Protection dans le workflow :
- `absences_2026.xlsx` entre par le **Secret** `ABSENCES_B64` (base64), reconstitué le temps du run puis supprimé.
- `jira_cache.pkl` est généré pendant le run puis supprimé avant tout commit.
- Seuls les livrables **non sensibles** sont publiés (dashboards, historique agrégé).

---

## Déploiement GitHub Actions

### Secrets à créer

*Settings → Secrets and variables → Actions → New repository secret*

| Secret | Valeur |
|--------|--------|
| `JIRA_URL` | `https://sogelink.atlassian.net` |
| `JIRA_EMAIL` | `mathilde.panier@sogelink.com` |
| `JIRA_API_TOKEN` | token Jira |
| `TEMPO_TOKEN` | token Tempo |
| `ABSENCES_B64` | `absences_2026.xlsx` encodé en base64 |

### Encoder le fichier absences (PowerShell)

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("absences_2026.xlsx")) | Set-Clipboard
```

Coller le résultat dans le Secret `ABSENCES_B64`. **À refaire chaque mois** après mise à jour des absences. (Si tu es dans CMD, tape d'abord `powershell`.)

### Lancer

- **Manuel** : *Actions* → *Pipeline mensuel* → *Run workflow* (choix `courant` / `precedent`)
- **Automatique** : selon le `cron` dans `.github/workflows/pipeline_mensuel.yml`

---

## Lancer en local

```bash
pip install jira requests pandas openpyxl holidays

# PowerShell
$env:JIRA_URL = "https://sogelink.atlassian.net"
$env:JIRA_EMAIL = "mathilde.panier@sogelink.com"
$env:JIRA_TOKEN = "ton_token_jira"
$env:TEMPO_TOKEN = "ton_token_tempo"

python collect_jira_worklogs.py --mois courant --refresh
python export_mensuel.py
python generate_dashboard_mensuel.py
```

Le token Jira est lu via `JIRA_API_TOKEN` (GitHub) **ou** `JIRA_TOKEN` (local).

---

## Équipe suivie

Bérénice Bossard · Marine Masingarbe · Duncan Hamelin · Maxime Pontonnier · Quentin Bordillon · Fabien Reutenauer · Flavie Bardin (départ, 0 %) · Rémy Vincent

---

## Chantiers à venir

- Contrôle de cohérence du CA (`Reconnaissance_CA`) avec alerte
- Snapshots figés : backlog début de mois, prévisionnel (1er vendredi)
- Suivi continu des temps par projet (`Temps_Clockwork`)
- Orchestration fine du workflow (quotidien + figement mensuel)
