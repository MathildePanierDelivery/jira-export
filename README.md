# Jira Export — Professional Services France

Tableau de bord automatisé pour le suivi de production, CA, backlog et temps de l'équipe PS France.
Les données sont extraites depuis Jira chaque matin et publiées sur GitHub Pages.

---

## Structure du dépôt

```
jira-export/
│
├── .github/workflows/          # Automatisations GitHub Actions
│   ├── workflow_quotidien.yml  # Export quotidien (lun-ven, 7h UTC)
│   ├── workflow_mensuel.yml    # Export mensuel (1er du mois)
│   └── workflow_tps_collab.yml # Temps collabs (à la demande)
│
├── archives/                   # Fichiers de référence mensuels
│   ├── 2026-06_previsionnel_Juin.xlsx
│   ├── 2026-06_final_Juin.xlsx
│   └── ...
│
├── Scripts d'export (Jira → Excel)
│   ├── export_production.py            # Suivi de production mensuel
│   ├── export_backlog.py               # Backlog par bon de commande
│   ├── export_commandes.py             # Commandes reçues depuis jan.
│   ├── export_worklog_historique2026.py # Temps Clockwork par projet/BDC
│   ├── export_temps_collab.py          # Temps par collaborateur
│   └── update_reconnaissance_ca_historique2026.py # Tableau CA reconnu
│
├── Scripts de génération (Excel → HTML)
│   ├── generate_rapport_global.py      # Dashboard global interactif
│   └── generate_suivi_objectifs.py     # Dashboard suivi des objectifs
│
├── Fichiers Excel de référence (entrées)
│   ├── absences.xlsx                   # Absences mensuelles (saisie manuelle)
│   ├── objectifs.xlsx                  # Objectifs annuels
│   ├── ca_historique.xlsx              # Historique CA
│   ├── performance_historique.xlsx     # Historique performance
│   └── charge_historique.xlsx          # Historique charge
│
├── Fichiers Excel générés (sorties quotidiennes)
│   ├── jira_export_YYYY-MM-DD.xlsx     # Export production du jour
│   ├── Backlog_projets_by_CA_*.xlsx    # Backlog du jour
│   └── export_ca_YYYY_*.xlsx           # Commandes du jour
│
├── Fichiers HTML (GitHub Pages)
│   ├── index.html                      # Portail de navigation
│   ├── rapport_global.html             # Dashboard global
│   └── objectifs.html                  # Suivi des objectifs
│
└── requirements.txt                    # Dépendances Python
```

---

## Workflows

### `workflow_quotidien.yml` — Export quotidien
**Déclenchement** : lundi au vendredi à 7h UTC (8h/9h heure française)
**Déclenchement manuel** : onglet Actions → Export quotidien → Run workflow

**Scripts exécutés dans l'ordre :**
1. `export_production.py` — extraction Jira (mode `mois_courant` par défaut)
2. `generate_rapport_global.py` — génération dashboard global
3. `generate_suivi_objectifs.py` — génération dashboard objectifs
4. `export_backlog.py` — extraction backlog
5. `export_commandes.py` — extraction commandes

**Logique d'archivage automatique :**
- **Premier vendredi du mois** → copie dans `archives/YYYY-MM_previsionnel_Mois.xlsx`
- **1er jour ouvré du mois** → export en mode `mois_precedent` + copie dans `archives/YYYY-MM_final_Mois.xlsx`

---

### `workflow_mensuel.yml` — Export mensuel
**Déclenchement** : le 1er de chaque mois à 8h UTC
**Déclenchement manuel** : onglet Actions → Export mensuel → Run workflow

**Scripts exécutés :**
1. `update_reconnaissance_ca_historique2026.py` — mise à jour tableau CA reconnu
2. `export_worklog_historique2026.py` — extraction temps Clockwork par projet

**Archivage automatique :** copie dans `archives/` avec suffixe `_Mois.xlsx`

---

### `workflow_tps_collab.yml` — Temps collaborateurs
**Déclenchement** : manuel uniquement
**Onglet Actions → Export temps collaborateurs → Run workflow**

Au lancement, GitHub propose de choisir la période :
- `mois_courant` — mois en cours
- `mois_precedent` — mois précédent (défaut)

**Script exécuté :** `export_temps_collab.py`

---

## Secrets GitHub requis

À configurer dans **Settings → Secrets and variables → Actions** :

| Secret | Description |
|---|---|
| `JIRA_URL` | `https://sogelink.atlassian.net` |
| `JIRA_EMAIL` | Email du compte Jira |
| `JIRA_API_TOKEN` | Token API Jira (généré sur id.atlassian.com) |

---

## Fichiers d'entrée à maintenir manuellement

| Fichier | Fréquence | Comment |
|---|---|---|
| `absences.xlsx` | Mensuelle | Via le formulaire HTML sur le portail |
| `objectifs.xlsx` | Annuelle | Mise à jour des objectifs en début d'année |
| `ca_historique.xlsx` | Automatique | Alimenté par `export_production.py` |
| `performance_historique.xlsx` | Automatique | Alimenté par `export_production.py` |
| `charge_historique.xlsx` | Automatique | Alimenté par `export_production.py` |

---

## GitHub Pages

URL du portail : `https://mathildepanierdelivery.github.io/jira-export/`

| Page | Description | Accès |
|---|---|---|
| `index.html` | Portail de navigation | Tous |
| `rapport_global.html` | Dashboard global (CA, charge, backlog) | Équipe |
| `objectifs.html` | Suivi CA annuel + Rework | Équipe |
| `backlog.html` | Backlog BDC + commandes + RAR | Équipe |
| `charge.html` | Temps par collaborateur | Management |

---

## Dépendances Python

```
jira
pandas
openpyxl
holidays
python-dateutil
requests
```
