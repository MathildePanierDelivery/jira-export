# Process de bascule mensuelle

Ce document décrit, étape par étape, comment clôturer un mois et démarrer le suivant.
À suivre **dans l'ordre**, sans sauter d'étape. En cas de doute, s'arrêter et demander.

> **Pourquoi c'est important :** le figement du mois écoulé lit des données dans Jira.
> Si on remet Jira à zéro **avant** d'avoir figé, les bonnes données sont perdues.
> L'ordre des étapes protège contre ça.

---

## Vue d'ensemble

```
1er du mois  →  🔒 Le verrou se pose TOUT SEUL (mois courant en pause)
                 ↓
            Ajustements dans Jira (mois écoulé)
                 ↓
            ▶️  Action "figer_precedent"  (fige le mois écoulé dans l'historique)
                 ↓
            Vérifier que le figement est bon
                 ↓
            Remise à zéro dans Jira (CA du mois → 0, décalage prévisionnels)
                 ↓
            ▶️  Action "terminer_bascule"  (lève le verrou, le mois courant reprend)
                 ↓
            🔓 Terminé — le suivi du nouveau mois redémarre automatiquement
```

Le bandeau **« Bascule mensuelle en cours »** est visible sur la page d'accueil
tant que le verrou est posé. Il disparaît une fois la bascule terminée.

---

## Étape par étape

### 0. Le 1er du mois — rien à faire (automatique)

Le 1er du mois, au premier passage automatique, le **verrou se pose tout seul**.
À partir de là :
- le mois courant est **en pause** (plus aucune mise à jour des données du mois en cours) ;
- le bandeau apparaît sur la page d'accueil.

Tu peux prendre le temps qu'il faut pour la suite (1 jour, ou plusieurs).

### 1. Faire les ajustements dans Jira (mois écoulé)

Corriger dans Jira tout ce qui doit l'être sur le **mois qui vient de se terminer**
(CA, saisies, etc.), jusqu'à ce que les chiffres soient définitifs.

⚠️ **Ne pas encore faire la remise à zéro.** On fige d'abord.

### 1 bis. Mettre à jour les absences du mois écoulé

⚠️ **Important :** les arrêts maladie ne s'anticipent pas. Avant de figer, vérifier
que `absences_2026.xlsx` contient bien **toutes** les absences réelles du mois écoulé
(arrêts maladie imprévus, congés ajoutés en cours de mois...). Sinon la capacité
productive du mois figé sera fausse.

Si des absences ont changé, mettre à jour le Secret (voir [Mettre à jour les absences](#mettre-à-jour-les-absences) en bas de ce document).

### 2. Figer le mois précédent

Une fois les ajustements terminés **et les absences à jour** :

1. Aller sur **GitHub → onglet Actions → "Pipeline mensuel"**
2. Cliquer **"Run workflow"** (à droite)
3. Dans le menu **Action**, choisir **`figer_precedent`**
4. Cliquer **"Run workflow"** (le bouton vert)

Le pipeline collecte le mois écoulé et l'enregistre dans l'historique.
Une copie de l'export complet est aussi archivée dans **`archives_mensuelles/`**,
nommée `AAAA-MM_Mois.xlsx` (ex : `2026-06_Juin.xlsx`).

### 3. Vérifier que le figement est bon

Attendre que le run se termine (≈ 3-4 min), puis **vérifier** :
- Ouvrir le **tableau de bord mensuel**, sélectionner le mois qu'on vient de figer.
- Contrôler que le CA, la charge, les clôtures correspondent aux chiffres attendus.

✅ Si tout est bon → continuer.
❌ Si quelque chose cloche → refaire les ajustements dans Jira (étape 1),
   puis relancer **`figer_precedent`** (étape 2). On peut figer autant de fois
   que nécessaire **tant que la remise à zéro n'est pas faite.**

### 4. Remise à zéro dans Jira

**Seulement après que le figement est validé :**
- Remettre le **« CA du mois en cours »** à 0.
- Décaler les champs **prévisionnels** d'un mois.

### 4 bis. Mettre à jour les absences du nouveau mois

Préparer `absences_2026.xlsx` pour le **nouveau mois** qui démarre : congés connus,
RTT, etc. (les absences se compléteront au fil du mois, mais autant partir avec ce
qu'on sait). Mettre à jour le Secret (voir [Mettre à jour les absences](#mettre-à-jour-les-absences)).

### 5. Terminer la bascule

1. Aller sur **GitHub → Actions → "Pipeline mensuel" → "Run workflow"**
2. Dans **Action**, choisir **`terminer_bascule`**
3. Cliquer **"Run workflow"**

Le verrou est levé, le bandeau disparaît, et le **suivi du nouveau mois reprend**
automatiquement (aux passages de 7h et 12h).

---

## Garde-fous

- **Si on oublie de terminer la bascule :** le verrou se lève **tout seul au bout
  de 10 jours** (sécurité), pour éviter que le mois courant reste gelé. Mais il
  vaut toujours mieux le faire manuellement, au bon moment (après la remise à zéro).
- **Le mois courant pendant la bascule :** il est en pause, c'est normal. Les
  dashboards montrent les dernières données connues jusqu'à la fin de la bascule.
- **Refaire une étape :** tant que la remise à zéro Jira n'est pas faite, on peut
  relancer `figer_precedent` sans risque. Après la remise à zéro, **ne plus** lancer
  `figer_precedent` (cela écraserait l'historique avec des données vidées).

---

## Qui peut faire la bascule ?

Les personnes ajoutées comme **collaboratrices du dépôt GitHub** peuvent lancer les
actions (`figer_precedent`, `terminer_bascule`). En cas d'absence de la responsable,
elles peuvent assurer la bascule en suivant ce document.

---

## Mettre à jour les absences

Le fichier `absences_2026.xlsx` (congés / maladie) ne peut **pas** être mis dans le
dépôt (données sensibles). Il entre dans le pipeline via un **Secret GitHub** encodé.
Quand les absences changent, il faut ré-encoder le fichier et mettre à jour le Secret :

1. Ouvrir **PowerShell** (sous Windows ; dans CMD, taper d'abord `powershell`).
2. Aller dans le dossier qui contient le fichier :
   ```powershell
   cd "C:\chemin\vers\le\dossier"
   ```
   *(astuce : taper `cd ` puis glisser-déposer le dossier dans la fenêtre pour coller le chemin. `dir` pour vérifier que `absences_2026.xlsx` est bien là.)*
3. Encoder le fichier (le résultat est copié dans le presse-papier) :
   ```powershell
   [Convert]::ToBase64String([IO.File]::ReadAllBytes((Resolve-Path "absences_2026.xlsx"))) | Set-Clipboard
   ```
   *(le `Resolve-Path` évite une erreur « fichier introuvable » : sans lui, PowerShell
   cherche parfois le fichier dans le mauvais dossier. Vérifier d'abord avec `dir absences_2026.xlsx`.)*
4. Sur GitHub : **Settings → Secrets and variables → Actions → `ABSENCES_B64` → Update**,
   coller, sauvegarder.

Le prochain run utilisera les absences à jour.

---

## Aide-mémoire express

| Quand | Action |
|---|---|
| 1er du mois | *(rien — le verrou se pose seul)* |
| Avant de figer | Vérifier/MAJ **absences du mois écoulé** (arrêts maladie !) |
| Ajustements Jira finis | Run workflow → **`figer_precedent`** |
| Après RAZ Jira | MAJ **absences du nouveau mois** |
| Figement vérifié + RAZ faite | Run workflow → **`terminer_bascule`** |
