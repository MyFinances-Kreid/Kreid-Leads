# Suivi leads — Pôle Distance Kreid · Agent de mise à jour

Dossier : `OneDrive-MyFinances/Pôle distance/Suivi leads/_agent`
Fichier piloté : `../Leads data - Rentabilite.xlsx`
Dernière mise à jour de cette documentation : 28/10/2026

## Ce que fait l'agent (chaque jour ouvré)
1. **Ajoute les nouveaux leads Kreid.fr** : mails « Demande d'informations depuis le site internet Kreid » (expéditeur WordPress) reçus depuis le dernier passage → nouvelles lignes en bas de l'onglet **Kreid.fr** (date, nom, prénom, téléphone, canal, prix 0). Doublon de téléphone = ignoré.
2. **Prend l'extraction CRM la plus récente** (`crm_extract_AAAA-MM-JJ.txt`) déposée dans `Suivi leads/Imports/CRM` par Corentin ou Yohan avec le favori Chrome « Lire le CRM Kreid » (voir `Installer le favori CRM.html`). L'agent n'ouvre jamais l'extranet lui-même (décision du 10/10/2026). Lecture humaine, lecture seule, une page / 0,6 s ; le fichier contient toutes les fiches, `build_base.py` ne garde que celles des leads suivis.
3. **Lit les bordereaux Budgeo** (PDF) de `Desktop/Pro/Administratif/Commission/Perçues/2026/Budgéo`.
4. **Régénère le fichier** avec `run.sh` : `build.py` (onglets CRM, Bordereaux Budgeo, Base leads, Relances, Leads, Méthode) → recalcul LibreOffice (`tools/recalc.py`, 0 erreur exigée) → TCD (`inject_pivots.py`) → contrôles (total leads = lignes base).
5. **Journalise** une ligne d'indicateurs agrégés dans `journal_runs.csv` (aucune donnée nominative) : c'est la courbe d'évolution publiée sur GitHub.
6. **Dépose** le fichier dans `Suivi leads/` (écrase la version précédente).

Les onglets **Weendeal et Finanzen** sont alimentés par `import_exports.py` à partir des exports déposés dans `Imports/` : l'agent ajoute en bas les leads absents (clé e-mail pour Weendeal, téléphone pour Finanzen), sans jamais modifier une ligne existante. Qwamplify n'est plus touché. L'onglet **Kreid.fr** est créé une fois (base reconstituée le 28/10/2026 : 36 demandes du formulaire + 2 leads « KREID » de SUIVI-DISTANCE, `kreid_seed.csv`), puis l'agent y ajoute seulement des lignes.

## Règles de gestion (décisions)
| Sujet | Règle | Décidé |
|---|---|---|
| CA réel | Montant « Paiement » des bordereaux Budgeo, jusqu'au passage MNE (correspond à la grille de rémunération Budgeo / banques) | Yohan, 28/10/2026 |
| Ajustements | Gestes commerciaux et reprises saisis dans `ajustements.csv` (montant négatif = CA rendu) | Yohan, 28/10/2026 |
| Rapprochement bordereau → lead | Même nom + accord CRM du lead (téléphone) entre la date d'achat et la date du bordereau (≤ 120 j). Exceptions dans `rapprochements_manuels.csv` | 28/10/2026 |
| CA en attente | CA saisi à la main dans l'onglet fournisseur pour un accord pas encore présent dans un bordereau | 28/10/2026 |
| Prix Qwamplify | 20 € HT = 24 € TTC (TVA non récupérable, à confirmer expert-comptable) | Yohan, 28/10/2026 |
| MAF | Montant « MAF » du bordereau Budgeo rattaché au lead (colonne AC de Base leads, TCD fournisseur, indicateur MAF par € de lead) | Yohan, 28/10/2026 |
| Kreid.fr | 4ᵉ source à coût 0, exclue du « Total payants ». Alerte « Absent du CRM » activée automatiquement dès qu'une extraction CRM inclut leurs téléphones | 28/10/2026 |
| Période | Cohorte = mois d'achat du lead | 28/10/2026 |
| Relances | Absent du CRM (Weendeal/Finanzen) · contact sans simulation > 2 j · simulation sans retour > 10 j · dossier reçu non déposé > 7 j · dépôt sans décision > 15 j · accord non payé. Leads < 120 j | 28/10/2026 |
| Diffusion | Onglet Relances sans origine des contacts ; fichier = données personnelles, usage interne (RGPD) | Yohan |

## Fichiers
- `build.py`, `build_base.py` : construction du classeur (openpyxl).
- `inject_pivots.py` : ajout des 4 TCD (cache vide + actualisation à l'ouverture, format natif Excel).
- `make_crm_bookmarklet.py` : génère `Installer le favori CRM.html` (favori de lecture du CRM, lancé à la main).
- `make_crm_snippet.py` : ancienne extraction pilotée par l'agent (inutilisée depuis le 29/10/2026).
- `run.sh` : chaîne complète. Usage : `bash run.sh <source.xlsx> <crm_extract.txt> <dossier_bordereaux> <sortie.xlsx>`
- `kreid_seed.csv` : base initiale des leads du site (utilisée seulement si l'onglet Kreid.fr n'existe pas).
- `ajustements.csv`, `rapprochements_manuels.csv` : corrections manuelles (à éditer avec Claude ou à la main, séparateur `;`).
- `crm_extract_AAAA-MM-JJ.txt` : dernière extraction CRM (historique).
- `tools/` : recalcul LibreOffice.

## Limites connues
- La session de l'extranet expire : si Yohan n'est pas connecté dans Chrome au moment de la tâche, l'agent s'arrête et le signale (il ne se connecte jamais à sa place). Le fichier garde alors la dernière extraction.
- Le Mac doit être allumé, Chrome ouvert, l'application Claude lancée.
- Les fiches CRM « faux contact / ne plus relancer » ne sont pas visibles dans la liste : elles peuvent apparaître en relance.
- Excel affiche les TCD vides dans un aperçu (Finder, Outlook) : ils se remplissent à l'ouverture dans Excel.
- Nouveau fournisseur : à ajouter dans `build_base.py`, `make_crm_snippet.py` et `inject_pivots.py` (avec Claude).

## Dépôt GitHub
Code, décisions et journal agrégé versionnés dans le dépôt privé MyFinances-Kreid/Kreid-Leads (sans aucune donnée personnelle : classeurs, bordereaux, extractions CRM, CSV de corrections et `kreid_seed.csv` restent ici).

## Partage avec Corentin (29/10/2026)
- Classeur partagé **en lecture seule** (il est réécrit à chaque passage de l'agent). Corentin voit tout le fichier, rentabilité comprise (décision Yohan).
- Corentin dépose ses fichiers dans `Imports/` (droit de modification sur ce dossier seulement) : `Imports/CRM` (favori), `Imports/Weendeal` et `Imports/Finanzen` (exports des portails fournisseurs, intégration automatique à venir).
- Qwamplify : achat arrêté pour le moment, onglet conservé pour l'historique.

## Imports et traçabilité (29/10/2026)
- `run.sh <classeur> - <bordereaux> <sortie> <dossier Imports>` : `import_exports.py` lit le dernier fichier de chaque sous-dossier d'`Imports/`, complète les onglets fournisseurs, choisit l'extraction CRM la plus récente, puis la génération habituelle tourne.
- **Journal** : `_agent/journal_imports.csv` (une ligne par événement : fichier lu + empreinte, AJOUT_ONGLET, TELEPHONE_MANQUANT, STATUT_PORTAIL, DISPARU_EXPORT, NOUVELLE_FICHE_LEAD, ETAPE_CRM, FICHE_DISPARUE, FICHIER_INCHANGE, SYNTHESE). Recopié dans l'onglet « Journal imports » du classeur (le plus récent en haut).
- **État de comparaison** : `_agent/etat_imports/` (dernier export de chaque source + dernière extraction CRM). Journal et état ne sont publiés que si tout le passage réussit.
- **Weendeal** : exporter en **« CSV Codé »** (toutes les pages) pour avoir le téléphone. Avec le CSV simple, le téléphone est repris du CRM (même nom + prénom), sinon TELEPHONE_MANQUANT.
- Données personnelles : journal et état restent dans OneDrive, jamais dans Git.
