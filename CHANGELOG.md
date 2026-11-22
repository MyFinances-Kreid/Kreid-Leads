# Évolution de l'agent

## v1.4.1 — 22/11/2026
- Couplage avec Kreid Pilotage : le classeur alimente la table `lead_entrant` du dashboard (connecteur `leads/charger_leads.py` du dépôt Kreid-Pilotage, branche `feat/couplage-leads`). Structure des onglets fournisseurs figée (documentée dans FONCTIONNEMENT.md).

## v1.4 — 18/11/2026
- **Onglet « KPI Distance »** (`kpi.py`) : synthèse (taux lead → accord / payé vs cible 6 %, ROI, coût par dossier payé), entonnoir par source, délais médians entre étapes, tableau par conseiller, cohortes par mois d'achat (accords à 30/60/90 j), indicateurs back office (dépôt ≤ 2 j, dossiers non déposés, accords non payés, leads absents du CRM).
- **Favori CRM enrichi** : montant simulé, date de paiement, CP, type de fiche (format 15 champs, compatible avec les anciens).
- Exports Weendeal / Finanzen du 18/11 intégrés (+39 / +37 leads), bordereau MY2026019.

## v1.3.1 — 18/11/2026
- Conseillers Distance arrivés le 02/11 (Pasqualini, Grelin, Deleplanque) reconnus ; fiches des personnes sorties des effectifs affichées « À réattribuer (ex-…) » dans la base et les Relances.
- Bordereau MY2026019 (04/11) intégré : 5 dossiers Finanzen rattachés.

## v1.3 — 29/10/2026
- **Intégration des exports fournisseurs** (`import_exports.py`) : dernier fichier de `Imports/Weendeal` et `Imports/Finanzen`, ajout en bas des onglets des leads absents (clé e-mail / téléphone), jamais de modification d'une ligne existante ; téléphone Weendeal repris du CRM si l'export n'en a pas.
- **Traçabilité** : `journal_imports.csv` (fichier lu + empreinte, ajouts, statuts portail, leads disparus, nouvelles fiches et étapes CRM franchies, fichiers inchangés, synthèses) + onglet « Journal imports » ; état de comparaison dans `etat_imports/`. Publiés seulement si tout le passage réussit.
- `run.sh` : 5ᵉ argument = dossier Imports (choix automatique de l'extraction CRM la plus récente).

## v1.2 — 29/10/2026
- **L'agent n'ouvre plus jamais l'extranet du CRM.** Lecture par un favori Chrome « Lire le CRM Kreid », lancé à la main par Corentin ou Yohan depuis leur session (lecture seule, 1 page / 0,6 s), qui télécharge l'extraction ; dépôt dans `Suivi leads/Imports/CRM`, reprise automatique aux passages de 9h15 et 17h45.
- Extraction complète filtrée par `build_base.py` (seules les fiches des leads suivis entrent dans le classeur).
- Préparation du partage avec Corentin : classeur en lecture seule, dossier `Imports/` pour ses dépôts (CRM, exports Weendeal et Finanzen).
- Qwamplify arrêté (onglet conservé pour l'historique).

## v1.1 — 28-29/10/2026
- **MAF** : colonne « MAF financé » (montant MAF du bordereau Budgeo rattaché au lead), champ dans le TCD fournisseur, indicateurs « MAF financé » et « MAF par € de lead ».
- **Kreid.fr, 4ᵉ source à coût 0** : onglet de saisie créé (base reconstituée depuis le formulaire du site + SUIVI-DISTANCE), ajout quotidien des nouvelles demandes depuis Outlook ; exclue du « Total payants ».
- Alerte « Absent du CRM » pour Kreid.fr activée seulement quand l'extraction CRM inclut leurs téléphones.
- Garde-fou : CA et MAF à 0 si le lead n'a pas de téléphone.
- **Lecture CRM suspendue par défaut** (interrupteur `_agent/crm_auto.txt` = OUI) et ralentie à une page / 0,6 s, suite à la décision du 10/10/2026.
- Données de départ Kreid.fr sorties du code (`kreid_seed.csv`, hors Git).

## v1.0 — 28/10/2026
- Onglets générés : Leads (4 TCD natifs, actualisés à l'ouverture), Relances (sans origine des contacts), CRM, Bordereaux Budgeo, Base leads, Méthode & sources.
- CA réel = bordereaux Budgeo (règle Yohan), ajustements manuels (geste commercial), rapprochement automatique bordereau → lead (nom + accord CRM ≤ 120 j) + exceptions manuelles.
- Extraction CRM en lecture seule dans Chrome (téléphones des leads, numéros multiples gérés).
- 3 tâches planifiées (9h15, rattrapage 9h30, 17h45), sans saisie de mot de passe.
- Correctifs Excel : TCD en XML natif ; pied de page < 255 caractères (erreur « réparation » à l'ouverture).

## v0 — avant le 28/10/2026
- Classeur construit à la main (ChatGPT) : onglets Weendeal et Finanzen, saisie manuelle quotidienne dans SUIVI-DISTANCE par Corentin.
