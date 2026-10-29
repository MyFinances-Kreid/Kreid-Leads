# Évolution de l'agent

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
