# Kreid · Agent de suivi des leads (Pôle Distance)

Dépôt **privé** de My Finances (Kreid). Il versionne l'agent IA qui tient à jour le fichier
« Leads data - Rentabilite.xlsx » : rentabilité par source de leads (Weendeal, Finanzen, Qwamplify, Kreid.fr),
entonnoir CRM, relances du jour, CA réel et MAF issus des bordereaux Budgeo.

> **Aucune donnée personnelle ici.** Ni classeur, ni bordereau, ni extraction CRM, ni nom de client ou téléphone.
> Ces fichiers vivent uniquement dans OneDrive (`Pôle distance/Suivi leads/_agent`). Le `.gitignore` et le hook
> `agent/check_rgpd.sh` bloquent tout commit qui en contiendrait.

## Où lire quoi
| Fichier | Contenu |
|---|---|
| `CHANGELOG.md` | Évolution de l'agent, version par version |
| `docs/DECISIONS.md` | Règles de gestion validées (qui, quand) |
| `docs/FONCTIONNEMENT.md` | Procédure exécutée par les tâches planifiées (9h15, rattrapage 9h30, 17h45) |
| `journal/runs.csv` | Indicateurs agrégés de chaque passage (sans donnée nominative) : courbe de l'agent dans le temps |
| `agent/` | Scripts : `build.py`, `build_base.py`, `inject_pivots.py`, `make_crm_snippet.py`, `run.sh`, `tools/` |
| `config/*.example.csv` | Format des fichiers de corrections manuelles (les vrais restent dans OneDrive) |

## Exécution
```bash
bash agent/run.sh "<Leads data - Rentabilite.xlsx>" <crm_extract.txt> <dossier_bordereaux> <sortie.xlsx>
```
Les fichiers `ajustements.csv`, `rapprochements_manuels.csv` et `kreid_seed.csv` sont lus à côté des scripts (copie OneDrive).

## Flux de mise à jour du dépôt
La référence d'exécution reste `_agent/` dans OneDrive. Après chaque évolution, Claude recopie les scripts ici,
met à jour `CHANGELOG.md` / `docs/DECISIONS.md` / `journal/runs.csv`, et pousse sur une branche + pull request.

---
My Finances – Kreid · SAS au capital de 1 000 € · RCS Lille Métropole 911 118 255 · ORIAS 22002570 · usage interne
