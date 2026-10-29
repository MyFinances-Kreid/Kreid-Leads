#!/bin/bash
# Régénère « Leads data - Rentabilite.xlsx » : build -> recalcul LibreOffice -> TCD
# Usage : bash run.sh <fichier_source.xlsx> <crm_extract.txt|-> <dossier_bordereaux> <fichier_sortie.xlsx> [<dossier Imports>]
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
SRC="$1"; CRM="$2"; BX="$3"; OUT="$4"; IMP="${5:-}"; TMP="$(mktemp -d)"
# 5e argument optionnel : dossier Imports -> intègre les exports fournisseurs, choisit l'extraction CRM la plus récente, journalise
if [ -n "$IMP" ]; then
  mkdir -p "$TMP/imp"
  CRM="$(python3 "$HERE/import_exports.py" "$SRC" "$IMP" "$HERE" "$TMP/imported.xlsx" "$TMP/imp")"
  # journal complet = journal publié + événements de ce passage (publié à la fin seulement si tout réussit)
  if [ -f "$HERE/journal_imports.csv" ]; then cp "$HERE/journal_imports.csv" "$TMP/journal_full.csv"
  else echo "horodatage;source;fichier;evenement;cle;nom;prenom;detail" > "$TMP/journal_full.csv"; fi
  cat "$TMP/imp/journal_ajout.csv" >> "$TMP/journal_full.csv"
  SRC="$TMP/imported.xlsx"
  echo "Extraction CRM utilisée : $CRM"
fi
JOURNAL_FOR_BUILD="${TMP}/journal_full.csv" python3 "$HERE/build.py" "$SRC" "$CRM" "$BX" "$HERE/ajustements.csv" "$TMP/stage.xlsx"
python3 "$HERE/tools/recalc.py" "$TMP/stage.xlsx" 300 | tee "$TMP/recalc.json"
grep -q '"total_errors": 0' "$TMP/recalc.json" || { echo "ERREUR : formules en erreur, fichier non publié"; exit 1; }
python3 "$HERE/inject_pivots.py" "$TMP/stage.xlsx" "$OUT"
python3 - "$OUT" <<'PY'
import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
B = wb['Base leads']; ctl = [B.cell(r, 31).value for r in range(4, 11)]
print('CONTROLE leads W/F/Q/total/base/relances/Kreid :', ctl)
assert ctl[3] == ctl[4], 'Total leads != lignes base'
PY
if [ -n "$IMP" ]; then   # passage réussi : on publie le journal et l'état des imports
  cp "$TMP/journal_full.csv" "$HERE/journal_imports.csv"
  mkdir -p "$HERE/etat_imports" && cp "$TMP/imp/etat_imports/"* "$HERE/etat_imports/"
fi
echo "OK -> $OUT"
