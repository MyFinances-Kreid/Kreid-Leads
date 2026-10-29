#!/bin/bash
# Régénère « Leads data - Rentabilite.xlsx » : build -> recalcul LibreOffice -> TCD
# Usage : bash run.sh <fichier_source.xlsx> <crm_extract.txt> <dossier_bordereaux> <fichier_sortie.xlsx>
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
SRC="$1"; CRM="$2"; BX="$3"; OUT="$4"; TMP="$(mktemp -d)"
python3 "$HERE/build.py" "$SRC" "$CRM" "$BX" "$HERE/ajustements.csv" "$TMP/stage.xlsx"
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
echo "OK -> $OUT"
