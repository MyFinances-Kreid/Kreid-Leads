#!/bin/bash
# Bloque un commit qui contiendrait un numéro de téléphone ou un fichier de données (RGPD).
# Installation : ln -s ../../agent/check_rgpd.sh .git/hooks/pre-commit
FILES=$(git diff --cached --name-only --diff-filter=ACM)
[ -z "$FILES" ] && exit 0
BAD=$(echo "$FILES" | grep -Ei '\.(xlsx|xls|pdf)$|crm_extract|kreid_seed\.csv|^config/(ajustements|rapprochements_manuels)\.csv$' || true)
[ -n "$BAD" ] && { echo "REFUS : fichiers de données personnelles : $BAD"; exit 1; }
HIT=$(git diff --cached -U0 -- $FILES | grep '^+' | grep -vE '^\+\+\+' | grep -nE '(\+33|0)[1-9]([ .-]?[0-9]{2}){4}|\b[67][0-9]{8}\b' | grep -v '600000000\|0600000000' || true)
[ -n "$HIT" ] && { echo "REFUS : numéro de téléphone détecté :"; echo "$HIT"; exit 1; }
exit 0
