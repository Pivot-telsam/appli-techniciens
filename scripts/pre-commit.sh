#!/bin/sh
# Hook pre-commit : lance toutes les verifications scripts/check-*.sh de ce
# depot, dans l'ordre alphabetique. La premiere qui echoue arrete le commit.
#
# Ce fichier existe parce qu'un depot a maintenant PLUSIEURS controles et qu'il
# n'y a qu'un seul hook pre-commit possible : installer directement un des
# scripts check-*.sh comme hook desactiverait silencieusement les autres.
# Un nouveau check-*.sh depose dans scripts/ est pris en compte sans rien
# modifier ici.
#
# Install (once per clone -- hooks are not versioned by git itself):
#   cp scripts/pre-commit.sh .git/hooks/pre-commit
#   chmod +x .git/hooks/pre-commit

RACINE=$(git rev-parse --show-toplevel)

for CHECK in "$RACINE"/scripts/check-*.sh; do
  [ -f "$CHECK" ] || continue
  sh "$CHECK" || exit 1
done

exit 0
