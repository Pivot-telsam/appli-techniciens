#!/bin/sh
# Les deux depots doivent dire la MEME chose sur PDP / PGO / IST / materiel.
#
# Pose le 14/09/2026 apres qu'un technicien a appele Patrice : son PGO
# s'affichait ROUGE dans l'appli alors que les dates etaient bonnes et qu'il
# avait le bon document dans App Tech. Le suivi disait "ok", l'appli "bad",
# sur 26-030 Rion des Landes, depuis plusieurs jours, et rien ne comparait
# les deux fichiers.
#
# Le detail du raisonnement est dans scripts/comparer-statuts.ps1.
# Pris automatiquement par scripts/pre-commit.sh (motif check-*.sh).

ici=$(cd "$(dirname "$0")" && pwd)
ps1="$ici/comparer-statuts.ps1"

if [ ! -f "$ps1" ]; then
  echo "  (statuts) comparer-statuts.ps1 introuvable : controle non joue."
  exit 0
fi

powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$(cygpath -w "$ps1" 2>/dev/null || echo "$ps1")"
exit $?
