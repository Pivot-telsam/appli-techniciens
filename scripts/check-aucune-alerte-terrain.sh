#!/bin/sh
# Pre-commit hook : refuse une fiche de l'appli technicien qui porte des ALERTES.
#
# POURQUOI IL EXISTE. Le 01/10/2026, Patrice a ouvert la carte de Lisieux - Vallee 1
# dans App Tech : PDP ind.4 et PGO ind.6 en vert, et juste dessous un encadre rouge
# « AUCUNE COUVERTURE A CE JOUR ... reclamer un nouveau PdP et un nouveau PGO a RTE
# avant tout depart », ecrit le 11/09 a l'epoque de l'indice 3 et jamais leve quand
# les nouveaux documents sont arrives le 18/09. Sa decision : « Les alertes doivent
# arriver dans le suivi, pas dans l'application des techniciens. »
#
# Le champ `alertes` d'une fiche est celui du BUREAU : il vit dans le suivi (colonne
# « Ce qui manque », dossier de l'affaire). L'appli ne l'affiche plus, et ce controle
# empeche qu'une fiche recopiee du suivi en ramene : une donnee qu'on n'affiche pas
# n'a rien a faire dans un depot public.
#
# Consignation et NIP restent : ce sont des faits du chantier, pas des alertes.
#
# Install : cf. scripts/pre-commit.sh

RACINE=$(git rev-parse --show-toplevel)
[ "$(basename "$RACINE")" = "appli-techniciens" ] || exit 0

git diff --cached --name-only --diff-filter=ACM | grep -qx 'index.html' || exit 0

N=$(git show :index.html | grep '^const SEED_DATA = ' | grep -o '"alertes":\[[^]]' | wc -l)
if [ "$N" -gt 0 ]; then
  echo "REFUSE : $N fiche(s) de SEED_DATA portent des alertes dans index.html."
  echo "Les alertes sont pour le bureau : elles vont dans le suivi, jamais dans App Tech."
  echo "Mettre \"alertes\":[] sur ces fiches (decision de Patrice, 01/10/2026)."
  exit 1
fi

if git show :index.html | grep -q 'c\.alertes||\[\])\.forEach'; then
  echo "REFUSE : index.html dessine de nouveau c.alertes sur les cartes chantier."
  echo "Les alertes sont pour le bureau : elles vont dans le suivi, jamais dans App Tech."
  exit 1
fi

exit 0
