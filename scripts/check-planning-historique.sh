#!/bin/sh
# Pre-commit hook : refuse un `PLANNING_TECH` dont la fenetre publiee COMMENCE
# plus tard qu'elle ne le doit -- c'est-a-dire qui efface des semaines passees.
#
# POURQUOI IL EXISTE (17/09/2026)
#
# `planning-rte.ps1` publiait la fenetre [lundi courant -> +6 semaines]. Chaque
# lundi, la semaine qui venait de finir sortait donc de ce qui est publie. Rien
# ne la rattrapait : `TECH_RANGES`, la recopie a la main, a cesse d'etre tenu le
# jour ou ce script est entre en service (03/09/2026). Resultat mesure dans
# l'appli le 17/09/2026 : les semaines 32 a 36 montraient 10 a 11 techniciens,
# et la semaine 37 n'en montrait plus que 3.
#
# Patrice : « les techniciens n'ont toujours pas la semaine 37 [...] il est
# important qu'ils se rappellent ou ils sont alles [...] fais en sorte qu'elle ne
# s'efface plus, elle et les suivantes. »
#
# CE QUI EST PROTEGE ICI, ET QUI NE SE VOIT PAS AUTREMENT. Le trou ne casse rien,
# n'affiche aucune erreur et ne se remarque qu'en allant regarder une semaine
# passee -- ce que personne ne fait souvent. Il a vecu deux semaines avant d'etre
# vu, et seulement parce que Patrice est alle sur sa version atelier. Une fenetre
# glissante de treize semaines aurait refait exactement le meme trou, en
# decembre : c'est pour ca que la date de depart est FIXE et qu'un controle la
# garde.
#
# CE QU'IL NE PROTEGE PAS : la fin de fenetre, qui elle DOIT glisser (six
# semaines devant, plafond applique par l'appli).
#
# Install : cf. scripts/pre-commit.sh

# Le lundi de la premiere semaine que la recopie manuelle ne couvre pas.
# Avant lui, `datesPour` lit `TECH_RANGES`, complet de la S02 a la S36.
ANCRE="2026-09-07"

FILE="index.html"

if ! git diff --cached --name-only | grep -qx "$FILE"; then
  exit 0
fi

BLOB=$(git rev-parse ":$FILE" 2>/dev/null) || exit 0
LIGNE=$(git cat-file blob "$BLOB" 2>/dev/null | grep -m1 '^const PLANNING_TECH = ')

# Pas de bloc publie (fichier en cours de refonte) : ce controle n'a rien a dire.
[ -n "$LIGNE" ] || exit 0

DU=$(printf '%s' "$LIGNE" | sed -n 's/.*"du":"\([0-9-]*\)".*/\1/p')

if [ -z "$DU" ]; then
  echo ""
  echo "ERREUR pre-commit : PLANNING_TECH ne porte pas de date de debut (\"du\")."
  echo "Sans elle, impossible de verifier que les semaines passees sont toujours"
  echo "publiees -- et c'est exactement ce qui s'est perdu en silence en septembre."
  echo ""
  exit 1
fi

# Les dates ISO se comparent comme du texte.
if [ "$DU" \> "$ANCRE" ]; then
  echo ""
  echo "ERREUR pre-commit : PLANNING_TECH commence le $DU, alors que l'ancre est $ANCRE."
  echo ""
  echo "Ce commit EFFACERAIT des semaines passees chez le technicien : tout ce qui"
  echo "est avant $DU ne serait plus visible, et la recopie manuelle"
  echo "(TECH_RANGES) ne le rattrape plus depuis le 03/09/2026."
  echo ""
  echo "-> Verifie \$DEBUT_HISTORIQUE dans suivi-chantiers/scripts/planning-rte.ps1 :"
  echo "   la fenetre doit partir d'une DATE FIXE, jamais du lundi courant."
  echo "   Si l'ancre doit vraiment bouger, c'est une decision de Patrice : change"
  echo "   les deux endroits ensemble, et dis-lui ce que les techniciens perdent."
  echo ""
  exit 1
fi

echo "  (historique) OK : le planning publie remonte au $DU, aucune semaine passee effacee."
exit 0
