#!/bin/sh
# Pre-commit hook: refuse a commit that changes a SHARED .md in one repo while the
# sibling repo's copy says something different.
#
# Reason: CLAUDE.md et JOURNAL.md sont LES MEMES documents dans les deux depots,
# dupliques pour que chaque clone soit complet. Rien ne les tenait ensemble : ils
# ont diverge en silence pendant 3 jours (constate le 31/08/2026). suivi-chantiers
# annoncait encore des liens "?tech=slugifiedname" personnels et permanents, alors
# que le portail par mot de passe du 27/08 avait supprime la lecture de ce
# parametre. Selon le depot dans lequel on travaillait, on lisait une regle
# perimee -- et on aurait pu rediffuser ces liens comme s'ils identifiaient encore
# un technicien.
#
# LE 14/09/2026 LE MEMO A ETE SCINDE EN DEUX : CLAUDE.md garde ce qui est en
# vigueur, JOURNAL.md garde le detail et l'histoire. Le controle porte donc sur
# les DEUX. Ne verifier que CLAUDE.md aurait laisse JOURNAL.md diverger en
# silence, c'est-a-dire exactement le defaut du 31/08 deplace d'un fichier.
#
# Ce controle ne se declenche que si l'un de ces fichiers fait partie du commit.
#
# Install : cf. scripts/pre-commit.sh

FICHIERS="CLAUDE.md JOURNAL.md"

RACINE=$(git rev-parse --show-toplevel)
ICI=$(basename "$RACINE")

case "$ICI" in
  appli-techniciens) AUTRE="suivi-chantiers" ;;
  suivi-chantiers)   AUTRE="appli-techniciens" ;;
  *) exit 0 ;;
esac

MIS_EN_SCENE=$(git diff --cached --name-only)
FAUTIFS=0

for FILE in $FICHIERS; do
  echo "$MIS_EN_SCENE" | grep -qx "$FILE" || continue

  AUTRE_FICHIER="$RACINE/../$AUTRE/$FILE"

  if [ ! -f "$AUTRE_FICHIER" ]; then
    echo ""
    echo "AVERTISSEMENT pre-commit : $AUTRE/$FILE est introuvable a cote de ce depot."
    echo "La coherence des deux copies n'a PAS pu etre verifiee -- le commit passe"
    echo "quand meme, mais pense a reporter la modification dans l'autre depot."
    echo "Chemin attendu : $AUTRE_FICHIER"
    echo ""
    continue
  fi

  # On compare le contenu MIS EN SCENE (l'index), pas le fichier du disque : c'est
  # lui qui partira dans le commit.
  #
  # PIEGE (rencontre en testant ce hook le 31/08/2026) : NE PAS utiliser
  # `git show :<fichier>`, qui applique la conversion de fins de ligne de
  # core.autocrlf et rend du CRLF, alors que le fichier voisin est lu tel quel sur
  # le disque. Les deux paraissent alors differer A CHAQUE LIGNE (2927 lignes
  # d'ecart pour un fichier identique) et le hook bloque tous les commits.
  # `git cat-file blob` rend les octets bruts ; le `tr -d '\r'` des deux cotes
  # acheve de rendre la comparaison insensible aux fins de ligne. Ce qui compte
  # ici est le TEXTE des consignes, pas sa representation.
  TMP_ICI=$(mktemp)
  TMP_AUTRE=$(mktemp)
  git cat-file blob "$(git rev-parse ":$FILE")" | tr -d '\r' > "$TMP_ICI"
  tr -d '\r' < "$AUTRE_FICHIER" > "$TMP_AUTRE"

  if cmp -s "$TMP_ICI" "$TMP_AUTRE"; then
    rm -f "$TMP_ICI" "$TMP_AUTRE"
    continue
  fi

  NB=$(diff "$TMP_ICI" "$TMP_AUTRE" | grep -c '^[<>]')

  echo ""
  echo "ERREUR pre-commit : apres ce commit, $FILE ne serait plus identique a celui"
  echo "de $AUTRE ($NB lignes d'ecart)."
  echo ""
  echo "Les deux depots partagent le MEME fichier. S'ils divergent, on lit une regle"
  echo "perimee selon le depot dans lequel on travaille."
  echo ""
  echo "Ecart (< la version de ce commit, > celle de $AUTRE) :"
  diff "$TMP_ICI" "$TMP_AUTRE" | head -30
  echo ""
  echo "-> Reporte la bonne version dans l'autre depot, puis recommence :"
  echo "     cp $FILE ../$AUTRE/$FILE      (si c'est CE depot qui est a jour)"
  echo "     cp ../$AUTRE/$FILE $FILE      (si c'est l'autre)"
  echo "   Copie octet pour octet : ne relis jamais le fichier pour le reecrire,"
  echo "   les accents se cassent (cf. le piege d'encodage dans CLAUDE.md)."
  echo "   (ou 'git commit --no-verify' pour forcer, deconseille)"
  echo ""

  rm -f "$TMP_ICI" "$TMP_AUTRE"
  FAUTIFS=$((FAUTIFS + 1))
done

[ "$FAUTIFS" -gt 0 ] && exit 1
exit 0
