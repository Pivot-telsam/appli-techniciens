#!/bin/sh
# Pre-commit hook: refuse a commit that changes CLAUDE.md in one repo while the
# sibling repo's CLAUDE.md says something different.
#
# Reason: appli-techniciens/CLAUDE.md and suivi-chantiers/CLAUDE.md are the SAME
# document, dupliquee dans les deux depots pour que chaque clone soit complet.
# Rien ne les tenait ensemble : elles ont diverge en silence pendant 3 jours
# (constate le 31/08/2026). suivi-chantiers annoncait encore des liens
# "?tech=slugifiedname" personnels et permanents, alors que le portail par mot de
# passe du 27/08 avait supprime la lecture de ce parametre. Selon le depot dans
# lequel on travaillait, on lisait une regle perimee -- et on aurait pu
# rediffuser ces liens comme s'ils identifiaient encore un technicien.
#
# Ce controle ne se declenche que si CLAUDE.md fait partie du commit.
#
# Install : cf. scripts/pre-commit.sh

FILE="CLAUDE.md"

git diff --cached --name-only | grep -qx "$FILE" || exit 0

RACINE=$(git rev-parse --show-toplevel)
ICI=$(basename "$RACINE")

case "$ICI" in
  appli-techniciens) AUTRE="suivi-chantiers" ;;
  suivi-chantiers)   AUTRE="appli-techniciens" ;;
  *) exit 0 ;;
esac

AUTRE_FICHIER="$RACINE/../$AUTRE/$FILE"

if [ ! -f "$AUTRE_FICHIER" ]; then
  echo ""
  echo "AVERTISSEMENT pre-commit : $AUTRE/$FILE est introuvable a cote de ce depot."
  echo "La coherence des deux CLAUDE.md n'a PAS pu etre verifiee -- le commit passe"
  echo "quand meme, mais pense a reporter la modification dans l'autre depot."
  echo "Chemin attendu : $AUTRE_FICHIER"
  echo ""
  exit 0
fi

# On compare le contenu MIS EN SCENE (l'index), pas le fichier du disque : c'est
# lui qui partira dans le commit.
#
# PIEGE (rencontre en testant ce hook le 31/08/2026) : NE PAS utiliser
# `git show :CLAUDE.md`, qui applique la conversion de fins de ligne de
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
  exit 0
fi

NB=$(diff "$TMP_ICI" "$TMP_AUTRE" | grep -c '^[<>]')

echo ""
echo "ERREUR pre-commit : apres ce commit, $FILE ne serait plus identique a celui"
echo "de $AUTRE ($NB lignes d'ecart)."
echo ""
echo "Les deux depots partagent le MEME fichier de consignes. S'ils divergent, on lit"
echo "une regle perimee selon le depot dans lequel on travaille."
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
exit 1
