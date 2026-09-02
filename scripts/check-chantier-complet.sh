#!/bin/sh
# Hook pre-commit : refuse un commit qui laisserait un chantier a moitie equipe.
#
# POURQUOI
# La checklist "nouveau chantier" de CLAUDE.md a ete appliquee a moitie DEUX fois,
# et les deux fois c'est quelqu'un d'autre qui l'a vu :
#   - 21/08/2026 : dossiers App Tech complets mais sans File Request ni depotTerrain
#     -> pas de bouton photos. Signale par Pascal Bonaventure.
#   - 02/09/2026 : dossier App Tech monte sans tachesVendues -> pas de bouton Suivi.
#     Signale par Patrice : "je vais etre oblige de verifier systematiquement ce que
#     tu fais, et ce n'est pas le but de cet outil".
# Une regle ecrite ne vaut que ce que vaut la memoire de celui qui la lit. Ce hook
# est le mecanisme qui la porte, au moment exact ou l'oubli serait livre.
#
# LA REGLE
# Des qu'une fiche NON TERMINEE porte "documentsAppTech" non vide, elle doit aussi
# porter "depotTerrain" non vide ET "tachesVendues" avec au moins une tache.
# Les trois boutons vont ensemble : Documents, Depot photos, Suivi.
#
# Ce controle ne regarde QUE le fichier mis en scene (l'index), c'est-a-dire ce qui
# partira vraiment dans le commit.
#
# Exception : ajouter un numero de chantier dans EXCEPTIONS ci-dessous, avec sa
# raison. Une exception doit etre un choix visible, jamais un oubli silencieux.

FILE="index.html"
EXCEPTIONS=""   # ex: "26-062:audit termine, pas de retour prevu"

git diff --cached --name-only | grep -qx "$FILE" || exit 0

TMP=$(mktemp)
git cat-file blob "$(git rev-parse ":$FILE")" > "$TMP" 2>/dev/null || { rm -f "$TMP"; exit 0; }

# SEED_DATA tient sur une seule ligne. On la sort, puis on decoupe fiche par fiche
# sur '{"id":"' : chaque morceau contient les champs d'un chantier.
# index() plutot que des expressions regulieres : en awk POSIX, '{' ouvre un
# intervalle et '\{' n'est pas portable. Premier essai du 02/09/2026 : le motif
# '"tachesVendues":\{[^}]*"taches":\[\{' ne correspondait a RIEN, et le hook
# accusait les 13 fiches equipees au lieu de la seule qui l'etait vraiment.
# Une recherche de chaine litterale ne peut pas se tromper ainsi.
PROBLEMES=$(awk '
  /^[[:space:]]*const[[:space:]]+SEED_DATA[[:space:]]*=/ {
    # Decoupe sur l identifiant de CHANTIER, qui commence toujours par "c_".
    # Decouper sur {"id":" tout court coupait la fiche en plein milieu de
    # tachesVendues, dont chaque tache porte aussi un "id" ({"id":"racc"...}) :
    # le hook accusait alors les 13 fiches equipees au lieu de la seule fautive.
    n = split($0, morceaux, /\{"id":"c_/)
    for (i = 2; i <= n; i++) {
      f = morceaux[i]
      numero = "?"
      if (match(f, /"numero":"[^"]*"/)) numero = substr(f, RSTART + 10, RLENGTH - 11)
      if (index(f, "\"termine\":true") > 0) continue
      if (index(f, "\"documentsAppTech\":\"http") == 0) continue

      manque = ""
      if (index(f, "\"depotTerrain\":\"http") == 0) manque = manque "+depotTerrain(bouton_photos)"
      if (index(f, "\"taches\":[{") == 0)            manque = manque "+tachesVendues(bouton_Suivi)"
      if (manque != "") print numero "|" manque
    }
  }
' "$TMP")
rm -f "$TMP"

[ -z "$PROBLEMES" ] && exit 0

# Retirer les exceptions declarees. Lecture ligne a ligne : un "for" sur une
# variable decoupe sur les espaces et melangeait tout a l'affichage.
RESTE=$(printf '%s\n' "$PROBLEMES" | while IFS= read -r LIGNE; do
  [ -z "$LIGNE" ] && continue
  NUM=${LIGNE%%|*}
  if printf '%s' "$EXCEPTIONS" | grep -q "$NUM:"; then continue; fi
  printf '%s\n' "$LIGNE"
done)
[ -z "$RESTE" ] && exit 0

echo ""
echo "ERREUR pre-commit : un ou plusieurs chantiers ont un dossier App Tech mais"
echo "pas tous les boutons qui vont avec. Un technicien qui ouvre sa fiche n'aura"
echo "pas de quoi deposer ses photos, ou pas de quoi declarer son avancement."
echo ""
printf '%s\n' "$RESTE" | while IFS= read -r LIGNE; do
  NUM=${LIGNE%%|*}
  MANQUE=$(printf '%s' "${LIGNE#*|}" | tr '+' ' ' | tr '_' ' ')
  echo "  $NUM -> manque :$MANQUE"
done
echo ""
echo "-> Completer la fiche (cf. CLAUDE.md, checklist 'Nouveau chantier') :"
echo "   depotTerrain  = URL de la File Request Dropbox vers 'Photos terrain'"
echo "   tachesVendues = taches relevees dans le devis"
echo "   (ou declarer une exception motivee dans scripts/check-chantier-complet.sh)"
echo ""
exit 1
