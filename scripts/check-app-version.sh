#!/bin/sh
# Pre-commit hook: refuse a commit that edits index.html without also bumping
# APP_VERSION in the same commit.
#
# Reason: the opening animation only replays when APP_VERSION changes. Without a
# bump, a technician who already opened the app sees nothing — and above all he
# loses the only signal telling him he received the new version and not the old
# one still sitting in his browser cache (cf. CLAUDE.md, section "Ecran
# d'ouverture anime"). Same class of trap as SEED_VERSION on suivi-chantiers,
# which was forgotten twice despite being documented.
#
# Also checks the format AAAA-MM-JJ-n, because the date shown to the technician
# under the logo is derived from APP_VERSION itself.
#
# Install (once per clone — hooks are not versioned by git itself):
#   cp scripts/pre-commit.sh .git/hooks/pre-commit
#   chmod +x .git/hooks/pre-commit
#
# (Ce script n'est plus installe directement comme hook : le depot a plusieurs
#  controles, et c'est scripts/pre-commit.sh qui les enchaine tous.)

FILE="index.html"

if ! git diff --cached --name-only | grep -qx "$FILE"; then
  exit 0
fi

DIFF=$(git diff --cached -- "$FILE")

# LES DONNEES ECRITES PAR LES SCRIPTS NE COMPTENT PAS COMME UNE MISE A JOUR.
#
# Depuis le 03/09/2026, la chaine du matin (7h30 et 13h00) reecrit deux lignes
# de ce fichier a chaque passage : PLANNING_TECH (le planning lu dans Teams) et
# AVANCEMENT_DECLARE (les cases deja cochees). Elles changent donc DEUX FOIS PAR
# JOUR, souvent pour le seul horodatage.
#
# Sans cette exception, deux issues, toutes deux mauvaises :
#   - soit APP_VERSION est bumpe chaque jour, et le technicien se prend l'ecran
#     d'ouverture tous les matins. La doc le dit : « insupportable en trois
#     jours, et le signal perdrait tout son sens a force d'etre vu » - donc on
#     detruirait precisement ce que ce hook protege ;
#   - soit tout commit suivant un passage du matin est bloque.
#
# Meme raisonnement que POSES_APPLI dans le suivi : une donnee regeneree vit a
# part justement pour ne pas declencher de bump.
#
# ON RETIRE CES LIGNES DU DIFF, ET ON REGARDE S'IL RESTE QUELQUE CHOSE. Ne pas
# remplacer ce test par un « si le diff CONTIENT ces lignes, on passe » : un
# vrai changement de code passerait alors en douce des qu'un passage du matin
# l'accompagne.
RESTE=$(printf '%s\n' "$DIFF" \
  | grep -E "^[+-]" \
  | grep -vE "^(\+\+\+|---)" \
  | grep -vE "^[+-]const (PLANNING_TECH|AVANCEMENT_DECLARE) = ")

if [ -z "$RESTE" ]; then
  exit 0
fi

ANCIENNE=$(printf '%s\n' "$DIFF" | grep -E "^-var APP_VERSION" | head -1 | sed "s/.*'\(.*\)'.*/\1/")
NOUVELLE=$(printf '%s\n' "$DIFF" | grep -E "^\+var APP_VERSION" | head -1 | sed "s/.*'\(.*\)'.*/\1/")

if [ -z "$NOUVELLE" ] || [ "$ANCIENNE" = "$NOUVELLE" ]; then
  echo ""
  echo "ERREUR pre-commit : $FILE a change mais APP_VERSION n'a pas ete incremente."
  echo "Sans ce bump, l'ecran d'ouverture ne se rejoue pas : le technicien n'a aucun"
  echo "signe qu'il a bien recu la nouvelle version et non l'ancienne restee en cache."
  echo "-> Incremente APP_VERSION (var APP_VERSION = 'AAAA-MM-JJ-n') avant de committer."
  echo "   (ou 'git commit --no-verify' pour forcer, deconseille)"
  echo ""
  exit 1
fi

# Une nouvelle version doit dire ce qu'elle change, sinon le technicien relit la
# phrase de la fois precedente — pire que pas de phrase du tout, parce qu'elle a
# l'air a jour.
NOUVEAUTE_CHANGEE=$(printf '%s\n' "$DIFF" | grep -cE "^\+var APP_NOUVEAUTE")

if [ "$NOUVEAUTE_CHANGEE" -eq 0 ]; then
  echo ""
  echo "ERREUR pre-commit : APP_VERSION passe a '$NOUVELLE' mais APP_NOUVEAUTE n'a pas change."
  echo "Le technicien lirait le message de la mise a jour precedente, en croyant"
  echo "qu'il decrit celle-ci."
  echo "-> Reecris APP_NOUVEAUTE : une phrase, de son point de vue, honnete meme"
  echo "   quand la reponse est 'rien qui te concerne cette fois'."
  echo "   (ou 'git commit --no-verify' pour forcer, deconseille)"
  echo ""
  exit 1
fi

# La date affichee sous le logo est deduite de APP_VERSION : un format casse
# donnerait une date absurde au technicien, sans erreur visible ailleurs.
if ! printf '%s' "$NOUVELLE" | grep -qE '^[0-9]{4}-[0-9]{2}-[0-9]{2}-[0-9]+$'; then
  echo ""
  echo "ERREUR pre-commit : APP_VERSION = '$NOUVELLE' n'a pas le format attendu."
  echo "Format : AAAA-MM-JJ-n  (exemple : 2026-08-26-1)"
  echo "La date affichee au technicien sous le logo en est deduite : un format"
  echo "different lui afficherait une date fausse, sans erreur ailleurs."
  echo ""
  exit 1
fi

exit 0
