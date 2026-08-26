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
#   cp scripts/check-app-version.sh .git/hooks/pre-commit
#   chmod +x .git/hooks/pre-commit

FILE="index.html"

if ! git diff --cached --name-only | grep -qx "$FILE"; then
  exit 0
fi

DIFF=$(git diff --cached -- "$FILE")

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
