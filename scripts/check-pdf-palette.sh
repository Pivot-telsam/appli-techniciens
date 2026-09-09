#!/bin/sh
# Pre-commit hook : refuse un commit qui change les couleurs du PDF de la
# feuille d'heures sans le dire explicitement.
#
# POURQUOI CE CONTROLE EXISTE
#
# Le 09/09/2026, AppTech est passe a la charte sombre « Telsam Ops » : navy,
# terracotta, indigo. Le PDF que l'appli envoie au bureau, lui, garde
# l'ANCIENNE palette (PDF_AMBER = ambre 255,176,0 sur PDF_DARK = 12,16,18).
#
# Ce n'est PAS un oubli. Question posee a Patrice le meme jour, reponse :
# « garde le meme ». Raison : ce PDF n'est pas un ecran, c'est un DOCUMENT
# que Karine et Pierre lisent depuis des semaines pour la paie. Le refaire
# aux couleurs de la charte ne leur apporte rien et leur change un repere
# visuel sans les prevenir.
#
# LE PIEGE CONTRE LEQUEL CE SCRIPT PROTEGE est une bonne intention : une
# session qui voit l'ecran refait et le PDF reste en ambre conclut qu'il
# reste du travail, et « finit la charte ». Le commentaire dans le code ne
# suffit pas -- c'est la lecon du 02/09/2026, ecrite dans CLAUDE.md : une
# regle sans mecanisme ne tient pas. Le mecanisme, c'est ce fichier.
#
# CE N'EST PAS UNE INTERDICTION DEFINITIVE. Le jour ou Patrice demande le
# contraire, il suffit de le dire dans le message de commit :
#
#   git commit -m "PDF de la feuille aux couleurs Telsam (demande de Patrice le JJ/MM)
#
#   PALETTE-PDF-ASSUMEE"
#
# Le mot-cle PALETTE-PDF-ASSUMEE dans le message laisse passer le commit.
# Il oblige seulement a ecrire QUI a demande le changement, ce qui est tout
# l'objet du controle.
#
# Install : cf. scripts/pre-commit.sh (qui enchaine tous les check-*.sh).

FILE="index.html"

git diff --cached --name-only | grep -qx "$FILE" || exit 0

# On ne regarde que les lignes AJOUTEES ou RETIREES qui definissent la
# palette du PDF. Une ligne qui les utilise (doc.setFillColor(PDF_DARK...))
# peut bouger librement : c'est la VALEUR des couleurs qu'on protege, pas le
# code qui s'en sert.
CHANGE=$(git diff --cached -- "$FILE" \
  | grep -E "^[+-]" \
  | grep -vE "^(\+\+\+|---)" \
  | grep -E "^[+-]const PDF_(AMBER|DARK|DARK2|TEXT_MUTED|BORDER) *=")

[ -z "$CHANGE" ] && exit 0

# Le feu vert explicite, dans le message de commit en preparation.
MSG_FILE=$(git rev-parse --git-dir)/COMMIT_EDITMSG
if [ -f "$MSG_FILE" ] && grep -q "PALETTE-PDF-ASSUMEE" "$MSG_FILE"; then
  exit 0
fi

echo ""
echo "ERREUR pre-commit : les couleurs du PDF de la feuille d'heures changent."
echo ""
printf '%s\n' "$CHANGE" | sed 's/^/  /'
echo ""
echo "Ce PDF garde volontairement l'ancienne palette ambre : c'est un DOCUMENT"
echo "que Karine et Pierre lisent depuis des semaines pour la paie, pas un ecran."
echo "Question posee a Patrice le 09/09/2026, reponse : « garde le meme »."
echo ""
echo "-> Si ce n'etait pas voulu : annule ces lignes (git checkout -p)."
echo "-> Si Patrice a change d'avis : ecris qui l'a demande dans le message de"
echo "   commit, avec le mot PALETTE-PDF-ASSUMEE sur une ligne."
echo ""
exit 1
