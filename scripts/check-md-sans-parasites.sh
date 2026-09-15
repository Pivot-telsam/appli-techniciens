#!/bin/sh
# Aucun fichier .md ne doit contenir de caractere de CONTROLE brut
# (tout octet < 0x20 autre que TAB/LF/CR, ou 0x7F).
#
# POSE LE 14/09/2026, APRES COUP. Le motif de detection du double encodage
# avait ete recopie dans CLAUDE.md avec ses octets REELS -- un NUL et un DEL
# au milieu de `[^...-...]` -- au lieu de leur notation `\x00` / `\x7F`.
#
# Ce qu'un seul octet de controle coute : grep, diff et les outils de lecture
# classent TOUT le fichier comme binaire et CESSENT D'AFFICHER a partir de la
# premiere ligne fautive. Sur CLAUDE.md, la coupure tombait ligne 6191 : les
# 1 800 lignes suivantes -- soit tout ce qui a ete ecrit depuis le 10/09 --
# etaient invisibles a une recherche normale. Le comptage, lui, restait juste :
# le fichier ne parait ni tronque ni casse, il se tait simplement sur sa fin.
#
# Ca s'est paye deux fois avant d'etre compris : le 11/09 (« bloc-test-notes
# n'etait pas mort : c'est ma lecture qui le tronquait »), puis le 14/09, quand
# j'ai affirme a Patrice que l'onglet Affaires lisait le classeur de son Bureau
# -- alors qu'il l'avait corrige le 11/09 pour lire Teams, et que la correction
# etait ecrite ligne 7071, APRES la coupure.
#
# Une regle ecrite dans le memo n'aurait rien empeche : c'est justement le memo
# qui n'etait plus lisible. D'ou un controle au commit.
#
# Pris automatiquement par scripts/pre-commit.sh (motif check-*.sh).

RACINE=$(git rev-parse --show-toplevel)
FAUTIFS=0

for MD in "$RACINE"/*.md; do
  [ -f "$MD" ] || continue

  # tr est la seule facon fiable de compter un NUL : les motifs grep/sed ne
  # peuvent pas le porter (le shell tronque toute chaine sur un octet nul).
  # On efface tout ce qui est LEGITIME, on compte ce qui reste.
  N=$(tr -d '\11\12\15\40-\176\200-\377' < "$MD" | wc -c | tr -d ' ')

  if [ "$N" -gt 0 ]; then
    FAUTIFS=$((FAUTIFS + 1))
    echo "  (parasites) $(basename "$MD") : $N caractere(s) de controle brut(s)."
    echo "              Une recherche s'arretera net a la premiere ligne fautive,"
    echo "              et tout ce qui suit sera invisible SANS AUCUNE ERREUR."
    # Reperage indicatif de la ligne (cat -v rend NUL en ^@ et DEL en ^?).
    cat -v "$MD" | grep -n '\^@\|\^?' | head -3 | while read -r L; do
      echo "              -> $(echo "$L" | cut -c1-100)"
    done
    echo "              Ecrire la notation (\\x00, \\x7F...), jamais l'octet lui-meme."
  fi
done

if [ "$FAUTIFS" -gt 0 ]; then
  echo "  (parasites) commit refuse."
  exit 1
fi

exit 0
