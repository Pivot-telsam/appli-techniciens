#!/bin/sh
# Pre-commit hook : refuse d'envoyer du COMMERCIAL dans ce depot, QUI EST PUBLIC.
#
# POURQUOI IL EXISTE. Le 15/09/2026, on a decouvert que `JOURNAL.md` -- 8995
# lignes, 635 Ko -- etait lisible sur github.com SANS AUCUN COMPTE : noms de
# clients, montants de devis, numeros et montants de factures, commandes RTE,
# methode de calcul des marges. Il y etait depuis le 20/08/2026, entre a 5 Ko et
# grossi jusqu'a 540 Ko sans que personne ne redecide jamais si ca avait sa place
# en ligne. Et GitHub Pages servait le fichier sous l'adresse de l'appli des
# techniciens, celle qu'un moteur de recherche indexe volontiers.
#
# CE QUE CET INCIDENT A APPRIS, ET QUE CE CONTROLE APPLIQUE. La cause n'etait pas
# une fausse manipulation : c'etait une regle juste (« les deux memos sont
# identiques dans les deux depots »), portee par un hook, appliquee sans que
# personne se demande OU elle envoyait ce qu'elle recopiait. Aucun des controles
# existants ne regardait le CONTENU : ils verifiaient des versions, des accents,
# des statuts. Il manquait celui-ci.
#
# CE QU'IL REFUSE (et pourquoi c'est ce niveau-la) :
#   - les deux memos, par leur nom -- ils n'ont plus rien a faire ici ;
#   - un montant en euros -- c'est le plus revelateur de tous ;
#   - un numero de commande RTE ou client (4500…, 6100…, 920…) ;
#   - les noms de colonnes du classeur commercial, signe qu'on a recopie une
#     ligne de suivi d'affaire.
#
# LES NUMEROS DE DEVIS SONT PASSES DE « TOLERES » A « REFUSES » (17/09/2026).
#
# Ils etaient comptes et nommes sans etre refuses, au motif qu'ils vivaient dans
# `tachesVendues` et qu'y toucher changerait l'outil du technicien. Patrice a
# tranche l'inverse : « supprime aussi les references du devis, je ne sais pas
# pourquoi nous les avons mis la, ils n'ont pas besoin de savoir d'ou ca vient ».
# Les 26 ont ete retirees le meme jour -- aucune n'etait affichee nulle part,
# c'etait de la donnee morte dans un fichier public.
#
# ET LE MOTIF A ETE ELARGI, PARCE QU'IL NE VOYAIT QU'UN DIXIEME DU PROBLEME :
# il ne cherchait que la forme a barres obliques (TELSAM/CC/…). Les references
# etaient ecrites avec des SOULIGNES dans 23 fiches sur 26, et passaient donc
# sans un mot. Le controle annoncait « 3 numeros » la ou il y en avait 26 -- et
# l'arbitrage rendu la-dessus portait sur un chiffre faux. Un controle qui ne
# voit qu'une forme d'ecriture ment avec l'accent de la certitude.
#
# UN CONTROLE QUI NE PEUT PAS PASSER DETRUIT LA SUITE ENTIERE : les motifs
# ci-dessous ont ete calibres sur le contenu reel du depot le 15/09/2026, et le
# depot passe au vert une fois les trois fuites trouvees ce jour-la corrigees
# (un montant de devis et deux references de devis dans des textes d'alerte).
#
# Install : cf. scripts/pre-commit.sh

RACINE=$(git rev-parse --show-toplevel)
ICI=$(basename "$RACINE")

# Ce controle ne vaut QUE pour le depot public. Dans `suivi-chantiers`, qui est
# prive et derriere un mot de passe, le commercial est chez lui.
[ "$ICI" = "appli-techniciens" ] || exit 0

MIS_EN_SCENE=$(git diff --cached --name-only --diff-filter=ACM)
[ -z "$MIS_EN_SCENE" ] && exit 0

FAUTIFS=0
DEVIS_VUS=0

for FILE in $MIS_EN_SCENE; do

  # --- 1. Les deux memos, par leur nom -------------------------------------
  case "$FILE" in
    CLAUDE.md|JOURNAL.md)
      echo ""
      echo "ERREUR pre-commit : $FILE ne doit plus etre versionne dans ce depot."
      echo ""
      echo "Ce depot est PUBLIC. Ces deux fichiers racontent tout le commercial de"
      echo "TELSAM et ont deja ete lisibles en ligne pendant 26 jours (15/09/2026)."
      echo "Ils restent sur le disque, a cote du depot, et leur exemplaire de"
      echo "reference vit dans le depot prive suivi-chantiers."
      echo ""
      echo "-> Retire-le de ce commit :   git reset HEAD $FILE"
      echo ""
      FAUTIFS=$((FAUTIFS + 1))
      continue
      ;;
  esac

  # Les images et le binaire ne se lisent pas ligne a ligne.
  case "$FILE" in
    *.png|*.jpg|*.jpeg|*.gif|*.ico|*.zip|*.pdf|*.xlsx|*.docx) continue ;;
  esac

  # LES CONTROLES EUX-MEMES SONT EXEMPTES, et il a fallu s'y reprendre a deux
  # fois pour le voir : ce fichier CONTIENT les motifs qu'il cherche -- il ne
  # peut pas nommer ce qu'il refuse sans l'ecrire. Sans cette exemption, il se
  # refusait lui-meme et bloquait tous les commits du depot, c'est-a-dire
  # exactement « un controle qui ne peut pas passer detruit la suite entiere ».
  # Le trou est nul : ces fichiers sont courts, relus, et ne portent aucune
  # valeur reelle -- les exemples ci-dessus sont inventes, verifie.
  case "$FILE" in
    scripts/check-*.sh) continue ;;
  esac

  # On lit le contenu MIS EN SCENE, pas le fichier du disque : c'est lui qui
  # partira dans le commit. `git cat-file blob` rend les octets bruts (cf. le
  # piege de core.autocrlf documente dans check-claude-md-sync.sh).
  BLOB=$(git rev-parse ":$FILE" 2>/dev/null) || continue
  TMP=$(mktemp)
  git cat-file blob "$BLOB" > "$TMP" 2>/dev/null

  # --- 2. Ce qui REFUSE ----------------------------------------------------
  # Un montant en euros, colle ou espace, avec ou sans centimes. Au moins trois
  # chiffres, pour ne pas se declencher sur un prix a un chiffre ni sur une
  # annee. (Exemples volontairement absents : ce fichier serait son propre
  # fautif -- cf. l'exemption des scripts/check-*.sh plus haut.)
  MONTANTS=$(grep -oE '[0-9][0-9 ]{2,}(,[0-9]{2})? ?€|€ ?[0-9][0-9 ]{2,}' "$TMP" | sort -u | head -5)
  # Une commande RTE ou client : 4500……, 6100……, 920……
  COMMANDES=$(grep -oE '\b(4500|6100)[0-9]{6}\b|\b920[0-9]{6}\b' "$TMP" | sort -u | head -5)
  # Les colonnes du classeur commercial, recopiees telles quelles.
  COLONNES=$(grep -oiE 'Montant devis HT|Factu Partiel|Montant Restant a Facturer|Sold[eé] O/N|Montant Achat' "$TMP" | sort -u | head -5)

  if [ -n "$MONTANTS" ] || [ -n "$COMMANDES" ] || [ -n "$COLONNES" ]; then
    echo ""
    echo "ERREUR pre-commit : $FILE contient du commercial, et CE DEPOT EST PUBLIC."
    echo ""
    [ -n "$MONTANTS" ]  && { echo "  Montant(s) en euros :"; echo "$MONTANTS"  | sed 's/^/    /'; }
    [ -n "$COMMANDES" ] && { echo "  Numero(s) de commande :"; echo "$COMMANDES" | sed 's/^/    /'; }
    [ -n "$COLONNES" ]  && { echo "  Colonne(s) du classeur commercial :"; echo "$COLONNES" | sed 's/^/    /'; }
    echo ""
    echo "  L'appli technicien dit le FAIT, jamais le commercial : ce qu'un"
    echo "  technicien doit faire, pas ce que ca coute ni sous quelle commande."
    echo "  Reecris le passage sans le chiffre -- l'information utile sur le"
    echo "  terrain survit toujours a ce retrait -- puis recommence."
    echo ""
    echo "  (ou 'git commit --no-verify' pour forcer, a ne pas faire ici)"
    echo ""
    FAUTIFS=$((FAUTIFS + 1))
  fi

  # --- 3. Les numeros de devis, sous TOUTES leurs ecritures -----------------
  # Barres obliques, soulignes, avec ou sans RTE, avec ou sans suffixe de
  # version : c'est la meme reference, et n'en chercher qu'une forme revient a
  # ne pas chercher (23 sur 26 passaient avant le 17/09/2026).
  DEVIS=$(grep -oE '(TELSAM|DATA)[_/]?CC[_/]?(RTE[_/]?)?[0-9][A-Z0-9_.-]*' "$TMP" | sort -u)
  if [ -n "$DEVIS" ]; then
    NB=$(echo "$DEVIS" | grep -c .)
    DEVIS_VUS=$((DEVIS_VUS + NB))
    echo ""
    echo "ERREUR pre-commit : $FILE porte $NB reference(s) de devis, et CE DEPOT EST PUBLIC."
    echo "$DEVIS" | sed 's/^/    /'
    echo ""
    echo "  Le technicien n'a pas besoin de savoir d'ou vient sa liste de taches"
    echo "  (arbitrage de Patrice, 17/09/2026). Retire la reference : ce qu'il doit"
    echo "  faire sur le terrain survit toujours a ce retrait."
    echo ""
    FAUTIFS=$((FAUTIFS + 1))
  fi

  rm -f "$TMP"
done

if [ "$FAUTIFS" -gt 0 ]; then
  exit 1
fi

echo "  (commercial) OK : aucun montant, aucune commande, aucun memo dans ce commit."
exit 0
