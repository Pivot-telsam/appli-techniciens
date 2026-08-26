# État d'avancement — sessions des 20, 21, 24, 25 et 26/08/2026

**Ne pas noter ici de numéro de commit ni de version** : ces valeurs sont fausses dès le commit
suivant, et un fichier de reprise qui se trompe dès sa troisième ligne est pire qu'un fichier
muet. Pour connaître l'état réel :
- derniers commits : `git log --oneline -1` dans chaque dépôt ;
- version en ligne de l'appli : `grep '^var APP_VERSION' appli-techniciens/index.html`.

Ce fichier récapitule une longue session de travail sur les deux dépôts. Objectif : pouvoir
reprendre le fil après compactage de la conversation, sans perdre le contexte.

## Ce qui est fait

### Session du 21/08/26 (suite)
- Copies de `CLAUDE.md` et `PROGRESS.md` ajoutées dans `suivi-chantiers` aussi (les deux dépôts
  sont interconnectés, doivent rester synchronisés sur ces deux fichiers de référence).
- **Correction importante du Gantt (`suivi-chantiers`)** : la barre bleue "Intervention TELSAM
  prévue" était construite à partir du texte des `fenetres` (plages larges, pas reconfirmées
  chaque semaine), pas à partir de `REAL_DAYS` (jours réellement confirmés par la couleur du
  planning). Conséquence concrète repérée par Patrice : Cantegrit/Fleyriat/Joncquiere/Lisieux/
  Chaineau-Cordy-Lamotte affichaient une "présence" TELSAM alors qu'aucun technicien n'y était
  placé en semaine 35 ; à l'inverse Audit Multi Postes (sans champ `fenetres`) n'affichait rien
  du tout malgré 2 techniciens confirmés. Corrigé : la barre est maintenant construite depuis
  `REAL_DAYS` (le texte des fenêtres n'alimente plus que le tooltip). Le tri du Gantt (chantiers
  "en cours" en tête) se corrige de facto. `mergePresenceSegs` supprimée (remplacée par
  `mergeCoveragePeriods`, déjà existante, réutilisée). Vérifié sans erreur console, testé sur
  tous les chantiers cités + quelques autres en sanity-check. Règle documentée dans `CLAUDE.md`.
- **Tri du Gantt corrigé aussi** : `chronoSortKey` faisait passer PGO/consignation/NIP à égalité
  avec la présence technicien réelle, donc Cantegrit (PGO seul) pouvait remonter au même niveau
  que Portet/DATA4/Audit (techniciens réellement placés). La présence réelle passe maintenant
  toujours en premier. Vérifié : Portet, DATA4-Marcoussis, Audit Multi Postes sont bien les 3
  premiers du Gantt en semaine 35.
- **PDP reçus pour DATA4-Marcoussis et Audit Garies-Pessac** (documents ajoutés par Patrice dans
  App Tech) : statuts pdp/pgo mis à jour dans les deux dépôts (voir commits `92b08aa`/`9f1f43f`).
  Point notable : le PGO de l'Audit (indice 1, .xlsm) est en réalité un **modèle vide** (aucune
  tâche/date renseignée dans les feuilles PGO/Gantt/Données) — statut mis à `warn` avec alerte
  dédiée plutôt que `ok`, à relancer auprès de RTE.
- **Rappel important (répété 2 fois ce jour-là)** : ajouter un fichier dans Dropbox "App Tech"
  ne met JAMAIS à jour automatiquement les champs `pdp`/`pgo`/`ist` de `SEED_DATA` — il faut
  toujours qu'une session interactive ouvre le document et saisisse l'indice/les dates.
- **Tentative d'automatisation abandonnée** : Patrice a demandé si la détection de nouveaux
  PDP/PGO/IST dans Dropbox pouvait être automatisée. Tentative de créer une routine planifiée
  cloud (skill `schedule` / `RemoteTrigger`) — bloquée à deux niveaux : (1) aucun moyen de
  récupérer le `connector_uuid` du connecteur Dropbox pour l'attacher à la routine (la page
  Connecteurs de Patrice montre des permissions par outil, pas un ID de connecteur ; aucun outil
  API ne liste les connecteurs disponibles), (2) `RemoteTrigger create` a été rejeté avec
  "Connect your GitHub account before saving a routine that uses a GitHub repository" — GitHub
  n'est pas connecté pour les routines non plus, et ni Patrice ni Claude n'ont trouvé où faire
  ça dans son interface. Décision : **on reste sur le traitement manuel en session**, comme
  avant. Ne pas retenter cette automatisation sans une vraie raison de penser que ces deux
  blocages sont résolus (cf. mémoire `feedback_no_cloud_automation_pdp_pgo`). Autre limite à
  garder en tête si retenté un jour : même connecteurs OK, une routine cloud ne pourrait
  toujours pas ouvrir les PGO/PDP `.xlsm` (Excel via COM PowerShell, Windows uniquement) ni
  générer les Briefs Techniciens (Word COM) — ça resterait à faire en session locale.

### Mise en place / structure
- Dépôt `appli-techniciens` cloné dans `TELSAM-apps/appli-techniciens/`.
- Dépôt `suivi-chantiers` créé (privé, GitHub `Pivot-telsam/suivi-chantiers`), contient
  `suivi_chantiers_205.html`.
- `TELSAM-apps/` (racine) n'est PAS un dépôt Git — juste un dossier contenant les deux dépôts
  en sous-dossiers séparés.
- `appli-techniciens/CLAUDE.md` créé et enrichi au fil de la session — c'est la référence
  vivante des règles métier (à toujours consulter en début de tâche liée à ces deux apps).

### Données chantiers
- PDP Fleyriat validé indice 3, valable jusqu'au 15/09/26.
- ~28 chantiers ont désormais un champ `baseVie` (adresse base de vie/dépôt extraite des NDS
  Dropbox) — le bouton "📍 Itinéraire" de `appli-techniciens` priorise cette adresse, sinon
  retombe sur le point géo. Environ 20 chantiers restent sans adresse trouvée (non prioritaire
  sauf demande).
- Fusions de fiches doublons (données fusionnées, historique tracé, fiche obsolète supprimée
  dans les deux dépôts) :
  - "Poste de Tivernon" (26-070) → "Dambron - Voves" (26-031)
  - "Vallee Du Louron (Aure-Loudenvielle)" (26-011) → "Aure - Loudenvielle" (26-024) — le
    `baseVie` de 26-011 n'a volontairement PAS été repris (mal attribué, venait du NDS d'un
    segment voisin Bordières-Tramezaygues)
- Renommages (sourcés PDP/devis signé, jamais sur la seule base du planning brut) :
  - 26-002 : "Icp / Ls Cross-Sausset-Goodman" → "Ls Cross-Sausset-Goodman"
  - 26-003 : "PS Cantegrit" → "Poste Cantegrit"
  - 26-007 : "CONSIGNATION 225kV CURBANS - SISTERON" → "Curbans - Sisteron (225kV)"
    (devis signé TELSAM/CC/RTE/25027)
- PGO Bradascou-Repaire-Rouverade : `ref`+`validFrom`/`validUntil` renseignés (les 2 fenêtres
  TELSAM du PGO ind.10 étaient déjà correctement calculées par l'app, juste masquées par
  défaut par la case "Afficher les semaines passées" décochée — pas un bug, rien d'autre à
  corriger côté code).
- Semaine 35 (24-28/08) : affectations techniciens mises à jour dans les deux dépôts
  (Poste de Portet, DATA4-Marcoussis, Audit Multi Postes = Audit Garies-Pessac).
- Dossiers "App Tech" créés et alignés sur la convention pour DATA4-Marcoussis et
  Audit Garies-Pessac (Brief Techniciens généré, NDS/MO/PGO/PDP, dossier "Photos terrain"),
  `documentsAppTech` renseigné dans `appli-techniciens`, liens Dropbox passés en public par
  Patrice. `termine:true` erroné corrigé en `false` pour DATA4-Marcoussis.

### Bugs / trous trouvés et corrigés
- **`SEED_VERSION`** (suivi-chantiers) n'avait jamais été incrémenté malgré de nombreux
  changements de `SEED_DATA` dans la session — corrigé (`v85-2026-08-19c` → `v86-2026-08-20`).
  Sans ça, les mises à jour étaient invisibles pour les collègues ayant déjà chargé le fichier.
- Règle du **codage couleur du Planning RTE 2026.xlsx** ratée une fois (affectations
  technicien/chantier lues via couleur de cellule, pas via texte) — corrigée et documentée en
  détail dans `CLAUDE.md`.
- `create_shared_link` (connecteur Dropbox) crée toujours des liens privés — Patrice n'a pas
  besoin d'accès admin, un simple clic "Partager" → réglage du "lien de consultation" sur
  "Toute personne disposant du lien" suffit, dossier par dossier.
- Audit complet des commentaires de code des deux fichiers HTML → plusieurs règles non
  documentées ajoutées à `CLAUDE.md` (trou de couverture PGO malgré statut "ok" global, sévérité
  du statut `nc` sauf phase touret, Gantt masquant les semaines passées par défaut,
  `MTFO_STANDALONE` = choix produit délibéré, badges PDP/PGO côté appli-techniciens liés à
  `documentsAppTech`/migration, code d'accès partagé = confort pas sécurité).

### Règles désormais actées dans CLAUDE.md
- Convention de nommage chantier (PDP/devis signé ou dossier Dropbox existant, jamais planning
  brut seul, sinon demander confirmation).
- Structure Dropbox `RTE > [Ligne aérienne | Ligne souterraine | Postes | Fibrage | Audit FO |
  Arteria] > dossier chantier > Documents Telsam`.
- Numérotation `26-0XX` (Patrice renomme progressivement les dossiers Dropbox avec ce préfixe —
  en cours de sa part, pas encore terminé sur l'ensemble du Dropbox).
- Règle Planning RTE 2026.xlsx (codage couleur).
- Contenu standard obligatoire d'un dossier "App Tech" : Brief Techniciens (skill
  `brief-techniciens`, secours Word/COM→PDF car pas de Python sur cette machine), MO/NDS,
  PDP/PGO pertinents, IST *seulement si signée RTE* (sinon alerte `warn` + `valideRTE:false`),
  sous-dossier "Photos terrain" vide — s'applique à tout chantier qui devient actif, pas
  seulement aux nouveaux.
- **Mise à jour continue obligatoire** : dès qu'un PDP/PGO/IST plus récent est connu pour un
  chantier ayant déjà un App Tech, remplacer l'ancienne version par la nouvelle (jamais de doc
  périmé qui traîne), et mettre à jour les champs structurés en cohérence dans les deux dépôts.
- Vérification systématique de la visibilité des liens Dropbox (`get_shared_link_metadata`)
  avant de committer `documentsAppTech`.
- Toujours montrer le résumé des changements avant de committer/pousser (sauf demande contraire
  explicite) ; ne jamais deviner un nom de chantier ambigu ou un chemin Dropbox — poser la
  question.

### Mémoire long-terme (persistante entre sessions, hors dépôts Git)
- Pas de Python sur cette machine → PowerShell + Excel/Word COM pour tout traitement
  xlsx/génération PDF.
- Règle couleur Planning RTE.
- Obligation de bump `SEED_VERSION`.
- Vérification systématique des liens Dropbox.


### Session du 24/08/26 — le pont feuilles d'heures (chantier majeur, terminé et validé)

Objectif de Patrice : que les feuilles d'heures saisies par les techniciens alimentent son récap
RH sans aucune ressaisie de sa part, tout en gardant son point de contrôle avant envoi aux RH.
**La chaîne complète est en place et validée avec de vraies feuilles.** Le détail des règles est
dans `CLAUDE.md`, section « Feuilles d'heures : de l'appli au récap RH » — ne pas le dupliquer ici.

Ce qui a été construit :
- `suivi-chantiers/scripts/integrer-feuilles-heures.ps1` : lit les feuilles déposées, montre un
  aperçu, n'écrit qu'avec `-Ecrire`. Idempotent, refuse d'écrire dans quatre situations à risque.
- L'appli technicien embarque les données de saisie **dans le PDF** (un seul fichier à partager).
- Dossier de dépôt `feuille d'heures/dépôts appli/` avec son `_LISEZ-MOI.txt`.
- N° d'index `26-0XX` propagé aux 68 fiches de `appli-techniciens` (il y était absent).
- Colonnes `Devis N` du récap renommées en `N° chantier N`, skill `recap-feuilles-heures` adapté
  en conséquence (accepte les deux formats, les anciens fichiers mensuels restent recalculables).
- Corrections issues du test terrain avec Ahmed Hamouch : boutons de déplacement rendus
  compréhensibles, totaux GD/PD/Nacelle affichés sur le PDF, avertissement avant l'envoi d'une
  feuille vide, une seule nacelle par jour, bouton « Effacer et recommencer cette semaine ».
- Logo TELSAM ajouté en en-tête de l'appli (écrans d'accueil et bandeau).

Validation : Pascal Bonaventure (35h sur DATA4-Marcoussis, GD 4, PD 1) et Ahmed Hamouch
(GD 4, PD 1 seulement) ont rempli sur leur téléphone ; leurs feuilles sont arrivées dans
`recap_FH-TEST.xlsx` sans ressaisie, et le skill régénère correctement les onglets 2 et 3.



## État des fichiers au 24/08/2026 (à lire avant de reprendre)

**`feuille d'heures/recap_FH-TEST.xlsx` — NE PAS ENVOYER AUX RH.** C'est le fichier d'essai.
Son bloc **S34 contient des données inventées par Claude** pour tester la mécanique (les heures
sont exactes car reprises des vraies feuilles, mais la répartition Chaineau/Dambron de Hamouch et
Moussa est une reconstitution de la saisie manuelle de Patrice, pas une donnée d'origine). Seul
le bloc **S35 est réel** : Pascal Bonaventure et Ahmed Hamouch, passés par toute la chaîne.
Patrice avait envisagé de renommer ce fichier en « septembre » pour l'envoyer — à ne pas faire
tant que la S34 n'a pas été remise à blanc et resaisie.

**`feuille d'heures/recap_FH-septembre.xlsx`** est le vrai fichier de production. Ses étiquettes
de semaines ont été corrigées le 24/08 (`S34 S35 S36 S37 S38` au lieu de `S34 S36 S37 S38 vide`) ;
aucune donnée n'a été déplacée, seul le bloc S34 est rempli, c'est la saisie manuelle de Patrice.
Sauvegarde d'avant correction : `recap FH/recap_FH-septembre_sauvegarde-avant-etiquettes.xlsx`.

**`feuille d'heures/dépôts appli/`** contient encore les deux PDF de test
(`S35-HAMOUCH_Ahmed.pdf`, `S35-BONAVENTURE_Pascal.pdf`). Ils ont été intégrés dans le fichier de
test, pas dans le vrai récap. À déplacer dans `traité` ou à supprimer avant le déploiement, sinon
ils seront réintégrés lors de la prochaine commande sur la S35.

**Documents Word remis à Patrice le 24/08** (sur son Bureau, régénérables à la demande) :
- `Liens_App_Techniciens_TELSAM.docx` — les 13 techniciens et leur lien personnel cliquable, à
  diffuser au moment du déploiement. Le code d'accès partagé n'y figure volontairement pas.
- `Chantiers_Numeros_Index_TELSAM.docx` — les 68 chantiers avec leur n° d'index, groupés par
  priorité de renommage Dropbox (A : techniciens déjà placés, B : fenêtre à venir, C : sans date,
  D : terminés).

## Décisions prises par Patrice (à respecter, ne pas rouvrir sans qu'il le demande)
- Garder la structure Dropbox actuelle (un "App Tech" par dossier de chantier) plutôt que de
  créer un dossier parent unique déjà public — accepte de refaire le clic manuel à chaque
  nouveau chantier plutôt que de sortir les docs techniciens de leur dossier respectif.
- "Aure - Loudenvielle" (26-024) reste distinct, pas d'autre fusion à chercher sur ce nom.
- Les adresses issues de NDS sous-traitant (EDEA pour Curbans-Sisteron, EQOS pour Bruges-Galus)
  sont valides, pas besoin de NDS TELSAM propre.
- Les ICP (visites préalables) ne justifient jamais la création d'une fiche chantier.
- **Le cumul des heures par chantier démarre à mai 2026** (26/08/26). Avant mai, les chantiers
  n'étaient pas renseignés sur le récap FH : janvier à avril ne seront pas intégrés, et il ne faut
  pas les reproposer. Conséquence à garder en tête : un chantier ouvert en début d'année paraîtra
  sous-chargé, ce n'est pas une anomalie.

### Session du 25/08/26 — surveillance automatique, portée des PDP, chantiers multi-lots

Journée dense. Le détail des règles est dans `CLAUDE.md` — ne pas le dupliquer ici.

**Ce qui a été construit**

- **Veille documentaire quotidienne** (`scripts/veille-documents.ps1`, tâche Windows « TELSAM -
  Veille documents RTE », 07h30). Balaie les 142 000 fichiers du Dropbox local en ~20 s, repère
  les PDP/PGO/IST/NDS/PPSPS nouveaux et les dossiers App Tech en retard. Écrit dans
  `Desktop\TELSAM-apps\veille\`. **En lecture seule**, ne modifie jamais Dropbox.
  La voie cloud, échouée en août, est définitivement abandonnée : tout le Dropbox est déjà
  synchronisé sur le disque, aucun connecteur n'est nécessaire.
- **Injection automatique du rapport au démarrage de session** (hook `SessionStart` dans
  `TELSAM-apps/.claude/settings.json` → `scripts/hook-veille-session.ps1`). Répond à la question
  de Patrice « qu'est-ce que tu appelles début de session ? » : la consigne écrite dans CLAUDE.md
  dépendait de la mémoire de Claude, comme la règle `SEED_VERSION` oubliée deux fois. Désormais
  c'est l'outil qui l'exécute. **Vérifié en conditions réelles** : le rapport est bien arrivé seul
  à l'ouverture de la session suivante.
- **Contrôle de couverture PDP par périmètre** (`perimetre` + `pdps` + `computePdpAlerts`).
  Un PDP ligne n'ouvre pas la porte d'un poste — sans PDP couvrant le poste, l'équipe est refoulée
  alors que la fiche affichait « PDP OK » en vert.
- **Vue par semaine alignée sur le Gantt** : les deux lisent maintenant la même source
  (`joursPresenceReels`, union de `REAL_DAYS` et `TECH_RANGES`). La S35 est passée de 9 chantiers
  affichés à 3 — Portet, DATA4, Audit Multi Postes.
- **Chaineau-Cordy-Lamotte scindé** en `26-036-1` (RODA) et `26-036-2` (SELT), chacun avec son
  dossier App Tech, son brief, sa NDS, son bouton photos et son lien public vérifié.

**Chantiers traités**

- Fleyriat (26-055) : PGO passé en indice 5. Tests Phase 1 avancés d'une semaine (31/08 au lieu du
  07/09) et deux consignations avancées. App Tech mis à jour.
- Portet (26-051) : IST 63 kV indice A transmise à RTE, tableau de validation encore vierge →
  statut `warn`, document volontairement PAS dans App Tech. Le poste 225 kV, où le technicien
  travaille, ne nécessite pas d'IST. PGO App Tech corrigé de V37 en V40 (deux semaines de retard).
- DATA4 (26-065) : **accès au poste de Villejust non couvert par un PDP**, alors que le devis
  prévoit une recette qui s'y termine. Alerte rouge active, laissée en l'état sur demande de
  Patrice. Un PPSPS y est aussi exigé par la VIC et n'existe pas.
- Aure-Loudenvielle (26-024) : MTFO annulée, rappel et alerte retirés, fiche conservée.

**Erreurs commises et corrigées**

- **Encodage cassé et poussé.** Un bump de `SEED_VERSION` fait avec `Get-Content -Raw` sans
  `-Encoding UTF8` a doublé l'encodage de tout `suivi_chantiers_205.html` : 4 202 séquences
  abîmées, plus un seul accent correct, commit `df964c0` poussé dans cet état. Réparé en
  `055aa2c`, avec vérification que les 67 fiches non concernées étaient identiques à la version
  saine. **Règle et procédure de réparation écrites dans CLAUDE.md.**
- **Généralisation abusive** : j'avais écrit que « sur un audit, le PDP s'établit sur place »
  comme si c'était une règle. Patrice a corrigé : c'est le cas de ce chantier-là, pas une
  généralité. Reformulé, et le marqueur `pdpSurPlace` est explicitement « à poser au cas par cas,
  jamais déduit ».
- **Explications trop techniques** : Patrice a dû me demander deux fois de reprendre en langage
  simple. À retenir : expliquer d'abord ce qui change pour lui, sans jargon, et ne donner le
  détail technique que s'il le demande.
- Une tâche de fond laissée tourner à vide 92 minutes après avoir tué le processus Word dont elle
  dépendait. Nettoyer derrière soi.

**Documents remis à Patrice** (sur son Bureau, régénérables)
- `Memo_Exploitation_TELSAM.docx` — qui fait quoi, à jour au 25/08.
- `Appli_TELSAM_Consigne_Technicien.docx` et `.pdf` — à joindre au mail de mise en service.
  L'ancien `Appli_TELSAM_Mode_emploi` a été supprimé du Bureau pour éviter d'envoyer la mauvaise
  version.

### Session du 26/08/26 — écran d'ouverture animé (appli-techniciens)

Demande de Patrice : un logo animé à l'ouverture, rejoué à chaque mise à jour. Trois propositions
lui ont été montrées sur maquette (Fibre / Sobre / Réseau) ; il a retenu **Réseau**, puis a eu
l'idée qui a fait la version finale : **les deux petites flèches du « s » du logo ne sont pas
décoratives, ce sont le début et la fin du trait.** L'animation trace donc le « s » entre elles
avant de faire apparaître le logo complet autour.

Détail des règles dans `CLAUDE.md`, section « Écran d'ouverture animé » — ne pas le dupliquer ici.
Ce qu'il faut retenir :
- **`APP_VERSION` est à incrémenter à chaque mise en ligne**, au même titre que `SEED_VERSION`
  côté suivi. C'est ce qui rejoue l'animation, et l'animation est le seul signal qui dit au
  technicien qu'il a bien reçu la nouvelle version et non l'ancienne restée en cache.
- Le tracé du « s » est un **relevé au pixel du logo** : si le logo change, refaire le relevé.

**Deux erreurs commises, toutes deux invisibles à un test automatique.**
- Le « s » retombait **entre le « a » et le « m »** au lieu de sa place. Repéré par Patrice à
  l'œil, alors que mon contrôle chiffré affichait « 0,2 pixel d'écart ». Le contrôle était faux :
  je comparais le tracé au logo **dans le même repère**, ce qui était vrai d'avance, au lieu de
  vérifier que le calque du tracé se superposait à l'image. Leçon à garder : **mesurer en
  coordonnées d'écran, et se demander si le contrôle peut seulement échouer.** Un test qui ne
  peut pas échouer ne prouve rien.
- Le porteur du logo avait une hauteur nulle tant que l'image n'était pas chargée, donc le « s »
  se dessinait décalé en début d'animation. Corrigé par `aspect-ratio` + les attributs
  `width`/`height` sur l'image.

Vérifié dans l'appli réelle, servie en local, à 375×812 : première ouverture (« Bienvenue »),
deuxième ouverture (pas d'animation), changement de version (« Mise à jour installée »), tape
pour passer, lien personnel `?tech=` qui arrive bien sur la fiche du technicien, zéro erreur
console, alignement du « s » à 0 pixel.

**Mis en ligne** : `appli-techniciens` `d410ddd` puis `926153c`, `suivi-chantiers` `b21176e` puis
`c872195`. `APP_VERSION` en ligne = `2026-08-26-1`.

**Garde-fou ajouté** (`scripts/check-app-version.sh`, installé en `.git/hooks/pre-commit`) :
bloque un commit qui touche `index.html` sans incrémenter `APP_VERSION`, et vérifie le format
`AAAA-MM-JJ-n` puisque la date montrée au technicien en est déduite. Les trois cas testés :
oubli bloqué, format cassé bloqué, bump correct accepté. Même remède que pour `SEED_VERSION`,
où la consigne écrite seule n'avait pas suffi.

**Question de Patrice en fin de matinée, à ne pas oublier :** il rechargeait la page avec la
flèche du navigateur et ne revoyait pas l'animation. **C'est le comportement voulu**, pas un
défaut — recharger n'est pas une mise à jour. Refait devant lui les trois étapes d'affilée pour
le prouver plutôt que l'affirmer. Règle écrite dans `CLAUDE.md` avec la mention explicite de ne
jamais « réparer » ça.

**Deuxième demande de Patrice, en fin de matinée :** l'animation dit « quelque chose a changé »
sans dire quoi. Une vraie notification de téléphone est impossible — l'appli n'a aucun serveur,
zéro appel réseau, c'est ce qui la rend increvable hors couverture. Retenu à la place : une
phrase `APP_NOUVEAUTE` affichée sous la date, réécrite à chaque mise en ligne. **Patrice a
tranché deux points** : plus de tape pour passer l'animation, et rien d'autre que l'écran
d'ouverture (pas de bandeau qui reste sur les chantiers). Comme la phrase n'aurait alors été
lisible qu'une demi-seconde, l'écran se termine désormais par un bouton **« Continuer »** :
il lit à son rythme et décide quand passer. Filet de sécurité de 30 s pour que personne ne
puisse rester bloqué. Le garde-fou refuse un bump d'`APP_VERSION` sans phrase nouvelle.

**Consigne technicien mise à jour** (`Appli_TELSAM_Consigne_Technicien.docx` + `.pdf` sur le
Bureau, celle qui part avec le mail aux 13) : la puce sur le cache a été remplacée par une puce
qui explique l'animation et en fait le signal — *« si tu ne vois PAS l'animation après une mise à
jour annoncée, tu as encore l'ancienne version »*. Document modifié en place, mise en forme
conservée, sauvegarde de la version précédente faite avant. Piège rencontré : dans ce document
la puce « • » fait partie du texte, pas du format — remplacer un paragraphe entier la fait
disparaître.

### Session du 26/08/26 (suite) — heures par chantier : le cumul ne se faisait pas

Patrice a demandé confirmation que les heures des récaps FH mensuels étaient bien cumulées
chantier par chantier. **Elles ne l'étaient pas.** Les trois récaps (juin, juillet, août) avaient
été chargés en une seule fois le 19/08/2026, à la création de `suivi_chantiers_205.html`, et rien
n'avait bougé depuis — l'historique Git le confirme (aucun commit ne touche `heuresTelsam` en
dehors de la scission de Chaineau). Aucune règle n'existait pour dire « à chaque envoi, mettre à
jour ». Elle existe maintenant : `CLAUDE.md`, section « Heures cumulées par chantier ».

**Deux erreurs trouvées en vérifiant** (en comparant le suivi aux onglets `Chantiers` des récaps,
ligne à ligne) :
- **Chaineau comptait double, 70 h fantômes.** Le récap distingue les lots : 42 h en juin sur le
  lot 2, 28 h en août sur le lot 1. La scission du 25/08/26 avait recopié les deux montants sur
  les deux fiches. Corrigé : lot 1 = 28 h (août), lot 2 = 42 h (juin) + 63 h (septembre).
- **Poste de Portet : 819 h enregistrées contre 817 h au récap.** Deux heures de trop, sans
  explication. Aligné sur le récap.

**Récap de septembre intégré** (374 h sur 6 chantiers) — c'est la période **en cours** (S34-S38),
donc un chiffre partiel qui sera réécrit aux prochains envois. Dambron - Voves n'avait aucune
ligne d'heures : elle a été créée. `SEED_VERSION` passée de `v97-2026-08-25i` à `v98-2026-08-26`.

**Contrôle final, écart zéro sur les quatre périodes** : juin 1905, juillet 1417, août 1335,
septembre 374 — chaque total égal à la somme de l'onglet `Chantiers` du récap correspondant,
lignes « Agence » déduites (14 h en juin, 63 h en juillet, non rattachables à un chantier).

**Découverte à retenir : les récaps FH sont des périodes de paie, pas des mois calendaires.**
Juin = S21-S25, juillet = S26-S29, août = S30-S33 (soit 20/07 → 16/08), septembre = S34-S38.
D'où Anneyron - St Vallier, terminé le 28/07, qui porte 112 h « en août ». Détail et tableau dans
`CLAUDE.md`.

### Session du 26/08/26 (suite) — quantités de boîtes WTC2 relevées dans les devis

Patrice préparait ses achats de boîtiers WC2 et sa location de nacelles. Le suivi ne chiffrait
que 5 chantiers ; les quantités réelles sont dans les devis, article par article. **19 devis lus**
(`Telsam Fibre\RTE`, dossiers `DEVIS` de chaque chantier).

Résultat : **173 boîtiers WTC2 chiffrés**, dont 29 déjà posés (phases 2025 de Lannemezan, 9
d'Arcomie en juillet) → **144 restent à poser**, plus **2 boîtiers de jonction OPPC** à Fleyriat
(pylônes 125 et 135). Export livré : `Boites_WTC2_Reste_A_Poser.xlsx` sur le Bureau, deux onglets
(reste à poser / points à trancher). Page de synthèse publiée en Artifact.

**Règle métier confirmée par les devis : les boîtiers WTC2 se posent en pied de pylône, donc sans
nacelle.** La nacelle sert pour l'OPPC, les tores PMC, les transitions isolantes et les LOP ADSS.
Elle est chiffrée sur Fleyriat (4x4, à notre charge), Jonquières-St Césaire (51 m), Arcomie et
Bruges-Galus ; elle est explicitement **exclue** du devis Colayrac-Gupie.

**Sept écarts entre documents signalés et non arbitrés** (dans l'onglet 2 de l'export) : Bradascou
7 ou 10, qui commande les boîtes OPPC de Fleyriat, le pylône 91 manquant chez Lannemezan INEO, le
pylône 55 compté deux fois entre les deux lots Lannemezan, les portiques de Chaineau lot 1, le
devis complément Colayrac déjà absorbé par le V5.

**Outil découvert : lire un PDF sans poppler.** `pdftoppm` n'est pas installé, donc l'outil Read
ne peut pas ouvrir les PDF. La parade qui marche : `Documents.Open($pdf)` en COM Word (conversion
PDF → texte intégrée à Word), puis `Content.Text`. Attention, `Content.Text` sur le `.doc`
d'origine ne suffit pas : le tableau du devis est un objet Excel incorporé, invisible pour lui —
il faut passer par le PDF.


## AVANT LA MISE EN SERVICE AUX 13 TECHNICIENS — reste à faire

Dans l'ordre où ça se fait :

1. **Affecter les techniciens de la semaine prochaine sur Chaineau** : dire à Claude qui va sur le
   lot 1 (RODA) et/ou le lot 2 (SELT). Sans ça le chantier n'apparaît pas dans leur appli.
2. **Remettre à blanc les blocs S34 et S35** du récap RH, sauvegarde d'abord. Claude le fait sur
   signal de Patrice. Objectif : galop d'essai grandeur nature, les 13 saisissent ces deux
   semaines. Patrice garde les feuilles papier de la S34, tout est rattrapable.
3. ~~Sortir les deux PDF de test du dossier `dépôts appli`~~ — **FAIT le 25/08/26.** Patrice les a supprimés : il avait reçu par ailleurs les feuilles Excel habituelles et les a rangées dans son dossier `feuille d'heures/S34`. Le dossier de dépôt ne contient plus que son `_LISEZ-MOI.txt`, ce qui est l'état correct. Les données de ces deux PDF n'avaient été écrites que dans `recap_FH-TEST.xlsx`, jamais dans le vrai fichier : rien n'est perdu.
   **À retenir pour le galop d'essai** : Patrice détient **11 feuilles Excel de la S34** (S34-BENZAMERA, BONAVENTURE, DENIS, DIRAT, EL ABBASSI, HAMOUCH Bilal, MOUSSA, PERRIN Didier/François/Vincent, Viry). Elles servent de référence pour vérifier que les techniciens ressaisissent bien la même chose dans l'appli après la remise à blanc.
4. **Demander à Ahmed Hamouch et Pascal Bonaventure d'effacer leur semaine d'essai** sur leur
   téléphone (bouton « Effacer et recommencer cette semaine »). Personne ne peut le faire à leur
   place : la saisie vit dans leur navigateur.
5. **Envoyer le mail** aux 13, avec le lien personnel de chacun
   (`Liens_App_Techniciens_TELSAM.docx`) et `Appli_TELSAM_Consigne_Technicien.pdf` en pièce
   jointe. Le code d'accès se transmet par un autre canal, il n'est dans aucun des deux documents.
   Prévenir que l'application Dropbox, si elle est installée, intercepte les liens et réclame une
   connexion : appui long → « ouvrir dans le navigateur », ou la désinstaller (cas Pascal).

## Points ouverts, sans urgence

- **DATA4** : accès au poste de Villejust non couvert, et PPSPS manquant. Alerte active.
- **Chaineau** : le PDP en vigueur est un « Projet … indice 2 ». La fiche annonçait un indice 3
  sous la référence `PP26-SOL-EEL-002D`, introuvable dans Dropbox. Si un PDP signé arrive sous
  cette référence, il remplace le projet.
- **Trois App Tech pas à jour**, sans technicien planifié donc sans urgence : PDP de
  Verney-St Guillerme (Ind 7 au lieu de Ind 8), PDP de Lamativie-La Mole, NDS de Fibrage Feyriat.
  Patrice doit d'abord vérifier auprès des chargés de travaux si ces chantiers sont terminés.
- **Périmètre PDP** renseigné sur les 3 chantiers actifs seulement. À compléter chantier par
  chantier au moment où chacun devient actif. Les six chantiers où Dropbox montre déjà un
  découpage poste/ligne sont les prochains candidats : Bagatelle-Issel, Bédarieux-Espondeilhan,
  Issel-Revel, Sèvres-St Vallier, Fibrage FO DI Lyon, Givors.
- **Renommage Dropbox** avec les n° d'index : en cours côté Patrice. La veille le tolère
  désormais sans produire de fausses alertes.
- **Habilitations** : dossier partagé à tous, chacun voit celles des autres. Signalé, pas traité.


## Ce qu'il restait à faire au 24/08/26 (repris ci-dessus, conservé pour l'historique)
- ~20 chantiers sur les 69 (suivi-chantiers) / 67 après fusions n'ont toujours pas de `baseVie`
  trouvée — non prioritaire, à reprendre seulement si Patrice le demande.
- Numérotation Dropbox `26-0XX` : Patrice la met à jour progressivement de son côté ; une fois
  généralisée, la recherche de dossiers/documents pourra se faire par numéro plutôt que par nom
  approximatif (plus fiable).
- Aucune IST non signée RTE identifiée à ce jour → aucune alerte de ce type encore créée ; la
  règle est prête à s'appliquer dès qu'un cas se présente.
- Pas de vérification en attente côté "documents App Tech périmés" pour les 6 dossiers déjà
  existants avant cette session (Portet, Fleyriat, Bissy, St-Guillerme, Lamativie, Hospitalet) —
  la règle de mise à jour continue s'applique surtout à partir de maintenant, pas de passage
  rétroactif fait sur ces 6-là.
- Rien d'autre en attente de validation à ce stade — tout ce qui a été discuté a été commité et
  poussé sur les deux dépôts.

- **Déploiement prévu la semaine prochaine** aux 13 techniciens. Prévoir de remettre à blanc les
  blocs S34 et S35 du récap (sauvegarde d'abord) et de demander à tous de saisir ces deux
  semaines, pour un galop d'essai grandeur nature. Patrice conserve les feuilles de la S34 dans
  son dossier dédié, il peut donc rattraper toute erreur.
- Ahmed Hamouch et Pascal Bonaventure doivent effacer leur semaine de test sur leur téléphone
  (bouton dédié) — la saisie vit dans leur navigateur, personne ne peut l'effacer à distance.
- Renommage progressif des dossiers Dropbox avec les n° d'index, et annotation du Planning RTE
  avec ces mêmes numéros (Patrice s'en charge) — voir le document
  `Chantiers_Numeros_Index_TELSAM.docx` remis le 24/08.
