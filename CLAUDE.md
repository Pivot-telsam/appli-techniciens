# Contexte TELSAM — appli-techniciens & suivi-chantiers

## Structure de travail
- `appli-techniciens/` : dépôt Git public, déployé sur GitHub Pages
  (https://pivot-telsam.github.io/appli-techniciens/). Fichier unique `index.html`.
- `suivi-chantiers/` : dépôt Git privé, PAS déployé. Contient `suivi_chantiers_205.html`
  (nom de fichier à garder tel quel, ne pas renommer/versionner par numéro — l'historique Git suffit).
  Patrice envoie ce fichier par Teams à ses collègues après chaque mise à jour.

## Rôle de chaque appli
- `appli-techniciens` : consommée par les techniciens terrain (lien perso + phrase de passe partagée).
  Affiche pour chaque techniciens ses chantiers de la semaine + documents + feuille d'heures.
- `suivi-chantiers` : source de vérité détaillée (statuts PDP/PGO/IST, alertes, Gantt, historique).
  `appli-techniciens` ne fait que consommer un sous-ensemble allégé de ces données.

## Champs protégés (JAMAIS écrasés lors d'une synchro depuis suivi-chantiers)
Dans `appli-techniciens/index.html`, ces champs par chantier n'existent que là et doivent toujours
être préservés/fusionnés, jamais régénérés depuis suivi-chantiers :
- `documentsAppTech` (lien Dropbox vers sous-dossier "App Tech", lecture seule technicien)
- `depotTerrain` (lien Dropbox File Request, dépôt photos sans compte)
- `client`
- `baseVie` (adresse base de vie/dépôt, utilisée en priorité par le bouton Itinéraire ;
  sinon fallback sur `geo.points[0]`)
- `documentsTerrain` (ancien système pdp/pgo/briefing, fallback pour chantiers non migrés —
  **en cours d'extinction, cf. règle ci-dessous**)

## Convention de nommage des chantiers — RÈGLE IMPORTANTE
Un nom de chantier n'est fiable que s'il vient :
1. d'un PDP ou devis TELSAM signé (source la plus fiable), ou
2. d'un dossier Dropbox déjà existant avec ce nom.
Un nom issu uniquement d'une ligne du Planning_RTE_2026.xlsx brut est PROVISOIRE — ne jamais
créer/renommer une fiche définitivement sur cette seule base sans demander confirmation à Patrice.

## Structure Dropbox (pour recherche de documents/adresses)
Toujours : `telsamfibre > RTE > [Ligne aérienne | Ligne souterraine | Postes | Audit FO | ARTERIA]`
puis dossier du chantier, puis sous-dossier `Documents Telsam` (contient les NDS).
Ne jamais sortir de cette arborescence pour ce type de recherche.

Précisions vérifiées le 25/08/26 (l'ancienne version de cette liste était inexacte) :
- **Il n'y a PAS de catégorie `Fibrage`.** « Fibrage » est un préfixe de nom de chantier à
  l'intérieur de `Ligne aérienne` : `Fibrage Bissy`, `Fibrage Feyriat`, `Fibrage FO DI Lyon`,
  `Fibrage Pressy`. Ne pas chercher un dossier `RTE\Fibrage`, il n'a jamais existé.
- Sous **`Audit FO`**, un niveau intermédiaire de direction interrégionale s'intercale avant le
  chantier : `Audit FO > DI Marseille | DI Toulouse > <chantier>`.
- `préparation documents sécurités` (premier niveau, hors arborescence chantier) sert de zone de
  passage : des IST y transitent, parfois déjà signées par RTE, avant d'être classées dans le
  dossier du chantier. À regarder si une IST attendue reste introuvable ailleurs.

Sous-dossiers frères possibles directement sous le dossier du chantier (confirmé 21/08/26) :
- `PDP` et `PGO` — c'est ICI que Patrice dépose les nouveaux PDP/PGO reçus, PAS dans "App Tech"
  ni dans "Documents Telsam". Un nouveau fichier déposé là ne se propage JAMAIS automatiquement
  vers "App Tech" — il faut activement l'y copier (en remplaçant l'ancienne version) à chaque
  mise à jour, cf. règle App Tech ci-dessous.
- `App Tech` — dossier curé à part, alimenté par copie depuis PDP/PGO/Documents Telsam, jamais
  la source primaire des documents.
Si je découvre un document par moi-même (recherche Dropbox pour une autre raison, pas signalé par
Patrice), penser à regarder aussi dans `PDP`/`PGO`, pas seulement dans "Documents Telsam" ou
"App Tech" — sinon une mise à jour peut être manquée.

**PIÈGE — orthographe instable des noms de dossiers.** Le nom d'un dossier Dropbox n'est pas une
source fiable pour identifier un chantier : la même affaire s'écrit différemment selon les
dossiers. Cas confirmé par Patrice le 25/08/26 : le chantier **Fleyriat** (26-055) a son dossier
racine orthographié `Fibrage Feyriat` (**sans le L**) alors que tous les documents RTE à
l'intérieur disent bien FLEYRIAT. Conséquences pratiques :
- Toujours chercher avec un motif tolérant (`fleyria|feyria`, jamais l'orthographe exacte seule),
  sinon un document existant ressort comme absent.
- Ne jamais conclure "chantier introuvable dans Dropbox" sur la base d'une seule graphie.
- C'est précisément ce que le **n° d'index** (`26-0XX`) vient résoudre : une fois le dossier
  renommé avec son préfixe, l'identification devient exacte et cette ambiguïté disparaît. Un
  dossier déjà renommé doit être identifié par son numéro, pas par son libellé.
- Ce chantier a aussi **deux dossiers `PGO` parallèles** (un au niveau racine `Fibrage Feyriat`,
  un dans le sous-dossier `Brou - La Cluse - Feyriat Roda`) : les nouveaux indices arrivent dans
  le second. Regarder les deux avant de conclure.

## Planning RTE 2026.xlsx — RÈGLE CRITIQUE (affectations technicien/chantier)
Fichier `C:\Users\patrice.pivot\Desktop\Planning RTE 2026.xlsx`, feuille "Feuil1" :
- 370 colonnes = 1 jour par colonne à partir du 1er janvier (col2 = 1er janvier ; formule :
  colonne = jour_de_l'année + 1). Lignes 1-4 = en-têtes (mois/semaine/lettre jour/n° jour).
- Lignes ~5-22 = un technicien par ligne (nom en colonne A).
- Lignes ~23 et suivantes = des **lignes-projet SANS nom en colonne A**. Chaque ligne-projet a
  une **couleur de fond** et un texte décrivant le chantier/la tâche, écrit UNE SEULE FOIS le
  lundi de la semaine où elle démarre/change (les jours suivants de la même semaine ont la même
  couleur mais pas de texte répété).
- **L'affectation technicien → chantier se lit en comparant la couleur de fond** de la cellule
  du technicien avec celle des lignes-projet le même jour — PAS le texte de la cellule
  technicien, qui est très souvent VIDE même quand le technicien est affecté. Un technicien
  affecté à un chantier a donc une cellule colorée sans texte ; il faut retrouver la ligne-projet
  de la même couleur ce jour-là pour savoir sur quel chantier.
- Piège vécu (20/08/26) : une première lecture qui ne regarde que le texte des cellules
  technicien peut conclure "personne n'est affecté" alors qu'en réalité plusieurs le sont
  (seule la couleur le montre) — erreur à conséquences réelles (sécurité, planification).
  Toujours vérifier via `Interior.Color` (PowerShell + Excel COM, cf. [[feedback_environment_no_python]])
  pour CHAQUE cellule de la semaine, technicien ET lignes-projet, avant de conclure qu'une
  semaine est vide.
- Autres couleurs rencontrées : CP/congé et CP Paternité = fond rouge (couleur COM 255) — jours
  non travaillés, à ignorer pour les affectations. "MTS longue distance" est un rappel de
  matériel intégré au texte d'une ligne-projet (ex. Audit Multi Postes), pas une tâche séparée —
  toujours vérifier si ce texte fait partie d'une ligne-projet plus large avant de l'ignorer ou
  de le traiter isolément.
- Les ICP (visites préalables) ne justifient pas la création d'un chantier.
- `REAL_DAYS`/`TECH_RANGES` (jours réels confirmés par ce code couleur) n'existent QUE pour les
  semaines où le planning fournit une couleur exploitable — absents pour les fenêtres issues de
  PGO/devis seuls. Une absence d'entrée ne veut donc pas dire "personne d'affecté", juste "pas de
  source colorée disponible" pour cette période.
- Modèle sous-jacent : un technicien n'a qu'une seule couleur (= un seul chantier) par jour dans
  le planning. Ne jamais lui affecter deux chantiers le même jour dans `TECH_RANGES`.

## Pièges connus dans le code (à ne pas "corriger" par erreur)
- **`SEED_VERSION` (suivi-chantiers, ligne ~1033) — ÉTAPE OBLIGATOIRE À CHAQUE MODIF DE `SEED_DATA`.**
  Au chargement, si la version stockée en IndexedDB == `SEED_VERSION`, l'app ne touche à RIEN et
  garde les données déjà en cache — même si `SEED_DATA` a changé dans le fichier. Oublier de
  l'incrémenter rend toute modification invisible pour quiconque a déjà ouvert une version
  antérieure du fichier (vécu le 20/08/26 : plusieurs commits d'affilée sans bump — corrigé après
  coup). À l'inverse, changer la version efface tout état local resté seulement dans l'UI et
  jamais reporté dans `SEED_DATA`. **Réflexe : incrémenter `SEED_VERSION` à la fin de CHAQUE
  session de modifications de `SEED_DATA`, juste avant de committer.**
- **CE QUE PATRICE SAISIT DANS L'APPLI NE VA PAS DANS LE FICHIER — à lui redire, jamais l'oublier.**
  `saveEdit()` → `saveChantier()` → `storeSet()` écrit dans **IndexedDB**, c'est-à-dire dans le
  navigateur de la machine où il travaille. Le fichier `suivi_chantiers_205.html`, celui qu'il
  renvoie à ses collègues, n'en sait rien. Trois conséquences, dans l'ordre de gravité :
  1. ses modifications **sont effacées** au prochain bump de `SEED_VERSION`, parce que
     `ensureSeeded()` fait `chantiers = {}` puis recharge tout depuis `SEED_DATA` ;
  2. ses collègues ne les voient jamais ;
  3. elles n'existent pas sur une autre machine ni dans un autre navigateur.
  Le formulaire de modification sert donc à **essayer**, pas à tenir des comptes. Signalé à Patrice
  le 26/08/2026, le jour où le champ `boites` a été ajouté avec sa saisie — il avait justement
  demandé « je pourrai modifier au fur et à mesure et ça s'enregistrera ? ». La réponse est non.
  **La seule voie durable : Patrice dit le chiffre, Claude l'écrit dans `SEED_DATA`, bump, commit.**
  Ne jamais lui proposer de saisir lui-même quelque chose qui doit durer ou être partagé.
- Le PGO peut afficher un statut global "ok" tout en ayant un trou de couverture précis sur une
  semaine où une équipe est présente (cas réel Fleyriat/Lisieux, 18/08/26) — toujours croiser
  `pgo.statut` ET les dates précises de `pgo.couverture`, jamais se fier au seul statut global.
- Statut `nc` (aucun PDP/PGO reçu) est le cas le plus grave, sauf pendant une phase de mesure de
  touret/MAP où c'est normal de ne pas encore en avoir — ne pas ignorer `nc` ailleurs qu'à ce stade.
- La vue Gantt (suivi-chantiers) masque les semaines passées par défaut ("Afficher les semaines
  passées" décoché) — une fenêtre/segment entièrement passé est invisible tant qu'on ne coche pas
  cette case, même s'il est bien présent dans les données (vécu avec le PGO Bradascou, 20/08/26).
  Toujours vérifier avec la case cochée avant de conclure qu'une donnée manque.
- **La barre bleue "Intervention TELSAM prévue" du Gantt est construite à partir de `REAL_DAYS`**
  (jours confirmés par le code couleur du planning RTE), PAS à partir du texte des `fenetres`
  (corrigé le 21/08/26 — avant ça, une fenêtre à plage large affichait une "présence" même sans
  aucun technicien réellement confirmé cette semaine-là : Cantegrit, Fleyriat, Joncquiere,
  Lisieux, Chaineau-Cordy-Lamotte en ont été victimes ; à l'inverse, un chantier sans champ
  `fenetres` du tout — ex. Audit Multi Postes — n'affichait jamais rien même avec des
  techniciens confirmés). Le texte des `fenetres` ne sert plus qu'à enrichir le tooltip. La
  couverture PGO (teal) et la validité PDP (verte) restent indépendantes de `REAL_DAYS` — un
  PGO signé couvre une période même avant confirmation couleur, c'est une info complémentaire,
  pas redondante. Ne pas revenir à une logique basée sur le texte des fenêtres seul.
- **`joursPresenceReels(id)` est la source UNIQUE de la présence technicien, pour la vue par
  semaine ET pour le Gantt** (mis en place le 25/08/26). C'est l'**union de `REAL_DAYS` et de
  `TECH_RANGES`** : les deux viennent du même code couleur du planning RTE mais ne sont pas
  toujours mis à jour ensemble — constaté le 25/08/26, en S34 `TECH_RANGES` portait des
  techniciens sur 26-045/26-036/26-051/26-060 alors que `REAL_DAYS` ne contenait qu'Arudy
  (terminé), et en S33 `TECH_RANGES` portait six chantiers pour un `REAL_DAYS` vide. Lire l'un
  sans l'autre fait donc disparaître des présences réelles. Ne jamais réintroduire d'accès
  direct à `REAL_DAYS[c.id]` dans une vue : passer par cette fonction, sinon les deux vues
  redivergent.
- **La vue par semaine liste les chantiers par présence confirmée, pas par texte de fenêtre**
  (demande de Patrice, 25/08/26). Avant, `buildWeekBuckets` bucketait chaque `fenetres` datée :
  la S35 affichait 9 chantiers (Arnage, Cantegrit, Chaineau, Fleyriat, Lisieux, Joncquiere,
  Lannemezan en plus) alors que 3 seulement avaient des techniciens sur place — une fenêtre est
  souvent une plage large écrite une fois, elle ne prouve aucune présence. Même piège que celui
  corrigé sur le Gantt le 21/08/26. **Repli assumé et nécessaire :** la présence confirmée
  s'arrête à la dernière semaine saisie au planning ; pour toute semaine sans AUCUNE donnée
  couleur, la vue retombe sur les fenêtres et marque le bloc `PRÉVISIONNEL`. Sans ce repli
  toutes les semaines à venir seraient vides et la vue perdrait son utilité de planification.
  Une semaine ayant au moins un jour confirmé n'est jamais mélangée avec du prévisionnel.
  Attention en modifiant le rendu : un chantier confirmé peut n'avoir AUCUNE fenêtre
  correspondante (cas d'Audit Multi Postes, sans champ `fenetres`) — le test « ligne de mesure
  de touret » doit donc vérifier `labels.length > 0` avant le `.every()`, sinon un tableau vide
  renvoie `true` et le chantier s'affiche à tort comme une mesure de touret.
- **Tri du Gantt (`chronoSortKey`) : la présence technicien réelle prime toujours sur PGO/
  consignation/NIP** (corrigé le 21/08/26). Avant, un chantier seulement couvert par un PGO en
  cours (sans technicien placé, ex. Cantegrit) pouvait remonter au même niveau qu'un chantier
  où des techniciens sont réellement envoyés (Portet/DATA4/Audit) — contraire au but du Gantt
  ("voir où sont les techs" d'un coup d'œil). Les segments 'normal'/'tourets' (présence réelle)
  sont maintenant toujours classés avant tout autre type de fenêtre réelle.
- `MTFO_STANDALONE` (suivi-chantiers) : rappels MTFO sans fiche chantier associée, décision
  volontaire de Patrice pour les lignes du planning sans PDP/PGO à gérer. Ne pas créer de fiche
  chantier pour "régulariser" ces lignes.
- Côté `appli-techniciens` : les badges de statut PDP/PGO ne s'affichent que si le chantier a un
  `documentsAppTech` renseigné (nouveau système) — un chantier encore sur l'ancien système
  (`documentsTerrain`) n'affiche jamais ces badges même si `pdp`/`pgo` sont remplis. Absence de
  badge = chantier non migré, pas forcément donnée manquante.
- **Accès à l'appli : portail par mot de passe par technicien** (depuis le 27/08/2026, remplace
  `GATE_CODE_HASH`). Il protège l'ENTRÉE de l'appli, PAS les documents Dropbox — voir la section
  « Accès à l'appli — portail par technicien ». Ne jamais le présenter comme une protection des
  documents techniciens/PDP.

## Système de numérotation (`numero`, format `26-0XX`)
C'est le numéro de référence que Patrice utilise pour organiser Dropbox (il renomme progressivement
les dossiers avec ce préfixe, ex: "26-055 - Fleyriat - Brou - La Cluse"). Pour tout nouveau chantier,
attribuer le prochain numéro disponible et le communiquer à Patrice.

Depuis le 24/08/26, ce numéro est présent dans **les deux dépôts** (68 fiches sur 68) : il sert
aussi d'identifiant du chantier dans le récap RH, à la place du n° de devis (cf. section Feuilles
d'heures). Objectif de Patrice : noter ce même numéro dans le Planning RTE à côté de chaque
chantier, pour que le rapprochement planning ↔ fiche devienne mécanique au lieu d'être fait à
l'œil. Renommage Dropbox et annotation du planning en cours de sa part, progressivement.

## Chantiers à plusieurs lots / sous-chantiers — RÈGLE (cas Chaineau-Cordy-Lamotte, 21/08/26)
Certains chantiers sont découpés en plusieurs lots attribués à des sous-traitants différents
(ex. Chaineau-Cordy-Lamotte : LOT 1 RODA / LOT 2 SELT dans Dropbox). Décision de Patrice (21/08/26) :
- Numérotation : un sous-chantier par lot, suffixe numérique sur le numéro parent — ex. `26-036-1`
  et `26-036-2` pour les deux lots de Chaineau-Cordy-Lamotte (parent `26-036`).
- **PDP et PGO sont communs** aux sous-chantiers d'un même chantier parent — un seul jeu de
  documents de sécurité, partagé, pas dupliqué ni divergent entre sous-chantiers.
- **NDS, IST et Brief technicien sont propres à chaque sous-chantier** — un jeu différent par lot
  (cohérent avec le fait que chaque lot a son propre sous-traitant, donc son propre périmètre de
  travaux techniciens).
- Donc chaque sous-chantier a son propre dossier "App Tech" (avec son propre Brief, sa propre IST
  si applicable, son propre "Photos terrain" + File Request + `depotTerrain`), mais le PDP et le
  PGO qui y figurent sont une copie du même document partagé entre les sous-chantiers du même
  parent — pas deux versions indépendantes qui pourraient diverger.
**APPLIQUÉ le 25/08/26** sur Chaineau-Cordy-Lamotte, une affectation étant possible la semaine
suivante. Convention de nommage Dropbox de Patrice, à reprendre pour les futurs cas :
- Dossier parent = **les deux numéros** : `Chaineau-Cordy-Lamotte 26-036-1-2`
- Un sous-dossier par lot = **son numéro** : `LOT 1 RODA 26-036-1`, `LOT 2 SELT 26-036-2`
- Le PDP et le PGO communs restent à la racine du parent (`PDP et  PGO`), les documents propres
  au lot dans le sous-dossier du lot.

Ce qui a été fait, à refaire à l'identique pour un futur multi-lots :
- La fiche unique `26-036` a été **scindée en `26-036-1` et `26-036-2`**. Les données communes
  (PDP, PGO, fenêtres, consignations, NIP, géo, base de vie) sont dupliquées telles quelles : même
  ouvrage, mêmes contraintes. Les tâches, elles, étaient déjà étiquetées par lot dans leur
  description (`[Lot Lamotte-Cordy-222]` = lot 1, `[Lot2 pyl222-poste Chaineau]` = lot 2) et ont
  été réparties exactement : 8 et 5.
- **L'identifiant `c_e5849v6z` a été conservé pour le lot 1**, pas régénéré : `TECH_RANGES` et
  `REAL_DAYS` le référencent, et créer deux nouveaux identifiants aurait fait disparaître
  silencieusement les affectations déjà enregistrées. La présence technicien de la S34 se retrouve
  donc rattachée au lot 1 — **rattachement confirmé exact par Patrice le 25/08/26** : ces
  techniciens étaient bien sur RODA. Point clos, ne pas le rouvrir. À retenir pour un futur
  multi-lots : conserver l'identifiant existant sur l'un des lots est le bon réflexe, mais toujours
  demander à Patrice à quel lot les affectations passées se rattachent au lieu de le supposer.
- Un dossier App Tech complet par lot (brief du lot, MO et NDS du lot, PGO ind.6 commun,
  `Photos terrain` + File Request + `depotTerrain`), et un `client` distinct par fiche
  (RODA / SELT France).

**Piège de lecture des noms de dossiers multi-lots** : `26-036-1-2` désigne le PARENT, pas le
lot 1. Deux suffixes ou plus = parent, un seul suffixe = ce lot (`NumeroDeDossier` dans
`scripts/veille-documents.ps1`). Sans ce cas particulier, l'extraction s'arrêtait au premier
suffixe et rattachait tout le chantier au lot 1.

## Contenu standard d'un dossier "App Tech" — RÈGLE SYSTÉMATIQUE (à appliquer TOUJOURS)
Chaque fois qu'un dossier "App Tech" est créé ou complété pour un chantier (nouveau chantier,
ou chantier existant qui devient actif — cf. workflow hebdomadaire ci-dessous), il doit contenir :
1. **Brief_Techniciens_RTE_<Nom>.pdf** — généré via le skill `brief-techniciens` à partir du
   devis TELSAM pertinent trouvé dans Dropbox. Jamais de prix/référence devis/conditions
   commerciales dans le brief, uniquement travaux + matériel. **Obligatoire à chaque fois**, pas
   seulement pour les nouveaux chantiers.
   - Ce skill s'appuie normalement sur ReportLab (Python) pour générer le PDF — **indisponible
     sur cette machine** (cf. [[feedback_environment_no_python]]). Solution de repli qui marche :
     construire le document via **Word (COM PowerShell)** avec la même charte visuelle
     (BLEU_F #1F4E79 titres, BLEU_M #2E75B6 sous-titre, marges 2cm = 56.7pt) puis
     `$doc.SaveAs2($path, 17)` (17 = wdFormatPDF) pour exporter directement en PDF.
2. **MO et NDS TELSAM** (si présents dans Dropbox pour ce chantier).
3. **PDP/PGO pertinents** (le plus récent indice / la version signée).
4. **IST — DEUX conditions, pas une : elle doit être une IST TELSAM *et* être signée/validée par
   RTE** (`ist.valideRTE == true`). Si une IST existe pour ce chantier mais n'est PAS signée par
   RTE : NE PAS la mettre dans App Tech, et à la place ajouter une alerte (niveau `warn`) dans
   `suivi-chantiers` signalant l'IST non signée par RTE, avec `ist.valideRTE:false`.

   **RÈGLE — une IST qui n'est pas de TELSAM n'entre JAMAIS dans App Tech, même signée par RTE.**
   Posée par Patrice le 27/08/2026. Une IST autorise **le personnel de l'entreprise qui l'a
   rédigée** pour une opération précise. Celle d'un co-traitant (INEO, Bouygues, Omexom, SELT,
   RODA…) ne couvre pas nos techniciens : la déposer dans App Tech ferait croire à un technicien
   qu'il est autorisé alors qu'il ne l'est pas. Les clients en envoient parfois **à titre
   d'exemple**, et elles atterrissent dans le dossier du chantier comme les autres.

   *Le cas qui a fait poser la règle* : le 27/08/2026 la veille signale deux IST sur Cantegrit
   26-003, dont une nommée « …signée RTE.pdf ». La règle de signature seule aurait dit « à
   déposer ». Lecture du document : **zéro occurrence de « TELSAM »**, une de « INEO », dossier
   `IST INEO`, n° de compte APS3T5108 (celui d'INEO). Ce sont des exemples du client.

   **Comment vérifier, dans l'ordre** : le nom du dossier et du fichier ; puis, toujours, ouvrir
   le document et y chercher « TELSAM » — une IST TELSAM nomme TELSAM. **En cas de doute, demander
   à Patrice. Ne jamais copier dans App Tech une IST dont l'auteur n'est pas identifié.**
5. **Sous-dossier "Photos terrain" vide** — toujours créé, pour que les techniciens y déposent
   leurs photos de chantier directement depuis l'app.
- Ne jamais inclure devis, contrats de sous-traitance, factures ou PV de réception commerciaux.
- Après avoir ajouté un Brief ou tout autre fichier : vérifier le lien de partage (cf. règle
  ci-dessous) avant de committer `documentsAppTech`.

**RÈGLE — le bouton "📷 Dépôt photos" (`depotTerrain`) fait partie intégrante de la création
d'App Tech, jamais une étape séparée qu'on remet à plus tard.** Incident vécu le 21/08/26 :
DATA4-Marcoussis et Audit Garies-Pessac avaient un dossier App Tech complet (avec "Photos
terrain" déjà créé) depuis plusieurs jours, mais sans File Request Dropbox associée ni champ
`depotTerrain` dans `appli-techniciens/index.html` — le bouton photo n'existait donc pas pour
les techniciens, sans que personne ne s'en rende compte jusqu'à ce que Pascal Bonaventure le
signale. Sur un audit du 21/08/26, seulement 2 chantiers sur 7 avec présence technicien réelle
avaient ce bouton opérationnel. Donc, à chaque création/complétion d'un dossier App Tech
(nouveau chantier, ou chantier existant qui devient actif) :
1. Créer le sous-dossier "Photos terrain" (déjà obligatoire ci-dessus).
2. **Dans le même geste**, créer une Dropbox File Request (`create_file_request`) dont la
   destination est ce sous-dossier "Photos terrain".
3. Câbler immédiatement le champ `depotTerrain` (URL de la File Request) dans
   `appli-techniciens/index.html` pour ce chantier — jamais laissé pour "plus tard".
4. Avant de clore une session de mise à jour hebdomadaire (cf. workflow ci-dessous), lister tous
   les chantiers ayant une présence technicien réelle cette semaine ou les 2 prochaines
   (`TECH_RANGES`) et vérifier qu'aucun n'a `documentsAppTech` sans `depotTerrain` — traiter ce
   contrôle comme aussi obligatoire que le bump de `SEED_VERSION`.

**RÈGLE — App Tech est le SEUL canal documentaire vers les techniciens (décision de Patrice,
25/08/26).** Les boutons individuels « 📄 PDP » / « 📄 PGO » / « 📋 Briefing » (champ
`documentsTerrain`) sont supprimés au fur et à mesure : un lien Dropbox pointe sur *un fichier*
précis, donc dès qu'un indice change le bouton devient mort ou, pire, ouvre silencieusement une
version périmée (constaté le 25/08/26 sur Portet : le bouton PGO pointait encore sur la V37 alors
que la V40 datait du 10/08). Le bouton « 📁 Documents » ouvre le dossier App Tech, dont le contenu
est tenu à jour — il n'a pas ce défaut. Donc :
- Ne JAMAIS recréer d'entrée `documentsTerrain` pour un chantier qui a un `documentsAppTech`.
- Avant de vider le `documentsTerrain` d'un chantier, vérifier deux choses : que les documents
  concernés sont bien présents dans son dossier App Tech, et que le lien App Tech est réellement
  `audience: "public"` (`get_shared_link_metadata`) — sinon on retire le seul accès qui marche.
- Un chantier sans `documentsAppTech` garde ses boutons tant qu'il n'est pas migré : les vider
  reviendrait à priver les techniciens de tout document.
- État au 25/08/26 : vidés pour Portet (26-051) et Givors (26-060). **26-062 Audit Bollène — Ste
  Cécile les Vignes — Carpentras** est le seul chantier restant sans dossier App Tech ni
  `depotTerrain`. **Décision de Patrice (25/08/26) : on n'y touche pas** — l'audit est terminé
  (intervention des 05-06/08/26), aucun retour sur site n'est prévu a priori. Ne pas proposer de
  le migrer spontanément ; si un retour se présente, en reparler à ce moment-là et créer alors son
  App Tech complet. Le code de rendu de `documentsTerrain` dans `appli-techniciens/index.html`
  reste donc nécessaire pour cette seule fiche — ne pas le supprimer comme du code mort.

**RÈGLE — tenir App Tech à jour dans la durée, pas seulement à la création.**
Dès que je prends connaissance d'un PDP, PGO ou IST nouveau/mis à jour pour un chantier qui a
déjà un dossier App Tech (que Patrice me le signale, ou que je le découvre en traitant autre
chose — ex. mise à jour du planning, recherche Dropbox) :
- Si ce chantier a un dossier App Tech existant, remplacer dans App Tech l'ancienne version du
  document par la nouvelle (supprimer l'ancienne, copier la nouvelle) — ne pas laisser une
  version obsolète à côté de la nouvelle.
- Une IST qui revient signée par RTE alors qu'elle n'y était pas encore : l'ajouter à App Tech à
  ce moment-là (elle ne l'était pas avant, cf. règle ci-dessus), et lever l'alerte `warn`
  correspondante dans suivi-chantiers + repasser `ist.valideRTE` à `true`.
- Mettre à jour en cohérence les champs structurés (`pdp`/`pgo`/`ist` indice, ref, dates) dans
  les deux dépôts, comme pour toute mise à jour de document déjà pratiquée (ex. Bradascou PGO).
- Le but : ce que le technicien ouvre dans App Tech doit toujours être la version en vigueur,
  jamais une version périmée qui traîne depuis la création du dossier.

## Feuilles d'heures : de l'appli au récap RH (mis en place le 24/08/2026)

**La chaîne, validée de bout en bout avec de vraies feuilles (Ahmed Hamouch et Pascal
Bonaventure, S35) :**
1. Le technicien remplit sa feuille dans l'appli et appuie sur « Envoyer ma feuille ».
2. Son téléphone partage **un seul fichier** : le PDF. Les données de saisie sont **cachées
   dedans** (propriété `Keywords` du PDF, préfixée `TELSAMDATA:`, JSON encodé en base64).
3. Patrice dépose ce PDF dans
   `TELSAM TEAM Dropbox/Patrice PIVOT/feuille d'heures/dépôts appli/` (un `_LISEZ-MOI.txt`
   y rappelle la marche à suivre).
4. Il demande « intègre la semaine NN ». Le script
   `suivi-chantiers/scripts/integrer-feuilles-heures.ps1` lit les dépôts, montre un **aperçu**,
   et n'écrit qu'avec `-Ecrire`.
5. Patrice relit, lance son skill `recap-feuilles-heures` pour régénérer les onglets 2 et 3,
   puis envoie aux RH.

**RÈGLE — un seul fichier à envoyer, JAMAIS deux.** Un `.json` séparé à côté du PDF a été
essayé le 24/08/26 : la plupart des téléphones ne savent pas partager deux fichiers d'un coup,
le partage natif échouait et l'appli retombait sur « télécharger les deux, ouvrir chacun, puis
partager » — trop de manipulations, les techniciens ne le feront pas (constaté avec Ahmed).
Une fois revenu à un seul fichier, le partage s'ouvre directement. Ne pas revenir en arrière.

**RÈGLE — le classeur n'a jamais qu'un seul écrivain.** Dropbox ne fusionne pas les `.xlsx` :
deux écritures concurrentes donnent une « copie en conflit », et Excel verrouille le fichier
quand il est ouvert. D'où le choix d'un fichier par soumission, jamais modifié, et d'une écriture
groupée déclenchée par Patrice quand le classeur est fermé. Le script refuse d'écrire si le
verrou `~$` existe, si un technicien est introuvable en colonne A, si une cellule contient déjà
une valeur différente (sauf `-Ecraser`), ou si la semaine demande plus de chantiers que le bloc
n'en a. Il est idempotent : le relancer ne duplique rien.

**Structure de l'onglet `Recap` — pièges.**
- Il ne s'agrandit pas vers le bas mais **vers la droite** : un « bloc » de colonnes par semaine,
  étiqueté en ligne 1 (`S34`, `S35`…), en-têtes en ligne 2, un technicien par ligne (3 à 15).
- **Les blocs n'ont pas tous la même largeur** : 2 ou 3 emplacements chantier selon les semaines.
  Donc **toujours repérer les colonnes par leur en-tête (ligne 2), jamais par leur position.**
  Erreur commise le 24/08/26 : copier des cellules d'un bloc à l'autre par position a envoyé le
  GD dans la colonne COM et perdu PD/Nacelle/COM — et la source avait été effacée avant
  vérification. Ne jamais effacer une source avant d'avoir confirmé que la destination a tout reçu.
- **Les étiquettes de semaines doivent être consécutives.** Le fichier de septembre 2026 sautait
  le `S35` (étiquettes `S34, S36, S37, S38, vide`), ce qui décalait tout et faisait écrire la S35
  dans le dernier bloc. Corrigé le 24/08/26. Vérifier ce point avant toute intégration.
- Ajouter une semaine **à droite** du dernier bloc est sans danger ; **élargir un bloc existant**
  (technicien avec 4 chantiers alors que le bloc en prévoit 3) impose d'insérer des colonnes au
  milieu et décale tout : le script s'arrête et prévient plutôt que de le faire seul.

**RÈGLE — les heures supplémentaires se lisent dans la colonne COM, elles ne sont PAS déduites
du total.** Le skill `recap-feuilles-heures` cherche le motif `(\d+)\s*h?\s*sup` dans COM
(« 8h sups », « 1j recup et 9h sups »). Écrire seulement « 43h » en heures totales ferait donc
apparaître **0h sup** dans l'onglet 2. Le script calcule l'écart au-delà de 35h et écrit
`8h sups` dans COM, en le préfixant au commentaire du technicien plutôt qu'en l'écrasant.

**Colonnes `Devis N` renommées en `N° chantier N` (24/08/2026).** Le n° d'index du chantier
(`26-0XX`) remplace le n° de devis comme identifiant : il est unique et stable, alors qu'un devis
change de version (`…V2`). Les colonnes `heures Devis N travaillées` ont suivi en
`heures chantier N travaillées`. Conséquences :
- Le n° d'index a été propagé de `suivi-chantiers` vers `appli-techniciens` (il y était totalement
  absent), pour que l'appli puisse l'émettre dans ses soumissions.
- Le skill a été adapté : sa détection accepte désormais `Devis N` **et** `N° chantier N`, donc les
  fichiers mensuels antérieurs se recalculent toujours à l'identique.
- **Le skill s'exécute côté Claude.ai, pas depuis Claude Code** (il tourne avec Python, absent de
  cette machine, cf. [[feedback_environment_no_python]]). Modifier le fichier local dans
  `AppData/.../skills/recap-feuilles-heures/` **ne change rien** : il faut demander la modification
  dans une session qui l'exécute. Une copie de sauvegarde est conservée dans
  `suivi-chantiers/scripts/skill-recap-feuilles-heures/`, ce cache `AppData` pouvant être écrasé.

**Conventions de saisie relevées dans les récaps existants** (à respecter, ne pas inventer) :
heures en texte (`35h`), GD/PD/Nacelle en nombres, `CP/RECUP` = marqueur texte (`CP`, `1 RTT`),
`ARRET` = `AM`/`AT`, `COM` = texte libre.

**Périodes mensuelles : du 20 au 20, en semaines entières.** La semaine qui contient le 20
bascule dans la période suivante — vérifié sur les fichiers de Patrice : août = S30→S33
(20/07→16/08), septembre = S34→S38 (17/08→20/09). C'est Claude qui crée le nouveau fichier vierge
avec ses semaines quand Patrice le demande.

**Cas particuliers de techniciens.**
- `TECHS_SANS_CHANTIER` dans `appli-techniciens/index.html` : profils sans chantier attribué
  (chargés d'affaires). **Ahmed HAMOUCH** y figure — sa semaine est forfaitairement à 35h, il ne
  compte pas d'heures sup, et **seuls ses GD, PD et primes nacelle intéressent Patrice**. Une
  feuille sans aucune heure est donc normale pour lui et ne déclenche pas l'avertissement de
  feuille vide. Pour tous les autres, cet avertissement demande confirmation avant l'envoi.
- **Une seule nacelle par jour au maximum** : la bascule est un simple oui/non. Elle montait à 2
  avant le 24/08/26, ce qui permettait de saisir une prime impossible.

**La saisie vit dans le navigateur du téléphone**, pas sur un serveur : ni Patrice ni Claude ne
peuvent l'effacer à distance. D'où le bouton « Effacer et recommencer cette semaine » dans la
feuille d'heures. Corollaire : après une mise à jour de l'appli, un technicien peut recevoir
l'ancienne version en cache — lui faire fermer complètement l'onglet et rouvrir son lien.

**Les feuilles Excel remplies à la main par les techniciens (ancien circuit) ne sont PAS
exploitables par une machine.** Constaté le 24/08/26 sur les 7 fichiers de la S34 : chantier écrit
sur la ligne « Qualité », ligne sans étiquette, et chez Moussa une disposition décalée d'une ligne
(sa case « PD » là où les autres ont « GD »). Un programme qui les lirait mettrait des heures dans
les mauvaises cases sans prévenir. De plus Patrice y ajoute sa propre connaissance en recopiant
(un « Orléans, 35h » devient 21h sur Chaineau + 14h sur Dambron). Ne pas écrire de lecteur
automatique pour ces fichiers : l'appli résout le problème à la source.

## Heures cumulées par chantier — RÈGLE (posée par Patrice le 26/08/2026)

**À chaque nouveau récap FH reçu, mettre à jour `heuresTelsam` dans `SEED_DATA`.** C'est le seul
endroit où le temps réellement passé s'accumule chantier par chantier, et il ne se remplit pas
tout seul : jusqu'au 26/08/2026 les trois premiers récaps avaient été chargés en une fois à la
création du fichier (19/08/2026) et plus rien n'avait bougé depuis. Patrice l'a demandé
explicitement : « dès qu'un nouveau fichier est ajouté, on met à jour la durée cumulée des
chantiers ».

**Où lire.** Onglet **`Chantiers`** de
`TELSAM TEAM Dropbox/Patrice PIVOT/feuille d'heures/recap FH/recap_FH-<mois>.xlsx`.
Le fichier de la période **en cours** vit à la racine de `feuille d'heures/`, les périodes closes
sont rangées dans le sous-dossier `recap FH/`. Trois colonnes : Chantier, N° Devis, Total heures.
Le rapprochement avec les fiches se fait sur le **numéro de devis**, pas sur le libellé (les noms
du récap ne sont pas ceux du suivi : « FIBRAGE FEYRIAT » pour Fleyriat, « LA VERNEY - ST
GUILLERMME » pour St-Guillerme). Le moyen le plus sûr de trancher un devis inconnu : chercher son
numéro dans `Telsam Fibre\RTE` — le dossier qui le contient porte souvent l'index du chantier
(`Ligne aérienne\Livière - Mas Nou 26-025\Devis_TELSAM_CC_26019_…pdf` ⇒ fiche 26-025).

**PIÈGE — les récaps antérieurs à juin 2026 sont à un AUTRE format.** Constaté le 26/08/2026 en
intégrant mai. Les fichiers de janvier à mai n'ont qu'un onglet (`Feuil1`), **pas d'onglet
`Chantiers` tout prêt**, et **une seule colonne « heures travaillées » par semaine**, partagée
entre `chantier 1` et `chantier 2`. Quand un technicien a fait deux chantiers dans la semaine, la
répartition est écrite **dans la cellule Devis, entre parenthèses** :
`TELSAMCC26036V3 (14h)` d'un côté, `TELSAMCC26012 (7h)` de l'autre.
`scripts/skill-recap-feuilles-heures/recalc.py` **ne sait pas les lire** : il cherche une colonne
`heures Devis N travaillées` juste après chaque colonne de devis, colonne qui n'apparaît qu'à
partir de juin. Sur un fichier ancien il ne trouve aucune paire et rend **zéro chantier, sans
lever d'erreur** — donc ne pas se fier à sa sortie. Ces mois-là se dépouillent à part, en lisant
les blocs de semaine repérés par leur en-tête `S17`, `S18`… en ligne 1.

**RÈGLE — des heures sans chantier ne s'attribuent pas au jugé.** Elles sont laissées de côté et
**signalées à Patrice**, jamais rattachées à la fiche qui paraît la plus proche. Les deux cas
rencontrés en mai ont été tranchés par Patrice le 26/08/2026 et sont **clos — ne pas les
rouvrir** :

- **Ahmed HAMOUCH : ses heures ne vont jamais sur un chantier.** Il est chargé d'affaires, pas
  technicien de chantier — c'est le seul nom de `TECHS_SANS_CHANTIER` (cf. « Cas particuliers de
  techniciens » plus haut). Vérifié sur les cinq récaps de mai à septembre : aucun numéro de devis
  sur sa ligne, donc rien ne lui a jamais été attribué à tort. Ses heures vivent au récap RH, pas
  dans `heuresTelsam`. Ne pas reposer la question à chaque mois.
- **Perche - La Tour de Carol et Postes de La Perche et La Tour de Carol : hors périmètre.** Ces
  deux devis (`TELSAMCC24172` et `TELSAMCC25009`, 273 h en mai) n'ont pas de fiche dans le suivi et
  n'en auront pas. Surtout, **ne pas les rattacher à Hospitalet - La Tour de Carol (26-027)** : le
  nom se ressemble mais c'est un autre chantier, avec son propre dossier Dropbox.

Pour tout **nouveau** cas d'heures sans chantier, la règle tient : signaler, ne pas décider seul.

**L'historique des heures commence à MAI 2026. Décision de Patrice le 26/08/2026 : ne pas
remonter plus loin.** Avant mai, les chantiers n'étaient pas renseignés sur le récap — les
fichiers de janvier à avril existent bien dans `feuille d'heures/recap FH/`, mais il n'y a rien à
en tirer par chantier. Ne pas proposer de les intégrer, ne pas s'étonner qu'un chantier démarré en
début d'année paraisse sous-chargé : ses heures d'avant mai n'existent nulle part.

**PIÈGE — « le mois » est une période de paie, pas un mois calendaire.** Vérifié le 26/08/2026 en
lisant les en-têtes de semaine de l'onglet `Recap` :

| fichier | semaines | période réelle |
| --- | --- | --- |
| juin | S21-S25 | 18/05 → 21/06 |
| juillet | S26-S29 | 22/06 → 19/07 |
| août | S30-S33 | 20/07 → 16/08 |
| septembre | S34-S38 | 17/08 → 20/09 |

On garde quand même l'étiquette du fichier (`"mois":"2026-08"`), c'est celle que Patrice emploie.
Mais **ne jamais en déduire une date de travaux** : c'est ce qui fait qu'Anneyron - St Vallier,
chantier terminé le 28/07/26, porte 112 h « en août ». Ce sont ses heures de fin juillet.

**RÈGLE — on REMPLACE la valeur d'une période, et on additionne les périodes entre elles.**
Formulée trop vite le 26/08/2026, cette règle était incompréhensible ; Patrice a demandé qu'elle
soit réécrite. L'exemple qui la rend claire, Dambron - Voves :

| moment | ce que dit le récap de septembre | ce qu'on écrit |
| --- | --- | --- |
| 26/08/26 | Dambron - Voves — 42 h | `{"mois":"2026-09","heures":42}` |
| une semaine plus tard, même fichier complété | Dambron - Voves — **105 h** | `{"mois":"2026-09","heures":105}` |

Le chiffre du récap **est déjà le cumul de la période** : le classeur fait l'addition lui-même,
semaine après semaine, dans le même fichier. Écrire `42 + 105 = 147` compterait donc les 42
premières heures deux fois — une fois seules, une fois incluses dans les 105.

Le cumul demandé par Patrice se fait **entre** les périodes, pas à l'intérieur de l'une d'elles :
Chaineau lot 2 = 42 h (juin, figée) + 63 h (septembre, encore mouvante) = 105 h, et ce total
grandira à chaque nouvelle période.

**RÈGLE — un lot = ses propres heures.** Le récap distingue les lots (« LA CHAINEAU-CORDY-LAMOTTE
LOT 1 » et « LOT 2 » sont deux lignes, avec deux numéros de devis). Ne jamais recopier le même
montant sur les deux fiches. Erreur commise le 25/08/26 en scindant Chaineau : 42 h de juin et
28 h d'août recopiées sur les deux lots, soit 70 h fantômes, corrigées le 26/08/26.

**Les lignes « Agence » / « contribution agence » ne se rattachent à aucun chantier** : les
laisser de côté (14 h en juin, 63 h en juillet). Conséquence : le total de `heuresTelsam` ne sera
jamais égal au « TOTAL GÉNÉRAL » du récap. Comparer mois par mois, agence déduite.

**Contrôle après intégration — c'est lui qui a trouvé les deux erreurs.** Additionner
`heuresTelsam` par mois sur toutes les fiches et comparer à la somme de l'onglet `Chantiers` du
récap correspondant, agence déduite. Écart attendu : **zéro**. État vérifié au 26/08/2026 :
mai 1229 (7 chantiers sur 9, voir ci-dessus), juin 1905, juillet 1417, août 1335,
septembre 374 (période en cours, chiffre partiel).
Vérifier aussi que le `total` de chaque fiche est bien la somme de ses `parMois`.

**Ne pas oublier `SEED_VERSION`.** Sans incrément, un collègue qui a déjà ouvert le fichier garde
les anciennes heures et la mise à jour est invisible. Le hook pre-commit le vérifie.

## Boîtes WTC2 et nacelle — onglet du suivi (mis en place le 26/08/2026)

Onglet **« Boîtes & nacelle »** de `suivi_chantiers_205.html`, demandé par Patrice pour préparer
ses achats de boîtiers et ses locations de nacelle. **Il n'existe que dans le suivi ; l'appli des
techniciens n'est pas concernée et ne doit pas l'être.**

**RÈGLE — à chaque nouveau chantier, renseigner les boîtes.** C'est la raison d'être de l'onglet.
Un chantier actif sans boîtes renseignées apparaît dans l'encart « À compléter » en haut de la vue,
et y reste jusqu'à ce qu'on le remplisse. **Un chantier sans aucune boîte à poser se renseigne
avec 0** — c'est ce qui le fait sortir de la liste. Laisser le champ vide n'est pas une réponse :
ça veut dire « pas encore regardé ».

**Où se trouve le chiffre.** Dans le **devis** du chantier (dossier `DEVIS` de son dossier
Dropbox), article « Fourniture et pose boîtier type WTC2 avec son support et ses brides ». La
quantité y est explicite et la liste des pylônes avec. **Ne jamais l'estimer** : la relever.
Stocker la référence du devis dans le champ prévu, pour pouvoir revérifier sans refaire la
recherche.

**Structure du champ** (`boites` sur chaque fiche) :
```
boites: {
  lots: [ { type, auDevis, posees, pylones, periode, devis, libelle } ],
  nacelle: "",   // ce que dit le devis, vide si aucune nacelle prévue
  remarque: ""
}
```
`lots` est un tableau : un chantier peut en avoir plusieurs. Fleyriat (26-055) a un lot THYM et un
lot OPPC ; Bérat-Portet (26-034) a le lot Bouygues et le lot Lebag ; Lannemezan (26-066) a INEO et
OMEXOM. Le champ `libelle` sert à les distinguer dans la vue.

**Les deux types, à ne pas confondre — c'est ce qui commande le budget nacelle :**
- **THYM** : boîtier WTC2, posé **en pied** de pylône ou de portique. **Pas de nacelle.**
- **OPPC** : boîtier posé **sur le conducteur** (CBJ, Donuts). **Nacelle obligatoire**, et le prix
  n'a rien à voir : 4 750 € la boîte OPPC nacelle comprise à Fleyriat, contre 1 175 € le WTC2.

Le nombre de boîtes ne dit donc rien du besoin en nacelle. Ce sont deux lectures différentes, d'où
la section « Nacelle » séparée dans la vue.

**Saisie.** Formulaire de modification d'une fiche, trois champs en bas : un textarea (une ligne
par lot, champs séparés par `|` dans l'ordre *type | au devis | posées | pylônes | période | devis
| libellé*), puis nacelle et remarque. Le parseur est volontairement tolérant : champ manquant =
vide, quantité non numérique = 0, type non reconnu = THYM. Mieux vaut une ligne imparfaite
enregistrée qu'une saisie refusée, la vue montre de toute façon ce qui a été compris.

**Les écarts entre documents ne vont PAS dans cet onglet** mais dans `aVerifier`, donc dans
l'onglet « À vérifier », avec les autres points en attente d'arbitrage. Cinq y ont été versés le
26/08/2026 (Bradascou 10 boîtes « pose » sans « fourniture » ; qui commande les OPPC de Fleyriat ;
pylône 91 manquant chez Lannemezan INEO ; pylône 55 compté dans les deux lots Lannemezan ;
portiques de Chaineau lot 1 selon devis ou CCTP).

**Trois types, pas deux.** Un troisième est apparu en lisant les devis des liaisons souterraines :
**CHAMBRE**, pour les boîtiers de jonction posés en chambre. Ce ne sont pas des WTC2 et le prix n'a
rien à voir (275 € contre 1 175 €) : les compter à part évite de fausser la commande de WTC2.
Attention aussi au WTC1,5 de Rion des Landes, qui n'est pas un WTC2 — et dont le devis exclut la
fourniture : 1 boîtier à poser, 0 à acheter.

**PIÈGE — « le devis mentionne un boîtier » ne veut pas dire « il y a un boîtier à acheter ».**
Quatorze devis sur les 28 relus le 26/08/2026 parlaient de boîtiers ; la moitié ne portait que sur
du raccordement de boîtiers DÉJÀ EXISTANTS (Cantegrit, Argia-Hernani, Manoire, Arteria Pessac,
Chafauds-Chaingy). Seule la ligne « **Fourniture** et pose » compte. De même, « hors fourniture
matériel » (Donuts de Blavozy et d'Asasp) veut dire que le client fournit : zéro à acheter, mais la
nacelle reste à prévoir. Et une ligne « Option : fourniture et pose… » ne se compte pas tant qu'elle
n'est pas commandée (St-Guillerme). Enfin, un devis peut ne contenir AUCUNE ligne de fourniture :
Jonquières-St Césaire ne fait que raccorder ou mettre en place du matériel existant ou fourni, seule
la nacelle y est facturée — zéro boîte à acheter.

**Contrôle systématique après chaque saisie.** Pour chaque quantité enregistrée, retrouver dans le
devis la ligne qui contient à la fois « fournitur » et « boît », et comparer le nombre qui suit la
cellule d'unité. Fait le 26/08/2026 sur les 26 lots chiffrés : 24 concordent, Hospitalet vérifié à
part, et Bradascou reste le seul écart connu (article « Pose » sans « Fourniture », déjà dans
À vérifier). Piège technique de ce contrôle : Word termine chaque cellule de tableau par CR + BEL
(caractère 7), donc découper le texte sur les seuls retours à la ligne laisse un caractère invisible
en tête de chaque ligne et plus rien ne correspond. Il faut découper en incluant le caractère 7.

**Mettre à jour les boîtes posées — accord passé avec Patrice le 26/08/2026.** Quand il le demande,
lire les dossiers **« Suivi »** / « suivi de chantier » des chantiers concernés dans le Dropbox, et :
- **écrire seul** ce qui est dit explicitement. Exemple qui se suffit à lui-même, Givors Ban du
  10/08/26 : « installation wtc2 portique et pylône 1 » ⇒ 2 posées, aucune interprétation.
- **soumettre à Patrice** tout ce qui demande de traduire. « BE au pylônes 63 et 55 fait »
  (St Christol) suppose que BE = boîte d'épissure ; « raccordement : pyl65 pyl66… » (Arcomie) ne dit
  pas que la boîte a été posée. Montrer ce qu'on a compris, laisser Patrice trancher.
- ce que les fichiers ne disent pas, **Patrice le complète lui-même** : ne pas le déduire.

**Attention à la couverture, elle est faible.** Au 26/08/2026 : 12 chantiers ont un dossier
« Suivi », et sur ceux qui ont des boîtes à poser, **3 seulement** contiennent un texte exploitable.
Bissy - Grand Île (8 boîtes) et Fleyriat (11 + 2) ont un dossier **vide**. Les fichiers sont aussi
souvent en retard : celui de Givors datait du 10/08 et ne couvrait pas les poses du 17 au 23/08.
Ne jamais conclure « rien n'a été posé » de l'absence d'écrit.

**À revoir quand le fichier normé existera.** Patrice prévoit de mettre en place un fichier de suivi
de chantier normalisé. Le jour où il est en place, cette règle est à réécrire : la lecture pourra
devenir automatique et la part « soumettre à Patrice » disparaîtra en grande partie.

**État au 26/08/2026** : 44 fiches renseignées sur 45 chantiers actifs, **167 boîtiers WTC2**,
2 boîtiers OPPC (Fleyriat) et 4 boîtiers en chambre restant à poser, 29 déjà posés. Reste une seule fiche à
compléter, 26-021 Gampaloup-Valence, parce qu'aucun devis n'est encore arrivé dessus.

## PDP : portée poste / ligne — RÈGLE CRITIQUE (posée par Patrice le 25/08/2026)

**Un PDP autorise un périmètre précis. UN PDP LIGNE N'OUVRE PAS LA PORTE D'UN POSTE.** Dès qu'il
y a une recette, il faut entrer dans le poste pour brancher les appareils : sans PDP couvrant ce
poste, l'équipe est refoulée à l'entrée et la journée est perdue — alors que la fiche affichait
« PDP OK » en vert. **Un vert mensonger est pire qu'un rouge**, parce que personne ne va vérifier
derrière. Corollaires :
- **Plusieurs PDP sur un chantier est NORMAL, pas une anomalie.** Une ligne qui traverse
  plusieurs postes demande souvent un PDP ligne PLUS un PDP par poste. Cas réels relevés dans
  Dropbox : Bagatelle-Issel (`LA ISSEL REVEL` + `PS AVIGNONET` + `PS ISSEL`),
  Bédarieux-Espondeilhan (`LA BEDARIEUX` + `PS BEDARIEUX` + `PS ESPONDHEILAN`), Fibrage FO DI Lyon,
  Sèvres-St Vallier, Givors. Marqueurs dans les noms de fichiers : `LA`/`Ligne` d'un côté,
  `PS`/`PO`/`Poste` de l'autre (sur 473 PDP, 91 nomment un poste et 59 une ligne).
- **Un PDP unique peut couvrir poste ET ligne** — dans ce cas il n'y a rien à faire, on le note
  `portee: 'ligne+poste'`.
- **Le devis dit ce dont on a besoin** : c'est lui qui indique s'il y a des recettes et depuis
  quels points. Toujours le lire pour renseigner le périmètre (indication de Patrice).

Deux champs par fiche, dans **les deux dépôts** :
- `perimetre` : `{ ligne: bool, postes: [noms des postes où il faut ENTRER], source: "..." }`
  = ce que le chantier EXIGE. `source` doit citer la preuve (article de devis, etc.).
- `pdps` : `[{ portee: 'ligne'|'poste'|'ligne+poste', poste|postes, tousPostes, ref, indice,
  statut, validFrom, validUntil, note }]` = ce qui est COUVERT.

`computePdpAlerts(c)` croise les deux et lève une alerte `bad` par poste non couvert. Elle est
**calculée à l'affichage, jamais stockée** — elle se corrige donc seule dès que les données
changent. `pdpBadge()` fait passer le badge à « PDP incomplet » (rouge) dans ce cas. Sur une fiche
sans `perimetre`, tout le comportement est **inchangé** : les 67 autres fiches ne bougent pas.

**`perimetre.pdpSurPlace` — marqueur au cas par cas, JAMAIS déduit.**
Sur certains chantiers le PDP n'est pas obtenu à l'avance : les techniciens l'établissent sur
place le matin, directement avec RTE, poste par poste. C'est le cas de **`26-061` AUDIT MULTI
POSTES** (Bruges, Cubnezais, Garies, Marquis, Mérignac, Pessac, Saucats), confirmé par Patrice le
25/08/26. Le marqueur `perimetre.pdpSurPlace: true` coupe alors le contrôle, sinon ce chantier
porterait une fausse alerte permanente — et une fausse alerte permanente finit par faire ignorer
les vraies.

**Mais ce n'est PAS une règle générale, et surtout pas une propriété des audits.** Patrice a
explicitement corrigé cette généralisation le 25/08/26 : « ça peut arriver de temps en temps mais
ce n'est pas une règle ». Rien ne permet de déduire ce mode de fonctionnement du type de chantier
ni du nombre de postes. Donc :
- Ne JAMAIS poser `pdpSurPlace` de sa propre initiative. Il se pose uniquement quand Patrice
  l'indique, chantier par chantier.
- Ne pas non plus renseigner `perimetre.postes` sur un chantier dont on ignore le mode de
  fonctionnement : demander plutôt que présumer, dans un sens comme dans l'autre.

**RÈGLE — ce contrôle vit uniquement dans `suivi-chantiers`, PAS dans l'appli technicien**
(décision de Patrice, 25/08/26). Il avait été mis dans les deux, puis retiré de l'appli le jour
même : aucun technicien n'est envoyé sans que les documents soient en règle, c'est Patrice qui
décide des départs. L'alerte a donc sa place dans son outil de pilotage, pas sur le terrain où
elle inquiéterait sans action possible. Ne pas la réintroduire dans `appli-techniciens` sans
demande explicite (une note dans le code, à l'endroit où elle se trouvait, le rappelle).

**Piège de comparaison des noms de postes** : `normNom()` réduit accents, séparateurs et préfixes
(`petit-bois` = `Petit Bois`, `PS ISSEL` = `Issel`, `Poste de Portet` = `PORTET`). Sans ça le
contrôle signalait des postes non couverts qui l'étaient — et un contrôle qui crie au loup est
abandonné en trois jours. Ne pas simplifier cette fonction.

**NDS et PPSPS : même besoin, deux régimes** (précision de Patrice, 25/08/26). La NDS relève du
**décret 92**, le PPSPS du **décret 94** quand le chantier est sous coordination SPS. Le PPSPS se
traite donc exactement comme une NDS : il conditionne le démarrage et doit figurer dans App Tech.
Cas vécu : la VIC DATA4 du 09/12/25 exige « aucune intervention ne pourra débuter sans
présentation d'un PPSPS à jour, validé » — et aucun PPSPS n'existait. Les deux types sont
désormais surveillés par la veille au même titre que PDP/PGO/IST.

## Veille documentaire automatique (mise en place le 25/08/2026)

**Le rapport de veille est injecté automatiquement au démarrage de chaque session** — plus besoin
d'y penser, ni pour Patrice ni pour moi. Il arrive dans le contexte avant le premier message, avec
sa date de génération et les trois derniers passages de la tâche.

Mise en place le 25/08/26, après une question de Patrice : « qu'est-ce que tu appelles début de
session exactement ? ». La réponse honnête était que la consigne, écrite ici seulement, dépendait
de ma mémoire — **exactement la faiblesse de la règle `SEED_VERSION`, oubliée deux fois malgré sa
documentation.** D'où le même remède : un mécanisme exécuté par l'outil, pas par moi.
- Hook `SessionStart` déclaré dans `TELSAM-apps/.claude/settings.json`, qui lance
  `suivi-chantiers/scripts/hook-veille-session.ps1`.
- Ce script émet `hookSpecificOutput.additionalContext` ; il **prévient si le rapport a plus de
  48 h** (la tâche planifiée n'a pas tourné) et ne fait jamais échouer une session en cas de
  problème — il sort silencieusement.
- **`TELSAM-apps/.claude/settings.json` n'est dans aucun des deux dépôts** (il est dans le dossier
  parent, qui n'est pas un dépôt Git) : il n'est donc pas sauvegardé par Git. S'il disparaît, le
  recréer avec un hook `SessionStart` de type `command` appelant le script ci-dessus.

Le rapport reste lisible à la main dans `veille\RAPPORT.md` (version exploitable : `dernier.json`).
Il liste les PDP/PGO/IST/NDS/PPSPS nouveaux ou modifiés depuis le passage précédent, avec pour
chacun le chantier concerné et s'il est déjà présent dans son dossier App Tech. Ne pas attendre que
Patrice signale un document : le vrai point faible du circuit n'était pas la mise à jour elle-même
mais le fait que tout reposait sur lui pour y penser.

Le script est `suivi-chantiers/scripts/veille-documents.ps1`, lancé chaque jour à 07h30 par
la tâche Windows **« TELSAM - Veille documents RTE »** (compte `patrice.pivot`, session ouverte
requise, aucun mot de passe stocké, rattrapage automatique si le PC était éteint). Il est en
**lecture seule** : il ne touche ni à Dropbox, ni aux dépôts, seulement à son dossier `veille\`.

**Pourquoi en local et pas dans le cloud.** La tâche planifiée cloud avait échoué en août sur
deux blocages (impossible d'attacher le connecteur Dropbox à une routine, compte GitHub non
connecté pour les routines — cf. [[feedback_no_cloud_automation_pdp_pgo]]). Le raisonnement qui
manquait : **tout le Dropbox est déjà synchronisé sur le disque**, entièrement, sans fichiers
« en ligne seulement ». Un balayage complet de l'arborescence RTE (142 000 fichiers) prend 18 s.
On n'a donc besoin d'aucun connecteur ni d'aucune permission cloud. Ne pas re-tenter la voie
cloud : elle n'apporterait rien de plus et rencontrerait les mêmes blocages.

Points de conception à connaître avant de le modifier :
- Il tient un **index de référence** (`veille\etat.json`) de TOUS les documents de sécurité,
  pas seulement des récents — sinon chaque passage rapporterait tout l'existant comme nouveau.
  Ne pas supprimer ce fichier sans raison ; s'il disparaît, le passage suivant repart en
  « première passe » et ne liste que les 7 derniers jours.
- **L'identité d'un document est `nom + taille + date`, JAMAIS son chemin** (corrigé le
  25/08/26). Patrice renomme progressivement ses dossiers Dropbox avec les n° d'index : avec une
  clé fondée sur le chemin, chaque dossier renommé faisait passer tout son contenu pour du neuf
  — constaté sur « Chaineau-Cordy-Lamotte 26-036-1-2 », 10 faux nouveaux d'un coup. Vérifié sur
  une arborescence d'essai : renommer un dossier ne rapporte rien, modifier réellement un
  document rapporte toujours. Ne pas revenir à une clé par chemin.
- `etat.json` porte un champ `format` (actuellement 2). **Si la forme des clés change, incrémenter
  `$FORMAT_ETAT`** : le passage suivant reconstruit alors l'index en silence au lieu de lister
  les ~1 700 documents comme nouveaux.
- Extensions volontairement restreintes (`.pdf .doc .docx .xls .xlsx .xlsm .zip`) : les `.dwg`,
  `.bak`, `.dwl` sont des fichiers de travail AutoCAD qui bougent en permanence et noyaient le
  rapport.
- Le motif IST est borné (`(^|[^a-z])ist([^a-z]|$)`) pour ne pas matcher « Liste », « Existant ».
- Dédoublonnage par nom+taille : un même document est souvent copié dans plusieurs dossiers
  (PGO racine, PGO du sous-chantier, App Tech) — il est rapporté une fois, avec tous ses chemins.
- Sous `Audit FO`, les chantiers sont rangés par direction interrégionale (`DI Marseille`,
  `DI Toulouse`) : le dossier chantier est donc un niveau plus bas que dans les autres catégories.
- Le dossier `préparation documents sécurités` (premier niveau, hors arborescence chantier) est
  balayé lui aussi : il contient des IST revenues signées par RTE avant classement. Un document
  qui n'existe QUE là apparaît dans une section « hors dossier de chantier — à rattacher ». S'il
  existe aussi dans un dossier de chantier, c'est cet exemplaire-là qui sert de référence, pour
  garder le nom du chantier et le contrôle App Tech.
- Un avertissement « dossiers attendus et introuvables » dans le rapport n'est PAS cosmétique :
  il veut dire qu'une catégorie entière n'a pas été balayée (dossier renommé ou déplacé). À
  traiter, pas à ignorer.
- **Une IST signalée par la veille n'est pas forcément une IST TELSAM.** La veille rapporte tout
  fichier du chantier dont le nom contient `ist` + `sign`/`valid` : elle ne sait pas qui l'a
  rédigée, et elle affiche même « — signe » dans son titre de section. Ne jamais enchaîner du
  rapport vers App Tech sans avoir ouvert le document et vérifié qu'il nomme TELSAM (cf. la règle
  du 27/08/2026 en section « Contenu standard d'un dossier App Tech »). Cas réel du 27/08/2026 :
  deux IST INEO sur Cantegrit, dont une nommée « signée RTE », envoyées par le client à titre
  d'exemple.

**Second contrôle : la cohérence des dossiers App Tech.** Indépendant des nouveautés, il compare
pour chaque chantier ayant un App Tech le document qui s'y trouve avec le plus récent disponible
ailleurs dans le dossier du chantier, et signale les retards. Motivé par un cas réel : le PGO de
Portet était resté en V37 dans App Tech alors que la V40 existait depuis le 10/08 — deux semaines
de retard découvertes par hasard le 25/08. La veille des nouveautés ne peut pas attraper ça,
puisqu'aucun fichier neuf n'arrive. Dès son premier passage, ce contrôle a trouvé le PDP de
Verney - St Guillerme resté en **Ind 7 (22/05)** dans App Tech alors que l'**Ind 8 (25/06)** était
disponible. Détails de conception :
- **Les IST ne sont signalées que si leur nom contient `sign` ou `valid`.** Sans ce filtre, le
  contrôle réclamerait la mise en ligne de brouillons — exactement l'inverse de la règle App Tech.
- **Une archive est jugée sur son contenu, pas sur son nom** (`NomsDansArchive`). Les PGO/PDP
  arrivent souvent en `.zip` transféré tel quel : comparer les noms faisait passer
  « …PGO ind.5.zip » pour un retard alors qu'App Tech contenait déjà le PDF extrait de cette
  archive. On ouvre donc le zip et on regarde si l'un de ses fichiers y figure déjà.
- Ces deux règles ont fait tomber les faux positifs de 4 à 2 au premier essai. Ne pas les retirer
  en "simplifiant".

**RÈGLE DE LECTURE — un retard App Tech ne vaut que si des techniciens sont planifiés**
(précision de Patrice, 25/08/26). Les premiers dossiers App Tech ont été montés il y a plusieurs
semaines pendant les essais ; ces chantiers-là sont aujourd'hui **terminés ou en suspens**, et
qu'ils ne soient pas à jour n'a aucune conséquence — personne n'y va. Ce qui compte, c'est
**les chantiers des semaines à venir où des techniciens sont planifiés**. Donc :
- Ne pas remonter spontanément un retard App Tech sur un chantier sans présence technicien
  prévue. Le signaler une fois suffit ; y revenir chaque matin transforme le contrôle en bruit,
  et un contrôle bruyant finit par être ignoré — c'est exactement ce qu'on cherche à éviter.
- En revanche, un retard sur un chantier où des techniciens sont placés (ou le seront dans les
  semaines qui viennent) est à traiter tout de suite : c'est le cas Portet V37/V40 du 25/08.
- Connus et acceptés au 25/08/26, sans action demandée : PDP de Verney - St Guillerme (Ind 7 au
  lieu de Ind 8), PDP de Lamativie - La Mole, NDS de Fibrage Feyriat. Patrice doit d'abord
  vérifier auprès des chargés de travaux et chargés d'affaires si ces chantiers sont terminés.

**Le rapport applique cette règle lui-même** (25/08/26) : la section App Tech est coupée en
« **A TRAITER** » (chantier avec des techniciens planifiés cette semaine ou après) et
« **Pour information** » (aucun technicien planifié — une ligne par cas, sans détail). Un chantier
n'est « prioritaire » que s'il a une présence à partir du lundi de la semaine en cours **et** n'est
pas marqué terminé. Le `RIEN A SIGNALER` ne regarde que les prioritaires.

**Rapprochement dossier Dropbox → fiche : deux stratégies, jamais d'à-peu-près.**
1. Par **n° d'index dans le nom du dossier** (`26-055 - …`) — voie normale à mesure que Patrice
   renomme ses dossiers.
2. Sinon par le champ **`dossierDropbox`** de la fiche : le chemin exact relevé via
   `get_shared_link_metadata`, pas déduit du libellé. Renseigné sur les 9 fiches qui ont un
   dossier App Tech.
**Ne jamais rapprocher approximativement sur le nom** : le chantier Fleyriat (26-055) a son
dossier orthographié « Fibrage Feyriat », sans le L — un rapprochement au libellé se tromperait
précisément là. Si un dossier n'est rapproché à aucune fiche, le rapport l'indique
(« fiche non rapprochée ») : c'est le signal qu'il faut mettre à jour `dossierDropbox`.

**Ce que la veille ne remplace pas — et pourquoi c'est délibéré.** Elle détecte et signale, elle
ne met JAMAIS à jour App Tech toute seule. Question posée par Patrice le 25/08/26 ; la réponse
tient à trois cas vécus le jour même :
- Le PGO Fleyriat est arrivé en `.zip` contenant trois fichiers (deux PDF, un MS Project) — il
  faut choisir lequel fait foi pour un technicien sur téléphone.
- L'IST de Portet paraissait complète ; seule la lecture de la page 23 sur 24 montrait que le
  tableau « Validation RTE » était vierge. Une copie automatique aurait laissé croire à Pascal
  Bonaventure qu'il pouvait attaquer le poste 63 kV. **C'est le seul cas où une erreur
  d'automatisation a une conséquence de sécurité, pas seulement de fichier.**
- Les conventions d'indice n'ont rien de commun d'un projet à l'autre (`Ind.3`→`Ind.5`,
  `V37`→`V40`, `indice A`, `V08`) : un script qui se trompe d'ancienne version supprime le
  document en vigueur et laisse le périmé.
Le déséquilibre est structurel : une détection ratée coûte une journée, une mise à jour
automatique fausse met un document périmé ou non validé entre les mains d'un technicien. La
lecture, la mise à jour des fiches et la copie dans App Tech restent donc faites en session. Deux
choses resteront de toute façon locales : les PGO/PDP `.xlsm` (Excel requis) et la génération des
briefs techniciens (Word requis).

## Accès à l'appli — portail par technicien (mis en place le 27/08/2026)

Remplace l'ancien **code d'accès partagé** (`GATE_CODE_HASH`, supprimé). Chaque technicien a
**son** mot de passe (deux mots : un mot + une couleur, ex. « nuage-emeraude ») : il ouvre l'appli,
tape son mot de passe, et l'appli le reconnaît et le connecte directement — plus de lien `?tech=`
transmissible, plus de code commun. Un mot de passe **« atelier »** (celui de Patrice) ouvre à la
place la liste des 13 noms, pour dépanner un technicien qui a oublié le sien. Demandé une seule
fois par téléphone (mémorisé en IndexedDB, clé `gate_ok` = nom du technicien ou `__ADMIN__`) ; le
bouton « Changer » **déconnecte** un technicien et redemande le mot de passe, alors qu'en accès
atelier il revient seulement à la liste des noms.

**Ce portail protège l'ENTRÉE de l'appli, PAS les documents Dropbox.** Les liens Dropbox sont
écrits dans `index.html`, page téléchargée AVANT l'écran d'accès : quelqu'un de technique les lit
dans le code source sans mot de passe. La vraie protection des documents est le mot de passe posé
sur chaque lien **côté Dropbox** (cf. section Accès Dropbox). Ne jamais présenter le portail comme
une protection des documents — dit à Patrice le 27/08/2026, après avoir ouvert un PDP RTE (avec
coordonnées personnelles d'agents RTE) sans aucun code.

**Stockage** : jamais le mot de passe en clair, seulement son empreinte **PBKDF2-SHA-256,
150000 tours, sel `telsam-app-tech-2026`** (`ACCESS_HASHES` indexé par nom exact de `TECHS`,
`ADMIN_HASH` à part). Normalisation `normPw` : minuscules + **suppression de tout caractère non
`a-z0-9`** (tirets, espaces, ponctuation) — donc « nuage-emeraude », « nuage emeraude »,
« Nuage Emeraude » et « nuageemeraude » donnent tous la même clé `nuageemeraude`. Fait ainsi
depuis le 27/08/2026 pour ne PAS obliger un technicien à taper le tiret sur son téléphone (première
version qui remplaçait les séparateurs par `-` obligeait un séparateur, pénible au clavier mobile).

**Changer ou ajouter un mot de passe** (procédure, à refaire à l'identique) :
1. Choisir un mot de passe **ASCII uniquement** (a-z 0-9) — sinon `normPw` transforme les accents
   en `-` et l'empreinte ne correspondra pas à ce que le technicien tape.
2. Calculer l'empreinte avec EXACTEMENT le même algorithme (même sel, mêmes tours, même `normPw`).
   Le plus sûr : dans un **contexte sécurisé** — `crypto.subtle` est absent d'une page `file://`
   ou `data:`, présent en `https://` et sur `http://localhost` — exécuter `deriveAccess(pw)`. C'est
   ainsi que les 14 empreintes ont été calculées le 27/08 (dans un onglet https, puis revérifiées
   en servant l'appli sur `http://localhost`).
3. Remplacer la valeur dans `ACCESS_HASHES` (ou `ADMIN_HASH`), **bumper `APP_VERSION`**, committer.
4. Redonner le nouveau mot de passe au technicien : pas de « mot de passe oublié », c'est Patrice
   qui le redit, et un changement suppose une republication de l'appli (pas faisable par lui seul).

Les 14 mots de passe en clair vivent dans `Mots_de_passe_App_Techniciens_TELSAM.docx` sur le
Bureau de Patrice (hors dépôt) — **jamais dans Git**. Chaque technicien ne reçoit que sa ligne.

## Accès Dropbox
Dropbox est connecté et utilisable directement (create_folder, create_shared_link, create_file_request,
search, list_folder, fetch). Toujours confirmer le plan exact avec Patrice avant toute création/modification
dans Dropbox (dossier, lien de partage, file request) — ne jamais deviner un chemin de destination.

**RÈGLE — liens de partage "App Tech" toujours privés par défaut, à vérifier systématiquement.**
`create_shared_link` de ce connecteur Dropbox est bridé en dur sur `audience: "no_one"` (lien
privé, accessible seulement aux personnes déjà autorisées) — il ne peut jamais créer un lien
public directement, quel que soit le réglage du compte Dropbox de Patrice (confirmé le 20/08/26 :
un lien "App Tech" plus ancien, Fleyriat, est bien `audience: "public"`, donc le compte autorise
les liens publics — c'est uniquement l'outil qui les crée restreints). Conséquence pratique :
- Après CHAQUE `create_shared_link` sur un nouveau dossier "App Tech", vérifier immédiatement
  avec `get_shared_link_metadata` si `audience` == `"public"`.
- Si ce n'est pas le cas (`"no_one"` ou autre), le signaler explicitement à Patrice dans la même
  réponse, avec le lien exact et l'étape à faire côté Dropbox : clic droit sur le dossier →
  Partager → dans "Paramètres de partage", régler le **lien de consultation** (pas la partie
  grisée "qui peut être ajouté", inutile pour des techniciens sans compte Dropbox) sur
  "Toute personne disposant du lien". Ne pas committer/pousser `documentsAppTech` sans avoir
  fait cette vérification et ce signalement — ne jamais supposer que le lien est public.
- Décision de Patrice (20/08/26) : on ne restructure PAS les dossiers Dropbox pour contourner ça
  (ex. dossier parent unique déjà public où tout hériterait automatiquement) — il préfère garder
  la structure actuelle et faire le clic manuel à chaque nouveau chantier plutôt que de sortir les
  docs techniciens de leur dossier de chantier respectif.

**RÈGLE — les liens des documents sont protégés par mot de passe côté Dropbox (serrure 2, mise en
place par Patrice le 27/08/2026).** C'est la **seule** protection réelle des documents : le portail
de l'appli ne garde que l'entrée (cf. « Accès à l'appli — portail par technicien »). Les 14 liens
de `appli-techniciens/index.html` — 12 dossiers « App Tech », le PDP et le brief de Bollène
(26-062), et le dossier Habilitations — ont été passés en « personnes disposant du lien + mot de
passe », vérifiés un à un avec `get_shared_link_metadata` (`password_protected: true`). Points à
retenir :
- **L'URL ne change pas** quand on ajoute un mot de passe (même `rlkey`) : l'appli continue de
  fonctionner sans retouche. (Un changement d'URL n'arrive que si on supprime puis recrée le lien.)
- **Vérifier chaque lien après coup, ne pas croire sur parole** : DATA4-Marcoussis (26-065) était
  resté public au premier passage, repéré uniquement par le contrôle `password_protected`.
- Le mot de passe est demandé **par lien, pas par document** : un dossier App Tech s'ouvre une
  fois, puis tous ses fichiers se lisent sans le retaper ; un autre chantier = un autre lien = à
  retaper (le même mot de passe partout). Dropbox le mémorise pour la session du navigateur.

**Pourquoi il n'y a pas de vraie serrure à l'entrée de l'appli** (analyse du 27/08/2026). GitHub
Pages sert le site à quiconque a l'URL — on ne peut pas y restreindre l'accès à une liste de
personnes. Passer le dépôt en privé ne change rien au site servi, et **casserait Pages sur un
compte GitHub gratuit** (Pages sur dépôt privé exige un plan payant). `Pivot-telsam` est un compte
personnel, plan non vérifiable de l'extérieur : **ne jamais proposer de basculer le dépôt en privé
sans avoir fait confirmer le plan par Patrice.** Une vraie porte d'entrée (Cloudflare Access :
liste d'e-mails, code reçu par mail, gratuit ≤ 50 personnes) supposerait de déménager l'appli sur
un domaine routé par Cloudflare — chantier d'une demi-journée étalé sur 1-2 jours à cause de
l'attente DNS, **repoussé après le déploiement** (décision de Patrice, 27/08/2026).

## Workflow hebdomadaire
1. Le lien personnalisé technicien (`?tech=slugifiedname`) est permanent — jamais régénéré à chaque semaine.
2. Ce qui change chaque semaine : les affectations technicien/chantier (TECH_RANGES côté appli-techniciens,
   structure équivalente côté suivi-chantiers) — à mettre à jour dans les deux dépôts en cohérence.
3. Les liens Documents/Dépôt photos/base de vie sont à créer/vérifier pour les NOUVEAUX chantiers
   apparaissant dans le planning, ET pour tout chantier déjà existant qui devient actif pour la
   PREMIÈRE FOIS cette semaine-là (première fois avec des techniciens réellement affectés,
   même si la fiche existait déjà sans jamais avoir eu de TECH_RANGES) — vécu le 20/08/26 avec
   DATA4-Marcoussis et Audit Multi Postes, oubliés lors d'une mise à jour de planning car leur
   fiche n'était pas "nouvelle". Ne pas se fier uniquement à la date de création de la fiche.
   Pas besoin de recréer/vérifier pour les chantiers déjà actifs les semaines précédentes.

## État au 20/08/2026
- 6 chantiers ont un bouton "📷 Dépôt photos" (Portet, Fleyriat, Fibrage THYM Bissy-Grand Ile,
  St-Guillerme, Lien Arteria Lamativie, Hospitalet-La Tour de Carol).
- Champ "Immatriculation du véhicule" ajouté à la feuille d'heures (app + export Excel cellule A52 + PDF).
- 28 chantiers ont un champ `baseVie` (adresse extraite des NDS) pour le bouton Itinéraire.
  ~20 chantiers encore non trouvés/à approfondir si besoin.
- Fusions terminées (fiches doublons supprimées dans les deux dépôts, données fusionnées) :
  "Poste de Tivernon" (26-070) → "Dambron - Voves" (26-031) ; "Vallee Du Louron (Aure-Loudenvielle)"
  (26-011) → "Aure - Loudenvielle" (26-024, baseVie de 26-011 volontairement PAS repris car
  mal attribué — provenait du NDS d'un segment voisin).
- Renommages terminés : 26-002 "Icp / Ls Cross-Sausset-Goodman" → "Ls Cross-Sausset-Goodman" ;
  26-003 "PS Cantegrit" → "Poste Cantegrit" ; 26-007 "CONSIGNATION 225kV CURBANS - SISTERON" →
  "Curbans - Sisteron (225kV)" (d'après devis signé TELSAM/CC/RTE/25027).
- Semaine 35 (24-28/08) mise à jour dans les deux dépôts : Poste de Portet, DATA4-Marcoussis,
  Audit Multi Postes — voir règle Planning RTE 2026.xlsx ci-dessus.

## Écran d'ouverture animé (appli-techniciens, mis en place le 26/08/2026)

Au lancement de l'appli, le logo TELSAM s'anime : des points s'allument comme des sites sur une
carte et se relient, puis le **« s » du logo se trace tout seul** entre sa flèche cyan (en bas à
gauche) et sa flèche rouge (en haut à droite), avant de reprendre sa taille réelle pendant que le
logo complet apparaît exactement autour de lui. Environ 3 s, une tape sur l'écran la passe.

**`APP_VERSION` (`appli-techniciens/index.html`, dans le `<head>`, format `AAAA-MM-JJ-n`) —
À INCRÉMENTER À CHAQUE MISE EN LIGNE.** C'est le même type de piège que `SEED_VERSION` :
- L'animation ne se joue qu'une fois par valeur d'`APP_VERSION`, mémorisée dans le
  `localStorage` du téléphone (clé `telsam_version_vue`). Sans bump, le technicien qui a déjà
  ouvert l'appli ne voit rien — et surtout il perd le seul signal lui disant qu'il a bien reçu la
  nouvelle version et non l'ancienne restée en cache (cf. la règle sur le cache navigateur, plus
  haut). C'est la raison d'être de cet écran autant que l'effet visuel.
- La **date affichée en dessous est déduite d'`APP_VERSION`** : un seul endroit à modifier, pas
  deux. Ne pas réintroduire de date écrite en dur.
- Première ouverture (rien en mémoire) → « Bienvenue — version du … » ; version différente →
  « Mise à jour installée — … ». La version est notée comme vue **dès le début** de l'animation,
  pour qu'un technicien qui ferme l'onglet en cours de route ne la revoie pas.
- Si le `localStorage` est indisponible, on n'affiche RIEN plutôt que de rejouer l'animation à
  chaque ouverture : une animation subie à chaque lancement serait vite insupportable.

**`APP_NOUVEAUTE` — la phrase qui dit CE QUI CHANGE, à réécrire à chaque mise en ligne**
(demandé par Patrice le 26/08/26 : « s'il voit une animation sans savoir, il va se demander
pourquoi »). Elle s'affiche sous la date, **uniquement sur une mise à jour** — à la première
ouverture il n'y a pas d'« avant » à comparer, donc rien à annoncer.
- **C'est Claude qui l'écrit, jamais Patrice.** Point confirmé par lui le 26/08/26 : il ne doit
  pas avoir à la dicter. C'est moi qui fais les modifications, donc je sais ce qui change — je
  rédige la phrase et je la fais figurer **dans le résumé montré avant de pousser**, au même
  titre que le reste. Patrice n'intervient que s'il veut la reformuler. **Ne jamais la lui
  réclamer comme un préalable** : ce serait lui rendre une corvée dont l'automatisation était
  justement le but.
- Écrite **de son point de vue**, pas du mien : « Tes chantiers de la semaine sont à jour »,
  « Nouveau PGO sur Portet — vérifie tes documents avant de partir », et non « refonte de
  `buildWeekBuckets` ». Une ligne ou deux, pas plus : c'est lu debout sur un parking.
- **Honnête à chaque fois, y compris « rien qui te concerne cette semaine ».** Une formule creuse
  répétée cesse d'être lue en trois semaines — même mécanique que les alertes qui crient au loup,
  déjà constatée sur le suivi.
- Le hook pre-commit **refuse un bump d'`APP_VERSION` sans changement d'`APP_NOUVEAUTE`** : sinon
  le technicien relirait le message de la fois précédente en le croyant à jour, ce qui est pire
  que pas de message du tout.

**L'écran ne peut PAS être sauté, et il ne se ferme que par le bouton « Continuer »**
(décision de Patrice, 26/08/26). La tape pour passer a été retirée : c'était exactement le geste
d'un technicien pressé, et elle effaçait la phrase avant qu'il l'ait lue. Le bouton n'apparaît
qu'à la fin de l'animation, donc une tape réflexe pendant celle-ci ne fait rien. Conséquence à
garder en tête si on retouche les délais : **la phrase doit rester affichée le temps qu'il
faut** — c'est le technicien qui décide quand passer, pas un minuteur.
**Filet de sécurité : un `setTimeout` de 30 s ferme l'écran quoi qu'il arrive.** Si le bouton ne
répondait pas, personne ne doit se retrouver bloqué devant un logo sur un chantier. Ne pas le
retirer en « simplifiant ».

**RECHARGER LA PAGE N'EST PAS UNE MISE À JOUR — et l'animation ne doit PAS s'y rejouer.**
C'est voulu, ce n'est pas un défaut à corriger. Patrice l'a signalé le 26/08/26 : il ouvrait
`index.html` en local, appuyait sur la flèche de rechargement du navigateur et ne revoyait pas
l'animation. Comportement correct : la version du fichier n'a pas changé, donc rien à annoncer.
Sans cette règle, un technicien se prendrait l'animation à chaque retour sur l'appli dans la
journée — insupportable en trois jours, et le signal perdrait tout son sens à force d'être vu.
Vérifié en conditions réelles le 26/08/26, les trois étapes d'affilée : première ouverture →
animation ; rechargement → rien ; changement réel d'`APP_VERSION` dans le fichier → animation
avec « Mise à jour installée ». **Ne jamais "réparer" l'absence de rejeu au rechargement.**

**Conséquence pratique pour Patrice, à lui redire si besoin :** si un technicien ne voit PAS
l'animation après une mise à jour annoncée, ce n'est pas que ça ne marche pas — c'est qu'il a
encore l'ancienne version en cache. Il ferme complètement l'onglet et rouvre son lien. Constaté
sur le site en ligne le 26/08/26 : la première ouverture, dix minutes après le push, servait
encore la version précédente.

**Les mémoires sont séparées par origine.** Le fichier ouvert en local (`file://`) et le site
`pivot-telsam.github.io` ont chacun leur `localStorage` : un essai de Patrice sur son PC
n'influence en rien ce que voient les techniciens, et inversement.

**Le tracé du « s » est un relevé au pixel du logo, pas un dessin à main levée.** Il vit dans le
même repère que l'image (`viewBox 0 0 550 291`) et son agrandissement est centré sur le centre du
« s » (52.18 % ; 52.41 %). En revenant à l'échelle 1 il se pose donc exactement sur le « s » du
logo — écart mesuré : 0 pixel. Les coordonnées (corps du « s » x 265→309, y 124→181 ; flèche
rouge tip (305.5,128) ; flèche cyan tip (264.5,178) ; épaisseur du trait 8.5) viennent d'un
balayage des pixels de `TELSAM_LOGO_B64`. **Si le logo change un jour, il faut refaire ce relevé**,
pas ajuster à l'œil.

**Deux pièges qui ne se voient qu'à l'écran, aucun test automatique ne les attrape :**
1. **La taille se pose sur `.sp-holder`, l'image la remplit en `width:100%`.** L'inverse
   (`width:240px; max-width:70vw` sur l'image) fait que le porteur garde 240 px pendant que
   l'image rétrécit : le calque du tracé, en `width:100%`, suit le porteur et le « s » retombe
   **entre le « a » et le « m »**. Vécu le 26/08/26, repéré par Patrice à l'œil alors que mon
   contrôle chiffré disait « 0,2 pixel » — parce que je comparais le tracé au logo **dans le même
   repère**, ce qui était vrai d'avance, au lieu de vérifier que le calque se superposait à
   l'image. Toujours mesurer en **coordonnées d'écran** (`getBoundingClientRect`), jamais dans le
   repère du dessin.
2. **`aspect-ratio:550/291` sur le porteur + `width="550" height="291"` sur l'image** réservent la
   hauteur AVANT que le logo soit chargé. Sans ça le porteur a une hauteur nulle au démarrage,
   le calque aussi, et le « s » se dessine décalé puis saute en place quand l'image arrive. Très
   visible ici : le logo est injecté par le script principal, tout à la fin d'un fichier de
   300 Ko. Ne pas retirer ces deux garde-fous.

Le reste est sans surprise : tout est en local (aucune police ni image à télécharger, ça marche
sans réseau), les classes sont préfixées `sp-` pour ne rien heurter dans l'appli, et un téléphone
réglé sur « réduire les animations » voit le logo fixe pendant une seconde, sans effets.

## Habilitations des techniciens
Le bouton « 📁 Ouvrir mes habilitations » de l'appli pointe sur **un dossier** :
`Telsam Fibre/HABILITATIONS/HABILITATIONS 2026 signées` (`HABILITATIONS_LINK` dans
`appli-techniciens/index.html`, lien public en lecture seule, le même pour tous).
- **Un lien vers un dossier sert toujours son contenu du moment.** Pour mettre à jour une
  habilitation, Patrice dépose ou remplace le fichier dans ce dossier : rien à recréer, rien à
  demander à Claude, aucune modification de l'appli.
- **Un ZIP figé y a été supprimé le 25/08/26.** Il datait du 13/04 et divergeait déjà des fichiers
  individuels (il contenait un `EL ABBASSI.pdf` que Patrice venait de remplacer par une version à
  jour). Un technicien qui l'aurait téléchargé aurait récupéré des habilitations périmées sans le
  savoir. **Ne pas en recréer** : les fichiers individuels suffisent, un technicien veut le sien.
- Le contrôle de suppression a refusé deux fois avant d'accepter, parce que le ZIP contenait un
  fichier introuvable ailleurs — c'est ce qui a permis de découvrir le remplacement en cours.
  Garder ce réflexe : ne jamais supprimer une archive sans avoir vérifié que TOUT son contenu
  existe encore hors de l'archive.
- Limite connue et assumée : ce dossier est partagé tel quel, donc **chaque technicien voit les
  habilitations de tous les autres**. Signalé à Patrice le 25/08/26 ; à re-proposer seulement s'il
  l'évoque. Beaucoup de fichiers portent des suffixes `-1`, `-2` sans qu'on sache ce qui les
  distingue — Patrice a commencé à les renommer explicitement (`El Abbassi Morad - pass RTE 2025`).

## Documents du Bureau destinés aux techniciens

**`Appli_TELSAM_Consigne_Technicien.docx` / `.pdf`** — la notice qui part avec le mail de mise en
service aux 13 techniciens.

**RÈGLE — elle est écrite au nom de l'équipe, jamais au nom de Patrice.** Réécrite ainsi le
27/08/2026 à sa demande, pour qu'il puisse la diffuser au nom de TELSAM : « le code d'accès que
**nous** t'avons communiqué », « **que nous tenons** à jour », « toute précision utile **pour
nous** », « **préviens-nous** ». Ne jamais réintroduire son prénom dans ce document.

**Elle décrit l'écran d'ouverture animé**, en deux puces : ce qu'il affiche (date de la version,
phrase de nouveauté, bouton « Continuer ») et ce qu'il signifie (seul repère qu'une mise à jour a
bien été reçue). **Toute modification de cet écran dans `appli-techniciens/index.html` oblige à
relire ces deux puces.** Cas réel : l'animation ne se passe plus d'une tape depuis le 26/08 — la
notice a gardé la phrase « une tape sur l'écran la passe » jusqu'au 27/08, et elle serait partie
telle quelle aux 13 techniciens.

**Deux pièges de mise en forme dans ce fichier** : la puce « •  » fait partie du **texte**, pas du
format — remplacer un paragraphe entier la fait disparaître ; et Word convertit en apostrophes
typographiques (’) le texte que l'on insère, si bien qu'un motif de recherche écrit avec une
apostrophe droite (') ne retrouve plus ce qu'on vient d'écrire. Préférer des motifs sans
apostrophe. Régénérer le PDF après chaque modification (`SaveAs2($chemin, 17)`).

**Mise à jour du 27/08/2026** (à la demande de Patrice, en même temps que le portail par mot de
passe) : section « 1. Te connecter » réécrite pour le nouvel accès (**mot de passe personnel**, il
n'y a plus de lien perso ni de code partagé) ; dépôt photos devenu « **envoi et nommage** » avec
l'exemple de nom (« Prénom Nom - lieu ») ; nacelle « une seule par jour **et par boîte** » ; phrase
« Ne pars pas sur un chantier dont les documents ne sont pas en règle » **retirée**. La correction
de la ligne du mot de passe a d'abord été tentée par Find/Replace et a **corrompu le paragraphe**
(cf. le piège des apostrophes en « Règles de prudence ») — refaite par remplacement de la plage du
paragraphe. Sauvegarde de l'original dans `Appli_TELSAM_Consigne_Technicien_backup_2708.docx`.

**`Memo_Exploitation_TELSAM.docx`** — le « qui fait quoi » de Patrice, 11 sections depuis le
27/08/2026. À tenir à jour dès qu'une tâche récurrente change de main : c'est le document qu'il
relit pour savoir ce qu'il doit faire et ce que je fais. Il n'en existe pas de PDF, et Patrice
n'en veut pas.

## Règles de prudence
- **PIÈGE D'ENCODAGE — ne JAMAIS relire un fichier accentué avec `Get-Content` sans `-Encoding
  UTF8`.** En PowerShell 5.1, `Get-Content -Raw` lit en ANSI (Windows-1252). Enchaîné avec
  `Set-Content -Encoding UTF8`, cela **double l'encodage de tout le fichier** : « scindée »
  devient « scindÃ©e ». Commis le 25/08/26 sur `suivi_chantiers_205.html` pour un simple bump de
  `SEED_VERSION` — 4 202 séquences abîmées, plus un seul accent correct, et **le fichier a été
  committé et poussé dans cet état**. Si Patrice l'avait envoyé par Teams entre-temps, ses
  collègues l'auraient reçu illisible.
  - **À utiliser à la place** : `[IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)` puis
    `[IO.File]::WriteAllText($f, $t, (New-Object Text.UTF8Encoding $false))`, ou `sed` via Bash,
    ou l'outil Edit. Ces trois voies ont été utilisées toute la journée sans incident.
  - **Réparation si ça se reproduit** : le double encodage s'inverse exactement —
    `[Text.Encoding]::UTF8.GetString([Text.Encoding]::GetEncoding(1252).GetBytes($t))`. Vérifier
    ensuite `Ã` = 0 ET `�` = 0, puis comparer fiche par fiche avec le dernier commit sain.
  - **Contrôle systématique après toute écriture sur un fichier accentué** : compter les `Ã` et
    les `�`. C'est une seule commande et ça aurait évité le commit fautif.
- **PIÈGE POWERSHELL — une fonction qui affiche un message POLLUE sa valeur de retour.** Tout ce
  qu'une fonction émet part dans le flux de sortie : `"  OK  $libelle"` suivi de
  `return $texte.Replace(...)` renvoie **les deux**, et `$fiche = MaFonction ...` reçoit un
  tableau de 2 éléments, pas une chaîne. PowerShell 5.1 ne proteste pas : `.Replace()` sur un
  tableau fait de l'énumération de membres et continue, puis la concaténation finale colle le
  message dans les données. **Commis le 27/08/2026** sur `suivi_chantiers_205.html` : six « OK … »
  injectés au milieu de `SEED_DATA`, JSON cassé, fichier restauré par `git checkout --`.
  - **À faire** : les messages passent par `Write-Host` (qui n'écrit pas dans le flux de sortie),
    jamais par une chaîne nue. Et vérifier le type avant de rendre :
    `if ($sortie -isnot [string]) { throw ... }`.
  - **Contrôle après écriture** : reparser `SEED_DATA` avec `ConvertFrom-Json`. Un fichier qui ne
    parse plus est le seul signal fiable — le script, lui, avait affiché « Ecrit. » sans broncher.
- **PIÈGE WORD — la limite de 255 caractères vaut aussi pour le TEXTE DE REMPLACEMENT.** Elle est
  connue pour le motif recherché ; elle s'applique de la même façon à `Replacement.Text`, et Word
  lève une `COMException` « Paramètre de la chaîne trop long ». Rencontré le 27/08/2026 en
  ajoutant deux puces au mémo d'exploitation. **Parade** : découper en plusieurs remplacements
  courts qui s'accrochent l'un à l'autre — insérer d'abord une phrase tronquée, puis la compléter
  en cherchant sa fin.
- **PIÈGE WORD — insérer un paragraphe vide puis lui affecter son texte NE MARCHE PAS.**
  `InsertParagraphAfter()` puis `$p.Range.Text = "…"` remplace la marque de paragraphe : le texte
  **fusionne avec le paragraphe voisin**. Résultat le 27/08/2026 sur le mémo d'exploitation :
  17 paragraphes ajoutés, 17 paragraphes collés à leurs voisins, document illisible
  (« 7. Chaque période de paie — les heures par chantier9. Ce que je ne fais pas tout seul »).
  - **À faire** : insérer **tout le bloc d'un coup**, chaque ligne portant sa propre marque de
    paragraphe — `$r = $doc.Paragraphs.Item($cible).Range; $r.Collapse(1); $r.InsertBefore($texte + "`r")` —
    puis boucler sur les paragraphes créés pour leur appliquer la mise en forme, copiée depuis un
    paragraphe existant du même rôle (`$p.Format = $ref.Format`, taille, gras, police).
  - **Contrôle** : compter les paragraphes avant et après. C'est ce qui a révélé l'erreur
    (72 au lieu de 87 attendus), pas la lecture du texte.
- **PIÈGE WORD — Find/Replace bute sur les apostrophes typographiques.** Un `Find.Execute` dont
  le texte recherché contient une apostrophe courbe (’, U+2019) peut ne remplacer qu'un fragment
  et **corrompre le paragraphe**. Vécu le 27/08/2026 sur la consigne technicien : « Entre le code
  d’accès… » est devenu « Accès’accès que nous… ». De plus, Word **convertit en apostrophe
  courbe** toute apostrophe droite (') insérée par script — donc un contrôle qui recherche
  `l'appli` (droite) ne retrouve pas `l’appli` (courbe) et croit à tort le remplacement raté.
  - **À faire pour une ligne contenant des apostrophes** : éviter Find/Replace ; localiser le
    paragraphe (`foreach ($p in $doc.Paragraphs)` sur un ancrage sans apostrophe, ex. un mot
    unique comme « communiqué »), puis remplacer son texte via la plage en **excluant la marque
    de paragraphe** — `$r = $p.Range; $r.End = $r.End - 1; $r.Text = "…"`. Insensible aux
    apostrophes, et pas de fusion de paragraphes.
  - **Contrôle** : réextraire le texte du `.docx` (dézipper `word/document.xml`) et vérifier que
    le neuf est présent ET l'ancien absent, en tenant compte des DEUX formes d'apostrophe.
- Toujours montrer le résumé des changements à Patrice AVANT de committer/pousser, sauf demande explicite contraire.
- En cas de nom de chantier ambigu ou de dossier Dropbox introuvable/multiple, ne jamais deviner —
  poser la question à Patrice.
- Valider la syntaxe JS (node --check) et le JSON (SEED_DATA) avant toute livraison.
- **Avant de committer un changement à `SEED_DATA` dans suivi-chantiers, incrémenter `SEED_VERSION`.**
  Sans ça, la modification est invisible pour les collègues qui ont déjà chargé une version antérieure.
  Cette règle a déjà été oubliée deux fois malgré sa documentation (20/08 puis 21/08) — un hook
  git local (`scripts/check-seed-version.sh`, installé dans `.git/hooks/pre-commit`) bloque
  maintenant tout commit qui modifie un enregistrement `SEED_DATA` sans toucher `SEED_VERSION`
  dans le même commit. À réinstaller après un nouveau clone du dépôt :
  `cp scripts/check-seed-version.sh .git/hooks/pre-commit && chmod +x .git/hooks/pre-commit`
  (les hooks ne sont pas versionnés par Git lui-même). Toujours annoncer explicitement dans le
  résumé de commit "SEED_VERSION bumpé : oui/non" quand `SEED_DATA` est modifié.
- **Avant de committer un changement à `appli-techniciens/index.html`, incrémenter `APP_VERSION`.**
  Même logique que `SEED_VERSION`, côté technicien : sans bump, l'écran d'ouverture ne se rejoue
  pas et le technicien n'a aucun signe qu'il a bien reçu la nouvelle version. Cf. la section
  « Écran d'ouverture animé ».
  Un hook git local (`scripts/check-app-version.sh`, installé dans `.git/hooks/pre-commit` le
  26/08/26) bloque tout commit qui touche `index.html` sans incrémenter `APP_VERSION`. Il vérifie
  aussi le **format `AAAA-MM-JJ-n`** : la date montrée au technicien sous le logo en est déduite,
  donc un format cassé lui afficherait une date fausse sans erreur nulle part ailleurs. Les trois
  cas (oubli / format cassé / bump correct) ont été testés en conditions réelles.
  À réinstaller après un nouveau clone du dépôt, comme celui de `suivi-chantiers` :
  `cp scripts/check-app-version.sh .git/hooks/pre-commit && chmod +x .git/hooks/pre-commit`
  (les hooks ne sont pas versionnés par Git lui-même — ce garde-fou ne protège que la machine où
  il est installé).
