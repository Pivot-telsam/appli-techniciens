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
Toujours : `telsamfibre > RTE > [Ligne aérienne | Ligne souterraine | Postes | Fibrage | Audit FO | Arteria]`
puis dossier du chantier, puis sous-dossier `Documents Telsam` (contient les NDS).
Ne jamais sortir de cette arborescence pour ce type de recherche.

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
- Le code d'accès partagé (`GATE_CODE_HASH`, hash SHA-256) est un confort d'accès, PAS une vraie
  sécurité — ne jamais le présenter comme une protection réelle des données techniciens/PDP.

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
- Renommage Dropbox pas encore fait pour Chaineau-Cordy-Lamotte au 21/08/26 (dossiers encore
  "LOT 1 RODA"/"LOT 2 SELT") — Patrice le fera progressivement, ne pas renommer à sa place sans
  demande explicite. Cette règle de structuration (numérotation, PDP/PGO commun, NDS/IST/Brief
  séparés) s'applique à ce cas ET à tout futur chantier multi-lots similaire.

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
4. **IST — uniquement si elle est signée/validée par RTE** (`ist.valideRTE == true`). Si une IST
   existe pour ce chantier mais n'est PAS signée par RTE : NE PAS la mettre dans App Tech, et à
   la place ajouter une alerte (niveau `warn`) dans `suivi-chantiers` signalant l'IST non signée
   par RTE, avec `ist.valideRTE:false`.
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

## Veille documentaire automatique (mise en place le 25/08/2026)

**RÈGLE — au début de toute session de mise à jour, lire le rapport de veille**
`C:\Users\patrice.pivot\Desktop\TELSAM-apps\veille\RAPPORT.md` (version exploitable :
`dernier.json` à côté). Il liste les PDP/PGO/IST nouveaux ou modifiés depuis le passage
précédent, avec pour chacun le chantier concerné et s'il est déjà présent dans son dossier
App Tech. Ne pas attendre que Patrice signale un document : le vrai point faible du circuit
n'était pas la mise à jour elle-même mais le fait que tout reposait sur lui pour y penser.

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
- Il tient un **index de référence** (`veille\etat.json`) de TOUS les documents de sécurité
  (~1 320), pas seulement des récents — sinon chaque passage rapporterait tout l'existant comme
  nouveau. Ne pas supprimer ce fichier sans raison ; s'il disparaît, le passage suivant repart en
  « première passe » et ne liste que les 7 derniers jours.
- Extensions volontairement restreintes (`.pdf .doc .docx .xls .xlsx .xlsm .zip`) : les `.dwg`,
  `.bak`, `.dwl` sont des fichiers de travail AutoCAD qui bougent en permanence et noyaient le
  rapport.
- Le motif IST est borné (`(^|[^a-z])ist([^a-z]|$)`) pour ne pas matcher « Liste », « Existant ».
- Dédoublonnage par nom+taille : un même document est souvent copié dans plusieurs dossiers
  (PGO racine, PGO du sous-chantier, App Tech) — il est rapporté une fois, avec tous ses chemins.
- Sous `Audit FO`, les chantiers sont rangés par direction interrégionale (`DI Marseille`,
  `DI Toulouse`) : le dossier chantier est donc un niveau plus bas que dans les autres catégories.
- `Fibrage` figure dans la liste des catégories de CLAUDE.md mais **n'existe pas** comme dossier
  de premier niveau (les chantiers Fibrage sont sous `Ligne aérienne`) — le rapport le signale en
  note de bas de page à chaque passage, c'est normal.

**Ce que la veille ne remplace pas.** Elle détecte et signale, elle ne met rien à jour. La lecture
des documents, la mise à jour des fiches et la copie dans App Tech restent faites en session, sous
le contrôle de Patrice. Et deux choses resteront toujours locales quoi qu'il arrive : les PGO/PDP
au format `.xlsm` (Excel requis) et la génération des briefs techniciens (Word requis).

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

## Règles de prudence
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
