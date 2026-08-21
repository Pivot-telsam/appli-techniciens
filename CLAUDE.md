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
- `documentsTerrain` (ancien système pdp/pgo/briefing, fallback pour chantiers non migrés)

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
- `MTFO_STANDALONE` (suivi-chantiers) : rappels MTFO sans fiche chantier associée, décision
  volontaire de Patrice pour les lignes du planning sans PDP/PGO à gérer. Ne pas créer de fiche
  chantier pour "régulariser" ces lignes.
- Côté `appli-techniciens` : les badges de statut PDP/PGO ne s'affichent que si le chantier a un
  `documentsAppTech` renseigné (nouveau système) — un chantier encore sur l'ancien système
  (`documentsTerrain`) n'affiche jamais ces badges même si `pdp`/`pgo` sont remplis. Absence de
  badge = chantier non migré, pas forcément donnée manquante.
- Le code d'accès partagé (`GATE_CODE_HASH`, hash SHA-256) est un confort d'accès, PAS une vraie
  sécurité — ne jamais le présenter comme une protection réelle des données techniciens/PDP.

## Système de numérotation (`numero` dans suivi-chantiers, format `26-0XX`)
C'est le numéro de référence que Patrice utilise pour organiser Dropbox (il renomme progressivement
les dossiers avec ce préfixe, ex: "26-055 - Fleyriat - Brou - La Cluse"). Pour tout nouveau chantier,
attribuer le prochain numéro disponible et le communiquer à Patrice.

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
