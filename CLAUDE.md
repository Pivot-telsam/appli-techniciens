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

**LE FICHIER QUI FAIT FOI EST CELUI DE TEAMS, ET IL SE LIT DIRECTEMENT** (mis en place le
01/09/2026) :
`https://telsam.sharepoint.com/sites/RTE/Shared Documents/RTE/Planning RTE 2026.xlsx`.
Excel l'ouvre **en lecture seule** avec la session Windows de Patrice — `Workbooks.Open($url, 0,
$true)` — donc les couleurs sont lisibles comme sur un fichier local. **Patrice n'a plus à copier
le planning sur son Bureau ni à me l'envoyer dans la conversation.** La copie du Bureau
(`C:\Users\patrice.pivot\Desktop\Planning RTE 2026.xlsx`) ne sert plus que de repli quand le réseau
TELSAM n'est pas joignable — elle est presque toujours en retard (le 01/09 elle datait du 27/08
alors que le fichier partagé avait été modifié le matin même).

**Le planning est un travail d'équipe dans Teams : Pierre Brillou, Christian Cazenave,
Ahmed Hamouch, François Vidal et Patrice le modifient ensemble.** Ils sont aussi les seuls à le
consulter au quotidien (le PDG y a accès). **Teams reste donc le maître** : la grille du suivi en
est un reflet — **modifiable depuis le 02/09/2026 via la base commune** (cf. « ÉTAPE 1 »), mais un
reflet : ce qu'on décide dans le suivi ne part PAS dans Teams et reste à y reporter. Cette ligne
disait jusqu'au 03/09/2026 « lecture seule, ne pas proposer de la rendre modifiable » — c'était vrai
avant l'étape 1, elle ne l'était plus.

Structure de la feuille "Feuil1" :
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
- **PIÈGE WEEK-END** : les colonnes samedi/dimanche ont un fond gris uniforme sur toutes les
  lignes, que la lecture couleur confond avec un chantier (« tout le monde le même chantier le
  S/D », y compris les encadrants). **Ne lire les affectations que sur les jours ouvrés (L-V).**
- **PIÈGE DE LA CELLULE SANS FOND — c'est `ColorIndex` qui fait foi, PAS `Color`** (relevé le
  01/09/2026). Une cellule **sans remplissage** renvoie `Interior.Color = 16777215`, exactement la
  même valeur qu'une cellule réellement peinte en blanc. Or **il existe des lignes-projet sans
  fond** : ce sont de simples notes (le 01/09 : « Joncquiere St Cézaire Thym et recette », « DATA4 -
  MARCOUSSIS fin chantier », « livraison Tourets Arnage Ecommoiy »). Les inclure dans la table des
  couleurs colle **tout le monde** sur le mauvais chantier — un premier essai plaçait ainsi les cinq
  encadrants sur Joncquière. Le test qui tranche est `Interior.ColorIndex = -4142` (`xlNone`).
  Valeurs relevées sur ce classeur : `-4142` = pas de fond, `3` = rouge (CP / CP paternité),
  `7` = gris du week-end, tout le reste = un chantier.
- **RÈGLE FONDAMENTALE POSÉE PAR PATRICE LE 01/09/2026 : UNE COULEUR NE VAUT QUE DANS SA SEMAINE.**
  « Les couleurs ne valent que pour la semaine en cours, c'est-à-dire dans la même colonne où elles
  sont. » Le jaune de la S35 et le jaune de la S36 désignent couramment **deux chantiers
  différents**. Ce n'est pas un défaut de tenue du fichier : les chantiers sont placés des mois à
  l'avance, et en coloriant une ligne pour dans quatre mois on ne peut pas savoir de quelle couleur
  seront ses voisines — **la couleur est donc choisie au hasard**, et les collisions sont inévitables.
  Conséquences, toutes obligatoires :
  - **Construire la table couleur → chantier SEMAINE PAR SEMAINE**, jamais pour l'année. Une table
    globale mélangerait des chantiers sans aucun rapport.
  - **Ne jamais reporter un libellé d'une semaine à une autre non contiguë.** Le libellé n'est écrit
    qu'au démarrage d'une ligne, donc un report est nécessaire — mais uniquement depuis la semaine
    **immédiatement précédente** (même ligne, même couleur), le seul cas légitime : un chantier qui
    se poursuit. `planning-rte.ps1` borne ce report ; sans cette borne, une ligne recoloriée du même
    vert quatre mois plus tard héritait du libellé de l'ancien chantier, en silence.
  - **Deux lignes-projet de la même couleur dans la MÊME semaine restent possibles** (vu le
    01/09/2026 : Chaineau lot RODA et lot SELT en vert, Fleyriat 1ʳᵉ phase et « oppc pyl125 » en
    jaune — 83 cases sur l'année). Là, la case d'un technicien n'a qu'une couleur et ne permet
    donc pas de trancher : le script garde les deux libellés (`a`) et la grille marque un ⚠ plutôt
    que de choisir au hasard. **Ne pas présenter ça à Patrice comme une erreur à corriger** — c'est
    la conséquence assumée de sa méthode. Le signaler seulement si les deux libellés désignent des
    chantiers vraiment sans rapport, là où l'ambiguïté aurait une conséquence.
- **`Nassime EL GARTILI` ne fait PLUS partie de l'effectif** (dit par Patrice le 27/08/2026) —
  **mais il reste coloré sur le planning RTE**, Patrice ne l'en retirera pas. C'est donc à moi de
  **l'exclure systématiquement** de toute lecture d'affectation : ne jamais l'ajouter à
  `TECH_RANGES`/`REAL_DAYS` ni à la liste des techniciens. Vu sur Portet (S35) et Fleyriat (S36) :
  ignoré. De même, **`Pierre Brillou` et `François VIDAL` sont des encadrants, pas des
  techniciens** — ne pas les traiter comme des techniciens (comme Christian Cazenave, Patrice,
  Ahmed Hamouch). **Nuance ajoutée le 02/09/2026 : François VIDAL est bien un encadrant pour la
  lecture du planning, mais il est aussi l'UN DES DEUX SEULS habilités MTFO, avec Vincent PERRIN.**
  Sur un chantier de mesures thermiques FO, c'est donc lui ou Vincent Perrin qui va sur site : ne
  pas l'écarter en le rangeant mécaniquement parmi les encadrants.
- **`Vincent BENIGAUD` N'EST PAS DE TELSAM : c'est un interlocuteur INEO, donc côté CLIENT.**
  Règle posée par Patrice le 02/09/2026 (« attention de ne pas faire la confusion, c'est
  important »). Il ne doit jamais entrer dans `TECH_RANGES`, dans le portail de l'appli, dans les
  feuilles d'heures, ni dans une phrase parlant de « notre effectif ». **Le seul Vincent de
  l'effectif est Vincent PERRIN.** L'erreur a été commise le 02/09/2026 : une alerte de la fiche
  26-002 annonçait « deux Vincent dans l'effectif, Vincent PERRIN et Vincent Benigaud » — corrigée
  le jour même. Vérification faite à cette occasion : il n'apparaissait nulle part ailleurs comme
  technicien, sa seule autre mention le nomme correctement « Vincent Benigaud (INEO) ».

## Grille Planning — onglet du suivi (mis en place le 01/09/2026)

L'onglet **« Planning »** du suivi (bouton `btnViewTechs`, vue `techs`) **remplace l'ancienne vue
« Techniciens »**, qui ne listait que des pastilles par technicien et dont Patrice disait qu'elle
« ne servait pas à grand-chose ». C'est maintenant une **grille technicien × jour ouvré**, comme le
planning RTE : une ligne par personne, cinq colonnes (L-V), la couleur du chantier en fond.

**Ce qu'elle apporte par rapport au fichier Excel** — c'est ce qui justifie de l'avoir faite :
- le **nom du chantier est écrit dans la case** (l'Excel ne met que la couleur, c'est ce qui avait
  fait rater 9 affectations le 20/08/2026) ;
- chaque case **ouvre la fiche du chantier** d'un clic (`ouvrirFiche(cid)`, utilisable depuis
  n'importe quelle vue) ;
- la colonne du jour est surlignée, les encadrants sont dans un bloc à part, les CP en rouge pâle,
  et une légende liste les chantiers de la semaine avec leurs avertissements.

**La donnée : `PLANNING_RTE`**, constante à côté de `SEED_DATA`, **refaite de zéro à chaque
passage** de `scripts/planning-rte.ps1` — ne jamais y écrire à la main. Elle ne vit PAS dans
IndexedDB : pas de `SEED_VERSION` à bumper, et aucun état local des collègues effacé. Forme :
`{maj, source, personnes:[{nom,role}], libelles:[{t,c}], cellules:{"AAAA-MM-JJ":{"Nom":{l|s,a?,n?}}}}`
— `l` = index de libellé, `s` = texte d'absence, `a` = libellés en concurrence sur la couleur,
`n` = note écrite dans la cellule du technicien. 90 Ko pour l'année entière.

**`scripts/planning-rte.ps1`** : `-Injecter` met à jour le HTML, `-Local` force la copie du Bureau.
~35 s (la couleur ne se lit pas en bloc : `Range.Interior.ColorIndex` revient vide dès que la plage
est mélangée, donc un appel COM par cellule ; les textes, eux, se lisent en un seul `Value2`).

**PIÈGE — Excel refuse d'ouvrir un classeur si un autre DU MÊME NOM est déjà ouvert**, même quand
l'un est sur SharePoint et l'autre sur le disque. Vécu le 01/09/2026 : Patrice avait
« Planning RTE 2026.xlsx » ouvert depuis 10h45, la lecture Teams a échoué et **le repli a servi une
donnée du 27/08 sans que ce soit flagrant**. Trois parades, toutes en place :
1. **Emprunter le classeur déjà ouvert** (`GetActiveObject('Excel.Application')` puis recherche du
   classeur par son nom) : c'est le fichier que Patrice a sous les yeux, donc la source la plus à
   jour. Vérifié le 01/09 : son `FullName` était bien l'URL Teams.
   **Dans ce cas il ne faut NI fermer le classeur NI quitter Excel** — c'est le rôle du drapeau
   `$emprunte` dans le bloc `finally`. Le fermer ferait perdre à Patrice son travail non enregistré.
   `GetActiveObject` est **peu fiable** (Excel n'est pas toujours inscrit au ROT, erreur
   `MK_E_UNAVAILABLE`) : c'est une chance à saisir, jamais une garantie.
2. **La grille dit d'où vient la donnée.** Sur repli, un bandeau rouge annonce que le fichier Teams
   n'a pas pu être lu et donne la date d'enregistrement de la copie. Champ `teams` (booléen) +
   `datePlanning` dans `PLANNING_RTE`. Les DEUX branches ont été testées à l'écran le 01/09/2026,
   `-Local` servant à provoquer le repli exprès.
3. La date n'est relevée que pour un fichier **local** (`LastWriteTime`) : sur un fichier SharePoint
   `BuiltinDocumentProperties('Last Save Time')` revient **vide** (vérifié), et de toute façon lire
   le fichier partagé donne par construction son état du moment. La date ne sert qu'au cas du repli.

La voie `\\telsam.sharepoint.com@SSL\DavWWWRoot\...` (copie WebDAV vers un nom temporaire, qui
contournerait le conflit de noms) a été essayée le 01/09/2026 : **le chemin n'est pas résolu**, même
avec le service WebClient démarré. Ne pas la re-tenter sans nouvelle information.

**Les noms des listes du script sont écrits SANS ACCENT et comparés sans accent** (`ClePersonne`).
PowerShell 5.1 lit un `.ps1` sans BOM comme de l'ANSI : `'François VIDAL'` écrit avec sa cédille ne
correspondait plus au nom lu dans la feuille, et **il s'est retrouvé classé parmi les techniciens**
le 01/09/2026 — repéré à l'écran, pas par un contrôle. Un garde-fou prévient maintenant si un nom
déclaré encadrant est absent de la feuille.

**Rattachement case → fiche, dans cet ordre** (`chantierDuLibelle` puis `chantierParPresence`) :
1. le **numéro complet** trouvé dans le libellé, sous-lot compris (`26-036-1` désigne le lot RODA,
   pas le lot SELT) ;
2. le numéro de base **s'il ne désigne qu'une seule fiche** ;
3. sinon `TECH_RANGES` : sur quel chantier cette personne était-elle ce jour-là, d'après les
   affectations déjà relevées — **et uniquement si la réponse est unique**.
Le point 3 n'est pas un luxe : Patrice n'a commencé à noter les numéros au planning que récemment,
donc sans lui les semaines antérieures à fin août n'auraient **aucun** lien cliquable (mesuré :
0 case liée sur 57 en S16, 55 sur 57 avec le repli). **Ne jamais ajouter un rapprochement par
ressemblance de libellé** : c'est ce qui enverrait un jour vers le mauvais lot.

**La lecture quotidienne EST branchée sur la tâche de 7h30** — étape « Planning RTE (Teams) » de
`matin.ps1`, volontairement l'avant-dernière (voir « La matinée automatique »). Cette ligne disait
le contraire jusqu'au 03/09/2026, alors que la refonte de `matin.ps1` du 02/09 l'avait intégrée :
une règle périmée dans ce fichier est aussi trompeuse qu'une règle absente.
Le fichier `scripts\Ajouter le planning a la veille du matin.cmd` n'a donc plus d'objet.
**L'action doit rester la DERNIÈRE de la chaîne** : c'est la seule qui ouvre Excel sur un fichier en
ligne, donc la seule qui pourrait rester bloquée sur une demande d'authentification — en dernière
position, un blocage ne retarderait ni la veille, ni le récap du matin, ni son envoi par mail.
**Leçon de forme, redite par Patrice le 01/09/2026** : une commande PowerShell collée dans une
réponse ne lui suffit pas (« pour ce genre de manip, il faut que tu sois beaucoup plus précis dans
tes demandes… ou que tu m'amènes directement au bon endroit »). Lui fournir un fichier à
double-cliquer ou un chemin exact, jamais un bloc de code à recopier.

### Le planning alimente le reste du suivi (branché le 01/09/2026)

Jusqu'à ce jour la même information vivait **deux fois** : le planning relu tout seul, et
`REAL_DAYS`/`TECH_RANGES` que je recopiais à la main chaque semaine — dans les DEUX dépôts. C'est ce
doublon qui a déjà divergé deux fois cette année. `planning-rte.ps1` publie donc désormais la
traduction « chantier → jours de présence », **calculée une seule fois** et consommée telle quelle
par tous :

| champ de `PLANNING_RTE` | forme | qui s'en sert |
| --- | --- | --- |
| `presence` | `{ id: [jours] }` | `joursPresenceReels()` (Gantt + vue par semaine), `veille-documents.ps1`, `recap-matin.ps1` |
| `techRanges` | `{ id: { personne: [jours] } }` — **la forme exacte de `TECH_RANGES`** | `veille-chantiers-actifs.ps1`, qui la fusionne sans rien réécrire |
| `conflits` | désaccords avec la recopie manuelle | affiché dans la vue Planning |

**UNION, jamais remplacement.** Le planning ne porte le numéro du chantier que sur les semaines que
Patrice a annotées (à partir de fin août) : avant, seul `TECH_RANGES` rattache un libellé à une
fiche, et il garde donc l'historique. Sur un contrôle de documents de sécurité, mieux vaut vérifier
un chantier de trop qu'en oublier un. Le jour où tout le planning portera les numéros,
`REAL_DAYS`/`TECH_RANGES` pourront cesser d'être tenus à la main.

**Le rattachement libellé → fiche n'existe qu'à UN endroit** (`planning-rte.ps1`), par le **numéro**
écrit dans le libellé — numéro complet d'abord (`26-036-1` = lot RODA), puis numéro de base s'il ne
désigne qu'une seule fiche. Ne jamais le refaire côté JavaScript ni dans un script de veille : ce
serait recréer le doublon qu'on vient de supprimer.

**Une case ambiguë ne nourrit PAS la présence.** Quand deux lignes-projet partagent la couleur cette
semaine-là, attribuer la présence à l'une des deux serait un tirage au sort. Au premier essai, la
lecture automatique déplaçait ainsi trois jours du lot 1 de Chaineau vers le lot 2. La grille garde
le ⚠, `TECH_RANGES` garde l'arbitrage humain, et 78 cases sur l'année sont écartées à ce titre.

**LE CONTRÔLE DE COHÉRENCE, ET CE QU'IL A TROUVÉ.** À chaque passage, le script compare, personne
par personne et jour par jour, ce que dit le planning et ce que dit la recopie manuelle. Un
désaccord n'est pas un détail : il veut dire qu'une personne apparaît sur **deux chantiers le même
jour**. Dès le premier passage, 17 désaccords, tous en **S34 (17-21/08)** :

- **François PERRIN**, 17-18/08 : planning = Portet 26-051, recopie = Arudy 26-056 ;
- **Bilal HAMOUCH, Younes MOUSSA, Hugues VIRY**, 17-19/08 : planning = Chaineau **lot 2 SELT**
  (26-036-2), recopie = **lot 1 RODA** ;
- **les mêmes**, 20-21/08 : planning = Dambron-Voves 26-031, recopie = lot 1 RODA.

**TRANCHÉ PAR PATRICE LE 01/09/2026 : c'est le PLANNING qui dit vrai, la recopie était fausse.**
Les deux réponses sont « version B ». Ce qui annule et remplace le point noté comme clos le
25/08/2026 (« ces techniciens étaient bien sur RODA ») : le planning partagé a été corrigé depuis,
et c'est lui qu'on relit. Corrigé dans `TECH_RANGES` **des deux dépôts** et dans `REAL_DAYS` :
Bilal / Younes / Hugues retirés du lot 1, placés sur le lot 2 SELT (17-19/08) puis Dambron-Voves
(20-21/08) ; François PERRIN retiré d'Arudy et placé sur Portet toute la semaine. Arudy garde les
17 et 18 dans `REAL_DAYS` (Didier, Benjamin DIRAT et Sid Ahmed y étaient bien) et perd les 19-21.
Contrôle après correction : **aucun désaccord**.

**Et ce que la comparaison des DEUX dépôts a révélé en plus** : l'appli mettait Didier PERRIN,
Benjamin DIRAT et Sid Ahmed BENZAMERA sur Arudy du 19 au 21/08 alors que le suivi ET le planning
disaient Portet. Corrigé aussi — pas besoin d'arbitrage, deux sources sur trois concordaient.
**Le contrôle de cohérence porte donc maintenant sur les deux dépôts**, et son message précise
lequel est en cause (`[suivi]` / `[appli]`). **Vérifié en y glissant une erreur exprès** : sans ce
test, un contrôle qui n'aurait regardé que le suivi aurait paru irréprochable en laissant passer
exactement ce cas.

**La leçon, plus large que ce cas** : la recopie manuelle a été fausse sur DEUX semaines et dans
DEUX dépôts sans que rien ne le signale pendant deux semaines. C'est l'argument le plus fort pour
que `REAL_DAYS`/`TECH_RANGES` cessent d'exister — chaque jour où ils survivent est un jour où ils
peuvent mentir en silence.

### La réserve du bas et le brouillon d'affectation (demandés par Patrice le 01/09/2026)

**La réserve** reproduit le bas de son fichier Excel : toutes les lignes-chantier de la semaine
affichée, **y compris celles où personne n'est encore placé**. C'est ce qui manquait le plus — sans
elle, une semaine à venir paraissait vide alors que les chantiers y sont posés depuis des mois
(S37 : 8 chantiers en attente, S40 : 3). Champ `reserves` de `PLANNING_RTE`, une entrée par semaine,
**dans l'ordre des lignes du fichier** — celui que Patrice a sous les yeux. Chaque pastille indique
combien de personnes sont déjà placées dessus.

**Le brouillon** permet de préparer les affectations comme dans l'Excel : glisser une pastille sur
une case, **ou** cliquer la pastille pour l'« armer » puis cliquer les cases (c'est ce second geste
qui sert vraiment quand on place cinq personnes d'affilée). Les deux gestes existent pour cette
raison, ne pas en retirer un en simplifiant.

**CE QUE LE BROUILLON N'EST PAS — à redire à Patrice, c'est le point qui compte.** Il ne part PAS
dans Teams, ses collègues ne le voient pas (il vit dans l'IndexedDB de SON navigateur), et le
planning partagé reste le maître. Quatre garde-fous, tous nécessaires :
1. il est stocké **à part** de `PLANNING_RTE` (clé `planning_brouillon`), jamais mélangé à la donnée
   venue de Teams ;
2. il est **dessiné différemment** — bordure pointillée, hachures, étiquette « BROUILLON », croix de
   retrait — pour qu'aucun coup d'œil ne puisse le confondre avec une affectation réelle ;
3. quand il se pose **par-dessus une affectation Teams**, il écrit ce qu'il remplace — sur la case
   (« brouillon · remplace 26-051 — Poste de Portet ») et dans le récapitulatif (« à la place de »).
   *Première version corrigée le 01/09/2026 : le dépôt sur une case déjà renseignée était REFUSÉ,
   au motif qu'on ne sait pas retirer l'affectation dans Teams. Conséquence non vue : en semaine
   en cours toutes les cases des techniciens sont remplies, donc **la semaine la plus utile à
   ajuster était la seule impossible à modifier** — Patrice l'a constaté tout de suite. Le bon
   remède n'était pas d'interdire mais de **dire ce qui est remplacé*** ;
4. il **disparaît tout seul** quand l'affectation apparaît dans le planning partagé
   (`nettoyerBrouillonPlanning`, appelée au démarrage). Le ménage compare le **libellé**, pas la
   simple présence d'une case : depuis qu'on peut préparer un changement par-dessus l'existant,
   « la case est remplie » ne veut plus dire « c'est reporté ». Sans ce ménage le récapitulatif
   « à reporter » grossirait sans fin et on cesserait de le lire.

**PIÈGE ÉVITÉ — le brouillon stocke le LIBELLÉ EN TOUTES LETTRES, jamais son index.** La table
`libelles` est reconstruite à chaque relecture du planning et les index se décalent (118 le matin,
172 après l'ajout de la réserve) : un brouillon qui aurait retenu un index se serait mis à désigner
**un autre chantier** du jour au lendemain, sans que rien ne le signale. `chargerBrouillonPlanning`
convertit les anciens brouillons au format index, et jette ceux qu'il ne peut pas convertir plutôt
que de désigner un chantier au hasard.

**Le clic sur une case est arbitré** : pinceau armé → on pose le chantier ; sinon → on ouvre la
fiche. Et tout est branché en JavaScript (`brancherGlisserDeposer`), **pas en `onclick=` dans le
HTML** : les libellés et les noms contiennent apostrophes et accents, et un attribut `onclick`
construit par concaténation finit par se casser en silence (piège déjà vécu le 28/08/2026 sur le
bouton Suivi de l'appli technicien).

Un récapitulatif en bas de la vue liste ce qui reste à reporter dans Teams, par personne.
**Tout a été vérifié à l'écran le 01/09/2026** : glisser-déposer, pinceau au clic, refus du dépôt sur
une case Teams, retrait par la croix, survie au rechargement, et effacement automatique après
rattrapage par Teams.

**Suite naturelle, PAS FAITE et à ne pas faire sans accord explicite de Patrice** : appliquer le
brouillon au fichier Teams en session (Excel COM sait colorier les cellules). Ce serait la première
fois qu'on **écrit** dans un fichier que toute l'équipe a ouvert — risque de conflit et de perte de
travail. À proposer, jamais à décider seul.

**Suite prévue — la grille sera modifiable, mais pas avant que le suivi soit une appli partagée.**
C'est la raison pour laquelle Patrice a demandé cette grille (échange du 01/09/2026) : il veut à
terme que ses collègues et lui modifient le planning dans le suivi, ce qui suppose de sortir les
données du fichier HTML pour les mettre dans un stockage commun (le Worker Cloudflare existe déjà,
il faut lui ajouter une base). **Tant que ce n'est pas fait, rendre la grille modifiable créerait
deux plannings divergents** — celui de Teams où l'équipe travaille, celui du suivi modifié dans un
coin — sur une donnée qui décide d'envoyer des hommes sur des postes. Les 7 personnes concernées
(les 5 ci-dessus + Carine Rambaud à la compta + Guillaume Heras, PDG) peuvent toutes voir le
commercial : **un seul niveau d'accès suffit**, inutile de prévoir des rôles.

## ÉTAPE 1 — le suivi devient une appli partagée (démarrée le 02/09/2026)

**Le but, dit par Patrice** : « une page en ligne où chacun ouvre, saisit, et voit la même chose en
direct. Fini les envois Teams, fini le SEED_VERSION, fini le une-personne-à-la-fois. » C'est la
raison pour laquelle il a demandé la grille Planning. **Le planning est le pilote** : c'est là qu'est
la vraie douleur (plusieurs personnes arrêtent le planning ensemble), il ne contient rien de
commercial, et l'Excel Teams reste tendu dessous comme filet pendant la bascule.

### Les décisions de Patrice (02/09/2026)

1. **Compte Cloudflare au nom de TELSAM**, pas son compte personnel — l'outil de l'entreprise ne
   doit pas dépendre de son compte à lui (le relais Dropbox, lui, est resté sur le compte perso :
   à migrer un jour).
2. **Porte d'entrée = un mot de passe par personne.** Il avait d'abord retenu le code à usage
   unique par mail (Cloudflare Access), puis **il est revenu en arrière le même jour** : Access
   réclame une carte bancaire pour son offre gratuite, et « si un jour nous sommes plus de
   cinquante, je vais être embêté si j'ai laissé mon code de carte bleue ». **Ne pas re-proposer
   Access sans qu'il le demande** — la raison n'est pas technique, elle est contractuelle, et elle
   ne changera pas parce qu'on la réexplique.
   *Ce que ça coûte, à savoir et à assumer* : avec Access il ajoutait ou retirait quelqu'un tout
   seul ; avec les mots de passe, il faut me le demander (recalcul de l'empreinte + push).*

Rappel du cadrage : les **7 personnes** (Brillou, Cazenave, Hamouch, Vidal, Pivot, Rambaud, Heras)
peuvent toutes voir le commercial → **un seul niveau d'accès**, pas de rôles à prévoir.

### L'architecture, et pourquoi celle-là

**Aucun outil à installer sur la machine : il n'y a NI Node, NI npm, NI wrangler** (vérifié le
02/09/2026). Tout choix qui suppose un déploiement en ligne de commande est donc hors jeu — d'où :

| brique | choix | pourquoi |
| --- | --- | --- |
| hébergement | **Cloudflare Pages** relié au dépôt GitHub privé | publie un dépôt **privé** sans plan payant, contrairement à GitHub Pages (cf. « Pourquoi il n'y a pas de vraie serrure »). Un push = un déploiement, rien à lancer à la main. |
| porte | **`functions/_middleware.js`**, devant tout le projet | la page elle-même n'est jamais servie à un inconnu — voir ci-dessous, c'est LA différence avec l'appli technicien. |
| API | **fonction Pages** `functions/api/planning.js` | même dépôt, même déploiement, aucune infrastructure de plus. |
| base | **Cloudflare D1** (SQL), liaison `DB` | écriture **ligne par ligne**. C'est le point clé, voir ci-dessous. |

**Le build ne publie QUE la page** : `mkdir -p public && cp suivi_chantiers_205.html public/index.html`,
sortie `public`. Sans ça Pages mettrait en ligne tout le dépôt (CLAUDE.md, scripts). Les fonctions,
elles, sont lues dans `functions/` à la racine quel que soit le dossier de sortie.

**POURQUOI UNE BASE ET PAS UN FICHIER PARTAGÉ.** Un fichier ne réglerait rien : c'est le défaut de
l'Excel actuel, où deux personnes qui préparent la même semaine s'écrasent mutuellement sans s'en
apercevoir. Ici **une ligne = une personne, un jour** : Ahmed qui place Bilal mardi et Patrice qui
place Didier jeudi ne se marchent jamais dessus. **La version partagée est donc plus sûre que
l'Excel, pas moins** — argument à redonner à Patrice, c'est contre-intuitif.

### Ce qui reste dehors, volontairement

- **Le planning lu dans Teams n'entre PAS dans la base.** Il continue d'être relu chaque matin et
  embarqué dans la page (`PLANNING_RTE`). La base ne contient que ce que l'équipe **décide** dans le
  suivi. C'est ce qui permet de faire tourner les deux en parallèle et de **voir les écarts**, au
  lieu de basculer d'un coup sur une source non éprouvée. On arrêtera l'Excel quand plus personne
  ne l'ouvrira, pas avant.
- **Le reste du suivi** (chantiers, boîtes, heures) fonctionne comme aujourd'hui, depuis `SEED_DATA`.
- **Le fichier doit continuer de marcher seul, sans réseau.** C'est le filet ; ne pas le retirer en
  « simplifiant » quand la version en ligne tournera.

### La porte — `functions/_middleware.js`

**LA SERRURE EST AVANT LA PAGE, PAS DEDANS. C'est le point à ne jamais défaire.** Dans l'appli
technicien le mot de passe est vérifié DANS la page : elle est donc téléchargée d'abord, et son
contenu est lisible dans le code source sans rien taper — acceptable là-bas, puisqu'elle ne contient
aucun secret (la vraie serrure des documents est côté Dropbox). Ici c'est l'inverse : **le suivi
contient le commercial**. Une serrure dans la page ne protégerait donc rien du tout. Le filtre
tourne côté Cloudflare et, sans session valide, ne renvoie que l'écran de mot de passe.

- **Empreintes PBKDF2-SHA-256, sel `telsam-suivi-partage-2026`, 100 000 tours.** Pas 150 000 comme
  l'appli technicien : **Cloudflare plafonne à 100 000**, au-delà l'appel échoue (piège déjà payé
  sur le relais Dropbox). Même `normPw` que l'appli — minuscules, tout ce qui n'est ni lettre ni
  chiffre retiré — donc les tirets sont facultatifs.
- **Trois mots au lieu de deux** : ces mots de passe se tapent sur un clavier d'ordinateur, pas sur
  un téléphone en haut d'un pylône, et ils gardent une adresse publique devant du commercial.
- **Session par cookie signé** (`mail|expiration|HMAC`), 7 jours, `HttpOnly` + `Secure` +
  `SameSite=Strict`. Le secret de signature vient de `SESSION_SECRET` (variable d'environnement
  Cloudflare, **jamais dans le dépôt**). **Absent ⇒ 503, rien n'est servi** : un réglage manquant
  ne doit pas ouvrir la porte.
- **Comparaisons à temps constant** (`memeChaine`) : une comparaison qui s'arrête au premier écart
  laisse deviner une empreinte octet par octet.
- **Compteur d'échecs par IP** (table `tentative`, 10 échecs / 10 min). Sept mots de passe derrière
  une adresse publique : sans compteur, rien n'empêche l'essai en masse.

**INCIDENT DU 02/09/2026 — NE JAMAIS CITER EN EXEMPLE UNE COMBINAISON QUE LE TIRAGE PEUT PRODUIRE.**
J'avais pris `ancre-jade-moulin` comme exemple dans le guide et comme mot de passe d'essai. Le tirage
des 7 est **tombé exactement dessus** : un mot de passe réel, public le jour de sa création. Ce n'est
pas de la malchance rare — 26×20×20 = 10 400 combinaisons, 7 tirages, plusieurs exemples cités : de
l'ordre d'une chance sur 1 500 par exemple. Les 7 ont été regénérés, le générateur refuse maintenant
les combinaisons interdites et les doublons, et le guide cite `velo-mangue-tiroir`, dont les trois
mots sont **absents des listes de tirage**. Règle générale : un exemple doit être **impossible** à
tirer, pas seulement improbable — et il faut vérifier après coup, ce qui est ce qui a rattrapé le cas.

**Les mots de passe en clair vivent dans `Mots_de_passe_Suivi_Partage_TELSAM.txt` sur le Bureau,
hors dépôt** — en `.txt` et non dans le document Word des mots de passe, pour ne pas risquer
d'abîmer un fichier qui contient déjà tout le reste (les pièges Word sont documentés plus bas).
Proposé à Patrice de l'y intégrer s'il préfère. Pour en changer un : recalculer avec
`scratchpad/mdp-suivi.ps1` (même sel, mêmes tours), remplacer dans `PERSONNES`, pousser.

**PIÈGE DE MÉTHODE — le navigateur ne peut PAS éprouver les cookies.** `Cookie` est un en-tête de
requête interdit et `Set-Cookie` est filtré des `Headers` d'une réponse : un essai qui passe par de
vrais objets `Request`/`Response` **ne prouve rien** et paraît pourtant réussir. Vécu le 02/09/2026 :
un premier jeu d'essais affichait « cookie falsifié refusé ✓ » alors que le cookie n'était jamais
arrivé — il refusait tout, y compris le cas valide. **Utiliser un objet de requête factice** (le code
n'appelle que `headers.get`, `method`, `url`, `formData`). C'est ainsi qu'ont été vérifiés, le
02/09/2026 : cookie valide accepté et identité posée, mail falsifié refusé, expiration repoussée
refusée, session expirée refusée, secret différent refusé, API sans session refusée en JSON,
blocage après 10 échecs, et le suivi jamais servi dans aucun cas de refus. Les attributs du cookie
(`HttpOnly`/`Secure`/`SameSite`), eux, sont contrôlés **dans le source** — le navigateur ne les
laisse pas relire.

### L'API — deux points à ne pas défaire

`functions/api/planning.js`. `GET ?du=&au=` lit une semaine, `POST {jour, personne, chantier, remplace}`
écrit une cellule (`chantier: null` retire).

1. **ELLE ÉCHOUE FERMÉE, ET NE CROIT AUCUN EN-TÊTE.** L'identité est lue dans
   `context.data.utilisateur`, posé par le filtre. **Absente ⇒ 503, on ne sert RIEN** : ça ne veut
   pas dire « visiteur anonyme » mais « le filtre n'est pas devant », donc une configuration qui
   exposerait du commercial.
   *Une première version lisait l'en-tête `Cf-Access-Authenticated-User-Email` parce que Cloudflare
   Access était prévu devant. Après le retour aux mots de passe, cet en-tête n'existe plus — et un
   en-tête est envoyé par le client, donc falsifiable : s'y fier laisserait n'importe qui écrire au
   nom de n'importe qui. **Ne jamais faire porter l'identité par un en-tête ici.***
2. **Le libellé du chantier est stocké EN TOUTES LETTRES, jamais son index** — même piège que le
   brouillon : `libelles` est reconstruite à chaque relecture et les index se décalent (118 puis
   172), une ligne qui retiendrait un index désignerait un autre chantier du jour au lendemain.

Refus volontaires : samedi/dimanche (le piège du week-end gris — une affectation acceptée là serait
invisible dans la grille, donc incorrigible) et libellé vide. Une couleur mal formée n'est pas un
refus, elle est simplement ignorée : mieux vaut une affectation enregistrée sans couleur qu'une
saisie perdue.

**Testée le 02/09/2026, 8 cas, avant tout déploiement** (fausse base D1 + vraies `Request`, jouées
dans le navigateur puisqu'il n'y a pas de Node) : refus sans Access, refus sans base, lecture bornée
à la semaine demandée, refus du samedi, pose avec libellé normalisé, journal avant/après, retrait,
couleur invalide ignorée, libellé vide refusé.

### OÙ ON EN EST — état au 02/09/2026, à relire avant de reprendre

**LE SUIVI EST EN LIGNE ET PROTÉGÉ : `https://suivi-telsam.pages.dev`.** Patrice s'y est connecté
avec son mot de passe et voit son suivi. **Ne jamais lui envoyer cette adresse en lien cliquable —
un lien fait planter sa machine et lui coupe l'accès à Claude** (posé fermement le 02/09/2026,
cf. [[feedback_manip_amener_au_bon_endroit]]).

Installé par Patrice, et **vérifié par moi sur le site réel** :

| brique | état |
| --- | --- |
| Compte Cloudflare TELSAM | créé sur `Suivichantier@telsam.com` |
| Projet Pages `suivi-telsam` | relié à `Pivot-telsam/suivi-chantiers` (autorisation GitHub limitée à ce seul dépôt), build `mkdir -p public && cp suivi_chantiers_205.html public/index.html`, sortie `public` |
| `SESSION_SECRET` | posé en variable de production |
| La porte | **vérifiée en direct** : écran de mot de passe, mauvais mot de passe refusé, 9 chemins testés (`/`, `/CLAUDE.md`, `/PROGRESS.md`, `/scripts/…`, `/partage/schema.sql`, `/functions/_middleware.js`, `/api/planning`, `/index.html`, `/suivi_chantiers_205.html`) tous bloqués, aucune fuite |
| Base D1 `suivi-telsam` | créée, schéma exécuté (4 tables) |
| Liaison `DB` → projet Pages | **faite et prouvée en direct** (voir ci-dessous) |

**L'INSTALLATION DE PATRICE EST TERMINÉE (02/09/2026).** Il n'a plus rien à faire.

**Comment la liaison `DB` a été prouvée sans mot de passe** — méthode à réutiliser, elle vérifie
trois choses d'un coup : envoyer **12 mauvais mots de passe** à `/entrer`. Les 10 premiers
répondent 401, le **11ᵉ répond 429 « Trop de tentatives »**. Or ce blocage n'est possible que si le
compteur d'échecs a pu **écrire puis relire** la table `tentative` : liaison en place, table
présente et accessible, protection contre l'essai en force fonctionnelle. Un simple appel à
`/api/planning` n'aurait rien prouvé (la porte le refuse avant d'atteindre la base).
**Effet de bord à annoncer** : ce test bloque l'adresse IP **de Patrice** pendant 10 minutes,
puisque le navigateur d'essai sort par sa connexion. Le lui dire avant, ou après, mais le dire.
Pour débloquer tout de suite : `DELETE FROM tentative;` dans la console D1.

*Petite friction d'interface, pour la prochaine fois : dans la liste **Workers & Pages**, la ligne
contient DEUX textes superposés — le nom du projet en gras noir (qui ouvre le projet) et l'adresse
`…pages.dev` en gris juste dessous (qui ouvre le site). Patrice cliquait la seconde. Et deux objets
portent le nom `suivi-telsam` : la base D1 et le projet Pages.*

### La grille est branchée sur la base (02/09/2026) — DEUX MODES

**`planningPartage`** dit dans quel mode on tourne, et **on ne le devine jamais** : on interroge
`/api/planning` une fois au démarrage et c'est sa réponse qui tranche.

| mode | quand | ce qui se passe |
| --- | --- | --- |
| **partagé** | en ligne, la base répond | l'écriture part dans la base, les 7 la voient, le suivi garde qui/quand |
| **fichier seul** | fichier ouvert en direct, pas de réseau, base en panne | brouillon local en IndexedDB, exactement comme avant |

**LE REPLI N'EST PAS UN RESTE DU PASSÉ, C'EST LE FILET.** Le suivi doit continuer de s'ouvrir quand
Cloudflare est en panne — c'est un outil qui sert à décider d'envoyer des hommes sur des postes. Ne
pas le retirer en « simplifiant » quand la version en ligne sera devenue l'habitude.

**IL NE FAUT JAMAIS LAISSER CROIRE QU'ON PARTAGE QUAND ON NE PARTAGE PAS.** C'est la règle qui a
dicté le reste. Quelqu'un qui prépare une semaine en croyant que l'équipe la voit, c'est un
technicien que personne ne prévient. D'où :
- un **bandeau en haut de la vue**, vert « Suivi partagé — tes six collègues le voient » ou ambre
  « Fichier seul — personne d'autre ne le voit ». Il se lit avant de toucher à la grille ;
- une base qui répond mal **le dit** (`planningErreurBase` affiché dans le bandeau ambre) au lieu de
  basculer en silence sur le local ;
- une écriture refusée **prévient et ne change rien** — vérifié : la base reste intacte et la case ne
  s'affiche pas comme si c'était passé.

**Trois choix à ne pas défaire :**
1. **Pas de bouton « Vider » en partagé.** Il effacerait d'un clic les décisions des six autres, sans
   qu'elles l'aient demandé. On retire une case à la fois, comme on l'a posée.
2. **Aucune suppression de ligne partagée déclenchée par une lecture.** Quand Teams rattrape une
   décision, on la marque `dejaDansTeams` et on cesse de l'afficher — on ne la supprime pas. Effacer
   des lignes communes pendant que sept postes lisent ferait disparaître le travail de quelqu'un.
   En mode local, en revanche, l'effacement reste (c'est le comportement d'avant).
3. **Pas de minuterie de rafraîchissement en fond.** On relit la base à l'ouverture de la vue et au
   changement de semaine — les moments où l'on s'attend à voir le travail des collègues. Toujours
   **dessiner d'abord, rafraîchir ensuite** : l'écran n'attend jamais le réseau.

**Ce que la case affiche en partagé** : le chantier, puis « par <prénom> » (l'auteur vient de la
base, jamais du navigateur) et « au lieu de <chantier> » quand la décision remplace ce que dit
Teams. Trait plein (`plPartage`) au lieu des hachures du brouillon : ce n'est plus un brouillon.
Le mot « brouillon » ne subsiste qu'en mode fichier seul.

**PIÈGES DE BANC D'ESSAI — deux fois le même genre d'erreur le 02/09/2026, à connaître avant
d'écrire un test ici :**
- **Le navigateur ne peut pas éprouver les cookies** (déjà noté plus haut) ni relire `Set-Cookie`.
- **Un faux `fetch` se contamine d'un essai à l'autre.** Un premier script avait remplacé
  `window.fetch` puis échoué avant d'exposer sa poignée de restauration ; le script suivant a
  « sauvegardé le vrai fetch » — qui était déjà le faux. Résultat : après restauration, la page
  croyait être en mode partagé avec les données du test précédent. **Recharger la page entre deux
  essais qui remplacent une fonction globale**, et ne jamais conclure sur un état hérité.
- **Une découpe de fichier laisse des accolades orphelines.** Le splice de ce jour a dupliqué le `}`
  final d'`armerPinceau` → tout le script cassait, et `storeSet is not defined` était le seul
  symptôme visible. Contrôle qui tranche : `new Function(code)` sur **chaque** bloc `<script>` de la
  page. Et méfiance avec la console du navigateur : **elle garde les erreurs des chargements
  précédents**, j'ai cru deux fois à une erreur déjà corrigée.

**Vérifié le 02/09/2026, les deux modes** : mode fichier seul inchangé (bandeau ambre, étiquette
« brouillon », bouton Vider, persistance IndexedDB) ; mode partagé (bandeau vert, identité lue dans
la base, POST au bon format avec `remplace`, étiquette « par Christian · au lieu de … », trait
plein, pas de bouton Vider, retrait qui envoie `chantier: null`) ; et les deux cas de panne
(écriture refusée → alerte + base intacte + rien affiché ; base muette au chargement → repli annoncé
dans le bandeau).

**À prévoir, non bloquant :**
- une **sauvegarde quotidienne** de la base vers Dropbox (le relais sait déjà y écrire) — aujourd'hui
  chaque copie Teams du fichier est une sauvegarde, une base unique n'en a aucune ;
- **inviter un deuxième administrateur** au compte Cloudflare (Christian ou Pierre) : c'est le vrai
  remède au « l'outil ne doit pas dépendre d'une seule personne », mieux qu'une adresse générique.

**Piège rencontré à l'inscription, à ne pas réoublier** : `Suivichantier@telsam.com` est un **groupe
Microsoft 365**, et un groupe M365 **refuse par défaut les expéditeurs externes** — le mail de
vérification Cloudflare n'arrivait donc jamais, alors que les mails internes arrivaient bien.
Diagnostiqué par l'erreur Graph `ErrorGroupIsUsedInNonGroupURI` en tentant de lire la boîte, pas par
supposition. À savoir pour tout futur service externe branché sur une adresse générique de TELSAM.

**Piège de l'interface Cloudflare** : deux systèmes coexistent. L'écran « Make something new »
(nouveau système Workers) **ignore silencieusement le dossier `functions`** — il aurait mis le suivi
en ligne **sans la serrure**. Il faut passer par le lien « Continue to Pages » en bas de cet écran.
C'est écrit dans `partage/INSTALLATION.md`.

### Les trois améliorations demandées par Patrice le 03/09/2026

**1. Des bulles, et NOTRE palette — pas celle de l'Excel.** Trois essais, et c'est le troisième qui
tient. D'abord les couleurs brutes du fichier : « trop flashy » (l'Excel colorie au hasard parmi
jaune pur, vert pur, orange pur). Puis ces mêmes couleurs adoucies en pastel : « quand même un peu
trop pastel ». Ce qu'il voulait, dit avec ses mots : **« celles du Gantt — sobres, mais bien
colorées »**.

**La phrase qui a débloqué le sujet** : « tu n'es pas obligé de prendre les mêmes couleurs que nous
mettons sur notre planning Teams, l'essentiel est que les couleurs correspondent aux chantiers
semaine par semaine comme nous fonctionnons sur Teams. » Autrement dit **la couleur de l'Excel n'a
aucune valeur en soi** : sa seule fonction est de dire « ces cases-là, c'est le même chantier ». On
peut donc la remplacer, à condition de rester stable sur la période regardée. Les deux premiers
essais partaient de la couleur du fichier ; c'était la contrainte inutile.

`PALETTE_PLANNING` : 16 teintes dans le registre de `.ganttBar` (fond franc, texte blanc, ombre
discrète). Deux couleurs sont **volontairement absentes** : le rouge vif et l'ambre des alertes,
parce que dans cette vue rouge pâle = absence et ambre = avertissement ; les réutiliser pour un
chantier brouillerait une lecture qui doit rester immédiate.

**LA COULEUR EST FIXE PAR CHANTIER, ET NE CHANGE PLUS D'UNE SEMAINE SUR L'AUTRE.**
Deuxième correction du 03/09/2026, sur remarque de Patrice : « ce serait quand même pas mal de
garder les mêmes d'une semaine sur l'autre. Ce n'est pas grave de les changer quand on réintervient
sur un chantier un ou deux mois plus tard, mais ça prête à confusion de les changer d'une semaine
sur l'autre. »
- La version précédente attribuait la palette **par affichage**, dans l'ordre d'apparition :
  changer de semaine redistribuait tout, et un chantier passait du bleu au vert sans avoir bougé.
- Désormais la teinte vient du **rang du chantier dans la liste de toutes les fiches, triée par
  numéro** (`ordreChantiers`) : déterministe, identique quelle que soit la semaine, la machine ou le
  moment. **Vérifié sur quatre semaines d'affilée : aucun chantier ne change de couleur.**
- `ordreChantiers` se recalcule quand le nombre de fiches change — sans ça une fiche créée en cours
  de session n'aurait pas de rang et retomberait sur le libellé.
- Pour une ligne de planning qu'aucune fiche ne rattache, le rang est dérivé du **libellé**
  (`_rangDuTexte`), donc stable aussi.

**CE QUE CE CHOIX COÛTE, ET POURQUOI C'EST LE BON.** Avec 16 teintes et 71 fiches, deux chantiers
partagent forcément une couleur et peuvent tomber dans la même vue. Mais le nom est écrit dans la
bulle : une couleur partagée fait hésiter une seconde, une couleur qui change fait croire à un
autre chantier. Patrice a tranché dans ce sens de lui-même. **Ne pas « améliorer » en revenant à une
attribution par vue** — c'est précisément ce qu'on vient de retirer.

*Défaut intermédiaire, réglé au passage : le Poste de Portet ressortait magenta en S36 et violet en
S37, parce que le planning le nomme autrement d'une semaine à l'autre. La clé étant maintenant la
fiche, le problème ne peut plus se poser. On ne retombe sur le libellé que pour les lignes
qu'aucune fiche ne rattache — et **il ne faut PAS les rapprocher par ressemblance de texte** pour
« finir le travail » : c'est la règle qui interdit d'envoyer un jour vers le mauvais lot. Ces
lignes-là se règlent en annotant le numéro dans le planning, ou en créant la fiche et son entrée
`TECH_RANGES` (ce qui a suffi pour Bollène-Plantades : la case est passée du libellé brut à
« 26-071 — Bollène - Plantades », reliée et cliquable, sans toucher au planning).*

La case redevient un fond neutre et c'est la **bulle** (`.plBulle`) qui porte la couleur — d'où
aussi une case vide qui se distingue enfin d'une case remplie.

**2. Deux semaines côte à côte, et la réserve enfin visible.** Dix colonnes (L-V, L-V), un en-tête
par semaine et un séparateur franc (`.plSep`) : sans lui on lit le vendredi de la semaine 1 et le
lundi de la semaine 2 comme deux jours consécutifs. Les boutons disent « Semaine précédente /
suivante » et décalent bien d'**une** semaine — le libellé doit décrire ce que fait le bouton.
La réserve, elle, est devenue une **barre collée en bas** (`position:sticky; bottom:0`). Le défaut
signalé était réel : avec la grille au-dessus, elle tombait sous la ligne de flottaison, donc on ne
pouvait pas glisser un chantier sans faire défiler et perdre de vue la case visée. **Vérifié à
cinq hauteurs de défilement** : la barre reste en bas de la fenêtre partout. Les pastilles sont
bornées en hauteur et défilent (`max-height:104px`) — une barre qui grandit sans fin mangerait
l'écran qu'elle est censée dégager.

**3. Une recherche parmi TOUS les chantiers.** « Aller chercher un chantier à affecter parmi nos
chantiers. » La réserve ne montre que ce que l'Excel propose ces semaines-là ; le champ cherche dans
les 70 fiches. **Capacité nouvelle** : placer quelqu'un sur un chantier pas encore posé au planning.

**ELLE NE CACHE RIEN — corrigé le 03/09/2026.** Patrice : « quand je tape c a s, je n'ai pas Poste de
Casteljaloux ». La recherche n'était pas cassée : **j'excluais les chantiers marqués `termine`**, et
la fiche 26-044 Casteljaloux l'est — alors que le planning RTE porte des travaux dessus en septembre.
Un chantier « terminé » peut donc reprendre, et c'est justement là qu'on a besoin de le chercher.
**Un outil de recherche qui omet en silence ce qu'on lui demande passe pour cassé, et il l'est en
pratique.** On rend donc tout, en marquant « ⚠ fiche marquée terminée » : c'est à Patrice de savoir,
pas à moi de choisir pour lui.

**Accents retirés des deux côtés**, via `normSearch()` — la fonction existait déjà dans le fichier
(barre de recherche des fiches), inutile d'en écrire une deuxième. Sans ça « severac » ne trouve pas
« Séverac » ni « vallee » « Vallée », et un chantier introuvable fait croire qu'il n'existe pas.
Vérifié sur les trois cas.

Le libellé enregistré est `26-039 — Lisieux - Vallée 1`, donc `chantierDuLibelle` le rattache à sa
fiche par le numéro : la bulle reste cliquable et prend la couleur de cette fiche. **Vérifié.**
- **La recherche ne redessine QUE la liste de pastilles**, jamais la vue : sinon le champ perdrait
  le focus à chaque lettre. Vérifié (`document.activeElement === champ` après saisie).

**Les pastilles ne portent qu'un INDEX dans leur HTML** (`data-chip="3"`, résolu via
`planningChips`). `escapeHtml` de ce fichier n'échappe **ni** les guillemets **ni** les apostrophes,
et les libellés du planning en sont truffés (« Racco d'un TI, 2 équipes ») : mettre un libellé dans
un attribut casserait l'attribut en silence — même piège que le bouton Suivi de l'appli le 28/08.

**Impression** : la règle « les couleurs sont l'information » vaut maintenant pour `.plBulle` et
`.plChip`, à ajouter au bloc `@media print` — l'oublier reproduirait exactement le défaut de grille
blanche que cette règle corrigeait. La barre passe en `position:static` à l'impression, sinon elle
se superpose au tableau sur le papier.

**Limite connue, et c'est la vérité du fichier** : sur une semaine où Patrice n'a pas encore écrit
le numéro dans la ligne du planning, la bulle affiche le libellé brut en capitales
(« POSTE DE PORT… ») au lieu de « 26-051 — Poste de Portet ». Ne pas « embellir » en retitrant le
texte : `LA 63kV` deviendrait `La 63kv`, et on abîmerait une information technique pour une
question d'allure. C'est l'annotation du planning qui résout ça, pas le code.



### Les couleurs du planning, deuxième passe (03/09/2026) — et l'erreur de raisonnement corrigée

Patrice : **« les couleurs de bollene et fleyria sont memes d'une semaine sur l'autre. il faut le
modifier. »**

**Ce qui était écrit ici avant, et qui était faux.** Le commentaire du code disait : « avec 16
teintes et 71 fiches, deux chantiers finissent forcément par partager une couleur, et ce n'est pas
grave puisque le nom est écrit dans la bulle ». Confortable, et démenti par la mesure. Les chiffres,
relevés sur l'année entière :

| | étiquettes simultanées au maximum |
|---|---|
| la grille seule (les cases) | **14** |
| la grille + la réserve du bas | 25 |

Pour 16 teintes, **la grille avait la place**. Ce n'était pas la palette qui manquait, c'était
`rang % 16` qui la gâchait, en donnant la même case à deux rangs distants de 16 :
Bollène rang 70, Fleyriat rang 54, `70 % 16 = 54 % 16 = 6`.

**Ce qui est fait maintenant.** Chaque étiquette garde sa couleur **préférée** (`rang % 16`, celle
qu'elle a toujours eue), sauf si une étiquette avec qui elle partage une quinzaine l'a déjà prise —
auquel cas elle glisse sur la première teinte libre. Le calcul porte sur **toute l'année d'un
coup**, jamais sur la semaine affichée : c'est ce qui garantit qu'une couleur ne bouge pas quand on
tourne les semaines. Prix payé, mesuré : 24 chantiers sur 58 changent de teinte **une fois**.

**TROIS DÉCISIONS QUI ONT CHACUNE COÛTÉ UN ESSAI RATÉ, ne pas les défaire :**

1. **L'ordre de passage est la date de PREMIÈRE APPARITION, pas le numéro.** Premier essai par
   numéro : trois collisions subsistaient. Le graphe a dit pourquoi — « POSTE DE CANTEGRIT » avait
   35 voisins et « consignation 225kV CURBANS » 71, parce qu'un chantier qui revient en février
   puis en mars puis en mai voisine avec tout ce qui passe entre-temps. Or ces voisinages ne sont
   pas simultanés : ce qui compte n'est pas le nombre total de voisins mais le nombre de chantiers
   présents **en même temps**. En traitant les étiquettes dans l'ordre de leur première semaine, on
   ne se heurte qu'à celles déjà commencées. C'est le résultat classique sur les graphes
   d'intervalles, et il vaut ici parce qu'un chantier occupe des semaines, pas des points isolés.
2. **La grille passe avant la réserve.** Deuxième essai : deux collisions restaient, et la mesure a
   montré que 25 étiquettes peuvent coexister pour 24 teintes — aucune palette de couleurs franches
   et distinctes ne couvre ça. On tranche donc par l'usage : dans la grille, la couleur EST
   l'information (« qui est où »), alors qu'une pastille de réserve est un bouton qu'on lit et
   qu'on fait glisser, avec son nom écrit dessus. Les étiquettes de la grille sont servies en
   premier et gardent la **garantie** ; les pastilles prennent ce qui reste. Résultat : **0
   collision dans la grille**, 4 sur les pastilles, assumées.
3. **Huit teintes de SECOURS** (`PALETTE_SECOURS_PLANNING`), jamais choisies spontanément, qui ne
   servent que quand les 16 premières sont prises. Elles évitent qu'un chantier à cheval sur
   plusieurs mois retombe sur une couleur déjà prise à côté de lui.

**Le graphe est gardé dans `_grapheTeintes`** après chaque construction (voisins, semaines,
première apparition, nombre de teintes bloquées au maximum, étiquettes à court de teintes). Ce n'est
pas du décor : les trois essais ci-dessus ont été tranchés en le lisant, pas en devinant. La
première tentative de diagnostic « à vue » avait produit une hypothèse fausse.

**Deux étiquettes = une seule couleur quand c'est le même chantier.** La clé est `f:<id de fiche>`
dès qu'une fiche est reconnue, et `t:<libellé normalisé>` sinon. C'est ce qui fait que les cinq
lignes « FLEYRIAT » du planning ne donnent qu'une teinte. La résolution utilise
`chantierDuLibelle(t) || chantierParPresence(nom, iso)`, **la même que la vue** : sans le repli par
présence, Bollène — dont la ligne de planning ne porte pas de numéro — n'aurait pas été reconnu
comme la fiche 26-071, et la collision signalée par Patrice serait restée invisible au calcul.

### L'avertissement « couleur ambiguë » ne crie plus au loup (03/09/2026)

Patrice : **« concernant fleyriat c'est le même chantier il n'y a pas besoin de mettre un
avertissement. j'ai rajouté le numéro d'indice dans teams. »**

Un même chantier occupe souvent **plusieurs lignes** du planning (une par phase : « contrôle
touret », « 1er phase de travaux », « oppc… »), toutes de la même couleur puisque c'est le même
chantier. L'avertissement « cette couleur sert à plusieurs lignes » se déclenchait alors pour rien —
20 cases sur les seules lignes Fleyriat.

**La règle, dans `planning-rte.ps1` : c'est le NUMÉRO écrit dans la ligne qui juge, et lui seul.**

- un seul numéro distinct parmi les lignes, les autres muettes ⇒ même chantier, on retient la ligne
  qui porte le numéro, **pas d'avertissement** ;
- deux numéros différents (ex. les deux lots `26-036-1` / `26-036-2`) ⇒ vraie ambiguïté, on avertit ;
- aucun numéro ⇒ on ne peut pas savoir, on avertit.

**Comparer les libellés entre eux serait une devinette** : « AUDIT LIVIERE - MAS NOU » et
« LANGEAC - PRATCLAUX » partagent une couleur et n'ont rien à voir. Et cette règle récompense le bon
geste : **annoter le planning fait disparaître l'avertissement, sans toucher au code.**

Effet mesuré : cases ambiguës 73 → 53, et **présence publiée 12 → 14 chantiers**, puis **16** le soir même quand Patrice a écrit les trois numéros manquants dans Teams — Fleyriat
alimente désormais le Gantt et la couverture PGO, ce qu'il ne faisait pas.

**Restent 4 groupes signalés, tous sans aucun numéro dans leurs lignes** — à dire à Patrice, le
remède est entre ses mains :
- CERDAGNE 2 (deux groupes, 43 cases) ;
- COLAYRAC-GUPIE (6 cases) ;
- AUDIT LIVIERE - MAS NOU / LANGEAC - PRATCLAUX (4 cases) — celui-là est un **vrai** avertissement,
  ce sont deux chantiers différents.

### Les notes libres dans les cases (demandé par Patrice le 03/09/2026)

Sa phrase : **« j'aimerais aussi que nous puissions rajouter manuellement des infos dans les cases
des encadrants (comme rajouter une ICP où prevoir une visite, etc..) »**.

**LA DÉCISION QUI COMPTE : une note n'est PAS un chantier, et ne doit jamais en devenir un.** Le
réflexe économique aurait été de la poser comme une affectation au libellé libre — rien à changer
dans la base, rien dans l'API. Ç'aurait été faux, et faux silencieusement :

- elle serait entrée dans l'effectif affiché sur les pastilles (« 3 personnes » au lieu de 2) ;
- elle serait apparue dans le récap « affectations à reporter dans Teams » ;
- le jour où le planning alimentera la **couverture PGO**, « prévoir une visite » compterait comme
  une présence d'homme sur un poste — c'est-à-dire une couverture documentaire fausse sur une
  donnée de sécurité.

D'où un champ **séparé à tous les étages** : colonne `note` dans `affectation`, clé `n` dans le
brouillon, `.plBulleNote` au rendu.

**Ce qui a été touché, et ce qu'il faut savoir avant de le retoucher :**

| Endroit | Ce qui change |
|---|---|
| `partage/schema.sql` | colonne `note TEXT`, `schema_version` = 3, historique des versions en pied |
| `functions/api/planning.js` | `assurerColonneNote()`, lecture de `note` → `n`, branche note dans `ecrire()` |
| `suivi_chantiers_205.html` | `poserNote()`, `demanderNote()`, `noteDe()`, `NOTE_MAX`, `.plBulleNote`, `.plNoteBtn` |

**La colonne se pose toute seule, et c'est volontaire.** `assurerColonneNote()` lance
`ALTER TABLE affectation ADD COLUMN note TEXT` au premier appel de l'isolat et absorbe **la seule**
erreur « duplicate column ».

**Elle ne fait JAMAIS échouer la requête qui l'appelle.** Une première version relançait l'erreur
dès qu'elle n'était pas « duplicate column » — défaut grave, corrigé avant la mise en ligne : le
jour où cet ordre échoue pour une raison passagère alors que la colonne existe déjà (posée par un
autre isolat), on aurait coupé la **lecture** du planning, donc toute la vue, pour les sept
personnes, à cause d'une migration qui n'avait plus rien à faire. **Une étape de rattrapage ne doit
pas pouvoir casser ce qui marchait sans elle.** On laisse la vraie requête trancher : si la colonne
manque pour de bon, c'est elle qui échoue, et le message reprend celui de la migration
(`_derniereErreurNote`) plus la ligne SQL de secours — un « no such column: note » nu enverrait
chercher au mauvais endroit. Le drapeau `_colonneNotePosee` ne se pose qu'en cas de succès ou de
doublon (une base momentanément indisponible est retentée au coup suivant, pas condamnée pour la
durée de vie de l'isolat). Raison du choix : la base est **en service**, et faire
exécuter une ligne SQL à Patrice dans la console Cloudflare voulait dire une manip de plus sur une
donnée vivante, un ordre de passage à respecter entre le push et la manip, et un déploiement en
panne tant qu'elle n'est pas faite. Ne pas supprimer cette fonction en croyant nettoyer : une base
restaurée depuis une sauvegarde d'avant le 03/09/2026 en aurait besoin.

**LES QUATRE PIÈGES, tous testés, tous avec leur contre-exemple :**

1. **`note: ''` ≠ `note` absente.** On distingue sur `hasOwnProperty`, pas sur la valeur. `''` veut
   dire « efface la note », absente veut dire « n'y touche pas ». Confondre les deux effacerait la
   note d'un encadrant chaque fois qu'un collègue déplace un chantier sur la même case.
2. **Retirer le chantier ne supprime plus forcément la ligne.** Si une note y reste, la ligne
   survit avec `chantier_libelle = NULL` (`UPDATE`, pas `DELETE`). La réponse renvoie
   `noteConservee`.
3. **Effacer la note d'une case qui ne portait QUE ça supprime la ligne.** Laisser une ligne
   entièrement vide donnerait une case qui a l'air libre mais qui bloque le prochain `INSERT` sur
   sa clé primaire.
4. **Toutes les boucles du brouillon supposaient `b.t` présent.** Elles ont chacune reçu un
   `|| !b.t` : teintes, `effectif`/`prepares`, `nbDecisions`, récap, et `nettoyerBrouillonPlanning`.
   Sans ça, une case ne portant qu'une note comptait comme une affectation et cherchait une couleur
   pour un libellé `undefined`.

**Le geste, côté écran** : survoler la case fait apparaître un `+` en bas à droite (un `✎` si une
note existe déjà). `prompt()` est gardé volontairement — c'est le seul champ de saisie identique
dans les deux modes, y compris sur le fichier ouvert en double-clic sans réseau, et il ne demande
aucune fenêtre à construire. Le bouton fait `stopPropagation` : sans lui, le clic remonte à la case
et, avec un pinceau armé, on peindrait un chantier au moment même où on voulait écrire une note.
Le crayon reste invisible jusqu'au survol — 190 cases (19 personnes × 10 jours) portant chacune un
bouton visible annuleraient la lisibilité gagnée le matin même.

**Plafond à 300 caractères, côté serveur ET côté page.** Le côté page seul ne suffit pas : l'API
est joignable directement.

### Comment tester ces deux dépôts sans Node ni navigateur interactif

Il n'y a **ni Node ni npm** sur la machine, et l'accès `file://` du navigateur intégré est accordé
fichier par fichier — donc pas toujours disponible. Chrome en mode `--headless` remplace les deux :

```bash
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless=new --disable-gpu \
  --allow-file-access-from-files --virtual-time-budget=20000 --dump-dom "file:///…/test.html"
```

Les harnais vivent dans `suivi-chantiers/partage/` — donc versionnés, et **jamais publiés** : le
build Cloudflare ne recopie que `suivi_chantiers_205.html` vers `public/index.html`, tout le reste
du dépôt reste hors ligne. Ils portent des chemins `file:///C:/Users/patrice.pivot/...` en dur : à
adapter si le dossier de travail change.

- **`verif-syntaxe.html`** — lit le HTML cible, découpe chaque bloc `<script>` et le passe à
  `new Function` (qui compile sans exécuter). **Piège** : `new Function` compile un *corps de
  fonction*, donc il refuse un `await` de premier niveau et signale « await is only valid in async
  functions ». Sur un fichier qui compilait la veille, ce message veut presque toujours dire qu'un
  **en-tête `async function` a été mangé** par une insertion — c'est exactement ce qui est arrivé le
  03/09/2026 à `poserBrouillon`.
- **`localiser.html`** — injecte le bloc dans un vrai `<script>` et attrape l'événement `error` de
  la fenêtre, qui donne **la ligne et la colonne** ; `new Function` ne sait dire que la nature de
  l'erreur, pas où elle est. C'est l'outil qui a trouvé l'en-tête manquant en une passe.
- **`test-notes.html` / `test-api-notes.html`** — tests fonctionnels. Le premier concatène la vraie
  page et un bloc de test (22 contrôles) ; le second charge `planning.js` en texte, retire les
  `export`, le compile avec `new Function` et remplace `env.DB` par un **mouchard** qui note les
  ordres SQL (36 contrôles). On ne teste **jamais** contre la vraie base D1 : sept personnes ont
  leurs décisions dedans.

**Chaque contrôle doit avoir son contre-exemple.** Les scénarios C et E du test d'API existent
uniquement pour cela : C vérifie qu'on NE supprime PAS une ligne qui porte un chantier, E qu'on la
supprime bien quand elle n'en porte pas. Un test qui ne peut pas échouer ne prouve rien.

### Le cas Rion des Landes S38 — quand la couleur dit le contraire de la vérité (03/09/2026)

Patrice : « je te confirme que nos techniciens sont bien sur Rion des landes en s38. je vois que mon
collègue n'a pas mis exactement la même couleur, c'est peut-être pour ça que tu ne l'as pas vu ? »

Son hypothèse était juste, mais le fond est **plus grave que « pas vu »**. Couleurs mesurées sur les
17 et 18/09 (colonnes 261-262 du fichier Teams) :

| Ligne | Couleur | `ColorIndex` |
|---|---|---|
| Benjamin SOUPA (L15) et Sid Ahmed BENZAMERA (L16) | `#9bc2e6` | 37 |
| ligne-projet « POSTE DE RION DES LANDES : TX AMIANTE SOUS SS4 » (L27) | `#bdd7ee` | 24 |
| ligne-projet « Racco et recette … Poste de Cantegrit » (L34) | **`#9bc2e6`** | **37** |

Les cases des deux techniciens correspondent **exactement** à la ligne Cantegrit. La lecture ne les
avait donc pas « oubliés » : elle les avait mis sur **le mauvais chantier**.

**CONSÉQUENCE POUR LE CODE : ne pas ajouter de tolérance sur la couleur.** C'était la correction
évidente, et elle n'aurait rien réglé ici — une correspondance *exacte* avec une autre ligne existe
déjà, donc tout algorithme de « couleur la plus proche » choisirait Cantegrit avant Rion. Une
tolérance rendrait en plus ambigus des cas aujourd'hui nets. Le vrai correctif est **humain** : la
case doit être recoloriée dans Teams avec `#bdd7ee`.

**Sur le moment, l'arbitrage a vécu dans `TECH_RANGES`** (`c_poste_de_rion_des_la` → Benjamin
SOUPA et Sid Ahmed BENZAMERA, 17 et 18/09), dans les deux dépôts. C'est le rôle de `TECH_RANGES` :
ce que l'humain a tranché prime sur ce que la couleur raconte.

**RÉGLÉ LE 04/09/2026 — et la règle qui en sort.** Patrice a recolorié la case dans Teams. La
ligne porte maintenant son numéro et le planning publie `26-030 POSTE DE RION DES LANDES` de
lui-même, les 16, 17 et 18/09. **L'arbitrage a donc été retiré des deux dépôts.**

> **Un arbitrage manuel se retire dès que le planning porte le numéro.** Tant qu'il survit, il
> n'ajoute plus rien — mais le jour où nous enlevons quelqu'un du planning, lui continue de le
> montrer au technicien. Un arbitrage qui survit à la ligne du planning finit par mentir, et il
> mentira sans rien afficher d'anormal. Donc : à chaque `-Injecter`, comparer les arbitrages de
> `TECH_RANGES` aux dates que le planning publie déjà, et supprimer ceux qui sont devenus
> redondants. **Le retrait doit être prouvé, pas supposé** : le nombre de jours-personne et le
> nombre « venus de la recopie manuelle » doivent être identiques avant et après. Le 04/09/2026 :
> 170 et 15 dans les deux cas (chiffres du matin, avant la correction décrite juste en dessous).

**La question des 14 et 15/09 reste ouverte** : ces deux jours portent le texte `CATEC` (une
formation) sur fond vert-Portet. Le PGO de Rion couvre TELSAM du 14 au 18/09, plus large que la
présence réelle. Le 16/09, en revanche, n'est plus une déduction : le planning le dit.


### Didier PERRIN, Cantegrit et Portet (04/09/2026) — j'ai corrigé une donnée juste, faute d'avoir demandé DE QUELLE SEMAINE il parlait

Séquence exacte, elle vaut d'être relue en entier.

1. Le matin, j'annonce à Patrice : « Cantegrit (26-003) : Pascal BONAVENTURE rejoint Didier PERRIN
   et Hugues VIRY, du lundi 14 au vendredi 18. » C'était **la semaine 38**, et c'était exact.
2. Il répond : « Didier Perrin n'est pas sur Cantegrit **lundi**, il n'y a que Pascal Bonaventure
   et Hugues Viry. Comment peux-tu te tromper ? C'est grave si j'envoie de mauvaises informations
   à l'application et donc au technicien. »
3. J'ai lu « lundi » comme le lundi de la semaine dont je venais de parler (le 14). **Il parlait du
   lundi suivant, le 7 — la semaine 37.** Et sur la semaine 37, il avait raison : Didier est en vert,
   sur Portet `26-051`, ce que l'appli affichait déjà correctement.
4. J'ai donc **retiré Didier de Cantegrit sur la semaine 38, où il y est réellement**. J'ai cassé une
   donnée juste pour répondre à une objection qui portait sur une autre semaine.
5. Sa capture d'écran a tranché : elle s'arrêtait au vendredi 11/09, donc elle ne montrait pas la
   semaine dont je parlais. C'est ce détail qui a permis de comprendre le malentendu.

> **RÈGLE 1 — quand Patrice conteste une affectation, établir la semaine AVANT de toucher quoi que
> ce soit.** Un jour de la semaine sans son numéro de jour (« lundi ») est ambigu dès que la
> conversation porte sur plusieurs semaines, et le planning en montre six. Demander, ou nommer les
> deux lectures possibles — jamais modifier les données sur une interprétation.

> **RÈGLE 2 — toujours écrire la semaine ET la date.** « lundi 14/09 (semaine 38) », jamais « lundi ».
> Vaut pour mes messages, pour `APP_NOUVEAUTE` et pour les récapitulatifs. C'est cette imprécision
> qui a produit tout le malentendu, et elle m'a coûté deux corrections dans le mauvais sens.

**Ce que la mesure disait, et qui était juste** — semaine 38, colonnes 258-262 :

| Case | Couleur | `ColorIndex` |
|---|---|---|
| Pascal BONAVENTURE (L10), les 5 jours | `#9bc2e6` | 37 |
| Didier PERRIN (L11), les 5 jours | `#9bc2e6` | 37 |
| Hugues VIRY (L22), les 5 jours | `#9bc2e6` | 37 |
| ligne-projet L34 « … Cantegrit: Epissures + Recettes : FO comptage » | `#9bc2e6` | 37 |

Les trois cases sont rigoureusement identiques : aucune lecture ne peut mettre deux de ces personnes
sur Cantegrit et pas la troisième. En semaine 37, la même mesure donne Didier à `#00ff00` = L33 =
`26-051` Portet. Les deux lectures étaient bonnes ; c'est moi qui ai mélangé les semaines.

**DÉNOUEMENT — et c'est le vrai enseignement technique.** Patrice a écrit `26-003` dans la ligne
Cantegrit de la semaine 38. Effet mesuré immédiatement : les libellés sans numéro passent de 96 à
95, l'avertissement « ligne que l'appli ne pourra pas montrer » disparaît, et surtout le planning
publie **170 jours-personne dont 0 venus de la recopie manuelle**. Tous les arbitrages `TECH_RANGES`
de la semaine 38 sont donc devenus inutiles et ont été retirés des deux dépôts (Pascal 42→37 dates,
Hugues 41→36) : preuve que le retrait ne coûte rien, 170 et 0 avant comme après.

> **Un numéro écrit dans la ligne du planning vaut mieux que n'importe quel arbitrage de ma part.**
> Un arbitrage est un doublon manuel qu'il faut ensuite penser à retirer ; le numéro, lui, règle le
> cas définitivement et pour toutes les personnes que Patrice ajoutera plus tard sur la ligne. Donc :
> devant une ligne non rattachée, **demander le numéro d'abord**, et ne poser un arbitrage que s'il
> ne peut pas être obtenu.

**Confirmé par lui le 04/09/2026** : Benjamin SOUPA et Sid Ahmed BENZAMERA sont bien sur Portet
`26-051` les lundi 14 et mardi 15/09, malgré le texte `CATEC câbles BRC` / `CATEC` écrit dans les
cases. Le texte est une précision, pas un chantier concurrent : la couleur reste la référence.

## Planning validé / prévisionnel (demandé par Patrice le 03/09/2026)

**Le problème, dans ses mots** : « les techniciens ont accès aux plusieurs semaines suivantes, ce
qui n'est pas une bonne chose, sachant que nous pouvons modifier entre-temps le planning et qu'il
ne faudrait pas qu'eux réservent par exemple un hôtel ou un logement alors qu'on va les échanger de
sites. » Et ce qu'il veut : « une fois que le planning est fixé le vendredi, nous disons c'est bon,
le planning est bon pour la semaine prochaine, et les techniciens le voient passer comme planning
validé ».

### CE SUR QUOI L'APPLI SE BASE — la réponse à sa première question

**L'appli technicien ne lit PAS Teams.** Le mot `PLANNING_RTE` n'apparaît nulle part dans
`appli-techniciens/index.html` (vérifié le 03/09/2026). Elle affiche les chantiers d'un technicien
depuis **`TECH_RANGES`**, recopié à la main. La lecture de Teams ne sert qu'à la grille du suivi et
au contrôle de cohérence du matin.

Mesuré le 03/09/2026, pour répondre précisément : un technicien voyait jusqu'au **18 septembre**
(Benjamin SOUPA et Sid Ahmed BENZAMERA, 3 semaines devant), 1 semaine pour les dix autres — et les
boutons de semaine n'avaient **aucune limite**. Sa crainte était donc fondée, mais bornée par
hasard, pas par une sécurité.

### CE QUI A ÉTÉ ÉCARTÉ, ET PAR QUI

J'avais proposé de garder une **empreinte du planning** au moment de la validation, de la
recomparer chaque matin, et de **dévalider automatiquement** une semaine modifiée depuis.
**Patrice l'a refusé**, avec sa raison : « ce n'est pas la peine de tout modifier et de faire un
contrôle quotidien et systématique des chantiers ; si un jour nous devons changer un chantier dans
la journée, nous préviendrons le technicien et ça ne sera pas la peine de revalider une semaine. »

**Ne pas le rajouter « pour bien faire ».** Ce serait une alerte quotidienne pour un cas qu'ils
règlent au téléphone — exactement le mécanisme qui crie au loup jusqu'à ce qu'on l'ignore.

**Ses deux arbitrages du 03/09/2026** :
1. une semaine non validée **reste visible**, marquée « prévisionnel » (l'autre option était de ne
   rien montrer du tout) ;
2. **les cinq qui modifient le planning** peuvent cocher — pas Carine Rambaud ni Guillaume Heras,
   qui entrent pourtant dans le suivi.

### L'architecture, et pourquoi pas de manip Cloudflare

| brique | rôle |
| --- | --- |
| table `semaine_validee` | un interrupteur par semaine ISO (`2026-S38`), plus qui et quand |
| `functions/api/semaines.js` | valider/dévalider, **derrière la porte**, réservé aux 5 |
| `functions/api/public/semaines.js` | **la seule adresse du suivi sans mot de passe** : lit la liste |
| `functions/_middleware.js` | ouvre cette seule adresse, en GET/OPTIONS |
| `appli-techniciens/index.html` | bandeau vert/ambre + plafond à 6 semaines |

**La porte de service, plutôt qu'une liaison D1 sur le Worker relais.** L'appli vit sur un autre
domaine et n'a aucun mot de passe du suivi. L'autre solution était de relier la base au Worker
`relais-telsam` que l'appli utilise déjà : une liaison à ajouter à la main dans la console
Cloudflare, donc **une manip de plus pour Patrice sur une base en service**. Ici, zéro manip — la
table se crée toute seule (`CREATE TABLE IF NOT EXISTS`, idempotent), comme la colonne `note`.

**Ce que cette adresse expose, et rien d'autre** : une liste de numéros de semaine. Pas de
chantier, pas de nom de personne, pas de date de validation. Un inconnu qui la trouverait
apprendrait que TELSAM a validé la semaine 38.

**TROIS RÈGLES À NE PAS DÉFAIRE sur cette exception :**
1. le chemin est comparé **en entier** (`===`), pas en `startsWith` : un préfixe ouvrirait tout ce
   qu'on ajouterait un jour sous `/api/public/` sans que personne ne le décide ;
2. **GET et OPTIONS seulement** ;
3. **aucun champ en plus.** La tentation sera d'y mettre « validé par Patrice le 03/09 » pour faire
   joli dans l'appli : ce serait publier des noms de salariés sur une adresse ouverte, pour de la
   décoration.

### LA RÈGLE QUI COMMANDE TOUT LE RESTE : on n'affirme jamais « validé » sans preuve

Trois états côté appli, et le doute penche toujours du côté prudent :

| ce que dit la base | ce que voit le technicien |
| --- | --- |
| la semaine y est | vert, « Planning validé — tu peux réserver ton logement » |
| la semaine n'y est pas | ambre, « Planning prévisionnel — ne réserve rien » |
| **elle ne répond pas** | ambre, « impossible de vérifier pour le moment. Ne réserve rien » |

Se rabattre sur « validé » en cas de panne serait le pire des replis : **une coupure réseau ferait
réserver un hôtel.** Le bandeau dit ce qu'il faut FAIRE, pas un état (« ne réserve rien » plutôt
que « non validé »).

**Le cache existe pour le terrain**, sept jours (`kvSet('semaines_validees')`) : un technicien sur
un poste sans réseau doit pouvoir lire l'état de sa semaine. Risque assumé, et c'est Patrice qui l'a
tranché — si une semaine validée change, « nous préviendrons le technicien ».

**Côté suivi, deux garde-fous du même ordre :**
- **la case n'existe pas en mode « fichier seul »** (« validation indisponible hors ligne »). Une
  validation qui ne vivrait que dans le navigateur de celui qui a coché ne serait vue ni par les
  collègues ni par les techniciens : un bouton qui mentirait sur une décision qui envoie des gens
  réserver des logements ;
- **c'est le serveur qui applique le droit de valider**, la page ne fait que griser la case. Une
  page se modifie par celui qui la regarde ; `VALIDATEURS` vit dans `functions/api/semaines.js`.
- **Confirmation à la décoche seulement.** Cocher est le geste courant du vendredi ; décocher
  retire aux techniciens une information sur laquelle ils ont pu réserver.

**Le plafond à 6 semaines** est dans l'appli (`SEMAINES_DEVANT_MAX`), et **le passé n'est pas
borné** : les feuilles d'heures des semaines écoulées doivent rester accessibles.


**CORRECTION DANS L'HEURE — aucun bandeau sur la semaine en cours ni sur le passé.**
Patrice, juste après la mise en ligne : « je vois qu'il faut valider le planning de la semaine en
cours. J'espère que les techniciens n'ont pas eu de changement sur leur appli ? » Ils en avaient
un : un bandeau ambre **« ne réserve rien »** sur la semaine qu'ils étaient en train de faire.
Vérifié en interrogeant le site en ligne (`APP_VERSION` déployée, et `/api/public/semaines` qui
renvoyait `{"validees":[]}`) — pas supposé.

Ce bandeau répond à **une** question : « est-ce que je peux réserver mon logement ? ». Elle ne se
pose que pour une semaine **à venir**. Sur la semaine en cours le technicien est déjà sur le
chantier ; sur une semaine passée la question est absurde. Un avertissement affiché là où il n'a
pas d'objet inquiète pour rien, et surtout **il apprend à ne plus lire le bandeau** — donc à ne
plus le lire le jour où il compte. `majBandeauValidation` reçoit désormais l'offset et ne dit rien
pour `offset <= 0`. **Ne pas « uniformiser » en le remettant partout.**

Contrôle ajouté avec son contre-exemple : rien sur l'offset 0 (même semaine validée), rien sur
l'offset -1, mais la semaine suivante parle bien.

### Ce qui reste vrai et qu'il faut savoir

**FAIT LE MÊME JOUR, une heure plus tard** : l'appli reçoit son planning de Teams (voir
« L'appli technicien reçoit enfin le planning de Teams »). Ce paragraphe disait que le branchement
restait à faire ; livrer le bandeau sans lui a produit exactement le défaut annonce - un bandeau
qui promet un planning absent.

### La vue atelier ouverte EN LOCAL ne peut pas vérifier la validation (04/09/2026)

Patrice : « mes techniciens voient bien que leur planning pour la semaine prochaine est validé.
Par contre, moi, sur la vue atelier, j'ai toujours le message d'alerte comme quoi le chantier
n'est pas validé. Est-ce normal ? Pourtant, j'ai rechargé ma page ? »

**Sa capture d'écran a tranché en une seconde, et c'est elle qui a évité une fausse piste.** Elle
montrait **semaine 37** — la bonne — et le message « **impossible de vérifier pour le moment** »,
pas « ne réserve rien, ça peut encore changer ». Ces deux phrases ambres ne disent pas la même
chose : la première veut dire que la base n'a pas répondu, la seconde que la semaine n'est pas
validée. **Sans cette distinction, le diagnostic aurait été « vous êtes sur la semaine 38 », ce qui
était faux.** Garder les deux formulations distinctes.

**La cause, mesurée** : `/api/public/semaines` répond toujours
`Access-Control-Allow-Origin: https://pivot-telsam.github.io`, quelle que soit l'origine qui
demande. Depuis un `index.html` ouvert en double-clic (origine `file://`), le navigateur **bloque**
donc la réponse : `fetch` échoue en « Failed to fetch », `semainesValidees` reste `null`, et le
bandeau annonçait un incident réseau alors qu'il n'y en avait aucun. Vérifié en jouant le `fetch`
depuis une page `file://` : bloqué ; depuis l'origine publiée : 200.

> **On ne corrige PAS ça en élargissant l'autorisation côté serveur.** `Origin: null` couvre tout
> contexte bac-à-sable, donc l'autoriser ouvrirait cette adresse — la seule du suivi sans mot de
> passe — à n'importe quelle page locale. La règle reste : garder l'exception aussi étroite que
> possible. C'est le MESSAGE qu'on corrige.

`majBandeauValidation` a donc une branche de plus, avant celle de `null` : si la page est en
`file:`, elle dit **pourquoi** elle ne peut pas vérifier et renvoie vers l'appli en ligne. **La
propriété de sûreté ne bouge pas** : on ne dit jamais « validé » sans preuve.

**Éprouvé dans les quatre combinaisons** (semaine 37, offset +1) : local + base muette ⇒ nouveau
message ; **en ligne + base muette ⇒ l'ancien message, celui du vrai incident, est conservé** ;
base qui répond ⇒ « validé » dans les deux cas. Le `location` est passé en paramètre au harnais
pour masquer le global et jouer les deux protocoles sans toucher au code de l'appli.

**À redire à Patrice si le cas revient** : pour voir ce que voient les techniciens, il faut ouvrir
l'appli **en ligne**, pas le fichier du dépôt. Le fichier local est ma copie de travail : il peut
être en cours de modification, et il ne sait pas interroger la base.

### Testé — 64 contrôles, chacun avec son contre-exemple

`partage/test-api-semaines.html` (27), `partage/bloc-test-validation-suivi.html` (16),
`appli-techniciens/scripts/test-validation.html` (21). Jamais contre la vraie base D1.

Les contrôles qui comptent vraiment :
- **la porte** : `/api/public/semaines` passe en GET ; le **POST** sur la même adresse ne passe
  pas ; `/api/public/semaines/autre` ne passe pas ; `/api/planning`, `/` et
  `/functions/_middleware.js` restent fermés ;
- **le droit de valider** : la compta reçoit 403 **et rien n'est écrit**, alors qu'elle voit l'état ;
- **la panne réseau** : sur une semaine réellement validée, base injoignable ⇒ le bandeau reste
  ambre. C'est le contrôle qui protège de l'hôtel réservé pour rien ;
- **le cache** : frais ⇒ on garde ce qu'on savait ; périmé de 30 jours ⇒ on ne sait plus, donc
  prévisionnel ; réponse inattendue ⇒ on ne sait pas (jamais « tout validé ») ;
- **le mode fichier seul** : aucune case à cocher dans l'en-tête.


## L'appli technicien reçoit enfin le planning de Teams (03/09/2026)

**Ce qui l'a déclenché.** Une heure après la mise en ligne du bandeau « planning prévisionnel »,
Patrice : « ils ont bien le message comme quoi le planning n'est pas validé. Par contre, **ils n'ont
pas le planning**. Donc il y a un problème d'affichage. »

**Ce n'était pas l'affichage.** Mesuré avant de toucher au code : pour la semaine du 07/09, l'appli
pouvait montrer **1 personne** quand Teams en plaçait **7** ; pour celle du 14/09, **2 contre 10**.
Le bandeau annonçait donc un planning que l'appli n'avait jamais reçu. La cause est celle notée la
veille sans être traitée : l'appli ne connaissait que `TECH_RANGES`, que je recopiais à la main.

**LEÇON DE MÉTHODE.** J'avais présenté ce branchement comme « le préalable » puis livré le bandeau
sans lui. Un bandeau qui promet une donnée absente est pire que pas de bandeau : il fait douter de
l'outil. **Quand une fonction dépend d'un préalable, on livre le préalable ou on ne livre rien.**

### Comment ça marche

`planning-rte.ps1 -Injecter` écrit désormais, dans l'appli, la constante **`PLANNING_TECH`** :
`{maj, du, au, tech:{ chantierId: { personne: [jours] } }}` — **six semaines** à partir du lundi
courant, le même plafond que celui appliqué côté appli (`SEMAINES_DEVANT_MAX`).

**UNE SEULE SOURCE PAR JOUR, JAMAIS L'UNION** (`datesPour` dans l'appli) : dans la fenêtre publiée,
c'est le planning qui dit tout ; avant elle, `TECH_RANGES`. Faire l'union aurait ramené le défaut
Chaineau du matin même — une personne sur deux chantiers le même jour, avec deux jeux de documents,
donc le risque de partir avec le mauvais dossier. Vérifié sur les 7 semaines publiées : aucune
personne sur deux chantiers le même jour.

**REPLI COMPLET SI LE BLOC EST VIDE** (`planningTechActif`) : tant que le script n'a jamais tourné,
ou s'il a échoué, l'appli retombe entièrement sur `TECH_RANGES`. Sans ce garde-fou, un bloc vide
effacerait la semaine en cours de **tout le monde** — douze techniciens sur le terrain devant
« rien de prévu ». Contrôle de non-régression dans le harnais : bloc neutralisé ⇒ 11 personnes
servies quand même.

**Le complément manuel ne comble que les trous.** `TECH_RANGES` n'entre dans la fenêtre que pour
les couples jour-personne que le planning ne place pas — un arbitrage humain là où le planning est
muet. Un jour où le planning place quelqu'un ailleurs écrase donc la vieille recopie, au lieu de
s'y ajouter.

**Les lignes sans numéro sont SIGNALÉES, jamais tues.** Le script liste ce que le planning place et
qu'il ne sait pas rattacher : le technicien ne verrait rien ce jour-là. Au premier passage,
15 cases sur 3 lignes — Bissy - Grand Île (26-064), MTFO Cross-Sausset (26-002) et « Racco et
recette réseau liaisons BI-BR Poste de Cantegrit » (26-003), toutes sans numéro dans leur libellé.
Arbitrage posé dans `TECH_RANGES` des deux dépôts, et **la vraie solution reste que Patrice écrive
le numéro dans Teams** : ce jour-là ces entrées deviennent redondantes et inoffensives.
*Défaut corrigé aussitôt dans ce même avertissement : il comparait à ce que le planning avait su
rattacher, donc il continuait de crier après que l'arbitrage avait couvert les cases. Il compare
maintenant à ce qui est RÉELLEMENT publié.*

Résultat mesuré, avant/après : semaine en cours 11 → 11, **semaine +1 : 1 → 10**,
**semaine +2 : 2 → 11**. La semaine +3 reste à 0 — Teams n'est pas rempli plus loin, et c'est la
vérité du fichier, pas un défaut.

## Les cases cochées restent cochées (03/09/2026)

**Demande de Patrice** : « quand les techniciens cochent sur leur suivi, le pylône coché ne reste
pas coché. Mettons, lundi il coche tel pylône ; le mardi, quand il va ouvrir le suivi, ce pylône
sera marqué décoché. Cela n'a pas d'incidence sur le rapport. Par contre ça peut les amener en
erreur, eux, ou **si on envoie une autre équipe sur le même chantier, ils auront l'impression
qu'aucun pylône n'a été fait** s'ils ne lisent pas l'avancement du chantier. »

Son diagnostic était exact, y compris sur le rapport : la consolidation fait une **union**, un
pylône déclaré ne se perd pas. Le défaut était d'affichage — mais un affichage qui montre un
chantier vierge alors qu'une équipe y a travaillé fait refaire du travail, ou douter de son envoi.

### DEUX ÉTAGES DE MÉMOIRE, et les deux sont nécessaires

| étage | contenu | fraîcheur | où |
| --- | --- | --- | --- |
| `AVANCEMENT_DECLARE` | ce que **tout le monde** a déclaré | refait chaque matin par `boites-posees.ps1` | constante dans l'appli |
| `kvSet('suivi_<numéro>')` | ce que **ce technicien** a envoyé | immédiat, et marche sans réseau | IndexedDB du téléphone |

Retirer l'un laisse un trou : sans le local, il ne revoit pas ses propres coches de la journée
(la consolidation ne repasse que le matin) ; sans le partagé, il ne voit pas celles des autres
équipes — le cas que Patrice a nommé.

**Ce qui est publié va au-delà des boîtes.** `boites-posees.ps1` ne comptait que les tâches reliées
à un lot par `BOITES_TACHES` ; `AVANCEMENT_DECLARE`, lui, publie **toutes** les tâches, y compris
celles sans pylône (statut « fait »/« en cours ») et les cantons. Un chantier dont le devis ne vend
aucune boîte doit quand même remontrer ses coches : c'est précisément là que le technicien se
demanderait s'il a bien envoyé. Premier passage : 11 cases sur 3 chantiers.

**À l'écran** : les cases déjà déclarées s'ouvrent cochées, celles d'un **autre** technicien en
bordure pointillée, l'infobulle disant « Déjà déclaré par Pascal BONAVENTURE le 02/09/26 », et un
rappel compte les cases (« ✔ 3 pylônes déjà déclarés. Décoche seulement si c'est une erreur. »).
Sans le « par qui », un technicien qui trouve une case cochée qu'il n'a pas cochée doute de l'outil.

**Union à l'écriture aussi** (`retenirSuivi`) : un envoi du mardi ne doit pas effacer ce qui a été
déclaré lundi. Et **on ne retient QUE ce qui est parti** — retenir avant l'envoi ferait croire que
le travail est déclaré alors qu'il est resté dans le téléphone.

### DÉFAUT TROUVÉ PAR LE BANC D'ESSAI : une promesse qui ne se règle jamais

Le harnais s'est **figé** sur `await kvGet(...)`. Cause : IndexedDB ne répond pas sous
`--virtual-time-budget`. Mais ce n'est pas un artefact de test — sur un vrai téléphone, IndexedDB
peut ne jamais répondre (navigation privée, stockage désactivé, base verrouillée par un autre
onglet). **Une promesse qui ne se règle pas ne lève aucune erreur** : le `try/catch` ne l'attrape
pas, et `ouvrirSuivi` restait bloqué AVANT d'afficher la fenêtre. Le technicien aurait appuyé sur
« Suivi » et **rien ne se serait passé**.

D'où `avecDelai(promesse, 1500)` : au-delà, on continue avec ce qu'on a. À garder sur tout accès à
IndexedDB dont dépend l'ouverture d'un écran.

**Autre correction du même passage** : `boites-posees.ps1` sortait sur « rien de nouveau » en ne
comparant que le suivi. Le jour où seules les cases de l'appli changent (une tâche sans pylône
passée de « en cours » à « fait », qui ne compte aucune boîte), l'appli n'aurait jamais été mise à
jour. Il compare maintenant **les deux fichiers**.

### Testé — 120 contrôles au total, 0 échec

`appli-techniciens/scripts/test-planning-appli.html` (21) s'ajoute aux suites existantes
(25 + 22 + 9 + 16 + 27), toutes rejouées sans régression. Les contrôles qui comptent :
- semaine +1 servie par au moins 6 personnes (elle en servait **1**) ;
- les deux lignes sans numéro rattachées à leur fiche ;
- **personne sur deux chantiers le même jour** sur les 7 semaines ;
- bloc planning neutralisé ⇒ repli complet, jamais d'écran vide ;
- cases pré-cochées **et** toutes les cases pas cochées (sinon le test ne prouverait rien) ;
- un chantier sans déclaration s'ouvre entièrement décoché ;
- la mémoire locale fait bien l'union de deux envois successifs.


## « Avec qui je suis » sur la fiche chantier (demandé par les techniciens, 03/09/2026)

**Ce n'est pas une demande du bureau, elle vient du terrain.** Patrice la relaie : « les techniciens
me remontent une info importante et intéressante, c'est qu'ils ne savent pas avec qui ils sont la
semaine suivante. Cela leur permettrait de **mutualiser les logements**. »

Elle n'a pu être satisfaite que parce que l'appli reçoit désormais le planning (section
précédente) : sans les semaines à venir, la question n'avait pas de réponse.

**LES JOURS DE CHACUN COMPTENT AUTANT QUE LE NOM.** C'est tout l'intérêt pour un logement : savoir
que Bilal est là toute la semaine et Younes seulement lundi et mardi change la réservation. Un
simple « avec Bilal et Younes » aurait laissé croire à une semaine complète à trois. La ligne porte
donc le nom à gauche et, à droite, soit ses jours (`Lun Mar`), soit « mêmes jours que toi ».

**« Tu es seul sur ce chantier cette semaine » s'affiche aussi.** Ne pas remplacer ça par un bloc
absent : « seul » est une réponse, une absence de bloc laisse se demander si l'information manque.

**Les noms viennent de la DONNÉE, pas de la liste `TECHS`** — c'est le planning qui dit qui est là.
Les encadrants en sont écartés **à la source** (`planning-rte.ps1` les exclut de la présence : une
visite ou une réunion n'est pas une équipe sur le chantier), donc aucun risque d'annoncer un
colocataire qui ne viendra pas dormir. Si un jour Patrice veut y voir François Vidal en MTFO, c'est
la règle d'exclusion du script qu'il faudra discuter, pas ce bloc.

`equipeSurChantier(cid, mon, sun, moi)` + `blocEquipe(...)`, posés sur la fiche juste après les
jours — le « quand » et le « avec qui » se lisent ensemble. Tout est calculé dans la page à partir
de `datesPour`, **aucune donnée nouvelle à publier** : le planning par technicien était déjà là.

**Un cas a été testé bien qu'aucune donnée réelle ne le produise** : un collègue présent 2 jours
quand j'y suis 5. C'est le plus important pour un logement — s'il s'affichait « mêmes jours que
toi », le technicien réserverait la semaine à deux pour rien. La branche est donc éprouvée avec un
cas fabriqué, puis la donnée réelle restaurée (et cette restauration est elle-même vérifiée). **Ne
pas laisser une branche non jouée sous prétexte que le planning du moment ne la produit pas.**

Vérifié sur la donnée réelle de la semaine du 07/09 : Anthony DENIS lit « Avec toi sur ce chantier
(5) » sur Portet 26-051, Morad EL ABBASSI lit « Tu es seul » sur Bollène 26-071.

**Les trois numéros manquants ont été écrits dans Teams par Patrice le 03/09/2026** (Bissy - Grand
Île, MTFO Cross-Sausset, Racco Cantegrit) : le script ne signale plus aucune ligne non rattachée, et
la présence publiée est passée de 14 à **16 chantiers**. Les entrées d'arbitrage correspondantes
dans `TECH_RANGES` deviennent redondantes dans la fenêtre publiée (le planning prime) et restent
justes pour les dates passées — les laisser.

## Ordre d'affichage des chantiers — Gantt ET Boîtes & nacelle (posé par Patrice le 01/09/2026)

**Les deux vues affichent les chantiers dans le MÊME ordre**, calculé par `chronoSortKey()` et
exposé hors du Gantt par `chronoKeyParChantier()` / `chronoKeyDe()`. Demande de Patrice :
« d'abord ceux où les techniciens sont placés, puis ceux où ils sont censés être placés ».
Deux niveaux, chronologiques à l'intérieur de chaque niveau :
1. **présence réelle** — segments `normal` / `tourets`, c'est-à-dire des jours confirmés par le
   code couleur du planning RTE (`REAL_DAYS`) ;
2. **tout le reste de ce qui est à venir** — fenêtre prévisionnelle (`prevu`), couverture PGO,
   consignation, NIP, rappel MTFO.
Puis, tout en bas : ce qui est entièrement passé, puis ce qui n'a aucune date. Un chantier terminé
ou absent du Gantt reçoit la clé maximale (`chronoKeyDe`) et finit donc en fin de liste.

**Le type de barre `prevu` (ajouté le 01/09/2026) — à ne pas confondre avec un retour en arrière
sur l'incident du 20/08/26.** Depuis cette correction, une fenêtre annoncée dans `fenetres` sans
aucun jour confirmé dans `REAL_DAYS` ne produit **pas** de barre bleue de présence : cette règle
tient toujours, elle est ce qui empêche de croire qu'un technicien est placé alors que non.
Son effet de bord, lui, était mauvais : **14 chantiers actifs sur 46 ne s'affichaient plus du tout**
sur le Gantt et retombaient en vrac dans la ligne « sans date connue » du bas. Ils réapparaissent
donc, mais avec un type de barre **distinct** : `prevu`, pâle, hachurée, bordure en pointillés,
texte en italique, infobulle « Prévu, aucun technicien encore placé au planning ». **Aucun calcul
d'alerte ne doit s'appuyer sur ce type** — les alertes PGO filtrent sur `'normal'`, les garder
ainsi. Si un jour on retouche ce point : la barre pleine = quelqu'un est placé, la barre hachurée
pâle = personne, et rien d'autre ne doit brouiller cette lecture.

**Vue Boîtes & nacelle** : le tri chronologique est le **défaut** ; l'ancien tri « plus gros reste
à poser d'abord » reste accessible dans le menu déroulant `boitesTri` (il sert à préparer une
commande de boîtiers, où les dates ne comptent pas). Les trois listes secondaires de la vue
(« À compléter », « Nacelle », « Remarques ») suivent le même ordre chronologique.

## Lire un PGO/PDP : la forme du document ne se négocie pas

**Le format des PGO et des PDP nous est imposé** — c'est le SPS ou l'AEU qui les édite et les
envoie, TELSAM n'a aucune main dessus (rappelé par Patrice le 02/09/2026). **Ne jamais demander
« la version Excel plutôt que le PDF »** : la question est sans objet et fait perdre du temps à
tout le monde. C'est à l'outil de s'adapter au document reçu, pas l'inverse.

**Ce qui marche, dans l'ordre à essayer :**
1. **`.xlsx` / `.xls`** → Excel COM. Modèle prêt à l'emploi dans le scratchpad :
   ouvrir en lecture seule, parcourir `UsedRange.Value2`, filtrer les lignes contenant « TELSAM ».
   Les dates ressortent en numérique série : `[DateTime]::FromOADate($v)`.
2. **`.doc` / `.docx`** → Word COM, `$doc.Content.Text`. **Attention, le tableau chiffré d'un devis
   TELSAM n'est pas dans le texte** : c'est un objet Excel embarqué
   (`InlineShapes.Item(1).OLEFormat.Object`, ProgID `Excel.Sheet.8`), à ouvrir comme un classeur.
   C'est comme ça qu'on lit les articles d'un devis.
3. **`.pdf` texte** → décompresser les flux du PDF et récupérer les chaînes entre parenthèses.
   **Piège : les polices espacent les caractères**, « TELSAM » y apparaît comme « T E L S A M ».
   Toujours chercher sur le texte **compacté** (`-replace '\s+',''`), sinon on conclut à tort que
   TELSAM n'est pas mentionné.
4. **`.pdf` image** (le PGO Sausset ind.47 en est un : 1 page, 1 table d'en-tête, 1 image) → Word
   n'en tire que l'en-tête. Là, seule la méthode 3 fonctionne ; si elle échoue aussi, le dire à
   Patrice plutôt que d'inventer.

## Pièges connus dans le code (à ne pas "corriger" par erreur)
- **METTRE À JOUR UN PGO/PDP = METTRE À JOUR LES DATES STRUCTURÉES, PAS SEULEMENT LE TEXTE DE
  L'INDICE.** Le Gantt dessine la couverture PGO à partir de `pgo.couverture` (tableau de
  « JJ/MM/AA - JJ/MM/AA : libellé ») et les alertes PDP à partir des champs `pdp`/`perimetre`, PAS
  à partir du texte `pgo.indice`. Vécu le 28/08/2026 : j'avais monté l'indice PGO de Bradascou de 15
  à 16 dans le texte mais pas ajouté la fenêtre 31/08-04/09 dans `pgo.couverture` → le nouveau PGO
  n'apparaissait pas sur le Gantt en S36, là où des techniciens sont planifiés (repéré par Patrice).
  Donc à chaque changement d'indice : mettre à jour `pgo.couverture` (et `validFrom`/`validUntil`)
  en cohérence avec les fenêtres réellement couvertes. Contrôle : ouvrir le Gantt et vérifier que la
  couverture apparaît bien aux semaines concernées.
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

**Le 07/09/2026, Patrice a numéroté 2025 : `25-001` à `25-085` dans la colonne F.** Ce ne sont
**PAS** des n° de chantier au sens de `26-0XX` : comme 21- à 24-, **ce sont les anciens devis de
« DATA », l'ancienne entité du groupe** — « tu peux les laisser de côté ». L'onglet Affaires les
écarte, et le dit. **Ne pas les reprendre comme clé de fiche, ni renommer un dossier Dropbox
avec.** Détail et vérification (DATACC vs TELSAMCC) dans la section « Onglet Affaires ».

Il a aussi numéroté **12 chantiers de plus en `26-`** ce jour-là, tous sur des lignes datées 2025 :
`26-002`, `26-016`, `26-019`, `26-031`, `26-036-2`, `26-052`, `26-053`, `26-054`, `26-055`,
`26-064`, `26-065`, `26-066`. **Vérifiés un par un contre les fiches : les 12 concordent** (libellé
Excel ↔ nom de fiche). Effet mesuré sur l'onglet Affaires : 41 → **53 affaires**, et les affaires
vivantes sans numéro 28 → **14**.

**DEUX NUMÉROS 25- ÉTAIENT UTILISÉS DEUX FOIS — tranché par Patrice et corrigé le 07/09/2026.**
Il garde le numéro pour Joncquière et pour La Perche, les deux autres chantiers en reçoivent un
neuf à la suite de la série :

| n° | ligne | chantier | décision |
|---|---|---|---|
| `25-045` | 385 | LA JONCQUIERES - ST CEZAIRE (DATACC25060) | **garde 25-045** |
| `25-086` | 383 | LS GENAS - MIONS (DATACC25056) | numéro neuf |
| `25-065` | 405 | TOURETS LA PERCHE-LA TOUR DE CAROL (DATACC25084) | **garde 25-065** |
| `25-087` | 406 | LS PEYBER CHRIST - QUIRIN 1&2 (DATACC25085) | numéro neuf |

Un même numéro sur **plusieurs lignes du même chantier** reste normal : une affaire porte 1..n
devis. Le doublon n'en était un que parce que les libellés désignaient deux chantiers différents —
c'est le test à refaire (`même numéro + libellés différents`), pas « même numéro ».

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

## Nouveau chantier — checklist systématique, et les DEUX mécanismes qui la portent

**Cette checklist a été appliquée à moitié deux fois, et les deux fois c'est quelqu'un d'autre qui
l'a vu.** Le 21/08/2026 : dossiers App Tech complets mais sans File Request ni `depotTerrain`,
donc pas de bouton photos — signalé par Pascal Bonaventure. Le 02/09/2026 : dossier App Tech monté
sans `tachesVendues`, donc pas de bouton Suivi — signalé par Patrice, avec la phrase qui résume
l'enjeu : **« je vais être obligé de vérifier systématiquement ce que tu fais, et ce n'est pas le
but de cet outil ».**

**Ne pas répondre à ça en réécrivant la règle.** Elle était déjà écrite ici, et ça n'a pas suffi —
c'est la leçon générale des « Règles de prudence » : une règle qu'aucun mécanisme ne porte ne tient
pas. Deux mécanismes la portent désormais :

**1. `appli-techniciens/scripts/check-chantier-complet.sh` — bloque le commit.**
Dès qu'une fiche non terminée porte `documentsAppTech`, elle DOIT porter `depotTerrain` et
`tachesVendues`. Les trois boutons vont ensemble ; il refuse le commit sinon, en nommant le
chantier et ce qui manque. Pris automatiquement par `scripts/pre-commit.sh`. **Testé sur les deux
oublis réels rejoués (21/08 et 02/09) et sur le cas sain.** Une exception se déclare dans le script
avec sa raison — jamais un oubli silencieux.
*Deux pièges rencontrés en l'écrivant* : en awk POSIX `\{` n'est pas portable (utiliser `index()`,
pas une expression régulière), et **découper `SEED_DATA` sur `{"id":"` coupe la fiche au milieu de
`tachesVendues`**, dont chaque tâche porte aussi un `id` — découper sur `{"id":"c_`.

**2. `suivi-chantiers/scripts/controle-chantiers.ps1` — audite tout, chaque matin.**
Étape de `matin.ps1`, résultat dans `veille/controle-chantiers.json`, repris **en ligne rouge dans
le récap** et **injecté au démarrage de session**. Il vérifie, chantier actif par chantier actif :
les trois champs de l'appli, `dossierDropbox` / boîtes / `heuresPrevues` côté suivi, et dans
Dropbox le dossier App Tech, le brief, le PDP, le PGO, le sous-dossier Photos terrain.
- **Priorité** : un chantier est « à traiter » si un technicien y est placé sous 21 jours **ou** si
  une fenêtre de `pgo.couverture` y commence sous 21 jours. Le reste tient en une ligne.
- **Ce réglage a demandé trois corrections le 02/09**, toutes trouvées en vérifiant les
  signalements au lieu de les croire : le dossier App Tech de Fleyriat est au niveau du
  regroupement et non du chantier (rapprocher dans les deux sens) ; le PDP de Portet s'appelle
  `2025-PYR-PORTET-P.SIM-363-V08_260724.pdf`, son nom ne contient pas « PDP » (comparer aussi à
  `pdp.ref`) ; « Plan de Prévention V2.pdf » ne correspondait pas au motif à cause du **é**
  (retirer les accents avant toute comparaison). Puis un quatrième resserrage : ne considérer que
  `pgo.couverture`, jamais `fenetres` — une fenêtre seulement annoncée est du prévisionnel, et la
  version large sortait 9 chantiers « urgents » dont 8 où personne n'ira. **Un contrôle bruyant
  finit ignoré : mieux vaut le régler jusqu'à ce que chaque ligne mérite d'être lue.**
- Quand un document ne peut être ni reconnu ni rapproché d'une référence, il est déclaré
  **« invérifiable »**, pas « absent », et le contrôle demande de renseigner `pdp.ref`/`pgo.ref`.

**3. `hook-veille-prompt.ps1` rappelle les chantiers prioritaires à CHAQUE message** (ajouté le
04/09/2026). Les deux mécanismes ci-dessus ne suffisaient pas, et voici la preuve.

**LE CONTRÔLE AVAIT RAISON, ET PERSONNE NE L'A LU.** Le 04/09/2026 à 07h58,
`controle-chantiers.json` signalait, en **prioritaire** :

```
26-003 Poste Cantegrit :
   bouton Documents (documentsAppTech) ; bouton Depot photos (depotTerrain) ;
   bouton Suivi (tachesVendues) ; dossierDropbox ; AUCUN dossier App Tech dans Dropbox
```

Trois techniciens y sont placés en S37 — **la semaine que Patrice venait de valider** — et en S38.
C'est lui qui l'a découvert, en ouvrant le dossier Dropbox : « Peux-tu m'expliquer pourquoi quand
je vais sur le dossier de cantegrit, je ne vois pas l'App Tech qui est créé ? Comment vont faire
les techniciens pour le suivi, les photos, et caetera. Tu es censé le faire systématiquement dès
que nouveau chantier est affecté au technicien. »

Le mécanisme a donc parfaitement fonctionné : la faute est entière de mon côté. Mais la CAUSE est
structurelle et déjà connue de ce fichier : **le résultat du contrôle n'apparaissait qu'une fois,
au démarrage de session, et dans le récap.** Après quoi le traiter dépendait de ma mémoire — c'est
mot pour mot le défaut qui avait fait oublier la veille deux fois (28 et 31/08) et `SEED_VERSION`
deux fois avant elle. **Un signal affiché une fois vaut une promesse.**

D'où l'extension du rappel insistant de la veille aux chantiers incomplets. Le hook a maintenant
**deux sources et deux extincteurs séparés** :

| source | ce qu'elle rappelle | extincteur |
|---|---|---|
| `veille/dernier.json` | PGO/PDP/IST nouveaux ou d'indice supérieur | `marquer-veille-traitee.ps1` |
| `veille/controle-chantiers.json` | chantiers **prioritaires** incomplets | `marquer-chantiers-traite.ps1` |

**QUATRE DÉCISIONS DE CONCEPTION, à ne pas défaire :**

1. **Seuls les prioritaires.** Au 04/09/2026 : 48 chantiers actifs, **37 incomplets**, **4
   prioritaires**. Rappeler les 37 à chaque message rendrait le rappel illisible en une matinée —
   et un contrôle bruyant finit ignoré, ce qui est exactement le mal qu'on soigne. La règle de
   lecture de Patrice (25/08/2026) s'applique ici : un chantier où personne ne va peut attendre.
2. **Chaque source a son propre `try`.** Un `controle-chantiers.json` abîmé ne doit pas faire taire
   le rappel de la veille — ce serait l'inverse du but. **Vérifié** en y mettant du JSON invalide :
   la veille sort toujours, le bloc chantiers se tait.
3. **La note reste obligatoire pour éteindre.** Sans elle, l'extincteur deviendrait un bouton
   « faire taire ». Constater qu'un chantier ne demande rien EST un traitement — mais il faut
   l'avoir constaté et l'écrire. `-Lister` montre sans marquer.
4. **La sortie du hook est écrite en octets UTF-8.** Défaut vu le jour même dans le rappel reçu :
   « NON TRAIT?E ». PowerShell 5.1 encode sa sortie standard dans la page de codes de la console,
   pas en UTF-8 : le JSON était juste, son encodage non. Ne pas revenir à un `ConvertTo-Json`
   rendu au pipeline.

**Éprouvé dans les cinq sens** le 04/09/2026 : les deux blocs sortent sur l'état réel ; dossier de
veille vide ⇒ zéro octet ; contrôle abîmé ⇒ la veille survit ; aucun prioritaire ⇒ pas de bloc
chantiers mais la veille sort ; note trop courte ⇒ refus, `-Lister` ⇒ rien d'écrit.

**ET UN DOUBLON ÉVITÉ DE JUSTESSE, qui vaut comme leçon.** Mon premier réflexe a été d'écrire un
nouveau script de contrôle. Il a reproduit, en une heure, **exactement les trois faux positifs que
`controle-chantiers.ps1` avait déjà corrigés le 02/09** : le PDP de Portet dont le nom ne contient
pas « PDP », le « Plan de Prévention V2.pdf » de Bissy, et le dossier App Tech de Fleyriat placé au
niveau du regroupement. Script supprimé. **Avant d'écrire un contrôle, chercher celui qui existe
déjà — ce fichier le décrit — et se demander si le défaut est dans le contrôle ou dans le fait que
personne ne le lit.** Ici, c'était le second.

**Ce qui reste fait EN SESSION, et pourquoi.** Monter un dossier App Tech suppose de choisir le bon
devis parmi plusieurs, de juger si une IST est bien de TELSAM, de lire le bon indice de PGO,
d'écrire un brief sans rien de commercial. Ces jugements ne s'automatisent pas sans risque — c'est
le même raisonnement que pour la veille (« une détection ratée coûte une journée, une copie
automatique fausse met un document périmé entre les mains d'un technicien »). Ce qui est garanti,
c'est **qu'un oubli ne peut plus passer inaperçu** : il bloque le commit, il ressort le matin, et il revient à chaque message tant qu'il n'est pas marqué traité.

**À faire À CHAQUE nouveau chantier, sans attendre qu'il le demande.** Patrice l'a redit le
28/08/2026 (« note bien la règle pour le faire systématiquement »). L'ordre :
1. **Attribuer le n° d'index** (prochain `26-0XX` libre) et le communiquer à Patrice.
2. **Créer la fiche** dans les DEUX dépôts (suivi = fiche riche, appli = fiche allégée, même `id`),
   avec `mesureTouret:true` si c'est une mesure de tourets.
3. **Prévisionnel** : calculer `heuresPrevues` depuis le devis (cf. « Temps prévisionnel par
   chantier ») et **relever les boîtes** (cf. « Boîtes WTC2 »). Si le chantier a des boîtes ET des
   tâches vendues avec des numéros de pylônes, **l'ajouter à `BOITES_TACHES`** dans le même geste —
   sans cette entrée, les poses déclarées par les techniciens ne seront jamais comptées.
4. **Monter le dossier App Tech complet** (cf. section ci-dessous) : Brief + MO/NDS TELSAM +
   PDP/PGO à jour + IST si TELSAM ET signée RTE + sous-dossier **Photos terrain**.
5. **Demande de fichiers** Dropbox vers `Photos terrain` → câbler `depotTerrain` (bouton Dépôt
   photos) dans l'appli. **Jamais reporté à plus tard.**
6. **Lien de partage** du dossier App Tech → `documentsAppTech` (appli) + `dossierDropbox` (suivi).
   L'outil crée le lien en **privé** : le **signaler à Patrice pour qu'il le passe en public +
   mot de passe** (cf. « Accès Dropbox » et la serrure 2). Le bouton Documents ne marche pour les
   techniciens qu'après.
7. **Bumper** `SEED_VERSION` (suivi) et `APP_VERSION` + `APP_NOUVEAUTE` (appli), committer, pousser.

Exemple fait le 27-28/08/2026 : **26-070 Campagnac - Séverac** (mesure de tourets) — fiche,
prévisionnel 241 h, 8 WTC2, App Tech (Brief généré + NDS + MO + PDP signé + Photos terrain),
demande de fichiers créée, lien App Tech à rendre public. PGO absent en PDF (seulement `.xlsm`) et
PDP nommé « CTEAM » (couverture TELSAM à confirmer) — signalés à Patrice.


### Doublon inter-lots : deux hommes comptés deux fois pendant cinq jours (03/09/2026)

**Ce qui s'était passé.** `TECH_RANGES` mettait Sid Ahmed BENZAMERA et Vincent PERRIN sur le
**LOT 1 RODA et sur le LOT 2 SELT** du 31/08 au 04/09, dans les deux dépôts. Pas un arbitrage :
l'union des deux. Le lot 2 avait été ajouté sans retirer les jours du lot 1. Effet : deux hommes
comptés deux fois en présence, en heures, et en **couverture PGO**.

**Rien à arbitrer, contrairement à la semaine 34.** Sur les lignes de Chaineau, le planning Teams
écrit le numéro **en clair dans le libellé** : « lot RODA :MAP 26-036-1 » les 31/08 et 01/09,
« SELT MAP) 26-036-2 » les 02, 03 et 04/09. Quand le numéro est écrit, c'est le planning qui
tranche ; c'est seulement quand le lot n'est porté que par la couleur qu'il faut demander à
Patrice.

**Le piège à ne pas répéter en corrigeant** : `REAL_DAYS` ne doit perdre un jour que si **plus
personne** ne reste sur ce lot ce jour-là. Retirer les jours « en même temps que » les personnes
effacerait la présence réelle d'un autre technicien.

**LE CONTRÔLE NE POUVAIT PAS VOIR CE DÉFAUT — c'est le vrai enseignement.** `planning-rte.ps1`
indexait la recopie manuelle par `"jour|personne"` → **UN** id, en écrasant :

```powershell
$manuel[$cle][$depot] = $cid        # AVANT : la deuxième affectation efface la première
```

Une personne sur deux chantiers le même jour ne laissait donc voir que le dernier lu. Résultat : le
contrôle signalait **2 jours sur 5** — ceux où ce dernier lu différait du planning — et se taisait
sur les trois autres. Un contrôle qui perd l'information avant de la comparer ne peut pas trouver
ce qu'il cherche.

Corrigé en trois points :
1. la recopie est indexée par **liste** d'ids, sans écrasement ;
2. le **doublon est une alerte à part entière** (« une personne inscrite sur PLUSIEURS chantiers le
   même jour »), indépendante du planning — elle aurait crié même si le planning avait été muet ;
3. un désaccord n'est déclaré que si **aucun** des chantiers recopiés ne correspond au planning,
   sinon un même fait sortait en deux alertes et l'ensemble se met à crier au loup.

**Vérifié dans les deux sens** (règle : un test qui ne peut pas échouer ne prouve rien) — passe 1
sur les données corrigées : « aucun désaccord, aucun doublon » ; passe 2 avec un doublon réinjecté
exprès : l'alerte nomme la bonne personne, le bon jour, les deux lots et le bon dépôt, et **ne**
sort **pas** de deuxième alerte en désaccord.

**Piège de vérification rencontré au passage** : un premier script de comptage affichait
« 0 cas » alors que sa lecture de fichier avait échoué — `Set-Location` n'avait pas pris effet et
le compteur restait à zéro. Dans un contrôle, toujours **chemins absolus + `throw` si le fichier
manque + afficher le nombre de lignes lues** : un zéro rassurant doit être impossible à obtenir par
accident.

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

## Tâches vendues et document d'avancement (mis en place le 31/08/2026)

Le technicien appuie sur « 📋 Suivi », coche ses pylônes, et le relais Cloudflare dépose un JSON
dans `<chantier>/App Tech/Suivi/`. Chaque matin à 7h30, `suivi-chantiers/scripts/avancement-suivi.ps1`
(deuxième action de la tâche Windows « TELSAM - Veille documents RTE ») consolide tous les envois
d'un chantier en **un seul PDF `Avancement <numero>.pdf`** déposé dans son dossier App Tech — donc
lisible par les techniciens ET le bureau, via le lien de partage qui existe déjà. Aucun nouvel
accès à créer. Journal : `veille/avancement.log`.

**`tachesVendues` vit UNIQUEMENT dans `appli-techniciens/index.html`**, jamais dans le suivi.
Sans ce champ, le bouton Suivi n'apparaît pas (condition : `tachesVendues.taches.length` +
`documentsAppTech`). Structure :
```
tachesVendues: { devis: "TELSAM_CC_RTE_…", taches: [ { id, libelle, unite?, pylones?: [...] } ] }
```
Une tâche **avec** `pylones` donne des cases à cocher ; **sans**, un bouton Pas fait / En cours /
Fait. Au 31/08/2026 : 12 chantiers actifs renseignés, 65 tâches, 149 cases.

**RÉPONSE À UNE QUESTION DE PATRICE (02/09/2026) : le bouton Suivi n'apparaît JAMAIS tout seul.**
Il a demandé, après la création du dossier App Tech de Cross-Sausset, si l'absence de suivi était
normale et si « ça se mettrait en place une fois le chantier créé ». Non : le chantier existe,
`documentsAppTech` est rempli, **il manquait `tachesVendues`** — et ce champ est relevé à la main
dans le devis, il n'existe aucun automatisme. À ajouter dans le même geste que le dossier App Tech,
sinon le technicien n'a pas de bouton et personne ne s'en aperçoit. Le sous-dossier `Suivi` de
Dropbox, lui, est créé par le relais au premier envoi — il n'a pas à être préparé.

**`unite` — le mot juste quand le chantier ne se compte pas en pylônes** (ajouté le 02/09/2026).
Le mécanisme de cases sert à tout repère du devis, pas seulement aux pylônes : sur **26-002
Cross-Sausset**, première MTFO souterraine équipée, les cases sont les **liaisons** (`LS 1`,
`LS 2`). L'appli disait alors « Coche les pylônes que tu as faits » devant LS 1 / LS 2. Le champ
`unite` (`"liaison"`) corrige les deux phrases de la fenêtre ; **sans lui, on garde « pylône »**,
qui reste le cas courant — les 12 autres chantiers sont inchangés (non-régression vérifiée à
l'écran sur Bradascou).

### RÈGLE — quel devis fait foi (posée par Patrice le 31/08/2026)

**Toujours la DERNIÈRE version (V2, V3…), PLUS les devis TS (travaux supplémentaires) quand il y
en a.** Un chantier n'a presque jamais un seul devis, et les TS ne remplacent rien : ils s'ajoutent.
- Lister TOUS les devis avant d'en lire un : dossier `DEVIS`, mais aussi la racine du chantier et
  `Acte ST` — ils traînent aux trois endroits.
- **PIÈGE — une variante `_option` peut être celle qui fait foi.** Vécu sur **26-045
  St-Guillerme** : `26008` et `26008_option` coexistent ; c'est le second, devis
  **`TELSAM/CC/RTE/26008-1`**, qui fait foi, et il ajoute le **pylône 1** au raccordement. Le
  suffixe du nom de fichier laisse croire à une pièce secondaire. À ne pas confondre avec une
  **ligne** « Option : … » à l'intérieur d'un devis, qui elle ne se compte pas tant qu'elle n'est
  pas commandée.
- Plusieurs devis sans hiérarchie évidente (variante `option`, devis d'un autre client comme le
  « Data Hertz » de Lamativie 26-058) : **demander à Patrice**, ne pas trancher seul.

### RÈGLE CHANGÉE — le DOO/DOE/DEO EST une tâche vendue

Le 28/08/2026 il avait été retiré (« pas de leur ressort »). **Patrice a changé d'avis le
31/08/2026 : on le garde.** Présent sur 11 chantiers sur 12 ; seul Portet (26-051) n'en a pas au
devis. Sur l'audit 26-061 il s'appelle « Dossier de recettes de toutes les liaisons ».

### Comment modéliser un chantier

- **Pose et raccordement du boîtier = UNE seule tâche**, jamais deux, même quand le devis
  distingue les deux lignes avec des listes différentes (choix de Patrice, 31/08/2026).
- Le **demi-boîtier WTC2** reste une tâche à part (156 et 229 à Bradascou, P112 à Fleyriat) :
  c'est une autre prestation, pas un autre état.
- Un chantier qui compte des **quantités et pas des pylônes** (Portet : 84 raccordements,
  20 000 m de câble) passe en **boutons d'état**, quantités précisées au commentaire. Idem quand
  le devis ne donne aucun numéro de pylône (Campagnac-Séverac).

### Pièges techniques

- **Le tableau « DEVIS ESTIMATIF » est un classeur Excel INCRUSTÉ dans le `.doc`** :
  `Document.Content.Text` de Word ne le voit pas et rend un devis sans aucune prestation, sans
  erreur. **Lire le PDF**, pas le Word.
- Le document d'avancement **n'affiche ni pourcentage ni référence de devis** : il est lu par les
  techniciens, et la règle TELSAM interdit d'y faire figurer quoi que ce soit de commercial (même
  règle que pour les briefs).
- **Un document qui existe est toujours refait**, même quand il n'y a plus aucun envoi. Sinon,
  supprimer un envoi erroné laisserait en place un document qui continue de l'affirmer.
- **Un chantier `termine` n'annonce jamais de « reste à faire »** : le détail par tâche est retiré
  et le document dit « Chantier terminé ». Afficher « 0 / 3 » sur des travaux qui SONT faits serait
  faux.
- Pour corriger l'avancement, il suffit de **supprimer le JSON fautif** (dans `Suivi` ou
  `Suivi/_bruts`) : le document est recalculé de zéro au passage suivant.
- **UNION des pylônes faits, jamais addition** — même règle que les récaps de feuilles d'heures.

### Écarts de devis connus, laissés en l'état à la demande de Patrice

- **26-060 Givors** : 7 cantons facturés, 4 listés (il manque le tronçon p1 → 39N).
- **26-055 Fleyriat** : 11 cantons facturés, 10 listés.
- **26-065 DATA4** (terminé) : un second devis `25115` couvre le pylône HZ291, non repris dans ses
  tâches vendues.

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

**GARDE-FOU — les feuilles qui attendent sont signalées au démarrage** (mis en place le
01/09/2026). `scripts/veille-feuilles-heures.ps1`, appelé par le hook `SessionStart` juste après le
rapport de veille. Il dit deux choses :
- les feuilles **déposées mais pas encore intégrées** au classeur, et depuis combien de jours ;
- les feuilles **manquantes** — un technicien sans ligne dans une semaine que les autres ont déjà
  remplie (congé ou oubli). C'est ce second point qui manquait le plus : rien ne disait qu'on
  attendait encore quelqu'un.

**Pourquoi.** Le 01/09/2026, **douze feuilles de la S35 dormaient depuis trois jours**, déposées
entre le 28 et le 31/08. Le classeur n'avait pas bougé depuis le 25/08, donc le suivi affichait
**374 h pour septembre au lieu de 759**. Rien ne le signalait : l'étape 4 ci-dessus dépendait
entièrement de quelqu'un qui y pense. Exactement la faiblesse de `SEED_VERSION` et de la veille —
même remède, un mécanisme exécuté par l'outil.

**Détails de conception** : lecture seule (le classeur est ouvert en lecture seule, rien n'est
jamais écrit), ~4 s, **silence complet quand il n'y a rien à signaler**. Il rapproche le nom des
dépôts (`Prénom NOM - SNN …pdf`) des blocs de semaine du classeur de la période en cours, et
normalise les noms comme le script d'intégration (accents et séparateurs retirés). Il ouvre Excel
seulement s'il y a au moins un dépôt. Le fichier est en **pur ASCII** — PowerShell 5.1 lit un
`.ps1` sans BOM comme de l'ANSI, et des accents dans le code se casseraient en silence ; le nom du
dossier `dépôts appli` est donc retrouvé par motif (`^d.p.ts appli$`) au lieu d'être écrit en dur.
**Les deux branches ont été testées**, dont l'arrivée d'une feuille manquante simulée par un faux
dépôt au nom de Benjamin SOUPA — un test qui pouvait échouer, pas une vérification de façade.

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

**Ce paragraphe remplace l'ancien « à revoir quand le fichier normé existera » : le fichier normé
existe, c'est le bouton « 📋 Suivi » de l'appli.** La lecture des dossiers « Suivi » en texte libre
reste valable pour les chantiers qui n'ont pas de tâches vendues, mais elle n'est plus le canal
principal — voir la section suivante.

### Mise à jour automatique des boîtes posées (mise en place le 01/09/2026)

**Ce que déclare un technicien dans l'appli met à jour tout seul l'onglet Boîtes & nacelle.**
Demandé par Patrice le 01/09/2026, à partir du cas Bradascou : Pascal BONAVENTURE avait déclaré le
pylône 195 sur la tâche « Raccordement boîtier WTC2 » le 31/08, et l'onglet affichait toujours
0 posée sur 10.

**La chaîne, de bout en bout :** bouton 📋 Suivi → relais Cloudflare → JSON dans
`<chantier>\App Tech\Suivi\` → `scripts/boites-posees.ps1` (3e action de la tâche Windows
« TELSAM - Veille documents RTE », 7h30) → constante `POSES_APPLI` dans `suivi_chantiers_205.html`
→ colonne « Posées » de l'onglet, avec le pictogramme 📱 et, au survol, qui a déclaré quel pylône
et quand.

**Deux constantes, deux natures — ne pas les confondre :**
- `BOITES_TACHES` : **tenue à la main**, comme `REAL_DAYS`. Elle dit quelles tâches vendues valent
  « une boîte posée », lot par lot. Sans entrée, un chantier n'est jamais compté automatiquement.
  **À compléter à chaque nouveau chantier** qui a des boîtes au devis ET des tâches à pylônes,
  sinon ses déclarations resteront invisibles (le script le signale dans le journal : « ATTENTION
  … BOITES_TACHES n'a pas d'entrée pour ce chantier »).
- `POSES_APPLI` : **refaite de zéro à chaque passage** du script. Ne jamais y écrire à la main.
  Pour corriger une déclaration fausse : supprimer le JSON fautif dans `Suivi` ou `Suivi\_bruts`
  et relancer — même principe que le PDF d'avancement.

**Les quatre règles de comptage, dans l'ordre où elles protègent :**
1. **Toutes les tâches ne valent pas une boîte.** Chez Bradascou, `racc` (8 pylônes) + `demi`
   (2 pylônes) font les 10 du devis ; `canton`, `touret`, `recette`, `doo` n'ont rien à voir. C'est
   `BOITES_TACHES` qui tranche, jamais le libellé.
2. **Le pylône déclaré doit figurer dans la liste de pylônes DU LOT.** C'est ce qui empêche de
   compter dans le lot 1 une boîte du lot 2 (Chaineau, pylône 222 : raccordé par le lot 2, mais sa
   boîte est fournie par le lot 1) et ce qui neutralise **le piège St-Guillerme (26-045)** — le
   technicien peut y cocher des pylônes alors que le devis ne vend AUCUNE boîte (option non
   commandée) : lot sans pylônes ⇒ rien n'est compté.
3. **Union, jamais addition.** Deux techniciens qui déclarent le même pylône = une seule boîte ;
   le premier déclarant fait foi. Même règle que les feuilles d'heures et le document d'avancement.
4. **`posees` (saisie manuelle) et les déclarations de l'appli ne s'additionnent PAS** : le
   compteur retient **le plus grand des deux**. La saisie manuelle est un nombre sans détail des
   pylônes, les déclarations sont une liste — les additionner compterait deux fois une boîte déjà
   relevée à la main. Le maximum peut sous-compter, jamais sur-compter : c'est le bon sens d'erreur
   pour un chiffre qui sert à commander du matériel. Dès que les deux valeurs diffèrent, la vue
   affiche « (N à la main / M appli) » pour que l'écart se voie.

**Comptage automatique possible au 01/09/2026 : 9 lots sur 8 chantiers** (Bradascou 26-054,
Chaineau 26-036-1 et 26-036-2, Fleyriat 26-055 THYM + OPPC, Givors 26-060, Bissy 26-064,
DATA4 26-065, Hospitalet 26-027). Restent hors comptage automatique, volontairement : Campagnac
26-070 (8 boîtes mais aucun numéro de pylône au devis — tâches en boutons d'état), St-Guillerme
26-045, Portet 26-051, Audit 26-061 et Lamativie 26-058 (aucune boîte vendue).

**Le script ne touche PAS à `SEED_DATA`** : `POSES_APPLI` vit à côté, donc pas de `SEED_VERSION` à
bumper chaque matin et aucun effacement de l'état local des collègues. Il ne commite pas non plus —
le fichier modifié reste visible dans `git status`. Options utiles : `-Simuler` (montre ce qui
serait écrit sans toucher au fichier), `-Chantier 26-054`, `-Racine`/`-Suivi` pour tester sur une
copie. Journal : `veille/avancement.log`, lignes préfixées `[boites]`.

### LE NUMÉRO DU CHANTIER SE LIT DANS L'ENVOI, JAMAIS DANS LE NOM DU DOSSIER (02/09/2026)

**L'incident.** Un technicien remplit son suivi sur Fleyriat, et c'est LUI qui prévient Patrice :
aucune alerte n'était partie. Deux envois perdus (Younes MOUSSA le 01/09, Morad EL ABBASSI le
02/09). Patrice : « comment est-ce encore possible avec tout ce que nous avons mis en place ? »

**La cause.** `boites-posees.ps1` et `avancement-suivi.ps1` identifiaient le chantier en cherchant
un `26-0XX` dans le nom du **dossier parent** du dossier `Suivi` :
```
if ($d.Parent.Parent.Name -match '(\d{2}-\d{3})') { $numero = $Matches[1] }
if (-not $numero) { continue }        # <-- saut SILENCIEUX
```
Or le dossier App Tech de Fleyriat est `Ligne aérienne\Fibrage Fleyriat\App Tech` : **son parent ne
porte aucun numéro** (le numéro est sur un dossier frère, en dessous). Le comptage sautait donc ce
dossier sans un mot, et **le détecteur d'orphelines ne pouvait rien signaler puisque le chantier
n'entrait jamais dans son champ de vision.** Un `continue` sans trace est invisible par
construction. `avancement-suivi.ps1`, lui, retombait sur le nom du dossier et écrivait
« Avancement Fibrage Fleyriat.pdf » en annonçant « pas de tâches vendues » — un document qui ne
disait rien pour un chantier qui a pourtant ses tâches.

**Le pire est que l'information était là depuis le début** : chaque envoi contient
`"chantier":"26-055"`, écrit par l'appli, et le relais met même le numéro dans le nom du fichier.
Les scripts lisent désormais **le champ de l'envoi** ; le nom du dossier n'est qu'un repli, et un
envoi qu'on n'arrive pas à rattacher est **signalé**, jamais ignoré (`nonRattaches`).

**C'est le troisième exemplaire du même défaut en un jour** : `controle-chantiers.ps1` avait le
même le matin même (« aucun dossier App Tech » sur Fleyriat, qui en a un). **Corriger un défaut de
rapprochement dans un script oblige à vérifier les autres** : ils partagent tous cette hypothèse
fausse que le dossier parent porte le numéro.

### Les commentaires des techniciens remontent maintenant (02/09/2026)

**Deuxième moitié de l'incident, indépendante de la première.** Même pour les envois correctement
comptés, **le commentaire n'arrivait nulle part** — il ne finissait que dans le PDF déposé dans
App Tech, que le bureau n'ouvre pas. Or c'est souvent là qu'est l'essentiel. Ce jour-là :
- Morad EL ABBASSI, Fleyriat : « Ts : pyl 19 a 48 sur demande du client une mesure du câble car il
  pensé à une cassure » — **une demande de travaux supplémentaires du client**, rien de coché ;
- Pascal BONAVENTURE, Bradascou : « Pylône 184 devient pylône 46 » — **une renumérotation**, qui
  invalide la liste de pylônes de la fiche.

`boites-posees.ps1` écrit donc `veille/suivis.json` (tous les envois, faits ET commentaires), lu
par le récap et par le hook de démarrage :
- un envoi de moins de 24 h va dans **« Fait »** du récap ;
- un envoi **avec commentaire** va AUSSI dans **« reste à faire »**, et y reste **7 jours** : une
  phrase du terrain demande presque toujours une suite (chiffrer un TS, corriger un numéro,
  rappeler le client), et elle ne doit pas disparaître au bout d'un jour ;
- un envoi non rattaché passe en **rouge**.

**Un envoi sans rien de coché n'est pas un envoi vide.** Celui de Morad ne cochait aucune case et
portait la demande la plus importante de la journée. Ne jamais filtrer sur « faits non vides ».

**Un commentaire arbitré doit CESSER de revenir — `scripts/marquer-suivi-traite.ps1`.**
Le rappel sur 7 jours était nécessaire, mais sans moyen de l'éteindre il aurait redit pendant une
semaine ce qui était déjà décidé, et le récap aurait perdu sa crédibilité comme tous les contrôles
bruyants. La **note est obligatoire** (même principe que `marquer-veille-traitee.ps1`) : elle dit
ce qui a été décidé et sert de trace dans `veille/suivis-traites.json`. `-Lister` montre ce qui
attend, `-Tous` solde tout d'un coup. **Le fait reste visible dans « Fait »** du récap : seul le
rappel « à traiter » s'éteint, on n'efface pas l'événement.

**Et une décision de Patrice ne vit PAS dans ce fichier de marquage : elle va dans `aVerifier` de
la fiche.** Le marquage éteint un rappel, il ne conserve rien de consultable. Les deux arbitrages
du 02/09/2026 sont donc inscrits dans l'onglet « À vérifier » :
- **26-055 Fleyriat** — le TS demandé par le client (mesure de câble pylônes 19 à 48) **se fera et
  sera facturé plus tard** : ne pas le chiffrer maintenant, ne pas l'oublier à la facturation ;
- **26-054 Bradascou** — « le pylône 184 devient le 46 » : **ne rien modifier**, ni la liste de
  pylônes, ni les tâches vendues, ni le lot de boîtes. La mise à jour se fera **dans le DOO une
  fois les travaux terminés**. Ne pas « corriger » cette numérotation entre-temps.

**Les poses que le comptage n'a pas su rattacher (« orphelines ») sortent à DEUX endroits.**
Ajouté le 01/09/2026 après une remarque de Patrice : le script écrivait déjà cet avertissement dans
`veille/avancement.log`, et il a demandé « de quel journal parles-tu ? » — **un avertissement rangé
dans un fichier que personne n'ouvre n'avertit personne.** Donc, désormais :
- **bandeau rouge en haut de l'onglet Boîtes & nacelle** (constante `POSES_ORPHELINES`, écrite par
  le même script), qui dit explicitement que le reste à poser affiché est trop élevé ;
- **bloc injecté au démarrage de session** via `scripts/veille-boites-orphelines.ps1`, appelé par
  `hook-veille-session.ps1` — à traiter en session comme la veille documentaire.

Deux causes possibles, distinguées dans le message : le chantier n'a pas d'entrée dans
`BOITES_TACHES`, ou le pylône déclaré ne figure dans la liste d'aucun lot (erreur de numéro du
technicien, ou liste du devis incomplète dans la fiche).

**Deux pièges rencontrés en l'écrivant**, à ne pas réintroduire :
- la ligne `const SEED_DATA` du **suivi** se termine par `;;` (double point-virgule, sans effet en
  JavaScript, fatal pour `ConvertFrom-Json`) — d'où le `TrimEnd(';')` et non un `Substring` ;
- en .NET, `.` avale le `\r` et `$` se place ENTRE le `\r` et le `\n` : un `Regex.Replace` sur
  `^const POSES_APPLI = .*$` mange le retour chariot et laisse une fin de ligne LF isolée au milieu
  d'un fichier en CRLF. Motif à utiliser : `^const POSES_APPLI = [^\r\n]*`, puis découpe/recolle
  (un `Replace` réinterpréterait un `$1` présent dans un nom).

**État au 26/08/2026** : 44 fiches renseignées sur 45 chantiers actifs, **167 boîtiers WTC2**,
2 boîtiers OPPC (Fleyriat) et 4 boîtiers en chambre restant à poser, 29 déjà posés. Reste une seule fiche à
compléter, 26-021 Gampaloup-Valence, parce qu'aucun devis n'est encore arrivé dessus.
**Au 01/09/2026** : 174 WTC2 restants, 30 posés dont **1 déclaré par un technicien** (Bradascou,
pylône 195, Pascal BONAVENTURE le 31/08).

## La matinée automatique — `scripts/matin.ps1` (refondu le 02/09/2026)

**L'incident qui a tout déclenché.** Le 02/09 au matin, Patrice branche son PC : pas de mail, pas
de planning à jour, et l'impression que « tout ce que l'on fait ne sert à rien ». Il avait raison.
Ce matin-là, sur les cinq actions de la tâche Windows : la veille est passée, **le document
d'avancement et le comptage des boîtes n'ont rien fait**, **la lecture du planning Teams s'est
bloquée sur Excel** au point que Windows a tué la tâche (`LastTaskResult = 0xC000013A`), et
**le mail est resté dans la boîte d'envoi** d'un Outlook démarré sans fenêtre.

**Mais le vrai défaut n'est aucun de ces quatre-là : c'est que RIEN NE L'A DIT.** Le récap
affichait « Scripts du matin passés » — parce qu'il regardait la *date du fichier journal*, touché
par la dernière étape. Un système qui se déclare en bonne santé quand il est en panne est pire que
pas de système : il consomme la confiance de celui qui s'y fie.

**La refonte, quatre règles :**
1. **Un chef d'orchestre.** La tâche n'a plus qu'une action, `matin.ps1`, qui enchaîne les étapes.
2. **Chaque étape a un délai maximum et tourne dans son propre processus.** Une étape bloquée est
   tuée, **les suivantes se font quand même**. Avant, Excel restait pendu et Windows finissait par
   tuer la tâche entière — donc aussi tout ce qui n'avait pas encore tourné.
3. **Chaque étape écrit son état** (`ok` / `echec` / `expire` / `ignore`, message, durée) dans
   `veille/matin.json`. **Le récap lit ce fichier** et fait une ligne rouge nommée par étape en
   défaut. **Un silence ne doit jamais pouvoir se lire comme un succès.**
4. **On attend Dropbox** (jusqu'à 10 min) avant les étapes qui en dépendent, au lieu d'échouer en
   une fraction de seconde. La tâche se déclenche à l'ouverture de session, quand Dropbox monte
   encore.

**Ménage Office — deux garde-fous.** `matin.ps1` ferme les Excel/Word restés **sans fenêtre**,
mais **uniquement ceux démarrés pendant son propre passage**, et **ne touche jamais à Outlook** :
un Outlook réduit dans la zone de notification peut n'avoir aucun titre de fenêtre, le tuer
fermerait la messagerie de Patrice et retiendrait son courrier.

**L'envoi du mail est sorti de la matinée** — tâche séparée « TELSAM - Envoi du recap », qui
repasse **toutes les 10 minutes de 7h30 à 13h00**. `recap-mail.ps1` **n'a plus le droit de démarrer
Outlook** : il n'envoie que si Outlook tourne DÉJÀ **avec une fenêtre** (preuve qu'il est vraiment
connecté), sinon il ne fait rien et retente. Et si Outlook disparaît juste après le `Send()`
(vécu : « Le serveur RPC n'est pas disponible »), **la journée n'est pas marquée comme envoyée** :
mieux vaut un doublon, qui se supprime en deux secondes, qu'un récap manquant que personne ne voit.

**Bug qui aurait tout figé, corrigé le même jour** : dans `marquer-veille-traitee.ps1`,
`@($x) | Where-Object {...}` rend un PSObject et non un tableau quand un seul élément passe le
filtre, et le `+=` suivant échouait sur `op_Addition`. Déclenché le jour où l'historique est tombé
à un seul élément : **plus aucune veille ne pouvait être marquée traitée**, donc le rappel serait
resté allumé indéfiniment. Le `@()` doit envelopper le RÉSULTAT du filtre.



### Deux passages par jour, et la fin de la tâche qui clignotait (03/09/2026)

**Trois choses tranchées par Patrice le 03/09/2026 au soir.**

**1. Le débat sur la publication est CLOS, et c'est lui qui l'a fermé.** J'avais soulevé que la
relecture du matin n'était pas poussée, donc pas vue par les techniciens. Sa réponse : « le planning
ne changera que rarement. Une fois que nous validons le planning le vendredi, il ne bouge pas, ou
alors c'est exceptionnel, et dans ce cas-là nous préviendrons nous-mêmes les techniciens et nous
modifierons ensuite le planning. **Il n'y a pas de débat là-dessus.** » **Ne pas re-proposer une
publication automatique.** Et il a raison sur le fond : la validation d'une semaine passe par la
base, elle est donc vue **tout de suite** sans aucun push — c'est seulement le CONTENU du planning
qui attend une publication, et il ne bouge presque pas.

**2. La tâche « TELSAM - Envoi du recap » est DÉSACTIVÉE.** Elle repassait toutes les 10 minutes de
7h30 à 13h00 : « j'ai des fois l'écran de Visual Planning, ou de Word, ou de Microsoft, un écran
noir qui s'ouvre et qui se referme automatiquement. Ça m'a fait ça toute la matinée. » C'était bien
elle — 34 lancements de `powershell.exe` dans la matinée, chacun pouvant faire clignoter une console
même en `-WindowStyle Hidden`.

**3. La chaîne tourne maintenant DEUX fois par jour : 7h30 et 13h00.** Sa demande : « qu'il y ait la
veille du matin et la veille de treize heures, pour que tout se mette à jour deux fois par jour. »
Fait en ajoutant un **second déclencheur à la tâche existante** (« TELSAM - Veille documents RTE »),
pas en créant une seconde tâche : une seule chaîne, un seul endroit où regarder quand ça ne tourne
pas. Vérifié après coup que le rattrapage (`StartWhenAvailable`) et le compte
(`LogonType Interactive`, aucun mot de passe stocké) ont survécu à la modification.

**Ce que ce second passage répare, en plus de ce qu'il a demandé** : le trou connu du 31/08/2026 —
la veille prenait une photo à 7h30, et deux PGO arrivés à 8h55 plus une IST à 10h34 restaient
invisibles jusqu'au lendemain. Avec le passage de 13h, ils remontent le jour même.

**Effet à connaître** : le passage de 13h ouvre Excel (lecture du planning SharePoint), donc **une**
fenêtre peut clignoter à 13h — au lieu de 34 dans la matinée.

**L'archive du récap porte désormais l'heure** (`RECAP-2026-09-03-1634.html`). Sans ça le passage de
13h **écrasait** l'archive du matin, alors que cette archive n'existe que pour comparer : on aurait
perdu en silence la moitié de ce qu'elle sert à garder. Vérifié : l'ancienne
`RECAP-2026-09-03.html` du matin est intacte à côté de la nouvelle.


### La bulle Windows remplace le mail — `scripts/recap-bulle.ps1` (03/09/2026)

Patrice a vu la bulle d'essai et a demandé de la mettre en place. Étape
**« Bulle de notification »** de `matin.ps1`, donc à 7h30 **et** à 13h00.
WinRT `Windows.UI.Notifications`, AppId de PowerShell, **aucun module à installer**.

**TROIS DÉFAUTS TROUVÉS EN LA CONSTRUISANT, tous par l'essai et non par relecture :**

1. **Elle annonçait des chiffres FAUX.** Première version : une expression régulière sur
   `RECAP.html`, qui a sorti « 0 fait, 26 à faire » quand le récap disait « 13 fait(s),
   31 reste(s) ». **Des chiffres faux annoncés avec aplomb sont pires que pas de chiffres.**
   `recap-matin.ps1` écrit désormais `veille/recap-compteurs.json` — les deux nombres sortent de
   celui qui les calcule, ou de nulle part. Ne pas revenir à une relecture de la page.
2. **Les accents étaient cassés** : « Â« Recap Â» ». Le `.ps1` était sans BOM, donc lu en ANSI par
   PowerShell 5.1 — le piège déjà documenté, repayé une fois de plus. Le fichier porte maintenant un
   BOM, **et** son message a été réécrit sans guillemets typographiques : du texte qui ne peut pas
   se casser, quel que soit l'outil qui réécrira le fichier un jour.
3. **Son statut n'aurait été signalé nulle part.** Posée après l'écriture de `matin.json`, son échec
   n'entrait dans aucun fichier — exactement le défaut que `matin.json` existe pour empêcher. L'état
   des étapes est donc écrit **tout à la fin**, après la dernière étape.

**Deux règles de placement, à ne pas défaire :**
- elle vient **après** le bloc qui écrit la page de remplacement. Si le récap a échoué, c'est cette
  page-là qui est du jour, et la bulle annonce alors une page qui dit franchement que la chaîne a
  échoué — au lieu de se taire, ce qui laisserait croire que rien n'a tourné ;
- **elle ne doit jamais faire échouer la chaîne** : délai de 2 minutes, toute erreur tracée et
  avalée, sortie en 0. Une bulle est un confort ; la veille et le planning sont le travail.

**Elle refuse de parler d'une page périmée** (même règle que celle posée pour le mail le
02/09/2026) : si `RECAP.html` n'est pas du jour, aucune bulle. Vérifié avec ses deux
contre-exemples — page antidatée d'un jour, et page absente.

**Essai réel de la chaîne complète le 03/09/2026** : 101 s, 7 étapes sur 7 en `ok`, bulle comprise,
et son statut bien présent dans `matin.json`.


**LE HOOK D'`APP_VERSION` IGNORE LES DEUX LIGNES ÉCRITES PAR LES SCRIPTS** (03/09/2026).
Depuis que la chaîne tourne deux fois par jour, elle réécrit `PLANNING_TECH` et
`AVANCEMENT_DECLARE` dans `index.html` à chaque passage — souvent pour le seul horodatage. Le hook
a bloqué le premier commit qui suivait, et il avait raison de le faire : il ne savait pas
distinguer. Mais le corriger n'était pas facultatif, car les deux issues étaient mauvaises :

- bumper `APP_VERSION` chaque jour ⇒ le technicien se prend l'écran d'ouverture **tous les matins**.
  La règle est écrite plus haut : « insupportable en trois jours, et le signal perdrait tout son
  sens à force d'être vu ». On aurait détruit exactement ce que ce hook protège ;
- ne pas le corriger ⇒ tout commit suivant un passage du matin est bloqué.

Le hook **retire ces deux lignes du diff et regarde s'il reste quelque chose**. Ne pas remplacer ce
test par « si le diff contient ces lignes, on passe » : un vrai changement de code passerait alors
en douce dès qu'un passage du matin l'accompagne. **Vérifié dans les trois sens** : lignes générées
seules ⇒ passe ; vrai changement sans bump ⇒ bloque ; vrai changement avec bump et nouveauté ⇒ passe.

Même raisonnement que `POSES_APPLI` dans le suivi : une donnée regénérée vit à part justement pour
ne pas déclencher de bump de version.

### On n'emprunte que le classeur SharePoint, jamais une copie locale (03/09/2026)

Question de Patrice sur le passage de 13h : « j'espère que ce n'est pas pour lire le planning qui
est dans mon bureau, mais bien celui sur Teams, qui est toujours à jour. »

**Réponse vérifiée, pas de mémoire** : `planning-rte.json` porte `teams = true`, la source est
l'adresse SharePoint, la copie du Bureau **n'existe plus**, et le classeur ouvert dans son Excel
pointe bien sur l'URL. Les trois voies restent : fichier Teams en direct, sinon copie figée du
classeur qu'il a ouvert, sinon (et seulement alors) la copie du Bureau avec bandeau rouge.

**Mais sa question a mis le doigt sur un vrai trou, désormais fermé.** La voie 2 ne testait que le
**NOM** du classeur ouvert — et la copie du Bureau porte **exactement le même nom**. Le jour où il
en rouvre une, on aurait lu une donnée périmée alors que le fichier partagé était joignable. Le test
porte maintenant sur `FullName -like 'http*'` : un classeur du même nom mais local est **ignoré avec
un message**, et on va chercher Teams.

### POURQUOI LE MAIL DU RÉCAP N'ARRIVERA JAMAIS AINSI — Patrice utilise le NOUVEL Outlook

Patrice : « je suis bien allé chercher mon rappel du matin, mais je n'ai jamais reçu le mail. »
Diagnostiqué, pas supposé. Le journal montre **34 tentatives, 34 fois** « Outlook n'est pas ouvert :
rien envoyé ». Et la détection était JUSTE : aucun processus `OUTLOOK.EXE` ne tournait.

**La cause : il lit son courrier dans le NOUVEL Outlook pour Windows — `olk.exe`**
(paquet `Microsoft.OutlookForWindows`, vérifié en liste de processus). Le nouvel Outlook **n'a
aucune interface COM** : `Outlook.Application` ne peut pas lui parler, ni maintenant ni plus tard.
L'Outlook classique est bien installé (`Office16\OUTLOOK.EXE`) mais **jamais lancé** — et le lancer
par COM est le cas déjà écarté le 02/09/2026 : il démarre sans fenêtre, ne se connecte pas, et le
message reste dans la boîte d'envoi.

**Donc `recap-mail.ps1` est un cul-de-sac.** Le script reste dans le dépôt (il marcherait pour
quelqu'un sous Outlook classique) mais **sa tâche est désactivée**. Ne pas essayer de le
« réparer » : le problème n'est pas dans le script.

**Ce qui est réellement faisable, mesuré :**

| voie | verdict |
| --- | --- |
| **Rien de plus** — il ouvre le raccourci du Bureau, ce qu'il fait déjà | marche, zéro pièce en plus |
| **Une bulle de notification Windows** à 7h30 et 13h00 | **RETENUE ET EN PLACE** : Patrice l'a vue, elle est l'étape « Bulle de notification » de matin.ps1 (voir ci-dessus) |
| Mail via un service externe (Resend/SendGrid) branché sur le relais Cloudflare | demande un compte et une clé de plus, et un mail venu d'un domaine extérieur vers telsam.com risque le courrier indésirable. Disproportionné |
| Message Teams par Graph | demande une inscription d'application côté M365. Pas simple, et il vit déjà dans Teams — à ne creuser que s'il le demande |

Le code d'une bulle, pour mémoire : WinRT `Windows.UI.Notifications`, `AppId`
`{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe`. Aucun module à
installer.

### LE TROU DE LA CHAÎNE : ce qui est relu chaque matin n'est PAS publié (relevé le 03/09/2026)

**Trouvé en auditant les règles à la demande de Patrice**, pas en le supposant : la tâche de 7h30
lance bien `planning-rte.ps1 -Injecter` et `boites-posees.ps1`, qui réécrivent
`suivi_chantiers_205.html` **et** `appli-techniciens/index.html`. Mais **aucun script ne commite ni
ne pousse**. Or :

- le suivi en ligne est construit par Cloudflare **depuis GitHub** ;
- l'appli technicien est servie par GitHub Pages **depuis GitHub**.

**Donc, tant que personne ne pousse, la relecture du matin ne sort pas de la machine de Patrice.**
Les techniciens gardent le planning et les cases cochées de la dernière publication, et **rien ne
le dit**. C'est le défaut de forme qui s'est répété toute la journée du 03/09 : un mécanisme dont
le dernier maillon est humain et silencieux.

**Ce n'est pas un oubli à corriger sans réfléchir.** Une publication automatique enverrait la
lecture du planning aux 13 techniciens sans qu'aucun humain l'ait regardée — et le 03/09 la lecture
a produit une attribution FAUSSE (Cantegrit au lieu de Rion des Landes) que seul l'œil de Patrice a
rattrapée. Il y a donc un vrai argument pour garder quelqu'un dans la boucle.

**TRANCHÉ PAR PATRICE LE 03/09/2026, LE MÊME SOIR : ON NE PUBLIE PAS AUTOMATIQUEMENT.** « Le planning
ne changera que rarement… il n'y a pas de débat là-dessus. » Voir la section « Deux passages par
jour » ci-dessus. Les deux voies étudiées, gardées pour mémoire :
1. **Publication automatique** en fin de chaîne, limitée aux lignes générées (`PLANNING_RTE`,
   `PLANNING_TECH`, `POSES_APPLI`, `POSES_ORPHELINES`, `AVANCEMENT_DECLARE`) et à elles seules —
   jamais un `git add -A`, qui emporterait un travail en cours.
2. **Un rappel en tête du récap du matin** : « le planning a changé ce matin, il n'est pas encore
   publié aux techniciens », avec le nombre de lignes en attente.

Tant que ce n'est pas tranché, **penser à pousser après toute session qui touche au planning ou aux
déclarations** — et le dire à Patrice, au lieu de le laisser croire que les techniciens ont la
dernière version.

## Récap du matin — `veille/RECAP.html` (mis en place le 01/09/2026)

**Une page par jour : ce qui a été fait, ce qui reste à faire.** Demandée par Patrice le
01/09/2026, et sa raison d'être est écrite dans sa phrase : « cela me permettra de constater s'il y
a des erreurs ou non dans le traitement de nos infos ». **Ce n'est pas un tableau de bord de plus,
c'est son moyen de contrôler le travail automatique et le mien.** Deux conséquences directes :
ne jamais y présenter comme traité quelque chose qui ne l'est pas, et toujours dire d'où vient
chaque ligne.

`scripts/recap-matin.ps1`, **4e action** de la tâche Windows de 7h30. Il ne recalcule rien : il lit
`veille/dernier.json` + `traitement.json`, les fiches du suivi, `POSES_APPLI`/`POSES_ORPHELINES`,
et `git log` des deux dépôts. Écrit `veille/RECAP.html` (toujours au même endroit) **et** une copie
datée dans `veille/recaps/` — c'est l'archive qui permet de comparer d'un jour sur l'autre.
Un raccourci « Recap TELSAM du matin » est sur le Bureau de Patrice.

**LE POINT QUI DÉCIDE DE TOUT : le tri du bruit.** Au 01/09/2026, les 46 chantiers actifs portent
**42 alertes orange**. La première version listait tout : 55 lignes, illisible, et le vrai problème
noyé. Règles appliquées, à ne pas défaire :
- **toutes les alertes rouges passent**, toujours ;
- **une alerte orange ne passe que si un technicien est placé sur ce chantier dans les 15 jours**
  (via `REAL_DAYS`) — un point de vigilance sur un chantier où personne ne va peut attendre
  l'onglet Alertes ;
- **le reste est compté en une ligne**, jamais supprimé silencieusement : « N points de vigilance
  sur des chantiers où personne n'est placé ». Cacher sans le dire serait pire que tout lister.
Résultat : 5 faits / 20 restes le premier jour, au lieu de 22 / 55.

**Un script qui n'a pas tourné est une information, pas un silence.** Si `avancement.log` n'a pas
bougé depuis 24 h, le récap l'écrit en rouge — sans ça, une page vide se lirait comme « rien à
signaler » alors qu'elle veut dire « rien ne fonctionne ».

### Fiabilité du récap : quatre trous bouchés le 02/09/2026

Patrice, après une journée où le récap l'avait trompé : **« le récap du matin me suffit, par contre
je veux qu'il soit fiable ».** Décision claire : **on n'ajoute plus rien à la chaîne** (le relevé
des suivis en cours de journée a été proposé et refusé) — on rend fiable ce qui existe. Audit de la
chaîne elle-même, et non de ce qu'elle rapporte :

1. **Si l'étape du récap échoue, la page ne doit PAS rester celle de la veille.** C'était le trou le
   plus grave : le raccourci du Bureau s'ouvrait normalement sur la page d'hier, titre daté d'hier,
   et rien ne disait qu'elle était périmée. `matin.ps1` écrit maintenant, à la place, **une page
   rouge courte** qui dit ce qui s'est passé et donne l'état de chaque étape. Elle se déclenche
   aussi quand l'étape se dit OK mais que la page n'a pas été réécrite ce jour-là.
2. **Le mail refuse de partir avec une page qui n'est pas du jour** (`recap-mail.ps1` compare la
   date du fichier). Sans ça il aurait envoyé le contenu d'hier sous un objet portant la date
   d'aujourd'hui et les compteurs d'hier — **un mail qui se trompe de jour est pire que pas de
   mail : il donne l'impression que tout a tourné.**
3. **`dernier.json` absent** ne fait plus disparaître la ligne « veille » en silence : ligne rouge
   « impossible de dire si un PGO est arrivé ».
4. **Un `git log` muet** ne rend plus « 0 modification » (qui se lit « rien n'a bougé ») mais une
   ligne disant qu'on n'a pas pu regarder.

**Les deux premiers ont été vérifiés en cassant réellement la chaîne** : page antidatée → le mail
refuse ; `recap-matin.ps1` renommé → la page de remplacement s'écrit. Un test qui ne peut pas
échouer ne prouve rien.

**Le trou qui RESTE, connu et assumé par Patrice** : un suivi envoyé par un technicien n'est vu
qu'au passage du matin. Celui de Morad, arrivé à 11h33 le 02/09, n'aurait remonté que le lendemain.
Un relevé toutes les dix minutes a été proposé (la tâche « TELSAM - Envoi du recap » existe déjà) ;
**Patrice l'a refusé** — le récap du matin lui suffit. Ne pas le remettre sur la table.

**L'envoi par mail — `scripts/recap-mail.ps1`, 5e action de la tâche (activé le 01/09/2026).**
Écrire un fichier local ne demande pas d'autorisation ; envoyer un mail au nom de Patrice, si.
Il l'a demandé explicitement (« mets le mail en place ») et **le seul destinataire est lui-même** :
`patrice.pivot@telsam.com`. **Ne jamais élargir la liste des destinataires sans une nouvelle
demande de sa part** — un récap contient des alertes internes et des écarts de devis.

- **Pourquoi Outlook COM et pas un envoi SMTP** : telsam.com est une boîte Microsoft 365,
  l'authentification simple y est désactivée, et un mot de passe dans un script serait à proscrire
  de toute façon. Outlook est déjà installé et déjà authentifié : aucun secret nulle part.
- **PIÈGE PRINCIPAL, vérifié le 01/09/2026 : un Outlook démarré par COM tourne SANS FENÊTRE et ne
  se connecte pas au compte Exchange.** `Send()` ne fait que déposer dans la boîte d'envoi, et
  **`SendAndReceive()` ne suffit PAS à l'en sortir** — testé, la boîte d'envoi contenait toujours
  1 message après l'appel, et les Éléments envoyés dataient encore du 08/06. J'avais entre-temps
  annoncé à Patrice que le mail était parti : il ne l'était pas. **Ne jamais conclure d'un `Send()`
  sans exception que le message est parti** ; la seule preuve est la boîte d'envoi vide (lecture
  lente, plusieurs minutes, quand Outlook est en mode sans fenêtre).
- **Choix retenu : on dépose quand même**, plutôt que de forcer une fenêtre Outlook à l'écran.
  Le message part dès que Patrice ouvre Outlook, ce qu'il fait tous les matins — et de toute façon
  la tâche ne tourne qu'à l'ouverture de sa session. Le journal distingue les deux cas, et le
  script **ne ferme jamais Outlook** : le fermer retiendrait le message.
- **Un seul envoi par jour** (`veille/recap-mail.json`) : la tâche se rattrape à l'ouverture de
  session et peut se déclencher plusieurs fois.
- **Un échec d'envoi n'arrête rien** : il est tracé dans `avancement.log` avec le préfixe `[mail]`,
  et la page HTML existe de toute façon. Options : `-Simuler`, `-Forcer`, `-Destinataire`.
- **Les compteurs sont dans l'objet du mail** (« 5 fait(s), 20 à faire ») : c'est ce qui se lit sur
  un téléphone sans ouvrir le message.

**À savoir sur l'heure d'arrivée.** La tâche est en `LogonType Interactive` : elle ne tourne QUE
quand la session Windows de Patrice est ouverte, et se rattrape à l'ouverture — le 31/08 elle est
passée à 08h10, le 01/09 à 08h14, jamais à 07h30. **Le mail arrive donc quand Patrice s'assied, pas
avant son arrivée.** Pour qu'il parte vraiment à 7h30 il faudrait enregistrer son mot de passe
Windows dans la tâche planifiée : à ne pas faire, et à ne pas proposer.

**Piège PowerShell rencontré ici (le même que dans `avancement-suivi.ps1`) : `@($null).Count`
vaut 1.** Une liste d'orphelines vide ressortait comme un élément fantôme, et le récap affichait
une ligne rouge sans texte. Filtrer explicitement (`Where-Object { $_ -and $_.numero }`).

**Encodage** : ce script-ci est en **UTF-8 AVEC BOM**, contrairement à `avancement-suivi.ps1` qui
est en pur ASCII avec des accents composés (`[char]0x00E9`). Les deux marchent ; le BOM est
préférable pour un script à texte dense. Ce qui casse, c'est un `.ps1` **sans BOM** contenant des
accents écrits directement.

## Temps prévisionnel par chantier — onglet Heures (mis en place le 27/08/2026)

L'onglet **Heures** de `suivi_chantiers_205.html` affiche, par chantier, **deux jauges** : le temps
**prévu** (calculé depuis le devis) et le temps **réalisé** (heures cumulées des feuilles d'heures,
champ `heuresTelsam`). But : voir d'un coup d'œil où en est chaque chantier par rapport à son budget
temps, et se remplir au fur et à mesure que les techniciens rentrent leurs heures.

**RÈGLE — à chaque chantier qui a (ou prend) des heures réelles, calculer son prévisionnel depuis
le devis, avec la compétence `temps-previsionnel-chantier`.** Champ `heuresPrevues` sur la fiche :
```
heuresPrevues: { heures, jours, devis, montantHT, materielAchat, note }
```
`heures = 0` tant que ce n'est pas calculé → la jauge affiche « prévu à calculer ». La `note` est
affichée en clair sous le nom du chantier (devis, jours, matériel déduit) : la renseigner utilement.

**MÉTHODE (financière, v2, décidée par Patrice le 27/08/2026) — c'est le `SKILL.md` du skill qui
fait foi, PAS la description courte.** Le résumé du skill parle encore d'une ancienne méthode « somme
des durées » : **obsolète**. La vraie méthode :
```
heures prévues = (montant total HT du devis − matériel au prix d'ACHAT) ÷ 90 €/h ; jours = heures ÷ 8
```
- Le **matériel** se déduit au **prix d'achat** du fichier `reference/bordereau_materiel_achat.csv`
  du skill (WTC2 D.01 = 800 €, boîtier chambre D.06 = 200 €, tore PMC C.16 = 1500 €, câbles COR/COSM
  D.7/D.8 = 2 €/ML — au mètre, à multiplier par le métrage —, nacelles H.13-15, etc.), ligne de
  devis par ligne de devis. Le reste (préparation, mesures, raccordement, recette, DOO/DOE) = main
  d'œuvre convertie en heures.
- Le script `build_previsionnel_financier.py` du skill exige **Python (absent de la machine)** :
  calcul **à la main**. Extraction du texte des devis PDF via **Word COM** (`pdf2txt` : Open en
  lecture seule puis `.Content.Text`) — le texte sort en un seul bloc, relever le total HT et les
  lignes « Fourniture… ».

**Pièges rencontrés le 27/08/2026, à connaître :**
- **Le prévu et le réalisé ne mesurent pas la même chose.** `heuresTelsam` additionne les heures de
  **chaque** technicien (2 techniciens un jour = 14 h) ; le prévu, non. Dépasser 100 % est donc
  **normal** quand l'équipe est nombreuse — la jauge passe en **orange** (`var(--amber-border)`,
  libellé « — dépassé »), ce n'est PAS une faute. Ne pas « corriger » ce dépassement.
- **Couleurs : n'utiliser que des variables définies dans le suivi.** Bug vécu : `var(--orange)` et
  `var(--gray)` **n'existent pas** ici → barre sans couleur, invisible (paraissait vide même à
  192 %). Utiliser `--amber-border` / `--gray-border`. Vérifier `getComputedStyle` sur une vraie
  barre, pas seulement le style écrit.
- **Fourchette haute** : là où une fourniture du devis n'a **pas** d'article correspondant au
  bordereau (surtout devis arteria/poste : tiroirs, cordons, pigtails, châssis), le skill la compte
  en **main d'œuvre** (comportement voulu) → prévu surestimé. Le noter dans `note`.
- **Anomalie du bordereau à faire corriger par Patrice** : l'article **brides C.13** y est à
  **1500 €** d'achat alors que le devis le vend **250 €** (prix d'achat > prix de vente,
  incohérent). **Non déduit** sur 26-019 et 26-041, signalé. Idem vérifier D.05 (350 € vs 275 €).
- **Remises** : certains devis portent une remise de 5 % (Cantegrit 26-003, Arcomie 26-035, Lisieux
  26-039) — prendre le **montant net** (après remise).
- **26-054 Bradascou** : le devis dit « **Pose** » et non « Fourniture et pose » (même prix 1175 €) —
  matériel supposé inclus et déduit, ambiguïté déjà notée dans « À vérifier ».
- **26-017 Chafauds-Courelles-Beaugency-Lestiou-Gribouzy : laissé de côté** (27/08/2026, décision de
  Patrice) — plusieurs devis possibles, lequel fait foi à confirmer avant de chiffrer.

**JAMAIS DE REMPLACEMENT GLOBAL DANS `SEED_DATA`, MÊME POUR UN SEUL CARACTÈRE.** Commis le
03/09/2026 en montant l'indice PGO de Rion des Landes : un `.Replace('(PGO ind.12)', '(PGO ind.13)')`
sur tout le fichier a **aussi modifié la fiche 26-053 Le Havre**, qui portait exactement le même
texte pour SON propre PGO — elle annonçait donc un indice qu'elle n'a jamais eu. Rattrapé parce que
le compteur affichait « 2 occurrences » là où un seul chantier était concerné : **toujours compter
les occurrences et s'arrêter si le nombre ne correspond pas à ce qu'on croit modifier.** Le contrôle
qui tranche est de reparser `SEED_DATA` et de lister les fiches portant la nouvelle valeur.
Réparation : réancrer sur un texte **unique** à la fiche (`'Havre-Rougemontier (PGO ind.13)'`),
jamais sur le fragment commun.

**Écriture dans `SEED_DATA`** : insertion ciblée par fiche (isoler la fiche via `"numero":"26-0XX"`
puis `LastIndexOf('{"id":"c_')`, remplacer `"heuresTelsam":` par `"heuresPrevues":{…},"heuresTelsam":`
DANS cette fiche seulement), lecture/écriture `[IO.File]` en UTF-8 sans BOM, **bump `SEED_VERSION`**,
reparser le JSON en contrôle. État au 27/08/2026 : **34 fiches sur 35** chiffrées (toutes celles qui
ont des heures réelles sauf 26-017).

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

**RÈGLE — TRAITER le rapport de veille dès le début de session, pas seulement le lire.** Manqué le
28/08/2026, signalé fermement par Patrice (« tu es censé vérifier les PGO/PDP dès la discussion du
matin »). Le rapport arrive dans le contexte au démarrage ; il faut alors, **sans attendre qu'il le
demande** :
1. Pour chaque PGO/PDP/IST **nouveau ou d'indice supérieur** : le lire (Excel via COM si `.xlsm`),
   mettre à jour la fiche dans les DEUX dépôts — l'**indice texte ET la couverture structurée**
   (`pgo.couverture` / `pdp` dates), cf. le piège Gantt ci-dessous — et le copier dans App Tech en
   remplaçant l'ancien.
2. Pour chaque chantier signalé « App Tech en retard » ou « techniciens planifiés sans App Tech » :
   compléter/créer le dossier App Tech (cf. checklist nouveau chantier).
Exemples manqués le 28/08 puis rattrapés : PGO Rion des Landes ind.12, PGO Bradascou ind.16, et le
dossier App Tech de Bradascou (techniciens en S36 sans App Tech).

### Deux garde-fous, parce que la règle seule n'a pas tenu (31/08/2026)

**Cette règle a été oubliée DEUX fois** — le 28/08, puis le 31/08 où Patrice a dû redemander
(« ça fait deux fois que je suis obligé de te rappeler des tâches à faire »). Même diagnostic que
`SEED_VERSION` : une consigne écrite ici ne vaut que ce que vaut la mémoire de celui qui la lit.
Même remède, deux mécanismes exécutés par l'outil.

**1. Le rappel revient à chaque message tant que ce n'est pas traité.**
`scripts/hook-veille-prompt.ps1`, déclaré en `UserPromptSubmit` dans `.claude/settings.json`.
Il compare l'horodatage de `veille/dernier.json` à celui de **`veille/traitement.json`** et, tant
qu'ils diffèrent, ajoute une ligne courte — dans le contexte de Claude **et à l'écran de Patrice**
(champ `systemMessage`), pour qu'il voie lui-même si le travail a été fait plutôt que d'avoir à le
demander. Il ne compte que ce qui est réellement actionnable : les nouveautés, plus les dossiers
App Tech en retard **marqués `Prioritaire`** (donc avec des techniciens placés — cf. la règle de
lecture ci-dessus). Un rapport sans rien d'actionnable ne déclenche aucun rappel.

**Pour l'éteindre : `scripts/marquer-veille-traitee.ps1 -Note "…"`.** La note est obligatoire —
elle force à dire ce qui a été fait et sert de trace. **Constater qu'un élément ne demande rien EST
un traitement valable**, mais il faut l'avoir constaté : ne jamais marquer sans avoir regardé.

**2. Le contrôle des chantiers actifs, pour ce que la veille ne peut pas voir.**
`scripts/veille-chantiers-actifs.ps1`, appelé par le hook `SessionStart` à la suite du rapport.
La veille prend une photo à 07h30 ; **tout ce qui arrive ensuite est invisible jusqu'au
lendemain**. Le 31/08/2026, deux PGO de Portet sont arrivés à 08h55 et l'IST à 10h34 : dans aucun
rapport, alors que trois techniciens y étaient toute la semaine — et le PGO du dossier App Tech
(V40) ne couvrait plus la période, ses lignes TELSAM s'arrêtant au 28/08. Le 43-1 ajoutait
justement « 31/08 → 04/09 : fin de raccordement BR SECTIONNEMENT ». C'est Patrice qui l'a vu.

Ce contrôle ne rejoue pas les 143 000 fichiers : il ne regarde que les dossiers des chantiers où
des techniciens sont placés (semaine en cours + 3 semaines), et n'y cherche que les documents de
sécurité plus récents que le dernier passage de la veille. **3 secondes.** Il dit pour chacun s'il
est déjà dans App Tech, et dédoublonne source et copie.

**`.claude/settings.json` n'est sauvegardé par aucun dépôt** (le dossier parent n'est pas un dépôt
Git). S'il disparaît, recréer les DEUX hooks : `SessionStart` → `hook-veille-session.ps1`,
`UserPromptSubmit` → `hook-veille-prompt.ps1`.

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
1. **Un seul lien pour tout le monde** : `https://pivot-telsam.github.io/appli-techniciens/`. Il ne
   change jamais. C'est le **mot de passe** qui identifie le technicien depuis le 27/08/2026, plus
   l'adresse : le commit du portail a supprimé la lecture de `?tech=` (`URLSearchParams`), il n'en
   reste rien dans le code. Les anciens liens personnels `?tech=slugifiedname` restent inoffensifs
   (le paramètre est ignoré, on tombe sur l'écran du mot de passe) — inutile de les récupérer, mais
   **ne plus jamais les diffuser ni les présenter comme personnels**. Question posée par Patrice le
   28/08/2026 en relisant le mail de mise en service : cette ligne de doc disait encore l'inverse.
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
paragraphe. (La sauvegarde `..._backup_2708.docx` a été supprimée le 28/08, cf. ci-dessous.)

**Mise à jour du 28/08/2026 — version envoyée aux 13 techniciens.** Quatre changements :
1. **Section 5 « Envoyer ta feuille » entièrement réécrite** : l'ancienne décrivait le partage du
   PDF puis un mail à `equipefibretelsam@telsam.com`. Nouveau texte : un appui, l'appli fabrique
   le PDF et l'envoie directement (relais Cloudflare → Dropbox), attendre le message de
   confirmation, plus de mail ni de pièce jointe ; note italique sur l'échec réseau (« ta saisie
   reste sur le téléphone, réappuie plus tard, rien n'est perdu »).
2. **Les DEUX mots de passe sont maintenant expliqués** (section 1) : celui de l'appli (personnel)
   et celui des documents (commun, demandé par Dropbox à la première ouverture d'un dossier de
   chantier). Ajouté aussi que **le tiret n'est pas obligatoire** (`normPw`). La puce
   « 📁 Documents » de la section 2 rappelle que Dropbox le demande une fois par chantier.
3. **Puce « 📋 Suivi » ajoutée** en section 2, formulée « sur certains chantiers seulement » — le
   bouton était alors en pilote sur DATA4-Marcoussis seul. **Réécrite le 31/08/2026**, tous les
   chantiers actifs étant équipés : la puce décrit maintenant le geste (cocher en fin de journée),
   le document « Avancement » qui se refait chaque matin dans Documents, et la consigne qui compte
   — *ne pas corriger le document, écrire la correction dans le commentaire*, puisqu'il est
   reconstruit de zéro chaque matin. **Seule la FIN de la puce a été remplacée** (plage bornée à
   partir de « sur certains ») : réécrire le paragraphe entier aurait fait perdre à l'emoji son run
   en `Segoe UI Emoji` et il serait devenu un carré. Contrôlé après coup : emoji présent, run
   toujours en Segoe UI Emoji, texte en Calibri, document toujours sur 1 page.
4. Section 6, première puce : « Tant que tu n'as pas **appuyé sur Envoyer** » (au lieu de « envoyé
   le PDF »), cohérent avec l'envoi direct.

**Méthode employée, à reprendre** : plutôt que Find/Replace dans Word (qui a corrompu un
paragraphe le 27/08), le `document.xml` a été extrait, découpé en paragraphes, les paragraphes
visés régénérés à l'identique du style voisin (`w:pPr`/`w:rPr` recopiés, runs emoji en
`Segoe UI Emoji`), puis réinjecté dans le `.docx` avec
`[System.IO.Compression.ZipFile]::Open(..., 'Update')` — pas de `zip` sur cette machine. Contrôles
faits : `<w:p>`/`<w:r>` appariés, zéro `Ã` et zéro caractère de remplacement, ouverture Word OK
(45 paragraphes, 1 page), puis `SaveAs2` docx + PDF.

**RÈGLE — pas de fichier `_backup_*.docx` qui traîne sur le Bureau.** Le 28/08/2026, juste avant
d'envoyer l'appli aux 13 techniciens, Patrice a demandé de supprimer les documents Word périmés
« pour éviter de me tromper » — le vrai risque du jour était d'ouvrir une vieille version et de
joindre le mauvais PDF au mail. Envoyés à la corbeille (récupérables, jamais d'effacement
définitif) : `Appli_TELSAM_Consigne_Technicien_backup_2708.docx`,
`..._backup_2808.docx`, `Memo_Exploitation_TELSAM_backup_2708-2.docx`. Lui-même avait déjà
supprimé `Liens_App_Techniciens_TELSAM.docx`. **Le Bureau ne doit garder que 4 documents Word** :
`Appli_TELSAM_Consigne_Technicien.docx` (+ son PDF), `Mots_de_passe_App_Techniciens_TELSAM.docx`,
`Memo_Exploitation_TELSAM.docx`, `Chantiers_Numeros_Index_TELSAM.docx`. Conséquence pour les
prochaines modifications de ces documents : **la sauvegarde va dans le scratchpad de la session,
pas sur le Bureau**, et on ne s'appuie plus sur un `_backup_` du Bureau comme filet.

**Le document est une page A3 portrait sur DEUX colonnes** (`w:pgSz 16838×23811`, `w:cols w:num=2`)
— pas un A4. Ne pas s'étonner d'un « 1 page » avec 800 mots, et ne pas ajouter de contenu sans
vérifier que ça tient encore sur la page. **Le PDF pèse ~4,3 Mo** : c'est la police
`SegoeUIEmoji` embarquée (un seul flux de 8 Mo avant compression, les emoji couleur ne se
sous-ensemblent pas vraiment). Normal, pas une régression ; si un jour il faut l'alléger, la seule
voie est de remplacer les emoji par du texte — au prix de la ressemblance avec l'écran du
téléphone.

**`Mail_general_mise_en_service_appli.txt`** (Bureau) — le mail de mise en service envoyé à tous
les techniciens le 28/08/2026, rédigé au nom de l'équipe. Il reste volontairement court : le lien
unique, la liste de ce qu'on trouve dans l'appli, le **principe des deux mots de passe** (personnel
pour l'appli / commun pour les documents, envoyés dans un mail individuel), l'ajout à l'écran
d'accueil, et le renvoi à la notice PDF pour tout le détail. Patrice envoie ensuite un mail
individuel par technicien avec ses deux mots de passe.

**`Memo_Exploitation_TELSAM.docx`** — le « qui fait quoi » de Patrice, **12 sections depuis le
28/08/2026** (A4 portrait, deux colonnes, 4 pages). À tenir à jour dès qu'une tâche récurrente
change de main : c'est le document qu'il relit pour savoir ce qu'il doit faire et ce que je fais.
Il n'en existe pas de PDF, et Patrice n'en veut pas.

**Mise à jour du 28/08/2026.** Ce mémo est le document qui vieillit le plus vite : il décrit des
circuits, et chaque changement de l'appli en périme un. Ce qui a changé :
- **Section 6 (feuilles d'heures)** : elle décrivait encore le technicien envoyant son PDF par mail
  à `equipefibretelsam@telsam.com`, et Patrice enregistrant les pièces jointes dans
  « feuille d'heures / dépôts appli ». Faux depuis le relais : le technicien appuie, la feuille
  arrive d'elle-même dans le dossier, Patrice n'a plus rien à enregistrer. Puce « Moi » ajoutée sur
  le relais Cloudflare (en cas de panne, le technicien garde sa saisie — rien à lui redemander).
- **Nouvelle section 7 « Le suivi d'avancement — nouveau, en pilote »**, insérée juste après les
  heures (même circuit : ce qui remonte du terrain par l'appli). Dit explicitement qu'un seul
  chantier est équipé, qu'ouvrir un chantier de plus demande son devis + la validation des tâches
  par Patrice, et les deux limites assumées (fichier déposé pas confortable à lire dans App Tech,
  regroupement multi-techniciens à faire).
- **Sections 7 à 11 renumérotées en 8 à 12**, et l'ancienne 11 « Avant le déploiement » réécrite en
  **12 « Le déploiement aux 13 techniciens »** (fait / toi / moi).
Les sections 1 à 5 et 10-11 n'ont pas bougé — la puce « deux serrures » de la section 11 était déjà
juste. Mêmes styles que les voisins (`H` / `SUB` / `BUL` / `NOTE`, cf. la méthode d'édition docx
ci-dessus). Sauvegarde dans le scratchpad de la session, pas sur le Bureau.

**Mise à jour du 03/09/2026 — 8 passages réécrits, une section 13 ajoutée.** Ce mémo décrit des
circuits, donc chaque changement de l'outil en périme un morceau : ce qui a été corrigé était
devenu **faux**, pas seulement incomplet.
- §1 : la chaîne tourne **deux fois par jour** (7h30 et 13h00) ; le rappel est une **bulle Windows**,
  le mail est abandonné (le nouvel Outlook ne se pilote pas) ; ce qui arrive dans la matinée est
  rattrapé à 13h.
- §2 : Patrice **n'a plus à me dire** que le planning est à jour, ni à renvoyer le suivi par Teams.
- §5 : écrire le numéro **dans la ligne** du planning, et pourquoi (sans lui le technicien ne voit
  rien ce jour-là).
- §11 : **trois** serrures, pas deux.
- **§13 nouvelle** : le planning en ligne, la case « Valider ce planning » du vendredi, qui peut
  cocher, et ce que le technicien voit en plus (avec qui il est, six semaines devant, les cases qui
  restent cochées).

**Méthode, et deux garde-fous qui ont servi le jour même** (script gardé dans le scratchpad, la
sauvegarde du mémo aussi — **pas sur le Bureau**, cf. la règle des `_backup_`) :
1. **chaque paragraphe visé est vérifié par son texte avant d'être touché**, et on n'écrit rien si
   un seul ancrage manque. Ça a rattrapé un ancrage que j'avais écrit sans accent — sans lui,
   j'aurais réécrit le mauvais paragraphe en silence ;
2. **le contrôle d'encodage était cassé, et cassé dans le sens qui ne se voit pas**. Écrit
   `if ($x -match [char]0xC3 + '|' + ...)`, PowerShell applique `-match` AVANT les `+` : la
   condition valait une chaîne non vide, donc **toujours vraie**. Il bloquait une écriture légitime,
   et il aurait laissé passer un vrai mojibake avec la même indifférence. **Parenthéser le motif**,
   et éprouver le contrôle sur un texte sain ET sur un texte abîmé.

Contrôles après écriture : Word ouvre le document (127 paragraphes, 4 pages, 2508 mots), les
nouveaux textes sont présents, les anciens absents, zéro caractère abîmé.

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
- **LE PIÈGE DU `.ps1` SANS BOM NE CONCERNE PAS QUE LES COMPARAISONS DE NOMS — il abîme aussi tout
  texte accentué que le script ÉCRIT.** Il est documenté plus haut pour `ClePersonne` (un
  « François » qui ne correspond plus), mais je m'y suis fait prendre autrement le 03/09/2026 en
  créant la fiche 26-071 : le script portait `nom = 'Bollène - Plantades'`, PowerShell 5.1 a lu le
  fichier en ANSI, et **le mojibake est parti dans `SEED_DATA` des deux dépôts** (« BollÃ¨ne »,
  « Â» »). Neuf séquences côté suivi, deux côté appli.
  - **Parade** : soit écrire le `.ps1` **avec BOM UTF-8**, soit composer les accents
    (`'Boll' + [char]0x00E8 + 'ne'`), soit — le plus simple — n'écrire que de l'ASCII dans le script
    et poser les libellés accentués avec l'outil Edit après coup.
  - **Ce qui l'a rattrapé** : le comptage systématique des `Ã` après écriture (règle ci-dessus).
    Il valait 1 par fichier au lieu de 0. **Ce contrôle n'est pas une formalité : c'est la seule
    chose qui a vu l'erreur.**
  - **Réparation ciblée, pas globale** : remplacer uniquement les séquences `Ã¨ Ã© Â« Â»`, qui
    n'existent jamais dans du français correct. Une réparation 1252→UTF-8 appliquée à tout le
    fichier aurait cassé les accents SAINS des 70 autres fiches.
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
- **« Montrer le résumé avant de committer » ne veut PAS dire « demander la permission de
  travailler ».** Précision posée par Patrice le 01/09/2026, après avoir dû redire une TROISIÈME
  fois une règle déjà écrite : « mets à jour systématiquement notre suivi quand un nouveau
  document arrive et remplace dans App Tech les documents obsolètes — tu devrais le faire
  systématiquement sans que je te le demande ». Sa crainte, dite mot pour mot : « si je suis
  obligé de répéter les choses systématiquement, cela va être compliqué d'utiliser cet outil ».
  - **Par défaut et sans rien demander** : lire le document, mettre à jour la fiche dans les DEUX
    dépôts (indice ET couverture structurée), remplacer la version périmée dans App Tech quand le
    dossier existe, bumper `SEED_VERSION`/`APP_VERSION`, puis montrer le résumé avant de committer.
  - **On ne demande que pour une DÉCISION**, jamais pour une mise à jour : rouvrir un chantier
    marqué terminé, trancher entre deux devis, diffuser une IST dont la validation RTE n'est pas
    certaine, choisir un chemin de destination Dropbox ambigu.
  - **Pourquoi la règle n'avait pas tenu — la vraie cause, vérifiée le 01/09/2026.** Les scripts
    de veille écrits la veille répétaient eux-mêmes « avec l'accord de Patrice » : à chaque
    message (`hook-veille-prompt.ps1`), à chaque démarrage de session (`hook-veille-session.ps1`),
    dans `veille-chantiers-actifs.ps1` et dans `marquer-veille-traitee.ps1`. La consigne écrite
    ici était donc contredite en permanence par le mécanisme censé la faire appliquer — et un
    texte lu à chaque message pèse plus lourd qu'une ligne lue une fois au démarrage. Les quatre
    formulations ont été réécrites le 01/09/2026.
  - **Leçon générale, valable au-delà de ce cas** : quand une règle ne tient pas, ce n'est presque
    jamais un problème de mémoire. C'est soit qu'aucun mécanisme ne la porte (cas `SEED_VERSION`,
    résolu par un hook), soit qu'un mécanisme dit le contraire (ce cas-ci). **Répondre « c'est
    noté » à Patrice ne vaut rien** : ce qui vaut, c'est de trouver quel mécanisme manque ou
    lequel contredit la règle, et de le corriger.
- En cas de nom de chantier ambigu ou de dossier Dropbox introuvable/multiple, ne jamais deviner —
  poser la question à Patrice.
- Valider la syntaxe JS (node --check) et le JSON (SEED_DATA) avant toute livraison.
- **Avant de committer un changement à `SEED_DATA` dans suivi-chantiers, incrémenter `SEED_VERSION`.**
  Sans ça, la modification est invisible pour les collègues qui ont déjà chargé une version antérieure.
  Cette règle a déjà été oubliée deux fois malgré sa documentation (20/08 puis 21/08) — un hook
  git local (`scripts/check-seed-version.sh`, installé dans `.git/hooks/pre-commit`) bloque
  maintenant tout commit qui modifie un enregistrement `SEED_DATA` sans toucher `SEED_VERSION`
  dans le même commit. À réinstaller après un nouveau clone du dépôt :
  `cp scripts/pre-commit.sh .git/hooks/pre-commit && chmod +x .git/hooks/pre-commit`
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
  `cp scripts/pre-commit.sh .git/hooks/pre-commit && chmod +x .git/hooks/pre-commit`
  (les hooks ne sont pas versionnés par Git lui-même — ce garde-fou ne protège que la machine où
  il est installé).
- **CE FICHIER EST LE MÊME DANS LES DEUX DÉPÔTS — un hook refuse de les laisser diverger**
  (mis en place le 31/08/2026). `appli-techniciens/CLAUDE.md` et `suivi-chantiers/CLAUDE.md`
  sont **un seul document**, dupliqué pour que chaque clone soit complet. Rien ne les tenait
  ensemble : ils ont divergé en silence pendant trois jours. `suivi-chantiers` annonçait encore
  des liens `?tech=slugifiedname` personnels et permanents, alors que le portail par mot de passe
  du 27/08 avait supprimé la lecture de ce paramètre — selon le dépôt dans lequel on travaillait,
  on lisait une règle périmée, et on aurait pu rediffuser ces liens comme s'ils identifiaient
  encore un technicien.
  `scripts/check-claude-md-sync.sh` (identique dans les deux dépôts) **bloque tout commit qui
  toucherait `CLAUDE.md` sans que l'autre dépôt dise déjà la même chose**, et affiche l'écart. La
  marche à suivre est donc : modifier un fichier, le **copier octet pour octet** dans l'autre
  (`cp CLAUDE.md ../<autre-dépôt>/CLAUDE.md` — jamais une relecture/réécriture, les accents se
  cassent), puis committer les deux. Si le dépôt voisin est absent (clone isolé), le hook
  **avertit et laisse passer** plutôt que de bloquer un travail qu'il ne peut pas vérifier.
- **Un seul hook `pre-commit` est possible : c'est `scripts/pre-commit.sh` qui enchaîne tous les
  contrôles.** Depuis le 31/08/2026 chaque dépôt en a plusieurs (version + cohérence des
  CLAUDE.md). Installer directement un `check-*.sh` comme hook, comme le disaient les lignes
  ci-dessus jusqu'à cette date, **désactiverait silencieusement les autres**. Le lanceur prend
  tout `scripts/check-*.sh` du dépôt : un nouveau contrôle déposé là est actif sans rien modifier.
  - **PIÈGE rencontré en testant ce hook** : `git show :CLAUDE.md` applique la conversion de fins
    de ligne de `core.autocrlf` (activé ici) et rend du CRLF, alors que le fichier voisin est lu
    tel quel sur le disque. Deux fichiers identiques paraissaient alors différer **à chaque
    ligne** (2927 lignes d'écart) et le hook bloquait tout. Utiliser `git cat-file blob`, qui rend
    les octets bruts, et neutraliser les `\r` des deux côtés.

## Onglet « Affaires » — le cycle de vie commercial (mis en place le 04/09/2026)

**UNE AFFAIRE = UN CHANTIER. Tranché par Patrice le 04/09/2026**, après un premier retour où je
proposais l'inverse (« une affaire = un devis »). **Ma proposition était FAUSSE et ne doit pas
revenir** : le n° de devis n'est pas unique. `TELSAM_CC_RTE_25107` désigne DEUX devis réels —
Bradascou (26-054) et Fleyriat (26-055) — et `25115` deux autres — DATA4 (26-065) et Bissy
(26-064), qui a même une commande client « Dev25115 ». L'Excel contourne en réécrivant `25106` et
`25114`, alors que `25106` est le devis d'Arbresles-Charpennay et que `25114` n'existe pas.
**J'ai failli « corriger » les fiches sur la foi de l'Excel : c'est Dropbox qui a tranché en faveur
des fiches.** Le n° de chantier `26-XXX` est la seule clé unique.

**Une affaire porte donc 1..n devis, et son étape s'AGRÈGE** — jamais celle d'une ligne. Cantegrit
26-003 porte quatre devis : un PERDU à 65 100 €, deux GAGNE, un sans statut. Six affaires sur 41
sont dans ce cas.

### La donnée — `AFFAIRES_RTE` et `scripts/affaires-rte.ps1`

Même nature que `PLANNING_RTE` et `POSES_APPLI` : **refaite de zéro à chaque passage**, vit **à
côté** de `SEED_DATA` (donc **pas de `SEED_VERSION` à bumper** et aucun état local des collègues
effacé). Ne jamais y écrire à la main. 38 Ko pour 41 affaires.

- Source : **`C:\Users\patrice.pivot\Desktop\SUIVI RTE & TELECOM.xlsx`** (le classeur maître),
  **complété par** `SUIVI RTE TELECOM - mise a jour_3.xlsx`. Les deux portent la colonne
  **`F N° Chantier`**, donc **toutes les colonnes suivantes sont décalées d'un rang** (N° Devis en
  G, Montant en H, statut en AO, PV EN COURS en AR…).
  *Cette ligne disait jusqu'au 07/09/2026 que `mise a jour_3` faisait foi « plus `SUIVI RTE &
  TELECOM.xlsx` » : c'était vrai le 04/09, quand seule la copie portait la colonne F.*
- **DEUX CLASSEURS, ET C'EST VOLONTAIRE (07/09/2026).** Patrice numérote la colonne F
  progressivement, et son travail s'est retrouvé réparti sur deux fichiers : **le maître porte la
  série 25- entière plus 28 numéros 26-**, la copie `mise a jour_3` porte **les 26- des lignes de
  2026** que le maître n'a pas encore. Comparés **cellule par cellule** le 07/09/2026 : les deux
  feuilles sont **identiques sauf la colonne F**, et **sur aucune ligne elles ne se contredisent**
  — l'une est vide là où l'autre porte un numéro. Donc :
  - on lit le maître et on **complète** avec la copie, ligne par ligne, **uniquement quand F y est
    vide** (paramètre `-Complement`, mettre `''` le jour où le maître portera tout) ;
  - le nombre de numéros empruntés est **publié** (`empruntes`) et **la vue le dit** : l'emprunt se
    voit, et la ligne disparaît d'elle-même quand un seul classeur porte toute la colonne ;
  - une contradiction éventuelle est **annoncée**, jamais arbitrée en silence — le maître gagne, et
    la vue nomme les deux numéros. Éprouvé en fabriquant une contradiction, puis en restaurant la
    donnée réelle ;
  - **ne jamais écrire dans les classeurs de Patrice pour « consolider » sans qu'il le demande** :
    lequel garder est sa décision (cf. [[feedback_fichiers_de_travail_de_patrice]]).

**CE QUI A ÉTÉ ÉCRIT DANS SES CLASSEURS LE 07/09/2026, À SA DEMANDE** (« j'avais oublié le fichier
mis à jour 3, tu peux le compléter avec celui que je viens de te renvoyer ») :
- dans **`mise a jour_3`** : **102 cellules** de la colonne F recopiées depuis le maître (87 de la
  série 25-, 15 de la série 26-). Ce fichier porte donc désormais **toute** la numérotation (476
  cellules) ;
- dans le **maître** : **les 2 corrections de doublon** seulement (F383, F406). Il **manque encore
  34 cellules** — les 26- des lignes de 2026, que seul `mise a jour_3` porte. Patrice n'a pas
  demandé de les y écrire : à proposer, pas à décider.

**LA MÉTHODE D'ÉCRITURE, à reprendre telle quelle** (Excel COM, le seul moyen d'écrire ; la lecture
reste par dézippage) :
1. **vérifier l'absence de verrou `~$` et qu'Excel ne tourne pas**, et faire une **copie de
   sauvegarde dans le scratchpad** — pas sur le Bureau (cf. la règle des `_backup_`) ;
2. **le script d'écriture est en PUR ASCII.** Premier essai perdu sur ce piège déjà documenté :
   `Worksheets.Item("Suivi Activité")` écrit avec son accent dans un `.ps1` sans BOM est lu en
   ANSI, devient `ActivitÃ©`, et Excel répond `DISP_E_BADINDEX`. La feuille se trouve par motif
   sans accent (`suiviactivit`), l'en-tête se compare **sans le caractère `°`** ;
3. **la colonne F porte un format de DATE (`jj/mm/aaaa`) alors que ses valeurs sont du TEXTE.** On
   ne touche pas au format de Patrice, mais **chaque cellule est relue après écriture** et le
   script refuse d'enregistrer si Excel a converti la valeur (`$relu -isnot [string]`) ;
4. **garde-fous avant d'écrire une cellule** : le libellé (colonne D) doit être **identique dans
   les deux classeurs** sur cette ligne, la cellule cible doit être **réellement vide**, et pour
   une correction la valeur trouvée doit être **celle qu'on croit remplacer**. Un seul écart et
   **rien n'est enregistré** — `$wb.Save()` n'est appelé qu'à la toute fin ;
5. **contrôle après coup, avec son contre-exemple** : les deux classeurs restent **identiques hors
   colonne F** (0 écart mesuré), plus aucun numéro 25- sur deux chantiers différents, aucun numéro
   du maître ne manque dans `mise a jour_3` — **et le contre-exemple qui prouve que le contrôle
   sait dire non** : le maître, lui, en manque bien encore 34.
- **Le classeur est lu en DÉZIPPANT le `.xlsx`, jamais via Excel COM** : Patrice peut l'avoir
  ouvert, rien n'est verrouillé et rien n'est écrit dans son fichier. La feuille est trouvée par
  son nom, pas par `sheet1.xml`.
- **Seuls les numéros `26-` sont pris** (`-Prefixe`) : la vue ne montre que **l'année en cours**.

  **LE PRÉFIXE EST L'ANNÉE, PAS L'ENTITÉ — corrigé par Patrice le 07/09/2026.**
  Sa phrase : « j'ai donné le chiffre 25 aux vieux devis de data parce que nous avions encore cette
  entreprise à ce moment-là. Depuis, nous sommes passés chez Telsam en 2025. C'est pour ça qu'en
  2025 il y a des devis Telsam et des devis data. »

  *J'avais écrit ici que « 21- à 25- numérotent les anciens devis de DATA », et j'ai construit une
  question entière là-dessus (« faut-il mélanger les devis TELSAM de 2025 dans la série 25- ? »).
  La question n'avait pas lieu d'être.* Si les séries 21- à 24- sont à 100 % sur des devis
  `DATACC`, c'est simplement que DATA était l'entreprise à l'époque ; **2025 est l'année de la
  bascule et porte donc les deux**. **NE PAS RÉÉCRIRE « 25- = DATA »** : le préfixe dit l'année,
  l'entité se lit dans le numéro du devis.
  Vérifié dans le fichier : sur 475 numéros, **437 ont pour préfixe l'année de leur ligne**, et les
  38 autres portent l'année **suivante** — un devis de fin 2022 numéroté `23-`, un devis de 2025
  numéroté `26-`. C'est logique : le numéro suit l'année où le chantier se fait.
  - **C'est vérifiable dans le fichier, et ça a été vérifié** : les séries 21-, 22-, 23- et 25-
    sont à **100 % sur des devis `DATACC…`**, la série 26- à **100 % sur des devis `TELSAMCC…`**.
    Seule la série 24- fait exception, avec **3 lignes `TELSAMCC` sur 146**.
  - La série 25- est apparue le **07/09/2026** (`25-001` à `25-085`, sans trou), posée dans l'ordre
    des lignes ; **le préfixe ne suit pas l'année** (`25-011` est sur une ligne de 2023). Les vrais
    numéros 26-, eux, ne sont **pas** dans l'ordre des lignes (r345 = 26-031, r381 = 26-007,
    r423 = 26-055). *Mon raisonnement du matin — « une suite par ligne » — était juste sur la forme
    mais passait à côté de la raison : ce sont les devis d'une autre entité.*
  - **Ne pas reprendre un numéro 21- à 25- comme clé de fiche, et ne pas proposer de renommer un
    dossier Dropbox avec.**
  - **La liste des préfixes écartés est construite depuis la donnée, jamais écrite en dur** : un
    intitulé « 21- à 24- » aurait continué de s'afficher en cachant 87 lignes.
- Les colonnes de dates sont converties **colonne par colonne** : un montant comme `46144` tomberait
  sinon dans la plage des dates.

### CE QUE LA VUE REFUSE DE FAIRE — et pourquoi chaque refus a coûté un essai

1. **Elle ne laisse jamais croire qu'elle est complète.** Bandeau permanent : les affaires vivantes
   sans numéro (28 le 04/09/2026, **14 le 07/09** — leur **répartition par année est comptée**, plus
   annoncée de mémoire), les chantiers actifs sans ligne numérotée (comptés **en direct sur les
   fiches**, pour rester justes si une fiche est ajoutée), les lignes écartées et leurs préfixes.
2. **L'étape « Accord verbal » n'est PAS calculée.** Aucune colonne ne la porte. La deviner depuis
   « GAGNE » mélangerait l'accord et la commande — la vue dit qu'elle n'est pas suivie.
3. **Un vide n'est pas un manque.** Trois états : renseigné, **sans objet**, manquant.
   `afRenseigne()` rejette `absent`, `NA`, `?`, `pas sur PGO`, et **toute phrase du genre « à priori
   non concerné »** — sans ça la consignation de Portet criait « à recaler » sur un chantier qui n'en
   a pas. Un voyant qui crie à tort finit par ne plus être lu.
4. **Une date invraisemblable reste du texte.** `afDate()` borne l'année à 2000-2100 : le fichier
   contient la coquille `17/06/206` (PGO de Portet), qui devenait l'an 206 — donc une date
   « largement passée » — et **faisait basculer le chantier en « travaux en cours » sur une faute de
   frappe**. Trouvé par le banc d'essai, pas par relecture.
5. **On n'affirme jamais « Facturé » sur un montant inconnu.** Il faut que TOUS les devis gagnés
   aient leur montant. Cantegrit 26-003 sortait « Facturé » avec 3 345 € de devis chiffré et
   58 345 € facturés.
6. **Les écarts entre le fichier commercial et le suivi sont DITS, jamais arbitrés.** C'est ce que
   la vue apporte et que ni l'Excel ni le suivi ne savent dire seuls : **12 affaires** en écart, dont
   9 marquées « Devis envoyé » alors que des heures y ont été passées ou que des techniciens y sont
   placés (Givors 26-060 : 49 h). La cause n'est pas un défaut de lecture — sur ces affaires **la
   seule ligne numérotée est un devis TS**, donc l'étape décrit le TS et pas le chantier. Deviner
   « en fait c'est en travaux » serait le vert mensonger que ce fichier interdit. Bandeau rouge en
   tête de vue + bloc dans la fiche, avec les deux corrections possibles laissées à Patrice.

### Ce qui est volontairement absent de cette première version

- ~~**Aucune saisie**~~ — **LEVÉ LE 07/09/2026, cf. « LOT 2 » ci-dessous.** La vue était en lecture
  seule en attendant la base commune ; la base est là, la saisie y vit.
- **Aucun blocage** : les points de passage sont des **voyants**. On regarde d'abord lesquels crient
  à tort, on ne transforme en barrage que ce qui le mérite. Mesuré dans son fichier : l'acte de
  sous-traitance est vide sur 664 lignes sur 763, le « PV en cours » sur 760 — des gates durs
  bloqueraient presque toutes les affaires dès le premier matin.
- **Pas de glisser-déposer** : il sert une dizaine de fois dans la vie d'une affaire, la lecture
  vingt fois par jour. À reparler après usage.

### LOT 2 — la vue devient saisissable (07/09/2026)

Demande de Patrice : les **deux numéros de commande** (RTE et client) séparés, des **menus
déroulants**, une **vue par chantier**, **l'affectation** (« affaire suivie par »), et un **menu
déroulant de statut** qui ouvre une **fenêtre de saisie** au changement (« si un devis est envoyé,
tu nous marques devis envoyé, ça nous donne une pop-up et nous remplissons de suite le numéro du
devis »). **« C'est toi qui définis le statut. Ne te fie pas à ce qui est marqué sur le suivi. »**

**Ses cinq arbitrages du 07/09/2026**, tous appliqués :
1. le premier numéro est la **commande RTE**, le second la **commande du client** — « oui quand il
   y a un client » (donc chez un client RTE en direct il n'y en a qu'un) ;
2. **un seul menu déroulant**, pas deux étages. Pour La Curbans (un devis gagné + un devis en
   attente), « le devis en attente doit être un devis de TS, je vais me renseigner » — **question
   ouverte, à ne pas trancher seul** ;
3. « suivie par » = **les 4 qui gèrent les travaux** : Brillou, Cazenave, Hamouch, Vidal.
   **Sans Carine Rambaud, sans Guillaume Heras, et sans Patrice lui-même** ;
4. **« Soldé » = « Terminé »** sur la fiche chantier ;
5. **saisie dans la base partagée**, et **tout le monde peut modifier**.

**LA DONNÉE — la colonne unique de l'Excel est SÉPARÉE, sans jamais rien jeter.**
La colonne I « N°Commande Client » porte les deux numéros dans une seule case
(`4500807063 / 920087601`, 30 cases sur 64), et la colonne J « COMMANDE » n'est que la **recopie du
premier segment** (30 fois sur 35) — colonne redondante, non reprise. Lecture vérifiée par client :
RTE en direct ⇒ **un** numéro (`6100…`, parfois `4500…` chez RTE ARTERIA) ; sous-traitance ⇒ **deux**,
le `4500…` de RTE puis le numéro propre du client (`920…` OMEXOM/VINCI, `24…`/`BOOOO…` INEO,
`900…`/`9001…` BOUYGUES ES, `C511…` EIFFAGE, `POSZ…` RODA, `P.079…` OMEXOM FG).

- **ON NE DÉCOUPE PAS SUR LE « / ».** CTEAM (26-070) écrit `12/2026/464`, qui est **UN** numéro : un
  découpage naïf le cassait en trois (vu à l'écran avant d'écrire la version retenue). On **cherche**
  le numéro RTE par son motif — 10 chiffres commençant par 45 ou 61 — et ce qui reste est la
  commande du client.
- **16 cases sur 64 ne contiennent pas un numéro mais une phrase** (« Devis envoyé », « Attente
  commande », « Attente avenant », « PERDU », « RTE demande à AXIANS de refaire les mesures ») : la
  case du numéro servait aussi de statut. C'est exactement ce que le menu déroulant supprime, et ce
  texte est **conservé** (`cdeTexte`, affiché dans la fiche), pas jeté.
- **La case brute reste dans `cdeClient`**, telle que Patrice l'a tapée, et un montant entre
  parenthèses ou un « attente AV1 » part dans `cdeNote`. **Un contrôle de conservation des chiffres**
  compare l'avant et l'après case par case ; il **n'arrête pas le script** (la case brute est gardée,
  un arrêt priverait l'équipe de tout l'onglet) mais le doute est publié (`cdeIncertaines`) et **la
  vue le dit**. Mesuré le 07/09/2026 : **0 chiffre perdu sur 63 cases**.
- **L'affectation existe déjà dans l'Excel** : colonne V « NOM », en initiales — `CC` = Christian
  CAZENAVE, `PB` = Pierre Brillou, `FV` = François VIDAL, `AH` = Ahmed HAMOUCH. Renseignée sur
  **17 affaires sur 53** seulement : point de départ, jamais vérité. La traduction initiales → noms
  vit **dans la page** (`AF_INITIALES`), pas dans le script, qui reste en pur ASCII.

**LES DOUZE STATUTS** (`AF_STATUTS`), construits sur ce que son fichier dit réellement — les 16
phrases ci-dessus donnent la liste des états qu'il écrit à la main :
`a_chiffrer`, `devis_envoye`, `en_nego`, `attente_cde`, `cde_recue`, `pret`, `travaux`,
`travaux_finis`, `a_facturer`, `facture`, `soldee`, `perdue`.
**À TENIR EN PHASE AVEC `STATUTS` DANS `functions/api/affaires.js`** — l'API refuse tout statut hors
liste, une dérive donnerait un refus incompréhensible. Le banc d'essai compare les deux listes.

**SEPT DÉCISIONS QUI ONT CHACUNE UNE RAISON, à ne pas défaire :**

1. **Le statut posé à la main GAGNE, le calculé devient une remarque.** Sans ça Patrice ne pourrait
   pas corriger un statut faux, ce qui est tout l'objet de sa demande ; à l'inverse, jeter le calculé
   ferait perdre le croisement Excel/suivi, la seule chose que cette vue apporte. Le désaccord
   s'affiche « fichier : … » sous la liste, et en clair dans la fiche.
2. **Un statut inconnu en base ne passe PAS pour une saisie.** Défaut trouvé par le banc d'essai :
   `saisi: !!brut` valait vrai même pour une valeur hors liste — la ligne se disait « posée à la
   main » en affichant en réalité le statut calculé, et la mention « calculé » disparaissait.
3. **La fenêtre de saisie est TOUJOURS PASSABLE.** Bloquer un changement de statut sur un numéro
   qu'on n'a pas sous la main ferait renoncer au changement, et le statut resterait faux — ce qu'on
   cherche justement à supprimer. « Le statut compte, le numéro suivra. »
4. **Un statut qui n'a rien à demander n'ouvre AUCUNE fenêtre** (`travaux`, `soldee`, `pret`…).
   Une fenêtre pour rien fait cliquer dans le vide et on finit par éviter la liste déroulante.
5. **Un doublon de numéro PRÉVIENT, il ne refuse jamais.** `TELSAMCC25107` désigne deux devis réels
   (Bradascou et Fleyriat) : refuser rejetterait une saisie juste, taire laisserait passer une faute
   de frappe. *À savoir : dans l'Excel les 64 références sont uniques — Patrice a contourné en
   renumérotant Fleyriat en 25106. Le risque de doublon vient donc de ce qui sera TAPÉ.*
6. **« Soldé ⇒ Terminé » est DÉRIVÉ, jamais écrit dans la fiche.** Les fiches vivent dans IndexedDB,
   poste par poste, et sont remises à zéro à chaque `SEED_VERSION` : une écriture y serait invisible
   pour les six autres et effacée à la mise à jour suivante. `appliquerSoldees()` dérive donc
   `termine` du statut partagé à chaque chargement — **dans les deux sens** (repasser en travaux
   décoche), et **seulement ce que la dérivation a coché elle-même** (une fiche terminée à la main
   reste terminée).
7. **Le mode est DIT, jamais deviné.** Hors ligne, les listes déroulantes sont **désactivées** et un
   bandeau l'annonce. Une liste qui a l'air de marcher sans rien enregistrer est pire que pas de
   liste du tout.

**AUCUN TEXTE D'AIDE N'EST UN NUMÉRO RÉEL.** Première version, vue **à l'écran** : les exemples des
cases étaient de vrais numéros du fichier (`4500810332` = Givors, `920087601` = Dambron-Voves,
`2604186` = une vraie facture). Grisés dans une case vide, ils se lisaient comme une valeur
renseignée — sur la fiche d'Aure-Lannemezan on croyait lire la commande RTE d'un autre chantier.
**Aucun contrôle chiffré ne pouvait le voir** ; c'est la capture d'écran qui l'a montré, et un
contrôle a été ajouté depuis (avec son contre-exemple). Même règle que les mots de passe d'exemple
du 02/09/2026 : un exemple doit être **impossible** à confondre, pas seulement improbable.

**LES DEUX PAVÉS D'AVERTISSEMENT SE REPLIENT** (`<details>`). À 16 affaires en écart, ils occupaient
tout le haut de la page et le tableau — l'outil — commençait sous la ligne de flottaison. La ligne de
titre reste toujours visible et porte le compte ; le détail des écarts est ouvert d'emblée tant qu'il
tient en cinq lignes. **Ne pas les supprimer pour « alléger »** : ce qu'ils disent est ce que la vue
apporte de plus utile.

**Ce qui a été touché** : `scripts/affaires-rte.ps1` (séparation des commandes + colonne V),
`functions/api/affaires.js` (nouveau), `partage/schema.sql` (tables `affaire` et `journal_affaire`,
`schema_version` = 5), `suivi_chantiers_205.html` (constantes, fenêtre, vue, CSS).
**Les deux tables sont créées par le code au premier appel** — aucune manip pour Patrice, même
principe que `assurerColonneNote`.

**Testé — 161 contrôles, 0 échec**, chacun avec son contre-exemple :
`partage/test-api-affaires.html` (39, fausse base D1), `partage/bloc-test-saisie-affaires.html`
(62, sur le vrai code de la page), `partage/bloc-test-affaires.html` (59 + 1, rejoué sans
régression — son contrôle du filtre lisait la pastille `.afEtape` que la liste déroulante remplace,
il a été mis à jour plutôt que laissé en échec permanent).
**Deux défauts ont été trouvés par ces bancs** (le statut inconnu du point 2, et une prémisse fausse
de mon propre test sur le doublon 25107) **et un troisième par la capture d'écran** (les textes
d'aide) — aucun des trois n'aurait été vu par relecture.

### LOT 3 — dates, jauge, tri/filtres, création d'affaire (07/09/2026)

Patrice : « il manque encore pas mal de choses pour que nous puissions nous en servir » — les dates
prévisionnelles, « une petite jauge, un pourcentage de l'avancement par chantier, quelque chose de
visuel, efficace et discret », « qu'on puisse filtrer et trier tous les champs, ça c'est
important », et « qu'on puisse créer un nouveau chantier, nous ne pouvons pas le faire ».

**Ses trois arbitrages** : dates = **le prévisionnel de l'Excel seulement** (j'avais proposé d'y
ajouter le planning réel et la couverture PGO) ; **tri et filtres retenus** d'une fois sur l'autre ;
**on commence par l'affaire**, sans n° de chantier.

**CE QUE LA MESURE A DIT AVANT DE CODER — c'est elle qui a dicté la forme de la jauge.**
Quatre sources d'avancement possibles, aucune ne couvre les 53 affaires :

| ce qu'on mesurerait | disponible le 07/09/2026 |
|---|---|
| travail réellement fait (cases cochées) | 19 chantiers outillés, **5 avec une déclaration** |
| heures passées / prévues | 30 |
| facturé / devis | 22 chiffrées, 9 facturées |
| boîtes posées / au devis | 22 renseignées, **2** posées |

D'où **deux couches**, et pas une :
1. **une barre de cycle en douze crans**, sous le nom, 3 px — le rang du statut. Couvre les 53, ne
   peut pas mentir (elle montre le statut posé, rien de déduit), et reste discrète ;
2. **le pourcentage physique** dans sa propre colonne triable, **là où il existe seulement**.

**LES HEURES N'ENTRENT PAS DANS LA JAUGE, ET C'EST UN REFUS RAISONNÉ.** Elles couvriraient 30
affaires au lieu de 5 — mais le réalisé compte chaque technicien et le prévu non, donc une jauge en
heures dépasse 100 % sur un chantier à moitié fait. Elles restent dans l'onglet Heures.

**`AVANCEMENT_CHANTIERS`, publiée par `scripts/boites-posees.ps1`.** Les deux moitiés de
l'information vivent dans **l'appli** (`tachesVendues` et `AVANCEMENT_DECLARE`) : le suivi n'en voit
aucune. Le calcul se fait donc là où la donnée est, une seule fois — même principe que
`presence`/`techRanges`. Trois règles dedans :
- **un pylône déclaré hors de la liste vendue ne compte pas** : il est mis à part (`hors`) et dit.
  Sans ça Fleyriat sortait à **123 %** dans ma première version ;
- **« en cours » vaut une demie**, ni zéro ni un — sinon la jauge saute d'un cran à chaque bouton ;
- **un chantier sans tâches vendues n'a PAS 0 %, il n'a RIEN** (« — »). Un zéro se lirait « rien
  n'a été fait » alors qu'il veut dire « personne n'a coché ». Et **un 0 % réel** (chantier outillé,
  aucune case cochée : 14 cas) est **écrit en gris** avec la phrase qui le dit — vu à l'écran.

**LES COLONNES SONT DÉCRITES UNE FOIS (`AF_COLONNES`), ET TOUT EN DÉCOULE** : en-tête, ligne de
filtres, clé de tri, cellule. Écrire les treize cellules à la main puis ajouter le tri par-dessus
aurait garanti qu'une colonne finisse triable sur autre chose que ce qu'elle affiche.
- **Une case vide va toujours en dernier, DANS LES DEUX SENS.** Trier par date prévisionnelle en
  croissant remplirait sinon le haut de l'écran des 7 affaires qui n'en ont pas.
- Les choix des filtres à liste sont **construits sur la donnée**, jamais écrits en dur.
- Tri et filtres vivent dans **IndexedDB, donc par personne** : ce sont des préférences d'affichage,
  les mettre dans la base commune imposerait le tri de l'un aux six autres.

**LES DATES PRÉVISIONNELLES DISENT PEU, ET LA VUE NE LE CACHE PAS.** Mesuré : début renseigné sur
46 affaires sur 53, fin sur 43, et surtout **29 fenêtres entièrement passées** contre 4 à venir. Une
fenêtre passée est affichée en gris avec son infobulle. C'est déjà la raison pour laquelle la
cascade d'étapes saute « commande reçue » et « vérifs ».

**CRÉER UNE AFFAIRE — et le piège du bouton qui existait déjà.**
Le bouton « + Nouveau chantier » de la barre du haut **crée une fiche dans IndexedDB seulement** :
invisible pour les collègues, effacée au prochain `SEED_VERSION`. Patrice avait raison de dire
qu'on ne pouvait pas le faire. **Ne pas le confondre avec le nouveau bouton de l'onglet Affaires.**

- **Une affaire naît SANS n° de chantier**, sous une clé `af_xxxxxxxx` tirée au sort **côté
  serveur** (une clé fournie par le navigateur permettrait d'écraser une affaire en devinant son
  identifiant). C'est l'arbitrage de Patrice du 04/09 : « un devis pas encore gagné n'a pas de n° de
  chantier et ne doit pas en recevoir » — 182 lignes sur 758 n'ont jamais donné de chantier.
- **`numeroter` DÉPLACE la ligne**, il ne la copie pas : laisser les deux vivre ferait apparaître
  l'affaire deux fois avec deux saisies qui divergent. Il **refuse** si le numéro visé porte déjà
  une saisie, si la source a déjà un vrai numéro, ou si le format est faux. Le journal est écrit
  **des deux côtés**, sinon l'historique paraîtrait commencer le jour de la numérotation.
- Le **prochain numéro libre** est calculé sur les fiches **ET** les affaires du fichier **ET** les
  affaires créées : le seul maximum des fiches redonnerait un numéro déjà pris dans l'Excel.
- **Une ligne de saisie sans `nom` n'est pas une affaire créée** — les 53 lignes de l'Excel ont une
  saisie sans nom, les compter comme créées les afficherait deux fois.

**Ce qui reste de mon côté, et pourquoi** : lire le devis, calculer le prévisionnel, relever boîtes
et tâches vendues, monter le dossier App Tech, créer la demande de photos. Ce ne sont pas des étapes
mécaniques — juger qu'une IST est bien de TELSAM *et* validée par RTE est ce qui évite d'envoyer un
technicien dans un poste où il n'est pas autorisé.

**Testé — 211 contrôles, 0 échec** : `test-api-affaires.html` (64), `bloc-test-saisie-affaires.html`
(88), `bloc-test-affaires.html` (59). Deux défauts trouvés par les bancs, **et deux par la capture
d'écran** (le 0 % muet, et le pavé d'avertissement qui repoussait le tableau hors de l'écran).
*Piège de harnais rencontré : une fausse base D1 qui répond la même ligne à toutes les requêtes ne
peut pas éprouver `numeroter`, qui en interroge deux (« la source existe ? » puis « la cible est
libre ? ») — mes quatre premiers ÉCHEC étaient des faux. D'où `firstPar`, qui répond selon les
paramètres.*

### LOT 4 — TOUT est numéroté (07/09/2026), et ce que ça change

**Demande de Patrice, qui renverse sa règle du 04/09** : « j'aimerais plutôt que l'on numérote tous
les chantiers, **même ceux qui sont perdus**, cela nous aidera pour les statistiques plus tard.
Ils seront comme cela tous répertoriés. »

**149 numéros écrits dans `SUIVI RTE TELECOM - mise a jour_3.xlsx`.** Le fichier passe de 476 à
625 lignes numérotées, et l'onglet Affaires de **53 à 131 affaires**. Surtout : **plus aucune
affaire vivante sans numéro** (il y en avait 14 le matin, 28 le 04/09) — c'était le but.

Méthode d'écriture : celle du 07/09 au matin, reprise telle quelle (sauvegarde dans le scratchpad,
script en pur ASCII, **tous les garde-fous joués AVANT la première écriture** — libellé et n° de
devis comparés sans accent, case cible réellement vide — chaque cellule **relue** après écriture,
`Save()` seulement à la fin). Contrôle après coup avec son contre-exemple : 625 numérotées, 3
laissées de côté, 0 converti en date, 0 doublon sur deux chantiers différents.

**TROIS LIGNES LAISSÉES SANS NUMÉRO, volontairement** : elles ressemblent à des TS d'affaires déjà
numérotées, et un numéro neuf scinderait l'affaire en deux. À arbitrer par Patrice :
`TELSAMCC25019` ARGIA-HERNANI, `TELSAMCC25129V2` CANTEGRIT, `TELSAMCC26111` DONZAC-LESQUIVE LOT 2.

**CE QUE ÇA CASSAIT, ET QUI EST CORRIGÉ :**
1. **`nextNumero()` aurait donné un numéro DÉJÀ ATTRIBUÉ.** Il prenait le maximum des seules
   fiches (71) alors que le fichier va maintenant à 149 : le bouton « + Nouveau chantier » aurait
   proposé `26-072`, et **deux chantiers auraient porté le même numéro** — définitif, un numéro ne
   se réutilise jamais. Il regarde désormais les fiches, `AFFAIRES_RTE` **et** les affaires créées.
   `afProchainNumero()` n'est plus qu'un appel à `nextNumero()` : **un seul endroit calcule**, deux
   fonctions qui font le même calcul finissent par diverger.
2. **« 78 numéros sans fiche » n'est plus une ANOMALIE, c'est le résultat voulu.** Le script criait
   78 fois ; il ne signale plus que celles qui **mériteraient** une fiche — gagnées et non soldées
   (3 le 07/09 : `26-077`, `26-080`, `26-086`) — et compte les autres en une ligne. Champ
   `sansFicheVivantes`. **Un contrôle bruyant finit ignoré.**
3. **La vue disait « anciens devis de DATA »** pour les préfixes écartés : faux depuis la
   correction ci-dessus. Elle dit maintenant « les années précédentes ».

**À SIGNALER À PATRICE, ET C'EST SA DÉCISION** : les 149 numéros n'ont été écrits que dans
`mise a jour_3`. Le classeur **maître ne porte plus que 25 des 131 numéros de l'année** — l'écart
entre ses deux fichiers s'est beaucoup creusé. La vue l'annonce (« 177 de ces numéros ne sont pas
encore dans… »), mais lequel garder reste son choix : ne pas consolider tout seul
(cf. [[feedback_fichiers_de_travail_de_patrice]]).

**Testé — 215 contrôles, 0 échec.** Cinq contrôles ont dû être RÉÉCRITS parce qu'ils encodaient
des chiffres devenus faux, et **deux d'entre eux échouaient sur un résultat réussi** : « les
affaires vivantes sans numéro sont annoncées » (il n'y en a plus) et « 53 affaires » (il y en a
131). *Un test qui vérifie l'ancien état fait échouer le banc sur un succès — le corriger fait
partie du travail, le contourner ne le fait pas.* Le troisième était un faux ÉCHEC de ma part :
« dates vides en dernier » jugeait le vide sur la chaîne brute alors que le tri le juge sur la clé
de tri, si bien qu'une date illisible (« ? ») comptait pour pleine.

### LE DÉFAUT LE PLUS GRAVE DE CET ONGLET — la virgule (07/09/2026)

**Trouvé par Patrice à l'œil, pas par un contrôle**, en relisant Cantegrit : « le V3, il n'y a pas
de montant. Le reste à facturer est de zéro. Comment se fait-il que tout soit mélangé ? »

**La cause.** Le XML d'un `.xlsx` écrit **toujours** ses nombres avec un **point** décimal, quelle
que soit la langue d'Excel. Cette machine est en **fr-FR**, où `[double]::TryParse` **sans culture
explicite attend une virgule**. `"112502.79"` était donc refusé et rendait `$null`, **en silence**.
Un entier passait, un montant à centimes non.

**L'ampleur, mesurée** : **278 montants de devis perdus sur 587** (8 935 298 €), plus 137
« Factu partiel » et 135 « Restant à facturer ». Cantegrit affichait **3 345 € de devis pour
58 345 € facturés** ; Portet annonçait « aucun montant » alors que son devis vaut 136 219,55 €.

**Pourquoi c'était pire qu'une panne franche** : la vue montrait une donnée **à moitié juste**. Les
affaires à montants ronds étaient exactes, celles à centimes fausses — rien ne distinguait les deux.

**Le correctif, et ses deux moitiés :**
1. `NombreInvariant()` parse en **culture invariante**, toujours. *Ne jamais revenir à un
   `TryParse` à un seul argument dans ce fichier.*
2. **Il faut accepter les DEUX séparateurs.** Le garde-fou ci-dessous a trouvé, dès son premier
   passage, la ligne 578 : `"21215,60"` avec une **virgule** — celle-là est une chaîne tapée à la
   main par Patrice, pas un nombre du XML. Une lecture qui n'accepterait que le point la perdrait
   aussi sûrement.
3. **GARDE-FOU** : toute case d'argent **non vide** qui ne donne pas un nombre est comptée
   (`argentIllisible`) et le script **crie en rouge**. Une règle sans mécanisme ne tient pas.

**CE QUE LA CORRECTION A FAIT DISPARAÎTRE, ET C'EST LA PREUVE :**
- **Givors 26-060** n'est plus « en écart ». Depuis le 04/09 la vue signalait « 49 h passées alors
  que le fichier ne donne aucun devis gagné », et **j'avais conclu que la cause était un devis TS —
  c'était faux**. La vraie cause est la règle suivante.
- **Cantegrit** n'affiche plus « facturé au-delà du devis chiffré » : l'écart n'existait que parce
  que le devis de 112 502,79 € était perdu.

### CE QUI FAIT QU'UN DEVIS EST ACCEPTÉ : LA COMMANDE (règle de Patrice, 07/09/2026)

« Tu devrais faire la somme de tous les devis en prenant compte le devis accepté. **Je suis basé sur
le numéro de commande.** » Je ne regardais que la colonne « statut » = `GAGNE`.

Mesuré sur les lignes 26- : **58 portent une vraie commande, 48 seulement sont marquées GAGNE**.
**14 devis réellement commandés étaient donc ignorés** — dont Givors, dont la ligne porte la
commande `4500810332` avec une colonne statut vide.

- `afAccepte(d)` : **pas clos** ET (`cdeRte` ou `cdeCli` renseigné **OU** statut `GAGNE`).
- **L'union, pas la commande seule** : 4 devis sont marqués GAGNE sans commande arrivée
  (« Attente avenant », « Devis envoyé »). Les écarter perdrait ce que son fichier affirme.
- **Un devis clos reste clos.** Vérifié : aucune ligne PERDU/REFUS/ANNULE ne porte de commande, donc
  les deux critères ne se contredisent jamais aujourd'hui — mais si cela arrivait, « clos » gagne.
- **La commande se lit dans `cdeRte`/`cdeCli`, jamais dans la case brute** : celle-ci contient
  souvent une phrase (« Devis envoyé »), et `afRenseigne` la prenait pour une valeur — l'affaire
  passait en « commande reçue » alors que la case disait l'inverse.

**« Facturé en partie » est devenu un statut à part** (13 au lieu de 12). Cantegrit sortait
« Facturé » avec 57 502,79 € encore à facturer : le vert mensonger, sur de l'argent.

**« Perdu mais soldé »** — sa question. Sa colonne « Soldé O/N » vaut OUI sur un devis PERDU (c'est
ainsi qu'il clôt une ligne perdue) et la fiche affichait « soldé », qui veut dire *payé*. On écrit
désormais **« clos »** sur un devis fermé.

**LEÇON DE MÉTHODE, LA PLUS IMPORTANTE DE LA JOURNÉE.** Le banc d'essai passait au **vert** sur ces
chiffres faux : ses contrôles avaient été écrits le 04/09 **à partir de ce que la page affichait**
(« total = 3345 », « reste = 0 », « son devis gagné n'a AUCUN montant »). Un test rédigé depuis la
sortie observée ne teste rien — il **fige le défaut** et lui donne l'apparence d'une règle. Huit
contrôles ont dû être réécrits, et leurs valeurs attendues viennent maintenant **du classeur**.
Deux garde-fous ajoutés : les montants à centimes sont lus (62 affaires en portent), et le V3 de
Cantegrit vaut bien 112 502,79 € — écrit en dur, pour que la disparition crie.

*Piège de banc d'essai rencontré ici :* une exception dans un `(async function(){…})()` devient une
promesse rejetée **avalée en silence** — le harnais restait sur « en cours » sans rien dire. Et mon
extraction du verdict reprenait le mot `verdictTest` présent aussi dans le source du script, donc
elle affichait le code au lieu du résultat : **j'ai cru à un plantage qui n'existait pas**. Lire le
premier bloc `<pre>` et lui seul (`partage/lire-verdict.sh` dans le scratchpad).

### L'AUDIT QUE PATRICE A DEMANDÉ (07/09/2026) — et ce qu'il a trouvé

Sa phrase, après le défaut de la virgule : **« Non, je pense que c'est toi qui vas regarder si tu
n'as pas fait d'autres erreurs. »** J'avais terminé en l'invitant à continuer de vérifier — c'est
lui rendre mon travail. **Ne pas refaire ça** : après un défaut trouvé par lui, l'audit est à moi.

**LA MÉTHODE QUI A MARCHÉ : recalculer par un chemin indépendant, et comparer.**
Les 131 affaires ont été recalculées **depuis le classeur** par un script qui ne partage **aucune
fonction** avec celui de production, puis comparées à ce que la page publie, sur quatre dimensions
(montant HT, facturé, reste, nombre de devis retenus). **0 écart** — et le contre-exemple prouve que
la comparaison sait dire non. `scratchpad/audit.ps1` + `compare.html`, à reprendre tels quels.

**CE QUE L'AUDIT A TROUVÉ DANS MON CODE :**
1. **`[double]'112502,79'` rend 11250279**, sans erreur. Le CAST PowerShell utilise la culture
   invariante, donc **une virgule y est lue comme un séparateur de milliers** : le miroir exact du
   défaut du matin, en pire (une valeur cent fois trop grande au lieu d'une valeur absente).
   Vérifié : les trois autres sites de cast du dépôt reçoivent des nombres, jamais des chaînes.
2. **`parseFloat('12 500,50')` rend 12** côté page — un espace suffit à diviser un montant par
   mille. D'où `afNombre()`, qui normalise et **rend `null` plutôt que `NaN`** (un NaN se propage en
   silence dans les additions et ressort en « — » sans qu'on sache pourquoi).
3. **Une phrase prise pour un numéro de commande.** `« Projet en 2028 »` contient une année, donc la
   séparation en faisait une référence client — et **depuis que la commande vaut acceptation, cela
   comptait le devis et son montant**. Règle ajoutée : un segment qui contient un espace **et** un
   mot de deux lettres ou plus est du texte. Contre-exemples qui doivent rester des références :
   `920086240 / 913053066?`, `913149925 / Q.0128590.3.40`, `F-26-0892`.
4. **Mon propre outil d'audit portait le défaut qu'il cherchait** : `Export-Csv` écrit `62003,5` en
   fr-FR, et mon comparateur JavaScript lisait `parseFloat('62003,5') = 62003` — **29 faux écarts**.
   Un outil de contrôle doit être écrit en format neutre, sinon il mesure sa propre langue.

**CE QUE L'AUDIT A TROUVÉ DANS LA DONNÉE — dit, jamais arbitré :**
- **4 affaires portent « perdu / annulé » dans la case du NUMÉRO DE COMMANDE**, colonne statut vide
  ou « EN ATTENTE » : `26-091`, `26-123`, `26-126`, `26-145`. Elles restent donc comptées comme
  vivantes. **On ne les ferme pas tout seul** — le mot est dans la mauvaise colonne, et clore une
  affaire sur une note serait exactement le genre de décision qui ne m'appartient pas.
- **1 incohérence de montants** : `26-047`, 11 250 € − 0 € devrait faire 11 250 €, la case
  « restant à facturer » dit 9 800 €. Les 103 autres lignes vérifiables sont cohérentes.
- 12 lignes portent une **année** au lieu d'une date prévisionnelle (`2028`, `Fin 2028`,
  `2027 ou 2028`) : affichées telles quelles, triées avec les cases vides.

État après audit : 15 écarts signalés sur 131 affaires, 0 affaire acceptée sans montant, 0 montant
négatif. **228 contrôles, 0 échec.**

**RESTE À FAIRE, et à ne pas commencer sans Patrice** : les 4 questions du 04/09 encore ouvertes
(l'onglet remplace-t-il l'appli Facturation ? qui est responsable de quelle colonne ? quand
arrête-t-on l'Excel ? les affaires vivantes non numérotables), plus la réponse sur le devis TS de La
Curbans. **L'ordre de tri met les affaires les plus avancées en tête** (héritage du lot 1) : à lui
demander s'il préfère l'inverse, ce sont les affaires en attente qui demandent une action.

### Testé — `partage/bloc-test-affaires.html`, 49 contrôles, 0 échec

Concaténé à la vraie page et joué en Chrome `--headless` (`chantiers` rempli depuis `SEED_DATA`,
IndexedDB ne répondant pas sous `--virtual-time-budget`). **Chaque contrôle a son contre-exemple** :
le devis PERDU de 65 100 € n'entre pas dans les montants **et** les gagnés y entrent bien ; Givors
est signalé en écart **et** Lacanau (aucune activité) ne l'est pas ; « à priori non concerné » n'est
pas une valeur **et** « signé le 12/09 » en est une. Les deux défauts des points 4 et 5 ci-dessus
ont été trouvés par ce banc d'essai.

**Pièges de rendu respectés** : rien en `onclick=` dans le HTML (les libellés portent apostrophes et
accents, et `escapeHtml` de ce fichier n'échappe pas les guillemets) ; les lignes ne portent qu'un
**index** résolu via `afAffairesAffichees`, posée **avant** le rendu ; la recherche rend le focus et
la position du curseur, sinon le champ le perd à chaque lettre.

### `controleDifferre` — un arbitrage de Patrice doit faire TAIRE le contrôle (07/09/2026)

**Les DEUX lignes prioritaires du contrôle de 07h55 étaient des faux positifs**, et chacune avait sa
raison déjà écrite dans la donnée :

- **26-009 Aure-Lannemezan** : une alerte de la fiche disait mot pour mot, depuis l'arbitrage de
  Patrice du 04/09/2026, « *ne pas traiter cette ligne du contrôle des chantiers comme un manque à
  combler* » — le PDP n'entre dans App Tech que lorsque l'intervention est datée, et le planning ne
  place personne (seule la couverture PGO 07/09→09/10 rendait la ligne prioritaire).
  **Le contrôle criait quand même : il ne lit pas les alertes.**
- **26-071 Bollène-Plantades** : `mesureTouret = true`, et la note de la fiche comme le libellé du
  planning concordent (« mesures de 3 tourets le lundi matin, déroulage par Bouygues à partir de
  mardi »). Ce fichier dit déjà qu'un statut `nc` est **normal** en phase mesure de tourets, et
  TELSAM ne revient en poste qu'au raccordement des 7 WTC2, après 8 semaines d'appro.

**Un contrôle qui crie après un arbitrage déjà rendu apprend à ne plus être lu** — c'est exactement
le mal que ce script existe pour éviter, et le remède est un mécanisme, pas ma mémoire.

**Le champ, sur la fiche du suivi** : `controleDifferre: { raison, jusqua }`. `jusqua` au format
`JJ/MM/AA`, vide si la reprise n'a pas de date connue.

**QUATRE BORNES, chacune nécessaire :**
1. **Il est posé À LA MAIN, jamais déduit** — même règle que `perimetre.pdpSurPlace`. Un différé est
   une décision de Patrice ; je ne fais que l'inscrire quand il l'a rendue.
2. **Un différé daté qui expire redevient prioritaire**, avec la mention de son échéance dans les
   manques. Sans ça un différé oublié masquerait un vrai manque pour toujours — le même piège qu'un
   arbitrage `TECH_RANGES` qui survit à la ligne du planning.
3. **Un différé ne disparaît pas du rapport** : le chantier sort des prioritaires et entre dans la
   liste `differes` de `controle-chantiers.json` (nouveau compteur `differes`). Le rappel de
   `hook-veille-prompt.ps1` en donne **le nombre en une ligne**, seulement quand il y a par ailleurs
   un vrai prioritaire — on ne cache jamais sans le dire, mais on ne remplace pas un bruit par un
   autre.
4. **Zéro prioritaire ⇒ aucun rappel**, même s'il reste des différés : ils vivent dans le récap du
   matin, pas à chaque message.

**Vérifié dans les deux sens le 07/09/2026** : état réel ⇒ 0 prioritaire, 2 différés, et le hook ne
sort aucun bloc chantiers ; puis **contre-examen sur une COPIE** (`-Base` dans le scratchpad, différé
de 26-071 antidaté au 01/09) ⇒ 26-071 redevient prioritaire avec « le differé du contrôle est échu »,
la ligne « pour mémoire » nomme 26-009, et le vrai `controle-chantiers.json` n'a pas bougé.

**PIÈGE DE BANC D'ESSAI rencontré ici** : `hook-veille-prompt.ps1` écrit ses octets **directement sur
la sortie standard** (c'est voulu, cf. le défaut d'encodage du 04/09) — donc `& .\hook… | Out-String`
capture **une chaîne vide** et tout contrôle par expression régulière dessus paraît échouer. Mes deux
premiers « ECHEC » étaient donc des faux : la sortie brute affichée à l'écran prouvait l'inverse.
**Lire la sortie réelle, ne pas conclure sur une capture vide.**

## Trois demandes de Patrice du 07/09/2026 — congé paternité, cible du crayon, et le côté « glamour »

Ses mots : « il faudrait rajouter les congés paternité… un technicien n'a pas pu remplir ses congés
car il n'y a pas la possibilité dans l'appli » ; « il est très difficile d'appuyer sur la petite
croix quand on veut créer une note en bas à droite des cases » ; « ce suivi n'est pas très
glamour… j'aimerais que les bulles fassent peut-être un peu plus bulle, quelque chose de plus fun
qui donne envie de l'utiliser… c'est important pour que nous puissions travailler sereinement ».

### 1. Congé paternité — LE MENU N'ÉTAIT PAS LE PROBLÈME, LA FEUILLE OFFICIELLE L'ÉTAIT

**Le modèle Excel de feuille d'heures n'a AUCUNE ligne « paternité ».** Ses étiquettes sont
imprimées dans le classeur (`sharedStrings.xml` relu : Formation, Administratif, Réunion
hiérarchique, Contribution agence, Qualité, Congés, RTT, Férié, Arrêt Maladie, Accident du travail,
Garde d'enfants, Chômage — et rien d'autre). On ne peut donc pas en inventer une sans que le modèle
change du côté de Karine et Pierre.

**CE QUI A ÉTÉ ÉCARTÉ, ET POURQUOI** — les rangs 25 à 38 du modèle sont des lignes chantier libres,
sans étiquette : y écrire « CONGÉS PATERNITÉ » aurait été tentant. C'est faux, et faux en silence :
la cellule `B45` du modèle calcule l'**Activité** par `SUM(R12:R17)+SUM(R25:R38)`, donc ces heures
d'absence auraient été comptées comme du **temps travaillé**, et le récap FH aurait vu apparaître un
chantier fantôme sans numéro de devis.

**CE QUI EST FAIT** : les heures partent sur la ligne **« Congés » (ligne 18)**, la seule case
d'absence payée existante, ET la mention part **automatiquement** dans les commentaires de la
feuille (`J47`), devant ce que le technicien a écrit :
`CONGÉS PATERNITÉ (à ne pas décompter des congés payés) : 21 h (Lun, Mar, Mer) — <son texte>`.

**La mention n'est pas décorative, c'est elle qui empêche l'erreur de paie.** Sans elle, Karine
verrait des heures sur la ligne « Congés » et décompterait des CP à tort. Elle est **calculée**,
jamais tapée : un technicien qui oublierait de l'écrire ne peut pas provoquer l'erreur.

**LE DÉFAUT QUE ÇA CRÉAIT, ET LE CORRECTIF** : deux codes de l'appli visent désormais la MÊME ligne
du modèle. `applyPayloadToSheetXML` écrivait absence par absence, donc la seconde **écrasait** la
première — 3 h de CP + 7 h de paternité le même lundi donnaient 7 h, et 3 h disparaissaient sans
rien signaler. Les heures sont maintenant **additionnées par cellule avant d'être écrites**
(`totauxAbsences`). Ne pas revenir à une écriture absence par absence.

Mécanique : un code d'`ABSENCE_CODES` peut porter `mention` (ce qui part dans les commentaires) et
`aide` (la phrase montrée au technicien sous le titre de l'absence, pour qu'il ne s'étonne pas de ne
pas voir « paternité » sur le document officiel). **Le jour où le modèle gagne une vraie ligne
« paternité », il suffit de changer `row` et de retirer `mention`.**

*À signaler à Patrice, c'est sa décision et pas la mienne : demander à Karine si elle préfère une
ligne dédiée dans le modèle. Tant qu'elle n'existe pas, la mention est le meilleur compromis —
visible par un humain, et jamais silencieuse.*

Le planning Teams, lui, connaît déjà le cas : `CP Paternité` y est en rouge (ColorIndex 3) et la
grille du suivi le lisait déjà (`/^(CP|RTT|cong|c pater|AM|recup|VM|Visite Med)/i`).

### 2. La cible du crayon de note — mesurée, pas estimée

Elle faisait **15 × 16 px**, pastille comprise : sous les 24 px conseillés pour une cible, et il
fallait viser au pixel dans une case de 40 px de haut. Elle fait maintenant **35 × 27 px**, soit
quatre fois la surface, **sans agrandir la pastille visible** (23 × 19 px) — un bouton plus gros
répété 190 fois à l'écran referait le damier illisible que le survol évite.

Le moyen est un pseudo-élément transparent (`.plNoteBtn::after`) : il appartient au bouton pour le
pointeur mais ne se voit pas. Il déborde **vers la gauche et vers le haut uniquement**, pour ne pas
empiéter sur la croix de retrait (`.plX`), qui vit dans le coin opposé — mesuré, 20 px les séparent.

**La mesure ne se fait PAS sur le rectangle dessiné** : on balaie la zone point par point avec
`elementFromPoint`, c'est-à-dire on demande au navigateur qui recevrait le clic. Deux harnais
versionnés, chacun avec ses contre-exemples : `partage/bloc-test-cible-note.html` (la cible est
grande, ET le haut de la case appartient toujours à la bulle, ET le coin bas-gauche à la case) et
`partage/bloc-test-cible-croix.html` (la croix de retrait reçoit toujours ses propres clics).

**PIÈGE DE BANC D'ESSAI, à connaître avant d'en écrire un autre** : `elementFromPoint` rend `null`
pour tout point **hors de la fenêtre**. Sans `scrollIntoView` préalable, mes trois premiers
contrôles paraissaient échouer tandis que les deux contre-exemples paraissaient **passer** — sur du
vide. Un contrôle qui interroge un point hors écran ne prouve rien.

### 3. « Que les bulles fassent un peu plus bulle »

Une demande de goût, traduite en trois choses précises : coin plus rond (11 px au lieu de 6),
**reflet** (dégradé clair en haut, ombre en bas) qui donne le volume, et **soulèvement au survol**
(`translateY(-1.5px) scale(1.02)`). Les pastilles de la réserve suivent le même registre — ce sont
les mêmes objets, on les déplace d'un endroit à l'autre.

**LE PIÈGE À NE PAS RÉINTRODUIRE : la couleur du chantier se pose en `background-color`, JAMAIS en
raccourci `background`.** Le reflet est un `background-image` ; un raccourci l'effacerait et la
bulle redeviendrait plate sans que rien ne le signale. Vaut pour le rendu des cases, des pastilles,
et pour `.plAbsence .plBulle` / `.plTexte .plBulle`.

### 3 bis. La palette du planning — CE QUE LA MESURE A DIT, ET UN MÉCANISME RETIRÉ

Patrice : « il y a des couleurs comme le violet et le bleu qui sont très proches ». Écart perceptif
(ΔE, sur les 120 paires de la palette) :

| | paire la plus proche |
|---|---|
| ancienne palette | **ΔE 6,8** — le cyan `#1f7f9f` et le pétrole `#0f6f8a` étaient à peu près la même couleur ; puis bleu/indigo 15,6 ; violet/pourpre 17,3 |
| nouvelle palette | **ΔE 18,2**, et cinq teintes seulement dans le secteur bleu → violet, contre sept avant |

**La paire qu'il a citée mesurait ΔE 28,6 — donc PAS la plus proche du lot.** Ce n'est pas lui qui
se trompe, c'est la mesure : ΔE sous-estime les écarts de teinte dans les bleus. Le chiffre sert à
écarter les cas manifestes, **il ne remplace pas son œil** (cf.
[[feedback_verifier_a_l_ecran_pas_dans_le_repere]]).

**UN MÉCANISME ESSAYÉ, MESURÉ, ET RETIRÉ — ne pas le refaire.** J'avais ajouté une notion de
« famille » de teinte pour que la répartition évite de placer deux voisines dans la même quinzaine.
Mesuré sur les 50 quinzaines de l'année : **36 paires de même famille avant, 35 après.** Aucun gain,
et pour cause — huit familles ne suffisent pas quand une quinzaine porte jusqu'à 14 chantiers
simultanés, le principe des tiroirs impose la réutilisation. Ça coûtait en revanche de la stabilité
(46 chantiers déplacés de leur couleur habituelle au lieu de 24). **C'est la palette qui règle le
problème, pas la répartition.** Le glouton est donc revenu à l'identique : préférence, puis première
teinte libre. Contrôle inchangé : **0 collision dans la grille sur 37 quinzaines**.

Les 16 teintes portent toutes du texte blanc à ≥ 4,5:1 — à revérifier avant d'en éclaircir une.
Les 8 teintes de secours sont plus proches des 16 premières (paire la plus proche des 24 : ΔE 10,8)
et **c'est inévitable** : après seize teintes bien réparties il ne reste plus de place franche dans
le registre « assez foncé pour porter du blanc ». Acceptable là et nulle part ailleurs — elles ne
servent qu'aux pastilles de la réserve, qui portent leur nom écrit dessus.

### 3 ter. L'onglet Affaires — des teintes VIVES qui s'ajoutent, sans toucher aux alertes

« Avec ses couleurs pastels, ça reste quand même un petit peu triste. »

**Les variables `--vif-*` sont AJOUTÉES, pas substituées.** Les `--*-bg` / `--*-border` portent les
**alertes** du suivi, où le rouge et l'ambre ont un sens précis : les saturer aurait rendu la page
criarde là où elle doit rester calme, et affaibli le contraste des vrais avertissements. Les
`--vif-*` ne servent qu'à ce qu'on **lit** (pastilles d'étape, jauges, tuiles), jamais à ce qui
**alerte**.

- **le liseré gauche de 4 px des listes déroulantes de statut** porte la couleur pleine. Le fond
  pastel reste nécessaire — c'est un menu, on y lit du texte toute la journée — mais il ne suffisait
  pas à distinguer douze étapes : d'une ligne à l'autre, quatre tons pâles se confondaient. **Ne pas
  remplacer le fond pastel par de la couleur pleine**, le texte deviendrait illisible ;
- **chaque tuile de filtre porte la couleur de SON étape** (`afTu-<cls>`, la classe de couleur que
  le statut porte déjà — on ne réinvente pas de correspondance ici, les deux finiraient par se
  contredire). La rangée devient du même coup la légende des couleurs du tableau, au lieu d'être
  huit cadres identiques ;
- jauges du cycle en dégradé (4 px, toujours « visuel, efficace et discret » comme il l'avait
  demandé), survol de ligne teinté de bleu et non de gris (le gris se confondait avec l'en-tête).

`.afEtape` n'est plus utilisé par aucun rendu depuis que la liste déroulante a remplacé la pastille
(lot 2) — CSS mort, laissé en place, à retirer un jour.

### Vérifié

**Suivi : 223 contrôles, 0 échec** (`bloc-test-couleurs` 9, `bloc-test-affaires` 74,
`bloc-test-saisie-affaires` 90, `bloc-test-notes` 22, `bloc-test-validation-suivi` 16, plus les deux
harnais de cible 6 + 6), syntaxe des deux blocs `<script>` compilée.
**Appli : 17 contrôles** dans `scripts/test-conges-paternite.html`, 0 échec.
**Et à l'écran**, ce que ni l'un ni l'autre n'aurait montré : la grille avant/après, la vue Affaires,
et la feuille d'heures d'un technicien avec « Congés paternité » et « Congés » côte à côte.

### DEUX HARNAIS DE L'APPLI ÉTAIENT DÉJÀ ROUGES AVANT CETTE SESSION — dont un qui compte

Vérifié en les rejouant sur le `index.html` du dernier commit : **mêmes échecs**, ils ne viennent pas
de cette session.

1. `test-planning-appli` — « la semaine +2 est servie aussi » attend ≥ 6 personnes, en trouve **0**.
   **Ce n'est pas un défaut** : le planning Teams est **vide à partir du 21/09** (compté :
   13 cases le 07/09, 13 le 14/09, 0 ensuite). C'est la vérité du fichier, et le contrôle encode un
   chiffre figé quand S38 était la semaine +2. **À réécrire pour qu'il compare à ce que le bloc
   publié place réellement, pas à un nombre gravé.**
2. `test-planning-appli` — « bloc neutralisé : repli complet sur `TECH_RANGES` » attend ≥ 8, trouve
   **3**. **CELUI-LÀ EST UNE VRAIE INFORMATION, à porter à Patrice.** Le filet existe pour qu'un
   échec de `planning-rte.ps1` ne laisse pas les techniciens devant « rien de prévu ». Or les
   arbitrages `TECH_RANGES` ont été retirés à mesure que le planning portait les numéros — règle
   qu'il a lui-même posée le 04/09 — donc **le filet ne couvre plus que 3 personnes sur 11**. Le
   filet s'est aminci comme conséquence directe d'une bonne règle. **C'est à lui de décider** s'il
   veut le regarnir ; ne pas abaisser le seuil du test en silence, ce serait effacer le signal.
3. `test-validation` — « base injoignable : on le DIT au technicien » : le harnais ne joue pas la
   branche `file:` ajoutée le 04/09. Test à mettre à jour, pas de défaut.

**Un banc d'essai rouge en permanence finit ignoré, exactement comme un contrôle qui crie au loup.**
Ces trois-là sont à solder.

## Veille du 08/09/2026 — l'IST indice B de Cantegrit, et deux chantiers sans fiche

### 26-003 Cantegrit — IST INDICE B : reçue, TOUJOURS PAS VALIDÉE, et l'échéance est lundi

Indice B déposé le 07/09/2026 à 15h47 (`.docx` 15h43 + `.pdf` 15h47, 13 pages) dans
`Postes\Cantegrit 26-003\Documents Telsam`. C'est bien une IST **TELSAM** (29 occurrences sur le
texte compacté), rédigée par TELSAM, C. CAZENAVE et Pascal BONAVENTURE chargé de travaux. Elle
remplace l'indice A du 31/08.

**ELLE N'EST PAS VALIDÉE PAR RTE, et cette fois c'est MESURÉ, pas « apparemment ».** Le tableau
« Validation RTE » (section i, page 13) a été lu **cellule par cellule via Word COM** sur le
`.docx` : 6 lignes, en-tête `Date | Nom | GMR | Modifications | Signature`, puis **5 lignes vides
sur 5**, et surtout **ZÉRO image dans ce tableau** — donc pas même une signature scannée, ce que
l'extraction de texte seule n'aurait pas pu exclure.

> **LA MÉTHODE À REPRENDRE POUR TOUTE IST** : ouvrir le `.docx` en lecture seule par Word COM,
> retrouver le tableau par son en-tête (`GMR`), lire ses cellules, **et compter
> `tb.Range.InlineShapes.Count`**. Le compte d'images est ce qui transforme « le texte est vide »
> en « il n'y a rien, texte ni signature ». Sur le PDF seul on ne peut que constater l'absence de
> texte, et une IST signée à la main aurait une signature en image.

Conséquences appliquées : `ist.indice` porté de A à **B** dans le suivi (l'appli ne stocke pas
l'indice IST, rien à y changer), `valideRTE` reste **false**, **rien copié dans App Tech**, alerte
réécrite en URGENT — le PGO ind.6 exige une IST sur la phase **14/09 → 02/10**, qui commence
**lundi 14/09 (semaine 38)** avec trois techniciens placés. `SEED_VERSION` v150.

### Deux PGO sur des chantiers qui n'ont AUCUNE fiche — décision de Patrice

**CRENEY - TROYES — c'est le point à traiter.** `PGO_20260703_RTE_TROYES-EST_ind5.xlsx`, **V5 du
07/09/2026**, dont l'objet de la mise à jour est littéralement « **ajout date TELSAM** ».
Opération : « RTE - TROYES EST - Remplacement du poteau béton n°29 ». Ligne 47 de l'onglet `PGO` :

| début | fin | entreprise | rôle | consignation | intervention |
|---|---|---|---|---|---|
| 28/09/26 | 09/10/26 | **TELSAM** | A1EC HT&BT | oui | Contrôle avant et après pose du canton pyl 28-pyl 29, **pose de boîtiers WTC2** |

Aucune fiche n'existe, donc : pas de numéro, pas de boîtes relevées, pas de prévisionnel, pas de
dossier App Tech, pas de tâches vendues. **Créer la fiche et attribuer un numéro est une DÉCISION**
(règle des noms provisoires + le numéro se communique à Patrice) : signalé, rien créé.

**CHAUSSE - REVIGNY — à confirmer, et je ne peux pas trancher seul.** Deux PGO du 07/09 à 17h50
(`SEPT 2026` et `OCT DEC 2026`), ligne 225 kV CHAUSSEE REVIGNY. TELSAM figure dans la légende des
entreprises (avec SEMI France, EGERI, DESSOLIN) et le bloc porte « Date révision : 07/09/2026 ».
Le PGO octobre-décembre contient une phase **TRAVAUX OPTIQUES** — déroulage / réglage / ancrage au
portique, raccordements optiques, test liaison optique, **test réflectométrie des tourets** — qui
est notre métier.

> **MAIS l'extraction de texte d'un PDF en TABLEAU ne préserve pas la correspondance
> ligne ↔ entreprise.** L'ordre de lecture du PDF n'est pas l'ordre visuel : les phases sortent
> groupées (`ASSEMBLAGES LEVAGES | SEMI SEMI SEMI | TRAVAUX SUR CABLES | SEMI SEMI SEMI | TRAVAUX
> OPTIQUES`) et la colonne entreprise de la phase optique ne sort pas. Je **ne peux affirmer ni
> que ces lignes sont pour TELSAM, ni le contraire.** À confirmer par Patrice ou par la source
> Excel du SPS. Ne pas conclure sur le comptage d'occurrences seul : TELSAM y apparaît 1 fois,
> mais SEMI France aussi, et SEMI a manifestement des lignes.

Faux positifs déjà tranchés, pour mémoire : PGO Casteljaloux ind.6.1 (ma propre copie du 07/09),
IST `APS3T5108` Cantegrit (celle d'INEO, règle du 27/08), PPSPS `3H223` Cross-Sausset (celui
d'INEO, question au SPS toujours ouverte), NDS Fleyriat (7e jour du même `.docx` identique au
`.pdf`).

## La palette du planning, TROISIÈME passe (08/09/2026) — et la mesure qui a débloqué le sujet

Patrice, après la deuxième passe : « les couleurs, je les aime bien, j'aime bien le léger dégradé
et le petit effet quand on passe la souris. Le problème, c'est qu'**on a encore trop de couleurs
proches côte à côte, des bleus et des violets qui se chevauchent**. Il faudrait vraiment
**trancher** pour qu'on ait du bleu à côté du rouge ou du jaune. »

### CE QUE J'AVAIS PRIS POUR UNE CONTRAINTE, ET QUI N'EN ÉTAIT PAS

Les deux premières passes cherchaient **seize** teintes distinctes, parce que ce fichier annonçait
« jusqu'à 14 chantiers simultanés dans la grille ». **Remesuré le 08/09/2026, quinzaine par
quinzaine sur les 50 de l'année : le maximum réel est 10, et la MÉDIANE est 5.** Une seule
quinzaine à 10, trois à 9, quatre à 8 — **90 % des quinzaines tiennent en 8 couleurs ou moins.**
*Le chiffre de 14 écrit plus haut dans ce fichier est donc périmé : c'est 10.*

Or seize teintes « assez foncées pour porter du blanc » ne peuvent PAS être toutes éloignées :
c'est ce qui entassait cinq à sept tons dans le secteur bleu → violet. **La contrainte n'était pas
la palette, c'était mon hypothèse sur le besoin.**

### DIX TEINTES FRANCHES, ET LE ROUGE ET L'OR RÉADMIS

Écart perceptif minimum (ΔE) de la palette : **6,8** (origine) → **18,2** (2ᵉ passe) → **27,7**
(celle-ci). Dix couleurs qui portent chacune un nom que tout le monde emploie : rouge, or, olive,
vert, bleu, violet, pourpre, magenta, brun, ardoise.

**Le rouge et l'or sont réadmis à la demande explicite de Patrice, et ce n'est pas une régression
sur la règle qui les excluait.** Le registre sépare ce que la teinte ne sépare pas : une absence
(CP, CP paternité) est un fond **pâle** à texte foncé, une note un fond **pâle** ambre ; un
chantier est une bulle **saturée** à texte blanc. **Si un jour une case CP se confond avec une
bulle rouge, c'est le fond pâle qu'il faut éclaircir, pas le rouge qu'il faut retirer.**

**L'ordre de `PALETTE_PLANNING` n'est pas décoratif** : la couleur préférée est `rang % 10`, donc
deux chantiers de numéros voisins prennent deux entrées voisines de la liste — elle alterne
chaud/froid pour que ces deux-là soient toujours franchement différents. Ne pas la retrier.

### LA MESURE QUI A TROUVÉ LE DÉFAUT RESTANT — et qu'il faut refaire à chaque retouche

Un minimum de palette élevé ne prouve rien si les deux teintes les plus proches tombent toujours
ensemble. Le contrôle qui compte est donc : **parmi les couleurs qui se retrouvent RÉELLEMENT dans
la même quinzaine, quelle est la paire la plus proche ?** (580 paires sur les 50 quinzaines.)

| | pire paire à l'écran | paires sous ΔE 25 |
|---|---|---|
| premier jet du 08/09 | **16,4** (vert `#17805A` / turquoise `#00857A`) | 19 sur 580 |
| après réordonnancement des secours | **23,4** (rouge brique / orange brûlé) | **6 sur 580** |

La cause du 16,4 : un chantier qui court sur plusieurs mois voisine avec **plus de dix** étiquettes
au total, épuise les dix franches, et prenait alors la PREMIÈRE teinte de secours — qui était un
turquoise voisin du vert. **`PALETTE_SECOURS_PLANNING` est donc ordonnée par un glouton max-min** :
à chaque rang, la teinte dont la distance minimale au noyau ET aux secours déjà placées est la plus
grande (36,7 puis 31,7 puis 26,2…). La grille ne consomme que ses premières entrées : ce sont elles
qui doivent être franches. **Ne pas retrier cette liste par goût.**

### CE QUE LE GRAPHE EXIGE VRAIMENT — 24 teintes, et pourquoi

`_grapheTeintes.maxPrises` = **24** : autour de l'étiquette la plus contrainte, 24 teintes sont
déjà prises. Ce n'est pas le maximum par quinzaine (10) qui dimensionne la palette, c'est le
**degré du graphe** : un chantier qui revient toute l'année a jusqu'à **158 voisins** (`c_y2edi4th`),
et comme il garde UNE couleur pour l'année entière, tous doivent différer de lui. Le nombre
chromatique dépasse donc largement la clique maximale. C'est pour ça que la grille descend dans
les secours et qu'il reste 6 paires sous 25.

> **LA SORTIE EXISTE, ET PATRICE L'A DÉJÀ AUTORISÉE SANS LE SAVOIR.** Sa règle du 03/09/2026 était
> plus étroite que ce que le code applique : « ce n'est pas grave de les changer quand on
> réintervient sur un chantier un ou deux mois plus tard, **mais ça prête à confusion de les
> changer d'une semaine sur l'autre** ». Le code, lui, fige une couleur par chantier **pour
> l'année**. En colorant par **plage continue de semaines** au lieu de par chantier, le graphe
> redevient un graphe d'intervalles : son nombre chromatique retombe à la clique maximale, soit
> **10** — exactement le noyau franc, et plus aucune descente dans les secours.
> Coût : un chantier qui revient après une interruption peut changer de teinte — ce qu'il autorise.
> Garde-fou obligatoire : **une nouvelle plage ne démarre qu'après un trou de 3 semaines ou plus**,
> pour que deux plages du même chantier ne soient JAMAIS visibles ensemble (la vue en montre deux).
> **À proposer à Patrice, pas à décider seul** : c'est sa règle de stabilité qu'on assouplit.

### Vérifié

**Suivi : 224 contrôles, 0 échec** — dont `bloc-test-couleurs` passé de 9 à 10 contrôles.
**0 collision de couleur dans la grille sur 37 quinzaines**, et la couleur ne bouge pas quand on
tourne les semaines. Prix payé, mesuré : 38 chantiers sur 58 changent de teinte une fois (24 avant,
conséquence attendue de `% 10` au lieu de `% 16`).

**UN CONTRÔLE A DÛ ÊTRE RÉÉCRIT, et c'est le cas d'école.** `bloc-test-couleurs` vérifiait que
Bollène et Fleyriat réclamaient la **même couleur préférée** (rangs 70 et 54 : `70 % 16 == 54 % 16`)
— c'était le contre-exemple qui prouvait que l'évitement travaillait vraiment. Avec dix teintes,
`70 % 10 ≠ 54 % 10` : ces deux fiches ne se disputent plus rien, et **le banc échouait sur un
succès**. Le contre-exemple cherche maintenant **dans la donnée** un couple qui partage sa couleur
préférée ET une quinzaine, puis vérifie qu'il a bien été séparé. Figer un couple d'identifiants
dans un test, c'est le condamner au prochain changement de palette.

## 26-117 Creney - Troyes, créé le 08/09/2026 — et cinq pièges payés d'un coup

Chantier apparu par la veille : le PGO ind.5 du 07/09 (V5, objet « ajout date TELSAM ») nous place
du **28/09 au 09/10**, et aucune fiche n'existait. Sous-traitance de **SEMI FRANCE** (commande
client F-26-0892, AST signé) sur l'opération RTE « TROYES EST — Remplacement du poteau béton n°29 ».
Devis TELSAM/CC/RTE/26100, 8 275 € HT. 2 WTC2 THYM à fournir et poser (pyl 28 et 29), 74 h / 9,3 j.

### 1. LE NUMÉRO EXISTAIT DÉJÀ, ET `nextNumero()` NE POUVAIT PAS LE SAVOIR

`nextNumero()` proposait **26-150**. Le chantier portait déjà **26-117** dans le fichier
commercial, parmi les 149 numéros que Patrice a écrits le 07/09. Créer 26-150 aurait scindé
l'affaire en deux, avec deux saisies qui divergent — et un numéro ne se réutilise jamais.

> **RÈGLE — avant d'attribuer un numéro à un chantier qui n'a pas de fiche, le CHERCHER dans
> `AFFAIRES_RTE`** (par libellé ET par n° de devis). L'absence de fiche ne veut pas dire l'absence
> de numéro : depuis le 07/09/2026, le fichier commercial numérote 131 affaires pour 72 fiches.
> `nextNumero()` sert à créer une affaire NEUVE, pas à retrouver celle qui existe.

### 2. TROIS ALERTES QUI NE DISAIENT RIEN DEPUIS CINQ JOURS — et c'était moi

`alertHtml()` lit `a.level` et `a.text`. En créant la fiche 26-071 Bollène le 03/09, je les avais
écrites `niveau` / `texte` : les trois alertes s'affichaient **`⚠️ ` suivi de rien** — dont
« aucun PDP ni PGO reçu à ce jour » sur un chantier actif. Corrigé, et un contrôle ajouté au
harnais : **aucune alerte ne doit manquer son champ `text`.** Il aurait crié dès le 03/09.

*Les 67 autres alertes du fichier étaient justes : c'est le genre de faute qu'une relecture ne voit
pas, parce que le JSON reste valide et la page ne plante pas.*

### 3. `perl -CSD -i -pe` AVEC DU TEXTE ACCENTUÉ DOUBLE L'ENCODAGE

« boîtiers » est parti en `boÃ®tiers` dans les **deux** dépôts, et « pylônes » en `pylÃ´nes`.
Le comptage systématique des `Ã` après écriture — et lui seul — l'a vu. Réparation **ciblée** sur
les deux séquences fautives, jamais une conversion 1252→UTF-8 globale qui aurait cassé les accents
sains des 71 autres fiches.

> **Pour écrire du texte accentué dans ces fichiers : l'outil Edit, ou `[IO.File]::WriteAllText`
> avec un `UTF8Encoding $false`.** Ni `perl -CSD -i`, ni `Set-Content`, ni un `.ps1` sans BOM.
> La méthode qui a bien marché ici : le script pose des repères ASCII (`ACCENT_NOM`, `ACCENT_AL1`…),
> puis l'outil Edit les remplace par le vrai texte. Contrôle final : `grep -c 'Ã'` = 0.

### 4. LE DEVIS EST EN POLICE EMBARQUÉE — et `OLEFormat.Object` est NUL en lecture seule

L'extraction de texte du PDF du devis rend **zéro** occurrence de `WTC`, `TELSAM` ou `TOTAL` :
14 polices embarquées, les caractères ne se décodent pas. Le tableau vit dans le **classeur Excel
incrusté** du `.doc` — mais `InlineShapes.Item(2).OLEFormat.Object` renvoie **`$null`** sur un
document ouvert en lecture seule.

> **La séquence qui marche** : copier le `.doc` dans le scratchpad (jamais le fichier de Patrice),
> l'ouvrir **en écriture**, appeler `$sh.OLEFormat.Activate()`, PUIS lire `OLEFormat.Object` et
> parcourir `UsedRange`. Ici : 55 lignes, les 11 articles et le total.

### 5. NE JAMAIS VÉRIFIER UN PDF QU'ON VIENT DE GÉNÉRER EN LE RÉOUVRANT DANS WORD

J'ai cru trois fois de suite avoir un brief cassé — numéros de section disparus, une puce collée à
la ligne précédente, « mesurées au / sol » coupé en deux, un `/` parasite. **Le document était
bon ; c'est la relecture qui était fausse.** Word *reconvertit* un PDF à l'ouverture : il
transforme les retours à la ligne VISUELS en paragraphes, fusionne, et fabrique des glyphes.
Et `Content.Text` n'inclut jamais les numéros de liste automatiques — donc « 0 numéro sur 5 » ne
prouvait rien.

> **La bonne vérification** : enregistrer le MÊME document en `.docx` dans le scratchpad, à côté du
> PDF et dans la même opération, puis relire le **témoin `.docx`** — aucune conversion, donc aucun
> artefact. Résultat ici : 31 paragraphes pour 31 lignes source, **0 écart**, les 5 titres
> numérotés, 14 puces, 6 triangles, 1 page.
>
> **Et construire le document par insertion d'un SEUL bloc** (`$doc.Content.InsertAfter` avec les
> lignes jointes par `\r`), puis mise en forme **par index de paragraphe**. La saisie par
> `Selection.TypeText` fusionnait réellement deux lignes — c'est la méthode déjà recommandée plus
> haut pour les `.docx`, elle vaut aussi pour un document créé de zéro. Contrôle **avant**
> `SaveAs2` : chaque paragraphe doit être égal à sa ligne source, sinon on n'enregistre pas.

### Ce qui reste à faire sur ce chantier

1. **Patrice : rendre le lien de partage du dossier App Tech PUBLIC + mot de passe.** Vérifié après
   création : `audience: no_one`, `password_protected: false` — le bouton Documents ne s'ouvre donc
   pas encore chez le technicien.
2. **Le PDP.** Il n'y en a AUCUN, et l'article 1 du devis vend un « plan de prévention ligne ».
   Alerte rouge sur la fiche, et le brief lui-même dit au technicien de ne pas partir sans vérifier.
   À réclamer à SEMI France.
3. Aucune NDS ni MO TELSAM sur ce chantier.

### Chausse - Revigny — toujours ouvert, et je ne peux pas trancher

Deux PGO du 07/09 sur un chantier sans fiche. Le PGO octobre-décembre porte une phase **TRAVAUX
OPTIQUES** (déroulage, raccordements optiques, test liaison, réflectométrie des tourets) et TELSAM
est dans la légende des entreprises avec « Date révision : 07/09/2026 ». **Mais l'ordre de lecture
d'un PDF en tableau ne préserve pas la correspondance ligne ↔ entreprise** : SEMI France n'y
apparaît qu'une fois aussi, et elle a manifestement des lignes. À confirmer par Patrice ou par la
source Excel du SPS. **Ne pas créer de fiche sur cette base.**

## Planning, QUATRIÈME passe (08/09/2026) — rayures, cible du crayon, et la palette enfin mesurée sur la bonne chose

Quatre demandes de Patrice dans un seul message : la petite croix des notes inatteignable **sur les
cases blanches du bas**, les congés « un peu fades » à rayer, « la même chose pour les ICP avec une
rayure un peu différente », et « les couleurs sont encore parfois un peu trop proches (Rion des
Landes / Fleyriat par exemple) ».

### 1. LA CIBLE DU CRAYON — le pointillé « libre » lui volait ses clics

**Mesuré avant de toucher au code**, avec `elementFromPoint` point par point, et c'est la mesure qui
a tout donné :

| sorte de case | zone cliquable du bouton de note |
|---|---|
| case AVEC chantier | 43 × 32 px, **972 points sur 972** |
| case VIDE (blanche) | 43 × 32 px, **157 points sur 972** |

Le rectangle était le bon dans les deux cas — c'est pour ça que le banc du 07/09 disait « TOUT
PASSE ». Mais dans une case vide, `.plVide::after` (le pointillé décoratif) est un **pseudo-élément
de la case** : il se peint PAR-DESSUS le bouton et lui prend ses clics, ne lui laissant qu'une
**lisière de 3 px** le long du bord. D'où « très difficile de pouvoir l'atteindre et c'est fatigant ».

Correctif : `pointer-events:none` sur ce pointillé (il est décoratif, il n'a aucune raison de
recevoir un clic, et le clic repart là où il allait déjà — la case, pour le pinceau et le dépôt),
plus `z-index:2` sur le bouton en ceinture. La cible passe à **1376 points sur 1376, identique pour
les quatre sortes de case**. Et un **anneau au repos** (`box-shadow 0 0 0 1px var(--border)`) : sur
une case blanche la pastille était blanche sur blanc, il ne restait qu'un « + » gris à viser — la
cible était à la fois volée ET invisible.

> **LA LEÇON, et elle vaut pour tout banc d'essai de ce dépôt.** Le harnais du 07/09 prenait
> `querySelector('.plNoteBtn')`, donc **le premier bouton de la grille** — toujours celui d'une case
> avec chantier. Il ne pouvait pas voir le défaut. C'est la même faute que « un test qui ne peut pas
> échouer ne prouve rien », sous une autre forme : **un banc qui n'échantillonne qu'une sorte
> d'objet ne prouve rien sur les autres.**
>
> `partage/bloc-test-cible-note.html` mesure donc maintenant **les quatre sortes de case** (chantier,
> vide, absence, texte), exige que la cible soit **PLEINE** (≥ 95 % de son rectangle, le contrôle qui
> manquait) et que les quatre soient **identiques à 5 % près**. **Rejoué sur le fichier d'avant
> correction : 2 échecs, « 157 points sur 972 (16 %) ».** Il sait donc dire non.

### 2. LES RAYURES — trois registres qui ne peuvent plus se confondre

Sa demande était de goût, mais elle se traduit en lecture : la grille se parcourt en diagonale pour
répondre à « qui est où », et une case pâle sans texture se lit comme un chantier dont on n'a pas
trouvé la couleur. La rayure la sort du registre des chantiers **avant** qu'on ait lu le mot dedans.

| registre | fond | rayure |
|---|---|---|
| chantier | bulle **saturée**, texte blanc | aucune |
| absence (CP, CP paternité, RTT, AM, récup, VM…) | rose pâle | **montantes, 45°**, fines et serrées |
| **ICP** (nouveau, `.plIcp`) | violet pâle | **descendantes, 135°**, larges et doublées |

**Le SENS de la rayure suffit à les séparer du coin de l'œil**, et il ne dépend pas de la couleur —
ce qui compte sur un portable au soleil et à l'impression en noir et blanc (d'où l'ajout de
`print-color-adjust:exact` sur les deux : sans lui l'impression les efface et la case redevient
« fade », exactement le défaut qu'on corrige).

**POURQUOI LE VIOLET POUR L'ICP, ET PAS L'AMBRE** : l'ambre est déjà le registre des notes libres
(`.plBulleNote`, le post-it au crayon). Une ICP venue de Teams et une note écrite ici ne sont pas la
même chose — l'une est un engagement, l'autre un pense-bête — et les confondre effacerait la
distinction que Patrice demande justement d'accentuer.

L'ICP tombait jusque-là dans `plTexte`, le gris italique fourre-tout, avec « agence », « Hauteur » et
« PASS RTE » — alors que c'est le rendez-vous qui CONDITIONNE le chantier et qu'on le cherche dans la
grille. Mesuré sur l'année : **14 cases ICP**. Le `\b` de `/^ICP\b/i` n'est pas décoratif (contrôle
avec son contre-exemple : « ICPESSAI » reste du texte), et « ICP TIVERNON annulée » entre bien dans
le lot — une ICP annulée reste une ligne d'ICP.

**`background-color` + `background-image`, JAMAIS le raccourci `background`** : le raccourci
effacerait la rayure ET le reflet de `.plBulle`. Un contrôle du banc vérifie que la bulle rayée garde
ses deux couches.

*Limite assumée, à dire à Patrice : la rayure ICP porte sur le texte de la case venu de **Teams**.
Une ICP écrite comme note dans le suivi garde le post-it ambre au crayon — c'est un autre objet.*

### 3. LES COULEURS — LE COUPLE QU'IL A CITÉ N'ÉTAIT PAS LE PIRE, ET DE LOIN

C'est la mesure qui compte de cette session. Écart **CIEDE2000** des **894 paires de chantiers
réellement co-visibles** dans une quinzaine, sur l'année entière :

| | pire paire | paires sous 15 | sous 20 | à court de teintes |
|---|---|---|---|---|
| avant | **9,5** | 53 | 157 | 2 |
| après | **14,0** | **10** | **112** | **0** |

**La pire paire d'avant était deux BLEUS — `#0C7CB5` et `#1D6FE0` à 9,5, au coude à coude sur cinq
quinzaines.** Rion / Fleyriat, le couple qu'il cite, mesurait 22,6. Son œil avait donc raison **bien
en dessous** du seuil où je croyais le problème réglé.

> **RÈGLE — quand Patrice dit que deux couleurs se ressemblent, la mesure à refaire n'est pas celle
> du couple qu'il cite : c'est celle de TOUTES les paires.** Deux fois de suite (07/09 puis 08/09)
> j'ai mesuré le couple nommé, trouvé un écart confortable, et conclu que la palette allait bien.
> Le couple qu'il cite est un **symptôme**, pas le diagnostic.

**POURQUOI CIEDE2000 ET PAS ΔE 76.** L'ancienne mesure disait Rion `#C4392C` / Fleyriat `#8C2F55` à
**44** — « très différentes » — alors que les deux sont des rouges sombres. L'essentiel de ces 44
est un écart de CLARTÉ, que l'œil pardonne beaucoup plus qu'un écart de teinte. En CIEDE2000 le même
couple tombe à **22,6**, ce qui correspond à ce qu'il voit. Table des 29 × 29 couples calculée **une
fois au chargement** : `repartirTeintes()` tourne à chaque rendu, recalculer la formule ferait payer
l'affichage.

**LE COMMENTAIRE QUI DISAIT « NE PAS REFAIRE » ÉTAIT FAUX, et il a bloqué la bonne piste 24 h.**
Le 07/09 j'avais écrit dans `construireCarteTeintes` : « NE PAS AJOUTER ICI UN ÉVITEMENT DES TEINTES
PROCHES — essayé et mesuré, aucun gain. » La tentative du 07/09 était mal outillée sur deux points,
et aucun n'était « l'idée est mauvaise » :
1. elle raisonnait sur des **noms de famille** de teinte, pas sur l'écart que l'œil voit ;
2. elle départageait un glouton **déjà saturé** — mesuré : 24 teintes bloquées autour de l'étiquette
   la plus contrainte, sur 24 disponibles, et deux étiquettes n'en trouvaient plus aucune. Un
   départage n'a rien à départager quand il ne reste rien de libre.

> **« Mesuré, aucun gain » n'autorise PAS à écrire « à ne pas refaire ».** Ça autorise à écrire
> « essayé COMME CECI, aucun gain, voici comment c'était fait ». La première formule ferme une piste
> pour de bon ; la seconde laisse la reprendre autrement. Vaut au-delà des couleurs.

**CE QUI EST FAIT, en deux temps et l'ordre compte** (`construireCarteTeintes`) :
1. si la couleur **préférée** du chantier est libre ET déjà à plus de `ECART_SUFFISANT` (20) de tous
   ses voisins, il la garde — c'est ce qui préserve la promesse du 03/09 là où elle ne coûte rien ;
2. sinon, parmi les teintes libres, celle dont l'écart **minimal** aux voisins est le **plus grand**.
   Pas la première libre : la première libre est ce qui posait deux bleus côte à côte.

**Le seuil n'est pas un compromis, c'est l'optimum des trois variantes mesurées** : ignorer
complètement la préférence donne la même pire paire (14,0) mais **126** paires sous 20 au lieu de 112,
et déplace 86 chantiers sur 93 au lieu de 75.

**LA PALETTE FRANCHE PASSE DE 10 À 15** (`PALETTE_GRILLE_APPOINT` : prune profond, aubergine, olive
sombre, sapin bleuté, bordeaux). Parce que la grille montre jusqu'à **15 chantiers en même temps** —
et ce chiffre a été annoncé 14 le 03/09, **10 le 08/09 au matin**, 15 le 08/09 au soir : les deux
premiers ont chacun servi à justifier une palette trop courte. **Ne pas le remplacer sans dire sur
quoi porte le nouveau compte** (ici : les étiquettes DE GRILLE co-visibles dans une quinzaine,
réserve exclue).

Les cinq ajouts sont choisis par glouton max-min en CIEDE2000 (17,7 / 17,3 / 17,2 / 16,9 / 14,0),
portent tous du blanc à ≥ 4,5:1 (9,6:1 à 14,8:1), et **aucun n'est plus proche des dix que les dix ne
le sont entre elles** : la pire paire du noyau est violet `#7B3FE4` / pourpre `#8E2BA0` à **12,1**.
*S'il faut resserrer un jour, c'est ce couple-là qu'il faut traiter, pas les ajouts.*
Elles sont toutes sombres (L\* 16 à 32 quand les dix vont de 31 à 49) : la clarté devient un second
indice, en plus de la teinte.

**ELLES NE SONT PAS AJOUTÉES À `PALETTE_PLANNING`, et c'est volontaire** : la couleur préférée est
`rang % PALETTE_PLANNING.length`, passer cette liste de 10 à 15 changerait le modulo et donc la
couleur habituelle de **presque tous** les chantiers d'un coup. Un contrôle du banc le vérifie.

**Résultat sur le couple qu'il a cité** : Rion des Landes `#6558B8` (lavande foncé) contre Fleyriat
`#365314` (olive sombre) — ΔE2000 **50,4**.

**Et à l'écran, sur la quinzaine la plus chargée de l'année** (01/06, 15 chantiers, 10 couleurs
affichées) : la paire la plus proche réellement visible est à **15,7** (`#7F1D1D` Casteljaloux contre
`#C4392C` St Christol), puis 16,9 (ardoise contre sapin). Contre 9,5 avant.

### CE QUI A ÉTÉ MESURÉ ET **PAS** RETENU — la coloration par plage de semaines

La troisième passe (ci-dessus) proposait de colorer par **plage continue de semaines** au lieu de par
chantier, ce que la règle de Patrice du 03/09 autorise (« pas grave de les changer quand on
réintervient un ou deux mois plus tard »). **Mesuré cette fois, avec une coupure à 5 semaines :**

| | pire paire | sous 15 | stabilité |
|---|---|---|---|
| par plage de semaines | 14,0 | 2 | un chantier peut changer de teinte entre deux campagnes |
| **par chantier (retenu)** | **14,0** | 10 | **une couleur par chantier, toute l'année** |

**La pire paire est la MÊME.** Le découpage en plages n'achète que 8 paires de moins sous 15, au prix
de la propriété que Patrice a demandée nommément. **Écarté** — et le paragraphe de la troisième passe
qui le présentait comme « la sortie » est donc périmé : la sortie était la palette et le critère de
choix, pas le découpage. *À reprendre seulement si Patrice trouve encore deux couleurs proches ET
accepte explicitement le changement entre campagnes.*

### Vérifié

**Suivi : 263 contrôles, 0 échec** — `bloc-test-rayures-ecarts` (22, **nouveau**), `bloc-test-couleurs`
(10), `bloc-test-cible-note` (23, passé de 6 à 23), `bloc-test-cible-croix` (6), `bloc-test-affaires`
(74), `bloc-test-saisie-affaires` (90), `bloc-test-notes` (22), `bloc-test-validation-suivi` (16).
Les deux blocs `<script>` compilent, et le comptage d'encodage donne **0 `Ã`, 0 caractère de
remplacement**.

Le nouveau banc porte ses contre-exemples : un chantier n'est **pas** rayé, une case texte ordinaire
non plus, « ICPESSAI » n'est pas une ICP, et surtout **il rejoue l'état exact d'avant** (ancienne
règle ET ancienne palette) pour vérifier qu'il sait trouver le défaut : 9,5 contre 14,0, 50 paires
sous 15 contre 10.

*Piège rencontré en l'écrivant : mon premier contre-exemple rejouait l'ancienne règle avec la palette
ÉLARGIE. Il donnait 12,1 au lieu de 9,5 et **le banc échouait de justesse sur un succès**. Un
contre-exemple doit reproduire l'état d'avant EN ENTIER, sinon il mesure un mélange qui n'a jamais
existé.*

**Et à l'écran**, ce qu'aucun de ces contrôles ne montre : la grille sur deux quinzaines, avec les
congés rayés en rose, les ICP rayées en violet, les cases blanches du bas et leur bouton visible.

**`SEED_DATA` n'est pas touché : pas de `SEED_VERSION` à bumper.** Tout est CSS et calcul de rendu.

## La charte « Telsam Ops » et la grille en B' (08/09/2026)

Patrice a lancé une refonte visuelle du suivi et de l'appli à partir d'une charte reçue par Teams :
**« Identité visuelle Telsam Ops.pdf »**, envoyée par Pierre Brillou (son OneDrive, dossier
« Fichiers de conversation Microsoft Teams »). Navy `#16175B`, terracotta `#CD8E81` réservé à
l'accent, indigo `#4B4EDB` pour les actions, League Spartan / DM Sans / IBM Plex Mono.

Les maquettes de travail vivent dans **`TELSAM-apps\maquette-charte\`** — dossier parent, donc
**hors des deux dépôts Git** : elles ne sont ni versionnées ni publiées, et c'est voulu.

### CE QUE LA CHARTE DEMANDAIT ET QUI NE PASSE PAS LA MESURE — ne pas re-proposer

La charte proposait de remplacer les couleurs pleines de chantier par **dix teintes très claires**
(`oklch(0.955 0.028 H)` pour le fond, `oklch(0.58 0.145 H)` pour le liseré), attribuées **par
hachage** du n° d'affaire. Trois choses, mesurées avant d'écrire une ligne, sur les **894 paires de
chantiers réellement co-visibles** dans une quinzaine de l'année :

| registre | pire paire co-visible | paires sous 15 |
|---|---|---|
| la palette actuelle (26 teintes du 08/09) | **14,0** | 10 / 894 |
| charte, fond `oklch(0.955 0.028)` | **4,7** | 433 / 894 |
| fond plus soutenu, `oklch(0.84 0.10)` | **7,9** | 127 / 894 |

**La cause est structurelle, pas un défaut de réglage** : quand une quinzaine porte 15 chantiers,
toutes les teintes sont côte à côte, donc l'écart co-visible ne peut pas dépasser celui de la
palette elle-même — et un registre clair n'a pas la place perceptive d'en tenir quinze. Corollaire
contre-intuitif à garder : **on ne rattrape pas un fond trop pâle en le saturant.** À clarté 0,955,
monter la saturation de 0,028 à 0,13 fait *baisser* l'écart minimal de 4,7 à 4,0, tout étant écrasé
contre le blanc.
Deux autres refus de la même veine : l'attribution **par hachage** ne garantit rien, là où le
glouton max-min garantit 0 collision dans la grille ; et le **terracotta ne peut pas porter de
texte** (2,70:1 sur blanc, 2,52:1 sur le fond de travail) — il ne sert qu'en filet, en marqueur et
en liseré d'onglet actif.

> **LA MESURE À FAIRE EST CELLE DES PAIRES CO-VISIBLES, jamais celle de la palette.** La palette de
> la charte annonçait 13,6 sur ses dix liserés — meilleur que nos 12,1 — et s'effondre à 4,7 dès
> qu'on la confronte au planning réel. Une palette se juge sur ce qui se retrouve ensemble à
> l'écran, pas sur ses paires théoriques.

### B' A ÉTÉ ESSAYÉE, POUSSÉE, ET RETIRÉE LE JOUR MÊME — ne pas y revenir

**Patrice, en voyant la grille livrée : « les couleurs ne sont absolument pas celles que tu m'as
présentées, c'est fade, on ne discerne pas les couleurs… c'est nul et inutilisable. Tu m'as vendu
ça. »** Il avait raison, et la mesure qui manquait le dit sans détour. Écart CIEDE2000 entre deux
chantiers **réellement co-visibles**, la teinte étant diluée vers le blanc :

| fond des cases | écart co-visible minimum |
|---|---|
| 8 % (ce qui a été poussé) | **1,4** |
| 22 % | 3,5 |
| 40 % | 5,9 |
| **bulle saturée (ce qui était en place)** | **14,0** |

> **DILUER VERS LE BLANC ÉCRASE LES ÉCARTS, ET AUCUN RÉGLAGE N'Y ÉCHAPPE.** Deux teintes séparées
> de 14 saturées ne le sont plus que de 5 une fois délavées à 60 % de blanc. Toute la famille
> « fond clair porte l'identité du chantier » est donc morte pour cette grille — pas seulement la
> valeur 8 %, la famille entière. **Ne pas re-proposer un fond clair pour les cases du planning.**

**ET VOICI L'ERREUR DE RAISONNEMENT QUI A COÛTÉ LA JOURNÉE, elle est instructive.** La maquette
montrait une variante B qui, elle, paraissait colorée — et je m'en suis servi pour vendre B'. Mais
la B de la maquette utilisait une palette **régénérée pour cette clarté** (`oklch(0.84 0.10 H)`),
pas les teintes de Patrice diluées. Deux objets différents sous une même étiquette. Et j'ai annoncé
« l'écart reste à 14,0 » : c'était vrai **du liseré**, alors que ce qu'un œil lit dans une case de
40 px, c'est le **fond**. Une mesure juste sur le mauvais objet ne vaut pas mieux qu'une mesure
fausse.

**CE QUI EST GARDÉ de ce passage, parce que ça ne dépendait pas du fond :** le n° d'affaire en mono
sur la première ligne et le nom sur la seconde **sans ellipse** (le défaut « Poste de Portet (P… »),
et les fonds d'absence et d'ICP **approfondis** (`#f2c9c9`, `#d5d0f4`) — plus distincts qu'avant,
ce que Patrice demandait le matin même. Le reste est revenu à l'identique : bulle saturée, texte
blanc, reflet et ombre d'origine. Les pastilles de la réserve n'ayant jamais changé, l'écran est de
nouveau cohérent — c'était son troisième reproche, et il tombait de lui-même avec le retour.

### CE QUI AVAIT ÉTÉ FAIT : LA VARIANTE B' (retirée, cf. ci-dessus)

Tranchée par Patrice à l'œil sur la maquette, après avoir vu les trois registres à 10 chantiers.
**La teinte ne change pas d'un iota** — ce sont les 26 teintes mesurées le 08/09, donc l'écart
co-visible reste à **14,0**. Ce qui change est la **surface** que la couleur occupe :

- la case chantier est un **fond très clair** (sa teinte mélangée à 8 % de blanc) plus un
  **liseré gauche de 4 px** à la teinte pleine, encre `var(--text)` à 11,9:1 ;
- le **n° d'affaire est en mono sur la première ligne**, le nom sur la seconde, **sans ellipse**
  (règle 7 de la charte : plus de « Poste de Portet (P… »). Les lignes de la grille sont donc un peu
  plus hautes, c'est le prix et il est voulu ;
- le numéro porte la teinte **assombrie de 20 %** : mesuré sur les 26 teintes, la teinte brute ne
  tenait 4,5:1 sur son propre fond que dans 20 cas sur 26 ; à −20 % le minimum passe à 5,4:1 et
  aucune n'échoue ;
- **les pastilles de la réserve du bas gardent leur aplat saturé.** Ce sont des boutons qu'on lit et
  qu'on fait glisser, avec leur nom écrit dessus — le raisonnement du 03/09 tient. Et la différence
  de registre devient utile : réserve = pastille pleine, grille = fond clair + liseré.

**LE PIÈGE QUE B' CRÉE, ET QU'IL A FALLU CORRIGER DANS LE MÊME GESTE.** Un fond de chantier pâle
entre dans le registre des **absences** et des **ICP**, qui sont pâles aussi. Mesuré : le fond du
rouge brique `#C4392C` tombait à **ΔE 1,0** du rose des congés, et l'indigo `#4B4FCE` à **1,5** du
violet des ICP — la même couleur. Or c'est exactement la distinction que Patrice demandait
d'ACCENTUER le matin même. Les deux fonds ont donc été **approfondis, localement à la grille** :
`.plAbsence .plBulle` passe à `#f2c9c9` et `.plIcp .plBulle` à `#d5d0f4` (écarts remontés à **11,8**
et **10,5**, textes à 6,9:1), et `.plTexte .plBulle` à `#ece8dc`. Ces valeurs sont **écrites en dur
dans la grille** et ne touchent pas `--red-bg` / `--purple-bg`, qui continuent de servir aux alertes
du suivi — là où aucune case de chantier ne les côtoie.

> **DANS LA GRILLE, LE FOND NE SÉPARE PLUS LES REGISTRES.** Ce sont le **liseré** (chantier
> seulement) et le **SENS de la rayure** (absence 45°, ICP 135°) qui les séparent — deux indices qui
> ne dépendent ni de la teinte ni de la couleur, donc qui tiennent au soleil et en noir et blanc.
> **Ré-éclaircir un fond d'absence ramène le défaut, et il revient SANS RIEN AFFICHER D'ANORMAL.**
> Écart résiduel assumé : **3,0** entre une case chantier or/brun et une case du registre `plTexte`
> (le gris italique fourre-tout). C'est le moins conséquent des quatre — on lit les mots — et une
> case de chantier y ajoute un liseré et un numéro en mono.

**UN SEUL ENDROIT POSE LA COULEUR** : `styleBulleChantier(fond)` et `contenuBulleChantier(texte, c)`.
Trois rendus en avaient besoin (case du planning, case du brouillon, et la réserve qui garde son
aplat) et les trois écrivaient leur `style=` à la main — ils avaient déjà divergé une fois. La
couleur arrive en **`background-color`** et **`border-left-color`**, jamais en raccourci : `background`
effacerait le reflet, `border-left` effacerait la largeur posée par la feuille de style et **le
liseré disparaîtrait sans que rien ne le signale**, la case restant simplement plate.
`.plNum` est ajouté au bloc `@media print` : sans lui le numéro sortait en noir et on perdait le
seul rappel de couleur qui survive à une impression économe.

### Le logo vectoriel existe — extrait du `.ai` le 08/09/2026

Patrice a fourni `Bureau\logo-Telsam-23-fondBleu.ai`. **Un `.ai` récent EST un PDF** : celui-ci est
un PDF 1.6, une page, sans police ni image — tout en tracés. Extrait en SVG dans
`maquette-charte\` : `logo-telsam.svg` (fidèle, 8,3 Ko) et **`logo-telsam-compact.svg`** (logotype
et flèches du « s » seulement, cadré serré, 7,2 Ko). **C'est le compact qu'il faut en en-tête** :
sous 40 px, les trois virgules du premier tombent à 0,07 px de trait et le logo paraît délavé. La
carte d'Europe du fond (13 005 segments) est volontairement laissée de côté — invisible sous 200 px.
Le marqueur « T » terracotta que la charte proposait en substitut est donc inutile.

*Méthode, si c'est à refaire : copier le `.ai` en `.pdf` dans le scratchpad, inflater le flux
`/Contents` avec `DeflateStream` (sauter les 2 octets d'entête zlib), puis traduire les opérateurs
`m l c v re h f f* S k K w cm` en chemins SVG, l'axe Y retourné par `translate(0 H) scale(1 -1)`.*

> **DEUX PIÈGES POWERSHELL PAYÉS ICI, TOUS DEUX DE CASSE.** `-eq` est **insensible à la casse** :
> mon test de l'opérateur `M` (limite d'onglet du PDF) avalait donc tous les `m` (moveto) et aucun
> chemin ne démarrait. **Utiliser `-ceq` pour comparer des opérateurs.** Et les **noms de variables
> le sont aussi** : le `$h` du rectangle écrasait le `$H` de la hauteur de page. Les deux se
> voyaient à l'écran — des bandes vides — et aucun ne se voyait à la relecture.

### L'OSSATURE, faite juste après (08/09/2026) — et l'erreur de lecture qui l'a retardée

**Patrice, en voyant le suivi poussé : « qu'est-ce que tu as fait ? je ne vois rien de ce que tu
m'as proposé ».** Il avait raison. Il avait validé la maquette — barre navy, typographie, en-tête —
puis dit « on part sur la B', commit et pousse » ; j'ai lu ça comme « pousse la variante de couleur
de la grille » et livré ce seul détail, en annonçant le reste comme « pas commencé ».

> **LA LEÇON, et elle vaut au-delà de ce cas : « on part sur X » après une maquette veut dire la
> maquette, pas la ligne X.** Quand une décision porte sur un point d'un ensemble qu'il vient de
> valider, c'est l'ensemble qui est commandé. Ne pas rétrécir la commande à la question posée.

Faite dans la foulée, et c'est le **geste n° 1** de la charte :
- **navigation latérale navy** groupée « Pilotage » / « Terrain », à la place de la rangée de dix
  boutons identiques ;
- **en-tête** avec le titre de l'écran, la fraîcheur de la donnée et **trois actions** (Nouveau
  chantier en indigo, Imprimer, Exporter) ;
- **typographie** League Spartan (titres) / DM Sans (interface) / IBM Plex Mono (chiffres) ;
- le **logotype vectoriel** en `data:` dans la variable `--logo-telsam`, posé en `mask` sur un
  aplat blanc — donc dans le fichier, donc hors ligne, et recolorable ;
- l'imprimante et la loupe **emoji** retirées.

**LES IDENTIFIANTS `btnView*` SONT GARDÉS TELS QUELS, ET C'EST LA SEULE CHOSE À NE PAS DÉFAIRE.**
`setView` bascule la classe `active` sur chacun des dix, par `getElementById`. Refaire le balisage
sans garder les ID casse la vue active **en silence** : l'écran s'afficherait, mais plus aucun
onglet ne se marquerait. Les anciens boutons ont donc été SUPPRIMÉS, pas cachés — deux éléments du
même ID auraient laissé `getElementById` en désigner un au hasard.
`TITRES_VUES` porte le titre de chaque écran : depuis que les vues sont dans la barre latérale,
l'en-tête est le seul endroit qui dit où l'on est.

*Deux compteurs de la charte ont été RETIRÉS du balisage plutôt que laissés vides : « Alertes » et
« À vérifier » n'ont pas de comptage fiable sous la main, et un compteur faux est pire qu'absent.
Seul celui d'Affaires est posé, depuis `AFFAIRES_RTE.affaires.length`.*

**Vérifié en plus des 263 contrôles** : les dix vues ouvertes une par une — titre juste, onglet
marqué, vue affichée, **zéro erreur JavaScript** — et aucun identifiant `btnView*` en double.

### Ce qui reste de la charte, et n'est PAS fait

**bandeau d'alerte unique** et repli du reste (l'onglet Affaires en empile encore trois),
**colonne conformité** PDP/PGO/IST/MAT, les trois derniers emoji (⚠️ ×3, 📄, 📋 ×2, tous dans du
texte d'alerte et non dans le chrome), et le **thème sombre d'AppTech** (charte fournie par
Patrice, maquette à quatre écrans dans `maquette-charte\maquette-apptech.html`).

**Deux points à ne pas trancher seul quand on reprendra :**
1. **Les listes déroulantes de statut de l'onglet Affaires.** La charte demande une pastille avec
   édition au survol ; c'est exactement ce que le lot 2 du 07/09 a **remplacé** par une liste
   déroulante à liseré de 4 px, parce qu'il fallait distinguer douze étapes. Y revenir défait un
   arbitrage éprouvé de la veille.
2. **Les polices viennent de Google.** Le suivi part par Teams et s'ouvre parfois hors ligne : dans
   ce cas il retombe sur Arial. Soit on les embarque dans le fichier (+250 Ko environ), soit on
   accepte le repli. Question posée à Patrice, sans réponse à ce jour.

**Vérifié — 263 contrôles, 0 échec** : `bloc-test-couleurs` 10, `bloc-test-rayures-ecarts` 22,
`bloc-test-cible-note` 23, `bloc-test-cible-croix` 6, `bloc-test-notes` 22,
`bloc-test-validation-suivi` 16, `bloc-test-affaires` 74, `bloc-test-saisie-affaires` 90 ; les deux
blocs `<script>` compilent, 0 `Ã` et 0 caractère de remplacement.
Le contrôle qui comptait le plus ici est **`bloc-test-cible-note`** : le rembourrage de `.plBulle` a
changé, et la cible du crayon reste à **1344 points sur 1344, identique sur les quatre sortes de
case**. Et `bloc-test-rayures-ecarts` confirme les **14,0** avec son contre-exemple (9,5 avec
l'ancienne règle).
**Et à l'écran** : la grille sur les semaines 37 et 38, les congés rayés nettement plus soutenus que
les cases de chantier, et les noms de chantier écrits en entier.

**`SEED_DATA` n'est pas touché : pas de `SEED_VERSION` à bumper.** Tout est CSS et rendu.
