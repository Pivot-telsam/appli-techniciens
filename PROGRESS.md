# État d'avancement — sessions des 20, 21 et 24/08/2026

Dernier commit `appli-techniciens` : `4f9d93a`
Dernier commit `suivi-chantiers` : `3865e66`

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

## Ce qu'il reste à faire / points ouverts
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
