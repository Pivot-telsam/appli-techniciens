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

## Système de numérotation (`numero` dans suivi-chantiers, format `26-0XX`)
C'est le numéro de référence que Patrice utilise pour organiser Dropbox (il renomme progressivement
les dossiers avec ce préfixe, ex: "26-055 - Fleyriat - Brou - La Cluse"). Pour tout nouveau chantier,
attribuer le prochain numéro disponible et le communiquer à Patrice.

## Accès Dropbox
Dropbox est connecté et utilisable directement (create_folder, create_shared_link, create_file_request,
search, list_folder, fetch). Toujours confirmer le plan exact avec Patrice avant toute création/modification
dans Dropbox (dossier, lien de partage, file request) — ne jamais deviner un chemin de destination.

## Workflow hebdomadaire
1. Le lien personnalisé technicien (`?tech=slugifiedname`) est permanent — jamais régénéré à chaque semaine.
2. Ce qui change chaque semaine : les affectations technicien/chantier (TECH_RANGES côté appli-techniciens,
   structure équivalente côté suivi-chantiers) — à mettre à jour dans les deux dépôts en cohérence.
3. Les liens Documents/Dépôt photos/base de vie ne sont à créer/vérifier que pour les NOUVEAUX chantiers
   apparaissant dans le planning, pas à chaque semaine pour les chantiers déjà configurés.

## État au 20/08/2026 — travaux en cours
- 6 chantiers ont un bouton "📷 Dépôt photos" (Portet, Fleyriat, Fibrage THYM Bissy-Grand Ile,
  St-Guillerme, Lien Arteria Lamativie, Hospitalet-La Tour de Carol).
- Champ "Immatriculation du véhicule" ajouté à la feuille d'heures (app + export Excel cellule A52 + PDF).
- 28 chantiers ont un champ `baseVie` (adresse extraite des NDS) pour le bouton Itinéraire.
  ~20 chantiers encore non trouvés/à approfondir si besoin.
- EN COURS, PAS ENCORE COMMITTÉ : fusion de "Poste de Tivernon" (26-070) dans "Dambron - Voves"
  (26-031) — même chantier physique confirmé par Patrice. À terminer côté suivi-chantiers ET
  appli-techniciens (supprimer la fiche Tivernon séparée là-bas aussi).
- Corrections de noms en attente (ne changent pas le `numero`) :
  - 26-002 : "Icp / Ls Cross-Sausset-Goodman" → "Ls Cross-Sausset-Goodman"
  - 26-003 : "PS Cantegrit" → "Poste Cantegrit"
  - 26-007 : "CONSIGNATION 225kV CURBANS - SISTERON" — nom incorrect (terme repris du planning brut,
    pas un vrai nom de chantier), à corriger via PDP/devis si trouvable, sinon demander à Patrice.
  - 26-011 : "Vallee Du Louron (Aure-Loudenvielle)" — possible doublon avec 26-024 "Aure - Loudenvielle",
    à vérifier avant fusion (ne pas fusionner sans confirmation).

## Règles de prudence
- Toujours montrer le résumé des changements à Patrice AVANT de committer/pousser, sauf demande explicite contraire.
- En cas de nom de chantier ambigu ou de dossier Dropbox introuvable/multiple, ne jamais deviner —
  poser la question à Patrice.
- Valider la syntaxe JS (node --check) et le JSON (SEED_DATA) avant toute livraison.
