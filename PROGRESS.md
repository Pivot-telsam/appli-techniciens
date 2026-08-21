# État d'avancement — session du 20/08/2026

Dernier commit `appli-techniciens` : `fe43bfb`
Dernier commit `suivi-chantiers` : `206c423`

Ce fichier récapitule une longue session de travail sur les deux dépôts. Objectif : pouvoir
reprendre le fil après compactage de la conversation, sans perdre le contexte.

## Ce qui est fait

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
