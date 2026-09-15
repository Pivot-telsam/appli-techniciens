<#
  COMPARER-STATUTS - les deux depots doivent dire la MEME chose sur PDP, PGO,
  IST et materiel.

  POURQUOI CE CONTROLE EXISTE (14/09/2026)
  Un technicien a appele Patrice : son PGO s'affichait ROUGE dans l'appli alors
  que les dates etaient bonnes et qu'il avait le bon document dans App Tech.
  Mesure : sur 26-030 Poste de Rion des Landes, le suivi disait "ok" et l'appli
  disait "bad". Le meme champ, deux valeurs, dans deux fichiers que personne ne
  compare. L'ecart existait depuis au moins le 09/09 - donc plusieurs jours
  pendant lesquels le technicien a vu un rouge qui ne voulait rien dire.

  CE QUE CA COUTE, ET C'EST LE VRAI ENJEU. Un rouge faux n'est pas un detail
  d'affichage : c'est un voyant de securite qui crie a tort. Repete, il apprend
  au technicien a ne plus le lire - donc a ne plus le lire le jour ou il compte.
  C'est exactement le mal que tous les controles de ce depot cherchent a eviter.

  Le suivi est la source de verite (CLAUDE.md : "appli-techniciens ne fait que
  consommer un sous-ensemble allege"). Mais rien ne le VERIFIAIT : les deux
  fiches se modifient a la main, l'une apres l'autre, et l'oubli de la seconde
  ne se voit nulle part.

  CE QU'IL COMPARE : le champ "statut" de pdp / pgo / ist / materiel, pour les
  fiches presentes dans LES DEUX depots, rapprochees par leur "id" (jamais par
  le numero, qui peut etre renumerote - cf. St-Guillerme 26-045 devenu 26-080
  le 10/09/2026).

  ET DEPUIS LE 15/09/2026, LE NUMERO D'INDICE AUSSI. L'audit de ce jour a trouve
  six fiches ou l'appli affichait un autre indice que le suivi et que le
  document present dans App Tech : Bradascou PDP "4" alors que l'indice 6 est
  dans le dossier, Chaineau lots 1 et 2 PDP "3" alors que seul l'indice 2
  existe, Rion des Landes PDP "1" pour un indice 2, Cantegrit PGO d'aout alors
  que celui de septembre est arrive. Le badge de l'appli AFFICHE cet indice
  (abregeIndice) : le technicien compare ce chiffre au document qu'il tient.
  Un chiffre faux, meme sous un badge vert, lui fait croire qu'il n'a pas le
  bon document - ou l'inverse.
  On ne compare que le PREMIER MOT de l'indice ("6", "6.1", "V03"), parce que
  l'appli porte volontairement un texte plus court derriere ; le reste du texte
  n'a pas a etre identique. Deux indices vides, ou un seul renseigne, ne sont
  pas un ecart : c'est le statut qui porte alors l'information.

  UNE FICHE ABSENTE DE L'APPLI N'EST PAS UNE ERREUR : toutes les fiches du suivi
  n'ont pas a etre publiees. On ne compare que l'intersection.

  LANCEMENT
    powershell -ExecutionPolicy Bypass -File comparer-statuts.ps1
  Sortie 0 = les deux depots concordent. Sortie 1 = au moins un ecart.
#>
param(
  [string]$Suivi = '',
  [string]$Appli = ''
)

$ErrorActionPreference = 'Stop'

if (-not $Suivi -or -not $Appli) {
  $racine = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
  if (-not $Suivi) { $Suivi = Join-Path $racine 'suivi-chantiers\suivi_chantiers_205.html' }
  if (-not $Appli) { $Appli = Join-Path $racine 'appli-techniciens\index.html' }
}

# Depot voisin absent (clone isole) : on AVERTIT et on laisse passer, plutot
# que de bloquer un travail qu'on ne peut pas verifier. Meme choix que
# check-claude-md-sync.sh.
foreach ($f in @($Suivi, $Appli)) {
  if (-not (Test-Path -LiteralPath $f)) {
    Write-Host "  (statuts) $f introuvable : controle non joue, rien n'est bloque."
    exit 0
  }
}

function LireSeed([string]$fichier) {
  $t = [System.IO.File]::ReadAllText($fichier, [System.Text.Encoding]::UTF8)
  $m = [regex]::Match($t, '(?m)^const SEED_DATA = (.*)$')
  if (-not $m.Success) { throw "const SEED_DATA introuvable dans $fichier" }
  # .Trim() AVANT ET APRES le point-virgule, et ce n'est pas de la coquetterie.
  # Le 14/09/2026 ce controle s'est TU pendant un commit : la copie de travail du
  # suivi etait passee en fins de ligne Windows (git checkout + core.autocrlf),
  # la capture `(.*)$` emportait donc un retour chariot final, `TrimEnd(';')` ne
  # retirait plus rien, et ConvertFrom-Json refusait la ligne. Le catch appelant
  # avale l'erreur et sort en 0 : le commit passait AVEC le controle eteint.
  # Un controle qui ne peut plus se declencher est pire que pas de controle.
  return ($m.Groups[1].Value.Trim().TrimEnd(';').Trim() | ConvertFrom-Json)
}

# Premier mot de l'indice, sans ponctuation de bord, en minuscules :
# "6 (25/08/26) - periode..." -> "6" ; "6.1 (07/09/26)" -> "6.1" ; "V03 du ..." -> "v03".
function PremierMot([string]$s) {
  if (-not $s) { return '' }
  $mots = $s.Trim() -split '\s+'
  if ($mots.Count -eq 0) { return '' }
  return $mots[0].Trim(',;:.()-').ToLower()
}

try {
  $S = LireSeed $Suivi
  $A = LireSeed $Appli
} catch {
  Write-Host "  (statuts) lecture impossible ($($_.Exception.Message)) : controle non joue."
  exit 0
}

$parId = @{}
foreach ($c in $S) { if ($c.id) { $parId[$c.id] = $c } }

$DOCS = @('pdp', 'pgo', 'ist', 'materiel')
$ecarts = @()
$compares = 0
$comparesIndice = 0

foreach ($c in $A) {
  if (-not $c.id) { continue }
  $s = $parId[$c.id]
  if (-not $s) { continue }
  foreach ($d in $DOCS) {
    $sv = ''; $av = ''
    if ($s.$d) { $sv = [string]$s.$d.statut }
    if ($c.$d) { $av = [string]$c.$d.statut }
    if (-not $sv -and -not $av) { continue }
    $compares++
    if ($sv -ne $av) {
      $ecarts += [pscustomobject]@{
        numero = [string]$c.numero
        nom    = [string]$c.nom
        doc    = $d.ToUpper()
        quoi   = 'statut'
        suivi  = $(if ($sv) { $sv } else { '(absent)' })
        appli  = $(if ($av) { $av } else { '(absent)' })
      }
    }
    # L'indice : premier mot seulement, et seulement si les deux cotes en ont un.
    $si = ''; $ai = ''
    if ($s.$d) { $si = [string]$s.$d.indice }
    if ($c.$d) { $ai = [string]$c.$d.indice }
    $sm = PremierMot $si
    $am = PremierMot $ai
    if ($sm -and $am) {
      $comparesIndice++
      if ($sm -ne $am) {
        $ecarts += [pscustomobject]@{
          numero = [string]$c.numero
          nom    = [string]$c.nom
          doc    = $d.ToUpper()
          quoi   = 'indice'
          suivi  = $si
          appli  = $ai
        }
      }
    }
  }
}

if ($ecarts.Count -eq 0) {
  Write-Host ("  (statuts) OK : {0} statuts et {1} indices compares, les deux depots concordent." -f $compares, $comparesIndice)
  exit 0
}

Write-Host ""
Write-Host "  LES DEUX DEPOTS NE DISENT PAS LA MEME CHOSE SUR UN DOCUMENT DE SECURITE"
Write-Host ""
foreach ($e in $ecarts) {
  Write-Host ("    {0}  {1}" -f $e.numero, $e.nom)
  Write-Host ("        {0} ({1}) : le suivi dit '{2}', l'appli dit '{3}'" -f $e.doc, $e.quoi, $e.suivi, $e.appli)
}
Write-Host ""
Write-Host "  C'est ce que le technicien voit sur son telephone : la couleur du voyant et"
Write-Host "  le numero d'indice sous le badge. Un rouge faux lui apprend a ne plus lire le"
Write-Host "  voyant ; un indice faux lui fait douter du document qu'il tient. Le suivi fait"
Write-Host "  foi : aligner l'appli, ou corriger le suivi si c'est lui qui se trompe."
Write-Host ""
exit 1
