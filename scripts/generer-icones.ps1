# Genere les icones de l'ecran d'accueil (manifest.json + apple-touch-icon) a
# partir du bandeau logo TELSAM.
#
#   powershell -ExecutionPolicy Bypass -File scripts\generer-icones.ps1
#
# Source : icons\logo-telsam-source.jpg  (ou .png). Bandeau, pas un carre :
# le logo TELSAM fait 550x291. Pour en faire une icone carree sans deformer le
# mot ni le poser sur un fond invente, on ETIRE VERTICALEMENT la premiere et la
# derniere ligne du bandeau. Ces deux lignes sont un aplat bleu marine : etire,
# il prolonge exactement le degrade du logo, sans raccord visible.
#
# Le mot "telsam" occupe alors 20%..87% en largeur et 41%..58% en hauteur de
# l'icone : il tient dans le cercle de securite de 80% qu'Android applique aux
# icones "maskable", et dans l'arrondi d'iOS. C'est pour ca que les icones sont
# declarees "any maskable" dans manifest.json -- si tu changes le cadrage,
# reverifie ce point avant de rediffuser.

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$racine = Split-Path -Parent $PSScriptRoot
$dossier = Join-Path $racine 'icons'

$source = Get-ChildItem -Path $dossier -Filter 'logo-telsam-source.*' |
          Where-Object { $_.Extension -match '^\.(png|jpg|jpeg)$' } |
          Select-Object -First 1
if (-not $source) {
  throw "Aucun logo trouve. Depose-le dans $dossier sous le nom logo-telsam-source.png (ou .jpg)."
}

$src = [System.Drawing.Image]::FromFile($source.FullName)
Write-Host "Source : $($source.Name)  $($src.Width)x$($src.Height)"

# Hauteur de la tranche recopiee en haut et en bas. Assez fine pour rester un
# aplat, assez epaisse pour que le lissage n'attrape pas le bord de l'image.
#
# Les DEUX bandes sont prises sur la PREMIERE ligne du logo, jamais la derniere :
# le bas du bandeau porte de fines diagonales blanches qui, etirees, donnent des
# trainees verticales sur toute la moitie basse de l'icone (essaye, ca se voit).
# Un miroir du bas est pire encore : il fait reapparaitre un "telsam" a l'envers.
$tranche = 3

function New-Icone {
  param([int]$Taille, [string]$Fichier)

  $bmp = New-Object System.Drawing.Bitmap($Taille, $Taille, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode  = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality

  # WrapMode TileFlipXY : sans lui, l'etirement d'une tranche de 3 pixels laisse
  # un liseré clair sur son bord (le lissage va chercher du vide hors image).
  $attr = New-Object System.Drawing.Imaging.ImageAttributes
  $attr.SetWrapMode([System.Drawing.Drawing2D.WrapMode]::TileFlipXY)

  $hBandeau = [int][Math]::Round($Taille * $src.Height / $src.Width)
  $y0 = [int][Math]::Round(($Taille - $hBandeau) / 2)

  # bande du haut = premiere ligne du logo, etiree
  if ($y0 -gt 0) {
    $g.DrawImage($src,
      (New-Object System.Drawing.Rectangle(0, 0, $Taille, $y0)),
      0, 0, $src.Width, $tranche,
      [System.Drawing.GraphicsUnit]::Pixel, $attr)
  }
  # bande du bas = derniere ligne du logo, etiree
  $yBas = $y0 + $hBandeau
  $hBas = $Taille - $yBas
  $ySrcBas = 0   # variante A : bande du bas prise en HAUT du logo
  if ($yBas -lt $Taille) {
    $g.DrawImage($src,
      (New-Object System.Drawing.Rectangle(0, $yBas, $Taille, $hBas)),
      0, $ySrcBas, $src.Width, $tranche,
      [System.Drawing.GraphicsUnit]::Pixel, $attr)
  }
  # le bandeau lui-meme, pleine largeur, centre
  $g.DrawImage($src,
    (New-Object System.Drawing.Rectangle(0, $y0, $Taille, $hBandeau)),
    0, 0, $src.Width, $src.Height,
    [System.Drawing.GraphicsUnit]::Pixel, $attr)

  $chemin = Join-Path $dossier $Fichier
  $bmp.Save($chemin, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose(); $attr.Dispose()
  Write-Host ("  {0,-24} {1}x{1}  {2} Ko" -f $Fichier, $Taille, [int]((Get-Item $chemin).Length / 1024))
}

New-Icone -Taille 512 -Fichier 'icon-512.png'
New-Icone -Taille 192 -Fichier 'icon-192.png'
New-Icone -Taille 180 -Fichier 'apple-touch-icon.png'   # iOS, ecran d'accueil
New-Icone -Taille 32  -Fichier 'favicon-32.png'         # onglet du navigateur

$src.Dispose()
Write-Host "Termine."
