<#
  Generates PWA icon PNGs (192, 512, 180-apple-touch, 32-favicon) from the official
  Arbor logo (logo-source.png — tree + goose, transparent background), the same
  source used by the desktop app's build/make-icon.ps1, so the iPhone home-screen
  icon matches the Windows app icon.

  The tree and goose are treated as two separate sprites and recomposed at
  matched height, rather than scaling the flat combined image. In the source
  art the tree is 58x151px (spans the full canvas height) while the goose is
  only 42x63px tucked in the lower-right - scaling the combined image as one
  unit always left the goose looking tiny relative to the tree, no matter how
  tightly the whole image was cropped. Rendering them as independent sprites
  at the same height keeps both clearly visible at small icon sizes.
#>
Add-Type -AssemblyName System.Drawing

function Col([string]$hex, [int]$a=255) {
  $c = [System.Drawing.ColorTranslator]::FromHtml($hex)
  return [System.Drawing.Color]::FromArgb($a, $c.R, $c.G, $c.B)
}

$sourcePath = Join-Path $PSScriptRoot "logo-source.png"
$sourceImg = [System.Drawing.Image]::FromFile($sourcePath)

# Measured directly from logo-source.png: the tree occupies columns 0-57, the
# goose columns 66-107 (a dead gap of empty columns 58-65 separates them).
$treeBounds = New-Object System.Drawing.Rectangle(0, 0, 58, 151)
$gooseBounds = New-Object System.Drawing.Rectangle(66, 89, 42, 63)

function RenderComposition([int]$size, [double]$sideMargin = 0.03, [double]$gap = 0.035) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)

  $treeAspect = $treeBounds.Width / $treeBounds.Height
  $gooseAspect = $gooseBounds.Width / $gooseBounds.Height

  # Both sprites are drawn at the same height H, side by side, so neither
  # reads as an afterthought next to the other. H is solved so tree+gap+goose
  # (plus side margins) exactly fills the square.
  $availW = ($size * (1 - 2 * $sideMargin)) - ($size * $gap)
  $H = $availW / ($treeAspect + $gooseAspect)
  $treeW = $H * $treeAspect
  $gooseW = $H * $gooseAspect

  $topMargin = ($size - $H) / 2
  $treeX = $size * $sideMargin
  $treeY = $topMargin
  $gooseX = $treeX + $treeW + ($size * $gap)
  $gooseY = $topMargin

  $destTree = New-Object System.Drawing.RectangleF($treeX, $treeY, $treeW, $H)
  $destGoose = New-Object System.Drawing.RectangleF($gooseX, $gooseY, $gooseW, $H)
  $g.DrawImage($sourceImg, $destTree, $treeBounds, [System.Drawing.GraphicsUnit]::Pixel)
  $g.DrawImage($sourceImg, $destGoose, $gooseBounds, [System.Drawing.GraphicsUnit]::Pixel)

  $g.Dispose()
  return $bmp
}

# iOS home-screen icons render on an opaque background (no transparency support in
# the springboard), so bake the dark app background behind the glyph for those sizes.
function RenderOpaque([int]$size, [double]$sideMargin = 0.03, [double]$gap = 0.035) {
  $logo = RenderComposition $size $sideMargin $gap
  $final = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $gf = [System.Drawing.Graphics]::FromImage($final)
  $gf.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $gf.Clear((Col "#080a0f"))
  $gf.DrawImage($logo, 0, 0, $size, $size)
  $gf.Dispose()
  $logo.Dispose()
  return $final
}

$outDir = Join-Path $PSScriptRoot "icons"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$targets = @(
  @{size=192; name="icon-192.png"; opaque=$false},
  @{size=512; name="icon-512.png"; opaque=$false},
  @{size=180; name="apple-touch-icon.png"; opaque=$true},
  @{size=32;  name="favicon-32.png"; opaque=$false},
  # Maskable icons (purpose="maskable" in manifest.json): Chromium/Windows
  # renders these full-bleed and trusts the source to keep content inside
  # a safe zone, instead of applying its own extra shrink to avoid clipping
  # by the taskbar's rounded-corner mask. Needs a solid background
  # (transparency would show through the mask).
  @{size=192; name="icon-192-maskable.png"; opaque=$true},
  @{size=512; name="icon-512-maskable.png"; opaque=$true}
)
foreach ($t in $targets) {
  $img = if ($t.opaque) { RenderOpaque $t.size } else { RenderComposition $t.size }
  $img.Save((Join-Path $outDir $t.name), [System.Drawing.Imaging.ImageFormat]::Png)
  $img.Dispose()
  Write-Output "Wrote $($t.name) ($($t.size)px, opaque=$($t.opaque))"
}

$sourceImg.Dispose()
