<#
  Generates PWA icon PNGs (192, 512, 180-apple-touch, 32-favicon) from the official
  Arbor logo (logo-source.png — tree + goose, transparent background), the same
  source used by the desktop app's build/make-icon.ps1, so the iPhone home-screen
  icon matches the Windows app icon.
#>
Add-Type -AssemblyName System.Drawing

function Col([string]$hex, [int]$a=255) {
  $c = [System.Drawing.ColorTranslator]::FromHtml($hex)
  return [System.Drawing.Color]::FromArgb($a, $c.R, $c.G, $c.B)
}

$sourcePath = Join-Path $PSScriptRoot "logo-source.png"
$sourceImg = [System.Drawing.Image]::FromFile($sourcePath)
$sourceBmp = New-Object System.Drawing.Bitmap $sourceImg

# The source PNG has dead transparent padding baked into its export around the
# actual tree+goose artwork. Find the tight bounding box of non-transparent
# pixels so we scale the *artwork*, not the padded canvas, to fill the icon.
function FindContentBounds([System.Drawing.Bitmap]$bmp) {
  $minX = $bmp.Width; $minY = $bmp.Height; $maxX = 0; $maxY = 0
  for ($y = 0; $y -lt $bmp.Height; $y++) {
    for ($x = 0; $x -lt $bmp.Width; $x++) {
      if ($bmp.GetPixel($x, $y).A -gt 10) {
        if ($x -lt $minX) { $minX = $x }
        if ($x -gt $maxX) { $maxX = $x }
        if ($y -lt $minY) { $minY = $y }
        if ($y -gt $maxY) { $maxY = $y }
      }
    }
  }
  return New-Object System.Drawing.Rectangle($minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1))
}
$contentBounds = FindContentBounds $sourceBmp

function RenderLogo([int]$size, [double]$margin = 0.01) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)

  $availW = $size * (1 - $margin * 2)
  $availH = $size * (1 - $margin * 2)
  # Full cover-crop cuts off the treetop, trunk, and goose feet - all
  # recognizable parts of the logo - because the art is portrait (109x152)
  # and a square cover-crop needs ~28% cropped off the height. Contain (Min)
  # is the non-destructive choice: it fills the full height edge-to-edge
  # (this art's constrained axis) and only leaves margin on the sides.
  $scale = [Math]::Min($availW / $contentBounds.Width, $availH / $contentBounds.Height)
  $drawW = $contentBounds.Width * $scale
  $drawH = $contentBounds.Height * $scale
  $drawX = ($size - $drawW) / 2
  $drawY = ($size - $drawH) / 2
  $destRect = New-Object System.Drawing.RectangleF($drawX, $drawY, $drawW, $drawH)
  $g.DrawImage($sourceImg, $destRect, $contentBounds, [System.Drawing.GraphicsUnit]::Pixel)

  $g.Dispose()
  return $bmp
}

# iOS home-screen icons render on an opaque background (no transparency support in
# the springboard), so bake the dark app background behind the glyph for those sizes.
function RenderOpaque([int]$size, [double]$margin = 0.01) {
  $logo = RenderLogo $size $margin
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
  @{size=192; name="icon-192.png"; opaque=$false; margin=0.01},
  @{size=512; name="icon-512.png"; opaque=$false; margin=0.01},
  @{size=180; name="apple-touch-icon.png"; opaque=$true; margin=0.01},
  @{size=32;  name="favicon-32.png"; opaque=$false; margin=0.01},
  # Maskable icons (purpose="maskable" in manifest.json): Chromium/Windows
  # renders these full-bleed and trusts the source to keep content inside
  # a safe zone, instead of applying its own extra shrink to avoid clipping
  # by the taskbar's rounded-corner mask - which is what was making the
  # regular icon look small even with an opaque background. Needs a solid
  # background (transparency would show through the mask) and enough margin
  # (0.1, i.e. content in the middle 80%) that the corner mask never bites in.
  @{size=192; name="icon-192-maskable.png"; opaque=$true; margin=0.1},
  @{size=512; name="icon-512-maskable.png"; opaque=$true; margin=0.1}
)
foreach ($t in $targets) {
  $img = if ($t.opaque) { RenderOpaque $t.size $t.margin } else { RenderLogo $t.size $t.margin }
  $img.Save((Join-Path $outDir $t.name), [System.Drawing.Imaging.ImageFormat]::Png)
  $img.Dispose()
  Write-Output "Wrote $($t.name) ($($t.size)px, opaque=$($t.opaque), margin=$($t.margin))"
}

$sourceImg.Dispose()
