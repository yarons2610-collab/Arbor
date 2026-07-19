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

function RenderLogo([int]$size) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)

  $margin = 0.05
  $availW = $size * (1 - $margin * 2)
  $availH = $size * (1 - $margin * 2)
  $scale = [Math]::Min($availW / $sourceImg.Width, $availH / $sourceImg.Height)
  $drawW = $sourceImg.Width * $scale
  $drawH = $sourceImg.Height * $scale
  $drawX = ($size - $drawW) / 2
  $drawY = ($size - $drawH) / 2
  $g.DrawImage($sourceImg, $drawX, $drawY, $drawW, $drawH)

  $g.Dispose()
  return $bmp
}

# iOS home-screen icons render on an opaque background (no transparency support in
# the springboard), so bake the dark app background behind the glyph for those sizes.
function RenderOpaque([int]$size) {
  $logo = RenderLogo $size
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
  @{size=32;  name="favicon-32.png"; opaque=$false}
)
foreach ($t in $targets) {
  $img = if ($t.opaque) { RenderOpaque $t.size } else { RenderLogo $t.size }
  $img.Save((Join-Path $outDir $t.name), [System.Drawing.Imaging.ImageFormat]::Png)
  $img.Dispose()
  Write-Output "Wrote $($t.name) ($($t.size)px, opaque=$($t.opaque))"
}

$sourceImg.Dispose()
