<#
  Generates PWA icon PNGs (192, 512, 180-apple-touch, 32-favicon) using the exact
  same dendritic-tree glyph as the desktop app's build/make-icon.ps1, so the
  iPhone home-screen icon matches the Windows app icon.
#>
Add-Type -AssemblyName System.Drawing

function Col([string]$hex, [int]$a=255) {
  $c = [System.Drawing.ColorTranslator]::FromHtml($hex)
  return [System.Drawing.Color]::FromArgb($a, $c.R, $c.G, $c.B)
}

function RenderTree([int]$size) {
  $SS = 8
  $big = $size * $SS
  $bmp = New-Object System.Drawing.Bitmap $big, $big, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.Clear([System.Drawing.Color]::Transparent)
  $cx = $big / 2.0
  $outlineCol = Col "#12240f"

  $trunkW = $big * 0.15
  $trunkTop = $big * 0.60
  $trunkBottom = $big * 0.88
  $hubX = $cx; $hubY = $trunkTop
  $branchW = $big * 0.065
  $canopyCy = $big * 0.30
  $lobes = @(
    @(0,     ($canopyCy+$big*0.08), ($big*0.29), "#1f9e5e"),
    @((-$big*0.23), ($canopyCy-$big*0.01), ($big*0.22), "#27b06c"),
    @(($big*0.23),  ($canopyCy-$big*0.01), ($big*0.22), "#27b06c"),
    @(0,     ($canopyCy-$big*0.21), ($big*0.22), "#3fe089")
  )
  $branchTargets = @(
    @((-$big*0.21), ($canopyCy+$big*0.02)),
    @(0,            ($canopyCy-$big*0.10)),
    @(($big*0.21),  ($canopyCy+$big*0.02))
  )

  function DrawPass([bool]$outline) {
    $extra = if ($outline) { $big*0.030 } else { 0 }
    if (-not $outline) {
      $shR = $big * 0.20; $shY = $big * 0.90
      $shadowBrush = New-Object System.Drawing.SolidBrush((Col "#000000" 60))
      $g.FillEllipse($shadowBrush, ($cx - $shR), ($shY - $shR*0.28), ($shR*2), ($shR*0.56))
    }
    $tw = $trunkW + $extra*2
    $tCol = if ($outline) { $outlineCol } else { Col "#6b4527" }
    $trunkPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $tr = New-Object System.Drawing.RectangleF(($cx - $tw/2), ($trunkTop-$extra), $tw, ($trunkBottom - $trunkTop + $extra*2))
    $rad = $tw * 0.5
    $trunkPath.AddArc($tr.X, $tr.Y, $rad*2, $rad*2, 180, 90)
    $trunkPath.AddArc(($tr.Right-$rad*2), $tr.Y, $rad*2, $rad*2, 270, 90)
    $trunkPath.AddArc(($tr.Right-$rad*2), ($tr.Bottom-$rad*2), $rad*2, $rad*2, 0, 90)
    $trunkPath.AddArc($tr.X, ($tr.Bottom-$rad*2), $rad*2, $rad*2, 90, 90)
    $trunkPath.CloseFigure()
    $g.FillPath((New-Object System.Drawing.SolidBrush($tCol)), $trunkPath)

    foreach ($lobe in $lobes) {
      $r = $lobe[2] + $extra
      $col = if ($outline) { $outlineCol } else { Col $lobe[3] }
      $b = New-Object System.Drawing.SolidBrush($col)
      $g.FillEllipse($b, ($cx + $lobe[0] - $r), ($lobe[1] - $r), ($r*2), ($r*2))
    }

    $bCol = if ($outline) { $outlineCol } else { Col "#7a5330" }
    $bw = $branchW + $extra*2
    foreach ($bt in $branchTargets) {
      $pen = New-Object System.Drawing.Pen($bCol, $bw)
      $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
      $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
      $g.DrawLine($pen, (New-Object System.Drawing.PointF($hubX,$hubY)), (New-Object System.Drawing.PointF(($cx+$bt[0]),$bt[1])))
    }

    if (-not $outline) {
      $r = $big*0.06
      $ax = $cx + $big*0.10; $ay = $canopyCy - $big*0.04
      $accentOutline = New-Object System.Drawing.SolidBrush($outlineCol)
      $g.FillEllipse($accentOutline, ($ax-$r-$big*0.015), ($ay-$r-$big*0.015), (($r+$big*0.015)*2), (($r+$big*0.015)*2))
      $b = New-Object System.Drawing.SolidBrush((Col "#ff5f70"))
      $g.FillEllipse($b, ($ax-$r), ($ay-$r), ($r*2), ($r*2))
    }
  }

  DrawPass $true
  DrawPass $false
  $g.Dispose()

  $final = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $gf = [System.Drawing.Graphics]::FromImage($final)
  $gf.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $gf.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $gf.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
  $gf.DrawImage($bmp, 0, 0, $size, $size)
  $gf.Dispose()
  $bmp.Dispose()
  return $final
}

# iOS home-screen icons render on an opaque background (no transparency support in
# the springboard), so bake the dark app background behind the glyph for those sizes.
function RenderOpaque([int]$size) {
  $tree = RenderTree $size
  $final = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $gf = [System.Drawing.Graphics]::FromImage($final)
  $gf.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $gf.Clear((Col "#080a0f"))
  $gf.DrawImage($tree, 0, 0, $size, $size)
  $gf.Dispose()
  $tree.Dispose()
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
  $img = if ($t.opaque) { RenderOpaque $t.size } else { RenderTree $t.size }
  $img.Save((Join-Path $outDir $t.name), [System.Drawing.Imaging.ImageFormat]::Png)
  $img.Dispose()
  Write-Output "Wrote $($t.name) ($($t.size)px, opaque=$($t.opaque))"
}
