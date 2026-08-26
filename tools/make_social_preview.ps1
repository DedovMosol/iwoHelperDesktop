# Генерирует соц-превью 1280x640 (GitHub Settings -> Social preview) — фирменную карточку
# для шеринга ссылки. Берёт логотип и палитру из docs/screenshots/banner.jpg, кладёт название
# и краткий тезис на синий градиент. Результат: docs/screenshots/social-preview.png.
# Загружать вручную (API для соц-превью нет). Запускать при смене баннера.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$root = Join-Path $PSScriptRoot '..'
$banner = [System.Drawing.Image]::FromFile((Join-Path $root 'docs\screenshots\banner.jpg'))
$probe = New-Object System.Drawing.Bitmap $banner
$c1 = $probe.GetPixel(15, 15)                                  # тёмно-синий (верхний угол)
$c2 = $probe.GetPixel($banner.Width - 15, $banner.Height - 15) # светлее (нижний угол)

$W = 1280; $H = 640
$card = New-Object System.Drawing.Bitmap $W, $H
$g = [System.Drawing.Graphics]::FromImage($card)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

$rect = New-Object System.Drawing.Rectangle 0, 0, $W, $H
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $c1, $c2, 40.0
$g.FillRectangle($brush, $rect)

# Логотип: вырезаем квадрат из баннера (в 1600px-баннере он ~ (90,60)..(410,395)) и кладём слева.
$logoSrc = New-Object System.Drawing.Rectangle 88, 58, 322, 338
$logoDst = New-Object System.Drawing.Rectangle 150, 175, 290, 290
$g.DrawImage($banner, $logoDst, $logoSrc, [System.Drawing.GraphicsUnit]::Pixel)

$white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$soft = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(220, 230, 240))
$fName = New-Object System.Drawing.Font 'Segoe UI', 49, ([System.Drawing.FontStyle]::Bold)
$fTag = New-Object System.Drawing.Font 'Segoe UI', 18, ([System.Drawing.FontStyle]::Regular)
$fSub = New-Object System.Drawing.Font 'Segoe UI', 20, ([System.Drawing.FontStyle]::Regular)

$tx = 470
$arrow = [char]0x2192  # -> без не-ASCII в исходнике (PS 5.1 ломается на не-ASCII в строках)
$dot = [char]0x00B7
$g.DrawString('iwo Helper Desktop', $fName, $white, [single]$tx, [single]252)
$g.DrawString("Excel merge   $dot   PDF merge / split / compare", $fTag, $soft, [single]($tx + 3), [single]334)
$g.DrawString("PDF $arrow Word / PPTX   $dot   More operations   $dot   offline & free", $fSub, $soft, [single]($tx + 3), [single]368)

$g.Dispose(); $banner.Dispose(); $probe.Dispose()
$out = Join-Path $root 'docs\screenshots\social-preview.png'
$card.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$card.Dispose()
Write-Host ("wrote {0} ({1} KB)" -f $out, [int]((Get-Item $out).Length / 1024))
