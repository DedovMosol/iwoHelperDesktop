# Ужимает README-баннер до разумной ширины (по умолчанию 1600px — хватает и для retina при
# показе на ~720px), высокое качество ресемпла, JPEG q90. Исходник у нас был 5918px (в ~8 раз
# больше нужного). Запускать при замене баннера; результат (docs/screenshots/banner.jpg) коммитить.
param([int]$Width = 1600, [int]$Quality = 90)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$path = Join-Path $PSScriptRoot '..\docs\screenshots\banner.jpg'
# Грузим из копии в памяти, чтобы не держать файл замапленным (сохраняем в тот же путь).
$bytes = [System.IO.File]::ReadAllBytes($path)
$ms = New-Object System.IO.MemoryStream (,$bytes)
$img = [System.Drawing.Image]::FromStream($ms)
$w = $Width
$h = [int][math]::Round($img.Height * $w / $img.Width)
Write-Host "resize $($img.Width)x$($img.Height) -> ${w}x${h}"
$bmp = New-Object System.Drawing.Bitmap $w, $h
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.DrawImage($img, 0, 0, $w, $h)
$g.Dispose(); $img.Dispose(); $ms.Dispose()

$enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$p = New-Object System.Drawing.Imaging.EncoderParameters 1
$p.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)
$bmp.Save($path, $enc, $p)
$bmp.Dispose()
$kb = [int]((Get-Item $path).Length / 1024)
Write-Host "wrote $path (${kb} KB)"
