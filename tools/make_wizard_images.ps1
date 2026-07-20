# Генерирует фирменные картинки мастера установки Inno Setup из логотипа img\Logo.jpeg:
#   installer\wizard.bmp        (164x314) — большая, на страницах приветствия/финиша;
#   installer\wizard_small.bmp  (55x55)   — маленькая, вверху остальных страниц.
# Фон — фирменный синий градиент (Theme.HubBlue -> HubBlueDark), по центру — логотип,
# на большой под ним — подпись «iwo / Helper Desktop». 24-битный BMP (формат для Inno).
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$root    = Split-Path $PSScriptRoot
$logoPath = Join-Path $root 'img\Logo.jpeg'
$outDir  = Join-Path $root 'installer'

$blue     = [System.Drawing.Color]::FromArgb(15, 108, 189) # HubBlue  #0F6CBD
$blueDark = [System.Drawing.Color]::FromArgb(10, 78, 134)  # HubBlueDark #0A4E86
$logoImg  = [System.Drawing.Image]::FromFile($logoPath)

function New-WizardBmp([int]$w, [int]$h, [string]$path, [int]$logoBox, [bool]$showText) {
    $bmp = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
    $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $blue, $blueDark, 90.0)
    $g.FillRectangle($br, $rect)

    # Вписываем логотип в квадрат logoBox С СОХРАНЕНИЕМ ПРОПОРЦИЙ (без искажений), по центру.
    $scale = [Math]::Min($logoBox / $logoImg.Width, $logoBox / $logoImg.Height)
    $dw = [int]($logoImg.Width * $scale)
    $dh = [int]($logoImg.Height * $scale)
    $lx = [int](($w - $dw) / 2)
    $ly = if ($showText) { [int]($h * 0.15) } else { [int](($h - $dh) / 2) }
    $g.DrawImage($logoImg, $lx, $ly, $dw, $dh)

    if ($showText) {
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $f1 = New-Object System.Drawing.Font('Segoe UI', 30.0, [System.Drawing.FontStyle]::Bold)
        $f2 = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Regular)
        $r1 = New-Object System.Drawing.RectangleF(0, ($ly + $dh + 18), $w, 50)
        $g.DrawString('iwo', $f1, [System.Drawing.Brushes]::White, $r1, $sf)
        $r2 = New-Object System.Drawing.RectangleF(0, ($ly + $dh + 66), $w, 26)
        $g.DrawString('Helper Desktop', $f2, [System.Drawing.Brushes]::White, $r2, $sf)
        $f1.Dispose(); $f2.Dispose(); $sf.Dispose()
    }

    $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $bmp.Dispose(); $br.Dispose()
    Write-Host ("Создано: " + $path + " (" + $w + "x" + $h + ")")
}

New-WizardBmp 164 314 (Join-Path $outDir 'wizard.bmp') 120 $true
New-WizardBmp 55 55 (Join-Path $outDir 'wizard_small.bmp') 51 $false
$logoImg.Dispose()
