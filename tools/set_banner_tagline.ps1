# Перерисовывает СТРОКУ ВОЗМОЖНОСТЕЙ на README-баннере, не трогая ни фон, ни логотип, ни
# заголовок. Нужен, потому что баннер — готовая картинка с текстом внутри: при появлении
# нового инструмента строка устаревает, а перерисовывать баннер целиком значит подбирать
# заново градиент и плитку логотипа.
#
# Фон под строкой не «заливается похожим цветом», а ВОССТАНАВЛИВАЕТСЯ: каждый пиксель полосы
# берётся линейной интерполяцией между чистыми строками выше и ниже неё. Градиент баннера
# при этом воспроизводится точно, каким бы он ни был, — формулу знать не нужно.
#
# Usage: tools\set_banner_tagline.ps1 -Text "Excel merge | ... | offline & free"
param(
    [string]$Text = 'Excel merge  |  PDF merge / split / compress  |  PDF → Word / PowerPoint  |  offline & free',
    [int]$Quality = 90
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$path = Join-Path $PSScriptRoot '..\docs\screenshots\banner.jpg'
# Полоса строки возможностей — измерена по самой картинке (светлые пиксели правее логотипа).
$bandTop = 256
$bandBottom = 296

# Грузим из копии в памяти: иначе файл остаётся замапленным и в него же не записать.
$bytes = [System.IO.File]::ReadAllBytes($path)
$ms = New-Object System.IO.MemoryStream (,$bytes)
$src = [System.Drawing.Image]::FromStream($ms)
$bmp = New-Object System.Drawing.Bitmap $src
$src.Dispose(); $ms.Dispose()

$above = $bandTop - 6
$below = $bandBottom + 6
if ($above -lt 0 -or $below -ge $bmp.Height) { throw "Полоса выходит за пределы баннера" }

# Восстанавливаем фон ТОЛЬКО правее плитки логотипа: полоса пересекает плитку по высоте, и
# интерполяция размазала бы по ней буквы логотипа сверху вниз. Правый край плитки измерен
# по самой картинке (яркий синий 0,102,255): 389 px, плюс запас.
$fromX = 400
for ($x = $fromX; $x -lt $bmp.Width; $x++) {
    $top = $bmp.GetPixel($x, $above)
    $bottom = $bmp.GetPixel($x, $below)
    for ($y = $bandTop; $y -le $bandBottom; $y++) {
        $t = ($y - $above) / [double]($below - $above)
        $r = [int][math]::Round($top.R + ($bottom.R - $top.R) * $t)
        $g = [int][math]::Round($top.G + ($bottom.G - $top.G) * $t)
        $b = [int][math]::Round($top.B + ($bottom.B - $top.B) * $t)
        $bmp.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($r, $g, $b))
    }
}

$g2 = [System.Drawing.Graphics]::FromImage($bmp)
$g2.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
$g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
# Кегль подбирается ПОД ШИРИНУ зоны: прежний соответствовал прежнему тексту, а новый
# длиннее и обрезался бы справа. Начинаем с исходных 28 px и уменьшаем, пока не влезет.
$zoneWidth = $bmp.Width - 420 - 40
$size = 28.0
$font = $null
while ($size -ge 16.0) {
    if ($font) { $font.Dispose() }
    $font = New-Object System.Drawing.Font('Segoe UI', $size, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $measured = $g2.MeasureString($Text, $font)
    if ($measured.Width -le $zoneWidth) { break }
    $size -= 0.5
}
Write-Host ("tagline font {0} px, width {1} of {2}" -f $size, [int]$g2.MeasureString($Text, $font).Width, $zoneWidth)
$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 245, 255))
$fmt = New-Object System.Drawing.StringFormat
$fmt.Alignment = [System.Drawing.StringAlignment]::Center
$fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
# Текст центрируется в той же зоне, что и заголовок: правее плитки логотипа.
$zone = New-Object System.Drawing.RectangleF 420, $bandTop, ($bmp.Width - 420 - 40), ($bandBottom - $bandTop + 1)
$g2.DrawString($Text, $font, $brush, $zone, $fmt)
$g2.Dispose(); $brush.Dispose(); $font.Dispose(); $fmt.Dispose()

$enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$p = New-Object System.Drawing.Imaging.EncoderParameters 1
$p.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]$Quality)
$bmp.Save($path, $enc, $p)
$bmp.Dispose()
Write-Host ("wrote {0} ({1} KB)" -f $path, [int]((Get-Item $path).Length / 1024))
