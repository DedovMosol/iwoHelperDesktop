# Генерирует флаги для стартового выбора языка в инсталляторе:
#   installer\flag_en.bmp (Union Jack) и installer\flag_ru.bmp (триколор РФ).
# Геометрия ЗЕРКАЛИТ src\Flags.cs (тот же дизайн, что в меню языка приложения), но крупнее
# (66x44) — под кнопки выбора языка мастера. 24-битный BMP (Inno TBitmapImage грузит надёжно).
# Запускать при изменении Flags.cs; результат (2 BMP) коммитить в installer\.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$W = 66; $H = 44
$outDir = Join-Path $PSScriptRoot '..\installer'
$fmt = [System.Drawing.Imaging.PixelFormat]::Format24bppRgb

function New-Canvas { New-Object System.Drawing.Bitmap $W, $H, $fmt }
function Save-Bmp([System.Drawing.Bitmap]$bmp, [string]$name) {
    $path = Join-Path $outDir $name
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $bmp.Dispose()
    Write-Host "wrote $path"
}
function Draw-Border([System.Drawing.Graphics]$g) {
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 150, 150))
    $g.DrawRectangle($pen, 0, 0, $W - 1, $H - 1)
    $pen.Dispose()
}

# --- Россия: три горизонтальные полосы (белая/синяя/красная) ---
$ru = New-Canvas
$g = [System.Drawing.Graphics]::FromImage($ru)
$band = [int]($H / 3)
$blue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(0, 57, 166))
$red  = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(213, 43, 30))
$g.FillRectangle([System.Drawing.Brushes]::White, 0, 0, $W, $band)
$g.FillRectangle($blue, 0, $band, $W, $band)
$g.FillRectangle($red, 0, 2 * $band, $W, $H - 2 * $band)
Draw-Border $g
$g.Dispose(); $blue.Dispose(); $red.Dispose()
Save-Bmp $ru 'flag_ru.bmp'

# --- Великобритания: Union Jack (диагонали под клипом, затем прямой крест) ---
$uk = New-Canvas
$g = [System.Drawing.Graphics]::FromImage($uk)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$bgBlue = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(1, 33, 105))
$g.FillRectangle($bgBlue, 0, 0, $W, $H)
$white = [System.Drawing.Color]::White
$redC  = [System.Drawing.Color]::FromArgb(200, 16, 46)
$tl = New-Object System.Drawing.Point 0, 0
$tr = New-Object System.Drawing.Point $W, 0
$bl = New-Object System.Drawing.Point 0, $H
$br = New-Object System.Drawing.Point $W, $H
# Диагонали не должны вылезать за флаг.
$g.Clip = New-Object System.Drawing.Region (New-Object System.Drawing.Rectangle 0, 0, $W, $H)
$wpen = New-Object System.Drawing.Pen $white, 11.0
$g.DrawLine($wpen, $tl, $br); $g.DrawLine($wpen, $tr, $bl); $wpen.Dispose()
$rpen = New-Object System.Drawing.Pen $redC, 5.0
$g.DrawLine($rpen, $tl, $br); $g.DrawLine($rpen, $tr, $bl); $rpen.Dispose()
$g.ResetClip()
# Прямой крест: белый шире, красный уже.
$wcross = New-Object System.Drawing.SolidBrush $white
$g.FillRectangle($wcross, [int]($W / 2) - 8, 0, 16, $H)
$g.FillRectangle($wcross, 0, [int]($H / 2) - 8, $W, 16)
$wcross.Dispose()
$rcross = New-Object System.Drawing.SolidBrush $redC
$g.FillRectangle($rcross, [int]($W / 2) - 5, 0, 11, $H)
$g.FillRectangle($rcross, 0, [int]($H / 2) - 5, $W, 11)
$rcross.Dispose()
Draw-Border $g
$g.Dispose(); $bgBlue.Dispose()
Save-Bmp $uk 'flag_en.bmp'

Write-Host 'OK'
