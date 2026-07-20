# Копирует минимальный рантайм Ghostscript в installer\gs\ для вшивания в установщик.
# Нужны: bin\gsdll64.dll (движок) + bin\gswin64c.exe (загрузчик) + lib\ + Resource\ +
# iccprofiles\ (без gsdll64.lib — это import-lib, в рантайме не нужна). Источник —
# установленный GS (Program Files\gs\gs* или %USERPROFILE%\gs*); можно задать -Source.
# installer\gs\ в .gitignore (крупные бинарники в репозиторий не кладём).
param(
    [string]$Source,
    [string]$Dest = (Join-Path (Split-Path $PSScriptRoot) 'installer\gs')
)
$ErrorActionPreference = 'Stop'

function Find-GsRoot {
    $roots = @()
    $bases = @()
    if ($env:ProgramFiles) { $bases += (Join-Path $env:ProgramFiles 'gs') }
    if (${env:ProgramFiles(x86)}) { $bases += (Join-Path ${env:ProgramFiles(x86)} 'gs') }
    if ($env:USERPROFILE) { $bases += $env:USERPROFILE }
    foreach ($b in $bases) {
        if (Test-Path $b) {
            Get-ChildItem -LiteralPath $b -Directory -Filter 'gs*' -ErrorAction SilentlyContinue | ForEach-Object {
                if (Test-Path (Join-Path $_.FullName 'bin\gsdll64.dll')) { $roots += $_.FullName }
            }
        }
    }
    # Самая свежая версия (сортировка по имени: gs10.07 > gs10.05 > gs9.56).
    return ($roots | Sort-Object -Descending | Select-Object -First 1)
}

if (-not $Source) { $Source = Find-GsRoot }
if (-not $Source -or -not (Test-Path $Source)) {
    Write-Host 'Ghostscript не найден. Установите его: https://ghostscript.com/releases/gsdnld.html'
    Write-Host 'или задайте -Source <корень gs, напр. C:\Program Files\gs\gs10.07>.'
    exit 1
}

$dll = Join-Path $Source 'bin\gsdll64.dll'
$exe = Join-Path $Source 'bin\gswin64c.exe'
if (-not (Test-Path $dll) -or -not (Test-Path $exe)) {
    Write-Host "В источнике нет bin\gsdll64.dll / bin\gswin64c.exe: $Source"
    exit 1
}

Write-Host "Источник Ghostscript: $Source"
if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $Dest 'bin') -Force | Out-Null

Copy-Item $dll (Join-Path $Dest 'bin\gsdll64.dll') -Force
Copy-Item $exe (Join-Path $Dest 'bin\gswin64c.exe') -Force
foreach ($sub in @('lib', 'Resource', 'iccprofiles')) {
    $s = Join-Path $Source $sub
    if (Test-Path $s) {
        Copy-Item $s (Join-Path $Dest $sub) -Recurse -Force
    } elseif ($sub -ne 'iccprofiles') {
        Write-Host "ВНИМАНИЕ: нет обязательного каталога '$sub' в $Source"
        exit 1
    }
}
# Лицензия GS (AGPL) — для комплаенса кладём рядом.
foreach ($lic in @('LICENSE', 'doc\COPYING', 'COPYING')) {
    $lp = Join-Path $Source $lic
    if (Test-Path $lp) { Copy-Item $lp (Join-Path $Dest 'LICENSE') -Force; break }
}

$size = [math]::Round(((Get-ChildItem $Dest -Recurse -File | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host "Ghostscript подготовлен в $Dest ($size МБ)."
exit 0
