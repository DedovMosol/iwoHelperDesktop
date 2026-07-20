# Локальная сборка ПОДПИСАННОГО установщика iwo Helper Desktop с вшитым Ghostscript.
# Шаги: build.cmd -> подпись exe -> stage_gs.ps1 -> ISCC -> подпись setup.exe.
# Подпись — самоподписанным сертификатом (tools\sign.ps1), только локально (сертификат
# в Cert:\CurrentUser\My, на CI его нет). Требуется установленный Inno Setup 6 (ISCC.exe)
# и Ghostscript (для stage). Дочерние .ps1 зовём отдельным процессом: их `exit` иначе
# завершил бы этот скрипт.
param(
    [switch]$SkipBuild,
    [string]$Iscc
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
$ps = (Get-Process -Id $PID).Path  # текущий хост PowerShell

function Invoke-Child([string]$script, [string[]]$scriptArgs) {
    & $ps -NoProfile -ExecutionPolicy Bypass -File $script @scriptArgs
    if ($LASTEXITCODE -ne 0) { throw "Ошибка: $script (код $LASTEXITCODE)" }
}

function Find-Iscc {
    # Любая версия Inno Setup (6/7/…) и любой диск установки. Порядок: реестр
    # (InstallLocation — надёжнее всего, не зависит от %ProgramFiles%) -> glob по
    # Program Files -> PATH. Выбираем свежайшую версию.
    $found = @()

    $unins = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    foreach ($k in $unins) {
        Get-ItemProperty $k -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like 'Inno Setup*' -and $_.InstallLocation } |
            ForEach-Object {
                $iscc = Join-Path $_.InstallLocation 'ISCC.exe'
                if (Test-Path $iscc) { $found += $iscc }
            }
    }

    $bases = @()
    if (${env:ProgramFiles(x86)}) { $bases += ${env:ProgramFiles(x86)} }
    if ($env:ProgramFiles) { $bases += $env:ProgramFiles }
    foreach ($b in $bases) {
        Get-ChildItem -LiteralPath $b -Directory -Filter 'Inno Setup *' -ErrorAction SilentlyContinue | ForEach-Object {
            $iscc = Join-Path $_.FullName 'ISCC.exe'
            if (Test-Path $iscc) { $found += $iscc }
        }
    }

    if ($found.Count) { return ($found | Sort-Object -Descending -Unique | Select-Object -First 1) }
    $cmd = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

# 1. Сборка приложения
if (-not $SkipBuild) {
    & (Join-Path $root 'build.cmd')
    if ($LASTEXITCODE -ne 0) { throw 'build.cmd failed' }
}
if (-not (Test-Path $exe)) { throw "нет $exe (сначала сборка)" }

# 2. Подпись exe (установщик вшивает уже подписанный файл)
Invoke-Child (Join-Path $PSScriptRoot 'sign.ps1') @('-ExePath', $exe)

# 3. Подготовка вшиваемого Ghostscript
Invoke-Child (Join-Path $PSScriptRoot 'stage_gs.ps1') @()

# 4. Версия (3 компонента) из версии exe — единый источник истины
$ver = (Get-Item $exe).VersionInfo.FileVersion   # напр. 1.13.0.0
$ver3 = ($ver -split '\.')[0..2] -join '.'

# 5. Компиляция установщика
if (-not $Iscc) { $Iscc = Find-Iscc }
if (-not $Iscc) {
    Write-Host 'Не найден ISCC.exe (Inno Setup). Установите: https://jrsoftware.org/isdl.php'
    exit 1
}
& $Iscc "/DAppVersion=$ver3" (Join-Path $root 'installer\iwoHelperDesktop.iss')
if ($LASTEXITCODE -ne 0) { throw 'ISCC failed' }

$setup = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3.exe")
if (-not (Test-Path $setup)) { throw "установщик не создан: $setup" }

# 6. Подпись установщика
Invoke-Child (Join-Path $PSScriptRoot 'sign.ps1') @('-ExePath', $setup)

Write-Host ""
Write-Host "OK: $setup"
exit 0
