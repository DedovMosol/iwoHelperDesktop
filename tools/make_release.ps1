# Локальная публикация релиза на GitHub: собирает и ПОДПИСЫВАЕТ portable-exe и
# установщик (make_installer.ps1), извлекает заметки из docs\CHANGELOG.md для текущей
# версии и создаёт релиз с обоими артефактами. Подпись — самоподписанным сертификатом
# (только локально), поэтому релиз собирается здесь, а не в CI.
#
# По умолчанию — сухой прогон (готовит артефакты и заметки, ничего не публикует).
# Реальная публикация: tools\make_release.ps1 -Publish (нужны git-доступ на push тега
# и авторизованный gh; тег vX.Y.Z будет создан из версии exe).
param(
    [switch]$Publish,
    [switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$ps = (Get-Process -Id $PID).Path

# 1. Сборка + подпись обоих артефактов
if (-not $SkipBuild) {
    & $ps -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'make_installer.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'make_installer.ps1 failed' }
}

# 2. Версия и пути (единый источник — версия exe)
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
if (-not (Test-Path $exe)) { throw "нет $exe" }
$ver = (Get-Item $exe).VersionInfo.FileVersion
$ver3 = ($ver -split '\.')[0..2] -join '.'
$tag = "v$ver3"
$installer = Join-Path $root ("dist\iwoHelperDesktop-setup-$ver3.exe")
if (-not (Test-Path $installer)) { throw "нет установщика $installer" }

# 3. Заметки релиза из секции CHANGELOG для этой версии
$notes = Join-Path $root ("dist\release-notes-$ver3.md")
$changelog = [IO.File]::ReadAllLines((Join-Path $root 'docs\CHANGELOG.md'), [Text.Encoding]::UTF8)
$section = New-Object System.Collections.Generic.List[string]
$inSection = $false
foreach ($line in $changelog) {
    if ($line -match '^##\s*\[') {
        if ($inSection) { break }                         # начало следующей версии — стоп
        if ($line -match ("\[" + [regex]::Escape($ver3) + "\]")) { $inSection = $true }
    }
    elseif ($inSection) { $section.Add($line) }
}
if ($section.Count -eq 0) { throw "в docs\CHANGELOG.md нет секции [$ver3]" }
$body = "# iwo Helper Desktop $tag`r`n" + (($section -join "`r`n").Trim()) + "`r`n`r`n" +
    "**Файлы:** ``iwoHelperDesktop.exe`` — portable (запуск одним файлом); " +
    "``iwoHelperDesktop-setup-$ver3.exe`` — установщик со встроенным Ghostscript " +
    "(установка для текущего пользователя без прав администратора)."
[IO.File]::WriteAllText($notes, $body, (New-Object Text.UTF8Encoding($false)))
Write-Host "Заметки релиза: $notes"

# 4. Публикация
if (-not $Publish) {
    Write-Host ""
    Write-Host "СУХОЙ ПРОГОН. Артефакты и заметки готовы:"
    Write-Host "  tag       = $tag"
    Write-Host "  exe       = $exe"
    Write-Host "  installer = $installer"
    Write-Host "  notes     = $notes"
    Write-Host "Для публикации на GitHub: tools\make_release.ps1 -Publish"
    exit 0
}

# Тег
& git -C $root rev-parse --verify --quiet "refs/tags/$tag" | Out-Null
if ($LASTEXITCODE -ne 0) {
    & git -C $root tag $tag
    & git -C $root push origin $tag
    if ($LASTEXITCODE -ne 0) { throw "не удалось запушить тег $tag" }
}

# Релиз (создать или обновить), оба подписанных артефакта
& gh release view $tag *> $null
if ($LASTEXITCODE -ne 0) {
    & gh release create $tag $exe $installer --title $tag --notes-file $notes
} else {
    & gh release upload $tag $exe $installer --clobber
    & gh release edit $tag --notes-file $notes
}
if ($LASTEXITCODE -ne 0) { throw 'gh release failed' }
Write-Host "OK: релиз $tag опубликован (portable + установщик)."
exit 0
