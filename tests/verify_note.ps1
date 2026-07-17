# Интеграционный тест записки Word: exe загружается как .NET-сборка,
# записка генерируется из синтетического результата и проверяется через Word COM.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$fails = @()

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'dist\iwoHelperDesktop.exe'))

$result = New-Object ExcelMerger.MergeResult
$result.OutputPath = Join-Path $PSScriptRoot 'out\Свод.xlsx'
$ok = New-Object ExcelMerger.FileResult
$ok.FileName = 'А.xlsx'; $ok.FullPath = 'C:\in\А.xlsx'; $ok.Ok = $true; $ok.SheetName = 'А'
$skip = New-Object ExcelMerger.FileResult
$skip.FileName = 'Б.xlsx'; $skip.FullPath = 'C:\in\Б.xlsx'; $skip.Note = 'файл защищён паролем'
$result.Files.Add($ok)
$result.Files.Add($skip)
$result.OkCount = 1
$result.SkipCount = 1

$options = New-Object ExcelMerger.MergeOptions
$options.AddToc = $true

$notePath = Join-Path $PSScriptRoot 'out\Записка.docx'
Remove-Item $notePath -Force -ErrorAction SilentlyContinue

$content = [ExcelMerger.NoteText]::Build($result, 'C:\in', $options, (Get-Date '2026-07-16 14:05'))
[ExcelMerger.WordNoteWriter]::Write($content, $notePath)

if (-not (Test-Path $notePath)) { $fails += 'файл записки не создан' }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($notePath, $false, $true)

    $text = $doc.Content.Text
    if ($text -notmatch 'СПРАВКА') { $fails += 'нет заголовка' }
    if ($text -notmatch '16 июля 2026') { $fails += 'нет периода' }
    if ($text -notmatch 'Обработано файлов: 2') { $fails += 'нет счётчиков' }
    if ($text -notmatch 'файл защищён паролем') { $fails += 'нет причины пропуска' }
    if ($text -notmatch 'Составитель') { $fails += 'нет подписи' }

    if ($doc.Tables.Count -ne 1) { $fails += "таблиц $($doc.Tables.Count), ожидалась 1" }
    elseif ($doc.Tables.Item(1).Rows.Count -ne 2) { $fails += "строк таблицы $($doc.Tables.Item(1).Rows.Count), ожидалось 2" }

    # стандарт: поля 30/15/20/20 мм и шрифт основного текста
    $ps = $doc.PageSetup
    if ([math]::Abs($ps.LeftMargin - 85.05) -gt 1) { $fails += "левое поле $($ps.LeftMargin)pt, ожидалось ~85" }
    if ([math]::Abs($ps.RightMargin - 42.55) -gt 1) { $fails += "правое поле $($ps.RightMargin)pt" }
    $first = $doc.Paragraphs.Item(1).Range
    if ($first.Font.Name -ne 'Times New Roman') { $fails += "шрифт «$($first.Font.Name)»" }
    if ($first.Font.Size -ne 14) { $fails += "кегль $($first.Font.Size)" }

    $doc.Close($false)
}
finally {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY NOTE OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
