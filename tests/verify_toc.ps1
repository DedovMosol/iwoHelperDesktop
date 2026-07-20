# Проверка свода, собранного с флагами --toc --values:
# лист «Содержание» с гиперссылками и статусами, формулы заменены значениями.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'out\Свод_toc.xlsx'
$fails = @()

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($out, 0, $true)

    # 10 перенесённых листов + «Содержание»
    if ($wb.Sheets.Count -ne 11) { $fails += "листов $($wb.Sheets.Count), ожидалось 11" }

    $toc = $wb.Sheets.Item(1)
    if ($toc.Name -ne 'Содержание') { $fails += "первый лист «$($toc.Name)», а не «Содержание»" }

    # Шапка оглавления
    if ($toc.Range("B1").Value2 -ne 'Лист') { $fails += "шапка оглавления не найдена" }
    if (-not $toc.Range("A1").Font.Bold) { $fails += "шапка не выделена" }

    # Строки: 12 обработанных файлов -> нумерация в A13 равна 12
    if ($toc.Range("A13").Value2 -ne 12) { $fails += "строк оглавления: A13=$($toc.Range('A13').Value2), ожидалось 12" }

    # Гиперссылки только для перенесённых листов
    if ($toc.Hyperlinks.Count -ne 10) { $fails += "гиперссылок $($toc.Hyperlinks.Count), ожидалось 10" }
    $hl = $toc.Hyperlinks.Item(1)
    if ($hl.SubAddress -ne "'Отчет _март 2026_'!A1") { $fails += "первая ссылка ведёт на «$($hl.SubAddress)»" }

    # Пропущенные файлы: первая строка данных — «Битый файл.xlsx», подсвечен, без ссылки
    if ($toc.Range("C2").Value2 -ne 'Битый файл.xlsx') { $fails += "строка 2 оглавления: $($toc.Range('C2').Value2)" }
    if ($toc.Range("B2").Value2 -ne '—') { $fails += "у пропущенного файла есть имя листа" }
    if ($toc.Range("A2").Font.Color -ne 2237106) { $fails += "пропущенный файл не подсвечен" }

    # Формулы заменены значениями (включая формулу в объединённой ячейке)
    $wsA = $wb.Sheets.Item('Отчет управления A_2')
    if ($wsA.Range("A4").HasFormula) { $fails += "A4 осталась формулой" }
    if ($wsA.Range("A4").Value2 -ne 30) { $fails += "A4: $($wsA.Range('A4').Value2), ожидалось 30" }
    if ($wsA.Range("D1").HasFormula) { $fails += "D1 (объединённая) осталась формулой" }
    if ($wsA.Range("D1").Value2 -ne 30) { $fails += "D1: $($wsA.Range('D1').Value2), ожидалось 30" }
    if (-not $wsA.Range("D1").MergeCells) { $fails += "D1 потеряла объединение" }

    # Строковые результаты формул записаны буквально, без повторного разбора Excel
    if ($wsA.Range("F1").HasFormula) { $fails += "F1 осталась формулой" }
    if ($wsA.Range("F1").Value2 -cne '=тест') { $fails += "F1: «$($wsA.Range('F1').Value2)», ожидалось «=тест»" }
    $g1 = $wsA.Range("G1").Value2
    if ($g1 -isnot [string] -or $g1 -cne '12 345') { $fails += "G1: строка «12 345» превратилась в «$g1» ($($g1.GetType().Name))" }
    $h1 = $wsA.Range("H1").Value2
    if ($h1 -isnot [string] -or $h1 -cne '01.02.2026') { $fails += "H1: строка-дата превратилась в «$h1» ($($h1.GetType().Name))" }

    # Форматирование не пострадало от замены значений
    if (-not $wsA.Range("A1").Font.Bold) { $fails += "A1 потеряла жирный шрифт" }
    if ($wsA.Range("A2").Interior.Color -ne 65535) { $fails += "A2 потеряла заливку" }

    # Кнопка-ссылка «К оглавлению» на листе данных: наша фигура с гиперссылкой на «Содержание»
    $hasBtn = $false
    foreach ($shp in $wsA.Shapes) { if ($shp.Name -eq 'iwoTocLink') { $hasBtn = $true; break } }
    if (-not $hasBtn) { $fails += "на листе данных нет кнопки-ссылки на оглавление (iwoTocLink)" }
    $backOk = $false
    foreach ($lnk in $wsA.Hyperlinks) { if ($lnk.SubAddress -like "*Содержание*") { $backOk = $true; break } }
    if (-not $backOk) { $fails += "кнопка на листе данных не ссылается на «Содержание»" }
    # На самом оглавлении обратной кнопки быть не должно
    $tocHasBtn = $false
    foreach ($shp in $toc.Shapes) { if ($shp.Name -eq 'iwoTocLink') { $tocHasBtn = $true; break } }
    if ($tocHasBtn) { $fails += "на листе «Содержание» лишняя кнопка-ссылка на само себя" }

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY TOC OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
