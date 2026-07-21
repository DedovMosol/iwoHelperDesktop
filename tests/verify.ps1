# Verify the output file: sheet set, values, a formula, formatting.
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'out\Свод.xlsx'
$fails = @()

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($out, 0, $true)

    $names = @()
    foreach ($s in $wb.Sheets) { $names += $s.Name }
    Write-Host ("SHEETS: " + ($names -join ' | '))

    # Natural order (like Explorer): "Отчет 2" before "Отчет 10".
    $expected = @(
        'Отчет _март 2026_',
        'Отчет 2',
        'Отчет 10',
        'Отчет управления A',
        'Отчет управления A_2',
        'Отчет управления B',
        'Отчет управления C',
        'Очень длинное имя файла отчета',
        'Пустой отчет',
        'Скрытый первый лист'
    )
    if (($names -join ';') -ne ($expected -join ';')) { $fails += "состав/порядок листов не совпал" }

    $wsA = $wb.Sheets.Item('Отчет управления A_2')  # source .xlsx with formatting
    if ($wsA.Range("A4").Value2 -ne 30) { $fails += "формула A4: ожидалось 30, получено $($wsA.Range('A4').Value2)" }
    if ($wsA.Range("A4").Formula -ne '=SUM(A2:A3)') { $fails += "A4 не формула: $($wsA.Range('A4').Formula)" }
    if (-not $wsA.Range("A1").MergeCells) { $fails += "A1 не объединена" }
    if (-not $wsA.Range("A1").Font.Bold) { $fails += "A1 не жирный" }
    if ($wsA.Range("A2").Interior.Color -ne 65535) { $fails += "A2 без заливки" }
    if ([math]::Abs($wsA.Columns.Item(1).ColumnWidth - 25) -gt 0.5) { $fails += "ширина колонки не перенесена: $($wsA.Columns.Item(1).ColumnWidth)" }
    if ($wsA.Range("D1").Formula -ne '=A2+A3') { $fails += "D1: формула в объединённой ячейке не перенесена" }
    if (-not $wsA.Range("D1").MergeCells) { $fails += "D1 не объединена" }
    # without the "values" option, string-result formulas stay formulas
    if (-not $wsA.Range("F1").HasFormula) { $fails += "F1 потеряла формулу" }
    if (-not $wsA.Range("G1").HasFormula) { $fails += "G1 потеряла формулу" }

    $wsH = $wb.Sheets.Item('Скрытый первый лист')
    if ($wsH.Range("A1").Value2 -ne 'видимый лист') { $fails += "взят не видимый лист: '$($wsH.Range('A1').Value2)'" }
    if ($wsH.Visible -ne -1) { $fails += "лист из файла со скрытым листом сам скрыт" }

    foreach ($s in $wb.Sheets) {
        if ($s.Visible -ne -1) { $fails += "скрытый лист в результате: $($s.Name)" }
    }

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

if ($fails.Count -eq 0) {
    Write-Host "VERIFY OK"
    exit 0
} else {
    $fails | ForEach-Object { Write-Host ("FAIL: " + $_) }
    exit 1
}
