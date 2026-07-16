# Генерация тестового корпуса Excel-файлов через COM установленного Excel.
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot 'testdata'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Get-ChildItem $dir -File | Remove-Item -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    # 1. Обычный файл: форматирование, объединённые ячейки, формула, заливка
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Sheets.Item(1)
    $ws.Range("A1").Value2 = "Отчет управления A"
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1:C1").Merge()
    $ws.Range("A2").Value2 = 10
    $ws.Range("A3").Value2 = 20
    $ws.Range("A4").Formula = "=SUM(A2:A3)"
    $ws.Range("A2").Interior.Color = 65535
    $ws.Columns.Item(1).ColumnWidth = 25
    $wb.SaveAs((Join-Path $dir "Отчет управления A.xlsx"), 51)
    $wb.Close($false)

    # 2. Пустой файл
    $wb = $xl.Workbooks.Add()
    $wb.SaveAs((Join-Path $dir "Пустой отчет.xlsx"), 51)
    $wb.Close($false)

    # 3. Старый бинарный формат .xls
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "старый формат xls"
    $wb.SaveAs((Join-Path $dir "Отчет управления B.xls"), 56)
    $wb.Close($false)

    # 4. Дубль базового имени с другим расширением (.xlsm) -> лист должен получить суффикс _2
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "дубль имени"
    $wb.SaveAs((Join-Path $dir "Отчет управления A.xlsm"), 52)
    $wb.Close($false)

    # 5. Двоичный формат .xlsb
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "формат xlsb"
    $wb.SaveAs((Join-Path $dir "Отчет управления C.xlsb"), 50)
    $wb.Close($false)

    # 6. Первый лист скрыт, второй видимый -> должен перенестись второй
    $wb = $xl.Workbooks.Add()
    $second = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
    $wb.Sheets.Item(1).Range("A1").Value2 = "скрытый лист"
    $second.Range("A1").Value2 = "видимый лист"
    $wb.Sheets.Item(1).Visible = 0
    $wb.SaveAs((Join-Path $dir "Скрытый первый лист.xlsx"), 51)
    $wb.Close($false)

    # 7. Очень длинное имя файла (>31 символа) — имя листа должно обрезаться
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "длинное имя"
    $wb.SaveAs((Join-Path $dir "Очень длинное имя файла отчета министерства за март 2026.xlsx"), 51)
    $wb.Close($false)

    # 7б. Имя файла со скобками [ ] — Excel не открывает такие файлы,
    #     программа должна пропустить его с причиной, не падая.
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "скобки"
    $wb.SaveAs((Join-Path $dir "Отчет скобки.xlsx"), 51)
    $wb.Close($false)
    Rename-Item (Join-Path $dir "Отчет скобки.xlsx") "Отчет [март 2026].xlsx"

    # 8. Запароленный файл -> должен быть пропущен с причиной
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "секрет"
    $wb.SaveAs((Join-Path $dir "Запароленный.xlsx"), 51, "secret123")
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

# 9. Битый файл (не Excel)
Set-Content -Path (Join-Path $dir "Битый файл.xlsx") -Value "this is not an excel file" -Encoding ASCII

# 10. Временный файл Excel — должен игнорироваться
Set-Content -Path (Join-Path $dir ('~$' + 'Отчет управления A.xlsx')) -Value "lock" -Encoding ASCII

Write-Host "TESTDATA OK:"
Get-ChildItem $dir -File | ForEach-Object { Write-Host ("  " + $_.Name) }
