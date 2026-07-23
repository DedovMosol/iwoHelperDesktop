# Generate the Excel test corpus via COM automation of an installed Excel.
$ErrorActionPreference = 'Stop'
$dir = Join-Path $PSScriptRoot 'testdata'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Get-ChildItem $dir -File | Remove-Item -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    # 1. Normal file: formatting, merged cells, a formula, a fill
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
    # formula inside a merged cell - exercises value replacement with fallback
    $ws.Range("D1:E1").Merge()
    $ws.Range("D1").Formula = "=A2+A3"
    # string-result formulas: on value replacement Excel must not re-parse the
    # string (formula injection, number, date)
    $ws.Range("F1").Formula = '="="&"тест"'
    $ws.Range("G1").Formula = '="12 345"'
    $ws.Range("H1").Formula = '="01.02.2026"'
    $wb.SaveAs((Join-Path $dir "Отчет управления A.xlsx"), 51)
    $wb.Close($false)

    # 2. Empty file
    $wb = $xl.Workbooks.Add()
    $wb.SaveAs((Join-Path $dir "Пустой отчет.xlsx"), 51)
    $wb.Close($false)

    # 3. Old binary .xls format
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "старый формат xls"
    $wb.SaveAs((Join-Path $dir "Отчет управления B.xls"), 56)
    $wb.Close($false)

    # 4. Duplicate base name with a different extension (.xlsm) -> the sheet must get suffix _2
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "дубль имени"
    $wb.SaveAs((Join-Path $dir "Отчет управления A.xlsm"), 52)
    $wb.Close($false)

    # 5. Binary .xlsb format
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "формат xlsb"
    $wb.SaveAs((Join-Path $dir "Отчет управления C.xlsb"), 50)
    $wb.Close($false)

    # 6. First sheet hidden, second visible -> the second one must be taken
    $wb = $xl.Workbooks.Add()
    $second = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
    $wb.Sheets.Item(1).Range("A1").Value2 = "скрытый лист"
    $second.Range("A1").Value2 = "видимый лист"
    $wb.Sheets.Item(1).Visible = 0
    $wb.SaveAs((Join-Path $dir "Скрытый первый лист.xlsx"), 51)
    $wb.Close($false)

    # 7. Very long file name (>31 chars) - the sheet name must be truncated
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "длинное имя"
    $wb.SaveAs((Join-Path $dir "Очень длинное имя файла отчета за первый квартал 2026.xlsx"), 51)
    $wb.Close($false)

    # 7b. File name with brackets [ ] - Excel cannot open such files,
    #     the app must skip it with a reason, without crashing.
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "скобки"
    $wb.SaveAs((Join-Path $dir "Отчет скобки.xlsx"), 51)
    $wb.Close($false)
    Rename-Item (Join-Path $dir "Отчет скобки.xlsx") "Отчет [март 2026].xlsx"

    # 8. Natural order: "Отчет 2" must come before "Отчет 10"
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "второй"
    $wb.SaveAs((Join-Path $dir "Отчет 2.xlsx"), 51)
    $wb.Close($false)
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "десятый"
    $wb.SaveAs((Join-Path $dir "Отчет 10.xlsx"), 51)
    $wb.Close($false)

    # 9. Password-protected file -> must be skipped with a reason
    $wb = $xl.Workbooks.Add()
    $wb.Sheets.Item(1).Range("A1").Value2 = "секрет"
    $wb.SaveAs((Join-Path $dir "Запароленный.xlsx"), 51, "secret123")
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}

# 9. Broken file (not Excel)
Set-Content -Path (Join-Path $dir "Битый файл.xlsx") -Value "this is not an excel file" -Encoding ASCII

# 10. Excel temp file - must be ignored
Set-Content -Path (Join-Path $dir ('~$' + 'Отчет управления A.xlsx')) -Value "lock" -Encoding ASCII

Write-Host "TESTDATA OK:"
Get-ChildItem $dir -File | ForEach-Object { Write-Host ("  " + $_.Name) }
