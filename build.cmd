@echo off
rem Сборка ExcelMerger.exe встроенным в Windows компилятором C# (.NET Framework 4.8).
rem Никаких внешних инструментов и упаковщиков — минимальный риск триггера антивируса.
setlocal

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo ERROR: csc.exe not found. .NET Framework 4.x is required.
    exit /b 1
)

"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 ^
    /out:"%~dp0ExcelMerger.exe" ^
    /win32manifest:"%~dp0app.manifest" ^
    /win32icon:"%~dp0app.ico" ^
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll ^
    /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll ^
    "%~dp0src\*.cs"

if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)
echo BUILD OK: %~dp0ExcelMerger.exe
