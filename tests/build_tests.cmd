@echo off
rem Сборка и запуск юнит-тестов: src\*.cs + tests\UnitTests.cs -> tests\UnitTests.exe.
rem Точка входа тестов выбирается ключом /main (в src есть свой Main).
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe

"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 ^
    /main:ExcelMerger.Tests.UnitTests ^
    /out:"%~dp0UnitTests.exe" ^
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll ^
    /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll ^
    /r:"%~dp0..\build\PdfSharp.dll" ^
    /resource:"%~dp0..\build\PdfSharp.dll",PdfSharp.dll ^
    "%~dp0..\src\*.cs" "%~dp0UnitTests.cs"
if errorlevel 1 (
    echo TESTS BUILD FAILED
    exit /b 1
)

"%~dp0UnitTests.exe"
exit /b %ERRORLEVEL%
