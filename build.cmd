@echo off
rem Сборка dist\iwoHelperDesktop.exe через dotnet SDK (.NET Framework 4.8, WinForms).
rem На выходе один exe: PdfSharp вшит ресурсом, WinRT-проекции — только компиляция.
rem На целевой машине ничего не ставится (нужен лишь .NET Framework 4.8, он есть).
setlocal
where dotnet >nul 2>&1
if errorlevel 1 (
    echo ERROR: не найден dotnet SDK. Установите .NET SDK 6+ для сборки.
    exit /b 1
)

dotnet build "%~dp0iwoHelperDesktop.csproj" -c Release -v minimal --nologo
if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)
echo BUILD OK: %~dp0dist\iwoHelperDesktop.exe
