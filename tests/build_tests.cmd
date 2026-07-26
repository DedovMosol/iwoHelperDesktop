@echo off
rem Сборка и запуск юнит-тестов через dotnet SDK.
rem Usage: build_tests.cmd [x86]  - по умолчанию x64. Разрядность важна: часть логики
rem ветвится по IntPtr.Size (размеры кэшей документов и страниц), и без x86-прогона
rem эти ветки не проверяются никогда.
setlocal
set ARCH=%~1
if "%ARCH%"=="" set ARCH=x64
if /I not "%ARCH%"=="x64" if /I not "%ARCH%"=="x86" (
    echo ERROR: unknown architecture "%ARCH%" ^(use x64 or x86^)
    exit /b 1
)
dotnet build "%~dp0UnitTests.csproj" -c Release -p:Arch=%ARCH% -v minimal --nologo
if errorlevel 1 (
    echo TESTS BUILD FAILED
    exit /b 1
)
if /I "%ARCH%"=="x86" (
    "%~dp0bin\x86\UnitTests.exe"
) else (
    "%~dp0bin\UnitTests.exe"
)
exit /b %ERRORLEVEL%
