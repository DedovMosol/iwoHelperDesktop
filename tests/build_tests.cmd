@echo off
rem Сборка и запуск юнит-тестов через dotnet SDK.
rem Usage: build_tests.cmd [x64|x86] [--filter подстрока] [--list] [--timeout секунды]
rem   по умолчанию x64. Разрядность важна: часть логики ветвится по IntPtr.Size (размеры
rem   кэшей документов и страниц), и без x86-прогона эти ветки не проверяются никогда.
rem   Остальные аргументы передаются раннеру как есть (фильтр по имени теста и т.п.).
setlocal
rem Каталог скрипта — ДО shift: shift сдвигает и %0, после него %~dp0 указывал бы не туда.
set HERE=%~dp0
set ARCH=%~1
if "%ARCH%"=="" set ARCH=x64
if /I not "%ARCH%"=="x64" if /I not "%ARCH%"=="x86" (
    echo ERROR: unknown architecture "%ARCH%" ^(use x64 or x86^)
    exit /b 1
)
set REST=
:collect
shift
if "%~1"=="" goto build
set REST=%REST% %1
goto collect
:build
dotnet build "%HERE%UnitTests.csproj" -c Release -p:Arch=%ARCH% -v minimal --nologo
if errorlevel 1 (
    echo TESTS BUILD FAILED
    exit /b 1
)
if /I "%ARCH%"=="x86" (
    "%HERE%bin\x86\UnitTests.exe"%REST%
) else (
    "%HERE%bin\UnitTests.exe"%REST%
)
exit /b %ERRORLEVEL%
