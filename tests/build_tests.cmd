@echo off
rem Сборка и запуск юнит-тестов через dotnet SDK.
setlocal
dotnet build "%~dp0UnitTests.csproj" -c Release -v minimal --nologo
if errorlevel 1 (
    echo TESTS BUILD FAILED
    exit /b 1
)
"%~dp0bin\UnitTests.exe"
exit /b %ERRORLEVEL%
