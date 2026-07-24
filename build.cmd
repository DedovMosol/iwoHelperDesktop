@echo off
rem Build dist\iwoHelperDesktop.exe with the dotnet SDK (.NET Framework 4.8, WinForms).
rem Output is a single exe: PdfSharp is embedded as a resource, WinRT projections are
rem compile-time only. Nothing is installed on the target machine (only .NET Framework 4.8
rem is required: it ships with Windows 10 1903+, and installs once on Windows 8.1).
rem Usage: build.cmd [x86]  - default is the x64 build (dist\), x86 goes to dist\x86\.
setlocal
set ARCH=%~1
if "%ARCH%"=="" set ARCH=x64
if /I not "%ARCH%"=="x64" if /I not "%ARCH%"=="x86" (
    echo ERROR: unknown architecture "%ARCH%" ^(use x64 or x86^)
    exit /b 1
)
where dotnet >nul 2>&1
if errorlevel 1 (
    echo ERROR: dotnet SDK not found. Install .NET SDK 6+ to build.
    exit /b 1
)

dotnet build "%~dp0iwoHelperDesktop.csproj" -c Release -p:Arch=%ARCH% -v minimal --nologo
if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)
if /I "%ARCH%"=="x86" (
    echo BUILD OK: %~dp0dist\x86\iwoHelperDesktop.exe
) else (
    echo BUILD OK: %~dp0dist\iwoHelperDesktop.exe
)
