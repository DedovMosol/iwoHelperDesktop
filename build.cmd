@echo off
rem Build dist\iwoHelperDesktop.exe with the dotnet SDK (.NET Framework 4.8, WinForms).
rem Output is a single exe: PdfSharp is embedded as a resource, WinRT projections are
rem compile-time only. Nothing is installed on the target machine (only .NET Framework 4.8
rem is required, and it ships with Windows 10 1903+).
setlocal
where dotnet >nul 2>&1
if errorlevel 1 (
    echo ERROR: dotnet SDK not found. Install .NET SDK 6+ to build.
    exit /b 1
)

dotnet build "%~dp0iwoHelperDesktop.csproj" -c Release -v minimal --nologo
if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)
echo BUILD OK: %~dp0dist\iwoHelperDesktop.exe
