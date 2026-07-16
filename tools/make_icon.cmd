@echo off
rem Перегенерация app.ico из tools\make_icon.cs (дизайн — assets\icon.svg).
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe

"%CSC%" /nologo /target:exe /codepage:65001 /out:"%TEMP%\make_icon_tmp.exe" ^
    /r:System.Drawing.dll "%~dp0make_icon.cs"
if errorlevel 1 exit /b 1

"%TEMP%\make_icon_tmp.exe" "%~dp0..\app.ico"
set RC=%ERRORLEVEL%
del "%TEMP%\make_icon_tmp.exe" >nul 2>&1
exit /b %RC%
