@echo off
rem Запуск полной пирамиды тестов (обёртка для двойного клика).
powershell -NoProfile -File "%~dp0run_all.ps1"
exit /b %ERRORLEVEL%
