@echo off
title Migracja CMT - Dataverse
cd /d "%~dp0"
echo Log sesji: %~dp0Logs\
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-CMTMigrationMenu.ps1"
if errorlevel 1 (
    echo.
    echo [BLAD] Aplikacja zakonczyla sie z bledem.
    echo Szczegoly powyzej lub w pliku Logs\CMT_Menu_*.log
    echo.
)
echo.
pause
