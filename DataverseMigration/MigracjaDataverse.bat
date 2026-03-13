@echo off
title Migracja Dataverse
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-MigrationGUI.ps1"
if errorlevel 1 pause
