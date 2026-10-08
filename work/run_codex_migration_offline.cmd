@echo off
title Codex Offline Data Migration
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_codex_migration_offline.ps1"
echo.
echo Exit code: %ERRORLEVEL%
pause
