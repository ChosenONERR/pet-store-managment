@echo off
title Verify Codex Data Migration
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0verify_codex_migration_offline.ps1"
echo.
echo Exit code: %ERRORLEVEL%
pause
