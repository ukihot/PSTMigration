@echo off
rem Close Outlook, then double-click this file. Runs detect -> copy -> split/7z/upload.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0PSTMigration-manual.ps1"
pause
