@echo off
rem Double-click entry point for the script launcher.
rem Forwards Arena env vars from the parent shell (if set) and runs run.ps1.
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run.ps1" %*
