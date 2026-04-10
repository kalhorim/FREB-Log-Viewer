@echo off
REM FREB Log Viewer Launcher
REM Uses -STA to satisfy WinForms Single-Threaded Apartment requirement.
REM Uses -File (not -Command) for idiomatic script execution and safe path handling.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Start-FrebViewer.ps1"
exit /b %ERRORLEVEL%
