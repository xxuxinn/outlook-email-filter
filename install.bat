@echo off
REM ====================================================================
REM Outlook Email Filter v2.0 - Quick Installer
REM ====================================================================
REM This batch file runs the PowerShell installer with execution policy
REM bypass, so you don't need to change your PowerShell settings.
REM
REM USAGE: Double-click this file, or run: install.bat
REM ====================================================================

echo.
echo ===============================================================
echo   Outlook Email Filter v2.0 - Quick Installer
echo ===============================================================
echo.
echo This will install the email filter using PowerShell automation.
echo.
echo IMPORTANT: Close Outlook before continuing!
echo.
pause

powershell.exe -ExecutionPolicy Bypass -File "%~dp0Install-OutlookFilter.ps1"

echo.
echo.
pause
