@echo off
REM Folder Navigator Phase 3.6 Launcher
echo ========================================
echo  Folder Navigator Phase 3.6
echo  Starting...
echo ========================================
echo.

REM Execute PowerShell script in STA mode
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0FolderNavigator_Phase3.6_Complete.ps1"

REM Check error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo  Error occurred
    echo  Error Code: %ERRORLEVEL%
    echo ========================================
    pause
)

