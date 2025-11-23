@echo off
REM Folder Navigator Phase 3.1 Launcher
echo ========================================
echo  Folder Navigator Phase 3.1
echo  Starting...
echo ========================================
echo.

REM Execute PowerShell script in STA mode
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%~dp0FolderNavigator_Phase3_Complete.ps1"

REM Check error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo  Error occurred
    echo  Error Code: %ERRORLEVEL%
    echo ========================================
    pause
)
