@echo off
setlocal EnableExtensions
REM Conditional Access Manager Launcher
REM This batch file can launch either the EXE or PowerShell script version

cd /d "%~dp0"

echo ========================================
echo Conditional Access Manager
echo ========================================
echo.
echo Choose launch method:
echo   1. PowerShell Script (Recommended - better authentication)
echo   2. Executable (EXE)
echo.
set /p choice="Enter choice (1 or 2): "

if "%choice%"=="1" goto :script
if "%choice%"=="2" goto :exe
goto :script

:script
echo.
echo Launching PowerShell script version...
echo.
if not exist "%~dp0ca2.ps1" (
    echo ERROR: ca2.ps1 not found in "%~dp0"
    pause
    exit /b 1
)
REM -NoProfile avoids running profile scripts; -File runs this script only
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0ca2.ps1"
) else (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0ca2.ps1"
)
set "SCRIPT_ERR=%ERRORLEVEL%"
goto :end

:exe
echo.
echo Launching EXE version...
echo Note: EXE version may have authentication limitations.
echo.
if not exist "%~dp0ca_manager.exe" (
    echo ERROR: ca_manager.exe not found in "%~dp0"
    pause
    exit /b 1
)
start "" "%~dp0ca_manager.exe"
set "SCRIPT_ERR=0"
goto :end

:end
if defined SCRIPT_ERR (
    if not "%SCRIPT_ERR%"=="0" (
        echo.
        echo Error launching the script ^(exit code %SCRIPT_ERR%^).
        pause
    )
)
endlocal
