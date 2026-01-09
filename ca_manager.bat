@echo off
REM Conditional Access Manager Launcher
REM This batch file can launch either the EXE or PowerShell script version

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
REM Check if PowerShell is available
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    REM Use PowerShell Core if available
    pwsh.exe -ExecutionPolicy Bypass -File "%~dp0ca2.ps1"
) else (
    REM Fall back to Windows PowerShell
    powershell.exe -ExecutionPolicy Bypass -File "%~dp0ca2.ps1"
)
goto :end

:exe
echo.
echo Launching EXE version...
echo Note: EXE version may have authentication limitations.
echo.
start "" "%~dp0ca_manager.exe"
goto :end

:end
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Error launching the application.
    pause
)
