@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title FortiGate Policy to Excel Exporter v1.0
cd /d "%~dp0"

echo ==============================================================================
echo   FortiGate Policy to Excel Exporter - Startup Launcher
echo ==============================================================================
echo [*] Detecting Python environment on your system...

set "PY_BIN="

REM 1. Check python in system PATH
where python >nul 2>&1
if %errorlevel% equ 0 (
    for /f "delims=" %%p in ('where python 2^>nul') do (
        if not defined PY_BIN set "PY_BIN=%%p"
    )
)

REM 2. Check Python Launcher for Windows (py.exe)
if not defined PY_BIN (
    where py >nul 2>&1
    if !errorlevel! equ 0 set "PY_BIN=py"
)

REM 3. Check Windows Registry
if not defined PY_BIN (
    for /f "tokens=2*" %%a in ('reg query "HKCU\Software\Python\PythonCore" /s /v "ExecutablePath" 2^>nul ^| findstr /i "ExecutablePath"') do (
        if exist "%%b" if not defined PY_BIN set "PY_BIN=%%b"
    )
)
if not defined PY_BIN (
    for /f "tokens=2*" %%a in ('reg query "HKLM\Software\Python\PythonCore" /s /v "ExecutablePath" 2^>nul ^| findstr /i "ExecutablePath"') do (
        if exist "%%b" if not defined PY_BIN set "PY_BIN=%%b"
    )
)

REM 4. Check Common LocalAppData and ProgramFiles paths
if not defined PY_BIN (
    for %%v in (313 312 311 310 39 38) do (
        if not defined PY_BIN (
            if exist "%LocalAppData%\Programs\Python\Python%%v\python.exe" set "PY_BIN=%LocalAppData%\Programs\Python\Python%%v\python.exe"
        )
        if not defined PY_BIN (
            if exist "%ProgramFiles%\Python%%v\python.exe" set "PY_BIN=%ProgramFiles%\Python%%v\python.exe"
        )
    )
)

REM 5. Check Conda / Anaconda
if not defined PY_BIN (
    if exist "%UserProfile%\anaconda3\python.exe" set "PY_BIN=%UserProfile%\anaconda3\python.exe"
    if exist "%UserProfile%\miniconda3\python.exe" set "PY_BIN=%UserProfile%\miniconda3\python.exe"
)

if not defined PY_BIN goto :py_not_found

echo [*] Python detected: %PY_BIN%

REM 6. Check required dependency (openpyxl)
"%PY_BIN%" -c "import openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [*] Installing required module: openpyxl...
    "%PY_BIN%" -m pip install openpyxl
)

REM 7. Run Application
echo [*] Launching FortiGate Policy to Excel Exporter...
echo ==============================================================================
"%PY_BIN%" fortigate_policy_to_excel.py %*
goto :finish

:py_not_found
echo.
echo [ERROR] Python was not found on this computer.
echo.
echo Please install Python [version 3.9 or higher]:
echo   1. Download: https://www.python.org/downloads/
echo   2. Make sure to check "Add python.exe to PATH" during installation.
echo.
pause
exit /b 1

:finish
if %errorlevel% neq 0 (
    echo.
    echo [NOTICE] Application terminated with exit code %errorlevel%.
    pause
)
