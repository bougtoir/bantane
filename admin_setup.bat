@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   Shift Optimizer Application
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"

echo Current directory: %CD%
echo.

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo Please install Python: https://www.python.org/downloads/
    echo.
    goto :end
)

REM Check Python version
echo Checking Python version...
python --version
echo.

REM Create virtual environment if not exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        echo Please extract the zip file before running.
        goto :end
    )
    echo Virtual environment created.
    echo.
    set NEED_INSTALL=1
) else (
    set NEED_INSTALL=0
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [ERROR] Failed to activate virtual environment.
    goto :end
)

REM Install dependencies if needed
if not exist ".deps_installed" (
    set NEED_INSTALL=1
)

if "%NEED_INSTALL%"=="1" (
    if exist "requirements.txt" (
        echo Installing dependencies...
        echo This may take a few minutes...
        echo.
        pip install -r requirements.txt
        if %errorlevel% neq 0 (
            echo [ERROR] Failed to install packages.
            goto :end
        )
        echo. > .deps_installed
        echo Dependencies installed successfully.
        echo.
    )
)

REM Launch application
echo Launching application...
echo.
python app.py
set APP_EXIT_CODE=%errorlevel%

echo.
echo ----------------------------------------
echo Application exit code: %APP_EXIT_CODE%

if not "%APP_EXIT_CODE%"=="0" (
    echo.
    echo [ERROR] Application terminated abnormally.
    if exist "shift_app.log" (
        echo.
        echo === Log file contents ===
        type shift_app.log
    )
)

:end
echo.
echo Press any key to close...
pause >nul
endlocal
