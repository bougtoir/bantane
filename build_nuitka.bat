@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   Shift Optimizer - Nuitka Build
echo   (Obfuscated Native Compilation)
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"

echo Working directory: %CD%
echo.

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found.
    echo https://www.python.org/downloads/
    goto :end
)

python --version
echo.

REM Create virtual environment if not exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        goto :end
    )
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    goto :end
)

REM Install Nuitka and ordered-set (speeds up compilation)
echo Installing Nuitka...
pip install nuitka ordered-set
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install Nuitka.
    goto :end
)

echo.
echo ========================================
echo   Building with Nuitka (this may take
echo   10-20 minutes on first run)...
echo ========================================
echo.

REM Nuitka compilation
REM   --onefile            : single .exe output
REM   --windows-console-mode=disable : no console window (GUI app)
REM   --enable-plugin=pyside6 : PySide6 support
REM   --include-package=pulp : include PuLP solver
REM   --include-data-dir    : include PuLP solver data
REM   --windows-icon-from-ico : app icon (if exists)
REM
REM Nuitka compiles Python to C then to native code,
REM making reverse-engineering extremely difficult.

set NUITKA_OPTS=--onefile --windows-console-mode=disable
set NUITKA_OPTS=%NUITKA_OPTS% --enable-plugin=pyside6
set NUITKA_OPTS=%NUITKA_OPTS% --include-package=pulp
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=cryptography
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=jpholiday
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=xlrd
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=openpyxl

REM Include PuLP solver data files
for /f "delims=" %%P in ('python -c "import pulp; import os; print(os.path.dirname(pulp.__file__))"') do set PULP_DIR=%%P
if defined PULP_DIR (
    set NUITKA_OPTS=%NUITKA_OPTS% --include-data-dir="%PULP_DIR%"=pulp
)

set NUITKA_OPTS=%NUITKA_OPTS% --output-filename=BantaneShiftOptimizer.exe
set NUITKA_OPTS=%NUITKA_OPTS% --output-dir=dist_nuitka
set NUITKA_OPTS=%NUITKA_OPTS% --remove-output

REM Add icon if available
if exist "app.ico" (
    set NUITKA_OPTS=%NUITKA_OPTS% --windows-icon-from-ico=app.ico
)

REM Company / product metadata
set NUITKA_OPTS=%NUITKA_OPTS% --product-name="Bantane Shift Optimizer"
set NUITKA_OPTS=%NUITKA_OPTS% --product-version=1.0.0

echo Running: python -m nuitka %NUITKA_OPTS% app.py
echo.

python -m nuitka %NUITKA_OPTS% app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Nuitka build failed.
    echo.
    echo Troubleshooting:
    echo   - Ensure a C compiler is installed (MinGW64 or MSVC)
    echo   - Run: python -m nuitka --version
    echo   - Try: pip install --upgrade nuitka
    goto :end
)

echo.
echo ========================================
echo   Build Complete (Nuitka)
echo ========================================
echo.
echo Output: dist_nuitka\BantaneShiftOptimizer.exe
echo.

REM Create release folder
if not exist "release" mkdir release
copy /Y dist_nuitka\BantaneShiftOptimizer.exe release\
if not exist "release\files" mkdir release\files

REM Copy setting files
if exist "files\setting*.xlsx" (
    xcopy /Y files\setting*.xlsx release\files\
) else (
    for %%f in (setting*.xlsx) do copy /Y "%%f" release\files\
)

echo.
echo release\ folder is ready for distribution.
echo.
echo Distribution checklist:
echo   1. Generate a license for the target PC:
echo      python generate_license.py
echo   2. Copy release\BantaneShiftOptimizer.exe
echo   3. Copy release\files\ (with setting*.xlsx)
echo   4. Copy the .license file to the same folder as the .exe
echo   5. Deliver the folder to the end user
echo.

:end
echo.
echo Press any key to close...
pause >nul
endlocal
