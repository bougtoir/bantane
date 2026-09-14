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
REM   --standalone         : self-contained folder output (no per-launch temp
REM                          extraction, so startup is faster than --onefile)
REM   --lto=yes            : link-time optimization for faster native code
REM   --jobs=N             : parallel C compilation
REM   --windows-console-mode=disable : no console window (GUI app)
REM   --enable-plugin=pyside6 : PySide6 support
REM   --include-package=pulp : include PuLP solver
REM   --include-module=highspy / --include-package-data=highspy : HiGHS solver
REM   --windows-icon-from-ico : app icon (if exists)
REM
REM Nuitka compiles Python to C then to native code,
REM making reverse-engineering extremely difficult.

set NUITKA_OPTS=--standalone --mingw64 --lto=yes --jobs=%NUMBER_OF_PROCESSORS% --windows-console-mode=disable
set NUITKA_OPTS=%NUITKA_OPTS% --enable-plugin=pyside6
set NUITKA_OPTS=%NUITKA_OPTS% --include-package=pulp
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=highspy
set NUITKA_OPTS=%NUITKA_OPTS% --include-package-data=highspy
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=cryptography
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=jpholiday
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=xlrd
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=openpyxl
set NUITKA_OPTS=%NUITKA_OPTS% --nofollow-import-to=pulp.tests

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
    echo [WARNING] Nuitka returned non-zero exit code.
    echo   If the exe was created, this may be a warning only.
    echo.
)

REM --standalone output is the folder dist_nuitka\app.dist with the exe inside.
set DIST_DIR=dist_nuitka\app.dist
if exist "%DIST_DIR%\BantaneShiftOptimizer.exe" (
    echo.
    echo ========================================
    echo   Build Complete - Nuitka standalone
    echo ========================================
    echo.
    echo Output folder: %DIST_DIR%
    echo.
    echo Creating runtime subfolders next to the exe...
    mkdir "%DIST_DIR%\files" 2>nul
    mkdir "%DIST_DIR%\input" 2>nul
    mkdir "%DIST_DIR%\output" 2>nul
    if exist "files\*setting*.xlsx" (
        xcopy /Y files\*setting*.xlsx "%DIST_DIR%\files\"
    ) else (
        for %%f in (*setting*.xlsx) do copy /Y "%%f" "%DIST_DIR%\files\"
    )
) else (
    echo.
    echo [ERROR] Nuitka build failed - exe not found.
    echo.
    echo Troubleshooting:
    echo   - MinGW64 is downloaded automatically by Nuitka on first build
    echo   - Run: python -m nuitka --version
    echo   - Try: pip install --upgrade nuitka
    goto :end
)

REM Build the shippable release folder = the whole standalone folder + tools.
echo Creating release folder...
if exist "release" rmdir /S /Q release
mkdir release
xcopy /E /I /Y "%DIST_DIR%" release
if not exist "release\files" mkdir release\files
if not exist "release\input" mkdir release\input
if not exist "release\output" mkdir release\output

REM Copy setting files
if exist "files\*setting*.xlsx" (
    xcopy /Y files\*setting*.xlsx release\files\
) else (
    for %%f in (*setting*.xlsx) do copy /Y "%%f" release\files\
)

REM Copy license tools
if exist "generate_license.bat" copy /Y generate_license.bat release\
if exist "generate_license.py" copy /Y generate_license.py release\
if exist "license_manager.py" copy /Y license_manager.py release\

echo.
echo ========================================
echo   Release folder ready
echo ========================================
echo.
echo release\ contains the standalone app folder:
echo   release\BantaneShiftOptimizer.exe (plus bundled DLLs)
echo   release\files\*setting*.xlsx
echo   release\generate_license.bat
echo   release\generate_license.py
echo   release\license_manager.py
echo.
echo Distribution checklist:
echo   1. Run generate_license.bat to issue a license
echo   2. Deliver the entire release\ folder to the end user
echo   3. Place the .license file in the same folder as the .exe
echo.

:end
echo.
echo Press any key to close...
pause >nul
endlocal
