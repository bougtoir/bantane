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

REM Install Nuitka with onefile support and ordered-set (speeds up compilation)
echo Installing Nuitka...
pip install "nuitka[onefile]" ordered-set
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

set NUITKA_OPTS=--onefile --mingw64 --windows-console-mode=disable
set NUITKA_OPTS=%NUITKA_OPTS% --enable-plugin=pyside6
set NUITKA_OPTS=%NUITKA_OPTS% --include-package=pulp
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=cryptography
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=jpholiday
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=xlrd
set NUITKA_OPTS=%NUITKA_OPTS% --include-module=openpyxl
set NUITKA_OPTS=%NUITKA_OPTS% --nofollow-import-to=pulp.tests

REM Include PuLP solver data files (cbc.exe etc.)
set NUITKA_OPTS=%NUITKA_OPTS% --include-package-data=pulp

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

REM Always create subfolders next to the exe in dist_nuitka
if exist "dist_nuitka\BantaneShiftOptimizer.exe" (
    echo.
    echo ========================================
    echo   Build Complete - Nuitka
    echo ========================================
    echo.
    echo Output: dist_nuitka\BantaneShiftOptimizer.exe
    echo.
    echo Creating dist_nuitka subfolders...
    mkdir dist_nuitka\files 2>nul
    mkdir dist_nuitka\input 2>nul
    mkdir dist_nuitka\output 2>nul
    if exist "files\*setting*.xlsx" (
        xcopy /Y files\*setting*.xlsx dist_nuitka\files\
    ) else (
        for %%f in (*setting*.xlsx) do copy /Y "%%f" dist_nuitka\files\
    )
) else (
    echo.
    echo [ERROR] Nuitka build failed - exe not found.
    echo.
    echo Troubleshooting:
    echo   - Ensure a C compiler is installed (MinGW64 or MSVC)
    echo   - Run: python -m nuitka --version
    echo   - Try: pip install --upgrade nuitka
    goto :end
)

REM Copy CBC solver binary next to exe
echo Copying CBC solver...
copy /Y venv\Lib\site-packages\pulp\solverdir\cbc\win\i64\cbc.exe dist_nuitka\
if exist "dist_nuitka\cbc.exe" (
    echo CBC solver: OK
) else (
    echo [WARNING] cbc.exe could not be copied
)

REM Create release folder structure
echo Creating release folder...
if not exist "release" mkdir release
if not exist "release\files" mkdir release\files
if not exist "release\input" mkdir release\input
if not exist "release\output" mkdir release\output

REM Copy exe and solver
copy /Y dist_nuitka\BantaneShiftOptimizer.exe release\
if exist "dist_nuitka\cbc.exe" copy /Y dist_nuitka\cbc.exe release\

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
echo release\ contents:
echo   release\BantaneShiftOptimizer.exe
echo   release\files\*setting*.xlsx
echo   release\generate_license.bat
echo   release\generate_license.py
echo   release\license_manager.py
echo.
echo Distribution checklist:
echo   1. Run generate_license.bat to issue a license
echo   2. Deliver the release\ folder to the end user
echo   3. Place the .license file in the same folder as the .exe
echo.

:end
echo.
echo Press any key to close...
pause >nul
endlocal
