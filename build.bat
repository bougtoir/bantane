@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   ばんたね病院 シフト最適化アプリ
echo   exe ビルドスクリプト
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"

echo 作業ディレクトリ: %CD%
echo.

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    goto :end
)

python --version
echo.

REM Create virtual environment if not exists
if not exist "venv" (
    echo 仮想環境を作成しています...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [エラー] 仮想環境の作成に失敗しました。
        goto :end
    )
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install dependencies
echo 依存パッケージをインストールしています...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [エラー] パッケージのインストールに失敗しました。
    goto :end
)

REM Install PyInstaller
echo PyInstallerをインストールしています...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo [エラー] PyInstallerのインストールに失敗しました。
    goto :end
)

echo.
echo ========================================
echo   exe をビルドしています...
echo   （数分かかる場合があります）
echo ========================================
echo.

pyinstaller --clean --onefile --windowed ^
    --name "BantaneShiftOptimizer" ^
    --hidden-import=PySide6.QtWidgets ^
    --hidden-import=PySide6.QtCore ^
    --hidden-import=PySide6.QtGui ^
    --collect-all=pulp ^
    --hidden-import=cryptography ^
    --hidden-import=cryptography.fernet ^
    app.py

if %errorlevel% neq 0 (
    echo.
    echo [エラー] ビルドに失敗しました。
    goto :end
)

echo.
echo ========================================
echo   ビルド完了！
echo ========================================
echo.

REM Create subfolders next to the exe in dist
if not exist "dist\files" mkdir dist\files
if not exist "dist\input" mkdir dist\input
if not exist "dist\output" mkdir dist\output

REM Copy setting files to dist\files
if exist "files\setting*.xlsx" (
    xcopy /Y files\setting*.xlsx dist\files\
) else (
    for %%f in (setting*.xlsx) do copy /Y "%%f" dist\files\
)

REM Create release folder structure
echo 配布用フォルダを作成しています...
if not exist "release" mkdir release
if not exist "release\files" mkdir release\files
if not exist "release\input" mkdir release\input
if not exist "release\output" mkdir release\output

REM Copy exe
copy /Y dist\BantaneShiftOptimizer.exe release\

REM Copy setting files
if exist "files\setting*.xlsx" (
    xcopy /Y files\setting*.xlsx release\files\
) else (
    for %%f in (setting*.xlsx) do copy /Y "%%f" release\files\
)

REM Copy license tools
if exist "generate_license.bat" copy /Y generate_license.bat release\
if exist "generate_license.py" copy /Y generate_license.py release\
if exist "license_manager.py" copy /Y license_manager.py release\

echo.
echo ========================================
echo   配布用フォルダ作成完了
echo ========================================
echo.
echo release\ フォルダの内容:
echo   release\BantaneShiftOptimizer.exe
echo   release\files\setting*.xlsx
echo   release\generate_license.bat
echo   release\generate_license.py
echo   release\license_manager.py
echo.
echo 配布手順:
echo   1. generate_license.bat でライセンスを発行
echo   2. release\ フォルダごと配布先に渡す
echo   3. .license ファイルを exe と同じフォルダに配置
echo.

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
