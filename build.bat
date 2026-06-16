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
echo 生成されたexeファイル:
echo   dist\BantaneShiftOptimizer.exe
echo.
echo 配布手順:
echo   1. 新しいフォルダを作成
echo   2. dist\BantaneShiftOptimizer.exe をコピー
echo   3. files\ フォルダをコピー（setting*.xlsx を含む）
echo   4. フォルダごと配布先に渡す
echo.

REM Create distribution folder
if not exist "release" mkdir release
copy /Y dist\BantaneShiftOptimizer.exe release\
if not exist "release\files" mkdir release\files
REM settingファイルをコピー（リポジトリ直下またはfiles/から探す）
if exist "files\setting*.xlsx" (
    xcopy /Y files\setting*.xlsx release\files\
) else (
    for %%f in (setting*.xlsx) do copy /Y "%%f" release\files\
)

echo.
echo release\ フォルダに配布用ファイルをまとめました。
echo.

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
