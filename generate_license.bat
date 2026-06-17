@echo off
chcp 65001 >nul 2>&1
setlocal

echo ========================================
echo   ライセンス発行ツール
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"
echo 実行ディレクトリ: %CD%

REM Activate venv if exists
if exist "venv\Scripts\activate.bat" (
    echo venv を有効化しています...
    call venv\Scripts\activate.bat
) else if exist "..\venv\Scripts\activate.bat" (
    echo venv を有効化しています（親フォルダ）...
    call ..\venv\Scripts\activate.bat
)

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    goto :end
)

REM Show Python version
python --version

REM Check cryptography module
python -c "from cryptography.fernet import Fernet" >nul 2>&1
if %errorlevel% neq 0 (
    echo [情報] cryptography モジュールが見つかりません。インストールします...
    echo.
    if exist "venv\Scripts\activate.bat" (
        pip install cryptography
    ) else (
        echo venv が見つかりません。venv を作成してインストールします...
        python -m venv venv
        call venv\Scripts\activate.bat
        pip install cryptography
    )
    echo.
    python -c "from cryptography.fernet import Fernet" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [エラー] cryptography のインストールに失敗しました。
        goto :end
    )
    echo [成功] cryptography をインストールしました。
    echo.
)

REM Check if generate_license.py exists
if not exist "generate_license.py" (
    echo [エラー] generate_license.py が見つかりません。
    echo このバッチファイルと同じフォルダに配置してください。
    echo 現在のフォルダ: %CD%
    dir /b *.py
    goto :end
)

echo.
echo generate_license.py を実行します...
echo.

REM Run the license generator
python generate_license.py %*

echo.
echo 終了コード: %errorlevel%

REM Check if .license was created
if exist ".license" (
    echo [成功] .license ファイルが生成されました: %CD%\.license
) else (
    echo [確認] .license ファイルが見つかりません。
    echo         出力先を確認してください。
)

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
