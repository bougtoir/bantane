@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   ばんたね病院 シフト最適化アプリ
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    echo.
    goto :end
)

REM Create virtual environment if not exists
if not exist "venv" (
    echo 仮想環境を作成しています...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [エラー] 仮想環境の作成に失敗しました。
        goto :end
    )
    echo.
    set NEED_INSTALL=1
) else (
    set NEED_INSTALL=0
)

REM Activate virtual environment
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo [エラー] 仮想環境の有効化に失敗しました。
    goto :end
)

REM Install dependencies if needed
if not exist ".deps_installed" (
    set NEED_INSTALL=1
)

if "%NEED_INSTALL%"=="1" (
    if exist "requirements.txt" (
        echo 依存パッケージをインストールしています...
        pip install -r requirements.txt
        if %errorlevel% neq 0 (
            echo [エラー] パッケージのインストールに失敗しました。
            goto :end
        )
        echo. > .deps_installed
        echo インストール完了。
        echo.
    )
)

REM Launch application
echo アプリを起動しています...
echo.
python app.py

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
