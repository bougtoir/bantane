@echo off
setlocal enabledelayedexpansion

echo ========================================
echo   ライセンス発行ツール
echo ========================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    goto :end
)

REM Check if generate_license.py exists
if not exist "generate_license.py" (
    echo [エラー] generate_license.py が見つかりません。
    echo このバッチファイルと同じフォルダに配置してください。
    goto :end
)

REM Run the license generator
python generate_license.py %*

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
