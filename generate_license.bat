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

REM Check Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/downloads/ からインストールしてください。
    goto :end
)

REM Show Python version
python --version

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

REM Check if .license was created in any dist*/files/ or release/files/
set LICENSE_FOUND=0
for /d %%D in (dist* release) do (
    if exist "%%D\files\.license" (
        echo [成功] .license: %CD%\%%D\files\.license
        set LICENSE_FOUND=1
    )
)
if "%LICENSE_FOUND%"=="0" (
    if exist "files\.license" (
        echo [成功] .license: %CD%\files\.license
    ) else if exist ".license" (
        echo [成功] .license: %CD%\.license
    ) else (
        echo [確認] .license ファイルが見つかりません
        echo         generate_license.py の出力先を確認してください
    )
)

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
