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

REM ---- Pick a Python that has pip (prefer the py launcher; skip pip-less venvs on PATH) ----
set PY=
py -3 -m pip --version >nul 2>&1
if %errorlevel% equ 0 set PY=py -3
if defined PY goto :python_ok

python -m pip --version >nul 2>&1
if %errorlevel% equ 0 set PY=python
if defined PY goto :python_ok

where python >nul 2>&1
if %errorlevel% neq 0 goto :no_python

REM python exists but has no pip: try ensurepip
python -m ensurepip --default-pip >nul 2>&1
python -m pip --version >nul 2>&1
if %errorlevel% equ 0 set PY=python
if defined PY goto :python_ok
goto :no_pip

:python_ok
%PY% --version
echo 使用する Python: %PY%

REM Check if generate_license.py exists
if not exist "generate_license.py" goto :no_script

REM ---- Ensure cryptography is available ----
%PY% -c "import cryptography" >nul 2>&1
if %errorlevel% equ 0 goto :run
echo cryptography パッケージが見つかりません。インストールします...
%PY% -m pip install cryptography
if %errorlevel% neq 0 goto :pip_failed
echo.

:run
echo.
echo generate_license.py を実行します...
echo.
%PY% generate_license.py %*
set GEN_RC=%errorlevel%

echo.
echo 終了コード: %GEN_RC%
if not "%GEN_RC%"=="0" goto :gen_failed

REM ---- Report where .license ended up ----
set LICENSE_FOUND=0
for /d %%D in (dist* release) do (
    if exist "%%D\files\.license" (
        echo [成功] .license: %CD%\%%D\files\.license
        set LICENSE_FOUND=1
    )
)
if "%LICENSE_FOUND%"=="1" goto :end
if exist "..\files\.license" (
    echo [成功] .license: %CD%\..\files\.license
    goto :end
)
if exist "files\.license" (
    echo [成功] .license: %CD%\files\.license
    goto :end
)
if exist ".license" (
    echo [成功] .license: %CD%\.license
    goto :end
)
echo [確認] .license ファイルが見つかりません
echo         generate_license.py の出力先を確認してください
goto :end

:no_python
echo [エラー] Pythonが見つかりません。
echo https://www.python.org/downloads/ からインストールしてください。
goto :end

:no_pip
echo [エラー] PATH 上の python に pip が入っていません（別アプリ用の venv の可能性）。
echo python.org の Python をインストールするか、py ランチャーを有効にしてください。
goto :end

:no_script
echo [エラー] generate_license.py が見つかりません。
echo このバッチファイルと同じフォルダに配置してください。
echo 現在のフォルダ: %CD%
dir /b *.py
goto :end

:pip_failed
echo [エラー] cryptography のインストールに失敗しました。
echo 手動で  %PY% -m pip install cryptography  を実行してください。
goto :end

:gen_failed
echo [エラー] ライセンス発行に失敗しました。上記のエラーを確認してください。
goto :end

:end
echo.
echo 終了するにはキーを押してください...
pause >nul
endlocal
