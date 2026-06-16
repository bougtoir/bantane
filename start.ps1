# ばんたね病院 シフト最適化アプリ 起動スクリプト
Write-Host "========================================"
Write-Host "  ばんたね病院 シフト最適化アプリ"
Write-Host "========================================"
Write-Host ""

# Change to script directory
Set-Location -Path $PSScriptRoot

# Check Python
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) {
    Write-Host "[エラー] Pythonが見つかりません。" -ForegroundColor Red
    Write-Host "https://www.python.org/downloads/ からインストールしてください。"
    Read-Host "終了するにはEnterを押してください"
    exit 1
}

# Create virtual environment if not exists
$needInstall = $false
if (-not (Test-Path "venv")) {
    Write-Host "仮想環境を作成しています..."
    python -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[エラー] 仮想環境の作成に失敗しました。" -ForegroundColor Red
        Read-Host "終了するにはEnterを押してください"
        exit 1
    }
    $needInstall = $true
}

# Activate virtual environment
& ".\venv\Scripts\Activate.ps1"

# Install dependencies if needed
if (-not (Test-Path ".deps_installed")) {
    $needInstall = $true
}

if ($needInstall -and (Test-Path "requirements.txt")) {
    Write-Host "依存パッケージをインストールしています..."
    pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[エラー] パッケージのインストールに失敗しました。" -ForegroundColor Red
        Read-Host "終了するにはEnterを押してください"
        exit 1
    }
    New-Item -Path ".deps_installed" -ItemType File -Force | Out-Null
    Write-Host "インストール完了。"
    Write-Host ""
}

# Launch application
Write-Host "アプリを起動しています..."
Write-Host ""
python app.py

Write-Host ""
Read-Host "終了するにはEnterを押してください"
