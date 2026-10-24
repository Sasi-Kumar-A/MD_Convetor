Set-Location -Path $PSScriptRoot

if (-not (Test-Path "venv")) {
    Write-Host "Creating virtual environment..." -ForegroundColor Cyan
    python -m venv venv
}

& ".\venv\Scripts\Activate.ps1"

Write-Host "Checking and installing dependencies..." -ForegroundColor Cyan
python.exe -m pip install --upgrade pip
pip install -r requirements.txt

Write-Host "Starting MD Converter Studio..." -ForegroundColor Green
python app.py