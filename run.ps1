# PowerShell Launcher for MD Converter Studio
Write-Host "Launching MD Converter Studio..." -ForegroundColor Cyan
python "$PSScriptRoot\app.py"
if ($LASTEXITCODE -ne 0) {
    Write-Host "`n[ERROR] Application exited with code $LASTEXITCODE." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
}
