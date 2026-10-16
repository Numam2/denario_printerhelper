# Genera dist\DenarioPrinter.exe
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
python -m pip install -r requirements.txt --quiet
python -m PyInstaller --noconfirm --clean PrinterHelper.spec
Write-Host "`nListo: $PSScriptRoot\dist\DenarioPrinter.exe" -ForegroundColor Green
