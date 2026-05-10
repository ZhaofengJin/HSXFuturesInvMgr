# Install dependencies and run main program
$pythonPath = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
$scriptPath = "C:\Users\77188\Desktop\期货库存\期货库存工具.py"

Write-Host "========================================"
Write-Host "  Futures Inventory Data Processing Tool"
Write-Host "========================================"
Write-Host ""

# Check Python
if (-not (Test-Path $pythonPath)) {
    Write-Host "[Error] Python not found" -ForegroundColor Red
    exit 1
}

Write-Host "[1/3] Installing dependency openpyxl..."
& $pythonPath -m pip install openpyxl --quiet --disable-pip-version-check
Write-Host "      Done" -ForegroundColor Green

Write-Host ""
Write-Host "[2/3] Running data processing tool..."
Write-Host ""

# Run main program
& $pythonPath $scriptPath
$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Processing completed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
} else {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "  Processing completed with warnings" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
}
