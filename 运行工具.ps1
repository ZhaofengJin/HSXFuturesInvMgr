# Install dependencies and run main program
$pythonPath = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
$pipPath = "C:\Users\77188\AppData\Local\Programs\Python\Python312\Scripts\pip.exe"
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
if ($LASTEXITCODE -eq 0) {
    Write-Host "      Dependency installed" -ForegroundColor Green
} else {
    Write-Host "      Dependency installation completed (warnings may appear)"
}

Write-Host ""
Write-Host "[2/3] Running data processing tool..."
Write-Host ""

# Set working directory
Set-Location "C:\Users\77188\Desktop\期货库存"

# Run main program
& $pythonPath $scriptPath

Write-Host ""
Write-Host "[3/3] Processing complete!"

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "  Processing completed successfully!"
    Write-Host "========================================"
    Write-Host ""
    Write-Host "Output file: 期货库存明细_处理结果.xlsx"
    Write-Host ""
} else {
    Write-Host "[Warning] Processing may have issues, please check output" -ForegroundColor Yellow
}

Read-Host "Press Enter to exit"
