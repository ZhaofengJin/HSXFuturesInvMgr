# ============================================================
# 期货库存数据处理工具 - 一键运行脚本
# ============================================================

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   Futures Inventory Data Processing Tool" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Run Python script with output capture
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
$psi.Arguments = "C:\Users\77188\Desktop\期货库存\期货库存工具_final.py"
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false

$proc = [System.Diagnostics.Process]::Start($psi)
$stdout = $proc.StandardOutput.ReadToEnd()
$stderr = $proc.StandardError.ReadToEnd()
$proc.WaitForExit()

Write-Host $stdout

if ($stderr) {
    Write-Host "Errors:" -ForegroundColor Red
    Write-Host $stderr -ForegroundColor Red
}

if ($proc.ExitCode -eq 0) {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "   Processing completed successfully!" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output file: FuturesInventory_Result.xlsx" -ForegroundColor Yellow
    Write-Host "Location: C:\Users\77188\Desktop\期货库存" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "   Processing failed with exit code: $($proc.ExitCode)" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
}
