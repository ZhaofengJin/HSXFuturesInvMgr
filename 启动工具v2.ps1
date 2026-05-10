# ============================================================
# 期货库存数据处理工具 v2.0 - 一键运行脚本
# ============================================================

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   Futures Inventory Data Processing Tool v2.0" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 运行Python脚本
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
$psi.Arguments = "C:\Users\77188\Desktop\期货库存\期货库存工具_v2.py"
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
    Write-Host "Output location: results folder" -ForegroundColor Yellow
    Write-Host "  - 期货库存明细_已更新.xlsx (updated data)" -ForegroundColor Yellow
    Write-Host "  - 期货库存明细_原始备份.xlsx (original backup)" -ForegroundColor Yellow
} else {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "   Processing failed with exit code: $($proc.ExitCode)" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
}
