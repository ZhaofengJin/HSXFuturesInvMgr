$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "C:\Users\77188\AppData\Local\Programs\Python\Python312\python.exe"
$psi.Arguments = "C:\Users\77188\Desktop\期货库存\期货库存工具_final.py"
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
$psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

$proc = [System.Diagnostics.Process]::Start($psi)
$stdout = $proc.StandardOutput.ReadToEnd()
$stderr = $proc.StandardError.ReadToEnd()
$proc.WaitForExit()

Write-Host "STDOUT:"
Write-Host $stdout
if ($stderr) {
    Write-Host "STDERR:"
    Write-Host $stderr
}
Write-Host "EXIT CODE: $($proc.ExitCode)"
