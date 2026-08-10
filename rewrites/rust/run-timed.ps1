$exe = "C:\Development\updateEverything\rewrites\rust\target\release\updateeverything.exe"
$logFile = "C:\Development\updateEverything\rewrites\rust\run-$(Get-Date -Format yyyyMMdd-HHmmss).log"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
& $exe --jobs 4 *>&1 | Tee-Object -FilePath $logFile
$sw.Stop()
Write-Host "ELAPSED: $($sw.Elapsed)" -ForegroundColor Cyan
