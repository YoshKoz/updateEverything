#Requires -Version 7
param()
$ErrorActionPreference = 'Stop'

# Replicate script's Invoke-WingetWithTimeout exactly
function Invoke-WingetWithTimeout {
    param([string[]]$Arguments, [int]$TimeoutSec = 300)
    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    try {
        $proc = Start-Process winget -ArgumentList $Arguments -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { taskkill.exe /PID $proc.Id /T /F 2>$null | Out-Null } catch {}
            try { $proc.Kill() } catch {}
            throw "winget timed out"
        }
        $stdout = Get-Content -LiteralPath $stdoutFile -Raw -ErrorAction SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrFile -Raw -ErrorAction SilentlyContinue
        $combined = (($stdout + $stderr) -replace "`0", '').Trim()
        return [pscustomobject]@{ Output = $combined; ExitCode = $proc.ExitCode }
    }
    finally { Remove-Item $stdoutFile, $stderrFile -Force -EA SilentlyContinue }
}

Write-Host "=== STEP 1: source update ==="
$r1 = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('source', 'update', '--disable-interactivity')
Write-Host "ExitCode=$($r1.ExitCode)"
Write-Host $r1.Output

Write-Host "`n=== STEP 2: scan ==="
$r2 = Invoke-WingetWithTimeout -TimeoutSec 300 -Arguments @('upgrade', '--include-unknown', '--include-pinned', '--accept-source-agreements', '--disable-interactivity')
Write-Host "ExitCode=$($r2.ExitCode)"
Write-Host $r2.Output

Write-Host "`n=== STEP 3: upgrade Git.Git (same args as script) ==="
$r3 = Invoke-WingetWithTimeout -TimeoutSec 300 -Arguments @('upgrade', '--id', 'Git.Git', '--source', 'winget', '--include-unknown', '--include-pinned', '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
Write-Host "ExitCode=$($r3.ExitCode)"
Write-Host $r3.Output
