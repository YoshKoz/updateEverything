#Requires -Version 7
$ErrorActionPreference = 'Continue'

function Invoke-WingetWithTimeout {
    param([string[]]$Arguments, [int]$TimeoutSec = 120)
    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    try {
        $proc = Start-Process winget -ArgumentList $Arguments -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $proc.WaitForExit($TimeoutSec * 1000) | Out-Null
        $stdout = Get-Content -LiteralPath $stdoutFile -Raw -EA SilentlyContinue
        $stderr = Get-Content -LiteralPath $stderrFile -Raw -EA SilentlyContinue
        [pscustomobject]@{ Output = (($stdout + $stderr) -replace "`0", '').Trim(); ExitCode = $proc.ExitCode }
    }
    finally { Remove-Item $stdoutFile, $stderrFile -Force -EA SilentlyContinue }
}

Write-Host "=== Run Git.Git Pre-hook ==="
Stop-Process -Name 'git', 'git-bash', 'bash', 'sh', 'mintty', 'git-credential-manager', 'wintoast', 'gitk', 'tig', 'gitui' -Force -EA SilentlyContinue
Start-Sleep 1

Write-Host "`n=== Upgrade Git.Git ==="
$r = Invoke-WingetWithTimeout -Arguments @('upgrade', '--id', 'Git.Git', '--source', 'winget', '--include-unknown', '--include-pinned', '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
Write-Host "ExitCode=$($r.ExitCode)"
Write-Host $r.Output
