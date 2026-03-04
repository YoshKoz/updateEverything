<#
.SYNOPSIS
    Quick one-liner installer for Update-Everything
.DESCRIPTION
    Downloads updatescript.ps1 to a permanent location and optionally
    registers a daily scheduled task.
.EXAMPLE
    irm https://raw.githubusercontent.com/YoshKoz/windows-update-script/main/install.ps1 | iex
#>

$installDir  = Join-Path $env:USERPROFILE 'scripts'
$scriptPath  = Join-Path $installDir 'updatescript.ps1'
$repoRaw     = 'https://raw.githubusercontent.com/YoshKoz/windows-update-script/main/updatescript.ps1'

# Create install directory
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}

# Download latest version
Write-Host "Downloading updatescript.ps1..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $repoRaw -OutFile $scriptPath -UseBasicParsing
Write-Host "[OK] Saved to: $scriptPath" -ForegroundColor Green

# Add to PATH if not already there
$userPath = [Environment]::GetEnvironmentVariable('PATH', 'User')
if ($userPath -notlike "*$installDir*") {
    [Environment]::SetEnvironmentVariable('PATH', "$userPath;$installDir", 'User')
    Write-Host "[OK] Added $installDir to user PATH (restart your terminal to use)" -ForegroundColor Green
} else {
    Write-Host "[OK] $installDir already in PATH" -ForegroundColor Green
}

Write-Host "`nUsage:" -ForegroundColor Cyan
Write-Host "  updatescript.ps1                      # Run all updates"
Write-Host "  updatescript.ps1 -DryRun              # Preview mode"
Write-Host "  updatescript.ps1 -AutoElevate         # Run as admin"
Write-Host "  updatescript.ps1 -Schedule            # Daily 3 AM task"
Write-Host "  updatescript.ps1 -FastMode            # Skip slow tools"
Write-Host ""
