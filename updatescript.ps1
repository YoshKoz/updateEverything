<#
.SYNOPSIS
    System-wide update script for Windows
.DESCRIPTION
    Updates all package managers, system components, and development tools.
    Supports parallel execution of independent package managers for faster runs.
.VERSION
    2.2.0
.NOTES
    Run as Administrator for full functionality (some features require elevation).
    Requires PowerShell 7+ for -Parallel support.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -Parallel
    .\updatescript.ps1 -DryRun
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
    .\updatescript.ps1 -SkipCleanup -SkipWindowsUpdate
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,                  # Skip slower managers
    [switch]$NoElevate,                 # Compatibility switch (prevents auto-elevation)
    [switch]$AutoElevate,               # Opt-in: relaunch elevated (opens a new window due UAC)
    [switch]$NoPause,                   # Skip "Press Enter to close" prompt (for VS Code / CI)
    [switch]$Parallel,                  # Enable parallel updates where supported (e.g. PSResourceGet, PS7+)
    [switch]$DryRun,                    # Show what would be updated without doing it
    [switch]$SkipWSL,
    [switch]$SkipWSLDistros,            # Skip updating WSL distro packages (apt/pacman etc.)
    [switch]$SkipDefender,
    [switch]$SkipStoreApps,
    [switch]$SkipUVTools,
    [switch]$SkipVSCodeExtensions,
    [switch]$SkipPoetry,
    [switch]$SkipComposer,
    [switch]$SkipRuby,
    [switch]$SkipPowerShellModules,
    [switch]$SkipCleanup,               # Skip the system cleanup section
    [int]$WingetTimeoutSec = 300,       # Per-call timeout for winget (seconds, default 5 min)
    [switch]$Schedule,                  # Register a daily scheduled task to run this script
    [string]$ScheduleTime = "03:00",    # Time for the scheduled task (default: 3 AM)
    [string]$LogPath
)

# ── Console encoding / external tool output ─────────────────────────────────────
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding  = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
} catch {
    # Best effort only; continue if host does not allow setting console encodings.
}

# ── Self-elevate ────────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and $AutoElevate -and -not $NoElevate) {
    $params = @()
    if ($SkipWindowsUpdate)     { $params += '-SkipWindowsUpdate' }
    if ($SkipReboot)            { $params += '-SkipReboot' }
    if ($SkipDestructive)       { $params += '-SkipDestructive' }
    if ($FastMode)              { $params += '-FastMode' }
    if ($Parallel)              { $params += '-Parallel' }
    if ($DryRun)                { $params += '-DryRun' }
    if ($SkipWSL)               { $params += '-SkipWSL' }
    if ($SkipWSLDistros)        { $params += '-SkipWSLDistros' }
    if ($SkipDefender)          { $params += '-SkipDefender' }
    if ($SkipStoreApps)         { $params += '-SkipStoreApps' }
    if ($SkipUVTools)           { $params += '-SkipUVTools' }
    if ($SkipVSCodeExtensions)  { $params += '-SkipVSCodeExtensions' }
    if ($SkipPoetry)            { $params += '-SkipPoetry' }
    if ($SkipComposer)          { $params += '-SkipComposer' }
    if ($SkipRuby)              { $params += '-SkipRuby' }
    if ($SkipPowerShellModules) { $params += '-SkipPowerShellModules' }
    if ($SkipCleanup)               { $params += '-SkipCleanup' }
    if ($NoPause)                   { $params += '-NoPause' }
    if ($WingetTimeoutSec -ne 300)  { $params += @('-WingetTimeoutSec', $WingetTimeoutSec) }
    if ($AutoElevate)               { $params += '-AutoElevate' }
    if ($LogPath)                   { $params += @('-LogPath', $LogPath) }

    $pwshArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + $params
    try {
        $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue

        if ($pwshCmd) {
            Start-Process -FilePath $pwshCmd.Source -Verb RunAs -ArgumentList $pwshArgs -Wait
        } else {
            Write-Host "WARNING: pwsh.exe not found. Falling back to Windows PowerShell." -ForegroundColor Yellow
            Start-Process powershell -Verb RunAs -ArgumentList $pwshArgs -Wait
        }
        exit
    } catch {
        Write-Host "WARNING: Could not elevate. Running without Administrator privileges.`n" -ForegroundColor Yellow
    }
} elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host "INFO: Running in the current terminal without elevation (admin-only tasks may be skipped)." -ForegroundColor DarkYellow
    Write-Host "      To run elevated, start the terminal as Administrator or use -AutoElevate (opens a new window)." -ForegroundColor DarkYellow
}

# ── Scheduled task registration ─────────────────────────────────────────────────
if ($Schedule) {
    $taskName = "DailySystemUpdate"
    $action   = New-ScheduledTaskAction -Execute 'pwsh.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipReboot"
    $trigger  = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
        -Settings $settings -RunLevel Highest -Force | Out-Null
    Write-Host "[OK] Scheduled task '$taskName' registered to run daily at $ScheduleTime." -ForegroundColor Green
    exit
}

$ErrorActionPreference = "Continue"
$startTime    = Get-Date
$commandCache = @{}

# ── Helper functions ─────────────────────────────────────────────────────────────

function Write-Section {
    param([string]$Title, [double]$ElapsedSeconds = 0)
    $elapsed = if ($ElapsedSeconds -gt 0) { "  ({0:F1}s)" -f $ElapsedSeconds } else { "" }
    Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
    Write-Host " $Title$elapsed" -ForegroundColor Cyan
    Write-Host "$('=' * 54)" -ForegroundColor DarkGray
}

function Test-Command([string]$Command) {
    if ($commandCache.ContainsKey($Command)) { return $commandCache[$Command] }
    $result = [bool](Get-Command $Command -ErrorAction SilentlyContinue)
    $commandCache[$Command] = $result
    return $result
}

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Success', 'Warning', 'Error', 'Info')]
        [string]$Type = 'Info'
    )
    $colors  = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]';  Warning = '[!]';    Error = '[X]'; Info = '[*]'  }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Write-FilteredOutput {
    param(
        [AllowNull()][string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    # Remove common ANSI escape sequences and normalize carriage-return progress updates.
    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    $normalized = $normalized -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.TrimEnd()
        if (-not $line) { continue }

        $compact = $line.Trim()

        # Drop spinner-only frames emitted by winget and installers.
        if ($compact -match '^[\\/\|\-]+$') { continue }

        # Drop garbled/Unicode progress bar rows (block chars, regardless of whether a size suffix follows).
        if ($compact -match '&#x393;&#xFB;|[&#x2588;&#x2593;&#x2592;&#x2591;&#x258F;&#x258E;&#x258D;&#x258C;&#x258B;&#x258A;&#x2589;&#x25A0;&#x25A1;&#x25AA;&#x25AB;]') { continue }

        # Drop noisy winget "cannot be determined" info lines.
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }

        Write-Host $line -ForegroundColor $Color
    }
}

function Get-WingetFailedPackageIds {
    param([AllowNull()][string]$Text)

    $failedIds = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $currentId = $null
    $normalized = $Text -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }

        if ($line -match '^\(\d+/\d+\)\s+Found\s+.+\s+\[(?<id>[^\]]+)\]\s+Version\b') {
            $currentId = $Matches['id']
            continue
        }

        if (-not $currentId) { continue }

        # Terminal states for the current package
        if ($line -match '^(Successfully installed|No applicable upgrade found|Package already installed)') {
            $currentId = $null
            continue
        }

        # Common winget failure markers (including portable-package edge cases)
        if ($line -match '(?i)installer failed with exit code' -or
            $line -match '(?i)unable to remove portable package' -or
            $line -match '(?i)upgrade failed') {
            if (-not $failedIds.Contains($currentId)) {
                $failedIds.Add($currentId)
            }
            $currentId = $null
            continue
        }
    }

    return @($failedIds)
}

function Invoke-WingetWithTimeout {
    <#
    .SYNOPSIS
        Runs winget with a hard timeout, writing output to temp files.
        Avoids both the pipeline-buffering hang and .NET event-handler thread crashes.
    #>
    param(
        [string[]]$Arguments,
        [int]$TimeoutSec = 300
    )

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()

    try {
        # Start-Process with file redirection: no in-memory buffers, no threading issues.
        $proc = Start-Process -FilePath 'winget' `
            -ArgumentList $Arguments `
            -RedirectStandardOutput $stdoutFile `
            -RedirectStandardError  $stderrFile `
            -NoNewWindow -PassThru

        $exited = $proc.WaitForExit($TimeoutSec * 1000)

        if (-not $exited) {
            try { $proc.Kill() } catch { }
            throw "winget timed out after ${TimeoutSec}s &#x2014; a hanging installer may require manual intervention"
        }

        $exitCode = $proc.ExitCode
        $stdout   = Get-Content -Raw -Path $stdoutFile -ErrorAction SilentlyContinue
        $stderr   = Get-Content -Raw -Path $stderrFile -ErrorAction SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()

        return [pscustomobject]@{ Output = $combined; ExitCode = $exitCode }
    } finally {
        Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
    }
}

$updateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[string]]::new()
}

# ── Logging ──────────────────────────────────────────────────────────────────────
$transcriptStarted = $false
if ($LogPath) {
    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if ($logDir) { New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null }
        Start-Transcript -Path $LogPath -Append -Force | Out-Null
        $transcriptStarted = $true
    } catch {
        Write-Host "WARNING: Could not start transcript at $LogPath. $_" -ForegroundColor Yellow
    }
}

# ── Startup banner ───────────────────────────────────────────────────────────────
Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
Write-Host "  Update-Everything v2.2.0  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Cyan
Write-Host "  $(if ($isAdmin) { 'Running as Administrator' } else { 'Running as standard user (some tasks skipped)' })" -ForegroundColor $(if ($isAdmin) { 'Green' } else { 'Yellow' })
Write-Host "$('=' * 54)" -ForegroundColor DarkGray

# ── Dry-run banner ───────────────────────────────────────────────────────────────
if ($DryRun) {
    Write-Host "`n[DRY-RUN MODE] No changes will be made.`n" -ForegroundColor Magenta
}

# ── Core update wrapper ──────────────────────────────────────────────────────────
function Invoke-Update {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$Title,
        [Parameter(Mandatory)][scriptblock]$Action,
        [string]$RequiresCommand,
        [string[]]$RequiresAnyCommand,
        [switch]$Disabled,
        [switch]$RequiresAdmin,
        [switch]$SlowOperation,
        [switch]$NoSection
    )
    if (-not $Title) { $Title = $Name }

    if ($Disabled)                                                                          { $updateResults.Skipped.Add($Name); return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand))                         { $updateResults.Skipped.Add($Name); return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { $updateResults.Skipped.Add($Name); return }
    if ($RequiresAdmin -and -not $isAdmin)                                                  { $updateResults.Skipped.Add($Name); return }
    if ($SlowOperation -and $FastMode)                                                      { $updateResults.Skipped.Add($Name); return }

    $sectionStart = Get-Date
    if (-not $NoSection) { Write-Section $Title }

    if ($DryRun) {
        Write-Status "[DRY-RUN] Would update: $Name" -Type Info
        $updateResults.Success.Add($Name)
        return
    }

    try {
        & $Action
        $elapsed = ((Get-Date) - $sectionStart).TotalSeconds
        Write-Status ("$Name updated ({0:F1}s)" -f $elapsed) -Type Success
        $updateResults.Success.Add($Name)
    } catch {
        Write-Status "$Name failed: $_" -Type Error
        $updateResults.Failed.Add($Name)
    }
}

# ════════════════════════════════════════════════════════════
#  PACKAGE MANAGERS
# ════════════════════════════════════════════════════════════

# ── Scoop ────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Scoop' -RequiresCommand 'scoop' -Action {
    scoop update 2>&1 | Out-Null
    $out = (scoop update '*' 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out
    scoop cleanup '*' 2>&1 | Out-Null
    scoop cache rm '*' 2>&1 | Out-Null
}

# ── Winget ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Winget' -RequiresCommand 'winget' -Action {
    $bulkHadFailures = $false
    $retryIds = [System.Collections.Generic.List[string]]::new()

    # Standard winget source
    Write-Host "  Upgrading all (winget source, timeout: ${WingetTimeoutSec}s)..." -ForegroundColor Gray
    try {
        $result = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
            'upgrade', '--all', '--source', 'winget',
            '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
        )
        Write-FilteredOutput $result.Output
        foreach ($id in (Get-WingetFailedPackageIds $result.Output)) {
            if (-not $retryIds.Contains($id)) { $retryIds.Add($id) }
        }
        if ($result.ExitCode -ne 0 -or $result.Output -match 'Installer failed with exit code') { $bulkHadFailures = $true }
    } catch {
        Write-Status "Winget (winget source): $_" -Type Warning
        $bulkHadFailures = $true
    }

    # MS Store source (skipped when -SkipStoreApps)
    if (-not $SkipStoreApps) {
        Write-Host "  Checking MS Store source for upgrades..." -ForegroundColor Gray
        try {
            $storeCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
                'upgrade', '--source', 'msstore',
                '--accept-source-agreements', '--disable-interactivity'
            )
            if ($storeCheck.Output -match 'upgrades available') {
                Write-Host "  Upgrading MS Store apps (timeout: ${WingetTimeoutSec}s)..." -ForegroundColor Gray
                $storeResult = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                    'upgrade', '--all', '--source', 'msstore',
                    '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
                )
                Write-FilteredOutput $storeResult.Output
                foreach ($id in (Get-WingetFailedPackageIds $storeResult.Output)) {
                    if (-not $retryIds.Contains($id)) { $retryIds.Add($id) }
                }
                if ($storeResult.ExitCode -ne 0 -or $storeResult.Output -match 'Installer failed with exit code') { $bulkHadFailures = $true }
            }
        } catch {
            Write-Status "Winget (msstore source): $_" -Type Warning
        }
    }

    # Retry only packages that actually failed during bulk execution.
    if ($retryIds.Count -gt 0) {
        Write-Host "  Retrying $($retryIds.Count) failed package(s) individually..." -ForegroundColor Gray
        foreach ($pkgId in $retryIds) {
            Write-Host "    Updating $pkgId..." -ForegroundColor Gray
            try {
                Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                    'upgrade', '--id', $pkgId,
                    '--accept-source-agreements', '--accept-package-agreements',
                    '--disable-interactivity', '--force'
                ) | Out-Null
            } catch {
                Write-Status "  Retry of $pkgId timed out or failed: $_" -Type Warning
            }
        }
    } elseif ($bulkHadFailures) {
        Write-Status 'Winget reported a bulk failure, but no specific failed package IDs were detected for retry' -Type Warning
    }
}

# ── Chocolatey ───────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Chocolatey' -RequiresCommand 'choco' -RequiresAdmin -Action {
    $out = (choco upgrade all -y 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out
}

# ════════════════════════════════════════════════════════════
#  WINDOWS COMPONENTS
# ════════════════════════════════════════════════════════════

# ── Windows Update ───────────────────────────────────────────────────────────────
if (-not $SkipWindowsUpdate) {
    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Invoke-Update -Name 'WindowsUpdate' -Title 'Windows Update' -RequiresAdmin -Action {
            Import-Module PSWindowsUpdate

            # 0x800704c7 (ERROR_CANCELLED) is usually caused by the WU agent aborting
            # due to a pending reboot or -AutoReboot conflicting with the install.
            # Fix: use -IgnoreReboot during install, then handle reboot separately.
            $maxRetries = 2
            $attempt    = 0
            $succeeded  = $false

            while ($attempt -lt $maxRetries -and -not $succeeded) {
                $attempt++
                try {
                    # IgnoreReboot prevents the WU agent from self-cancelling mid-install.
                    # RecurseCycle handles chained dependencies that require multiple passes.
                    $wuParams = @{
                        Install       = $true
                        AcceptAll     = $true
                        NotCategory   = 'Drivers'
                        IgnoreReboot  = $true
                        RecurseCycle  = 3
                        Verbose       = $false
                        Confirm       = $false
                    }
                    $results = Get-WindowsUpdate @wuParams

                    if ($results) {
                        Write-Host "  Installed $($results.Count) update(s)." -ForegroundColor Gray
                    } else {
                        Write-Host "  No updates available." -ForegroundColor Gray
                    }
                    $succeeded = $true
                } catch {
                    $hresult = $_.Exception.HResult
                    $msg     = $_.Exception.Message

                    # 0x800704c7 = -2147023673 (ERROR_CANCELLED)
                    if ($hresult -eq -2147023673 -or $msg -match '0x800704c7') {
                        Write-Status "Windows Update cancelled by system (0x800704c7), attempt $attempt/$maxRetries" -Type Warning

                        if ($attempt -lt $maxRetries) {
                            Write-Host "  Restarting Windows Update service before retry..." -ForegroundColor Gray
                            Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 5
                        }
                    } else {
                        # Unknown error &#x2014; don't retry
                        throw
                    }
                }
            }

            if (-not $succeeded) {
                throw "Windows Update failed after $maxRetries attempts (0x800704c7). A pending reboot may be required."
            }

            # Handle reboot separately after successful install
            if (-not $SkipReboot) {
                $rebootRequired = (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)
                if ($rebootRequired) {
                    Write-Status 'A reboot is required to finish installing updates.' -Type Warning
                }
            }
        }
    } else {
        Write-Section 'Windows Update'
        Write-Status 'PSWindowsUpdate module not found. Install with: Install-Module PSWindowsUpdate -Force' -Type Warning
        $updateResults.Skipped.Add('WindowsUpdate')
    }
} else {
    $updateResults.Skipped.Add('WindowsUpdate')
}

# ── Microsoft Store Apps ─────────────────────────────────────────────────────────
Invoke-Update -Name 'StoreApps' -Title 'Microsoft Store Apps' -Disabled:$SkipStoreApps -RequiresAdmin -Action {
    $timeoutSeconds = 45
    Write-Host "  Triggering Store app update scan (timeout: ${timeoutSeconds}s)..." -ForegroundColor Gray

    $job = Start-Job -ScriptBlock {
        $appMgmt = Get-CimInstance -Namespace "Root\cimv2\mdm\dmmap" `
            -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" `
            -OperationTimeoutSec 15
        $null = $appMgmt | Invoke-CimMethod -MethodName UpdateScanMethod -OperationTimeoutSec 30 -ErrorAction Stop
        'UpdateScanMethod invoked'
    }

    try {
        if (-not (Wait-Job -Job $job -Timeout $timeoutSeconds)) {
            Stop-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
            Write-Status "Timed out after ${timeoutSeconds}s &#x2014; MDM scan unreliable without Intune enrollment (Store already covered by winget msstore)" -Type Warning
            return
        }

        $jobErrors = @($job.ChildJobs | ForEach-Object { $_.Error } | Where-Object { $_ })
        if ($jobErrors) {
            Write-Status "Store scan unavailable (MDM method not supported without Intune enrollment)" -Type Warning
            return
        }
        $jobOutput = (Receive-Job -Job $job -ErrorAction SilentlyContinue 2>&1 | Out-String).Trim()
        if ($jobOutput) { Write-FilteredOutput $jobOutput }
    } finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

# ── WSL ──────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'WSL' -Title 'Windows Subsystem for Linux' -Disabled:$SkipWSL -RequiresCommand 'wsl' -RequiresAdmin -Action {
    # Update the WSL kernel/platform itself
    $out = (wsl --update 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }

    # Optionally update packages inside each running distro
    if (-not $SkipWSLDistros) {
        $distros = wsl --list --quiet 2>&1 | Where-Object { $_ -and $_ -notmatch '^\s*$' }
        foreach ($distro in $distros) {
            $distroName = $distro.Trim() -replace '\x00', ''   # strip null bytes from UTF-16 output
            if (-not $distroName) { continue }
            Write-Host "  Updating packages in WSL distro: $distroName" -ForegroundColor Gray
            # Try apt-get first, then pacman (Arch), then zypper (openSUSE)
            $aptResult = wsl -d $distroName -- bash -c "command -v apt-get &>/dev/null && sudo apt-get update -qq && sudo apt-get upgrade -y -qq 2>&1 | tail -3" 2>&1
            if ($LASTEXITCODE -ne 0) {
                $pacResult = wsl -d $distroName -- bash -c "command -v pacman &>/dev/null && sudo pacman -Syu --noconfirm 2>&1 | tail -3" 2>&1
                if ($LASTEXITCODE -ne 0) {
                    wsl -d $distroName -- bash -c "command -v zypper &>/dev/null && sudo zypper refresh && sudo zypper update -y" 2>&1 | Out-Null
                } else {
                    if ($pacResult) { Write-Host "    $pacResult" -ForegroundColor Gray }
                }
            } else {
                if ($aptResult) { Write-Host "    $aptResult" -ForegroundColor Gray }
            }
        }
    }
}

# ── Defender Signatures ──────────────────────────────────────────────────────────
Invoke-Update -Name 'DefenderSignatures' -Title 'Microsoft Defender Signatures' -Disabled:$SkipDefender -RequiresCommand 'Update-MpSignature' -RequiresAdmin -Action {
    $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($mpStatus -and -not ($mpStatus.AMServiceEnabled -and $mpStatus.AntivirusEnabled)) {
        Write-Status 'Microsoft Defender is not the active AV; skipping signature update' -Type Info
        return
    }
    try {
        Update-MpSignature -ErrorAction Stop 2>&1 | Out-Null
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'completed with errors') {
            $mpStatusAfter = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($mpStatusAfter -and $mpStatusAfter.AntivirusSignatureLastUpdated) {
                if (((Get-Date) - $mpStatusAfter.AntivirusSignatureLastUpdated).TotalDays -lt 2) {
                    Write-Status 'Defender partial source errors, but signatures appear current' -Type Warning
                    return
                }
            }
        }
        throw
    }
}

# ════════════════════════════════════════════════════════════
#  DEVELOPMENT TOOLS
# ════════════════════════════════════════════════════════════

# ── npm ──────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'npm' -Title 'npm (Node.js)' -RequiresCommand 'npm' -Action {
    $currentNpm = (npm --version 2>&1).Trim()
    $latestNpm  = (npm view npm version 2>&1).Trim()
    if ($currentNpm -ne $latestNpm) { npm install -g npm@latest 2>&1 | Out-Null }
    $outdatedJson = (npm outdated -g --json 2>$null | Out-String).Trim()
    if ($outdatedJson) {
        $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
            # BUG FIX: was `@pkgs` (splatting syntax) &#x2014; use $pkgs so the array expands correctly
            $pkgs = $outdated.PSObject.Properties.Name
            Write-Host "  Updating $($pkgs.Count) package(s): $($pkgs -join ', ')" -ForegroundColor Gray
            npm install -g $pkgs 2>&1 | Out-Null
        } else {
            Write-Host "  All global packages are up to date" -ForegroundColor Gray
        }
    } else {
        Write-Host "  All global packages are up to date" -ForegroundColor Gray
    }
    npm cache clean --force *> $null
}

# ── pnpm ─────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'pnpm' -RequiresCommand 'pnpm' -SlowOperation -Action {
    pnpm update -g 2>&1 | Out-Null
}

# ── Bun ──────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Bun' -RequiresCommand 'bun' -SlowOperation -Action {
    bun upgrade 2>&1 | Out-Null
}

# ── Deno ─────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Deno' -RequiresCommand 'deno' -SlowOperation -Action {
    $denoPath = (Get-Command deno).Source
    if ($denoPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
        return
    }
    if ($denoPath -like '*WinGet*' -or $denoPath -like '*winget*') {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
        return
    }

    $env:NO_COLOR = '1'
    try {
        $out = (cmd /c "deno upgrade 2>&1" | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    } finally {
        Remove-Item Env:\NO_COLOR -ErrorAction SilentlyContinue
    }
}

# ── Python / pip ─────────────────────────────────────────────────────────────────
Write-Section 'Python / pip'
$pythonCmd = $null
if (Test-Command 'python') {
    $pythonCmd = 'python'
} else {
    foreach ($ver in @(314, 313, 312, 311, 310)) {
        $py = "C:\Program Files\Python$ver\python.exe"
        if (Test-Path $py) { $pythonCmd = $py; break }
    }
}
if ($pythonCmd) {
    Invoke-Update -Name 'pip' -NoSection -Action {
        Write-Host "  Using: $pythonCmd" -ForegroundColor Gray
        & $pythonCmd -m pip install --upgrade pip 2>&1 | Out-Null
    }
} else {
    Write-Status 'Python not found' -Type Warning
    $updateResults.Skipped.Add('pip')
}

# ── uv ───────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv' -Title 'UV Package Manager' -RequiresCommand 'uv' -Action {
    $uvPath = (Get-Command uv).Source
    if ($uvPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($uvPath -like "*pip*" -or $uvPath -like "*Python*") {
        Write-Host "  Managed by pip (update with: pip install --upgrade uv)" -ForegroundColor Gray
    } else {
        $out = (uv self update 2>&1 | Out-String).Trim()
        if ($out -match 'error') {
            Write-Host "  uv self-update not supported; trying winget..." -ForegroundColor Gray
            Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                'upgrade', '--id', 'astral-sh.uv',
                '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
            ) | Out-Null
        } elseif ($out) {
            Write-Host $out -ForegroundColor Gray
        }
    }
}

# ── uv tools ─────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv-tools' -Title 'uv Tool Installs' -Disabled:$SkipUVTools -RequiresCommand 'uv' -SlowOperation -Action {
    $out = (cmd /c "uv tool upgrade --all 2>&1" | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -and $out -match '(?i)unknown command|unrecognized') {
        Write-Status 'uv tool upgrade --all not supported by this version' -Type Info
        return
    }
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Rust ─────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Rust' -RequiresCommand 'rustup' -Action {
    $out = (cmd /c "rustup update 2>&1" | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Cargo global binaries ────────────────────────────────────────────────────────
if ((Test-Command 'cargo') -and -not $FastMode) {
    $hasCargoUpdate = $false
    try {
        $null = cargo install-update --version 2>$null
        $hasCargoUpdate = ($LASTEXITCODE -eq 0)
    } catch { }

    if ($hasCargoUpdate) {
        Invoke-Update -Name 'cargo-binaries' -Title 'Cargo Global Binaries' -Action {
            $out = (cargo install-update -a 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    } else {
        Write-Status 'cargo-update not installed (run: cargo install cargo-update)' -Type Info
        $updateResults.Skipped.Add('cargo-binaries')
    }
} elseif (Test-Command 'cargo') {
    $updateResults.Skipped.Add('cargo-binaries')
}

# ── Go ───────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Go' -RequiresCommand 'go' -Action {
    $goPath = (Get-Command go).Source
    if ($goPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($goPath -like '*winget*' -or $goPath -like '*WinGet*') {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
    } else {
        # Try updating via winget
        try {
            $goResult = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                'upgrade', '--id', 'GoLang.Go',
                '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
            )
            if ($goResult.ExitCode -eq 0 -and $goResult.Output -match 'Successfully installed') {
                Write-Host "  Updated via winget (GoLang.Go)" -ForegroundColor Gray
            } else {
                Write-Host "  No newer version available or not managed by winget" -ForegroundColor Gray
            }
        } catch {
            Write-Status "Go winget upgrade: $_" -Type Warning
        }
    }
    if (-not $SkipDestructive) {
        Write-Host "  Cleaning module cache (use -SkipDestructive to skip)..." -ForegroundColor Gray
        go clean -modcache 2>&1 | Out-Null
    }
}

# ── .NET tools ───────────────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet' -Title '.NET Tools' -RequiresCommand 'dotnet' -Action {
    $toolLines = dotnet tool list -g 2>&1 | Select-Object -Skip 2 | Where-Object { $_ -match '\S' }
    if (-not $toolLines) { Write-Host "  No global .NET tools installed" -ForegroundColor Gray; return }

    $updatedCount = 0
    foreach ($line in $toolLines) {
        $parts = $line -split '\s+', 3
        if ($parts.Count -lt 2) { continue }
        $toolId      = $parts[0].Trim()
        $currentVer  = $parts[1].Trim()

        try {
            $meta      = Invoke-RestMethod "https://api.nuget.org/v3-flatcontainer/$($toolId.ToLower())/index.json" -TimeoutSec 10 -ErrorAction Stop
            $latestVer = $meta.versions | Where-Object { $_ -notmatch '-' } | Select-Object -Last 1
            if (-not $latestVer) { $latestVer = $meta.versions | Select-Object -Last 1 }
        } catch { $latestVer = $null }

        if (-not $latestVer -or $latestVer -eq $currentVer) { continue }

        $out = (dotnet tool update -g $toolId 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
        $updatedCount++
    }

    if ($updatedCount -eq 0) { Write-Host "  All .NET tools are up to date" -ForegroundColor Gray }
}

# ── .NET workloads (NEW) ─────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet-workloads' -Title '.NET Workloads' -RequiresCommand 'dotnet' -Action {
    $workloads = (dotnet workload list 2>&1 | Out-String)
    if ($workloads -notmatch 'No workloads are installed') {
        $out = (dotnet workload update 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    } else {
        Write-Status 'No .NET workloads installed' -Type Info
    }
}

# ── GitHub CLI extensions ────────────────────────────────────────────────────────
Invoke-Update -Name 'gh-extensions' -Title 'GitHub CLI Extensions' -RequiresCommand 'gh' -Action {
    $ghExt = (gh extension list 2>&1 | Out-String).Trim()
    if ($ghExt) {
        gh extension upgrade --all 2>&1 | Out-Null
    } else {
        Write-Status 'No gh extensions installed' -Type Info
    }
}

# ── VS Code extensions ───────────────────────────────────────────────────────────
Invoke-Update -Name 'vscode-extensions' -Title 'VS Code Extensions' -Disabled:$SkipVSCodeExtensions -RequiresAnyCommand @('code', 'code-insiders') -SlowOperation -Action {
    $codeCmd = if (Test-Command 'code') { 'code' } elseif (Test-Command 'code-insiders') { 'code-insiders' } else { $null }
    $out = (& $codeCmd --update-extensions 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── pipx ─────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'pipx' -Title 'pipx (Python CLI Tools)' -RequiresCommand 'pipx' -Action {
    pipx upgrade-all 2>&1 | Out-Null
}

# ── Poetry ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Poetry' -Title 'Poetry (Python Packaging)' -Disabled:$SkipPoetry -RequiresCommand 'poetry' -Action {
    $out = (poetry self update 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}

# ── Composer ─────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Composer' -Title 'Composer (PHP)' -Disabled:$SkipComposer -RequiresCommand 'composer' -Action {
    $selfOut   = (composer self-update --no-interaction 2>&1 | Out-String).Trim()
    if ($selfOut)   { Write-Host $selfOut   -ForegroundColor Gray }
    $globalOut = (composer global update --no-interaction 2>&1 | Out-String).Trim()
    if ($globalOut) { Write-Host $globalOut -ForegroundColor Gray }
}

# ── RubyGems ─────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'RubyGems' -Title 'RubyGems' -Disabled:$SkipRuby -RequiresCommand 'gem' -SlowOperation -Action {
    $sysOut = (gem update --system 2>&1 | Out-String).Trim()
    if ($sysOut) { Write-Host $sysOut -ForegroundColor Gray }
    $gemOut = (gem update 2>&1 | Out-String).Trim()
    if ($gemOut) { Write-Host $gemOut -ForegroundColor Gray }
}

# ── Volta ────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Volta' -Title 'Volta (Node Version Manager)' -RequiresCommand 'volta' -SlowOperation -Action {
    $voltaPath = (Get-Command volta).Source
    if ($voltaPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } else {
        Write-Status 'Volta installed standalone - update via its installer' -Type Info
    }
}

# ── Git Credential Manager (NEW) ─────────────────────────────────────────────────
Invoke-Update -Name 'git-credential-manager' -Title 'Git Credential Manager' -RequiresCommand 'git-credential-manager' -Action {
    # GCM ships with Git for Windows and has no self-update command; update via winget Git.Git
    $gcmPath = (Get-Command git-credential-manager).Source
    if ($gcmPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } else {
        try {
            Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                'upgrade', '--id', 'Git.Git',
                '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
            ) | Out-Null
            Write-Host "  Updated via winget (Git.Git)" -ForegroundColor Gray
        } catch {
            Write-Status "GCM winget upgrade: $_" -Type Warning
        }
    }
}

# ── Oh My Posh (NEW) ─────────────────────────────────────────────────────────────
Invoke-Update -Name 'oh-my-posh' -Title 'Oh My Posh' -RequiresCommand 'oh-my-posh' -Action {
    $ompPath = (Get-Command oh-my-posh).Source
    if ($ompPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($ompPath -like "*winget*") {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
    } elseif ($ompPath -like "*WindowsApps*") {
        Write-Status "Installed as MSIX &#x2014; update via Microsoft Store or reinstall with winget: winget install JanDeDobbeleer.OhMyPosh" -Type Info
    } else {
        $out = (oh-my-posh upgrade 2>&1 | Out-String).Trim()
        if ($out -match '(?i)not supported|error|&#x274C;|failed') {
            Write-Status "oh-my-posh upgrade failed: $out" -Type Warning
        } elseif ($out) {
            Write-Host $out -ForegroundColor Gray
        }
    }
}

# ── yt-dlp (NEW) ─────────────────────────────────────────────────────────────────
Invoke-Update -Name 'yt-dlp' -Title 'yt-dlp' -RequiresCommand 'yt-dlp' -Action {
    $ytdlpPath = (Get-Command yt-dlp).Source
    if ($ytdlpPath -like "*scoop*") {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($ytdlpPath -like "*pip*" -or $ytdlpPath -like "*Python*" -or $ytdlpPath -like "*Scripts*") {
        if ($pythonCmd) {
            & $pythonCmd -m pip install --upgrade yt-dlp 2>&1 | Out-Null
            Write-Host "  Updated via pip" -ForegroundColor Gray
        } else {
            Write-Status "yt-dlp installed via pip but Python not found" -Type Warning
        }
    } else {
        $out = (yt-dlp -U 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    }
}

# ── fnm (Fast Node Manager) ──────────────────────────────────────────────────────
Invoke-Update -Name 'fnm' -Title 'fnm (Fast Node Manager)' -RequiresCommand 'fnm' -SlowOperation -Action {
    $fnmPath = (Get-Command fnm).Source
    if ($fnmPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($fnmPath -like '*winget*' -or $fnmPath -like '*WinGet*') {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
    } else {
        Write-Status 'fnm installed standalone - update via its installer or scoop' -Type Info
    }
}

# ── mise (polyglot tool version manager) ─────────────────────────────────────────
Invoke-Update -Name 'mise' -Title 'mise (Tool Version Manager)' -RequiresCommand 'mise' -SlowOperation -Action {
    $misePath = (Get-Command mise).Source
    if ($misePath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($misePath -like '*winget*' -or $misePath -like '*WinGet*') {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
    } else {
        $out = (mise self-update --yes 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    }
    # Also update all mise-managed runtimes
    $pluginsOut = (mise plugins ls 2>&1 | Out-String).Trim()
    if ($pluginsOut) {
        Write-Host "  Upgrading mise plugins and tools..." -ForegroundColor Gray
        mise plugins upgrade 2>&1 | Out-Null
    }
}

# ── juliaup ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'juliaup' -Title 'Julia (juliaup)' -RequiresCommand 'juliaup' -SlowOperation -Action {
    $juliaupPath = (Get-Command juliaup).Source
    if ($juliaupPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
    } elseif ($juliaupPath -like '*winget*' -or $juliaupPath -like '*WinGet*') {
        Write-Host "  Managed by winget (already updated)" -ForegroundColor Gray
    } else {
        $out = (juliaup update 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray }
    }
}

# ── PowerShell modules / PSResources ─────────────────────────────────────────────
Invoke-Update -Name 'pwsh-resources' -Title 'PowerShell Modules / Resources' -Disabled:$SkipPowerShellModules -Action {
    $usedProvider = $false

    # PSResourceGet (modern, PS 7.4+) &#x2014; parallelize per-resource updates
    if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
        $usedProvider = $true
        $resources = Get-InstalledPSResource -ErrorAction SilentlyContinue
        if ($resources) {
            Write-Host "  Updating $(@($resources).Count) PSResource(s)..." -ForegroundColor Gray
            $supportsAcceptLicense = (Get-Command Update-PSResource -ErrorAction SilentlyContinue)?.Parameters.ContainsKey('AcceptLicense')

            # Use parallel if PS7 and more than a handful of resources
            if ($PSVersionTable.PSVersion.Major -ge 7 -and @($resources).Count -gt 3) {
                $resources | ForEach-Object -Parallel {
                    $psrArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                    if ($using:supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { }
                } -ThrottleLimit 4
            } else {
                foreach ($resource in $resources) {
                    $psrArgs = @{ Name = $resource.Name; ErrorAction = 'SilentlyContinue' }
                    if ($supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { }
                }
            }
        } else {
            Write-Status 'No installed PSResources found' -Type Info
        }
    }

    # PowerShellGet (legacy fallback)
    if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
        $usedProvider = $true
        $modules = Get-InstalledModule -ErrorAction SilentlyContinue
        if ($modules) {
            Write-Host "  Updating $(@($modules).Count) PowerShellGet module(s)..." -ForegroundColor Gray
            $supportsAcceptLicense = (Get-Command Update-Module -ErrorAction SilentlyContinue)?.Parameters.ContainsKey('AcceptLicense')
            $modules | ForEach-Object {
                $updateArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                if ($supportsAcceptLicense) { $updateArgs['AcceptLicense'] = $true }
                Update-Module @updateArgs 2>&1 | Out-Null
            }
        } else {
            Write-Status 'No installed PowerShellGet modules found' -Type Info
        }
    }

    if (-not $usedProvider) {
        Write-Status 'No PowerShell module update provider found (PSResourceGet or PowerShellGet)' -Type Info
    }
}

# ════════════════════════════════════════════════════════════
#  CLEANUP
# ════════════════════════════════════════════════════════════
if ($SkipCleanup) {
    Write-Section 'System Cleanup'
    Write-Status 'Skipped (use -SkipCleanup to re-enable)' -Type Info
    $updateResults.Skipped.Add('cleanup')
} else {

Write-Section 'System Cleanup'

# Temp files (older than 7 days)
try {
    $tempPath = $env:TEMP
    if ($tempPath -and (Test-Path $tempPath) -and $tempPath -ne (Split-Path $tempPath -Qualifier)) {
        if ($PSCmdlet.ShouldProcess($tempPath, 'Remove temp files older than 7 days')) {
            $cutoff = (Get-Date).AddDays(-7)
            Get-ChildItem -Path $tempPath -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "Temp files cleared (older than 7 days)" -Type Success
        }
    }
} catch {
    Write-Status "Temp cleanup partially failed (normal)" -Type Warning
}

# DNS cache
try {
    Clear-DnsClientCache -ErrorAction SilentlyContinue
    Write-Status "DNS cache flushed" -Type Success
} catch {
    Write-Status "DNS cache flush failed" -Type Warning
}

# Recycle Bin
try {
    if ($PSCmdlet.ShouldProcess('Recycle Bin', 'Empty')) {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
        Write-Status "Recycle Bin emptied" -Type Success
    }
} catch {
    Write-Status "Recycle Bin cleanup failed" -Type Warning
}

# Crash dumps
$crashDumpPath = Join-Path $env:LOCALAPPDATA 'CrashDumps'
if (Test-Path $crashDumpPath) {
    try {
        $dumps = Get-ChildItem -Path $crashDumpPath -ErrorAction SilentlyContinue
        if ($dumps) {
            $dumps | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Status "Crash dumps cleared ($($dumps.Count) files)" -Type Success
        }
    } catch {
        Write-Status "Crash dump cleanup failed" -Type Warning
    }
}

# Windows Error Reporting queue (NEW)
$werPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\WER\ReportQueue'
if (Test-Path $werPath) {
    try {
        Get-ChildItem -Path $werPath -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-Status "WER report queue cleared" -Type Success
    } catch {
        Write-Status "WER cleanup failed" -Type Warning
    }
}

# Admin-only cleanup
if ($isAdmin) {
    try {
        Write-Host "  Cleaning WinSxS component store..." -ForegroundColor Gray
        DISM /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-Null
        Write-Status "DISM component store cleaned" -Type Success
    } catch {
        Write-Status "DISM cleanup failed" -Type Warning
    }

    try {
        Delete-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
        Write-Status "Delivery Optimization cache cleared" -Type Success
    } catch {
        Write-Status "Delivery Optimization cleanup failed" -Type Warning
    }

    # Prefetch cleanup (NEW - optional, rarely needed but useful on old HDDs)
    $prefetchPath = 'C:\Windows\Prefetch'
    if ((Test-Path $prefetchPath) -and -not $SkipDestructive) {
        try {
            Get-ChildItem -Path $prefetchPath -Filter '*.pf' -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Status "Prefetch files cleared" -Type Success
        } catch {
            Write-Status "Prefetch cleanup failed (normal if Superfetch is disabled)" -Type Warning
        }
    }
}

$updateResults.Success.Add('cleanup')

} # end -not $SkipCleanup

# ════════════════════════════════════════════════════════════
#  SUMMARY
# ════════════════════════════════════════════════════════════
$duration = (Get-Date) - $startTime
$dryTag   = if ($DryRun) { ' [DRY-RUN]' } else { '' }

Write-Host "`n$('=' * 54)" -ForegroundColor Green
Write-Host (" UPDATE COMPLETE{0} -- {1}" -f $dryTag, $duration.ToString('hh\:mm\:ss')) -ForegroundColor Green
Write-Host "$('=' * 54)" -ForegroundColor Green

if ($updateResults.Success.Count -gt 0) {
    Write-Host "`n[OK] Succeeded ($($updateResults.Success.Count)): $($updateResults.Success -join ', ')" -ForegroundColor Green
}
if ($updateResults.Failed.Count -gt 0) {
    Write-Host "[X] Failed    ($($updateResults.Failed.Count)): $($updateResults.Failed -join ', ')" -ForegroundColor Red
}
if ($updateResults.Skipped.Count -gt 0) {
    Write-Host "[!] Skipped   ($($updateResults.Skipped.Count)): $($updateResults.Skipped -join ', ')" -ForegroundColor Yellow
}

Write-Host ""

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { }
}

if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }
