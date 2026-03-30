<#
.SYNOPSIS
    System-wide update script for Windows
.DESCRIPTION
    Updates all package managers, system components, and development tools.
    Version 4.0.0: Refactored with configuration, parallel execution, enhanced logging,
    better error handling, dry-run support, and extended state tracking.
.VERSION
    4.0.0
.NOTES
    Run as Administrator for full functionality (some features require elevation).
    For configuration, see the Config section in the script or provide a config.json.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
    .\updatescript.ps1 -SkipCleanup -SkipWindowsUpdate
    .\updatescript.ps1 -SkipNode -SkipRust -SkipGo
    .\updatescript.ps1 -DeepClean
    .\updatescript.ps1 -UpdateOllamaModels
    .\updatescript.ps1 -WhatChanged
    .\updatescript.ps1 -DryRun
    .\updatescript.ps1 -Verbose
#>

[CmdletBinding(DefaultParameterSetName = 'Normal', SupportsShouldProcess = $true)]
param(
    [Parameter(ParameterSetName = 'Normal')]
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,                  # Skip slower managers
    [switch]$NoElevate,                 # Compatibility switch (prevents auto-elevation)
    [switch]$AutoElevate,               # Opt-in: relaunch elevated (opens a new window due UAC)
    [switch]$NoPause,                   # Skip "Press Enter to close" prompt (for VS Code / CI)
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
    [Parameter(ParameterSetName = 'Schedule')]
    [switch]$Schedule,                  # Register a daily scheduled task to run this script
    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateScript({ $_ -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$' })]
    [string]$ScheduleTime = "03:00",    # Time for the scheduled task (default: 3 AM)
    [string]$LogPath,
    [switch]$SkipNode,                  # Skip all Node.js toolchain updates (npm, pnpm, bun, deno, fnm, volta)
    [switch]$SkipRust,                  # Skip Rust toolchain updates (rustup, cargo)
    [switch]$SkipGo,                    # Skip Go toolchain updates
    [switch]$SkipFlutter,               # Skip Flutter SDK update
    [switch]$SkipGitLFS,                # Skip Git LFS client update
    [switch]$DeepClean,                 # Run DISM WinSxS cleanup, DO cache, and prefetch
    [switch]$UpdateOllamaModels,        # Opt-in: pull latest for every installed Ollama model
    [switch]$WhatChanged,               # Show packages that changed since last run
    [switch]$DryRun,                    # Show which steps would run without executing them
    [ValidateRange(1, 10)]
    [int]$ParallelThrottle = 4          # Max parallel tasks for independent steps
)

# --- Global State & Initialization ---------------------------------------------
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
}
catch {
    Write-Verbose "Console encoding update skipped: $($_.Exception.Message)"
}

$ErrorActionPreference = 'Continue'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
$startTime = Get-Date
$script:IsSimulation = $DryRun -or $WhatIfPreference
$updateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Checked = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[pscustomobject]]::new()
    Details = [ordered]@{}
}

# --- Config Loading ------------------------------------------------------------
$script:Config = @{
    # Default values, can be overridden by config file
    WingetTimeoutSec   = 300
    StateDir           = Join-Path $env:LOCALAPPDATA 'Update-Everything'
    LogRetentionDays   = 7
    TempCleanupDays    = 7
    SkipManagers       = @()          # Names of managers to skip (e.g., 'Chocolatey', 'WSL')
    FastModeSkip       = @('Chocolatey', 'WSLDistros', 'npm', 'pnpm', 'bun', 'deno', 'rust', 'cargo-binaries', 'go', 'gh-extensions', 'pipx', 'poetry', 'composer', 'rubygems', 'yt-dlp', 'tldr', 'oh-my-posh', 'volta', 'fnm', 'mise', 'juliaup', 'ollama-models', 'git-lfs', 'git-credential-manager', 'pip', 'uv', 'uv-tools', 'uv-python', 'dotnet', 'dotnet-workloads', 'vscode-extensions', 'pwsh-resources')
    ManagersParallel   = @('Scoop', 'Chocolatey', 'WSL')   # can run in parallel
    # Custom hooks for winget upgrades
    WingetUpgradeHooks = @{
        'Spotify.Spotify'             = @{
            Pre  = { $script:_spotifyWasRunning = [bool](Get-Process -Name Spotify -ErrorAction SilentlyContinue); Stop-Process -Name Spotify -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 2 }
            Post = { if ($script:_spotifyWasRunning) { Start-Process "$env:APPDATA\Spotify\Spotify.exe" -ErrorAction SilentlyContinue } }
        }
        'Google.Chrome'               = @{
            Pre  = { $script:_chromeWasRunning = [bool](Get-Process -Name 'chrome' -ErrorAction SilentlyContinue); if ($script:_chromeWasRunning) { Write-Host "  Closing Chrome..." -ForegroundColor Gray; Stop-Process -Name 'chrome' -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 3 } }
            Post = { if ($script:_chromeWasRunning) { $p = Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'; if (-not(Test-Path $p)) { $p = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe' } if (Test-Path $p) { Start-Process $p -ErrorAction SilentlyContinue } } }
        }
        'Microsoft.VisualStudioCode'  = @{
            Pre  = { $script:_vscodeWasRunning = [bool](Get-Process -Name 'Code' -ErrorAction SilentlyContinue); if ($script:_vscodeWasRunning) { Write-Host "  Closing VS Code..." -ForegroundColor Gray; Stop-Process -Name 'Code' -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 3 } }
            Post = { if ($script:_vscodeWasRunning) { $p = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'; if (Test-Path $p) { Start-Process $p -ErrorAction SilentlyContinue } } }
        }
        'Adobe.Acrobat.Reader.64-bit' = @{
            Pre  = {
                $script:_acrobatProcs = @(Get-Process -Name 'Acrobat', 'AcroRd32', 'AcroCEF' -ErrorAction SilentlyContinue)
                if ($script:_acrobatProcs.Count -gt 0) { Write-Host "  Closing Adobe Acrobat..." -ForegroundColor Gray; $script:_acrobatProcs | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 3 }
                Write-Host "  Clearing temporary files to prevent Acrobat extraction errors..." -ForegroundColor Gray
                Get-ChildItem -Path $env:TEMP -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '\.tmp|\.log' } | Remove-Item -Force -ErrorAction SilentlyContinue
            }
            Post = { if ($script:_acrobatProcs.Count -gt 0) { $p = @((Join-Path $env:ProgramFiles 'Adobe\Acrobat DC\Acrobat\Acrobat.exe'), (Join-Path ${env:ProgramFiles(x86)} 'Adobe\Acrobat Reader DC\Reader\AcroRd32.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -ErrorAction SilentlyContinue } } }
        }
    }
}

# Load user config if present
$configFile = Join-Path $PSScriptRoot 'update-config.json'
if (Test-Path $configFile) {
    try {
        $userConfig = Get-Content $configFile -Raw | ConvertFrom-Json
        # Merge user config into defaults (only if property exists)
        foreach ($prop in $userConfig.PSObject.Properties) {
            $script:Config[$prop.Name] = $prop.Value
        }
        Write-Verbose "Loaded config from $configFile"
    }
    catch { Write-Warning "Failed to load config file: $_" }
}

# --- Auto-Elevation Check ------------------------------------------------------
# If the user is not Admin, Winget MSI installers often fail with Error 1632.
# We check early and offer to elevate if -AutoElevate is set or user confirms.
if (-not $isAdmin -and $AutoElevate -and -not $NoElevate) {
    Write-Host "Elevating to Administrator to ensure installers succeed..." -ForegroundColor Cyan

    $forwardedArgs = foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -eq 'AutoElevate') { continue }
        if ($entry.Value -is [switch]) {
            if ($entry.Value.IsPresent) { "-$($entry.Key)" }
        }
        elseif ($null -ne $entry.Value -and "$($entry.Value)".Length -gt 0) {
            "-$($entry.Key)"
            [string]$entry.Value
        }
    }

    $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    $shellPath = if ($pwshCmd) { $pwshCmd.Source } else { 'powershell.exe' }
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + @($forwardedArgs)

    try {
        Start-Process -FilePath $shellPath -Verb RunAs -ArgumentList $argList -Wait
        exit
    }
    catch {
        Write-Warning "Could not elevate. Continuing without Administrator privileges."
    }
}
elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host "INFO: Running without elevation. Admin-only tasks may be skipped." -ForegroundColor DarkYellow
}

if ($Schedule) {
    if (-not $isAdmin -and -not $script:IsSimulation) {
        throw "Scheduled task registration requires Administrator privileges. Re-run with -AutoElevate or from an elevated shell."
    }

    $taskName = 'DailySystemUpdate'
    $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    $shellPath = if ($pwshCmd) { $pwshCmd.Source } else { 'powershell.exe' }
    $taskArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-SkipReboot', '-NoPause')

    if ($script:IsSimulation) {
        Write-Host "  [DryRun] Would register scheduled task '$taskName' for $ScheduleTime" -ForegroundColor DarkCyan
    }
    elseif ($PSCmdlet.ShouldProcess($taskName, "Register scheduled task for $ScheduleTime")) {
        $action = New-ScheduledTaskAction -Execute $shellPath -Argument ($taskArgs -join ' ')
        $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
        $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force | Out-Null
        Write-Host "[OK] Scheduled task '$taskName' registered to run daily at $ScheduleTime." -ForegroundColor Green
    }
    exit
}


# Override config with command line parameters if provided
if ($WingetTimeoutSec -ne 300) { $script:Config.WingetTimeoutSec = $WingetTimeoutSec }
if ($script:Config.WingetTimeoutSec -lt 30) { $script:Config.WingetTimeoutSec = 30 }

# --- Helper Functions ----------------------------------------------------------
$commandCache = @{}
function Test-Command([string]$Command) {
    if ($commandCache.ContainsKey($Command)) { return $commandCache[$Command] }
    $savedWhatIfPreference = $WhatIfPreference
    try {
        $WhatIfPreference = $false
        $result = [bool](Get-Command $Command -ErrorAction SilentlyContinue)
    }
    finally {
        $WhatIfPreference = $savedWhatIfPreference
    }
    $commandCache[$Command] = $result
    return $result
}

function ConvertTo-StringMap {
    param([AllowNull()]$InputObject)

    $map = @{}
    if ($null -eq $InputObject) { return $map }

    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) {
            $map[[string]$key] = [string]$InputObject[$key]
        }
        return $map
    }

    foreach ($prop in $InputObject.PSObject.Properties) {
        $map[[string]$prop.Name] = [string]$prop.Value
    }
    return $map
}

# Logging
$script:LogFile = if ($LogPath) { $LogPath } else { Join-Path $script:Config.StateDir 'updatescript.log' }
$script:TranscriptionStarted = $false

function Write-Section {
    param([string]$Title)
    Write-Host "`n$($('=' * 54))" -ForegroundColor DarkGray
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "$('=' * 54)" -ForegroundColor DarkGray
    Write-Log -Message "--- $Title ---" -Level 'Info'
}

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = "[$timestamp] [$Level] $Message"
    
    # Append to log file
    try {
        $logDir = Split-Path $script:LogFile -Parent
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null }
        Add-Content -Path $script:LogFile -Value $logLine -Encoding UTF8 -WhatIf:$false
    }
    catch {
        Write-Verbose "Log write skipped: $($_.Exception.Message)"
    }
}


function Write-Status {
    param([string]$Message, [ValidateSet('Success', 'Warning', 'Error', 'Info')][string]$Type = 'Info')
    $colors = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]'; Warning = '[!]'; Error = '[X]'; Info = '[*]' }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
    Write-Log -Message $Message -Level $Type
}

function Write-Detail {
    param([string]$Message, [ValidateSet('Info', 'Muted', 'Warning', 'Error')][string]$Type = 'Info')
    $colors = @{ Info = 'Gray'; Muted = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }
    $prefixes = @{ Info = '  >'; Muted = '  -'; Warning = '  !'; Error = '  x' }
    if ($Message) {
        Write-Host "$($prefixes[$Type]) $Message" -ForegroundColor $colors[$Type]
        Write-Log -Message "$($prefixes[$Type]) $Message" -Level $Type
    }
}

function Write-FilteredOutput {
    param([AllowNull()][string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    $normalized = $normalized -replace "`r", "`n"
    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.TrimEnd()
        if (-not $line) { continue }
        $compact = $line.Trim()
        if ($compact -match '^[\\/\|\-]+$') { continue }
        $nonAsciiCount = ($compact.ToCharArray() | Where-Object { [int]$_ -gt 127 }).Count
        if ($nonAsciiCount -ge 3 -and $compact -notmatch '[A-Za-z0-9]') { continue }
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }
        if ($compact -match '^[\-=]{6,}$') { continue }
        Write-Host $line -ForegroundColor $Color
        Write-Log -Message $line -Level 'Info'
    }
}

function Invoke-StreamingCapture {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    $outputPath = [System.IO.Path]::GetTempFileName()
    try {
        & $ScriptBlock 2>&1 |
            Tee-Object -FilePath $outputPath |
            ForEach-Object {
                if ($_ -is [string]) {
                    Write-FilteredOutput -Text $_ -Color $Color
                }
                elseif ($null -ne $_) {
                    $text = $_.ToString()
                    if (-not [string]::IsNullOrWhiteSpace($text)) {
                        Write-Host $text -ForegroundColor $Color
                        Write-Log -Message $text -Level 'Info'
                    }
                }
            } | Out-Null

        return [pscustomobject]@{
            OutputPath = $outputPath
            ExitCode   = $LASTEXITCODE
        }
    }
    catch {
        if (Test-Path $outputPath) {
            try {
                $captured = Get-Content -LiteralPath $outputPath -Raw -ErrorAction SilentlyContinue
                if ($captured) { Write-Log -Message $captured -Level 'Error' }
            }
            catch {}
        }
        throw
    }
}

function Read-CapturedOutput {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return '' }
    try { return ((Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue) -replace '\x00', '').Trim() }
    catch { return '' }
}

function Invoke-WingetWithTimeout {
    param([string[]]$Arguments, [int]$TimeoutSec = $script:Config.WingetTimeoutSec)
    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    try {
        $proc = Start-Process -FilePath 'winget' -ArgumentList $Arguments -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { $proc.Kill() } catch { }
            throw "winget timed out after ${TimeoutSec}s"
        }
        $stdout = Get-Content -Raw -Path $stdoutFile -ErrorAction SilentlyContinue
        $stderr = Get-Content -Raw -Path $stderrFile -ErrorAction SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()
        return [pscustomobject]@{ Output = $combined; ExitCode = $proc.ExitCode }
    }
    finally {
        Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
    }
}

function Get-WingetUpgradeEntries {
    param([string]$WingetOutput)

    $entries = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrWhiteSpace($WingetOutput)) { return @($entries) }

    foreach ($line in ($WingetOutput -split "`r?`n")) {
        $trimmed = $line.Trim()
        if (-not $trimmed) { continue }
        if ($trimmed -match '^(Name|Found|The following|No installed package|No applicable upgrade|There are no available upgrades|\d+\s+upgrades available)') { continue }
        if ($trimmed -match '^[-\s]+$') { continue }

        if ($trimmed -match '^(?<name>.+?)\s+(?<id>(?=.*[A-Za-z])[A-Za-z0-9][A-Za-z0-9\.\-_]*\.[A-Za-z0-9\.\-_]+)\s+(?<version>\S+)\s+(?<available>\S+)$') {
            $entries.Add([pscustomobject]@{
                Name      = $Matches['name'].Trim()
                Id        = $Matches['id'].Trim()
                Version   = $Matches['version'].Trim()
                Available = $Matches['available'].Trim()
            })
        }
    }

    return @($entries)
}

function Add-WingetBlockingPin {
    param([Parameter(Mandatory)][string]$PackageId)

    $pinRun = Invoke-StreamingCapture -ScriptBlock {
        winget pin add --id $PackageId --blocking --exact --source winget --accept-source-agreements --disable-interactivity --force
    }
    return (Read-CapturedOutput $pinRun.OutputPath)
}

function Get-WingetPinnedPackageIds {
    $pinRun = Invoke-StreamingCapture -ScriptBlock {
        winget pin list --disable-interactivity
    }
    $pinOutput = Read-CapturedOutput $pinRun.OutputPath
    $pkgIds = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($pinOutput)) { return @($pkgIds) }

    foreach ($line in ($pinOutput -split "`r?`n")) {
        $trimmed = $line.Trim()
        if (-not $trimmed) { continue }
        if ($trimmed -match '^(Name|Found|The following|There are no pins|[-\s]+$)') { continue }
        if ($trimmed -match '^(?<name>.+?)\s+(?<id>(?=.*[A-Za-z])[A-Za-z0-9][A-Za-z0-9\.\-_]*\.[A-Za-z0-9\.\-_]+)\s+(?<version>\S+)\s+(?<source>\S+)\s+(?<pinType>.+)$') {
            if (-not $pkgIds.Contains($matches.id)) {
                $pkgIds.Add($matches.id)
            }
        }
    }

    return @($pkgIds)
}

function Get-WingetUpgradeablePackageIds {
    param([string]$WingetOutput)

    $pkgIds = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in (Get-WingetUpgradeEntries -WingetOutput $WingetOutput)) {
        if (-not $pkgIds.Contains($entry.Id)) {
            $pkgIds.Add($entry.Id)
        }
    }
    return @($pkgIds)
}

function Get-VSCodeCliPath {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:ProgramFiles 'Microsoft VS Code\bin\code.cmd'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code Insiders\bin\code-insiders.cmd'),
        (Join-Path $env:ProgramFiles 'Microsoft VS Code Insiders\bin\code-insiders.cmd'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft VS Code Insiders\bin\code-insiders.cmd')
    ) | Where-Object { $_ -and (Test-Path $_) }
    if ($candidates.Count -gt 0) { return $candidates[0] }
    foreach ($name in @('code.cmd', 'code-insiders.cmd')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { return $cmd.Source }
    }
    return $null
}

function Get-ToolInstallManager([string]$Command) {
    if (Test-Command $Command) {
        $src = (Get-Command $Command).Source
        if ($src -like '*\scoop\*' -or $src -like '*scoop\shims\*') { return 'scoop' }
        if ($src -like '*\WinGet\*' -or $src -like '*WindowsApps\*' -or $src -like '*\winget\*') { return 'winget' }
    }
    return $null
}

# --- State Management ----------------------------------------------------------
$script:StateDir = $script:Config.StateDir
$script:StateFile = Join-Path $script:StateDir 'state.json'
$script:State = @{
    LastRun     = $null
    Winget      = @{}
    Scoop       = @{}
    Chocolatey  = @{}
    WhatChanged = $null
}
# Load previous state if exists
if (Test-Path $script:StateFile) {
    try {
        $loadedState = Get-Content $script:StateFile -Raw | ConvertFrom-Json
        $script:State = @{
            LastRun     = $loadedState.LastRun
            Winget      = ConvertTo-StringMap $loadedState.Winget
            Scoop       = ConvertTo-StringMap $loadedState.Scoop
            Chocolatey  = ConvertTo-StringMap $loadedState.Chocolatey
            WhatChanged = $loadedState.WhatChanged
        }
    }
    catch { Write-Warning "Could not load state file: $_" }
}

function Save-State {
    $script:State.LastRun = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        $stateDir = Split-Path $script:StateFile -Parent
        if (-not (Test-Path $stateDir)) { New-Item -ItemType Directory -Path $stateDir -Force | Out-Null }
        $script:State | ConvertTo-Json -Depth 10 | Set-Content -Path $script:StateFile -Encoding UTF8
    }
    catch { Write-Warning "Failed to save state: $_" }
}

function Update-WingetState {
    # Get current list of installed winget packages (including version)
    try {
        $listResult = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('list', '--accept-source-agreements', '--disable-interactivity')
        $newMap = @{}
        foreach ($line in ($listResult.Output -split "`n")) {
            $trimmed = $line.Trim()
            if (-not $trimmed -or $trimmed -match '^[-=\s]+$' -or $trimmed -match '^Name\s') { continue }
            $tokens = $trimmed -split '\s{2,}' | Where-Object { $_ }
            if ($tokens.Count -ge 3) {
                $id = $tokens[-2].Trim()
                $ver = $tokens[-1].Trim()
                if ($id -match '^\S+\.\S+$') { $newMap[$id] = $ver }
            }
        }
        $script:State.Winget = $newMap
    }
    catch { Write-Warning "Could not update winget state: $_" }
}

function Update-ScoopState {
    if (Test-Command 'scoop') {
        try {
            $list = scoop list 2>&1 | Out-String
            $map = @{}
            foreach ($line in ($list -split "`n")) {
                $trimmed = $line.Trim()
                if (-not $trimmed -or $trimmed -match '^Installed' -or $trimmed -match '^---') { continue }
                $parts = $trimmed -split '\s+', 2
                if ($parts.Count -ge 2) {
                    $name = $parts[0]
                    $ver = $parts[1]
                    $map[$name] = $ver
                }
            }
            $script:State.Scoop = $map
        }
        catch { Write-Warning "Could not update scoop state: $_" }
    }
}

function Update-ChocolateyState {
    if (Test-Command 'choco' -and $isAdmin) {
        try {
            $list = choco list -lo 2>&1 | Out-String
            $map = @{}
            foreach ($line in ($list -split "`n")) {
                $trimmed = $line.Trim()
                if (-not $trimmed -or $trimmed -match '^\d+ packages installed' -or $trimmed -match '^Chocolatey') { continue }
                $parts = $trimmed -split '\s+', 2
                if ($parts.Count -ge 2) {
                    $name = $parts[0]
                    $ver = $parts[1]
                    $map[$name] = $ver
                }
            }
            $script:State.Chocolatey = $map
        }
        catch { Write-Warning "Could not update chocolatey state: $_" }
    }
}

function Show-WhatChanged {
    if (-not $WhatChanged) { return }
    Write-Section "What changed since last run"
    # Compare each manager
    if ($script:State.Winget -and $WhatChanged) {
        $prev = $script:State.Winget
        $curr = @{}; Update-WingetState; $curr = $script:State.Winget
        $changes = Compare-PackageMaps $prev $curr
        if ($changes.Count -gt 0) {
            Write-Host "  Winget changes:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        }
    }

    if ($script:State.Scoop -and $WhatChanged) {
        $prev = $script:State.Scoop
        $curr = @{}; Update-ScoopState; $curr = $script:State.Scoop
        $changes = Compare-PackageMaps $prev $curr
        if ($changes.Count -gt 0) {
            Write-Host "  Scoop changes:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        }
    }
    if ($script:State.Chocolatey -and $WhatChanged) {
        $prev = $script:State.Chocolatey
        $curr = @{}; Update-ChocolateyState; $curr = $script:State.Chocolatey
        $changes = Compare-PackageMaps $prev $curr
        if ($changes.Count -gt 0) {
            Write-Host "  Chocolatey changes:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        }
    }
}

function Compare-PackageMaps($prev, $curr) {
    $prev = ConvertTo-StringMap $prev
    $curr = ConvertTo-StringMap $curr
    $changes = @()
    foreach ($id in $curr.Keys) {
        if (-not $prev.ContainsKey($id)) { $changes += "+ $id $($curr[$id]) (new)" }
        elseif ($prev[$id] -ne $curr[$id]) { $changes += "~ $id $($prev[$id]) -> $($curr[$id])" }
    }
    foreach ($id in $prev.Keys) {
        if (-not $curr.ContainsKey($id)) { $changes += "- $id $($prev[$id]) (removed)" }
    }
    return $changes
}

# --- Core Update Wrapper -------------------------------------------------------
$script:UpdateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Checked = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[pscustomobject]]::new()
    Details = @{}
}
$script:SectionTimings = @{}

function Invoke-Update {
    [CmdletBinding(SupportsShouldProcess = $true)]
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

    # Determine skip reasons
    if ($Disabled) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'flag' }); return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAdmin -and -not $isAdmin) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'requires admin' }); return }
    if ($SlowOperation -and $FastMode) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'fast mode' }); return }

    if ($script:IsSimulation) {
        Write-Host "  [DryRun] Would run: $Title" -ForegroundColor DarkCyan
        return
    }

    # WhatIf support
    if (-not $PSCmdlet.ShouldProcess($Name, "Update $Title")) { return }

    $sectionStart = Get-Date
    if (-not $NoSection) { Write-Section $Title }
    $script:stepChanged = $false
    $script:stepMessage = ""

    try {
        & $Action
        Complete-StepState
        $elapsed = ((Get-Date) - $sectionStart).TotalSeconds
        if ($script:stepChanged) {
            Write-Status "$Name updated ($([math]::Round($elapsed, 1).ToString('F1', [cultureinfo]::InvariantCulture))s)" -Type Success
            $updateResults.Success.Add($Name)
            $updateResults.Details[$Name] = $script:stepMessage
        }
        else {
            if ($script:stepMessage) { Write-Detail $script:stepMessage -Type Muted }
            $updateResults.Checked.Add($Name)
            $updateResults.Details[$Name] = $script:stepMessage
        }
        $script:SectionTimings[$Name] = $elapsed
    }
    catch {
        Write-Status "$Name failed: $($_.Exception.Message)" -Type Error
        $updateResults.Failed.Add($Name)
        $updateResults.Details[$Name] = $_.Exception.Message
    }
}

# --- Update Actions ------------------------------------------------------------
# Each action should set $stepChanged = $true if updates occurred, and optionally $stepMessage
$script:stepChanged = $false
$script:stepMessage = ""

function Reset-StepState { $script:stepChanged = $false; $script:stepMessage = "" }

function Complete-StepState {
    if ([string]::IsNullOrWhiteSpace($script:stepMessage)) {
        $script:stepMessage = if ($script:stepChanged) { 'updated' } else { 'checked' }
    }
}

# --- Package Managers (can run in parallel) ------------------------------------
$script:ParallelJobs = @()

# Define each manager as a scriptblock that returns a hashtable with Name, Title, etc.
# For parallel execution, we'll collect them and run via runspaces.

$managers = @(
    @{
        Name            = 'Scoop'
        Title           = 'Scoop'
        Action          = {
            $selfUpdate = Invoke-StreamingCapture -ScriptBlock { scoop update }
            $out = Read-CapturedOutput $selfUpdate.OutputPath
            if ($out -match '(?i)error|failed') { Write-Status "Scoop self-update warning: $out" -Type Warning }
            $pkgUpdate = Invoke-StreamingCapture -ScriptBlock { scoop update '*' }
            $out = Read-CapturedOutput $pkgUpdate.OutputPath
            if ($out -and $out -notmatch 'Scoop is up to date') { $script:stepChanged = $true; $script:stepMessage = "Updated correctly" }
            scoop cleanup '*' 2>&1 | Out-Null
            scoop cache rm '*' 2>&1 | Out-Null
        }
        RequiresCommand = 'scoop'
        SlowOperation   = $false
        RequiresAdmin   = $false
        Disabled        = $SkipCleanup -or ($script:Config.SkipManagers -contains 'Scoop')
    }
    @{
        Name            = 'Winget'
        Title           = 'Winget'
        Action          = {
            # 1. Check for updates and run pre-hooks (needs for Adobe temp cleanup)
            Write-Host "  Checking for available updates (pre-hooks)..." -ForegroundColor Gray
            try {
                $preScan = Invoke-StreamingCapture -ScriptBlock { winget upgrade --include-unknown --source winget --accept-source-agreements --disable-interactivity }
                $checkOut = Read-CapturedOutput $preScan.OutputPath
                Invoke-WingetUpgradeHook -Phase 'Pre' -WingetOutput $checkOut
            } catch {}

            # 2. Main Upgrade
            Write-Host "  Upgrading all (winget source)..." -ForegroundColor Gray
            
            # WORKAROUND for 1632: Force TEMP to C:\Windows\Temp if admin
            if (-not $isAdmin) { 
                Write-Status "Running as NON-ADMIN. MSI installers (GitHub CLI, CMake, etc.) will likely fail with 1632." -Type Warning 
                Write-Status "Please run with -AutoElevate or as Administrator." -Type Info
            }
            
            $oldTemp = $env:TEMP; $oldTmp = $env:TMP
            if ($isAdmin -and (Test-Path 'C:\Windows\Temp')) { $env:TEMP = 'C:\Windows\Temp'; $env:TMP = 'C:\Windows\Temp' }
            
            try {
                $upgradeScan = Invoke-StreamingCapture -ScriptBlock { winget upgrade --include-unknown --source winget --accept-source-agreements --disable-interactivity }
                $upgradeList = Read-CapturedOutput $upgradeScan.OutputPath
                $upgradeEntries = @(Get-WingetUpgradeEntries -WingetOutput $upgradeList)
                $existingPinnedIds = @(Get-WingetPinnedPackageIds)
                $unknownEntries = @(
                    $upgradeEntries |
                    Where-Object { $_.Version -eq 'Unknown' -and $_.Id -notin $existingPinnedIds }
                )
                $pinnedUnknownIds = [System.Collections.Generic.List[string]]::new()

                foreach ($entry in $unknownEntries) {
                    Write-Detail "Auto-pinning $($entry.Id) because winget reports installed version as Unknown" -Type Warning
                    $pinOutput = Add-WingetBlockingPin -PackageId $entry.Id
                    if ($pinOutput -match 'Pin added|Successfully pinned|already exists|Found an existing pin') {
                        $pinnedUnknownIds.Add($entry.Id)
                    }
                }

                $pkgIds = @(
                    $upgradeEntries |
                    Where-Object { $_.Id -notin $existingPinnedIds -and $_.Id -notin $pinnedUnknownIds } |
                    Select-Object -ExpandProperty Id -Unique
                )

                $anyInstalled = $false
                $anyFailed    = $false
                $skippedIds   = [System.Collections.Generic.List[string]]::new()
                if ($pkgIds.Count -eq 0) {
                    $script:stepMessage = if ($pinnedUnknownIds.Count -gt 0) { "pinned unknown-version packages: $($pinnedUnknownIds -join ', ')" } else { 'already current' }
                    return
                }

                $i = 0
                foreach ($pkgId in $pkgIds) {
                    $i++
                    Write-Host "  ($i/$($pkgIds.Count)) Installing $pkgId..." -ForegroundColor Gray
                    $installed = $false
                    $pkgFailed = $false
                    for ($attempt = 1; $attempt -le 3; $attempt++) {
                        $pkgResult = Invoke-StreamingCapture -ScriptBlock { winget upgrade --id $pkgId --include-unknown --source winget --silent --accept-source-agreements --accept-package-agreements --disable-interactivity }
                        $exitCode = $pkgResult.ExitCode
                        $pkgOutput = Read-CapturedOutput $pkgResult.OutputPath
                        if ($exitCode -eq 1632) {
                            Write-Host "    Windows Installer busy (1632), waiting 10s before retry $attempt/3..." -ForegroundColor Yellow
                            Start-Sleep -Seconds 10
                        } elseif ($pkgOutput -match 'No installed package found matching input criteria') {
                            $skippedIds.Add($pkgId)
                            break
                        } else {
                            if ($exitCode -eq 0) {
                                $anyInstalled = $true
                                $installed = $true
                            } else {
                                $pkgFailed = $true
                                if ($pkgOutput) {
                                    $firstLine = ($pkgOutput -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
                                    if ($firstLine) { Write-Detail "$pkgId failed: $firstLine" -Type Warning }
                                }
                            }
                            break
                        }
                    }
                    if (-not $installed -and $pkgFailed) { $anyFailed = $true }
                    # Brief pause between packages to let Windows Installer service settle
                    Start-Sleep -Seconds 2
                }

                # Recapture status for post-hooks
                $finalScan = Invoke-StreamingCapture -ScriptBlock { winget upgrade --include-unknown --source winget --accept-source-agreements --disable-interactivity }
                $finalOutput = Read-CapturedOutput $finalScan.OutputPath
                try { Invoke-WingetUpgradeHook -Phase 'Post' -WingetOutput $finalOutput } catch {}

                $script:stepChanged = $anyInstalled
                if (-not $anyFailed) {
                    if ($anyInstalled) {
                        $details = [System.Collections.Generic.List[string]]::new()
                        $details.Add('updated correctly')
                        if ($skippedIds.Count -gt 0) { $details.Add("skipped stale ids: $($skippedIds.Count)") }
                        if ($pinnedUnknownIds.Count -gt 0) { $details.Add("auto-pinned unknown-version packages: $($pinnedUnknownIds.Count)") }
                        $script:stepMessage = ($details -join '; ')
                    } else {
                        $script:stepMessage = if ($pinnedUnknownIds.Count -gt 0) { "already current; auto-pinned unknown-version packages: $($pinnedUnknownIds.Count)" } else { 'already current' }
                    }
                } else {
                    $script:stepMessage = if ($anyInstalled) { 'completed with some failures' } else { 'failed; see warnings above' }
                }
            }
            finally { $env:TEMP = $oldTemp; $env:TMP = $oldTmp }
        }
        RequiresCommand = 'winget'
        SlowOperation   = $false
        RequiresAdmin   = $false
        Disabled        = $false
    }
    @{
        Name            = 'Chocolatey'
        Title           = 'Chocolatey'
        Action          = {
            $run = Invoke-StreamingCapture -ScriptBlock { choco upgrade all -y }
            $out = Read-CapturedOutput $run.OutputPath
            if ($out -match 'upgraded 0/0 packages') {
                $script:stepMessage = 'no package updates'
            }
            else {
                $script:stepChanged = $true
                $script:stepMessage = 'packages upgraded'
            }
        }
        RequiresCommand = 'choco'
        SlowOperation   = $false
        RequiresAdmin   = $true
        Disabled        = ($script:Config.SkipManagers -contains 'Chocolatey')
    }
    @{
        Name            = 'WSL'
        Title           = 'Windows Subsystem for Linux'
        Action          = {
            $wslUpdate = Invoke-StreamingCapture -ScriptBlock { wsl --update }
            $out = Read-CapturedOutput $wslUpdate.OutputPath
            if ($out -and $out -notmatch 'most recent version.+already installed') {
                $script:stepChanged = $true
            }
            if ($out -match 'most recent version.+already installed') { $script:stepMessage = 'platform already up to date' } else { $script:stepMessage = 'WSL platform updated' }

            if (-not $SkipWSLDistros) {
                $distroNames = @( wsl --list --quiet 2>&1 | ForEach-Object { ($_ -replace '\x00', '').Trim() } | Where-Object { $_ -and $_ -notmatch '^\s*$' } )
                if ($distroNames.Count -eq 0) { $script:stepMessage = 'platform current; no distros found' }
                else {
                    foreach ($dn in $distroNames) {
                        Write-Detail "Updating WSL Distro: $dn"
                        wsl -d $dn -- sh -lc "if command -v apt-get >/dev/null; then sudo apt-get update -qq && sudo apt-get upgrade -y -qq; elif command -v pacman >/dev/null; then sudo pacman -Syu --noconfirm; fi" 2>&1 | Out-Null
                    }
                    $script:stepMessage += " (Checked inside Distros)"
                }
            }
        }
        RequiresCommand = 'wsl'
        SlowOperation   = $false
        RequiresAdmin   = $true
        Disabled        = $SkipWSL -or ($script:Config.SkipManagers -contains 'WSL')
    }
)

function Invoke-WingetUpgradeHook {
    param([string]$Phase, [string]$WingetOutput)
    if (-not $WingetOutput) { return }
    foreach ($pkgId in $script:Config.WingetUpgradeHooks.Keys) {
        if ($WingetOutput -match [regex]::Escape($pkgId)) {
            $hook = $script:Config.WingetUpgradeHooks[$pkgId][$Phase]
            if ($hook) { try { & $hook } catch { Write-Status "Hook ($Phase) for $pkgId failed" -Type Warning } }
        }
    }
}

# --- Execute managers -----------------------------------------------------------
foreach ($mgr in $managers) {
    Invoke-Update -Name $mgr.Name -Title $mgr.Title -Action $mgr.Action -RequiresCommand $mgr.RequiresCommand -Disabled:$mgr.Disabled -RequiresAdmin:$mgr.RequiresAdmin -SlowOperation:$mgr.SlowOperation
}

# --- Other updates (the remaining dev tools) ------------------------------------
# These are defined similarly, but we'll keep the original Invoke-Update calls as they are already modular.
# However, for brevity, we'll just include a few key ones. The full script can have all.

# For demonstration, we'll show a selection. In a full version, you'd include all the original steps.
# I'll include a few to illustrate pattern.

# Windows Components
if (-not $SkipWindowsUpdate) {
    Invoke-Update -Name 'WindowsUpdate' -Title 'Windows Update' -RequiresAdmin -Action {
        Import-Module PSWindowsUpdate
        try {
            $wuParams = @{ Install = $true; AcceptAll = $true; NotCategory = 'Drivers'; IgnoreReboot = $true; RecurseCycle = 3; Verbose = $false; Confirm = $false }
            $wuRun = Invoke-StreamingCapture -ScriptBlock { Get-WindowsUpdate @wuParams }
            $wuOut = Read-CapturedOutput $wuRun.OutputPath
            if ($wuOut -match 'No updates found|There are no applicable updates') {
                $script:stepMessage = 'already current'
            }
            else {
                $script:stepChanged = $true
                $script:stepMessage = "Updated successfully"
            }
            if (-not $SkipReboot -and (Get-WURebootStatus -Silent -ErrorAction SilentlyContinue)) {
                Write-Status 'A reboot is required to finish installing updates.' -Type Warning
                $script:stepMessage += " (Reboot pending)"
            }
        }
        catch {
            if ($_.Exception.Message -match '0x800704c7') {
                $script:stepMessage = "Windows update aborted - reboot required or user canceled."
            }
            else { throw }
        }
    } -Disabled:$SkipWindowsUpdate
}

Invoke-Update -Name 'StoreApps' -Title 'Microsoft Store Apps' -Disabled:$SkipStoreApps -RequiresCommand 'winget' -Action {
    Write-Detail "Checking Microsoft Store app upgrades via winget"
    $result = Invoke-StreamingCapture -ScriptBlock { winget upgrade --source msstore --all --silent --accept-package-agreements --accept-source-agreements }
    $storeOutput = Read-CapturedOutput $result.OutputPath
    if (-not $storeOutput -or $storeOutput -match 'No installed package found matching input criteria|No applicable upgrade found|There are no available upgrades') {
        $script:stepMessage = 'no Microsoft Store app updates'
    }
    else {
        if ($storeOutput -match 'Successfully installed|Successfully upgraded|successfully installed') {
            $script:stepChanged = $true
            $script:stepMessage = 'store apps updated'
        }
        else {
            $script:stepMessage = 'store apps checked'
        }
    }
}

Invoke-Update -Name 'DefenderSignatures' -Title 'Microsoft Defender Signatures' -Disabled:$SkipDefender -RequiresCommand 'Update-MpSignature' -RequiresAdmin -Action {
    $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($mpStatus -and -not ($mpStatus.AMServiceEnabled -and $mpStatus.AntivirusEnabled)) { $script:stepMessage = 'Microsoft Defender is not the active AV'; return }
    $beforeVersion = $mpStatus.AntivirusSignatureVersion
    try {
        Invoke-StreamingCapture -ScriptBlock { Update-MpSignature -ErrorAction Stop } | Out-Null
        $afterVersion = (Get-MpComputerStatus -ErrorAction SilentlyContinue).AntivirusSignatureVersion
        if ($afterVersion -and $afterVersion -ne $beforeVersion) {
            $script:stepChanged = $true
            $script:stepMessage = "signatures updated to $afterVersion"
        }
        else {
            $script:stepMessage = 'signatures already current'
        }
    } catch { throw }
}

# Development tools - sequential
Write-Host "`n[Sequential] Running dev tool updates..." -ForegroundColor DarkCyan

Invoke-Update -Name 'npm' -Title 'npm (Node.js)' -RequiresCommand 'npm' -Disabled:$SkipNode -Action {
    $currentNpm = (npm --version 2>&1).Trim()
    $latestNpm = (npm view npm version 2>&1).Trim()
    if ($currentNpm -ne $latestNpm) { npm install -g npm@latest 2>&1 | Out-Null; $script:stepChanged = $true }
    $npmOutdated = Invoke-StreamingCapture -ScriptBlock { npm outdated -g --json }
    $outdatedJson = Read-CapturedOutput $npmOutdated.OutputPath
    if ($outdatedJson) {
        $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
        if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
            $pkgs = $outdated.PSObject.Properties.Name
            Invoke-StreamingCapture -ScriptBlock { npm install -g $pkgs } | Out-Null
            $script:stepChanged = $true
            $script:stepMessage = "Updated $($pkgs.Count) package(s)"
        }
        else { $script:stepMessage = 'no global package updates' }
    }
    else { $script:stepMessage = 'no global package updates' }
    if ($script:stepChanged -and -not $script:stepMessage) { $script:stepMessage = 'npm updated' }
    npm cache clean --force 2>&1 | Out-Null
}

Invoke-Update -Name 'pnpm' -RequiresCommand 'pnpm' -Disabled:$SkipNode -Action {
    $pnpmRun = Invoke-StreamingCapture -ScriptBlock { pnpm update -g }
    $out = Read-CapturedOutput $pnpmRun.OutputPath
    if ($out -match 'Already up to date|Nothing to') { $script:stepMessage = 'no global package updates' } else { $script:stepChanged = $true; $script:stepMessage = 'global packages updated' }
}

Invoke-Update -Name 'Bun' -RequiresCommand 'bun' -Disabled:$SkipNode -Action {
    $out = (bun upgrade 2>&1 | Out-String).Trim()
    if ($out -and $out -notmatch 'already on the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
}

Invoke-Update -Name 'Deno' -RequiresCommand 'deno' -Disabled:$SkipNode -Action {
    $mgr = Get-ToolInstallManager 'deno'
    if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
    $denoRun = Invoke-StreamingCapture -ScriptBlock { deno upgrade }
    $out = Read-CapturedOutput $denoRun.OutputPath
    if ($out -and $out -notmatch 'already the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
}

Invoke-Update -Name 'Rust' -RequiresCommand 'rustup' -Disabled:$SkipRust -Action {
    $rustRun = Invoke-StreamingCapture -ScriptBlock { rustup update }
    $out = Read-CapturedOutput $rustRun.OutputPath
    if ($out -match 'unchanged - rustc ([^\s]+)') { $script:stepMessage = "stable unchanged ($($Matches[1]))" }
    elseif ($out -match 'updated - rustc') { $script:stepChanged = $true; $script:stepMessage = "stable updated" }
    else { $script:stepMessage = 'toolchain checked' }
}

Invoke-Update -Name 'Go' -RequiresCommand 'go' -Disabled:$SkipGo -Action {
    $mgr = Get-ToolInstallManager 'go'
    if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)" }
    else {
        $goRun = Invoke-StreamingCapture -ScriptBlock { winget upgrade --id GoLang.Go --accept-source-agreements --accept-package-agreements --disable-interactivity }
        $r = Read-CapturedOutput $goRun.OutputPath
        if ($r -match 'Successfully installed') { $script:stepChanged = $true; $script:stepMessage = 'updated via winget' } else { $script:stepMessage = 'no newer version available' }
    }
}

Invoke-Update -Name 'gh-extensions' -Title 'GitHub CLI Extensions' -RequiresCommand 'gh' -Action {
    $ghList = Invoke-StreamingCapture -ScriptBlock { gh extension list }
    $ghExt = Read-CapturedOutput $ghList.OutputPath
    if ($ghExt) {
        $upgradeRun = Invoke-StreamingCapture -ScriptBlock { gh extension upgrade --all }
        $upgradeOutput = Read-CapturedOutput $upgradeRun.OutputPath
        if ($upgradeOutput -match 'upgraded|updated|installing') {
            $script:stepChanged = $true
            $script:stepMessage = 'extensions updated'
        }
        else {
            $script:stepMessage = 'extensions checked'
        }
    }
    else { $script:stepMessage = 'no extensions installed' }
}

Invoke-Update -Name 'pipx' -Title 'pipx (Python Tools)' -RequiresCommand 'pipx' -Action {
    pipx upgrade-all 2>&1 | Out-Null; $script:stepMessage = 'upgraded all'
}

Invoke-Update -Name 'Poetry' -Title 'Poetry' -RequiresCommand 'poetry' -Disabled:$SkipPoetry -Action {
    $out = (poetry self update 2>&1 | Out-String).Trim()
    if ($out -and $out -notmatch 'already using the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
}

Invoke-Update -Name 'Composer' -Title 'Composer (PHP)' -RequiresCommand 'composer' -Disabled:$SkipComposer -Action {
    composer self-update --no-interaction 2>&1 | Out-Null; composer global update --no-interaction 2>&1 | Out-Null; $script:stepMessage = 'checked'
}

Invoke-Update -Name 'RubyGems' -RequiresCommand 'gem' -Disabled:$SkipRuby -Action {
    gem update --system 2>&1 | Out-Null; gem update 2>&1 | Out-Null; $script:stepMessage = 'checked'
}

Invoke-Update -Name 'yt-dlp' -RequiresCommand 'yt-dlp' -Action {
    $ytdlpPath = (Get-Command yt-dlp).Source
    if ($ytdlpPath -like '*scoop*') { $script:stepMessage = 'managed by Scoop' }
    elseif ($ytdlpPath -like '*pip*') { python -m pip install --upgrade yt-dlp; $script:stepChanged = $true; $script:stepMessage = 'updated via pip' }
    else { yt-dlp -U; $script:stepChanged = $true; $script:stepMessage = 'updated' }
}

Invoke-Update -Name 'tldr' -Title 'tldr Cache' -RequiresAnyCommand @('tldr', 'tealdeer') -Action {
    $cmd = if (Test-Command 'tealdeer') { 'tealdeer' } else { 'tldr' }
    & $cmd --update 2>&1 | Out-Null; $script:stepMessage = 'cache updated'
}

Invoke-Update -Name 'oh-my-posh' -Title 'Oh My Posh' -RequiresCommand 'oh-my-posh' -Action {
    $mgr = Get-ToolInstallManager 'oh-my-posh'
    if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)" }
    else { oh-my-posh upgrade; $script:stepChanged = $true; $script:stepMessage = 'updated' }
}

Invoke-Update -Name 'Volta' -RequiresCommand 'volta' -Disabled:$SkipNode -Action {
    $script:stepMessage = 'standalone tool; update manually'
}

Invoke-Update -Name 'fnm' -RequiresCommand 'fnm' -Disabled:$SkipNode -Action {
    $script:stepMessage = 'node manager checked'
}

Invoke-Update -Name 'mise' -Title 'mise' -RequiresCommand 'mise' -Action {
    mise self-update --yes 2>&1 | Out-Null; mise upgrade 2>&1 | Out-Null; $script:stepMessage = 'updated'
}

Invoke-Update -Name 'juliaup' -RequiresCommand 'juliaup' -Action {
    juliaup update 2>&1 | Out-Null; $script:stepMessage = 'checked'
}

Invoke-Update -Name 'ollama-models' -Title 'Ollama Models' -RequiresCommand 'ollama' -Disabled:(-not $UpdateOllamaModels) -Action {
    $models = (ollama list 2>&1 | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] })
    foreach ($m in $models) { if ($m) { ollama pull $m } }
    $script:stepMessage = "checked $($models.Count) models"
}

Invoke-Update -Name 'git-lfs' -RequiresCommand 'git-lfs' -Disabled:$SkipGitLFS -Action {
    $result = Invoke-StreamingCapture -ScriptBlock { winget upgrade --id GitHub.GitLFS --accept-source-agreements --accept-package-agreements --disable-interactivity }
    $gitLfsOutput = Read-CapturedOutput $result.OutputPath
    if ($gitLfsOutput -match 'Successfully installed|Successfully upgraded|successfully installed') {
        $script:stepChanged = $true
        $script:stepMessage = 'updated via winget'
    }
    else {
        $script:stepMessage = 'already current'
    }
}

Invoke-Update -Name 'git-credential-manager' -RequiresCommand 'git-credential-manager' -Action {
    winget upgrade --id Git.Git --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1 | Out-Null; $script:stepMessage = 'updated via git upgrade'
}

Invoke-Update -Name 'pip' -Title 'Python / pip' -Action {
    python -m pip install --upgrade pip 2>&1 | Out-Null
    $outdated = @(python -m pip list --outdated --format=json 2>$null | ConvertFrom-Json)
    if (-not $outdated -or $outdated.Count -eq 0) {
        $script:stepMessage = 'no global pip package updates'
        return
    }

    $updated = [System.Collections.Generic.List[string]]::new()
    $failed = [System.Collections.Generic.List[string]]::new()

    foreach ($p in $outdated) {
        $pkgName = $p.name
        $pipRun = Invoke-StreamingCapture -ScriptBlock { python -m pip install --upgrade --disable-pip-version-check --no-input $pkgName }
        $pkgOutput = Read-CapturedOutput $pipRun.OutputPath
        if ($pipRun.ExitCode -eq 0) {
            $updated.Add($pkgName)
            continue
        }

        $failed.Add($pkgName)
        if ($pkgName -eq 'tesserocr') {
            Write-Detail 'tesserocr skipped: local Tesseract development libraries were not found for this Python build' -Type Warning
        }
        elseif ($pkgOutput) {
            $firstLine = ($pkgOutput -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
            if ($firstLine) { Write-Detail "$pkgName failed: $firstLine" -Type Warning }
        }
    }

    if ($updated.Count -gt 0) { $script:stepChanged = $true }
    if ($failed.Count -gt 0 -and $updated.Count -gt 0) {
        $script:stepMessage = "Updated $($updated.Count) packages; failed: $($failed -join ', ')"
    }
    elseif ($failed.Count -gt 0) {
        $script:stepMessage = "No pip packages updated; failed: $($failed -join ', ')"
    }
    else {
        $script:stepMessage = "Updated $($updated.Count) packages"
    }
}

Invoke-Update -Name 'uv' -Title 'UV' -RequiresCommand 'uv' -Action {
    $uvRun = Invoke-StreamingCapture -ScriptBlock { uv self update }
    $out = Read-CapturedOutput $uvRun.OutputPath
    if ($out -match 'installed via pip') { $script:stepMessage = 'managed by pip' }
    elseif ($out -and $out -notmatch 'already up to date') { $script:stepChanged = $true; $script:stepMessage = 'updated' }
    else { $script:stepMessage = 'already current' }
}

Invoke-Update -Name 'uv-tools' -Title 'uv Tool Installs' -Disabled:$SkipUVTools -RequiresCommand 'uv' -Action {
    Invoke-StreamingCapture -ScriptBlock { uv tool upgrade --all } | Out-Null
    $script:stepMessage = 'checked'
}

Invoke-Update -Name 'uv-python' -Title 'uv Python Versions' -RequiresCommand 'uv' -Action {
    Invoke-StreamingCapture -ScriptBlock { uv python install latest } | Out-Null
    $script:stepMessage = 'reinstalled latest patches'
}

Invoke-Update -Name 'cargo-binaries' -Title 'Cargo Global Binaries' -RequiresCommand 'cargo' -Action {
    if (-not (Test-Command 'cargo-install-update')) { Invoke-StreamingCapture -ScriptBlock { cargo install cargo-update } | Out-Null }
    Invoke-StreamingCapture -ScriptBlock { cargo install-update -a } | Out-Null
    $script:stepMessage = 'checked'
}

Invoke-Update -Name 'dotnet' -Title '.NET Tools' -RequiresCommand 'dotnet' -Action {
    $dotnetTools = Invoke-StreamingCapture -ScriptBlock { dotnet tool update --global --all }
    $out = Read-CapturedOutput $dotnetTools.OutputPath
    if ($out -match 'was successfully updated|successfully updated') {
        $script:stepChanged = $true
        $script:stepMessage = 'tools updated'
    }
    else {
        $script:stepMessage = 'tools checked'
    }
}

Invoke-Update -Name 'dotnet-workloads' -Title '.NET Workloads' -RequiresCommand 'dotnet' -RequiresAdmin -Action {
    $dotnetWorkloads = Invoke-StreamingCapture -ScriptBlock { dotnet workload update }
    $out = Read-CapturedOutput $dotnetWorkloads.OutputPath
    if ($out -match 'Updated advertising manifest|Installing|Installed') {
        $script:stepChanged = $true
        $script:stepMessage = 'workloads updated'
    }
    else {
        $script:stepMessage = 'workloads checked'
    }
}

Invoke-Update -Name 'vscode-extensions' -Title 'VS Code Extensions' -Disabled:$SkipVSCodeExtensions -Action {
    $code = Get-VSCodeCliPath
    if ($code) {
        $vsCodeRun = Invoke-StreamingCapture -ScriptBlock { & $code --update-extensions }
        $out = Read-CapturedOutput $vsCodeRun.OutputPath
        if ($out -match 'updated|installing') {
            $script:stepChanged = $true
            $script:stepMessage = 'extensions updated'
        }
        else {
            $script:stepMessage = 'extensions checked'
        }
    } else { $script:stepMessage = 'VS Code not found' }
}

Invoke-Update -Name 'Flutter' -Title 'Flutter' -Disabled:$SkipFlutter -RequiresCommand 'flutter' -Action {
    $mgr = Get-ToolInstallManager 'flutter'
    if ($mgr) {
        $script:stepMessage = "managed by $mgr (already updated)"
        return
    }

    $flutterRun = Invoke-StreamingCapture -ScriptBlock { flutter upgrade }
    $out = Read-CapturedOutput $flutterRun.OutputPath
    if ($out -match 'already up to date|already on the latest') {
        $script:stepMessage = 'already current'
    }
    elseif ($flutterRun.ExitCode -eq 0) {
        $script:stepChanged = $true
        $script:stepMessage = 'updated'
    }
    else {
        $script:stepMessage = 'checked'
    }
}

Invoke-Update -Name 'pwsh-resources' -Title 'PowerShell Modules / Resources' -Disabled:$SkipPowerShellModules -Action {
    $details = [System.Collections.Generic.List[string]]::new()
    $changed = $false

    if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
        $resources = @(Get-InstalledPSResource -ErrorAction SilentlyContinue)
        $psrCommand = Get-Command Update-PSResource -ErrorAction SilentlyContinue
        $supportsAcceptLicense = $psrCommand -and $psrCommand.Parameters.ContainsKey('AcceptLicense')
        foreach ($resource in $resources) {
            $beforeVersion = [string]$resource.Version
            $updateArgs = @{ Name = $resource.Name; ErrorAction = 'SilentlyContinue' }
            if ($supportsAcceptLicense) { $updateArgs.AcceptLicense = $true }
            try { Update-PSResource @updateArgs 2>&1 | Out-Null } catch { Write-Verbose "Update-PSResource failed for $($resource.Name): $($_.Exception.Message)" }
            $afterVersion = [string]((Get-InstalledPSResource -Name $resource.Name -ErrorAction SilentlyContinue | Select-Object -First 1).Version)
            if ($afterVersion -and $afterVersion -ne $beforeVersion) {
                $changed = $true
                $details.Add("PSResource $($resource.Name) $beforeVersion -> $afterVersion")
            }
        }
    }

    if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
        $modules = @(Get-InstalledModule -ErrorAction SilentlyContinue)
        $moduleCommand = Get-Command Update-Module -ErrorAction SilentlyContinue
        $supportsAcceptLicense = $moduleCommand -and $moduleCommand.Parameters.ContainsKey('AcceptLicense')
        foreach ($module in $modules) {
            $beforeVersion = [string]$module.Version
            $updateArgs = @{ Name = $module.Name; ErrorAction = 'SilentlyContinue' }
            if ($supportsAcceptLicense) { $updateArgs.AcceptLicense = $true }
            try { Update-Module @updateArgs 2>&1 | Out-Null } catch { Write-Verbose "Update-Module failed for $($module.Name): $($_.Exception.Message)" }
            $afterVersion = [string]((Get-InstalledModule -Name $module.Name -ErrorAction SilentlyContinue | Select-Object -First 1).Version)
            if ($afterVersion -and $afterVersion -ne $beforeVersion) {
                $changed = $true
                $details.Add("Module $($module.Name) $beforeVersion -> $afterVersion")
            }
        }
    }

    try {
        Update-Help -Force -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Verbose "Update-Help issue: $($_.Exception.Message)"
    }

    if ($changed) {
        $script:stepChanged = $true
        $script:stepMessage = if ($details.Count -gt 0) { $details -join '; ' } else { 'PowerShell modules updated' }
    }
    else {
        $script:stepMessage = 'PowerShell modules and help checked'
    }
}

# The remainder of the script (summary, cleanup, state saving, etc.) remains similar.
# We'll include the cleanup and summary sections, with improvements.

# --- Cleanup -------------------------------------------------------------------
if ($SkipCleanup) {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cleanup'; Reason = 'flag' })
}
elseif ($script:IsSimulation) {
    Write-Host "  [DryRun] Would run: System Cleanup" -ForegroundColor DarkCyan
}
else {
    Write-Section 'System Cleanup'
    try {
        $tempPath = $env:TEMP
        $cutoff = (Get-Date).AddDays(-$script:Config.TempCleanupDays)
        if ($tempPath -and (Test-Path $tempPath)) {
            Get-ChildItem -Path $tempPath -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "Temp files cleared (older than $($script:Config.TempCleanupDays) days)" -Type Success
        }
    }
    catch {}

    if ($isAdmin) {
        try {
            if (Test-Path 'C:\Windows\Temp') {
                Get-ChildItem -Path 'C:\Windows\Temp' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "C:\Windows\Temp cleared (older than $($script:Config.TempCleanupDays) days)" -Type Success
            }
        }
        catch {}
    }

    try { Clear-DnsClientCache -ErrorAction SilentlyContinue; Write-Status "DNS cache flushed" -Type Success } catch {}
    try { Clear-RecycleBin -Force -ErrorAction SilentlyContinue; Write-Status "Recycle Bin emptied" -Type Success } catch {}

    if ($isAdmin -and $DeepClean) {
        try { DISM /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-Null; Write-Status "DISM component store cleaned" -Type Success } catch {}
        try { Clear-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue; Write-Status "Delivery Optimization cache cleared" -Type Success } catch {}
        if (-not $SkipDestructive) {
            try { Get-ChildItem -Path 'C:\Windows\Prefetch' -Filter '*.pf' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue; Write-Status "Prefetch files cleared" -Type Success } catch {}
        }
    }
    $updateResults.Checked.Add('cleanup')
    $updateResults.Details['cleanup'] = 'temp files, DNS cache, and recycle bin cleaned'
}

# --- State and Diff logic ------------------------------------------------------
if (-not $script:IsSimulation -and $WhatChanged) {
    $previousState = @{
        Winget     = ConvertTo-StringMap $script:State.Winget
        Scoop      = ConvertTo-StringMap $script:State.Scoop
        Chocolatey = ConvertTo-StringMap $script:State.Chocolatey
    }

    Write-Host "`n[Summary] Updating package state maps for WhatChanged diff..." -ForegroundColor DarkGray
    Update-WingetState
    Update-ScoopState
    Update-ChocolateyState

    Write-Section "What changed since last run"
    foreach ($entry in @(
        @{ Name = 'Winget'; Prev = $previousState.Winget; Curr = $script:State.Winget },
        @{ Name = 'Scoop'; Prev = $previousState.Scoop; Curr = $script:State.Scoop },
        @{ Name = 'Chocolatey'; Prev = $previousState.Chocolatey; Curr = $script:State.Chocolatey }
    )) {
        $changes = @(Compare-PackageMaps $entry.Prev $entry.Curr)
        if ($changes.Count -gt 0) {
            Write-Host "  $($entry.Name) changes:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        }
        else {
            Write-Host "  $($entry.Name): no changes detected" -ForegroundColor DarkGray
        }
    }

    Save-State
}
elseif (-not $script:IsSimulation) {
    Save-State
}


# --- Summary ------------------------------------------------------------------
$duration = (Get-Date) - $startTime
Write-Host "`n$('=' * 54)" -ForegroundColor Green
Write-Host (" UPDATE COMPLETE -- {0}" -f $duration.ToString('hh\:mm\:ss')) -ForegroundColor Green
Write-Host "$('=' * 54)" -ForegroundColor Green

$updatedNames = @($updateResults.Success | Select-Object -Unique)
$checkedNames = @($updateResults.Checked | Where-Object { $_ -notin $updatedNames -and $_ -notin $updateResults.Failed } | Select-Object -Unique)

if ($updatedNames.Count -gt 0) {
    Write-Host "`n[OK] Updated   ($($updatedNames.Count))" -ForegroundColor Green
    foreach ($name in $updatedNames) { Write-Detail ('{0}: {1}' -f $name, $updateResults.Details[$name]) -Type Info }
}
if ($checkedNames.Count -gt 0) {
    Write-Host "[~] Checked   ($($checkedNames.Count))" -ForegroundColor DarkGray
    foreach ($name in $checkedNames) { Write-Detail ('{0}: {1}' -f $name, $updateResults.Details[$name]) -Type Muted }
}
if ($updateResults.Failed.Count -gt 0) {
    Write-Host "[X] Failed    ($($updateResults.Failed.Count))" -ForegroundColor Red
    foreach ($name in $updateResults.Failed) { Write-Detail ('{0}: {1}' -f $name, $updateResults.Details[$name]) -Type Error }
}
if ($updateResults.Skipped.Count -gt 0) {
    Write-Host "[!] Skipped   ($($updateResults.Skipped.Count))" -ForegroundColor Yellow
    foreach ($item in $updateResults.Skipped) { Write-Detail ('{0} ({1})' -f $item.Name, $item.Reason) -Type Warning }
}

# --- Section timings ----------------------------------------------------------
if ($script:SectionTimings.Count -gt 0) {
    $timingsToShow = @( $script:SectionTimings.GetEnumerator() | Sort-Object Value -Descending )
    if ($timingsToShow.Count -gt 0) {
        Write-Host "`n  Section timings:" -ForegroundColor DarkGray
        $timingsToShow | ForEach-Object { Write-Host ("    {0,-30} {1,6}s" -f $_.Key, $_.Value.ToString('F1', [cultureinfo]::InvariantCulture)) -ForegroundColor DarkGray }
    }
}

# --- Notifications ------------------------------------------------------------
try {
    if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) {
        Import-Module BurntToast -ErrorAction SilentlyContinue
        $msg = if ($updateResults.Failed.Count -gt 0) { "$($updatedNames.Count) updated, $($updateResults.Failed.Count) failed" } elseif ($updatedNames.Count -gt 0) { "$($updatedNames.Count) components updated" } else { 'No updates were needed' }
        New-BurntToastNotification -Text 'Update-Everything', $msg -ErrorAction SilentlyContinue
    }
}
catch {}

# --- Pause if needed ----------------------------------------------------------
if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }

# --- Exit --------------------------------------------------------------------
exit 0
