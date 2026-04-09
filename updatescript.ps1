<#
.SYNOPSIS
    System-wide update script for Windows
.DESCRIPTION
    Updates all package managers, system components, and development tools.
    Version 5.0.0: Parallel execution for system components and dev tools.
.VERSION
    5.0.0
.NOTES
    Run as Administrator for full functionality. PowerShell 7 is preferred,
    but the script can fall back to Windows PowerShell 5.1-compatible paths.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -UltraFast
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
    .\updatescript.ps1 -SkipCleanup -SkipWindowsUpdate
    .\updatescript.ps1 -SkipNode -SkipRust -SkipGo
    .\updatescript.ps1 -DeepClean
    .\updatescript.ps1 -UpdateOllamaModels
    .\updatescript.ps1 -WhatChanged
    .\updatescript.ps1 -DryRun
#>

[CmdletBinding(DefaultParameterSetName = 'Normal', SupportsShouldProcess = $true)]
param(
    [Parameter(ParameterSetName = 'Normal')]
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,
    [switch]$UltraFast,
    [switch]$NoElevate,
    [switch]$AutoElevate,
    [switch]$NoPause,
    [switch]$SkipWSL,
    [switch]$SkipWSLDistros,
    [switch]$SkipDefender,
    [switch]$SkipStoreApps,
    [switch]$SkipUVTools,
    [switch]$SkipVSCodeExtensions,
    [switch]$SkipPoetry,
    [switch]$SkipComposer,
    [switch]$SkipRuby,
    [switch]$SkipPowerShellModules,
    [switch]$SkipCleanup,
    [int]$WingetTimeoutSec = 300,
    [Parameter(ParameterSetName = 'Schedule')]
    [switch]$Schedule,
    [Parameter(ParameterSetName = 'Schedule')]
    [ValidateScript({ $_ -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$' })]
    [string]$ScheduleTime = '03:00',
    [string]$LogPath,
    [switch]$SkipNode,
    [switch]$SkipRust,
    [switch]$SkipGo,
    [switch]$SkipFlutter,
    [switch]$SkipGitLFS,
    [switch]$DeepClean,
    [switch]$UpdateOllamaModels,
    [switch]$WhatChanged,
    [switch]$DryRun,
    [ValidateRange(1, 10)]
    [int]$ParallelThrottle = 4
)

# --- Encoding & globals --------------------------------------------------------
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding  = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding           = $utf8NoBom
} catch { Write-Verbose "Console encoding update skipped: $($_.Exception.Message)" }

$ErrorActionPreference = 'Continue'
$isAdmin    = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')
$startTime  = Get-Date
$script:IsSimulation = $DryRun -or $WhatIfPreference
$updateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Checked = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[pscustomobject]]::new()
    Details = [ordered]@{}
}

# --- Config --------------------------------------------------------------------
$script:Config = @{
    WingetTimeoutSec = 300
    StateDir         = Join-Path $env:LOCALAPPDATA 'Update-Everything'
    LogRetentionDays = 7
    LogMaxSizeMB     = 10
    TempCleanupDays  = 7
    PipAllowPackages = @()
    PipSkipPackages  = @('tesserocr')
    SkipManagers     = @()
    FastModeSkip     = @('Chocolatey','WSLDistros','npm','pnpm','bun','deno','rust','cargo-binaries',
                         'go','gh-extensions','pipx','poetry','composer','rubygems','yt-dlp','tldr',
                         'oh-my-posh','volta','fnm','mise','juliaup','ollama-models','git-lfs',
                         'git-credential-manager','pip','uv','uv-tools','uv-python','dotnet',
                         'dotnet-workloads','vscode-extensions','pwsh-resources')
    UltraFastSkip    = @('StoreApps','cleanup','WSL','WindowsUpdate','DefenderSignatures')
    WingetUpgradeHooks = @{
        'Spotify.Spotify' = @{
            Pre  = { $script:_spotifyWasRunning = [bool](Get-Process -Name Spotify -EA SilentlyContinue); Stop-Process -Name Spotify -Force -EA SilentlyContinue; Start-Sleep 2 }
            Post = { if ($script:_spotifyWasRunning) { Start-Process "$env:APPDATA\Spotify\Spotify.exe" -EA SilentlyContinue } }
        }
        'Google.Chrome' = @{
            Pre  = {
                $script:_chromeWasRunning = [bool](Get-Process -Name chrome -EA SilentlyContinue)
                if ($script:_chromeWasRunning) { Write-Host '  Closing Chrome...' -ForegroundColor Gray }
                # Kill Chrome and ALL Google background processes — any of these can cause 1603
                Stop-Process -Name 'chrome','chrome_crashpad_handler','GoogleUpdate','GoogleUpdateSetup','GoogleCrashHandler','GoogleCrashHandler64','GoogleUpdateComRegisterShell64' -Force -EA SilentlyContinue
                Start-Sleep 3
            }
            Post = { if ($script:_chromeWasRunning) { $p = @((Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),(Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'Google.QuickShare' = @{
            Pre  = { Stop-Process -Name 'QuickShare','QuickShareAgent','NearbyShare' -Force -EA SilentlyContinue; Start-Sleep 2 }
            Post = {}
        }
        'Microsoft.VisualStudioCode' = @{
            Pre  = { $script:_vscodeWasRunning = [bool](Get-Process -Name Code -EA SilentlyContinue); if ($script:_vscodeWasRunning) { Write-Host '  Closing VS Code...' -ForegroundColor Gray; Stop-Process -Name Code -Force -EA SilentlyContinue; Start-Sleep 3 } }
            Post = { if ($script:_vscodeWasRunning) { $p = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'; if (Test-Path $p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'Foxit.FoxitReader' = @{
            Pre  = { $script:_foxitWasRunning = [bool](Get-Process -Name 'FoxitPDFReader','FoxitReader' -EA SilentlyContinue); if ($script:_foxitWasRunning) { Write-Host '  Closing Foxit Reader...' -ForegroundColor Gray; Stop-Process -Name 'FoxitPDFReader','FoxitReader' -Force -EA SilentlyContinue; Start-Sleep 2 } }
            Post = { if ($script:_foxitWasRunning) { $p = @((Join-Path $env:ProgramFiles 'Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe'),(Join-Path ${env:ProgramFiles(x86)} 'Foxit Software\Foxit Reader\FoxitReader.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'GoLang.Go' = @{
            Pre  = { Stop-Process -Name 'go','gopls','dlv','golangci-lint' -Force -EA SilentlyContinue; Start-Sleep 2 }
            Post = {}
        }
        'Adobe.Acrobat.Reader.64-bit' = @{
            Pre  = {
                $script:_acrobatProcs = @(Get-Process -Name 'Acrobat','AcroRd32','AcroCEF' -EA SilentlyContinue)
                if ($script:_acrobatProcs.Count -gt 0) { Write-Host '  Closing Adobe Acrobat...' -ForegroundColor Gray; $script:_acrobatProcs | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 3 }
                Write-Host '  Clearing temp files to prevent Acrobat extraction errors...' -ForegroundColor Gray
                Get-ChildItem -Path $env:TEMP -File -Force -EA SilentlyContinue | Where-Object { $_.Extension -match '\.tmp|\.log' } | Remove-Item -Force -EA SilentlyContinue
            }
            Post = { if ($script:_acrobatProcs.Count -gt 0) { $p = @((Join-Path $env:ProgramFiles 'Adobe\Acrobat DC\Acrobat\Acrobat.exe'),(Join-Path ${env:ProgramFiles(x86)} 'Adobe\Acrobat Reader DC\Reader\AcroRd32.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
    }
}

$configFile = Join-Path $PSScriptRoot 'update-config.json'
if (Test-Path $configFile) {
    try {
        $userConfig = Get-Content $configFile -Raw | ConvertFrom-Json
        foreach ($prop in $userConfig.PSObject.Properties) { $script:Config[$prop.Name] = $prop.Value }
        Write-Verbose "Loaded config from $configFile"
    } catch { Write-Warning "Failed to load config file: $_" }
}

# --- Elevation -----------------------------------------------------------------
if (-not $isAdmin -and $AutoElevate -and -not $NoElevate) {
    Write-Host 'Elevating to Administrator...' -ForegroundColor Cyan
    $forwardedArgs = foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -eq 'AutoElevate') { continue }
        if ($entry.Value -is [switch]) { if ($entry.Value.IsPresent) { "-$($entry.Key)" } }
        elseif ($null -ne $entry.Value -and "$($entry.Value)".Length -gt 0) { "-$($entry.Key)"; [string]$entry.Value }
    }
    $pwshCommand = Get-Command pwsh.exe -EA SilentlyContinue
    if ($pwshCommand -and $pwshCommand.Source) { $shell = $pwshCommand.Source } else { $shell = 'powershell.exe' }
    try { Start-Process -FilePath $shell -Verb RunAs -ArgumentList (@('-NoProfile','-ExecutionPolicy','Bypass','-File',$PSCommandPath) + @($forwardedArgs)) -Wait; exit }
    catch { Write-Warning 'Could not elevate. Continuing without Administrator privileges.' }
} elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host 'INFO: Running without elevation. Admin-only tasks may be skipped.' -ForegroundColor DarkYellow
}

if ($Schedule) {
    if (-not $isAdmin -and -not $script:IsSimulation) { throw 'Scheduled task registration requires Administrator.' }
    $pwshCommand = Get-Command pwsh.exe -EA SilentlyContinue
    if ($pwshCommand -and $pwshCommand.Source) { $shell = $pwshCommand.Source } else { $shell = 'powershell.exe' }
    $taskArgs = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-SkipReboot','-NoPause')
    if ($script:IsSimulation) { Write-Host "  [DryRun] Would register 'DailySystemUpdate' at $ScheduleTime" -ForegroundColor DarkCyan }
    elseif ($PSCmdlet.ShouldProcess('DailySystemUpdate', "Register scheduled task for $ScheduleTime")) {
        Register-ScheduledTask -TaskName 'DailySystemUpdate' -Force `
            -Action  (New-ScheduledTaskAction  -Execute $shell -Argument ($taskArgs -join ' ')) `
            -Trigger (New-ScheduledTaskTrigger -Daily -At $ScheduleTime) `
            -Settings(New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable) `
            -RunLevel Highest | Out-Null
        Write-Host "[OK] Scheduled 'DailySystemUpdate' daily at $ScheduleTime." -ForegroundColor Green
    }
    exit
}

if ($WingetTimeoutSec -ne 300) { $script:Config.WingetTimeoutSec = $WingetTimeoutSec }
if ($script:Config.WingetTimeoutSec -lt 30) { $script:Config.WingetTimeoutSec = 30 }
if ($SkipUVTools) { $script:Config.SkipUVTools = $true }
if ($UltraFast) { $FastMode = $true }
if (-not $PSBoundParameters.ContainsKey('ParallelThrottle')) {
    $ParallelThrottle = [Math]::Max(4, [Math]::Min([Environment]::ProcessorCount, 8))
}

# --- Helper Functions ----------------------------------------------------------
$commandCache = @{}
function Test-Command([string]$Command) {
    if ($commandCache.ContainsKey($Command)) { return $commandCache[$Command] }
    $saved = $WhatIfPreference; $WhatIfPreference = $false
    try { $result = [bool](Get-Command $Command -EA SilentlyContinue) }
    finally { $WhatIfPreference = $saved }
    $commandCache[$Command] = $result
    return $result
}

function ConvertTo-StringMap {
    param([AllowNull()]$InputObject)
    $map = @{}
    if ($null -eq $InputObject) { return $map }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($key in $InputObject.Keys) { $map[[string]$key] = [string]$InputObject[$key] }
        return $map
    }
    foreach ($prop in $InputObject.PSObject.Properties) { $map[[string]$prop.Name] = [string]$prop.Value }
    return $map
}

function Normalize-PackageName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return (($Name -replace '[-_.]+', '-').ToLowerInvariant())
}

function Invoke-WingetUpgradeHook {
    param([string]$Phase, [string]$WingetOutput)
    if (-not $WingetOutput) { return }
    foreach ($pkgId in $script:Config.WingetUpgradeHooks.Keys) {
        if ($WingetOutput -match [regex]::Escape($pkgId)) {
            $hook = $script:Config.WingetUpgradeHooks[$pkgId][$Phase]
            if ($hook) { try { & $hook } catch { Write-Status "Hook ($Phase/$pkgId) failed: $_" -Type Warning } }
        }
    }
}

function Get-VSCodeCliPath {
    $candidates = @((
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:ProgramFiles  'Microsoft VS Code\bin\code.cmd'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft VS Code\bin\code.cmd'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code Insiders\bin\code-insiders.cmd'),
        (Join-Path $env:ProgramFiles  'Microsoft VS Code Insiders\bin\code-insiders.cmd')
    ) | Where-Object { $_ -and (Test-Path $_) })
    if ($candidates.Count -gt 0) { return $candidates[0] }
    foreach ($name in 'code.cmd','code-insiders.cmd') {
        $cmd = Get-Command $name -EA SilentlyContinue
        if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { return $cmd.Source }
    }
    return $null
}

function Get-ToolInstallManager([string]$Command) {
    if (Test-Command $Command) {
        $src = (Get-Command $Command).Source
        if ($src -like '*\scoop\*') { return 'scoop' }
        if ($src -like '*\WinGet\*' -or $src -like '*WindowsApps\*') { return 'winget' }
    }
    return $null
}

# --- Logging -------------------------------------------------------------------
$script:LogFile = if ($LogPath) { $LogPath } else { Join-Path $script:Config.StateDir 'updatescript.log' }

$logDir = Split-Path $script:LogFile -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false | Out-Null }
if ((Test-Path $script:LogFile) -and (Get-Item $script:LogFile).Length -gt ($script:Config.LogMaxSizeMB * 1MB)) {
    Rename-Item $script:LogFile "$script:LogFile.old" -Force -EA SilentlyContinue
}

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    try {
        Add-Content -Path $script:LogFile -Value "[$( Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" -Encoding UTF8 -WhatIf:$false
    } catch { Write-Verbose "Log write skipped: $($_.Exception.Message)" }
}

function Write-Section([string]$Title) {
    Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
    Write-Host "  $Title"      -ForegroundColor Cyan
    Write-Host "$('=' * 54)"   -ForegroundColor DarkGray
    Write-Log "--- $Title ---"
}

function Write-Status {
    param([string]$Message, [ValidateSet('Success','Warning','Error','Info')][string]$Type = 'Info')
    $colors  = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]';  Warning = '[!]';    Error = '[X]'; Info = '[*]' }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
    Write-Log $Message -Level $Type
}

function Write-Detail {
    param([string]$Message, [ValidateSet('Info','Muted','Warning','Error')][string]$Type = 'Info')
    $colors   = @{ Info = 'Gray'; Muted = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }
    $prefixes = @{ Info = '  >'; Muted = '  -'; Warning = '  !'; Error = '  x' }
    if ($Message) {
        Write-Host "$($prefixes[$Type]) $Message" -ForegroundColor $colors[$Type]
        Write-Log "$($prefixes[$Type]) $Message" -Level $Type
    }
}

function Write-FilteredOutput {
    param([AllowNull()][string]$Text, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    # Split on \r\n, bare \r, or bare \n to handle winget's carriage-return progress updates
    foreach ($rawLine in ($normalized -split '\r\n|\r|\n')) {
        $line = $rawLine.TrimEnd()
        if (-not $line) { continue }
        $compact = $line.Trim()
        if ($compact -match '^[\\/\|\-]+$') { continue }
        if ([regex]::Matches($compact, '[^\x00-\x7F]').Count -ge 3 -and $compact -notmatch '[A-Za-z0-9]') { continue }
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }
        if ($compact -match 'package\(s\) have pins that prevent upgrade') { continue }
        if ($compact -match '^[\-=]{6,}$') { continue }
        Write-Host $line -ForegroundColor $Color
        Write-Log $line
    }
}

function Invoke-StreamingCapture {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock, [ConsoleColor]$Color = [ConsoleColor]::Gray)
    $outputPath = [System.IO.Path]::GetTempFileName()
    try {
        & $ScriptBlock 2>&1 | Tee-Object -FilePath $outputPath | ForEach-Object {
            if ($_ -is [string]) { Write-FilteredOutput -Text $_ -Color $Color }
            elseif ($null -ne $_) {
                $text = $_.ToString()
                if (-not [string]::IsNullOrWhiteSpace($text)) { Write-Host $text -ForegroundColor $Color; Write-Log $text }
            }
        } | Out-Null
        return [pscustomobject]@{ OutputPath = $outputPath; ExitCode = $LASTEXITCODE }
    } catch {
        $captured = if (Test-Path $outputPath) { Get-Content -LiteralPath $outputPath -Raw -EA SilentlyContinue } else { '' }
        if ($captured) { Write-Log $captured -Level Error }
        throw
    }
}

function Read-CapturedOutput([string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return '' }
    try { return ((Get-Content -LiteralPath $Path -Raw -EA SilentlyContinue) -replace '\x00', '').Trim() }
    catch { return '' }
}

function Invoke-WingetWithTimeout {
    param([string[]]$Arguments, [int]$TimeoutSec = $script:Config.WingetTimeoutSec)
    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    try {
        $proc   = Start-Process winget -ArgumentList $Arguments -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { taskkill.exe /PID $proc.Id /T /F | Out-Null } catch {}
            try { $proc.Kill() } catch {}
            throw "winget timed out after ${TimeoutSec}s"
        }
        $stdout  = Get-Content -Raw -Path $stdoutFile -Encoding UTF8 -EA SilentlyContinue
        $stderr  = Get-Content -Raw -Path $stderrFile -Encoding UTF8 -EA SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()
        # Strip progress bar lines - split on `r too since winget uses carriage returns for in-place updates
        $combined = (($combined -split '\r\n|\r|\n') | Where-Object { $_ -notmatch '^[\s\p{S}\p{P}]*\d+%\s*$' }) -join "`n"
        return [pscustomobject]@{ Output = $combined; ExitCode = $proc.ExitCode }
    } finally {
        Remove-Item $stdoutFile, $stderrFile -Force -EA SilentlyContinue
    }
}

function Get-WingetUpgradeEntries([string]$WingetOutput) {
    $entries = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrWhiteSpace($WingetOutput)) { return @($entries) }
    foreach ($line in ($WingetOutput -split '\r\n|\r|\n')) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t -match '^(Name|Found|The following|No installed package|No applicable upgrade|There are no available upgrades|\d+\s+upgrades available)') { continue }
        if ($t -match '^[-\s]+$') { continue }
        if ($t -match '^(?<name>.+?)\s+(?<id>(?=.*[A-Za-z])[A-Za-z0-9][A-Za-z0-9\.\-_]*\.[A-Za-z0-9\.\-_]+)\s+(?<version>\S+)\s+(?<available>\S+)$') {
            $entries.Add([pscustomobject]@{ Name = $Matches.name.Trim(); Id = $Matches.id.Trim(); Version = $Matches.version; Available = $Matches.available })
        }
    }
    return @($entries)
}

function Test-WingetUpgradeListsMatch {
    param(
        [AllowNull()]$First,
        [AllowNull()]$Second
    )
    $firstList = @($First)
    $secondList = @($Second)
    if ($firstList.Count -ne $secondList.Count) { return $false }
    for ($i = 0; $i -lt $firstList.Count; $i++) {
        if ($firstList[$i].Id -ne $secondList[$i].Id) { return $false }
        if ($firstList[$i].Version -ne $secondList[$i].Version) { return $false }
        if ($firstList[$i].Available -ne $secondList[$i].Available) { return $false }
    }
    return $true
}

function Remove-WingetDuplicateUpgradeListing {
    param(
        [AllowNull()][string]$WingetOutput,
        [AllowNull()]$KnownEntries
    )
    if ([string]::IsNullOrWhiteSpace($WingetOutput)) { return $WingetOutput }
    $entries = @($KnownEntries)
    if ($entries.Count -eq 0) { return $WingetOutput }

    $lines = @($WingetOutput -split '\r\n|\r|\n')
    if ($lines.Count -eq 0) { return $WingetOutput }

    $installIndex = -1
    $summaryIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $trimmed = $lines[$i].Trim()
        if ($installIndex -lt 0 -and $trimmed -match '^\(\d+/\d+\)\s+Found ') { $installIndex = $i }
        if ($summaryIndex -lt 0 -and $trimmed -match '^\d+\s+upgrades available\.$') { $summaryIndex = $i }
    }

    if ($summaryIndex -lt 0 -or $installIndex -le $summaryIndex) { return $WingetOutput }

    $prefix = $lines[0..$summaryIndex]
    $matchedIds = 0
    foreach ($entry in $entries) {
        if ($prefix -match [regex]::Escape($entry.Id)) { $matchedIds++ }
    }

    if ($matchedIds -lt $entries.Count) { return $WingetOutput }
    return (($lines[($summaryIndex + 1)..($lines.Count - 1)] -join "`n").Trim())
}

# --- State Management ----------------------------------------------------------
$script:StateDir  = $script:Config.StateDir
$script:StateFile = Join-Path $script:StateDir 'state.json'
$script:State     = @{ LastRun = $null; Winget = @{}; Scoop = @{}; Chocolatey = @{}; WhatChanged = $null }

if (Test-Path $script:StateFile) {
    try {
        $loaded = Get-Content $script:StateFile -Raw | ConvertFrom-Json
        $script:State = @{
            LastRun     = $loaded.LastRun
            Winget      = ConvertTo-StringMap $loaded.Winget
            Scoop       = ConvertTo-StringMap $loaded.Scoop
            Chocolatey  = ConvertTo-StringMap $loaded.Chocolatey
            WhatChanged = $loaded.WhatChanged
        }
    } catch { Write-Warning "Could not load state: $_" }
}

function Save-State {
    $script:State.LastRun = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        $dir = Split-Path $script:StateFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $script:State | ConvertTo-Json -Depth 10 | Set-Content $script:StateFile -Encoding UTF8
    } catch { Write-Warning "Failed to save state: $_" }
}

function Update-WingetState {
    try {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) "winget-state-$([guid]::NewGuid().ToString('N')).json"
        winget export -o $tmp --include-versions --accept-source-agreements --disable-interactivity | Out-Null
        $data = Get-Content $tmp -Raw | ConvertFrom-Json
        $map  = @{}
        foreach ($src in @($data.Sources)) {
            foreach ($pkg in @($src.Packages)) {
                if ($pkg.PackageIdentifier -and $pkg.Version) { $map[[string]$pkg.PackageIdentifier] = [string]$pkg.Version }
            }
        }
        $script:State.Winget = $map
    } catch { Write-Warning "Could not update winget state: $_" }
    finally { Remove-Item $tmp -Force -EA SilentlyContinue }
}

function Update-ScoopState {
    if (-not (Test-Command 'scoop')) { return }
    try {
        $map = @{}
        scoop list 2>&1 | Out-String | ForEach-Object { $_ -split "`n" } | ForEach-Object {
            $t = $_.Trim()
            if ($t -and $t -notmatch '^Installed|^---') {
                $p = $t -split '\s+', 2
                if ($p.Count -ge 2) { $map[$p[0]] = $p[1] }
            }
        }
        $script:State.Scoop = $map
    } catch { Write-Warning "Could not update scoop state: $_" }
}

function Update-ChocolateyState {
    if (-not (Test-Command 'choco') -or -not $isAdmin) { return }
    try {
        $map = @{}
        choco list -lo 2>&1 | Out-String | ForEach-Object { $_ -split "`n" } | ForEach-Object {
            $t = $_.Trim()
            if ($t -and $t -notmatch '^\d+ packages|\^Chocolatey') {
                $p = $t -split '\s+', 2
                if ($p.Count -ge 2) { $map[$p[0]] = $p[1] }
            }
        }
        $script:State.Chocolatey = $map
    } catch { Write-Warning "Could not update chocolatey state: $_" }
}

function Compare-PackageMaps($prev, $curr) {
    $prev = ConvertTo-StringMap $prev; $curr = ConvertTo-StringMap $curr
    $changes = @()
    foreach ($id in $curr.Keys) {
        if (-not $prev.ContainsKey($id)) { $changes += "+ $id $($curr[$id]) (new)" }
        elseif ($prev[$id] -ne $curr[$id]) { $changes += "~ $id $($prev[$id]) -> $($curr[$id])" }
    }
    foreach ($id in $prev.Keys) { if (-not $curr.ContainsKey($id)) { $changes += "- $id $($prev[$id]) (removed)" } }
    return $changes
}

# --- Core Update Wrapper -------------------------------------------------------
$script:SectionTimings = @{}
$script:stepChanged    = $false
$script:stepMessage    = ''

function Complete-StepState {
    if ([string]::IsNullOrWhiteSpace($script:stepMessage)) {
        $script:stepMessage = if ($script:stepChanged) { 'updated' } else { 'checked' }
    }
}

function Invoke-Update {
    [CmdletBinding(SupportsShouldProcess)]
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
    if ($Disabled)                   { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'flag' }); return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAdmin -and -not $isAdmin) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'requires admin' }); return }
    if ($SlowOperation -and $FastMode)    { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = if ($UltraFast) { 'ultra fast mode' } else { 'fast mode' } }); return }
    if ($FastMode -and $script:Config.FastModeSkip -contains $Name) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = if ($UltraFast) { 'ultra fast mode' } else { 'fast mode' } }); return }
    if ($UltraFast -and $script:Config.UltraFastSkip -contains $Name) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'ultra fast mode' }); return }
    if ($script:IsSimulation) { Write-Host "  [DryRun] Would run: $Title" -ForegroundColor DarkCyan; return }
    if (-not $PSCmdlet.ShouldProcess($Name, "Update $Title")) { return }

    $sectionStart = Get-Date
    if (-not $NoSection) { Write-Section $Title }
    $script:stepChanged = $false
    $script:stepMessage = ''
    try {
        & $Action
        Complete-StepState
        $elapsed = ((Get-Date) - $sectionStart).TotalSeconds
        if ($script:stepChanged) {
            Write-Status "$Name updated ($([math]::Round($elapsed,1).ToString('F1',[cultureinfo]::InvariantCulture))s)" -Type Success
            $updateResults.Success.Add($Name)
        } else {
            if ($script:stepMessage) { Write-Detail $script:stepMessage -Type Muted }
            $updateResults.Checked.Add($Name)
        }
        $updateResults.Details[$Name]    = $script:stepMessage
        $script:SectionTimings[$Name]    = $elapsed
    } catch {
        Write-Status "$Name failed: $($_.Exception.Message)" -Type Error
        $updateResults.Failed.Add($Name)
        $updateResults.Details[$Name] = $_.Exception.Message
    }
}

# --- Parallel Batch Execution --------------------------------------------------
# Runs independent tasks concurrently via Start-ThreadJob.
# Each job captures Write-Host output (with color) to a temp file; the main
# thread replays them in order once all jobs finish.

# Build init script once - serializes helper functions into each runspace
$script:BatchInitScript = [scriptblock]::Create((
    @('Write-FilteredOutput','Write-Detail','Write-Status','Write-Log','Invoke-StreamingCapture',
      'Read-CapturedOutput','Get-ToolInstallManager','Test-Command','Normalize-PackageName',
      'Complete-StepState','Get-VSCodeCliPath') | ForEach-Object {
        $item = Get-Item "Function:$_" -EA SilentlyContinue
        if ($item) { "function $_ {`n$($item.ScriptBlock)`n}" }
    }
) -join "`n`n")

function Invoke-UpdateBatch {
    param(
        [Parameter(Mandatory)][hashtable[]]$Tasks,
        [int]$ThrottleLimit = 4
    )

    function Process-UpdateBatchResult {
        param(
            [Parameter(Mandatory)]$Info,
            $Result
        )

        $sectionTitle = $Info.Task.Name
        if ($Result -and $Result.PSObject.Properties['Title'] -and -not [string]::IsNullOrWhiteSpace([string]$Result.Title)) {
            $sectionTitle = [string]$Result.Title
        }
        Write-Section $sectionTitle

        if (Test-Path $Info.OutFile) {
            foreach ($raw in (Get-Content $Info.OutFile -Encoding UTF8)) {
                $idx = $raw.IndexOf("`t")
                if ($idx -gt 0 -and $idx -le 2 -and $raw.Substring(0, $idx) -match '^\d+$') {
                    Write-Host $raw.Substring($idx + 1) -ForegroundColor ([System.ConsoleColor][int]$raw.Substring(0, $idx))
                } else { Write-Host $raw }
            }
            Remove-Item $Info.OutFile -Force -EA SilentlyContinue
        }

        if (-not $Result) {
            Write-Status "$($Info.Task.Name): no result returned" -Type Warning
            $updateResults.Checked.Add($Info.Task.Name)
        } elseif ($Result.Error) {
            Write-Status "$($Result.Name) failed: $($Result.Error)" -Type Error
            $updateResults.Failed.Add($Result.Name)
            $updateResults.Details[$Result.Name] = $Result.Error
        } elseif ($Result.Changed) {
            $e = [math]::Round($Result.Elapsed, 1).ToString('F1', [cultureinfo]::InvariantCulture)
            Write-Status "$($Result.Name) updated (${e}s)" -Type Success
            $updateResults.Success.Add($Result.Name)
            $updateResults.Details[$Result.Name] = $Result.Message
        } else {
            if ($Result.Message) { Write-Detail $Result.Message -Type Muted }
            $updateResults.Checked.Add($Result.Name)
            $updateResults.Details[$Result.Name] = $Result.Message
        }
        if ($Result) { $script:SectionTimings[$Result.Name] = $Result.Elapsed }
    }

    # Evaluate skip conditions before spawning any jobs
    $active = foreach ($t in $Tasks) {
        $disabled = $t.Disabled -eq $true
        $missingCmd = $t.RequiresCommand -and -not (Test-Command $t.RequiresCommand)
        $missingAny = $t.RequiresAnyCommand -and -not ($t.RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)
        $needsAdmin = $t.RequiresAdmin -and -not $isAdmin
        $fastSkip   = $FastMode -and $script:Config.FastModeSkip -contains $t.Name
        $ultraSkip  = $UltraFast -and $script:Config.UltraFastSkip -contains $t.Name
        $skipped    = $disabled -or $missingCmd -or $missingAny -or $needsAdmin -or $fastSkip -or $ultraSkip
        if ($skipped) {
            $reason = if ($disabled) { 'flag' } elseif ($ultraSkip) { 'ultra fast mode' } elseif ($fastSkip) { if ($UltraFast) { 'ultra fast mode' } else { 'fast mode' } } elseif ($needsAdmin) { 'requires admin' } else { 'not installed' }
            $updateResults.Skipped.Add([pscustomobject]@{ Name = $t.Name; Reason = $reason })
        } else { $t }
    }
    if (-not $active) { return }

    $initScript = $script:BatchInitScript
    $workerScript = {
        param($task, $outFile, $logFile, $cfg)
        $script:LogFile  = $logFile
        $script:Config   = $cfg
        $commandCache    = @{}

        # Shadow Write-Host to capture colored output to a file
        function Write-Host {
            param($Object = '', $ForegroundColor, $BackgroundColor, [switch]$NoNewline)
            $c = if ($ForegroundColor -is [System.ConsoleColor]) { [int]$ForegroundColor }
                 elseif ($ForegroundColor) { try { [int][System.ConsoleColor]$ForegroundColor } catch { 7 } }
                 else { 7 }
            [System.IO.File]::AppendAllText($outFile, "$c`t$Object`n", [System.Text.Encoding]::UTF8)
        }

        $script:stepChanged = $false
        $script:stepMessage = ''
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $errMsg = $null
        try {
            $actionBlock = [scriptblock]::Create($task.Action.ToString())
            & $actionBlock
            Complete-StepState
        }
        catch { $errMsg = $_.Exception.Message }

        [pscustomobject]@{
            Name    = $task.Name
            Title   = if ($task.Title) { $task.Title } else { $task.Name }
            Changed = $script:stepChanged
            Message = $script:stepMessage
            Error   = $errMsg
            Elapsed = $sw.Elapsed.TotalSeconds
        }
    }

    if (-not (Get-Command Start-ThreadJob -EA SilentlyContinue)) {
        try { Import-Module ThreadJob -EA SilentlyContinue | Out-Null } catch {}
    }
    $canUseThreadJobs = [bool](Get-Command Start-ThreadJob -EA SilentlyContinue)

    if (-not $canUseThreadJobs) {
        foreach ($task in $active) {
            $outFile = [System.IO.Path]::GetTempFileName()
            $info = [pscustomobject]@{ OutFile = $outFile; Task = $task }
            $result = & $workerScript $task $outFile $script:LogFile $script:Config
            Process-UpdateBatchResult -Info $info -Result $result
        }
        return
    }

    # Launch all jobs
    $jobInfos = foreach ($task in $active) {
        $outFile  = [System.IO.Path]::GetTempFileName()
        $logFile  = $script:LogFile
        $cfg      = $script:Config

        $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit -InitializationScript $initScript -ArgumentList $task, $outFile, $logFile, $cfg -ScriptBlock $workerScript
        [pscustomobject]@{ Job = $job; OutFile = $outFile; Task = $task }
    }

    # Wait for all, then replay output in submission order
    $null = $jobInfos.Job | Wait-Job

    foreach ($info in $jobInfos) {
        $result = Receive-Job $info.Job -EA SilentlyContinue
        Remove-Job $info.Job -Force
        Process-UpdateBatchResult -Info $info -Result $result
    }
}

# ==============================================================================
# PACKAGE MANAGERS
# ==============================================================================

Invoke-Update -Name 'Scoop' -RequiresCommand 'scoop' -Disabled:($script:Config.SkipManagers -contains 'Scoop') -Action {
    $out = Read-CapturedOutput (Invoke-StreamingCapture { scoop update }).OutputPath
    if ($out -match '(?i)error|failed') { Write-Status "Scoop self-update warning: $out" -Type Warning }
    $out = Read-CapturedOutput (Invoke-StreamingCapture { scoop update '*' }).OutputPath
    if ($out -and $out -notmatch 'Scoop is up to date') { $script:stepChanged = $true; $script:stepMessage = 'updated correctly' }
    scoop cleanup '*' 2>&1 | Out-Null
    scoop cache rm '*' 2>&1 | Out-Null
}

Invoke-Update -Name 'Winget' -RequiresCommand 'winget' -Action {
    if (-not $isAdmin) {
        Write-Status 'Running as NON-ADMIN. MSI installers will likely fail with 1632.' -Type Warning
        Write-Status 'Re-run with -AutoElevate or as Administrator.' -Type Info
    }
    $oldTemp = $env:TEMP; $oldTmp = $env:TMP
    if ($isAdmin -and (Test-Path 'C:\Windows\Temp')) { $env:TEMP = 'C:\Windows\Temp'; $env:TMP = 'C:\Windows\Temp' }
    try {
        # Restart Windows Installer service — prevents 1603 errors from stale msiexec state
        if ($isAdmin) {
            try { Restart-Service msiserver -Force -EA Stop; Start-Sleep 1 }
            catch { Write-Detail "Could not restart Windows Installer service: $_" -Type Warning }
        }

        # Warn if a reboot is pending — MSI installers (Go, Chrome, etc.) will fail with 1603 until rebooted
        $rebootPending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
                         (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                         (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -EA SilentlyContinue)
        if ($rebootPending) {
            Write-Status 'A reboot is pending — MSI installers (Go, Chrome, etc.) will fail with 1603 until you reboot.' -Type Warning
        }

        # Refresh source catalog before scanning unless UltraFast mode is prioritizing speed over metadata freshness
        if (-not $UltraFast) {
            Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('source','update','--disable-interactivity') | Out-Null
        }

        $scan = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @('upgrade','--include-unknown','--source','winget','--accept-source-agreements','--disable-interactivity')
        if ($scan.Output) { Write-FilteredOutput -Text $scan.Output -Color Gray }
        $scanEntries = @(Get-WingetUpgradeEntries $scan.Output)
        $script:Config.WingetScannedIds = @($scanEntries | ForEach-Object { $_.Id })
        # Share the scanned package IDs so dedicated tool sections can skip redundant winget calls
        $script:_wingetScannedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $scanEntries) { $null = $script:_wingetScannedIds.Add($e.Id) }
        try { Invoke-WingetUpgradeHook -Phase 'Pre' -WingetOutput $scan.Output } catch {}

        if ($UltraFast -and $scanEntries.Count -gt 0 -and ($rebootPending -or -not $isAdmin)) {
            Write-Detail 'UltraFast mode: skipping the winget install pass because reboot/elevation issues would likely waste time.' -Type Warning
            $script:stepMessage = "found $($scanEntries.Count) upgrade(s); install pass skipped for speed"
            return
        }

        # Upgrade all — cap at 3× per-package timeout; fall back per-package if bulk times out or fails
        $bulkTimeout = [math]::Max($script:Config.WingetTimeoutSec * 3, 600)
        $upgradeResult = Invoke-WingetWithTimeout -TimeoutSec $bulkTimeout -Arguments @('upgrade','--all','--include-unknown','--source','winget','--silent','--accept-source-agreements','--accept-package-agreements','--disable-interactivity')
        $upgradeDisplayOutput = Remove-WingetDuplicateUpgradeListing -WingetOutput $upgradeResult.Output -KnownEntries $scanEntries
        if ($upgradeDisplayOutput) { Write-FilteredOutput -Text $upgradeDisplayOutput -Color Gray }

        $bulkTimedOut = $upgradeResult.Output -match 'timed out'
        $bulkHadKnownInstallerFailures = $upgradeResult.Output -match 'Installer failed with exit code: 1603'
        $shouldRetryIndividually = $bulkTimedOut -or ($upgradeResult.ExitCode -ne 0 -and
            $upgradeResult.Output -notmatch 'Successfully installed|No applicable upgrade|already installed' -and
            -not $bulkHadKnownInstallerFailures -and -not $rebootPending)
        if ($shouldRetryIndividually) {
            if ($scanEntries.Count -gt 0) {
                Write-Detail 'Bulk upgrade timed out or failed unexpectedly — retrying packages individually' -Type Warning
                foreach ($entry in $scanEntries) {
                    $perResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @('upgrade','--id',$entry.Id,'--include-unknown','--source','winget','--silent','--accept-source-agreements','--accept-package-agreements','--disable-interactivity')
                    if ($perResult.Output) { Write-FilteredOutput -Text $perResult.Output -Color Gray }
                }
                $upgradeResult = $perResult  # use last result for exit-code check below
            } else {
                Write-Detail 'Bulk upgrade failed unexpectedly, but no individual packages were parsed for retry.' -Type Warning
            }
        } elseif ($bulkHadKnownInstallerFailures -and $rebootPending) {
            Write-Detail 'Skipping per-package retries because a pending reboot is blocking MSI-based installers.' -Type Warning
        }

        $finalScan = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @('upgrade','--include-unknown','--source','winget','--accept-source-agreements','--disable-interactivity')
        $finalEntries = @(Get-WingetUpgradeEntries $finalScan.Output)
        if ($finalScan.Output -and -not (Test-WingetUpgradeListsMatch -First $scanEntries -Second $finalEntries)) {
            Write-Detail 'Remaining upgrades after winget run:' -Type Muted
            Write-FilteredOutput -Text $finalScan.Output -Color Gray
        }
        try { Invoke-WingetUpgradeHook -Phase 'Post' -WingetOutput $finalScan.Output } catch {}

        $anyInstalled = $upgradeResult.Output -match 'Successfully installed'
        $script:stepChanged = $anyInstalled
        $script:stepMessage = if ($upgradeResult.ExitCode -ne 0 -and -not $anyInstalled) { 'completed with some failures' }
                              elseif ($anyInstalled) { 'updated correctly' }
                              else { 'already current' }
    } finally { $env:TEMP = $oldTemp; $env:TMP = $oldTmp }
}

Invoke-Update -Name 'Chocolatey' -RequiresCommand 'choco' -RequiresAdmin -Disabled:($script:Config.SkipManagers -contains 'Chocolatey') -Action {
    $out = Read-CapturedOutput (Invoke-StreamingCapture { choco upgrade all -y }).OutputPath
    if ($out -match 'upgraded 0/') { $script:stepMessage = 'no package updates' }
    else { $script:stepChanged = $true; $script:stepMessage = 'packages upgraded' }
}

Invoke-Update -Name 'WSL' -Title 'Windows Subsystem for Linux' -RequiresCommand 'wsl' -RequiresAdmin -Disabled:($SkipWSL -or $script:Config.SkipManagers -contains 'WSL') -Action {
    # wsl --update emits UTF-16LE; correct encoding for the call
    $out = Read-CapturedOutput (Invoke-StreamingCapture {
        $saved = [Console]::OutputEncoding
        [Console]::OutputEncoding = [System.Text.Encoding]::Unicode
        try { wsl --update } finally { [Console]::OutputEncoding = $saved }
    }).OutputPath
    $script:stepChanged = $out -and $out -notmatch 'most recent version.+already installed'
    $script:stepMessage = if ($script:stepChanged) { 'WSL platform updated' } else { 'platform already up to date' }
    if (-not $SkipWSLDistros) {
        $distros = @(wsl --list --quiet 2>&1 | ForEach-Object { ($_ -replace '\x00','').Trim() } | Where-Object { $_ })
        if ($distros.Count -eq 0) { $script:stepMessage = 'platform current; no distros found' }
        else {
            foreach ($d in $distros) {
                Write-Detail "Updating WSL Distro: $d"
                wsl -d $d -- sh -lc 'if command -v apt-get >/dev/null; then sudo apt-get update -qq && sudo apt-get upgrade -y -qq; elif command -v pacman >/dev/null; then sudo pacman -Syu --noconfirm; fi' 2>&1 | Out-Null
            }
            $script:stepMessage += ' (Checked inside Distros)'
        }
    }
}

# ==============================================================================
# SYSTEM COMPONENTS - run in parallel
# ==============================================================================

Write-Host "`n[Parallel] Running system component updates..." -ForegroundColor DarkCyan

Invoke-UpdateBatch -ThrottleLimit 3 -Tasks @(
    @{
        Name           = 'WindowsUpdate'
        Title          = 'Windows Update'
        RequiresAdmin  = $true
        Disabled       = $SkipWindowsUpdate
        Action         = {
            Import-Module PSWindowsUpdate
            $params = @{ Install = $true; AcceptAll = $true; NotCategory = 'Drivers'; IgnoreReboot = $true; RecurseCycle = 3; Verbose = $false; Confirm = $false }
            $out = Read-CapturedOutput (Invoke-StreamingCapture ({ Get-WindowsUpdate @params }.GetNewClosure())).OutputPath
            if ($out -match 'No updates found|There are no applicable updates') { $script:stepMessage = 'already current' }
            else {
                $script:stepChanged = $true; $script:stepMessage = 'Updated successfully'
                if (-not $SkipReboot -and (Get-WURebootStatus -Silent -EA SilentlyContinue)) { $script:stepMessage += ' (Reboot pending)' }
            }
        }
    }
    @{
        Name            = 'StoreApps'
        Title           = 'Microsoft Store Apps'
        Disabled        = $SkipStoreApps
        RequiresCommand = 'winget'
        Action          = {
            Write-Detail 'Checking Microsoft Store app upgrades via winget'
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --source msstore --all --silent --accept-package-agreements --accept-source-agreements }).OutputPath
            if (-not $out -or $out -match 'No installed package found|No applicable upgrade|There are no available upgrades') { $script:stepMessage = 'no Microsoft Store app updates' }
            elseif ($out -match 'Successfully installed|Successfully upgraded') { $script:stepChanged = $true; $script:stepMessage = 'store apps updated' }
            else { $script:stepMessage = 'store apps checked' }
        }
    }
    @{
        Name            = 'DefenderSignatures'
        Title           = 'Microsoft Defender Signatures'
        Disabled        = $SkipDefender
        RequiresCommand = 'Update-MpSignature'
        RequiresAdmin   = $true
        Action          = {
            $status = Get-MpComputerStatus -EA SilentlyContinue
            if (-not $status -or -not ($status.AMServiceEnabled -and $status.AntivirusEnabled)) { $script:stepMessage = 'Microsoft Defender is not the active AV'; return }
            $before = $status.AntivirusSignatureVersion
            Invoke-StreamingCapture { Update-MpSignature -EA Stop } | Out-Null
            $after  = (Get-MpComputerStatus -EA SilentlyContinue).AntivirusSignatureVersion
            if ($after -and $after -ne $before) { $script:stepChanged = $true; $script:stepMessage = "signatures updated to $after" }
            else { $script:stepMessage = 'signatures already current' }
        }
    }
)

# ==============================================================================
# DEV TOOLS - independent tools run in parallel
# ==============================================================================

Write-Host "`n[Parallel] Running dev tool updates..." -ForegroundColor DarkCyan

Invoke-UpdateBatch -ThrottleLimit $ParallelThrottle -Tasks @(
    @{
        Name            = 'npm'
        Title           = 'npm (Node.js)'
        RequiresCommand = 'npm'
        Disabled        = $SkipNode
        Action          = {
            $current = (npm --version 2>&1).Trim(); $latest = (npm view npm version 2>&1).Trim()
            if ($current -ne $latest) {
                $run = Invoke-StreamingCapture { npm install -g npm@latest }
                if ($run.ExitCode -eq 0 -and (npm --version 2>&1).Trim() -eq $latest) { $script:stepChanged = $true }
                else {
                    $out = Read-CapturedOutput $run.OutputPath
                    Write-Detail "npm self-update failed: $($out -split '\r\n|\r|\n' | Where-Object { $_.Trim() } | Select-Object -First 1)" -Type Warning
                }
            }
            $outdatedJson = Read-CapturedOutput (Invoke-StreamingCapture { npm outdated -g --json }).OutputPath
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -EA SilentlyContinue
                if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
                    $pkgs = $outdated.PSObject.Properties.Name
                    $run  = Invoke-StreamingCapture ({ npm install -g $pkgs }.GetNewClosure())
                    if ($run.ExitCode -eq 0) { $script:stepChanged = $true; $script:stepMessage = "Updated $($pkgs.Count) package(s)" }
                    else { Write-Detail "npm package update failed" -Type Warning }
                } else { $script:stepMessage = 'no global package updates' }
            } else { $script:stepMessage = 'no global package updates' }
            npm cache clean --force 2>&1 | Out-Null
        }
    }
    @{
        Name            = 'pnpm'
        RequiresCommand = 'pnpm'
        Disabled        = $SkipNode
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { pnpm update -g }).OutputPath
            if ($out -match 'Already up to date|Nothing to|No global packages found') { $script:stepMessage = 'no global package updates' }
            else { $script:stepChanged = $true; $script:stepMessage = 'global packages updated' }
        }
    }
    @{
        Name            = 'Bun'
        RequiresCommand = 'bun'
        Disabled        = $SkipNode
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { bun upgrade }).OutputPath
            if ($out -notmatch 'already on the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
        }
    }
    @{
        Name            = 'Deno'
        RequiresCommand = 'deno'
        Disabled        = $SkipNode
        Action          = {
            $mgr = Get-ToolInstallManager 'deno'
            if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { deno upgrade }).OutputPath
            if ($out -notmatch 'already the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
        }
    }
    @{
        Name            = 'Rust'
        RequiresCommand = 'rustup'
        Disabled        = $SkipRust
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { rustup update }).OutputPath
            if ($out -match 'unchanged - rustc ([^\s]+)') { $script:stepMessage = "stable unchanged ($($Matches[1]))" }
            elseif ($out -match 'updated - rustc') { $script:stepChanged = $true; $script:stepMessage = 'stable updated' }
            else { $script:stepMessage = 'toolchain checked' }
        }
    }
    @{
        Name            = 'cargo-binaries'
        Title           = 'Cargo Global Binaries'
        RequiresCommand = 'cargo'
        Disabled        = $SkipRust
        Action          = {
            if (-not (Test-Command 'cargo-install-update')) { Invoke-StreamingCapture { cargo install cargo-update } | Out-Null }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { cargo install-update -a }).OutputPath
            if ($out -match 'Overall updated [1-9]') { $script:stepChanged = $true; $script:stepMessage = 'cargo binaries updated' }
            elseif ($out -match 'Failed to update|could not compile|link.exe') { $script:stepMessage = 'cargo binaries checked; update failed' }
            else { $script:stepMessage = 'cargo binaries checked' }
        }
    }
    @{
        Name            = 'Go'
        RequiresCommand = 'go'
        Disabled        = $SkipGo
        Action          = {
            $mgr = Get-ToolInstallManager 'go'
            if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
            # GoLang.Go is a winget package — skip if the main Winget section already attempted it
            $wingetScannedIds = @($script:Config.WingetScannedIds)
            if ($wingetScannedIds -contains 'GoLang.Go') {
                $script:stepMessage = 'managed by winget (already updated)'; return
            }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --id GoLang.Go --accept-source-agreements --accept-package-agreements --disable-interactivity }).OutputPath
            if ($out -match 'Successfully installed') { $script:stepChanged = $true; $script:stepMessage = 'updated via winget' }
            elseif ($out -match 'Installer failed|exit code') { $script:stepMessage = "upgrade failed: $([regex]::Match($out,'exit code[:\s]+(\S+)').Groups[1].Value)" }
            else { $script:stepMessage = 'no newer version available' }
        }
    }
    @{
        Name            = 'gh-extensions'
        Title           = 'GitHub CLI Extensions'
        RequiresCommand = 'gh'
        Action          = {
            $list = Read-CapturedOutput (Invoke-StreamingCapture { gh extension list }).OutputPath
            if (-not $list) { $script:stepMessage = 'no extensions installed'; return }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { gh extension upgrade --all }).OutputPath
            if ($out -match 'upgraded|updated|installing') { $script:stepChanged = $true; $script:stepMessage = 'extensions updated' }
            else { $script:stepMessage = 'extensions checked' }
        }
    }
    @{
        Name            = 'yt-dlp'
        RequiresCommand = 'yt-dlp'
        Action          = {
            $src = (Get-Command yt-dlp).Source
            if ($src -like '*scoop*') { $script:stepMessage = 'managed by Scoop'; return }
            if ($src -like '*pip*') {
                $out = Read-CapturedOutput (Invoke-StreamingCapture { python -m pip install --upgrade yt-dlp }).OutputPath
                if ($out -notmatch 'Requirement already satisfied') { $script:stepChanged = $true; $script:stepMessage = 'updated via pip' } else { $script:stepMessage = 'already current' }
                return
            }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { yt-dlp -U }).OutputPath
            if ($out -match 'is up to date|Latest version') { $script:stepMessage = 'already current' }
            else { $script:stepChanged = $true; $script:stepMessage = 'updated' }
        }
    }
    @{
        Name               = 'tldr'
        Title              = 'tldr Cache'
        RequiresAnyCommand = @('tldr','tealdeer')
        Action             = {
            $cmd = if (Test-Command 'tealdeer') { 'tealdeer' } else { 'tldr' }
            & $cmd --update 2>&1 | Out-Null; $script:stepMessage = 'cache updated'
        }
    }
    @{
        Name            = 'oh-my-posh'
        Title           = 'Oh My Posh'
        RequiresCommand = 'oh-my-posh'
        Action          = {
            $mgr = Get-ToolInstallManager 'oh-my-posh'
            if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
            oh-my-posh upgrade 2>&1 | Out-Null; $script:stepChanged = $true; $script:stepMessage = 'updated'
        }
    }
    @{
        Name            = 'Volta'
        RequiresCommand = 'volta'
        Disabled        = $SkipNode
        Action          = { $script:stepMessage = 'standalone tool; update manually' }
    }
    @{
        Name            = 'fnm'
        RequiresCommand = 'fnm'
        Disabled        = $SkipNode
        Action          = { $script:stepMessage = 'node manager checked' }
    }
    @{
        Name            = 'mise'
        RequiresCommand = 'mise'
        Action          = { mise self-update --yes 2>&1 | Out-Null; mise upgrade 2>&1 | Out-Null; $script:stepMessage = 'updated' }
    }
    @{
        Name            = 'juliaup'
        RequiresCommand = 'juliaup'
        Action          = { juliaup update 2>&1 | Out-Null; $script:stepMessage = 'checked' }
    }
    @{
        Name            = 'git-lfs'
        RequiresCommand = 'git-lfs'
        Disabled        = $SkipGitLFS
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --id GitHub.GitLFS --accept-source-agreements --accept-package-agreements --disable-interactivity }).OutputPath
            if ($out -match 'Successfully installed|Successfully upgraded') { $script:stepChanged = $true; $script:stepMessage = 'updated via winget' } else { $script:stepMessage = 'already current' }
        }
    }
    @{
        Name            = 'git-credential-manager'
        RequiresCommand = 'git-credential-manager'
        Action          = {
            Invoke-WingetWithTimeout -Arguments @('upgrade','--id','Git.Git','--accept-source-agreements','--accept-package-agreements','--disable-interactivity') | Out-Null
            $script:stepMessage = 'updated via git upgrade'
        }
    }
    @{
        Name            = 'Flutter'
        Title           = 'Flutter'
        RequiresCommand = 'flutter'
        Disabled        = $SkipFlutter
        Action          = {
            $mgr = Get-ToolInstallManager 'flutter'
            if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { flutter upgrade }).OutputPath
            if ($out -match 'already up to date|already on the latest') { $script:stepMessage = 'already current' }
            elseif ($?) { $script:stepChanged = $true; $script:stepMessage = 'updated' }
            else { $script:stepMessage = 'checked' }
        }
    }
    @{
        Name            = 'ollama-models'
        Title           = 'Ollama Models'
        RequiresCommand = 'ollama'
        Disabled        = (-not $UpdateOllamaModels)
        Action          = {
            $models = @(ollama list 2>&1 | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] } | Where-Object { $_ })
            foreach ($m in $models) { ollama pull $m }
            $script:stepMessage = "checked $($models.Count) models"
        }
    }
    @{
        Name            = 'pipx'
        Title           = 'pipx (Python Tools)'
        RequiresCommand = 'pipx'
        Action          = { pipx upgrade-all 2>&1 | Out-Null; $script:stepMessage = 'upgraded all' }
    }
    @{
        Name            = 'Poetry'
        RequiresCommand = 'poetry'
        Disabled        = $SkipPoetry
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { poetry self update }).OutputPath
            if ($out -notmatch 'already using the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
        }
    }
    @{
        Name            = 'Composer'
        Title           = 'Composer (PHP)'
        RequiresCommand = 'composer'
        Disabled        = $SkipComposer
        Action          = { composer self-update --no-interaction 2>&1 | Out-Null; composer global update --no-interaction 2>&1 | Out-Null; $script:stepMessage = 'checked' }
    }
    @{
        Name            = 'RubyGems'
        RequiresCommand = 'gem'
        Disabled        = $SkipRuby
        Action          = { gem update --system 2>&1 | Out-Null; gem update 2>&1 | Out-Null; $script:stepMessage = 'checked' }
    }
    @{
        Name            = 'vscode-extensions'
        Title           = 'VS Code Extensions'
        Disabled        = $SkipVSCodeExtensions
        Action          = {
            $cli = Get-VSCodeCliPath
            if (-not $cli) { $script:stepMessage = 'VS Code not found'; return }
            $out = Read-CapturedOutput (Invoke-StreamingCapture ({ & $cli --update-extensions }.GetNewClosure())).OutputPath
            if ($out -match 'updated|installing') { $script:stepChanged = $true; $script:stepMessage = 'extensions updated' } else { $script:stepMessage = 'extensions checked' }
        }
    }
)

# ==============================================================================
# REMAINING TOOL UPDATES - parallelize heavy independent work
# ==============================================================================

Write-Host "`n[Parallel] Running remaining tool updates..." -ForegroundColor DarkCyan

Invoke-UpdateBatch -ThrottleLimit ([Math]::Min($ParallelThrottle, 4)) -Tasks @(
    @{
        Name            = 'uv'
        Title           = 'UV Toolchain'
        RequiresCommand = 'uv'
        Action          = {
            $messages = [System.Collections.Generic.List[string]]::new()

            $out = Read-CapturedOutput (Invoke-StreamingCapture { uv self update }).OutputPath
            if ($out -match 'installed via pip|standalone installation scripts') { $messages.Add('uv managed by another installer') }
            elseif ($out -match 'updated|successfully') { $script:stepChanged = $true; $messages.Add('uv updated') }
            else { $messages.Add('uv already current') }

            if (-not $script:Config.SkipUVTools) {
                $toolOut = Read-CapturedOutput (Invoke-StreamingCapture { uv tool upgrade --all }).OutputPath
                if ($toolOut -match 'Updated|Installed|upgraded') { $script:stepChanged = $true; $messages.Add('tools upgraded') }
                elseif ($toolOut -match 'Nothing to upgrade|no tools') { $messages.Add('no tools to upgrade') }
                else { $messages.Add('tools checked') }
            }

            $pythonOut = Read-CapturedOutput (Invoke-StreamingCapture { uv python upgrade }).OutputPath
            if ($pythonOut -match 'Upgraded|Installed|updated') { $script:stepChanged = $true; $messages.Add('Python versions updated') }
            else { $messages.Add('Python versions checked') }

            $script:stepMessage = $messages -join '; '
        }
    }
    @{
        Name  = 'pip'
        Title = 'Python / pip'
        Action = {
            function Invoke-PipHealthCheck {
                $out = (python -m pip check 2>&1) | Out-String
                if ($out -and $out.Trim() -notmatch '^No broken requirements') {
                    ($out -split '\r\n|\r|\n' | Where-Object { $_.Trim() }) | ForEach-Object { Write-Detail $_ -Type Warning }
                }
            }

            python -m pip install --upgrade pip 2>&1 | Out-Null
            $outdated = @(python -m pip list --outdated --format=json 2>$null | ConvertFrom-Json)
            if (-not $outdated -or $outdated.Count -eq 0) { Invoke-PipHealthCheck; $script:stepMessage = 'no global pip package updates'; return }

            $notRequired = @(python -m pip list --not-required --format=json 2>$null | ConvertFrom-Json)
            $topLevel    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($p in $notRequired) { $null = $topLevel.Add((Normalize-PackageName $p.name)) }

            if ($topLevel.Count -eq 0) {
                Write-Detail 'Cannot determine top-level packages; skipping upgrades to avoid breaking deps.' -Type Warning
                Invoke-PipHealthCheck; $script:stepMessage = 'pip checked; skipped for safety'; return
            }

            $skip    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $allow   = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($p in $script:Config.PipSkipPackages)  { $null = $skip.Add((Normalize-PackageName $p)) }
            foreach ($p in $script:Config.PipAllowPackages) { $null = $allow.Add((Normalize-PackageName $p)) }

            $updated  = [System.Collections.Generic.List[string]]::new()
            $failed   = [System.Collections.Generic.List[string]]::new()
            $excluded = [System.Collections.Generic.List[string]]::new()
            $batch    = @()
            $nSkip = 0; $nTransitive = 0

            foreach ($pkg in $outdated) {
                $n = Normalize-PackageName $pkg.name
                if ($skip.Contains($n))                                    { $excluded.Add($pkg.name); Write-Detail "$($pkg.name) skipped: excluded" -Type Warning; continue }
                if (-not $topLevel.Contains($n))                           { $nTransitive++; continue }
                if ($allow.Count -gt 0 -and -not $allow.Contains($n))     { $nSkip++; continue }
                $batch += $pkg.name
            }
            if ($nTransitive -gt 0) { Write-Detail "$nTransitive transitive dependenc$(if($nTransitive -eq 1){'y'}else{'ies'}) skipped" }
            if ($nSkip       -gt 0) { Write-Detail "$nSkip package$(if($nSkip -eq 1){''} else{'s'}) not in PipAllowPackages allowlist" }

            if ($batch.Count -eq 0) {
                Invoke-PipHealthCheck
                $parts = @()
                if ($failed.Count   -gt 0) { $parts += "failed: $($failed -join ', ')" }
                if ($excluded.Count -gt 0) { $parts += "excluded: $($excluded -join ', ')" }
                $script:stepMessage = if ($parts.Count -gt 0) { "No pip packages updated; $($parts -join '; ')" } else { 'no global pip package updates' }
                return
            }

            $run = Invoke-StreamingCapture ({ python -m pip install --upgrade --quiet --disable-pip-version-check --no-input $batch }.GetNewClosure())
            if ($run.ExitCode -eq 0) {
                foreach ($p in $batch) { $updated.Add($p) }
            } else {
                foreach ($p in $batch) {
                    $r = Invoke-StreamingCapture ({ python -m pip install --upgrade --quiet --disable-pip-version-check --no-input $p }.GetNewClosure())
                    if ($r.ExitCode -eq 0) { $updated.Add($p) }
                    else {
                        $failed.Add($p)
                        $firstLine = (Read-CapturedOutput $r.OutputPath) -split '\r\n|\r|\n' | Where-Object { $_.Trim() } | Select-Object -First 1
                        if ($firstLine) { Write-Detail "$p failed: $firstLine" -Type Warning }
                    }
                }
            }

            Invoke-PipHealthCheck
            if ($updated.Count -gt 0) { $script:stepChanged = $true }
            $parts = @()
            if ($updated.Count  -gt 0) { $parts += "Updated $($updated.Count) package$(if($updated.Count -eq 1){''}else{'s'})" }
            elseif ($failed.Count -eq 0 -and $excluded.Count -eq 0) { $parts += if ($nTransitive -gt 0) { "Updated 0 top-level packages" } else { 'no global pip package updates' } }
            if ($failed.Count   -gt 0) { $parts += "failed: $($failed -join ', ')" }
            if ($excluded.Count -gt 0) { $parts += "excluded: $($excluded -join ', ')" }
            if ($nTransitive    -gt 0 -and $updated.Count -gt 0) { $parts += "$nTransitive transitive deps skipped" }
            $script:stepMessage = $parts -join '; '
        }
    }
    @{
        Name            = 'dotnet'
        Title           = '.NET Tools'
        RequiresCommand = 'dotnet'
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { dotnet tool update --global --all }).OutputPath
            if ($out -match 'was successfully updated|successfully updated|reinstalled with the stable') {
                $script:stepChanged = $true; $script:stepMessage = 'tools updated'
            } elseif ($out -match 'Failed to uninstall tool package.*Could not find a part of the path') {
                $broken = [regex]::Matches($out, "Tool '([\w.-]+)' failed") | ForEach-Object { $_.Groups[1].Value }
                $repairedOk = [System.Collections.Generic.List[string]]::new()
                $repairFailed = [System.Collections.Generic.List[string]]::new()
                foreach ($tool in $broken) {
                    Write-Detail "Repairing $tool (uninstall + reinstall)..." -Type Warning
                    dotnet tool uninstall -g $tool 2>&1 | Out-Null
                    $installOut = (dotnet tool install -g $tool 2>&1) | Out-String
                    if ($installOut -match 'successfully installed|already installed') { $repairedOk.Add($tool) }
                    else { $repairFailed.Add($tool); Write-Detail "$tool reinstall failed: $(($installOut -split '\r\n|\r|\n' | Where-Object { $_.Trim() } | Select-Object -First 1))" -Type Warning }
                }
                $script:stepChanged = $repairedOk.Count -gt 0
                $repairParts = @()
                if ($repairedOk.Count   -gt 0) { $repairParts += "repaired: $($repairedOk -join ', ')" }
                if ($repairFailed.Count -gt 0) { $repairParts += "repair failed: $($repairFailed -join ', ')" }
                $script:stepMessage = if ($repairParts.Count -gt 0) { $repairParts -join '; ' } else { 'tool update errors encountered' }
            } elseif ($out -match 'failed to update|Failed to uninstall|Tool .* failed') {
                $script:stepMessage = 'tool update errors encountered'
            } else {
                $script:stepMessage = 'tools checked'
            }
        }
    }
    @{
        Name            = 'dotnet-workloads'
        Title           = '.NET Workloads'
        RequiresCommand = 'dotnet'
        RequiresAdmin   = $true
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { dotnet workload update }).OutputPath
            if ($out -match 'Updated advertising manifest|Installing|Installed') { $script:stepChanged = $true; $script:stepMessage = 'workloads updated' }
            else { $script:stepMessage = 'workloads checked' }
        }
    }
    @{
        Name     = 'pwsh-resources'
        Title    = 'PowerShell Modules / Resources'
        Disabled = $SkipPowerShellModules
        Action   = {
            $details = [System.Collections.Generic.List[string]]::new()
            $changed = $false
            $psResourceNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
                $cmd   = Get-Command Update-PSResource -EA SilentlyContinue
                $splat = @{ Name = '*'; ErrorAction = 'SilentlyContinue' }
                if ($cmd -and $cmd.Parameters.ContainsKey('AcceptLicense')) { $splat.AcceptLicense = $true }
                $before = @{}
                @(Get-InstalledPSResource -EA SilentlyContinue) | ForEach-Object {
                    $before[$_.Name] = [string]$_.Version
                    $null = $psResourceNames.Add($_.Name)
                }
                try { Update-PSResource @splat 2>&1 | Out-Null } catch { Write-Verbose "Update-PSResource: $($_.Exception.Message)" }
                @(Get-InstalledPSResource -EA SilentlyContinue) | ForEach-Object {
                    $prev = $before[$_.Name]; $curr = [string]$_.Version
                    if ($prev -and $curr -ne $prev) { $changed = $true; $details.Add("PSResource $($_.Name) $prev -> $curr") }
                }
            }

            if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
                $cmd   = Get-Command Update-Module -EA SilentlyContinue
                $splat = @{ ErrorAction = 'SilentlyContinue' }
                if ($cmd -and $cmd.Parameters.ContainsKey('AcceptLicense')) { $splat.AcceptLicense = $true }
                $legacyModules = @(Get-InstalledModule -EA SilentlyContinue | Where-Object { -not $psResourceNames.Contains($_.Name) })
                if ($legacyModules.Count -gt 0) {
                    $before = @{}
                    foreach ($module in $legacyModules) { $before[$module.Name] = [string]$module.Version }
                    $splat.Name = @($legacyModules | Select-Object -ExpandProperty Name)
                    try { Update-Module @splat 2>&1 | Out-Null } catch { Write-Verbose "Update-Module: $($_.Exception.Message)" }
                    foreach ($module in @(Get-InstalledModule -EA SilentlyContinue | Where-Object { $before.ContainsKey($_.Name) })) {
                        $prev = $before[$module.Name]; $curr = [string]$module.Version
                        if ($prev -and $curr -ne $prev) { $changed = $true; $details.Add("Module $($module.Name) $prev -> $curr") }
                    }
                } else {
                    Write-Detail 'Legacy PowerShellGet module pass skipped: PSResourceGet already covers installed modules.'
                }
            }

            if ($changed) { $script:stepChanged = $true; $script:stepMessage = if ($details.Count) { $details -join '; ' } else { 'PowerShell modules updated' } }
            else { $script:stepMessage = 'PowerShell modules checked' }
        }
    }
)

# ==============================================================================
# CLEANUP
# ==============================================================================

if ($SkipCleanup -or $UltraFast) {
    $reason = if ($SkipCleanup) { 'flag' } else { 'ultra fast mode' }
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cleanup'; Reason = $reason })
} elseif (-not $script:IsSimulation) {
    Write-Section 'System Cleanup'
    $cutoff = (Get-Date).AddDays(-$script:Config.TempCleanupDays)
    try { if ($env:TEMP -and (Test-Path $env:TEMP)) { Get-ChildItem $env:TEMP -EA SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Recurse -Force -EA SilentlyContinue; Write-Status "Temp files cleared (older than $($script:Config.TempCleanupDays) days)" -Type Success } } catch {}
    if ($isAdmin) {
        try { if (Test-Path 'C:\Windows\Temp') { Get-ChildItem 'C:\Windows\Temp' -EA SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff } | Remove-Item -Recurse -Force -EA SilentlyContinue; Write-Status "C:\Windows\Temp cleared (older than $($script:Config.TempCleanupDays) days)" -Type Success } } catch {}
    }
    try { Clear-DnsClientCache  -EA SilentlyContinue; Write-Status 'DNS cache flushed' -Type Success } catch {}
    try { Clear-RecycleBin -Force -EA SilentlyContinue; Write-Status 'Recycle Bin emptied' -Type Success } catch {}
    if ($isAdmin -and $DeepClean) {
        try { DISM /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-Null; Write-Status 'DISM component store cleaned' -Type Success } catch {}
        try { Clear-DeliveryOptimizationCache -Force -EA SilentlyContinue; Write-Status 'Delivery Optimization cache cleared' -Type Success } catch {}
        if (-not $SkipDestructive) {
            try { Get-ChildItem 'C:\Windows\Prefetch' -Filter '*.pf' -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue; Write-Status 'Prefetch files cleared' -Type Success } catch {}
        }
    }
    $updateResults.Checked.Add('cleanup')
    $updateResults.Details['cleanup'] = 'temp files, DNS cache, and recycle bin cleaned'
}

# ==============================================================================
# WHATCHANGED / STATE SAVE
# ==============================================================================

if (-not $script:IsSimulation -and $WhatChanged) {
    $prev = @{ Winget = ConvertTo-StringMap $script:State.Winget; Scoop = ConvertTo-StringMap $script:State.Scoop; Chocolatey = ConvertTo-StringMap $script:State.Chocolatey }
    Write-Host "`n[Summary] Updating package state maps..." -ForegroundColor DarkGray
    Update-WingetState; Update-ScoopState; Update-ChocolateyState
    Write-Section 'What changed since last run'
    foreach ($entry in @(
        @{ Name = 'Winget';     Prev = $prev.Winget;     Curr = $script:State.Winget }
        @{ Name = 'Scoop';      Prev = $prev.Scoop;      Curr = $script:State.Scoop }
        @{ Name = 'Chocolatey'; Prev = $prev.Chocolatey; Curr = $script:State.Chocolatey }
    )) {
        $changes = @(Compare-PackageMaps $entry.Prev $entry.Curr)
        if ($changes.Count -gt 0) { Write-Host "  $($entry.Name) changes:" -ForegroundColor Cyan; $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray } }
        else { Write-Host "  $($entry.Name): no changes detected" -ForegroundColor DarkGray }
    }
    Save-State
} elseif (-not $script:IsSimulation) {
    Save-State
}

# ==============================================================================
# SUMMARY
# ==============================================================================

$duration      = (Get-Date) - $startTime
$updatedNames  = @($updateResults.Success | Select-Object -Unique)
$checkedNames  = @($updateResults.Checked | Where-Object { $_ -notin $updatedNames -and $_ -notin $updateResults.Failed } | Select-Object -Unique)

Write-Host "`n$('=' * 54)" -ForegroundColor Green
Write-Host (" UPDATE COMPLETE -- {0}" -f $duration.ToString('hh\:mm\:ss')) -ForegroundColor Green
Write-Host "$('=' * 54)" -ForegroundColor Green

if ($updatedNames.Count -gt 0) {
    Write-Host "`n[OK] Updated   ($($updatedNames.Count))" -ForegroundColor Green
    foreach ($n in $updatedNames) { Write-Detail "$n`: $($updateResults.Details[$n])" }
}
if ($checkedNames.Count -gt 0) {
    Write-Host "[~] Checked   ($($checkedNames.Count))" -ForegroundColor DarkGray
    foreach ($n in $checkedNames) { Write-Detail "$n`: $($updateResults.Details[$n])" -Type Muted }
}
if ($updateResults.Failed.Count -gt 0) {
    Write-Host "[X] Failed    ($($updateResults.Failed.Count))" -ForegroundColor Red
    foreach ($n in $updateResults.Failed) { Write-Detail "$n`: $($updateResults.Details[$n])" -Type Error }
}
if ($updateResults.Skipped.Count -gt 0) {
    Write-Host "[!] Skipped   ($($updateResults.Skipped.Count))" -ForegroundColor Yellow
    foreach ($item in $updateResults.Skipped) { Write-Detail "$($item.Name) ($($item.Reason))" -Type Warning }
}

if ($script:SectionTimings.Count -gt 0) {
    Write-Host "`n  Section timings:" -ForegroundColor DarkGray
    $script:SectionTimings.GetEnumerator() | Sort-Object Value -Descending |
        ForEach-Object { Write-Host ("    {0,-30} {1,6}s" -f $_.Key, $_.Value.ToString('F1',[cultureinfo]::InvariantCulture)) -ForegroundColor DarkGray }
}

# --- Toast notification -------------------------------------------------------
try {
    if (Get-Module -ListAvailable BurntToast -EA SilentlyContinue) {
        Import-Module BurntToast -EA SilentlyContinue
        $msg = if ($updateResults.Failed.Count -gt 0) { "$($updatedNames.Count) updated, $($updateResults.Failed.Count) failed" }
               elseif ($updatedNames.Count -gt 0) { "$($updatedNames.Count) components updated" }
               else { 'No updates were needed' }
        New-BurntToastNotification -Text 'Update-Everything', $msg -EA SilentlyContinue
    }
} catch {}

if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }
exit 0
