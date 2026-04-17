<#
.SYNOPSIS
    Update Everything for Windows
.DESCRIPTION
    Updates package managers, Windows components, and common development toolchains.
.VERSION
    5.1.1
.NOTES
    Run as Administrator for full functionality. PowerShell 7 is recommended.
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
    .\updatescript.ps1 -WhatChanged -DryRun
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
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
}
catch { Write-Verbose "Console encoding update skipped: $($_.Exception.Message)" }

$ErrorActionPreference = 'Continue'
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')
$startTime = Get-Date
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
    WingetTimeoutSec        = 300
    WingetInstallRetryCount = 3
    MsiBusyWaitSec          = 600
    StateDir                = Join-Path $env:LOCALAPPDATA 'Update-Everything'
    LogRetentionDays        = 7
    LogMaxSizeMB            = 10
    TempCleanupDays         = 7
    PipAllowPackages        = @()
    PipSkipPackages         = @()
    PipIgnoreHealthPackages = @()
    WingetSkipPackages      = @()
    SkipManagers            = @()
    FastModeSkip            = @('Chocolatey', 'WSLDistros', 'npm', 'pnpm', 'bun', 'deno', 'rust', 'cargo-binaries',
        'go', 'gh-extensions', 'pipx', 'poetry', 'composer', 'rubygems', 'yt-dlp', 'tldr',
        'oh-my-posh', 'volta', 'fnm', 'mise', 'juliaup', 'ollama-models', 'git-lfs',
        'git-credential-manager', 'pip', 'uv', 'uv-tools', 'uv-python', 'dotnet',
        'dotnet-workloads', 'vscode-extensions', 'pwsh-resources')
    UltraFastSkip           = @('StoreApps', 'cleanup', 'WSL', 'WindowsUpdate', 'DefenderSignatures')
    # Pre/Post hooks: kill conflicting processes before upgrade, relaunch after
    WingetUpgradeHooks      = @{
        'Spotify.Spotify'             = @{
            Pre  = { $script:_spotifyWasRunning = [bool](Get-Process -Name Spotify -EA SilentlyContinue); Stop-Process -Name Spotify -Force -EA SilentlyContinue; Start-Sleep 1 }
            Post = { if ($script:_spotifyWasRunning) { Start-Process "$env:APPDATA\Spotify\Spotify.exe" -EA SilentlyContinue } }
        }
        'Google.Chrome'               = @{
            Pre  = {
                $script:_chromeWasRunning = [bool](Get-Process -Name chrome -EA SilentlyContinue)
                if ($script:_chromeWasRunning) { Write-Host '  Closing Chrome...' -ForegroundColor Gray }
                Stop-Process -Name 'chrome', 'chrome_crashpad_handler', 'GoogleUpdate', 'GoogleUpdateSetup', 'GoogleCrashHandler', 'GoogleCrashHandler64', 'GoogleUpdateComRegisterShell64' -Force -EA SilentlyContinue
                Start-Sleep 1
            }
            Post = { if ($script:_chromeWasRunning) { $p = @((Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'), (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'Google.QuickShare'           = @{
            Pre  = {
                Stop-Process -Name 'QuickShare', 'QuickShareAgent', 'NearbyShare', 'QuickShareService' -Force -EA SilentlyContinue
                Get-Service -EA SilentlyContinue | Where-Object { $_.DisplayName -match 'Quick Share|Nearby Share|NearShare' } | ForEach-Object { Stop-Service $_.Name -Force -EA SilentlyContinue }
                Start-Sleep 2
            }
            Post = {}
        }
        'Microsoft.VisualStudioCode'  = @{
            Pre  = { $script:_vscodeWasRunning = [bool](Get-Process -Name Code -EA SilentlyContinue); if ($script:_vscodeWasRunning) { Write-Host '  Closing VS Code...' -ForegroundColor Gray; Stop-Process -Name Code -Force -EA SilentlyContinue; Start-Sleep 1 } }
            Post = { if ($script:_vscodeWasRunning) { $p = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'; if (Test-Path $p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'Foxit.FoxitReader'           = @{
            Pre  = { $script:_foxitWasRunning = [bool](Get-Process -Name 'FoxitPDFReader', 'FoxitReader' -EA SilentlyContinue); if ($script:_foxitWasRunning) { Write-Host '  Closing Foxit Reader...' -ForegroundColor Gray; Stop-Process -Name 'FoxitPDFReader', 'FoxitReader' -Force -EA SilentlyContinue; Start-Sleep 1 } }
            Post = { if ($script:_foxitWasRunning) { $p = @((Join-Path $env:ProgramFiles 'Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe'), (Join-Path ${env:ProgramFiles(x86)} 'Foxit Software\Foxit Reader\FoxitReader.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
        'GoLang.Go'                   = @{
            Pre  = { Stop-Process -Name 'go', 'gopls', 'dlv', 'golangci-lint' -Force -EA SilentlyContinue; Start-Sleep 1 }
            Post = {}
        }
        'Adobe.Acrobat.Reader.64-bit' = @{
            Pre  = {
                $script:_acrobatProcs = @(Get-Process -Name 'Acrobat', 'AcroRd32', 'AcroCEF' -EA SilentlyContinue)
                if ($script:_acrobatProcs.Count -gt 0) { Write-Host '  Closing Adobe Acrobat...' -ForegroundColor Gray; $script:_acrobatProcs | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 1 }
                Write-Host '  Skipping broad TEMP cleanup to avoid deleting unrelated app files...' -ForegroundColor Gray
            }
            Post = { if ($script:_acrobatProcs.Count -gt 0) { $p = @((Join-Path $env:ProgramFiles 'Adobe\Acrobat DC\Acrobat\Acrobat.exe'), (Join-Path ${env:ProgramFiles(x86)} 'Adobe\Acrobat Reader DC\Reader\AcroRd32.exe')) | Where-Object { Test-Path $_ } | Select-Object -First 1; if ($p) { Start-Process $p -EA SilentlyContinue } } }
        }
    }
}

$configFile = Join-Path $PSScriptRoot 'update-config.json'
if (Test-Path $configFile) {
    try {
        $userConfig = Get-Content $configFile -Raw | ConvertFrom-Json
        foreach ($prop in $userConfig.PSObject.Properties) { $script:Config[$prop.Name] = $prop.Value }
        Write-Verbose "Loaded config from $configFile"
    }
    catch { Write-Warning "Failed to load config file: $_" }
}

# Force skip lists empty regardless of defaults or update-config.json: update everything.
$script:Config.WingetSkipPackages = @()
$script:Config.PipSkipPackages = @()
$script:Config.SkipManagers = @()

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
    try { Start-Process -FilePath $shell -Verb RunAs -ArgumentList (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + @($forwardedArgs)) -Wait; exit }
    catch { Write-Warning 'Could not elevate. Continuing without Administrator privileges.' }
}
elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host 'INFO: Running without elevation. Admin-only tasks may be skipped.' -ForegroundColor DarkYellow
}

if ($Schedule) {
    if (-not $isAdmin -and -not $script:IsSimulation) { throw 'Scheduled task registration requires Administrator.' }
    $pwshCommand = Get-Command pwsh.exe -EA SilentlyContinue
    if ($pwshCommand -and $pwshCommand.Source) { $shell = $pwshCommand.Source } else { $shell = 'powershell.exe' }
    $taskArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-SkipReboot', '-NoPause', '-SkipWSL', '-SkipWindowsUpdate')
    if ($script:IsSimulation) { Write-Host "  [DryRun] Would register 'DailySystemUpdate' at $ScheduleTime" -ForegroundColor DarkCyan }
    elseif ($PSCmdlet.ShouldProcess('DailySystemUpdate', "Register scheduled task for $ScheduleTime")) {
        Register-ScheduledTask -TaskName 'DailySystemUpdate' -Force `
            -Action  (New-ScheduledTaskAction  -Execute $shell -Argument ($taskArgs -join ' ')) `
            -Trigger (New-ScheduledTaskTrigger -Daily -At $ScheduleTime) `
            -Settings(New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)) `
            -Principal (New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest) | Out-Null
        Write-Host "[OK] Scheduled 'DailySystemUpdate' daily at $ScheduleTime (runs elevated, skips WSL + Windows Update)." -ForegroundColor Green
    }
    exit
}

if ($WingetTimeoutSec -ne 300) { $script:Config.WingetTimeoutSec = $WingetTimeoutSec }
if ($script:Config.WingetTimeoutSec -lt 30) { $script:Config.WingetTimeoutSec = 30 }
if ($SkipUVTools) { $script:Config.SkipUVTools = $true }
$script:Config.SkipWSLDistros = [bool]$SkipWSLDistros.IsPresent
if ($UltraFast) { $FastMode = $true }
if (-not $PSBoundParameters.ContainsKey('ParallelThrottle')) {
    $ParallelThrottle = [Math]::Max(4, [Math]::Min([Environment]::ProcessorCount, 8))
}

# --- Helper Functions ----------------------------------------------------------
$script:commandCache = @{}
function Test-Command([string]$Command) {
    if (-not $script:commandCache) { $script:commandCache = @{} }
    if ($script:commandCache.ContainsKey($Command)) { return $script:commandCache[$Command] }
    $saved = $WhatIfPreference; $WhatIfPreference = $false
    try { $result = [bool](Get-Command $Command -EA SilentlyContinue) }
    finally { $WhatIfPreference = $saved }
    $script:commandCache[$Command] = $result
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

function ConvertTo-NormalizedPackageName([string]$Name) {
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

function Test-HardRebootPending {
    return (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
    (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
}

# Check if Windows Installer (msiexec) is currently running an install
function Test-MsiInstallInProgress {
    try {
        $mutex = [System.Threading.Mutex]::OpenExisting('Global\_MSIExecute')
        if ($mutex) {
            $mutex.Dispose()
            return $true
        }
    }
    catch {
        return $false
    }
    return $false
}

function Wait-WindowsInstallerIdle {
    param(
        [int]$TimeoutSec = 600,
        [int]$PollSec = 5,
        [string]$Reason = 'before package install'
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (-not (Test-MsiInstallInProgress)) { return $true }
        Write-Detail "Windows Installer is busy ($Reason). Waiting ${PollSec}s..." -Type Muted
        Start-Sleep -Seconds $PollSec
    }
    return (-not (Test-MsiInstallInProgress))
}

# Extract installer log file path from winget output
function Get-WingetInstallerLogPath([string]$WingetOutput) {
    if ([string]::IsNullOrWhiteSpace($WingetOutput)) { return $null }
    $match = [regex]::Match($WingetOutput, 'Installer log is available at:\s*(.+)$', [System.Text.RegularExpressions.RegexOptions]::Multiline)
    if (-not $match.Success) { return $null }
    $path = $match.Groups[1].Value.Trim()
    if ($path -and (Test-Path -LiteralPath $path)) { return $path }
    return $null
}

# Detect MSI "source not found" failure (error 1714 + 1612 in installer log)
function Test-WingetMissingMsiSourceFailure {
    param(
        [AllowNull()][string]$WingetOutput,
        [AllowNull()][string]$LogPath
    )
    if ($WingetOutput -match 'Installer failed with exit code:\s*1603') {
        if ($LogPath -and (Test-Path -LiteralPath $LogPath)) {
            $logText = Get-Content -LiteralPath $LogPath -Raw -EA SilentlyContinue
            if ($logText -match 'Error 1714' -and $logText -match 'System Error 1612') { return $true }
        }
    }
    return $false
}

# Look up a program's registry uninstall entry by matching winget package metadata
function Get-InstalledUninstallEntryForWinget {
    param([Parameter(Mandatory)]$Entry)
    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $candidates = @(Get-ItemProperty $roots -EA SilentlyContinue | Where-Object {
            $_.DisplayName -and (
                $_.DisplayName -eq $Entry.Name -or
                $_.DisplayName -like "$($Entry.Name)*" -or
                ($Entry.Id -eq 'Google.Chrome' -and $_.DisplayName -eq 'Google Chrome') -or
                ($Entry.Id -eq 'Google.QuickShare' -and $_.DisplayName -eq 'Quick Share') -or
                ($Entry.Id -eq 'GoLang.Go' -and $_.DisplayName -like 'Go Programming Language*')
            )
        })
    if ($Entry.Version) {
        $exact = @($candidates | Where-Object { [string]$_.DisplayVersion -eq [string]$Entry.Version })
        if ($exact.Count -gt 0) { return $exact[0] }
    }
    return @($candidates | Select-Object -First 1)[0]
}

# Download and recache MSI installer to fix broken Windows Installer source references
function Invoke-WingetMsiSourceRepair {
    param([Parameter(Mandatory)]$Entry)

    $installed = Get-InstalledUninstallEntryForWinget -Entry $Entry
    if (-not $installed) { return $false }

    $version = if ($installed.DisplayVersion) { [string]$installed.DisplayVersion } else { [string]$Entry.Version }
    if ([string]::IsNullOrWhiteSpace($version)) { return $false }

    $downloadRoot = Join-Path $script:Config.StateDir 'msi-repair-cache'
    $targetDir = Join-Path $downloadRoot ("{0}-{1}" -f $Entry.Id, $version)
    New-Item -ItemType Directory -Path $targetDir -Force -EA SilentlyContinue | Out-Null

    $msi = @(Get-ChildItem -Path $targetDir -Recurse -Filter *.msi -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
    if (-not $msi) {
        Write-Detail "$($Entry.Name): downloading cached installer for version $version" -Type Warning
        $downloadArgs = @(
            'download', '--id', $Entry.Id, '--exact', '--source', 'winget', '--version', $version,
            '--download-directory', $targetDir,
            '--accept-source-agreements', '--disable-interactivity'
        )
        $download = Invoke-ProcessWithHeartbeat -FilePath 'winget' -ArgumentList $downloadArgs -TimeoutSec ([Math]::Max($script:Config.WingetTimeoutSec, 180)) -HeartbeatMessage "$($Entry.Name): still downloading cached installer"
        if ($download.Output) { Write-FilteredOutput -Text $download.Output -Color Gray }
        if ($download.ExitCode -ne 0) {
            Write-Detail "$($Entry.Name): installer download failed during MSI source repair" -Type Warning
            return $false
        }

        $msi = @(Get-ChildItem -Path $targetDir -Recurse -Filter *.msi -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
    }
    else {
        Write-Detail "$($Entry.Name): reusing cached installer $($msi.Name)" -Type Muted
    }
    if (-not $msi) { return $false }

    Write-Detail "$($Entry.Name): recaching MSI metadata from $($msi.Name)" -Type Warning
    $repairLog = Join-Path $targetDir 'msi-recache.log'
    $repair = Invoke-ProcessWithHeartbeat -FilePath 'msiexec.exe' -ArgumentList @('/fvomus', "`"$($msi.FullName)`"", '/qn', '/L*v', "`"$repairLog`"") -TimeoutSec ([Math]::Max($script:Config.WingetTimeoutSec, 180)) -HeartbeatMessage "$($Entry.Name): still recaching MSI metadata"
    if ($repair.Output) { Write-FilteredOutput -Text $repair.Output -Color Gray }
    if ($repair.ExitCode -eq 0) {
        Write-Detail "$($Entry.Name): MSI cache repaired" -Type Warning
    }
    else {
        Write-Detail "$($Entry.Name): MSI recache exited with code $($repair.ExitCode)" -Type Warning
    }
    return ($repair.ExitCode -eq 0)
}

# --- MSI Repair Failure Cache (skip known-broken repairs for 14 days) --------
function Get-MsiRepairFailureStatePath {
    return (Join-Path $script:Config.StateDir 'msi-repair-failures.json')
}

function Get-MsiRepairFailureState {
    if ($script:MsiRepairFailureState) { return $script:MsiRepairFailureState }
    $script:MsiRepairFailureState = @{}
    $statePath = Get-MsiRepairFailureStatePath
    if (Test-Path -LiteralPath $statePath) {
        try {
            $raw = Get-Content -LiteralPath $statePath -Raw -EA SilentlyContinue | ConvertFrom-Json
            if ($raw) {
                foreach ($prop in $raw.PSObject.Properties) {
                    $script:MsiRepairFailureState[[string]$prop.Name] = [string]$prop.Value
                }
            }
        }
        catch {
            Write-Verbose "Could not parse MSI repair failure state: $($_.Exception.Message)"
        }
    }
    return $script:MsiRepairFailureState
}

function Save-MsiRepairFailureState {
    $statePath = Get-MsiRepairFailureStatePath
    $dir = Split-Path -Path $statePath -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force -EA SilentlyContinue | Out-Null
    }
    try {
        (Get-MsiRepairFailureState) | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $statePath -Encoding UTF8
    }
    catch {
        Write-Verbose "Could not save MSI repair failure state: $($_.Exception.Message)"
    }
}

function Get-MsiRepairFailureKey {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Version)
    return ("{0}|{1}" -f [string]$Entry.Id, [string]$Version)
}

function Test-MsiRepairFailureCached {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Version, [int]$MaxAgeDays = 14)
    $state = Get-MsiRepairFailureState
    $key = Get-MsiRepairFailureKey -Entry $Entry -Version $Version
    if (-not $state.ContainsKey($key)) { return $false }
    [datetime]$when = [datetime]::MinValue
    if (-not [datetime]::TryParse([string]$state[$key], [ref]$when)) { return $false }
    return ((Get-Date) - $when).TotalDays -lt $MaxAgeDays
}

function Set-MsiRepairFailureCache {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Version)
    $state = Get-MsiRepairFailureState
    $state[(Get-MsiRepairFailureKey -Entry $Entry -Version $Version)] = (Get-Date).ToString('o')
    Save-MsiRepairFailureState
}

function Clear-MsiRepairFailureCache {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)][string]$Version)
    $state = Get-MsiRepairFailureState
    $key = Get-MsiRepairFailureKey -Entry $Entry -Version $Version
    if ($state.ContainsKey($key)) {
        [void]$state.Remove($key)
        Save-MsiRepairFailureState
    }
}

# Convert a standard GUID to the Windows Installer packed GUID used in Classes\Installer registry keys.
# Reverses chars in Data1/2/3, swaps pairs in Data4/5.
function ConvertTo-MsiPackedGuid([string]$Guid) {
    $g = $Guid.Trim('{}').ToUpper()
    if ($g.Length -ne 36) { return $null }
    -join ($g[7], $g[6], $g[5], $g[4], $g[3], $g[2], $g[1], $g[0],
        $g[12], $g[11], $g[10], $g[9],
        $g[17], $g[16], $g[15], $g[14],
        $g[20], $g[19], $g[22], $g[21],
        $g[25], $g[24], $g[27], $g[26], $g[29], $g[28], $g[31], $g[30], $g[33], $g[32], $g[35], $g[34])
}

# Remove every Windows Installer registry trace of a broken package so the new installer
# does not hit RemoveExistingProducts → 1714/1612 failures.
# Cleans: Uninstall (HKLM x2 + HKCU), Classes\Installer\Products, \Features, UserData\*\Products
# Falls back to searching the Products hive by ProductName when Guid is absent or already removed.
function Invoke-MsiFullRegistryCleanse {
    param([string]$Guid, [string]$ProductName)
    $packedGuid = $null
    $canonGuid = $null
    if ($Guid) {
        $canonGuid = if ($Guid -match '^\{') { $Guid } else { "{$Guid}" }
        $packedGuid = ConvertTo-MsiPackedGuid $canonGuid
    }
    if (-not $packedGuid -and $ProductName) {
        $found = Get-ChildItem 'HKLM:\SOFTWARE\Classes\Installer\Products' -EA SilentlyContinue |
        Where-Object { (Get-ItemProperty $_.PSPath -Name ProductName -EA SilentlyContinue).ProductName -eq $ProductName } |
        Select-Object -First 1
        if ($found) { $packedGuid = $found.PSChildName }
    }
    $removed = 0
    if ($canonGuid) {
        foreach ($rp in @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$canonGuid",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$canonGuid",
                "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$canonGuid"
            )) { if (Test-Path -LiteralPath $rp -EA SilentlyContinue) { Remove-Item -LiteralPath $rp -Recurse -Force -EA SilentlyContinue; $removed++ } }
    }
    if ($packedGuid) {
        foreach ($rp in @(
                "HKLM:\SOFTWARE\Classes\Installer\Products\$packedGuid",
                "HKLM:\SOFTWARE\Classes\Installer\Features\$packedGuid"
            )) { if (Test-Path -LiteralPath $rp -EA SilentlyContinue) { Remove-Item -LiteralPath $rp -Recurse -Force -EA SilentlyContinue; $removed++ } }
        Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData' -EA SilentlyContinue | ForEach-Object {
            $udp = "$($_.PSPath)\Products\$packedGuid"
            if (Test-Path -LiteralPath $udp -EA SilentlyContinue) { Remove-Item -LiteralPath $udp -Recurse -Force -EA SilentlyContinue; $removed++ }
        }
    }
    return $removed
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
    foreach ($name in 'code.cmd', 'code-insiders.cmd') {
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
$script:LogFile = $null
$script:LoggingEnabled = $false

function Initialize-Logging {
    $preferredLog = if ($LogPath) { $LogPath } else { Join-Path $script:Config.StateDir 'updatescript.log' }
    $fallbackLogs = @(
        $preferredLog
        (Join-Path $PSScriptRoot 'updatescript.log')
        (Join-Path ([System.IO.Path]::GetTempPath()) 'updatescript.log')
    ) | Where-Object { $_ } | Select-Object -Unique

    foreach ($candidate in $fallbackLogs) {
        try {
            $logDir = Split-Path $candidate -Parent
            if ($logDir -and -not (Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false -ErrorAction Stop | Out-Null
            }
            if ((Test-Path $candidate) -and (Get-Item $candidate -ErrorAction Stop).Length -gt ($script:Config.LogMaxSizeMB * 1MB)) {
                Rename-Item $candidate "$candidate.old" -Force -EA SilentlyContinue -WhatIf:$false
            }

            $stream = [System.IO.File]::Open($candidate, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $stream.Dispose()

            $script:LogFile = $candidate
            $script:LoggingEnabled = $true
            return
        }
        catch {
            $script:LoggingEnabled = $false
        }
    }
}

Initialize-Logging

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    if (-not $script:LoggingEnabled -or [string]::IsNullOrWhiteSpace($script:LogFile)) { return }
    try {
        Add-Content -Path $script:LogFile -Value "[$( Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" -Encoding UTF8 -WhatIf:$false -ErrorAction Stop
    }
    catch {
        $script:LoggingEnabled = $false
        Write-Verbose "Logging disabled after write failure: $($_.Exception.Message)"
    }
}

function Write-Section([string]$Title) {
    Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
    Write-Host "  $Title"      -ForegroundColor Cyan
    Write-Host "$('=' * 54)"   -ForegroundColor DarkGray
    Write-Log "--- $Title ---"
}

function Write-Status {
    param([string]$Message, [ValidateSet('Success', 'Warning', 'Error', 'Info')][string]$Type = 'Info')
    $colors = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]'; Warning = '[!]'; Error = '[X]'; Info = '[*]' }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
    Write-Log $Message -Level $Type
}

function Write-Detail {
    param([string]$Message, [ValidateSet('Info', 'Muted', 'Warning', 'Error')][string]$Type = 'Info')
    $colors = @{ Info = 'Gray'; Muted = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }
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
    foreach ($rawLine in ($normalized -split '\r\n|\r|\n')) {
        $line = $rawLine.TrimEnd()
        if (-not $line) { continue }
        $compact = $line.Trim()
        if ($compact -match '^[\\/\|\-]+$') { continue }
        if ([regex]::Matches($compact, '[^\x00-\x7F]').Count -ge 3 -and $compact -notmatch '[A-Za-z0-9]') { continue }
        if ($compact -eq 'System.__ComObject') { continue }
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }
        if ($compact -match 'package\(s\) have pins that prevent upgrade') { continue }
        if ($compact -match '^Installer failed with exit code:\s*1603$') { continue }
        if ($compact -match '^Installer log is available at:\s+') { continue }
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
    }
    catch {
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
        $proc = Start-Process winget -ArgumentList $Arguments -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { taskkill.exe /PID $proc.Id /T /F | Out-Null } catch {}
            try { $proc.Kill() } catch {}
            throw "winget timed out after ${TimeoutSec}s"
        }
        $stdout = Get-Content -Raw -Path $stdoutFile -Encoding UTF8 -EA SilentlyContinue
        $stderr = Get-Content -Raw -Path $stderrFile -Encoding UTF8 -EA SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()
        $combined = (($combined -split '\r\n|\r|\n') | Where-Object { $_ -notmatch '^[\s\p{S}\p{P}]*\d+%\s*$' }) -join "`n"  # strip progress lines
        return [pscustomobject]@{ Output = $combined; ExitCode = $proc.ExitCode }
    }
    finally {
        Remove-Item $stdoutFile, $stderrFile -Force -EA SilentlyContinue
    }
}

function Invoke-ProcessWithHeartbeat {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$ArgumentList,
        [int]$TimeoutSec = 300,
        [string]$HeartbeatMessage = 'Still running...',
        [int]$HeartbeatSec = 5
    )

    $stdoutFile = [System.IO.Path]::GetTempFileName()
    $stderrFile = [System.IO.Path]::GetTempFileName()
    $streamStates = @(
        [pscustomobject]@{ Path = $stdoutFile; Position = 0L }
        [pscustomobject]@{ Path = $stderrFile; Position = 0L }
    )
    try {
        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile -NoNewWindow -PassThru
        $start = Get-Date
        $nextHeartbeat = $start.AddSeconds($HeartbeatSec)

        while (-not $proc.HasExited) {
            Start-Sleep -Milliseconds 500

            foreach ($stream in $streamStates) {
                if (Test-Path -LiteralPath $stream.Path) {
                    $currentLength = (Get-Item -LiteralPath $stream.Path -EA SilentlyContinue).Length
                    if ($currentLength -gt $stream.Position) {
                        $fs = $null
                        $reader = $null
                        try {
                            $fs = [System.IO.File]::Open($stream.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                            $fs.Seek($stream.Position, [System.IO.SeekOrigin]::Begin) | Out-Null
                            $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8, $true, 1024, $true)
                            $chunk = $reader.ReadToEnd()
                            if ($chunk) { Write-FilteredOutput -Text $chunk -Color Gray }
                        }
                        catch {} finally {
                            if ($reader) { $reader.Dispose() }
                            if ($fs) { $fs.Dispose() }
                        }
                        $stream.Position = $currentLength
                    }
                }
            }

            if ((Get-Date) -ge $nextHeartbeat) {
                $elapsed = [int](((Get-Date) - $start).TotalSeconds)
                Write-Detail "$HeartbeatMessage (${elapsed}s elapsed)" -Type Muted
                $nextHeartbeat = (Get-Date).AddSeconds($HeartbeatSec)
            }

            if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSec) {
                try { taskkill.exe /PID $proc.Id /T /F | Out-Null } catch {}
                try { $proc.Kill() } catch {}
                throw "$FilePath timed out after ${TimeoutSec}s"
            }
        }

        $stdout = if (Test-Path -LiteralPath $stdoutFile) { Get-Content -LiteralPath $stdoutFile -Raw -EA SilentlyContinue } else { '' }
        $stderr = if (Test-Path -LiteralPath $stderrFile) { Get-Content -LiteralPath $stderrFile -Raw -EA SilentlyContinue } else { '' }
        return [pscustomobject]@{
            ExitCode = $proc.ExitCode
            Output   = (($stdout + $stderr) -replace '\x00', '').Trim()
        }
    }
    finally {
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
        if ($t -match '^(?<name>.+?)\s+(?<id>(?=.*[A-Za-z])[A-Za-z0-9][A-Za-z0-9\.\-_]*\.[A-Za-z0-9\.\-_]+)\s+(?<version>(?:<\s+)?\S+)\s+(?<available>\S+)$') {
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


# --- State Management ----------------------------------------------------------
$script:StateDir = $script:Config.StateDir
$script:StateFile = Join-Path $script:StateDir 'state.json'
$script:State = @{ LastRun = $null; Winget = @{}; Scoop = @{}; Chocolatey = @{}; WhatChanged = $null }

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
    }
    catch { Write-Warning "Could not load state: $_" }
}

function Save-State {
    $script:State.LastRun = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    try {
        $dir = Split-Path $script:StateFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $script:State | ConvertTo-Json -Depth 10 | Set-Content $script:StateFile -Encoding UTF8
    }
    catch { Write-Warning "Failed to save state: $_" }
}

function Update-WingetState {
    try {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) "winget-state-$([guid]::NewGuid().ToString('N')).json"
        winget export -o $tmp --include-versions --accept-source-agreements --disable-interactivity | Out-Null
        $data = Get-Content $tmp -Raw | ConvertFrom-Json
        $map = @{}
        foreach ($src in @($data.Sources)) {
            foreach ($pkg in @($src.Packages)) {
                if ($pkg.PackageIdentifier -and $pkg.Version) { $map[[string]$pkg.PackageIdentifier] = [string]$pkg.Version }
            }
        }
        $script:State.Winget = $map
    }
    catch { Write-Warning "Could not update winget state: $_" }
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
    }
    catch { Write-Warning "Could not update scoop state: $_" }
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
    }
    catch { Write-Warning "Could not update chocolatey state: $_" }
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
$script:stepChanged = $false
$script:stepMessage = ''

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
    if ($Disabled) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'flag' }); return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'not installed' }); return }
    if ($RequiresAdmin -and -not $isAdmin) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = 'requires admin' }); return }
    if ($SlowOperation -and $FastMode) { $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = if ($UltraFast) { 'ultra fast mode' } else { 'fast mode' } }); return }
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
        }
        else {
            if ($script:stepMessage) { Write-Detail $script:stepMessage -Type Muted }
            $updateResults.Checked.Add($Name)
        }
        $updateResults.Details[$Name] = $script:stepMessage
        $script:SectionTimings[$Name] = $elapsed
    }
    catch {
        Write-Status "$Name failed: $($_.Exception.Message)" -Type Error
        $updateResults.Failed.Add($Name)
        $updateResults.Details[$Name] = $_.Exception.Message
    }
}

# --- Parallel Batch Execution --------------------------------------------------
# Jobs run via Start-ThreadJob; Write-Host is shadowed per-job to a temp file
# so output replays sequentially after all jobs complete.
$script:BatchInitScript = [scriptblock]::Create((@(
            'function Write-Host { param($Object = '''', $ForegroundColor, $BackgroundColor, [switch]$NoNewline); $c = if ($ForegroundColor -is [System.ConsoleColor]) { [int]$ForegroundColor } elseif ($ForegroundColor) { try { [int][System.ConsoleColor]$ForegroundColor } catch { 7 } } else { 7 }; if ($script:_batchOutFile) { [System.IO.File]::AppendAllText($script:_batchOutFile, "$c`t$Object`n", [System.Text.Encoding]::UTF8) } }'
        ) + @(
            @('Write-FilteredOutput', 'Write-Detail', 'Write-Status', 'Write-Log', 'Invoke-StreamingCapture',
                'Read-CapturedOutput', 'Invoke-WingetWithTimeout', 'Get-ToolInstallManager', 'Test-Command',
                'ConvertTo-NormalizedPackageName', 'Complete-StepState', 'Get-VSCodeCliPath') | ForEach-Object {
                $item = Get-Item "Function:$_" -EA SilentlyContinue
                if ($item) { "function $_ {`n$($item.ScriptBlock)`n}" }
            }
        )) -join "`n`n")

function Invoke-UpdateBatch {
    param(
        [Parameter(Mandatory)][hashtable[]]$Tasks,
        [int]$ThrottleLimit = 4,
        [int]$JobTimeoutSec = 600
    )

    if ($script:IsSimulation) {
        foreach ($t in $Tasks) { Write-Host "  [DryRun] Would run: $(if ($t.Title) { $t.Title } else { $t.Name })" -ForegroundColor DarkCyan }
        return
    }

    function Write-UpdateBatchResult {
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
                }
                else { Write-Host $raw }
            }
            Remove-Item $Info.OutFile -Force -EA SilentlyContinue
        }

        if (-not $Result) {
            Write-Status "$($Info.Task.Name): no result returned" -Type Warning
            $updateResults.Checked.Add($Info.Task.Name)
        }
        elseif ($Result.Error) {
            Write-Status "$($Result.Name) failed: $($Result.Error)" -Type Error
            $updateResults.Failed.Add($Result.Name)
            $updateResults.Details[$Result.Name] = $Result.Error
        }
        elseif ($Result.Changed) {
            $e = [math]::Round($Result.Elapsed, 1).ToString('F1', [cultureinfo]::InvariantCulture)
            Write-Status "$($Result.Name) updated (${e}s)" -Type Success
            $updateResults.Success.Add($Result.Name)
            $updateResults.Details[$Result.Name] = $Result.Message
        }
        else {
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
        $fastSkip = $FastMode -and $script:Config.FastModeSkip -contains $t.Name
        $ultraSkip = $UltraFast -and $script:Config.UltraFastSkip -contains $t.Name
        $skipped = $disabled -or $missingCmd -or $missingAny -or $needsAdmin -or $fastSkip -or $ultraSkip
        if ($skipped) {
            $reason = if ($disabled) { 'flag' } elseif ($ultraSkip) { 'ultra fast mode' } elseif ($fastSkip) { if ($UltraFast) { 'ultra fast mode' } else { 'fast mode' } } elseif ($needsAdmin) { 'requires admin' } else { 'not installed' }
            $updateResults.Skipped.Add([pscustomobject]@{ Name = $t.Name; Reason = $reason })
        }
        else { $t }
    }
    if (-not $active) { return }

    try {
        $initScript = $script:BatchInitScript
        $workerScript = {
            param($task, $outFile, $logFile, $cfg)
            $script:LogFile = $logFile
            $script:Config = $cfg
            $script:commandCache = @{}

            $script:_batchOutFile = $outFile

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
                Write-UpdateBatchResult -Info $info -Result $result
            }
            return
        }

        # Launch all jobs
        $jobInfos = foreach ($task in $active) {
            $outFile = [System.IO.Path]::GetTempFileName()
            $logFile = $script:LogFile
            $cfg = $script:Config

            $job = Start-ThreadJob -ThrottleLimit $ThrottleLimit -InitializationScript $initScript -ArgumentList $task, $outFile, $logFile, $cfg -ScriptBlock $workerScript
            [pscustomobject]@{ Job = $job; OutFile = $outFile; Task = $task }
        }

        # Wait for all jobs with a timeout so a hung manager can't stall the script
        $null = $jobInfos.Job | Wait-Job -Timeout $JobTimeoutSec

        foreach ($info in $jobInfos) {
            if ($info.Job.State -eq 'Running') {
                # Job exceeded timeout — force-stop and report
                Stop-Job $info.Job -EA SilentlyContinue
                $result = [pscustomobject]@{
                    Name    = $info.Task.Name
                    Title   = if ($info.Task.Title) { $info.Task.Title } else { $info.Task.Name }
                    Changed = $false
                    Message = ''
                    Error   = "timed out after ${JobTimeoutSec}s"
                    Elapsed = $JobTimeoutSec
                }
            }
            else {
                $result = Receive-Job $info.Job -EA SilentlyContinue
            }
            Remove-Job $info.Job -Force
            Write-UpdateBatchResult -Info $info -Result $result
        }
    }
    catch {
        # Mark un-processed tasks as failed so the script can continue
        foreach ($t in $active) {
            if ($t.Name -notin $updateResults.Success -and $t.Name -notin $updateResults.Checked -and $t.Name -notin $updateResults.Failed) {
                $updateResults.Failed.Add($t.Name)
                $updateResults.Details[$t.Name] = "batch error: $($_.Exception.Message)"
            }
        }
        Write-Status "Parallel batch failed: $($_.Exception.Message)" -Type Error
    }
}

# ==============================================================================
# PACKAGE MANAGERS — lightweight managers run in parallel, Winget sequential
# ==============================================================================

Write-Host "`n[Parallel] Running lightweight package manager updates..." -ForegroundColor DarkCyan

Invoke-UpdateBatch -ThrottleLimit 3 -Tasks @(
    @{
        Name            = 'Scoop'
        RequiresCommand = 'scoop'
        Disabled        = $script:Config.SkipManagers -contains 'Scoop'
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { scoop update }).OutputPath
            if ($out -match '(?i)error|failed') { Write-Status "Scoop self-update warning: $out" -Type Warning }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { scoop update '*' }).OutputPath
            if ($out -and $out -notmatch 'Scoop is up to date') { $script:stepChanged = $true; $script:stepMessage = 'updated correctly' }
            scoop cleanup '*' 2>&1 | Out-Null
            scoop cache rm '*' 2>&1 | Out-Null
        }
    }
    @{
        Name            = 'Chocolatey'
        RequiresCommand = 'choco'
        RequiresAdmin   = $true
        Disabled        = $script:Config.SkipManagers -contains 'Chocolatey'
        Action          = {
            $run = Invoke-StreamingCapture { choco upgrade all -y --exclude-prerelease --no-progress }
            $out = Read-CapturedOutput $run.OutputPath
            if ($run.ExitCode -ne 0) {
                $logPath = "$env:ProgramData\chocolatey\logs\chocolatey.log"
                $logTail = if (Test-Path $logPath) { ((Get-Content $logPath -Tail 20 -ErrorAction SilentlyContinue) -join "`n").Trim() } else { '' }
                $detail = if ($logTail) { $logTail } else { $out }
                if ($detail) { Write-Detail "Chocolatey errors detected:`n$detail" -Type Warning }
                $script:stepMessage = 'Chocolatey reported errors (see output/log)'
            }
            elseif ($out -match 'upgraded 0/') {
                $script:stepMessage = 'no package updates'
            }
            else { $script:stepChanged = $true; $script:stepMessage = 'packages upgraded' }
        }
    }
    @{
        Name            = 'WSL'
        Title           = 'Windows Subsystem for Linux'
        RequiresCommand = 'wsl'
        RequiresAdmin   = $true
        Disabled        = $SkipWSL -or $script:Config.SkipManagers -contains 'WSL'
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture {
                    $saved = [Console]::OutputEncoding
                    [Console]::OutputEncoding = [System.Text.Encoding]::Unicode
                    try { wsl --update } finally { [Console]::OutputEncoding = $saved }
                }).OutputPath
            $script:stepChanged = $out -and $out -notmatch 'most recent version.+already installed'
            $script:stepMessage = if ($script:stepChanged) { 'WSL platform updated' } else { 'platform already up to date' }
            if (-not $script:Config.SkipWSLDistros) {
                $distros = @(wsl --list --quiet 2>&1 | ForEach-Object { ($_ -replace '\x00', '').Trim() } | Where-Object { $_ })
                if ($distros.Count -eq 0) { $script:stepMessage = 'platform current; no distros found' }
                else {
                    $hadDistroWarnings = $false
                    $skippedBrokenDistros = [System.Collections.Generic.List[string]]::new()
                    foreach ($d in $distros) {
                        Write-Detail "Updating WSL Distro: $d"
                        $distroOut = wsl -d $d -- sh -lc 'if command -v apt-get >/dev/null; then sudo apt-get update -qq && sudo apt-get upgrade -y -qq; elif command -v pacman >/dev/null; then sudo pacman -Syu --noconfirm; fi' 2>&1
                        $distroText = ((@($distroOut) | ForEach-Object { ($_.ToString() -replace '\x00', '').TrimEnd() } | Where-Object { $_ }) -join "`n").Trim()
                        $missingDistroPath = $distroText -match '(?im)\bERROR_PATH_NOT_FOUND\b|The system cannot find the path specified'
                        $distroFailed = ($LASTEXITCODE -ne 0) -or ($distroText -match '(?im)\b(error|failed)\b')
                        if ($missingDistroPath) {
                            $skippedBrokenDistros.Add($d)
                            Write-Detail "  ${d}: skipped (WSL distro storage path is missing)" -Type Muted
                        }
                        elseif ($distroFailed) {
                            $hadDistroWarnings = $true
                            Write-Detail "  ${d}: update error: $distroText" -Type Warning
                        }
                        elseif ($distroText) {
                            Write-Detail "  ${d}: updated or already current"
                        }
                        else {
                            Write-Detail "  ${d}: update completed"
                        }
                    }
                    $distroNotes = [System.Collections.Generic.List[string]]::new()
                    if ($skippedBrokenDistros.Count -gt 0) { $distroNotes.Add("skipped inaccessible distros: $($skippedBrokenDistros -join ', ')") }
                    if ($hadDistroWarnings) { $distroNotes.Add('warnings detected') }
                    $script:stepMessage += if ($distroNotes.Count -gt 0) { " (Checked inside Distros; $($distroNotes -join '; '))" } else { ' (Checked inside Distros)' }
                }
            }
        }
    }
)

Invoke-Update -Name 'Winget' -RequiresCommand 'winget' -Action {
    if (-not $isAdmin) {
        Write-Status 'Running as NON-ADMIN. MSI installers will likely fail with 1632.' -Type Warning
        Write-Status 'Re-run with -AutoElevate or as Administrator.' -Type Info
    }
    $oldTemp = $env:TEMP; $oldTmp = $env:TMP
    if ($isAdmin -and (Test-Path 'C:\Windows\Temp')) { $env:TEMP = 'C:\Windows\Temp'; $env:TMP = 'C:\Windows\Temp' }
    try {
        # Start Windows Installer service if stopped
        if ($isAdmin) {
            try {
                $msiService = Get-Service msiserver -EA SilentlyContinue
                if ($msiService) {
                    if ($msiService.Status -eq 'Stopped') {
                        Start-Service msiserver -EA Stop
                    }
                    else {
                        Write-Detail 'Windows Installer service already running; skipping restart.' -Type Muted
                    }
                }
            }
            catch { Write-Detail "Could not initialize Windows Installer service: $_" -Type Warning }
        }

        if (-not (Wait-WindowsInstallerIdle -TimeoutSec $script:Config.MsiBusyWaitSec -Reason 'before winget scan')) {
            throw "Windows Installer remained busy for $($script:Config.MsiBusyWaitSec)s before winget scan"
        }

        $rebootPending = Test-HardRebootPending
        if ($rebootPending) {
            Write-Status 'Windows reports a pending reboot through servicing/update markers.' -Type Warning
        }

        # Refresh winget source catalog
        if (-not $UltraFast) {
            Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('source', 'update', '--disable-interactivity') | Out-Null
        }

        $scan = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @('upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity')
        if ($scan.Output) { Write-FilteredOutput -Text $scan.Output -Color Gray }
        $scanEntries = @(Get-WingetUpgradeEntries $scan.Output)
        $script:Config.WingetScannedIds = @($scanEntries | ForEach-Object { $_.Id })
        $script:_wingetScannedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $scanEntries) { $null = $script:_wingetScannedIds.Add($e.Id) }
        if ($UltraFast -and $scanEntries.Count -gt 0 -and ($rebootPending -or -not $isAdmin)) {
            Write-Detail 'UltraFast mode: skipping the winget install pass because reboot/elevation issues would likely waste time.' -Type Warning
            $script:stepMessage = "found $($scanEntries.Count) upgrade(s); install pass skipped for speed"
            return
        }

        $updatedPackages = [System.Collections.Generic.List[string]]::new()
        $failedPackages = [System.Collections.Generic.List[string]]::new()

        foreach ($entry in $scanEntries) {
            $skipList = @($script:Config.WingetSkipPackages)
            if ($skipList.Count -gt 0 -and ($skipList -contains $entry.Id -or $skipList -contains $entry.Name)) {
                Write-Detail "$($entry.Name): skipped (in WingetSkipPackages list)" -Type Muted
                continue
            }

            try { Invoke-WingetUpgradeHook -Phase 'Pre' -WingetOutput $entry.Id } catch {}

            if (-not (Wait-WindowsInstallerIdle -TimeoutSec $script:Config.MsiBusyWaitSec -Reason "before installing $($entry.Name)")) {
                $failedPackages.Add($entry.Name)
                Write-Detail "$($entry.Name): skipped because Windows Installer remained busy" -Type Warning
                continue
            }

            $wingetArgs = @('upgrade', '--id', $entry.Id, '--include-unknown', '--source', 'winget', '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
            $maxAttempts = [Math]::Max(1, [int]$script:Config.WingetInstallRetryCount)
            $attempt = 1
            $perResult = $null
            $logPath = $null

            while ($attempt -le $maxAttempts) {
                $perResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments $wingetArgs
                $logPath = Get-WingetInstallerLogPath -WingetOutput $perResult.Output

                $installerBusy = ($perResult.Output -match 'Another installation is already in progress') -or
                ($perResult.Output -match 'Installer failed with exit code:\s*1618')

                if ($installerBusy -and $attempt -lt $maxAttempts) {
                    $waitSec = [Math]::Min(180, 20 * $attempt)
                    Write-Detail "$($entry.Name): another installer is active (1618). Retrying in ${waitSec}s (attempt $($attempt + 1)/$maxAttempts)" -Type Warning
                    [void](Wait-WindowsInstallerIdle -TimeoutSec $waitSec -Reason "retrying $($entry.Name)")
                    $attempt++
                    continue
                }
                break
            }
            $repaired = $false

            if (Test-WingetMissingMsiSourceFailure -WingetOutput $perResult.Output -LogPath $logPath) {
                $repairVersion = if ($entry.Version) { [string]$entry.Version } else { '' }
                if ($repairVersion -and (Test-MsiRepairFailureCached -Entry $entry -Version $repairVersion)) {
                    Write-Detail "$($entry.Name): skipping MSI source recache attempt (recently failed for version $repairVersion); cleaning registry for fresh install" -Type Warning
                    $unArgs2 = @('uninstall', '--id', $entry.Id, '--silent', '--accept-source-agreements', '--disable-interactivity')
                    $unResult2 = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments $unArgs2
                    if ($unResult2.Output) { Write-FilteredOutput -Text $unResult2.Output -Color Gray }
                    # Clean all MSI registry entries — works even when winget uninstall fails (MSI DB gone)
                    $regEntry2 = Get-InstalledUninstallEntryForWinget -Entry $entry
                    $cleanGuid2 = if ($regEntry2) { $regEntry2.PSChildName } else { $null }
                    $cleaned2 = Invoke-MsiFullRegistryCleanse -Guid $cleanGuid2 -ProductName $entry.Name
                    if ($cleaned2 -gt 0) { Write-Detail "$($entry.Name): cleared $cleaned2 broken MSI registry entries" -Type Warning }
                    $perResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @(
                        'install', '--id', $entry.Id, '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
                    )
                    $logPath = Get-WingetInstallerLogPath -WingetOutput $perResult.Output
                }
                else {
                    Write-Detail "$($entry.Name): repairing missing MSI source cache for installed version $($entry.Version) (this can take a minute or two)" -Type Warning
                    $repaired = Invoke-WingetMsiSourceRepair -Entry $entry
                    if ($repaired) {
                        if ($repairVersion) { Clear-MsiRepairFailureCache -Entry $entry -Version $repairVersion }
                        Write-Detail "$($entry.Name): MSI source repaired; retrying upgrade" -Type Warning
                        $perResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments $wingetArgs
                        $logPath = Get-WingetInstallerLogPath -WingetOutput $perResult.Output
                    }
                    else {
                        if ($repairVersion) { Set-MsiRepairFailureCache -Entry $entry -Version $repairVersion }
                        Write-Detail "$($entry.Name): MSI source repair failed; cleaning registry for fresh install" -Type Warning
                        $unArgs = @('uninstall', '--id', $entry.Id, '--silent', '--accept-source-agreements', '--disable-interactivity')
                        $unResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments $unArgs
                        if ($unResult.Output) { Write-FilteredOutput -Text $unResult.Output -Color Gray }
                        # Clean all MSI registry entries regardless of winget uninstall result
                        $regEntry = Get-InstalledUninstallEntryForWinget -Entry $entry
                        $cleanGuid = if ($regEntry) { $regEntry.PSChildName } else { $null }
                        $cleaned = Invoke-MsiFullRegistryCleanse -Guid $cleanGuid -ProductName $entry.Name
                        if ($cleaned -gt 0) { Write-Detail "$($entry.Name): cleared $cleaned broken MSI registry entries" -Type Warning }
                        Write-Detail "$($entry.Name): running fresh install" -Type Warning
                        $perResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @(
                            'install', '--id', $entry.Id, '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
                        )
                        $logPath = Get-WingetInstallerLogPath -WingetOutput $perResult.Output
                    }
                }
            }

            if ($perResult.Output) { Write-FilteredOutput -Text $perResult.Output -Color Gray }
            try { Invoke-WingetUpgradeHook -Phase 'Post' -WingetOutput $entry.Id } catch {}

            if ($perResult.Output -match 'Successfully installed') {
                $updatedPackages.Add($entry.Name)
                continue
            }

            if ($perResult.Output -match 'No applicable upgrade found|No newer package versions are available|already installed') {
                if ($entry.Available -and $perResult.Output -match 'No applicable upgrade found') {
                    Write-Detail "$($entry.Name): upgrade not applicable; trying direct install of $($entry.Available)" -Type Warning
                    $fallbackArgs = @('install', '--id', $entry.Id, '--version', $entry.Available, '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
                    $fallbackResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments $fallbackArgs
                    if ($fallbackResult.Output) { Write-FilteredOutput -Text $fallbackResult.Output -Color Gray }
                    if ($fallbackResult.Output -match 'No applicable installer found') {
                        Write-Detail "$($entry.Name): no applicable installer; retrying with --force" -Type Warning
                        $fallbackResult = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments ($fallbackArgs + @('--force'))
                        if ($fallbackResult.Output) { Write-FilteredOutput -Text $fallbackResult.Output -Color Gray }
                    }
                    if ($fallbackResult.Output -match 'Successfully installed') { $updatedPackages.Add($entry.Name) }
                }
                continue
            }

            if (Test-WingetMissingMsiSourceFailure -WingetOutput $perResult.Output -LogPath $logPath) {
                $failedPackages.Add($entry.Name)
                Write-Detail "$($entry.Name): upgrade still blocked because Windows cannot remove the old MSI package source (1714 / 1612)" -Type Warning
                continue
            }

            if ($perResult.Output -match 'Another installation is already in progress|Installer failed with exit code:\s*1618') {
                $failedPackages.Add($entry.Name)
                Write-Detail "$($entry.Name): installer busy (1618) after $maxAttempts attempt(s)" -Type Warning
                continue
            }

            if ($perResult.ExitCode -ne 0) {
                $failedPackages.Add($entry.Name)
                Write-Detail "$($entry.Name): upgrade failed" -Type Warning
            }
        }

        # Re-scan only when failures occurred
        if ($failedPackages.Count -gt 0) {
            $finalScan = Invoke-WingetWithTimeout -TimeoutSec $script:Config.WingetTimeoutSec -Arguments @('upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity')
            $finalEntries = @(Get-WingetUpgradeEntries $finalScan.Output)
            if ($finalScan.Output -and -not (Test-WingetUpgradeListsMatch -First $scanEntries -Second $finalEntries)) {
                Write-Detail 'Remaining upgrades after winget run:' -Type Muted
                Write-FilteredOutput -Text $finalScan.Output -Color Gray
            }
        }

        $script:stepChanged = $updatedPackages.Count -gt 0
        if ($failedPackages.Count -gt 0 -and $updatedPackages.Count -gt 0) {
            $script:stepMessage = "updated $($updatedPackages.Count) package(s); failed: $($failedPackages -join ', ')"
        }
        elseif ($failedPackages.Count -gt 0) {
            $script:stepMessage = "completed with some failures: $($failedPackages -join ', ')"
        }
        elseif ($updatedPackages.Count -gt 0) {
            $script:stepMessage = 'updated correctly'
        }
        else {
            $script:stepMessage = 'already current'
        }
    }
    finally { $env:TEMP = $oldTemp; $env:TMP = $oldTmp }
}

# ==============================================================================
# SYSTEM COMPONENTS - run in parallel
# ==============================================================================

Write-Host "`n[Parallel] Running system component updates..." -ForegroundColor DarkCyan

Invoke-Update -Name 'WindowsUpdate' -Title 'Windows Update' -RequiresAdmin -Disabled:$SkipWindowsUpdate -Action {
    Import-Module PSWindowsUpdate -ErrorAction Stop
    $params = @{
        Install         = $true
        AcceptAll       = $true
        IgnoreUserInput = $true
        NotCategory     = 'Drivers'
        IgnoreReboot    = $true
        RecurseCycle    = 3
        Verbose         = $false
        Confirm         = $false
        ErrorAction     = 'Stop'
        WarningAction   = 'SilentlyContinue'
        ProgressAction  = 'SilentlyContinue'
    }
    $run = Invoke-StreamingCapture ({ Get-WindowsUpdate @params | Out-String -Width 220 }.GetNewClosure())
    $out = Read-CapturedOutput $run.OutputPath
    if (-not $out -or $out -match 'No updates found|There are no applicable updates|No updates available') {
        $script:stepMessage = 'already current'
    }
    else {
        $script:stepChanged = $true
        $script:stepMessage = 'Updated successfully'
        if (-not $SkipReboot -and (Get-WURebootStatus -Silent -EA SilentlyContinue)) { $script:stepMessage += ' (Reboot pending)' }
    }
}

Invoke-UpdateBatch -ThrottleLimit 3 -Tasks @(
    @{
        Name            = 'StoreApps'
        Title           = 'Microsoft Store Apps'
        Disabled        = $SkipStoreApps
        RequiresCommand = 'winget'
        Action          = {
            Write-Detail 'Checking Microsoft Store app upgrades via winget'
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --source msstore --all --silent --accept-package-agreements --accept-source-agreements --disable-interactivity }).OutputPath
            if (-not $out -or $out -match 'No applicable upgrade|There are no available upgrades') {
                $script:stepMessage = 'no Microsoft Store app updates'
            }
            elseif ($out -match 'No installed package found') {
                $script:stepMessage = 'no Microsoft Store-managed apps found'
            }
            elseif ($out -match 'Successfully installed|Successfully upgraded') {
                $script:stepChanged = $true; $script:stepMessage = 'store apps updated'
            }
            else {
                $script:stepMessage = 'store apps checked'
            }
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
            $after = (Get-MpComputerStatus -EA SilentlyContinue).AntivirusSignatureVersion
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
            $current = ([string](npm --version 2>&1 | Select-Object -First 1)).Trim(); $latest = ([string](npm view npm version 2>&1 | Select-Object -First 1)).Trim()
            if ($current -ne $latest) {
                $run = Invoke-StreamingCapture { npm install -g npm@latest }
                if ($run.ExitCode -eq 0 -and ([string](npm --version 2>&1 | Select-Object -First 1)).Trim() -eq $latest) { $script:stepChanged = $true }
                else {
                    $out = Read-CapturedOutput $run.OutputPath
                    Write-Detail "npm self-update failed: $($out -split '\r\n|\r|\n' | Where-Object { $_.Trim() } | Select-Object -First 1)" -Type Warning
                }
            }
            $outdatedJson = Read-CapturedOutput (Invoke-StreamingCapture { npm outdated -g --json }).OutputPath
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -EA SilentlyContinue
                if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
                    $pkgs = @($outdated.PSObject.Properties.Name | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    $updated = [System.Collections.Generic.List[string]]::new()
                    $failed = [System.Collections.Generic.List[string]]::new()
                    foreach ($pkg in $pkgs) {
                        $run = Invoke-StreamingCapture ({ npm install -g $pkg }.GetNewClosure())
                        if ($run.ExitCode -eq 0) {
                            $updated.Add($pkg)
                        }
                        else {
                            $failed.Add($pkg)
                            Write-Detail "npm update failed for $pkg" -Type Warning
                        }
                    }
                    if ($updated.Count -gt 0) {
                        $script:stepChanged = $true
                        $script:stepMessage = "updated $($updated.Count) package(s)"
                        if ($failed.Count -gt 0) {
                            $script:stepMessage += "; failed: $($failed -join ', ')"
                        }
                    }
                    elseif ($failed.Count -gt 0) {
                        $script:stepMessage = "failed: $($failed -join ', ')"
                    }
                    else {
                        $script:stepMessage = 'no global package updates'
                    }
                }
                else { $script:stepMessage = 'no global package updates' }
            }
            else { $script:stepMessage = 'no global package updates' }
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
            $out = Read-CapturedOutput (Invoke-StreamingCapture { bun upgrade --stable }).OutputPath
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
            $out = Read-CapturedOutput (Invoke-StreamingCapture { deno upgrade --stable }).OutputPath
            if ($out -notmatch 'already the latest') { $script:stepChanged = $true; $script:stepMessage = 'updated' } else { $script:stepMessage = 'already current' }
        }
    }
    @{
        Name            = 'Rust'
        RequiresCommand = 'rustup'
        Disabled        = $SkipRust
        Action          = {
            $out = Read-CapturedOutput (Invoke-StreamingCapture { rustup update stable }).OutputPath
            if ($out -match 'error: failure removing component') {
                $errorLine = @($out -split '\r\n|\r|\n' | Where-Object { $_ -match 'error:' } | Select-Object -First 1)[0]
                if (-not $errorLine) { $errorLine = @($out -split '\r\n|\r|\n' | Where-Object { $_.Trim() } | Select-Object -First 1)[0] }
                if ($errorLine) { Write-Detail "Rustup error: $($errorLine.Trim())" -Type Warning }
                else { Write-Detail 'Rustup error encountered during update' -Type Warning }
                $script:stepMessage = 'toolchain checked; warning: rustup error removing component (see output)'
            }
            elseif ($out -match 'unchanged - rustc ([^\s]+)') {
                $script:stepMessage = "stable unchanged ($($Matches[1]))"
            }
            elseif ($out -match 'updated - rustc') {
                $script:stepChanged = $true; $script:stepMessage = 'stable updated'
            }
            else {
                $script:stepMessage = 'toolchain checked'
            }
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
            $wingetScannedIds = @($script:Config.WingetScannedIds)
            if ($wingetScannedIds -contains 'GoLang.Go') {
                $script:stepMessage = 'managed by winget (already updated)'; return
            }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --id GoLang.Go --silent --accept-source-agreements --accept-package-agreements --disable-interactivity }).OutputPath
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
        RequiresAnyCommand = @('tldr', 'tealdeer')
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
            $mgr = Get-ToolInstallManager 'git-lfs'
            if ($mgr) { $script:stepMessage = "managed by $mgr (already updated)"; return }
            $out = Read-CapturedOutput (Invoke-StreamingCapture { winget upgrade --id GitHub.GitLFS --silent --accept-source-agreements --accept-package-agreements --disable-interactivity }).OutputPath
            if ($out -match 'No installed package found') {
                $script:stepMessage = 'installed outside winget; skipping package-manager update'
            }
            elseif ($out -match 'Successfully installed|Successfully upgraded') {
                $script:stepChanged = $true; $script:stepMessage = 'updated via winget'
            }
            else {
                $script:stepMessage = 'already current'
            }
        }
    }
    @{
        Name            = 'git-credential-manager'
        RequiresCommand = 'git-credential-manager'
        Action          = {
            Invoke-WingetWithTimeout -Arguments @('upgrade', '--id', 'Git.Git', '--silent', '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity') | Out-Null
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
        Action          = { composer self-update --stable --no-interaction 2>&1 | Out-Null; composer global update --no-interaction 2>&1 | Out-Null; $script:stepMessage = 'checked' }
    }
    @{
        Name            = 'RubyGems'
        RequiresCommand = 'gem'
        Disabled        = $SkipRuby
        Action          = { gem update --system 2>&1 | Out-Null; gem update 2>&1 | Out-Null; $script:stepMessage = 'checked' }
    }
    @{
        Name     = 'vscode-extensions'
        Title    = 'VS Code Extensions'
        Disabled = $SkipVSCodeExtensions
        Action   = {
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
        Name   = 'pip'
        Title  = 'Python / pip'
        Action = {
            function Invoke-PipHealthCheck {
                $out = (python -m pip check 2>&1) | Out-String
                if ($out -and $out.Trim() -notmatch '^No broken requirements') {
                    $ignoreList = @($script:Config.PipIgnoreHealthPackages | ForEach-Object { (ConvertTo-NormalizedPackageName $_) })
                    ($out -split '\r\n|\r|\n' | Where-Object { $_.Trim() }) | ForEach-Object {
                        $line = $_
                        $suppressed = $ignoreList.Count -gt 0 -and ($ignoreList | Where-Object { $line -match [regex]::Escape($_) })
                        if (-not $suppressed) { Write-Detail $line -Type Warning }
                    }
                }
            }

            python -m pip install --upgrade pip 2>&1 | Out-Null
            $outdated = @(python -m pip list --outdated --format=json 2>$null | ConvertFrom-Json)
            if (-not $outdated -or $outdated.Count -eq 0) { Invoke-PipHealthCheck; $script:stepMessage = 'no global pip package updates'; return }

            $notRequired = @(python -m pip list --not-required --format=json 2>$null | ConvertFrom-Json)
            $topLevel = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($p in $notRequired) { $null = $topLevel.Add((ConvertTo-NormalizedPackageName $p.name)) }

            if ($topLevel.Count -eq 0) {
                Write-Detail 'Cannot determine top-level packages; skipping upgrades to avoid breaking deps.' -Type Warning
                Invoke-PipHealthCheck; $script:stepMessage = 'pip checked; skipped for safety'; return
            }

            $skip = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $allow = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($p in $script:Config.PipSkipPackages) { $null = $skip.Add((ConvertTo-NormalizedPackageName $p)) }
            foreach ($p in $script:Config.PipAllowPackages) { $null = $allow.Add((ConvertTo-NormalizedPackageName $p)) }

            $updated = [System.Collections.Generic.List[string]]::new()
            $failed = [System.Collections.Generic.List[string]]::new()
            $batch = @()
            $nSkip = 0; $nTransitive = 0

            foreach ($pkg in $outdated) {
                $n = ConvertTo-NormalizedPackageName $pkg.name
                if ($skip.Contains($n)) { continue }
                if (-not $topLevel.Contains($n)) { $nTransitive++; continue }
                if ($allow.Count -gt 0 -and -not $allow.Contains($n)) { $nSkip++; continue }
                $batch += $pkg.name
            }
            if ($nSkip -gt 0) { Write-Detail "$nSkip package$(if($nSkip -eq 1){''} else{'s'}) not in PipAllowPackages allowlist" }

            if ($batch.Count -eq 0) {
                Invoke-PipHealthCheck
                $parts = @()
                if ($failed.Count -gt 0) { $parts += "failed: $($failed -join ', ')" }
                $script:stepMessage = if ($parts.Count -gt 0) { "No pip packages updated; $($parts -join '; ')" } else { 'no global pip package updates' }
                return
            }

            $run = Invoke-StreamingCapture ({ python -m pip install --upgrade --quiet --disable-pip-version-check --no-input $batch }.GetNewClosure())
            if ($run.ExitCode -eq 0) {
                foreach ($p in $batch) { $updated.Add($p) }
            }
            else {
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
            if ($updated.Count -gt 0) { $parts += "Updated $($updated.Count) package$(if($updated.Count -eq 1){''}else{'s'})" }
            elseif ($failed.Count -eq 0) { $parts += if ($nTransitive -gt 0) { "Updated 0 top-level packages" } else { 'no global pip package updates' } }
            if ($failed.Count -gt 0) { $parts += "failed: $($failed -join ', ')" }
            if ($nTransitive -gt 0 -and $updated.Count -gt 0) { $parts += "$nTransitive transitive deps skipped" }
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
            }
            elseif ($out -match 'Failed to uninstall tool package.*Could not find a part of the path') {
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
                if ($repairedOk.Count -gt 0) { $repairParts += "repaired: $($repairedOk -join ', ')" }
                if ($repairFailed.Count -gt 0) { $repairParts += "repair failed: $($repairFailed -join ', ')" }
                $script:stepMessage = if ($repairParts.Count -gt 0) { $repairParts -join '; ' } else { 'tool update errors encountered' }
            }
            elseif ($out -match 'failed to update|Failed to uninstall|Tool .* failed') {
                $script:stepMessage = 'tool update errors encountered'
            }
            else {
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
            if ($out -match 'Successfully (updated|installed) workload\(s\):\s+[A-Za-z]') { $script:stepChanged = $true; $script:stepMessage = 'workloads updated' }
            elseif ($out -match 'Updated advertising manifest') { $script:stepMessage = 'workload manifests refreshed' }
            else { $script:stepMessage = 'workloads checked' }
        }
    }
    @{
        Name     = 'pwsh-resources'
        Title    = 'PowerShell Modules / Resources'
        Disabled = $SkipPowerShellModules
        Action   = {
            $details = [System.Collections.Generic.List[string]]::new()
            $psResourceNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

            if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
                $cmd = Get-Command Update-PSResource -EA SilentlyContinue
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
                    if ($prev -and $curr -ne $prev) { $details.Add("PSResource $($_.Name) $prev -> $curr") }
                }
            }

            if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
                $cmd = Get-Command Update-Module -EA SilentlyContinue
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
                        if ($prev -and $curr -ne $prev) { $details.Add("Module $($module.Name) $prev -> $curr") }
                    }
                }
                else {
                    Write-Detail 'Legacy PowerShellGet module pass skipped: PSResourceGet already covers installed modules.'
                }
            }

            if ($details.Count -gt 0) { $script:stepChanged = $true; $script:stepMessage = $details -join '; ' }
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
}
elseif (-not $script:IsSimulation) {
    Write-Section 'System Cleanup'
    $cutoff = (Get-Date).AddDays(-$script:Config.TempCleanupDays)
    $preserveTempPath = {
        param($item)
        if (-not $item) { return $false }
        return ($item.FullName -match '\\WinGet(\\|$)')
    }
    try {
        if ($env:TEMP -and (Test-Path $env:TEMP)) {
            Get-ChildItem $env:TEMP -EA SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff -and -not (& $preserveTempPath $_) } |
            Remove-Item -Recurse -Force -EA SilentlyContinue
            Write-Status "Temp files cleared (older than $($script:Config.TempCleanupDays) days, preserving WinGet caches)" -Type Success
        }
    }
    catch {}
    if ($isAdmin) {
        try {
            if (Test-Path 'C:\Windows\Temp') {
                Get-ChildItem 'C:\Windows\Temp' -EA SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff -and -not (& $preserveTempPath $_) } |
                Remove-Item -Recurse -Force -EA SilentlyContinue
                Write-Status "C:\Windows\Temp cleared (older than $($script:Config.TempCleanupDays) days, preserving WinGet caches)" -Type Success
            }
        }
        catch {}
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
            @{ Name = 'Winget'; Prev = $prev.Winget; Curr = $script:State.Winget }
            @{ Name = 'Scoop'; Prev = $prev.Scoop; Curr = $script:State.Scoop }
            @{ Name = 'Chocolatey'; Prev = $prev.Chocolatey; Curr = $script:State.Chocolatey }
        )) {
        $changes = @(Compare-PackageMaps $entry.Prev $entry.Curr)
        if ($changes.Count -gt 0) { Write-Host "  $($entry.Name) changes:" -ForegroundColor Cyan; $changes | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray } }
        else { Write-Host "  $($entry.Name): no changes detected" -ForegroundColor DarkGray }
    }
    Save-State
}
elseif (-not $script:IsSimulation) {
    Save-State
}

# ==============================================================================
# SUMMARY
# ==============================================================================

$duration = (Get-Date) - $startTime
$updatedNames = @($updateResults.Success | Select-Object -Unique)
$checkedNames = @($updateResults.Checked | Where-Object { $_ -notin $updatedNames -and $_ -notin $updateResults.Failed } | Select-Object -Unique)
$skippedItems = @($updateResults.Skipped)
$failedNames = @($updateResults.Failed | Select-Object -Unique)

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
if ($failedNames.Count -gt 0) {
    Write-Host "[X] Failed    ($($failedNames.Count))" -ForegroundColor Red
    foreach ($n in $failedNames) { Write-Detail "$n`: $($updateResults.Details[$n])" -Type Error }
}
if ($skippedItems.Count -gt 0) {
    Write-Host "[-] Skipped   ($($skippedItems.Count))" -ForegroundColor DarkGray
    foreach ($s in $skippedItems) {
        $reason = if ($s.Reason) { $s.Reason } else { 'unknown reason' }
        Write-Detail "$($s.Name): $reason" -Type Muted
    }
}

# --- One-line health verdict ---
$healthParts = @()
if ($updatedNames.Count -gt 0) { $healthParts += "$($updatedNames.Count) updated" }
if ($checkedNames.Count -gt 0) { $healthParts += "$($checkedNames.Count) checked" }
if ($skippedItems.Count -gt 0) { $healthParts += "$($skippedItems.Count) skipped" }
if ($failedNames.Count -gt 0) { $healthParts += "$($failedNames.Count) FAILED" }
$healthColor = if ($failedNames.Count -gt 0) { 'Yellow' } else { 'Green' }
Write-Host ("`nHealth: " + ($healthParts -join ', ')) -ForegroundColor $healthColor
Write-Log ("Health: " + ($healthParts -join ', '))

if ($script:SectionTimings.Count -gt 0) {
    Write-Host "`n  Section timings:" -ForegroundColor DarkGray
    $script:SectionTimings.GetEnumerator() | Sort-Object Value -Descending |
    ForEach-Object { Write-Host ("    {0,-30} {1,6}s" -f $_.Key, $_.Value.ToString('F1', [cultureinfo]::InvariantCulture)) -ForegroundColor DarkGray }
}

# --- Toast notification -------------------------------------------------------
try {
    if (Get-Module -ListAvailable BurntToast -EA SilentlyContinue) {
        Import-Module BurntToast -EA SilentlyContinue
        $msg = if ($failedNames.Count -gt 0) { "$($updatedNames.Count) updated, $($failedNames.Count) failed" }
        elseif ($updatedNames.Count -gt 0) { "$($updatedNames.Count) components updated" }
        else { 'No updates were needed' }
        New-BurntToastNotification -Text 'Update-Everything', $msg -EA SilentlyContinue
    }
}
catch {}

if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }
exit 0
