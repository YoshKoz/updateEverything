<#
.SYNOPSIS
    System-wide update script for Windows
.DESCRIPTION
    Updates all package managers, system components, and development tools.
    Supports parallel execution of independent package managers for faster runs.
.VERSION
    2.7.0
.NOTES
    Run as Administrator for full functionality (some features require elevation).
    Requires PowerShell 7+ for -Parallel support.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -Parallel
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
    .\updatescript.ps1 -SkipCleanup -SkipWindowsUpdate
    .\updatescript.ps1 -SkipNode -SkipRust -SkipGo
    .\updatescript.ps1 -DeepClean
    .\updatescript.ps1 -UpdateOllamaModels
    .\updatescript.ps1 -WhatChanged
    .\updatescript.ps1 -DryRun
    .\updatescript.ps1 -NoParallel
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive script uses colorized host output intentionally.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '', Justification = 'Job scriptblocks receive values via parameters or generated initialization script; analyzer flags false positives here.')]
[CmdletBinding(SupportsShouldProcess)]
param(
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
    [switch]$Schedule,                  # Register a daily scheduled task to run this script
    [string]$ScheduleTime = "03:00",    # Time for the scheduled task (default: 3 AM)
    [string]$LogPath,
    [switch]$SkipNode,                  # Skip all Node.js toolchain updates (npm, pnpm, bun, deno, fnm, volta)
    [switch]$SkipRust,                  # Skip Rust toolchain updates (rustup, cargo)
    [switch]$SkipGo,                    # Skip Go toolchain updates
    [switch]$SkipFlutter,               # Skip Flutter SDK update
    [switch]$SkipGitLFS,                # Skip Git LFS client update
    [switch]$DeepClean,                 # Run DISM WinSxS cleanup, DO cache, and prefetch (adds ~7 min)
    [switch]$UpdateOllamaModels,        # Opt-in: pull latest for every installed Ollama model
    [switch]$WhatChanged,               # Show packages that changed since last run
    [switch]$NoParallel,                # Disable parallel execution of independent tools (PS7+)
    [switch]$DryRun                     # Show which steps would run without executing them
)

# ── Console encoding / external tool output ─────────────────────────────────────
try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
}
catch {
    Write-Verbose "Console encoding update skipped: $($_.Exception.Message)"
}

# ── Self-elevate ────────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and $AutoElevate -and -not $NoElevate) {
    $params = $PSBoundParameters.GetEnumerator() | ForEach-Object {
        if ($_.Value -is [switch]) { if ($_.Value.IsPresent) { "-$($_.Key)" } }
        else { "-$($_.Key)"; [string]$_.Value }
    }
    $pwshArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + @($params)
    try {
        $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue

        if ($pwshCmd) {
            Start-Process -FilePath $pwshCmd.Source -Verb RunAs -ArgumentList $pwshArgs -Wait
        }
        else {
            Write-Host "WARNING: pwsh.exe not found. Falling back to Windows PowerShell." -ForegroundColor Yellow
            Start-Process powershell -Verb RunAs -ArgumentList $pwshArgs -Wait
        }
        exit
    }
    catch {
        Write-Host "WARNING: Could not elevate. Running without Administrator privileges.`n" -ForegroundColor Yellow
    }
}
elseif (-not $isAdmin -and -not $NoElevate) {
    Write-Host "INFO: Running in the current terminal without elevation (admin-only tasks may be skipped)." -ForegroundColor DarkYellow
    Write-Host "      To run elevated, start the terminal as Administrator or use -AutoElevate (opens a new window)." -ForegroundColor DarkYellow
}

# ── Scheduled task registration ─────────────────────────────────────────────────
if ($Schedule) {
    $taskName = "DailySystemUpdate"
    $action = New-ScheduledTaskAction -Execute 'pwsh.exe' `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -SkipReboot -SkipWindowsUpdate"
    $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
        -Settings $settings -RunLevel Highest -Force | Out-Null
    Write-Host "[OK] Scheduled task '$taskName' registered to run daily at $ScheduleTime." -ForegroundColor Green
    exit
}

if ($WingetTimeoutSec -lt 30) { $WingetTimeoutSec = 30 }
$ErrorActionPreference = "Continue"
$startTime = Get-Date
$commandCache = @{}
$script:sectionTimings = [ordered]@{}      # Name -> elapsed seconds

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

function Get-VSCodeCliPath {
    # Always prefer the CLI shim (.cmd) to avoid launching Code.exe UI windows.
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

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('Success', 'Warning', 'Error', 'Info')]
        [string]$Type = 'Info'
    )
    $colors = @{ Success = 'Green'; Warning = 'Yellow'; Error = 'Red'; Info = 'Gray' }
    $symbols = @{ Success = '[OK]'; Warning = '[!]'; Error = '[X]'; Info = '[*]' }
    Write-Host "$($symbols[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Write-Detail {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Muted', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )

    $colors = @{ Info = 'Gray'; Muted = 'DarkGray'; Warning = 'Yellow'; Error = 'Red' }
    $prefixes = @{ Info = '  >'; Muted = '  -'; Warning = '  !'; Error = '  x' }
    Write-Host "$($prefixes[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Get-ToolInstallManager([string]$Command) {
    $src = (Get-Command $Command -ErrorAction SilentlyContinue).Source
    if (-not $src) { return $null }
    if ($src -like '*\scoop\*' -or $src -like '*scoop\shims\*') { return 'scoop' }
    if ($src -like '*\WinGet\*' -or $src -like '*WindowsApps\*' -or $src -like '*\winget\*') { return 'winget' }
    return $null
}

function Invoke-WingetPackageUpgrade([string]$Id, [int]$TimeoutSec = 120) {
    return Invoke-WingetWithTimeout -TimeoutSec $TimeoutSec -Arguments @(
        'upgrade', '--id', $Id,
        '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
    )
}

function Write-IndentedOutput {
    param(
        [AllowNull()][string]$Text,
        [string]$Prefix = '  >',
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    $normalized = $Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', ''
    $normalized = $normalized -replace "`r", "`n"

    foreach ($rawLine in ($normalized -split "`n")) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }
        Write-Host "$Prefix $line" -ForegroundColor $Color
    }
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
        if ($compact -match 'Γû|[█▓▒░▏▎▍▌▋▊▉■□▪▫]') { continue }

        # Drop noisy winget "cannot be determined" info lines.
        if ($compact -match 'package\(s\) have version numbers that cannot be determined') { continue }

        # Drop table separators and blank winget headers that add noise without new information.
        if ($compact -match '^[\-=]{6,}$') { continue }

        Write-Host $line -ForegroundColor $Color
    }
}

function Get-WingetFailedPackageId {
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
            try { $proc.Kill() } catch { Write-Verbose "Failed to terminate timed-out winget process: $($_.Exception.Message)" }
            throw "winget timed out after ${TimeoutSec}s — a hanging installer may require manual intervention"
        }

        $exitCode = $proc.ExitCode
        $stdout = Get-Content -Raw -Path $stdoutFile -ErrorAction SilentlyContinue
        $stderr = Get-Content -Raw -Path $stderrFile -ErrorAction SilentlyContinue
        $combined = (($stdout + $stderr) -replace '\x00', '').Trim()

        return [pscustomobject]@{ Output = $combined; ExitCode = $exitCode }
    }
    finally {
        Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
    }
}

$updateResults = @{
    Success = [System.Collections.Generic.List[string]]::new()
    Checked = [System.Collections.Generic.List[string]]::new()
    Failed  = [System.Collections.Generic.List[string]]::new()
    Skipped = [System.Collections.Generic.List[pscustomobject]]::new()   # @{ Name; Reason }
}

$script:updateDetails = @{}   # Name -> @{ Summary; Updated; Info; Warning; Changed }

function Get-UpdateDetailBucket {
    param([Parameter(Mandatory)][string]$Name)

    if (-not $script:updateDetails.ContainsKey($Name)) {
        $script:updateDetails[$Name] = @{
            Summary = $null
            Updated = [System.Collections.Generic.List[string]]::new()
            Info    = [System.Collections.Generic.List[string]]::new()
            Warning = [System.Collections.Generic.List[string]]::new()
            Changed = $false
        }
    }

    return $script:updateDetails[$Name]
}

function Set-UpdateSummary {
    param(
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][string]$Summary,
        [switch]$Changed
    )

    if ([string]::IsNullOrWhiteSpace($Summary)) { return }
    $bucket = Get-UpdateDetailBucket -Name $Name
    $bucket['Summary'] = $Summary.Trim()
    if ($Changed) { $bucket['Changed'] = $true }
}

function Add-UpdateItem {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Item,
        [ValidateSet('Updated', 'Info', 'Warning')]
        [string]$Kind = 'Updated'
    )

    if ([string]::IsNullOrWhiteSpace($Item)) { return }

    $bucket = Get-UpdateDetailBucket -Name $Name
    $list = $bucket[$Kind]
    $value = $Item.Trim()
    if (-not $list.Contains($value)) {
        $list.Add($value)
    }
    if ($Kind -eq 'Updated') {
        $bucket['Changed'] = $true
    }
}

# ── Thin wrappers used by parallel action blocks in sequential mode ───────────────
function Write-UpdateSummaryMarker {
    param([string]$Section, [string]$Summary, [switch]$Changed)
    Set-UpdateSummary -Name $Section -Summary $Summary -Changed:$Changed
}
function Write-UpdateItemMarker {
    param([string]$Section, [string]$Item,
        [ValidateSet('Updated', 'Info', 'Warning')][string]$Kind = 'Updated')
    Add-UpdateItem -Name $Section -Kind $Kind -Item $Item
}

function Test-SectionChanged {
    param([Parameter(Mandatory)][string]$Name)

    if (-not $script:updateDetails.ContainsKey($Name)) { return $false }
    return [bool]$script:updateDetails[$Name]['Changed']
}

function ConvertFrom-Base64Utf8 {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
}

function Format-UpdateItemList {
    param(
        [AllowEmptyCollection()][string[]]$Items,
        [int]$MaxItems = 6
    )

    $list = @($Items | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($list.Count -eq 0) { return $null }
    if ($list.Count -le $MaxItems) { return ($list -join '; ') }
    return ('{0}; +{1} more' -f ($list[0..($MaxItems - 1)] -join '; '), ($list.Count - $MaxItems))
}

function Get-UpdateSummaryText {
    param([Parameter(Mandatory)][string]$Name)

    if (-not $script:updateDetails.ContainsKey($Name)) { return $null }

    $bucket = $script:updateDetails[$Name]
    if ($bucket['Summary']) { return $bucket['Summary'] }

    $parts = [System.Collections.Generic.List[string]]::new()

    $updated = Format-UpdateItemList -Items $bucket['Updated']
    if ($updated) { $parts.Add("updated: $updated") }

    $warning = Format-UpdateItemList -Items $bucket['Warning'] -MaxItems 3
    if ($warning) { $parts.Add("warnings: $warning") }

    if ($parts.Count -eq 0) { return $null }
    return ($parts -join ' | ')
}

function Get-WingetUpgradeCandidate {
    param(
        [AllowNull()][string]$Output,
        [switch]$IncludeUnknown
    )

    if ([string]::IsNullOrWhiteSpace($Output)) { return @() }

    $packages = [System.Collections.Generic.List[object]]::new()
    $seenIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($rawLine in ($Output -split '\r?\n')) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }
        if ($line -match '^(Name|[-=\s]{4,}|\d+\s+upgrades available\.|No installed package found matching input criteria\.)') { continue }
        if ($line -match '^(Found\s|This application is licensed|Microsoft is not responsible|Downloading |Successfully |Starting package |Extracting archive|Installer log )') { continue }

        if ($line -notmatch '^(?<name>.+?)\s+(?<id>(?=\S*[A-Za-z])\S+\.\S+)\s+(?<version>Unknown|\S+)\s+(?<available>\S+)(?:\s+(?<source>\S+))?\s*$') {
            continue
        }

        $id = $Matches['id'].Trim()
        if (-not $seenIds.Add($id)) { continue }

        $version = $Matches['version'].Trim()
        $available = $Matches['available'].Trim()
        if (-not $IncludeUnknown -and $version -eq 'Unknown') { continue }

        $packages.Add([pscustomobject]@{
                Name      = $Matches['name'].Trim()
                Id        = $id
                Version   = $version
                Available = $available
            })
    }

    return @($packages)
}

function Get-WingetFoundPackage {
    param([AllowNull()][string]$Output)

    if ([string]::IsNullOrWhiteSpace($Output)) { return @() }

    $packages = [System.Collections.Generic.List[object]]::new()
    foreach ($rawLine in ($Output -split '\r?\n')) {
        $line = $rawLine.Trim()
        if ($line -match '^\(\d+/\d+\)\s+Found\s+(.+?)\s+\[((?=\S*[A-Za-z])\S+\.\S+)\]\s+Version\s+(\S+)\s*$') {
            $packages.Add([pscustomobject]@{
                    Name      = $Matches[1].Trim()
                    Id        = $Matches[2].Trim()
                    Version   = $null
                    Available = $Matches[3].Trim()
                })
        }
    }

    return @($packages)
}

function Get-WingetUnknownUpgradeCandidate {
    param([AllowNull()][string]$Output)

    if ([string]::IsNullOrWhiteSpace($Output)) { return @() }

    $packages = [System.Collections.Generic.List[object]]::new()
    $seenIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($rawLine in ($Output -split '\r?\n')) {
        $line = $rawLine.Trim()
        if (-not $line) { continue }
        if ($line -match '^(Name|[-=\s]{4,}|\d+\s+upgrades available\.|No installed package found matching input criteria\.)') { continue }

        if ($line -match '^(?<name>.+?)\s+(?<displayVersion>\d+(?:\.\d+)+(?:[-+][^\s]+)?)\s+(?<id>(?=\S*[A-Za-z])\S+\.\S+)\s+Unknown\s+(?<available>\S+)(?:\s+(?<source>\S+))?\s*$') {
            $id = $Matches['id'].Trim()
            if ($seenIds.Add($id)) {
                $packages.Add([pscustomobject]@{
                        Name                 = $Matches['name'].Trim()
                        Id                   = $id
                        DisplayVersion       = $Matches['displayVersion'].Trim()
                        Available            = $Matches['available'].Trim()
                        InstalledLooksCurrent = ($Matches['displayVersion'].Trim() -eq $Matches['available'].Trim())
                    })
            }
            continue
        }

        if ($line -match '^(?<name>.+?)\s+(?<id>(?=\S*[A-Za-z])\S+\.\S+)\s+Unknown\s+(?<available>\S+)(?:\s+(?<source>\S+))?\s*$') {
            $id = $Matches['id'].Trim()
            if ($seenIds.Add($id)) {
                $packages.Add([pscustomobject]@{
                        Name                 = $Matches['name'].Trim()
                        Id                   = $id
                        DisplayVersion       = $null
                        Available            = $Matches['available'].Trim()
                        InstalledLooksCurrent = $false
                    })
            }
        }
    }

    return @($packages)
}

function Clear-WingetInstallerCache {
    param([Parameter(Mandatory)][string]$PackageId)

    $cacheRoot = Join-Path $env:LOCALAPPDATA 'Temp\WinGet'
    if (-not (Test-Path $cacheRoot)) { return }

    $targets = Get-ChildItem -Path $cacheRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$PackageId*" }

    foreach ($target in @($targets)) {
        try {
            Remove-Item -LiteralPath $target.FullName -Recurse -Force -ErrorAction Stop
            Write-Detail "Cleared stale winget cache: $($target.Name)" -Type Muted
        }
        catch {
            Write-Status "Could not clear winget cache for $PackageId at $($target.FullName): $($_.Exception.Message)" -Type Warning
        }
    }
}

function Get-WingetCachedInstallerPath {
    param([Parameter(Mandatory)][string]$PackageId)

    $cacheRoot = Join-Path $env:LOCALAPPDATA 'Temp\WinGet'
    if (-not (Test-Path $cacheRoot)) { return $null }

    $packageDirs = Get-ChildItem -Path $cacheRoot -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "$PackageId*" } |
    Sort-Object LastWriteTime -Descending

    foreach ($dir in @($packageDirs)) {
        $installer = Get-ChildItem -Path $dir.FullName -File -Filter '*.msi' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
        if ($installer) { return $installer.FullName }
    }

    return $null
}

function Get-PendingRestartReasons {
    $reasons = [System.Collections.Generic.List[string]]::new()

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        $null = $reasons.Add('Windows Update')
    }

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        $null = $reasons.Add('Component Based Servicing')
    }

    $renameOps = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
    if ($renameOps) {
        $null = $reasons.Add('pending file rename operations')
    }

    return @($reasons)
}

$script:WingetMsiFallbacks = @{
    'Google.Chrome' = @{
        AdditionalArguments = @()
    }
    'Kitware.CMake' = @{
        AdditionalArguments = @('ADD_CMAKE_TO_PATH=System')
    }
}

function Invoke-WingetPackageHook {
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [Parameter(Mandatory)][ValidateSet('Pre', 'Post')][string]$Phase
    )

    if (-not $script:WingetUpgradeHooks.ContainsKey($PackageId)) { return }

    $packageHooks = $script:WingetUpgradeHooks[$PackageId]
    if (-not $packageHooks -or -not $packageHooks.ContainsKey($Phase)) { return }

    $hook = $packageHooks[$Phase]
    if (-not $hook) { return }

    try {
        & $hook
    }
    catch {
        Write-Status "Hook ($Phase) for $PackageId failed: $_" -Type Warning
    }
}

function Invoke-WingetMsiFallback {
    param(
        [Parameter(Mandatory)][string]$PackageId,
        [int]$TimeoutSec = 300
    )

    $fallback = $script:WingetMsiFallbacks[$PackageId]
    if (-not $fallback) { return $null }

    $installerPath = Get-WingetCachedInstallerPath -PackageId $PackageId
    if (-not $installerPath) {
        Write-Status "No cached MSI found for $PackageId after winget retry" -Type Warning
        return [pscustomobject]@{ Success = $false; ExitCode = $null; LogPath = $null; InstallerPath = $null }
    }

    $diagDir = Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir'
    New-Item -ItemType Directory -Path $diagDir -Force -ErrorAction SilentlyContinue | Out-Null
    $logPath = Join-Path $diagDir ('{0}.msiexec-{1}.log' -f $PackageId, (Get-Date -Format 'yy-MM-dd-HH-mm-ss'))

    $arguments = @('/i', $installerPath, '/passive', '/norestart', '/L*V!', $logPath)
    if ($fallback.AdditionalArguments) { $arguments += $fallback.AdditionalArguments }

    Write-Detail "Retrying $PackageId via msiexec fallback..." -Type Muted
    $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -PassThru
    $exited = $proc.WaitForExit($TimeoutSec * 1000)
    if (-not $exited) {
        try { $proc.Kill() } catch { }
        Write-Status "msiexec fallback timed out for $PackageId after ${TimeoutSec}s" -Type Warning
        return [pscustomobject]@{ Success = $false; ExitCode = $null; LogPath = $logPath; InstallerPath = $installerPath }
    }

    $success = $proc.ExitCode -in @(0, 1641, 3010)
    if ($success) {
        Write-Detail "$PackageId msiexec fallback succeeded with exit code $($proc.ExitCode)" -Type Muted
    }
    else {
        Write-Status "msiexec fallback failed for $PackageId with exit code $($proc.ExitCode); log: $logPath" -Type Warning
        if ($proc.ExitCode -eq 1632) {
            $pendingRestartReasons = @(Get-PendingRestartReasons)
            if ($pendingRestartReasons.Count -gt 0) {
                Write-Detail ("Detected reboot-pending state: {0}. Reboot before retrying $PackageId." -f ($pendingRestartReasons -join ', ')) -Type Warning
            }
        }
    }

    return [pscustomobject]@{
        Success       = $success
        ExitCode      = $proc.ExitCode
        LogPath       = $logPath
        InstallerPath = $installerPath
    }
}

function Invoke-ZedUserInstallerFallback {
    param([int]$TimeoutSec = 300)

    $showResult = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
        'show', '--id', 'ZedIndustries.Zed',
        '--accept-source-agreements', '--disable-interactivity'
    )
    if ($showResult.ExitCode -ne 0 -or [string]::IsNullOrWhiteSpace($showResult.Output)) {
        Write-Status 'Could not retrieve Zed package metadata for direct installer fallback' -Type Warning
        return $false
    }

    $installerUrl = $null
    $version = $null
    foreach ($rawLine in ($showResult.Output -split '\r?\n')) {
        $line = $rawLine.Trim()
        if (-not $installerUrl -and $line -match '^Installer Url:\s+(?<url>\S+)$') {
            $installerUrl = $Matches['url'].Trim()
        }
        if (-not $version -and $line -match '^Version:\s+(?<version>\S+)$') {
            $version = $Matches['version'].Trim()
        }
    }

    if (-not $installerUrl) {
        Write-Status 'Zed direct installer fallback could not find an installer URL in winget metadata' -Type Warning
        return $false
    }

    $zedWasRunning = [bool](Get-Process -Name 'Zed' -ErrorAction SilentlyContinue)
    if ($zedWasRunning) {
        Write-Detail 'Closing Zed for direct installer fallback...' -Type Muted
        Stop-Process -Name 'Zed' -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }

    $downloadPath = Join-Path $env:TEMP ("Zed-{0}-x86_64.exe" -f $(if ($version) { $version } else { 'latest' }))
    $logPath = Join-Path $env:TEMP ("Zed-install-{0}.log" -f (Get-Date -Format 'yy-MM-dd-HH-mm-ss'))

    try {
        Invoke-WebRequest -Uri $installerUrl -OutFile $downloadPath
        $proc = Start-Process -FilePath $downloadPath -ArgumentList @(
            '/VERYSILENT',
            '/SUPPRESSMSGBOXES',
            '/NORESTART',
            '/CURRENTUSER',
            ('/LOG="{0}"' -f $logPath)
        ) -PassThru
        $exited = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $exited) {
            try { $proc.Kill() } catch { }
            Write-Status "Zed direct installer fallback timed out after ${TimeoutSec}s" -Type Warning
            return $false
        }

        if ($proc.ExitCode -eq 0) {
            Write-Detail ('Zed direct installer fallback succeeded{0}' -f $(if ($version) { " ($version)" } else { '' })) -Type Muted
            return $true
        }

        Write-Status "Zed direct installer fallback failed with exit code $($proc.ExitCode); log: $logPath" -Type Warning
        return $false
    }
    catch {
        Write-Status "Zed direct installer fallback failed: $($_.Exception.Message)" -Type Warning
        return $false
    }
}

# ── Winget upgrade hooks (pre/post per package ID) ──────────────────────────────
# Each key is a winget package ID.  Value is a hashtable with optional Pre / Post scriptblocks.
$script:WingetUpgradeHooks = @{
    'Spotify.Spotify'                     = @{
        Pre  = {
            $script:_spotifyWasRunning = [bool](Get-Process -Name Spotify -ErrorAction SilentlyContinue)
            Stop-Process -Name Spotify -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
        Post = {
            if ($script:_spotifyWasRunning) {
                Start-Process "$env:APPDATA\Spotify\Spotify.exe" -ErrorAction SilentlyContinue
            }
        }
    }
    'Google.Chrome'                       = @{
        Pre  = {
            $script:_chromeWasRunning = [bool](Get-Process -Name 'chrome' -ErrorAction SilentlyContinue)
            if ($script:_chromeWasRunning) {
                Write-Host "  Closing Google Chrome for upgrade..." -ForegroundColor Gray
                Stop-Process -Name 'chrome' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
            }
        }
        Post = {
            if ($script:_chromeWasRunning) {
                $chromePath = Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'
                if (-not (Test-Path $chromePath)) {
                    $chromePath = Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'
                }
                if (Test-Path $chromePath) { Start-Process $chromePath -ErrorAction SilentlyContinue }
            }
        }
    }
    'Microsoft.VisualStudioCode'          = @{
        Pre  = {
            $script:_vscodeWasRunning = [bool](Get-Process -Name 'Code' -ErrorAction SilentlyContinue)
            if ($script:_vscodeWasRunning) {
                Write-Host "  Closing VS Code for upgrade..." -ForegroundColor Gray
                Stop-Process -Name 'Code' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
            }
        }
        Post = {
            if ($script:_vscodeWasRunning) {
                $codePath = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\Code.exe'
                if (Test-Path $codePath) { Start-Process $codePath -ErrorAction SilentlyContinue }
            }
        }
    }
    'Microsoft.VisualStudioCode.Insiders' = @{
        Pre  = {
            $script:_vscodeInsidersWasRunning = [bool](Get-Process -Name 'Code - Insiders' -ErrorAction SilentlyContinue)
            if ($script:_vscodeInsidersWasRunning) {
                Write-Host "  Closing VS Code Insiders for upgrade..." -ForegroundColor Gray
                Stop-Process -Name 'Code - Insiders' -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
            }
        }
        Post = {
            if ($script:_vscodeInsidersWasRunning) {
                $codePath = Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code Insiders\Code - Insiders.exe'
                if (Test-Path $codePath) { Start-Process $codePath -ErrorAction SilentlyContinue }
            }
        }
    }
}

function Invoke-WingetUpgradeHook {
    <#
    .SYNOPSIS  Run Pre or Post hooks for any package IDs detected in winget output.
    #>
    param(
        [string]$Phase,          # 'Pre' or 'Post'
        [AllowNull()][string]$WingetOutput
    )
    if (-not $WingetOutput) { return }
    foreach ($pkgId in $script:WingetUpgradeHooks.Keys) {
        if ($WingetOutput -match [regex]::Escape($pkgId)) {
            $hook = $script:WingetUpgradeHooks[$pkgId][$Phase]
            if ($hook) {
                try { & $hook } catch {
                    Write-Status "Hook ($Phase) for $pkgId failed: $_" -Type Warning
                }
            }
        }
    }
}

# ── Logging ──────────────────────────────────────────────────────────────────────
$transcriptStarted = $false
if ($LogPath) {
    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if ($logDir) { New-Item -ItemType Directory -Path $logDir -Force -ErrorAction SilentlyContinue | Out-Null }
        Start-Transcript -Path $LogPath -Append -Force | Out-Null
        $transcriptStarted = $true
    }
    catch {
        Write-Host "WARNING: Could not start transcript at $LogPath. $_" -ForegroundColor Yellow
    }
}

# ── Startup banner ───────────────────────────────────────────────────────────────
Write-Host "`n$('=' * 54)" -ForegroundColor DarkGray
Write-Host "  Update-Everything v2.7.0  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Cyan
Write-Host "  $(if ($isAdmin) { 'Running as Administrator' } else { 'Running as standard user (some tasks skipped)' })" -ForegroundColor $(if ($isAdmin) { 'Green' } else { 'Yellow' })
Write-Host "$('=' * 54)" -ForegroundColor DarkGray

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

    # Helper to record a skip with a reason
    $addSkip = { param([string]$r) $updateResults.Skipped.Add([pscustomobject]@{ Name = $Name; Reason = $r }) }

    if ($Disabled) { & $addSkip 'flag'; return }
    if ($RequiresCommand -and -not (Test-Command $RequiresCommand)) { & $addSkip 'not installed'; return }
    if ($RequiresAnyCommand -and -not ($RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1)) { & $addSkip 'not installed'; return }
    if ($RequiresAdmin -and -not $isAdmin) { & $addSkip 'requires admin'; return }
    if ($SlowOperation -and $FastMode) { & $addSkip 'fast mode'; return }

    if ($DryRun) {
        Write-Host "  [DryRun] Would run: $Title" -ForegroundColor DarkCyan
        return
    }

    $sectionStart = Get-Date
    if (-not $NoSection) { Write-Section $Title }

    try {
        & $Action
        $elapsed = ((Get-Date) - $sectionStart).TotalSeconds
        if (Test-SectionChanged -Name $Name) {
            Write-Status ("$Name updated ({0:F1}s)" -f $elapsed) -Type Success
            $updateResults.Success.Add($Name)
        }
        else {
            $updateResults.Checked.Add($Name)
        }
        $script:sectionTimings[$Name] = $elapsed
    }
    catch {
        Write-Status "$Name failed: $_" -Type Error
        $updateResults.Failed.Add($Name)
    }
}

# ── Parallel update runner ───────────────────────────────────────────────────────
function Invoke-UpdateParallel {
    <#
    .SYNOPSIS
        Runs multiple update sections in parallel using Start-Job.
        Each section must be self-contained (no calls to parent script functions).
        Falls back to sequential if -NoParallel is set or PS version < 7.
    .PARAMETER Sections
        Array of hashtables: @{ Name; Title; RequiresCommand; RequiresAnyCommand; Disabled; SlowOperation; Action }
    #>
    param([hashtable[]]$Sections)

    # Sequential fallback
    if ($NoParallel -or $PSVersionTable.PSVersion.Major -lt 7) {
        foreach ($sec in $Sections) {
            # Inject $secName so parallel-style action blocks can reference it
            $escapedName = $sec.Name -replace "'", "''"
            $patchedAction = [scriptblock]::Create(
                "`$secName = '$escapedName'" + [System.Environment]::NewLine + $sec.Action.ToString()
            )
            $invokeArgs = @{ Name = $sec.Name; Action = $patchedAction }
            if ($sec.Title) { $invokeArgs['Title'] = $sec.Title }
            if ($sec.RequiresCommand) { $invokeArgs['RequiresCommand'] = $sec.RequiresCommand }
            if ($sec.RequiresAnyCommand) { $invokeArgs['RequiresAnyCommand'] = $sec.RequiresAnyCommand }
            if ($sec.ContainsKey('Disabled')) { $invokeArgs['Disabled'] = $sec.Disabled }
            if ($sec.ContainsKey('SlowOperation')) { $invokeArgs['SlowOperation'] = $sec.SlowOperation }
            Invoke-Update @invokeArgs
        }
        return
    }

    # Parallel path: Start-ThreadJob per section, collect output in order
    # Serialize helper functions so thread jobs can call them
    $fnInit = [scriptblock]::Create(@"
`$commandCache = @{}
function Test-Command([string]`$c) {
    if (`$commandCache.ContainsKey(`$c)) { return `$commandCache[`$c] }
    `$r = [bool](Get-Command `$c -ErrorAction SilentlyContinue)
    `$commandCache[`$c] = `$r; return `$r
}
function Write-Section([string]`$Title, [double]`$ElapsedSeconds = 0) {
    `$e = if (`$ElapsedSeconds -gt 0) { '  ({0:F1}s)' -f `$ElapsedSeconds } else { '' }
    Write-Host "``n`$('=' * 54)" -ForegroundColor DarkGray
    Write-Host " `$Title`$e"      -ForegroundColor Cyan
    Write-Host "`$('=' * 54)"     -ForegroundColor DarkGray
}
function Write-Status([string]`$Message, [ValidateSet('Success','Warning','Error','Info')][string]`$Type='Info') {
    `$c = @{ Success='Green'; Warning='Yellow'; Error='Red'; Info='Gray' }
    `$s = @{ Success='[OK]';  Warning='[!]';    Error='[X]'; Info='[*]' }
    Write-Host "`$(`$s[`$Type]) `$Message" -ForegroundColor `$c[`$Type]
}
function Write-FilteredOutput([AllowNull()][string]`$Text, [ConsoleColor]`$Color = [ConsoleColor]::Gray) {
    if ([string]::IsNullOrWhiteSpace(`$Text)) { return }
    `$n = (`$Text -replace '\x1b\[[0-9;?]*[ -/]*[@-~]', '') -replace '`r', '`n'
    foreach (`$raw in (`$n -split '`n')) {
        `$l = `$raw.TrimEnd(); if (-not `$l) { continue }
        `$c2 = `$l.Trim()
        if (`$c2 -match '^[\\\/\|\-]+`$')                                                            { continue }
        if (`$c2 -match '[█▓▒░▏▎▍▌▋▊▉■□▪▫]')                                                       { continue }
        if (`$c2 -match 'package\(s\) have version numbers that cannot be determined')               { continue }
        Write-Host `$l -ForegroundColor `$Color
    }
}
function Write-UpdateSummaryMarker([string]`$Section, [string]`$Summary, [switch]`$Changed) {
    if ([string]::IsNullOrWhiteSpace(`$Summary)) { return }
    `$encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(`$Summary.Trim()))
    `$changedFlag = if (`$Changed) { '1' } else { '0' }
    Write-Output "__SUMMARY__`${Section}__`${changedFlag}__`${encoded}"
}
function Write-UpdateItemMarker([string]`$Section, [string]`$Item, [ValidateSet('Updated','Info','Warning')][string]`$Kind='Updated') {
    if ([string]::IsNullOrWhiteSpace(`$Item)) { return }
    `$encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(`$Item.Trim()))
    Write-Output "__DETAIL__`${Section}__`${Kind}__`${encoded}"
}
function Invoke-WingetWithTimeout {
    param([string[]]`$Arguments, [int]`$TimeoutSec = 300)
    `$outFile = [System.IO.Path]::GetTempFileName()
    `$errFile = [System.IO.Path]::GetTempFileName()
    try {
        `$proc = Start-Process -FilePath 'winget' -ArgumentList `$Arguments ``
            -RedirectStandardOutput `$outFile -RedirectStandardError `$errFile -NoNewWindow -PassThru
        `$exited = `$proc.WaitForExit(`$TimeoutSec * 1000)
        if (-not `$exited) { try { `$proc.Kill() } catch { }; throw "winget timed out after `${TimeoutSec}s" }
        `$combined = ((Get-Content -Raw `$outFile -ErrorAction SilentlyContinue) + (Get-Content -Raw `$errFile -ErrorAction SilentlyContinue) -replace '\x00','').Trim()
        return [pscustomobject]@{ Output = `$combined; ExitCode = `$proc.ExitCode }
    } finally { Remove-Item `$outFile, `$errFile -Force -ErrorAction SilentlyContinue }
}
function Get-ToolInstallManager([string]`$Command) {
    `$src = (Get-Command `$Command -ErrorAction SilentlyContinue).Source
    if (-not `$src) { return `$null }
    if (`$src -like '*\scoop\*' -or `$src -like '*scoop\shims\*') { return 'scoop' }
    if (`$src -like '*\WinGet\*' -or `$src -like '*WindowsApps\*' -or `$src -like '*\winget\*') { return 'winget' }
    return `$null
}
function Invoke-WingetPackageUpgrade([string]`$Id, [int]`$TimeoutSec = 120) {
    return Invoke-WingetWithTimeout -TimeoutSec `$TimeoutSec -Arguments @(
        'upgrade', '--id', `$Id,
        '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
    )
}
"@)

    $jobs = [ordered]@{}   # Name -> Job
    $jobTitles = @{}            # Name -> Title

    foreach ($sec in $Sections) {
        $name = $sec.Name
        $title = if ($sec.Title) { $sec.Title } else { $name }
        $jobTitles[$name] = $title

        # Apply guards (same as Invoke-Update)
        if ($sec.Disabled) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'flag' })
            continue
        }
        if ($sec.RequiresCommand -and -not (Test-Command $sec.RequiresCommand)) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'not installed' })
            continue
        }
        if ($sec.RequiresAnyCommand) {
            $anyFound = $sec.RequiresAnyCommand | Where-Object { Test-Command $_ } | Select-Object -First 1
            if (-not $anyFound) {
                $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'not installed' })
                continue
            }
        }
        if ($sec.SlowOperation -and $FastMode) {
            $script:updateResults.Skipped.Add([pscustomobject]@{ Name = $name; Reason = 'fast mode' })
            continue
        }
        if ($DryRun) {
            Write-Host "  [DryRun] Would run: $title" -ForegroundColor DarkCyan
            continue
        }

        $actionStr = $sec.Action.ToString()   # pass as string; rebuilt with Create() in the thread job
        $secName = $name
        $secTitle = $title
        # Variables the action blocks may reference
        $varTable = @{
            WingetTimeoutSec = $WingetTimeoutSec
            SkipDestructive  = [bool]$SkipDestructive
        }
        $jobs[$name] = Start-ThreadJob -Name "update-$name" `
            -InitializationScript $fnInit `
            -ScriptBlock {
            param($actionStr, $secName, $secTitle, $varTable)
            # Inject needed script-scope variables so actions can use them directly
            foreach ($kv in $varTable.GetEnumerator()) { Set-Variable -Name $kv.Key -Value $kv.Value }
            # Reconstruct the scriptblock natively in this runspace so cmdlets resolve correctly
            $action = [ScriptBlock]::Create($actionStr)
            $t = Get-Date
            Write-Section $secTitle
            try {
                & $action
                $elapsed = ((Get-Date) - $t).TotalSeconds
                Write-Status ("{0} updated ({1:F1}s)" -f $secName, $elapsed) -Type Success
                # sentinel on stdout so main thread knows elapsed
                "__TIMING__${secName}__${elapsed}"
            }
            catch {
                Write-Host "[X] $secName failed: $($_.Exception.Message)" -ForegroundColor Red
                "__FAILED__${secName}"
            }
        } -ArgumentList $actionStr, $secName, $secTitle, $varTable
    }

    if ($jobs.Count -eq 0) { return }

    $null = Wait-Job -Job @($jobs.Values) -Timeout ([Math]::Max(600, $WingetTimeoutSec * 2))

    foreach ($name in $jobs.Keys) {
        $job = $jobs[$name]
        $jobSucceeded = $false
        $jobElapsed = $null

        # Replay Information-stream items (Write-Host calls from thread job)
        $rawItems = Receive-Job $job -ErrorAction SilentlyContinue 2>&1 6>&1
        foreach ($item in $rawItems) {
            if ($item -is [System.Management.Automation.InformationRecord]) {
                $hm = $item.MessageData
                if ($hm -is [System.Management.Automation.HostInformationMessage]) {
                    # ForegroundColor can be null or -1 in thread jobs — fall back to Gray safely
                    $fg = try {
                        $c = $hm.ForegroundColor
                        if ($null -ne $c -and [int]$c -ge 0 -and [int]$c -le 15) { [ConsoleColor]$c }
                        else { [ConsoleColor]::Gray }
                    }
                    catch { [ConsoleColor]::Gray }
                    try { Write-Host $hm.Message -ForegroundColor $fg -NoNewline:$hm.NoNewLine }
                    catch { Write-Host $hm.Message }
                }
                else {
                    Write-Host $hm.ToString()
                }
            }
            elseif ($item -is [string]) {
                if ($item -match '^__TIMING__(.+)__([0-9.,]+)$') {
                    $jobSucceeded = $true
                    $jobElapsed = [double]($Matches[2] -replace ',', '.')
                }
                elseif ($item -match '^__SUMMARY__(.+?)__(0|1)__(.+)$') {
                    Set-UpdateSummary -Name $Matches[1] -Summary (ConvertFrom-Base64Utf8 $Matches[3]) -Changed:($Matches[2] -eq '1')
                }
                elseif ($item -match '^__DETAIL__(.+?)__(Updated|Info|Warning)__(.+)$') {
                    Add-UpdateItem -Name $Matches[1] -Kind $Matches[2] -Item (ConvertFrom-Base64Utf8 $Matches[3])
                }
                elseif ($item -match '^__FAILED__(.+)$') {
                    $script:updateResults.Failed.Add($Matches[1])
                }
                # control strings are not printed
            }
            elseif ($null -ne $item) {
                try { Write-Host ($item | Out-String).TrimEnd() } catch { Write-Verbose "Failed to render job output item: $($_.Exception.Message)" }
            }
        }

        if ($job.State -eq 'Running') {
            Stop-Job $job -ErrorAction SilentlyContinue
            Write-Status "$name timed out" -Type Warning
            $script:updateResults.Failed.Add($name)
        }
        elseif ($jobSucceeded) {
            if (Test-SectionChanged -Name $name) {
                $script:updateResults.Success.Add($name)
            }
            else {
                $script:updateResults.Checked.Add($name)
            }
            if ($null -ne $jobElapsed) {
                $script:sectionTimings[$name] = $jobElapsed
            }
        }
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
}

# ════════════════════════════════════════════════════════════
#  PACKAGE MANAGERS
# ════════════════════════════════════════════════════════════

# ── Scoop ────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Scoop' -RequiresCommand 'scoop' -Action {
    $scoopSelfOut = (scoop update 2>&1 | Out-String).Trim()
    if ($scoopSelfOut -match '(?i)error|failed') { Write-Status "Scoop self-update warning: $scoopSelfOut" -Type Warning }
    $out = (scoop update '*' 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out
    scoop cleanup '*' 2>&1 | Out-Null
    scoop cache rm '*' 2>&1 | Out-Null
}

# Packages known to require a reboot to complete — retrying with --force can leave them broken.
$script:WingetRetryBlocklist = @(
    'Microsoft.VCRedist.2015+.x64',
    'Microsoft.VCRedist.2015+.x86',
    'Microsoft.DotNet.Runtime.8',
    'Microsoft.DotNet.Runtime.9',
    'Microsoft.DotNet.DesktopRuntime.8',
    'Microsoft.DotNet.DesktopRuntime.9',
    'Microsoft.WindowsTerminal'
)

# ── Winget ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Winget' -RequiresCommand 'winget' -Action {
    $bulkHadFailures = $false
    $retryIds = [System.Collections.Generic.List[string]]::new()
    $finalFailedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $skippedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pendingPackages = @()
    $pendingPackageMap = @{}
    $unknownPackages = @()
    $unknownUpgrades = [System.Collections.Generic.List[string]]::new()

    # Refresh package sources to ensure the latest catalog is available
    Write-Host "  Refreshing winget sources..." -ForegroundColor Gray
    try { Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('source', 'update') | Out-Null } catch { Write-Verbose "Winget source refresh failed: $($_.Exception.Message)" }

    # Run Pre hooks for packages with known conflicts (e.g. Spotify exit-code 23)
    try {
        $pendingCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity'
        )
        $pendingPackages = @(Get-WingetUpgradeCandidate -Output $pendingCheck.Output)
        foreach ($pkg in $pendingPackages) { $pendingPackageMap[$pkg.Id] = $pkg }
        $unknownPackages = @(Get-WingetUnknownUpgradeCandidate -Output $pendingCheck.Output)
        $currentUnknowns = @($unknownPackages | Where-Object { $_.InstalledLooksCurrent })
        $actionableUnknowns = @($unknownPackages | Where-Object { -not $_.InstalledLooksCurrent })
        if ($pendingPackages.Count -gt 0) {
            Write-Detail ("Pending winget upgrades: {0}" -f ($pendingPackages.Name -join ', ')) -Type Muted
        }
        if ($actionableUnknowns.Count -gt 0) {
            Write-Detail ("Pending winget unknown-version upgrades: {0}" -f ($actionableUnknowns.Name -join ', ')) -Type Muted
        }
        foreach ($pkg in $currentUnknowns) {
            Add-UpdateItem -Name 'Winget' -Kind Info -Item ('{0}: installed version already matches available {1} (winget reports installed version as Unknown)' -f $pkg.Name, $pkg.Available)
        }
        Invoke-WingetUpgradeHook -Phase 'Pre' -WingetOutput $pendingCheck.Output
    }
    catch { Write-Verbose "Winget preflight scan failed: $($_.Exception.Message)" }   # best-effort; don't block upgrades if list fails

    # Standard winget source — known-version packages only (safe, no unnecessary reinstalls)
    Write-Host "  Upgrading all (winget source, timeout: ${WingetTimeoutSec}s)..." -ForegroundColor Gray
    try {
        $result = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
            'upgrade', '--all', '--source', 'winget',
            '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity'
        )
        Write-FilteredOutput $result.Output
        Invoke-WingetUpgradeHook -Phase 'Post' -WingetOutput $result.Output
        foreach ($id in (Get-WingetFailedPackageId $result.Output)) {
            # Check if this specific package failed with 1632 (needs elevation)
            if (-not $isAdmin -and $result.Output -match [regex]::Escape($id) -and $result.Output -match 'Installer failed with exit code:\s*1632') {
                Write-Status "  $id requires admin to upgrade (exit code 1632) — skipping retry" -Type Warning
                $updateResults.Skipped.Add([pscustomobject]@{ Name = $id; Reason = 'requires admin' })
                $null = $skippedIds.Add($id)
                continue
            }
            $null = $finalFailedIds.Add($id)
            if (-not $retryIds.Contains($id)) { $retryIds.Add($id) }
        }
        if ($result.ExitCode -ne 0 -or $result.Output -match 'Installer failed with exit code') { $bulkHadFailures = $true }
    }
    catch {
        Write-Status "Winget (winget source): $_" -Type Warning
        $bulkHadFailures = $true
    }

    # Unknown-version packages — only upgrade when a NEW Available version appears.
    # Winget can't detect the installed version for some apps (e.g. Node.js, ImageMagick).
    # We track which Available version we last upgraded to and skip if unchanged.
    $unknownStateFile = Join-Path (Join-Path $env:LOCALAPPDATA 'Update-Everything') 'unknown_versions.json'
    $unknownState = @{}
    if (Test-Path $unknownStateFile) {
        try { $unknownState = Get-Content $unknownStateFile -Raw | ConvertFrom-Json -AsHashtable } catch { $unknownState = @{} }
    }
    try {
        $unknownCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget',
            '--accept-source-agreements', '--disable-interactivity'
        )
        # Parse unknown-version entries: lines where the Version column is "Unknown"
        # Winget IDs always contain dots (publisher.package), so we anchor on that pattern.
        $unknownPkgs = @(Get-WingetUnknownUpgradeCandidate -Output $unknownCheck.Output | Where-Object { -not $_.InstalledLooksCurrent })
        if ($unknownPkgs.Count -gt 0) {
            $toUpgrade = @()
            foreach ($pkg in $unknownPkgs) {
                $prev = $unknownState[$pkg.Id]
                if ($prev -ne $pkg.Available) {
                    $toUpgrade += $pkg
                }
            }
            if ($toUpgrade.Count -gt 0) {
                Write-Host "  Upgrading $($toUpgrade.Count) unknown-version package(s)..." -ForegroundColor Gray
                foreach ($pkg in $toUpgrade) {
                    Write-Host "    $($pkg.Id) → $($pkg.Available) (installed version unknown)" -ForegroundColor Gray
                    try {
                        $upResult = Invoke-WingetPackageUpgrade $pkg.Id $WingetTimeoutSec
                        Write-FilteredOutput $upResult.Output
                        if ($upResult.ExitCode -eq 0 -and $upResult.Output -notmatch 'Installer failed with exit code') {
                            # Record the version we just installed
                            $unknownState[$pkg.Id] = $pkg.Available
                            $unknownUpgrades.Add(('{0} -> {1} (installed version unknown)' -f $pkg.Id, $pkg.Available))
                        }
                        else {
                            $null = $finalFailedIds.Add($pkg.Id)
                        }
                    }
                    catch {
                        Write-Status "  Unknown-version upgrade of $($pkg.Id) failed: $_" -Type Warning
                        $null = $finalFailedIds.Add($pkg.Id)
                    }
                }
            }
            # Save state
            try {
                $dir = Split-Path $unknownStateFile
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction SilentlyContinue | Out-Null
                $unknownState | ConvertTo-Json -Compress | Set-Content -Path $unknownStateFile -Force
            }
            catch { Write-Verbose "Failed to save winget unknown-version state: $($_.Exception.Message)" }
        }
    }
    catch { Write-Verbose "Winget unknown-version scan failed: $($_.Exception.Message)" }

    # Retry only packages that actually failed during bulk execution.
    if ($retryIds.Count -gt 0) {
        Write-Host "  Retrying $($retryIds.Count) failed package(s) individually..." -ForegroundColor Gray
        foreach ($pkgId in $retryIds) {
            if ($script:WingetRetryBlocklist -contains $pkgId) {
                Write-Status "  Skipping retry for $pkgId (blocklisted — may require reboot)" -Type Warning
                continue
            }
            Write-Host "    Updating $pkgId..." -ForegroundColor Gray
            try {
                Invoke-WingetPackageHook -PackageId $pkgId -Phase 'Pre'
                Clear-WingetInstallerCache -PackageId $pkgId
                $retryResult = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
                    'upgrade', '--id', $pkgId,
                    '--accept-source-agreements', '--accept-package-agreements',
                    '--disable-interactivity', '--force'
                )
                if ($retryResult.ExitCode -eq 0 -and $retryResult.Output -notmatch 'Installer failed with exit code') {
                    $null = $finalFailedIds.Remove($pkgId)
                }
                elseif ($retryResult.Output -match 'Installer failed with exit code:\s*1632') {
                    # 1632 can mean MSI temp/access trouble, but repeated admin failures are
                    # commonly caused by reboot-pending servicing state.
                    if (-not $isAdmin) {
                        Write-Status "  $pkgId requires admin to upgrade (exit code 1632) — skipping" -Type Warning
                        $null = $finalFailedIds.Remove($pkgId)
                        $updateResults.Skipped.Add([pscustomobject]@{ Name = $pkgId; Reason = 'requires admin' })
                        $null = $skippedIds.Add($pkgId)
                    }
                    else {
                        $pendingRestartReasons = @(Get-PendingRestartReasons)
                        if ($pendingRestartReasons.Count -gt 0) {
                            Write-Status "  $pkgId hit MSI exit code 1632 and Windows is reporting reboot-pending state: $($pendingRestartReasons -join ', ')." -Type Warning
                            Write-Detail "Reboot the machine, then retry $pkgId." -Type Warning
                            $null = $finalFailedIds.Remove($pkgId)
                            $updateResults.Skipped.Add([pscustomobject]@{ Name = $pkgId; Reason = 'reboot required' })
                            $null = $skippedIds.Add($pkgId)
                            Add-UpdateItem -Name 'Winget' -Kind Warning -Item ("${pkgId}: reboot required before retry")
                        }
                        else {
                            $fallbackResult = Invoke-WingetMsiFallback -PackageId $pkgId -TimeoutSec $WingetTimeoutSec
                            if ($fallbackResult -and $fallbackResult.Success) {
                                $null = $finalFailedIds.Remove($pkgId)
                            }
                            else {
                                $null = $finalFailedIds.Add($pkgId)
                            }
                        }
                    }
                }
                else {
                    $null = $finalFailedIds.Add($pkgId)
                }
            }
            catch {
                Write-Status "  Retry of $pkgId timed out or failed: $_" -Type Warning
                $null = $finalFailedIds.Add($pkgId)
            }
            finally {
                Invoke-WingetPackageHook -PackageId $pkgId -Phase 'Post'
            }
        }
    }
    elseif ($bulkHadFailures) {
        Write-Status 'Winget reported a bulk failure, but no specific failed package IDs were detected for retry' -Type Warning
    }

    # If bulk mode aborted before updating later packages, retry any known-version
    # packages that are still pending after the normal retry flow.
    $postKnownRemaining = @()
    try {
        $postCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity'
        )
        $postKnownRemaining = @(Get-WingetUpgradeCandidate -Output $postCheck.Output)
    }
    catch {
        Write-Verbose "Winget post-check failed before targeted retries: $($_.Exception.Message)"
    }

    foreach ($pkg in $postKnownRemaining) {
        if ($skippedIds.Contains($pkg.Id) -or $finalFailedIds.Contains($pkg.Id)) { continue }
        if ($script:WingetRetryBlocklist -contains $pkg.Id) { continue }

        Write-Host "  Targeted retry for still-pending package $($pkg.Id)..." -ForegroundColor Gray
        try {
            Invoke-WingetPackageHook -PackageId $pkg.Id -Phase 'Pre'
            $targetedResult = Invoke-WingetPackageUpgrade $pkg.Id $WingetTimeoutSec
            Write-FilteredOutput $targetedResult.Output
            if ($targetedResult.Output -match 'No applicable upgrade found' -and $pkg.Id -eq 'ZedIndustries.Zed') {
                if (Invoke-ZedUserInstallerFallback -TimeoutSec $WingetTimeoutSec) {
                    $null = $finalFailedIds.Remove($pkg.Id)
                }
                else {
                    $null = $finalFailedIds.Add($pkg.Id)
                }
            }
            elseif ($targetedResult.ExitCode -ne 0 -or $targetedResult.Output -match 'Installer failed with exit code') {
                $null = $finalFailedIds.Add($pkg.Id)
            }
        }
        catch {
            Write-Status "  Targeted retry of $($pkg.Id) failed: $_" -Type Warning
            $null = $finalFailedIds.Add($pkg.Id)
        }
        finally {
            Invoke-WingetPackageHook -PackageId $pkg.Id -Phase 'Post'
        }
    }

    $finalKnownRemaining = @()
    $finalKnownRemainingIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        $finalCheck = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @(
            'upgrade', '--include-unknown', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity'
        )
        $finalKnownRemaining = @(Get-WingetUpgradeCandidate -Output $finalCheck.Output)
        foreach ($pkg in $finalKnownRemaining) { $null = $finalKnownRemainingIds.Add($pkg.Id) }
        foreach ($pkg in (Get-WingetUnknownUpgradeCandidate -Output $finalCheck.Output | Where-Object { $_.InstalledLooksCurrent })) {
            Add-UpdateItem -Name 'Winget' -Kind Info -Item ('{0}: installed version already matches available {1} (winget reports installed version as Unknown)' -f $pkg.Name, $pkg.Available)
        }
    }
    catch {
        Write-Verbose "Winget final verification failed: $($_.Exception.Message)"
    }

    foreach ($pkg in $pendingPackages) {
        if ($skippedIds.Contains($pkg.Id) -or $finalFailedIds.Contains($pkg.Id) -or $finalKnownRemainingIds.Contains($pkg.Id)) { continue }
        Add-UpdateItem -Name 'Winget' -Kind Updated -Item ('{0} {1} -> {2}' -f $pkg.Name, $pkg.Version, $pkg.Available)
    }
    foreach ($item in $unknownUpgrades) {
        Add-UpdateItem -Name 'Winget' -Kind Updated -Item $item
    }
    foreach ($pkgId in $finalFailedIds) {
        Add-UpdateItem -Name 'Winget' -Kind Warning -Item ("failed: $pkgId")
    }

    if ($pendingPackages.Count -eq 0 -and $unknownPackages.Count -eq 0 -and $unknownUpgrades.Count -eq 0 -and $finalFailedIds.Count -eq 0 -and $finalKnownRemainingIds.Count -eq 0) {
        Set-UpdateSummary -Name 'Winget' -Summary 'no package updates'
    }
}

# Save a winget package snapshot for -WhatChanged diffing
$script:WingetSnapshotDir = Join-Path $env:LOCALAPPDATA 'Update-Everything'
$script:WingetSnapshotCurrent = Join-Path $script:WingetSnapshotDir 'update_log.txt'
$script:WingetSnapshotPrev = Join-Path $script:WingetSnapshotDir 'update_log_prev.txt'
if ((-not $DryRun) -and (Test-Command 'winget')) {
    try {
        New-Item -ItemType Directory -Path $script:WingetSnapshotDir -Force -ErrorAction SilentlyContinue | Out-Null
        # Rotate: current -> prev
        if (Test-Path $script:WingetSnapshotCurrent) {
            Copy-Item -Path $script:WingetSnapshotCurrent -Destination $script:WingetSnapshotPrev -Force
        }
        # Capture current state
        $listResult = Invoke-WingetWithTimeout -TimeoutSec 60 -Arguments @('list', '--accept-source-agreements', '--disable-interactivity')
        if ($listResult.Output) {
            Set-Content -Path $script:WingetSnapshotCurrent -Value $listResult.Output -Force
        }
    }
    catch {
        Write-Verbose "Could not save winget snapshot: $_"
    }
}

# ── Sequential: Chocolatey (admin-only) ─────────────────────────────────────────
Invoke-Update -Name 'Chocolatey' -RequiresCommand 'choco' -RequiresAdmin -Action {
    $pendingPackages = @()
    try {
        $outdated = (choco outdated --limit-output 2>&1 | Out-String).Trim()
        foreach ($line in ($outdated -split '\r?\n')) {
            if ($line -notmatch '\|') { continue }
            $parts = $line.Split('|')
            if ($parts.Count -lt 3) { continue }
            $pendingPackages += [pscustomobject]@{
                Name      = $parts[0].Trim()
                Version   = $parts[1].Trim()
                Available = $parts[2].Trim()
            }
        }
    }
    catch { Write-Verbose "Chocolatey outdated package pre-scan failed: $($_.Exception.Message)" }

    $out = (choco upgrade all -y 2>&1 | Out-String).Trim()
    Write-FilteredOutput $out

    if ($pendingPackages.Count -gt 0) {
        foreach ($pkg in $pendingPackages) {
            Add-UpdateItem -Name 'Chocolatey' -Kind Updated -Item ('{0} {1} -> {2}' -f $pkg.Name, $pkg.Version, $pkg.Available)
        }
    }
    else {
        Set-UpdateSummary -Name 'Chocolatey' -Summary 'no package updates'
    }
}

# ════════════════════════════════════════════════════════════
#  WINDOWS COMPONENTS
# ════════════════════════════════════════════════════════════

# ── Windows Update ───────────────────────────────────────────────────────────────
if (-not $SkipWindowsUpdate) {
    if (-not $isAdmin) {
        $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'requires admin' })
    }
    elseif (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Invoke-Update -Name 'WindowsUpdate' -Title 'Windows Update' -RequiresAdmin -Action {
            Import-Module PSWindowsUpdate

            # 0x800704c7 (ERROR_CANCELLED) is usually caused by the WU agent aborting
            # due to a pending reboot or -AutoReboot conflicting with the install.
            # Fix: use -IgnoreReboot during install, then handle reboot separately.
            $maxRetries = 2
            $attempt = 0
            $succeeded = $false

            $wuTimeoutSec = 600   # 10 min hard limit to prevent indefinite hangs

            while ($attempt -lt $maxRetries -and -not $succeeded) {
                $attempt++
                try {
                    # Run Get-WindowsUpdate in a job with a timeout to prevent hangs
                    $wuJob = Start-Job -ScriptBlock {
                        Import-Module PSWindowsUpdate
                        $wuParams = @{
                            Install      = $true
                            AcceptAll    = $true
                            NotCategory  = 'Drivers'
                            IgnoreReboot = $true
                            RecurseCycle = 3
                            Verbose      = $false
                            Confirm      = $false
                        }
                        Get-WindowsUpdate @wuParams
                    }

                    Write-Detail "Windows Update can stay quiet for several minutes while the scan runs. Waiting up to ${wuTimeoutSec}s..." -Type Muted
                    $completed = $wuJob | Wait-Job -Timeout $wuTimeoutSec
                    if (-not $completed) {
                        $wuJob | Stop-Job -PassThru | Remove-Job -Force -ErrorAction SilentlyContinue
                        throw "Windows Update timed out after ${wuTimeoutSec}s — the WU service may be stuck. Try rebooting or run with -SkipWindowsUpdate."
                    }

                    # Check for errors in the job
                    $jobError = $wuJob.ChildJobs[0].Error
                    if ($jobError -and $jobError.Count -gt 0) {
                        $errMsg = $jobError[0].ToString()
                        $wuJob | Remove-Job -Force -ErrorAction SilentlyContinue
                        throw $errMsg
                    }

                    $results = $wuJob | Receive-Job
                    $wuJob | Remove-Job -Force -ErrorAction SilentlyContinue

                    if ($results) {
                        Write-Host "  Installed $($results.Count) update(s)." -ForegroundColor Gray
                        foreach ($update in @($results)) {
                            $kb = $null
                            foreach ($propName in @('KB', 'KBArticleID', 'KBArticleIDs', 'HotFixID')) {
                                if ($update.PSObject.Properties.Name -contains $propName) {
                                    $kbValue = $update.$propName
                                    if ($kbValue) {
                                        $kb = if ($kbValue -is [System.Array]) { ($kbValue -join ',') } else { [string]$kbValue }
                                        break
                                    }
                                }
                            }
                            $title = if ($update.PSObject.Properties.Name -contains 'Title') { [string]$update.Title } else { [string]$update }
                            Add-UpdateItem -Name 'WindowsUpdate' -Kind Updated -Item ($(if ($kb) { "$kb $title" } else { $title }))
                        }
                    }
                    else {
                        Write-Host "  Windows Update scan completed: no updates available." -ForegroundColor Gray
                        Set-UpdateSummary -Name 'WindowsUpdate' -Summary 'no updates available'
                    }
                    $succeeded = $true
                }
                catch {
                    $msg = $_.Exception.Message

                    # 0x800704c7 = ERROR_CANCELLED
                    if ($msg -match '0x800704c7' -or ($_.Exception.HResult -eq -2147023673)) {
                        Write-Status "Windows Update cancelled by system (0x800704c7), attempt $attempt/$maxRetries" -Type Warning

                        if ($attempt -lt $maxRetries) {
                            Write-Host "  Restarting Windows Update service before retry..." -ForegroundColor Gray
                            Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 5
                        }
                    }
                    elseif ($msg -match 'timed out') {
                        # Timeout — don't retry, just fail cleanly
                        throw
                    }
                    else {
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
    }
    else {
        Write-Section 'Windows Update'
        Write-Status 'PSWindowsUpdate module not found. Install with: Install-Module PSWindowsUpdate -Force' -Type Warning
        $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'not installed' })
    }
}
else {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'WindowsUpdate'; Reason = 'flag' })
}

# ── Microsoft Store Apps ─────────────────────────────────────────────────────────
Invoke-Update -Name 'StoreApps' -Title 'Microsoft Store Apps' -Disabled:$SkipStoreApps -Action {
    try {
        Write-Detail "Checking Microsoft Store app upgrades via winget"
        $result = Invoke-WingetWithTimeout -TimeoutSec $WingetTimeoutSec -Arguments @(
            'upgrade', '--source', 'msstore', '--all', '--silent',
            '--accept-package-agreements', '--accept-source-agreements'
        )
        $storePackages = @(Get-WingetUpgradeCandidate -Output $result.Output)
        if ($storePackages.Count -eq 0) {
            $storePackages = @(Get-WingetFoundPackage -Output $result.Output)
        }

        if ($result.Output -match 'No installed package found matching input criteria\.') {
            Set-UpdateSummary -Name 'StoreApps' -Summary 'no Microsoft Store app updates'
        }
        else {
            Write-FilteredOutput $result.Output
            foreach ($pkg in $storePackages) {
                Add-UpdateItem -Name 'StoreApps' -Kind Updated -Item $(if ($pkg.Version) {
                        '{0} {1} -> {2}' -f $pkg.Name, $pkg.Version, $pkg.Available
                    }
                    else {
                        '{0} -> {1}' -f $pkg.Name, $pkg.Available
                    })
            }
        }

        if ($result.ExitCode -ne 0) {
            Write-Status "winget msstore exited with code $($result.ExitCode)" -Type Warning
        }
    }
    catch {
        throw "StoreApps: $_"
    }
}

# ── WSL ──────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'WSL' -Title 'Windows Subsystem for Linux' -Disabled:$SkipWSL -RequiresCommand 'wsl' -RequiresAdmin -Action {
    # Update the WSL kernel/platform itself (sequential prerequisite)
    $out = (wsl --update 2>&1 | Out-String).Trim() -replace '\x00', ''
    if ($out -and $out -notmatch 'most recent version.+already installed') {
        Write-IndentedOutput $out -Prefix '  >' -Color ([ConsoleColor]::Gray)
    }
    if ($out -match 'most recent version.+already installed') {
        Add-UpdateItem -Name 'WSL' -Kind Info -Item 'platform already up to date'
    }
    else {
        Add-UpdateItem -Name 'WSL' -Kind Updated -Item 'WSL platform updated'
    }
    if ($LASTEXITCODE -ne 0) {
        throw "wsl --update failed with exit code $LASTEXITCODE"
    }

    # Optionally update packages inside each distro — in parallel via Start-Job
    if (-not $SkipWSLDistros) {
        $distros = wsl --list --quiet 2>&1 | Where-Object { $_ -and $_ -notmatch '^\s*$' }
        $distroNames = @($distros | ForEach-Object { ($_.Trim() -replace '\x00', '') } | Where-Object { $_ })

        if ($distroNames.Count -eq 0) {
            Set-UpdateSummary -Name 'WSL' -Summary 'platform current; no distros found'
        }
        else {
            Write-Detail "Updating packages in $($distroNames.Count) distro(s) in parallel"

            $jobs = @()
            foreach ($distroName in $distroNames) {
                $jobs += Start-Job -Name "wsl-$distroName" -ArgumentList $distroName -ScriptBlock {
                    param($dn)
                    function Invoke-WslPackageUpdate {
                        param(
                            [string]$DistroName,
                            [string]$Manager,
                            [string]$Command
                        )

                        $normalizedCommand = ($Command -replace "`r`n?", "`n").Trim()
                        $result = (wsl -d $DistroName -- sh -lc $normalizedCommand 2>&1 | Out-String).Trim()
                        $exitCode = $LASTEXITCODE

                        return [pscustomobject]@{
                            Distro   = $DistroName
                            Manager  = $Manager
                            ExitCode = $exitCode
                            Output   = $result
                        }
                    }

                    $probeToken = '__UPDATE_EVERYTHING_WSL_READY__'
                    $probeOutput = (wsl -d $dn -- sh -lc "printf '$probeToken'" 2>&1 | Out-String).Trim()
                    $probeExitCode = $LASTEXITCODE
                    if ($probeExitCode -ne 0 -or $probeOutput -notmatch [regex]::Escape($probeToken)) {
                        $normalizedProbeOutput = (($probeOutput -replace '\x00', '') -replace '\s+', ' ').Trim()
                        $probeReason = if ($normalizedProbeOutput -match 'ERROR_PATH_NOT_FOUND|MountDisk|Failed to attach disk') {
                            'broken-distro'
                        }
                        elseif ($normalizedProbeOutput -match 'E_ACCESSDENIED|Access is denied') {
                            'inaccessible'
                        }
                        else {
                            'unreachable'
                        }

                        return [pscustomobject]@{
                            Distro  = $dn
                            Success = $true
                            Manager = 'wsl'
                            Warning = $true
                            Reason  = $probeReason
                            Output  = $(if ($probeOutput) { $probeOutput } else { 'Failed to start WSL distro' })
                        }
                    }

                    $attempts = @(
                        @{
                            Manager = 'apt'
                            Command = @'
if ! command -v apt-get >/dev/null 2>&1; then
  exit 127
fi
prefix=""
if [ "$(id -u)" -eq 0 ]; then
  prefix=""
elif command -v sudo >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    prefix="sudo -n"
  else
    printf "sudo password is required for apt updates\n" >&2
    exit 126
  fi
else
  printf "sudo is required for apt updates\n" >&2
  exit 126
fi
$prefix apt-get update -qq &&
$prefix apt-get upgrade -y -qq 2>&1 | tail -5
'@
                        },
                        @{
                            Manager = 'pacman'
                            Command = @'
if ! command -v pacman >/dev/null 2>&1; then
  exit 127
fi
prefix=""
if [ "$(id -u)" -eq 0 ]; then
  prefix=""
elif command -v sudo >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    prefix="sudo -n"
  else
    printf "sudo password is required for pacman updates\n" >&2
    exit 126
  fi
else
  printf "sudo is required for pacman updates\n" >&2
  exit 126
fi
$prefix pacman -Syu --noconfirm 2>&1 | tail -5
'@
                        },
                        @{
                            Manager = 'zypper'
                            Command = @'
if ! command -v zypper >/dev/null 2>&1; then
  exit 127
fi
prefix=""
if [ "$(id -u)" -eq 0 ]; then
  prefix=""
elif command -v sudo >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    prefix="sudo -n"
  else
    printf "sudo password is required for zypper updates\n" >&2
    exit 126
  fi
else
  printf "sudo is required for zypper updates\n" >&2
  exit 126
fi
$prefix zypper refresh &&
$prefix zypper update -y 2>&1 | tail -5
'@
                        }
                    )

                    foreach ($attempt in $attempts) {
                        $result = Invoke-WslPackageUpdate -DistroName $dn -Manager $attempt.Manager -Command $attempt.Command
                        if ($result.ExitCode -eq 0) {
                            return [pscustomobject]@{
                                Distro  = $dn
                                Success = $true
                                Warning = $false
                                Manager = $attempt.Manager
                                Reason  = 'updated'
                                Output  = $result.Output
                            }
                        }

                        if ($result.ExitCode -eq 126) {
                            return [pscustomobject]@{
                                Distro  = $dn
                                Success = $true
                                Warning = $true
                                Manager = $attempt.Manager
                                Reason  = 'requires-elevation'
                                Output  = $result.Output
                            }
                        }

                        if ($result.ExitCode -ne 127) {
                            return [pscustomobject]@{
                                Distro  = $dn
                                Success = $false
                                Warning = $false
                                Manager = $attempt.Manager
                                Reason  = 'failed'
                                Output  = $result.Output
                            }
                        }
                    }

                    return [pscustomobject]@{
                        Distro  = $dn
                        Success = $true
                        Warning = $true
                        Manager = 'none'
                        Reason  = 'no-manager'
                        Output  = 'No supported package manager detected (apt, pacman, zypper)'
                    }
                }
            }

            # Wait for all distro jobs to complete (5 min timeout)
            $null = Wait-Job -Job $jobs -Timeout 300
            $wslFailures = [System.Collections.Generic.List[string]]::new()

            foreach ($job in $jobs) {
                if ($job.State -eq 'Running') {
                    Stop-Job -Job $job -ErrorAction SilentlyContinue
                    Write-Status "  $($job.Name) timed out" -Type Warning
                    $null = $wslFailures.Add($job.Name -replace '^wsl-', '')
                    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                    continue
                }

                $jobResult = Receive-Job -Job $job -ErrorAction SilentlyContinue
                foreach ($result in @($jobResult)) {
                    if ($null -eq $result) { continue }

                    $label = "[$($result.Distro)] $($result.Manager)"
                    if ($result.Success -and -not $result.Warning) {
                        Write-Detail "$label update completed" -Type Info
                        Write-IndentedOutput $result.Output -Prefix '    -' -Color ([ConsoleColor]::DarkGray)
                        $pkgNames = @()
                        foreach ($line in ($result.Output -split '\r?\n')) {
                            if ($line -match '(?i)\bupgrading\s+([A-Za-z0-9._+\-]+)') {
                                $pkgNames += $Matches[1]
                            }
                        }
                        if ($pkgNames.Count -gt 0) {
                            Add-UpdateItem -Name 'WSL' -Kind Updated -Item ('{0}: {1}' -f $result.Distro, (($pkgNames | Select-Object -Unique) -join ', '))
                        }
                        else {
                            Add-UpdateItem -Name 'WSL' -Kind Info -Item ("$($result.Distro): packages updated")
                        }
                    }
                    elseif ($result.Warning) {
                        if ($result.Reason -eq 'requires-elevation') {
                            Write-Detail "[$($result.Distro)] $($result.Manager) requires non-interactive sudo/root; skipped" -Type Warning
                            Write-IndentedOutput $result.Output -Prefix '    !' -Color ([ConsoleColor]::Yellow)
                            Add-UpdateItem -Name 'WSL' -Kind Warning -Item ("$($result.Distro): $($result.Manager) requires passwordless sudo or root")
                        }
                        elseif ($result.Reason -in @('broken-distro', 'inaccessible', 'unreachable')) {
                            $issueLabel = switch ($result.Reason) {
                                'broken-distro' { 'WSL distro is broken or missing its backing disk' }
                                'inaccessible' { 'WSL distro could not be accessed from this session' }
                                default { 'WSL distro could not be started' }
                            }
                            Write-Detail "[$($result.Distro)] $issueLabel; skipped" -Type Warning
                            Write-IndentedOutput ($result.Output -replace '\x00', '') -Prefix '    !' -Color ([ConsoleColor]::Yellow)
                            Add-UpdateItem -Name 'WSL' -Kind Warning -Item ("$($result.Distro): $issueLabel")
                        }
                        else {
                            Write-Detail "[$($result.Distro)] No supported package manager detected" -Type Warning
                            Add-UpdateItem -Name 'WSL' -Kind Warning -Item ("$($result.Distro): no supported package manager")
                        }
                    }
                    else {
                        Write-Detail "$label update failed" -Type Error
                        Write-IndentedOutput $result.Output -Prefix '    x' -Color ([ConsoleColor]::Red)
                        $failureLine = (($result.Output -split '\r?\n' | Where-Object { $_ } | Select-Object -First 1) -join '').Trim()
                        Add-UpdateItem -Name 'WSL' -Kind Warning -Item $(if ($failureLine) {
                                '{0}: {1}' -f $result.Distro, $failureLine
                            }
                            else {
                                "$($result.Distro): update failed"
                            })
                    }

                    if (-not $result.Success) {
                        $null = $wslFailures.Add($result.Distro)
                    }
                }

                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }

            if ($wslFailures.Count -gt 0) {
                throw "WSL distro update failed for: $($wslFailures -join ', ')"
            }
        }
    }
}

# ── Defender Signatures ──────────────────────────────────────────────────────────
Invoke-Update -Name 'DefenderSignatures' -Title 'Microsoft Defender Signatures' -Disabled:$SkipDefender -RequiresCommand 'Update-MpSignature' -RequiresAdmin -Action {
    $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($mpStatus -and -not ($mpStatus.AMServiceEnabled -and $mpStatus.AntivirusEnabled)) {
        Set-UpdateSummary -Name 'DefenderSignatures' -Summary 'Microsoft Defender is not the active AV'
        return
    }
    try {
        Update-MpSignature -ErrorAction Stop 2>&1 | Out-Null
        Add-UpdateItem -Name 'DefenderSignatures' -Kind Updated -Item 'signatures updated'
    }
    catch {
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
#  DEVELOPMENT TOOLS (parallel batch)
# ════════════════════════════════════════════════════════════
Write-Host "`n$(if (-not $NoParallel -and $PSVersionTable.PSVersion.Major -ge 7) { '[Parallel]' } else { '[Sequential]' }) Running dev tool updates..." -ForegroundColor DarkCyan

Invoke-UpdateParallel @(
    @{
        Name = 'npm'; Title = 'npm (Node.js)'; RequiresCommand = 'npm'; Disabled = $SkipNode
        Action = {
            $currentNpm = (npm --version 2>&1).Trim()
            $latestNpm = (npm view npm version 2>&1).Trim()
            if ($currentNpm -ne $latestNpm) {
                Write-UpdateItemMarker $secName ("npm $currentNpm -> $latestNpm")
                npm install -g npm@latest 2>&1 | Out-Null
            }
            $outdatedJson = (npm outdated -g --json 2>$null | Out-String).Trim()
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($outdated -and $outdated.PSObject.Properties.Count -gt 0) {
                    $pkgs = $outdated.PSObject.Properties.Name
                    foreach ($pkg in $pkgs) {
                        $meta = $outdated.PSObject.Properties[$pkg].Value
                        Write-UpdateItemMarker $secName ('{0} {1} -> {2}' -f $pkg, $meta.current, $meta.latest)
                    }
                    Write-Host "  Updating $($pkgs.Count) package(s): $($pkgs -join ', ')" -ForegroundColor Gray
                    npm install -g $pkgs 2>&1 | Out-Null
                }
                else { Write-UpdateSummaryMarker $secName 'no global package updates' }
            }
            else {
                Write-UpdateSummaryMarker $secName 'no global package updates'
            }
            npm cache clean --force 2>&1 | Out-Null
        }
    }
    @{
        Name = 'pnpm'; RequiresCommand = 'pnpm'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $out = (pnpm update -g 2>&1 | Out-String).Trim()
            if ($out -match 'No global packages found') {
                Write-UpdateSummaryMarker $secName 'no global packages found'
            }
            elseif ($out -match 'Already up to date|Nothing to (do|upgrade)') {
                Write-UpdateSummaryMarker $secName 'no global package updates'
            }
            elseif ($out) {
                Write-Host $out -ForegroundColor Gray
                Write-UpdateSummaryMarker $secName 'global packages updated' -Changed
            }
        }
    }
    @{
        Name = 'Bun'; RequiresCommand = 'bun'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $out = (bun upgrade 2>&1 | Out-String).Trim()
            if ($out) {
                Write-Host $out -ForegroundColor Gray
                Write-UpdateSummaryMarker $secName 'updated' -Changed
            }
        }
    }
    @{
        Name = 'Deno'; RequiresCommand = 'deno'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $mgr = Get-ToolInstallManager 'deno'
            if ($mgr) { Write-UpdateSummaryMarker $secName "managed by $mgr (already updated)"; return }
            $env:NO_COLOR = '1'
            try {
                $out = (deno upgrade 2>&1 | Out-String).Trim()
                if ($out -match 'up to date|already the latest') {
                    Write-UpdateSummaryMarker $secName 'already up to date'
                }
                elseif ($out) {
                    Write-Host $out -ForegroundColor Gray
                    Write-UpdateSummaryMarker $secName 'updated' -Changed
                }
            }
            finally { Remove-Item Env:\NO_COLOR -ErrorAction SilentlyContinue }
        }
    }
    @{
        Name = 'Rust'; RequiresCommand = 'rustup'; Disabled = $SkipRust
        Action = {
            $out = (rustup update 2>&1 | Out-String).Trim()
            if ($out -match 'unchanged - rustc\s+([0-9][^\s)]*)') {
                Write-UpdateSummaryMarker $secName ("stable unchanged (rustc $($Matches[1]))")
            }
            elseif ($out -match 'updated - rustc\s+([0-9][^\s)]*)') {
                if ($out) { Write-Host $out -ForegroundColor Gray }
                Write-UpdateSummaryMarker $secName ("stable updated (rustc $($Matches[1]))") -Changed
            }
            elseif ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    @{
        Name = 'Go'; RequiresCommand = 'go'; Disabled = $SkipGo
        Action = {
            $mgr = Get-ToolInstallManager 'go'
            if ($mgr) { Write-UpdateSummaryMarker $secName "managed by $mgr (already updated)" }
            else {
                try {
                    $r = Invoke-WingetPackageUpgrade 'GoLang.Go' $WingetTimeoutSec
                    if ($r.ExitCode -eq 0 -and $r.Output -match 'Successfully installed') {
                        Write-Host "  Updated via winget (GoLang.Go)" -ForegroundColor Gray
                        Write-UpdateSummaryMarker $secName 'updated via winget (GoLang.Go)' -Changed
                    }
                    else { Write-UpdateSummaryMarker $secName 'no newer version available or not managed by winget' }
                }
                catch { Write-Status "Go winget upgrade: $_" -Type Warning }
            }
            if (-not $SkipDestructive) {
                Write-Host "  Cleaning module cache..." -ForegroundColor Gray
                go clean -modcache 2>&1 | Out-Null
            }
        }
    }
    @{
        Name = 'gh-extensions'; Title = 'GitHub CLI Extensions'; RequiresCommand = 'gh'
        Action = {
            $ghExt = (gh extension list 2>&1 | Out-String).Trim()
            if ($ghExt) {
                gh extension upgrade --all 2>&1 | Out-Null
                Write-UpdateSummaryMarker $secName 'extensions checked and upgraded if needed'
            }
            else {
                Write-UpdateSummaryMarker $secName 'no gh extensions installed'
            }
        }
    }
    @{
        Name = 'pipx'; Title = 'pipx (Python CLI Tools)'; RequiresCommand = 'pipx'
        Action = { pipx upgrade-all 2>&1 | Out-Null }
    }
    @{
        Name = 'Poetry'; Title = 'Poetry (Python Packaging)'; RequiresCommand = 'poetry'; Disabled = $SkipPoetry
        Action = {
            $out = (poetry self update 2>&1 | Out-String).Trim()
            if ($out -and $out -notmatch 'already using the latest version|up to date') {
                Write-Host $out -ForegroundColor Gray
                Write-UpdateSummaryMarker $secName 'updated' -Changed
            }
        }
    }
    @{
        Name = 'Composer'; Title = 'Composer (PHP)'; RequiresCommand = 'composer'; Disabled = $SkipComposer
        Action = {
            $selfOut = (composer self-update --no-interaction 2>&1 | Out-String).Trim()
            if ($selfOut) { Write-Host $selfOut -ForegroundColor Gray }
            $globalOut = (composer global update --no-interaction 2>&1 | Out-String).Trim()
            if ($globalOut) { Write-Host $globalOut -ForegroundColor Gray }
            if ($selfOut -or $globalOut) { Write-UpdateSummaryMarker $secName 'updated' -Changed }
        }
    }
    @{
        Name = 'RubyGems'; RequiresCommand = 'gem'; Disabled = $SkipRuby; SlowOperation = $true
        Action = {
            $sysOut = (gem update --system 2>&1 | Out-String).Trim()
            if ($sysOut) { Write-Host $sysOut -ForegroundColor Gray }
            $gemOut = (gem update 2>&1 | Out-String).Trim()
            if ($gemOut) { Write-Host $gemOut -ForegroundColor Gray }
            if ($sysOut -or $gemOut) { Write-UpdateSummaryMarker $secName 'updated' -Changed }
        }
    }
    @{
        Name = 'yt-dlp'; RequiresCommand = 'yt-dlp'
        Action = {
            $ytdlpPath = (Get-Command yt-dlp).Source
            if ($ytdlpPath -like '*scoop*') { Write-UpdateSummaryMarker $secName 'managed by Scoop (already updated)' }
            elseif ($ytdlpPath -like '*pip*' -or $ytdlpPath -like '*Python*' -or $ytdlpPath -like '*Scripts*') {
                if (Test-Command 'python') { python -m pip install --upgrade yt-dlp 2>&1 | Out-Null; Write-Host "  Updated via pip" -ForegroundColor Gray; Write-UpdateSummaryMarker $secName 'updated via pip' -Changed }
                else { Write-Status "yt-dlp installed via pip but Python not found" -Type Warning }
            }
            else {
                $out = (yt-dlp -U 2>&1 | Out-String).Trim()
                if ($out -match 'up to date') {
                    Write-UpdateSummaryMarker $secName 'already up to date'
                }
                elseif ($out) {
                    Write-Host $out -ForegroundColor Gray
                    Write-UpdateSummaryMarker $secName 'updated' -Changed
                }
            }
        }
    }
    @{
        Name = 'tldr'; Title = 'tldr Cache'; RequiresAnyCommand = @('tldr', 'tealdeer')
        Action = {
            $tldrCmd = if (Test-Command 'tealdeer') { 'tealdeer' } else { 'tldr' }
            $out = (& $tldrCmd --update 2>&1 | Out-String).Trim()
            if ($out) {
                Write-Host $out -ForegroundColor Gray
                Write-UpdateSummaryMarker $secName 'cache updated' -Changed
            }
        }
    }
    @{
        Name = 'oh-my-posh'; Title = 'Oh My Posh'; RequiresCommand = 'oh-my-posh'
        Action = {
            $mgr = Get-ToolInstallManager 'oh-my-posh'
            if ($mgr) { Write-UpdateSummaryMarker $secName "managed by $mgr (already updated)" }
            else {
                $out = (oh-my-posh upgrade 2>&1 | Out-String).Trim()
                if ($out -match '(?i)not supported|error|failed') { Write-Status "oh-my-posh upgrade failed: $out" -Type Warning }
                elseif ($out) { Write-Host $out -ForegroundColor Gray; Write-UpdateSummaryMarker $secName 'updated' -Changed }
            }
        }
    }
    @{
        Name = 'Volta'; Title = 'Volta (Node Version Manager)'; RequiresCommand = 'volta'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            if ((Get-ToolInstallManager 'volta') -eq 'scoop') { Write-UpdateSummaryMarker $secName 'managed by scoop (already updated)' }
            else { Write-UpdateSummaryMarker $secName 'standalone install; update via its installer' }
        }
    }
    @{
        Name = 'fnm'; Title = 'fnm (Fast Node Manager)'; RequiresCommand = 'fnm'; Disabled = $SkipNode; SlowOperation = $true
        Action = {
            $mgr = Get-ToolInstallManager 'fnm'
            if ($mgr) { Write-UpdateSummaryMarker $secName "managed by $mgr (already updated)" }
            else { Write-UpdateSummaryMarker $secName 'standalone install; update via installer or Scoop' }
        }
    }
    @{
        Name = 'mise'; Title = 'mise (Tool Version Manager)'; RequiresCommand = 'mise'; SlowOperation = $true
        Action = {
            $mgr = Get-ToolInstallManager 'mise'
            if (-not $mgr) {
                $out = (mise self-update --yes 2>&1 | Out-String).Trim()
                if ($out) { Write-Host $out -ForegroundColor Gray }
            }
            $pluginsOut = (mise plugins ls 2>&1 | Out-String).Trim()
            if ($pluginsOut) {
                Write-Host "  Upgrading mise plugins..." -ForegroundColor Gray
                mise plugins upgrade 2>&1 | Out-Null
            }
            Write-Host "  Upgrading mise-managed runtimes..." -ForegroundColor Gray
            $upgradeOut = (mise upgrade 2>&1 | Out-String).Trim()
            if ($upgradeOut) { Write-Host $upgradeOut -ForegroundColor Gray }
            if ($pluginsOut -or $upgradeOut) { Write-UpdateSummaryMarker $secName 'updated' -Changed }
        }
    }
    @{
        Name = 'juliaup'; Title = 'Julia (juliaup)'; RequiresCommand = 'juliaup'; SlowOperation = $true
        Action = {
            $mgr = Get-ToolInstallManager 'juliaup'
            if ($mgr) { Write-UpdateSummaryMarker $secName "managed by $mgr (already updated)" }
            else {
                $out = (juliaup update 2>&1 | Out-String).Trim()
                if ($out) { Write-Host $out -ForegroundColor Gray; Write-UpdateSummaryMarker $secName 'updated' -Changed }
            }
        }
    }
    @{
        Name = 'Ollama Models'; Title = 'Ollama Models'; RequiresCommand = 'ollama'
        Disabled = (-not $UpdateOllamaModels); SlowOperation = $true
        Action = {
            $listOut = (ollama list 2>&1 | Out-String).Trim()
            if (-not $listOut -or $listOut -match '(?i)error') { Write-Status 'Could not list Ollama models' -Type Warning; return }
            $modelNames = ($listOut -split "`n" | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] } | Where-Object { $_ -and $_ -notmatch '^\s*$' })
            if (-not $modelNames -or $modelNames.Count -eq 0) { Write-Host "  No Ollama models installed" -ForegroundColor Gray; return }
            Write-Host "  Pulling updates for $($modelNames.Count) model(s)..." -ForegroundColor Gray
            $updatedCount = 0; $currentCount = 0
            foreach ($model in $modelNames) {
                Write-Host "    Pulling $model..." -ForegroundColor Gray
                $pullOut = (ollama pull $model 2>&1 | Out-String).Trim()
                if ($pullOut -match '(?i)up to date') { $currentCount++ } else { $updatedCount++ }
            }
            Write-Host "  $updatedCount updated, $currentCount already current" -ForegroundColor Gray
        }
    }
    @{
        Name = 'git-lfs'; Title = 'Git LFS'; RequiresCommand = 'git-lfs'; Disabled = $SkipGitLFS
        Action = {
            if ((Get-ToolInstallManager 'git-lfs') -eq 'scoop') { Write-UpdateSummaryMarker $secName 'managed by scoop (already updated)' }
            else {
                try {
                    Invoke-WingetPackageUpgrade 'GitHub.GitLFS' 120 | Out-Null
                    Write-Host "  Updated via winget (GitHub.GitLFS)" -ForegroundColor Gray
                    Write-UpdateSummaryMarker $secName 'updated via winget (GitHub.GitLFS)' -Changed
                }
                catch { Write-Status "git-lfs winget upgrade: $_" -Type Warning }
            }
        }
    }
    @{
        Name = 'git-credential-manager'; Title = 'Git Credential Manager'; RequiresCommand = 'git-credential-manager'
        Action = {
            if ((Get-ToolInstallManager 'git-credential-manager') -eq 'scoop') { Write-UpdateSummaryMarker $secName 'managed by scoop (already updated)' }
            else {
                try {
                    # GCM ships with Git for Windows; upgrading Git.Git updates GCM
                    Invoke-WingetPackageUpgrade 'Git.Git' $WingetTimeoutSec | Out-Null
                    Write-Host "  Updated via winget (Git.Git)" -ForegroundColor Gray
                    Write-UpdateSummaryMarker $secName 'updated via winget (Git.Git)' -Changed
                }
                catch { Write-Status "GCM winget upgrade: $_" -Type Warning }
            }
        }
    }
)

# ── Python / pip ─────────────────────────────────────────────────────────────────
$pythonCmd = if (Test-Command 'python') { 'python' }
else { 314, 313, 312, 311, 310 | ForEach-Object { "C:\Program Files\Python$_\python.exe" } | Where-Object { Test-Path $_ } | Select-Object -First 1 }
if ($pythonCmd) {
    Invoke-Update -Name 'pip' -Title 'Python / pip' -Action {
        Write-Host "  Using: $pythonCmd" -ForegroundColor Gray
        try {
            $pipBeforeRaw = (& $pythonCmd -m pip --version 2>&1 | Out-String).Trim()
            $pipBefore = if ($pipBeforeRaw -match 'pip\s+([0-9][^\s]*)') { $Matches[1] } else { $null }
        }
        catch {
            $pipBefore = $null
        }

        & $pythonCmd -m pip install --upgrade pip 2>&1 | Out-Null

        try {
            $pipAfterRaw = (& $pythonCmd -m pip --version 2>&1 | Out-String).Trim()
            $pipAfter = if ($pipAfterRaw -match 'pip\s+([0-9][^\s]*)') { $Matches[1] } else { $null }
            if ($pipBefore -and $pipAfter -and $pipBefore -ne $pipAfter) {
                Add-UpdateItem -Name 'pip' -Kind Updated -Item ("pip $pipBefore -> $pipAfter")
            }
        }
        catch { Write-Verbose "Unable to determine pip version after self-upgrade: $($_.Exception.Message)" }

        # Upgrade outdated global packages
        try {
            $outdatedJson = (& $pythonCmd -m pip list --outdated --format=json 2>$null | Out-String).Trim()
            if ($outdatedJson) {
                $outdated = $outdatedJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($outdated -and @($outdated).Count -gt 0) {
                    $pkgNames = @($outdated | ForEach-Object { $_.name })
                    Write-Host "  Upgrading $($pkgNames.Count) outdated package(s): $($pkgNames -join ', ')" -ForegroundColor Gray
                    foreach ($pkg in @($outdated)) {
                        Add-UpdateItem -Name 'pip' -Kind Updated -Item ('{0} {1} -> {2}' -f $pkg.name, $pkg.version, $pkg.latest_version)
                        & $pythonCmd -m pip install --upgrade $pkg.name 2>&1 | Out-Null
                    }
                }
                else {
                    if (-not (Get-UpdateSummaryText -Name 'pip')) {
                        Set-UpdateSummary -Name 'pip' -Summary 'no global pip package updates'
                    }
                }
            }
            elseif (-not (Get-UpdateSummaryText -Name 'pip')) {
                Set-UpdateSummary -Name 'pip' -Summary 'no global pip package updates'
            }
        }
        catch {
            Write-Status "pip global upgrade check failed: $_" -Type Warning
        }
    }
}
else {
    Write-Section 'Python / pip'
    Write-Status 'Python not found' -Type Warning
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'pip'; Reason = 'not installed' })
}

# ── uv ───────────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv' -Title 'UV Package Manager' -RequiresCommand 'uv' -Action {
    $uvPath = (Get-Command uv).Source
    if ($uvPath -like "*scoop*") { Set-UpdateSummary -Name 'uv' -Summary 'managed by scoop (already updated)' }
    elseif ($uvPath -like "*pip*" -or $uvPath -like "*Python*") { Set-UpdateSummary -Name 'uv' -Summary 'managed by pip' }
    else {
        $out = (uv self update 2>&1 | Out-String).Trim()
        if ($out -match 'error') {
            Write-Host "  uv self-update not supported; trying winget..." -ForegroundColor Gray
            Invoke-WingetPackageUpgrade 'astral-sh.uv' $WingetTimeoutSec | Out-Null
            Set-UpdateSummary -Name 'uv' -Summary 'updated via winget fallback' -Changed
        }
        elseif ($out) {
            if ($out -notmatch 'already up to date|latest version') { Write-Host $out -ForegroundColor Gray }
            Set-UpdateSummary -Name 'uv' -Summary $(if ($out -match 'already up to date|latest version') { 'already up to date' } else { 'updated' }) -Changed:($out -notmatch 'already up to date|latest version')
        }
    }
}

# ── uv tools ─────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv-tools' -Title 'uv Tool Installs' -Disabled:$SkipUVTools -RequiresCommand 'uv' -SlowOperation -Action {
    $out = (uv tool upgrade --all 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0 -and $out -match '(?i)unknown command|unrecognized') {
        Set-UpdateSummary -Name 'uv-tools' -Summary 'uv tool upgrade not supported by this version'
        return
    }
    if ($out -and $out -notmatch 'Nothing to upgrade|Nothing to do') { Write-Host $out -ForegroundColor Gray }
    if ($out -match 'Nothing to upgrade|Nothing to do') {
        Set-UpdateSummary -Name 'uv-tools' -Summary 'no uv tool updates'
    }
    elseif ($out) {
        Set-UpdateSummary -Name 'uv-tools' -Summary 'uv tools updated' -Changed
    }
}

# ── Claude CLI ────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Claude CLI' -Title 'Claude CLI (Anthropic)' -RequiresCommand 'npm' -Action {
    $currentVer = $null
    try {
        $rawVer = (claude --version 2>&1 | Out-String).Trim()
        # Strip labels like " (Claude Code)" to get the bare semver
        if ($rawVer -match '[\d]+\.[\d]+\.[\d]+') { $currentVer = $Matches[0] }
    }
    catch { Write-Verbose "Unable to determine installed Claude CLI version: $($_.Exception.Message)" }
    $latestVer = (npm show @anthropic-ai/claude-code version 2>&1 | Out-String).Trim()

    if (-not $latestVer -or $latestVer -match '(?i)error|ERR!') {
        Write-Status 'Could not determine latest Claude CLI version from npm' -Type Warning
        return
    }

    if (-not $currentVer -or $currentVer -ne $latestVer) {
        Write-Host "  Updating Claude CLI: $currentVer -> $latestVer" -ForegroundColor Gray
        Add-UpdateItem -Name 'Claude CLI' -Kind Updated -Item ("Claude CLI $currentVer -> $latestVer")

        # Check if the npm global prefix is in a protected location (requires admin)
        $npmPrefix = (npm prefix -g 2>&1 | Out-String).Trim()
        $needsAdmin = (-not $isAdmin) -and ($npmPrefix -like "${env:ProgramFiles}*" -or $npmPrefix -like "${env:ProgramFiles(x86)}*" -or $npmPrefix -like "C:\Program Files*")

        if ($needsAdmin) {
            # Try installing to the user-level npm prefix instead
            $userNpmDir = Join-Path $env:APPDATA 'npm'
            if (-not (Test-Path $userNpmDir)) { New-Item -ItemType Directory -Path $userNpmDir -Force | Out-Null }
            Write-Host "  npm global prefix is admin-protected; installing to user prefix ($userNpmDir)..." -ForegroundColor Gray
            $out = (npm install -g @anthropic-ai/claude-code --prefix "$userNpmDir" 2>&1 | Out-String).Trim()
            if ($out -match '(?i)EPERM|EACCES|error') {
                Write-Status "Claude CLI update requires admin (npm prefix: $npmPrefix). Run elevated or change npm prefix." -Type Warning
            }
            elseif ($out) { Write-Host $out -ForegroundColor Gray }
        }
        else {
            $out = (npm install -g @anthropic-ai/claude-code 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
        }
    }
    else {
        Set-UpdateSummary -Name 'Claude CLI' -Summary "already up to date ($currentVer)"
    }
}

# ── uv Python versions ────────────────────────────────────────────────────────────
Invoke-Update -Name 'uv-python' -Title 'uv Python Versions' -Disabled:$SkipUVTools -RequiresCommand 'uv' -SlowOperation -Action {
    $pyList = (uv python list --only-installed 2>&1 | Out-String).Trim()
    if (-not $pyList -or $pyList -match '(?i)no python|error') {
        Set-UpdateSummary -Name 'uv-python' -Summary 'no uv-managed Python versions found'
        return
    }
    Write-Host "  Upgrading uv-managed Python installs..." -ForegroundColor Gray

    # Extract unique major.minor versions from uv python list output and reinstall
    # to pull the latest patch releases.  Lines look like:
    #   cpython-3.13.2-windows-x86_64-none    C:\Users\...\python.exe
    $majMinors = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($line in ($pyList -split '\r?\n')) {
        if ($line -match 'cpython-(\d+\.\d+)') {
            $null = $majMinors.Add($Matches[1])
        }
    }

    if ($majMinors.Count -eq 0) {
        Set-UpdateSummary -Name 'uv-python' -Summary 'no uv-managed Python installs detected'
        return
    }

    $versions = @($majMinors | Sort-Object)
    Write-Host "  Reinstalling latest patches for: $($versions -join ', ')" -ForegroundColor Gray
    $out = (uv python install @versions 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
    Set-UpdateSummary -Name 'uv-python' -Summary ('reinstalled latest patches for: {0}' -f ($versions -join ', ')) -Changed
}

# ── Cargo global binaries ────────────────────────────────────────────────────────
$skipCargoReason = if ($SkipRust) { 'flag' } elseif (-not (Test-Command 'cargo')) { 'not installed' } elseif ($FastMode) { 'fast mode' } else { $null }
if (-not $skipCargoReason) {
    $hasInstallUpdate = $false
    try { $null = cargo install-update --version 2>$null; $hasInstallUpdate = ($LASTEXITCODE -eq 0) } catch { Write-Verbose "cargo-install-update probe failed: $($_.Exception.Message)" }
    if ($hasInstallUpdate) {
        Invoke-Update -Name 'cargo-binaries' -Title 'Cargo Global Binaries' -Action {
            $out = (cargo install-update -a 2>&1 | Out-String).Trim()
            if ($out) { Write-Host $out -ForegroundColor Gray }
            Set-UpdateSummary -Name 'cargo-binaries' `
                -Summary $(if ($out -match 'No packages need updating') { 'no cargo binary updates' } else { 'cargo binaries updated' }) `
                -Changed:($out -and $out -notmatch 'No packages need updating')
        }
    }
    else { $skipCargoReason = 'not installed' }
}
if ($skipCargoReason) {
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cargo-binaries'; Reason = $skipCargoReason })
}

# ── .NET tools ───────────────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet' -Title '.NET Tools' -RequiresCommand 'dotnet' -Action {
    # Clean up broken .store entries that cause DirectoryNotFoundException
    $storePath = Join-Path $env:USERPROFILE '.dotnet\tools\.store'
    if (Test-Path $storePath) {
        Get-ChildItem -Path $storePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $toolDir = $_
            # A valid store entry has a versioned subfolder with a 'tools' directory
            $hasValidLayout = Get-ChildItem -Path $toolDir.FullName -Directory -ErrorAction SilentlyContinue | Where-Object {
                Test-Path (Join-Path $_.FullName "$($toolDir.Name)\$($_.Name)\tools") -ErrorAction SilentlyContinue
            }
            if (-not $hasValidLayout) {
                Write-Host "  Removing broken tool store entry: $($toolDir.Name)" -ForegroundColor Gray
                Remove-Item -Path $toolDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $toolLines = dotnet tool list -g 2>&1 | Select-Object -Skip 2 | Where-Object { $_ -match '\S' }
    if (-not $toolLines) { Set-UpdateSummary -Name 'dotnet' -Summary 'no global .NET tools installed'; return }

    $updatedCount = 0
    $errorCount = 0
    foreach ($line in $toolLines) {
        $parts = $line -split '\s+', 3
        if ($parts.Count -lt 2) { continue }
        $toolId = $parts[0].Trim()
        $currentVer = $parts[1].Trim()

        try {
            $meta = Invoke-RestMethod "https://api.nuget.org/v3-flatcontainer/$($toolId.ToLower())/index.json" -TimeoutSec 10 -ErrorAction Stop
            $latestVer = $meta.versions | Where-Object { $_ -notmatch '-' } | Select-Object -Last 1
            if (-not $latestVer) { $latestVer = $meta.versions | Select-Object -Last 1 }
        }
        catch { $latestVer = $null }

        if (-not $latestVer -or $latestVer -eq $currentVer) { continue }

        try {
            $out = (dotnet tool update -g $toolId 2>&1 | Out-String).Trim()
            if ($out -match 'Unhandled exception|DirectoryNotFoundException') {
                Write-Status "  $toolId update failed (broken store entry); uninstalling..." -Type Warning
                dotnet tool uninstall -g $toolId 2>&1 | Out-Null
                $out = (dotnet tool install -g $toolId 2>&1 | Out-String).Trim()
            }
            if ($out) { Write-Host $out -ForegroundColor Gray }
            Add-UpdateItem -Name 'dotnet' -Kind Updated -Item ('{0} {1} -> {2}' -f $toolId, $currentVer, $latestVer)
            $updatedCount++
        }
        catch {
            Write-Status "  $toolId failed: $_" -Type Warning
            $errorCount++
        }
    }

    if ($updatedCount -eq 0 -and $errorCount -eq 0) { Set-UpdateSummary -Name 'dotnet' -Summary 'all .NET tools are up to date' }
    if ($errorCount -gt 0) { Write-Status "$errorCount tool(s) had errors" -Type Warning }
}

# ── .NET workloads ───────────────────────────────────────────────────────────────
Invoke-Update -Name 'dotnet-workloads' -Title '.NET Workloads' -RequiresCommand 'dotnet' -Action {
    $workloads = (dotnet workload list 2>&1 | Out-String)
    if ($workloads -notmatch 'No workloads are installed') {
        $out = (dotnet workload update 2>&1 | Out-String).Trim()
        if ($out) {
            $lines = $out -split "`n"
            $filtered = $lines | Where-Object { $_ -notmatch 'Updated advertising manifest' }
            $verbose = $lines | Where-Object { $_ -match 'Updated advertising manifest' }
            if ($verbose) { Write-Verbose ($verbose -join "`n") }
            $filteredText = ($filtered -join "`n").Trim()
            if ($filteredText) { Write-Host $filteredText -ForegroundColor Gray }
        }
        Set-UpdateSummary -Name 'dotnet-workloads' -Summary 'workloads checked and updated if needed'
    }
    else {
        Set-UpdateSummary -Name 'dotnet-workloads' -Summary 'no .NET workloads installed'
    }
}

# ── VS Code extensions ───────────────────────────────────────────────────────────
Invoke-Update -Name 'vscode-extensions' -Title 'VS Code Extensions' -Disabled:$SkipVSCodeExtensions -RequiresAnyCommand @('code', 'code-insiders', 'code.cmd', 'code-insiders.cmd') -SlowOperation -Action {
    $codeCli = Get-VSCodeCliPath
    if (-not $codeCli) {
        Write-Status 'VS Code CLI shim not found (code.cmd/code-insiders.cmd); skipping extension update to avoid launching UI' -Type Warning
        Set-UpdateSummary -Name 'vscode-extensions' -Summary 'VS Code CLI shim not found'
        return
    }
    $out = (& $codeCli --update-extensions 2>&1 | Out-String).Trim()
    if ($out -and $out -notmatch 'No extension to update') { Write-Host $out -ForegroundColor Gray }
    if ($out -match 'No extension to update') {
        Set-UpdateSummary -Name 'vscode-extensions' -Summary 'no extension updates'
    }
    elseif ($out) {
        Set-UpdateSummary -Name 'vscode-extensions' -Summary 'extensions updated' -Changed
    }
}

# ── Flutter ───────────────────────────────────────────────────────────────────────
Invoke-Update -Name 'Flutter' -Disabled:$SkipFlutter -RequiresCommand 'flutter' -SlowOperation -Action {
    if ((Get-ToolInstallManager 'flutter') -eq 'scoop') { Set-UpdateSummary -Name 'Flutter' -Summary 'managed by scoop (already updated)' }
    else {
        $out = (flutter upgrade 2>&1 | Out-String).Trim()
        if ($out) { Write-Host $out -ForegroundColor Gray; Set-UpdateSummary -Name 'Flutter' -Summary 'updated' -Changed }
    }
}

# ── PowerShell modules / PSResources ─────────────────────────────────────────────
Invoke-Update -Name 'pwsh-resources' -Title 'PowerShell Modules / Resources' -Disabled:$SkipPowerShellModules -Action {
    $usedProvider = $false
    $resourceBefore = @{}
    $resourceAfter = @{}
    $moduleBefore = @{}
    $moduleAfter = @{}

    # PSResourceGet (modern, PS 7.4+) — parallelize per-resource updates
    if ((Test-Command 'Get-InstalledPSResource') -and (Test-Command 'Update-PSResource')) {
        $usedProvider = $true
        $resources = Get-InstalledPSResource -ErrorAction SilentlyContinue
        if ($resources) {
            foreach ($resource in @($resources)) {
                $resourceBefore[$resource.Name] = [string]$resource.Version
            }
            Write-Host "  Updating $(@($resources).Count) PSResource(s)..." -ForegroundColor Gray
            $_psrCmd = Get-Command Update-PSResource -ErrorAction SilentlyContinue
            $supportsAcceptLicense = if ($_psrCmd) { $_psrCmd.Parameters.ContainsKey('AcceptLicense') } else { $false }

            # Use parallel if PS7 and more than a handful of resources
            if ($PSVersionTable.PSVersion.Major -ge 7 -and @($resources).Count -gt 3) {
                $resources | ForEach-Object -Parallel {
                    $psrArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                    if ($using:supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { Write-Verbose "Update-PSResource failed for $($_.Name): $($_.Exception.Message)" }
                } -ThrottleLimit 4
            }
            else {
                foreach ($resource in $resources) {
                    $psrArgs = @{ Name = $resource.Name; ErrorAction = 'SilentlyContinue' }
                    if ($supportsAcceptLicense) { $psrArgs['AcceptLicense'] = $true }
                    try { Update-PSResource @psrArgs 2>&1 | Out-Null } catch { Write-Verbose "Update-PSResource failed for $($resource.Name): $($_.Exception.Message)" }
                }
            }
        }
        else {
            Set-UpdateSummary -Name 'pwsh-resources' -Summary 'no installed PSResources found'
        }

        foreach ($resource in @(Get-InstalledPSResource -ErrorAction SilentlyContinue)) {
            $resourceAfter[$resource.Name] = [string]$resource.Version
        }
    }

    # PowerShellGet (legacy fallback)
    if ((Test-Command 'Get-InstalledModule') -and (Test-Command 'Update-Module')) {
        $usedProvider = $true
        $modules = Get-InstalledModule -ErrorAction SilentlyContinue
        if ($modules) {
            foreach ($module in @($modules)) {
                $moduleBefore[$module.Name] = [string]$module.Version
            }
            Write-Host "  Updating $(@($modules).Count) PowerShellGet module(s)..." -ForegroundColor Gray
            $_modCmd = Get-Command Update-Module -ErrorAction SilentlyContinue
            $supportsAcceptLicense = if ($_modCmd) { $_modCmd.Parameters.ContainsKey('AcceptLicense') } else { $false }
            $modules | ForEach-Object {
                $updateArgs = @{ Name = $_.Name; ErrorAction = 'SilentlyContinue' }
                if ($supportsAcceptLicense) { $updateArgs['AcceptLicense'] = $true }
                Update-Module @updateArgs 2>&1 | Out-Null
            }
        }
        else {
            Set-UpdateSummary -Name 'pwsh-resources' -Summary 'no installed PowerShellGet modules found'
        }

        foreach ($module in @(Get-InstalledModule -ErrorAction SilentlyContinue)) {
            $moduleAfter[$module.Name] = [string]$module.Version
        }
    }

    if (-not $usedProvider) {
        Set-UpdateSummary -Name 'pwsh-resources' -Summary 'no PowerShell module update provider found'
    }
    else {
        foreach ($name in $resourceAfter.Keys) {
            if ($resourceBefore.ContainsKey($name) -and $resourceBefore[$name] -ne $resourceAfter[$name]) {
                Add-UpdateItem -Name 'pwsh-resources' -Kind Updated -Item ('PSResource {0} {1} -> {2}' -f $name, $resourceBefore[$name], $resourceAfter[$name])
            }
        }
        foreach ($name in $moduleAfter.Keys) {
            if ($moduleBefore.ContainsKey($name) -and $moduleBefore[$name] -ne $moduleAfter[$name]) {
                Add-UpdateItem -Name 'pwsh-resources' -Kind Updated -Item ('Module {0} {1} -> {2}' -f $name, $moduleBefore[$name], $moduleAfter[$name])
            }
        }
        if (-not (Get-UpdateSummaryText -Name 'pwsh-resources')) {
            Set-UpdateSummary -Name 'pwsh-resources' -Summary 'no PowerShell module/resource version changes'
        }
    }
}

# ════════════════════════════════════════════════════════════
#  CLEANUP
# ════════════════════════════════════════════════════════════
if ($SkipCleanup) {
    Write-Section 'System Cleanup'
    Write-Status 'Cleanup skipped (-SkipCleanup flag set)' -Type Info
    $updateResults.Skipped.Add([pscustomobject]@{ Name = 'cleanup'; Reason = 'flag' })
}
elseif ($DryRun) {
    Write-Host "  [DryRun] Would run: System Cleanup" -ForegroundColor DarkCyan
}
else {

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
    }
    catch {
        Write-Status "Temp cleanup partially failed (normal)" -Type Warning
    }

    # Windows system temp (older than 7 days, admin only)
    if ($isAdmin) {
        try {
            $winTempPath = 'C:\Windows\Temp'
            if (Test-Path $winTempPath) {
                $cutoff = (Get-Date).AddDays(-7)
                Get-ChildItem -Path $winTempPath -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "C:\Windows\Temp cleared (older than 7 days)" -Type Success
            }
        }
        catch {
            Write-Status "C:\Windows\Temp cleanup partially failed (normal)" -Type Warning
        }
    }

    # DNS cache
    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        Write-Status "DNS cache flushed" -Type Success
    }
    catch {
        Write-Status "DNS cache flush failed" -Type Warning
    }

    # Recycle Bin
    try {
        if ($PSCmdlet.ShouldProcess('Recycle Bin', 'Empty')) {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            Write-Status "Recycle Bin emptied" -Type Success
        }
    }
    catch {
        Write-Status "Recycle Bin cleanup failed" -Type Warning
    }

    # Crash dumps
    $crashDumpPath = Join-Path $env:LOCALAPPDATA 'CrashDumps'
    if (Test-Path $crashDumpPath) {
        try {
            $dumps = Get-ChildItem -Path $crashDumpPath -ErrorAction SilentlyContinue
            if ($dumps) {
                $dumps | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Crash dumps cleared ($($dumps.Count) files)" -Type Success
            }
        }
        catch {
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
        }
        catch {
            Write-Status "WER cleanup failed" -Type Warning
        }
    }

    # Admin-only deep cleanup (opt-in via -DeepClean)
    if ($isAdmin -and $DeepClean) {
        try {
            Write-Host "  Cleaning WinSxS component store..." -ForegroundColor Gray
            DISM /Online /Cleanup-Image /StartComponentCleanup 2>&1 | Out-Null
            Write-Status "DISM component store cleaned" -Type Success
        }
        catch {
            Write-Status "DISM cleanup failed" -Type Warning
        }

        try {
            Clear-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
            Write-Status "Delivery Optimization cache cleared" -Type Success
        }
        catch {
            Write-Status "Delivery Optimization cleanup failed" -Type Warning
        }

        # Prefetch cleanup (optional, rarely needed but useful on old HDDs)
        $prefetchPath = 'C:\Windows\Prefetch'
        if ((Test-Path $prefetchPath) -and -not $SkipDestructive) {
            try {
                Get-ChildItem -Path $prefetchPath -Filter '*.pf' -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
                Write-Status "Prefetch files cleared" -Type Success
            }
            catch {
                Write-Status "Prefetch cleanup failed (normal if Superfetch is disabled)" -Type Warning
            }
        }
    }

    $updateResults.Checked.Add('cleanup')

} # end -not $SkipCleanup

# ════════════════════════════════════════════════════════════
#  SUMMARY
# ════════════════════════════════════════════════════════════
$duration = (Get-Date) - $startTime
Write-Host "`n$('=' * 54)" -ForegroundColor Green
Write-Host (" UPDATE COMPLETE -- {0}" -f $duration.ToString('hh\:mm\:ss')) -ForegroundColor Green
Write-Host "$('=' * 54)" -ForegroundColor Green

$updatedNames = @($updateResults.Success | Select-Object -Unique)
$checkedNames = @($updateResults.Checked | Where-Object { $_ -notin $updatedNames -and $_ -notin $updateResults.Failed } | Select-Object -Unique)

if ($updatedNames.Count -gt 0) {
    Write-Host "`n[OK] Updated   ($($updatedNames.Count))" -ForegroundColor Green
    foreach ($name in $updatedNames) {
        $summary = Get-UpdateSummaryText -Name $name
        if (-not $summary) { $summary = 'updated' }
        Write-Detail ('{0}: {1}' -f $name, $summary) -Type Info
    }
}
if ($checkedNames.Count -gt 0) {
    Write-Host "[~] Checked   ($($checkedNames.Count))" -ForegroundColor DarkGray
    foreach ($name in $checkedNames) {
        $summary = Get-UpdateSummaryText -Name $name
        if (-not $summary) { $summary = 'checked; no changes needed' }
        Write-Detail ('{0}: {1}' -f $name, $summary) -Type Muted
    }
}
if ($updateResults.Failed.Count -gt 0) {
    Write-Host "[X] Failed    ($($updateResults.Failed.Count))" -ForegroundColor Red
    foreach ($name in $updateResults.Failed) {
        $summary = Get-UpdateSummaryText -Name $name
        if ($summary) {
            Write-Detail ('{0}: {1}' -f $name, $summary) -Type Error
        }
        else {
            Write-Detail $name -Type Error
        }
    }
}
if ($updateResults.Skipped.Count -gt 0) {
    $optInNames = @('Ollama Models')
    $skippedDisplay = $updateResults.Skipped | ForEach-Object {
        $item = $_
        $label = switch ($item.Reason) {
            'flag' {
                if ($item.Name -in $optInNames) {
                    'opt-in: -Update{0}' -f ($item.Name -replace '\s', '')
                }
                else {
                    'flag: -Skip{0}' -f ($item.Name -replace '\s', '')
                }
            }
            'not installed' { 'not installed' }
            'requires admin' { 'requires admin' }
            'fast mode' { 'fast mode' }
            default { $item.Reason }
        }
        '{0} ({1})' -f $item.Name, $label
    }
    Write-Host "[!] Skipped   ($($updateResults.Skipped.Count))" -ForegroundColor Yellow
    foreach ($item in $skippedDisplay) {
        Write-Detail $item -Type Warning
    }
}
if ($updatedNames.Count -eq 0 -and $checkedNames.Count -eq 0 -and $updateResults.Failed.Count -eq 0 -and $updateResults.Skipped.Count -eq 0) {
    Write-Status 'No updates were needed' -Type Info
}
# Per-section timing table
if ($script:sectionTimings.Count -gt 0) {
    $timingNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $updatedNames) { $null = $timingNames.Add($name) }
    foreach ($name in $checkedNames) { $null = $timingNames.Add($name) }
    foreach ($name in $updateResults.Failed) { $null = $timingNames.Add($name) }
    $timingsToShow = @(
        $script:sectionTimings.GetEnumerator() |
        Where-Object { $timingNames.Contains($_.Key) } |
        Sort-Object Value -Descending
    )
    if ($timingsToShow.Count -gt 0) {
        Write-Host "`n  Section timings:" -ForegroundColor DarkGray
        $timingsToShow | ForEach-Object {
            Write-Host ("    {0,-30} {1,6:F1}s" -f $_.Key, $_.Value) -ForegroundColor DarkGray
        }
    }
}

# Completion notification
try {
    if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) {
        Import-Module BurntToast -ErrorAction SilentlyContinue
        $successCount = $updatedNames.Count
        $failCount = $updateResults.Failed.Count
        $msg = if ($failCount -gt 0) { "$successCount updated, $failCount failed" } elseif ($successCount -gt 0) { "$successCount components updated" } else { 'No updates were needed' }
        New-BurntToastNotification -Text 'Update-Everything', $msg -ErrorAction SilentlyContinue
    }
    else {
        [System.Console]::Beep(880, 200)
        Start-Sleep -Milliseconds 50
        [System.Console]::Beep(1100, 300)
    }
}
catch { Write-Verbose "Completion beep skipped: $($_.Exception.Message)" }

Write-Host ""

# ── WhatChanged diff ─────────────────────────────────────────────────────────────
if ($WhatChanged -and (Test-Path $script:WingetSnapshotCurrent) -and (Test-Path $script:WingetSnapshotPrev)) {
    Write-Section 'What changed since last run'
    try {
        # Parse winget list output into a hashtable of PackageId -> Version
        function ConvertFrom-WingetList {
            param([string]$RawText)
            $map = @{}
            foreach ($line in ($RawText -split "`n")) {
                $trimmed = $line.Trim()
                if (-not $trimmed -or $trimmed -match '^[-=\s]+$' -or $trimmed -match '^Name\s') { continue }
                # winget list columns are variable-width; extract the last two non-empty tokens as Id and Version
                $tokens = $trimmed -split '\s{2,}' | Where-Object { $_ }
                if ($tokens.Count -ge 3) {
                    $id = $tokens[-2].Trim()
                    $ver = $tokens[-1].Trim()
                    if ($id -match '^\S+\.\S+$') { $map[$id] = $ver }
                }
            }
            return $map
        }

        $prevContent = Get-Content -Raw -Path $script:WingetSnapshotPrev
        $currContent = Get-Content -Raw -Path $script:WingetSnapshotCurrent
        $prevMap = ConvertFrom-WingetList $prevContent
        $currMap = ConvertFrom-WingetList $currContent

        $changes = @()
        foreach ($id in $currMap.Keys) {
            if (-not $prevMap.ContainsKey($id)) {
                $changes += "  + $id $($currMap[$id]) (new)"
            }
            elseif ($prevMap[$id] -ne $currMap[$id]) {
                $changes += "  ~ $id $($prevMap[$id]) -> $($currMap[$id])"
            }
        }
        foreach ($id in $prevMap.Keys) {
            if (-not $currMap.ContainsKey($id)) {
                $changes += "  - $id $($prevMap[$id]) (removed)"
            }
        }

        if ($changes.Count -eq 0) {
            Write-Host "  No package changes detected." -ForegroundColor Gray
        }
        else {
            Write-Host "  $($changes.Count) change(s) detected:" -ForegroundColor Cyan
            $changes | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        }
    }
    catch {
        Write-Status "WhatChanged diff failed: $_" -Type Warning
    }
}
elseif ($WhatChanged) {
    Write-Host "`n[!] No previous snapshot found for comparison. Run the script once without -WhatChanged first." -ForegroundColor Yellow
}

if ($transcriptStarted) {
    try { Stop-Transcript | Out-Null } catch { Write-Verbose "Stop-Transcript failed: $($_.Exception.Message)" }
}

if (-not $NoPause -and $AutoElevate) { Read-Host "`nPress Enter to close" }
