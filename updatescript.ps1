#Requires -Version 7.4
<#
.SYNOPSIS
  Windows update orchestrator: dev tools, winget apps, system cleanup.
.DESCRIPTION
  Parallel dev-tool updates, graceful app close before winget upgrades,
  guaranteed app restart via try/finally, robust winget output parsing,
  age-based temp cleanup, recycle bin + DNS flush.
.PARAMETER CleanupDays
  Files in temp older than N days are deleted. Default 7.
.PARAMETER GracefulCloseSec
  Seconds to wait for graceful window close before force-kill. Default 5.
.NOTES
  PS 7.4+. Run elevated for full system temp cleanup.
#>

[CmdletBinding()]
param(
    [ValidateRange(0,3650)]
    [int]$CleanupDays      = 7,
    [ValidateRange(0,60)]
    [int]$GracefulCloseSec = 5,
    [ValidateRange(1,16)]
    [int]$ThrottleLimit    = 4,
    [Alias('SkipPackageManagers')]
    [switch]$SkipDevTools,
    [switch]$SkipSystemAreas,
    [switch]$SkipWinget,
    [switch]$SkipCleanup,
    [string]$LogPath        = '',
    [ValidateRange(30,3600)]
    [int]$WingetTimeoutSec  = 300
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$start   = Get-Date
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
          ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($LogPath) {
    $logDir = Split-Path -Path $LogPath -Parent
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Start-Transcript -Path $LogPath -Append -ErrorAction SilentlyContinue
}

$WG_NO_UPDATE = -1978335189

# Winget output lines that are pure noise (spinner chars, progress bars)
$WG_NOISE_PATTERN = '^\s*[-\\|/]\s*$|[█▒]'

# === Apps to close before upgrade (winget Id, process names, restart path) ===
# Ollama: 'ollama app' (tray, parent) respawns 'ollama' (server) — kill both
$AppHooks = @(
    [pscustomobject]@{ Id='Google.Chrome';              Proc=@('chrome');              Path="$env:ProgramFiles\Google\Chrome\Application\chrome.exe" }
    [pscustomobject]@{ Id='Spotify.Spotify';            Proc=@('Spotify');             Path="$env:APPDATA\Spotify\Spotify.exe" }
    [pscustomobject]@{ Id='Microsoft.VisualStudioCode'; Proc=@('Code');                Path="$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe" }
    [pscustomobject]@{ Id='Ollama.Ollama';              Proc=@('ollama app','ollama'); Path="$env:LOCALAPPDATA\Programs\Ollama\ollama app.exe" }
    [pscustomobject]@{ Id='ZedIndustries.Zed';           Proc=@('Zed');                 Path="$env:LOCALAPPDATA\Programs\Zed\Zed.exe" }
    [pscustomobject]@{ Id='Anysphere.Cursor';            Proc=@('Cursor');              Path="$env:LOCALAPPDATA\Programs\cursor\Cursor.exe" }
    [pscustomobject]@{ Id='OpenWhisperSystems.Signal';   Proc=@('Signal');              Path="$env:LOCALAPPDATA\Programs\signal-desktop\Signal.exe" }
)

# Packages winget detects but can't auto-upgrade (install technology changed).
$ManualUpgradeIds = @('GoLang.Go', 'gerardog.gsudo')

# === Logging helpers ===
function Write-Section { param($t) Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Write-Info    { param($t) Write-Host "  $t"        -ForegroundColor Gray }
function Write-Ok      { param($t) Write-Host "  OK $t"     -ForegroundColor Green }
function Write-Skip    { param($t) Write-Host "  -- $t"     -ForegroundColor DarkGray }
function Write-Warn2   { param($t) Write-Host "  !! $t"     -ForegroundColor Yellow }

function Test-Command {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Format-Duration {
    param([Parameter(Mandatory)][timespan]$Duration)
    if ($Duration.TotalSeconds -lt 1) { return ('{0} ms' -f [int]$Duration.TotalMilliseconds) }
    return ([math]::Round($Duration.TotalSeconds, 1).ToString('0.0') + ' s')
}

# Shared implementation — used both in main scope and injected into parallel runspaces
$InvokeUpdateCommandDef = {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Bin,
        [Parameter(Mandatory)][string[]]$Args,
        [switch]$RequiresAdmin,
        [int[]]$SkipExitCodes = @(),
        [string]$SkipMessage  = 'not applicable'
    )
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($RequiresAdmin -and -not $script:IsAdmin) {
        return [pscustomobject]@{ Name=$Name; Status='skip'; Msg='needs admin'; Ms=$sw.ElapsedMilliseconds }
    }
    if (-not (Get-Command $Bin -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{ Name=$Name; Status='skip'; Msg='not installed'; Ms=$sw.ElapsedMilliseconds }
    }
    try {
        & $Bin @Args *>&1 | Out-Null
        $code = if ($null -ne $LASTEXITCODE) { $LASTEXITCODE } else { 0 }
        $sw.Stop()
        if ($code -eq 0)                        { return [pscustomobject]@{ Name=$Name; Status='ok';   Msg='';          Ms=$sw.ElapsedMilliseconds } }
        if ($SkipExitCodes -contains $code)     { return [pscustomobject]@{ Name=$Name; Status='skip'; Msg=$SkipMessage; Ms=$sw.ElapsedMilliseconds } }
        return [pscustomobject]@{ Name=$Name; Status='err'; Msg="exit $code"; Ms=$sw.ElapsedMilliseconds }
    } catch {
        $sw.Stop()
        return [pscustomobject]@{ Name=$Name; Status='err'; Msg=$_.Exception.Message; Ms=$sw.ElapsedMilliseconds }
    }
}

function Invoke-UpdateCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Bin,
        [Parameter(Mandatory)][string[]]$Args,
        [switch]$RequiresAdmin,
        [int[]]$SkipExitCodes = @(),
        [string]$SkipMessage  = 'not applicable'
    )
    & $script:InvokeUpdateCommandDef @PSBoundParameters
}

# === Graceful close: WM_CLOSE first, force-kill on timeout ===
function Stop-AppGracefully {
    param([string[]]$Names, [int]$TimeoutSec)
    $procs = Get-Process -Name $Names -ErrorAction SilentlyContinue
    if (-not $procs) { return $false }
    foreach ($p in $procs) {
        try { if ($p.MainWindowHandle -ne 0) { $null = $p.CloseMainWindow() } } catch {}
    }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline -and (Get-Process -Name $Names -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 250
    }
    $left = Get-Process -Name $Names -ErrorAction SilentlyContinue
    if ($left) { $left | Stop-Process -Force -ErrorAction SilentlyContinue }
    return $true
}

# === Parse winget upgrade output by header column positions ===
function Get-WingetUpgradeInfo {
    param([int]$TimeoutSec = 300)
    $wgJob = Start-Job -ScriptBlock { winget upgrade --include-unknown --disable-interactivity --accept-source-agreements 2>&1 | Out-String }
    if (-not ($wgJob | Wait-Job -Timeout $TimeoutSec)) {
        $wgJob | Stop-Job -PassThru | Remove-Job -Force
        Write-Warning "winget upgrade listing timed out after ${TimeoutSec}s — skipping winget section"
        return @{ Ids = @(); Manual = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
    }
    $raw = ($wgJob | Receive-Job -ErrorAction SilentlyContinue) -split "`r?`n"
    $wgJob | Remove-Job -Force
    $ids    = [System.Collections.Generic.List[string]]::new()
    $manual = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $idCol  = -1
    $verCol = -1
    $prevId = $null

    foreach ($line in $raw) {
        if ($line -match '^\s*Name\s+Id\s+Version') {
            $idCol  = $line.IndexOf('Id')
            $verCol = $line.IndexOf('Version')
            continue
        }
        if ($idCol -lt 0)                                     { continue }
        if ($line -match '^\s*-+\s*$')                        { continue }
        if ([string]::IsNullOrWhiteSpace($line))              { continue }
        if ($line -match 'upgrades?\s+available' -or
            $line -match 'require explicit targeting')        { $idCol = -1; continue }
        # Winget prints "install technology is different" on the line after the package
        if ($line -match 'install technology is different' -and $prevId) {
            $null = $manual.Add($prevId); continue
        }
        if ($line.Length -le $idCol) { continue }
        $endCol = [Math]::Min($verCol, $line.Length)
        $id = $line.Substring($idCol, $endCol - $idCol).Trim()
        if ($id -match '^[A-Za-z0-9._+-]+$') {
            $ids.Add($id)
            $prevId = $id
        }
    }
    return @{
        Ids    = ($ids | Select-Object -Unique)
        Manual = $manual
    }
}

# === Resolve app launch path: configured path, else PATH lookup ===
function Resolve-AppPath {
    param($Hook)
    if ($Hook.Path -and (Test-Path $Hook.Path)) { return $Hook.Path }
    $cmd = Get-Command $Hook.Proc[0] -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    return $null
}

# === Snapshot global npm packages (name@version) ===
function Get-NpmGlobalSnapshot {
    $out = npm list -g --depth=0 --json 2>$null | ConvertFrom-Json -ErrorAction SilentlyContinue
    if (-not $out) { return @{} }
    $map = @{}
    if (-not $out.dependencies) { return $map }
    foreach ($key in $out.dependencies.PSObject.Properties.Name) {
        $map[$key] = $out.dependencies.$key.version
    }
    return $map
}

# === Spin a Braille animation on the current line while a background job runs ===
# Clears the spinner line when done; caller writes the final status line.
function Wait-JobWithSpinner {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Job]$Job,
        [Parameter(Mandatory)][string]$Label,
        [int]$TimeoutSec = 300,
        [int]$FrameMs    = 80
    )
    $frames  = [char[]]@(0x280B, 0x2819, 0x2839, 0x2838, 0x283C, 0x2834, 0x2826, 0x2827, 0x2807, 0x280F)
    $fi      = 0
    $sw      = [Diagnostics.Stopwatch]::StartNew()
    $prefix  = "  "

    # Hide cursor while spinning to avoid flicker
    try { [Console]::CursorVisible = $false } catch {}

    $completed = $false
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        $frame = $frames[$fi % $frames.Length]
        $fi++
        $elapsed = Format-Duration $sw.Elapsed
        $line = "$prefix$frame $Label  $elapsed"
        # Overwrite the current line in place
        [Console]::Write("`r" + $line.PadRight([Math]::Max($line.Length, 60)))
        if ($Job | Wait-Job -Timeout 0) { $completed = $true; break }
        Start-Sleep -Milliseconds $FrameMs
    }

    try { [Console]::CursorVisible = $true } catch {}
    # Clear the spinner line so the caller's Write-Ok/Warn2 starts clean
    [Console]::Write("`r" + (' ' * ([Math]::Max(70, $Host.UI.RawUI.WindowSize.Width - 1))) + "`r")
    return $completed
}

# === 1. Parallel package manager / toolchain updates ===
if (-not $SkipDevTools) {
    Write-Section 'Package Managers + Toolchains (parallel)'

    # Remove leftover npm temp dirs (e.g. @scope/.pkg-XXXXXXXX) that cause EINVALIDPACKAGENAME
    $npmGlobalMods = "$env:APPDATA\npm\node_modules"
    $npmStaleDirsRemaining = @()
    if (Test-Path -LiteralPath $npmGlobalMods) {
        $npmStaleDirs = @(Get-ChildItem -LiteralPath $npmGlobalMods -Directory -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object -FilterScript { $_.Parent.Name -like '@*' -and $_.Name -match '^\.[A-Za-z0-9._-]+$' })
        $npmStaleDirs | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        $npmStaleDirsRemaining = @(Get-ChildItem -LiteralPath $npmGlobalMods -Directory -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object -FilterScript { $_.Parent.Name -like '@*' -and $_.Name -match '^\.[A-Za-z0-9._-]+$' })
    }

    # Snapshot npm globals before update so we can diff after
    $npmBefore = if (Test-Command npm) { Get-NpmGlobalSnapshot } else { @{} }
    $preResults = @()
    if (Test-Command npm.cmd) {
        $preResults += Invoke-UpdateCommand -Name 'npm itself' -Bin 'npm.cmd' -Args @('install','-g','npm@latest')
    }
    if ($npmStaleDirsRemaining.Count -gt 0) {
        $preResults += [pscustomobject]@{ Name='npm globals'; Status='skip'; Msg='stale npm temp dir locked'; Ms=0 }
    }

    $tools = [System.Collections.Generic.List[hashtable]]::new()
    if ($npmStaleDirsRemaining.Count -eq 0) {
        $tools.Add(@{ Name='npm globals'; Bin='npm.cmd'; Args=@('update','-g') })
    }
    $tools.AddRange([hashtable[]]@(
        # Windows package managers
        @{ Name='scoop';                    Bin='scoop';     Args=@('update','*') }
        @{ Name='scoop cleanup';            Bin='scoop';     Args=@('cleanup','*') }
        @{ Name='chocolatey';               Bin='choco';     Args=@('upgrade','all','-y','--no-progress'); Admin=$true }
        @{ Name='homebrew';                 Bin='brew';      Args=@('upgrade') }
        @{ Name='MSYS2 pacman';             Bin='pacman';    Args=@('-Syu','--noconfirm') }
        @{ Name='chezmoi';                  Bin='chezmoi';   Args=@('upgrade') }
        @{ Name='mise';                     Bin='mise';      Args=@('upgrade') }
        @{ Name='mise plugins';             Bin='mise';      Args=@('plugins','update') }
        @{ Name='asdf plugins';             Bin='asdf';      Args=@('plugin','update','--all') }
        @{ Name='fnm node';                 Bin='fnm';       Args=@('install','--latest') }

        # JavaScript / TypeScript
        @{ Name='pnpm globals';             Bin='pnpm';      Args=@('update','-g') }
        @{ Name='yarn globals';             Bin='yarn';      Args=@('global','upgrade') }
        @{ Name='bun';                      Bin='bun';       Args=@('upgrade') }
        @{ Name='deno';                     Bin='deno';      Args=@('upgrade') }
        @{ Name='corepack';                 Bin='corepack';  Args=@('enable') }
        @{ Name='volta';                    Bin='volta';     Args=@('install','node@latest','npm@latest','yarn@latest','pnpm@latest') }

        # Python
        @{ Name='pip base';                 Bin='python';    Args=@('-m','pip','install','--upgrade','pip','setuptools','wheel') }
        @{ Name='pipx apps';                Bin='pipx';      Args=@('upgrade-all') }
        @{ Name='uv';                       Bin='uv';        Args=@('self','update') }
        @{ Name='uv tools';                 Bin='uv';        Args=@('tool','upgrade','--all') }
        @{ Name='poetry';                   Bin='poetry';    Args=@('self','update') }
        @{ Name='pdm';                      Bin='pdm';       Args=@('self','update') }
        @{ Name='hatch';                    Bin='hatch';     Args=@('self','update') }
        @{ Name='rye';                      Bin='rye';       Args=@('self','update') }
        @{ Name='conda';                    Bin='conda';     Args=@('update','-n','base','-c','defaults','conda','-y') }
        @{ Name='mamba';                    Bin='mamba';     Args=@('update','--all','-y') }
        @{ Name='pixi';                     Bin='pixi';      Args=@('self-update') }

        # Rust / Go / .NET / JVM / Ruby / PHP
        @{ Name='rustup self';              Bin='rustup';    Args=@('self','update') }
        @{ Name='rustup';                   Bin='rustup';    Args=@('update') }
        @{ Name='cargo installs';           Bin='cargo';     Args=@('install-update','-a') }
        @{ Name='go tools';                 Bin='go';        Args=@('install','golang.org/x/tools/gopls@latest') }
        @{ Name='dotnet tools';             Bin='dotnet';    Args=@('tool','update','--global','--all') }
        @{ Name='dotnet workloads';         Bin='dotnet';    Args=@('workload','update'); Admin=$true }
        @{ Name='sdkman';                   Bin='pwsh';      Args=@('-NoProfile','-Command','if ($env:SDKMAN_DIR -and (Test-Path "$env:SDKMAN_DIR\bin\sdkman-init.ps1")) { . "$env:SDKMAN_DIR\bin\sdkman-init.ps1"; sdk selfupdate force; sdk update } else { exit 127 }'); SkipExitCodes=@(127); SkipMessage='not installed' }
        @{ Name='gem';                      Bin='gem';       Args=@('update','--system') }
        @{ Name='ruby gems';                Bin='gem';       Args=@('update') }
        @{ Name='composer';                 Bin='composer';  Args=@('self-update') }
        @{ Name='composer globals';         Bin='composer';  Args=@('global','update') }
        @{ Name='juliaup';                  Bin='juliaup';   Args=@('update') }
        @{ Name='flutter';                  Bin='flutter';   Args=@('upgrade') }
        @{ Name='ghcup';                    Bin='ghcup';     Args=@('upgrade') }
        @{ Name='stack';                    Bin='stack';     Args=@('upgrade','--binary-only') }
        @{ Name='cabal index';              Bin='cabal';     Args=@('update') }
        @{ Name='opam';                     Bin='opam';      Args=@('upgrade','-y') }
        @{ Name='nimble';                   Bin='nimble';    Args=@('refresh') }
        @{ Name='zef';                      Bin='zef';       Args=@('upgrade') }

        # Shell / CLI ecosystems
        @{ Name='PowerShell resources';     Bin='pwsh';      Args=@('-NoProfile','-Command','$resources = if (Get-Command Get-InstalledPSResource -ErrorAction SilentlyContinue) { @(Get-InstalledPSResource -ErrorAction SilentlyContinue) } else { @() }; if ($resources.Count -gt 0 -and (Get-Command Update-PSResource -ErrorAction SilentlyContinue)) { Update-PSResource -Name $resources.Name -TrustRepository -AcceptLicense -ErrorAction SilentlyContinue; exit 0 }; $modules = if (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue) { @(Get-InstalledModule -ErrorAction SilentlyContinue) } else { @() }; if ($modules.Count -gt 0 -and (Get-Command Update-Module -ErrorAction SilentlyContinue)) { Update-Module -Name $modules.Name -AcceptLicense -Force -ErrorAction SilentlyContinue; exit 0 }; exit 0') }
        @{ Name='oh-my-posh';               Bin='oh-my-posh';Args=@('upgrade') }
        @{ Name='starship';                 Bin='starship';  Args=@('self','update'); SkipExitCodes=@(2); SkipMessage='self-update unsupported; use winget/choco installer' }
        @{ Name='tldr';                     Bin='tldr';      Args=@('--update') }
        @{ Name='gh extensions';            Bin='gh';        Args=@('extension','upgrade','--all') }
        @{ Name='git lfs';                  Bin='git';       Args=@('lfs','install','--skip-repo') }
        @{ Name='vcpkg';                    Bin='vcpkg';     Args=@('upgrade','--no-dry-run') }

        # Editors / IDE plugin surfaces
        @{ Name='VS Code extensions';       Bin='code';      Args=@('--update-extensions') }
        @{ Name='Cursor extensions';        Bin='cursor';    Args=@('--update-extensions') }
        @{ Name='Codium extensions';        Bin='codium';    Args=@('--update-extensions') }
        @{ Name='JetBrains Toolbox';        Bin='jetbrains'; Args=@('update') }

        # Containers / infra CLIs
        @{ Name='kubectl krew';             Bin='kubectl';   Args=@('krew','upgrade') }
        @{ Name='helm repos';               Bin='helm';      Args=@('repo','update') }
        @{ Name='terraform';                Bin='tfupdate';  Args=@('terraform','--version','latest') }
    ))

    $isAdminForParallel   = $IsAdmin
    $invokeDefStr         = $InvokeUpdateCommandDef.ToString()
    $results = @($preResults) + @($tools | ForEach-Object -Parallel {
        $script:IsAdmin = $using:isAdminForParallel
        ${function:Invoke-UpdateCommand} = [scriptblock]::Create($using:invokeDefStr)
        Invoke-UpdateCommand -Name $_.Name -Bin $_.Bin -Args $_.Args `
            -RequiresAdmin:([bool]$_.Admin) `
            -SkipExitCodes @($_.SkipExitCodes) `
            -SkipMessage ($_.SkipMessage ?? 'not applicable')
    } -ThrottleLimit $ThrottleLimit)

    foreach ($r in $results) {
        $time = Format-Duration ([timespan]::FromMilliseconds($r.Ms))
        switch ($r.Status) {
            'ok'   { Write-Ok    "$($r.Name) ($time)" }
            'skip' { Write-Skip  "$($r.Name): $($r.Msg)" }
            'err'  { Write-Warn2 "$($r.Name): $($r.Msg) ($time)" }
        }
    }

    # Diff npm globals and report what changed
    if ($npmBefore.Count -gt 0) {
        $npmAfter   = Get-NpmGlobalSnapshot
        $npmChanged = @()
        foreach ($pkg in $npmAfter.Keys) {
            if (-not $npmBefore.ContainsKey($pkg)) {
                $npmChanged += "  + $pkg@$($npmAfter[$pkg]) (new)"
            } elseif ($npmBefore[$pkg] -ne $npmAfter[$pkg]) {
                $npmChanged += "  $pkg  $($npmBefore[$pkg]) -> $($npmAfter[$pkg])"
            }
        }
        if ($npmChanged.Count -gt 0) {
            Write-Info 'npm changes:'
            $npmChanged | ForEach-Object -Process { Write-Info $_ }
        }
    }
}

# === 1b. System update surfaces ===
if (-not $SkipSystemAreas) {
    Write-Section 'System Areas'
    $systemJobs = @(
        @{ Name='Windows Update scan'; Bin='UsoClient.exe'; Args=@('StartScan') }
        @{ Name='Defender signatures'; Bin='powershell';    Args=@('-NoProfile','-Command','Update-MpSignature') }
        @{ Name='WSL kernel';          Bin='wsl.exe';       Args=@('--update') }
        @{ Name='Store app scan';      Bin='winget';        Args=@('source','update') }
    )
    foreach ($job in $systemJobs) {
        $r    = Invoke-UpdateCommand -Name $job.Name -Bin $job.Bin -Args $job.Args
        $time = Format-Duration ([timespan]::FromMilliseconds($r.Ms))
        switch ($r.Status) {
            'ok'   { Write-Ok    "$($r.Name) ($time)" }
            'skip' { Write-Skip  "$($r.Name): $($r.Msg)" }
            'err'  { Write-Warn2 "$($r.Name): $($r.Msg) ($time)" }
        }
    }
}

# === 2. Winget upgrades with hook-based close/restart ===
if (-not $SkipWinget) {
    Write-Section 'Winget'
    if (-not (Test-Command winget)) {
        Write-Warn2 'winget not found, skipping'
    } else {
        Write-Info 'querying upgrades...'
        $wgInfo     = Get-WingetUpgradeInfo -TimeoutSec $WingetTimeoutSec
        $upgradeIds = @($wgInfo.Ids)
        $manualIds  = $wgInfo.Manual
        Write-Info "$($upgradeIds.Count) upgrade(s) found"

        # Deduplicate manual-reinstall IDs from both static list and winget-detected set
        $allManual = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in $ManualUpgradeIds) { $null = $allManual.Add($id) }
        foreach ($id in $manualIds)        { $null = $allManual.Add($id) }

        # Auto-reinstall technology-mismatch packages (uninstall then install)
        $noiseRx = $WG_NOISE_PATTERN
        foreach ($id in @($upgradeIds | Where-Object { $allManual.Contains($_) })) {
            Write-Info "$id: technology mismatch — reinstalling..."
            $phases = @(
                @{ Label='uninstalling'; Args=@('uninstall','--id',$id,'--silent','--disable-interactivity','--accept-source-agreements') }
                @{ Label='installing';   Args=@('install',  '--id',$id,'--silent','--disable-interactivity','--accept-source-agreements','--accept-package-agreements') }
            )
            $abortReinstall = $false
            foreach ($phase in $phases) {
                if ($abortReinstall) { break }
                $rArgs = $phase.Args
                $rJob  = Start-Job -ScriptBlock { param($a) & winget @a 2>&1; $LASTEXITCODE } -ArgumentList @(,$rArgs)
                $rsw   = [Diagnostics.Stopwatch]::StartNew()
                $rdone = Wait-JobWithSpinner -Job $rJob -Label "$($phase.Label) $id" -TimeoutSec $WingetTimeoutSec
                $rsw.Stop()
                if (-not $rdone) {
                    $rJob | Stop-Job -PassThru | Remove-Job -Force
                    Get-Process -Name 'winget' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                    Write-Warn2 "$id $($phase.Label) timed out — skipping"
                    $abortReinstall = $true; continue
                }
                $rOut   = @($rJob | Receive-Job -ErrorAction SilentlyContinue)
                $rJob   | Remove-Job -Force
                $rCode  = if ($rOut.Count -gt 0 -and $rOut[-1] -is [int]) { [int]$rOut[-1] } else { 0 }
                $rLines = if ($rOut.Count -gt 1) { $rOut[0..($rOut.Count - 2)] } else { @() }
                $rLines | Where-Object { $_ -and $_ -notmatch $noiseRx } | ForEach-Object { Write-Info $_ }
                if ($phase.Label -eq 'uninstalling' -and $rCode -ne 0) {
                    Write-Warn2 "$id uninstall failed (exit $rCode) — skipping reinstall"
                    $abortReinstall = $true; continue
                }
                if ($phase.Label -eq 'installing') {
                    $t = Format-Duration $rsw.Elapsed
                    switch ($rCode) {
                        0       { Write-Ok   "$id reinstalled ($t)" }
                        default { Write-Warn2 "$id reinstall exit $rCode ($t)" }
                    }
                }
            }
        }

        $upgradeIds = $upgradeIds | Where-Object { -not $allManual.Contains($_) }

        $closed = [System.Collections.Generic.List[object]]::new()
        try {
            foreach ($hook in $AppHooks) {
                $hit = $upgradeIds | Where-Object { $_ -eq $hook.Id -or $_ -like "$($hook.Id).*" }
                if (-not $hit) { continue }
                if (Stop-AppGracefully -Names $hook.Proc -TimeoutSec $GracefulCloseSec) {
                    Write-Info "closed $($hook.Id)"
                    $closed.Add($hook)
                }
            }

            foreach ($id in $upgradeIds) {
                $wgArgs = @('upgrade','--id',$id,'--silent','--include-unknown',
                            '--disable-interactivity','--accept-source-agreements',
                            '--accept-package-agreements')
                $wgJob  = Start-Job -ScriptBlock { param($a) & winget @a 2>&1; $LASTEXITCODE } -ArgumentList @(,$wgArgs)
                $sw     = [Diagnostics.Stopwatch]::StartNew()

                $completed = Wait-JobWithSpinner -Job $wgJob -Label $id -TimeoutSec $WingetTimeoutSec
                $sw.Stop()
                $time = Format-Duration $sw.Elapsed

                if (-not $completed) {
                    $wgJob | Stop-Job -PassThru | Remove-Job -Force
                    Write-Warn2 "$id (timed out after ${WingetTimeoutSec}s)"
                    continue
                }
                $wgOut = @($wgJob | Receive-Job -ErrorAction SilentlyContinue)
                $wgJob | Remove-Job -Force

                $code  = if ($wgOut.Count -gt 0 -and $wgOut[-1] -is [int]) { [int]$wgOut[-1] } else { 0 }
                $lines = if ($wgOut.Count -gt 1) { $wgOut[0..($wgOut.Count - 2)] } else { @() }

                # Print only meaningful output lines (skip spinner/progress-bar noise)
                $lines | Where-Object { $_ -and $_ -notmatch $WG_NOISE_PATTERN } |
                    ForEach-Object { Write-Info $_ }

                switch ($code) {
                    0             { Write-Ok   "$id ($time)" }
                    $WG_NO_UPDATE { Write-Skip "$id (no applicable update, $time)" }
                    default       { Write-Warn2 "$id (exit $code, $time)" }
                }
            }
        } finally {
            foreach ($hook in $closed) {
                $path = Resolve-AppPath -Hook $hook
                if ($path) {
                    Start-Process -FilePath $path -ErrorAction SilentlyContinue
                    Write-Info "restarted $($hook.Id)"
                } else {
                    Write-Warn2 "no launch path for $($hook.Id)"
                }
            }
        }
    }
}

# === 3. Cleanup: temp + recycle bin + DNS ===
if (-not $SkipCleanup) {
    Write-Section 'Cleanup'
    $paths   = @($env:TEMP, 'C:\Windows\Temp')
    $cutoff  = (Get-Date).AddDays(-$CleanupDays)
    $freed   = 0L
    $removed = 0

    foreach ($p in $paths) {
        if (-not (Test-Path $p)) { continue }
        if ($p -eq 'C:\Windows\Temp' -and -not $IsAdmin) {
            Write-Warn2 "skip $p (needs admin)"
            continue
        }
        Get-ChildItem -LiteralPath $p -Recurse -Force -File -Attributes !ReparsePoint -ErrorAction SilentlyContinue |
            Where-Object -FilterScript { $_.LastWriteTime -lt $cutoff } |
            ForEach-Object -Process {
                $f = $_
                try {
                    $sz = $f.Length
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                    $freed += $sz
                    $removed++
                } catch {}
            }
        Write-Ok $p
    }
    Write-Info ('removed {0} file(s), freed {1} MB from temp' -f $removed, ([math]::Round($freed / 1MB, 1).ToString('0.0')))

    try { Clear-RecycleBin -Force -Confirm:$false -ErrorAction Stop; Write-Ok 'recycle bin' }
    catch { Write-Warn2 "recycle bin: $($_.Exception.Message)" }

    try { Clear-DnsClientCache -ErrorAction Stop; Write-Ok 'dns cache' }
    catch { Write-Warn2 "dns: $($_.Exception.Message)" }
}

# === Summary ===
$elapsed = (Get-Date) - $start
Write-Host ("`nDone in {0:mm\:ss}." -f $elapsed) -ForegroundColor Green
if ($LogPath) { Stop-Transcript -ErrorAction SilentlyContinue }
