<#
.SYNOPSIS
    Update Everything for Windows.
.DESCRIPTION
    Runs system, package manager, and developer-tool updates from one PowerShell entrypoint.
    The script discovers available tools, applies skip/only filters, executes tasks with
    throttled parallel jobs, records logs, and writes a machine-readable run summary.
.VERSION
    6.1.0
.NOTES
    PowerShell 7+ is required. Run as Administrator for all system-level tasks.
    Uses Start-ThreadJob (Microsoft.PowerShell.ThreadJob) for fast parallel execution.
.EXAMPLE
    .\updatescript.ps1
    .\updatescript.ps1 -DryRun
    .\updatescript.ps1 -SelfTest
    .\updatescript.ps1 -ListTasks
    .\updatescript.ps1 -FastMode
    .\updatescript.ps1 -UltraFast
    .\updatescript.ps1 -Only winget,npm,pipx
    .\updatescript.ps1 -Skip windows-update,cleanup
    .\updatescript.ps1 -AutoElevate
    .\updatescript.ps1 -Schedule -ScheduleTime "03:00"
#>

#Requires -Version 7.0

[CmdletBinding(SupportsShouldProcess = $true)]
param(
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
    [switch]$SkipNode,
    [switch]$SkipRust,
    [switch]$SkipGo,
    [switch]$SkipFlutter,
    [switch]$SkipGitLFS,
    [switch]$DeepClean,
    [switch]$UpdateOllamaModels,
    [switch]$WhatChanged,
    [switch]$DryRun,
    [switch]$ListTasks,
    [switch]$SelfTest,
    [switch]$NoParallel,
    [switch]$Quiet,
    [switch]$Schedule,

    [ValidateScript({ $_ -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$' })]
    [string]$ScheduleTime = '03:00',

    [ValidateRange(30, 7200)]
    [int]$WingetTimeoutSec = 600,

    [ValidateRange(60, 14400)]
    [int]$TaskTimeoutSec = 1800,

    [ValidateRange(0, 16)]
    [int]$ParallelThrottle = 0,

    [ValidateRange(0, 5)]
    [int]$RetryCount = 0,

    [string]$LogPath,
    [string]$JsonSummaryPath,
    [string]$StateDir,
    [string[]]$Only = @(),
    [string[]]$Skip = @()
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Continue'

try {
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
}
catch {
    Write-Verbose "Console encoding setup skipped: $($_.Exception.Message)"
}

$script:StartTime = Get-Date
$script:RunId = $script:StartTime.ToString('yyyyMMdd-HHmmss-fff')
$script:CommandCache = @{}
if (-not $StateDir) {
    $StateDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Update-Everything'
}
$script:StateDir = $StateDir
$script:LogDir = Join-Path $script:StateDir 'logs'
$script:DefaultJsonSummaryPath = Join-Path $script:StateDir 'last-run.json'
$script:PreviousJsonSummaryPath = Join-Path $script:StateDir 'previous-run.json'
$script:IsSimulation = $DryRun -or $WhatIfPreference

if (-not $LogPath) {
    $LogPath = Join-Path $script:LogDir ("update-everything-{0}.log" -f $script:RunId)
}
if (-not $JsonSummaryPath) {
    $JsonSummaryPath = $script:DefaultJsonSummaryPath
}

$script:Config = [ordered]@{
    FastModeSkip       = @(
        'chocolatey', 'wsl-distros', 'npm', 'pnpm', 'yarn', 'bun', 'deno',
        'rustup', 'cargo', 'go', 'pip', 'pipx', 'uv', 'uv-tools',
        'poetry', 'composer', 'ruby-gems', 'flutter', 'dotnet-tools',
        'dotnet-workloads', 'vscode-extensions', 'powershell-modules',
        'ollama-models'
    )
    UltraFastSkip      = @('windows-update', 'store-apps', 'wsl', 'wsl-distros', 'defender', 'cleanup')
    SkipManagers       = @()
    WingetSkipPackages = @()
    PipSkipPackages    = @()
    NpmSkipPackages    = @()
    LogRetentionDays   = 14
    TempCleanupDays    = 7
}

function ConvertTo-StringArray {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) { return @($Value) }
    return @($Value | ForEach-Object { [string]$_ })
}

function ConvertTo-FilterList {
    param([AllowNull()]$Value)

    return @(ConvertTo-StringArray $Value |
        ForEach-Object { $_ -split ',' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ })
}

function Import-UpdateConfig {
    $configPath = Join-Path $PSScriptRoot 'update-config.json'
    if (-not (Test-Path -LiteralPath $configPath)) { return }

    try {
        $configJson = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in $configJson.PSObject.Properties) {
            if (-not $script:Config.Contains($property.Name)) {
                $script:Config[$property.Name] = $property.Value
                continue
            }

            if ($script:Config[$property.Name] -is [array]) {
                $script:Config[$property.Name] = @(ConvertTo-FilterList $property.Value)
            }
            else {
                $script:Config[$property.Name] = $property.Value
            }
        }
        Write-Verbose "Loaded config from $configPath"
    }
    catch {
        Write-Warning "Failed to load update-config.json: $($_.Exception.Message)"
    }
}

function Initialize-RunStorage {
    foreach ($path in @($script:StateDir, $script:LogDir)) {
        if (-not (Test-Path -LiteralPath $path)) {
            New-Item -Path $path -ItemType Directory -Force -WhatIf:$false | Out-Null
        }
    }

    try {
        if (Test-Path -LiteralPath $script:DefaultJsonSummaryPath) {
            Copy-Item -LiteralPath $script:DefaultJsonSummaryPath -Destination $script:PreviousJsonSummaryPath -Force -WhatIf:$false -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Verbose "Could not rotate previous run summary: $($_.Exception.Message)"
    }

    $retentionDays = [int]$script:Config.LogRetentionDays
    if ($retentionDays -gt 0) {
        try {
            $cutoff = (Get-Date).AddDays(-$retentionDays)
            Get-ChildItem -LiteralPath $script:LogDir -Filter 'update-everything-*.log' -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff } |
            Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue
        }
        catch {
            Write-Verbose "Log retention cleanup skipped: $($_.Exception.Message)"
        }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Muted')]
        [string]$Level = 'Info'
    )

    $line = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level.ToUpperInvariant(), $Message
    try { Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8 -WhatIf:$false } catch {}
}

function Write-Status {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Muted')]
        [string]$Level = 'Info'
    )

    Write-Log -Message $Message -Level $Level
    if ($Quiet -and $Level -notin @('Warning', 'Error')) { return }

    $color = switch ($Level) {
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Muted' { 'DarkGray' }
        default { 'Cyan' }
    }

    Write-Host $Message -ForegroundColor $color
}

function Test-IsAdmin {
    if (-not $IsWindows) { return $false }
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Test-Command {
    param([Parameter(Mandatory)][string]$Name)

    if ($script:CommandCache.ContainsKey($Name)) {
        return $script:CommandCache[$Name]
    }

    $found = [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
    $script:CommandCache[$Name] = $found
    return $found
}

function Get-CommandPath {
    param([Parameter(Mandatory)][string]$Name)

    $command = Get-Command -Name $Name -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    return $null
}

function ConvertTo-TaskId {
    param([Parameter(Mandatory)][string]$Value)

    return (($Value.Trim().ToLowerInvariant()) -replace '[^a-z0-9]+', '-').Trim('-')
}

function Test-NameMatch {
    param(
        [Parameter(Mandatory)]$Task,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        $rawPattern = $pattern.Trim()
        $needle = ConvertTo-TaskId $rawPattern
        $taskName = ConvertTo-TaskId $Task.Name

        if ($Task.Id -eq $needle -or $taskName -eq $needle) {
            return $true
        }
        if ($Task.Category -and ((ConvertTo-TaskId $Task.Category) -eq $needle)) {
            return $true
        }
        if ($Task.Tags -contains $needle) {
            return $true
        }
        if ($rawPattern -match '[*?]' -and $Task.Name -like $rawPattern) {
            return $true
        }
    }

    return $false
}

function New-UpdateTask {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][scriptblock]$Script,
        [string[]]$RequiresCommand = @(),
        [string[]]$Tags = @(),
        [switch]$RequiresAdmin,
        [switch]$Disabled,
        [string]$DisabledReason,
        [int]$TimeoutSec = $TaskTimeoutSec
    )

    [pscustomobject]@{
        Id              = ConvertTo-TaskId $Name
        Name            = $Name
        Category        = $Category
        Script          = $Script
        RequiresCommand = @($RequiresCommand)
        Tags            = @($Tags | ForEach-Object { ConvertTo-TaskId $_ })
        RequiresAdmin   = [bool]$RequiresAdmin
        Disabled        = [bool]$Disabled
        DisabledReason  = $DisabledReason
        TimeoutSec      = $TimeoutSec
    }
}

function Join-QuotedList {
    param([string[]]$Values)

    return (@($Values) | ForEach-Object { "'$($_.Replace("'", "''"))'" }) -join ', '
}

function Get-UpdateTasks {
    $wingetSkip = ConvertTo-StringArray $script:Config.WingetSkipPackages
    $pipSkip = ConvertTo-StringArray $script:Config.PipSkipPackages
    $npmSkip = ConvertTo-FilterList $script:Config.NpmSkipPackages
    $tempCleanupDays = [int]$script:Config.TempCleanupDays

    $tasks = [System.Collections.Generic.List[object]]::new()

    $wingetScript = [scriptblock]::Create(@"
param([string[]]`$SkipPackages)
`$allArgs = @(
    'upgrade', '--all', '--include-unknown', '--silent', '--disable-interactivity',
    '--accept-package-agreements', '--accept-source-agreements'
)

if (`$SkipPackages.Count -eq 0) {
    & winget @allArgs
    return
}

`$skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach (`$package in `$SkipPackages) {
    if (-not [string]::IsNullOrWhiteSpace(`$package)) { [void]`$skipSet.Add(`$package.Trim()) }
}

Write-Output "Winget skip list detected: `$(`$skipSet -join ', ')"
`$listOutput = & winget upgrade --include-unknown --accept-source-agreements 2>&1
`$wingetListExit = `$LASTEXITCODE
if (`$wingetListExit -ne 0) {
    throw "winget upgrade list failed with exit code `$wingetListExit"
}

`$ids = [System.Collections.Generic.List[string]]::new()
`$inTable = `$false
`$idColumnStart = `$null
`$idColumnEnd = `$null
foreach (`$line in `$listOutput) {
    `$text = [string]`$line
    if (`$text -match '^\s*-{3,}') {
        `$inTable = `$true
        `$columnMatches = [regex]::Matches(`$text, '-{3,}')
        if (`$columnMatches.Count -ge 2) {
            `$idColumnStart = `$columnMatches[1].Index
            if (`$columnMatches.Count -ge 3) {
                `$idColumnEnd = `$columnMatches[2].Index
            }
        }
        continue
    }
    if (-not `$inTable -or [string]::IsNullOrWhiteSpace(`$text)) { continue }
    if (`$null -eq `$idColumnStart -or `$text.Length -le `$idColumnStart) { continue }
    if (`$null -ne `$idColumnEnd -and `$idColumnEnd -gt `$idColumnStart) {
        `$width = [Math]::Min(`$idColumnEnd - `$idColumnStart, `$text.Length - `$idColumnStart)
        `$id = `$text.Substring(`$idColumnStart, `$width).Trim()
    }
    else {
        `$id = `$text.Substring(`$idColumnStart).Trim()
    }
    if (`$id -and `$id -notmatch '^(Id|Version|-)') {
        [void]`$ids.Add(`$id)
    }
}

if (`$ids.Count -eq 0) {
    Write-Output 'No winget upgrades found.'
    return
}

`$failed = [System.Collections.Generic.List[string]]::new()
foreach (`$id in `$ids) {
    if (`$skipSet.Contains(`$id)) {
        Write-Output "Skipping winget package: `$id"
        continue
    }

    Write-Output "Upgrading winget package: `$id"
    & winget upgrade --id `$id --exact --include-unknown --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
    if (`$LASTEXITCODE -ne 0) { [void]`$failed.Add(`$id) }
}

if (`$failed.Count -gt 0) {
    throw "winget failed packages: `$(`$failed -join ', ')"
}
"@)

    $tasks.Add((New-UpdateTask -Name 'winget' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec $WingetTimeoutSec -Script $wingetScript -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'scoop' -Category 'package-manager' -RequiresCommand 'scoop' -Script {
                scoop update
                if ($LASTEXITCODE -ne 0) { throw "scoop update failed with exit code $LASTEXITCODE" }
                scoop update *
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'chocolatey' -Category 'package-manager' -RequiresCommand 'choco' -RequiresAdmin -Script {
                choco upgrade all -y --no-progress --limit-output
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'store-apps' -Category 'system' -RequiresCommand 'winget' -Disabled:$SkipStoreApps -DisabledReason 'disabled by -SkipStoreApps' -TimeoutSec $WingetTimeoutSec -Script {
                Write-Output 'store-apps command: winget upgrade --all --source msstore --include-unknown --silent --disable-interactivity --accept-package-agreements --accept-source-agreements'
                winget upgrade --all --source msstore --include-unknown --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
            } -Tags @('windows', 'store'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'windows-update' -Category 'system' -RequiresAdmin -Disabled:$SkipWindowsUpdate -DisabledReason 'disabled by -SkipWindowsUpdate' -TimeoutSec 7200 -Script {
                if (Get-Command Install-WindowsUpdate -ErrorAction SilentlyContinue) {
                    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -ErrorAction Continue
                }
                elseif (Get-Command UsoClient.exe -ErrorAction SilentlyContinue) {
                    UsoClient.exe StartScan
                    UsoClient.exe StartDownload
                    UsoClient.exe StartInstall
                    Write-Output 'Windows Update started through UsoClient. Completion may continue in the background.'
                }
                else {
                    throw 'Neither PSWindowsUpdate nor UsoClient.exe is available.'
                }
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'defender' -Category 'system' -RequiresCommand 'Update-MpSignature' -Disabled:$SkipDefender -DisabledReason 'disabled by -SkipDefender' -Script {
                Update-MpSignature -ErrorAction Continue
            } -Tags @('windows', 'security'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl' -Category 'system' -RequiresCommand 'wsl' -Disabled:$SkipWSL -DisabledReason 'disabled by -SkipWSL' -Script {
                wsl --update
            } -Tags @('windows', 'linux'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl-distros' -Category 'system' -RequiresCommand 'wsl' -Disabled:($SkipWSL -or $SkipWSLDistros) -DisabledReason 'disabled by WSL skip switch' -TimeoutSec 3600 -Script {
                $distros = @(wsl -l -q 2>$null |
                    ForEach-Object { ([string]$_).Replace([string][char]0, '').Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                if ($distros.Count -eq 0) {
                    Write-Output 'No WSL distros found.'
                    return
                }

                $failedDistros = [System.Collections.Generic.List[string]]::new()
                foreach ($distro in $distros) {
                    Write-Output "Updating WSL distro: $distro"
                    wsl --distribution $distro --exec sh -lc 'if command -v apt >/dev/null 2>&1; then sudo -n true >/dev/null 2>&1 && sudo -n apt update && sudo -n DEBIAN_FRONTEND=noninteractive apt -y upgrade && sudo -n apt -y autoremove || echo "Skipping apt: sudo requires a password"; elif command -v pacman >/dev/null 2>&1; then sudo -n true >/dev/null 2>&1 && sudo -n pacman -Syu --noconfirm || echo "Skipping pacman: sudo requires a password"; else echo "No supported package manager found"; fi'
                    if ($LASTEXITCODE -ne 0) { [void]$failedDistros.Add($distro) }
                }

                if ($failedDistros.Count -gt 0) {
                    throw "WSL distro updates failed: $($failedDistros -join ', ')"
                }
            } -Tags @('windows', 'linux'))) | Out-Null

    $npmScript = {
        param([string[]]$SkipPackages)

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($package in $SkipPackages) {
            if (-not [string]::IsNullOrWhiteSpace($package)) { [void]$skipSet.Add($package.Trim()) }
        }

        $root = (npm root -g 2>$null | Select-Object -First 1)
        if ($root -and (Test-Path -LiteralPath $root)) {
            $dotFolders = @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Name.StartsWith('.') } |
                Select-Object -ExpandProperty Name)
            if ($dotFolders.Count -gt 0) {
                Write-Output "Ignoring stale npm temp folders: $($dotFolders -join ', ')"
            }
        }

        $listJson = (npm ls -g --depth=0 --json 2>$null | Out-String).Trim()
        if (-not $listJson) {
            throw 'npm did not return a global package list.'
        }

        $tree = $listJson | ConvertFrom-Json -ErrorAction Stop
        $packageNames = @()
        if ($tree.PSObject.Properties['dependencies'] -and $tree.dependencies) {
            $packageNames = @($tree.dependencies.PSObject.Properties.Name)
        }

        if ($packageNames.Count -eq 0) {
            Write-Output 'No global npm packages found.'
            return
        }

        $validNamePattern = '^(?:@[a-z0-9][a-z0-9._~-]*/)?[a-z0-9][a-z0-9._~-]*$'
        $failed = [System.Collections.Generic.List[string]]::new()
        foreach ($name in ($packageNames | Sort-Object -Unique)) {
            if ($skipSet.Contains($name)) {
                Write-Output "Skipping npm package from config: $name"
                continue
            }
            if ($name -notmatch $validNamePattern) {
                Write-Output "Skipping invalid npm package name: $name"
                continue
            }

            $spec = "$name@latest"
            Write-Output "Updating npm package: $spec"
            npm install -g $spec --no-fund --no-audit
            if ($LASTEXITCODE -ne 0) { [void]$failed.Add($name) }
        }

        if ($failed.Count -gt 0) {
            throw "npm failed packages: $($failed -join ', ')"
        }
    }

    $tasks.Add((New-UpdateTask -Name 'npm' -Category 'javascript' -RequiresCommand 'npm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script $npmScript -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pnpm' -Category 'javascript' -RequiresCommand 'pnpm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                pnpm self-update
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'yarn' -Category 'javascript' -RequiresCommand 'yarn' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                yarn global upgrade
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'bun' -Category 'javascript' -RequiresCommand 'bun' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                bun upgrade
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'deno' -Category 'javascript' -RequiresCommand 'deno' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                deno upgrade
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'mise' -Category 'version-manager' -RequiresCommand 'mise' -Script {
                mise self-update --yes
                if ($LASTEXITCODE -ne 0) { throw "mise self-update failed with exit code $LASTEXITCODE" }
                mise upgrade --yes
            } -Tags @('version-manager'))) | Out-Null

    $pipScript = [scriptblock]::Create(@"
param([string[]]`$SkipPackages)
python -m pip install --upgrade pip
`$exitCode = `$LASTEXITCODE
if (`$exitCode -ne 0) { throw "pip self-upgrade failed with exit code `$exitCode" }

`$skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach (`$pkg in `$SkipPackages) {
    if (-not [string]::IsNullOrWhiteSpace(`$pkg)) { [void]`$skipSet.Add(`$pkg.Trim()) }
}
if (`$SkipPackages.Count -gt 0) {
    Write-Output "Configured pip package skip list: `$(`$SkipPackages -join ', ')"
}

`$outdated = python -m pip list --outdated --format=json 2>`$null | ConvertFrom-Json -ErrorAction SilentlyContinue
if (-not `$outdated) { Write-Output 'No outdated pip packages found.'; return }

`$failed = [System.Collections.Generic.List[string]]::new()
foreach (`$pkg in `$outdated) {
    if (`$skipSet.Contains(`$pkg.name)) { Write-Output "Skipping pip package: `$(`$pkg.name)"; continue }
    Write-Output "Upgrading pip package: `$(`$pkg.name) `$(`$pkg.version) -> `$(`$pkg.latest_version)"
    python -m pip install --upgrade `$pkg.name
    if (`$LASTEXITCODE -ne 0) { [void]`$failed.Add(`$pkg.name) }
}
if (`$failed.Count -gt 0) { throw "pip failed packages: `$(`$failed -join ', ')" }
"@)

    $tasks.Add((New-UpdateTask -Name 'pip' -Category 'python' -RequiresCommand 'python' -Disabled:$false -Script $pipScript -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pipx' -Category 'python' -RequiresCommand 'pipx' -Script {
                pipx upgrade-all
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv' -Category 'python' -RequiresCommand 'uv' -Script {
                $out = uv self update 2>&1 | Out-String
                $ec  = $LASTEXITCODE
                if ($ec -ne 0 -and $out -match 'standalone installation') {
                    # uv installed via pip/Python package — self-update not supported; skip cleanly
                    $uvPath = (Get-Command uv -ErrorAction SilentlyContinue).Source
                    Write-Output "uv self-update skipped: '$uvPath' is a managed install (use pip/uv-pip to upgrade)"
                    $global:LASTEXITCODE = 0
                    return
                }
                if ($ec -ne 0) { throw "uv self update failed with exit code $ec" }
                Write-Output $out.Trim()
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv-tools' -Category 'python' -RequiresCommand 'uv' -Disabled:$SkipUVTools -DisabledReason 'disabled by -SkipUVTools' -Script {
                uv tool upgrade --all
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'poetry' -Category 'python' -RequiresCommand 'poetry' -Disabled:$SkipPoetry -DisabledReason 'disabled by -SkipPoetry' -Script {
                poetry self update
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'rustup' -Category 'systems-language' -RequiresCommand 'rustup' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                rustup update
            } -Tags @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'cargo' -Category 'systems-language' -RequiresCommand 'cargo' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                if (-not (Get-Command cargo-install-update -ErrorAction SilentlyContinue)) {
                    cargo install cargo-update -q
                    if ($LASTEXITCODE -ne 0) { throw "cargo-update install failed with exit code $LASTEXITCODE" }
                }
                cargo install-update -a
            } -Tags @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'go' -Category 'systems-language' -RequiresCommand 'go' -Disabled:$SkipGo -DisabledReason 'disabled by -SkipGo' -Script {
                go install golang.org/x/tools/gopls@latest
            } -Tags @('go'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'flutter' -Category 'systems-language' -RequiresCommand 'flutter' -Disabled:$SkipFlutter -DisabledReason 'disabled by -SkipFlutter' -Script {
                flutter upgrade
            } -Tags @('flutter'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-tools' -Category 'dotnet' -RequiresCommand 'dotnet' -Script {
                dotnet tool update --global --all
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-workloads' -Category 'dotnet' -RequiresCommand 'dotnet' -RequiresAdmin -Script {
                dotnet workload update
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'ruby-gems' -Category 'runtime' -RequiresCommand 'gem' -Disabled:$SkipRuby -DisabledReason 'disabled by -SkipRuby' -Script {
                gem update --system
                if ($LASTEXITCODE -ne 0) { throw "gem update --system failed with exit code $LASTEXITCODE" }
                gem update
            } -Tags @('ruby'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'composer' -Category 'runtime' -RequiresCommand 'composer' -Disabled:$SkipComposer -DisabledReason 'disabled by -SkipComposer' -Script {
                composer self-update
            } -Tags @('php'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'vscode-extensions' -Category 'dev-tools' -RequiresCommand 'code' -Disabled:$SkipVSCodeExtensions -DisabledReason 'disabled by -SkipVSCodeExtensions' -Script {
                code --update-extensions
            } -Tags @('editor'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'git-lfs' -Category 'dev-tools' -RequiresCommand 'git-lfs' -Disabled:$SkipGitLFS -DisabledReason 'disabled by -SkipGitLFS' -Script {
                git lfs install --skip-repo
            } -Tags @('git'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'gh-extensions' -Category 'dev-tools' -RequiresCommand 'gh' -Script {
                $extensions = @(gh extension list 2>$null |
                    ForEach-Object { ($_ -split '\s+')[0] } |
                    Where-Object { $_ -match '^gh-' })
                foreach ($extension in $extensions) {
                    gh extension upgrade $extension
                }
            } -Tags @('github'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'powershell-modules' -Category 'powershell' -Disabled:$SkipPowerShellModules -DisabledReason 'disabled by -SkipPowerShellModules' -Script {
                if (Get-Command Update-PSResource -ErrorAction SilentlyContinue) {
                    $moduleArgs = @{ Name = '*'; ErrorAction = 'Continue' }
                    $cmd = Get-Command Update-PSResource -ErrorAction SilentlyContinue
                    if ($cmd.Parameters.ContainsKey('AcceptLicense')) { $moduleArgs.AcceptLicense = $true }
                    Update-PSResource @moduleArgs
                }
                elseif (Get-Command Update-Module -ErrorAction SilentlyContinue) {
                    Get-InstalledModule -ErrorAction SilentlyContinue | ForEach-Object {
                        Update-Module -Name $_.Name -Force -ErrorAction Continue
                    }
                }
                else {
                    Write-Output 'No PowerShell module updater found.'
                }
            } -Tags @('powershell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'powershell-help' -Category 'powershell' -Script {
                Update-Help -Force -ErrorAction Continue
            } -Tags @('powershell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'ollama-models' -Category 'ai' -RequiresCommand 'ollama' -Disabled:$false -TimeoutSec 7200 -Script {
                $models = @(ollama list 2>$null | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] } | Where-Object { $_ })
                foreach ($model in $models) {
                    ollama pull $model
                }
            } -Tags @('ai'))) | Out-Null

    $cleanupScript = [scriptblock]::Create(@"
param([int]`$Days, [bool]`$Deep, [bool]`$SkipDestructive)
`$cutoff = (Get-Date).AddDays(-`$Days)
`$paths = @(`$env:TEMP)
if (`$IsWindows -and (Test-Path 'C:\Windows\Temp')) { `$paths += 'C:\Windows\Temp' }

foreach (`$path in `$paths | Where-Object { `$_ -and (Test-Path `$_) }) {
    Write-Output "Cleaning temp files older than `$Days day(s): `$path"
    Get-ChildItem -LiteralPath `$path -Force -ErrorAction SilentlyContinue |
        Where-Object { `$_.LastWriteTime -lt `$cutoff -and `$_.FullName -notmatch '\\WinGet(\\|`$)' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

if (`$IsWindows) {
    Clear-DnsClientCache -ErrorAction SilentlyContinue
    if (-not `$SkipDestructive) {
        Clear-RecycleBin -Force -ErrorAction SilentlyContinue
    }
    if (`$Deep) {
        DISM /Online /Cleanup-Image /StartComponentCleanup
        Clear-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
    }
}
"@)

    $tasks.Add((New-UpdateTask -Name 'cleanup' -Category 'maintenance' -Disabled:$SkipCleanup -DisabledReason 'disabled by -SkipCleanup' -TimeoutSec 3600 -Script $cleanupScript -Tags @('maintenance'))) | Out-Null

    foreach ($task in $tasks) {
        if ($task.Id -eq 'winget') {
            $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $wingetSkip } -Force
        }
        elseif ($task.Id -eq 'pip') {
            $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $pipSkip } -Force
        }
        elseif ($task.Id -eq 'npm') {
            $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $npmSkip } -Force
        }
        elseif ($task.Id -eq 'cleanup') {
            $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                Days            = $tempCleanupDays
                Deep            = [bool]$DeepClean
                SkipDestructive = [bool]$SkipDestructive
            } -Force
        }
        else {
            $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
        }
    }

    return @($tasks)
}

function Get-FilteredTasks {
    param(
        [Parameter(Mandatory)][object[]]$Tasks,
        [Parameter(Mandatory)][bool]$IsAdmin
    )

    $onlyPatterns = @(ConvertTo-FilterList $Only)
    $skipPatterns = @()
    $skipPatterns += ConvertTo-FilterList $script:Config.SkipManagers
    $skipPatterns += ConvertTo-FilterList $Skip

    if ($FastMode -or $UltraFast) {
        $skipPatterns += ConvertTo-FilterList $script:Config.FastModeSkip
    }
    if ($UltraFast) {
        $skipPatterns += ConvertTo-FilterList $script:Config.UltraFastSkip
    }

    $planned = [System.Collections.Generic.List[object]]::new()
    $skipped = [System.Collections.Generic.List[object]]::new()

    foreach ($task in $Tasks) {
        $reason = $null

        if ($onlyPatterns.Count -gt 0 -and -not (Test-NameMatch -Task $task -Patterns $onlyPatterns)) {
            $reason = 'not selected by -Only'
        }
        elseif ($skipPatterns.Count -gt 0 -and (Test-NameMatch -Task $task -Patterns $skipPatterns)) {
            $reason = 'skipped by filter'
        }
        elseif ($task.Disabled) {
            $reason = if ($task.DisabledReason) { $task.DisabledReason } else { 'disabled' }
        }
        elseif ($task.RequiresAdmin -and -not $IsAdmin) {
            $reason = 'requires Administrator'
        }
        else {
            foreach ($command in $task.RequiresCommand) {
                if (-not (Test-Command $command)) {
                    $reason = "missing command: $command"
                    break
                }
            }
        }

        if ($reason) {
            $skipped.Add([pscustomobject]@{
                    Name     = $task.Name
                    Id       = $task.Id
                    Category = $task.Category
                    Status   = 'Skipped'
                    Reason   = $reason
                }) | Out-Null
        }
        else {
            $planned.Add($task) | Out-Null
        }
    }

    return [pscustomobject]@{
        Planned = @($planned)
        Skipped = @($skipped)
    }
}

function Show-TaskList {
    param(
        [object[]]$Planned,
        [object[]]$Skipped
    )

    Write-Host ''
    Write-Host 'Planned tasks' -ForegroundColor Cyan
    if ($Planned.Count -eq 0) {
        Write-Host '  none' -ForegroundColor DarkGray
    }
    else {
        $Planned |
        Sort-Object Category, Name |
        Format-Table Name, Category, TimeoutSec, RequiresAdmin -AutoSize
    }

    Write-Host ''
    Write-Host 'Skipped tasks' -ForegroundColor DarkGray
    if ($Skipped.Count -eq 0) {
        Write-Host '  none' -ForegroundColor DarkGray
    }
    else {
        $Skipped |
        Sort-Object Category, Name |
        Format-Table Name, Category, Reason -AutoSize
    }
}

function New-TaskResult {
    param(
        [Parameter(Mandatory)]$Task,
        [Parameter(Mandatory)][string]$Status,
        [int]$ExitCode = 0,
        [double]$DurationSeconds = 0,
        [int]$Attempts = 1,
        [string[]]$Output = @(),
        [string]$Reason
    )

    [pscustomobject]@{
        Name            = $Task.Name
        Id              = $Task.Id
        Category        = $Task.Category
        Status          = $Status
        ExitCode        = $ExitCode
        DurationSeconds = [Math]::Round($DurationSeconds, 2)
        Attempts        = $Attempts
        Reason          = $Reason
        OutputPreview   = @($Output | Select-Object -First 20)
    }
}

function Start-UpdateTaskJob {
    param(
        [Parameter(Mandatory)]$Task,
        [Parameter(Mandatory)][int]$Attempt
    )

    $scriptText = $Task.Script.ToString()
    $argumentMap = if ($Task.PSObject.Properties['Arguments']) { $Task.Arguments } else { @{} }

    Start-ThreadJob -Name $Task.Name -ScriptBlock {
        param($TaskName, $TaskId, $Category, $ScriptText, $ArgumentMap, $Attempt)

        $ErrorActionPreference = 'Continue'
        $start = Get-Date
        $output = [System.Collections.Generic.List[string]]::new()
        $exitCode = 0
        $status = 'Succeeded'
        $reason = $null

        try {
            $block = [scriptblock]::Create($ScriptText)
            & $block @ArgumentMap 2>&1 | ForEach-Object {
                $line = if ($null -eq $_) { '' } else { ([string]$_).Replace([string][char]0, '').Trim() }
                $line = [regex]::Replace($line, '\x1b\[[0-9;?]*[a-zA-Z]', '')
                $line = [regex]::Replace($line, '\x1b\[[0-9;]*m', '')
                if ($line.Trim()) { $output.Add($line) | Out-Null }
            }

            if ($null -ne $global:LASTEXITCODE) {
                $exitCode = [int]$global:LASTEXITCODE
            }

            if ($exitCode -ne 0) {
                $status = 'Failed'
                $reason = "exit code $exitCode"
            }
        }
        catch {
            $status = 'Failed'
            $exitCode = 1
            $reason = $_.Exception.Message
            $output.Add($_.Exception.Message) | Out-Null
        }

        [pscustomobject]@{
            Name            = $TaskName
            Id              = $TaskId
            Category        = $Category
            Status          = $status
            ExitCode        = $exitCode
            DurationSeconds = [Math]::Round(((Get-Date) - $start).TotalSeconds, 2)
            Attempts        = $Attempt
            Reason          = $reason
            Output          = @($output)
        }
    } -ArgumentList $Task.Name, $Task.Id, $Task.Category, $scriptText, $argumentMap, $Attempt
}

function Invoke-TaskQueue {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][object[]]$Tasks,
        [Parameter(Mandatory)][int]$Throttle
    )

    $queue = [System.Collections.Queue]::new()
    foreach ($task in $Tasks) { $queue.Enqueue([pscustomobject]@{ Task = $task; Attempt = 1 }) }

    $running = @()
    $results = [System.Collections.Generic.List[object]]::new()

    while ($queue.Count -gt 0 -or $running.Count -gt 0) {
        while ($running.Count -lt $Throttle -and $queue.Count -gt 0) {
            $entry = $queue.Dequeue()
            $task = $entry.Task

            if (-not $PSCmdlet.ShouldProcess($task.Name, 'Run update task')) {
                $results.Add((New-TaskResult -Task $task -Status 'Skipped' -Reason 'ShouldProcess declined')) | Out-Null
                continue
            }

            Write-Status ("[{0}] starting (attempt {1})" -f $task.Name, $entry.Attempt) -Level Info
            $job = Start-UpdateTaskJob -Task $task -Attempt $entry.Attempt
            $running += [pscustomobject]@{
                Job     = $job
                Task    = $task
                Attempt = $entry.Attempt
                Started = Get-Date
            }
        }

        foreach ($item in @($running)) {
            $job = $item.Job
            $task = $item.Task
            $timedOut = ((Get-Date) - $item.Started).TotalSeconds -gt $task.TimeoutSec

            if ($timedOut) {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

                $result = New-TaskResult -Task $task -Status 'TimedOut' -ExitCode 124 -DurationSeconds $task.TimeoutSec -Attempts $item.Attempt -Reason "timeout after $($task.TimeoutSec)s"
                $results.Add($result) | Out-Null
                Write-Status ("[{0}] timed out after {1}s" -f $task.Name, $task.TimeoutSec) -Level Warning
                $running = @($running | Where-Object { $_.Job.Id -ne $job.Id })
                continue
            }

            if ($job.State -notin @('Completed', 'Failed', 'Stopped')) { continue }

            $raw = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $running = @($running | Where-Object { $_.Job.Id -ne $job.Id })

            $payload = $raw | Where-Object { $_.PSObject.Properties['Status'] } | Select-Object -Last 1
            $output = @($raw | Where-Object { -not $_.PSObject.Properties['Status'] } | ForEach-Object { $_.ToString() })

            if (-not $payload) {
                $payload = [pscustomobject]@{
                    Name            = $task.Name
                    Id              = $task.Id
                    Category        = $task.Category
                    Status          = 'Failed'
                    ExitCode        = 1
                    DurationSeconds = [Math]::Round(((Get-Date) - $item.Started).TotalSeconds, 2)
                    Attempts        = $item.Attempt
                    Reason          = "job ended in state $($job.State)"
                    Output          = $output
                }
            }

            $combinedOutput = @()
            if ($payload.PSObject.Properties['Output']) { $combinedOutput += @($payload.Output) }
            $combinedOutput += $output

            $result = New-TaskResult -Task $task -Status $payload.Status -ExitCode $payload.ExitCode -DurationSeconds $payload.DurationSeconds -Attempts $payload.Attempts -Output $combinedOutput -Reason $payload.Reason

            if ($result.Status -eq 'Failed' -and $item.Attempt -le $RetryCount) {
                Write-Status ("[{0}] failed; queueing retry {1}/{2}" -f $task.Name, $item.Attempt, $RetryCount) -Level Warning
                $queue.Enqueue([pscustomobject]@{ Task = $task; Attempt = ($item.Attempt + 1) })
                continue
            }

            $results.Add($result) | Out-Null

            if ($result.Status -eq 'Succeeded') {
                Write-Status ("[{0}] ok ({1:N1}s)" -f $task.Name, $result.DurationSeconds) -Level Success
            }
            else {
                Write-Status ("[{0}] {1}: {2}" -f $task.Name, $result.Status, $result.Reason) -Level Warning
            }

            if ($combinedOutput.Count -gt 0) {
                foreach ($line in $combinedOutput) {
                    Write-Log -Message ("[{0}] {1}" -f $task.Name, $line) -Level Muted
                }
                Write-Host "  [$($task.Name) output]" -ForegroundColor DarkGray
                foreach ($line in $combinedOutput) {
                    Write-Host "    $line" -ForegroundColor DarkGray
                }
            }
            else {
                Write-Host "  [$($task.Name) output] (no output)" -ForegroundColor DarkGray
            }
        }

        if ($running.Count -gt 0) {
            $done = $results.Count
            $total = $Tasks.Count
            $runningNames = ($running | ForEach-Object { $_.Task.Name }) -join ', '
            $percent = if ($total -gt 0) { [Math]::Min(99, [int](($done / $total) * 100)) } else { 100 }
            Write-Progress -Activity 'Update Everything' -Status "$done/$total done. Running: $runningNames" -PercentComplete $percent
            Start-Sleep -Milliseconds 250
        }
    }

    Write-Progress -Activity 'Update Everything' -Completed
    return @($results)
}

function Save-RunSummary {
    param(
        [object[]]$Results,
        [object[]]$Skipped,
        [object[]]$Planned
    )

    $duration = [Math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 2)
    $summary = [ordered]@{
        RunId            = $script:RunId
        StartedAt        = $script:StartTime.ToString('o')
        FinishedAt       = (Get-Date).ToString('o')
        DurationSeconds  = $duration
        DryRun           = [bool]$script:IsSimulation
        FastMode         = [bool]$FastMode
        UltraFast        = [bool]$UltraFast
        ParallelThrottle = $ParallelThrottle
        LogPath          = $LogPath
        PlannedCount     = $Planned.Count
        SucceededCount   = @($Results | Where-Object Status -eq 'Succeeded').Count
        FailedCount      = @($Results | Where-Object { $_.Status -in @('Failed', 'TimedOut') }).Count
        SkippedCount     = $Skipped.Count
        Results          = @($Results)
        Skipped          = @($Skipped)
    }

    $json = $summary | ConvertTo-Json -Depth 8
    try {
        $summaryDir = Split-Path -Parent $JsonSummaryPath
        if ($summaryDir -and -not (Test-Path -LiteralPath $summaryDir)) {
            New-Item -Path $summaryDir -ItemType Directory -Force -WhatIf:$false | Out-Null
        }
        Set-Content -LiteralPath $JsonSummaryPath -Value $json -Encoding utf8 -WhatIf:$false
    }
    catch {
        Write-Status "Could not write JSON summary: $($_.Exception.Message)" -Level Warning
    }

    return [pscustomobject]$summary
}

function Show-WhatChanged {
    param([Parameter(Mandatory)]$CurrentSummary)

    if (-not (Test-Path -LiteralPath $script:PreviousJsonSummaryPath)) {
        Write-Status 'No previous run summary found for -WhatChanged.' -Level Warning
        return
    }

    try {
        $previous = Get-Content -LiteralPath $script:PreviousJsonSummaryPath -Raw | ConvertFrom-Json
        $previousMap = @{}
        foreach ($item in @($previous.Results)) { $previousMap[$item.Id] = $item }

        $changes = [System.Collections.Generic.List[string]]::new()
        foreach ($item in @($CurrentSummary.Results)) {
            if (-not $previousMap.ContainsKey($item.Id)) {
                $changes.Add("new task: $($item.Name) -> $($item.Status)") | Out-Null
                continue
            }

            $old = $previousMap[$item.Id]
            if ($old.Status -ne $item.Status) {
                $changes.Add("$($item.Name): $($old.Status) -> $($item.Status)") | Out-Null
            }
        }

        if ($changes.Count -eq 0) {
            Write-Status 'WhatChanged: task statuses match the previous run.' -Level Muted
        }
        else {
            Write-Status 'WhatChanged:' -Level Info
            foreach ($change in $changes) {
                Write-Status "  $change" -Level Muted
            }
        }
    }
    catch {
        Write-Status "WhatChanged failed: $($_.Exception.Message)" -Level Warning
    }
}

function Register-UpdateSchedule {
    if (-not $IsWindows) { throw 'Scheduling is only supported on Windows.' }
    if (-not (Test-IsAdmin)) { throw 'Scheduled task registration requires Administrator.' }

    $pwsh = Get-CommandPath 'pwsh.exe'
    if (-not $pwsh) { $pwsh = 'powershell.exe' }

    $taskName = 'DailySystemUpdate'
    $taskArgs = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath),
        '-SkipReboot',
        '-NoPause',
        '-SkipWSL',
        '-SkipWindowsUpdate',
        '-Quiet'
    ) -join ' '

    if ($script:IsSimulation) {
        Write-Status "[DryRun] Would register scheduled task $taskName at $ScheduleTime." -Level Info
        return
    }

    $action = New-ScheduledTaskAction -Execute $pwsh -Argument $taskArgs
    $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 4)
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Highest

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
    Write-Status "Scheduled $taskName daily at $ScheduleTime." -Level Success
}

function Invoke-SelfElevation {
    if (-not $AutoElevate -or $NoElevate -or -not $IsWindows -or (Test-IsAdmin)) { return $false }

    $pwsh = Get-CommandPath 'pwsh.exe'
    if (-not $pwsh) { $pwsh = 'powershell.exe' }

    $forwarded = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Key -eq 'AutoElevate') { continue }
        if ($entry.Value -is [switch]) {
            if ($entry.Value.IsPresent) { $forwarded.Add("-$($entry.Key)") | Out-Null }
        }
        elseif ($entry.Value -is [array]) {
            if ($entry.Value.Count -gt 0) {
                $forwarded.Add("-$($entry.Key)") | Out-Null
                $quoted = @($entry.Value | ForEach-Object { '"{0}"' -f ([string]$_).Replace('"', '\"') })
                $forwarded.Add($quoted -join ',') | Out-Null
            }
        }
        else {
            $forwarded.Add("-$($entry.Key)") | Out-Null
            $forwarded.Add(('"{0}"' -f ([string]$entry.Value).Replace('"', '\"'))) | Out-Null
        }
    }

    $elevateArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) + @($forwarded)
    Start-Process -FilePath $pwsh -Verb RunAs -ArgumentList ($elevateArgs -join ' ') -Wait
    return $true
}

Import-UpdateConfig
Initialize-RunStorage

if (Invoke-SelfElevation) {
    exit 0
}

$isAdmin = Test-IsAdmin
if (-not $isAdmin -and -not $NoElevate) {
    Write-Status 'Running without Administrator. Admin-only tasks will be skipped. Use -AutoElevate for a full run.' -Level Warning
}

if ($Schedule) {
    Register-UpdateSchedule
    exit 0
}

if ($UltraFast) { $FastMode = $true }
if ($ParallelThrottle -lt 1) {
    $ParallelThrottle = [Math]::Max(2, [Math]::Min([Environment]::ProcessorCount, 6))
}
if ($NoParallel) { $ParallelThrottle = 1 }

$allTasks = @(Get-UpdateTasks)
$filtered = Get-FilteredTasks -Tasks $allTasks -IsAdmin $isAdmin
$plannedTasks = @($filtered.Planned)
$skippedTasks = @($filtered.Skipped)

if ($SelfTest) {
    $selfTestTask = New-UpdateTask -Name 'self-test' -Category 'diagnostics' -Script {
        Write-Output 'Scheduler, logging, and summary path are working.'
    }
    $selfTestTask | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
    $plannedTasks = @($selfTestTask)
    $skippedTasks = @()
}

Write-Status ("Update-Everything v6.1.0 | {0:yyyy-MM-dd HH:mm} | throttle {1}" -f $script:StartTime, $ParallelThrottle) -Level Info

if ($ListTasks) {
    Show-TaskList -Planned $plannedTasks -Skipped $skippedTasks
    $summary = Save-RunSummary -Results @() -Skipped $skippedTasks -Planned $plannedTasks
    Write-Status "Task list written to $JsonSummaryPath" -Level Muted
    exit 0
}

if ($plannedTasks.Count -eq 0) {
    Write-Status 'No runnable update tasks were found.' -Level Warning
    $summary = Save-RunSummary -Results @() -Skipped $skippedTasks -Planned $plannedTasks
    exit 0
}

if ($script:IsSimulation) {
    Write-Status 'Dry run: no update commands will be executed.' -Level Info
    foreach ($task in $plannedTasks) {
        Write-Status ("[DryRun] {0} ({1})" -f $task.Name, $task.Category) -Level Muted
    }

    $dryResults = @($plannedTasks | ForEach-Object {
            New-TaskResult -Task $_ -Status 'DryRun' -Reason 'preview only'
        })
    $summary = Save-RunSummary -Results $dryResults -Skipped $skippedTasks -Planned $plannedTasks
    if ($WhatChanged) { Show-WhatChanged -CurrentSummary $summary }
    Write-Status "Dry-run summary written to $JsonSummaryPath" -Level Success
    exit 0
}

Write-Status ("Dispatching {0} task(s). Skipped: {1}" -f $plannedTasks.Count, $skippedTasks.Count) -Level Info
$results = Invoke-TaskQueue -Tasks $plannedTasks -Throttle $ParallelThrottle
$summary = Save-RunSummary -Results $results -Skipped $skippedTasks -Planned $plannedTasks

if ($WhatChanged) {
    Show-WhatChanged -CurrentSummary $summary
}

$failed = @($results | Where-Object { $_.Status -in @('Failed', 'TimedOut') })
$succeeded = @($results | Where-Object { $_.Status -eq 'Succeeded' })
$elapsed = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')

Write-Host ''
if ($failed.Count -gt 0) {
    Write-Status ("Completed with {0} succeeded, {1} failed/timed out, {2} skipped in {3}." -f $succeeded.Count, $failed.Count, $skippedTasks.Count, $elapsed) -Level Warning
    Write-Status "Summary: $JsonSummaryPath" -Level Muted
    Write-Status "Log: $LogPath" -Level Muted
    exit 1
}

Write-Status ("All runnable tasks completed: {0} succeeded, {1} skipped in {2}." -f $succeeded.Count, $skippedTasks.Count, $elapsed) -Level Success
Write-Status "Summary: $JsonSummaryPath" -Level Muted
Write-Status "Log: $LogPath" -Level Muted

if (-not $SkipReboot -and $IsWindows) {
    $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
    (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    if ($pendingReboot) {
        Write-Status 'A reboot appears to be pending.' -Level Warning
    }
}
