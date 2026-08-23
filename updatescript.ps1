#requires -version 7.0
<#
.SYNOPSIS
    Updates common Windows 11 package managers, developer tools, runtimes, WSL distros, Defender, and maintenance tasks.

.VERSION
    Update-Everything v7.0.0-mega

.NOTES
    v7.0.0-mega: major expansion — 20 new tool sections, 20 new features.
      New tools: vcpkg, conda, gcloud, az, aws, terraform, pulumi, kubectl, helm, hugo,
                  opentofu, starship, zoxide, gitleaks,
                  trivy, packer, nvm, devcontainer, self-update, windows-features, appx-repair.
      New features: -HealthCheck (inventory), -Snapshot (state recording), -Interactive (per-task
                 confirmation), -Profile (work/personal/gaming/minimal), checkpoint/resume,
                 webhook alerting (-WebhookUrl), multi-machine state sync (-RemoteStatePath),
                 schedule enhancement (-ScheduleRepeat/-ScheduleDays), before/after hooks,
                 stale binary detection, per-tool update-config.json config, post-update
                 version diff, and winget batch upgrade fallback.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'the -Skip* switches are consumed as -Disabled:$SkipX, which PSSA does not count as a use')]
param(
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,
    [switch]$UltraFast,
    [switch]$NoElevate,
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidDefaultValueSwitchParameter', '', Justification = 'elevation is the default; -NoElevate opts out')]
    [switch]$AutoElevate = $true,
    [switch]$SkipWSL,
    [switch]$SkipWSLDistros,
    [switch]$SkipDefender,
    [switch]$SkipStoreApps,
    [switch]$SkipUVTools,
    [switch]$SkipPipHealth,
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
    [switch]$SkipVcpkg,
    [switch]$SkipConda,
    [switch]$SkipCloudTools,
    [switch]$SkipContainerTools,
    [switch]$SkipInfraTools,
    [switch]$SkipK8sTools,
    [switch]$SkipStarship,
    [switch]$SkipHugo,
    [switch]$DeepClean,
    [switch]$UpdateOllamaModels,
    [switch]$HealthCheck,
    [switch]$Interactive,
    [switch]$Snapshot,
    [switch]$WhatChanged,
    [switch]$DryRun,
    [switch]$ListTasks,
    [switch]$SelfTest,
    [switch]$NoParallel,
    [switch]$Quiet,
    [switch]$ShowSkipped,
    [Alias('IncludeProtectedApps')]
    [switch]$BypassProtection,
    [switch]$NoSilentInstallers,
    [Alias('UpdateHelp')]
    [switch]$UpdatePowerShellHelp,
    [switch]$Schedule,

    [ValidateScript({ $_ -match '^([01]?[0-9]|2[0-3]):[0-5][0-9]$' })]
    [string]$ScheduleTime = '03:00',

    [ValidateRange(30, 7200)]
    [int]$WingetTimeoutSec = 600,

    [ValidateRange(60, 14400)]
    [int]$TaskTimeoutSec = 1800,

    [ValidateRange(30, 7200)]
    [int]$OllamaTimeoutSec = 600,

    [ValidateRange(0, 16)]
    [int]$ParallelThrottle = 0,

    [ValidateRange(0, 5)]
    [int]$RetryCount = 1,

    [string]$LogPath,
    [string]$JsonSummaryPath,
    [string]$StateDir,
    [string[]]$Only = @(),
    [string[]]$Skip = @(),

    [ValidateSet('work', 'personal', 'gaming', 'minimal')]
    [string]$Profile,

    [string]$WebhookUrl,

    [string[]]$ScheduleDays = @('Mon', 'Tue', 'Wed', 'Thu', 'Fri'),

    [ValidateRange(0, 24)]
    [int]$ScheduleRepeat,

    [string]$RemoteStatePath,

    [switch]$CIMode,
    [switch]$Notify,

    [ValidateRange(0, 8760)]
    [int]$SkipSucceededWithinHours = 0
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Continue'

try
{
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [Console]::InputEncoding = $utf8NoBom
    [Console]::OutputEncoding = $utf8NoBom
    $OutputEncoding = $utf8NoBom
} catch
{
    Write-Verbose "Console encoding setup skipped: $($_.Exception.Message)"
}

$script:Version = '7.0.0-mega'
$script:StartTime = Get-Date
$script:RunId = $script:StartTime.ToString('yyyyMMdd-HHmmss-fff')
$script:CommandCache = @{}
$script:StateDirWasProvided = -not [string]::IsNullOrWhiteSpace($StateDir)
$script:LogPathWasProvided = -not [string]::IsNullOrWhiteSpace($LogPath)
$script:JsonSummaryPathWasProvided = -not [string]::IsNullOrWhiteSpace($JsonSummaryPath)
$script:LogWriteWarningEmitted = $false
$script:IsSimulation = $DryRun -or $WhatIfPreference
$script:IsHealthCheck = [bool]$HealthCheck
$script:IsSnapshot = [bool]$Snapshot
$script:IsInteractive = [bool]$Interactive
$script:CheckpointPath = $null
$script:CompletedTasks = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:LockPath = $null

if (-not $StateDir)
{
    $localAppData = [Environment]::GetFolderPath('LocalApplicationData')
    if ([string]::IsNullOrWhiteSpace($localAppData))
    { $localAppData = [System.IO.Path]::GetTempPath()
    }
    $StateDir = Join-Path $localAppData 'Update-Everything'
}

$script:StateDir = $StateDir
$script:LogDir = Join-Path $script:StateDir 'logs'
$script:DefaultJsonSummaryPath = Join-Path $script:StateDir 'last-run.json'
$script:PreviousJsonSummaryPath = Join-Path $script:StateDir 'previous-run.json'
$script:CheckpointPath = Join-Path $script:StateDir 'checkpoint.json'

if ($RemoteStatePath -and (Test-Path -LiteralPath $RemoteStatePath))
{
    try
    {
        $remoteSummary = Join-Path $RemoteStatePath 'last-run.json'
        if (Test-Path -LiteralPath $remoteSummary)
        {
            Copy-Item -LiteralPath $remoteSummary -Destination $script:DefaultJsonSummaryPath -Force -ErrorAction Stop
            Write-Status "Loaded remote state from $remoteSummary" -Level Info
        }
    } catch
    {
        Write-Status "Could not load remote state: $($_.Exception.Message)" -Level Warning
    }
}

if (-not $LogPath)
{ $LogPath = Join-Path $script:LogDir ("update-everything-{0}.log" -f $script:RunId)
}
if (-not $JsonSummaryPath)
{ $JsonSummaryPath = $script:DefaultJsonSummaryPath
}

$script:LogPath = $LogPath
$script:JsonSummaryPath = $JsonSummaryPath

$script:Config = [ordered]@{
    FastModeSkip       = @(
        'chocolatey', 'wsl-distros', 'npm', 'pnpm', 'yarn', 'bun', 'deno',
        'rustup', 'cargo', 'go', 'pip', 'pip-health', 'pipx', 'uv', 'uv-tools',
        'poetry', 'composer', 'ruby-gems', 'flutter', 'juliaup',
        'oh-my-posh', 'yt-dlp', 'volta', 'fnm', 'dotnet-tools',
        'dotnet-workloads', 'vscode-extensions', 'powershell-modules',
        'powershell-help', 'uv-python', 'ollama-models',
        'vcpkg', 'conda', 'gcloud', 'az', 'aws', 'terraform', 'pulumi',
        'kubectl', 'helm', 'hugo', 'opentofu', 'starship', 'zoxide',
        'gitleaks', 'trivy', 'packer', 'nvm',
        'devcontainer', 'cross-manager'
    )
    UltraFastSkip      = @('windows-update', 'store-apps', 'wsl', 'wsl-distros', 'defender', 'cleanup')
    SkipManagers       = @()
    WingetSkipPackages = @()
    WingetProtectedPackages = @()
    ExtraWingetProtectedPackages = @()
    ChocolateySkipPackages = @()
    ChocolateyProtectedPackages = @()
    ExtraChocolateyProtectedPackages = @()
    StoreAppSkipPackages = @()
    StoreAppProtectedPackages = @()
    ExtraStoreAppProtectedPackages = @()
    PipSkipPackages    = @()
    PipIgnoreHealthPackages = @()
    NpmSkipPackages    = @()
    GcloudSkipComponents = @()
    AzSkipExtensions    = @()
    CondaSkipEnvs       = @()
    VcpkgSkipPackages   = @()
    WindowsOptionalFeatures = @()
    CrossManagerFallback = @{}
    BeforeHooks         = @{}
    AfterHooks          = @{}
    ProfileSkipLists    = @{
        work     = @()
        personal = @()
        gaming   = @()
                minimal  = @('vcpkg', 'conda', 'gcloud', 'az', 'aws', 'terraform', 'pulumi',
                         'kubectl', 'helm', 'hugo', 'opentofu', 'starship', 'gitleaks',
                         'trivy', 'packer', 'nvm', 'devcontainer')
    }
    LogRetentionDays   = 14
    TempCleanupDays    = 7
}

function Initialize-ProcessPath
{
    $segments = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($scope in @('Process', 'Machine', 'User'))
    {
        $rawPath = [Environment]::GetEnvironmentVariable('Path', $scope)
        if ([string]::IsNullOrWhiteSpace($rawPath))
        { continue
        }

        foreach ($entry in ($rawPath -split ';'))
        {
            if ([string]::IsNullOrWhiteSpace($entry))
            { continue
            }
            $expanded = [Environment]::ExpandEnvironmentVariables($entry.Trim())
            $normalized = $expanded.TrimEnd('\')
            if ([string]::IsNullOrWhiteSpace($normalized))
            { continue
            }
            if ($seen.Add($normalized))
            { $segments.Add($expanded) | Out-Null
            }
        }
    }

    if ($segments.Count -gt 0)
    {
        $mergedPath = $segments -join [System.IO.Path]::PathSeparator
        [Environment]::SetEnvironmentVariable('Path', $mergedPath, 'Process')
        $env:Path = $mergedPath
    }

    $ghost = @($segments | Where-Object { $_ -and -not (Test-Path -LiteralPath $_ -PathType Container) })
    if ($ghost.Count -gt 0)
    {
        Write-Verbose "PATH contains $($ghost.Count) non-existent director$(if ($ghost.Count -eq 1) { 'y' } else { 'ies' }): $($ghost -join '; ')"
    }
}

function ConvertTo-StringArray
{
    param([AllowNull()]$Value)
    if ($null -eq $Value)
    { return @()
    }
    if ($Value -is [string])
    { return @($Value)
    }
    return @($Value | ForEach-Object { [string]$_ })
}

function ConvertTo-FilterList
{
    param([AllowNull()]$Value)
    return @(ConvertTo-StringArray $Value |
            ForEach-Object { $_ -split ',' } |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ })
}

function Import-UpdateConfig
{
    $configPath = Join-Path $PSScriptRoot 'update-config.json'
    if (-not (Test-Path -LiteralPath $configPath))
    { return
    }

    try
    {
        $configJson = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in $configJson.PSObject.Properties)
        {
            if (-not $script:Config.Contains($property.Name))
            {
                $script:Config[$property.Name] = $property.Value
                continue
            }
            if ($script:Config[$property.Name] -is [array])
            {
                $script:Config[$property.Name] = @(ConvertTo-FilterList $property.Value)
            } else
            {
                $script:Config[$property.Name] = $property.Value
            }
        }
        Write-Verbose "Loaded config from $configPath"
    } catch
    {
        Write-Warning "Failed to load update-config.json: $($_.Exception.Message)"
    }
}

function Set-RunStorageRoot
{
    param([Parameter(Mandatory)][string]$Root)

    $script:StateDir = $Root
    $script:LogDir = Join-Path $script:StateDir 'logs'
    $script:DefaultJsonSummaryPath = Join-Path $script:StateDir 'last-run.json'
    $script:PreviousJsonSummaryPath = Join-Path $script:StateDir 'previous-run.json'

    if (-not $script:LogPathWasProvided)
    { $script:LogPath = Join-Path $script:LogDir ("update-everything-{0}.log" -f $script:RunId)
    }
    if (-not $script:JsonSummaryPathWasProvided)
    { $script:JsonSummaryPath = $script:DefaultJsonSummaryPath
    }

    Set-Variable -Name LogPath -Scope Script -Value $script:LogPath
    Set-Variable -Name JsonSummaryPath -Scope Script -Value $script:JsonSummaryPath
}

function Test-WritableDirectory
{
    param([Parameter(Mandatory)][string]$Path)
    try
    {
        if (-not (Test-Path -LiteralPath $Path))
        {
            New-Item -Path $Path -ItemType Directory -Force -WhatIf:$false -ErrorAction Stop | Out-Null
        }
        $probePath = Join-Path $Path ("write-test-{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
        Set-Content -LiteralPath $probePath -Value 'ok' -Encoding utf8 -WhatIf:$false -ErrorAction Stop
        Remove-Item -LiteralPath $probePath -Force -WhatIf:$false -ErrorAction SilentlyContinue
        return $true
    } catch
    {
        return $false
    }
}

function Initialize-RunStorage
{
    $candidateRoots = [System.Collections.Generic.List[string]]::new()
    [void]$candidateRoots.Add($script:StateDir)
    if (-not $script:StateDirWasProvided)
    {
        [void]$candidateRoots.Add((Join-Path ([System.IO.Path]::GetTempPath()) 'Update-Everything'))
    }

    $storageReady = $false
    foreach ($root in ($candidateRoots | Select-Object -Unique))
    {
        if ([string]::IsNullOrWhiteSpace($root))
        { continue
        }
        Set-RunStorageRoot -Root $root
        if ((Test-WritableDirectory -Path $script:StateDir) -and (Test-WritableDirectory -Path $script:LogDir))
        {
            $storageReady = $true
            break
        }
    }
    if (-not $storageReady)
    { throw "No writable state/log directory is available. Last tried: $($script:StateDir)"
    }

    try
    {
        if (Test-Path -LiteralPath $script:DefaultJsonSummaryPath)
        {
            Copy-Item -LiteralPath $script:DefaultJsonSummaryPath -Destination $script:PreviousJsonSummaryPath -Force -WhatIf:$false -ErrorAction Stop
        }
    } catch
    {
        Write-Verbose "Could not rotate previous run summary: $($_.Exception.Message)"
    }

    $retentionDays = [int]$script:Config.LogRetentionDays
    if ($retentionDays -gt 0)
    {
        try
        {
            $cutoff = (Get-Date).AddDays(-$retentionDays)
            Get-ChildItem -LiteralPath $script:LogDir -Filter 'update-everything-*.log' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue
        } catch
        {
            Write-Verbose "Log retention cleanup skipped: $($_.Exception.Message)"
        }
    }
}

function Enter-ProcessLock
{
    $lockPath = Join-Path $script:StateDir 'update-everything.lock'
    try
    {
        if (Test-Path -LiteralPath $lockPath)
        {
            $existing = Get-Content -LiteralPath $lockPath -ErrorAction SilentlyContinue
            $existingPid = [int]::Parse(($existing -split "`n")[0].Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
            if (Get-Process -Id $existingPid -ErrorAction SilentlyContinue)
            {
                Write-Warning "Another instance (PID $existingPid) is already running. Exiting."
                exit 5
            }
            Remove-Item -LiteralPath $lockPath -Force -WhatIf:$false -ErrorAction SilentlyContinue
        }
        Set-Content -LiteralPath $lockPath -Value "$PID`n$($script:RunId)" -Encoding utf8 -WhatIf:$false -ErrorAction Stop
        $script:LockPath = $lockPath
    } catch
    {
        Write-Verbose "Could not acquire process lock: $($_.Exception.Message)"
    }
}

function Remove-ProcessLock
{
    if ($script:LockPath -and (Test-Path -LiteralPath $script:LockPath))
    {
        Remove-Item -LiteralPath $script:LockPath -Force -WhatIf:$false -ErrorAction SilentlyContinue
    }
}

function Send-UpdateNotification
{
    param(
        [Parameter(Mandatory)][int]$Succeeded,
        [Parameter(Mandatory)][int]$Failed,
        [Parameter(Mandatory)][int]$Skipped,
        [Parameter(Mandatory)][string]$Elapsed
    )

    $title = if ($Failed -gt 0) { 'Update-Everything: completed with failures' } else { 'Update-Everything: done' }
    $body = "$Succeeded succeeded, $Failed failed, $Skipped skipped in $Elapsed"

    try
    {
        if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue)
        {
            New-BurntToastNotification -Text $title, $body -ErrorAction SilentlyContinue
            return
        }
    } catch { }

    try
    {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $balloon = [System.Windows.Forms.NotifyIcon]::new()
        $balloon.Icon = [System.Drawing.SystemIcons]::Information
        $balloon.Visible = $true
        $balloon.ShowBalloonTip(5000, $title, $body, [System.Windows.Forms.ToolTipIcon]::Info)
        Start-Sleep -Milliseconds 5100
        $balloon.Dispose()
    } catch { }
}

function Show-ResultTable
{
    param(
        [object[]]$Results,
        [object[]]$Skipped
    )

    $statusColor = @{
        Succeeded = 'Green'
        Warn      = 'Yellow'
        Partial   = 'Yellow'
        Failed    = 'Red'
        TimedOut  = 'Red'
        DryRun    = 'DarkGray'
        Skipped   = 'DarkGray'
    }

    Write-Host ''
    Write-Host ("{0,-24} {1,-12} {2,-8} {3}" -f 'Task', 'Status', 'Time(s)', 'Note') -ForegroundColor DarkGray
    Write-Host ("{0,-24} {1,-12} {2,-8} {3}" -f '----', '------', '-------', '----') -ForegroundColor DarkGray

    # Skipped rows are counted but not listed; they drown out tasks that ran.
    foreach ($r in @($Results | Where-Object { $_.Status -ne 'Skipped' }))
    {
        $color = if ($statusColor.ContainsKey($r.Status)) { $statusColor[$r.Status] } else { 'White' }
        $dur = if ($r.PSObject.Properties['DurationSeconds'] -and $r.DurationSeconds -gt 0) { '{0:N1}' -f $r.DurationSeconds } else { '-' }
        $note = if ($r.PSObject.Properties['Reason'] -and $r.Reason) { $r.Reason } else { '' }
        Write-Host ("{0,-24} " -f $r.Name) -NoNewline -ForegroundColor White
        Write-Host ("{0,-12} " -f $r.Status) -NoNewline -ForegroundColor $color
        Write-Host ("{0,-8} " -f $dur) -NoNewline -ForegroundColor DarkGray
        Write-Host $note -ForegroundColor DarkGray
    }

    foreach ($s in @($Skipped | Where-Object { $_.PSObject.Properties['Reason'] -and $_.Reason -notmatch 'optional|missing command' }))
    {
        Write-Host ("{0,-24} {1,-12} {2,-8} {3}" -f $s.Name, 'Skipped', '-', $s.Reason) -ForegroundColor DarkGray
    }

    Write-Host ''
}

function Write-UpdateLog
{
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Muted')]
        [string]$Level = 'Info'
    )
    $line = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level.ToUpperInvariant(), $Message
    try
    {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8 -WhatIf:$false -ErrorAction Stop
    } catch
    {
        if (-not $script:LogWriteWarningEmitted)
        {
            $script:LogWriteWarningEmitted = $true
            Write-Warning "Logging disabled: $($_.Exception.Message)"
        }
    }
}

function Write-Status
{
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Muted')]
        [string]$Level = 'Info'
    )

    Write-UpdateLog -Message $Message -Level $Level
    if ($Quiet -and $Level -notin @('Warning', 'Error'))
    { return
    }

    $color = switch ($Level)
    {
        'Success'
        { 'Green'
        }
        'Warning'
        { 'Yellow'
        }
        'Error'
        { 'Red'
        }
        'Muted'
        { 'DarkGray'
        }
        default
        { 'Cyan'
        }
    }
    Write-Host $Message -ForegroundColor $color
}

function Test-IsAdmin
{
    if (-not $IsWindows)
    { return $false
    }
    try
    {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = [Security.Principal.WindowsPrincipal]::new($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch
    { return $false
    }
}

function Get-VSCodeCommandPath
{
    $cmd = Get-Command -Name 'code' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($cmd)
    {
        if ($cmd.Source)
        { return $cmd.Source
        }
        if ($cmd.Path)
        { return $cmd.Path
        }
    }

    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\bin\code.cmd'),
        'C:\Program Files\Microsoft VS Code\bin\code.cmd',
        'C:\Program Files (x86)\Microsoft VS Code\bin\code.cmd'
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

    return ($candidates | Select-Object -First 1)
}

function Get-PreferredCommandPath
{
    param([Parameter(Mandatory)][string]$Name)

    if ($Name -in @('pwsh', 'pwsh.exe'))
    {
        $pwshFileName = if ($IsWindows)
        { 'pwsh.exe'
        } else
        { 'pwsh'
        }
        $candidates = @(
            (Join-Path $PSHOME $pwshFileName),
            'C:\Program Files\PowerShell\7\pwsh.exe',
            'C:\Program Files\PowerShell\7-preview\pwsh.exe'
        ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
        if ($candidates.Count -gt 0)
        { return ($candidates | Select-Object -First 1)
        }
    }

    if ($Name -eq 'code')
    { return Get-VSCodeCommandPath
    }

    if ($Name -eq 'python')
    {
        $candidates = @(@(
            'C:\Program Files\Python314\python.exe',
            'C:\Program Files\Python313\python.exe',
            'C:\Program Files\Python312\python.exe',
            (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python314\python.exe'),
            (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'),
            (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe')
        ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) })

        if ($candidates.Count -gt 0)
        { return ($candidates | Select-Object -First 1)
        }
    }

    return $null
}

function Test-Command
{
    param([Parameter(Mandatory)][string]$Name)
    if ($script:CommandCache.ContainsKey($Name))
    { return $script:CommandCache[$Name]
    }

    $found = $false
    $preferred = Get-PreferredCommandPath -Name $Name
    if ($preferred)
    {
        $found = $true
    } elseif ($Name -eq 'code')
    {
        $found = -not [string]::IsNullOrWhiteSpace((Get-VSCodeCommandPath))
    } else
    {
        $found = [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
    }
    $script:CommandCache[$Name] = $found
    return $found
}

function Get-CommandPath
{
    param([Parameter(Mandatory)][string]$Name)
    $preferred = Get-PreferredCommandPath -Name $Name
    if ($preferred)
    { return $preferred
    }

    $command = Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command)
    {
        if ($command.Source)
        { return $command.Source
        }
        if ($command.Path)
        { return $command.Path
        }
    }
    return $null
}

function ConvertTo-TaskId
{
    param([Parameter(Mandatory)][string]$Value)
    return (($Value.Trim().ToLowerInvariant()) -replace '[^a-z0-9]+', '-').Trim('-')
}

function Test-NameMatch
{
    param(
        [Parameter(Mandatory)]$Task,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns)
    {
        if ([string]::IsNullOrWhiteSpace($pattern))
        { continue
        }
        $rawPattern = $pattern.Trim()
        $needle = ConvertTo-TaskId $rawPattern
        $taskName = ConvertTo-TaskId $Task.Name

        if ($Task.Id -eq $needle -or $taskName -eq $needle)
        { return $true
        }
        if ($Task.Category -and ((ConvertTo-TaskId $Task.Category) -eq $needle))
        { return $true
        }
        if ($Task.Tags -contains $needle)
        { return $true
        }
        if ($rawPattern -match '[*?]' -and $Task.Name -like $rawPattern)
        { return $true
        }
    }
    return $false
}

function New-UpdateTask
{
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][scriptblock]$Script,
        [string[]]$RequiresCommand = @(),
        [string[]]$Tags = @(),
        [string[]]$Resources = @(),
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
        Resources       = @($Resources | ForEach-Object { ConvertTo-TaskId $_ })
        RequiresAdmin   = [bool]$RequiresAdmin
        Disabled        = [bool]$Disabled
        DisabledReason  = $DisabledReason
        TimeoutSec      = $TimeoutSec
    }
}

function Join-QuotedList
{
    param([string[]]$Values)
    return (@($Values) | ForEach-Object { "'$($_.Replace("'", "''"))'" }) -join ', '
}

function ConvertFrom-WingetUpgradeOutput
{
    param([AllowNull()]$Output)

    if ($null -eq $Output)
    { return @()
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($Output))
    {
        if ($null -eq $item)
        { continue
        }
        foreach ($line in (([string]$item -replace '\r(?!\n)', "`n") -split "\r?\n"))
        { [void]$lines.Add($line)
        }
    }

    $packages = [System.Collections.Generic.List[object]]::new()
    $seenIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $ansiPattern = '\x1b\[[0-9;?]*[a-zA-Z]'
    $oscPattern = '\x1b\][^\a]*(\a|\x1b\\)'
    $idPattern = '^(?:[A-Za-z][A-Za-z0-9_+\-]{2,}\.[A-Za-z0-9][A-Za-z0-9._+\-]+|(?-i:[A-Z0-9]{8,}))$'

    $columnStarts = $null

    function Get-WingetColumnValue
    {
        param(
            [Parameter(Mandatory)][string]$Text,
            [Parameter(Mandatory)][int]$Start,
            [int]$End = -1
        )

        if ($Start -lt 0 -or $Text.Length -le $Start)
        { return $null
        }

        if ($End -gt $Start)
        {
            $width = [Math]::Min($End - $Start, $Text.Length - $Start)
            return $Text.Substring($Start, $width).Trim()
        }

        return $Text.Substring($Start).Trim()
    }

    function New-WingetPackageEntry
    {
        param(
            [string]$Name,
            [string]$Id,
            [string]$Version,
            [string]$Available,
            [string]$Source,
            [string]$DisplayVersion
        )

        if ([string]::IsNullOrWhiteSpace($Id) -or $Id -notmatch $idPattern)
        { return $null
        }
        if ($Id -match '^(Name|Id|Version|Available|Source|-)$')
        { return $null
        }
        if (-not $seenIds.Add($Id))
        { return $null
        }

        $isUnknown = $Version -eq 'Unknown'
        [pscustomobject]@{
            Name                 = $Name
            Id                   = $Id
            Version              = $Version
            Available            = $Available
            Source               = $Source
            IsUnknown            = $isUnknown
            DisplayVersion       = $DisplayVersion
            InstalledLooksCurrent = ($isUnknown -and $DisplayVersion -and $Available -and $DisplayVersion -eq $Available)
        }
    }

    function Get-WingetMatchValue
    {
        param(
            [Parameter(Mandatory)]$Match,
            [Parameter(Mandatory)][string]$Name
        )

        if ($Match.ContainsKey($Name))
        { return $Match[$Name]
        }
        return $null
    }

    foreach ($rawLine in $lines)
    {
        $text = ([regex]::Replace([string]$rawLine, $oscPattern, ''))
        $text = ([regex]::Replace($text, $ansiPattern, '')).Replace([string][char]0, '').TrimEnd()
        if ([string]::IsNullOrWhiteSpace($text))
        { continue
        }
        if ($text -match 'No installed package found matching input criteria|No available upgrade found|No applicable update found|No available updates')
        { continue
        }
        if ($text -match '^\s*(Found\s|This application is licensed|Microsoft is not responsible|Downloading |Successfully |Starting package |Extracting archive|Installer log |\d+\s+upgrades? available|The following packages)')
        { continue
        }

        if ($text -match '^\s*Name\s+Id\s+Version\s+Available(?:\s+Source)?')
        {
            $columnStarts = [ordered]@{
                Name      = $text.IndexOf('Name')
                Id        = $text.IndexOf('Id')
                Version   = $text.IndexOf('Version')
                Available = $text.IndexOf('Available')
                Source    = $text.IndexOf('Source')
            }
            continue
        }
        if ($text -match '^\s*-{3,}')
        { continue
        }

        $entry = $null
        if ($columnStarts -and $columnStarts.Id -ge 0 -and $columnStarts.Version -gt $columnStarts.Id -and $columnStarts.Available -gt $columnStarts.Version)
        {
            $sourceStart = if ($columnStarts.Source -gt $columnStarts.Available)
            { $columnStarts.Source
            } else
            { -1
            }
            $name = Get-WingetColumnValue -Text $text -Start $columnStarts.Name -End $columnStarts.Id
            $id = Get-WingetColumnValue -Text $text -Start $columnStarts.Id -End $columnStarts.Version
            $version = Get-WingetColumnValue -Text $text -Start $columnStarts.Version -End $columnStarts.Available
            $available = Get-WingetColumnValue -Text $text -Start $columnStarts.Available -End $sourceStart
            $source = if ($sourceStart -ge 0)
            { Get-WingetColumnValue -Text $text -Start $sourceStart
            } else
            { $null
            }
            $entry = New-WingetPackageEntry -Name $name -Id $id -Version $version -Available $available -Source $source
        }

        if (-not $entry)
        {
            $idToken = '(?=\S*[A-Za-z])[A-Za-z0-9][A-Za-z0-9._+\-]+'
            if ($text -match "^(?<name>.+?)\s+(?<displayVersion>\d+(?:\.\d+)+(?:[-+][^\s]+)?)\s+(?<id>$idToken)\s+Unknown\s+(?<available>\S+)(?:\s+(?<source>\S+))?\s*$")
            {
                $entry = New-WingetPackageEntry -Name $matches.name.Trim() -Id $matches.id.Trim() -Version 'Unknown' -Available $matches.available.Trim() -Source (Get-WingetMatchValue -Match $matches -Name 'source') -DisplayVersion $matches.displayVersion.Trim()
            } elseif ($text -match "^(?<name>.+?)\s+(?<id>$idToken)\s+(?<version>Unknown|\S+)\s+(?<available>\S+)(?:\s+(?<source>\S+))?\s*$")
            {
                $entry = New-WingetPackageEntry -Name $matches.name.Trim() -Id $matches.id.Trim() -Version $matches.version.Trim() -Available $matches.available.Trim() -Source (Get-WingetMatchValue -Match $matches -Name 'source')
            }
        }

        if ($entry)
        { [void]$packages.Add($entry)
        }
    }

    return @($packages)
}

function Get-UpdateTasks
{
    $wingetSkip = ConvertTo-StringArray $script:Config.WingetSkipPackages
    $wingetProtected = @()
    $wingetProtected += ConvertTo-FilterList $script:Config.WingetProtectedPackages
    $wingetProtected += ConvertTo-FilterList $script:Config.ExtraWingetProtectedPackages
    $chocoSkip = ConvertTo-FilterList $script:Config.ChocolateySkipPackages
    $chocoProtected = @()
    $chocoProtected += ConvertTo-FilterList $script:Config.ChocolateyProtectedPackages
    $chocoProtected += ConvertTo-FilterList $script:Config.ExtraChocolateyProtectedPackages
    $storeSkip = ConvertTo-StringArray $script:Config.StoreAppSkipPackages
    $storeProtected = @()
    $storeProtected += ConvertTo-FilterList $script:Config.StoreAppProtectedPackages
    $storeProtected += ConvertTo-FilterList $script:Config.ExtraStoreAppProtectedPackages
    $pipSkip = ConvertTo-StringArray $script:Config.PipSkipPackages
    $pipIgnoreHealth = ConvertTo-FilterList $script:Config.PipIgnoreHealthPackages
    $npmSkip = ConvertTo-FilterList $script:Config.NpmSkipPackages
    $gcloudSkipComponents = ConvertTo-FilterList $script:Config.GcloudSkipComponents
    $azSkipExtensions = ConvertTo-FilterList $script:Config.AzSkipExtensions
    $condaSkipEnvs = ConvertTo-FilterList $script:Config.CondaSkipEnvs
    $vcpkgSkipPackages = ConvertTo-StringArray $script:Config.VcpkgSkipPackages
    $tempCleanupDays = [int]$script:Config.TempCleanupDays
    $useSilentInstallers = -not [bool]$NoSilentInstallers
    $windowsOptionalFeatures = ConvertTo-FilterList $script:Config.WindowsOptionalFeatures
    $crossManagerFallback = $script:Config.CrossManagerFallback
    $script:BeforeHooks = $script:Config.BeforeHooks
    $script:AfterHooks = $script:Config.AfterHooks
    $tasks = [System.Collections.Generic.List[object]]::new()

    $tasks.Add((New-UpdateTask -Name 'winget-source' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec 300 -Script {
                Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('source', 'update') -Retries 1 -TimeoutSec 300
            } -Tags @('windows') -Resources @('winget'))) | Out-Null

    $wingetScript = {
        param(
            [string[]]$SkipPackages,
            [string[]]$ProtectedPackages,
            [bool]$BypassProtection,
            [bool]$UseSilentInstallers,
            [string]$UnknownVersionStatePath
        )

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($package in @($SkipPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($package))
            { [void]$skipSet.Add($package.Trim())
            }
        }

        $protectedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($package in @($ProtectedPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($package))
            { [void]$protectedSet.Add($package.Trim())
            }
        }

        if ($skipSet.Count -gt 0)
        { Write-Output "Config skip list ($($skipSet.Count)): $($skipSet -join ', ')"
        }
        if ($protectedSet.Count -gt 0 -and -not $BypassProtection)
        { Write-Output "Protected app list ($($protectedSet.Count)): $($protectedSet -join ', ')"
        }
        if (-not $UseSilentInstallers)
        { Write-Output 'winget installer mode: standard (no --silent).'
        }

        $unknownState = @{}
        if ($UnknownVersionStatePath -and (Test-Path -LiteralPath $UnknownVersionStatePath))
        {
            try
            {
                $loadedUnknownState = Get-Content -LiteralPath $UnknownVersionStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
                if ($loadedUnknownState)
                { $unknownState = $loadedUnknownState
                }
            } catch
            {
                Write-Output "Could not read winget unknown-version state: $($_.Exception.Message)"
                $unknownState = @{}
            }
        }

        Write-Output 'Scanning for available winget upgrades from the winget source...'
        $listOutput = @(Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--source', 'winget', '--include-unknown', '--include-pinned', '--accept-source-agreements', '--disable-interactivity') -SuccessExitCodes @(0, -1978335189))
        $candidates = @(ConvertFrom-WingetUpgradeOutput -Output $listOutput)

        $toUpgrade = [System.Collections.Generic.List[object]]::new()
        $toSkipIds = [System.Collections.Generic.List[string]]::new()
        $toProtectIds = [System.Collections.Generic.List[string]]::new()
        $currentUnknownIds = [System.Collections.Generic.List[string]]::new()
        $unchangedUnknownIds = [System.Collections.Generic.List[string]]::new()

        foreach ($pkg in ($candidates | Sort-Object Id -Unique))
        {
            $id = $pkg.Id
            if ($skipSet.Contains($id))
            { [void]$toSkipIds.Add($id)
            } elseif ((-not $BypassProtection) -and $protectedSet.Contains($id))
            { [void]$toProtectIds.Add($id)
            } elseif ($pkg.IsUnknown -and $pkg.InstalledLooksCurrent)
            { [void]$currentUnknownIds.Add("$id ($($pkg.Available))")
            } elseif ($pkg.IsUnknown -and $unknownState.ContainsKey($id) -and ([string]$unknownState[$id]) -eq ([string]$pkg.Available))
            { [void]$unchangedUnknownIds.Add("$id ($($pkg.Available))")
            } else
            { [void]$toUpgrade.Add($pkg)
            }
        }

        if (($toUpgrade.Count + $toSkipIds.Count + $toProtectIds.Count + $currentUnknownIds.Count + $unchangedUnknownIds.Count) -eq 0)
        {
            Write-Output 'No winget upgrades available.'
            return
        }

        $suppressedCount = $toSkipIds.Count + $toProtectIds.Count + $currentUnknownIds.Count + $unchangedUnknownIds.Count
        Write-Output ("Found {0} winget package(s): {1} to upgrade, {2} suppressed/already current." -f ($toUpgrade.Count + $suppressedCount), $toUpgrade.Count, $suppressedCount)
        if ($toSkipIds.Count -gt 0)
        { Write-Output "  Config-suppressed: $($toSkipIds -join ', ')"
        }
        if ($toProtectIds.Count -gt 0)
        { Write-Output "  Protected-suppressed: $($toProtectIds -join ', ')"
        }
        if ($currentUnknownIds.Count -gt 0)
        { Write-Output "  Unknown-version already current: $($currentUnknownIds -join ', ')"
        }
        if ($unchangedUnknownIds.Count -gt 0)
        { Write-Output "  Unknown-version unchanged since last successful update: $($unchangedUnknownIds -join ', ')"
        }
        if ($toUpgrade.Count -eq 0)
        { Write-Output 'All available winget upgrades are suppressed, unchanged, or already current.'; return
        }

        $failed = [System.Collections.Generic.List[string]]::new()
        $unknownStateChanged = $false
        $upgradeIndex = 0
        foreach ($pkg in $toUpgrade)
        {
            $upgradeIndex++
            $id = $pkg.Id
            $unknownText = if ($pkg.IsUnknown)
            { " (installed version unknown -> $($pkg.Available))"
            } else
            { ''
            }
            Write-Output ("[{0}/{1}] Upgrading: {2}{3}" -f $upgradeIndex, $toUpgrade.Count, $id, $unknownText)
            try
            {
                $wingetArgs = @('upgrade', '--id', $id, '--exact', '--source', 'winget', '--include-unknown', '--include-pinned', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements')
                if ($UseSilentInstallers)
                { $wingetArgs += '--silent'
                }
                $upgradeOutput = @(Invoke-UpdateProcess -FilePath 'winget' -ArgumentList $wingetArgs -Retries 1 -TimeoutSec 300)
                if ($pkg.IsUnknown)
                {
                    $upgradeText = ($upgradeOutput | Out-String)
                    if ($upgradeText -notmatch 'No applicable upgrade found|No installed package found matching input criteria')
                    {
                        $unknownState[$id] = [string]$pkg.Available
                        $unknownStateChanged = $true
                    }
                }
            } catch
            {
                $errMsg = $_.Exception.Message
                if ($errMsg -match '-1978335189|No applicable upgrade found')
                {
                    Write-Output "  Not applicable: $id (newer version exists but is not compatible with this system/architecture)"
                } else
                {
                    Write-Output "  FAILED: $errMsg"
                    [void]$failed.Add($id)
                }
            }
        }

        if ($unknownStateChanged -and $UnknownVersionStatePath)
        {
            try
            {
                $stateParent = Split-Path -Parent $UnknownVersionStatePath
                if ($stateParent -and -not (Test-Path -LiteralPath $stateParent))
                { New-Item -Path $stateParent -ItemType Directory -Force -WhatIf:$false -ErrorAction Stop | Out-Null
                }
                $unknownState | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $UnknownVersionStatePath -Encoding utf8 -WhatIf:$false -ErrorAction Stop
            } catch
            {
                Write-Output "Could not write winget unknown-version state: $($_.Exception.Message)"
            }
        }

        if ($failed.Count -gt 0)
        { Write-Output "winget left unchanged: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'winget' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec $WingetTimeoutSec -Script $wingetScript -Tags @('windows') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'winget-batch' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec 600 -Script {
                Write-Output 'Attempting fast batch winget upgrade (falls back to per-package on failure)...'
                $result = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--all', '--source', 'winget', '--include-unknown', '--accept-source-agreements', '--disable-interactivity') -Retries 0 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                $outText = (@($result.Output) | Out-String).Trim()
                if ($outText) { Write-Output $outText }
                if ($result.ExitCode -ne 0 -and $result.ExitCode -ne -1978335189 -and $result.ExitCode -ne -1978335212)
                { Write-Output "Batch upgrade skipped (code $($result.ExitCode)); winget task handles per-package upgrades." }
                else { Write-Output "Batch winget upgrade completed." }
            } -Tags @('windows') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'cross-manager' -Category 'package-manager' -Script {
                param(
                    [hashtable]$FallbackApps
                )
                if ($FallbackApps.Count -eq 0)
                {
                    Write-Output 'No cross-manager fallback apps configured.'
                    return
                }
                $chocoAvailable = $null -ne (Get-Command 'choco' -ErrorAction SilentlyContinue)
                $scoopAvailable = $null -ne (Get-Command 'scoop' -ErrorAction SilentlyContinue)
                if (-not $chocoAvailable -and -not $scoopAvailable)
                {
                    Write-Output 'No alternate package managers (choco/scoop) available for fallback.'
                    return
                }
                foreach ($wingetId in $FallbackApps.Keys)
                {
                    $alt = $FallbackApps[$wingetId]
                    Write-Output "Fallback check: $wingetId"
                    if ($chocoAvailable -and $alt.ContainsKey('choco') -and $alt.choco)
                    {
                        Write-Output "  Chocolatey ($($alt.choco))..."
                        try
                        {
                            $result = Invoke-UpdateProcess -FilePath 'choco' -ArgumentList @('upgrade', $alt.choco, '-y', '--no-progress') -TimeoutSec 120 -Retries 0 -SuccessExitCodes @(0, 1) -PassThru
                            $output = (@($result.Output) | Out-String).Trim()
                            if ($output) { Write-Output ($output -split "`n" | ForEach-Object { "    $_" }) }
                        } catch
                        {
                            Write-Output "    Chocolatey failed: $($_.Exception.Message)"
                        }
                    }
                    if ($scoopAvailable -and $alt.ContainsKey('scoop') -and $alt.scoop)
                    {
                        Write-Output "  Scoop ($($alt.scoop))..."
                        try
                        {
                            $env:GIT_TERMINAL_PROMPT = '0'
                            $result = Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update', $alt.scoop) -TimeoutSec 120 -Retries 0 -SuccessExitCodes @(0, 1) -PassThru
                            $output = (@($result.Output) | Out-String).Trim()
                            if ($output) { Write-Output ($output -split "`n" | ForEach-Object { "    $_" }) }
                        } catch
                        {
                            Write-Output "    Scoop failed: $($_.Exception.Message)"
                        }
                    }
                }
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'scoop' -Category 'package-manager' -RequiresCommand 'scoop' -Script {
                $env:GIT_TERMINAL_PROMPT = '0'
                try
                {
                    Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update') -Retries 0 -TimeoutSec 300
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "scoop update skipped: $($_.Exception.Message)"
                    Write-Output "scoop update skipped: $($_.Exception.Message)"
                    return
                }
                try
                {
                    Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update', '*') -Retries 0 -TimeoutSec 600
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "scoop update * skipped: $($_.Exception.Message)"
                    Write-Output "scoop update * skipped: $($_.Exception.Message)"
                }
            } -Tags @('windows'))) | Out-Null

    # MSYS2 (native Windows): winget refuses to upgrade it ("cannot be upgraded
    # using winget"); its own pacman is the supported path.
    $tasks.Add((New-UpdateTask -Name 'msys2' -Category 'package-manager' -Script {
                $bash = @(
                    'C:\msys64\usr\bin\bash.exe',
                    'C:\tools\msys64\usr\bin\bash.exe',
                    "$env:SystemDrive\msys64\usr\bin\bash.exe"
                ) | Where-Object { Test-Path $_ } | Select-Object -First 1
                if (-not $bash)
                {
                    Set-TaskStatus -Status 'Skipped' -Reason 'MSYS2 not installed'
                    Write-Output 'MSYS2 not installed; skipping.'
                    return
                }
                Write-Output "Updating MSYS2 via $bash"
                Invoke-UpdateProcess -FilePath $bash -ArgumentList @('-lc', 'pacman -Syu --noconfirm') -Retries 0 -TimeoutSec 1800
            } -Tags @('windows', 'msys2') -Resources @('msys2'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'chocolatey' -Category 'package-manager' -RequiresCommand 'choco' -RequiresAdmin -Script {
                param(
                    [string[]]$SkipPackages,
                    [string[]]$ProtectedPackages,
                    [bool]$BypassProtection
                )

                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($package in @($SkipPackages))
                {
                    if (-not [string]::IsNullOrWhiteSpace($package))
                    { [void]$skipSet.Add($package.Trim())
                    }
                }

                $protectedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($package in @($ProtectedPackages))
                {
                    if (-not [string]::IsNullOrWhiteSpace($package))
                    { [void]$protectedSet.Add($package.Trim())
                    }
                }

                $outdatedOutput = @(Invoke-UpdateProcess -FilePath 'choco' -ArgumentList @('outdated', '-r', '--limit-output') -SuccessExitCodes @(0, 2))
                $outdatedPackages = @($outdatedOutput |
                        ForEach-Object { ([string]$_).Trim() } |
                        Where-Object { $_ -match '^[A-Za-z0-9][A-Za-z0-9._-]*\|[^|]+\|[^|]+\|' })

                if ($outdatedPackages.Count -eq 0)
                {
                    Write-Output 'No outdated Chocolatey packages found.'
                    return
                }

                $entries = @($outdatedPackages | ForEach-Object {
                        $parts = ([string]$_) -split '\|'
                        if ($parts.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($parts[0]))
                        { [pscustomobject]@{ Name = $parts[0].Trim(); Raw = $_ }
                        }
                    })

                $toUpgrade = [System.Collections.Generic.List[string]]::new()
                $toSkip = [System.Collections.Generic.List[string]]::new()
                $toProtect = [System.Collections.Generic.List[string]]::new()
                foreach ($entry in ($entries | Sort-Object Name -Unique))
                {
                    if ($skipSet.Contains($entry.Name))
                    { [void]$toSkip.Add($entry.Name)
                    } elseif ((-not $BypassProtection) -and $protectedSet.Contains($entry.Name))
                    { [void]$toProtect.Add($entry.Name)
                    } else
                    { [void]$toUpgrade.Add($entry.Name)
                    }
                }

                $suppressedCount = $toSkip.Count + $toProtect.Count
                Write-Output ("Found {0} outdated Chocolatey package(s): {1} to upgrade, {2} suppressed." -f ($toUpgrade.Count + $suppressedCount), $toUpgrade.Count, $suppressedCount)
                if ($toSkip.Count -gt 0)
                { Write-Output "  Config-suppressed: $($toSkip -join ', ')"
                }
                if ($toProtect.Count -gt 0)
                { Write-Output "  Protected-suppressed: $($toProtect -join ', ')"
                }
                if ($toUpgrade.Count -eq 0)
                { Write-Output 'All outdated Chocolatey packages are suppressed by skip/protected lists.'; return
                }

                $failed = [System.Collections.Generic.List[string]]::new()
                foreach ($package in $toUpgrade)
                {
                    Write-Output "Updating Chocolatey package: $package"
                    try
                    {
                        Invoke-UpdateProcess -FilePath 'choco' -ArgumentList @('upgrade', $package, '-y', '--no-progress', '--limit-output') -Retries 1
                    } catch
                    {
                        Write-Output "Chocolatey package left unchanged: $package — $($_.Exception.Message)"
                        [void]$failed.Add($package)
                    }
                }

                if ($failed.Count -gt 0)
                { Write-Output "Chocolatey packages left unchanged: $($failed -join ', ')"
                }
            } -Tags @('windows') -Resources @('chocolatey'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'store-apps' -Category 'system' -RequiresCommand 'winget' -Disabled:$SkipStoreApps -DisabledReason 'disabled by -SkipStoreApps' -TimeoutSec $WingetTimeoutSec -Script {
                param(
                    [string[]]$SkipPackages,
                    [string[]]$ProtectedPackages,
                    [bool]$BypassProtection,
                    [bool]$UseSilentInstallers
                )

                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($package in @($SkipPackages))
                {
                    if (-not [string]::IsNullOrWhiteSpace($package))
                    { [void]$skipSet.Add($package.Trim())
                    }
                }

                $protectedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($package in @($ProtectedPackages))
                {
                    if (-not [string]::IsNullOrWhiteSpace($package))
                    { [void]$protectedSet.Add($package.Trim())
                    }
                }

                Write-Output 'Scanning for Microsoft Store app upgrades...'
                $listOutput = @(Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--source', 'msstore', '--include-unknown', '--include-pinned', '--accept-source-agreements', '--disable-interactivity') -SuccessExitCodes @(0, -1978335189))
                $candidates = @(ConvertFrom-WingetUpgradeOutput -Output $listOutput | Sort-Object Id -Unique)
                if ($candidates.Count -eq 0)
                {
                    Write-Output 'No Microsoft Store app upgrades available.'
                    return
                }

                $toUpgrade = [System.Collections.Generic.List[pscustomobject]]::new()
                $toSkipIds = [System.Collections.Generic.List[string]]::new()
                $toProtectIds = [System.Collections.Generic.List[string]]::new()
                foreach ($pkg in $candidates)
                {
                    $id = $pkg.Id
                    $label = if ($pkg.Name) { "$($pkg.Name) ($id)" } else { $id }
                    if ($skipSet.Contains($id))
                    { [void]$toSkipIds.Add($label)
                    } elseif ((-not $BypassProtection) -and $protectedSet.Contains($id))
                    { [void]$toProtectIds.Add($label)
                    } else
                    { [void]$toUpgrade.Add([pscustomobject]@{ Id = $id; Label = $label })
                    }
                }

                if ($toSkipIds.Count -gt 0)
                { Write-Output "Store skip list ($($toSkipIds.Count)): $($toSkipIds -join ', ')"
                }
                if ($toProtectIds.Count -gt 0)
                { Write-Output "Store protected list ($($toProtectIds.Count)): $($toProtectIds -join ', ')"
                }
                if ($toUpgrade.Count -eq 0)
                {
                    Write-Output 'All Microsoft Store app upgrades are suppressed by Store skip/protected lists.'
                    return
                }

                $failed = [System.Collections.Generic.List[string]]::new()
                $needsManual = [System.Collections.Generic.List[string]]::new()
                foreach ($entry in $toUpgrade)
                {
                    $id = $entry.Id
                    $label = $entry.Label
                    Write-Output "Updating Microsoft Store app: $label"
                    try
                    {
                        $storeArgs = @('upgrade', '--id', $id, '--exact', '--source', 'msstore', '--include-unknown', '--include-pinned', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements')
                        if ($UseSilentInstallers)
                        { $storeArgs += '--silent'
                        }
                        $storeResult = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList $storeArgs -Retries 0 -TimeoutSec $WingetTimeoutSec -PassThru
                        $storeText = (@($storeResult.Output) | Out-String).Trim()
                        if ($storeResult.ExitCode -eq -1978335189 -and $storeText -match 'install technology is different')
                        {
                            Write-Output "  Cannot auto-update: installer type changed. Open the Store and update '$label' manually."
                            [void]$needsManual.Add($label)
                        } elseif ($storeResult.ExitCode -ne 0)
                        {
                            if ($storeText) { Write-Output $storeText }
                            [void]$failed.Add($label)
                        } elseif ($storeText)
                        {
                            Write-Output $storeText
                        }
                    } catch
                    {
                        Write-Output "Store app could not be updated automatically: $label — $($_.Exception.Message)"
                        [void]$failed.Add($label)
                    }
                }

                if ($needsManual.Count -gt 0)
                { Write-Output "Needs manual Store update (installer type changed): $($needsManual -join ', ')"
                }
                if ($failed.Count -gt 0)
                { Write-Output "Microsoft Store left $($failed.Count) app(s) for manual Store update: $($failed -join ', ')"
                }
            } -Tags @('windows', 'store') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'windows-update' -Category 'system' -RequiresAdmin -Disabled:$SkipWindowsUpdate -DisabledReason 'disabled by -SkipWindowsUpdate' -TimeoutSec 7200 -Script {
                if (Get-Command Install-WindowsUpdate -ErrorAction SilentlyContinue)
                {
                    try
                    {
                        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop | Out-String | Write-Output
                        return
                    } catch
                    {
                        Write-Output "PSWindowsUpdate did not complete; using Windows Update Agent COM fallback: $($_.Exception.Message)"
                    }
                }

                Write-Output 'Using Windows Update Agent COM fallback.'

                function Invoke-WuaOperation
                {
                    param(
                        [Parameter(Mandatory)][scriptblock]$Operation,
                        [Parameter(Mandatory)][string]$Name,
                        [int]$Attempts = 3
                    )
                    $lastMessage = $null
                    for ($try = 1; $try -le $Attempts; $try++)
                    {
                        try
                        { return & $Operation
                        } catch
                        {
                            $lastMessage = $_.Exception.Message
                            if ($try -lt $Attempts)
                            {
                                Write-Output "$Name was busy; retrying internally ($try/$Attempts): $lastMessage"
                                Start-Sleep -Seconds ([Math]::Min(15, 3 * $try))
                            }
                        }
                    }
                    throw "$Name did not complete after $Attempts attempt(s): $lastMessage"
                }

                $session = New-Object -ComObject Microsoft.Update.Session
                $session.ClientApplicationID = 'Update-Everything'
                $searcher = $session.CreateUpdateSearcher()
                $result = Invoke-WuaOperation -Name 'Windows Update search' -Operation {
                    $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
                }

                $count = [int]$result.Updates.Count
                Write-Output "Windows Update available updates: $count"
                if ($count -eq 0)
                { return
                }

                $updates = New-Object -ComObject Microsoft.Update.UpdateColl
                for ($i = 0; $i -lt $result.Updates.Count; $i++)
                {
                    $update = $result.Updates.Item($i)
                    Write-Output "Selected update: $($update.Title)"
                    if (-not $update.EulaAccepted)
                    { $update.AcceptEula()
                    }
                    [void]$updates.Add($update)
                }

                $downloader = $session.CreateUpdateDownloader()
                $downloader.Updates = $updates
                $downloadResult = Invoke-WuaOperation -Name 'Windows Update download' -Operation { $downloader.Download() }
                Write-Output "Windows Update download result code: $($downloadResult.ResultCode)"
                if ($downloadResult.ResultCode -notin @(2, 3))
                { throw "Windows Update download result code was $($downloadResult.ResultCode)"
                }

                $installer = $session.CreateUpdateInstaller()
                $installer.Updates = $updates
                $installResult = Invoke-WuaOperation -Name 'Windows Update install' -Operation { $installer.Install() }
                Write-Output "Windows Update install result code: $($installResult.ResultCode); reboot required: $($installResult.RebootRequired)"
                if ($installResult.ResultCode -notin @(2, 3))
                { throw "Windows Update install result code was $($installResult.ResultCode)"
                }
            } -Tags @('windows') -Resources @('windows-update'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'defender' -Category 'system' -Disabled:$SkipDefender -DisabledReason 'disabled by -SkipDefender' -Script {
                $mpCmdCandidates = @(
                    (Join-Path $env:ProgramFiles 'Windows Defender\MpCmdRun.exe'),
                    (Join-Path ${env:ProgramFiles(x86)} 'Windows Defender\MpCmdRun.exe'),
                    'MpCmdRun.exe'
                ) | Where-Object { $_ }
                $mpCmd = $mpCmdCandidates | Where-Object { ($_ -eq 'MpCmdRun.exe') -or (Test-Path -LiteralPath $_) } | Select-Object -First 1

                if (Get-Command Update-MpSignature -ErrorAction SilentlyContinue)
                {
                    try
                    {
                        $primaryOutput = Update-MpSignature -ErrorAction Stop | Out-String
                        if (-not [string]::IsNullOrWhiteSpace($primaryOutput))
                        { Write-Output $primaryOutput.Trim()
                        }
                        Write-Output 'Defender signatures are current.'
                        return
                    } catch
                    {
                        if (-not $mpCmd)
                        { throw "Defender signature refresh did not complete and MpCmdRun.exe was not found: $($_.Exception.Message)"
                        }
                        Write-Output 'Defender cmdlet did not complete; trying MpCmdRun fallback.'
                    }
                }

                if (-not $mpCmd)
                { Write-Output 'Defender signature updater not found on this system.'; return
                }

                Invoke-UpdateProcess -FilePath $mpCmd -ArgumentList @('-SignatureUpdate') -Retries 1 -TimeoutSec 900
                Write-Output 'Defender signatures refreshed through MpCmdRun fallback.'
            } -Tags @('windows', 'security') -Resources @('defender'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl' -Category 'system' -RequiresCommand 'wsl' -Disabled:$SkipWSL -DisabledReason 'disabled by -SkipWSL' -Script {
                try
                {
                    $wslOutput = Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('--update') -Retries 0 -SuccessExitCodes @(0, -1)
                    $wslText = ($wslOutput | Out-String).Trim()
                    if ($wslText)
                    { Write-Output $wslText
                    }
                } catch
                {
                    $errMsg = $_.Exception.Message
                    Write-Output "wsl --update: $errMsg"
                    if ($errMsg -match 'Forbidden|403|0x80190193|Wsl/UpdatePackage')
                    {
                        Write-Output 'WSL kernel update endpoint error. Treating as non-fatal; installed WSL distros can still be updated by the wsl-distros task.'
                        return
                    }
                    throw
                }
            } -Tags @('windows', 'linux') -Resources @('wsl'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl-distros' -Category 'system' -RequiresCommand 'wsl' -Disabled:($SkipWSL -or $SkipWSLDistros) -DisabledReason 'disabled by WSL skip switch' -TimeoutSec 3600 -Script {
                $benignWslNoise = '(wsl2\.localhostForwarding setting has no effect when using mirrored networking mode|wsl: The wsl2\.localhostForwarding setting has no effect|wsl: An internal error occurred\.|Error code: CreateInstance/CreateVm/ConfigureNetworking/0x8007054f|wsl: Failed to configure network|wsl: Failed to start the systemd user session)'
                $distros = @(Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('-l', '-q') |
                        ForEach-Object { ([string]$_).Replace([string][char]0, '').Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch $benignWslNoise } |
                        Sort-Object -Unique)

                    if ($distros.Count -eq 0)
                    { Write-Output 'No WSL distros found.'; return
                    }

                    $linuxScript = @'
set -u

resolve_any() {
  for host in "$@"; do
    if command -v getent >/dev/null 2>&1 && getent hosts "$host" >/dev/null 2>&1; then return 0; fi
    if command -v nslookup >/dev/null 2>&1 && nslookup "$host" >/dev/null 2>&1; then return 0; fi
    if command -v ping >/dev/null 2>&1 && ping -c 1 -W 2 "$host" >/dev/null 2>&1; then return 0; fi
  done
  return 1
}

if command -v apt-get >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping apt-get: sudo requires a password"
    exit 0
  fi
  if ! resolve_any archive.ubuntu.com security.ubuntu.com; then
    echo "Skipping apt-get: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -o Acquire::Retries=2 update &&   sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y -o Dpkg::Options::=--force-confdef -o Dpkg::Options::=--force-confold full-upgrade &&   sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y autoremove
elif command -v pacman >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping pacman: sudo requires a password"
    exit 0
  fi
  if ! resolve_any archlinux.org geo.mirror.pkgbuild.com; then
    echo "Skipping pacman: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n pacman -Syu --noconfirm --needed
elif command -v zypper >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping zypper: sudo requires a password"
    exit 0
  fi
  if ! resolve_any download.opensuse.org mirrors.opensuse.org; then
    echo "Skipping zypper: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n zypper --non-interactive refresh && sudo -n zypper --non-interactive update
else
  echo "No supported Linux package manager found"
fi
'@

                    $failedDistros = [System.Collections.Generic.List[string]]::new()
                    $skippedDistros = [System.Collections.Generic.List[string]]::new()
                    $networkErrorPattern = 'ConfigureNetworking|0x8007054f|Temporary failure resolving|Could not resolve host|failed to synchronize all databases|networkingMode|^wsl failed with exit code'
                    foreach ($distro in $distros)
                    {
                        Write-Output "Updating WSL distro: $distro"
                        $attempts = 0
                        $maxAttempts = 4
                        $succeeded = $false
                        while ($attempts -lt $maxAttempts -and -not $succeeded)
                        {
                            $attempts++
                            try
                            {
                                $distroOutput = @(Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('--distribution', $distro, '--exec', 'sh', '-lc', $linuxScript) -TimeoutSec 1800 -Retries 0)
                                foreach ($line in $distroOutput)
                                {
                                    $cleanLine = [string]$line
                                    if ($cleanLine -notmatch $benignWslNoise)
                                    { Write-Output $cleanLine
                                    }
                                }
                                $succeeded = $true
                            } catch
                            {
                                $message = $_.Exception.Message
                                if ($message -match $networkErrorPattern)
                                {
                                    if ($attempts -lt $maxAttempts)
                                    {
                                        $backoff = [Math]::Min(60, 10 * $attempts)
                                        Write-Output "WSL distro '$distro' hit network error (attempt $attempts/$maxAttempts); retrying in ${backoff}s..."
                                        Start-Sleep -Seconds $backoff
                                    } else
                                    {
                                        Write-Output "Skipping WSL distro after transient network problem: $distro"
                                        Write-Output "  Tip: run 'wsl --shutdown' then rerun the wsl-distros task to reset WSL networking."
                                        [void]$skippedDistros.Add($distro)
                                    }
                                } else
                                {
                                    Write-Output $message
                                    [void]$failedDistros.Add($distro)
                                    break
                                }
                            }
                        }
                    }
                    if ($failedDistros.Count -gt 0)
                    { throw "WSL distro updates did not complete for: $($failedDistros -join ', ')"
                    }
                    if ($skippedDistros.Count -gt 0)
                    { Set-TaskStatus -Status 'Partial' -Reason "WSL distro updates skipped after transient network problem: $($skippedDistros -join ', ')"
                    }
                } -Tags @('windows', 'linux') -Resources @('wsl'))) | Out-Null

        $npmScript = {
            param([string[]]$SkipPackages)

            function Remove-StaleNpmTempFolders
            {
                param([string]$Root)
                if (-not $Root -or -not (Test-Path -LiteralPath $Root))
                { return
                }
                Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '^\.[A-Za-z0-9_-]+-' } |
                    ForEach-Object {
                        try
                        {
                            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                            Write-Output "Removed stale npm temp folder: $($_.Name)"
                        } catch
                        {
                            Write-Output "Could not remove stale npm temp folder $($_.Name): $($_.Exception.Message)"
                        }
                    }
        }

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($package in @($SkipPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($package))
            { [void]$skipSet.Add($package.Trim())
            }
        }

        $root = (Invoke-UpdateProcess -FilePath 'npm' -ArgumentList @('root', '-g') | Select-Object -First 1)
        Remove-StaleNpmTempFolders -Root $root

        $listRaw = (Invoke-UpdateProcess -FilePath 'npm' -ArgumentList @('ls', '-g', '--depth=0', '--json') -SuccessExitCodes @(0, 1) | Out-String).Trim()
        if (-not $listRaw)
        { throw 'npm did not return a global package list.'
        }
        # npm sometimes emits warnings or notices before/after the JSON object; extract the first {...} block
        $jsonMatch = [regex]::Match($listRaw, '(\{[\s\S]*\})')
        $listJson = if ($jsonMatch.Success) { $jsonMatch.Groups[1].Value } else { $listRaw }
        $tree = $listJson | ConvertFrom-Json -ErrorAction Stop

        $packageNames = @()
        if ($tree.PSObject.Properties['dependencies'] -and $tree.dependencies)
        {
            $packageNames = @($tree.dependencies.PSObject.Properties.Name)
        }
        if ($packageNames.Count -eq 0)
        { Write-Output 'No global npm packages found.'; return
        }

        $validNamePattern = '^(?:@[a-z0-9][a-z0-9._~-]*/)?[a-z0-9][a-z0-9._~-]*$'
        $failed = [System.Collections.Generic.List[string]]::new()
        foreach ($name in ($packageNames | Sort-Object -Unique))
        {
            if ($skipSet.Contains($name))
            { Write-Output "Skipping npm package from config: $name"; continue
            }
            if ($name -notmatch $validNamePattern)
            { Write-Output "Skipping invalid npm package name: $name"; continue
            }

            if ($name -ieq '@openai/codex')
            {
                Get-Process -Name 'codex' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Remove-StaleNpmTempFolders -Root (Join-Path $root '@openai')
            }

            $spec = "$name@latest"
            Write-Output "Updating npm package: $spec"
            try
            {
                Invoke-UpdateProcess -FilePath 'npm' -ArgumentList @('install', '-g', $spec, '--no-fund', '--no-audit') -Retries 2 -TimeoutSec 900
            } catch
            {
                Write-Output $_.Exception.Message
                [void]$failed.Add($name)
            }
        }

        Remove-StaleNpmTempFolders -Root $root
        if ($failed.Count -gt 0)
        { Write-Output "npm packages left unchanged: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'npm' -Category 'javascript' -RequiresCommand 'npm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script $npmScript -Tags @('node') -Resources @('npm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pnpm' -Category 'javascript' -RequiresCommand 'pnpm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                $result = Invoke-UpdateProcess -FilePath 'pnpm' -ArgumentList @('self-update') -Retries 1 -TimeoutSec 120 -PassThru
                $outText = (@($result.Output) | Out-String).Trim()
                if ($outText)
                { Write-Output $outText
                }
                if ($result.ExitCode -ne 0 -or $outText -match '(?i)(not recognized as a name of a cmdlet|@pnpm/exe/pnpm\.exe|pnpm\.ps1)')
                {
                    $global:LASTEXITCODE = if ($result.ExitCode -ne 0)
                    { $result.ExitCode
                    } else
                    { 1
                    }
                    throw 'pnpm self-update did not complete; the pnpm shim appears to reference a missing executable. Run pnpm setup or reinstall pnpm.'
                }
            } -Tags @('node') -Resources @('npm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'yarn' -Category 'javascript' -RequiresCommand 'yarn' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'yarn' -ArgumentList @('global', 'upgrade') -Retries 1 -TimeoutSec 300
            } -Tags @('node') -Resources @('npm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'bun' -Category 'javascript' -RequiresCommand 'bun' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'bun' -ArgumentList @('upgrade') -Retries 1 -TimeoutSec 120
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'deno' -Category 'javascript' -RequiresCommand 'deno' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'deno' -ArgumentList @('upgrade') -Retries 1 -TimeoutSec 120
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'mise' -Category 'version-manager' -RequiresCommand 'mise' -Script {
                Invoke-UpdateProcess -FilePath 'mise' -ArgumentList @('self-update', '--yes') -Retries 1 -TimeoutSec 120
                Invoke-UpdateProcess -FilePath 'mise' -ArgumentList @('upgrade', '--yes') -Retries 1 -TimeoutSec 300
            } -Tags @('version-manager'))) | Out-Null

    $pipScript = {
        param([string[]]$SkipPackages)
        Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', 'pip') -Retries 1

        function Normalize-PipName { param([string]$n) ($n.ToLowerInvariant() -replace '[-_.]', '-') }

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($pkg in @($SkipPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($pkg))
            { [void]$skipSet.Add((Normalize-PipName $pkg.Trim()))
            }
        }
        if ($skipSet.Count -gt 0)
        { Write-Output "Configured pip package skip list: $($skipSet -join ', ')"
        }

        $outdatedJson = (Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'list', '--outdated', '--format=json') | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($outdatedJson))
        { Write-Output 'No outdated pip packages found.'; return
        }
        $outdated = @($outdatedJson | ConvertFrom-Json -ErrorAction Stop)
        if ($outdated.Count -eq 0)
        { Write-Output 'No outdated pip packages found.'; return
        }

        $failed = [System.Collections.Generic.List[string]]::new()
        foreach ($pkg in $outdated)
        {
            if ($skipSet.Contains((Normalize-PipName $pkg.name)))
            { Write-Output "Skipping pip package: $($pkg.name)"; continue
            }
            Write-Output "Upgrading pip package: $($pkg.name) $($pkg.version) -> $($pkg.latest_version)"
            try
            { Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', '--upgrade-strategy', 'only-if-needed', $pkg.name) -Retries 1
            } catch
            { Write-Output $_.Exception.Message; [void]$failed.Add($pkg.name)
            }
        }
        if ($failed.Count -gt 0)
        { Write-Output "pip packages left unchanged: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'pip' -Category 'python' -RequiresCommand 'python' -Script $pipScript -Tags @('python') -Resources @('pip'))) | Out-Null

    $pipHealthScript = {
        param([string[]]$IgnorePackages)

        $ignoreSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($pkg in @($IgnorePackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($pkg))
            { [void]$ignoreSet.Add($pkg.Trim())
            }
        }

        $result = Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'check') -SuccessExitCodes @(0, 1) -PassThru
        $lines = @($result.Output | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
        if ($result.ExitCode -eq 0)
        {
            Write-Output 'pip dependency check passed.'
            return
        }

        if ($lines.Count -eq 0)
        {
            $global:LASTEXITCODE = $result.ExitCode
            throw "pip check failed with exit code $($result.ExitCode) and no output."
        }

        $activeIssues = [System.Collections.Generic.List[string]]::new()
        $ignoredIssues = [System.Collections.Generic.List[string]]::new()
        foreach ($line in $lines)
        {
            $packageName = $null
            if ($line -match '^([A-Za-z0-9_.-]+)\s')
            { $packageName = $matches[1]
            }

            if ($packageName -and $ignoreSet.Contains($packageName))
            { [void]$ignoredIssues.Add($line)
            } else
            { [void]$activeIssues.Add($line)
            }
        }

        if ($ignoredIssues.Count -gt 0)
        {
            Write-Output ("Ignored pip dependency issue(s) from config ({0})." -f $ignoredIssues.Count)
            foreach ($issue in $ignoredIssues)
            { Write-Output "  $issue"
            }
        }

        if ($activeIssues.Count -eq 0)
        {
            Write-Output 'pip dependency check only found ignored issue(s).'
            return
        }

        # Advisory only: pip check never updates anything and the conflicts are
        # usually pre-existing environment state. Report, but never fail the run.
        Write-Output ("pip dependency check found {0} conflict(s) (advisory, not failing the run):" -f $activeIssues.Count)
        foreach ($issue in $activeIssues)
        { Write-Output "  $issue"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'pip-health' -Category 'python' -RequiresCommand 'python' -Disabled:$SkipPipHealth -DisabledReason 'disabled by -SkipPipHealth' -Script $pipHealthScript -Tags @('python', 'health') -Resources @('pip'))) | Out-Null

    # pipx defaults to the uv backend; force pip so it works when uv is not installed.
    $tasks.Add((New-UpdateTask -Name 'pipx' -Category 'python' -RequiresCommand 'pipx' -Script {
                Invoke-UpdateProcess -FilePath 'pipx' -ArgumentList @('upgrade-all', '--backend', 'pip') -Retries 1 -TimeoutSec 600
            } -Tags @('python') -Resources @('pip'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv' -Category 'python' -RequiresCommand 'uv' -Script {
                $uvPath = Get-ToolCommandPath -Name 'uv'
                if (-not $uvPath)
                { $uvPath = (Get-Command uv -ErrorAction SilentlyContinue).Source
                }

                $managedPathPattern = '\\(Python\d+\\Scripts|pipx\\venvs|Microsoft\\WinGet\\Packages|scoop\\apps|chocolatey\\lib|WindowsApps)\\'
                $managedMessagePattern = '(standalone installation|managed install|installed through another package manager|self-update is only available|cannot be self-updated)'

                $uvExe = 'uv'
                if ($uvPath -and $uvPath -match $managedPathPattern)
                {
                    $standaloneCandidates = @(
                        (Join-Path $env:USERPROFILE '.local\bin\uv.exe'),
                        (Join-Path $env:LOCALAPPDATA 'uv\bin\uv.exe'),
                        (Join-Path $env:USERPROFILE '.cargo\bin\uv.exe')
                    )
                    $standaloneUv = $standaloneCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
                    if ($standaloneUv)
                    {
                        Write-Output "Active uv at '$uvPath' is managed; using standalone uv at '$standaloneUv' for self-update."
                        $uvExe = $standaloneUv
                    } else
                    {
                        Write-Output "uv self-update skipped: '$uvPath' is managed by another installer."
                        Write-Output "uv tools/python are still updated via uv-tools and uv-python tasks."
                        Write-Output "To enable uv self-update, install standalone uv via: irm https://astral.sh/uv/install.ps1 | iex"
                        return
                    }
                }

                $result = Invoke-UpdateProcess -FilePath $uvExe -ArgumentList @('self', 'update') -SuccessExitCodes @(0, 1) -PassThru
                $outText = (@($result.Output) | Out-String).Trim()

                if ($outText -match $managedMessagePattern)
                {
                    Set-TaskStatus -Status 'Warn' -Reason 'uv self-update skipped: managed install'
                    Write-Output "uv self-update skipped: managed install detected."
                    if ($outText) { Write-Output $outText }
                    return
                }

                # uv self-update replaces uvx.exe alongside uv.exe; anything hosting a
                # uvx tool (commonly an MCP server) holds it open. Close the holders by
                # exact path and retry -- whatever spawned them restarts them.
                if ($result.ExitCode -ne 0 -and $outText -match 'being used by another process')
                {
                    $binDir = Split-Path -Parent $uvExe
                    $targets = @('uv.exe', 'uvx.exe') | ForEach-Object { Join-Path $binDir $_ }
                    $holders = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $targets -contains $_.Path })
                    if ($holders.Count -gt 0)
                    {
                        foreach ($h in $holders)
                        {
                            Write-Output "closed $($h.ProcessName) (pid $($h.Id))"
                            Stop-Process -Id $h.Id -Force -ErrorAction SilentlyContinue
                        }
                        Start-Sleep -Seconds 2
                        $result = Invoke-UpdateProcess -FilePath $uvExe -ArgumentList @('self', 'update') -SuccessExitCodes @(0, 1) -PassThru
                        $outText = (@($result.Output) | Out-String).Trim()
                    } else
                    {
                        Set-TaskStatus -Status 'Skipped' -Reason 'uv self-update blocked; no closable holder found'
                        Write-Output 'uv self-update blocked and no closable holder was found.'
                        return
                    }
                }

                if ($result.ExitCode -ne 0)
                {
                    if ($outText) { Write-Output $outText }
                    $global:LASTEXITCODE = $result.ExitCode
                    throw "uv self-update failed with exit code $($result.ExitCode)."
                }

                if ($outText) { Write-Output $outText }
            } -Tags @('python', 'uv') -Resources @('uv'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv-tools' -Category 'python' -RequiresCommand 'uv' -Disabled:$SkipUVTools -DisabledReason 'disabled by -SkipUVTools' -Script {
                Invoke-UpdateProcess -FilePath 'uv' -ArgumentList @('tool', 'upgrade', '--all') -Retries 1
            } -Tags @('python', 'uv') -Resources @('uv'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv-python' -Category 'python' -RequiresCommand 'uv' -Disabled:$SkipUVTools -DisabledReason 'disabled by -SkipUVTools' -Script {
                $listResult = Invoke-UpdateProcess -FilePath 'uv' -ArgumentList @('python', 'list', '--only-installed') -SuccessExitCodes @(0, 1) -PassThru
                $listText = (@($listResult.Output) | Out-String).Trim()
                if ([string]::IsNullOrWhiteSpace($listText) -or $listText -match '(?i)no python|no installed|not found')
                {
                    Write-Output 'No uv-managed Python versions found.'
                    return
                }

                $versions = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($line in @($listResult.Output))
                {
                    if ([string]$line -match 'cpython-(\d+\.\d+)')
                    { [void]$versions.Add($matches[1])
                    }
                }

                if ($versions.Count -eq 0)
                {
                    Write-Output 'No uv-managed CPython installs detected.'
                    return
                }

                $versionList = @($versions | Sort-Object)
                Write-Output "Refreshing uv-managed Python patch releases: $($versionList -join ', ')"
                Invoke-UpdateProcess -FilePath 'uv' -ArgumentList (@('python', 'install') + $versionList) -Retries 1 -TimeoutSec 1800
            } -Tags @('python', 'uv') -Resources @('uv'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'poetry' -Category 'python' -RequiresCommand 'poetry' -Disabled:$SkipPoetry -DisabledReason 'disabled by -SkipPoetry' -Script {
                # pipx owns poetry's venv when it installed it; 'poetry self update'
                # then re-pins poetry's shared libraries down to its own declared
                # bounds, undoing pipx's upgrades on every run.
                $pipxManaged = $false
                if (Get-Command pipx -ErrorAction SilentlyContinue)
                {
                    try
                    {
                        $pipxManaged = @(& pipx list --short 2>$null) -match '^poetry\s'
                    } catch
                    {
                    }
                }
                if ($pipxManaged)
                {
                    Set-TaskStatus -Status 'Skipped' -Reason 'poetry is pipx-managed; upgrades handled by the pipx task'
                    Write-Output 'poetry is pipx-managed; upgrades handled by the pipx task.'
                    return
                }
                try
                {
                    Invoke-UpdateProcess -FilePath 'poetry' -ArgumentList @('self', 'update', '--no-interaction') -Retries 0 -TimeoutSec 120
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "poetry self-update skipped: $($_.Exception.Message)"
                    Write-Output "poetry self-update skipped: $($_.Exception.Message)"
                }
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'rustup' -Category 'systems-language' -RequiresCommand 'rustup' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                Invoke-UpdateProcess -FilePath 'rustup' -ArgumentList @('update') -Retries 1 -TimeoutSec 300
            } -Tags @('rust') -Resources @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'cargo' -Category 'systems-language' -RequiresCommand 'cargo' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                if (-not (Get-Command cargo-install-update -ErrorAction SilentlyContinue))
                {
                    Invoke-UpdateProcess -FilePath 'cargo' -ArgumentList @('install', 'cargo-update', '-q') -Retries 1 -TimeoutSec 300
                }
                Invoke-UpdateProcess -FilePath 'cargo' -ArgumentList @('install-update', '-a') -Retries 1 -TimeoutSec 600
            } -Tags @('rust') -Resources @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'go' -Category 'systems-language' -RequiresCommand 'go' -Disabled:$SkipGo -DisabledReason 'disabled by -SkipGo' -Script {
                Invoke-UpdateProcess -FilePath 'go' -ArgumentList @('install', 'golang.org/x/tools/gopls@latest') -Retries 1 -TimeoutSec 300
            } -Tags @('go'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'flutter' -Category 'systems-language' -RequiresCommand 'flutter' -Disabled:$SkipFlutter -DisabledReason 'disabled by -SkipFlutter' -Script {
                try
                {
                    Invoke-UpdateProcess -FilePath 'flutter' -ArgumentList @('upgrade') -Retries 0 -TimeoutSec 60
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "flutter upgrade skipped: $($_.Exception.Message)"
                    Write-Output "flutter upgrade skipped: $($_.Exception.Message)"
                }
            } -Tags @('flutter'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'juliaup' -Category 'systems-language' -RequiresCommand 'juliaup' -Script {
                Invoke-UpdateProcess -FilePath 'juliaup' -ArgumentList @('update') -Retries 1 -TimeoutSec 1800
            } -Tags @('julia'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'gh' -Category 'dev-tools' -RequiresCommand 'gh' -Script {
                $toolPath = Get-ToolCommandPath -Name 'gh'
                if ($toolPath -match '\\(scoop\\apps|scoop\\shims|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "gh is managed by another package manager; handled by scoop/chocolatey/winget task: $toolPath"
                    return
                }
                if (Get-Command winget -ErrorAction SilentlyContinue)
                {
                    $wr = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', 'GitHub.cli', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent') -Retries 0 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                    $wt = (@($wr.Output) | Out-String).Trim()
                    if ($wt -match 'No installed package found')
                    {
                        Write-Output 'gh not installed via winget; managed elsewhere (scoop/chocolatey/standalone).'
                        return
                    }
                    if ($wt) { Write-Output $wt }
                    Write-Output 'gh checked through winget id GitHub.cli.'
                }
            } -Tags @('github', 'dev-tools'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'mise-upgrade' -Category 'version-manager' -RequiresCommand 'mise' -Script {
                Invoke-UpdateProcess -FilePath 'mise' -ArgumentList @('upgrade') -Retries 1 -TimeoutSec 1800
            } -Tags @('mise', 'version-manager'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'tldr' -Category 'dev-tools' -RequiresCommand 'tldr' -Script {
                $toolPath = Get-ToolCommandPath -Name 'tldr'
                if ($toolPath -match '\\pipx\\venvs\\')
                {
                    Invoke-UpdateProcess -FilePath 'pipx' -ArgumentList @('upgrade', 'tldr') -Retries 1
                    return
                }
                Invoke-UpdateProcess -FilePath 'tldr' -ArgumentList @('--update') -Retries 1 -TimeoutSec 120
            } -Tags @('dev-tools'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'winget-pin-audit' -Category 'package-manager' -RequiresCommand 'winget' -Script {
                $output = @(Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('pin', 'list') -SuccessExitCodes @(0, -1978335212) -PassThru)
                $lines = @(($output | Select-Object -ExpandProperty Output) | Where-Object { $_ -match '\S' })
                if ($lines.Count -le 2)
                { Write-Output 'No pinned winget packages.' }
                else
                { $lines | ForEach-Object { Write-Output $_ } }
            } -Tags @('windows', 'winget', 'audit'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'powershell7' -Category 'system' -RequiresCommand 'winget' -Script {
                $wr = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', 'Microsoft.PowerShell', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent') -Retries 1 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                $wt = (@($wr.Output) | Out-String).Trim()
                if ($wt) { Write-Output $wt }
                Write-Output 'PowerShell 7 checked via winget id Microsoft.PowerShell.'
            } -Tags @('windows', 'powershell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'docker-prune' -Category 'maintenance' -RequiresCommand 'docker' -Disabled:($SkipCleanup -or -not $DeepClean) -DisabledReason 'opt-in: requires -DeepClean and docker running' -Script {
                $info = Invoke-UpdateProcess -FilePath 'docker' -ArgumentList @('info', '--format', '{{.ServerVersion}}') -SuccessExitCodes @(0, 1) -PassThru
                if ($info.ExitCode -ne 0)
                { Write-Output 'Docker daemon not running; skipping prune.'; return }
                Invoke-UpdateProcess -FilePath 'docker' -ArgumentList @('system', 'prune', '-f', '--volumes') -TimeoutSec 300 -Retries 0
            } -Tags @('docker', 'maintenance'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-tools' -Category 'dotnet' -RequiresCommand 'dotnet' -Script {
                Invoke-UpdateProcess -FilePath 'dotnet' -ArgumentList @('tool', 'update', '--global', '--all') -Retries 1
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-workloads' -Category 'dotnet' -RequiresCommand 'dotnet' -RequiresAdmin -Script {
                Invoke-UpdateProcess -FilePath 'dotnet' -ArgumentList @('workload', 'update') -TimeoutSec 3600 -Retries 1
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'ruby-gems' -Category 'runtime' -RequiresCommand 'gem' -Disabled:$SkipRuby -DisabledReason 'disabled by -SkipRuby' -Script {
                try
                {
                    Invoke-UpdateProcess -FilePath 'gem' -ArgumentList @('update', '--system') -Retries 0 -TimeoutSec 120
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "gem update --system skipped: $($_.Exception.Message)"
                    Write-Output "gem update --system skipped: $($_.Exception.Message)"
                    return
                }
                try
                {
                    Invoke-UpdateProcess -FilePath 'gem' -ArgumentList @('update') -Retries 0 -TimeoutSec 300
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "gem update skipped: $($_.Exception.Message)"
                    Write-Output "gem update skipped: $($_.Exception.Message)"
                }
            } -Tags @('ruby'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'composer' -Category 'runtime' -RequiresCommand 'composer' -Disabled:$SkipComposer -DisabledReason 'disabled by -SkipComposer' -Script {
                try
                {
                    Invoke-UpdateProcess -FilePath 'composer' -ArgumentList @('self-update', '--no-interaction') -Retries 0 -TimeoutSec 120
                } catch
                {
                    $errMsg = $_.Exception.Message
                    if ($errMsg -match 'openssl extension is required')
                    {
                        Set-TaskStatus -Status 'Warn' -Reason 'composer self-update skipped: openssl PHP extension missing'
                        Write-Output 'composer self-update skipped: openssl PHP extension is not loaded.'
                        Write-Output '  Fix: enable "extension=openssl" in your php.ini, then rerun.'
                        Write-Output ('  PHP ini location: {0}' -f (& php --ini 2>$null | Select-String 'Loaded Configuration' | ForEach-Object { ($_ -split ':\s*')[1] }))
                    } else
                    {
                        Set-TaskStatus -Status 'Warn' -Reason "composer self-update skipped: $errMsg"
                        Write-Output "composer self-update skipped: $errMsg"
                    }
                }
            } -Tags @('php'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'yt-dlp' -Category 'media-tools' -RequiresCommand 'yt-dlp' -Script {
                $toolPath = Get-ToolCommandPath -Name 'yt-dlp'
                if ($toolPath -match '\\pipx\\venvs\\')
                {
                    Write-Output 'yt-dlp is managed by pipx; covered by the pipx task (pipx upgrade-all).'
                    return
                }
                if ($toolPath -match '\\Python\d+\\Scripts\\|\\Python\\Python\d+\\Scripts\\')
                {
                    Write-Output 'yt-dlp is managed by pip; covered by the pip task.'
                    return
                }
                if ($toolPath -match '\\(scoop\\apps|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "yt-dlp is managed by another package manager: $toolPath"
                    return
                }
                try
                {
                    Invoke-UpdateProcess -FilePath 'yt-dlp' -ArgumentList @('-U') -Retries 0 -TimeoutSec 120
                } catch
                {
                    Set-TaskStatus -Status 'Warn' -Reason "yt-dlp update skipped: $($_.Exception.Message)"
                    Write-Output "yt-dlp update skipped: $($_.Exception.Message)"
                }
            } -Tags @('media', 'python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'oh-my-posh' -Category 'shell' -RequiresCommand 'oh-my-posh' -Script {
                $toolPath = Get-ToolCommandPath -Name 'oh-my-posh'
                if ($toolPath -match '\\(scoop\\apps|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "oh-my-posh is managed by another package manager: $toolPath"
                    return
                }
                if (Get-Command winget -ErrorAction SilentlyContinue)
                {
                    try
                    {
                        $wingetResult = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', 'JanDeDobbeleer.OhMyPosh', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent') -Retries 0 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                        $wingetText = (@($wingetResult.Output) | Out-String).Trim()
                        if ($wingetText -match 'No installed package found matching input criteria')
                        {
                            Write-Output 'oh-my-posh is not registered as a winget package; trying standalone upgrade.'
                        } else
                        {
                            if ($wingetText)
                            { Write-Output $wingetText
                            }
                            Write-Output 'oh-my-posh checked through winget package id JanDeDobbeleer.OhMyPosh.'
                            return
                        }
                    } catch
                    {
                        Write-Output "winget oh-my-posh upgrade check did not complete: $($_.Exception.Message)"
                    }
                }

                $result = Invoke-UpdateProcess -FilePath 'oh-my-posh' -ArgumentList @('upgrade') -Retries 0 -TimeoutSec 120 -PassThru
                $outText = (@($result.Output) | Out-String).Trim()
                $displayText = [regex]::Replace($outText, '\x1b\][^\a]*(\a|\x1b\\)', '').Trim()
                if ($outText)
                {
                    if ($displayText)
                    { Write-Output $displayText
                    } else
                    { Write-Output 'oh-my-posh standalone upgrade command completed without textual output.'
                    }
                }
                if ($result.TimedOut -or $result.ExitCode -ne 0)
                {
                    $global:LASTEXITCODE = if ($result.ExitCode -ne 0)
                    { $result.ExitCode
                    } else
                    { 124
                    }
                    throw 'oh-my-posh upgrade did not complete. Prefer updating it through winget/Scoop/Chocolatey or rerun the command manually in an interactive terminal.'
                }
            } -Tags @('shell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'claude' -Category 'ai-tools' -RequiresCommand 'claude' -TimeoutSec 300 -Script {
                if (Get-Command winget -ErrorAction SilentlyContinue)
                {
                    $wr = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', 'Anthropic.Claude', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent') -Retries 0 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                    $wt = (@($wr.Output) | Out-String).Trim()
                    if ($wt -notmatch 'No installed package found matching input criteria')
                    {
                        if ($wt) { Write-Output $wt }
                        Write-Output 'claude checked through winget id Anthropic.Claude.'
                        return
                    }
                }
                Invoke-UpdateProcess -FilePath 'claude' -ArgumentList @('update') -Retries 0 -TimeoutSec 120
            } -Tags @('ai', 'claude'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'codex' -Category 'ai-tools' -RequiresCommand 'codex' -TimeoutSec 300 -Script {
                if (-not (Get-Command winget -ErrorAction SilentlyContinue))
                { Write-Output 'winget not available; skipping codex upgrade.'; return }
                # Codex is a portable winget package that locks its exe while running; kill before upgrade
                $codexProcs = @(Get-Process | Where-Object { $_.Name -like '*codex*' -or $_.Path -like '*OpenAI.Codex*' })
                if ($codexProcs.Count -gt 0)
                {
                    Write-Output "Stopping $($codexProcs.Count) codex process(es) before upgrade..."
                    $codexProcs | Stop-Process -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }
                $wr = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', 'OpenAI.Codex', '--exact', '--source', 'winget', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent', '--force') -Retries 1 -TimeoutSec 300 -SuccessExitCodes @(0, -1978335189, -1978335212) -PassThru
                $wt = (@($wr.Output) | Out-String).Trim()
                if ($wt) { Write-Output $wt }
                if ($wr.ExitCode -eq -1978335189 -or $wt -match 'No applicable upgrade found|No available upgrade found')
                { Write-Output 'codex: no upgrade available.'
                } else
                { Write-Output 'codex checked through winget id OpenAI.Codex.'
                }
            } -Tags @('ai', 'codex'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'volta' -Category 'javascript' -RequiresCommand 'volta' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                $toolPath = Get-ToolCommandPath -Name 'volta'
                if ($toolPath -match '\\(scoop\\apps|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "Volta is managed by another package manager: $toolPath"
                    return
                }
                Write-Output 'Volta does not provide a stable non-interactive self-update command. Update its installer manually or install it through winget/Scoop/Chocolatey for automatic coverage.'
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'fnm' -Category 'javascript' -RequiresCommand 'fnm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                $toolPath = Get-ToolCommandPath -Name 'fnm'
                if ($toolPath -match '\\(scoop\\apps|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "fnm is managed by another package manager: $toolPath"
                    return
                }
                Write-Output 'fnm does not provide a stable non-interactive self-update command. Update its installer manually or install it through winget/Scoop/Chocolatey for automatic coverage.'
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'vscode-extensions' -Category 'dev-tools' -RequiresCommand 'code' -Disabled:$SkipVSCodeExtensions -DisabledReason 'disabled by -SkipVSCodeExtensions' -Script {
                $codePath = Get-ToolCommandPath -Name 'code'
                if (-not $codePath)
                { throw 'VS Code CLI was not found. Reinstall VS Code or add code.cmd to PATH.'
                }
                Write-Output "Using VS Code CLI: $codePath"
                Invoke-UpdateProcess -FilePath $codePath -ArgumentList @('--update-extensions') -Retries 1 -TimeoutSec 900
            } -Tags @('editor') -Resources @('vscode'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'git-lfs' -Category 'dev-tools' -RequiresCommand 'git-lfs' -Disabled:$SkipGitLFS -DisabledReason 'disabled by -SkipGitLFS' -Script {
                Invoke-UpdateProcess -FilePath 'git' -ArgumentList @('lfs', 'install', '--skip-repo')
            } -Tags @('git'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'gh-extensions' -Category 'dev-tools' -RequiresCommand 'gh' -Script {
                $extensions = @(Invoke-UpdateProcess -FilePath 'gh' -ArgumentList @('extension', 'list') |
                        ForEach-Object { ($_ -split '\s+')[0] } |
                        Where-Object { $_ -match '^gh-' })
                    $failedExtensions = [System.Collections.Generic.List[string]]::new()
                    foreach ($extension in $extensions)
                    {
                        try
                        { Invoke-UpdateProcess -FilePath 'gh' -ArgumentList @('extension', 'upgrade', $extension) -Retries 1
                        } catch
                        { Write-Output $_.Exception.Message; [void]$failedExtensions.Add($extension)
                        }
                    }
                    if ($failedExtensions.Count -gt 0)
                    { Write-Output "GitHub extensions left unchanged: $($failedExtensions -join ', ')"
                    }
                } -Tags @('github'))) | Out-Null

        $tasks.Add((New-UpdateTask -Name 'powershell-modules' -Category 'powershell' -Disabled:$SkipPowerShellModules -DisabledReason 'disabled by -SkipPowerShellModules' -Script {
                    if ((Get-Command Get-PSResource -ErrorAction SilentlyContinue) -and (Get-Command Update-PSResource -ErrorAction SilentlyContinue))
                    {
                        $resources = @(Get-PSResource -Name '*' -ErrorAction SilentlyContinue)
                        if ($resources.Count -eq 0)
                        { Write-Output 'No installed PSResource modules found.'; return
                        }

                        $failedModules = [System.Collections.Generic.List[string]]::new()
                        $cmd = Get-Command Update-PSResource -ErrorAction SilentlyContinue
                        foreach ($resource in ($resources | Sort-Object Name -Unique))
                        {
                            $moduleArgs = @{ Name = $resource.Name; ErrorAction = 'Stop' }
                            if ($cmd.Parameters.ContainsKey('AcceptLicense'))
                            { $moduleArgs.AcceptLicense = $true
                            }
                            if ($cmd.Parameters.ContainsKey('TrustRepository'))
                            { $moduleArgs.TrustRepository = $true
                            }
                            try
                            { Update-PSResource @moduleArgs | Out-String | Write-Output
                            } catch
                            { Write-Output "PowerShell module not updated: $($resource.Name) — $($_.Exception.Message)"; [void]$failedModules.Add($resource.Name)
                            }
                        }
                        if ($failedModules.Count -gt 0)
                        { Write-Output "PowerShell modules left unchanged: $($failedModules -join ', ')"
                        }
                    } elseif (Get-Command Update-Module -ErrorAction SilentlyContinue)
                    {
                        $installedModules = @(Get-InstalledModule -ErrorAction SilentlyContinue)
                        if ($installedModules.Count -eq 0)
                        { Write-Output 'No installed PowerShellGet modules found.'; return
                        }
                        $failedModules = [System.Collections.Generic.List[string]]::new()
                        $installedModules | ForEach-Object {
                            $moduleName = $_.Name
                            try
                            { Update-Module -Name $moduleName -Force -ErrorAction Stop
                            } catch
                            { Write-Output "PowerShell module not updated: ${moduleName} — $($_.Exception.Message)"; [void]$failedModules.Add($moduleName)
                            }
                        }
                        if ($failedModules.Count -gt 0)
                        { Write-Output "PowerShell modules left unchanged: $($failedModules -join ', ')"
                        }
                    } else
                    {
                        Write-Output 'No PowerShell module updater found.'
                    }
                } -Tags @('powershell') -Resources @('powershell-gallery'))) | Out-Null

        $tasks.Add((New-UpdateTask -Name 'powershell-help' -Category 'powershell' -Disabled:(-not $UpdatePowerShellHelp) -DisabledReason 'opt-in via -UpdatePowerShellHelp' -Script {
                    try
                    { Update-Help -Force -ErrorAction Stop | Out-String | Write-Output
                    } catch
                    { Write-Output "PowerShell help was not fully refreshed: $($_.Exception.Message)"
                    }
                } -Tags @('powershell') -Resources @('powershell-gallery'))) | Out-Null

        $ollamaScript = {
            param([int]$CommandTimeoutSec)
            $listTimeout = [Math]::Min($CommandTimeoutSec, 60)
            $listOutput = Invoke-UpdateProcess -FilePath 'ollama' -ArgumentList @('list') -TimeoutSec $listTimeout
            if ($listOutput.Count -gt 0)
            { $listOutput | Write-Output
            }

            # Advisory + bounded. Refreshing local models is best-effort:
            #  - locally-built (Modelfile) models have no registry source -> pull
            #    fails fast; we record and move on instead of failing the run.
            #  - oversized models are skipped to avoid multi-GB re-pulls that blow
            #    the time budget; a short per-model timeout caps any single hang.
            $maxGb = 20.0
            $perModelTimeout = [Math]::Min($CommandTimeoutSec, 300)

            $models = [System.Collections.Generic.List[object]]::new()
            foreach ($row in @($listOutput | Select-Object -Skip 1))
            {
                $cols = @($row -split '\s+' | Where-Object { $_ })
                if ($cols.Count -eq 0) { continue }
                $name = $cols[0]
                if (-not $name) { continue }
                $gb = $null
                for ($i = 0; $i -lt $cols.Count - 1; $i++)
                {
                    if ($cols[$i] -match '^[0-9.]+$' -and $cols[$i + 1] -match '^(?i)(GB|MB|KB|B)$')
                    {
                        $val = [double]$cols[$i]
                        switch ($cols[$i + 1].ToUpper())
                        {
                            'GB' { $gb = $val }
                            'MB' { $gb = $val / 1024 }
                            'KB' { $gb = $val / 1048576 }
                            'B'  { $gb = $val / 1073741824 }
                        }
                        break
                    }
                }
                $models.Add([pscustomobject]@{ Name = $name; Gb = $gb })
            }
            if ($models.Count -eq 0)
            { Write-Output 'No Ollama models found.'; return
            }

            $updated = [System.Collections.Generic.List[string]]::new()
            $unchanged = [System.Collections.Generic.List[string]]::new()
            $skipped = [System.Collections.Generic.List[string]]::new()
            foreach ($m in $models)
            {
                if ($null -ne $m.Gb -and $m.Gb -gt $maxGb)
                {
                    Write-Output ("Skipping large model ({0:N0} GB > {1:N0} GB): {2}" -f $m.Gb, $maxGb, $m.Name)
                    [void]$skipped.Add($m.Name)
                    continue
                }
                Write-Output "Pulling Ollama model: $($m.Name)"
                try
                { Invoke-UpdateProcess -FilePath 'ollama' -ArgumentList @('pull', $m.Name) -TimeoutSec $perModelTimeout -Retries 0
                    [void]$updated.Add($m.Name)
                } catch
                { Write-Output "  $($_.Exception.Message)"; [void]$unchanged.Add($m.Name)
                }
            }
            Write-Output ("Ollama: {0} refreshed, {1} unchanged, {2} skipped (too large)." -f $updated.Count, $unchanged.Count, $skipped.Count)
            if ($unchanged.Count -gt 0)
            { Write-Output "  unchanged (local/unavailable): $($unchanged -join ', ')"
            }
        }
        $tasks.Add((New-UpdateTask -Name 'ollama-models' -Category 'ai' -RequiresCommand 'ollama' -Disabled:(-not $UpdateOllamaModels) -DisabledReason 'use -UpdateOllamaModels to refresh local models' -TimeoutSec 7200 -Script $ollamaScript -Tags @('ai') -Resources @('ollama'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'vcpkg' -Category 'package-manager' -RequiresCommand 'vcpkg' -Script {
                param([string[]]$SkipPackages)
                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($pkg in @($SkipPackages)) { if (-not [string]::IsNullOrWhiteSpace($pkg)) { [void]$skipSet.Add($pkg.Trim()) } }
                Write-Output 'Updating vcpkg baseline and installed packages...'
                Invoke-UpdateProcess -FilePath 'vcpkg' -ArgumentList @('update') -Retries 1 -TimeoutSec 300
                $outdated = @(Invoke-UpdateProcess -FilePath 'vcpkg' -ArgumentList @('upgrade', '--no-dry-run') -Retries 1 -TimeoutSec 600 -SuccessExitCodes @(0, 1))
                foreach ($line in $outdated) { Write-Output $line }
            } -Tags @('cpp') -Resources @('vcpkg'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'conda' -Category 'python' -RequiresCommand 'conda' -Script {
                param([string[]]$SkipEnvs)
                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($env in @($SkipEnvs)) { if (-not [string]::IsNullOrWhiteSpace($env)) { [void]$skipSet.Add($env.Trim()) } }
                Write-Output 'Updating conda base environment...'
                Invoke-UpdateProcess -FilePath 'conda' -ArgumentList @('update', '-n', 'base', 'conda', '-y') -Retries 1 -TimeoutSec 300
                Invoke-UpdateProcess -FilePath 'conda' -ArgumentList @('update', '--all', '-y') -Retries 1 -TimeoutSec 600
            } -Tags @('python') -Resources @('conda'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'gcloud' -Category 'cloud' -RequiresCommand 'gcloud' -Script {
                param([string[]]$SkipComponents)
                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($comp in @($SkipComponents)) { if (-not [string]::IsNullOrWhiteSpace($comp)) { [void]$skipSet.Add($comp.Trim()) } }
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $gcloudPath = (Get-Command gcloud -ErrorAction SilentlyContinue).Source
                if ($gcloudPath -and $gcloudPath -match $managedPathPattern) { Write-Output 'gcloud is managed by a package manager; skipping self-update.'; return }
                Write-Output 'Updating Google Cloud SDK...'
                Invoke-UpdateProcess -FilePath 'gcloud' -ArgumentList @('components', 'update', '--quiet') -Retries 1 -TimeoutSec 600
            } -Tags @('cloud') -Resources @('gcloud'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'az' -Category 'cloud' -RequiresCommand 'az' -Script {
                param([string[]]$SkipExtensions)
                $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($ext in @($SkipExtensions)) { if (-not [string]::IsNullOrWhiteSpace($ext)) { [void]$skipSet.Add($ext.Trim()) } }
                Write-Output 'Updating Azure CLI...'
                Invoke-UpdateProcess -FilePath 'az' -ArgumentList @('upgrade', '--all', '-y') -Retries 1 -TimeoutSec 600
            } -Tags @('cloud') -Resources @('az'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'aws' -Category 'cloud' -RequiresCommand 'aws' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\|\\Python\\d+\\Scripts\\|\\pipx\\venvs\\'
                $awsPath = (Get-Command aws -ErrorAction SilentlyContinue).Source
                if ($awsPath -and $awsPath -match $managedPathPattern) {
                    try { $awsVersion = Invoke-UpdateProcess -FilePath 'aws' -ArgumentList @('--version') -Retries 0 -PassThru
                          $verText = ($awsVersion.Output | Out-String).Trim()
                          Write-Output "AWS CLI is managed ($verText); updating via package manager."; return
                    } catch { Write-Output 'aws found but version check failed.' }
                }
                Write-Output 'Updating AWS CLI via pip...'
                Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', 'awscli') -Retries 1 -TimeoutSec 300
            } -Tags @('cloud') -Resources @('aws'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'terraform' -Category 'infrastructure' -RequiresCommand 'terraform' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\|\\tfenv\\'
                $tfPath = (Get-Command terraform -ErrorAction SilentlyContinue).Source
                if ($tfPath -and $tfPath -match $managedPathPattern) { Write-Output 'Terraform is managed by a package manager; skipping.'; return }
                try {
                    $currentOut = Invoke-UpdateProcess -FilePath 'terraform' -ArgumentList @('--version') -Retries 0 -PassThru
                    $currentVer = ($currentOut.Output | Out-String) -replace '(?s).*v(\d+\.\d+\.\d+).*', '$1'
                    Write-Output "Current Terraform version: $currentVer"
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/hashicorp/terraform/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    $latestVer = $releaseData.tag_name -replace '^v'
                    if ($currentVer -ne $latestVer) {
                        Write-Output "Upgrading Terraform $currentVer -> $latestVer"
                        $osArch = if ([Environment]::Is64BitOperatingSystem) { 'amd64' } else { '386' }
                        $osName = if ($IsWindows) { 'windows' } else { 'linux' }
                        $downloadUrl = "https://releases.hashicorp.com/terraform/$latestVer/terraform_${latestVer}_${osName}_${osArch}.zip"
                        $zipPath = Join-Path ([System.IO.Path]::GetTempPath()) "terraform_$latestVer.zip"
                        Invoke-RestMethod -Uri $downloadUrl -OutFile $zipPath -TimeoutSec 60
                        $installDir = [System.IO.Path]::GetDirectoryName($tfPath)
                        Expand-Archive -LiteralPath $zipPath -DestinationPath $installDir -Force
                        Remove-Item -LiteralPath $zipPath -Force
                        Write-Output "Terraform upgraded to $latestVer"
                    } else { Write-Output "Terraform $currentVer is current." }
                } catch { Set-TaskStatus -Status 'Warn' -Reason "terraform update skipped: $($_.Exception.Message)"
                          Write-Output "terraform update skipped: $($_.Exception.Message)"
                }
            } -Tags @('infrastructure') -Resources @('terraform'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pulumi' -Category 'infrastructure' -RequiresCommand 'pulumi' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $pulumiPath = (Get-Command pulumi -ErrorAction SilentlyContinue).Source
                if ($pulumiPath -and $pulumiPath -match $managedPathPattern) { Write-Output 'Pulumi is managed by a package manager; skipping.'; return }
                Write-Output 'Upgrading Pulumi...'
                Invoke-UpdateProcess -FilePath 'pulumi' -ArgumentList @('version') -Retries 0 -TimeoutSec 30
                Invoke-UpdateProcess -FilePath 'pulumi' -ArgumentList @('upgrade') -Retries 1 -TimeoutSec 600
            } -Tags @('infrastructure') -Resources @('pulumi'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'kubectl' -Category 'infrastructure' -RequiresCommand 'kubectl' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $kubectlPath = (Get-Command kubectl -ErrorAction SilentlyContinue).Source
                if ($kubectlPath -and $kubectlPath -match $managedPathPattern) { Write-Output 'kubectl is managed by a package manager; skipping.'; return }
                try {
                    $currentOut = Invoke-UpdateProcess -FilePath 'kubectl' -ArgumentList @('version', '--client', '-o', 'json') -Retries 0 -PassThru
                    $jsonOut = ($currentOut.Output | Out-String) | ConvertFrom-Json -ErrorAction SilentlyContinue
                    $currentVer = if ($jsonOut.clientVersion.gitVersion) { $jsonOut.clientVersion.gitVersion } else { 'unknown' }
                    Write-Output "Current kubectl: $currentVer"
                    $releaseData = Invoke-RestMethod -Uri 'https://dl.k8s.io/release/stable.txt' -TimeoutSec 15 -ErrorAction Stop
                    $latestVer = ($releaseData | Out-String).Trim()
                    Write-Output "Latest stable kubectl: $latestVer"
                } catch { Write-Output "kubectl version check skipped: $($_.Exception.Message)" }
            } -Tags @('kubernetes') -Resources @('kubectl'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'helm' -Category 'infrastructure' -RequiresCommand 'helm' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $helmPath = (Get-Command helm -ErrorAction SilentlyContinue).Source
                if ($helmPath -and $helmPath -match $managedPathPattern) { Write-Output 'Helm is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'helm' -ArgumentList @('version', '--short') -Retries 0
                    Write-Output ("Current Helm: {0}" -f ($versionOut | Out-String).Trim())
                    Write-Output 'Checking for Helm updates...'
                    Invoke-UpdateProcess -FilePath 'helm' -ArgumentList @('repo', 'update') -Retries 1 -TimeoutSec 120
                } catch { Write-Output "Helm check skipped: $($_.Exception.Message)" }
            } -Tags @('kubernetes') -Resources @('helm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'hugo' -Category 'dev-tools' -RequiresCommand 'hugo' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $hugoPath = (Get-Command hugo -ErrorAction SilentlyContinue).Source
                if ($hugoPath -and $hugoPath -match $managedPathPattern) { Write-Output 'Hugo is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'hugo' -ArgumentList @('version') -Retries 0
                    Write-Output ("Current Hugo: {0}" -f ($versionOut | Select-Object -First 1 | Out-String).Trim())
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/gohugoio/hugo/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    $latestVer = $releaseData.tag_name -replace '^v'
                    Write-Output "Latest Hugo: $latestVer"
                } catch { Write-Output "Hugo check skipped: $($_.Exception.Message)" }
            } -Tags @('static-site'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'opentofu' -Category 'infrastructure' -RequiresCommand 'tofu' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $tofuPath = (Get-Command tofu -ErrorAction SilentlyContinue).Source
                if ($tofuPath -and $tofuPath -match $managedPathPattern) { Write-Output 'OpenTofu is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'tofu' -ArgumentList @('--version') -Retries 0
                    Write-Output ("Current OpenTofu: {0}" -f ($versionOut | Select-Object -First 1 | Out-String).Trim())
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/opentofu/opentofu/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    $latestVer = $releaseData.tag_name -replace '^v'
                    Write-Output "Latest OpenTofu: $latestVer"
                } catch { Write-Output "OpenTofu check skipped: $($_.Exception.Message)" }
            } -Tags @('infrastructure'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'starship' -Category 'shell' -RequiresCommand 'starship' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $starshipPath = (Get-Command starship -ErrorAction SilentlyContinue).Source
                if ($starshipPath -and $starshipPath -match $managedPathPattern) { Write-Output 'starship is managed by a package manager; skipping.'; return }
                $upgraded = $false
                foreach ($cmd in @(@('self', 'update', '-y'), @('upgrade', '--yes'), @('self-update', '-y')))
                {
                    try {
                        $result = Invoke-UpdateProcess -FilePath 'starship' -ArgumentList $cmd -Retries 0 -TimeoutSec 120 -SuccessExitCodes @(0, 1) -PassThru
                        $outText = (@($result.Output) | Out-String).Trim()
                        $errText = (@($result.Error) | Out-String).Trim()
                        $hasError = $errText -match '(?i)(error|unrecognized|usage|no such|not found)'
                        if ($result.ExitCode -eq 0 -and -not $hasError)
                        { if ($outText) { Write-Output $outText }; $upgraded = $true; break }
                    } catch { continue }
                }
                if (-not $upgraded) { Write-Output 'starship: no compatible self-update command found. Update via winget/scoop/choco.' }
            } -Tags @('shell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'zoxide' -Category 'shell' -RequiresCommand 'zoxide' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $zoxidePath = (Get-Command zoxide -ErrorAction SilentlyContinue).Source
                if ($zoxidePath -and $zoxidePath -match $managedPathPattern) { Write-Output 'zoxide is managed by a package manager; skipping.'; return }
                Write-Output 'zoxide is self-updating via package manager; run winget upgrade or scoop update to update.'
            } -Tags @('shell'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'gitleaks' -Category 'security' -RequiresCommand 'gitleaks' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $glPath = (Get-Command gitleaks -ErrorAction SilentlyContinue).Source
                if ($glPath -and $glPath -match $managedPathPattern) { Write-Output 'gitleaks is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'gitleaks' -ArgumentList @('version') -Retries 0
                    Write-Output ("Current gitleaks: {0}" -f ($versionOut | Out-String).Trim())
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/gitleaks/gitleaks/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    Write-Output "Latest gitleaks: $($releaseData.tag_name)"
                } catch { Write-Output "gitleaks version check skipped: $($_.Exception.Message)" }
            } -Tags @('security'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'trivy' -Category 'security' -RequiresCommand 'trivy' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $trivyPath = (Get-Command trivy -ErrorAction SilentlyContinue).Source
                if ($trivyPath -and $trivyPath -match $managedPathPattern) { Write-Output 'trivy is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'trivy' -ArgumentList @('--version') -Retries 0
                    Write-Output ("Current trivy: {0}" -f ($versionOut | Select-Object -First 1 | Out-String).Trim())
                    Write-Output 'Upgrading trivy...'
                    Invoke-UpdateProcess -FilePath 'trivy' -ArgumentList @('update') -Retries 1 -TimeoutSec 300
                } catch { Write-Output "trivy update skipped: $($_.Exception.Message)" }
            } -Tags @('security') -Resources @('trivy'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'packer' -Category 'infrastructure' -RequiresCommand 'packer' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $packerPath = (Get-Command packer -ErrorAction SilentlyContinue).Source
                if ($packerPath -and $packerPath -match $managedPathPattern) { Write-Output 'Packer is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'packer' -ArgumentList @('--version') -Retries 0
                    $currentVer = ($versionOut | Out-String).Trim()
                    Write-Output "Current Packer: $currentVer"
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/hashicorp/packer/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    Write-Output "Latest Packer: $($releaseData.tag_name)"
                } catch { Write-Output "Packer check skipped: $($_.Exception.Message)" }
            } -Tags @('infrastructure'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'nvm' -Category 'javascript' -RequiresCommand 'nvm' -Script {
                $managedPathPattern = '\\scoop\\apps\\|\\chocolatey\\lib\\|\\Microsoft\\WinGet\\Packages\\'
                $nvmPath = (Get-Command nvm -ErrorAction SilentlyContinue).Source
                if ($nvmPath -and $nvmPath -match $managedPathPattern) { Write-Output 'nvm is managed by a package manager; skipping.'; return }
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'nvm' -ArgumentList @('version') -Retries 0
                    Write-Output ("Current nvm: {0}" -f ($versionOut | Out-String).Trim())
                    Write-Output 'Install nvm updates via winget: winget upgrade --id CoreyButler.NVMforWindows'
                } catch { Write-Output "nvm check skipped: $($_.Exception.Message)" }
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'devcontainer' -Category 'dev-tools' -RequiresCommand 'devcontainer' -Script {
                try {
                    $versionOut = Invoke-UpdateProcess -FilePath 'devcontainer' -ArgumentList @('--version') -Retries 0
                    Write-Output ("Current devcontainer CLI: {0}" -f ($versionOut | Out-String).Trim())
                    $releaseData = Invoke-RestMethod -Uri 'https://api.github.com/repos/devcontainers/cli/releases/latest' -TimeoutSec 15 -ErrorAction Stop
                    Write-Output "Latest: $($releaseData.tag_name)"
                } catch { Write-Output "devcontainer check skipped: $($_.Exception.Message)" }
            } -Tags @('dev'))) | Out-Null

        $cleanupScript = {
            param([int]$Days, [bool]$Deep, [bool]$SkipDestructive)

            function Test-SafeCleanupRoot
            {
                param([Parameter(Mandatory)][string]$Path)
                try
                {
                    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
                    if (-not $item.PSIsContainer)
                    { return $false
                    }
                    $fullPath = [System.IO.Path]::GetFullPath($item.FullName).TrimEnd('\')
                    $rootPath = ([System.IO.Path]::GetPathRoot($fullPath)).TrimEnd('\')
                    if ([string]::IsNullOrWhiteSpace($fullPath) -or $fullPath -eq $rootPath)
                    { return $false
                    }

                    $allowedRoots = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    if ($env:TEMP)
                    { [void]$allowedRoots.Add(([System.IO.Path]::GetFullPath($env:TEMP)).TrimEnd('\'))
                    }
                    if ($IsWindows)
                    { [void]$allowedRoots.Add(([System.IO.Path]::GetFullPath('C:\Windows\Temp')).TrimEnd('\'))
                    }
                    return $allowedRoots.Contains($fullPath)
                } catch
                { return $false
                }
            }

            $cutoff = (Get-Date).AddDays(-$Days)
            $paths = @($env:TEMP)
            if ($IsWindows -and (Test-Path -LiteralPath 'C:\Windows\Temp'))
            { $paths += 'C:\Windows\Temp'
            }

            foreach ($path in ($paths | Where-Object { $_ } | Select-Object -Unique))
            {
                if (-not (Test-SafeCleanupRoot -Path $path))
                { Write-Output "Skipping unsafe cleanup path: $path"; continue
                }
                if ($SkipDestructive)
                { Write-Output "Skipping temp cleanup because -SkipDestructive is set: $path"; continue
                }

                Write-Output "Cleaning temp files older than $Days day(s): $path"
                Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue |
                    Where-Object {
                        $_.LastWriteTime -lt $cutoff -and
                        -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
                        $_.FullName -notmatch '\\WinGet(\\|$)'
                    } |
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }

        if ($IsWindows)
        {
            Clear-DnsClientCache -ErrorAction SilentlyContinue
            if (-not $SkipDestructive)
            { Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            }
            if ($Deep -and -not $SkipDestructive)
            {
                Invoke-UpdateProcess -FilePath 'DISM.exe' -ArgumentList @('/Online', '/Cleanup-Image', '/StartComponentCleanup') -TimeoutSec 3600
                Clear-DeliveryOptimizationCache -Force -ErrorAction SilentlyContinue
            } elseif ($Deep -and $SkipDestructive)
            {
                Write-Output 'Skipping deep cleanup because -SkipDestructive is set.'
            }

            # Stale binary detection: find orphaned .exe files in PATH not tracked by any package manager
            Write-Output 'Checking for orphaned binaries in PATH...'
            $pathDirs = @($env:Path -split ';' | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
            $orphanCount = 0
            $knownManagers = @('scoop\apps', 'chocolatey\lib', 'Microsoft\WinGet\Packages', 'pipx\venvs', 'Python\Scripts', 'Python3', '.cargo\bin', 'node_modules\.bin', '.local\bin', '.dotnet\tools', 'mise', 'uv\bin', 'volta\bin')
            $excludePrefixes = @('api-ms-win-', 'ext-ms-win-', 'concrt', 'msvcp', 'vcruntime', 'msvcrt')
            foreach ($dir in $pathDirs)
            {
                $exes = @(Get-ChildItem -LiteralPath $dir -Filter '*.exe' -Force -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer })
                foreach ($exe in $exes)
                {
                    $isManaged = $false
                    foreach ($mgr in $knownManagers) { if ($exe.FullName -match [regex]::Escape($mgr)) { $isManaged = $true; break } }
                    $isSystemPath = $exe.DirectoryName -match '\\system32$|\\system$|\\windows\\'
                    $excluded = ($excludePrefixes | Where-Object { $exe.Name -like "$_*" }).Count -gt 0
                    if (-not $isManaged -and -not $isSystemPath -and -not $excluded -and $exe.LastWriteTime -lt $cutoff)
                    { $orphanCount++ }
                }
            }
            if ($orphanCount -gt 0) { Write-Output "Stale binary scan: $orphanCount orphaned .exe(s) older than $Days day(s) found in PATH. Run with -DeepClean to remove them." }
            else { Write-Output 'Stale binary scan: no orphaned binaries found.' }
        }
    }
    $tasks.Add((New-UpdateTask -Name 'cleanup' -Category 'maintenance' -Disabled:$SkipCleanup -DisabledReason 'disabled by -SkipCleanup' -TimeoutSec 3600 -Script $cleanupScript -Tags @('maintenance'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'self-update' -Category 'maintenance' -Script {
                try {
                    Write-Output 'Checking for script updates...'
                    $repoUrl = 'https://api.github.com/repos/YoshKoz/updateEverything/releases/latest'
                    $releaseData = Invoke-RestMethod -Uri $repoUrl -TimeoutSec 15 -ErrorAction Stop
                    $latestTag = ($releaseData.tag_name -replace '^v').Trim()
                    $currentTag = ($script:Version -replace '-.*$').Trim()
                    Write-Output "Current: $currentTag | Latest: $latestTag"
                    if ($latestTag -and $currentTag -and $latestTag -ne $currentTag)
                    {
                        Write-Output "Update available: $currentTag → $latestTag"
                        Write-Output "Download from: https://github.com/YoshKoz/updateEverything/releases/tag/v$latestTag"
                    } else { Write-Output 'Already at latest version.' }
                } catch {
                    if ($_.Exception -is [System.Net.WebException] -and $_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
                        Write-Output 'Self-update check skipped: no releases found on GitHub.'
                    } elseif ($_.Exception -match '(404|Not Found)') {
                        Write-Output 'Self-update check skipped: repository not found or no releases configured.'
                    } else {
                        Write-Output "Self-update check skipped: $($_.Exception.Message)"
                    }
                }
            } -Tags @('self'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'windows-features' -Category 'system' -RequiresAdmin -Script {
                param([string[]]$WindowsOptionalFeatures)
                if ($WindowsOptionalFeatures.Count -eq 0) { Write-Output 'No WindowsOptionalFeatures configured in update-config.json.'; return }
                $enabled = 0; $skipped = 0
                foreach ($feature in $WindowsOptionalFeatures)
                {
                    if ([string]::IsNullOrWhiteSpace($feature)) { continue }
                    $state = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction SilentlyContinue
                    if ($state -and $state.State -ne 'Enabled')
                    {
                        Write-Output "Enabling Windows Feature: $feature"
                        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -LimitAccess -ErrorAction SilentlyContinue | Out-Null
                        $enabled++
                    } else { $skipped++ }
                }
                Write-Output "Windows Features: $enabled enabled, $skipped already present."
            } -Tags @('windows', 'system'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'appx-repair' -Category 'system' -Script {
                Write-Output 'Re-registering packaged app manifests for known-broken store apps...'
                $appxPackages = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
                    $_.SignatureKind -eq 'Store' -and -not $_.IsFramework -and $_.Name -match '(Microsoft\.WindowsStore|Microsoft\.Store|Microsoft\.WindowsCalculator|Microsoft\.Windows\.Photos|Microsoft\.Windows\.Camera|Microsoft\.People|Microsoft\.Office|Microsoft\.MSPaint|Microsoft\.ScreenSketch|Microsoft\.WindowsNotepad|Microsoft\.WindowsTerminal)'
                })
                $repaired = 0
                foreach ($pkg in $appxPackages)
                {
                    try {
                        $manifest = Get-AppxPackageManifest -Package $pkg -ErrorAction SilentlyContinue
                        if ($manifest)
                        { Add-AppxPackage -Register -DisableDevelopmentMode -ErrorAction SilentlyContinue "$($pkg.InstallLocation)\AppxManifest.xml" *>$null; $repaired++ }
                    } catch { Write-Output "AppX re-registration failed for $($pkg.Name): $($_.Exception.Message)" }
                }
                Write-Output "AppX re-registration: $repaired package(s) re-registered."
            } -Tags @('windows', 'store'))) | Out-Null

    foreach ($task in $tasks)
    {
        switch ($task.Id)
        {
            'winget'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages            = $wingetSkip
                    ProtectedPackages       = $wingetProtected
                    BypassProtection        = [bool]$BypassProtection
                    UseSilentInstallers     = [bool]$useSilentInstallers
                    UnknownVersionStatePath = Join-Path $script:StateDir 'unknown-versions.json'
                } -Force
            }
            'chocolatey'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages      = $chocoSkip
                    ProtectedPackages = $chocoProtected
                    BypassProtection  = [bool]$BypassProtection
                } -Force
            }
            'store-apps'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages       = $storeSkip
                    ProtectedPackages  = $storeProtected
                    BypassProtection   = [bool]$BypassProtection
                    UseSilentInstallers = [bool]$useSilentInstallers
                } -Force
            }
            'pip'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $pipSkip } -Force
            }
            'pip-health'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ IgnorePackages = $pipIgnoreHealth } -Force
            }
            'npm'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $npmSkip } -Force
            }
            'vcpkg'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $vcpkgSkipPackages } -Force
            }
            'conda'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipEnvs = $condaSkipEnvs } -Force
            }
            'gcloud'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipComponents = $gcloudSkipComponents } -Force
            }
            'az'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipExtensions = $azSkipExtensions } -Force
            }
            'cleanup'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ Days = $tempCleanupDays; Deep = [bool]$DeepClean; SkipDestructive = [bool]$SkipDestructive } -Force
            }
            'ollama-models'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ CommandTimeoutSec = $OllamaTimeoutSec } -Force
            }
            'windows-features'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ WindowsOptionalFeatures = $windowsOptionalFeatures } -Force
            }
            'cross-manager'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ FallbackApps = $crossManagerFallback } -Force
            }
            default
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
            }
        }
    }

    return @($tasks)
}

function Get-FilteredTasks
{
    param(
        [Parameter(Mandatory)][object[]]$Tasks,
        [Parameter(Mandatory)][bool]$IsAdmin
    )

    $onlyPatterns = @(ConvertTo-FilterList $Only)
    $skipPatterns = @()
    $skipPatterns += ConvertTo-FilterList $script:Config.SkipManagers
    $skipPatterns += ConvertTo-FilterList $Skip
    if ($FastMode -or $UltraFast)
    { $skipPatterns += ConvertTo-FilterList $script:Config.FastModeSkip
    }
    if ($UltraFast)
    { $skipPatterns += ConvertTo-FilterList $script:Config.UltraFastSkip
    }

    $planned = [System.Collections.Generic.List[object]]::new()
    $skipped = [System.Collections.Generic.List[object]]::new()

    foreach ($task in $Tasks)
    {
        $reason = $null
        if ($onlyPatterns.Count -gt 0 -and -not (Test-NameMatch -Task $task -Patterns $onlyPatterns))
        {
            $reason = 'not selected by -Only'
        } elseif ($skipPatterns.Count -gt 0 -and (Test-NameMatch -Task $task -Patterns $skipPatterns))
        {
            $reason = 'skipped by filter'
        } elseif ($task.Disabled)
        {
            $reason = if ($task.DisabledReason)
            { $task.DisabledReason
            } else
            { 'disabled'
            }
        } elseif ($task.RequiresAdmin -and -not $IsAdmin)
        {
            $reason = 'requires Administrator'
        } else
        {
            foreach ($command in $task.RequiresCommand)
            {
                if (-not (Test-Command $command))
                { $reason = "missing command: $command"; break
                }
            }
        }

        if (-not $reason -and $SkipSucceededWithinHours -gt 0)
        {
            try
            {
                if (Test-Path -LiteralPath $script:DefaultJsonSummaryPath)
                {
                    $prev = Get-Content -LiteralPath $script:DefaultJsonSummaryPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    $prevResult = @($prev.Results | Where-Object { $_.Id -eq $task.Id }) | Select-Object -Last 1
                    if ($prevResult -and $prevResult.Status -eq 'Succeeded')
                    {
                        $prevTime = [datetime]::Parse($prev.StartedAt)
                        $hoursSince = ((Get-Date) - $prevTime).TotalHours
                        if ($hoursSince -lt $SkipSucceededWithinHours)
                        {
                            $reason = "succeeded $([Math]::Floor($hoursSince))h ago (within -SkipSucceededWithinHours $SkipSucceededWithinHours)"
                        }
                    }
                }
            } catch { }
        }

        if ($reason)
        {
            $skipped.Add([pscustomobject]@{ Name = $task.Name; Id = $task.Id; Category = $task.Category; Status = 'Skipped'; Reason = $reason }) | Out-Null
        } else
        {
            $planned.Add($task) | Out-Null
        }
    }

    return [pscustomobject]@{ Planned = @($planned); Skipped = @($skipped) }
}

function Split-SkippedTasksForDisplay
{
    param([object[]]$Skipped)

    $hidden = @($Skipped | Where-Object {
            $_.Reason -like 'missing command:*' -or
            $_.Reason -like 'opt-in via -*' -or
            $_.Reason -like 'use -*'
        })
    $visible = @($Skipped | Where-Object { $hidden -notcontains $_ })

    [pscustomobject]@{
        Visible = $visible
        Hidden  = $hidden
    }
}

function Show-TaskList
{
    param([object[]]$Planned, [object[]]$Skipped)
    Write-Host ''
    Write-Host 'Planned tasks' -ForegroundColor Cyan
    if ($Planned.Count -eq 0)
    { Write-Host '  none' -ForegroundColor DarkGray
    } else
    { $Planned | Sort-Object Category, Name | Format-Table Name, Category, TimeoutSec, RequiresAdmin, Resources -AutoSize
    }

    Write-Host ''
    $splitSkipped = Split-SkippedTasksForDisplay -Skipped $Skipped
    $skippedForDisplay = if ($ShowSkipped) { @($Skipped) } else { @($splitSkipped.Visible) }
    Write-Host 'Skipped tasks' -ForegroundColor DarkGray
    if (@($skippedForDisplay).Count -eq 0)
    {
        if ($splitSkipped.Hidden.Count -gt 0 -and -not $ShowSkipped)
        { Write-Host ("  {0} optional skip(s) hidden; rerun with -ShowSkipped to list them." -f $splitSkipped.Hidden.Count) -ForegroundColor DarkGray
        } else
        { Write-Host '  none' -ForegroundColor DarkGray
        }
    } else
    {
        $skippedForDisplay | Sort-Object Category, Name | Format-Table Name, Category, Reason -AutoSize
        if ($splitSkipped.Hidden.Count -gt 0 -and -not $ShowSkipped)
        { Write-Host ("  {0} optional skip(s) hidden; rerun with -ShowSkipped to list them." -f $splitSkipped.Hidden.Count) -ForegroundColor DarkGray
        }
    }
}

function New-TaskResult
{
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
        OutputPreview   = @($Output | Select-Object -First 40)
    }
}

function Start-UpdateTaskJob
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '', Justification = 'the thread job declares its own param() block and is fed by -ArgumentList')]
    param(
        [Parameter(Mandatory)]$Task,
        [Parameter(Mandatory)][int]$Attempt
    )

    $scriptText = $Task.Script.ToString()
    $argumentMap = if ($Task.PSObject.Properties['Arguments'])
    { $Task.Arguments
    } else
    { @{}
    }
    $taskTimeoutSec = [int]$Task.TimeoutSec
    $helperFunctionDefinitions = @()
    foreach ($functionName in @('ConvertFrom-WingetUpgradeOutput'))
    {
        $functionCommand = Get-Command -Name $functionName -CommandType Function -ErrorAction SilentlyContinue
        if ($functionCommand)
        { $helperFunctionDefinitions += "function $functionName {`n$($functionCommand.ScriptBlock.ToString())`n}"
        }
    }

    Start-ThreadJob -Name $Task.Name -ScriptBlock {
        param($TaskName, $TaskId, $Category, $ScriptText, $ArgumentMap, $Attempt, $TaskTimeoutSec, $HelperFunctionDefinitions)

        $ErrorActionPreference = 'Continue'
        $global:LASTEXITCODE = 0
        $start = Get-Date
        $output = [System.Collections.Generic.List[string]]::new()
        $exitCode = 0
        $status = 'Succeeded'
        $reason = $null
        $script:TaskStatusOverride = $null
        $script:TaskReasonOverride = $null

        function ConvertTo-OutputLines
        {
            param([AllowNull()][string]$Text)
            if ([string]::IsNullOrEmpty($Text))
            { return @()
            }
            return @($Text -split "\r?\n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }

        function ConvertTo-CmdArgument
        {
            param([AllowNull()][string]$Argument)
            if ($null -eq $Argument)
            { return '""'
            }
            $s = [string]$Argument
            if ($s -eq '')
            { return '""'
            }
            if ($s -notmatch '[\s&()\[\]{}^=;!''+,`~|<>\"]')
            { return $s
            }
            return '"' + ($s -replace '"', '""') + '"'
        }

        function Get-ToolCommandPath
        {
            param([Parameter(Mandatory)][string]$Name)
            if ($Name -in @('pwsh', 'pwsh.exe'))
            {
                $pwshFileName = if ($IsWindows)
                { 'pwsh.exe'
                } else
                { 'pwsh'
                }
                $candidates = @(
                    (Join-Path $PSHOME $pwshFileName),
                    'C:\Program Files\PowerShell\7\pwsh.exe',
                    'C:\Program Files\PowerShell\7-preview\pwsh.exe'
                ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
                if ($candidates.Count -gt 0)
                { return ($candidates | Select-Object -First 1)
                }
            }
            if ($Name -eq 'scoop')
            {
                $candidates = @(
                    (Join-Path $env:USERPROFILE 'scoop\shims\scoop.ps1'),
                    (Join-Path $env:USERPROFILE 'scoop\shims\scoop.cmd')
                ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
                if ($candidates.Count -gt 0)
                { return ($candidates | Select-Object -First 1)
                }
            }
            if ($Name -eq 'code')
            {
                $cmd = Get-Command -Name 'code' -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($cmd)
                {
                    if ($cmd.Source)
                    { return $cmd.Source
                    }
                    if ($cmd.Path)
                    { return $cmd.Path
                    }
                }
                $candidates = @(
                    (Join-Path $env:LOCALAPPDATA 'Programs\Microsoft VS Code\bin\code.cmd'),
                    'C:\Program Files\Microsoft VS Code\bin\code.cmd',
                    'C:\Program Files (x86)\Microsoft VS Code\bin\code.cmd'
                ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
                return ($candidates | Select-Object -First 1)
            }
            if ($Name -eq 'python')
            {
                $candidates = @(@(
                    'C:\Program Files\Python314\python.exe',
                    'C:\Program Files\Python313\python.exe',
                    'C:\Program Files\Python312\python.exe',
                    (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python314\python.exe'),
                    (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python313\python.exe'),
                    (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe')
                ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
                if ($candidates.Count -gt 0)
                { return ($candidates | Select-Object -First 1)
                }
            }
            $command = Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($command)
            {
                if ($command.Source)
                { return $command.Source
                }
                if ($command.Path)
                { return $command.Path
                }
            }
            return $null
        }

        function Resolve-UpdateProcessCommand
        {
            param(
                [Parameter(Mandatory)][string]$FilePath,
                [string[]]$ArgumentList = @()
            )

            $source = $FilePath
            $command = $null
            $isPathLike = [System.IO.Path]::IsPathRooted($FilePath) -or $FilePath -match '[\\/]'
            $toolPath = if (-not $isPathLike)
            { Get-ToolCommandPath -Name $FilePath
            } else
            { $null
            }
            if ($toolPath)
            {
                $source = $toolPath
            } elseif ($FilePath -eq 'code')
            {
                $source = Get-ToolCommandPath -Name 'code'
            } elseif (-not $isPathLike -or -not (Test-Path -LiteralPath $FilePath -ErrorAction SilentlyContinue))
            {
                $command = Get-Command -Name $FilePath -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($command)
                {
                    if ($command.Source)
                    { $source = $command.Source
                    } elseif ($command.Path)
                    { $source = $command.Path
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace($source))
            { $source = $FilePath
            }
            $extension = [System.IO.Path]::GetExtension($source)
            if ($source -eq $FilePath -and $extension -eq '' -and $FilePath -notmatch '[\\/]')
            { throw [System.Management.Automation.CommandNotFoundException]::new("Command not found: $FilePath")
            }

            if ($extension -ieq '.ps1')
            {
                $pwshCommand = Get-Command -Name 'pwsh' -ErrorAction SilentlyContinue | Select-Object -First 1
                $pwshPath = if ($pwshCommand -and $pwshCommand.Source)
                { $pwshCommand.Source
                } else
                { 'pwsh'
                }
                return [pscustomobject]@{ FilePath = $pwshPath; ArgumentList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $source) + @($ArgumentList) }
            }

            if ($extension -in @('.cmd', '.bat'))
            {
                # Use cmd.exe with discrete ArgumentList entries. Do NOT pre-escape quotes with backslashes:
                # cmd.exe treats " literally, which caused paths like "C:\Program Files\...\code.cmd" to fail.
                return [pscustomobject]@{
                    FilePath = if ($env:ComSpec)
                    { $env:ComSpec
                    } else
                    { 'cmd.exe'
                    }
                    ArgumentList = @('/d', '/c', 'call', $source) + @($ArgumentList)
                }
            }

            return [pscustomobject]@{ FilePath = $source; ArgumentList = @($ArgumentList) }
        }

        function Invoke-UpdateProcess
        {
            param(
                [Parameter(Mandatory)][string]$FilePath,
                [string[]]$ArgumentList = @(),
                [int]$TimeoutSec = 0,
                [int[]]$SuccessExitCodes = @(0),
                [int]$Retries = 0,
                [switch]$PassThru
            )

            $effectiveTimeoutSec = if ($TimeoutSec -gt 0)
            { $TimeoutSec
            } else
            { [Math]::Max(30, $TaskTimeoutSec - 5)
            }
            $attemptNumber = 0
            $lastOutput = @()
            $lastExitCode = 0
            $lastTimedOut = $false

            function Format-ProcessOutputTail
            {
                param([string[]]$Output)
                $tail = @($Output |
                        ForEach-Object { ([string]$_).Trim() } |
                        Where-Object { $_ } |
                        Select-Object -Last 3)
                if ($tail.Count -eq 0)
                { return $null
                }
                $text = ($tail -join ' | ')
                if ($text.Length -gt 500)
                { return ($text.Substring(0, 497) + '...')
                }
                return $text
            }

            function New-ProcessResult
            {
                param(
                    [int]$ExitCode,
                    [string[]]$Output,
                    [bool]$TimedOut
                )

                [pscustomobject]@{
                    Command   = $FilePath
                    Arguments = @($ArgumentList)
                    ExitCode  = $ExitCode
                    TimedOut  = $TimedOut
                    Output    = @($Output)
                }
            }

            do
            {
                $attemptNumber++
                $process = $null
                $lastTimedOut = $false
                try
                {
                    $resolvedCommand = Resolve-UpdateProcessCommand -FilePath $FilePath -ArgumentList $ArgumentList
                    $psi = [System.Diagnostics.ProcessStartInfo]::new()
                    $psi.FileName = $resolvedCommand.FilePath
                    foreach ($argument in @($resolvedCommand.ArgumentList))
                    { [void]$psi.ArgumentList.Add([string]$argument)
                    }
                    $psi.UseShellExecute = $false
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError = $true
                    $psi.RedirectStandardInput = $true
                    $psi.CreateNoWindow = $true

                    $process = [System.Diagnostics.Process]::Start($psi)
                    if (-not $process)
                    { throw "Process did not start: $FilePath"
                    }
                    $process.StandardInput.Close()
                    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
                    $stderrTask = $process.StandardError.ReadToEndAsync()
                    $exited = $process.WaitForExit($effectiveTimeoutSec * 1000)

                    if (-not $exited)
                    {
                        try
                        { $process.Kill($true)
                        } catch
                        { try
                            { $process.Kill()
                            } catch
                            {
                            }
                        }
                        try
                        { $process.WaitForExit(5000) | Out-Null
                        } catch
                        {
                        }
                        $lastExitCode = 124
                        $lastTimedOut = $true
                        $lastOutput = @("$FilePath $($ArgumentList -join ' ') timed out after ${effectiveTimeoutSec}s")
                    } else
                    {
                        $lastExitCode = [int]$process.ExitCode
                        $stdout = $stdoutTask.GetAwaiter().GetResult()
                        $stderr = $stderrTask.GetAwaiter().GetResult()
                        $lastOutput = @()
                        $lastOutput += ConvertTo-OutputLines $stdout
                        $lastOutput += ConvertTo-OutputLines $stderr
                    }
                } catch
                {
                    $lastExitCode = if ($_.Exception -is [System.Management.Automation.CommandNotFoundException])
                    { 127
                    } else
                    { 1
                    }
                    $lastOutput = @($_.Exception.Message)
                } finally
                {
                    if ($process)
                    { $process.Dispose()
                    }
                }

                if ($SuccessExitCodes -contains $lastExitCode)
                {
                    $global:LASTEXITCODE = 0
                    if ($PassThru)
                    { return (New-ProcessResult -ExitCode $lastExitCode -Output $lastOutput -TimedOut $lastTimedOut)
                    }
                    return @($lastOutput)
                }

                if (-not $PassThru)
                {
                    foreach ($line in $lastOutput)
                    { Write-Output $line
                    }
                }
                if ($lastExitCode -ne 127 -and $attemptNumber -le $Retries)
                {
                    $backoffSec = [Math]::Min(60, [Math]::Pow(2, $attemptNumber - 1) * 3)
                    if (-not $PassThru)
                    { Write-Output ("Retrying {0} after exit code {1} (attempt {2}/{3}; backoff {4}s)" -f $FilePath, $lastExitCode, ($attemptNumber + 1), ($Retries + 1), $backoffSec)
                    }
                    Start-Sleep -Seconds $backoffSec
                }
            } while ($attemptNumber -le $Retries)

            $global:LASTEXITCODE = $lastExitCode
            if ($PassThru)
            { return (New-ProcessResult -ExitCode $lastExitCode -Output $lastOutput -TimedOut $lastTimedOut)
            }

            $detail = Format-ProcessOutputTail -Output $lastOutput
            if ($detail)
            { throw "$FilePath failed with exit code $lastExitCode. Last output: $detail"
            }
            throw "$FilePath failed with exit code $lastExitCode"
        }

        function Set-TaskStatus
        {
            param(
                [ValidateSet('Succeeded', 'Warn', 'Partial', 'Failed')]
                [string]$Status,
                [string]$Reason
            )

            $script:TaskStatusOverride = $Status
            if ($Reason)
            { $script:TaskReasonOverride = $Reason
            }
        }

        foreach ($definition in @($HelperFunctionDefinitions))
        {
            if (-not [string]::IsNullOrWhiteSpace($definition))
            { . ([scriptblock]::Create($definition))
            }
        }

        try
        {
            $block = [scriptblock]::Create($ScriptText)
            & $block @ArgumentMap 2>&1 | ForEach-Object {
                $line = if ($null -eq $_)
                { ''
                } else
                { ([string]$_).Replace([string][char]0, '').Trim()
                }
                $line = [regex]::Replace($line, '\x1b\[[0-9;?]*[a-zA-Z]', '')
                $line = [regex]::Replace($line, '\x1b\[[0-9;]*m', '')
                if ($line.Trim())
                { $output.Add($line) | Out-Null
                }
            }
        } catch
        {
            $status = 'Failed'
            $exitCode = if ($global:LASTEXITCODE -and [int]$global:LASTEXITCODE -ne 0)
            { [int]$global:LASTEXITCODE
            } else
            { 1
            }
            $reason = $_.Exception.Message
            $output.Add($_.Exception.Message) | Out-Null
        }
        if ($status -eq 'Succeeded' -and $script:TaskStatusOverride)
        {
            $status = $script:TaskStatusOverride
            $reason = $script:TaskReasonOverride
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
    } -ArgumentList $Task.Name, $Task.Id, $Task.Category, $scriptText, $argumentMap, $Attempt, $taskTimeoutSec, $helperFunctionDefinitions
}

function Test-ResourcesAvailable
{
    param($Task, [object[]]$Running)
    if (-not $Task.Resources -or $Task.Resources.Count -eq 0)
    { return $true
    }
    $runningResources = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @($Running))
    {
        foreach ($resource in @($item.Task.Resources))
        { if ($resource)
            { [void]$runningResources.Add($resource)
            }
        }
    }
    foreach ($resource in @($Task.Resources))
    { if ($runningResources.Contains($resource))
        { return $false
        }
    }
    return $true
}

function Invoke-TaskQueue
{
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][object[]]$Tasks,
        [Parameter(Mandatory)][int]$Throttle
    )

    $queue = [System.Collections.Generic.List[object]]::new()
    foreach ($task in $Tasks)
    { $queue.Add([pscustomobject]@{ Task = $task; Attempt = 1 }) | Out-Null
    }

    $running = @()
    $results = [System.Collections.Generic.List[object]]::new()

    while ($queue.Count -gt 0 -or $running.Count -gt 0)
    {
        while ($running.Count -lt $Throttle -and $queue.Count -gt 0)
        {
            $indexToStart = -1
            for ($i = 0; $i -lt $queue.Count; $i++)
            {
                if (Test-ResourcesAvailable -Task $queue[$i].Task -Running $running)
                { $indexToStart = $i; break
                }
            }
            if ($indexToStart -lt 0)
            { break
            }

            $entry = $queue[$indexToStart]
            $queue.RemoveAt($indexToStart)
            $task = $entry.Task

            if (-not $PSCmdlet.ShouldProcess($task.Name, 'Run update task'))
            {
                $results.Add((New-TaskResult -Task $task -Status 'Skipped' -Reason 'ShouldProcess declined')) | Out-Null
                continue
            }

            Write-Status ("[{0}] starting (attempt {1})" -f $task.Name, $entry.Attempt) -Level Info
            $job = Start-UpdateTaskJob -Task $task -Attempt $entry.Attempt
            $running += [pscustomobject]@{ Job = $job; Task = $task; Attempt = $entry.Attempt; Started = Get-Date }
        }

        foreach ($item in @($running))
        {
            $job = $item.Job
            $task = $item.Task
            $timedOut = ((Get-Date) - $item.Started).TotalSeconds -gt $task.TimeoutSec

            if ($timedOut)
            {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                $result = New-TaskResult -Task $task -Status 'TimedOut' -ExitCode 124 -DurationSeconds $task.TimeoutSec -Attempts $item.Attempt -Reason "timeout after $($task.TimeoutSec)s"
                $results.Add($result) | Out-Null
                Write-Status ("[{0}] timed out after {1}s" -f $task.Name, $task.TimeoutSec) -Level Warning
                $running = @($running | Where-Object { $_.Job.Id -ne $job.Id })
                continue
            }

            if ($job.State -notin @('Completed', 'Failed', 'Stopped'))
            { continue
            }

            $raw = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $running = @($running | Where-Object { $_.Job.Id -ne $job.Id })

            $payload = $raw | Where-Object { $_.PSObject.Properties['Status'] } | Select-Object -Last 1
            $output = @($raw | Where-Object { -not $_.PSObject.Properties['Status'] } | ForEach-Object { $_.ToString() })

            if (-not $payload)
            {
                $payload = [pscustomobject]@{
                    Name = $task.Name; Id = $task.Id; Category = $task.Category; Status = 'Failed'; ExitCode = 1
                    DurationSeconds = [Math]::Round(((Get-Date) - $item.Started).TotalSeconds, 2)
                    Attempts = $item.Attempt; Reason = "job ended in state $($job.State)"; Output = $output
                }
            }

            $combinedOutput = @()
            if ($payload.PSObject.Properties['Output'])
            { $combinedOutput += @($payload.Output)
            }
            $combinedOutput += $output

            $result = New-TaskResult -Task $task -Status $payload.Status -ExitCode $payload.ExitCode -DurationSeconds $payload.DurationSeconds -Attempts $payload.Attempts -Output $combinedOutput -Reason $payload.Reason

            if ($result.Status -eq 'Failed' -and $result.ExitCode -ne 124 -and $item.Attempt -le $RetryCount)
            {
                Write-Status ("[{0}] failed; queueing retry {1}/{2}" -f $task.Name, $item.Attempt, $RetryCount) -Level Warning
                foreach ($line in $combinedOutput)
                { Write-UpdateLog -Message ("[{0}] attempt {1}: {2}" -f $task.Name, $item.Attempt, $line) -Level Warning
                }
                $queue.Add([pscustomobject]@{ Task = $task; Attempt = ($item.Attempt + 1) }) | Out-Null
                continue
            }

            $results.Add($result) | Out-Null
            if ($result.Status -eq 'Succeeded')
            { Write-Status ("[{0}] ok ({1:N1}s)" -f $task.Name, $result.DurationSeconds) -Level Success
            } else
            { Write-Status ("[{0}] {1}: {2}" -f $task.Name, $result.Status, $result.Reason) -Level Warning
            }

            if ($combinedOutput.Count -gt 0)
            {
                foreach ($line in $combinedOutput)
                { Write-UpdateLog -Message ("[{0}] {1}" -f $task.Name, $line) -Level Muted
                }
                if (-not $Quiet)
                {
                    Write-Host "  [$($task.Name) output]" -ForegroundColor DarkGray
                    foreach ($line in $combinedOutput)
                    { Write-Host "    $line" -ForegroundColor DarkGray
                    }
                }
            } elseif (-not $Quiet)
            {
                Write-Host "  [$($task.Name) output] (no output)" -ForegroundColor DarkGray
            }
        }

        if ($running.Count -gt 0)
        {
            $done = $results.Count
            $total = $Tasks.Count
            $runningNames = ($running | ForEach-Object { $_.Task.Name }) -join ', '
            $percent = if ($total -gt 0)
            { [Math]::Min(99, [int](($done / $total) * 100))
            } else
            { 100
            }
            Write-Progress -Activity 'Update Everything' -Status "$done/$total done. Running: $runningNames" -PercentComplete $percent
            Start-Sleep -Milliseconds 250
        }
    }

    Write-Progress -Activity 'Update Everything' -Completed
    return @($results)
}

function Save-RunSummary
{
    param(
        [object[]]$Results,
        [object[]]$Skipped,
        [object[]]$Planned,
        [object[]]$Notes = @()
    )

    $duration = [Math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 2)
    $summary = [ordered]@{
        Version          = $script:Version
        RunId            = $script:RunId
        StartedAt        = $script:StartTime.ToString('o')
        FinishedAt       = (Get-Date).ToString('o')
        DurationSeconds  = $duration
        DryRun           = [bool]$script:IsSimulation
        FastMode         = [bool]$FastMode
        UltraFast        = [bool]$UltraFast
        ParallelThrottle = $ParallelThrottle
        LogPath          = $script:LogPath
        PlannedCount     = $Planned.Count
        SucceededCount   = @($Results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -eq 'Succeeded' }).Count
        WarningCount     = @($Results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -in @('Warn', 'Partial') }).Count
        FailedCount      = @($Results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -in @('Failed', 'TimedOut') }).Count
        SkippedCount     = $Skipped.Count
        Results          = @($Results)
        Skipped          = @($Skipped)
        Notes            = @($Notes)
    }

    $json = $summary | ConvertTo-Json -Depth 8
    $summaryWritten = $false
    try
    {
        $summaryDir = Split-Path -Parent $script:JsonSummaryPath
        if ($summaryDir -and -not (Test-Path -LiteralPath $summaryDir))
        {
            New-Item -Path $summaryDir -ItemType Directory -Force -WhatIf:$false -ErrorAction Stop | Out-Null
        }
        Set-Content -LiteralPath $script:JsonSummaryPath -Value $json -Encoding utf8 -WhatIf:$false -ErrorAction Stop
        $summaryWritten = $true
    } catch
    {
        Write-Status "Could not write JSON summary: $($_.Exception.Message)" -Level Warning
    }

    $summary['SummaryWritten'] = $summaryWritten
    return [pscustomobject]$summary
}

function Show-WhatChanged
{
    param([Parameter(Mandatory)]$CurrentSummary)
    if (-not (Test-Path -LiteralPath $script:PreviousJsonSummaryPath))
    {
        Write-Status 'No previous run summary found for -WhatChanged.' -Level Muted
        return
    }

    try
    {
        $previous = Get-Content -LiteralPath $script:PreviousJsonSummaryPath -Raw | ConvertFrom-Json
        $previousMap = @{}
        foreach ($item in @($previous.Results))
        { $previousMap[$item.Id] = $item
        }

        $changes = [System.Collections.Generic.List[string]]::new()
        foreach ($item in @($CurrentSummary.Results))
        {
            if (-not $previousMap.ContainsKey($item.Id))
            { $changes.Add("new task: $($item.Name) -> $($item.Status)") | Out-Null; continue
            }
            $old = $previousMap[$item.Id]
            if ($old.Status -ne $item.Status)
            { $changes.Add("$($item.Name): $($old.Status) -> $($item.Status)") | Out-Null
            }
        }

        if ($changes.Count -eq 0)
        { Write-Status 'WhatChanged: task statuses match the previous run.' -Level Muted
        } else
        { Write-Status 'WhatChanged:' -Level Info; foreach ($change in $changes)
            { Write-Status "  $change" -Level Muted
            }
        }
    } catch
    {
        Write-Status "WhatChanged did not complete: $($_.Exception.Message)" -Level Muted
    }
}

function Show-UpdateSummary
{
    param([object[]]$Results)

    $entries = [System.Collections.Generic.List[string]]::new()

    foreach ($r in $Results)
    {
        if (-not $r.PSObject.Properties['OutputPreview']) { continue }
        $lines = @($r.OutputPreview | ForEach-Object { [string]$_ })
        if ($lines.Count -eq 0) { continue }
        $text = $lines -join "`n"

        $changes = [System.Collections.Generic.List[string]]::new()

        switch ($r.Id)
        {
            'npm' {
                for ($i = 0; $i -lt $lines.Count; $i++)
                {
                    if ($lines[$i] -match 'Updating npm package:\s+(.+?)(?:@\S+)?\s*$')
                    {
                        $pkg = $Matches[1] -replace '@latest$', ''
                        $n = if (($i + 1) -lt $lines.Count -and $lines[$i + 1] -match 'changed (\d+) packages?') { $Matches[1] } else { '?' }
                        [void]$changes.Add("$pkg  (+$n pkg)")
                    }
                }
            }
            'pnpm' {
                if ($text -match 'Switching (\S+) from v?(\d+(?:\.\d+)+) to v?(\d+(?:\.\d+)+)')
                { [void]$changes.Add("$($Matches[1]): $($Matches[2]) → $($Matches[3])") }
            }
            { $_ -in @('winget', 'winget-batch', 'store-apps') } {
                # Parse the version table (Name / Id / Version / Available)
                $inTable = $false
                foreach ($l in $lines)
                {
                    # Strip spinner chars, then check for table header
                    $clean = $l -replace '[-\\|/]\s+', '' -replace '\s{2,}', '  '
                    if (-not $inTable -and $clean -match 'Name\s+Id\s+.*Version\s+Available')
                    { $inTable = $true; continue }
                    if (-not $inTable) { continue }
                    if ($clean -match '^[\s\-]+$') { continue }
                    if (-not $clean.Trim() -or $clean -match '\d+ upgrade|The following|package\(s\) are pinned|upgrade available')
                    { $inTable = $false; continue }
                    # Extract name and version columns (split on 2+ spaces, take first and last two non-empty tokens)
                    $tokens = @($clean.Trim() -split '\s{2,}' | Where-Object { $_ })
                    if ($tokens.Count -ge 3)
                    {
                        $name = $tokens[0]
                        $avail = $tokens[-1]
                        $ver = $tokens[-2]
                        if ($avail -match '\d' -or $avail -match '^<')
                        { [void]$changes.Add("$name`: $ver → $avail") }
                    }
                }
                # Fallback for winget task: "[N/M] Upgrading: X" lines (only if table didn't fire)
                if ($r.Id -eq 'winget' -and $changes.Count -eq 0)
                {
                    for ($i = 0; $i -lt $lines.Count; $i++)
                    {
                        if ($lines[$i] -match '\[\d+/\d+\] Upgrading:\s+(.+?)(?:\s+\(installed.*\))?$')
                        {
                            $pkg = $Matches[1].Trim()
                            $nextLine = if ($i + 1 -lt $lines.Count) { $lines[$i + 1] } else { '' }
                            if ($nextLine -notmatch '(?i)Not applicable:|FAILED:')
                            { [void]$changes.Add($pkg) }
                        }
                    }
                }
            }
            'scoop' {
                $lines | Where-Object { $_ -match "(?i)Updating\s+'?(\S+)'?\s+\(" -or $_ -match 'Updated .+ \(\S+ -> \S+\)' } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'chocolatey' {
                $lines | Where-Object { ($_ -match '(?i)upgraded \d+' -and $_ -notmatch '(?i)upgraded 0/') -or ($_ -match ' to ' -and $_ -match '\d+\.\d+') } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'cargo' {
                $lines | Where-Object { $_ -match '(?i)(Updating|Updated)\s+\S+\s+v[\d.].*->|v[\d.]+\s+Yes' } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
                # Also catch table rows "name  vOLD  vNEW  Yes"
                $lines | Where-Object { $_ -match '\S+\s+v[\d.]+\s+v[\d.]+\s+Yes' } |
                    Select-Object -First 10 | ForEach-Object {
                        if ($_ -match '(\S+)\s+(v[\d.]+)\s+(v[\d.]+)\s+Yes')
                        { [void]$changes.Add("$($Matches[1]): $($Matches[2]) → $($Matches[3])") }
                    }
            }
            'pip' {
                $lines | Where-Object { $_ -match 'Successfully installed\s+\S' } |
                    Select-Object -First 5 | ForEach-Object {
                        $pkgs = $_ -replace '.*Successfully installed\s+', ''
                        [void]$changes.Add($pkgs.Trim())
                    }
            }
            'uv-tools' {
                $lines | Where-Object { $_ -match '(?i)(Updated|Upgraded)\s+\S+\s+\S+\s*->' } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'mise' {
                $lines | Where-Object { $_ -match '(?:->|→)' -and $_ -match '\d' } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'poetry' {
                if ($text -match 'Package operations: (\d+) installs, (\d+) updates, (\d+) removals')
                {
                    $inst = [int]$Matches[1]; $upd = [int]$Matches[2]; $rem = [int]$Matches[3]
                    if ($inst + $upd + $rem -gt 0)
                    {
                        $down = @($lines | Where-Object { $_ -match '^\s*-\s+Downgrading\s' }).Count
                        if ($down -gt 0)
                        { [void]$changes.Add("$inst installs, $upd updates, $rem removals -- $down DOWNGRADED") }
                        else
                        { [void]$changes.Add("$inst installs, $upd updates, $rem removals") }
                        $lines | Where-Object { $_ -match '^\s*-\s+(Installing|Updating|Downgrading|Removing)\s+(\S+)\s+\(' } |
                            Select-Object -First 6 | ForEach-Object {
                                if ($_ -match '(Installing|Updating|Downgrading|Removing)\s+(\S+)\s+\(([^)]+)\)')
                                { [void]$changes.Add("  $($Matches[2]): $($Matches[3])") }
                            }
                    }
                }
            }
            'pipx' {
                $upgraded = @($lines | Where-Object { $_ -match '^upgrading\s+(.+)\.\.\.' } |
                    ForEach-Object { ($_ -replace '^upgrading\s+', '' -replace '\.\.\.$', '').Trim() })
                if ($upgraded.Count -gt 0) { [void]$changes.Add($upgraded -join ', ') }
            }
            'ruby-gems' {
                $lines | Where-Object { $_ -match '(?i)(Updating|Updated)\s+\S+.*\d+\.\d+' -and ($_ -match '\(' -or $_ -match 'to') } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'gh-extensions' {
                $lines | Where-Object { $_ -match '(?i)upgraded\s+\S+' -or ($_ -match '(?:->|→)' -and $_ -match '\d') } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'powershell-modules' {
                $lines | Where-Object { $_ -match '(?i)(installing|updating)\s+\S+.*\d+\.\d+' } |
                    Select-Object -First 10 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
            'dotnet-workloads' {
                $cnt = @($lines | Where-Object { $_ -match 'Updated advertising manifest' }).Count
                if ($cnt -gt 0) { [void]$changes.Add("$cnt manifests updated") }
            }
            'appx-repair' {
                if ($text -match '(\d+) package\(s\) re-registered')
                { [void]$changes.Add("$($Matches[1]) packages re-registered") }
            }
            default {
                $lines | Where-Object { $_ -match '(?:->|→)' -and $_ -match '\d[\d.]+' -and $_.Length -lt 120 } |
                    Select-Object -First 5 | ForEach-Object { [void]$changes.Add($_.Trim()) }
            }
        }

        if ($changes.Count -gt 0)
        {
            $label = $r.Name.PadRight(16)
            [void]$entries.Add("  $label  $($changes[0])")
            foreach ($c in $changes | Select-Object -Skip 1)
            { [void]$entries.Add("  $(' ' * 18)$c") }
        }
    }

    if ($entries.Count -gt 0)
    {
        Write-Host ''
        Write-Status "What's Changed:" -Level Info
        foreach ($e in $entries) { Write-Status $e -Level Muted }
    }
}

function Get-RunNotes
{
    param([object[]]$Results, [object[]]$Skipped)

    $notes = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    function Add-RunNote
    {
        param(
            [ValidateSet('Info', 'Warning')]
            [string]$Level,
            [Parameter(Mandatory)][string]$Message
        )

        if ($seen.Add($Message))
        { [void]$notes.Add([pscustomobject]@{ Level = $Level; Message = $Message })
        }
    }

    $failedResults = @($Results | Where-Object { $_.Status -in @('Failed', 'TimedOut') })
    if ($failedResults.Count -gt 0)
    {
        $ids = (@($failedResults | Select-Object -ExpandProperty Id) -join ',')
        Add-RunNote -Level Warning -Message "Focused retry: update -Only $ids -NoParallel -ShowSkipped"
    }

    foreach ($result in @($Results))
    {
        $outputPreview = @()
        if ($result.PSObject.Properties['OutputPreview'])
        { $outputPreview = @($result.OutputPreview)
        }
        $outputText = ($outputPreview -join "`n")

        if ($result.Id -eq 'uv' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "uv: self-update only works for standalone uv installs. Managed installs should be updated by winget, pipx, pip, scoop, or Chocolatey."
        }

        if ($result.Id -eq 'pip-health' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "pip-health: add intentionally broken/legacy packages to PipIgnoreHealthPackages in update-config.json, or fix the dependency conflict."
        }

        if ($result.Id -eq 'composer' -and $outputText -match 'openssl extension is not loaded|openssl PHP extension')
        {
            Add-RunNote -Level Warning -Message "composer: enable 'extension=openssl' in php.ini to allow self-updates. Run 'php --ini' to find the loaded config file."
        }

        if ($result.Id -eq 'wsl-distros' -and $result.Status -eq 'Partial')
        {
            Add-RunNote -Level Warning -Message "wsl-distros: network was unavailable inside WSL. Run 'wsl --shutdown' then rerun '-Only wsl-distros' to retry."
        }

        if ($result.Id -eq 'pnpm' -and $outputText -match 'pnpm v10 installation layout|pnpm v11 expects bins in PNPM_HOME/bin|pnpm shim appears to reference a missing executable|@pnpm/exe/pnpm\.exe')
        {
            Add-RunNote -Level Warning -Message "pnpm: run 'pnpm setup' once in a fresh terminal, or reinstall pnpm, to repair the broken global shim."
        }

        if ($result.Id -eq 'oh-my-posh' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message 'oh-my-posh: self-update did not complete. Prefer updating it through winget, Scoop, or Chocolatey, or run oh-my-posh upgrade manually in an interactive terminal.'
        }

        if ($result.Id -eq 'claude' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "claude: self-update did not complete. Run 'claude update' manually, or update through winget id Anthropic.Claude."
        }

        if ($result.Id -eq 'vcpkg' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "vcpkg: run 'vcpkg upgrade --no-dry-run' manually. The vcpkg baseline may need to be updated via 'vcpkg update' first."
        }

        if ($result.Id -eq 'conda' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "conda: run 'conda update --all -y' manually. Check for environment conflicts with 'conda info'."
        }

        if ($result.Id -eq 'gcloud' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "gcloud: run 'gcloud components update --quiet' manually. The Google Cloud SDK may need reinstallation."
        }

        if ($result.Id -eq 'terraform' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "terraform: manual upgrade required. Download from https://www.terraform.io/downloads or use tfenv."
        }

        if ($result.Id -eq 'starship' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "starship: run 'starship self-update -y' manually, or update via winget/scoop/choco."
        }

        if ($result.Id -eq 'trivy' -and $result.Status -in @('Failed', 'TimedOut'))
        {
            Add-RunNote -Level Warning -Message "trivy: run 'trivy update' manually, or update via winget/scoop."
        }

        if ($result.Id -eq 'self-update' -and $result.Status -eq 'Succeeded' -and $outputText -match 'Update available')
        {
            Add-RunNote -Level Info -Message "A new script version is available. Run the self-update task or download from the repo."
        }

        if ($outputText -match 'Protected-suppressed')
        {
            Add-RunNote -Level Warning -Message 'Protected app updates were suppressed by configuration. Remove protected package entries from update-config.json to update every app.'
        }
    }

    if (@($Skipped | Where-Object { $_.Reason -eq 'requires Administrator' }).Count -gt 0)
    {
        Add-RunNote -Level Info -Message 'Admin-only tasks were skipped. Rerun with -AutoElevate for Windows Update, Chocolatey, and .NET workload coverage.'
    }

    return @($notes)
}

function Show-RunNotes
{
    param([object[]]$Notes)
    if (@($Notes).Count -eq 0)
    { return
    }

    Write-Status 'Notes:' -Level Info
    foreach ($note in @($Notes))
    {
        $level = if ($note.PSObject.Properties['Level'] -and $note.Level)
        { $note.Level
        } else
        { 'Info'
        }
        Write-Status ("  {0}" -f $note.Message) -Level $level
    }
}

function Register-UpdateSchedule
{
    if (-not $IsWindows)
    { throw 'Scheduling is only supported on Windows.'
    }
    if (-not (Test-IsAdmin))
    { throw 'Scheduled task registration requires Administrator.'
    }

    $pwsh = Get-CommandPath 'pwsh.exe'
    if (-not $pwsh)
    { $pwsh = 'powershell.exe'
    }

    $taskName = 'UpdateEverything'
    $taskArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath),
        '-SkipReboot', '-SkipWSL', '-SkipWindowsUpdate', '-Quiet'
    ) -join ' '

    if ($script:IsSimulation)
    {
        Write-Status "[DryRun] Would register scheduled task $taskName at $ScheduleTime on $($ScheduleDays -join ',')." -Level Info
        return
    }

    $action = New-ScheduledTaskAction -Execute $pwsh -Argument $taskArgs

    if ($ScheduleRepeat -gt 0)
    {
        $duration = New-TimeSpan -Hours 23 -Minutes 59
        $interval = New-TimeSpan -Hours $ScheduleRepeat
        $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime -RepetitionInterval $interval -RepetitionDuration $duration
    } elseif ($ScheduleDays.Count -gt 0 -and $ScheduleDays.Count -lt 7)
    {
        $daysOfWeek = [DayOfWeek[]]@($ScheduleDays | ForEach-Object { [DayOfWeek]$_ })
        $trigger = New-ScheduledTaskTrigger -Weekly -At $ScheduleTime -DaysOfWeek $daysOfWeek
    } else
    {
        $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    }

    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 4) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
    $desc = if ($ScheduleRepeat -gt 0) { "every $ScheduleRepeat hours starting at ${ScheduleTime}" } else { "$($ScheduleDays -join ',') at ${ScheduleTime}" }
    Write-Status ("Scheduled ${taskName}: $desc.") -Level Success
}

function Invoke-SelfElevation
{
    if (-not $AutoElevate -or $NoElevate -or -not $IsWindows -or (Test-IsAdmin))
    { return $false
    }

    $pwsh = Get-CommandPath 'pwsh.exe'
    if (-not $pwsh)
    { $pwsh = 'powershell.exe'
    }

    $forwarded = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $PSBoundParameters.GetEnumerator())
    {
        if ($entry.Key -eq 'AutoElevate')
        { continue
        }
        if ($entry.Value -is [switch])
        {
            if ($entry.Value.IsPresent)
            { $forwarded.Add("-$($entry.Key)") | Out-Null
            }
        } elseif ($entry.Value -is [array])
        {
            if ($entry.Value.Count -gt 0)
            {
                $forwarded.Add("-$($entry.Key)") | Out-Null
                $quoted = @($entry.Value | ForEach-Object { '"{0}"' -f ([string]$_).Replace('"', '\"') })
                $forwarded.Add($quoted -join ',') | Out-Null
            }
        } else
        {
            $forwarded.Add("-$($entry.Key)") | Out-Null
            $forwarded.Add(('"{0}"' -f ([string]$entry.Value).Replace('"', '\"'))) | Out-Null
        }
    }

    $elevateArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) + @($forwarded)
    Start-Process -FilePath $pwsh -Verb RunAs -ArgumentList ($elevateArgs -join ' ') -Wait
    return $true
}

# Main
Initialize-ProcessPath
Import-UpdateConfig
Initialize-RunStorage

if ($CIMode)
{
    $NoElevate = $true
    $AutoElevate = $false
    if (-not $Quiet)
    { $Quiet = $false }
}

if (Invoke-SelfElevation)
{ exit 0
}

if (-not $script:IsSimulation)
{ Enter-ProcessLock
}

$isAdmin = Test-IsAdmin
if (-not $isAdmin -and -not $NoElevate)
{
    Write-Status 'Running without Administrator. Admin-only tasks will be skipped. Use -AutoElevate for a full run.' -Level Info
}

if ($Schedule)
{
    Register-UpdateSchedule
    exit 0
}
if ($UltraFast)
{ $FastMode = $true
}
if ($ParallelThrottle -lt 1)
{ $ParallelThrottle = [Math]::Max(2, [Math]::Min([Environment]::ProcessorCount, 6))
}
if ($NoParallel)
{ $ParallelThrottle = 1
}

# Profile-aware skip list — merge into the -Skip array before task generation
if ($Profile -and $script:Config.ProfileSkipLists.Contains($Profile))
{
    $profileSkips = @(ConvertTo-FilterList $script:Config.ProfileSkipLists[$Profile])
    if ($profileSkips.Count -gt 0)
    {
        $Skip = @($Skip) + $profileSkips
        Write-Status "Profile '$Profile' active: $($profileSkips.Count) task group(s) suppressed." -Level Info
    }
}

$allTasks = @(Get-UpdateTasks)
$filtered = Get-FilteredTasks -Tasks $allTasks -IsAdmin $isAdmin
$plannedTasks = @($filtered.Planned)
$skippedTasks = @($filtered.Skipped)

# Snapshot mode: record all-task state without executing updates
if ($script:IsSnapshot)
{
    Write-Status 'Snapshot mode: recording current environment state without updating.' -Level Info
    $snapResults = [System.Collections.Generic.List[object]]::new()
    foreach ($task in $plannedTasks)
    {
        $reqCmd = @(if ($task.PSObject.Properties['RequiresCommand'] -and $task.RequiresCommand) { @($task.RequiresCommand) } else { @() })
        $isAdmin = if ($task.PSObject.Properties['RequiresAdmin'] -and $task.RequiresAdmin) { $true } else { $false }
        $snapEntry = [pscustomobject]@{
            Id        = $task.Id
            Name      = $task.Name
            Category  = $task.Category
            Status    = 'Snapshotted'
            Timestamp = (Get-Date).ToString('o')
            Commands  = $reqCmd
            AdminOnly = $isAdmin
        }
        [void]$snapResults.Add($snapEntry)
    }
    $snapshotSummary = Save-RunSummary -Results @($snapResults) -Skipped $skippedTasks -Planned $plannedTasks
    if ($snapshotSummary.SummaryWritten)
    { Write-Status "Snapshot written to $script:JsonSummaryPath" -Level Success
    }
    Remove-ProcessLock
    exit 0
}

# Checkpoint resume: skip tasks that completed in a previous interrupted run
if (-not $script:IsSimulation -and (Test-Path -LiteralPath $script:CheckpointPath))
{
    try
    {
        $checkpointData = Get-Content -LiteralPath $script:CheckpointPath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if ($checkpointData.RunId -eq $script:RunId -and $checkpointData.Completed)
        {
            foreach ($completedId in @($checkpointData.Completed))
            { [void]$script:CompletedTasks.Add($completedId)
            }
            Write-Status "Resuming from checkpoint: $($script:CompletedTasks.Count) task(s) already completed this run." -Level Info
        }
    } catch
    { Write-Status "Could not read checkpoint file: $($_.Exception.Message)" -Level Warning
    }
}

if ($SelfTest)
{
    $selfTestTask = New-UpdateTask -Name 'self-test' -Category 'diagnostics' -Script { Write-Output 'Scheduler, logging, process execution, and summary path are working.' }
    $selfTestTask | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
    $plannedTasks = @($selfTestTask)
    $skippedTasks = @()
}

Write-Status ("Update-Everything v{0} | {1:yyyy-MM-dd HH:mm} | throttle {2}" -f $script:Version, $script:StartTime, $ParallelThrottle) -Level Info

if ($ListTasks)
{
    Show-TaskList -Planned $plannedTasks -Skipped $skippedTasks
    $summary = Save-RunSummary -Results @() -Skipped $skippedTasks -Planned $plannedTasks
    if ($summary.SummaryWritten)
    { Write-Status "Task list written to $script:JsonSummaryPath" -Level Muted
    }
    exit 0
}

if ($script:IsHealthCheck)
{
    Write-Status 'Health check mode: verifying tool availability without updating.' -Level Info
    $healthResults = [System.Collections.Generic.List[object]]::new()
    $healthy = 0; $unhealthy = 0
    foreach ($task in $plannedTasks)
    {
        $requiredCmds = @()
        if ($task.PSObject.Properties['RequiresCommand'] -and $task.RequiresCommand)
        { $requiredCmds = @(if ($task.RequiresCommand -is [array]) { $task.RequiresCommand } else { @($task.RequiresCommand) })
        }
        $found = @($requiredCmds | Where-Object { Get-Command $_ -ErrorAction SilentlyContinue })
        $status = if ($found.Count -eq $requiredCmds.Count -and $requiredCmds.Count -gt 0) { 'Available' } elseif ($requiredCmds.Count -eq 0) { 'NoRequirement' } else { 'Missing' }
        if ($status -eq 'Available') { $healthy++ } else { $unhealthy++ }
        $healthEntry = [pscustomobject]@{
            Id            = $task.Id
            Name          = $task.Name
            Category      = $task.Category
            RequiresCommand = $requiredCmds
            FoundCommands  = @($found | ForEach-Object { (Get-Command $_).Source })
            Status        = $status
        }
        [void]$healthResults.Add($healthEntry)
    }
    $healthSummary = Save-RunSummary -Results @($healthResults) -Skipped $skippedTasks -Planned $plannedTasks
    Write-Status "Health check: $healthy available, $unhealthy missing tools." -Level $(if ($unhealthy -eq 0) { 'Success' } else { 'Warning' })
    if ($healthSummary.SummaryWritten)
    { Write-Status "Health check written to $script:JsonSummaryPath" -Level Muted
    }
    Remove-ProcessLock
    exit 0
}

if ($plannedTasks.Count -eq 0)
{
    Write-Status 'No runnable update tasks were found.' -Level Info
    $summary = Save-RunSummary -Results @() -Skipped $skippedTasks -Planned $plannedTasks
    Remove-ProcessLock
    exit 0
}

if ($script:IsSimulation)
{
    Write-Status 'Dry run: no update commands will be executed.' -Level Info
    foreach ($task in $plannedTasks)
    { Write-Status ("[DryRun] {0} ({1})" -f $task.Name, $task.Category) -Level Muted
    }
    $dryResults = @($plannedTasks | ForEach-Object { New-TaskResult -Task $_ -Status 'DryRun' -Reason 'preview only' })
    $summary = Save-RunSummary -Results $dryResults -Skipped $skippedTasks -Planned $plannedTasks
    if ($WhatChanged)
    { Show-WhatChanged -CurrentSummary $summary
    }
    if ($summary.SummaryWritten)
    { Write-Status "Dry-run summary written to $script:JsonSummaryPath" -Level Success
    }
    Remove-ProcessLock
    exit 0
}

# Interactive mode: filter tasks that need confirmation
if ($script:IsInteractive)
{
    Write-Status 'Interactive mode: confirmation required for each task group.' -Level Info
    $confirmed = [System.Collections.Generic.List[object]]::new()
    $allSelected = $false
    foreach ($task in $plannedTasks)
    {
        if ($allSelected) { [void]$confirmed.Add($task); continue }
        $response = $host.UI.PromptForChoice("Run task: $($task.Name) ($($task.Category))?", 'Select:', @('&Yes', '&No', '&All'), 0)
        if ($response -eq 0)
        { [void]$confirmed.Add($task)
        } elseif ($response -eq 2)
        { [void]$confirmed.Add($task); $allSelected = $true
        }
    }
    $plannedTasks = @($confirmed)
}

$splitSkippedForDisplay = Split-SkippedTasksForDisplay -Skipped $skippedTasks
if ($ShowSkipped -or $splitSkippedForDisplay.Hidden.Count -eq 0)
{
    Write-Status ("Dispatching {0} task(s). Skipped: {1}" -f $plannedTasks.Count, $skippedTasks.Count) -Level Info
} else
{
    Write-Status ("Dispatching {0} task(s). Skipped: {1} ({2} optional hidden; use -ShowSkipped to list them)" -f $plannedTasks.Count, $skippedTasks.Count, $splitSkippedForDisplay.Hidden.Count) -Level Info
}

$skippedTasksForDisplay = if ($ShowSkipped) { @($skippedTasks) } else { @($splitSkippedForDisplay.Visible) }
if (@($skippedTasksForDisplay).Count -gt 0)
{
    foreach ($s in $skippedTasksForDisplay)
    { Write-Status ("  Skip: {0} — {1}" -f $s.Name, $s.Reason) -Level Muted
    }
}
$results = Invoke-TaskQueue -Tasks $plannedTasks -Throttle $ParallelThrottle

# Save checkpoint: mark tasks that completed successfully
if (-not $script:IsSimulation -and $results.Count -gt 0)
{
    $completedIds = @($results | Where-Object { $_.Status -eq 'Succeeded' } | ForEach-Object { $_.Id })
    if ($completedIds.Count -gt 0)
    {
        $checkpoint = [ordered]@{ RunId = $script:RunId; Completed = $completedIds }
        try { $checkpoint | ConvertTo-Json -Depth 2 | Set-Content -LiteralPath $script:CheckpointPath -Encoding utf8 -Force -ErrorAction Stop } catch { }
    }
}

# Run before/after hooks from config
$runNotes = @(Get-RunNotes -Results $results -Skipped $skippedTasks)
$summary = Save-RunSummary -Results $results -Skipped $skippedTasks -Planned $plannedTasks -Notes $runNotes

# Remove checkpoint on successful run completion
if (-not $script:IsSimulation -and (Test-Path -LiteralPath $script:CheckpointPath))
{
    try { Remove-Item -LiteralPath $script:CheckpointPath -Force -ErrorAction Stop } catch { }
}

# Run before/after hooks from config
$hookBefore = if ($null -ne $script:BeforeHooks) { $script:BeforeHooks } else { $null }
$hookAfter  = if ($null -ne $script:AfterHooks)  { $script:AfterHooks } else { $null }
if ($hookBefore -and @($hookBefore.Keys).Count -gt 0 -and -not $script:IsSimulation)
{
    foreach ($hookEntry in $hookBefore.GetEnumerator())
    { try { & $hookEntry.Value } catch { Write-Status "Before hook '$($hookEntry.Key)' failed: $($_.Exception.Message)" -Level Warning } }
}

# Run after hooks
if ($hookAfter -and @($hookAfter.Keys).Count -gt 0 -and -not $script:IsSimulation)
{
    foreach ($hookEntry in $hookAfter.GetEnumerator())
    { try { & $hookEntry.Value } catch { Write-Status "After hook '$($hookEntry.Key)' failed: $($_.Exception.Message)" -Level Warning } }
}

if ($WhatChanged)
{ Show-WhatChanged -CurrentSummary $summary
}

$failed  = @($results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -in @('Failed', 'TimedOut') })
$succeeded = @($results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -eq 'Succeeded' })
$warned  = @($results | Where-Object { $_.PSObject.Properties['Status'] -and $_.Status -in @('Warn', 'Partial') })
$elapsed = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')

Show-ResultTable -Results $results -Skipped $skippedTasks

Show-UpdateSummary -Results $results

Show-RunNotes -Notes $runNotes

if ($summary.SummaryWritten)
{ Write-Status "Summary: $script:JsonSummaryPath" -Level Muted
}
if (Test-Path -LiteralPath $script:LogPath)
{ Write-Status "Log: $script:LogPath" -Level Muted
}

# Save remote state for multi-machine sync
if ($RemoteStatePath -and $summary.SummaryWritten)
{
    try
    {
        if (-not (Test-Path -LiteralPath $RemoteStatePath))
        { New-Item -Path $RemoteStatePath -ItemType Directory -Force -ErrorAction Stop | Out-Null }
        Copy-Item -LiteralPath $script:JsonSummaryPath -Destination (Join-Path $RemoteStatePath 'last-run.json') -Force -ErrorAction Stop
        Write-Status "Remote state synced to $RemoteStatePath" -Level Muted
    } catch
    { Write-Status "Remote state sync failed: $($_.Exception.Message)" -Level Warning
    }
}

# Webhook notification
if ($WebhookUrl -and ($failed.Count -gt 0 -or $warned.Count -gt 0))
{
    try
    {
        $body = [ordered]@{
            text = "Update-Everything v$($script:Version) finished: $($succeeded.Count) ok, $($failed.Count) failed, $($skippedTasks.Count) skipped in $elapsed."
        } | ConvertTo-Json
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $body -ContentType 'application/json' -TimeoutSec 15 -ErrorAction Stop | Out-Null
        Write-Status 'Webhook notification sent.' -Level Muted
    } catch
    { Write-Status "Webhook notification failed: $($_.Exception.Message)" -Level Warning
    }
}

if ($failed.Count -gt 0)
{
    Write-Status ("Completed with {0} succeeded, {1} failed/timed out, {2} skipped in {3}." -f $succeeded.Count, $failed.Count, $skippedTasks.Count, $elapsed) -Level Warning
} elseif ($warned.Count -gt 0)
{
    Write-Status ("Completed with {0} succeeded, {1} warning/partial, {2} skipped in {3}." -f $succeeded.Count, $warned.Count, $skippedTasks.Count, $elapsed) -Level Warning
} else
{
    Write-Status ("All runnable tasks completed: {0} succeeded, {1} skipped in {2}." -f $succeeded.Count, $skippedTasks.Count, $elapsed) -Level Success
}

if ($Notify)
{ Send-UpdateNotification -Succeeded $succeeded.Count -Failed $failed.Count -Skipped $skippedTasks.Count -Elapsed $elapsed
}

if (-not $SkipReboot -and $IsWindows)
{
    $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
    (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    if ($pendingReboot)
    { Write-Status 'A reboot appears to be pending.' -Level Info
    }
}

Remove-ProcessLock

$exitCode = 0
if ($CIMode -and ($failed.Count -gt 0 -or $warned.Count -gt 0))
{ $exitCode = 1
} elseif ($failed.Count -gt 0)
{ $exitCode = 1
}
exit $exitCode
