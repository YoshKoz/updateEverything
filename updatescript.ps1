#requires -version 7.0
<#
.SYNOPSIS
    Updates common Windows 11 package managers, developer tools, runtimes, WSL distros, Defender, and maintenance tasks.

.VERSION
    Update-Everything v6.5.0-expanded-tooling

.NOTES
    Main fixes versus v6.2.0:
      - Refreshes process PATH from Machine/User environment values before task discovery.
      - Prefers real Python installs over the Microsoft Store app execution alias.
      - Correctly launches .cmd/.bat shims such as VS Code's code.cmd through cmd.exe.
      - Finds VS Code CLI from PATH and common install locations.
      - Avoids stale $LASTEXITCODE leakage between tasks.
      - Uses WUA COM as the primary Windows Update fallback instead of relying only on UsoClient.
      - Adds Defender MpCmdRun fallback.
      - Handles npm locked @openai/codex installs and stale npm temp folders more gracefully.
      - Prevents winget and msstore winget tasks from running at the same time.
      - Makes WSL distro updates non-interactive and treats sudo-password skips as non-fatal.
      - v6.4.1: hides optional skipped tools unless -ShowSkipped, adds Store skip list, suppresses WSL mirrored-networking noise, makes Update-Help opt-in, and quiets Chocolatey no-op runs.
      - v6.4.2: preserves failing command output in errors, skips managed uv self-updates cleanly, adds pip dependency health checks, and prints actionable run notes.
      - v6.4.3: protects self-updating/stateful desktop apps by default, upgrades Chocolatey packages individually, and adds a non-silent winget option.
      - v6.5.0: refreshes winget sources, restores broader standalone developer-tool coverage, adds uv Python patch refreshes, and supports zypper-based WSL distros.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$SkipWindowsUpdate,
    [switch]$SkipReboot,
    [switch]$SkipDestructive,
    [switch]$FastMode,
    [switch]$UltraFast,
    [switch]$NoElevate,
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
    [switch]$DeepClean,
    [switch]$UpdateOllamaModels,
    [switch]$WhatChanged,
    [switch]$DryRun,
    [switch]$ListTasks,
    [switch]$SelfTest,
    [switch]$NoParallel,
    [switch]$Quiet,
    [switch]$ShowSkipped,
    [switch]$IncludeProtectedApps = $true,
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
    [string[]]$Skip = @()
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

$script:Version = '6.5.0-expanded-tooling'
$script:StartTime = Get-Date
$script:RunId = $script:StartTime.ToString('yyyyMMdd-HHmmss-fff')
$script:CommandCache = @{}
$script:StateDirWasProvided = -not [string]::IsNullOrWhiteSpace($StateDir)
$script:LogPathWasProvided = -not [string]::IsNullOrWhiteSpace($LogPath)
$script:JsonSummaryPathWasProvided = -not [string]::IsNullOrWhiteSpace($JsonSummaryPath)
$script:LogWriteWarningEmitted = $false
$script:IsSimulation = $DryRun -or $WhatIfPreference

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
        'powershell-help', 'uv-python', 'ollama-models'
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

function Write-Log
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

    Write-Log -Message $Message -Level $Level
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
    $tempCleanupDays = [int]$script:Config.TempCleanupDays
    $useSilentInstallers = -not [bool]$NoSilentInstallers
    $tasks = [System.Collections.Generic.List[object]]::new()

    $tasks.Add((New-UpdateTask -Name 'winget-source' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec 300 -Script {
                Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('source', 'update') -Retries 1 -TimeoutSec 300
            } -Tags @('windows') -Resources @('winget'))) | Out-Null

    $wingetScript = {
        param(
            [string[]]$SkipPackages,
            [string[]]$ProtectedPackages,
            [bool]$IncludeProtected,
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

        if ($skipSet.Count -gt 0)
        { Write-Output "Config skip list ($($skipSet.Count)): $($skipSet -join ', ')"
        }
        if ($protectedSet.Count -gt 0 -and -not $IncludeProtected)
        { Write-Output "Protected app list ($($protectedSet.Count)): $($protectedSet -join ', ')"
        }
        if (-not $UseSilentInstallers)
        { Write-Output 'winget installer mode: standard (no --silent).'
        }

        Write-Output 'Scanning for available winget upgrades...'
        $listOutput = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--include-unknown', '--include-pinned', '--accept-source-agreements') -SuccessExitCodes @(0, -1978335189)
        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        $idColumnStart = $null
        $idColumnEnd = $null
        $ansiPattern = '\x1b\[[0-9;?]*[a-zA-Z]'

        foreach ($line in $listOutput)
        {
            $text = ([regex]::Replace([string]$line, $ansiPattern, '')).TrimEnd()
            if ($text -match 'No installed package found matching input criteria|No available upgrade found|No applicable update found')
            {
                continue
            }
            if ($text -match '^\s*Name\s+Id\s+Version\s+Available')
            {
                $idColumnStart = $text.IndexOf('Id')
                $idColumnEnd = $text.IndexOf('Version', $idColumnStart + 2)
                $inTable = $true
                continue
            }
            if ($text -match '^\s*-{3,}')
            {
                $inTable = $true
                $columnMatches = [regex]::Matches($text, '-{3,}')
                if ($columnMatches.Count -ge 2)
                {
                    $idColumnStart = $columnMatches[1].Index
                    if ($columnMatches.Count -ge 3)
                    { $idColumnEnd = $columnMatches[2].Index
                    }
                }
                continue
            }
            if (-not $inTable -or [string]::IsNullOrWhiteSpace($text))
            { continue
            }
            if ($null -eq $idColumnStart -or $text.Length -le $idColumnStart)
            { continue
            }

            if ($null -ne $idColumnEnd -and $idColumnEnd -gt $idColumnStart)
            {
                $width = [Math]::Min($idColumnEnd - $idColumnStart, $text.Length - $idColumnStart)
                $id = $text.Substring($idColumnStart, $width).Trim()
            } else
            {
                $id = $text.Substring($idColumnStart).Trim()
            }
            if ($id -and $id -notmatch '^(Id|Version|-)' -and $id -match '^[A-Za-z0-9][A-Za-z0-9._+\-]+$')
            {
                [void]$ids.Add($id)
            }
        }

        $toUpgrade = [System.Collections.Generic.List[string]]::new()
        $toSkipIds = [System.Collections.Generic.List[string]]::new()
        $toProtectIds = [System.Collections.Generic.List[string]]::new()
        foreach ($id in ($ids | Sort-Object -Unique))
        {
            if ($skipSet.Contains($id))
            { [void]$toSkipIds.Add($id)
            } elseif ((-not $IncludeProtected) -and $protectedSet.Contains($id))
            { [void]$toProtectIds.Add($id)
            } else
            { [void]$toUpgrade.Add($id)
            }
        }

        if (($toUpgrade.Count + $toSkipIds.Count + $toProtectIds.Count) -eq 0)
        {
            Write-Output 'No winget upgrades available.'
            return
        }

        $suppressedCount = $toSkipIds.Count + $toProtectIds.Count
        Write-Output ("Found {0} package(s): {1} to upgrade, {2} suppressed." -f ($toUpgrade.Count + $suppressedCount), $toUpgrade.Count, $suppressedCount)
        if ($toSkipIds.Count -gt 0)
        { Write-Output "  Config-suppressed: $($toSkipIds -join ', ')"
        }
        if ($toProtectIds.Count -gt 0)
        { Write-Output "  Protected-suppressed: $($toProtectIds -join ', ')"
        }
        if ($toUpgrade.Count -eq 0)
        { Write-Output 'All available upgrades are suppressed by skip/protected lists.'; return
        }

        $failed = [System.Collections.Generic.List[string]]::new()
        $upgradeIndex = 0
        foreach ($id in $toUpgrade)
        {
            $upgradeIndex++
            Write-Output ("[{0}/{1}] Upgrading: {2}" -f $upgradeIndex, $toUpgrade.Count, $id)
            try
            {
                $wingetArgs = @('upgrade', '--id', $id, '--exact', '--include-unknown', '--include-pinned', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements')
                if ($UseSilentInstallers)
                { $wingetArgs += '--silent'
                }
                Invoke-UpdateProcess -FilePath 'winget' -ArgumentList $wingetArgs -Retries 1
            } catch
            {
                Write-Output "  FAILED: $($_.Exception.Message)"
                [void]$failed.Add($id)
            }
        }

        if ($failed.Count -gt 0)
        { Write-Output "winget left unchanged: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'winget' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec $WingetTimeoutSec -Script $wingetScript -Tags @('windows') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'scoop' -Category 'package-manager' -RequiresCommand 'scoop' -Script {
                Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update')
                Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update', '*') -Retries 1
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'chocolatey' -Category 'package-manager' -RequiresCommand 'choco' -RequiresAdmin -Script {
                param(
                    [string[]]$SkipPackages,
                    [string[]]$ProtectedPackages,
                    [bool]$IncludeProtected
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
                    } elseif ((-not $IncludeProtected) -and $protectedSet.Contains($entry.Name))
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
                    [bool]$IncludeProtected,
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
                $listOutput = @(Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--source', 'msstore', '--include-unknown', '--include-pinned', '--accept-source-agreements') -SuccessExitCodes @(0, -1978335189))
                $ids = [System.Collections.Generic.List[string]]::new()
                $ansiPattern = '\x1b\[[0-9;?]*[a-zA-Z]'

                foreach ($line in $listOutput)
                {
                    $text = ([regex]::Replace([string]$line, $ansiPattern, '')).Trim()
                    if ([string]::IsNullOrWhiteSpace($text))
                    { continue
                    }
                    if ($text -match 'No installed package found matching input criteria|No available upgrade found|No applicable update found|No available updates')
                    { continue
                    }

                    # Microsoft Store package IDs are usually compact alphanumeric IDs. The table format can shift
                    # when package names contain spaces, so use the version columns at the end as anchors.
                    if ($text -match '\s([A-Za-z0-9]{8,})\s+[^\s]+\s+[^\s]+\s*$')
                    {
                        $candidate = $matches[1]
                        if ($candidate -notmatch '^(Name|Id|Version|Available)$')
                        { [void]$ids.Add($candidate)
                        }
                    }
                }

                $ids = @($ids | Sort-Object -Unique)
                if ($ids.Count -eq 0)
                {
                    Write-Output 'No Microsoft Store app upgrades available.'
                    return
                }

                $toUpgrade = [System.Collections.Generic.List[string]]::new()
                $toSkipIds = [System.Collections.Generic.List[string]]::new()
                $toProtectIds = [System.Collections.Generic.List[string]]::new()
                foreach ($id in $ids)
                {
                    if ($skipSet.Contains($id))
                    { [void]$toSkipIds.Add($id)
                    } elseif ((-not $IncludeProtected) -and $protectedSet.Contains($id))
                    { [void]$toProtectIds.Add($id)
                    } else
                    { [void]$toUpgrade.Add($id)
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
                foreach ($id in $toUpgrade)
                {
                    Write-Output "Updating Microsoft Store app: $id"
                    try
                    {
                        $storeArgs = @('upgrade', '--id', $id, '--exact', '--source', 'msstore', '--include-unknown', '--include-pinned', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements')
                        if ($UseSilentInstallers)
                        { $storeArgs += '--silent'
                        }
                        Invoke-UpdateProcess -FilePath 'winget' -ArgumentList $storeArgs -Retries 1 -TimeoutSec $WingetTimeoutSec
                    } catch
                    {
                        Write-Output "Store app could not be updated automatically: $id — $($_.Exception.Message)"
                        [void]$failed.Add($id)
                    }
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
  sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -o Acquire::Retries=2 update &&   sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y -o Dpkg::Options::=--force-confdef -o Dpkg::Options::=--force-confold upgrade &&   sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y autoremove
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
                    foreach ($distro in $distros)
                    {
                        Write-Output "Updating WSL distro: $distro"
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
                        } catch
                        {
                            $message = $_.Exception.Message
                            if ($message -match 'ConfigureNetworking|0x8007054f|Temporary failure resolving|Could not resolve host|failed to synchronize all databases|networkingMode|^wsl failed with exit code')
                            {
                                Write-Output "Skipping WSL distro after transient network problem: $distro"
                                [void]$skippedDistros.Add($distro)
                                continue
                            }
                            Write-Output $message
                            [void]$failedDistros.Add($distro)
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

        $listJson = (Invoke-UpdateProcess -FilePath 'npm' -ArgumentList @('ls', '-g', '--depth=0', '--json') -SuccessExitCodes @(0, 1) | Out-String).Trim()
        if (-not $listJson)
        { throw 'npm did not return a global package list.'
        }
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
                $result = Invoke-UpdateProcess -FilePath 'pnpm' -ArgumentList @('self-update') -Retries 1 -PassThru
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
                Invoke-UpdateProcess -FilePath 'yarn' -ArgumentList @('global', 'upgrade') -Retries 1
            } -Tags @('node') -Resources @('npm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'bun' -Category 'javascript' -RequiresCommand 'bun' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'bun' -ArgumentList @('upgrade') -Retries 1
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'deno' -Category 'javascript' -RequiresCommand 'deno' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'deno' -ArgumentList @('upgrade') -Retries 1
            } -Tags @('node'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'mise' -Category 'version-manager' -RequiresCommand 'mise' -Script {
                Invoke-UpdateProcess -FilePath 'mise' -ArgumentList @('self-update', '--yes') -Retries 1
                Invoke-UpdateProcess -FilePath 'mise' -ArgumentList @('upgrade', '--yes') -Retries 1
            } -Tags @('version-manager'))) | Out-Null

    $pipScript = {
        param([string[]]$SkipPackages)
        Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', 'pip') -Retries 1

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($pkg in @($SkipPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($pkg))
            { [void]$skipSet.Add($pkg.Trim())
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
            if ($skipSet.Contains($pkg.name))
            { Write-Output "Skipping pip package: $($pkg.name)"; continue
            }
            Write-Output "Upgrading pip package: $($pkg.name) $($pkg.version) -> $($pkg.latest_version)"
            try
            { Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', $pkg.name) -Retries 1
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

        Write-Output ("pip dependency check found {0} active issue(s):" -f $activeIssues.Count)
        foreach ($issue in $activeIssues)
        { Write-Output "  $issue"
        }
        $global:LASTEXITCODE = 1
        throw 'pip check found active dependency issues.'
    }
    $tasks.Add((New-UpdateTask -Name 'pip-health' -Category 'python' -RequiresCommand 'python' -Disabled:$SkipPipHealth -DisabledReason 'disabled by -SkipPipHealth' -Script $pipHealthScript -Tags @('python', 'health') -Resources @('pip'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pipx' -Category 'python' -RequiresCommand 'pipx' -Script {
                Invoke-UpdateProcess -FilePath 'pipx' -ArgumentList @('upgrade-all') -Retries 1
            } -Tags @('python') -Resources @('pip'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv' -Category 'python' -RequiresCommand 'uv' -Script {
                $uvPath = Get-ToolCommandPath -Name 'uv'
                if (-not $uvPath)
                { $uvPath = (Get-Command uv -ErrorAction SilentlyContinue).Source
                }

                $result = Invoke-UpdateProcess -FilePath 'uv' -ArgumentList @('self', 'update') -SuccessExitCodes @(0, 1) -PassThru
                $outText = (@($result.Output) | Out-String).Trim()
                $managedMessagePattern = '(standalone installation|managed install|installed through another package manager|self-update is only available|cannot be self-updated)'
                $managedPathPattern = '\\(Python\d+\\Scripts|pipx\\venvs|Microsoft\\WinGet\\Packages|scoop\\apps|chocolatey\\lib|WindowsApps)\\'

                if ($outText -match $managedMessagePattern -or ($result.ExitCode -ne 0 -and $uvPath -and $uvPath -match $managedPathPattern))
                {
                    Set-TaskStatus -Status 'Warn' -Reason 'uv self-update skipped because active uv is managed by another installer'
                    Write-Output "uv self-update skipped: '$uvPath' is managed by another installer."
                    if ($outText)
                    { Write-Output $outText
                    }
                    return
                }

                if ($result.ExitCode -ne 0)
                {
                    if ($outText)
                    { Write-Output $outText
                    }
                    $global:LASTEXITCODE = $result.ExitCode
                    throw "uv self-update failed with exit code $($result.ExitCode)."
                }

                if ($outText)
                { Write-Output $outText
                }
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
                Invoke-UpdateProcess -FilePath 'poetry' -ArgumentList @('self', 'update') -Retries 1
            } -Tags @('python'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'rustup' -Category 'systems-language' -RequiresCommand 'rustup' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                Invoke-UpdateProcess -FilePath 'rustup' -ArgumentList @('update') -Retries 1
            } -Tags @('rust') -Resources @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'cargo' -Category 'systems-language' -RequiresCommand 'cargo' -Disabled:$SkipRust -DisabledReason 'disabled by -SkipRust' -Script {
                if (-not (Get-Command cargo-install-update -ErrorAction SilentlyContinue))
                {
                    Invoke-UpdateProcess -FilePath 'cargo' -ArgumentList @('install', 'cargo-update', '-q') -Retries 1
                }
                Invoke-UpdateProcess -FilePath 'cargo' -ArgumentList @('install-update', '-a') -Retries 1
            } -Tags @('rust') -Resources @('rust'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'go' -Category 'systems-language' -RequiresCommand 'go' -Disabled:$SkipGo -DisabledReason 'disabled by -SkipGo' -Script {
                Invoke-UpdateProcess -FilePath 'go' -ArgumentList @('install', 'golang.org/x/tools/gopls@latest') -Retries 1
            } -Tags @('go'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'flutter' -Category 'systems-language' -RequiresCommand 'flutter' -Disabled:$SkipFlutter -DisabledReason 'disabled by -SkipFlutter' -Script {
                Invoke-UpdateProcess -FilePath 'flutter' -ArgumentList @('upgrade') -Retries 1
            } -Tags @('flutter'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'juliaup' -Category 'systems-language' -RequiresCommand 'juliaup' -Script {
                Invoke-UpdateProcess -FilePath 'juliaup' -ArgumentList @('update') -Retries 1 -TimeoutSec 1800
            } -Tags @('julia'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-tools' -Category 'dotnet' -RequiresCommand 'dotnet' -Script {
                Invoke-UpdateProcess -FilePath 'dotnet' -ArgumentList @('tool', 'update', '--global', '--all') -Retries 1
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'dotnet-workloads' -Category 'dotnet' -RequiresCommand 'dotnet' -RequiresAdmin -Script {
                Invoke-UpdateProcess -FilePath 'dotnet' -ArgumentList @('workload', 'update') -TimeoutSec 3600 -Retries 1
            } -Tags @('dotnet'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'ruby-gems' -Category 'runtime' -RequiresCommand 'gem' -Disabled:$SkipRuby -DisabledReason 'disabled by -SkipRuby' -Script {
                Invoke-UpdateProcess -FilePath 'gem' -ArgumentList @('update', '--system') -Retries 1
                Invoke-UpdateProcess -FilePath 'gem' -ArgumentList @('update') -Retries 1
            } -Tags @('ruby'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'composer' -Category 'runtime' -RequiresCommand 'composer' -Disabled:$SkipComposer -DisabledReason 'disabled by -SkipComposer' -Script {
                Invoke-UpdateProcess -FilePath 'composer' -ArgumentList @('self-update') -Retries 1
            } -Tags @('php'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'yt-dlp' -Category 'media-tools' -RequiresCommand 'yt-dlp' -Script {
                $toolPath = Get-ToolCommandPath -Name 'yt-dlp'
                if ($toolPath -match '\\pipx\\venvs\\')
                {
                    if (Get-Command pipx -ErrorAction SilentlyContinue)
                    {
                        Invoke-UpdateProcess -FilePath 'pipx' -ArgumentList @('upgrade', 'yt-dlp') -Retries 1
                        return
                    }
                    Write-Output 'yt-dlp appears to be managed by pipx, but pipx was not found.'
                    return
                }
                if ($toolPath -match '\\Python\d+\\Scripts\\|\\Python\\Python\d+\\Scripts\\')
                {
                    if (Get-Command python -ErrorAction SilentlyContinue)
                    {
                        Invoke-UpdateProcess -FilePath 'python' -ArgumentList @('-m', 'pip', 'install', '--upgrade', 'yt-dlp') -Retries 1
                        return
                    }
                    Write-Output 'yt-dlp appears to be managed by pip, but python was not found.'
                    return
                }
                if ($toolPath -match '\\(scoop\\apps|chocolatey\\lib|Microsoft\\WinGet\\Packages|WindowsApps)\\')
                {
                    Write-Output "yt-dlp is managed by another package manager: $toolPath"
                    return
                }
                Invoke-UpdateProcess -FilePath 'yt-dlp' -ArgumentList @('-U') -Retries 1 -TimeoutSec 900
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

            $models = @($listOutput | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\s+')[0] } | Where-Object { $_ })
            if ($models.Count -eq 0)
            { Write-Output 'No Ollama models found.'; return
            }

            $failed = [System.Collections.Generic.List[string]]::new()
            foreach ($model in $models)
            {
                Write-Output "Updating Ollama model: $model"
                try
                { Invoke-UpdateProcess -FilePath 'ollama' -ArgumentList @('pull', $model) -TimeoutSec $CommandTimeoutSec -Retries 1
                } catch
                { Write-Output $_.Exception.Message; [void]$failed.Add($model)
                }
            }
            if ($failed.Count -gt 0)
            { Write-Output "Ollama models left unchanged: $($failed -join ', ')"
            }
        }
        $tasks.Add((New-UpdateTask -Name 'ollama-models' -Category 'ai' -RequiresCommand 'ollama' -Disabled:(-not $UpdateOllamaModels) -DisabledReason 'use -UpdateOllamaModels to refresh local models' -TimeoutSec 7200 -Script $ollamaScript -Tags @('ai') -Resources @('ollama'))) | Out-Null

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
        }
    }
    $tasks.Add((New-UpdateTask -Name 'cleanup' -Category 'maintenance' -Disabled:$SkipCleanup -DisabledReason 'disabled by -SkipCleanup' -TimeoutSec 3600 -Script $cleanupScript -Tags @('maintenance'))) | Out-Null

    foreach ($task in $tasks)
    {
        switch ($task.Id)
        {
            'winget'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages       = $wingetSkip
                    ProtectedPackages  = $wingetProtected
                    IncludeProtected   = [bool]$IncludeProtectedApps
                    UseSilentInstallers = [bool]$useSilentInstallers
                } -Force
            }
            'chocolatey'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages      = $chocoSkip
                    ProtectedPackages = $chocoProtected
                    IncludeProtected  = [bool]$IncludeProtectedApps
                } -Force
            }
            'store-apps'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{
                    SkipPackages       = $storeSkip
                    ProtectedPackages  = $storeProtected
                    IncludeProtected   = [bool]$IncludeProtectedApps
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
            'cleanup'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ Days = $tempCleanupDays; Deep = [bool]$DeepClean; SkipDestructive = [bool]$SkipDestructive } -Force
            }
            'ollama-models'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ CommandTimeoutSec = $OllamaTimeoutSec } -Force
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

    Start-ThreadJob -Name $Task.Name -ScriptBlock {
        param($TaskName, $TaskId, $Category, $ScriptText, $ArgumentMap, $Attempt, $TaskTimeoutSec)

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
                    $psi.CreateNoWindow = $true

                    $process = [System.Diagnostics.Process]::Start($psi)
                    if (-not $process)
                    { throw "Process did not start: $FilePath"
                    }
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
                    if (-not $PassThru)
                    { Write-Output ("Retrying {0} after exit code {1} (attempt {2}/{3})" -f $FilePath, $lastExitCode, ($attemptNumber + 1), ($Retries + 1))
                    }
                    Start-Sleep -Seconds ([Math]::Min(10, [Math]::Max(1, $attemptNumber * 2)))
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
    } -ArgumentList $Task.Name, $Task.Id, $Task.Category, $scriptText, $argumentMap, $Attempt, $taskTimeoutSec
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

            if ($result.Status -eq 'Failed' -and $item.Attempt -le $RetryCount)
            {
                Write-Status ("[{0}] failed; queueing retry {1}/{2}" -f $task.Name, $item.Attempt, $RetryCount) -Level Warning
                foreach ($line in $combinedOutput)
                { Write-Log -Message ("[{0}] attempt {1}: {2}" -f $task.Name, $item.Attempt, $line) -Level Warning
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
                { Write-Log -Message ("[{0}] {1}" -f $task.Name, $line) -Level Muted
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
        SucceededCount   = @($Results | Where-Object Status -eq 'Succeeded').Count
        WarningCount     = @($Results | Where-Object { $_.Status -in @('Warn', 'Partial') }).Count
        FailedCount      = @($Results | Where-Object { $_.Status -in @('Failed', 'TimedOut') }).Count
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

    $taskName = 'DailySystemUpdate'
    $taskArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath),
        '-SkipReboot', '-SkipWSL', '-SkipWindowsUpdate', '-Quiet'
    ) -join ' '

    if ($script:IsSimulation)
    { Write-Status "[DryRun] Would register scheduled task $taskName at $ScheduleTime." -Level Info; return
    }

    $action = New-ScheduledTaskAction -Execute $pwsh -Argument $taskArgs
    $trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 4)
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
    Write-Status "Scheduled $taskName daily at $ScheduleTime." -Level Success
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

if (Invoke-SelfElevation)
{ exit 0
}

$isAdmin = Test-IsAdmin
if (-not $isAdmin -and -not $NoElevate)
{
    Write-Status 'Running without Administrator. Admin-only tasks will be skipped. Use -AutoElevate for a full run.' -Level Info
}

if ($Schedule)
{ Register-UpdateSchedule; exit 0
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

$allTasks = @(Get-UpdateTasks)
$filtered = Get-FilteredTasks -Tasks $allTasks -IsAdmin $isAdmin
$plannedTasks = @($filtered.Planned)
$skippedTasks = @($filtered.Skipped)

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

if ($plannedTasks.Count -eq 0)
{
    Write-Status 'No runnable update tasks were found.' -Level Info
    $summary = Save-RunSummary -Results @() -Skipped $skippedTasks -Planned $plannedTasks
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
    exit 0
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
$runNotes = @(Get-RunNotes -Results $results -Skipped $skippedTasks)
$summary = Save-RunSummary -Results $results -Skipped $skippedTasks -Planned $plannedTasks -Notes $runNotes

if ($WhatChanged)
{ Show-WhatChanged -CurrentSummary $summary
}

$failed = @($results | Where-Object { $_.Status -in @('Failed', 'TimedOut') })
$succeeded = @($results | Where-Object { $_.Status -eq 'Succeeded' })
$warned = @($results | Where-Object { $_.Status -in @('Warn', 'Partial') })
$elapsed = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')

Write-Host ''
if ($failed.Count -gt 0)
{
    Write-Status ("Completed with {0} succeeded, {1} failed/timed out, {2} skipped in {3}." -f $succeeded.Count, $failed.Count, $skippedTasks.Count, $elapsed) -Level Warning
    Show-RunNotes -Notes $runNotes
    if ($summary.SummaryWritten)
    { Write-Status "Summary: $script:JsonSummaryPath" -Level Muted
    }
    if (Test-Path -LiteralPath $script:LogPath)
    { Write-Status "Log: $script:LogPath" -Level Muted
    }
    exit 1
}

if ($warned.Count -gt 0)
{
    Write-Status ("Completed with {0} succeeded, {1} warning/partial, {2} skipped in {3}." -f $succeeded.Count, $warned.Count, $skippedTasks.Count, $elapsed) -Level Warning
} else
{
    Write-Status ("All runnable tasks completed: {0} succeeded, {1} skipped in {2}." -f $succeeded.Count, $skippedTasks.Count, $elapsed) -Level Success
}
Show-RunNotes -Notes $runNotes
if ($summary.SummaryWritten)
{ Write-Status "Summary: $script:JsonSummaryPath" -Level Muted
}
if (Test-Path -LiteralPath $script:LogPath)
{ Write-Status "Log: $script:LogPath" -Level Muted
}

if (-not $SkipReboot -and $IsWindows)
{
    $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
    (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
    if ($pendingReboot)
    { Write-Status 'A reboot appears to be pending.' -Level Info
    }
}
