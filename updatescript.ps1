#requires -version 7.0
<#
.SYNOPSIS
    Updates common Windows 11 package managers, developer tools, runtimes, WSL distros, Defender, and maintenance tasks.

.VERSION
    Update-Everything v6.3.1-fixed

.NOTES
    Main fixes versus v6.2.0:
      - Correctly launches .cmd/.bat shims such as VS Code's code.cmd through cmd.exe.
      - Finds VS Code CLI from PATH and common install locations.
      - Avoids stale $LASTEXITCODE leakage between tasks.
      - Uses WUA COM as the primary Windows Update fallback instead of relying only on UsoClient.
      - Adds Defender MpCmdRun fallback.
      - Handles npm locked @openai/codex installs and stale npm temp folders more gracefully.
      - Prevents winget and msstore winget tasks from running at the same time.
      - Makes WSL distro updates non-interactive and treats sudo-password skips as non-fatal.
#>

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

$script:Version = '6.3.1-fixed'
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
        'rustup', 'cargo', 'go', 'pip', 'pipx', 'uv', 'uv-tools',
        'poetry', 'composer', 'ruby-gems', 'flutter', 'dotnet-tools',
        'dotnet-workloads', 'vscode-extensions', 'powershell-modules',
        'powershell-help', 'ollama-models'
    )
    UltraFastSkip      = @('windows-update', 'store-apps', 'wsl', 'wsl-distros', 'defender', 'cleanup')
    SkipManagers       = @()
    WingetSkipPackages = @('Microsoft.VisualStudio.BuildTools')
    PipSkipPackages    = @()
    NpmSkipPackages    = @()
    LogRetentionDays   = 14
    TempCleanupDays    = 7
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

function Test-Command
{
    param([Parameter(Mandatory)][string]$Name)
    if ($script:CommandCache.ContainsKey($Name))
    { return $script:CommandCache[$Name]
    }

    $found = $false
    if ($Name -eq 'code')
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
    if ($Name -eq 'code')
    { return Get-VSCodeCommandPath
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
    $pipSkip = ConvertTo-StringArray $script:Config.PipSkipPackages
    $npmSkip = ConvertTo-FilterList $script:Config.NpmSkipPackages
    $tempCleanupDays = [int]$script:Config.TempCleanupDays
    $tasks = [System.Collections.Generic.List[object]]::new()

    $wingetScript = {
        param([string[]]$SkipPackages)

        $skipSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($package in @($SkipPackages))
        {
            if (-not [string]::IsNullOrWhiteSpace($package))
            { [void]$skipSet.Add($package.Trim())
            }
        }

        if ($skipSet.Count -gt 0)
        { Write-Output "Winget skip list detected: $($skipSet -join ', ')"
        }

        $listOutput = Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--include-unknown', '--accept-source-agreements') -SuccessExitCodes @(0, -1978335189)
        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        $idColumnStart = $null
        $idColumnEnd = $null

        foreach ($line in $listOutput)
        {
            $text = [string]$line
            if ($text -match 'No installed package found matching input criteria|No available upgrade found|No applicable update found')
            {
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

        if ($ids.Count -eq 0)
        {
            Write-Output 'No winget upgrades found.'
            return
        }

        $failed = [System.Collections.Generic.List[string]]::new()
        foreach ($id in ($ids | Sort-Object -Unique))
        {
            if ($skipSet.Contains($id))
            {
                Write-Output "Skipping winget package: $id"
                continue
            }

            Write-Output "Upgrading winget package: $id"
            try
            {
                Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--id', $id, '--exact', '--include-unknown', '--silent', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements') -Retries 1
            } catch
            {
                Write-Output $_.Exception.Message
                [void]$failed.Add($id)
            }
        }

        if ($failed.Count -gt 0)
        { throw "winget failed packages: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'winget' -Category 'package-manager' -RequiresCommand 'winget' -TimeoutSec $WingetTimeoutSec -Script $wingetScript -Tags @('windows') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'scoop' -Category 'package-manager' -RequiresCommand 'scoop' -Script {
                Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update')
                Invoke-UpdateProcess -FilePath 'scoop' -ArgumentList @('update', '*') -Retries 1
            } -Tags @('windows'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'chocolatey' -Category 'package-manager' -RequiresCommand 'choco' -RequiresAdmin -Script {
                Invoke-UpdateProcess -FilePath 'choco' -ArgumentList @('upgrade', 'all', '-y', '--no-progress', '--limit-output') -Retries 1
            } -Tags @('windows') -Resources @('chocolatey'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'store-apps' -Category 'system' -RequiresCommand 'winget' -Disabled:$SkipStoreApps -DisabledReason 'disabled by -SkipStoreApps' -TimeoutSec $WingetTimeoutSec -Script {
                Write-Output 'store-apps command: winget upgrade --all --source msstore --include-unknown --silent --disable-interactivity --accept-package-agreements --accept-source-agreements'
                Invoke-UpdateProcess -FilePath 'winget' -ArgumentList @('upgrade', '--all', '--source', 'msstore', '--include-unknown', '--silent', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements') -Retries 1
            } -Tags @('windows', 'store') -Resources @('winget'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'windows-update' -Category 'system' -RequiresAdmin -Disabled:$SkipWindowsUpdate -DisabledReason 'disabled by -SkipWindowsUpdate' -TimeoutSec 7200 -Script {
                if (Get-Command Install-WindowsUpdate -ErrorAction SilentlyContinue)
                {
                    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop | Out-String | Write-Output
                    return
                }

                Write-Output 'PSWindowsUpdate not found; using Windows Update Agent COM fallback.'
                $session = New-Object -ComObject Microsoft.Update.Session
                $session.ClientApplicationID = 'Update-Everything'
                $searcher = $session.CreateUpdateSearcher()
                $result = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
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
                $downloadResult = $downloader.Download()
                Write-Output "Windows Update download result code: $($downloadResult.ResultCode)"
                if ($downloadResult.ResultCode -notin @(2, 3))
                { throw "Windows Update download failed with result code $($downloadResult.ResultCode)"
                }

                $installer = $session.CreateUpdateInstaller()
                $installer.Updates = $updates
                $installResult = $installer.Install()
                Write-Output "Windows Update install result code: $($installResult.ResultCode); reboot required: $($installResult.RebootRequired)"
                if ($installResult.ResultCode -notin @(2, 3))
                { throw "Windows Update install failed with result code $($installResult.ResultCode)"
                }
            } -Tags @('windows') -Resources @('windows-update'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'defender' -Category 'system' -RequiresCommand 'Update-MpSignature' -Disabled:$SkipDefender -DisabledReason 'disabled by -SkipDefender' -Script {
                try
                {
                    $primaryOutput = Update-MpSignature -ErrorAction Stop | Out-String
                    if (-not [string]::IsNullOrWhiteSpace($primaryOutput))
                    { Write-Output $primaryOutput.Trim()
                    }
                    Write-Output 'Defender signature update completed through Update-MpSignature.'
                } catch
                {
                    Write-Output "Update-MpSignature reported an issue; trying MpCmdRun fallback: $($_.Exception.Message)"
                    $mpCmdCandidates = @(
                        (Join-Path $env:ProgramFiles 'Windows Defender\MpCmdRun.exe'),
                        (Join-Path ${env:ProgramFiles(x86)} 'Windows Defender\MpCmdRun.exe'),
                        'MpCmdRun.exe'
                    ) | Where-Object { $_ }
                    $mpCmd = $mpCmdCandidates | Where-Object { ($_ -eq 'MpCmdRun.exe') -or (Test-Path -LiteralPath $_) } | Select-Object -First 1
                    Invoke-UpdateProcess -FilePath $mpCmd -ArgumentList @('-SignatureUpdate') -Retries 1 -TimeoutSec 900
                    Write-Output 'Defender signature update completed through MpCmdRun fallback.'
                }
            } -Tags @('windows', 'security') -Resources @('defender'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl' -Category 'system' -RequiresCommand 'wsl' -Disabled:$SkipWSL -DisabledReason 'disabled by -SkipWSL' -Script {
                $wslOutput = Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('--update') -Retries 0 -SuccessExitCodes @(0, -1)
                $wslText = ($wslOutput | Out-String).Trim()
                if ($wslText)
                { Write-Output $wslText
                }
                $wslExitCode = [int]$global:LASTEXITCODE
                if ($wslExitCode -ne 0)
                {
                    if ($wslText -match 'Forbidden \(403\)|0x80190193|Wsl/UpdatePackage')
                    {
                        Write-Output 'WSL kernel/package update endpoint returned 403. Treating as non-fatal; installed WSL distros can still be updated.'
                        $global:LASTEXITCODE = 0
                        return
                    }
                    throw "wsl --update failed with exit code $wslExitCode"
                }
            } -Tags @('windows', 'linux') -Resources @('wsl'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'wsl-distros' -Category 'system' -RequiresCommand 'wsl' -Disabled:($SkipWSL -or $SkipWSLDistros) -DisabledReason 'disabled by WSL skip switch' -TimeoutSec 3600 -Script {
                $distros = @(Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('-l', '-q') |
                        ForEach-Object { ([string]$_).Replace([string][char]0, '').Trim() } |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

                    if ($distros.Count -eq 0)
                    { Write-Output 'No WSL distros found.'; return
                    }

                    $linuxScript = @'
set -u
if command -v apt >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    sudo -n apt update && sudo -n DEBIAN_FRONTEND=noninteractive apt -y upgrade && sudo -n apt -y autoremove
  else
    echo "Skipping apt: sudo requires a password"
  fi
elif command -v pacman >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    sudo -n pacman -Syu --noconfirm
  else
    echo "Skipping pacman: sudo requires a password"
  fi
else
  echo "No supported package manager found"
fi
'@

                    $failedDistros = [System.Collections.Generic.List[string]]::new()
                    foreach ($distro in $distros)
                    {
                        Write-Output "Updating WSL distro: $distro"
                        try
                        {
                            Invoke-UpdateProcess -FilePath 'wsl' -ArgumentList @('--distribution', $distro, '--exec', 'sh', '-lc', $linuxScript) -TimeoutSec 1800 -Retries 0
                        } catch
                        {
                            Write-Output $_.Exception.Message
                            [void]$failedDistros.Add($distro)
                        }
                    }
                    if ($failedDistros.Count -gt 0)
                    { throw "WSL distro updates failed: $($failedDistros -join ', ')"
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
        { throw "npm failed packages: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'npm' -Category 'javascript' -RequiresCommand 'npm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script $npmScript -Tags @('node') -Resources @('npm'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pnpm' -Category 'javascript' -RequiresCommand 'pnpm' -Disabled:$SkipNode -DisabledReason 'disabled by -SkipNode' -Script {
                Invoke-UpdateProcess -FilePath 'pnpm' -ArgumentList @('self-update') -Retries 1
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
        { throw "pip failed packages: $($failed -join ', ')"
        }
    }
    $tasks.Add((New-UpdateTask -Name 'pip' -Category 'python' -RequiresCommand 'python' -Script $pipScript -Tags @('python') -Resources @('pip'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'pipx' -Category 'python' -RequiresCommand 'pipx' -Script {
                Invoke-UpdateProcess -FilePath 'pipx' -ArgumentList @('upgrade-all') -Retries 1
            } -Tags @('python') -Resources @('pip'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv' -Category 'python' -RequiresCommand 'uv' -Script {
                $out = Invoke-UpdateProcess -FilePath 'uv' -ArgumentList @('self', 'update') -SuccessExitCodes @(0, 1)
                $outText = ($out | Out-String).Trim()
                if ($outText -match 'standalone installation')
                {
                    $uvPath = (Get-Command uv -ErrorAction SilentlyContinue).Source
                    Write-Output "uv self-update skipped: '$uvPath' is a managed install."
                    return
                }
                if ($outText)
                { Write-Output $outText
                }
            } -Tags @('python') -Resources @('uv'))) | Out-Null

    $tasks.Add((New-UpdateTask -Name 'uv-tools' -Category 'python' -RequiresCommand 'uv' -Disabled:$SkipUVTools -DisabledReason 'disabled by -SkipUVTools' -Script {
                Invoke-UpdateProcess -FilePath 'uv' -ArgumentList @('tool', 'upgrade', '--all') -Retries 1
            } -Tags @('python') -Resources @('uv'))) | Out-Null

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
                    { throw "GitHub extension updates failed: $($failedExtensions -join ', ')"
                    }
                } -Tags @('github'))) | Out-Null

        $tasks.Add((New-UpdateTask -Name 'powershell-modules' -Category 'powershell' -Disabled:$SkipPowerShellModules -DisabledReason 'disabled by -SkipPowerShellModules' -Script {
                    if (Get-Command Update-PSResource -ErrorAction SilentlyContinue)
                    {
                        $moduleArgs = @{ Name = '*'; ErrorAction = 'Stop' }
                        $cmd = Get-Command Update-PSResource -ErrorAction SilentlyContinue
                        if ($cmd.Parameters.ContainsKey('AcceptLicense'))
                        { $moduleArgs.AcceptLicense = $true
                        }
                        Update-PSResource @moduleArgs | Out-String | Write-Output
                    } elseif (Get-Command Update-Module -ErrorAction SilentlyContinue)
                    {
                        $failedModules = [System.Collections.Generic.List[string]]::new()
                        Get-InstalledModule -ErrorAction SilentlyContinue | ForEach-Object {
                            $moduleName = $_.Name
                            try
                            { Update-Module -Name $moduleName -Force -ErrorAction Stop
                            } catch
                            { Write-Output "PowerShell module update failed for ${moduleName}: $($_.Exception.Message)"; [void]$failedModules.Add($moduleName)
                            }
                        }
                        if ($failedModules.Count -gt 0)
                        { throw "PowerShell module updates failed: $($failedModules -join ', ')"
                        }
                    } else
                    {
                        Write-Output 'No PowerShell module updater found.'
                    }
                } -Tags @('powershell') -Resources @('powershell-gallery'))) | Out-Null

        $tasks.Add((New-UpdateTask -Name 'powershell-help' -Category 'powershell' -Script {
                    Update-Help -Force -ErrorAction Stop | Out-String | Write-Output
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
            { throw "Ollama model updates failed: $($failed -join ', ')"
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
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $wingetSkip } -Force
            }
            'pip'
            { $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{ SkipPackages = $pipSkip } -Force
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
    Write-Host 'Skipped tasks' -ForegroundColor DarkGray
    if ($Skipped.Count -eq 0)
    { Write-Host '  none' -ForegroundColor DarkGray
    } else
    { $Skipped | Sort-Object Category, Name | Format-Table Name, Category, Reason -AutoSize
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
            if ($FilePath -eq 'code')
            {
                $source = Get-ToolCommandPath -Name 'code'
            } elseif (-not (Test-Path -LiteralPath $FilePath -ErrorAction SilentlyContinue))
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
                [int]$Retries = 0
            )

            $effectiveTimeoutSec = if ($TimeoutSec -gt 0)
            { $TimeoutSec
            } else
            { [Math]::Max(30, $TaskTimeoutSec - 5)
            }
            $attemptNumber = 0
            $lastOutput = @()
            $lastExitCode = 0

            do
            {
                $attemptNumber++
                $process = $null
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
                    $lastExitCode = 1
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
                    return @($lastOutput)
                }

                foreach ($line in $lastOutput)
                { Write-Output $line
                }
                if ($attemptNumber -le $Retries)
                {
                    Write-Output ("Retrying {0} after exit code {1} (attempt {2}/{3})" -f $FilePath, $lastExitCode, ($attemptNumber + 1), ($Retries + 1))
                    Start-Sleep -Seconds ([Math]::Min(10, [Math]::Max(1, $attemptNumber * 2)))
                }
            } while ($attemptNumber -le $Retries)

            $global:LASTEXITCODE = $lastExitCode
            throw "$FilePath failed with exit code $lastExitCode"
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
    param([object[]]$Results, [object[]]$Skipped, [object[]]$Planned)

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
        FailedCount      = @($Results | Where-Object { $_.Status -in @('Failed', 'TimedOut') }).Count
        SkippedCount     = $Skipped.Count
        Results          = @($Results)
        Skipped          = @($Skipped)
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
        Write-Status 'No previous run summary found for -WhatChanged.' -Level Warning
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
        Write-Status "WhatChanged failed: $($_.Exception.Message)" -Level Warning
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
        '-SkipReboot', '-NoPause', '-SkipWSL', '-SkipWindowsUpdate', '-Quiet'
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

Write-Status ("Dispatching {0} task(s). Skipped: {1}" -f $plannedTasks.Count, $skippedTasks.Count) -Level Info
$results = Invoke-TaskQueue -Tasks $plannedTasks -Throttle $ParallelThrottle
$summary = Save-RunSummary -Results $results -Skipped $skippedTasks -Planned $plannedTasks

if ($WhatChanged)
{ Show-WhatChanged -CurrentSummary $summary
}

$failed = @($results | Where-Object { $_.Status -in @('Failed', 'TimedOut') })
$succeeded = @($results | Where-Object { $_.Status -eq 'Succeeded' })
$elapsed = ((Get-Date) - $script:StartTime).ToString('hh\:mm\:ss')

Write-Host ''
if ($failed.Count -gt 0)
{
    Write-Status ("Completed with {0} succeeded, {1} failed/timed out, {2} skipped in {3}." -f $succeeded.Count, $failed.Count, $skippedTasks.Count, $elapsed) -Level Warning
    if ($summary.SummaryWritten)
    { Write-Status "Summary: $script:JsonSummaryPath" -Level Muted
    }
    if (Test-Path -LiteralPath $script:LogPath)
    { Write-Status "Log: $script:LogPath" -Level Muted
    }
    exit 1
}

Write-Status ("All runnable tasks completed: {0} succeeded, {1} skipped in {2}." -f $succeeded.Count, $skippedTasks.Count, $elapsed) -Level Success
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
    { Write-Status 'A reboot appears to be pending.' -Level Warning
    }
}
