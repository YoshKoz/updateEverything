#Requires -Version 7.0
#Requires -Modules Pester

BeforeAll {
    $script:ScriptPath = Join-Path $PSScriptRoot 'updatescript.ps1'
    $script:PwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    if (-not $script:PwshPath) {
        $pwshFileName = if ($IsWindows) { 'pwsh.exe' } else { 'pwsh' }
        $script:PwshPath = Join-Path $PSHOME $pwshFileName
    }
    $script:InvokeTestPwsh = {
        param([Parameter(Mandatory)][string[]]$Arguments)

        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName = $script:PwshPath
        foreach ($argument in $Arguments) {
            [void]$psi.ArgumentList.Add($argument)
        }
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($psi)
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()

        [pscustomobject]@{
            ExitCode = $process.ExitCode
            Output   = (($stdout.GetAwaiter().GetResult() + $stderr.GetAwaiter().GetResult()) -replace "`0", '')
        }
    }
    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "Could not parse updatescript.ps1: $($errors[0].Message)"
    }

    foreach ($funcDef in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        . ([scriptblock]::Create($funcDef.Extent.Text))
    }

    $script:Version = 'test-version'
    $script:StartTime = Get-Date
    $script:RunId = 'test-run'
    $script:CommandCache = @{}
    $script:StateDirWasProvided = $true
    $script:LogPathWasProvided = $true
    $script:JsonSummaryPathWasProvided = $true
    $script:LogWriteWarningEmitted = $false
    $script:IsSimulation = $false
    $script:Config = [ordered]@{
        FastModeSkip                    = @('npm', 'cargo')
        UltraFastSkip                   = @('windows-update', 'cleanup')
        SkipManagers                    = @()
        WingetSkipPackages              = @()
        WingetProtectedPackages         = @()
        ExtraWingetProtectedPackages    = @()
        ChocolateySkipPackages          = @()
        ChocolateyProtectedPackages     = @()
        ExtraChocolateyProtectedPackages = @()
        StoreAppSkipPackages            = @()
        StoreAppProtectedPackages       = @()
        ExtraStoreAppProtectedPackages  = @()
        PipSkipPackages                 = @()
        PipIgnoreHealthPackages         = @()
        NpmSkipPackages                 = @()
        LogRetentionDays                = 14
        TempCleanupDays                 = 7
    }

    $script:TestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("update-everything-tests-" + [guid]::NewGuid().ToString('N'))
    $script:StateDir = Join-Path $script:TestRoot 'state'
    $script:LogDir = Join-Path $script:StateDir 'logs'
    $script:DefaultJsonSummaryPath = Join-Path $script:StateDir 'last-run.json'
    $script:PreviousJsonSummaryPath = Join-Path $script:StateDir 'previous-run.json'
    $script:LogPath = Join-Path $script:LogDir 'test.log'
    $script:JsonSummaryPath = Join-Path $script:StateDir 'summary.json'

    $script:FastMode = $false
    $script:UltraFast = $false
    $script:Only = @()
    $script:Skip = @()
    $script:ParallelThrottle = 2
    $script:ShowSkipped = $false
}

AfterAll {
    Remove-Item -LiteralPath $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ── ConvertTo-TaskId ──────────────────────────────────────────────────────────

Describe 'ConvertTo-TaskId' {
    It 'normalizes names into task ids' {
        ConvertTo-TaskId 'My Tool / Name' | Should -Be 'my-tool-name'
    }

    It 'lowercases and strips leading/trailing dashes' {
        ConvertTo-TaskId '  --UPPER--  ' | Should -Be 'upper'
    }

    It 'collapses multiple separators into one dash' {
        ConvertTo-TaskId 'a  b__c' | Should -Be 'a-b-c'
    }

    It 'handles single word unchanged' {
        ConvertTo-TaskId 'winget' | Should -Be 'winget'
    }
}

# ── New-UpdateTask ────────────────────────────────────────────────────────────

Describe 'New-UpdateTask' {
    It 'stores normalized id, tags, resources, and timeout' {
        $task = New-UpdateTask -Name 'VS Code Extensions' -Category 'dev-tools' -Script {} -Tags @('VSCode') -Resources @('VS Code') -TimeoutSec 42
        $task.Id | Should -Be 'vs-code-extensions'
        $task.Tags | Should -Be @('vscode')
        $task.Resources | Should -Be @('vs-code')
        $task.TimeoutSec | Should -Be 42
    }

    It 'sets RequiresAdmin to false by default' {
        $task = New-UpdateTask -Name 'simple' -Category 'test' -Script {}
        $task.RequiresAdmin | Should -BeFalse
    }

    It 'stores Disabled and DisabledReason' {
        $task = New-UpdateTask -Name 'skip me' -Category 'test' -Script {} -Disabled -DisabledReason 'not installed'
        $task.Disabled | Should -BeTrue
        $task.DisabledReason | Should -Be 'not installed'
    }

    It 'normalizes multiple tags' {
        $task = New-UpdateTask -Name 't' -Category 'c' -Script {} -Tags @('Tag One', 'TAG-TWO')
        $task.Tags | Should -Contain 'tag-one'
        $task.Tags | Should -Contain 'tag-two'
    }
}

# -- ConvertFrom-WingetUpgradeOutput ------------------------------------------------

Describe 'ConvertFrom-WingetUpgradeOutput' {
    It 'parses regular winget upgrade table rows' {
        $output = @(
            'Name                     Id                Version Available Source',
            '------------------------------------------------------------------',
            'Git                      Git.Git           2.44.0  2.45.0    winget'
        )

        $packages = @(ConvertFrom-WingetUpgradeOutput -Output $output)

        $packages | Should -HaveCount 1
        $packages[0].Id | Should -Be 'Git.Git'
        $packages[0].Version | Should -Be '2.44.0'
        $packages[0].Available | Should -Be '2.45.0'
        $packages[0].IsUnknown | Should -BeFalse
    }

    It 'parses unknown-version rows with a display version before the package id' {
        $output = @(
            'Node.js 22.14.0 OpenJS.NodeJS Unknown 22.15.0 winget'
        )

        $packages = @(ConvertFrom-WingetUpgradeOutput -Output $output)

        $packages | Should -HaveCount 1
        $packages[0].Name | Should -Be 'Node.js'
        $packages[0].Id | Should -Be 'OpenJS.NodeJS'
        $packages[0].DisplayVersion | Should -Be '22.14.0'
        $packages[0].IsUnknown | Should -BeTrue
        $packages[0].InstalledLooksCurrent | Should -BeFalse
    }

    It 'marks unknown-version rows as current when display and available versions match' {
        $packages = @(ConvertFrom-WingetUpgradeOutput -Output @('Node.js 22.15.0 OpenJS.NodeJS Unknown 22.15.0 winget'))

        $packages | Should -HaveCount 1
        $packages[0].InstalledLooksCurrent | Should -BeTrue
    }

    It 'parses compact Microsoft Store ids from source-filtered output' {
        $output = @(
            'Name                     Id           Version Available Source',
            '-------------------------------------------------------------',
            'Windows Terminal         9N0DX20HK701 1.20.0  1.21.0    msstore'
        )

        $packages = @(ConvertFrom-WingetUpgradeOutput -Output $output)

        $packages | Should -HaveCount 1
        $packages[0].Id | Should -Be '9N0DX20HK701'
        $packages[0].Source | Should -Be 'msstore'
    }

    It 'parses rows when winget omits the Source column' {
        $packages = @(ConvertFrom-WingetUpgradeOutput -Output @(
                'Name                     Id                Version Available',
                '----------------------------------------------------------',
                'Git                      Git.Git           2.44.0  2.45.0'
            ))

        $packages | Should -HaveCount 1
        $packages[0].Id | Should -Be 'Git.Git'
        $packages[0].Source | Should -BeNullOrEmpty
    }
}

# ── Test-NameMatch ────────────────────────────────────────────────────────────

Describe 'Test-NameMatch' {
    BeforeAll {
        $script:SampleTask = New-UpdateTask -Name 'Windows Update' -Category 'system tools' -Script {} -Tags @('windows', 'system')
    }

    It 'matches by exact task id' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-update') | Should -BeTrue
    }

    It 'matches by category id' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('system-tools') | Should -BeTrue
    }

    It 'matches by wildcard pattern' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-*') | Should -BeTrue
    }

    It 'matches by tag' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows') | Should -BeTrue
    }

    It 'returns false for unrelated pattern' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('npm') | Should -BeFalse
    }

    It 'ignores blank/whitespace patterns' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('', '  ') | Should -BeFalse
    }
}

# ── Get-FilteredTasks ─────────────────────────────────────────────────────────

Describe 'Get-FilteredTasks' {
    BeforeEach {
        $script:FastMode = $false
        $script:UltraFast = $false
        $script:Only = @()
        $script:Skip = @()
        $script:CommandCache = @{}

        $script:TasksUnderTest = @(
            (New-UpdateTask -Name 'winget' -Category 'package-manager' -Script {} -RequiresCommand @('winget')),
            (New-UpdateTask -Name 'windows-update' -Category 'system' -Script {} -RequiresAdmin),
            (New-UpdateTask -Name 'npm' -Category 'javascript' -Script {}),
            (New-UpdateTask -Name 'cleanup' -Category 'maintenance' -Script {}),
            (New-UpdateTask -Name 'disabled-tool' -Category 'misc' -Script {} -Disabled -DisabledReason 'disabled for test')
        )
    }

    It 'skips disabled tasks and admin-only tasks for non-admin runs' {
        Mock Test-Command { $true }

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $false

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Contain 'winget'
        ($result.Skipped | Where-Object Name -eq 'windows-update').Reason | Should -Be 'requires Administrator'
        ($result.Skipped | Where-Object Name -eq 'disabled-tool').Reason | Should -Be 'disabled for test'
    }

    It 'applies -Only filter' {
        Mock Test-Command { $true }
        $script:Only = @('npm')

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Be @('npm')
        ($result.Skipped | Where-Object Name -eq 'cleanup').Reason | Should -Be 'not selected by -Only'
    }

    It 'applies -Skip filter' {
        Mock Test-Command { $true }
        $script:Skip = @('cleanup')

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Not -Contain 'cleanup'
        ($result.Skipped | Where-Object Name -eq 'cleanup').Reason | Should -Be 'skipped by filter'
    }

    It 'applies FastMode skip list from config' {
        Mock Test-Command { $true }
        $script:FastMode = $true

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Skipped | Where-Object Name -eq 'npm').Reason | Should -Be 'skipped by filter'
    }

    It 'applies UltraFast skip list from config (superset of FastMode)' {
        Mock Test-Command { $true }
        $script:UltraFast = $true

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Skipped | Where-Object Name -eq 'cleanup').Reason | Should -Be 'skipped by filter'
        ($result.Skipped | Where-Object Name -eq 'npm').Reason | Should -Be 'skipped by filter'
    }

    It 'skips tasks whose required command is missing' {
        Mock Test-Command {
            param($Name)
            return $Name -ne 'winget'
        }

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Skipped | Where-Object Name -eq 'winget').Reason | Should -Be 'missing command: winget'
    }

    It 'applies SkipManagers from config' {
        Mock Test-Command { $true }
        $script:Config.SkipManagers = @('npm')

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Skipped | Where-Object Name -eq 'npm').Reason | Should -Be 'skipped by filter'

        $script:Config.SkipManagers = @()
    }

    It 'plans admin tasks when IsAdmin is true' {
        Mock Test-Command { $true }

        $result = Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Contain 'windows-update'
    }
}

# ── Test-ResourcesAvailable ───────────────────────────────────────────────────

Describe 'Test-ResourcesAvailable' {
    It 'returns true when task has no resources' {
        $task = New-UpdateTask -Name 'no-res' -Category 'test' -Script {}
        Test-ResourcesAvailable -Task $task -Running @() | Should -BeTrue
    }

    It 'returns true when resource is not held by any running task' {
        $task = New-UpdateTask -Name 'a' -Category 'test' -Script {} -Resources @('winget')
        $other = New-UpdateTask -Name 'b' -Category 'test' -Script {} -Resources @('npm')
        $running = @([pscustomobject]@{ Task = $other })
        Test-ResourcesAvailable -Task $task -Running $running | Should -BeTrue
    }

    It 'returns false when resource is already held by a running task' {
        $task = New-UpdateTask -Name 'a' -Category 'test' -Script {} -Resources @('winget')
        $holder = New-UpdateTask -Name 'b' -Category 'test' -Script {} -Resources @('winget')
        $running = @([pscustomobject]@{ Task = $holder })
        Test-ResourcesAvailable -Task $task -Running $running | Should -BeFalse
    }
}

# ── New-TaskResult ────────────────────────────────────────────────────────────

Describe 'New-TaskResult' {
    It 'stores all fields and rounds duration' {
        $task = New-UpdateTask -Name 'my task' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded' -ExitCode 0 -DurationSeconds 1.23456 -Attempts 2 -Output @('line1', 'line2')

        $result.Name | Should -Be 'my task'
        $result.Id | Should -Be 'my-task'
        $result.Status | Should -Be 'Succeeded'
        $result.DurationSeconds | Should -Be 1.23
        $result.Attempts | Should -Be 2
        $result.OutputPreview | Should -Contain 'line1'
        $result.OutputPreview | Should -Contain 'line2'
    }

    It 'caps OutputPreview at 40 lines' {
        $task = New-UpdateTask -Name 't' -Category 'c' -Script {}
        $lines = 1..60 | ForEach-Object { "line $_" }
        $result = New-TaskResult -Task $task -Status 'Succeeded' -Output $lines
        $result.OutputPreview.Count | Should -Be 40
    }
}

# ── Get-RunNotes ──────────────────────────────────────────────────────────────

Describe 'Get-RunNotes' {
    It 'emits retry note when tasks failed' {
        $task = New-UpdateTask -Name 'cargo' -Category 'rust' -Script {}
        $result = New-TaskResult -Task $task -Status 'Failed'
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        ($notes | Where-Object { $_.Message -match 'cargo' }).Count | Should -BeGreaterThan 0
    }

    It 'emits uv note when uv task fails' {
        $task = New-UpdateTask -Name 'uv' -Category 'python' -Script {}
        $result = New-TaskResult -Task $task -Status 'Failed'
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        ($notes | Where-Object { $_.Message -match 'uv:' }).Count | Should -BeGreaterThan 0
    }

    It 'emits pip-health note when pip-health fails' {
        $task = New-UpdateTask -Name 'pip-health' -Category 'python' -Script {}
        $result = New-TaskResult -Task $task -Status 'Failed'
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        ($notes | Where-Object { $_.Message -match 'pip-health' }).Count | Should -BeGreaterThan 0
    }

    It 'emits pnpm note when output contains broken shim text' {
        $task = New-UpdateTask -Name 'pnpm' -Category 'javascript' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded' -Output @('@pnpm/exe/pnpm.exe was not found')
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        ($notes | Where-Object { $_.Message -match 'pnpm' }).Count | Should -BeGreaterThan 0
    }

    It 'emits oh-my-posh note when oh-my-posh times out' {
        $task = New-UpdateTask -Name 'oh-my-posh' -Category 'shell' -Script {}
        $result = New-TaskResult -Task $task -Status 'TimedOut'
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        ($notes | Where-Object { $_.Message -match 'oh-my-posh' }).Count | Should -BeGreaterThan 0
    }

    It 'emits admin note when admin tasks were skipped' {
        $skipped = [pscustomobject]@{ Name = 'windows-update'; Id = 'windows-update'; Category = 'system'; Status = 'Skipped'; Reason = 'requires Administrator' }
        $notes = Get-RunNotes -Results @() -Skipped @($skipped)
        ($notes | Where-Object { $_.Message -match 'Admin-only' }).Count | Should -BeGreaterThan 0
    }

    It 'deduplicates identical notes' {
        $task1 = New-UpdateTask -Name 'uv' -Category 'python' -Script {}
        $task2 = New-UpdateTask -Name 'uv' -Category 'python' -Script {}
        $r1 = New-TaskResult -Task $task1 -Status 'Failed'
        $r2 = New-TaskResult -Task $task2 -Status 'Failed'
        $notes = Get-RunNotes -Results @($r1, $r2) -Skipped @()
        $uvNotes = @($notes | Where-Object { $_.Message -match 'uv:' })
        $uvNotes.Count | Should -Be 1
    }

    It 'returns empty array when everything succeeded' {
        $task = New-UpdateTask -Name 'npm' -Category 'javascript' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded'
        $notes = Get-RunNotes -Results @($result) -Skipped @()
        $notes.Count | Should -Be 0
    }
}

# ── Split-SkippedTasksForDisplay ──────────────────────────────────────────────

Describe 'Split-SkippedTasksForDisplay' {
    It 'hides missing-command skips by default' {
        $s1 = [pscustomobject]@{ Name = 'a'; Reason = 'missing command: yarn' }
        $s2 = [pscustomobject]@{ Name = 'b'; Reason = 'disabled for test' }
        $split = Split-SkippedTasksForDisplay -Skipped @($s1, $s2)
        $split.Hidden | Should -Contain $s1
        $split.Visible | Should -Contain $s2
    }

    It 'hides opt-in skips' {
        $s = [pscustomobject]@{ Name = 'x'; Reason = 'opt-in via -UpdateOllamaModels' }
        $split = Split-SkippedTasksForDisplay -Skipped @($s)
        $split.Hidden | Should -Contain $s
        $split.Visible.Count | Should -Be 0
    }
}

# ── Initialize-RunStorage / Write-Log / Save-RunSummary ───────────────────────

Describe 'Initialize-RunStorage' {
    It 'creates writable state and log directories' {
        Initialize-RunStorage
        Test-Path -LiteralPath $script:StateDir | Should -BeTrue
        Test-Path -LiteralPath $script:LogDir | Should -BeTrue
    }
}

Describe 'Write-Log' {
    It 'writes log lines with INFO level tag' {
        Initialize-RunStorage
        Write-Log -Message 'test message' -Level Info
        $content = Get-Content -LiteralPath $script:LogPath -Raw
        $content | Should -Match 'test message'
        $content | Should -Match '\[INFO\]'
    }

    It 'writes WARNING level tag' {
        Initialize-RunStorage
        Write-Log -Message 'warn message' -Level Warning
        $content = Get-Content -LiteralPath $script:LogPath -Raw
        $content | Should -Match '\[WARNING\]'
    }
}

Describe 'Save-RunSummary' {
    It 'writes a JSON summary containing version and success counts' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'dummy' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded' -DurationSeconds 1.234

        $summary = Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)

        $summary.Version | Should -Be 'test-version'
        $summary.SummaryWritten | Should -BeTrue

        $json = Get-Content -LiteralPath $script:JsonSummaryPath -Raw | ConvertFrom-Json
        $json.RunId | Should -Be 'test-run'
        $json.SucceededCount | Should -Be 1
        $json.FailedCount | Should -Be 0
    }

    It 'counts failed/timed-out results correctly' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'x' -Category 'test' -Script {}
        $r1 = New-TaskResult -Task $task -Status 'Failed'
        $r2 = New-TaskResult -Task $task -Status 'TimedOut'
        $r3 = New-TaskResult -Task $task -Status 'Succeeded'

        $summary = Save-RunSummary -Results @($r1, $r2, $r3) -Skipped @() -Planned @($task, $task, $task)
        $summary.FailedCount | Should -Be 2
        $summary.SucceededCount | Should -Be 1
    }

    It 'embeds Notes in the written JSON' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'n' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded'
        $note = [pscustomobject]@{ Level = 'Info'; Message = 'test note' }

        Save-RunSummary -Results @($result) -Skipped @() -Planned @($task) -Notes @($note) | Out-Null

        $json = Get-Content -LiteralPath $script:JsonSummaryPath -Raw | ConvertFrom-Json
        ($json.Notes | Where-Object { $_.Message -eq 'test note' }).Count | Should -Be 1
    }
}

# ── Invoke-TaskQueue ──────────────────────────────────────────────────────────

Describe 'Invoke-TaskQueue' {
    It 'runs a simple task and returns Succeeded' {
        $task = New-UpdateTask -Name 'self-test' -Category 'diagnostics' -Script { Write-Output 'ok' } -TimeoutSec 30
        $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force

        $results = Invoke-TaskQueue -Tasks @($task) -Throttle 1 -Confirm:$false

        $results | Should -HaveCount 1
        $results[0].Status | Should -Be 'Succeeded'
        $results[0].OutputPreview | Should -Contain 'ok'
    }

    It 'makes helper functions available inside task jobs' {
        $task = New-UpdateTask -Name 'helper-self-test' -Category 'diagnostics' -Script {
            $pwshPath = Get-ToolCommandPath -Name 'pwsh'
            if (-not $pwshPath) {
                throw 'pwsh command was not resolved inside the task job'
            }
            Invoke-UpdateProcess -FilePath $pwshPath -ArgumentList @('-NoProfile', '-Command', 'Write-Output helper-ok')
        } -TimeoutSec 30
        $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force

        $results = Invoke-TaskQueue -Tasks @($task) -Throttle 1 -Confirm:$false

        $results | Should -HaveCount 1
        $results[0].Status | Should -Be 'Succeeded'
        $results[0].OutputPreview | Should -Contain 'helper-ok'
    }

    It 'marks task as TimedOut when timeout is exceeded' {
        $task = New-UpdateTask -Name 'slow-task' -Category 'test' -Script { Start-Sleep -Seconds 60 } -TimeoutSec 2
        $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force

        $results = Invoke-TaskQueue -Tasks @($task) -Throttle 1 -Confirm:$false

        $results | Should -HaveCount 1
        $results[0].Status | Should -Be 'TimedOut'
    }

    It 'runs multiple tasks and respects resource exclusion' {
        $taskA = New-UpdateTask -Name 'task-a' -Category 'test' -Script { Write-Output 'a' } -Resources @('shared-res') -TimeoutSec 30
        $taskB = New-UpdateTask -Name 'task-b' -Category 'test' -Script { Write-Output 'b' } -Resources @('shared-res') -TimeoutSec 30
        $taskA | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
        $taskB | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force

        $results = Invoke-TaskQueue -Tasks @($taskA, $taskB) -Throttle 2 -Confirm:$false

        $results | Should -HaveCount 2
        ($results | Where-Object Status -eq 'Succeeded').Count | Should -Be 2
    }
}

# ── Show-WhatChanged ──────────────────────────────────────────────────────────

Describe 'Show-WhatChanged' {
    It 'does not throw when no previous summary exists' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'x' -Category 'c' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded'
        $summary = Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)
        { Show-WhatChanged -CurrentSummary $summary } | Should -Not -Throw
    }

    It 'detects a status change between runs' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'flaky' -Category 'test' -Script {}

        $r1 = New-TaskResult -Task $task -Status 'Succeeded'
        $prev = Save-RunSummary -Results @($r1) -Skipped @() -Planned @($task)
        $prevJson = $prev | ConvertTo-Json -Depth 8
        Set-Content -LiteralPath $script:PreviousJsonSummaryPath -Value $prevJson -Encoding utf8

        $r2 = New-TaskResult -Task $task -Status 'Failed'
        $curr = [pscustomobject]@{ Results = @($r2) }

        $output = & { Show-WhatChanged -CurrentSummary $curr } *>&1 | Out-String
        $output | Should -Match 'flaky'
    }
}

# ── Script integration ────────────────────────────────────────────────────────

Describe 'Script integration' {
    It 'runs SelfTest cleanly' {
        $result = & $script:InvokeTestPwsh -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:ScriptPath, '-SelfTest', '-NoElevate')
        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'All runnable tasks completed'
        $result.Output | Should -Not -Match 'WARNING:'
    }

    It 'exits 0 and reports no tasks when -Only filter matches nothing' {
        $result = & $script:InvokeTestPwsh -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:ScriptPath, '-DryRun', '-NoElevate', '-Only', 'definitely-not-a-real-update-task')
        $result.ExitCode | Should -Be 0
        $result.Output | Should -Match 'No runnable update tasks were found'
    }

    It 'parses with no syntax errors' {
        $errs = $null
        $tok = $null
        [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tok, [ref]$errs) | Out-Null
        $errs.Count | Should -Be 0
    }

    It '-ListTasks exits 0 and names expected tasks' {
        $stateDir = Join-Path $script:TestRoot 'listtasks-state'
        $result = & $script:InvokeTestPwsh -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:ScriptPath, '-ListTasks', '-NoElevate', '-ShowSkipped', '-StateDir', $stateDir)
        $output = $result.Output
        $result.ExitCode | Should -Be 0
        $output | Should -Match 'winget-source'
        $output | Should -Match 'uv-python'
        $output | Should -Match 'oh-my-posh'
        $output | Should -Match 'juliaup'
        $output | Should -Match 'yt-dlp'
        $output | Should -Match 'volta'
        $output | Should -Match 'fnm'
    }

    It '-SkipNode hides volta and fnm in task list' {
        $stateDir = Join-Path $script:TestRoot 'skipnode-state'
        $result = & $script:InvokeTestPwsh -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:ScriptPath, '-ListTasks', '-NoElevate', '-ShowSkipped', '-SkipNode', '-StateDir', $stateDir)
        $output = $result.Output
        $result.ExitCode | Should -Be 0
        ($output | Select-String 'volta' | Where-Object { $_ -match 'Skipped|skipped|disabled' }).Count | Should -BeGreaterThan 0
        ($output | Select-String 'fnm' | Where-Object { $_ -match 'Skipped|skipped|disabled' }).Count | Should -BeGreaterThan 0
    }

    It '-SkipUVTools hides uv, uv-tools, and uv-python in task list' {
        $stateDir = Join-Path $script:TestRoot 'skipuv-state'
        $result = & $script:InvokeTestPwsh -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:ScriptPath, '-ListTasks', '-NoElevate', '-ShowSkipped', '-SkipUVTools', '-StateDir', $stateDir)
        $output = $result.Output
        $result.ExitCode | Should -Be 0
        $uvSkips = @($output -split '\n' | Where-Object { $_ -match '\buv\b|\buv-tools\b|\buv-python\b' } | Where-Object { $_ -match 'Skipped|skipped|disabled' })
        $uvSkips.Count | Should -BeGreaterThan 0
    }
}
