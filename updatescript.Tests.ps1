#Requires -Version 7.0
#Requires -Modules Pester

BeforeAll {
    $script:ScriptPath = Join-Path $PSScriptRoot 'updatescript.ps1'
    $errors = $null
    $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) {
        throw "Could not parse updatescript.ps1: $($errors[0].Message)"
    }

    foreach ($funcDef in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        Invoke-Expression $funcDef.Extent.Text
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
        FastModeSkip       = @('npm', 'cargo')
        UltraFastSkip      = @('windows-update', 'cleanup')
        SkipManagers       = @()
        WingetSkipPackages = @('Microsoft.VisualStudio.BuildTools')
        PipSkipPackages    = @()
        NpmSkipPackages    = @()
        LogRetentionDays   = 14
        TempCleanupDays    = 7
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
}

AfterAll {
    Remove-Item -LiteralPath $script:TestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-TaskId' {
    It 'normalizes names into task ids' {
        ConvertTo-TaskId 'My Tool / Name' | Should -Be 'my-tool-name'
    }
}

Describe 'New-UpdateTask' {
    It 'stores normalized ids, tags, resources, and timeout' {
        $task = New-UpdateTask -Name 'VS Code Extensions' -Category 'dev-tools' -Script {} -Tags @('VSCode') -Resources @('VS Code') -TimeoutSec 42
        $task.Id | Should -Be 'vs-code-extensions'
        $task.Tags | Should -Be @('vscode')
        $task.Resources | Should -Be @('vs-code')
        $task.TimeoutSec | Should -Be 42
    }
}

Describe 'Test-NameMatch' {
    BeforeAll {
        $script:SampleTask = New-UpdateTask -Name 'Windows Update' -Category 'system tools' -Script {} -Tags @('windows')
    }

    It 'matches by id, category, and wildcard patterns' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-update') | Should -BeTrue
        Test-NameMatch -Task $script:SampleTask -Patterns @('system-tools') | Should -BeTrue
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-*') | Should -BeTrue
    }
}

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

        $result = & {
            $Only = $script:Only
            $Skip = $script:Skip
            $FastMode = $script:FastMode
            $UltraFast = $script:UltraFast
            Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $false
        }

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Contain 'winget'
        ($result.Skipped | Where-Object Name -eq 'windows-update').Reason | Should -Be 'requires Administrator'
        ($result.Skipped | Where-Object Name -eq 'disabled-tool').Reason | Should -Be 'disabled for test'
    }

    It 'applies Only, Skip, and FastMode filters' {
        Mock Test-Command { $true }
        $script:Only = @('npm', 'cleanup')
        $script:Skip = @('cleanup')
        $script:FastMode = $true

        $result = & {
            $Only = $script:Only
            $Skip = $script:Skip
            $FastMode = $script:FastMode
            $UltraFast = $script:UltraFast
            Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true
        }

        ($result.Planned | Select-Object -ExpandProperty Name) | Should -Be @()
        ($result.Skipped | Where-Object Name -eq 'npm').Reason | Should -Be 'skipped by filter'
        ($result.Skipped | Where-Object Name -eq 'cleanup').Reason | Should -Be 'skipped by filter'
    }

    It 'skips tasks whose required command is missing' {
        Mock Test-Command {
            param($Name)
            return $Name -ne 'winget'
        }

        $result = & {
            $Only = $script:Only
            $Skip = $script:Skip
            $FastMode = $script:FastMode
            $UltraFast = $script:UltraFast
            Get-FilteredTasks -Tasks $script:TasksUnderTest -IsAdmin $true
        }

        ($result.Skipped | Where-Object Name -eq 'winget').Reason | Should -Be 'missing command: winget'
    }
}

Describe 'Initialize-RunStorage' {
    It 'creates writable state and log directories' {
        Initialize-RunStorage
        Test-Path -LiteralPath $script:StateDir | Should -BeTrue
        Test-Path -LiteralPath $script:LogDir | Should -BeTrue
    }
}

Describe 'Write-Log' {
    It 'writes log lines to the configured log file' {
        Initialize-RunStorage
        Write-Log -Message 'test message' -Level Info
        $content = Get-Content -LiteralPath $script:LogPath -Raw
        $content | Should -Match 'test message'
        $content | Should -Match '\[INFO\]'
    }
}

Describe 'Save-RunSummary' {
    It 'writes a summary containing version and success counts' {
        Initialize-RunStorage
        $task = New-UpdateTask -Name 'dummy' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded' -DurationSeconds 1.234

        $summary = Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)

        $summary.Version | Should -Be 'test-version'
        $summary.SummaryWritten | Should -BeTrue

        $json = Get-Content -LiteralPath $script:JsonSummaryPath -Raw | ConvertFrom-Json
        $json.RunId | Should -Be 'test-run'
        $json.SucceededCount | Should -Be 1
    }
}

Describe 'Invoke-TaskQueue' {
    It 'runs a simple task through the queue successfully' {
        $task = New-UpdateTask -Name 'self-test' -Category 'diagnostics' -Script { Write-Output 'ok' } -TimeoutSec 30
        $task | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force

        $results = Invoke-TaskQueue -Tasks @($task) -Throttle 1 -Confirm:$false

        $results | Should -HaveCount 1
        $results[0].Status | Should -Be 'Succeeded'
        $results[0].OutputPreview | Should -Contain 'ok'
    }
}

Describe 'Script integration' {
    It 'runs SelfTest cleanly' {
        $output = & pwsh -NoProfile -ExecutionPolicy Bypass -File $script:ScriptPath -SelfTest -NoPause -NoElevate 2>&1
        $LASTEXITCODE | Should -Be 0
        ($output | Out-String) | Should -Match 'All runnable tasks completed'
        ($output | Out-String) | Should -Not -Match 'WARNING:'
    }

    It 'treats empty filtered runs as informational' {
        $output = & pwsh -NoProfile -ExecutionPolicy Bypass -File $script:ScriptPath -DryRun -NoPause -NoElevate -Only winget 2>&1
        $LASTEXITCODE | Should -Be 0
        ($output | Out-String) | Should -Match 'No runnable update tasks were found'
        ($output | Out-String) | Should -Not -Match 'WARNING:'
    }
}
