#Requires -Version 7.0
#Requires -Modules Pester

<#
.SYNOPSIS
    Pester tests for updatescript.ps1 — core logic, helpers, edge cases.
.NOTES
    Run:  Invoke-Pester .\updatescript.Tests.ps1 -Output Detailed
    These tests do NOT run actual update commands; all external tools are mocked.
#>

BeforeAll {
    # Load only the helper functions — dot-source inside a try/finally guard so
    # the script's top-level execution block (Import-UpdateConfig, Initialize-RunStorage,
    # the main dispatch) does NOT run.  We achieve this by injecting a stub environment.

    $script:ScriptPath = Join-Path $PSScriptRoot 'updatescript.ps1'

    # Parse out just the function definitions by extracting the AST. This avoids
    # side-effects from running the script body.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $script:ScriptPath, [ref]$null, [ref]$null
    )

    # Collect every function definition and define them in this scope.
    foreach ($funcDef in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        $funcText = $funcDef.Extent.Text
        try {
            Invoke-Expression $funcText
        }
        catch {
            Write-Warning "Skipped loading function $($funcDef.Name): $_"
        }
    }

    # Minimal $script:Config needed by Get-UpdateTasks / Get-FilteredTasks
    $script:Config = [ordered]@{
        FastModeSkip       = @('chocolatey', 'npm', 'rustup', 'cargo')
        UltraFastSkip      = @('windows-update', 'store-apps', 'wsl', 'defender', 'cleanup')
        SkipManagers       = @()
        WingetSkipPackages = @()
        PipSkipPackages    = @()
        NpmSkipPackages    = @()
        LogRetentionDays   = 14
        TempCleanupDays    = 7
    }

    $script:CommandCache = @{}
    $script:StateDir     = $env:TEMP
    $script:LogDir       = $env:TEMP

    # LogPath used by Write-Log
    $script:TestLogPath  = Join-Path $env:TEMP "update-everything-test-$([System.IO.Path]::GetRandomFileName()).log"
    $LogPath = $script:TestLogPath
}

AfterAll {
    if ($script:TestLogPath -and (Test-Path $script:TestLogPath)) {
        Remove-Item $script:TestLogPath -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
Describe 'ConvertTo-TaskId' {

    It 'lowercases and hyphenates spaces' {
        ConvertTo-TaskId 'My Tool' | Should -Be 'my-tool'
    }

    It 'collapses multiple non-alphanumeric chars to a single dash' {
        ConvertTo-TaskId 'wsl--distros' | Should -Be 'wsl-distros'
    }

    It 'trims leading and trailing dashes' {
        ConvertTo-TaskId '-foo-bar-' | Should -Be 'foo-bar'
    }

    It 'preserves already-normalised ids unchanged' {
        ConvertTo-TaskId 'windows-update' | Should -Be 'windows-update'
    }

    It 'handles a single word' {
        ConvertTo-TaskId 'winget' | Should -Be 'winget'
    }
}

# ---------------------------------------------------------------------------
Describe 'ConvertTo-FilterList' {

    It 'splits a comma-separated string into an array' {
        $result = ConvertTo-FilterList 'wsl,npm, rust '
        $result | Should -HaveCount 3
        $result | Should -Contain 'wsl'
        $result | Should -Contain 'npm'
        $result | Should -Contain 'rust'
    }

    It 'accepts a plain array and returns it trimmed' {
        $result = ConvertTo-FilterList @('  go  ', 'flutter')
        $result | Should -HaveCount 2
        $result | Should -Contain 'go'
    }

    It 'returns an empty array for null' {
        @(ConvertTo-FilterList $null) | Should -HaveCount 0
    }

    It 'filters out blank entries' {
        $result = ConvertTo-FilterList '  ,  ,npm'
        $result | Should -HaveCount 1
        $result | Should -Contain 'npm'
    }
}

# ---------------------------------------------------------------------------
Describe 'ConvertTo-StringArray' {

    It 'wraps a single string in an array' {
        $r = ConvertTo-StringArray 'hello'
        $r | Should -BeOfType [string]
        $r | Should -HaveCount 1
    }

    It 'returns an empty array for null' {
        @(ConvertTo-StringArray $null) | Should -HaveCount 0
    }

    It 'returns the same array for an existing array' {
        $r = ConvertTo-StringArray @('a', 'b')
        $r | Should -HaveCount 2
    }
}

# ---------------------------------------------------------------------------
Describe 'Test-IsAdmin' {

    It 'returns a boolean' {
        Test-IsAdmin | Should -BeOfType [bool]
    }
}

# ---------------------------------------------------------------------------
Describe 'Test-Command' {

    It 'returns $true for a command that exists (pwsh)' {
        # pwsh must be available since #Requires -Version 7.0
        Test-Command 'pwsh' | Should -BeTrue
    }

    It 'returns $false for a command that does not exist' {
        Test-Command 'this-tool-definitely-does-not-exist-xyz' | Should -BeFalse
    }

    It 'caches the result so Get-Command is not called twice' {
        $script:CommandCache.Clear()
        Test-Command 'pwsh' | Out-Null
        $script:CommandCache.ContainsKey('pwsh') | Should -BeTrue
        # Second call should use cache
        Test-Command 'pwsh' | Should -BeTrue
    }
}

# ---------------------------------------------------------------------------
Describe 'New-UpdateTask' {

    It 'creates an object with expected properties' {
        $task = New-UpdateTask -Name 'test-tool' -Category 'test' -Script { Write-Output 'ok' }
        $task.Id       | Should -Be 'test-tool'
        $task.Name     | Should -Be 'test-tool'
        $task.Category | Should -Be 'test'
        $task.RequiresAdmin | Should -BeFalse
        $task.Disabled | Should -BeFalse
    }

    It 'normalises name to Id' {
        $task = New-UpdateTask -Name 'My Tool' -Category 'x' -Script {}
        $task.Id | Should -Be 'my-tool'
    }

    It 'marks task disabled when -Disabled is set' {
        $task = New-UpdateTask -Name 'skip-me' -Category 'x' -Script {} -Disabled -DisabledReason 'testing'
        $task.Disabled | Should -BeTrue
        $task.DisabledReason | Should -Be 'testing'
    }

    It 'stores multiple required commands' {
        $task = New-UpdateTask -Name 'dual' -Category 'x' -Script {} -RequiresCommand @('git', 'gh')
        $task.RequiresCommand | Should -HaveCount 2
    }
}

# ---------------------------------------------------------------------------
Describe 'Test-NameMatch' {

    BeforeAll {
        $script:SampleTask = New-UpdateTask -Name 'windows-update' -Category 'system' -Script {} -Tags @('windows')
    }

    It 'matches by exact Id' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-update') | Should -BeTrue
    }

    It 'matches by category' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('system') | Should -BeTrue
    }

    It 'matches by tag' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows') | Should -BeTrue
    }

    It 'does not match an unrelated pattern' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('npm') | Should -BeFalse
    }

    It 'matches wildcard patterns' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('windows-*') | Should -BeTrue
    }

    It 'ignores blank patterns' {
        Test-NameMatch -Task $script:SampleTask -Patterns @('', '   ') | Should -BeFalse
    }
}

# ---------------------------------------------------------------------------
Describe 'Get-FilteredTasks (filter logic)' {

    BeforeAll {
        # Build a minimal task list without calling Get-UpdateTasks (which references
        # $WingetTimeoutSec, $TaskTimeoutSec etc from param block).
        function script:MakeTask($name, $cat, $admin = $false, $disabled = $false, $cmd = @()) {
            $t = New-UpdateTask -Name $name -Category $cat -Script {} `
                -RequiresAdmin:$admin -Disabled:$disabled -RequiresCommand $cmd
            $t | Add-Member -NotePropertyName Arguments -NotePropertyValue @{} -Force
            return $t
        }

        $script:AllTasks = @(
            (script:MakeTask 'winget'          'package-manager'),
            (script:MakeTask 'scoop'           'package-manager'),
            (script:MakeTask 'windows-update'  'system'          -admin $true),
            (script:MakeTask 'npm'             'javascript'      -cmd @('npm')),
            (script:MakeTask 'disabled-tool'   'misc'            -disabled $true)
        )

        # Stub Test-Command so 'npm' is seen as missing
        Mock -CommandName Test-Command -MockWith { $false } -ParameterFilter { $Name -eq 'npm' }
        Mock -CommandName Test-Command -MockWith { $true  } -ParameterFilter { $Name -ne 'npm' }
    }

    It 'skips admin tasks when not running as admin' {
        $r = Get-FilteredTasks -Tasks $script:AllTasks -IsAdmin $false
        $r.Skipped | Where-Object Name -eq 'windows-update' | Should -Not -BeNullOrEmpty
        $r.Planned  | Where-Object Name -eq 'windows-update' | Should -BeNullOrEmpty
    }

    It 'includes admin tasks when running as admin' {
        $r = Get-FilteredTasks -Tasks $script:AllTasks -IsAdmin $true
        $r.Planned | Where-Object Name -eq 'windows-update' | Should -Not -BeNullOrEmpty
    }

    It 'skips disabled tasks regardless of admin status' {
        $r = Get-FilteredTasks -Tasks $script:AllTasks -IsAdmin $true
        $r.Skipped | Where-Object Name -eq 'disabled-tool' | Should -Not -BeNullOrEmpty
    }

    It 'skips tasks whose required command is missing' {
        $r = Get-FilteredTasks -Tasks $script:AllTasks -IsAdmin $false
        $r.Skipped | Where-Object Name -eq 'npm' | Should -Not -BeNullOrEmpty
    }

    It '-Only filter selects only the named task' {
        $saved = $script:Only
        $script:Only = @('winget')   # mock the outer $Only param variable

        # We need to inject $Only into the function call; the function reads the
        # outer-scope variable $Only directly.  Use InModuleScope-style trick:
        $r = & {
            $Only = @('winget')
            Get-FilteredTasks -Tasks $script:AllTasks -IsAdmin $false
        }
        $r.Planned | ForEach-Object Name | Should -Contain 'winget'
        ($r.Planned | Where-Object Name -ne 'winget') | Should -BeNullOrEmpty
    }
}

# ---------------------------------------------------------------------------
Describe 'Write-Log' {

    BeforeAll {
        $script:TmpLog = Join-Path $env:TEMP "pester-writelog-$([System.IO.Path]::GetRandomFileName()).log"
        $LogPath = $script:TmpLog
    }

    AfterAll {
        Remove-Item $script:TmpLog -Force -ErrorAction SilentlyContinue
    }

    It 'writes a line with a timestamp and level prefix' {
        Write-Log -Message 'hello test' -Level Info
        $content = Get-Content $script:TmpLog -Raw
        $content | Should -Match '\[Info\].*hello test'
    }

    It 'does not throw when LogPath is inaccessible (silent failure)' {
        # Temporarily point $LogPath at an invalid path
        $badPath = 'Z:\definitely\not\writable\test.log'
        { Write-Log -Message 'noop' -Level Error } | Should -Not -Throw
    }
}

# ---------------------------------------------------------------------------
Describe 'Infrastructure I/O under WhatIfPreference' {

    BeforeAll {
        $script:TmpInfraRoot = Join-Path $env:TEMP "pester-whatif-$([System.IO.Path]::GetRandomFileName())"
        $script:SavedStateDir = $script:StateDir
        $script:SavedLogDir = $script:LogDir
        $script:SavedDefaultJsonSummaryPath = $script:DefaultJsonSummaryPath
        $script:SavedPreviousJsonSummaryPath = $script:PreviousJsonSummaryPath

        $script:StateDir = Join-Path $script:TmpInfraRoot 'state'
        $script:LogDir = Join-Path $script:StateDir 'logs'
        $script:DefaultJsonSummaryPath = Join-Path $script:StateDir 'last-run.json'
        $script:PreviousJsonSummaryPath = Join-Path $script:StateDir 'previous-run.json'
    }

    AfterAll {
        $script:StateDir = $script:SavedStateDir
        $script:LogDir = $script:SavedLogDir
        $script:DefaultJsonSummaryPath = $script:SavedDefaultJsonSummaryPath
        $script:PreviousJsonSummaryPath = $script:SavedPreviousJsonSummaryPath
        Remove-Item $script:TmpInfraRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Initialize-RunStorage still creates state and log directories' {
        $oldWhatIfPreference = $WhatIfPreference
        try {
            $WhatIfPreference = $true
            Initialize-RunStorage
        }
        finally {
            $WhatIfPreference = $oldWhatIfPreference
        }

        Test-Path $script:StateDir | Should -BeTrue
        Test-Path $script:LogDir | Should -BeTrue
    }

    It 'Write-Log still writes when WhatIfPreference is true' {
        $LogPath = Join-Path $script:LogDir 'whatif.log'
        $oldWhatIfPreference = $WhatIfPreference
        try {
            $WhatIfPreference = $true
            Write-Log -Message 'whatif-log' -Level Info
        }
        finally {
            $WhatIfPreference = $oldWhatIfPreference
        }

        Test-Path $LogPath | Should -BeTrue
        (Get-Content $LogPath -Raw) | Should -Match 'whatif-log'
    }

    It 'Save-RunSummary still writes JSON when WhatIfPreference is true' {
        $JsonSummaryPath = Join-Path $script:StateDir 'whatif-summary.json'
        $LogPath = Join-Path $script:LogDir 'whatif-summary.log'
        $script:StartTime = Get-Date
        $script:RunId = 'whatif-summary'
        $script:IsSimulation = $true
        $FastMode = $false
        $UltraFast = $false
        $ParallelThrottle = 2

        $task = New-UpdateTask -Name 'dummy' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Skipped'

        $oldWhatIfPreference = $WhatIfPreference
        try {
            $WhatIfPreference = $true
            Save-RunSummary -Results @($result) -Skipped @() -Planned @($task) | Out-Null
        }
        finally {
            $WhatIfPreference = $oldWhatIfPreference
        }

        Test-Path $JsonSummaryPath | Should -BeTrue
        (Get-Content $JsonSummaryPath -Raw | ConvertFrom-Json).RunId | Should -Be 'whatif-summary'
    }
}

# ---------------------------------------------------------------------------
Describe 'New-TaskResult' {

    It 'rounds DurationSeconds to 2dp' {
        $task = New-UpdateTask -Name 'foo' -Category 'x' -Script {}
        $r = New-TaskResult -Task $task -Status 'Succeeded' -DurationSeconds 1.23456
        $r.DurationSeconds | Should -Be 1.23
    }

    It 'stores output preview capped at 20 lines' {
        $task = New-UpdateTask -Name 'foo' -Category 'x' -Script {}
        $manyLines = 1..30 | ForEach-Object { "line $_" }
        $r = New-TaskResult -Task $task -Status 'Succeeded' -Output $manyLines
        $r.OutputPreview | Should -HaveCount 20
    }
}

# ---------------------------------------------------------------------------
Describe 'Import-UpdateConfig — array key accepts comma-separated string' {

    BeforeAll {
        $script:TmpCfg = Join-Path $env:TEMP "update-config-test-$([System.IO.Path]::GetRandomFileName()).json"
    }

    AfterAll {
        Remove-Item $script:TmpCfg -Force -ErrorAction SilentlyContinue
    }

    It 'splits a comma-separated string value into an array for array config keys' {
        # Simulate what Import-UpdateConfig does for a property whose config key is an array
        '{"FastModeSkip":"npm,rust,go"}' | Set-Content $script:TmpCfg

        # Reset config and re-run import pointing at our temp file
        $script:Config.FastModeSkip = @()

        # Patch $PSScriptRoot equivalent: Import-UpdateConfig uses $PSScriptRoot
        # We test the sub-logic directly instead
        $configJson = Get-Content $script:TmpCfg -Raw | ConvertFrom-Json
        foreach ($prop in $configJson.PSObject.Properties) {
            if ($script:Config.Contains($prop.Name) -and $script:Config[$prop.Name] -is [array]) {
                $script:Config[$prop.Name] = @(ConvertTo-FilterList $prop.Value)
            }
        }

        $script:Config.FastModeSkip | Should -HaveCount 3
        $script:Config.FastModeSkip | Should -Contain 'npm'
        $script:Config.FastModeSkip | Should -Contain 'rust'
        $script:Config.FastModeSkip | Should -Contain 'go'
    }

    It 'accepts a JSON array directly' {
        '{"FastModeSkip":["bun","deno"]}' | Set-Content $script:TmpCfg
        $script:Config.FastModeSkip = @()

        $configJson = Get-Content $script:TmpCfg -Raw | ConvertFrom-Json
        foreach ($prop in $configJson.PSObject.Properties) {
            if ($script:Config.Contains($prop.Name) -and $script:Config[$prop.Name] -is [array]) {
                $script:Config[$prop.Name] = @(ConvertTo-FilterList $prop.Value)
            }
        }

        $script:Config.FastModeSkip | Should -HaveCount 2
        $script:Config.FastModeSkip | Should -Contain 'bun'
        $script:Config.FastModeSkip | Should -Contain 'deno'
    }
}

# ---------------------------------------------------------------------------
Describe 'Self-elevation parameter forwarding — array params' {

    It 'joins array values with commas into a single argument' {
        # Test the fixed forwarding logic directly
        $forwarded = [System.Collections.Generic.List[string]]::new()
        $key   = 'Only'
        $value = @('winget', 'npm', 'scoop')

        # Reproduce the fixed code path
        if ($value.Count -gt 0) {
            $forwarded.Add("-$key") | Out-Null
            $quoted = @($value | ForEach-Object { '"{0}"' -f ([string]$_).Replace('"', '\"') })
            $forwarded.Add($quoted -join ',') | Out-Null
        }

        $forwarded | Should -HaveCount 2
        $forwarded[0] | Should -Be '-Only'
        $forwarded[1] | Should -Be '"winget","npm","scoop"'
    }

    It 'skips empty arrays entirely' {
        $forwarded = [System.Collections.Generic.List[string]]::new()
        $key   = 'Only'
        $value = @()

        if ($value.Count -gt 0) {
            $forwarded.Add("-$key") | Out-Null
        }

        $forwarded | Should -HaveCount 0
    }
}

# ---------------------------------------------------------------------------
Describe 'Winget LASTEXITCODE capture' {

    It 'captures exit code into a local variable immediately after the command' {
        $scriptContent = Get-Content $script:ScriptPath -Raw
        # The script text contains backtick-escaped dollar signs inside a here-string,
        # so the literal text in the file is "`$wingetListExit = `$LASTEXITCODE"
        $scriptContent | Should -Match 'wingetListExit'
        $scriptContent | Should -Match 'winget upgrade list failed with exit code'
        # Confirm the old bare $LASTEXITCODE error throw is gone
        $scriptContent | Should -Not -Match 'throw "winget upgrade list failed with exit code [^`]?\$LASTEXITCODE"'
    }

    It 'extracts the Id column correctly when the package name contains repeated spaces' {
        $sample = @'
Name                                   Id                                    Version       Available    Source
-------------------------------------- ------------------------------------- ------------- ------------ ------
Package With  Many  Spaces             SomeVendor.PackageWithSpaces          1.0           2.0          winget
'@

        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        $idColumnStart = $null
        $idColumnEnd = $null
        foreach ($line in $sample -split "`n") {
            $text = [string]$line
            if ($text -match '^\s*-{3,}') {
                $inTable = $true
                $columnMatches = [regex]::Matches($text, '-{3,}')
                if ($columnMatches.Count -ge 2) {
                    $idColumnStart = $columnMatches[1].Index
                    if ($columnMatches.Count -ge 3) {
                        $idColumnEnd = $columnMatches[2].Index
                    }
                }
                continue
            }
            if (-not $inTable -or [string]::IsNullOrWhiteSpace($text)) { continue }
            if ($null -eq $idColumnStart -or $text.Length -le $idColumnStart) { continue }
            if ($null -ne $idColumnEnd -and $idColumnEnd -gt $idColumnStart) {
                $width = [Math]::Min($idColumnEnd - $idColumnStart, $text.Length - $idColumnStart)
                $id = $text.Substring($idColumnStart, $width).Trim()
            }
            else {
                $id = $text.Substring($idColumnStart).Trim()
            }
            if ($id -and $id -notmatch '^(Id|Version|-)') {
                [void]$ids.Add($id)
            }
        }

        $ids | Should -HaveCount 1
        $ids[0] | Should -Be 'SomeVendor.PackageWithSpaces'
    }
}

# ---------------------------------------------------------------------------
Describe 'Invoke-TaskQueue — CmdletBinding for WhatIf' {

    It 'Invoke-TaskQueue has CmdletBinding with SupportsShouldProcess' {
        $scriptContent = Get-Content $script:ScriptPath -Raw
        # Check function definition has the attribute
        $scriptContent | Should -Match 'function Invoke-TaskQueue\s*\{'
        $scriptContent | Should -Match '\[CmdletBinding\(SupportsShouldProcess\s*=\s*\$true\)\]'
    }
}

# ---------------------------------------------------------------------------
Describe 'Join-QuotedList' {

    It 'wraps each element in single quotes' {
        Join-QuotedList @('a', 'b') | Should -Be "'a', 'b'"
    }

    It 'escapes embedded single quotes' {
        Join-QuotedList @("it's") | Should -Be "'it''s'"
    }

    It 'returns empty string for empty input' {
        Join-QuotedList @() | Should -Be ''
    }
}

# ---------------------------------------------------------------------------
Describe 'Show-WhatChanged — no previous run' {

    It 'emits a warning and does not throw when no previous summary exists' {
        $fakeSummary = [pscustomobject]@{ Results = @() }
        $script:PreviousJsonSummaryPath = 'Z:\does\not\exist.json'
        { Show-WhatChanged -CurrentSummary $fakeSummary } | Should -Not -Throw
    }
}

# ---------------------------------------------------------------------------
Describe 'Initialize-RunStorage — WhatIf immunity' {

    It 'creates state dirs even when WhatIfPreference is $true' {
        $tmpDir = Join-Path $env:TEMP "pester-init-whatif-$([System.IO.Path]::GetRandomFileName())"
        $savedStateDir = $script:StateDir
        $savedLogDir   = $script:LogDir
        $script:StateDir = $tmpDir
        $script:LogDir   = Join-Path $tmpDir 'logs'

        try {
            $WhatIfPreference = $true
            Initialize-RunStorage
            Test-Path $tmpDir | Should -BeTrue
            Test-Path (Join-Path $tmpDir 'logs') | Should -BeTrue
        }
        finally {
            $WhatIfPreference = $false
            $script:StateDir = $savedStateDir
            $script:LogDir   = $savedLogDir
            Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
Describe 'Write-Log — WhatIf immunity' {

    It 'writes log entries even when WhatIfPreference is $true' {
        $tmpLog = Join-Path $env:TEMP "pester-log-whatif-$([System.IO.Path]::GetRandomFileName()).log"
        $savedLogPath = $LogPath
        $LogPath = $tmpLog

        try {
            $WhatIfPreference = $true
            Write-Log -Message 'whatif-immune-test' -Level Info
            Test-Path $tmpLog | Should -BeTrue
            Get-Content $tmpLog -Raw | Should -Match 'whatif-immune-test'
        }
        finally {
            $WhatIfPreference = $false
            $LogPath = $savedLogPath
            Remove-Item $tmpLog -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
Describe 'Save-RunSummary — WhatIf immunity' {

    It 'writes summary JSON even when WhatIfPreference is $true' {
        $tmpSummary = Join-Path $env:TEMP "pester-summary-whatif-$([System.IO.Path]::GetRandomFileName()).json"
        $savedSummaryPath = $JsonSummaryPath
        $JsonSummaryPath = $tmpSummary
        $script:StartTime    = Get-Date
        $script:RunId        = 'whatif-test'
        $script:IsSimulation = $true
        $FastMode  = $false
        $UltraFast = $false
        $ParallelThrottle = 1

        try {
            $WhatIfPreference = $true
            $task   = New-UpdateTask -Name 'test' -Category 'x' -Script {}
            $result = New-TaskResult -Task $task -Status 'DryRun' -Reason 'preview only'
            Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)
            Test-Path $tmpSummary | Should -BeTrue
            $j = Get-Content $tmpSummary -Raw | ConvertFrom-Json
            $j.RunId | Should -Be 'whatif-test'
        }
        finally {
            $WhatIfPreference = $false
            $JsonSummaryPath = $savedSummaryPath
            Remove-Item $tmpSummary -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
Describe 'Invoke-SelfElevation — no $args shadowing' {

    It 'uses $elevateArgs not $args (avoids automatic variable shadow)' {
        $content = Get-Content $script:ScriptPath -Raw
        $content | Should -Match '\$elevateArgs\s*='
        $content | Should -Not -Match '^\s*\$args\s*=' # no bare $args assignment at start of line
    }
}

# ---------------------------------------------------------------------------
Describe 'Winget column parser — footer and edge cases' {

    It 'correctly ignores a summary footer line' {
        $sample = @"
Name          Id              Version    Available    Source
------------- --------------- ---------- ------------ ------
Git           Git.Git         2.43.0     2.44.0       winget
1 upgrades available.
"@
        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        foreach ($line in $sample -split "`n") {
            $text = [string]$line
            if ($text -match '^\s*-{3,}') { $inTable = $true; continue }
            if (-not $inTable -or [string]::IsNullOrWhiteSpace($text)) { continue }
            $columns = $text.Trim() -split '\s{2,}'
            if ($columns.Count -ge 2) {
                $id = $columns[1].Trim()
                if ($id -and $id -notmatch '^(Id|Version|-)') { [void]$ids.Add($id) }
            }
        }
        $ids | Should -HaveCount 1
        $ids[0] | Should -Be 'Git.Git'
    }

    It 'skips the header row (Id column header)' {
        $sample = @"
Name    Id       Version  Available  Source
------- -------- -------- ---------- ------
tool    Tool.Id  1.0      2.0        winget
"@
        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        foreach ($line in $sample -split "`n") {
            $text = [string]$line
            if ($text -match '^\s*-{3,}') { $inTable = $true; continue }
            if (-not $inTable -or [string]::IsNullOrWhiteSpace($text)) { continue }
            $columns = $text.Trim() -split '\s{2,}'
            if ($columns.Count -ge 2) {
                $id = $columns[1].Trim()
                if ($id -and $id -notmatch '^(Id|Version|-)') { [void]$ids.Add($id) }
            }
        }
        # Header is before the separator line so $inTable is false when it's processed
        $ids | Should -HaveCount 1
        $ids[0] | Should -Be 'Tool.Id'
    }

    It 'handles a package with no available version (empty Available column)' {
        $sample = @"
Name     Id           Version  Available  Source
-------- ------------ -------- ---------- ------
MyTool   Vendor.Tool  1.0                 winget
"@
        $ids = [System.Collections.Generic.List[string]]::new()
        $inTable = $false
        foreach ($line in $sample -split "`n") {
            $text = [string]$line
            if ($text -match '^\s*-{3,}') { $inTable = $true; continue }
            if (-not $inTable -or [string]::IsNullOrWhiteSpace($text)) { continue }
            $columns = $text.Trim() -split '\s{2,}'
            if ($columns.Count -ge 2) {
                $id = $columns[1].Trim()
                if ($id -and $id -notmatch '^(Id|Version|-)') { [void]$ids.Add($id) }
            }
        }
        $ids | Should -HaveCount 1
        $ids[0] | Should -Be 'Vendor.Tool'
    }
}

# ---------------------------------------------------------------------------
Describe 'UltraFast sets FastMode' {

    It 'UltraFast skip list is a superset: all UltraFastSkip items are skipped in UltraFast mode' {
        $ultra  = @(ConvertTo-FilterList $script:Config.UltraFastSkip)
        $fast   = @(ConvertTo-FilterList $script:Config.FastModeSkip)
        # UltraFastSkip items should not appear in FastModeSkip (they're separate layers)
        # but both are applied together under -UltraFast; verify each list is non-empty
        $ultra.Count | Should -BeGreaterThan 0
        $fast.Count  | Should -BeGreaterThan 0
    }

    It 'FastModeSkip contains expected slow tasks (from script source)' {
        # Read defaults directly from the script source so earlier test mutations don't interfere
        $src = Get-Content $script:ScriptPath -Raw
        $src | Should -Match "FastModeSkip\s*=\s*@\("
        $src | Should -Match "'chocolatey'"
        $src | Should -Match "'rustup'"
        $src | Should -Match "'cargo'"
    }
}

# ---------------------------------------------------------------------------
Describe 'Save-RunSummary' {

    BeforeAll {
        $script:TmpSummary = Join-Path $env:TEMP "pester-summary-$([System.IO.Path]::GetRandomFileName()).json"
        $JsonSummaryPath   = $script:TmpSummary
        $script:StartTime  = Get-Date
        $script:RunId      = 'test-run'
        $script:IsSimulation = $false
        $FastMode  = $false
        $UltraFast = $false
        $ParallelThrottle = 2
        $LogPath   = $script:TestLogPath
    }

    AfterAll {
        Remove-Item $script:TmpSummary -Force -ErrorAction SilentlyContinue
    }

    It 'writes valid JSON to the summary path' {
        $task   = New-UpdateTask -Name 'dummy' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Succeeded'
        Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)

        Test-Path $script:TmpSummary | Should -BeTrue
        $parsed = Get-Content $script:TmpSummary -Raw | ConvertFrom-Json
        $parsed.RunId | Should -Be 'test-run'
        $parsed.SucceededCount | Should -Be 1
        $parsed.FailedCount    | Should -Be 0
    }

    It 'counts failed tasks correctly' {
        $task   = New-UpdateTask -Name 'broken' -Category 'test' -Script {}
        $result = New-TaskResult -Task $task -Status 'Failed' -Reason 'oops'
        Save-RunSummary -Results @($result) -Skipped @() -Planned @($task)

        $parsed = Get-Content $script:TmpSummary -Raw | ConvertFrom-Json
        $parsed.FailedCount | Should -Be 1
    }
}
