#requires -Version 7
# Runs the four verification gates and writes each gate's raw output into -OutDir.
# Exit 0 only when all four are green.
param(
    [Parameter(Mandatory)][string]$OutDir
)

$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$zig = Join-Path $repo 'rewrites\zig'
$mainZig = Join-Path $zig 'main.zig'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# Multiline string constants in main.zig are the shipped scripts. Pull each one
# back out of the `\\`-prefixed literal so it can be handed to a real parser.
function Get-EmbeddedBlocks {
    $src = Get-Content $mainZig
    $blocks = @()
    for ($i = 0; $i -lt $src.Count; $i++) {
        if ($src[$i] -notmatch '^const\s+([A-Za-z0-9_]+)\s*=\s*$') { continue }
        $name = $matches[1]
        $j = $i + 1
        $lines = @()
        while ($j -lt $src.Count -and $src[$j] -match '^\s*\\\\') {
            $lines += ($src[$j] -replace '^\s*\\\\', '')
            $j++
        }
        if ($lines.Count -gt 2) { $blocks += [pscustomobject]@{ Name = $name; Lines = $lines } }
        $i = $j
    }
    return $blocks
}

$blocks = Get-EmbeddedBlocks
$byName = @{}
foreach ($b in $blocks) { $byName[$b.Name] = $b }
$results = [ordered]@{}

# G1 — zig unit tests
Push-Location $zig
$g1 = & zig build test --summary all 2>&1 | Out-String
$g1Code = $LASTEXITCODE
Pop-Location
$g1 | Set-Content (Join-Path $OutDir 'G1-zig-test.txt')
$results['G1 zig build test'] = ($g1Code -eq 0)

# G2 — every embedded PowerShell body must parse.
# help_text is CLI usage text, not a script; it only looks like one because of the flags.
$notScripts = @('help_text')
$psReport = @()
$psFail = 0
# Some bodies start mid-expression because the generator splices a list in front of
# them (see wingetPerPackageScript / storeAppsScript in main.zig). Reproduce that
# opening fragment so the parser sees the same text the task actually runs.
$psPrefix = @{
    winget_per_package_body = "`$deferFailures = `$True`n`$skip = @('x'"
    store_apps_body         = "`$skip = @('x'"
}
foreach ($b in $blocks) {
    if ($b.Name -in $notScripts) { continue }
    if ($b.Name -like '*_post') { continue }
    $text = $b.Lines -join "`n"
    if ($b.Name -like '*_pre') {
        $post = $byName[($b.Name -replace '_pre$', '_post')]
        if (-not $post) { continue }
        $text = $text + "`n`$spliced = 1`n" + ($post.Lines -join "`n")
    }
    if ($text -notmatch '(Write-Host|Write-Output|Get-\w+|Set-\w+|\$LASTEXITCODE|Invoke-RestMethod)') { continue }
    if ($text -match '(?m)^\s*import\s') { continue }
    if ($psPrefix.ContainsKey($b.Name)) { $text = $psPrefix[$b.Name] + "`n" + $text }
    # Remaining bodies expect the generator's variable declarations up front.
    $decls = @('repo', 'dir', 'assetRe', 'verRe', 'tagName', 'scheme', 'tagRe', 'preCmd', 'postCmd',
        'verCmd', 'skip', 'ignored', 'managed', 'days', 'deep', 'features', 'projectId', 'ref', 'job') |
        ForEach-Object { "`$$_ = ''" }
    $probe = ($decls + $text) -join "`n"
    $errs = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($probe, [ref]$null, [ref]$errs)
    $hard = @($errs | Where-Object { $_.ErrorId -ne 'MissingEndCurlyBrace' })
    if ($hard.Count) {
        $psFail++
        $psReport += "FAIL $($b.Name): $($hard.Count) parse error(s)"
        $psReport += ($hard | ForEach-Object { "  line $($_.Extent.StartLineNumber): $($_.Message)" })
    }
    else {
        $psReport += "OK   $($b.Name) ($($b.Lines.Count) lines)"
    }
}
$psReport -join "`n" | Set-Content (Join-Path $OutDir 'G2-powershell-parse.txt')
$results['G2 PowerShell parse'] = ($psFail -eq 0)

# G3 — every embedded python body must compile.
# `_pre` / `_post` pairs are fragments: the generator splices a literal set between
# them, so only the reassembled pair is valid python.
$pyReport = @()
$pyFail = 0
foreach ($b in $blocks) {
    if ($b.Name -like '*_post') { continue }
    $text = $b.Lines -join "`n"
    if ($b.Name -like '*_pre') {
        $post = $byName[($b.Name -replace '_pre$', '_post')]
        if (-not $post) { continue }
        $text = $text + "`n'placeholder',`n" + ($post.Lines -join "`n")
    }
    if ($text -notmatch '(?m)^\s*(import|from)\s+\w') { continue }
    if ($text -match 'Write-Host|Invoke-RestMethod') { continue }
    $tmp = Join-Path $env:TEMP ("ue_gate_" + $b.Name + ".py")
    Set-Content -Path $tmp -Value $text -Encoding UTF8
    $out = & python -m py_compile $tmp 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { $pyFail++; $pyReport += "FAIL $($b.Name): $out" }
    else { $pyReport += "OK   $($b.Name) ($($b.Lines.Count) lines)" }
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}
$pyReport -join "`n" | Set-Content (Join-Path $OutDir 'G3-python-compile.txt')
$results['G3 python py_compile'] = ($pyFail -eq 0)

# G4 — the binary must still build every task
Push-Location $zig
$build = & zig build 2>&1 | Out-String
$buildCode = $LASTEXITCODE
Pop-Location
Push-Location $repo
$dry = & (Join-Path $zig 'zig-out\bin\updateeverything-zig.exe') --dry-run --config (Join-Path $repo 'update-config.json') 2>&1 | Out-String
Pop-Location
($build + $dry) | Set-Content (Join-Path $OutDir 'G4-dry-run.txt')
# Rows are colour-coded, so the status word is not at a fixed offset.
$dryPlain = $dry -replace "`e\[[0-9;]*m", ''
$dryCount = ([regex]::Matches($dryPlain, '(?m)^\s*\S*\s*(dry|skip)\s+\S')).Count
$results['G4 dry-run'] = ($buildCode -eq 0 -and $dryCount -ge 80)

$summary = $results.GetEnumerator() | ForEach-Object {
    "{0,-24} {1}" -f $_.Key, $(if ($_.Value) { 'PASS' } else { 'FAIL' })
}
$summary += "tasks listed by dry-run: $dryCount"
$summary -join "`n" | Tee-Object -FilePath (Join-Path $OutDir 'gates-summary.txt')
if ($results.Values -contains $false) { exit 1 } else { exit 0 }
