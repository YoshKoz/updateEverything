# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Single-file PowerShell script (`updatescript.ps1`, ~2.6k lines) that updates Windows package managers, Windows components, and dev toolchains in one shot. Helpers: `fix_installer.ps1`, `force_reinstall.ps1`. `old.ps1` is a prior-version snapshot kept for reference — do not edit.

## Running

```powershell
./updatescript.ps1
./updatescript.ps1 -AutoElevate        # self-elevate via UAC
./updatescript.ps1 -FastMode           # skip slow/rarely-changing tools
./updatescript.ps1 -UltraFast          # FastMode + skip StoreApps/cleanup/WSL/WindowsUpdate/DefenderSignatures
./updatescript.ps1 -DryRun             # simulate; no state written
./updatescript.ps1 -WhatChanged        # diff package state vs last run
./updatescript.ps1 -Schedule -ScheduleTime "03:00"   # register scheduled task (admin)
./updatescript.ps1 -LogPath C:\Temp\update.log       # override log location
```

No build, no tests, no lint. Manual testing per `CONTRIBUTING.md`:
1. Run in controlled shell, confirm section detected.
2. Run non-admin → `-RequiresAdmin` sections skip gracefully.
3. Run `-FastMode` → `-SlowOperation` sections skip.
4. Run on system without the tool → section skips (not fail).

Use `$PSVersionTable.PSVersion` to report PS version when filing issues. PowerShell 7+ recommended; script handles 5.1 but prefers pwsh for elevation.

## Architecture

### Execution flow (top-to-bottom of `updatescript.ps1`)

1. `param(...)` block — all flags.
2. Encoding/globals/`$updateResults` buckets (`Success`, `Failed`, `Checked`, `Skipped`, `Details`).
3. `$script:Config` hashtable (line 88) — timeouts, skip lists, `FastModeSkip`/`UltraFastSkip` name lists, `WingetUpgradeHooks` (per-package Pre/Post scriptblocks).
4. `update-config.json` merged into `$script:Config` if present. **Note:** lines 215–217 force-clear `WingetSkipPackages`, `PipSkipPackages`, `SkipManagers` after merge — user config exclusions in those fields are intentionally ignored (comment line 214: "update everything"). Do not "fix" this without asking.
5. Elevation / `-Schedule` branch / param normalization.
6. Helper functions (~line 266–1143): command detection, winget parsing, MSI recovery, logging, state persistence.
7. `Invoke-Update` (line 1145) — core wrapper for all sequential sections.
8. `Invoke-UpdateBatch` (line 1215) — parallel variant; uses `Start-ThreadJob`, shadows `Write-Host` per-job to a tempfile via `$BatchInitScript` (line 1202), replays sequentially post-join.
9. Section calls in order (line 1414 onward):
   - **Parallel** lightweight managers (Scoop, Chocolatey, etc.) via `Invoke-UpdateBatch` at line 1420.
   - **Sequential** Winget at line 1520.
   - Gated block `if (-not $script:PackageManagersOnly)` at line 1798 → WindowsUpdate → StoreApps/DefenderSignatures batch → dev-tools batch.
   - Cleanup (line 2478), WhatChanged/Save-State (line 2528), summary/toast.

**`$script:PackageManagersOnly = $false`** (line 85) — default runs everything. Set to `$true` to short-circuit after Winget (skips Windows components + dev-tool section).

### Adding a new update section

Use the `Invoke-Update` pattern (also documented in `CONTRIBUTING.md`):

```powershell
Invoke-Update -Name 'ToolName' -Title 'Display' -RequiresCommand 'tool-bin' -Action {
    # Detect Scoop/winget managed installs → return early to avoid double-updating
    $src = (Get-Command tool-bin).Source
    if ($src -like '*scoop*') { Write-Host '  Managed by Scoop' -ForegroundColor Gray; return }
    $out = (tool-bin update 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
    # Set $script:stepChanged = $true when an actual update occurred
    # Set $script:stepMessage = '...' for summary detail
}
```

Flags: `-RequiresCommand` / `-RequiresAnyCommand` (skip if binary absent), `-RequiresAdmin` (skip non-elevated), `-SlowOperation` (skip in FastMode), `-Disabled:$SkipX` (user flag), `-NoSection` (suppress section header for nested calls).

For parallel-safe sections put them in an `Invoke-UpdateBatch -Tasks @(@{...})`. Only these helpers cross the job boundary (see `$BatchInitScript`, line 1202): `Write-FilteredOutput`, `Write-Detail`, `Write-Status`, `Write-Log`, `Invoke-StreamingCapture`, `Read-CapturedOutput`, `Invoke-WingetWithTimeout`, `Get-ToolInstallManager`, `Test-Command`, `ConvertTo-NormalizedPackageName`, `Complete-StepState`, `Get-VSCodeCliPath`. Anything else must be inlined into the task's `Action` or added to that list.

### Output helpers (always use these; don't call `Write-Host` directly in sections)

- `Write-Section $Title` — banner between sections.
- `Write-Status $Msg -Type Success|Warning|Error|Info` — top-level result.
- `Write-Detail $Msg -Type Info|Muted|Warning|Error` — indented child line.
- `Write-FilteredOutput $text` — strips ANSI/progress noise from external tool output.
- `Invoke-StreamingCapture { ... }` / `Read-CapturedOutput $path` — run a block, capture stdout+stderr to a temp file, return `{OutputPath, ExitCode}`. Preferred over `2>&1 | Out-String` for any long-running tool.

### Winget specifics

Winget is the most fragile section. The script carries extensive recovery logic:

- `Invoke-WingetWithTimeout` — wraps winget calls so a hang can't freeze the run.
- `Wait-WindowsInstallerIdle` — checks `Global\_MSIExecute` mutex; blocks before winget scan/install.
- Install-technology-mismatch (exe→msi or vice versa) detection + remediation via `Invoke-WingetTechnologyChangeReinstall` (uninstall old, install new).
- MSI "source not found" (1603 + log errors 1714/1612) detection → `Invoke-WingetMsiSourceRepair` / `Invoke-MsiFullRegistryCleanse`. Repeated failures cached in state file to avoid thrashing (`Get-/Save-MsiRepairFailureState`, `Set-/Clear-MsiRepairFailureCache`).
- Per-package pre/post hooks in `$script:Config.WingetUpgradeHooks` (exact key or prefix match, e.g. `JetBrains.`) — kill conflicting processes before upgrade, relaunch after. Extend this map to add new hookable packages.

When debugging winget: `force_reinstall.ps1` is the canonical escape hatch — edit the `$guids` hashtable with the broken package's uninstall GUID, run elevated; it scrubs registry entries and reinstalls via winget.

### State & logging

- State file: `$env:LOCALAPPDATA\Update-Everything\state.json` — winget/scoop/choco package maps + `LastRun` + MSI repair failure cache. Used by `-WhatChanged`.
- Log: `$env:LOCALAPPDATA\Update-Everything\updatescript.log` by default (falls back to script dir, then `%TEMP%`). Rotates at `LogMaxSizeMB` (10 MB).

### Regex/perf

Pre-compiled regexes live in `$script:RePatterns` (line 266). Reuse these; don't inline new `[regex]::new(...)` calls in hot paths.

## Repo conventions

- CHANGELOG.md is maintained manually — update it for user-visible behavior changes.
- `.VERSION` tag in the script synopsis (line 6) is the source of truth for script version.
- Don't commit `updatescript.log` or `run-elevated.log` changes; they're local artifacts. (`.gitignore` may not cover them — check.)