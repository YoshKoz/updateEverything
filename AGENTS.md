# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## What this is

`updateEverything` updates every package manager and dev toolchain on a Windows machine in one run.
The shipping product is **one PowerShell script** (`updatescript.ps1`, ~4070 lines). Two experimental
rewrites under `rewrites/` (`rust/`, `go/`) re-implement a subset each; neither is a replacement.

## Commands

```powershell
# Run (admin recommended for full coverage)
.\updatescript.ps1
.\updatescript.ps1 -DryRun            # print every action, change nothing
.\updatescript.ps1 -FastMode          # skip slow managers
.\updatescript.ps1 -ListTasks         # show planned/skipped tasks, no execution
.\updatescript.ps1 -Only winget,pip   # run only matching task names/categories/tags
.\updatescript.ps1 -Skip rust,go
.\updatescript.ps1 -SelfTest          # run one trivial task through the scheduler

# Tests (Pester v5)
Invoke-Pester .\updatescript.Tests.ps1
Invoke-Pester .\updatescript.Tests.ps1 -Output Detailed
# single block:
Invoke-Pester .\updatescript.Tests.ps1 -FullNameFilter '*ConvertFrom-WingetUpgradeOutput*'

# Lint
Invoke-ScriptAnalyzer -Path .\updatescript.ps1 -Settings .\PSScriptAnalyzerSettings.psd1

# Rust runner
cd rewrites\rust
cargo run -- --list-tasks
cargo run -- --dry-run --only winget-source,juliaup
cargo build --release

# Go runner (note: single-dash flags, different names than the PS1 switches)
cd rewrites\go
go build -o update-everything.exe .
go run . -list-tasks
go run . -dry-run -only juliaup
go test ./...
go test -run TestParseWingetIDs ./...
```

## Architecture — PowerShell script

The script is a **task scheduler**, not a linear list of update commands. Understanding the task model
is the key to working in it.

- **Task definition** — `Get-UpdateTasks` (huge function, ~line 946–2890) builds an array of task
  objects via `New-UpdateTask`. Each task carries: `Name`, `Category`, `Tags`, `RequiresCommand`,
  `Resources`, `TimeoutSec`, `Disabled`/`DisabledReason`, and a `Script` scriptblock. To add or change
  an updater, edit this function — do not add inline update logic elsewhere.
- **Filtering** — `Get-FilteredTasks` applies `-Only`/`-Skip` (matched by name, category, or tag via
  `Test-NameMatch`/`ConvertTo-TaskId`), speed modes, `-Skip*` switches, and missing-command checks.
- **Execution** — `Invoke-TaskQueue` runs tasks concurrently with `Start-ThreadJob` (each task =
  `Start-UpdateTaskJob`), bounded by `-ParallelThrottle` (0 = auto from CPU count). `Resources` are
  mutual-exclusion locks: two tasks sharing a resource (e.g. `winget`, `pip`) never run at once —
  `Test-ResourcesAvailable` gates this. Use a shared `Resources` tag when tasks touch the same manager
  or lock.
- **Results** — `New-TaskResult` normalizes outcomes; `Save-RunSummary` writes JSON; `Show-UpdateSummary`
  / `Show-WhatChanged` / `Show-RunNotes` render the end-of-run report. `Get-RunNotes` extracts
  actionable lines from task output.
- **Lifecycle** — `Initialize-RunStorage` (logs/summaries dir), `Enter-ProcessLock`/`Remove-ProcessLock`
  (single-instance), `Invoke-SelfElevation` (`-AutoElevate`), `Register-UpdateSchedule` (`-Schedule`).

### Config

- `$script:Config` (ordered hashtable, ~line 189) holds all defaults. `Import-UpdateConfig` overlays
  `update-config.json` (next to the script); any key present overrides the default.
- Array keys are normalized through `ConvertTo-FilterList`/`ConvertTo-StringArray`.
- Notable config: per-manager `*SkipPackages` / `*ProtectedPackages`, `PipIgnoreHealthPackages`,
  `GithubTools` (declarative GitHub/GitLab release & artifact installers — see `update-config.json`),
  `CrossManagerFallback`, and `BeforeHooks`/`AfterHooks`.

## Architecture — Rust runner (`rewrites/rust/src/main.rs`)

Single-file (~3550 lines). `build_tasks` mirrors `Get-UpdateTasks`; `filter_tasks` mirrors
`Get-FilteredTasks`; `run_tasks`/`run_task_streaming` mirror `Invoke-TaskQueue`. Most updaters are
emitted as PowerShell scriptblocks (`pwsh_cmd` + the `*_args`/`*_script` helpers) — the Rust side
orchestrates, PowerShell still executes. Reads the same `update-config.json`. Intentionally does **not**
cover Windows Update, Store apps, scheduled-task registration, or admin-only cleanup yet.

## Architecture — Go runner (`rewrites/go/`)

Multi-file, unlike the Rust runner. `main.go` parses flags and wires everything; `runner.go` is the
scheduler; task definitions are split by platform — `tasks_windows.go` (winget/scoop/choco/WSL/Defender),
`tasks_devtools.go` (language toolchains), `tasks_nowindows.go` (build-only stub). `process.go` +
`process_windows.go`/`process_nowindows.go` handle process spawn/kill and job-object cleanup.
`tasks_windows_test.go` holds the only unit tests in the repo outside the Pester suite (winget output
parsing). Flags use Go's single-dash style (`-fast`, `-throttle`, `-skip-node`) — do not assume PS1
parameter names carry over.

## Conventions specific to this repo

- Targets **PowerShell 7+** but falls back to Windows PowerShell 5.1; keep both paths working.
- PSScriptAnalyzer suppresses `PSAvoidUsingWriteHost`, `PSUseShouldProcessForStateChangingFunctions`,
  `PSUseApprovedVerbs`, empty-catch, etc. (see `PSScriptAnalyzerSettings.psd1`) — `Write-Host`,
  non-approved verbs, and intentional empty catches are accepted by design.
- When parity matters, a change to an updater in `updatescript.ps1` should be reflected in the Rust
  runner's corresponding `*_args`/`build_tasks` code (recent commits track "Rust parity").
- Keep PS1 file encoding stable (commits have fought encoding regressions).
- Helper scripts (`fix_installer.ps1`, `force_reinstall.ps1`, `bitwarden_cleanup.ps1`,
  `github-release-watcher.ps1`) are standalone, not part of the main flow.
- `rewrites/rust/target/` and `staging/` are build/run scratch — don't treat as source. So are the
  `rewrites/rust/run-*.log` files and the committed Go binaries.
- `CONTRIBUTING.md` is **stale**: it documents an `Invoke-Update -Name ... -Action { }` pattern with
  `-RequiresAdmin`/`-SlowOperation` flags. That function no longer exists — the script moved to
  `New-UpdateTask` + the scheduler. Follow `Get-UpdateTasks`, not CONTRIBUTING, when adding an updater.
