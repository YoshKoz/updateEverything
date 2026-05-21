<!-- markdownlint-configure-file {"MD024": {"siblings_only": true}} -->

# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [6.5.0] — 2026-05-18

### Added

- **Winget source refresh** — runs `winget source update` as its own serialized package-manager task before winget upgrade scans.
- **Broader developer-tool coverage** — adds tasks for `yt-dlp`, `oh-my-posh`, `juliaup`, `volta`, and `fnm`.
- **uv-managed Python refresh** — detects installed uv CPython major/minor versions and reinstalls latest patch releases.
- **WSL zypper support** — WSL distro updates now cover openSUSE-style distros in addition to apt and pacman.

### Changed

- Fast mode now skips the new slower optional tool refresh tasks.
- Protected desktop apps are no longer skipped by default; the built-in protected app lists are empty and `-IncludeProtectedApps` remains only for backward compatibility.

### Fixed

- **oh-my-posh timeout reporting** — avoids 15-minute silent self-update hangs where possible, checks the winget package first, and fails the task if standalone upgrade times out.
- **pnpm broken shim detection** — treats missing `@pnpm/exe/pnpm.exe` shim output as a task failure even when the wrapper exits successfully.

## [2.6.3] — 2026-03-23

### Fixed

- **Winget pending parsing** — long package names no longer disappear from the preflight upgrade list or update summary when winget collapses table spacing
- **WSL broken-distro reporting** — distro startup failures are now detected before package-manager probing, so broken VHDX mounts report as WSL reachability failures instead of misleading `apt`/`pacman` errors
- **WSL sudo hangs** — distro package updates now use non-interactive sudo checks and fail fast with a warning when a password prompt would block the run

### Changed

- Final summary once again distinguishes `updated`, `checked`, `failed`, and `skipped` sections instead of collapsing everything non-failed into a single bucket

## [2.6.2] — 2026-03-19

### Fixed

- **WSL result handling** — a failed or timed-out distro update now marks the WSL section as failed instead of incorrectly reporting success
- **WSL platform errors** — `wsl --update` now checks its exit code and fails the section when the platform update itself errors
- **Dry-run safety** — `-DryRun` no longer runs cleanup tasks or writes winget snapshot state

### Changed

- README refreshed for current script behavior, version, and newer switches (`-NoParallel`, `-DryRun`, tool-specific skip flags)

## [2.5.0] — 2026-03-16

### Added

- **Smart unknown-version package upgrades** — detects apps whose installed version winget cannot determine (e.g. Node.js, Everything, ImageMagick, Stella) and upgrades them only when a new Available version appears in the winget catalog. Tracks state in `%LOCALAPPDATA%\Update-Everything\unknown_versions.json` to avoid unnecessary reinstalls. Replaces the need for Patch My PC Home Updater.
- **Winget source refresh** (`winget source update`) before upgrade passes to ensure the package catalog is current

### Changed

- Pending-check listing now includes `--include-unknown` so pre-hooks fire for unknown-version packages too

## [2.2.0] — 2026-03-04

### Added

- **Windows Update resilience** — `0x800704c7` (ERROR_CANCELLED) retry logic with automatic `wuauserv` service restart
- `-RecurseCycle 3` for chained Windows Update dependencies
- `-IgnoreReboot` to prevent WU agent from self-cancelling mid-install
- Dedicated reboot status check via `Get-WURebootStatus` after install completes
- `.NET Workloads` update section (`dotnet workload update`)
- **Git Credential Manager** — updates GCM via winget `Git.Git`
- **Oh My Posh** — detects Scoop/winget/MSIX installs, self-upgrades standalone
- **yt-dlp** — detects pip/scoop/standalone installs, updates appropriately
- **fnm** (Fast Node Manager) — detects install method
- **mise** (polyglot tool manager) — self-updates and upgrades all plugins
- **juliaup** — updates Julia via juliaup
- **WER cleanup** — clears Windows Error Reporting queue
- **Prefetch cleanup** — clears prefetch files (admin only, skipped with `-SkipDestructive`)
- `-SkipCleanup` switch to skip the entire system cleanup section
- Managed-install detection for Deno, Go, uv, Oh My Posh, Volta, fnm, mise, juliaup (Scoop/winget)

### Changed

- Winget timeout wrapper now uses `Start-Process` with file-based I/O redirection to avoid pipeline-buffering hangs and .NET event-handler thread crashes
- `.NET Tools` now checks NuGet API for latest version before updating (avoids unnecessary reinstalls)
- `Write-FilteredOutput` strips Unicode progress bar characters and ANSI escape sequences more aggressively

### Fixed

- npm global update bug: was using `@pkgs` (splatting syntax) instead of `$pkgs`
- Windows Update `0x800704c7` error caused by `-AutoReboot` conflicting with install

## [2.1.0] — 2026-02-15

### Added

- `-Parallel` switch for PS7+ parallel execution of PSResource updates
- `-Schedule` / `-ScheduleTime` to register a daily scheduled task
- `-LogPath` for transcript logging
- Winget failed-package detection and individual retry logic
- MS Store source support in winget
- Cargo global binaries via `cargo-update`
- GitHub CLI extension updates
- Defender signature update with active-AV detection

### Changed

- Self-elevation now forwards all parameters to the elevated process
- Scoop cleanup and cache removal after updates

## [2.0.0] — 2026-01-20

### Added

- Complete rewrite with `Invoke-Update` wrapper function
- Per-section timing with elapsed seconds
- Color-coded summary report (succeeded / failed / skipped)
- `-FastMode` to skip slow operations
- `-AutoElevate` for opt-in UAC elevation
- `-NoPause` for CI / VS Code terminal compatibility
- Smart tool detection — missing tools are silently skipped
- WSL distro package updates (apt, pacman, zypper)
- UTF-8 BOM-less console encoding

### Changed

- Modular architecture with `Invoke-Update` replacing inline blocks
- Winget calls now have configurable timeout (`-WingetTimeoutSec`)

## [1.0.0] — 2025-12-01

### Added

- Initial release
- Scoop, Winget, Chocolatey updates
- Windows Update via PSWindowsUpdate
- npm, pip, Rust, .NET tools
- Basic temp file and DNS cleanup
