# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

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
- `-DryRun` mode to preview what would be updated
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
