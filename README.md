# updateEverything

Simple Windows update script for package managers, system components, and dev tools.

## Quick start

```powershell
./updatescript.ps1
./updatescript.ps1 -AutoElevate
./updatescript.ps1 -FastMode
./updatescript.ps1 -UltraFast
```

## What it updates

- Package managers: Winget, Scoop, Chocolatey
- System: Windows Update, Store apps, WSL, Defender
- Dev tools: npm/pnpm/bun/deno, rust, go, dotnet, python/pip, uv, pipx, gh extensions, VS Code extensions, and more
- Cleanup: temp files, DNS cache, recycle bin, optional deep clean

## Common options

```powershell
./updatescript.ps1 -SkipWindowsUpdate -SkipCleanup
./updatescript.ps1 -SkipNode -SkipRust -SkipGo
./updatescript.ps1 -DryRun
./updatescript.ps1 -WhatChanged
./updatescript.ps1 -Schedule -ScheduleTime "03:00"
```

## Parameters (actual script)

- SkipWindowsUpdate
- SkipReboot
- SkipDestructive
- FastMode
- UltraFast
- NoElevate
- AutoElevate
- NoPause
- SkipWSL
- SkipWSLDistros
- SkipDefender
- SkipStoreApps
- SkipUVTools
- SkipVSCodeExtensions
- SkipPoetry
- SkipComposer
- SkipRuby
- SkipPowerShellModules
- SkipCleanup
- WingetTimeoutSec (default: 300)
- Schedule
- ScheduleTime (default: 03:00)
- LogPath
- SkipNode
- SkipRust
- SkipGo
- SkipFlutter
- SkipGitLFS
- DeepClean
- UpdateOllamaModels
- WhatChanged
- DryRun
- ParallelThrottle (1-10, default: 4)

## Requirements

- Windows
- PowerShell 7+ recommended
- Admin rights recommended for full run
- PSWindowsUpdate module for Windows Update section

## Notes

- This repo includes helper scripts for installer repair:
  - fix_installer.ps1
  - force_reinstall.ps1

## License

MIT. See LICENSE.

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.
