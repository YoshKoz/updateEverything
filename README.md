# Windows Update Script (Update-Everything)

A comprehensive, all-in-one PowerShell script that updates **everything** on your Windows system in a single run — package managers, system components, development tools, and more.

![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell) ![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows) ![License](https://img.shields.io/badge/License-MIT-green)

---

## Features

### Package Managers
| Manager | What it does |
|---|---|
| **Winget** | Upgrades all packages (winget + MS Store sources) with timeout protection, automatic retry of failed packages |
| **Scoop** | Updates all buckets and packages, cleans up old versions and cache |
| **Chocolatey** | Upgrades all packages (requires admin) |

### Windows Components
| Component | What it does |
|---|---|
| **Windows Update** | Installs all non-driver updates via PSWindowsUpdate with retry logic for 0x800704c7 errors |
| **Microsoft Store Apps** | Triggers Store app update scans via MDM/CIM |
| **WSL** | Updates the WSL kernel and optionally runs `apt-get upgrade` / `pacman -Syu` / `zypper update` inside each distro |
| **Microsoft Defender** | Updates antivirus signatures |

### Development Tools
| Tool | What it does |
|---|---|
| **npm** | Updates npm itself and all global packages |
| **pnpm** | Updates global pnpm packages |
| **Bun** | Self-upgrades Bun |
| **Deno** | Self-upgrades Deno (detects Scoop/winget-managed installs) |
| **Python / pip** | Upgrades pip itself (auto-detects Python install location) |
| **uv** | Self-updates the uv package manager |
| **uv tools** | Upgrades all uv-managed tool installs |
| **Rust / rustup** | Runs `rustup update` |
| **Cargo binaries** | Updates global Cargo binaries via `cargo-update` |
| **Go** | Updates Go via winget (detects Scoop-managed installs) |
| **.NET tools** | Updates all global .NET tools, checks NuGet API for latest versions |
| **.NET workloads** | Runs `dotnet workload update` |
| **GitHub CLI extensions** | Runs `gh extension upgrade --all` |
| **VS Code extensions** | Runs `code --update-extensions` |
| **pipx** | Upgrades all pipx-managed Python CLI tools |
| **Poetry** | Self-updates Poetry |
| **Composer** | Self-updates Composer and global PHP packages |
| **RubyGems** | Updates RubyGems system and all gems |
| **Oh My Posh** | Self-upgrades Oh My Posh |
| **yt-dlp** | Updates yt-dlp (detects pip/scoop/standalone installs) |
| **fnm** | Detects and reports fnm management method |
| **mise** | Self-updates mise and upgrades all plugins/tools |
| **juliaup** | Updates Julia via juliaup |
| **PowerShell modules** | Updates all installed PSResources (PSResourceGet) and PowerShellGet modules, with parallel support |

### System Cleanup
| Action | What it does |
|---|---|
| **Temp files** | Removes temp files older than 7 days |
| **DNS cache** | Flushes the DNS client cache |
| **Recycle Bin** | Empties the Recycle Bin |
| **Crash dumps** | Clears local crash dump files |
| **WER reports** | Clears Windows Error Reporting queue |
| **DISM** | Cleans the WinSxS component store (admin only) |
| **Delivery Optimization** | Clears the DO cache (admin only) |
| **Prefetch** | Clears prefetch files (admin only, unless `-SkipDestructive`) |

---

## Requirements

- **PowerShell 7+** (recommended for `-Parallel` support; falls back to Windows PowerShell)
- **Administrator** privileges for: Windows Update, Chocolatey, WSL, Defender, Store apps, DISM cleanup
- **PSWindowsUpdate** module for Windows Update (`Install-Module PSWindowsUpdate -Force`)

---

## Installation

```powershell
# Clone this repo
git clone https://github.com/YoshKoz/windows-update-script.git
cd windows-update-script

# Or just download the script directly
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YoshKoz/windows-update-script/main/updatescript.ps1" -OutFile updatescript.ps1
```

---

## Usage

### Basic (run in current terminal)
```powershell
.\updatescript.ps1
```

### Run elevated (opens a new UAC-prompt window)
```powershell
.\updatescript.ps1 -AutoElevate
```

### Quick run (skip slow operations)
```powershell
.\updatescript.ps1 -FastMode
```

### Dry run (preview what would be updated)
```powershell
.\updatescript.ps1 -DryRun
```

### Skip specific components
```powershell
.\updatescript.ps1 -SkipWindowsUpdate -SkipWSL -SkipCleanup
```

### Schedule daily automatic updates
```powershell
.\updatescript.ps1 -Schedule -ScheduleTime "03:00"
```

### Log output to file
```powershell
.\updatescript.ps1 -LogPath "C:\Logs\update.log"
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-SkipWindowsUpdate` | Switch | `$false` | Skip Windows Update |
| `-SkipReboot` | Switch | `$false` | Suppress automatic reboot after updates |
| `-SkipDestructive` | Switch | `$false` | Skip destructive cleanup tasks (Go mod cache, prefetch) |
| `-FastMode` | Switch | `$false` | Skip slow operations (RubyGems, Volta, fnm, mise, etc.) |
| `-NoElevate` | Switch | `$false` | Prevent auto-elevation even with `-AutoElevate` |
| `-AutoElevate` | Switch | `$false` | Relaunch the script as Administrator (opens new window) |
| `-NoPause` | Switch | `$false` | Skip the "Press Enter to close" prompt (for CI / VS Code) |
| `-Parallel` | Switch | `$false` | Enable parallel updates where supported (PS7+) |
| `-DryRun` | Switch | `$false` | Show what would be updated without making changes |
| `-SkipWSL` | Switch | `$false` | Skip WSL kernel update |
| `-SkipWSLDistros` | Switch | `$false` | Skip updating packages inside WSL distros |
| `-SkipDefender` | Switch | `$false` | Skip Defender signature update |
| `-SkipStoreApps` | Switch | `$false` | Skip Microsoft Store app scan |
| `-SkipUVTools` | Switch | `$false` | Skip uv tool upgrades |
| `-SkipVSCodeExtensions` | Switch | `$false` | Skip VS Code extension updates |
| `-SkipPoetry` | Switch | `$false` | Skip Poetry self-update |
| `-SkipComposer` | Switch | `$false` | Skip Composer updates |
| `-SkipRuby` | Switch | `$false` | Skip RubyGems updates |
| `-SkipPowerShellModules` | Switch | `$false` | Skip PowerShell module updates |
| `-SkipCleanup` | Switch | `$false` | Skip the system cleanup section |
| `-WingetTimeoutSec` | Int | `300` | Per-call timeout for winget in seconds |
| `-Schedule` | Switch | `$false` | Register a daily scheduled task |
| `-ScheduleTime` | String | `"03:00"` | Time for the scheduled task |
| `-LogPath` | String | | Path to write a transcript log file |

---

## How It Works

1. **Self-elevation** — If `-AutoElevate` is passed and the script isn't running as admin, it relaunches itself elevated via UAC, forwarding all parameters.
2. **Smart detection** — Each update section checks whether its tool is installed (`Test-Command`) before attempting updates. Missing tools are silently skipped.
3. **Managed-install detection** — For tools like Deno, Go, uv, Oh My Posh, etc., the script detects whether they're managed by Scoop or winget and skips redundant self-updates.
4. **Winget timeout protection** — Winget calls use `Start-Process` with file-based I/O redirection and a configurable timeout to prevent hanging on stuck installers.
5. **Automatic retry** — Failed winget packages are detected by parsing output and retried individually with `--force`.
6. **Windows Update resilience** — Uses `-IgnoreReboot` to prevent `0x800704c7` (ERROR_CANCELLED), with automatic service restart and retry.
7. **Clean output** — ANSI escape sequences, progress bars, and spinner frames are stripped from tool output for clean terminal display.
8. **Summary report** — At the end, a color-coded summary shows succeeded, failed, and skipped components with total elapsed time.

---

## Example Output

```
======================================================
  Update-Everything v2.2.0  |  2026-03-04 14:30
  Running as Administrator
======================================================

======================================================
 Scoop
======================================================
[OK] Scoop updated (12.3s)

======================================================
 Winget
======================================================
  Upgrading all (winget source, timeout: 300s)...
  Successfully installed Package.Name v1.2.3
[OK] Winget updated (45.2s)

...

======================================================
 UPDATE COMPLETE -- 00:03:42
======================================================

[OK] Succeeded (15): Scoop, Winget, Chocolatey, WindowsUpdate, ...
[!] Skipped   (3): WSL, Poetry, Composer
```

---

## Scheduling

To run the script automatically every day at 3 AM:

```powershell
.\updatescript.ps1 -Schedule -ScheduleTime "03:00"
```

This creates a Windows Scheduled Task named `DailySystemUpdate` that runs elevated with `-SkipReboot`.

---

## Troubleshooting

| Issue | Solution |
|---|---|
| `0x800704c7` during Windows Update | Already handled with retry logic; ensure no pending reboots |
| Winget hangs indefinitely | Increase `-WingetTimeoutSec` or check for stuck installers in Task Manager |
| `PSWindowsUpdate module not found` | Run `Install-Module PSWindowsUpdate -Force` |
| Admin-only tasks skipped | Run with `-AutoElevate` or start terminal as Administrator |
| Scoop/winget not found | Install them first: [scoop.sh](https://scoop.sh), [winget](https://aka.ms/getwinget) |

---

## License

MIT License — free to use, modify, and distribute.
