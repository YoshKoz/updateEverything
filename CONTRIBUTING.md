# Contributing to Windows Update Script

Thanks for your interest in contributing! Here's how you can help.

## Reporting Issues

- **Search existing issues** before opening a new one
- Include your **PowerShell version** (`$PSVersionTable.PSVersion`)
- Include the **error output** — run with `-LogPath "C:\Temp\update.log"` and attach the log
- Mention which **update section** failed (e.g., "Winget", "WindowsUpdate")

## Suggesting New Tools

Want a new package manager or dev tool added? Open an issue with:

1. **Tool name** and what it does
2. **Update command** (e.g., `toolname update` or `toolname self-update`)
3. **How to detect** if it's installed (binary name, registry key, etc.)
4. Whether it **requires admin** privileges

## Pull Requests

### Adding a new update section

Each tool follows the same pattern using `Invoke-Update`:

```powershell
Invoke-Update -Name 'ToolName' -Title 'Display Name' -RequiresCommand 'tool-binary' -Action {
    $toolPath = (Get-Command tool-binary).Source
    if ($toolPath -like '*scoop*') {
        Write-Host "  Managed by Scoop (already updated)" -ForegroundColor Gray
        return
    }
    $out = (tool-binary update 2>&1 | Out-String).Trim()
    if ($out) { Write-Host $out -ForegroundColor Gray }
}
```

**Key flags:**
- `-RequiresCommand` — binary name to check; section is skipped if not found
- `-RequiresAdmin` — skip when not running elevated
- `-SlowOperation` — skip when `-FastMode` is used
- `-Disabled:$SkipSomething` — tie to a `-Skip*` parameter

### Code Style

- Use `Write-Status` for status messages (`Success`, `Warning`, `Error`, `Info`)
- Use `Write-FilteredOutput` for external tool output (strips ANSI codes)
- Detect Scoop/winget managed installs to avoid redundant updates
- Suppress noisy output with `2>&1 | Out-Null` where appropriate
- Add your section in the correct category (Package Managers / Windows Components / Development Tools)

### Testing

Before submitting:

1. Run with `-DryRun` to verify your section is detected
2. Run without admin to confirm `-RequiresAdmin` sections are skipped gracefully
3. Run with `-FastMode` if your section uses `-SlowOperation`
4. Test on a system where the tool is **not** installed to verify it's skipped

## Code of Conduct

Be respectful. Keep discussions constructive and focused on the code.
