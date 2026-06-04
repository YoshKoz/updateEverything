# updateeverything Rust runner

Experimental Rust orchestrator for `updateEverything`.

The current runner is intentionally separate from `updatescript.ps1`. It handles:

- `--list-tasks`
- `--dry-run`
- `--only` and `--skip` filters
- `--fast` and `--ultra-fast` filters
- `update-config.json` reads for package skip lists
- command discovery and missing-tool skips
- per-task process execution
- JSON run summaries

Examples:

```powershell
cargo run -- --list-tasks
cargo run -- --dry-run
cargo run -- --dry-run --only winget-source,juliaup
cargo run -- --only juliaup --json-summary ..\staging\rust-summary.json
```

This does not yet replace the PowerShell script. Windows Update, Store app updates,
scheduled task registration, and admin-only cleanup should stay in PowerShell bridge
tasks until the Rust runner has equivalent coverage.
