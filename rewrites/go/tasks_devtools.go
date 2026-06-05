package main

import (
	"encoding/json"
	"fmt"
	"strings"
)

// ---- system tasks (windows update, ollama) ----

func taskWindowsUpdate() *Task {
	script := `
$ErrorActionPreference = 'Stop'
if (Get-Command Install-WindowsUpdate -EA SilentlyContinue) {
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -EA Stop | Out-String
    return
}
Write-Output 'PSWindowsUpdate not found; using WUA COM fallback.'
$session = New-Object -ComObject Microsoft.Update.Session
$session.ClientApplicationID = 'Update-Everything'
$searcher = $session.CreateUpdateSearcher()
$result = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
$count = [int]$result.Updates.Count
Write-Output "Windows Update available updates: $count"
if ($count -eq 0) { return }
$updates = New-Object -ComObject Microsoft.Update.UpdateColl
for ($i = 0; $i -lt $result.Updates.Count; $i++) {
    $u = $result.Updates.Item($i)
    Write-Output "Selected: $($u.Title)"
    if (-not $u.EulaAccepted) { $u.AcceptEula() }
    [void]$updates.Add($u)
}
$dl = $session.CreateUpdateDownloader(); $dl.Updates = $updates
$dlResult = $dl.Download()
Write-Output "Download result: $($dlResult.ResultCode)"
if ($dlResult.ResultCode -notin @(2,3)) { throw "Download failed: $($dlResult.ResultCode)" }
$inst = $session.CreateUpdateInstaller(); $inst.Updates = $updates
$instResult = $inst.Install()
Write-Output "Install result: $($instResult.ResultCode); reboot required: $($instResult.RebootRequired)"
if ($instResult.ResultCode -notin @(2,3)) { throw "Install failed: $($instResult.ResultCode)" }
`
	return &Task{
		ID: "windows-update", Name: "windows-update", Category: "system",
		RequiresAdmin:   true,
		RequiresCommand: []string{"pwsh"},
		TimeoutSec:      7200,
		Run: func(tc *TaskContext) error {
			res, err := Run("pwsh", RunOpts{
				Args:       []string{"-NoProfile", "-Command", script},
				TimeoutSec: 7200,
			})
			tc.Log(res.Lines...)
			return err
		},
	}
}

func taskOllamaModels(timeoutSec int) *Task {
	return &Task{
		ID: "ollama-models", Name: "ollama-models", Category: "ai",
		RequiresCommand: []string{"ollama"},
		TimeoutSec:      timeoutSec,
		Run: func(tc *TaskContext) error {
			listRes, err := Run("ollama", RunOpts{Args: []string{"list"}, TimeoutSec: 60})
			tc.Log(listRes.Lines...)
			if err != nil {
				return err
			}
			var models []string
			for _, l := range listRes.Lines[1:] {
				parts := strings.Fields(l)
				if len(parts) > 0 && parts[0] != "" {
					models = append(models, parts[0])
				}
			}
			if len(models) == 0 {
				tc.Log("No Ollama models found.")
				return nil
			}
			var failed []string
			for _, model := range models {
				tc.Log("Updating Ollama model: " + model)
				res, err := tc.RunCmd("ollama", RunOpts{
					Args:       []string{"pull", model},
					TimeoutSec: timeoutSec,
					Retries:    1,
				})
				tc.Log(res.Lines...)
				if err != nil {
					failed = append(failed, model)
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("ollama model updates failed: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

func taskPip(skipPackages []string) *Task {
	skip := map[string]bool{}
	for _, p := range skipPackages {
		skip[strings.TrimSpace(p)] = true
	}
	return &Task{
		ID: "pip", Name: "pip", Category: "python",
		RequiresCommand: []string{"python"},
		TimeoutSec:      900,
		Resources:       []string{"pip"},
		Run: func(tc *TaskContext) error {
			if _, err := tc.RunCmd("python", RunOpts{
				Args: []string{"-m", "pip", "install", "--upgrade", "pip"}, Retries: 1,
			}); err != nil {
				return err
			}
			res, err := Run("python", RunOpts{
				Args:       []string{"-m", "pip", "list", "--outdated", "--format=json"},
				TimeoutSec: 60,
			})
			if err != nil {
				return err
			}
			raw := strings.Join(res.Lines, "\n")
			var outdated []struct {
				Name          string `json:"name"`
				Version       string `json:"version"`
				LatestVersion string `json:"latest_version"`
			}
			if err := json.Unmarshal([]byte(raw), &outdated); err != nil || len(outdated) == 0 {
				tc.Log("No outdated pip packages found.")
				return nil
			}

			var toUpgrade []string
			for _, pkg := range outdated {
				if skip[pkg.Name] {
					tc.Log("Skipping pip package: " + pkg.Name)
					continue
				}
				tc.Log(fmt.Sprintf("Will upgrade pip package: %s %s -> %s", pkg.Name, pkg.Version, pkg.LatestVersion))
				toUpgrade = append(toUpgrade, pkg.Name)
			}

			if len(toUpgrade) == 0 {
				tc.Log("No pip packages to update.")
				return nil
			}

			tc.Log(fmt.Sprintf("Batch upgrading %d pip packages...", len(toUpgrade)))
			args := []string{"-m", "pip", "install", "--upgrade"}
			args = append(args, toUpgrade...)
			_, batchErr := tc.RunCmd("python", RunOpts{
				Args:       args,
				TimeoutSec: 600,
				Retries:    1,
			})
			if batchErr == nil {
				return nil
			}

			tc.Log("Batch pip upgrade had failures, retrying individually...")
			var failed []string
			for _, pkg := range toUpgrade {
				tc.Log("Upgrading pip package: " + pkg)
				_, err := tc.RunCmd("python", RunOpts{
					Args: []string{"-m", "pip", "install", "--upgrade", pkg}, Retries: 1,
				})
				if err != nil {
					failed = append(failed, pkg)
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("pip failed packages: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

func taskPipx() *Task {
	return &Task{
		ID: "pipx", Name: "pipx", Category: "python",
		RequiresCommand: []string{"pipx"}, TimeoutSec: 600,
		Resources: []string{"pip"},
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("pipx", RunOpts{Args: []string{"upgrade-all"}, Retries: 1})
			return err
		},
	}
}

func taskUV() *Task {
	return &Task{
		ID: "uv", Name: "uv", Category: "python",
		RequiresCommand: []string{"uv"}, TimeoutSec: 180,
		Run: func(tc *TaskContext) error {
			res, err := Run("uv", RunOpts{
				Args:            []string{"self", "update"},
				SuccessExitCode: []int{0, 1},
				TimeoutSec:      120,
			})
			tc.Log(res.Lines...)
			for _, l := range res.Lines {
				if strings.Contains(l, "standalone installation") {
					tc.Log("uv self-update skipped: managed install.")
					return nil
				}
			}
			return err
		},
	}
}

func taskUVTools() *Task {
	return &Task{
		ID: "uv-tools", Name: "uv-tools", Category: "python",
		RequiresCommand: []string{"uv"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("uv", RunOpts{Args: []string{"tool", "upgrade", "--all"}, Retries: 1})
			return err
		},
	}
}

func taskPoetry() *Task {
	return &Task{
		ID: "poetry", Name: "poetry", Category: "python",
		RequiresCommand: []string{"poetry"}, TimeoutSec: 300,
		Run: func(tc *TaskContext) error {
			if commandExists("pipx") {
				tc.Log("Poetry is pipx-managed, using pipx upgrade")
				_, err := tc.RunCmd("pipx", RunOpts{Args: []string{"upgrade", "poetry"}, Retries: 2})
				return err
			}
			_, err := tc.RunCmd("poetry", RunOpts{Args: []string{"self", "update"}, Retries: 2})
			return err
		},
	}
}

func taskRustup() *Task {
	return &Task{
		ID: "rustup", Name: "rustup", Category: "systems-language",
		RequiresCommand: []string{"rustup"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("rustup", RunOpts{Args: []string{"update"}, Retries: 1})
			return err
		},
	}
}

func taskCargo() *Task {
	return &Task{
		ID: "cargo", Name: "cargo", Category: "systems-language",
		RequiresCommand: []string{"cargo"}, TimeoutSec: 1800,
		Run: func(tc *TaskContext) error {
			if !commandExists("cargo-install-update") {
				if _, err := tc.RunCmd("cargo", RunOpts{
					Args: []string{"install", "cargo-update", "-q"}, Retries: 1,
				}); err != nil {
					return err
				}
			}
			_, err := tc.RunCmd("cargo", RunOpts{Args: []string{"install-update", "-a"}, Retries: 1})
			return err
		},
	}
}

func taskGo() *Task {
	return &Task{
		ID: "go", Name: "go", Category: "systems-language",
		RequiresCommand: []string{"go"}, TimeoutSec: 900,
		Run: func(tc *TaskContext) error {
			tools := []string{
				"golang.org/x/tools/gopls@latest",
				"golang.org/x/tools/cmd/goimports@latest",
				"golang.org/x/tools/cmd/godoc@latest",
				"golang.org/x/lint/golint@latest",
				"github.com/go-delve/delve/cmd/dlv@latest",
			}

			res, err := Run("go", RunOpts{
				Args:       []string{"version"},
				TimeoutSec: 10,
			})
			tc.Log(res.Lines...)
			if err != nil {
				return err
			}

			var failed []string
			for _, tool := range tools {
				tc.Log("Installing/updating Go tool: " + tool)
				_, err := tc.RunCmd("go", RunOpts{
					Args:       []string{"install", tool},
					TimeoutSec: 300,
					Retries:    1,
				})
				if err != nil {
					failed = append(failed, tool)
				}
			}

			tc.Log("Cleaning Go module cache...")
			tc.RunCmd("go", RunOpts{
				Args:       []string{"clean", "-modcache"},
				TimeoutSec: 60,
			})

			if len(failed) > 0 {
				return fmt.Errorf("go failed tools: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

func taskFlutter() *Task {
	return &Task{
		ID: "flutter", Name: "flutter", Category: "systems-language",
		RequiresCommand: []string{"flutter"}, TimeoutSec: 900,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("flutter", RunOpts{Args: []string{"upgrade", "--force"}, Retries: 1})
			return err
		},
	}
}

func taskDotnetTools() *Task {
	return &Task{
		ID: "dotnet-tools", Name: "dotnet-tools", Category: "dotnet",
		RequiresCommand: []string{"dotnet"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("dotnet", RunOpts{
				Args: []string{"tool", "update", "--global", "--all"}, Retries: 1,
			})
			return err
		},
	}
}

func taskDotnetWorkloads() *Task {
	return &Task{
		ID: "dotnet-workloads", Name: "dotnet-workloads", Category: "dotnet",
		RequiresCommand: []string{"dotnet"},
		RequiresAdmin:   true,
		TimeoutSec:      3600,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("dotnet", RunOpts{
				Args: []string{"workload", "update"}, TimeoutSec: 3600, Retries: 1,
			})
			return err
		},
	}
}

func taskRubyGems() *Task {
	return &Task{
		ID: "ruby-gems", Name: "ruby-gems", Category: "runtime",
		RequiresCommand: []string{"gem"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			if _, err := tc.RunCmd("gem", RunOpts{Args: []string{"update", "--system"}, Retries: 1}); err != nil {
				return err
			}
			_, err := tc.RunCmd("gem", RunOpts{Args: []string{"update"}, Retries: 1})
			return err
		},
	}
}

func taskComposer() *Task {
	return &Task{
		ID: "composer", Name: "composer", Category: "runtime",
		RequiresCommand: []string{"composer"}, TimeoutSec: 300,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("composer", RunOpts{Args: []string{"self-update"}, Retries: 1})
			return err
		},
	}
}

func taskVSCodeExtensions() *Task {
	return &Task{
		ID: "vscode-extensions", Name: "vscode-extensions", Category: "dev-tools",
		RequiresCommand: []string{"code"}, TimeoutSec: 900,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("code", RunOpts{
				Args: []string{"--update-extensions"},
				Env:  []string{"NODE_NO_WARNINGS=1"},
				Retries: 1,
			})
			return err
		},
	}
}

func taskGitLFS() *Task {
	return &Task{
		ID: "git-lfs", Name: "git-lfs", Category: "dev-tools",
		RequiresCommand: []string{"git-lfs"}, TimeoutSec: 120,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("git", RunOpts{Args: []string{"lfs", "install", "--skip-repo"}})
			return err
		},
	}
}

func taskGHExtensions() *Task {
	return &Task{
		ID: "gh-extensions", Name: "gh-extensions", Category: "dev-tools",
		RequiresCommand: []string{"gh"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			tc.Log("Trying gh extension upgrade --all...")
			_, err := tc.RunCmd("gh", RunOpts{
				Args:            []string{"extension", "upgrade", "--all"},
				TimeoutSec:      300,
				SuccessExitCode: []int{0, 1},
				Retries:         1,
			})
			if err == nil {
				return nil
			}

			tc.Log("Bulk gh extension upgrade failed, trying individually...")
			res, listErr := Run("gh", RunOpts{Args: []string{"extension", "list"}, TimeoutSec: 30})
			if listErr != nil {
				return listErr
			}
			var exts []string
			for _, l := range res.Lines {
				parts := strings.Fields(l)
				if len(parts) > 0 && strings.HasPrefix(parts[0], "gh-") {
					exts = append(exts, parts[0])
				}
			}
			var failed []string
			for _, ext := range exts {
				_, err := tc.RunCmd("gh", RunOpts{Args: []string{"extension", "upgrade", ext}, Retries: 1})
				if err != nil {
					failed = append(failed, ext)
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("gh extension updates failed: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

func taskPSModules() *Task {
	return &Task{
		ID: "powershell-modules", Name: "powershell-modules", Category: "powershell",
		RequiresCommand: []string{"pwsh"},
		TimeoutSec:      1800,
		Run: func(tc *TaskContext) error {
			script := `
$ErrorActionPreference = 'Stop'
if (Get-Command Update-PSResource -EA SilentlyContinue) {
    $args = @{ Name = '*'; ErrorAction = 'Stop' }
    $cmd = Get-Command Update-PSResource -EA SilentlyContinue
    if ($cmd.Parameters.ContainsKey('AcceptLicense')) { $args.AcceptLicense = $true }
    Update-PSResource @args | Out-String
} elseif (Get-Command Update-Module -EA SilentlyContinue) {
    $failed = @()
    Get-InstalledModule -EA SilentlyContinue | ForEach-Object {
        try { Update-Module -Name $_.Name -Force -EA Stop }
        catch { Write-Output "Failed: $($_.Name): $($_.Exception.Message)"; $failed += $_.Name }
    }
    if ($failed) { throw "Module update failures: $($failed -join ', ')" }
} else { Write-Output 'No PS module updater found.' }
`
			res, err := Run("pwsh", RunOpts{
				Args:       []string{"-NoProfile", "-Command", script},
				TimeoutSec: 1800,
			})
			tc.Log(res.Lines...)
			return err
		},
	}
}

func taskPSHelp() *Task {
	return &Task{
		ID: "powershell-help", Name: "powershell-help", Category: "powershell",
		RequiresCommand: []string{"pwsh"},
		TimeoutSec:      600,
		Run: func(tc *TaskContext) error {
			script := `
$ErrorActionPreference = 'SilentlyContinue'
Update-Help -Force -ErrorAction SilentlyContinue | Out-String
`
			res, err := Run("pwsh", RunOpts{
				Args:       []string{"-NoProfile", "-Command", script},
				TimeoutSec: 600,
			})
			tc.Log(res.Lines...)
			return err
		},
	}
}

func taskCleanup(tempDays int, deep bool, skipDestructive bool) *Task {
	return &Task{
		ID: "cleanup", Name: "cleanup", Category: "maintenance",
		RequiresCommand: []string{"pwsh"},
		TimeoutSec:      3600,
		Run: func(tc *TaskContext) error {
			script := fmt.Sprintf(`
$days = %d; $deep = $%v; $skipDestructive = $%v
$cutoff = (Get-Date).AddDays(-$days)
foreach ($path in @($env:TEMP, 'C:\Windows\Temp') | Where-Object { $_ -and (Test-Path -LiteralPath $_) }) {
    if ($skipDestructive) { Write-Output "Skip temp cleanup (SkipDestructive): $path"; continue }
    Write-Output "Cleaning temp files older than $days day(s): $path"
    Get-ChildItem -LiteralPath $path -Force -EA SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and $_.FullName -notmatch '\\WinGet(\\|$)' } |
        Remove-Item -Recurse -Force -EA SilentlyContinue
}
Clear-DnsClientCache -EA SilentlyContinue
if (-not $skipDestructive) { Clear-RecycleBin -Force -EA SilentlyContinue }
if ($deep -and -not $skipDestructive) {
    & DISM.exe /Online /Cleanup-Image /StartComponentCleanup
    Clear-DeliveryOptimizationCache -Force -EA SilentlyContinue
}
`, tempDays, deep, skipDestructive)
			res, err := Run("pwsh", RunOpts{
				Args:       []string{"-NoProfile", "-Command", script},
				TimeoutSec: 3600,
			})
			tc.Log(res.Lines...)
			return err
		},
	}
}
