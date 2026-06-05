package main

import (
	"encoding/json"
	"fmt"
	"os"
	"regexp"
	"strings"
)

func taskWinget(timeoutSec int, skipPackages []string) *Task {
	skip := map[string]bool{}
	for _, p := range skipPackages {
		skip[strings.TrimSpace(p)] = true
	}
	return &Task{
		ID: "winget", Name: "winget", Category: "package-manager",
		RequiresCommand: []string{"winget"},
		TimeoutSec:      timeoutSec,
		Resources:       []string{"winget"},
		Run: func(tc *TaskContext) error {
			tc.Log("Updating winget sources...")
			srcRes, srcErr := Run("winget", RunOpts{
				Args:       []string{"source", "update", "--name", "winget"},
				TimeoutSec: 120,
			})
			tc.Log(srcRes.Lines...)
			if srcErr != nil {
				tc.Log("Source update warning: " + srcErr.Error())
			}

			tc.Log("Getting list of available upgrades...")
			listRes, listErr := Run("winget", RunOpts{
				Args:       []string{"upgrade", "--include-unknown"},
				TimeoutSec: 120,
			})
			if listErr != nil {
				return fmt.Errorf("failed to get upgrade list: %w", listErr)
			}

			ids := parseWingetIDs(listRes.Lines)
			if len(ids) == 0 {
				tc.Log(listRes.Lines...)
				tc.Log("No upgrades available")
				return nil
			}

			var toUpgrade []string
			for _, id := range ids {
				if skip[id] {
					tc.Log("Skipping winget package: " + id)
					continue
				}
				toUpgrade = append(toUpgrade, id)
			}

			if len(toUpgrade) == 0 {
				tc.Log("No winget packages to upgrade after skip list")
				return nil
			}

			if len(skip) == 0 {
				tc.Log(fmt.Sprintf("Found %d packages to upgrade, attempting bulk upgrade...", len(toUpgrade)))
				_, allErr := tc.RunCmd("winget", RunOpts{
					Args:            []string{"upgrade", "--all", "--include-unknown", "--silent", "--disable-interactivity", "--accept-package-agreements", "--accept-source-agreements", "--force"},
					SuccessExitCode: []int{0, -1978335189},
				})

				if allErr == nil {
					return nil
				}

				tc.Log("Bulk upgrade had failures, retrying individually...")
			} else {
				tc.Log(fmt.Sprintf("Found %d winget package(s) to upgrade after skip list.", len(toUpgrade)))
			}

			var failed []string
			for _, id := range toUpgrade {
				tc.Log("Upgrading winget package: " + id)
				res, err := tc.RunCmd("winget", RunOpts{
					Args:            []string{"upgrade", "--id", id, "--exact", "--include-unknown", "--silent", "--disable-interactivity", "--accept-package-agreements", "--accept-source-agreements", "--force"},
					SuccessExitCode: []int{0, -1978335189},
					Retries:         1,
				})
				if err != nil {
					locked := false
					for _, l := range res.Lines {
						if strings.Contains(l, "Access is denied") {
							locked = true
							break
						}
					}
					if locked {
						tc.Log("Skipping " + id + ": file locked (process is running)")
					} else {
						failed = append(failed, id)
					}
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("winget failed packages: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

var (
	sepRe   = regexp.MustCompile(`-{3,}`)
	validID = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._+\-]+$`)
)

func parseWingetIDs(lines []string) []string {
	var headerLine, sepLine string
	headerIdx := -1
	sepIdx := -1

	for i, line := range lines {
		if strings.Contains(line, "No installed package") || strings.Contains(line, "No available upgrade") {
			continue
		}
		if strings.Contains(line, "Name") && strings.Contains(line, "Id") && strings.Contains(line, "Version") {
			headerLine = line
			headerIdx = i
		}
		if sepRe.MatchString(line) && headerIdx >= 0 && i == headerIdx+1 {
			sepLine = line
			sepIdx = i
			break
		}
	}

	if headerLine == "" || sepLine == "" {
		return nil
	}

	idIdx := strings.Index(headerLine, "Id")
	if idIdx < 0 {
		return nil
	}

	verIdx := strings.Index(headerLine[idIdx:], "Version")
	if verIdx < 0 {
		return nil
	}
	idEnd := idIdx + verIdx

	seen := map[string]bool{}
	var ids []string

	for i := sepIdx + 1; i < len(lines); i++ {
		line := lines[i]
		if strings.Contains(line, "upgrades available") || strings.Contains(line, "package(s) have pins") {
			continue
		}
		if len(line) <= idIdx {
			continue
		}

		var id string
		if idEnd > idIdx && idEnd <= len(line) {
			id = strings.TrimSpace(line[idIdx:idEnd])
		} else {
			id = strings.TrimSpace(line[idIdx:])
		}

		if id != "" && validID.MatchString(id) && !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	}

	return ids
}

func taskScoop() *Task {
	return &Task{
		ID: "scoop", Name: "scoop", Category: "package-manager",
		RequiresCommand: []string{"scoop"},
		TimeoutSec:      900,
		Run: func(tc *TaskContext) error {
			if _, err := tc.RunCmd("scoop", RunOpts{Args: []string{"update"}}); err != nil {
				return err
			}
			if _, err := tc.RunCmd("scoop", RunOpts{Args: []string{"update", "*"}, Retries: 1}); err != nil {
				return err
			}
			tc.Log("Running scoop cleanup...")
			_, _ = tc.RunCmd("scoop", RunOpts{Args: []string{"cleanup", "*"}, TimeoutSec: 300})
			tc.Log("Running scoop cache rm *...")
			_, _ = tc.RunCmd("scoop", RunOpts{Args: []string{"cache", "rm", "*"}, TimeoutSec: 120})
			return nil
		},
	}
}

func taskChocolatey() *Task {
	return &Task{
		ID: "chocolatey", Name: "chocolatey", Category: "package-manager",
		RequiresCommand: []string{"choco"},
		RequiresAdmin:   true,
		TimeoutSec:      900,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("choco", RunOpts{
				Args:    []string{"upgrade", "all", "-y", "--no-progress", "--limit-output"},
				Retries: 1,
			})
			return err
		},
	}
}

func taskStoreApps(timeoutSec int) *Task {
	return &Task{
		ID: "store-apps", Name: "store-apps", Category: "system",
		RequiresCommand: []string{"winget"},
		TimeoutSec:      timeoutSec,
		Resources:       []string{"winget"},
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("winget", RunOpts{
				Args:            []string{"upgrade", "--all", "--source", "msstore", "--include-unknown", "--silent", "--disable-interactivity", "--accept-package-agreements", "--accept-source-agreements"},
				SuccessExitCode: []int{0, -1978335189},
				Retries:         1,
			})
			return err
		},
	}
}

func taskDefender() *Task {
	return &Task{
		ID: "defender", Name: "defender", Category: "system",
		RequiresAdmin:   true,
		RequiresCommand: []string{"pwsh"},
		TimeoutSec:      900,
		Run: func(tc *TaskContext) error {
			res, err := Run("pwsh", RunOpts{
				Args:       []string{"-NoProfile", "-Command", "Update-MpSignature -ErrorAction Stop"},
				TimeoutSec: 300,
			})
			tc.Log(res.Lines...)
			if err == nil {
				tc.Log("Defender signature updated via Update-MpSignature.")
				return nil
			}
			tc.Log("Update-MpSignature failed, trying MpCmdRun.exe fallback.")
			mpCmd := ""
			for _, candidate := range []string{
				os.Getenv("ProgramFiles") + `\Windows Defender\MpCmdRun.exe`,
				os.Getenv("ProgramFiles(x86)") + `\Windows Defender\MpCmdRun.exe`,
			} {
				if _, e := os.Stat(candidate); e == nil {
					mpCmd = candidate
					break
				}
			}
			if mpCmd == "" {
				return fmt.Errorf("MpCmdRun.exe not found")
			}
			res2, err2 := Run(mpCmd, RunOpts{Args: []string{"-SignatureUpdate"}, TimeoutSec: 900, Retries: 1})
			tc.Log(res2.Lines...)
			return err2
		},
	}
}

func taskWSL() *Task {
	return &Task{
		ID: "wsl", Name: "wsl", Category: "system",
		RequiresCommand: []string{"wsl"},
		TimeoutSec:      120,
		Run: func(tc *TaskContext) error {
			res, err := Run("wsl", RunOpts{
				Args:            []string{"--update"},
				SuccessExitCode: []int{0, -1},
				TimeoutSec:      120,
			})
			tc.Log(res.Lines...)
			if err != nil {
				for _, l := range res.Lines {
					if strings.Contains(l, "403") || strings.Contains(l, "0x80190193") {
						tc.Log("WSL update endpoint returned 403 — non-fatal.")
						return nil
					}
				}
				return err
			}
			return nil
		},
	}
}

func taskWSLDistros() *Task {
	linuxScript := `set -u
if command -v apt >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    sudo -n apt update && sudo -n DEBIAN_FRONTEND=noninteractive apt -y upgrade && sudo -n apt -y autoremove
  else
    echo "Skipping apt: sudo requires a password"
  fi
elif command -v pacman >/dev/null 2>&1; then
  if sudo -n true >/dev/null 2>&1; then
    sudo -n pacman -Syu --noconfirm
  else
    echo "Skipping pacman: sudo requires a password"
  fi
else
  echo "No supported package manager found"
fi`

	return &Task{
		ID: "wsl-distros", Name: "wsl-distros", Category: "system",
		RequiresCommand: []string{"wsl"},
		TimeoutSec:      3600,
		Resources:       []string{"wsl"},
		Run: func(tc *TaskContext) error {
			res, err := Run("wsl", RunOpts{Args: []string{"-l", "-q"}, TimeoutSec: 30})
			if err != nil {
				return err
			}
			var distros []string
			for _, l := range res.Lines {
				l = strings.Map(func(r rune) rune {
					if r == 0 {
						return -1
					}
					return r
				}, l)
				l = strings.TrimSpace(l)
				if l != "" {
					distros = append(distros, l)
				}
			}
			if len(distros) == 0 {
				tc.Log("No WSL distros found.")
				return nil
			}
			var failed []string
			for _, distro := range distros {
				tc.Log("Updating WSL distro: " + distro)
				r, err := Run("wsl", RunOpts{
					Args:       []string{"--distribution", distro, "--exec", "sh", "-lc", linuxScript},
					TimeoutSec: 1800,
				})
				tc.Log(r.Lines...)
				if err != nil {
					failed = append(failed, distro)
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("WSL distro updates failed: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

var validNpmName = regexp.MustCompile(`^(?:@[a-z0-9][a-z0-9._~-]*/)?[a-z0-9][a-z0-9._~-]*$`)

func taskNPM(skipPackages []string) *Task {
	skip := map[string]bool{}
	for _, p := range skipPackages {
		skip[strings.TrimSpace(p)] = true
	}

	return &Task{
		ID: "npm", Name: "npm", Category: "javascript",
		RequiresCommand: []string{"npm"},
		TimeoutSec:      900,
		Resources:       []string{"npm"},
		Run: func(tc *TaskContext) error {
			res, err := Run("npm", RunOpts{
				Args:            []string{"ls", "-g", "--depth=0", "--json"},
				SuccessExitCode: []int{0, 1},
				TimeoutSec:      60,
			})
			if err != nil {
				return err
			}
			raw := strings.Join(res.Lines, "\n")
			var tree struct {
				Dependencies map[string]json.RawMessage `json:"dependencies"`
			}
			if err := json.Unmarshal([]byte(raw), &tree); err != nil {
				return fmt.Errorf("npm ls parse error: %w", err)
			}
			if len(tree.Dependencies) == 0 {
				tc.Log("No global npm packages found.")
				return nil
			}

			var specs []string
			for name := range tree.Dependencies {
				if skip[name] {
					tc.Log("Skipping npm package: " + name)
					continue
				}
				if !validNpmName.MatchString(name) {
					tc.Log("Skipping invalid npm package name: " + name)
					continue
				}
				specs = append(specs, name+"@latest")
			}

			if len(specs) == 0 {
				tc.Log("No npm packages to update.")
				return nil
			}

			tc.Log(fmt.Sprintf("Batch updating %d npm packages...", len(specs)))
			args := []string{"install", "-g", "--no-fund", "--no-audit"}
			args = append(args, specs...)
			_, batchErr := tc.RunCmd("npm", RunOpts{
				Args:       args,
				TimeoutSec: 600,
				Retries:    1,
			})
			if batchErr == nil {
				return nil
			}

			tc.Log("Batch npm install had failures, retrying individually...")
			var failed []string
			for _, spec := range specs {
				tc.Log("Updating npm package: " + spec)
				_, err := tc.RunCmd("npm", RunOpts{
					Args:       []string{"install", "-g", spec, "--no-fund", "--no-audit"},
					TimeoutSec: 300,
					Retries:    2,
				})
				if err != nil {
					failed = append(failed, spec)
				}
			}
			if len(failed) > 0 {
				return fmt.Errorf("npm failed packages: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}

func taskPNPM() *Task {
	return &Task{
		ID: "pnpm", Name: "pnpm", Category: "javascript",
		RequiresCommand: []string{"pnpm"}, TimeoutSec: 300,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("pnpm", RunOpts{Args: []string{"self-update"}, Retries: 1})
			return err
		},
	}
}

func taskYarn() *Task {
	return &Task{
		ID: "yarn", Name: "yarn", Category: "javascript",
		RequiresCommand: []string{"yarn"}, TimeoutSec: 300,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("yarn", RunOpts{Args: []string{"global", "upgrade"}, Retries: 1})
			return err
		},
	}
}

func taskBun() *Task {
	return &Task{
		ID: "bun", Name: "bun", Category: "javascript",
		RequiresCommand: []string{"bun"}, TimeoutSec: 180,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("bun", RunOpts{Args: []string{"upgrade"}, Retries: 1})
			return err
		},
	}
}

func taskDeno() *Task {
	return &Task{
		ID: "deno", Name: "deno", Category: "javascript",
		RequiresCommand: []string{"deno"}, TimeoutSec: 180,
		Run: func(tc *TaskContext) error {
			_, err := tc.RunCmd("deno", RunOpts{Args: []string{"upgrade"}, Retries: 1})
			return err
		},
	}
}

func taskMise() *Task {
	return &Task{
		ID: "mise", Name: "mise", Category: "version-manager",
		RequiresCommand: []string{"mise"}, TimeoutSec: 300,
		Run: func(tc *TaskContext) error {
			if _, err := tc.RunCmd("mise", RunOpts{Args: []string{"self-update", "--yes"}, Retries: 1}); err != nil {
				return err
			}
			_, err := tc.RunCmd("mise", RunOpts{Args: []string{"upgrade", "--yes"}, Retries: 1})
			return err
		},
	}
}

func taskVolta() *Task {
	return &Task{
		ID: "volta", Name: "volta", Category: "version-manager",
		RequiresCommand: []string{"volta"}, TimeoutSec: 600,
		Run: func(tc *TaskContext) error {
			tc.Log("Updating Volta...")
			if _, err := tc.RunCmd("volta", RunOpts{Args: []string{"install", "volta"}, Retries: 1}); err != nil {
				return err
			}

			tc.Log("Updating Node.js via Volta...")
			tc.RunCmd("volta", RunOpts{Args: []string{"install", "node"}, TimeoutSec: 300})

			tc.Log("Listing Volta-managed packages...")
			res, err := Run("volta", RunOpts{
				Args:       []string{"list", "all", "--format", "plain"},
				TimeoutSec: 30,
			})
			if err != nil {
				return err
			}

			var packages []string
			for _, line := range res.Lines {
				if strings.HasPrefix(line, "package ") {
					parts := strings.Fields(line)
					if len(parts) >= 2 {
						pkgSpec := parts[1]
						if idx := strings.Index(pkgSpec, "@"); idx > 0 {
							pkgName := pkgSpec[:idx]
							packages = append(packages, pkgName)
						}
					}
				}
			}

			if len(packages) == 0 {
				tc.Log("No Volta-managed packages found.")
				return nil
			}

			tc.Log(fmt.Sprintf("Updating %d Volta-managed packages...", len(packages)))
			var failed []string
			for _, pkg := range packages {
				tc.Log("Updating Volta package: " + pkg)
				_, err := tc.RunCmd("volta", RunOpts{
					Args:       []string{"install", pkg},
					TimeoutSec: 120,
					Retries:    1,
				})
				if err != nil {
					failed = append(failed, pkg)
				}
			}

			if len(failed) > 0 {
				return fmt.Errorf("volta failed packages: %s", strings.Join(failed, ", "))
			}
			return nil
		},
	}
}
