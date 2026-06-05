//go:build !windows

package main

func taskWinget(_ int, _ []string) *Task {
	return &Task{ID: "winget", Name: "winget", Category: "package-manager", Disabled: true, DisabledReason: "windows only"}
}

func taskScoop() *Task {
	return &Task{ID: "scoop", Name: "scoop", Category: "package-manager", Disabled: true, DisabledReason: "windows only"}
}

func taskChocolatey() *Task {
	return &Task{ID: "chocolatey", Name: "chocolatey", Category: "package-manager", Disabled: true, DisabledReason: "windows only"}
}

func taskStoreApps(_ int) *Task {
	return &Task{ID: "store-apps", Name: "store-apps", Category: "system", Disabled: true, DisabledReason: "windows only"}
}

func taskDefender() *Task {
	return &Task{ID: "defender", Name: "defender", Category: "system", Disabled: true, DisabledReason: "windows only"}
}

func taskWSL() *Task {
	return &Task{ID: "wsl", Name: "wsl", Category: "system", Disabled: true, DisabledReason: "windows only"}
}

func taskWSLDistros() *Task {
	return &Task{ID: "wsl-distros", Name: "wsl-distros", Category: "system", Disabled: true, DisabledReason: "windows only"}
}

func taskNPM(_ []string) *Task {
	return &Task{ID: "npm", Name: "npm", Category: "javascript", Disabled: true, DisabledReason: "windows only"}
}

func taskPNPM() *Task {
	return &Task{ID: "pnpm", Name: "pnpm", Category: "javascript", Disabled: true, DisabledReason: "windows only"}
}

func taskYarn() *Task {
	return &Task{ID: "yarn", Name: "yarn", Category: "javascript", Disabled: true, DisabledReason: "windows only"}
}

func taskBun() *Task {
	return &Task{ID: "bun", Name: "bun", Category: "javascript", Disabled: true, DisabledReason: "windows only"}
}

func taskDeno() *Task {
	return &Task{ID: "deno", Name: "deno", Category: "javascript", Disabled: true, DisabledReason: "windows only"}
}

func taskMise() *Task {
	return &Task{ID: "mise", Name: "mise", Category: "version-manager", Disabled: true, DisabledReason: "windows only"}
}

func taskVolta() *Task {
	return &Task{ID: "volta", Name: "volta", Category: "version-manager", Disabled: true, DisabledReason: "windows only"}
}
