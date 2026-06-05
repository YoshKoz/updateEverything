//go:build !windows

package main

import (
	"os"
	"os/exec"
)

func setWindowsNoWindow(cmd *exec.Cmd) {}

func isAdminProcess() bool {
	return os.Getuid() == 0
}

func isAdminWindows() bool {
	return false
}
