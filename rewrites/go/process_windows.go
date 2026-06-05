package main

import (
	"os/exec"
	"syscall"
	"unsafe"
)

func setWindowsNoWindow(cmd *exec.Cmd) {
	if cmd.SysProcAttr == nil {
		cmd.SysProcAttr = &syscall.SysProcAttr{}
	}
	cmd.SysProcAttr.CreationFlags = 0x08000000
}

func isAdminProcess() bool {
	return isAdminWindows()
}

func isAdminWindows() bool {
	defer func() { recover() }()

	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	advapi32 := syscall.NewLazyDLL("advapi32.dll")
	getCurrentProcess := kernel32.NewProc("GetCurrentProcess")
	openProcessToken := advapi32.NewProc("OpenProcessToken")
	getTokenInformation := advapi32.NewProc("GetTokenInformation")
	closeHandle := kernel32.NewProc("CloseHandle")

	const TOKEN_QUERY = 0x0008
	const TokenElevation = 20

	handle, _, _ := getCurrentProcess.Call()
	var token syscall.Handle
	ret, _, _ := openProcessToken.Call(handle, TOKEN_QUERY, uintptr(unsafe.Pointer(&token)))
	if ret == 0 {
		return false
	}
	defer closeHandle.Call(uintptr(token))

	var elevation uint32
	var returnedLength uint32
	ret, _, _ = getTokenInformation.Call(
		uintptr(token),
		TokenElevation,
		uintptr(unsafe.Pointer(&elevation)),
		uintptr(unsafe.Sizeof(elevation)),
		uintptr(unsafe.Pointer(&returnedLength)),
	)
	return ret != 0 && elevation != 0
}
