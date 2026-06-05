package main

import (
	"bytes"
	"context"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"sync"
	"time"
)

var ansiRe = regexp.MustCompile(`\x1b\[[0-9;?]*[a-zA-Z]`)

func stripAnsi(s string) string {
	return ansiRe.ReplaceAllString(s, "")
}

func splitLines(s string) []string {
	normalized := strings.ReplaceAll(s, "\r\n", "\n")
	normalized = strings.ReplaceAll(normalized, "\r", "\n")
	raw := strings.Split(normalized, "\n")
	out := make([]string, 0, len(raw))
	for _, l := range raw {
		l = strings.TrimSpace(stripAnsi(l))
		if l != "" {
			out = append(out, l)
		}
	}
	return out
}

type resolvedCmd struct {
	exe  string
	args []string
}

var resolveCache = map[string]resolvedCmd{}
var resolveCacheMu sync.Mutex

func resolveExe(name string) (resolvedCmd, error) {
	resolveCacheMu.Lock()
	if rc, ok := resolveCache[name]; ok {
		resolveCacheMu.Unlock()
		return rc, nil
	}
	resolveCacheMu.Unlock()

	rc, err := resolveExeUncached(name)
	if err != nil {
		return rc, err
	}

	resolveCacheMu.Lock()
	resolveCache[name] = rc
	resolveCacheMu.Unlock()
	return rc, nil
}

func resolveExeUncached(name string) (resolvedCmd, error) {
	if name == "code" {
		return resolveVSCode()
	}
	if name == "winget" {
		comSpec := os.Getenv("ComSpec")
		if comSpec == "" {
			comSpec = "cmd.exe"
		}
		return resolvedCmd{exe: comSpec, args: []string{"/d", "/c", "winget"}}, nil
	}
	path, err := exec.LookPath(name)
	if err != nil {
		if name == "MpCmdRun" || name == "MpCmdRun.exe" {
			for _, candidate := range []string{
				filepath.Join(os.Getenv("ProgramFiles"), "Windows Defender", "MpCmdRun.exe"),
				filepath.Join(os.Getenv("ProgramFiles(x86)"), "Windows Defender", "MpCmdRun.exe"),
			} {
				if _, err2 := os.Stat(candidate); err2 == nil {
					return resolvedCmd{exe: candidate}, nil
				}
			}
		}
		return resolvedCmd{}, fmt.Errorf("command not found: %s", name)
	}
	ext := strings.ToLower(filepath.Ext(path))
	if ext == ".cmd" || ext == ".bat" {
		comSpec := os.Getenv("ComSpec")
		if comSpec == "" {
			comSpec = "cmd.exe"
		}
		return resolvedCmd{exe: comSpec, args: []string{"/d", "/c", "call", path}}, nil
	}
	return resolvedCmd{exe: path}, nil
}

func resolveVSCode() (resolvedCmd, error) {
	candidates := []string{
		filepath.Join(os.Getenv("LOCALAPPDATA"), "Programs", "Microsoft VS Code", "bin", "code.cmd"),
		`C:\Program Files\Microsoft VS Code\bin\code.cmd`,
		`C:\Program Files (x86)\Microsoft VS Code\bin\code.cmd`,
	}
	if p, err := exec.LookPath("code"); err == nil {
		candidates = append([]string{p}, candidates...)
	}
	for _, c := range candidates {
		if _, err := os.Stat(c); err == nil {
			ext := strings.ToLower(filepath.Ext(c))
			if ext == ".cmd" || ext == ".bat" {
				comSpec := os.Getenv("ComSpec")
				if comSpec == "" {
					comSpec = "cmd.exe"
				}
				return resolvedCmd{exe: comSpec, args: []string{"/d", "/c", "call", c}}, nil
			}
			return resolvedCmd{exe: c}, nil
		}
	}
	return resolvedCmd{}, fmt.Errorf("VS Code CLI not found")
}

type RunOpts struct {
	Args            []string
	TimeoutSec      int
	SuccessExitCode []int
	Retries         int
	Stdin           io.Reader
	Env             []string
}

type RunResult struct {
	Lines    []string
	ExitCode int
}

var commandExistsCache = map[string]bool{}
var commandExistsMu sync.Mutex

func commandExists(name string) bool {
	commandExistsMu.Lock()
	if v, ok := commandExistsCache[name]; ok {
		commandExistsMu.Unlock()
		return v
	}
	commandExistsMu.Unlock()

	var found bool
	if name == "code" {
		_, err := resolveVSCode()
		found = err == nil
	} else {
		_, err := exec.LookPath(name)
		found = err == nil
	}

	commandExistsMu.Lock()
	commandExistsCache[name] = found
	commandExistsMu.Unlock()
	return found
}

func Run(name string, opts RunOpts) (RunResult, error) {
	if len(opts.SuccessExitCode) == 0 {
		opts.SuccessExitCode = []int{0}
	}
	if opts.TimeoutSec <= 0 {
		opts.TimeoutSec = 1800
	}

	rc, err := resolveExe(name)
	if err != nil {
		return RunResult{}, err
	}
	allArgs := append(rc.args, opts.Args...)

	var lastErr error
	var lastResult RunResult

	for attempt := 0; attempt <= opts.Retries; attempt++ {
		if attempt > 0 {
			delay := time.Duration(attempt*2) * time.Second
			if delay > 10*time.Second {
				delay = 10 * time.Second
			}
			time.Sleep(delay)
		}

		exitCode, lines := func() (int, []string) {
			ctx, cancel := context.WithTimeout(globalCtx, time.Duration(opts.TimeoutSec)*time.Second)
			defer cancel()
			cmd := exec.CommandContext(ctx, rc.exe, allArgs...)
			cmd.Stdout = &bytes.Buffer{}
			cmd.Stderr = &bytes.Buffer{}
			if opts.Stdin != nil {
				cmd.Stdin = opts.Stdin
			}
			if len(opts.Env) > 0 {
				cmd.Env = append(os.Environ(), opts.Env...)
			}
			if name != "winget" {
				setSysProcAttr(cmd)
			}
			runErr := cmd.Run()
			code := 0
			if runErr != nil {
				if exitErr, ok := runErr.(*exec.ExitError); ok {
					code = int(int32(exitErr.ExitCode()))
				} else if ctx.Err() == context.DeadlineExceeded {
					code = 124
				} else {
					code = 1
				}
			}
			stdout := cmd.Stdout.(*bytes.Buffer).String()
			stderr := cmd.Stderr.(*bytes.Buffer).String()
			return code, splitLines(stdout + "\n" + stderr)
		}()

		lastResult = RunResult{Lines: lines, ExitCode: exitCode}

		for _, code := range opts.SuccessExitCode {
			if exitCode == code {
				return lastResult, nil
			}
		}

		lastErr = fmt.Errorf("%s failed with exit code %d", name, exitCode)
	}

	return lastResult, lastErr
}

func setSysProcAttr(cmd *exec.Cmd) {
	if runtime.GOOS == "windows" {
		setWindowsNoWindow(cmd)
	}
}
