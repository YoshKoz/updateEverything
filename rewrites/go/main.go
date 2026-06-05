package main

import (
	"context"
	"flag"
	"fmt"
	"math"
	"os"
	"os/signal"
	"path/filepath"
	"runtime"
	"strings"
	"syscall"
	"time"
)

const version = "1.0.0"

func main() {
	setupSignalHandler()
	// --- flags ---
	skipWindowsUpdate := flag.Bool("skip-windows-update", false, "")
	skipWSL := flag.Bool("skip-wsl", false, "")
	skipWSLDistros := flag.Bool("skip-wsl-distros", false, "")
	skipDefender := flag.Bool("skip-defender", false, "")
	skipStoreApps := flag.Bool("skip-store-apps", false, "")
	skipVSCode := flag.Bool("skip-vscode-extensions", false, "")
	skipPSModules := flag.Bool("skip-powershell-modules", false, "")
	skipNode := flag.Bool("skip-node", false, "")
	skipRust := flag.Bool("skip-rust", false, "")
	skipGo := flag.Bool("skip-go", false, "")
	skipFlutter := flag.Bool("skip-flutter", false, "")
	skipGitLFS := flag.Bool("skip-git-lfs", false, "")
	skipPoetry := flag.Bool("skip-poetry", false, "")
	skipComposer := flag.Bool("skip-composer", false, "")
	skipRuby := flag.Bool("skip-ruby", false, "")
	skipUVTools := flag.Bool("skip-uv-tools", false, "")
	skipVolta := flag.Bool("skip-volta", false, "")
	skipCleanup := flag.Bool("skip-cleanup", false, "")
	skipDestructive := flag.Bool("skip-destructive", false, "")
	deepClean := flag.Bool("deep-clean", false, "")
	fastMode := flag.Bool("fast", false, "skip slow/optional tasks")
	ultraFast := flag.Bool("ultra-fast", false, "skip everything slow")
	noParallel := flag.Bool("no-parallel", false, "")
	quietFlag := flag.Bool("quiet", false, "")
	dryRun := flag.Bool("dry-run", false, "")
	listTasks := flag.Bool("list-tasks", false, "")
	updateOllama := flag.Bool("update-ollama-models", false, "")

	throttle := flag.Int("throttle", 0, "max parallel tasks (0=auto)")
	wingetTimeout := flag.Int("winget-timeout", 600, "")
	taskTimeout := flag.Int("task-timeout", 1800, "")
	tempDays := flag.Int("temp-cleanup-days", 7, "")
	logPath := flag.String("log", "", "log file path")
	summaryPath := flag.String("summary", "", "JSON summary path")
	stateDir := flag.String("state-dir", "", "state directory")
	only := flag.String("only", "", "comma-separated task IDs to run")
	skip := flag.String("skip", "", "comma-separated task IDs to skip")

	flag.Parse()

	quietMode = *quietFlag

	// --- paths ---
	sd := *stateDir
	if sd == "" {
		localAppData := os.Getenv("LOCALAPPDATA")
		if localAppData == "" {
			localAppData = os.TempDir()
		}
		sd = filepath.Join(localAppData, "Update-Everything")
	}
	logDir := filepath.Join(sd, "logs")
	runID := time.Now().Format("20060102-150405-000")

	lp := *logPath
	if lp == "" {
		lp = filepath.Join(logDir, "update-everything-"+runID+".log")
	}
	sp := *summaryPath
	if sp == "" {
		sp = filepath.Join(sd, "last-run.json")
	}

	if err := os.MkdirAll(logDir, 0755); err != nil {
		fmt.Fprintf(os.Stderr, "warn: cannot create log dir: %v\n", err)
	}
	if err := initLog(lp); err != nil {
		fmt.Fprintf(os.Stderr, "warn: logging disabled: %v\n", err)
	}
	defer closeLog()

	// --- admin check ---
	isAdmin := isAdminProcess()
	if !isAdmin {
		logWarn("Running without Administrator. Admin-only tasks will be skipped.")
	}

	// --- throttle ---
	thr := *throttle
	if thr < 1 {
		cpus := runtime.NumCPU()
		thr = int(math.Min(float64(cpus), 6))
		if thr < 2 {
			thr = 2
		}
	}
	if *noParallel {
		thr = 1
	}

	if *ultraFast {
		*fastMode = true
	}

	config := loadUpdateConfig()

	// --- build task list ---
	allTasks := buildTasks(buildConfig{
		wingetTimeout:   *wingetTimeout,
		taskTimeout:     *taskTimeout,
		tempDays:        *tempDays,
		deepClean:       *deepClean,
		skipDestructive: *skipDestructive,
		updateOllama:    *updateOllama,
		// per-task skip flags
		skipWindowsUpdate:  *skipWindowsUpdate,
		skipWSL:            *skipWSL,
		skipWSLDistros:     *skipWSLDistros,
		skipDefender:       *skipDefender,
		skipStoreApps:      *skipStoreApps,
		skipVSCode:         *skipVSCode,
		skipPSModules:      *skipPSModules,
		skipNode:           *skipNode,
		skipRust:           *skipRust,
		skipGo:             *skipGo,
		skipFlutter:        *skipFlutter,
		skipGitLFS:         *skipGitLFS,
		skipPoetry:         *skipPoetry,
		skipComposer:       *skipComposer,
		skipRuby:           *skipRuby,
		skipUVTools:        *skipUVTools,
		skipVolta:          *skipVolta,
		skipCleanup:        *skipCleanup,
		wingetSkipPackages: config.WingetSkipPackages,
		pipSkipPackages:    config.PipSkipPackages,
	})

	// --- filter ---
	onlySet := parseCSV(*only)
	skipSet := parseCSV(*skip)
	skipSet = append(skipSet, config.SkipManagers...)
	if *fastMode {
		skipSet = append(skipSet, fastModeSkip...)
	}
	if *ultraFast {
		skipSet = append(skipSet, ultraFastSkip...)
	}

	var planned []*Task
	var skipped []SkippedTask

	for _, t := range allTasks {
		reason := filterReason(t, isAdmin, onlySet, skipSet)
		if reason != "" {
			skipped = append(skipped, SkippedTask{Name: t.Name, Category: t.Category, Reason: reason})
		} else {
			planned = append(planned, t)
		}
	}

	start := time.Now()
	logInfo(fmt.Sprintf("Update-Everything v%s | %s | throttle %d", version, start.Format("2006-01-02 15:04"), thr))

	if *listTasks {
		fmt.Println("\nPlanned tasks:")
		for _, t := range planned {
			fmt.Printf("  %-30s %s\n", t.Name, t.Category)
		}
		fmt.Println("\nSkipped tasks:")
		for _, s := range skipped {
			fmt.Printf("  %-30s %s\n", s.Name, s.Reason)
		}
		return
	}

	if *dryRun {
		logInfo("Dry run: no update commands will be executed.")
		for _, t := range planned {
			logMuted(fmt.Sprintf("[DryRun] %s (%s)", t.Name, t.Category))
		}
		return
	}

	if len(planned) == 0 {
		logWarn("No runnable update tasks found.")
		return
	}

	logInfo(fmt.Sprintf("Dispatching %d task(s). Skipped: %d", len(planned), len(skipped)))

	runner := NewRunner(thr)
	results := runner.Run(planned)

	elapsed := time.Since(start)

	succeeded := 0
	failed := 0
	for _, r := range results {
		switch r.Status {
		case "Succeeded":
			succeeded++
		case "Failed", "TimedOut":
			failed++
		}
	}

	summary := RunSummary{
		Version:         version,
		RunID:           runID,
		StartedAt:       start,
		FinishedAt:      time.Now(),
		DurationSeconds: elapsed.Seconds(),
		DryRun:          *dryRun,
		FastMode:        *fastMode,
		Throttle:        thr,
		LogPath:         lp,
		PlannedCount:    len(planned),
		SucceededCount:  succeeded,
		FailedCount:     failed,
		SkippedCount:    len(skipped),
		Results:         results,
		Skipped:         skipped,
	}
	if err := saveSummary(sp, summary); err != nil {
		logWarn("Could not write summary: " + err.Error())
	}

	exitCode := printFinalReport(results, skipped, elapsed, sp, lp)
	os.Exit(exitCode)
}

// buildConfig holds all construction-time settings for tasks.
type buildConfig struct {
	wingetTimeout      int
	taskTimeout        int
	tempDays           int
	deepClean          bool
	skipDestructive    bool
	updateOllama       bool
	skipWindowsUpdate  bool
	skipWSL            bool
	skipWSLDistros     bool
	skipDefender       bool
	skipStoreApps      bool
	skipVSCode         bool
	skipPSModules      bool
	skipNode           bool
	skipRust           bool
	skipGo             bool
	skipFlutter        bool
	skipGitLFS         bool
	skipPoetry         bool
	skipComposer       bool
	skipRuby           bool
	skipUVTools        bool
	skipVolta          bool
	skipCleanup        bool
	wingetSkipPackages []string
	pipSkipPackages    []string
}

var fastModeSkip = []string{
	"chocolatey", "wsl-distros", "npm", "pnpm", "yarn", "bun", "deno",
	"rustup", "cargo", "go", "pip", "pipx", "uv", "uv-tools",
	"poetry", "composer", "ruby-gems", "flutter", "dotnet-tools",
	"dotnet-workloads", "vscode-extensions", "powershell-modules",
	"powershell-help", "ollama-models", "volta",
}

var ultraFastSkip = []string{
	"windows-update", "store-apps", "wsl", "wsl-distros", "defender", "cleanup",
}

func buildTasks(cfg buildConfig) []*Task {
	tasks := []*Task{
		taskWinget(cfg.wingetTimeout, cfg.wingetSkipPackages),
		taskScoop(),
		taskChocolatey(),
		taskStoreApps(cfg.wingetTimeout),
		taskWindowsUpdate(),
		taskDefender(),
		taskWSL(),
		taskWSLDistros(),
		taskNPM(nil),
		taskPNPM(),
		taskYarn(),
		taskBun(),
		taskDeno(),
		taskMise(),
		taskVolta(),
		taskPip(cfg.pipSkipPackages),
		taskPipx(),
		taskUV(),
		taskUVTools(),
		taskPoetry(),
		taskRustup(),
		taskCargo(),
		taskGo(),
		taskFlutter(),
		taskDotnetTools(),
		taskDotnetWorkloads(),
		taskRubyGems(),
		taskComposer(),
		taskVSCodeExtensions(),
		taskGitLFS(),
		taskGHExtensions(),
		taskPSModules(),
		taskPSHelp(),
		taskOllamaModels(600),
		taskCleanup(cfg.tempDays, cfg.deepClean, cfg.skipDestructive),
	}

	// apply flag-based disabled state
	disabled := map[string]string{}
	if cfg.skipWindowsUpdate {
		disabled["windows-update"] = "-skip-windows-update"
	}
	if cfg.skipWSL || cfg.skipWSLDistros {
		disabled["wsl-distros"] = "-skip-wsl / -skip-wsl-distros"
	}
	if cfg.skipWSL {
		disabled["wsl"] = "-skip-wsl"
	}
	if cfg.skipDefender {
		disabled["defender"] = "-skip-defender"
	}
	if cfg.skipStoreApps {
		disabled["store-apps"] = "-skip-store-apps"
	}
	if cfg.skipVSCode {
		disabled["vscode-extensions"] = "-skip-vscode-extensions"
	}
	if cfg.skipPSModules {
		disabled["powershell-modules"] = "-skip-powershell-modules"
	}
	if cfg.skipNode {
		for _, id := range []string{"npm", "pnpm", "yarn", "bun", "deno"} {
			disabled[id] = "-skip-node"
		}
	}
	if cfg.skipRust {
		disabled["rustup"] = "-skip-rust"
		disabled["cargo"] = "-skip-rust"
	}
	if cfg.skipGo {
		disabled["go"] = "-skip-go"
	}
	if cfg.skipFlutter {
		disabled["flutter"] = "-skip-flutter"
	}
	if cfg.skipGitLFS {
		disabled["git-lfs"] = "-skip-git-lfs"
	}
	if cfg.skipPoetry {
		disabled["poetry"] = "-skip-poetry"
	}
	if cfg.skipComposer {
		disabled["composer"] = "-skip-composer"
	}
	if cfg.skipRuby {
		disabled["ruby-gems"] = "-skip-ruby"
	}
	if cfg.skipUVTools {
		disabled["uv-tools"] = "-skip-uv-tools"
	}
	if cfg.skipVolta {
		disabled["volta"] = "-skip-volta"
	}
	if cfg.skipCleanup {
		disabled["cleanup"] = "-skip-cleanup"
	}
	if !cfg.updateOllama {
		disabled["ollama-models"] = "use -update-ollama-models to refresh local models"
	}

	for _, t := range tasks {
		if reason, ok := disabled[t.ID]; ok {
			t.Disabled = true
			t.DisabledReason = reason
		}
	}
	return tasks
}

func filterReason(t *Task, isAdmin bool, only, skip []string) string {
	if len(only) > 0 && !contains(only, t.ID) && !contains(only, t.Category) {
		return "not in -only list"
	}
	if contains(skip, t.ID) || contains(skip, t.Category) {
		return "skipped by filter"
	}
	if t.Disabled {
		if t.DisabledReason != "" {
			return t.DisabledReason
		}
		return "disabled"
	}
	if t.RequiresAdmin && !isAdmin {
		return "requires Administrator"
	}
	for _, cmd := range t.RequiresCommand {
		if !commandExists(cmd) {
			return "missing command: " + cmd
		}
	}
	return ""
}

func parseCSV(s string) []string {
	if s == "" {
		return nil
	}
	parts := strings.Split(s, ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			out = append(out, p)
		}
	}
	return out
}

func contains(list []string, s string) bool {
	for _, v := range list {
		if strings.EqualFold(v, s) {
			return true
		}
	}
	return false
}

var globalCancel context.CancelFunc

func setupSignalHandler() {
	var ctx context.Context
	ctx, globalCancel = context.WithCancel(context.Background())
	_ = ctx
	sigCh := make(chan os.Signal, 1)
	signal.Notify(sigCh, os.Interrupt, syscall.SIGTERM)
	go func() {
		<-sigCh
		logWarn("Received interrupt signal, shutting down gracefully...")
		globalCancel()
	}()
}
