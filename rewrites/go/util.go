package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"sync"
	"time"
)

// ---- logging ----

type Level int

const (
	LevelInfo Level = iota
	LevelSuccess
	LevelWarning
	LevelError
	LevelMuted
)

var logMu sync.Mutex
var logFile *os.File
var quietMode bool

const maxLogSizeMB = 10

func initLog(path string) error {
	if err := os.MkdirAll(filepath.Dir(path), 0755); err != nil {
		return err
	}
	if info, err := os.Stat(path); err == nil {
		if info.Size() > maxLogSizeMB*1024*1024 {
			rotatedPath := path + ".1"
			os.Remove(rotatedPath)
			os.Rename(path, rotatedPath)
		}
	}
	f, err := os.OpenFile(path, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	if err != nil {
		return err
	}
	logFile = f
	return nil
}

func closeLog() {
	if logFile != nil {
		logFile.Close()
	}
}

var levelNames = map[Level]string{
	LevelInfo:    "INFO",
	LevelSuccess: "SUCCESS",
	LevelWarning: "WARNING",
	LevelError:   "ERROR",
	LevelMuted:   "MUTED",
}

var levelColors = map[Level]string{
	LevelInfo:    "\033[36m",
	LevelSuccess: "\033[32m",
	LevelWarning: "\033[33m",
	LevelError:   "\033[31m",
	LevelMuted:   "\033[90m",
}

const colorReset = "\033[0m"

func writeLog(msg string, lvl Level) {
	logMu.Lock()
	defer logMu.Unlock()
	line := fmt.Sprintf("[%s] [%s] %s", time.Now().Format("2006-01-02 15:04:05"), levelNames[lvl], msg)
	if logFile != nil {
		fmt.Fprintln(logFile, line)
	}
	if quietMode && lvl != LevelWarning && lvl != LevelError {
		return
	}
	color := levelColors[lvl]
	fmt.Fprintf(os.Stdout, "%s%s%s\n", color, msg, colorReset)
}

func logInfo(msg string)    { writeLog(msg, LevelInfo) }
func logSuccess(msg string) { writeLog(msg, LevelSuccess) }
func logWarn(msg string)    { writeLog(msg, LevelWarning) }
func logError(msg string)   { writeLog(msg, LevelError) }
func logMuted(msg string)   { writeLog(msg, LevelMuted) }

func logTaskOutput(taskName string, lines []string) {
	if quietMode {
		return
	}
	logMu.Lock()
	defer logMu.Unlock()
	fmt.Printf("\033[90m  [%s output]\033[0m\n", taskName)
	for _, l := range lines {
		fmt.Printf("\033[90m    %s\033[0m\n", l)
	}
}

// ---- summary ----

type RunSummary struct {
	Version         string        `json:"version"`
	RunID           string        `json:"runId"`
	StartedAt       time.Time     `json:"startedAt"`
	FinishedAt      time.Time     `json:"finishedAt"`
	DurationSeconds float64       `json:"durationSeconds"`
	DryRun          bool          `json:"dryRun"`
	FastMode        bool          `json:"fastMode"`
	Throttle        int           `json:"parallelThrottle"`
	LogPath         string        `json:"logPath"`
	PlannedCount    int           `json:"plannedCount"`
	SucceededCount  int           `json:"succeededCount"`
	FailedCount     int           `json:"failedCount"`
	SkippedCount    int           `json:"skippedCount"`
	Results         []TaskResult  `json:"results"`
	Skipped         []SkippedTask `json:"skipped"`
}

func saveSummary(path string, s RunSummary) error {
	dir := filepath.Dir(path)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}
	b, err := json.MarshalIndent(s, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, b, 0644)
}

func printFinalReport(results []TaskResult, skipped []SkippedTask, elapsed time.Duration, summaryPath, logPath string) int {
	succeeded := 0
	var failed []TaskResult
	for _, r := range results {
		switch r.Status {
		case "Succeeded":
			succeeded++
		case "Failed", "TimedOut":
			failed = append(failed, r)
		}
	}

	fmt.Println()
	if len(failed) > 0 {
		logWarn(fmt.Sprintf("Completed with %d succeeded, %d failed/timed out, %d skipped in %s.",
			succeeded, len(failed), len(skipped), fmtDuration(elapsed)))
		for _, f := range failed {
			logWarn(fmt.Sprintf("  FAILED: %s — %s", f.TaskName, f.Reason))
		}
	} else {
		logSuccess(fmt.Sprintf("All runnable tasks completed: %d succeeded, %d skipped in %s.",
			succeeded, len(skipped), fmtDuration(elapsed)))
	}
	if summaryPath != "" {
		logMuted("Summary: " + summaryPath)
	}
	if logPath != "" {
		if _, err := os.Stat(logPath); err == nil {
			logMuted("Log: " + logPath)
		}
	}

	if len(failed) > 0 {
		return 1
	}
	return 0
}

func fmtDuration(d time.Duration) string {
	h := int(d.Hours())
	m := int(d.Minutes()) % 60
	s := int(d.Seconds()) % 60
	return fmt.Sprintf("%02d:%02d:%02d", h, m, s)
}

// ---- config ----

type updateConfig struct {
	WingetSkipPackages []string
	PipSkipPackages    []string
	SkipManagers       []string
}

func loadUpdateConfig() updateConfig {
	path := findUpdateConfig()
	if path == "" {
		return updateConfig{}
	}
	data, err := os.ReadFile(path)
	if err != nil {
		logWarn("Could not read update-config.json: " + err.Error())
		return updateConfig{}
	}
	var cfg updateConfig
	if err := json.Unmarshal(data, &cfg); err != nil {
		logWarn("Could not parse update-config.json: " + err.Error())
		return updateConfig{}
	}
	return cfg
}

func findUpdateConfig() string {
	dir, err := os.Getwd()
	if err != nil {
		return ""
	}
	for {
		candidate := filepath.Join(dir, "update-config.json")
		if _, err := os.Stat(candidate); err == nil {
			return candidate
		}
		parent := filepath.Dir(dir)
		if parent == dir {
			return ""
		}
		dir = parent
	}
}
