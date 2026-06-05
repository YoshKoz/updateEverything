package main

import (
	"fmt"
	"strings"
	"sync"
	"time"
)

// Task is a named update unit.
type Task struct {
	ID              string
	Name            string
	Category        string
	RequiresCommand []string
	RequiresAdmin   bool
	Disabled        bool
	DisabledReason  string
	TimeoutSec      int
	Resources       []string // mutex tokens — tasks sharing a resource run serially
	Run             func(ctx *TaskContext) error
}

// TaskContext carries per-task helpers.
type TaskContext struct {
	Output     []string
	mu         sync.Mutex
	TimeoutSec int
}

func (tc *TaskContext) Log(lines ...string) {
	tc.mu.Lock()
	tc.Output = append(tc.Output, lines...)
	tc.mu.Unlock()
}

func (tc *TaskContext) RunCmd(name string, opts RunOpts) (RunResult, error) {
	if opts.TimeoutSec <= 0 {
		opts.TimeoutSec = tc.TimeoutSec
	}
	res, err := Run(name, opts)
	tc.Log(res.Lines...)
	return res, err
}

// TaskResult records outcome.
type TaskResult struct {
	Task            *Task    `json:"-"`
	TaskName        string   `json:"name"`
	TaskCategory    string   `json:"category"`
	Status          string   `json:"status"`
	Reason          string   `json:"reason,omitempty"`
	DurationSeconds float64  `json:"durationSeconds"`
	Attempts        int      `json:"attempts"`
	Output          []string `json:"output,omitempty"`
}

// SkippedTask records why a task was not run.
type SkippedTask struct {
	Name     string
	Category string
	Reason   string
}

// Runner executes tasks in parallel up to throttle, respecting resource locks.
type Runner struct {
	throttle int
}

func NewRunner(throttle int) *Runner {
	return &Runner{throttle: throttle}
}

type inflight struct {
	task   *Task
	result chan TaskResult
	start  time.Time
}

func (r *Runner) Run(tasks []*Task) []TaskResult {
	type queueEntry struct {
		task    *Task
		attempt int
	}

	queue := make([]queueEntry, 0, len(tasks))
	for _, t := range tasks {
		queue = append(queue, queueEntry{t, 1})
	}

	var results []TaskResult
	var resultsMu sync.Mutex

	// active resource locks
	resourceLock := map[string]bool{}
	var lockMu sync.Mutex

	acquireResources := func(t *Task) bool {
		lockMu.Lock()
		defer lockMu.Unlock()
		for _, r := range t.Resources {
			if resourceLock[r] {
				return false
			}
		}
		for _, r := range t.Resources {
			resourceLock[r] = true
		}
		return true
	}

	releaseResources := func(t *Task) {
		lockMu.Lock()
		defer lockMu.Unlock()
		for _, r := range t.Resources {
			delete(resourceLock, r)
		}
	}

	var wg sync.WaitGroup
	sem := make(chan struct{}, r.throttle)
	var qMu sync.Mutex

	var dispatch func()
	dispatch = func() {
		for {
			qMu.Lock()
			if len(queue) == 0 {
				qMu.Unlock()
				return
			}
			// find first entry whose resources are free
			idx := -1
			for i, e := range queue {
				if acquireResources(e.task) {
					idx = i
					break
				}
			}
			if idx < 0 {
				qMu.Unlock()
				return
			}
			entry := queue[idx]
			queue = append(queue[:idx], queue[idx+1:]...)
			qMu.Unlock()

			sem <- struct{}{}
			wg.Add(1)
			go func(e queueEntry) {
				defer func() {
					releaseResources(e.task)
					<-sem
					wg.Done()
					dispatch() // try to start next waiting task
				}()

				logInfo(fmt.Sprintf("[%s] starting (attempt %d)", e.task.Name, e.attempt))
				start := time.Now()
				tc := &TaskContext{TimeoutSec: e.task.TimeoutSec}
				err := e.task.Run(tc)
				elapsed := time.Since(start).Seconds()

				status := "Succeeded"
				reason := ""
				if err != nil {
					status = "Failed"
					reason = err.Error()
				}

				result := TaskResult{
					Task:            e.task,
					TaskName:        e.task.Name,
					TaskCategory:    e.task.Category,
					Status:          status,
					Reason:          reason,
					DurationSeconds: elapsed,
					Attempts:        e.attempt,
					Output:          tc.Output,
				}

				resultsMu.Lock()
				results = append(results, result)
				resultsMu.Unlock()

				if status == "Succeeded" {
					logSuccess(fmt.Sprintf("[%s] ok (%.1fs)", e.task.Name, elapsed))
				} else {
					logWarn(fmt.Sprintf("[%s] %s: %s", e.task.Name, status, reason))
				}

				if len(tc.Output) > 0 {
					logTaskOutput(e.task.Name, tc.Output)
				}
			}(entry)
		}
	}

	dispatch()
	wg.Wait()
	return results
}

func makeID(name string) string {
	return strings.ToLower(strings.NewReplacer(" ", "-", "_", "-").Replace(name))
}
