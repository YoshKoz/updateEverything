use std::{
    collections::{BTreeMap, BTreeSet, HashMap, HashSet, VecDeque},
    env, fs,
    io::{BufRead, BufReader},
    path::{Path, PathBuf},
    process::{Command, ExitStatus, Stdio},
    sync::{Arc, Condvar, Mutex, mpsc},
    thread,
    time::{Duration, Instant},
};

use anyhow::{Context, Result};
use clap::Parser;
use serde::{Deserialize, Serialize};
use time::{OffsetDateTime, macros::format_description};

#[derive(Debug, Parser)]
#[command(version, about = "Rust updateEverything runner")]
struct Cli {
    #[arg(long)]
    dry_run: bool,

    #[arg(long)]
    list_tasks: bool,

    #[arg(long)]
    fast: bool,

    #[arg(long)]
    ultra_fast: bool,

    #[arg(long, value_delimiter = ',')]
    only: Vec<String>,

    #[arg(long, value_delimiter = ',')]
    skip: Vec<String>,

    #[arg(long)]
    config: Option<PathBuf>,

    #[arg(long)]
    json_summary: Option<PathBuf>,

    #[arg(long)]
    quiet: bool,

    #[arg(long, default_value_t = 1800)]
    task_timeout_sec: u64,

    #[arg(long, default_value_t = 1)]
    jobs: usize,

    #[arg(long)]
    ci: bool,

    #[arg(long, default_value_t = 0.0)]
    since_hours: f64,

    #[arg(long)]
    state_dir: Option<PathBuf>,

    // --- skip flags (mirror PS1 params) ---
    #[arg(long)]
    skip_windows_update: bool,

    #[arg(long)]
    skip_wsl: bool,

    #[arg(long)]
    skip_wsl_distros: bool,

    #[arg(long)]
    skip_defender: bool,

    #[arg(long)]
    skip_store_apps: bool,

    #[arg(long)]
    skip_powershell_modules: bool,

    #[arg(long)]
    skip_node: bool,

    #[arg(long)]
    skip_rust: bool,

    #[arg(long)]
    skip_go: bool,

    #[arg(long)]
    skip_flutter: bool,

    #[arg(long)]
    skip_ruby: bool,

    #[arg(long)]
    skip_composer: bool,

    #[arg(long)]
    skip_poetry: bool,

    #[arg(long)]
    skip_uv_tools: bool,

    #[arg(long)]
    skip_cleanup: bool,

    #[arg(long)]
    deep_clean: bool,

    #[arg(long)]
    skip_destructive: bool,

    #[arg(long)]
    skip_git_lfs: bool,

    #[arg(long)]
    skip_vcpkg: bool,

    #[arg(long)]
    skip_conda: bool,

    #[arg(long)]
    skip_cloud_tools: bool,

    #[arg(long)]
    skip_infra_tools: bool,

    #[arg(long)]
    skip_k8s_tools: bool,

    #[arg(long)]
    skip_starship: bool,

    #[arg(long)]
    skip_hugo: bool,

    #[arg(long)]
    skip_vscode_extensions: bool,

    #[arg(long)]
    skip_pip_health: bool,

    /// Pull/upgrade all local Ollama models (opt-in, slow)
    #[arg(long)]
    update_ollama_models: bool,

    /// Run Update-Help for PowerShell (opt-in)
    #[arg(long)]
    update_powershell_help: bool,

    /// Download latest GitHub-release binaries for tools in config.GithubTools (opt-in)
    #[arg(long)]
    update_github_tools: bool,

    /// Include apps protected by package managers in upgrades
    #[arg(long)]
    bypass_protection: bool,

    /// Run extra cleanup steps (DISM, delivery-opt cache, orphan scan)
    #[arg(long = "deep-clean")]
    // already set above — alias for clarity in help
    #[arg(long, default_value_t = 600)]
    winget_timeout_sec: u64,

    #[arg(long, default_value_t = 600)]
    ollama_timeout_sec: u64,

    #[arg(long, default_value_t = 1)]
    retry_count: u32,

    /// Work/personal/gaming/minimal preset
    #[arg(long)]
    profile: Option<String>,

    #[arg(long)]
    show_skipped: bool,

    /// Register a Windows scheduled task to run at logon (use --schedule-time HH:MM for daily instead)
    #[arg(long)]
    schedule: bool,

    /// When scheduling: trigger daily at HH:MM instead of at logon
    #[arg(long)]
    schedule_time: Option<String>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "PascalCase")]
#[allow(dead_code)]
struct Config {
    #[serde(default)]
    winget_skip_packages: Vec<String>,
    #[serde(default)]
    skip_managers: Vec<String>,
    #[serde(default)]
    pip_skip_packages: Vec<String>,
    #[serde(default)]
    pip_ignore_health_packages: Vec<String>,
    #[serde(default)]
    npm_skip_packages: Vec<String>,
    #[serde(default)]
    chocolatey_skip_packages: Vec<String>,
    #[serde(default)]
    store_app_skip_packages: Vec<String>,
    #[serde(default)]
    gcloud_skip_components: Vec<String>,
    #[serde(default)]
    az_skip_extensions: Vec<String>,
    #[serde(default)]
    conda_skip_envs: Vec<String>,
    #[serde(default)]
    vcpkg_skip_packages: Vec<String>,
    #[serde(default)]
    windows_optional_features: Vec<String>,
    #[serde(default)]
    cross_manager_fallback: BTreeMap<String, BTreeMap<String, String>>,
    #[serde(default = "default_cleanup_days")]
    temp_cleanup_days: u32,
    #[serde(default = "default_log_retention")]
    log_retention_days: u32,
    #[serde(default)]
    github_tools: Vec<GithubTool>,
}

/// A manually-installed tool whose binaries come from a release/CI asset (no
/// package manager, no local git build). The task queries the latest version,
/// compares against the locally-installed version, and downloads + extracts the
/// matching asset when behind. Supports three providers:
///   - "github"          (default): GitHub releases/latest, asset by AssetRegex
///   - "gitlab"          : GitLab releases permalink/latest, asset link by AssetRegex
///   - "gitlab-artifact" : latest successful CI pipeline artifact for a Job
#[derive(Clone, Debug, Deserialize)]
#[serde(rename_all = "PascalCase")]
struct GithubTool {
    /// "owner/name", e.g. "ggml-org/llama.cpp" (also used for task id + marker)
    repo: String,
    /// Install directory to extract/copy into
    install_dir: String,
    /// Regex matched against release asset names to pick the download
    /// (unused for gitlab-artifact)
    #[serde(default)]
    asset_regex: String,
    /// "github" (default), "gitlab", or "gitlab-artifact"
    #[serde(default)]
    provider: Option<String>,
    /// Numeric GitLab project id (required for gitlab / gitlab-artifact)
    #[serde(default)]
    project_id: Option<String>,
    /// Git ref for gitlab-artifact pipelines (default "master")
    #[serde(default, rename = "Ref")]
    git_ref: Option<String>,
    /// CI job name for gitlab-artifact (e.g. "Windows 64")
    #[serde(default)]
    job: Option<String>,
    /// Optional command run inside install_dir to print the local version
    #[serde(default)]
    version_cmd: Option<String>,
    /// Regex with one capture group extracting a numeric local version
    #[serde(default)]
    version_regex: Option<String>,
    /// Optional task id override (defaults to "gh-<repo name>")
    #[serde(default)]
    id: Option<String>,
}

fn default_cleanup_days() -> u32 {
    7
}
fn default_log_retention() -> u32 {
    14
}

#[derive(Clone, Debug)]
struct Task {
    id: &'static str,
    category: &'static str,
    tags: &'static [&'static str],
    command: &'static str,
    args: Vec<String>,
    requires: &'static str,
    resource: &'static str,
    /// Task ids that must finish (any status) before this one starts. Resources only
    /// guarantee mutual exclusion, not ordering, so tasks sharing a resource that must
    /// run in a specific sequence (e.g. pinning before `winget upgrade`) need this too.
    depends_on: &'static [&'static str],
    skip_reason: Option<String>,
    timeout_override: Option<Duration>,
    acceptable_exit_codes: Vec<i32>,
    ok_on_timeout: bool,
}

#[derive(Debug, Serialize)]
struct TaskSummary {
    id: String,
    category: String,
    status: String,
    duration_ms: u128,
    exit_code: Option<i32>,
    command: String,
    args: Vec<String>,
    output_tail: Vec<String>,
}

#[derive(Debug, Serialize)]
struct RunSummary {
    version: String,
    started_at: String,
    duration_ms: u128,
    dry_run: bool,
    results: Vec<TaskSummary>,
}

#[derive(Debug, Deserialize)]
struct PrevResult {
    #[serde(rename = "Id", alias = "id")]
    id: String,
    #[serde(rename = "Status", alias = "status")]
    status: String,
}

#[derive(Debug, Deserialize)]
struct PrevSummary {
    #[serde(rename = "Results", alias = "results")]
    results: Vec<PrevResult>,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Handle --schedule before anything else
    if cli.schedule {
        let exe = std::env::current_exe().context("could not resolve current executable")?;
        schedule_task(&exe, cli.schedule_time.as_deref())?;
        return Ok(());
    }

    let start = Instant::now();
    let started_at = now_string();
    let repo_root = find_repo_root()?;
    let state_dir = get_state_dir(&cli);
    let config_path = cli
        .config
        .clone()
        .unwrap_or_else(|| repo_root.join("update-config.json"));
    let config = load_config(&config_path)?;

    // Prevent concurrent runs (non-fatal: if lock fails we still continue)
    let _lock = ProcessLock::acquire(&state_dir);

    let tasks = build_tasks(&config, &cli);
    let prev_summary = if cli.since_hours > 0.0 {
        load_prev_summary(&cli, &repo_root)
    } else {
        None
    };
    let selected = filter_tasks(tasks, &cli, &config, prev_summary.as_ref());

    if cli.list_tasks {
        print_task_list(&selected);
        return Ok(());
    }

    let parallel = cli.jobs.max(1);
    let timeout = Duration::from_secs(cli.task_timeout_sec);
    let results = run_tasks(selected, &cli, parallel, timeout);

    let summary = RunSummary {
        version: env!("CARGO_PKG_VERSION").to_string(),
        started_at,
        duration_ms: start.elapsed().as_millis(),
        dry_run: cli.dry_run,
        results,
    };

    // Auto-save last-run.json to state dir (enables --since-hours on next run)
    if !cli.dry_run {
        let auto_path = state_dir.join("last-run.json");
        if let Err(e) = write_summary(&auto_path, &summary) {
            eprintln!("warn: could not save {}: {e}", auto_path.display());
        }
    }

    // Also save to explicit --json-summary path if given
    if let Some(path) = &cli.json_summary {
        write_summary(path, &summary)?;
        if !cli.quiet {
            println!("summary {}", path.display());
        }
    }

    print_summary(&summary);

    if !cli.dry_run {
        print_update_summary(&summary.results);
    }

    if cli.ci {
        let any_failure = summary
            .results
            .iter()
            .any(|r| matches!(r.status.as_str(), "Failed" | "TimedOut"));
        if any_failure {
            std::process::exit(1);
        }
    }

    Ok(())
}

fn run_tasks(tasks: Vec<Task>, cli: &Cli, jobs: usize, timeout: Duration) -> Vec<TaskSummary> {
    let mut results: Vec<TaskSummary> = Vec::with_capacity(tasks.len());

    if cli.dry_run {
        for task in &tasks {
            if let Some(reason) = &task.skip_reason {
                if !cli.quiet {
                    println!("skip {:<22} {}", task.id, reason);
                }
                results.push(make_summary(task, "Skipped", 0, None, vec![]));
            } else {
                println!(
                    "dry  {:<22} {} {}",
                    task.id,
                    task.command,
                    shell_join_brief(&task.args)
                );
                results.push(make_summary(task, "DryRun", 0, None, vec![]));
            }
        }
        return results;
    }

    let (skipped, to_run): (Vec<_>, Vec<_>) =
        tasks.into_iter().partition(|t| t.skip_reason.is_some());

    for task in &skipped {
        let reason = task.skip_reason.as_deref().unwrap_or("");
        if !cli.quiet {
            println!("skip {:<22} {}", task.id, reason);
        }
        results.push(make_summary(task, "Skipped", 0, None, vec![]));
    }

    let retry_count = cli.retry_count;

    if jobs <= 1 {
        for task in to_run {
            let mut r = run_task_streaming(&task, cli.quiet, timeout, false, &HashMap::new());
            for attempt in 1..=retry_count {
                if !matches!(r.status.as_str(), "Failed" | "TimedOut") {
                    break;
                }
                let delay = Duration::from_secs(3 * (1u64 << attempt.min(4)));
                eprintln!("retry {attempt}/{retry_count} {} (waiting {}s)", task.id, delay.as_secs());
                thread::sleep(delay);
                r = run_task_streaming(&task, cli.quiet, timeout, false, &HashMap::new());
            }
            results.push(r);
        }
    } else {
        let mut resource_locks: HashMap<String, Arc<Mutex<()>>> = HashMap::new();
        for task in &to_run {
            if !task.resource.is_empty() {
                resource_locks
                    .entry(task.resource.to_string())
                    .or_insert_with(|| Arc::new(Mutex::new(())));
            }
        }
        let resource_locks = Arc::new(resource_locks);

        // Tasks may declare `depends_on` on other task ids that share a resource but
        // need a strict run-before/run-after order (a resource mutex only guarantees
        // exclusion, not sequencing). Only wait on dependencies that actually appear
        // in this run's task set, so filtering (--only/--skip) can't deadlock a task
        // waiting on a dependency that was never scheduled.
        let runnable_ids: HashSet<String> = to_run.iter().map(|t| t.id.to_string()).collect();
        let finished: Arc<(Mutex<HashSet<String>>, Condvar)> =
            Arc::new((Mutex::new(HashSet::new()), Condvar::new()));

        let queue: Arc<Mutex<VecDeque<Task>>> = Arc::new(Mutex::new(to_run.into_iter().collect()));
        let out: Arc<Mutex<Vec<TaskSummary>>> = Arc::new(Mutex::new(Vec::new()));

        let mut handles = vec![];
        for _ in 0..jobs {
            let queue = Arc::clone(&queue);
            let out = Arc::clone(&out);
            let locks = Arc::clone(&resource_locks);
            let finished = Arc::clone(&finished);
            let runnable_ids = runnable_ids.clone();
            let quiet = cli.quiet;
            let h = thread::spawn(move || {
                loop {
                    let task = {
                        let mut q = queue.lock().unwrap();
                        // Find the first queued task whose dependencies (that were
                        // actually scheduled this run) have all finished; requeue
                        // anything skipped over so other workers can still find it.
                        let (done_lock, _) = &*finished;
                        let pos = {
                            let done = done_lock.lock().unwrap();
                            q.iter().position(|t| {
                                t.depends_on
                                    .iter()
                                    .all(|dep| !runnable_ids.contains(*dep) || done.contains(*dep))
                            })
                        };
                        match pos {
                            Some(i) => q.remove(i).unwrap(),
                            None if q.is_empty() => break,
                            None => {
                                // Everything left is blocked on an in-flight dependency;
                                // wait to be woken when a task finishes, then re-check.
                                let guard = q;
                                drop(
                                    finished
                                        .1
                                        .wait_timeout(guard, Duration::from_millis(200))
                                        .unwrap(),
                                );
                                continue;
                            }
                        }
                    };
                    let mut r = run_task_streaming(&task, quiet, timeout, true, &locks);
                    for attempt in 1..=retry_count {
                        if !matches!(r.status.as_str(), "Failed" | "TimedOut") {
                            break;
                        }
                        let delay = Duration::from_secs(3 * (1u64 << attempt.min(4)));
                        eprintln!("retry {attempt}/{retry_count} {} (waiting {}s)", task.id, delay.as_secs());
                        thread::sleep(delay);
                        r = run_task_streaming(&task, quiet, timeout, true, &locks);
                    }
                    let task_id = task.id.to_string();
                    out.lock().unwrap().push(r);
                    finished.0.lock().unwrap().insert(task_id);
                    finished.1.notify_all();
                }
            });
            handles.push(h);
        }
        for h in handles {
            let _ = h.join();
        }
        let parallel_results = Arc::try_unwrap(out).unwrap().into_inner().unwrap();
        results.extend(parallel_results);
    }

    results
}

/// Task scripts print a line starting with this when they deliberately took no
/// action. Without it an intentional no-op reports as Succeeded, which reads as
/// "updated" in the summary.
const SKIP_PREFIX: &str = "SKIPPED:";

/// Prepended to python task scripts that need to free a locked executable.
/// Matches on the exact resolved path so an unrelated same-named binary
/// elsewhere on PATH is never touched.
const CLOSE_BLOCKERS_PY: &str = r#"
import os as _os, subprocess as _sp

def close_locking_processes(bindir, names):
    targets = [_os.path.join(bindir, n) for n in names]
    quoted = ",".join("'" + t.replace("'", "''") + "'" for t in targets)
    ps = (
        "$t=@(" + quoted + ");"
        "Get-Process -ErrorAction SilentlyContinue |"
        " Where-Object { $t -contains $_.Path } |"
        " ForEach-Object {"
        "  Write-Output ('closed ' + $_.ProcessName + ' (pid ' + $_.Id + ')');"
        "  Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }"
    )
    r = _sp.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
        capture_output=True, text=True,
    )
    out = (r.stdout or "").strip()
    if out:
        print(out)
    return bool(out)
"#;

fn run_task_streaming(
    task: &Task,
    quiet: bool,
    timeout: Duration,
    prefix_output: bool,
    resource_locks: &HashMap<String, Arc<Mutex<()>>>,
) -> TaskSummary {
    if !quiet {
        println!(
            "run  {:<22} {} {}",
            task.id,
            task.command,
            shell_join_brief(&task.args)
        );
    }

    let _resource_guard = if !task.resource.is_empty() {
        resource_locks.get(task.resource).map(|m| m.lock().unwrap())
    } else {
        None
    };

    let timeout = task.timeout_override.unwrap_or(timeout);
    let start = Instant::now();
    let mut command = new_task_command(task.command);
    command
        .args(&task.args)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    #[cfg(unix)]
    {
        use std::os::unix::process::CommandExt;
        command.process_group(0);
    }
    let mut child = match command.spawn() {
        Ok(child) => child,
        Err(err) => {
            eprintln!("fail {:<22} spawn failed: {err}", task.id);
            return make_summary(
                task,
                "Failed",
                start.elapsed().as_millis(),
                None,
                vec![format!("spawn failed: {err}")],
            );
        }
    };

    let stdout = child.stdout.take().unwrap();
    let stderr = child.stderr.take().unwrap();
    let task_id = task.id.to_string();

    let lines_out: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));
    let lines_err: Arc<Mutex<Vec<String>>> = Arc::new(Mutex::new(Vec::new()));

    let lo = Arc::clone(&lines_out);
    let tid_out = task_id.clone();
    let quiet_out = quiet;
    let prefix_out = prefix_output;
    let stdout_handle = thread::spawn(move || {
        let reader = BufReader::new(stdout);
        for line in reader.lines().map_while(|l| l.ok()) {
            if !quiet_out {
                if prefix_out {
                    println!("[{tid_out}] {line}");
                } else {
                    println!("  {line}");
                }
            }
            lo.lock().unwrap().push(line);
        }
    });

    let le = Arc::clone(&lines_err);
    let tid_err = task_id.clone();
    let quiet_err = quiet;
    let prefix_err = prefix_output;
    let stderr_handle = thread::spawn(move || {
        let reader = BufReader::new(stderr);
        for line in reader.lines().map_while(|l| l.ok()) {
            if !quiet_err {
                if prefix_err {
                    eprintln!("[{tid_err}] {line}");
                } else {
                    eprintln!("  {line}");
                }
            }
            le.lock().unwrap().push(line);
        }
    });

    let mut timed_out = false;
    loop {
        match child.try_wait() {
            Ok(Some(_)) => break,
            Ok(None) if start.elapsed() >= timeout => {
                timed_out = true;
                kill_task(&mut child);
                break;
            }
            Ok(None) => thread::sleep(Duration::from_millis(200)),
            Err(err) => {
                return make_summary(
                    task,
                    "Failed",
                    start.elapsed().as_millis(),
                    None,
                    vec![err.to_string()],
                );
            }
        }
    }

    let exit_status = child.wait().ok();
    let duration_ms = start.elapsed().as_millis();
    if !timed_out {
        let (tx, rx) = mpsc::sync_channel::<()>(2);
        let tx2 = tx.clone();
        thread::spawn(move || {
            let _ = stdout_handle.join();
            let _ = tx.send(());
        });
        thread::spawn(move || {
            let _ = stderr_handle.join();
            let _ = tx2.send(());
        });
        let drain_deadline = Instant::now() + Duration::from_secs(3);
        for _ in 0..2 {
            let remaining = drain_deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                break;
            }
            let _ = rx.recv_timeout(remaining);
        }
    }
    let mut all_lines = lines_out.lock().unwrap().clone();
    all_lines.extend(lines_err.lock().unwrap().iter().cloned());

    let code = exit_status.and_then(|s| s.code());
    let status = if timed_out && task.ok_on_timeout {
        "Succeeded".to_string()
    } else if timed_out {
        "TimedOut".to_string()
    } else if code.is_some_and(|c| task.acceptable_exit_codes.contains(&c)) {
        "Succeeded".to_string()
    } else {
        exit_status
            .map(status_name)
            .unwrap_or_else(|| "Failed".to_string())
    };
    let status = if status == "Succeeded"
        && all_lines
            .iter()
            .any(|l| l.trim_start().starts_with(SKIP_PREFIX))
    {
        "Skipped".to_string()
    } else {
        status
    };

    if !quiet {
        println!(
            "done {:<22} {} ({:.1}s)",
            task.id,
            status,
            duration_ms as f64 / 1000.0
        );
    }

    // Normalize whitelisted exit codes to 0 in the summary so the report column
    // doesn't show an alarming raw code (e.g. winget -1978335189 "no applicable
    // upgrade") for a task that succeeded.
    let reported_code = match code {
        Some(c) if task.acceptable_exit_codes.contains(&c) => Some(0),
        other => other,
    };
    make_summary(task, &status, duration_ms, reported_code, cap_output(all_lines, 300))
}

#[cfg(unix)]
fn kill_task(child: &mut std::process::Child) {
    unsafe {
        libc::kill(-(child.id() as i32), libc::SIGKILL);
    }
    let _ = child.kill();
}

#[cfg(windows)]
fn kill_task(child: &mut std::process::Child) {
    let _ = child.kill();
}

fn make_summary(
    task: &Task,
    status: &str,
    duration_ms: u128,
    exit_code: Option<i32>,
    output_tail: Vec<String>,
) -> TaskSummary {
    TaskSummary {
        id: task.id.to_string(),
        category: task.category.to_string(),
        status: status.to_string(),
        duration_ms,
        exit_code,
        command: task.command.to_string(),
        args: task.args.clone(),
        output_tail,
    }
}

// ─── Task list ───────────────────────────────────────────────────────────────

fn build_tasks(config: &Config, cli: &Cli) -> Vec<Task> {
    let mut pip_skip = config.pip_skip_packages.clone();
    for pkg in &config.pip_ignore_health_packages {
        if !pip_skip.contains(pkg) {
            pip_skip.push(pkg.clone());
        }
    }

    let skip_node = cli.skip_node;
    let skip_rust = cli.skip_rust;
    let skip_go = cli.skip_go;

    let winget_upgrade_args = winget_upgrade_args(&config.winget_skip_packages);

    let mut tasks = vec![
        // ── winget ──────────────────────────────────────────────────────────
        Task::new(
            "winget-source",
            "package-manager",
            &["windows", "winget"],
            "winget",
            &["source", "update"],
        )
        .with_resource("winget")
        .with_timeout(300),
        // Kill portable-app processes that hold their own exe before winget upgrade
        Task::new(
            "winget-pre",
            "package-manager",
            &["windows", "winget"],
            "cmd",
            &["/c", "taskkill /F /IM codex-x86_64-pc-windows-msvc.exe 2>nul & exit 0"],
        )
        .with_resource("winget"),
        // Git upgrades abort while any Git-shipped exe is running; Claude Code's
        // statusline respawns bash.exe every few seconds, so kill-loops and
        // /FORCECLOSEAPPLICATIONS both lose the race. Rename bash.exe away so
        // respawns fail harmlessly, upgrade, then clean up.
        Task::new_vec(
            "winget-git",
            "package-manager",
            &["windows", "winget", "git"],
            "pwsh",
            pwsh_cmd(winget_git_script()),
        )
        .with_resource("winget")
        .with_timeout(900)
        .with_requires("winget"),
        // Pin the skip-packages so the upgrade passes leave them alone. winget
        // has no per-package exclude for `upgrade --all`, so pinning is the only
        // reliable mechanism. `pin add` on an already-pinned id is a no-op.
        Task::new_vec(
            "winget-pin-skip",
            "package-manager",
            &["windows", "winget"],
            "pwsh",
            pwsh_cmd(&winget_pin_skip_script(&config.winget_skip_packages)),
        )
        .with_resource("winget")
        .with_timeout(300),
        Task::new_vec(
            "winget",
            "package-manager",
            &["windows", "winget"],
            "pwsh",
            winget_upgrade_args,
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212])
        .with_depends_on(&["winget-pin-skip", "winget-git"]),
        Task::new_vec(
            "winget-batch",
            "package-manager",
            &["windows", "winget"],
            "pwsh",
            // Second pass over whatever the first pass could not clear; `--all` is unusable
            // here for the same reason (see winget_per_package_script).
            pwsh_cmd(&winget_per_package_script(&config.winget_skip_packages)),
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212])
        .with_depends_on(&["winget-pin-skip", "winget-git"]),
        // Elevated winget refuses to upgrade user-scope (zip/portable) packages
        // ("cannot be uninstalled when running with administrator privileges",
        // e.g. charmbracelet.crush). Re-run the upgrade de-elevated via
        // `runas /trustlevel:0x20000` to catch those.
        Task::new_vec(
            "winget-userscope",
            "package-manager",
            &["windows", "winget"],
            "pwsh",
            pwsh_cmd(&winget_userscope_script(&config.winget_skip_packages)),
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_requires("winget")
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212]),
        Task::new_vec(
            "winget-pin-audit",
            "package-manager",
            &["windows", "winget"],
            "winget",
            vec!["pin".into(), "list".into()],
        )
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212]),
        Task::new_vec(
            "cross-manager",
            "package-manager",
            &["windows", "winget", "scoop", "choco"],
            "python",
            cross_manager_args(&config.cross_manager_fallback),
        )
        .with_resource("package-manager"),
        // ── store apps ───────────────────────────────────────────────────────
        Task::new_vec(
            "store-apps",
            "system",
            &["windows", "store"],
            "winget",
            {
                let mut a = vec![
                    "upgrade".into(), "--all".into(), "--source".into(), "msstore".into(),
                    "--include-unknown".into(), "--include-pinned".into(),
                    "--accept-source-agreements".into(), "--disable-interactivity".into(),
                ];
                for pkg in &config.store_app_skip_packages {
                    a.push("--exclude".into());
                    a.push(pkg.clone());
                }
                a
            },
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212])
        .with_requires("winget")
        .with_skip_if(cli.skip_store_apps, "disabled by --skip-store-apps"),
        // ── scoop ────────────────────────────────────────────────────────────
        // scoop is a .ps1 shim — must be invoked via pwsh
        Task::new(
            "scoop",
            "package-manager",
            &["windows", "scoop"],
            "pwsh",
            &[
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "scoop update; if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }; scoop update *",
            ],
        )
        .with_requires("scoop"),
        // ── chocolatey ───────────────────────────────────────────────────────
        Task::new_vec(
            "chocolatey",
            "package-manager",
            &["windows", "choco"],
            "pwsh",
            {
                // Build the `--except pkg` suffix, then wrap the whole choco call
                // in a pwsh admin-guard. choco requires elevation; when not
                // elevated it prints a warning and blocks ~20s on a "continue?"
                // prompt. Skip cleanly instead (the scheduled task runs elevated).
                let mut except = String::new();
                for pkg in &config.chocolatey_skip_packages {
                    except.push_str(&format!(" --except '{}'", pkg.replace('\'', "''")));
                }
                let script = format!(
                    "if(-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){{Write-Output 'Chocolatey skipped: requires elevation (run elevated to upgrade choco packages).';exit 0}}; \
                     & choco upgrade all -y --no-progress{except}; exit $LASTEXITCODE"
                );
                pwsh_cmd(&script)
            },
        )
        .with_requires("choco"),
        // ── Windows Update ───────────────────────────────────────────────────
        Task::new_vec(
            "windows-update",
            "system",
            &["windows"],
            "pwsh",
            pwsh_cmd(windows_update_script()),
        )
        .with_resource("windows-update")
        .with_timeout(7200)
        .with_requires("pwsh")
        .with_skip_if(cli.skip_windows_update, "disabled by --skip-windows-update"),
        // ── Defender ─────────────────────────────────────────────────────────
        Task::new_vec(
            "defender",
            "system",
            &["windows", "security"],
            "pwsh",
            pwsh_cmd(
                "$mode = try { (Get-MpComputerStatus -ErrorAction Stop).AMRunningMode } catch { $null }; \
                 if ($mode -and $mode -ne 'Normal') { Write-Output \"Defender signature update skipped: AMRunningMode='$mode' (third-party AV active, Defender passive).\"; return }; \
                 try { Update-MpSignature -ErrorAction Stop; Write-Output 'Defender signatures updated.' } \
                 catch { Write-Output \"Defender update skipped: $($_.Exception.Message)\" }",
            ),
        )
        .with_resource("defender")
        .with_requires("pwsh")
        .with_skip_if(cli.skip_defender, "disabled by --skip-defender"),
        // ── WSL ──────────────────────────────────────────────────────────────
        Task::new(
            "wsl",
            "system",
            &["windows", "linux"],
            "wsl",
            &["--update"],
        )
        .with_acceptable_exit_codes(&[-1])
        .with_skip_if(cli.skip_wsl, "disabled by --skip-wsl"),
        Task::new_vec(
            "wsl-distros",
            "system",
            &["windows", "linux"],
            "pwsh",
            pwsh_cmd(wsl_distros_script()),
        )
        .with_resource("wsl")
        .with_timeout(3600)
        .with_requires("wsl")
        .with_skip_if(
            cli.skip_wsl || cli.skip_wsl_distros,
            "disabled by WSL skip flag",
        ),
        // ── Windows Features ─────────────────────────────────────────────────
        Task::new_vec(
            "windows-features",
            "system",
            &["windows", "system"],
            "pwsh",
            pwsh_cmd(&windows_features_script(&config.windows_optional_features)),
        )
        .with_requires("pwsh"),
        // ── AppX repair ──────────────────────────────────────────────────────
        Task::new_vec(
            "appx-repair",
            "system",
            &["windows", "store"],
            "pwsh",
            pwsh_cmd(appx_repair_script()),
        )
        .with_requires("pwsh"),
        // ── Linux packages (Arch/WSL) ─────────────────────────────────────────
        Task::new_with_requires(
            "pacman",
            "package-manager",
            &["linux", "arch"],
            "sudo",
            vec!["pacman".into(), "-Syu".into(), "--noconfirm".into()],
            "pacman",
        ),
        // ── MSYS2 (native Windows) ───────────────────────────────────────────
        // winget refuses to upgrade MSYS2 ("cannot be upgraded using winget");
        // its own pacman is the supported path.
        Task::new_vec(
            "msys2",
            "package-manager",
            &["windows", "msys2"],
            "pwsh",
            pwsh_cmd(msys2_script()),
        )
        .with_resource("msys2")
        .with_timeout(1800)
        .with_requires("pwsh"),
        // ── JavaScript ───────────────────────────────────────────────────────
        Task::new_vec(
            "npm",
            "javascript",
            &["node"],
            "python",
            npm_upgrade_args(&config.npm_skip_packages),
        )
        .with_requires("npm")
        .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new("pnpm", "javascript", &["node"], "pnpm", &["self-update"])
            .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new(
            "yarn",
            "javascript",
            &["node"],
            "yarn",
            &["global", "upgrade"],
        )
        .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new("bun", "javascript", &["node"], "bun", &["upgrade"])
            .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new("deno", "javascript", &["node"], "deno", &["upgrade"])
            .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new_vec(
            "volta",
            "javascript",
            &["node"],
            "python",
            advisory_script(
                "volta",
                "Volta does not provide a stable non-interactive self-update. \
                 Install/update via winget/scoop/chocolatey for automatic coverage.",
            ),
        )
        .with_requires("volta")
        .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new_vec(
            "fnm",
            "javascript",
            &["node"],
            "python",
            advisory_script(
                "fnm",
                "fnm does not provide a self-update command. \
                 Install/update via winget/scoop/chocolatey for automatic coverage.",
            ),
        )
        .with_requires("fnm")
        .with_skip_if(skip_node, "disabled by --skip-node"),
        Task::new_vec(
            "nvm",
            "javascript",
            &["node"],
            "python",
            advisory_script(
                "nvm",
                "nvm (for Windows): update via winget upgrade --id CoreyButler.NVMforWindows",
            ),
        )
        .with_requires("nvm")
        .with_skip_if(skip_node, "disabled by --skip-node"),
        // ── Python ───────────────────────────────────────────────────────────
        Task::new_vec("pip", "python", &["python"], "python", pip_upgrade_args(&pip_skip)),
        Task::new_vec(
            "pip-health",
            "python",
            &["python", "health"],
            "python",
            pip_health_args(&config.pip_ignore_health_packages),
        )
        .with_skip_if(cli.skip_pip_health, "disabled by --skip-pip-health"),
        // pipx defaults to the uv backend; force pip so it works when uv is not installed.
        Task::new("pipx", "python", &["python"], "pipx", &["upgrade-all", "--backend", "pip"]),
        Task::new_vec("uv", "python", &["python"], "python", uv_self_update_args()),
        Task::new(
            "uv-tools",
            "python",
            &["python"],
            "uv",
            &["tool", "upgrade", "--all"],
        )
        .with_skip_if(cli.skip_uv_tools, "disabled by --skip-uv-tools"),
        Task::new_vec(
            "uv-python",
            "python",
            &["python", "uv"],
            "python",
            uv_python_upgrade_args(),
        )
        .with_requires("uv")
        .with_skip_if(cli.skip_uv_tools, "disabled by --skip-uv-tools"),
        Task::new_vec(
            "poetry",
            "python",
            &["python"],
            "python",
            poetry_self_update_args(),
        )
        .with_requires("poetry")
        .with_skip_if(cli.skip_poetry, "disabled by --skip-poetry"),
        Task::new_vec("conda", "python", &["python"], "python", conda_upgrade_args(&config.conda_skip_envs))
            .with_requires("conda")
            .with_resource("conda")
            .with_skip_if(cli.skip_conda, "disabled by --skip-conda"),
        // ── Rust ─────────────────────────────────────────────────────────────
        Task::new("rustup", "systems-language", &["rust"], "rustup", &["update"])
            .with_resource("rust")
            .with_skip_if(skip_rust, "disabled by --skip-rust"),
        Task::new_with_requires(
            "cargo",
            "systems-language",
            &["rust"],
            "cargo",
            vec!["install-update".into(), "-a".into()],
            "cargo-install-update",
        )
        .with_resource("rust")
        .with_skip_if(skip_rust, "disabled by --skip-rust"),
        // ── Go ───────────────────────────────────────────────────────────────
        Task::new(
            "go",
            "systems-language",
            &["go"],
            "go",
            &["install", "golang.org/x/tools/gopls@latest"],
        )
        .with_skip_if(skip_go, "disabled by --skip-go"),
        // ── PHP ──────────────────────────────────────────────────────────────
        Task::new(
            "composer",
            "runtime",
            &["php"],
            "composer",
            &["self-update", "--no-interaction"],
        )
        .with_skip_if(cli.skip_composer, "disabled by --skip-composer"),
        // ── Ruby ─────────────────────────────────────────────────────────────
        Task::new("ruby-gems", "runtime", &["ruby"], "gem", &["update"])
            .with_skip_if(cli.skip_ruby, "disabled by --skip-ruby"),
        // ── Flutter ──────────────────────────────────────────────────────────
        Task::new(
            "flutter",
            "systems-language",
            &["flutter"],
            "flutter",
            &["upgrade"],
        )
        .with_skip_if(cli.skip_flutter, "disabled by --skip-flutter"),
        // ── Julia ────────────────────────────────────────────────────────────
        Task::new("juliaup", "systems-language", &["julia"], "juliaup", &["update"]),
        Task::new_vec(
            "gh",
            "dev-tools",
            &["github", "dev-tools"],
            "python",
            gh_upgrade_args(),
        )
        .with_requires("gh")
        .with_resource("winget")
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212]),
        // ── .NET ─────────────────────────────────────────────────────────────
        Task::new(
            "dotnet-workloads",
            "dotnet",
            &["dotnet"],
            "dotnet",
            &["workload", "update"],
        )
        .with_timeout(3600),
        Task::new(
            "dotnet-tools",
            "dotnet",
            &["dotnet"],
            "dotnet",
            &["tool", "update", "--global", "--all"],
        ),
        // ── PowerShell ───────────────────────────────────────────────────────
        Task::new_vec(
            "powershell7",
            "system",
            &["windows", "powershell"],
            "winget",
            vec![
                "upgrade".into(), "--id".into(), "Microsoft.PowerShell".into(),
                "--exact".into(), "--include-unknown".into(), "--disable-interactivity".into(),
                "--accept-package-agreements".into(), "--accept-source-agreements".into(),
                "--silent".into(),
            ],
        )
        .with_timeout(300)
        .with_acceptable_exit_codes(&[-1978335188, -1978335189, -1978335212])
        .with_requires("winget"),
        Task::new_vec(
            "powershell-modules",
            "powershell",
            &["powershell"],
            "pwsh",
            pwsh_cmd(powershell_modules_script()),
        )
        .with_resource("powershell-gallery")
        .with_requires("pwsh")
        .with_skip_if(cli.skip_powershell_modules, "disabled by --skip-powershell-modules"),
        Task::new_vec(
            "powershell-help",
            "powershell",
            &["powershell"],
            "pwsh",
            pwsh_cmd(
                "try { Update-Help -Force -ErrorAction Stop } \
                 catch { Write-Output \"PowerShell help not fully refreshed: $($_.Exception.Message)\" }",
            ),
        )
        .with_resource("powershell-gallery")
        .with_requires("pwsh")
        .with_skip_if(!cli.update_powershell_help, "opt-in via --update-powershell-help"),
        // ── Editor ───────────────────────────────────────────────────────────
        Task::new_vec(
            "vscode-extensions",
            "editor",
            &["vscode"],
            "python",
            vscode_extensions_args(),
        )
        .with_requires("code")
        .with_skip_if(cli.skip_vscode_extensions, "disabled by --skip-vscode-extensions"),
        // ── Dev tools ────────────────────────────────────────────────────────
        Task::new_with_requires(
            "git-lfs",
            "dev-tools",
            &["git"],
            "git",
            vec!["lfs".into(), "install".into(), "--skip-repo".into()],
            "git-lfs",
        )
        .with_skip_if(cli.skip_git_lfs, "disabled by --skip-git-lfs"),
        Task::new(
            "gh-extensions",
            "dev-tools",
            &["github"],
            "gh",
            &["extension", "upgrade", "--all"],
        ),
        Task::new_vec(
            "devcontainer",
            "dev-tools",
            &["dev"],
            "python",
            github_version_check_args("devcontainer", &["--version"], "devcontainers/cli"),
        )
        .with_requires("devcontainer"),
        // ── Version managers ─────────────────────────────────────────────────
        Task::new_vec(
            "mise",
            "version-manager",
            &["mise"],
            "python",
            vec![
                "-c".to_string(),
                [
                    "import shutil, subprocess, sys",
                    r#"p = (shutil.which("mise") or "").replace("\\", "/").lower()"#,
                    r#"if "winget" in p or "/microsoft/" in p or "/scoop/" in p or "/homebrew/" in p:"#,
                    r#"    print("mise is package-manager-managed; update handled by winget/scoop task")"#,
                    "    sys.exit(0)",
                    r#"r = subprocess.run(["mise", "self-upgrade"])"#,
                    "sys.exit(r.returncode)",
                ]
                .join("\n"),
            ],
        )
        .with_requires("mise"),
        Task::new("mise-upgrade", "version-manager", &["mise"], "mise", &["upgrade"]),
        // ── Media tools ──────────────────────────────────────────────────────
        Task::new_vec("yt-dlp", "media-tools", &["media"], "python", yt_dlp_upgrade_args())
            .with_requires("yt-dlp"),
        // ── Shell tools ──────────────────────────────────────────────────────
        Task::new_vec(
            "oh-my-posh",
            "shell",
            &["shell"],
            "python",
            oh_my_posh_upgrade_args(),
        )
        .with_requires("oh-my-posh"),
        Task::new_vec(
            "starship",
            "shell",
            &["shell"],
            "python",
            starship_upgrade_args(),
        )
        .with_requires("starship")
        .with_skip_if(cli.skip_starship, "disabled by --skip-starship"),
        Task::new_vec(
            "zoxide",
            "shell",
            &["shell"],
            "python",
            advisory_script(
                "zoxide",
                "zoxide: update via your package manager (winget/scoop/choco/brew).",
            ),
        )
        .with_requires("zoxide"),
        Task::new("tldr", "dev-tools", &["dev-tools"], "tldr", &["--update"]),
        // ── Cloud tools ──────────────────────────────────────────────────────
        Task::new_vec("gcloud", "cloud", &["cloud"], "python", gcloud_upgrade_args())
            .with_resource("gcloud")
            .with_timeout(600)
            .with_requires("gcloud")
            .with_skip_if(cli.skip_cloud_tools, "disabled by --skip-cloud-tools"),
        Task::new("az", "cloud", &["cloud"], "az", &["upgrade", "--all", "-y"])
            .with_resource("az")
            .with_timeout(600)
            .with_skip_if(cli.skip_cloud_tools, "disabled by --skip-cloud-tools"),
        Task::new_vec("aws", "cloud", &["cloud"], "python", aws_upgrade_args())
            .with_requires("aws")
            .with_resource("aws")
            .with_skip_if(cli.skip_cloud_tools, "disabled by --skip-cloud-tools"),
        // ── Infrastructure tools ─────────────────────────────────────────────
        Task::new_vec("terraform", "infrastructure", &["infrastructure"], "python", terraform_upgrade_args())
            .with_resource("terraform")
            .with_requires("terraform")
            .with_skip_if(cli.skip_infra_tools, "disabled by --skip-infra-tools"),
        Task::new("pulumi", "infrastructure", &["infrastructure"], "pulumi", &["upgrade"])
            .with_timeout(600)
            .with_resource("pulumi")
            .with_skip_if(cli.skip_infra_tools, "disabled by --skip-infra-tools"),
        Task::new_vec("opentofu", "infrastructure", &["infrastructure"], "python",
            github_version_check_args("tofu", &["--version"], "opentofu/opentofu"))
            .with_requires("tofu")
            .with_skip_if(cli.skip_infra_tools, "disabled by --skip-infra-tools"),
        Task::new_vec("packer", "infrastructure", &["infrastructure"], "python",
            github_version_check_args("packer", &["--version"], "hashicorp/packer"))
            .with_requires("packer")
            .with_skip_if(cli.skip_infra_tools, "disabled by --skip-infra-tools"),
        // ── Kubernetes tools ─────────────────────────────────────────────────
        Task::new_vec("kubectl", "infrastructure", &["kubernetes"], "python",
            kubectl_check_args())
            .with_resource("kubectl")
            .with_requires("kubectl")
            .with_skip_if(cli.skip_k8s_tools, "disabled by --skip-k8s-tools"),
        Task::new("helm", "infrastructure", &["kubernetes"], "helm", &["repo", "update"])
            .with_timeout(120)
            .with_resource("helm")
            .with_skip_if(cli.skip_k8s_tools, "disabled by --skip-k8s-tools"),
        // ── Security tools ───────────────────────────────────────────────────
        Task::new_vec("gitleaks", "security", &["security"], "python",
            github_version_check_args("gitleaks", &["version"], "gitleaks/gitleaks"))
            .with_requires("gitleaks"),
        Task::new("trivy", "security", &["security"], "trivy", &["update"])
            .with_timeout(300)
            .with_resource("trivy"),
        // ── Static site ──────────────────────────────────────────────────────
        Task::new_vec("hugo", "dev-tools", &["static-site"], "python",
            github_version_check_args("hugo", &["version"], "gohugoio/hugo"))
            .with_requires("hugo")
            .with_skip_if(cli.skip_hugo, "disabled by --skip-hugo"),
        // ── Package managers (C++) ────────────────────────────────────────────
        Task::new_vec("vcpkg", "package-manager", &["cpp"], "python",
            vcpkg_upgrade_args(&config.vcpkg_skip_packages))
            .with_resource("vcpkg")
            .with_requires("vcpkg")
            .with_skip_if(cli.skip_vcpkg, "disabled by --skip-vcpkg"),
        // ── AI tools ─────────────────────────────────────────────────────────
        Task::new_vec("claude", "ai-tools", &["ai", "claude"], "python", claude_upgrade_args())
            .with_timeout(300)
            .with_requires("claude"),
        Task::new_vec("codex", "ai-tools", &["ai", "codex"], "python", codex_upgrade_args())
            .with_timeout(300)
            .with_requires("codex"),
        Task::new_vec(
            "ollama-models",
            "ai",
            &["ai"],
            "python",
            ollama_models_upgrade_args(cli.ollama_timeout_sec),
        )
        .with_timeout(7200)
        .with_resource("ollama")
        .with_requires("ollama")
        .with_skip_if(!cli.update_ollama_models, "use --update-ollama-models to refresh local models"),
        // ── Docker ───────────────────────────────────────────────────────────
        Task::new(
            "docker-prune",
            "maintenance",
            &["docker", "maintenance"],
            "docker",
            &["system", "prune", "-f", "--volumes"],
        )
        .with_timeout(300)
        .with_skip_if(!cli.deep_clean, "opt-in: requires --deep-clean and docker running"),
        // ── Maintenance ──────────────────────────────────────────────────────
        Task::new_vec(
            "cleanup",
            "maintenance",
            &["maintenance"],
            "python",
            cleanup_args(config.temp_cleanup_days, cli.deep_clean, cli.skip_destructive),
        )
        .with_timeout(3600)
        .with_skip_if(cli.skip_cleanup, "disabled by --skip-cleanup"),
        Task::new_vec(
            "self-update",
            "maintenance",
            &["self"],
            "python",
            self_update_check_args(),
        ),
    ];

    // ── GitHub-release tools (manually-installed, config-driven, opt-in) ──────
    for tool in &config.github_tools {
        let name = tool.id.clone().unwrap_or_else(|| {
            tool.repo
                .rsplit('/')
                .next()
                .unwrap_or(&tool.repo)
                .to_string()
        });
        let id: &'static str = Box::leak(format!("gh-{name}").into_boxed_str());
        tasks.push(
            Task::new_vec(
                id,
                "github-tools",
                &["github", "tools"],
                "pwsh",
                github_release_args(tool),
            )
            .with_resource("github-tools")
            .with_timeout(900)
            .with_skip_if(
                !cli.update_github_tools,
                "opt-in: use --update-github-tools",
            ),
        );
    }

    tasks.into_iter().map(mark_missing).collect()
}

// ─── Task impl ───────────────────────────────────────────────────────────────

impl Task {
    fn new(
        id: &'static str,
        category: &'static str,
        tags: &'static [&'static str],
        command: &'static str,
        args: &[&str],
    ) -> Self {
        Self::new_vec(
            id,
            category,
            tags,
            command,
            args.iter().map(|arg| arg.to_string()).collect(),
        )
    }

    fn new_vec(
        id: &'static str,
        category: &'static str,
        tags: &'static [&'static str],
        command: &'static str,
        args: Vec<String>,
    ) -> Self {
        Self {
            id,
            category,
            tags,
            command,
            args,
            requires: command,
            resource: "",
            depends_on: &[],
            skip_reason: None,
            timeout_override: None,
            acceptable_exit_codes: vec![],
            ok_on_timeout: false,
        }
    }

    fn new_with_requires(
        id: &'static str,
        category: &'static str,
        tags: &'static [&'static str],
        command: &'static str,
        args: Vec<String>,
        requires: &'static str,
    ) -> Self {
        Self {
            id,
            category,
            tags,
            command,
            args,
            requires,
            resource: "",
            depends_on: &[],
            skip_reason: None,
            timeout_override: None,
            acceptable_exit_codes: vec![],
            ok_on_timeout: false,
        }
    }

    fn with_resource(mut self, resource: &'static str) -> Self {
        self.resource = resource;
        self
    }

    fn with_requires(mut self, requires: &'static str) -> Self {
        self.requires = requires;
        self
    }

    fn with_depends_on(mut self, depends_on: &'static [&'static str]) -> Self {
        self.depends_on = depends_on;
        self
    }

    fn with_timeout(mut self, secs: u64) -> Self {
        self.timeout_override = Some(Duration::from_secs(secs));
        self
    }

    fn with_acceptable_exit_codes(mut self, codes: &[i32]) -> Self {
        self.acceptable_exit_codes = codes.to_vec();
        self
    }

    /// Set skip_reason when condition is true and no reason is already set.
    fn with_skip_if(mut self, condition: bool, reason: &str) -> Self {
        if condition && self.skip_reason.is_none() {
            self.skip_reason = Some(reason.to_string());
        }
        self
    }
}

// ─── Script/args builders ────────────────────────────────────────────────────

/// Wrap a PowerShell command string as pwsh -NoProfile -NonInteractive -Command args.
fn pwsh_cmd(script: &str) -> Vec<String> {
    vec![
        "-NoProfile".into(),
        "-NonInteractive".into(),
        "-Command".into(),
        script.to_string(),
    ]
}

/// Dispatch to the right provider builder for a managed tool.
fn cross_manager_args(fallback: &BTreeMap<String, BTreeMap<String, String>>) -> Vec<String> {
    let json = serde_json::to_string(fallback).unwrap_or_else(|_| "{}".to_string());
    vec![
        "-c".to_string(),
        format!(
            r#"import json, shutil, subprocess, sys
fallback = json.loads({json:?})
if not fallback:
    print('No cross-manager fallback apps configured.')
    sys.exit(0)
choco = shutil.which('choco')
scoop = shutil.which('scoop')
if not choco and not scoop:
    print('No alternate package managers (choco/scoop) available for fallback.')
    sys.exit(0)
failed = False
for winget_id, alt in fallback.items():
    print(f'Fallback check: {{winget_id}}')
    if choco and alt.get('choco'):
        result = subprocess.run(['choco', 'upgrade', alt['choco'], '-y', '--no-progress'], text=True)
        failed = failed or result.returncode not in (0, 1)
    if scoop and alt.get('scoop'):
        result = subprocess.run(['scoop', 'update', alt['scoop']], text=True)
        failed = failed or result.returncode != 0
sys.exit(1 if failed else 0)
"#
        ),
    ]
}

fn gh_upgrade_args() -> Vec<String> {
    vec![
        "-c".to_string(),
        r#"import shutil, subprocess, sys
path = (shutil.which('gh') or '').replace('\\', '/').lower()
managed_markers = ['/scoop/apps/', '/scoop/shims/', '/chocolatey/lib/', '/microsoft/winget/packages/', '/windowsapps/']
if any(marker in path for marker in managed_markers):
    print(f'gh is managed by another package manager; handled elsewhere: {path}')
    sys.exit(0)
if not shutil.which('winget'):
    print('winget missing; gh standalone update skipped.')
    sys.exit(0)
cmd = ['winget', 'upgrade', '--id', 'GitHub.cli', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent']
result = subprocess.run(cmd, capture_output=True, text=True)
out = (result.stdout + result.stderr).strip()
if out:
    print(out)
benign = ['No installed package found', 'No applicable update', 'No available upgrade', 'No newer package']
if any(b in out for b in benign):
    print('gh not tracked by winget (or already current); nothing to upgrade.')
    sys.exit(0)
sys.exit(result.returncode)
"#.to_string(),
    ]
}
fn github_release_args(tool: &GithubTool) -> Vec<String> {
    match tool.provider.as_deref().unwrap_or("github") {
        "gitlab" => gitlab_release_args(tool),
        "gitlab-artifact" => gitlab_artifact_args(tool),
        _ => github_release_inner_args(tool),
    }
}

/// Shared PowerShell helpers (version compare + marker) prepended to each script.
fn tool_script_prelude() -> &'static str {
    r#"$ErrorActionPreference = 'Stop'
function Test-UpToDate($l, $r) {
  if ($null -eq $l -or $null -eq $r) { return $false }
  $lv = $null; $rv = $null
  if ([version]::TryParse($l, [ref]$lv) -and [version]::TryParse($r, [ref]$rv)) {
    return $lv -ge $rv
  }
  $ln = ($l -replace '\D', ''); $rn = ($r -replace '\D', '')
  if ($ln -and $rn) { return [int64]$ln -ge [int64]$rn }
  return $false
}
"#
}

/// GitLab release-asset provider: releases/permalink/latest, asset link by AssetRegex.
fn gitlab_release_args(tool: &GithubTool) -> Vec<String> {
    let project = tool.project_id.clone().unwrap_or_default();
    let script = format!(
        "{prelude}\
$repo    = '{repo}'
$dir     = '{dir}'
$assetRe = '{asset}'
$proj    = '{proj}'
$marker  = Join-Path $dir (\".ue-\" + ($repo -replace '[\\\\/:]', '_') + \".version\")

$local = if (Test-Path $marker) {{ (Get-Content $marker -Raw).Trim() }} else {{ $null }}
$rel = Invoke-RestMethod -Uri \"https://gitlab.com/api/v4/projects/$proj/releases/permalink/latest\"
$tag = $rel.tag_name
Write-Host \"$repo  local=$local  latest=$tag\"
if (Test-UpToDate $local $tag) {{ Write-Host 'up to date'; exit 0 }}

$dl = $rel.assets.links | Where-Object {{ $_.name -match $assetRe }} | Select-Object -First 1
if (-not $dl) {{ Write-Host \"no asset matched /$assetRe/\"; exit 1 }}
$url = if ($dl.direct_asset_url) {{ $dl.direct_asset_url }} else {{ $dl.url }}
$tmp = Join-Path $env:TEMP $dl.name
Write-Host \"downloading $($dl.name)\"
Invoke-WebRequest -Uri $url -OutFile $tmp
if (-not (Test-Path $dir)) {{ New-Item -ItemType Directory -Force -Path $dir | Out-Null }}
if ($dl.name -match '\\.zip$') {{ Expand-Archive -Path $tmp -DestinationPath $dir -Force }}
else {{ Copy-Item $tmp (Join-Path $dir $dl.name) -Force }}
Remove-Item $tmp -Force -ErrorAction SilentlyContinue
Set-Content -Path $marker -Value $tag -NoNewline
Write-Host \"updated $repo -> $tag\"
",
        prelude = tool_script_prelude(),
        repo = tool.repo,
        dir = tool.install_dir,
        asset = tool.asset_regex,
        proj = project,
    );
    pwsh_cmd(&script)
}

/// GitLab CI-artifact provider: latest successful pipeline artifact for a Job.
/// Used by tools (e.g. OpenRGB) that publish builds as pipeline artifacts, not
/// release assets. Version is the pipeline commit sha, stored in the marker.
fn gitlab_artifact_args(tool: &GithubTool) -> Vec<String> {
    let project = tool.project_id.clone().unwrap_or_default();
    let git_ref = tool.git_ref.clone().unwrap_or_else(|| "master".to_string());
    let job = tool.job.clone().unwrap_or_default();
    let script = format!(
        "{prelude}\
$repo = '{repo}'
$dir  = '{dir}'
$proj = '{proj}'
$ref  = '{git_ref}'
$job  = '{job}'
$marker = Join-Path $dir (\".ue-\" + ($repo -replace '[\\\\/:]', '_') + \".version\")

$pl = Invoke-RestMethod -Uri \"https://gitlab.com/api/v4/projects/$proj/pipelines?ref=$ref&status=success&per_page=1\"
if (-not $pl) {{ Write-Host 'no successful pipeline found'; exit 1 }}
$latest = $pl[0].sha
$local = if (Test-Path $marker) {{ (Get-Content $marker -Raw).Trim() }} else {{ $null }}
Write-Host \"$repo  local=$local  latest=$latest\"
if ($local -eq $latest) {{ Write-Host 'up to date'; exit 0 }}

$enc = [uri]::EscapeDataString($job)
$url = \"https://gitlab.com/api/v4/projects/$proj/jobs/artifacts/$ref/download?job=$enc\"
$tmp = Join-Path $env:TEMP ((($repo -replace '[\\\\/:]', '_')) + '-artifact.zip')
Write-Host \"downloading artifact (job=$job)\"
Invoke-WebRequest -Uri $url -OutFile $tmp
if (-not (Test-Path $dir)) {{ New-Item -ItemType Directory -Force -Path $dir | Out-Null }}
Expand-Archive -Path $tmp -DestinationPath $dir -Force
Remove-Item $tmp -Force -ErrorAction SilentlyContinue
Set-Content -Path $marker -Value $latest -NoNewline
Write-Host \"updated $repo -> $latest\"
",
        prelude = tool_script_prelude(),
        repo = tool.repo,
        dir = tool.install_dir,
        proj = project,
        git_ref = git_ref,
        job = job,
    );
    pwsh_cmd(&script)
}

/// Build the pwsh args for a GitHub-release tool update: compare local version
/// to the latest release, download + extract the matching asset when behind.
fn github_release_inner_args(tool: &GithubTool) -> Vec<String> {
    let version_cmd = tool.version_cmd.clone().unwrap_or_default();
    let version_regex = tool
        .version_regex
        .clone()
        .unwrap_or_else(|| r"(\d+)".to_string());
    let script = format!(
        r#"$ErrorActionPreference = 'Stop'
$repo    = '{repo}'
$dir     = '{dir}'
$assetRe = '{asset}'
$verRe   = '{vre}'
$verCmd  = @'
{vcmd}
'@

$marker = Join-Path $dir (".ue-" + ($repo -replace '[\\/:]', '_') + ".version")

$local = $null
if ($verCmd.Trim()) {{
  try {{
    Push-Location $dir
    $out = & ([scriptblock]::Create($verCmd)) 2>&1 | Out-String
    Pop-Location
    if ($out -match $verRe) {{ $local = $matches[1] }}
  }} catch {{ Write-Host "version probe failed: $_" }}
}}
# Fall back to the marker written by a previous run (tools with no queryable version).
if ($null -eq $local -and (Test-Path $marker)) {{
  $local = (Get-Content $marker -Raw).Trim()
}}

$hdr = @{{ 'User-Agent' = 'updateEverything' }}
$rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers $hdr
$tag = $rel.tag_name
$latest = if ($tag -match '([\d.]+)') {{ $matches[1] }} else {{ $null }}

# Compare local vs latest: semver via [version] when both parse, else numeric.
function Test-UpToDate($l, $r) {{
  if ($null -eq $l -or $null -eq $r) {{ return $false }}
  $lv = $null; $rv = $null
  if ([version]::TryParse($l, [ref]$lv) -and [version]::TryParse($r, [ref]$rv)) {{
    return $lv -ge $rv
  }}
  $ln = ($l -replace '\D', ''); $rn = ($r -replace '\D', '')
  if ($ln -and $rn) {{ return [int64]$ln -ge [int64]$rn }}
  return $false
}}

Write-Host "$repo  local=$local  latest=$tag"
if (Test-UpToDate $local $latest) {{
  Write-Host "up to date"; exit 0
}}

$dl = $rel.assets | Where-Object {{ $_.name -match $assetRe }} | Select-Object -First 1
if (-not $dl) {{ Write-Host "no asset matched /$assetRe/"; exit 1 }}

$tmp = Join-Path $env:TEMP $dl.name
Write-Host "downloading $($dl.name)"
Invoke-WebRequest -Uri $dl.browser_download_url -OutFile $tmp -Headers $hdr
if (-not (Test-Path $dir)) {{ New-Item -ItemType Directory -Force -Path $dir | Out-Null }}
if ($dl.name -match '\.zip$') {{
  Write-Host "extracting to $dir"
  Expand-Archive -Path $tmp -DestinationPath $dir -Force
}} else {{
  Copy-Item $tmp (Join-Path $dir $dl.name) -Force
}}
Remove-Item $tmp -Force -ErrorAction SilentlyContinue
Set-Content -Path $marker -Value $tag -NoNewline
Write-Host "updated $repo -> $tag"
"#,
        repo = tool.repo,
        dir = tool.install_dir,
        asset = tool.asset_regex,
        vre = version_regex,
        vcmd = version_cmd,
    );
    pwsh_cmd(&script)
}

fn winget_pin_skip_script(skip_packages: &[String]) -> String {
    // Pin each skip package so `winget upgrade --all` (without --include-pinned)
    // leaves it untouched. Idempotent: pinning an already-pinned id just warns.
    let list = skip_packages
        .iter()
        .map(|p| format!("'{}'", p.replace('\'', "''")))
        .collect::<Vec<_>>()
        .join(", ");
    format!(
        "$pkgs = @({list})\n\
         if (-not $pkgs) {{ Write-Output 'No winget skip-packages to pin.'; exit 0 }}\n\
         foreach ($p in $pkgs) {{\n\
         \x20   winget pin add --id $p --exact --source winget --accept-source-agreements --disable-interactivity 2>&1 | Out-Null\n\
         \x20   Write-Output (\"pinned: {{0}}\" -f $p)\n\
         }}\n\
         exit 0"
    )
}

fn winget_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    pwsh_cmd(&winget_per_package_script(skip_packages))
}

/// PowerShell that upgrades each outdated winget package individually.
///
/// `winget upgrade --all` aborts the whole batch without installing anything when any
/// listed package needs an install location (e.g. Blizzard.BattleNet): it prints the
/// table, then exits -1978335137 with zero installs. Per-package keeps one unusable
/// entry from blocking the rest.
///
/// NOTE: `winget upgrade` has no `--exclude`. Packages in winget_skip_packages are
/// excluded by pinning them (see winget-pin-skip task) and omitting `--include-pinned`
/// here, so pinned packages are left untouched.
fn winget_per_package_script(skip_packages: &[String]) -> String {
    let skip_list = skip_packages
        .iter()
        .map(|s| format!("'{}'", s.replace('\'', "''")))
        .collect::<Vec<_>>()
        .join(", ");

    let script = format!(
        r#"
$skip = @({skip_list})
$raw = winget upgrade --include-unknown --disable-interactivity 2>&1 | Out-String
$lines = $raw -split "`r?`n"
$hdr = $null
for ($i = 0; $i -lt $lines.Count; $i++) {{
    if ($lines[$i] -match '^Name\s+Id\s+Version') {{ $hdr = $i; break }}
}}
if ($null -eq $hdr) {{ Write-Output 'winget: no upgrade table found'; exit 0 }}
$idStart = $lines[$hdr].IndexOf('Id')
$verStart = $lines[$hdr].IndexOf('Version')
$ids = @()
for ($i = $hdr + 1; $i -lt $lines.Count; $i++) {{
    $line = $lines[$i]
    if ($line -match '^-{{3,}}$') {{ continue }}
    if ([string]::IsNullOrWhiteSpace($line)) {{ break }}
    if ($line.Length -le $idStart) {{ continue }}
    $len = [Math]::Min($verStart - $idStart, $line.Length - $idStart)
    $id = $line.Substring($idStart, $len).Trim()
    if ($id -and $id -notmatch '\s' -and $id -notin $skip) {{ $ids += $id }}
}}
$ids = $ids | Select-Object -Unique
Write-Output ("winget: {{0}} package(s) to upgrade" -f $ids.Count)
# Codes winget returns when a listed upgrade cannot apply to this system; not our failure.
$tolerated = @(-1978335188, -1978335189, -1978335212, -1978335215)
# Package-level blockers no unattended run can clear: installer needs an install
# location (-1978335137), newer version uses a different install technology so it
# needs uninstall+install (-1978335090), or the de-elevated pass cannot touch a
# machine-scope package (-1978335226). Reported, but they do not fail the task.
$manual = @(-1978335137, -1978335090, -1978335226)
$ok = @(); $skipped = @(); $needsManual = @(); $failed = @()
foreach ($id in $ids) {{
    winget upgrade --id $id --exact --source winget --include-unknown --accept-package-agreements --accept-source-agreements --disable-interactivity --silent
    $code = $LASTEXITCODE
    if ($code -eq 0) {{ $ok += $id }}
    elseif ($tolerated -contains $code) {{ $skipped += ("{{0}} ({{1}})" -f $id, $code) }}
    elseif ($manual -contains $code) {{ $needsManual += ("{{0}} ({{1}})" -f $id, $code) }}
    else {{ $failed += ("{{0}} ({{1}})" -f $id, $code) }}
}}
Write-Output ("upgraded ({{0}}): {{1}}" -f $ok.Count, ($ok -join ', '))
Write-Output ("not applicable ({{0}}): {{1}}" -f $skipped.Count, ($skipped -join ', '))
Write-Output ("needs manual action ({{0}}): {{1}}" -f $needsManual.Count, ($needsManual -join ', '))
Write-Output ("failed ({{0}}): {{1}}" -f $failed.Count, ($failed -join ', '))
if ($failed.Count -gt 0) {{ exit 1 }}
exit 0
"#
    );

    script
}

fn pip_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let skip_set = skip_packages
        .iter()
        .map(|s| format!("\"{}\"", s.to_lowercase().replace('"', "\\\"")))
        .collect::<Vec<_>>()
        .join(", ");

    let script = format!(
        r#"
import json, os, subprocess, sys, sysconfig
if os.path.exists(os.path.join(sysconfig.get_path("stdlib"), "EXTERNALLY-MANAGED")):
    print("pip: externally managed environment (PEP 668); skipping")
    sys.exit(0)
skip = {{{skip_set}}}
subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=False)
r = subprocess.run([sys.executable, "-m", "pip", "list", "--outdated", "--format=json"], capture_output=True, text=True)
pkgs = [p["name"] for p in json.loads(r.stdout or "[]") if p["name"].lower() not in skip]
if not pkgs:
    print("All pip packages up to date")
    sys.exit(0)
failed = []
for p in pkgs:
    rc = subprocess.run([sys.executable, "-m", "pip", "install", "-U", "--upgrade-strategy", "only-if-needed", p], check=False).returncode
    print(("upgraded " if rc == 0 else "FAILED ") + p)
    if rc != 0:
        failed.append(p)
sys.exit(1 if failed else 0)
"#
    );
    vec!["-c".into(), script]
}

fn pip_health_args(ignore_packages: &[String]) -> Vec<String> {
    let ignore_set = ignore_packages
        .iter()
        .map(|s| format!("\"{}\"", s.to_lowercase().replace('"', "\\\"")))
        .collect::<Vec<_>>()
        .join(", ");
    // Advisory only: pip check never updates anything, and dep conflicts are
    // usually pre-existing environment state. Report, but never fail the run.
    let script = format!(
        r#"
import subprocess, sys
ignore = {{{ignore_set}}}
r = subprocess.run([sys.executable, "-m", "pip", "check"], capture_output=True, text=True)
lines = [l for l in r.stdout.strip().splitlines() if l.strip()]
shown = [l for l in lines if not any(pkg in l.lower() for pkg in ignore)]
if not lines:
    print("pip check: no issues found")
else:
    if shown:
        print("pip check found dependency conflicts (advisory, not failing the run):")
        for l in shown:
            print("  " + l)
    ignored = len(lines) - len(shown)
    if ignored:
        print(f"pip check: {{ignored}} known/ignored conflict(s) suppressed")
if r.stderr.strip():
    print(r.stderr.strip())
sys.exit(0)
"#
    );
    vec!["-c".into(), script]
}

fn uv_self_update_args() -> Vec<String> {
    let script = r#"
import os, shutil, subprocess, sys, time

uv = shutil.which("uv") or ""
p = uv.replace("\\", "/").lower()
if "/python" in p or "/scripts/" in p:
    print("SKIPPED: uv is pip-managed; update handled by pip task.")
    sys.exit(0)

def update():
    r = subprocess.run(["uv", "self", "update"], capture_output=True, text=True)
    return r, (r.stdout or "") + (r.stderr or "")

r, out = update()
print(out.strip())
# `uv self update` replaces uvx.exe alongside uv.exe. Anything hosting a uvx
# tool (commonly an MCP server) holds that file open and the installer fails.
# Close the holders by exact path and retry; whatever spawned them restarts them.
if r.returncode != 0 and "being used by another process" in out:
    closed = close_locking_processes(os.path.dirname(uv), ["uv.exe", "uvx.exe"])
    if closed:
        time.sleep(2)
        r, out = update()
        print(out.strip())
    else:
        print("SKIPPED: uv self-update blocked and no closable holder was found.")
        sys.exit(0)
sys.exit(r.returncode)
"#;
    vec!["-c".into(), format!("{}{}", CLOSE_BLOCKERS_PY, script)]
}

fn poetry_self_update_args() -> Vec<String> {
    // pipx owns poetry's venv when it installed it; `poetry self update` then
    // re-pins poetry's shared libraries down to its own declared bounds,
    // undoing pipx's upgrades on every run. Let pipx keep them current.
    let script = r#"
import shutil, subprocess, sys
if shutil.which("pipx"):
    r = subprocess.run(["pipx", "list", "--short"], capture_output=True, text=True)
    for line in (r.stdout or "").splitlines():
        if line.split()[:1] == ["poetry"]:
            print("SKIPPED: poetry is pipx-managed; upgrades handled by the pipx task.")
            sys.exit(0)
r = subprocess.run(
    ["poetry", "self", "update", "--no-interaction"], capture_output=True, text=True
)
print(((r.stdout or "") + (r.stderr or "")).strip())
sys.exit(r.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn uv_python_upgrade_args() -> Vec<String> {
    let script = r#"
import subprocess, sys
r = subprocess.run(["uv", "python", "list", "--only-installed"], capture_output=True, text=True)
if r.returncode != 0:
    print("uv python list failed:", r.stderr.strip())
    sys.exit(0)
lines = [l.strip() for l in r.stdout.splitlines() if l.strip() and not l.startswith("cpython") or True]
versions = []
for line in r.stdout.splitlines():
    parts = line.split()
    if parts:
        ver = parts[0]
        if ver.startswith("cpython-") or ver.startswith("pypy-"):
            ver_num = ver.split("-")[1] if "-" in ver else ver
            if ver_num not in versions:
                versions.append(ver_num)
if not versions:
    print("No uv-managed Python versions installed")
    sys.exit(0)
print(f"Upgrading {len(versions)} uv Python version(s): {', '.join(versions)}")
r2 = subprocess.run(["uv", "python", "install"] + versions)
sys.exit(r2.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn npm_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let skip_json = skip_packages
        .iter()
        .map(|s| format!("\"{}\"", s.replace('"', "\\\"")))
        .collect::<Vec<_>>()
        .join(", ");

    let script = format!(
        r#"
import json, shutil, subprocess, sys
skip = [{skip_json}]
# Bare "npm" is npm.cmd on Windows; subprocess needs the resolved path (or
# shell=True) or CreateProcess fails with WinError 2.
npm = shutil.which("npm")
if not npm:
    print("npm not found in PATH; skipping.")
    sys.exit(0)
r = subprocess.run([npm, "ls", "-g", "--depth=0", "--json"], capture_output=True, text=True)
try:
    data = json.loads(r.stdout or "{{}}")
    pkgs = [k for k in data.get("dependencies", {{}}).keys() if k not in skip and not k.startswith("npm")]
except Exception:
    pkgs = []
if not pkgs:
    print("npm: no global packages to upgrade (or npm ls failed)")
    subprocess.run([npm, "update", "-g"])
    sys.exit(0)
failed = []
for p in pkgs:
    rc = subprocess.run([npm, "install", "-g", p]).returncode
    print(("upgraded " if rc == 0 else "FAILED ") + p)
    if rc != 0:
        failed.append(p)
sys.exit(1 if failed else 0)
"#
    );
    vec!["-c".into(), script]
}

fn oh_my_posh_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("oh-my-posh") or "").replace("\\", "/").lower()
managed = any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "windowsapps"])
if managed:
    print(f"oh-my-posh is managed by another package manager: {p}")
    sys.exit(0)
# Try winget first
if shutil.which("winget"):
    r = subprocess.run(
        ["winget", "upgrade", "--id", "JanDeDobbeleer.OhMyPosh", "--exact",
         "--include-unknown", "--disable-interactivity",
         "--accept-package-agreements", "--accept-source-agreements", "--silent"],
        capture_output=True, text=True
    )
    if r.returncode == 0 and "No installed package found" not in (r.stdout + r.stderr):
        out = (r.stdout + r.stderr).strip()
        if out:
            print(out)
        print("oh-my-posh checked via winget id JanDeDobbeleer.OhMyPosh.")
        sys.exit(0)
# Standalone upgrade
r2 = subprocess.run(["oh-my-posh", "upgrade"], capture_output=True, text=True)
import re
out = re.sub(r'\x1b\][^\a]*(\a|\x1b\\)', '', r2.stdout + r2.stderr).strip()
if out:
    print(out)
sys.exit(r2.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn vscode_extensions_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
if not shutil.which("code"):
    print("code not in PATH")
    sys.exit(0)
r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Code.exe", "/NH"],
                   capture_output=True, text=True, errors="ignore")
if "Code.exe" not in r.stdout:
    print("VSCode not running; skipping extension update (run with VSCode open)")
    sys.exit(0)
r2 = subprocess.run(["code", "--update-extensions"])
sys.exit(r2.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn yt_dlp_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("yt-dlp") or "").replace("\\", "/").lower()
if "pipx/venvs" in p:
    print("yt-dlp is managed by pipx; covered by the pipx task.")
    sys.exit(0)
if "/python" in p or "/scripts/" in p:
    print("yt-dlp is managed by pip; covered by the pip task.")
    sys.exit(0)
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "windowsapps"]):
    print(f"yt-dlp is managed by another package manager: {p}")
    sys.exit(0)
r = subprocess.run(["yt-dlp", "-U"])
sys.exit(r.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn gcloud_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("gcloud") or "").replace("\\", "/").lower()
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    print("gcloud is managed by a package manager; skipping self-update.")
    sys.exit(0)
r = subprocess.run(["gcloud", "components", "update", "--quiet"])
sys.exit(r.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn aws_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("aws") or "").replace("\\", "/").lower()
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "python", "scripts"]):
    r = subprocess.run(["aws", "--version"], capture_output=True, text=True)
    print(f"AWS CLI is managed ({r.stdout.strip()}); updating via package manager.")
    sys.exit(0)
print("Updating AWS CLI via pip...")
r = subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "awscli"])
sys.exit(r.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn terraform_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("terraform") or "").replace("\\", "/").lower()
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "tfenv"]):
    print("Terraform is managed by a package manager; skipping.")
    sys.exit(0)
try:
    r = subprocess.run(["terraform", "--version"], capture_output=True, text=True)
    lines = r.stdout.strip().splitlines()
    cur = lines[0] if lines else "unknown"
    print(f"Current Terraform: {cur}")
    import urllib.request, json as _json
    with urllib.request.urlopen("https://api.github.com/repos/hashicorp/terraform/releases/latest", timeout=15) as resp:
        data = _json.loads(resp.read())
    latest = data["tag_name"].lstrip("v")
    print(f"Latest Terraform: {latest}")
    if cur and latest and latest in cur:
        print("Terraform is current.")
    else:
        print(f"New version available: {latest}. Update via winget/scoop or download from https://releases.hashicorp.com/terraform/{latest}/")
except Exception as e:
    print(f"terraform check skipped: {e}")
sys.exit(0)
"#;
    vec!["-c".into(), script.into()]
}

fn kubectl_check_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("kubectl") or "").replace("\\", "/").lower()
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    print("kubectl is managed by a package manager; skipping.")
    sys.exit(0)
try:
    r = subprocess.run(["kubectl", "version", "--client", "-o", "json"], capture_output=True, text=True)
    import json
    d = json.loads(r.stdout or "{}")
    cur = d.get("clientVersion", {}).get("gitVersion", "unknown")
    print(f"Current kubectl: {cur}")
    import urllib.request
    with urllib.request.urlopen("https://dl.k8s.io/release/stable.txt", timeout=15) as resp:
        latest = resp.read().decode().strip()
    print(f"Latest stable kubectl: {latest}")
    if cur != "unknown" and latest in cur:
        print("kubectl is current.")
    else:
        print(f"Update available. Install via: winget upgrade --id Kubernetes.kubectl")
except Exception as e:
    print(f"kubectl check skipped: {e}")
sys.exit(0)
"#;
    vec!["-c".into(), script.into()]
}

fn starship_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
p = (shutil.which("starship") or "").replace("\\", "/").lower()
if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    print("starship is managed by a package manager; skipping.")
    sys.exit(0)
cmds = [["starship", "self", "update", "-y"], ["starship", "upgrade", "--yes"], ["starship", "self-update", "-y"]]
for cmd in cmds:
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        err = (r.stderr or "").lower()
        has_err = any(x in err for x in ["error", "unrecognized", "usage", "no such", "not found"])
        if r.returncode == 0 and not has_err:
            out = (r.stdout + r.stderr).strip()
            if out:
                print(out)
            print("starship upgraded.")
            sys.exit(0)
    except Exception:
        continue
print("starship: no compatible self-update command found. Update via winget/scoop/choco.")
sys.exit(0)
"#;
    vec!["-c".into(), script.into()]
}

fn github_version_check_args(bin: &str, ver_args: &[&str], repo: &str) -> Vec<String> {
    let bin = bin.to_string();
    let ver_args_json = ver_args
        .iter()
        .map(|a| format!("\"{}\"", a))
        .collect::<Vec<_>>()
        .join(", ");
    let script = format!(
        r#"
import subprocess, sys, urllib.request, json
try:
    r = subprocess.run(["{bin}", {ver_args_json}], capture_output=True, text=True, timeout=30)
    cur = (r.stdout + r.stderr).strip().splitlines()[0] if (r.stdout + r.stderr).strip() else "unknown"
    print(f"Current {bin}: {{cur}}")
    with urllib.request.urlopen(f"https://api.github.com/repos/{repo}/releases/latest", timeout=15) as resp:
        data = json.loads(resp.read())
    print(f"Latest {bin}: {{data['tag_name']}}")
except Exception as e:
    print(f"{bin} version check skipped: {{e}}")
sys.exit(0)
"#
    );
    vec!["-c".into(), script]
}

fn vcpkg_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let skip_json = skip_packages
        .iter()
        .map(|s| format!("\"{}\"", s))
        .collect::<Vec<_>>()
        .join(", ");
    let script = format!(
        r#"
import subprocess, sys
skip = [{skip_json}]
print("Updating vcpkg baseline...")
r1 = subprocess.run(["vcpkg", "update"])
print("Upgrading vcpkg packages...")
r2 = subprocess.run(["vcpkg", "upgrade", "--no-dry-run"])
sys.exit(max(r1.returncode, r2.returncode))
"#
    );
    vec!["-c".into(), script]
}

fn conda_upgrade_args(skip_envs: &[String]) -> Vec<String> {
    let skip_json = skip_envs
        .iter()
        .map(|s| format!("\"{}\"", s))
        .collect::<Vec<_>>()
        .join(", ");
    let script = format!(
        r#"
import subprocess, sys
skip = [{skip_json}]
print("Updating conda base environment...")
r1 = subprocess.run(["conda", "update", "-n", "base", "conda", "-y"])
r2 = subprocess.run(["conda", "update", "--all", "-y"])
sys.exit(max(r1.returncode, r2.returncode))
"#
    );
    vec!["-c".into(), script]
}

fn ollama_models_upgrade_args(timeout_sec: u64) -> Vec<String> {
    // Advisory + bounded. Refreshing local models is best-effort:
    //  - locally-built (Modelfile) models have no registry source -> pull fails
    //    fast; we record and move on instead of erroring the run.
    //  - oversized models are skipped to avoid multi-GB re-pulls that blow the
    //    time budget; a short per-model timeout caps any single hang.
    //  - never fails the run (exit 0); reports what changed/was skipped.
    let script = format!(
        r#"
import re, subprocess, sys
MAX_GB = 20.0                       # skip models larger than this
per_model = min({timeout_sec}, 300) # hard cap per pull
try:
    r = subprocess.run(["ollama", "list"], capture_output=True, text=True, timeout=min({timeout_sec}, 60))
except Exception as e:
    print(f"ollama list failed: {{e}}")
    sys.exit(0)
print(r.stdout.strip())

def size_gb(parts):
    # SIZE is two columns like "24 GB" / "5.2 GB" near the end
    for i in range(len(parts) - 1):
        if re.fullmatch(r"[0-9.]+", parts[i]) and parts[i + 1].upper() in ("GB", "MB", "KB", "B"):
            v = float(parts[i])
            u = parts[i + 1].upper()
            return v if u == "GB" else v / 1024 if u == "MB" else v / 1048576 if u == "KB" else v / 1073741824
    return None

models = []
for line in r.stdout.strip().splitlines()[1:]:
    parts = line.split()
    if parts:
        models.append((parts[0], size_gb(parts)))
if not models:
    print("No Ollama models found.")
    sys.exit(0)

updated, unchanged, skipped = [], [], []
for name, gb in models:
    if gb is not None and gb > MAX_GB:
        print(f"Skipping large model ({{gb:.0f}} GB > {{MAX_GB:.0f}} GB): {{name}}")
        skipped.append(name)
        continue
    print(f"Pulling Ollama model: {{name}}")
    try:
        rc = subprocess.run(["ollama", "pull", name], timeout=per_model).returncode
        (updated if rc == 0 else unchanged).append(name)
    except subprocess.TimeoutExpired:
        print(f"  timed out after {{per_model}}s; left unchanged: {{name}}")
        unchanged.append(name)
    except Exception as e:
        print(f"  {{name}}: {{e}}")
        unchanged.append(name)

print(f"Ollama: {{len(updated)}} refreshed, {{len(unchanged)}} unchanged, {{len(skipped)}} skipped (too large).")
if unchanged:
    print(f"  unchanged (local/unavailable): {{', '.join(unchanged)}}")
sys.exit(0)
"#
    );
    vec!["-c".into(), script]
}

fn claude_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
if shutil.which("winget"):
    r = subprocess.run(
        ["winget", "upgrade", "--id", "Anthropic.Claude", "--exact",
         "--include-unknown", "--disable-interactivity",
         "--accept-package-agreements", "--accept-source-agreements", "--silent"],
        capture_output=True, text=True
    )
    out = (r.stdout + r.stderr).strip()
    if "No installed package found" not in out:
        if out:
            print(out)
        print("claude checked via winget id Anthropic.Claude.")
        sys.exit(0)
r2 = subprocess.run(["claude", "update"])
sys.exit(r2.returncode)
"#;
    vec!["-c".into(), script.into()]
}

fn codex_upgrade_args() -> Vec<String> {
    let script = r#"
import shutil, subprocess, sys
if not shutil.which("winget"):
    print("winget not available; skipping codex upgrade.")
    sys.exit(0)
# Kill codex processes that lock the exe
procs = subprocess.run(["tasklist", "/FI", "IMAGENAME eq codex-x86_64-pc-windows-msvc.exe", "/NH"],
                        capture_output=True, text=True, errors="ignore")
if "codex" in procs.stdout.lower():
    subprocess.run(["taskkill", "/F", "/IM", "codex-x86_64-pc-windows-msvc.exe"], capture_output=True)
    import time; time.sleep(2)
r = subprocess.run(
    ["winget", "upgrade", "--id", "OpenAI.Codex", "--exact", "--source", "winget",
     "--include-unknown", "--disable-interactivity",
     "--accept-package-agreements", "--accept-source-agreements", "--silent", "--force"],
    capture_output=True, text=True
)
out = (r.stdout + r.stderr).strip()
if out:
    print(out)
# winget exit codes arrive unsigned on Windows; normalize to signed int32
rc = r.returncode
if rc and rc > 0x7FFFFFFF:
    rc -= 0x100000000
low = out.lower()
benign = rc in (0, -1978335188, -1978335189, -1978335212) or \
    "no available upgrade" in low or "no newer package" in low or \
    "no installed package found" in low
if benign:
    print("codex checked via winget id OpenAI.Codex.")
    sys.exit(0)
sys.exit(rc)
"#;
    vec!["-c".into(), script.into()]
}

fn advisory_script(bin: &str, msg: &str) -> Vec<String> {
    let script = format!(
        r#"import shutil, sys
p = shutil.which("{bin}")
if p:
    print(f"{bin}: {{p}}")
print("{msg}")
sys.exit(0)
"#
    );
    vec!["-c".into(), script]
}

fn cleanup_args(days: u32, deep: bool, skip_destructive: bool) -> Vec<String> {
    let script = format!(
        r#"
import os, sys, time, pathlib, platform
days = {days}
deep = {deep_py}
skip_destructive = {skip_py}

cutoff = time.time() - days * 86400
temp_dirs = []
t = os.environ.get("TEMP") or os.environ.get("TMP")
if t:
    temp_dirs.append(t)
if platform.system() == "Windows":
    temp_dirs.append(r"C:\Windows\Temp")

for td in temp_dirs:
    p = pathlib.Path(td)
    if not p.is_dir():
        continue
    if skip_destructive:
        print(f"Skipping temp cleanup (--skip-destructive): {{td}}")
        continue
    print(f"Cleaning temp files older than {{days}} day(s): {{td}}")
    cleaned = 0
    for item in p.iterdir():
        try:
            st = item.stat()
            if st.st_mtime < cutoff and "WinGet" not in str(item):
                if item.is_dir():
                    import shutil
                    shutil.rmtree(item, ignore_errors=True)
                else:
                    item.unlink(missing_ok=True)
                cleaned += 1
        except Exception:
            pass
    print(f"  Cleaned {{cleaned}} item(s)")

if platform.system() == "Windows":
    import subprocess
    subprocess.run(["ipconfig", "/flushdns"], capture_output=True)
    print("DNS cache flushed.")
    if not skip_destructive:
        subprocess.run(["powershell", "-NoProfile", "-Command", "Clear-RecycleBin -Force -ErrorAction SilentlyContinue"], capture_output=True)
        print("Recycle bin emptied.")
    if deep and not skip_destructive:
        print("Running DISM component cleanup (this may take a while)...")
        subprocess.run(["DISM.exe", "/Online", "/Cleanup-Image", "/StartComponentCleanup"])

# Stale binary scan
print("Checking for orphaned binaries in PATH...")
path_dirs = [p for p in os.environ.get("PATH", "").split(os.pathsep) if p and os.path.isdir(p)]
managed = ["scoop\\\\apps", "chocolatey\\\\lib", "Microsoft\\\\WinGet", "pipx\\\\venvs",
           "Python\\\\Scripts", ".cargo\\\\bin", "node_modules\\\\.bin", ".dotnet\\\\tools",
           "mise", "uv\\\\bin", "volta\\\\bin"]
exclude_pfx = ["api-ms-win-", "ext-ms-win-", "concrt", "msvcp", "vcruntime", "msvcrt"]
orphans = 0
for d in path_dirs:
    for exe in pathlib.Path(d).glob("*.exe"):
        is_managed = any(m in str(exe) for m in managed)
        is_sys = any(x in str(exe).lower() for x in ["\\\\system32", "\\\\system\\\\", "\\\\windows\\\\"])
        is_excl = any(exe.name.startswith(p) for p in exclude_pfx)
        if not is_managed and not is_sys and not is_excl:
            try:
                if os.path.getmtime(exe) < cutoff:
                    orphans += 1
            except Exception:
                pass
if orphans:
    print(f"Stale binary scan: {{orphans}} orphaned .exe(s) older than {{days}} day(s) found in PATH.")
else:
    print("Stale binary scan: no orphaned binaries found.")
sys.exit(0)
"#,
        deep_py = if deep { "True" } else { "False" },
        skip_py = if skip_destructive { "True" } else { "False" },
    );
    vec!["-c".into(), script]
}

fn self_update_check_args() -> Vec<String> {
    let script = r#"
import sys, urllib.request, json
try:
    print("Checking for script updates...")
    with urllib.request.urlopen(
        "https://api.github.com/repos/YoshKoz/updateEverything/releases/latest", timeout=15
    ) as resp:
        data = json.loads(resp.read())
    latest = data["tag_name"].lstrip("v").strip()
    current = "7.0.0"
    print(f"Current: {current} | Latest: {latest}")
    if latest and latest != current:
        print(f"Update available: {current} → {latest}")
        print(f"Download from: https://github.com/YoshKoz/updateEverything/releases/tag/v{latest}")
    else:
        print("Already at latest version.")
except Exception as e:
    msg = str(e)
    if "404" in msg or "Not Found" in msg:
        print("Self-update check skipped: no releases found on GitHub.")
    else:
        print(f"Self-update check skipped: {e}")
sys.exit(0)
"#;
    vec!["-c".into(), script.into()]
}

fn winget_git_script() -> &'static str {
    r#"
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Output 'Git upgrade skipped: requires elevation.'; exit 0 }
$gitDir = 'C:\Program Files\Git'
if (-not (Test-Path (Join-Path $gitDir 'bin\bash.exe'))) { Write-Output 'Git for Windows not found; skipping.'; exit 0 }
$list = winget list --id Git.Git --exact --upgrade-available --accept-source-agreements --disable-interactivity 2>&1 | Out-String
if ($list -notmatch 'Git\.Git') { Write-Output 'Git: no upgrade available.'; exit 0 }
Write-Output 'Git upgrade available. Stopping bash.exe and parking it (statusline respawns it)...'
$bashPaths = @((Join-Path $gitDir 'usr\bin\bash.exe'), (Join-Path $gitDir 'bin\bash.exe'))
Get-Process | Where-Object { $_.Path -in $bashPaths } | Stop-Process -Force -ErrorAction SilentlyContinue
$renamed = @()
foreach ($p in $bashPaths) {
    if (Test-Path $p) {
        try { Rename-Item $p ($p + '.ue-hold') -Force; $renamed += $p }
        catch { Write-Output "rename failed: $p -- $($_.Exception.Message)" }
    }
}
Get-Process | Where-Object { $_.Path -in $bashPaths } | Stop-Process -Force -ErrorAction SilentlyContinue
winget upgrade --id Git.Git --exact --source winget --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
$code = $LASTEXITCODE
foreach ($p in $renamed) {
    if (Test-Path $p) { Remove-Item ($p + '.ue-hold') -Force -ErrorAction SilentlyContinue }
    else { Rename-Item ($p + '.ue-hold') $p -Force -ErrorAction SilentlyContinue }
}
if ($code -eq 0) { Write-Output 'Git upgraded.' } else { Write-Output "Git upgrade exit: $code" }
exit $code
"#
}

fn winget_userscope_script(skip_packages: &[String]) -> String {
    // Same per-package loop as the main pass; `--all` installs nothing when any listed
    // package needs an install location.
    let loop_script = winget_per_package_script(skip_packages);
    format!(
        r#"
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {{ Write-Output 'Not elevated: user-scope packages already covered by main winget pass.'; exit 0 }}
$dir = Join-Path $env:TEMP ('ue-userscope-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$ps1File = Join-Path $dir 'run.ps1'
$innerFile = Join-Path $dir 'upgrade.ps1'
$log = Join-Path $dir 'run.log'
# Inner script owns the exit code; run.ps1 only invokes it and records the result,
# so the loop's own `exit` cannot swallow the DONE marker the outer wait polls for.
$body = @'
{loop_script}
'@
Set-Content -Path $innerFile -Value $body -Encoding utf8
@(
    ('pwsh -NoProfile -ExecutionPolicy Bypass -File "' + $innerFile + '" *> "' + $log + '"'),
    ('"exit: $LASTEXITCODE" | Add-Content -Path "' + $log + '"'),
    ('"DONE" | Add-Content -Path "' + $log + '"')
) | Set-Content -Path $ps1File -Encoding utf8
runas /trustlevel:0x20000 "pwsh -NoProfile -ExecutionPolicy Bypass -File $ps1File" | Out-Null
$deadline = (Get-Date).AddSeconds(900)
while ((Get-Date) -lt $deadline) {{
    if ((Test-Path $log) -and (Select-String -Path $log -Pattern '^DONE' -Quiet -ErrorAction SilentlyContinue)) {{ break }}
    Start-Sleep 5
}}
if (-not (Test-Path $log)) {{ Write-Output 'De-elevated winget produced no log (runas launch failed?).'; exit 1 }}
Get-Content $log | Select-Object -Last 40
$m = Select-String -Path $log -Pattern '^exit: (-?\d+)' | Select-Object -Last 1
$code = if ($m) {{ [int]$m.Matches[0].Groups[1].Value }} else {{ 1 }}
Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
if ($code -in 0, -1978335188, -1978335189, -1978335212) {{ exit 0 }}
exit $code
"#
    )
}

fn windows_update_script() -> &'static str {
    "$s=New-Object -ComObject Microsoft.Update.Session;\
     $r=$s.CreateUpdateSearcher().Search(\"IsInstalled=0 and Type='Software'\");\
     if($r.Updates.Count-eq 0){Write-Output 'Windows: no updates available.';return};\
     Write-Output \"Found $($r.Updates.Count) update(s). Downloading...\";\
     $d=New-Object -ComObject Microsoft.Update.Downloader;\
     $d.Updates=$r.Updates;$d.Download();\
     $i=New-Object -ComObject Microsoft.Update.Installer;\
     $i.Updates=$r.Updates;$ir=$i.Install();\
     Write-Output \"Windows Update: $($r.Updates.Count) update(s) installed. Reboot required: $($ir.RebootRequired)\""
}

fn msys2_script() -> &'static str {
    "$bash=@('C:\\msys64\\usr\\bin\\bash.exe','C:\\tools\\msys64\\usr\\bin\\bash.exe',\
     \"$env:SystemDrive\\msys64\\usr\\bin\\bash.exe\")|Where-Object{Test-Path $_}|Select-Object -First 1;\
     if(-not $bash){Write-Output 'MSYS2 not installed; skipping.';return};\
     Write-Output \"Updating MSYS2 via $bash\";\
     & $bash -lc 'if [ -f /var/lib/pacman/db.lck ] && ! pgrep -x pacman >/dev/null 2>&1; then echo \"removing stale pacman lock\"; rm -f /var/lib/pacman/db.lck; fi; pacman -Syu --noconfirm';\
     if($LASTEXITCODE-ne 0){Write-Output \"MSYS2 pacman exit $LASTEXITCODE\"}"
}

fn wsl_distros_script() -> &'static str {
    // NOTE: kept in parity with updatescript.ps1 'wsl-distros' task.
    // Every package manager is gated behind `sudo -n true` so a distro whose
    // user lacks passwordless sudo (e.g. Arch) is skipped instead of hanging
    // forever on a password prompt under -NonInteractive.
    r#"$prev=[Console]::OutputEncoding
[Console]::OutputEncoding=[System.Text.Encoding]::Unicode
$raw=wsl -l -q 2>$null
[Console]::OutputEncoding=$prev
$benign='(wsl2\.localhostForwarding setting has no effect|wsl: An internal error occurred\.|CreateInstance/CreateVm/ConfigureNetworking/0x8007054f|wsl: Failed to configure network|wsl: Failed to start the systemd user session)'
$distros=@($raw|ForEach-Object{($_ -replace "`0",'').Trim()}|Where-Object{$_ -and $_ -notmatch 'docker-desktop' -and $_ -notmatch $benign}|Sort-Object -Unique)
if($distros.Count -eq 0){Write-Output 'No WSL distros found.';return}
$linuxScript=@'
set -u

resolve_any() {
  for host in "$@"; do
    if command -v getent >/dev/null 2>&1 && getent hosts "$host" >/dev/null 2>&1; then return 0; fi
    if command -v nslookup >/dev/null 2>&1 && nslookup "$host" >/dev/null 2>&1; then return 0; fi
    if command -v ping >/dev/null 2>&1 && ping -c 1 -W 2 "$host" >/dev/null 2>&1; then return 0; fi
  done
  return 1
}

if command -v apt-get >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping apt-get: sudo requires a password"
    exit 0
  fi
  if ! resolve_any archive.ubuntu.com security.ubuntu.com; then
    echo "Skipping apt-get: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -o Acquire::Retries=2 update && sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y -o Dpkg::Options::=--force-confdef -o Dpkg::Options::=--force-confold full-upgrade && sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y autoremove
elif command -v pacman >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping pacman: sudo requires a password"
    exit 0
  fi
  if ! resolve_any archlinux.org geo.mirror.pkgbuild.com; then
    echo "Skipping pacman: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n pacman -Syu --noconfirm --needed
elif command -v dnf >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping dnf: sudo requires a password"
    exit 0
  fi
  sudo -n dnf -y upgrade
elif command -v zypper >/dev/null 2>&1; then
  if ! sudo -n true >/dev/null 2>&1; then
    echo "Skipping zypper: sudo requires a password"
    exit 0
  fi
  if ! resolve_any download.opensuse.org mirrors.opensuse.org; then
    echo "Skipping zypper: WSL DNS/network is unavailable"
    exit 0
  fi
  sudo -n zypper --non-interactive refresh && sudo -n zypper --non-interactive update
else
  echo "No supported Linux package manager found"
fi
'@
foreach($d in $distros){
  Write-Output "Updating WSL distro: $d"
  wsl --distribution $d --exec sh -lc $linuxScript
}"#
}

fn powershell_modules_script() -> &'static str {
    "if(Get-Command Update-PSResource -EA SilentlyContinue){\
     $mods=@(Get-PSResource -Name '*' -EA SilentlyContinue);\
     if($mods.Count-eq 0){Write-Output 'No installed PSResource modules found.';return};\
     $failed=@();\
     foreach($m in ($mods|Sort-Object Name -Unique)){\
     try{Update-PSResource -Name $m.Name -AcceptLicense -EA Stop|Out-String|Write-Output}\
     catch{Write-Output \"Not updated: $($m.Name) — $($_.Exception.Message)\";$failed+=$m.Name}};\
     if($failed){Write-Output \"Left unchanged: $($failed -join ', ')\"}\
     }elseif(Get-Command Update-Module -EA SilentlyContinue){\
     $mods=@(Get-InstalledModule -EA SilentlyContinue);\
     if($mods.Count-eq 0){Write-Output 'No installed PowerShellGet modules found.';return};\
     $failed=@();\
     foreach($m in $mods){\
     try{Update-Module -Name $m.Name -Force -EA Stop}\
     catch{Write-Output \"Not updated: $($m.Name) — $($_.Exception.Message)\";$failed+=$m.Name}};\
     if($failed){Write-Output \"Left unchanged: $($failed -join ', ')\"}\
     }else{Write-Output 'No PowerShell module updater found.'}"
}

fn windows_features_script(features: &[String]) -> String {
    if features.is_empty() {
        return "Write-Output 'No WindowsOptionalFeatures configured in update-config.json.'"
            .into();
    }
    let list = features
        .iter()
        .map(|f| format!("'{}'", f.replace('\'', "''")))
        .collect::<Vec<_>>()
        .join(",");
    format!(
        "$features=@({list});\
         $enabled=0;$skipped=0;\
         foreach($f in $features){{\
         $state=Get-WindowsOptionalFeature -Online -FeatureName $f -EA SilentlyContinue;\
         if($state -and $state.State-ne 'Enabled'){{\
         Write-Output \"Enabling: $f\";\
         Enable-WindowsOptionalFeature -Online -FeatureName $f -All -LimitAccess -EA SilentlyContinue|Out-Null;\
         $enabled++\
         }}else{{$skipped++}}}}\
         Write-Output \"Windows Features: $enabled enabled, $skipped already present.\""
    )
}

fn appx_repair_script() -> &'static str {
    "if(-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){\
     Write-Output 'AppX repair skipped: requires elevation (Get-AppxPackage -AllUsers needs admin).';return};\
     $pkgs=@(Get-AppxPackage -AllUsers -EA SilentlyContinue|\
     Where-Object{$_.SignatureKind-eq 'Store'-and-not $_.IsFramework-and\
     $_.Name-match '(Microsoft\\.WindowsStore|Microsoft\\.Store|Microsoft\\.WindowsCalculator|\
     Microsoft\\.Windows\\.Photos|Microsoft\\.Windows\\.Camera|Microsoft\\.People|\
     Microsoft\\.MSPaint|Microsoft\\.ScreenSketch|Microsoft\\.WindowsNotepad|\
     Microsoft\\.WindowsTerminal)'});\
     $repaired=0;\
     foreach($p in $pkgs){\
     try{$m=Get-AppxPackageManifest -Package $p -EA SilentlyContinue;\
     if($m){Add-AppxPackage -Register -DisableDevelopmentMode -EA SilentlyContinue \"$($p.InstallLocation)\\AppxManifest.xml\" *>$null;$repaired++}}\
     catch{Write-Output \"AppX repair failed for $($p.Name): $($_.Exception.Message)\"}};\
     Write-Output \"AppX re-registration: $repaired package(s) re-registered.\""
}

// ─── Task filtering ───────────────────────────────────────────────────────────

fn mark_missing(mut task: Task) -> Task {
    // Only mark missing if no explicit skip_reason already set (e.g. disabled by flag)
    if task.skip_reason.is_none() && !command_exists(task.requires) {
        task.skip_reason = Some(format!("missing command: {}", task.requires));
    }
    task
}

fn filter_tasks(
    tasks: Vec<Task>,
    cli: &Cli,
    config: &Config,
    prev: Option<&PrevSummary>,
) -> Vec<Task> {
    let only = normalize_set(&cli.only);
    let skip = normalize_set(&cli.skip);
    let config_skip = normalize_set(&config.skip_managers);

    // Matches PS1 FastModeSkip
    let fast_skip = normalize_slice(&[
        "chocolatey",
        "wsl-distros",
        "npm",
        "pnpm",
        "yarn",
        "bun",
        "deno",
        "rustup",
        "cargo",
        "go",
        "pip",
        "pip-health",
        "pipx",
        "uv",
        "uv-tools",
        "poetry",
        "composer",
        "ruby-gems",
        "flutter",
        "juliaup",
        "oh-my-posh",
        "yt-dlp",
        "volta",
        "fnm",
        "dotnet-tools",
        "dotnet-workloads",
        "vscode-extensions",
        "powershell-modules",
        "powershell-help",
        "uv-python",
        "ollama-models",
        "vcpkg",
        "conda",
        "gcloud",
        "az",
        "aws",
        "terraform",
        "pulumi",
        "kubectl",
        "helm",
        "hugo",
        "opentofu",
        "starship",
        "zoxide",
        "gitleaks",
        "trivy",
        "packer",
        "nvm",
        "devcontainer",
        "cross-manager",
        "mise-upgrade",
        "tldr",
    ]);

    // Matches PS1 UltraFastSkip
    let ultra_skip = normalize_slice(&[
        "windows-update",
        "store-apps",
        "wsl",
        "wsl-distros",
        "defender",
        "cleanup",
        "winget",
        "winget-source",
        "winget-userscope",
        "scoop",
    ]);

    // Profile-based skip presets
    let profile_skip: BTreeSet<String> = match cli.profile.as_deref() {
        Some("minimal") => normalize_slice(&[
            "vcpkg",
            "conda",
            "gcloud",
            "az",
            "aws",
            "terraform",
            "pulumi",
            "kubectl",
            "helm",
            "hugo",
            "opentofu",
            "starship",
            "gitleaks",
            "trivy",
            "packer",
            "nvm",
            "devcontainer",
        ]),
        Some("work") => BTreeSet::new(),
        Some("personal") => BTreeSet::new(),
        Some("gaming") => normalize_slice(&[
            "vcpkg",
            "conda",
            "gcloud",
            "az",
            "aws",
            "terraform",
            "pulumi",
            "kubectl",
            "helm",
            "hugo",
            "opentofu",
            "gitleaks",
            "trivy",
            "packer",
            "nvm",
            "devcontainer",
            "powershell-modules",
        ]),
        _ => BTreeSet::new(),
    };

    let prev_succeeded: BTreeSet<String> = if cli.since_hours > 0.0 {
        if let Some(p) = prev {
            p.results
                .iter()
                .filter(|r| r.status == "Succeeded")
                .map(|r| normalize(&r.id))
                .collect()
        } else {
            BTreeSet::new()
        }
    } else {
        BTreeSet::new()
    };

    tasks
        .into_iter()
        .filter(|task| only.is_empty() || matches_any(task, &only))
        .map(|mut task| {
            if task.skip_reason.is_some() {
                return task;
            }
            let id = normalize(task.id);
            if skip.contains(&id) || config_skip.contains(&id) {
                task.skip_reason = Some("filtered by skip".to_string());
            } else if profile_skip.contains(&id) {
                task.skip_reason = Some(format!(
                    "skipped by --profile {}",
                    cli.profile.as_deref().unwrap_or("")
                ));
            } else if cli.fast && fast_skip.contains(&id) {
                task.skip_reason = Some("filtered by fast mode".to_string());
            } else if cli.ultra_fast && (fast_skip.contains(&id) || ultra_skip.contains(&id)) {
                task.skip_reason = Some("filtered by ultra-fast mode".to_string());
            } else if !prev_succeeded.is_empty() && prev_succeeded.contains(&id) {
                task.skip_reason = Some(format!(
                    "succeeded within last {:.0}h (--since-hours)",
                    cli.since_hours
                ));
            }
            task
        })
        .collect()
}

// ─── State / config I/O ──────────────────────────────────────────────────────

fn load_prev_summary(cli: &Cli, repo_root: &Path) -> Option<PrevSummary> {
    let candidates: Vec<PathBuf> = {
        let mut v = vec![];
        if let Some(ref dir) = cli.state_dir {
            v.push(dir.join("last-run.json"));
        }
        if let Some(local) = env::var_os("LOCALAPPDATA") {
            v.push(
                PathBuf::from(local)
                    .join("Update-Everything")
                    .join("last-run.json"),
            );
        }
        v.push(repo_root.join("staging").join("rust-run-summary.json"));
        v
    };

    for path in candidates {
        if !path.exists() {
            continue;
        }
        if cli.since_hours > 0.0
            && let Ok(meta) = fs::metadata(&path)
            && let Ok(modified) = meta.modified()
            && let Ok(age) = modified.elapsed()
            && age.as_secs_f64() / 3600.0 > cli.since_hours
        {
            continue;
        }
        if let Ok(text) = fs::read_to_string(&path)
            && let Ok(summary) = serde_json::from_str::<PrevSummary>(&text)
        {
            return Some(summary);
        }
    }
    None
}

fn status_name(status: ExitStatus) -> String {
    if status.success() {
        "Succeeded".to_string()
    } else {
        "Failed".to_string()
    }
}

fn load_config(path: &Path) -> Result<Config> {
    if !path.exists() {
        // Deserialize an empty object instead of Config::default() so the
        // serde field defaults apply (e.g. temp_cleanup_days = 7). A raw
        // Default would give temp_cleanup_days = 0, which makes the cleanup
        // task delete *all* temp files regardless of age.
        return Ok(serde_json::from_str("{}").expect("empty config is valid"));
    }
    let text = fs::read_to_string(path)
        .with_context(|| format!("failed to read config {}", path.display()))?;
    let config = serde_json::from_str(&text)
        .with_context(|| format!("failed to parse config {}", path.display()))?;
    Ok(config)
}

fn write_summary(path: &Path, summary: &RunSummary) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed to create {}", parent.display()))?;
    }
    let json = serde_json::to_string_pretty(summary)?;
    fs::write(path, json).with_context(|| format!("failed to write {}", path.display()))?;
    Ok(())
}

// ─── Output ──────────────────────────────────────────────────────────────────

fn print_task_list(tasks: &[Task]) {
    println!("{:<26} {:<20} State", "Task", "Category");
    println!("{:<26} {:<20} -----", "----", "--------");
    for task in tasks {
        let state = task.skip_reason.as_deref().unwrap_or("planned");
        println!("{:<26} {:<20} {}", task.id, task.category, state);
    }
}

fn print_summary(summary: &RunSummary) {
    let mut counts = BTreeMap::<&str, usize>::new();
    for result in &summary.results {
        *counts.entry(&result.status).or_default() += 1;
    }

    println!();
    println!("{:<26} {:<12} {:<8} Exit", "Task", "Status", "Time(s)");
    println!("{:<26} {:<12} {:<8} ----", "----", "------", "-------");
    // Skipped rows are almost all "tool not installed"; they drown out the tasks
    // that actually ran. Counted below, just not listed.
    for r in summary.results.iter().filter(|r| r.status != "Skipped") {
        let exit = r
            .exit_code
            .map(|c| c.to_string())
            .unwrap_or_else(|| "-".to_string());
        let secs = r.duration_ms as f64 / 1000.0;
        println!(
            "{:<26} {:<12} {:<8} {}",
            r.id,
            r.status,
            format!("{:.1}", secs),
            exit
        );
    }
    println!();

    let total = summary.results.len();
    let succeeded = counts.get("Succeeded").copied().unwrap_or_default();
    let failed = counts.get("Failed").copied().unwrap_or_default();
    let timed_out = counts.get("TimedOut").copied().unwrap_or_default();
    let skipped = counts.get("Skipped").copied().unwrap_or_default();
    let dry = counts.get("DryRun").copied().unwrap_or_default();
    let total_secs = summary.duration_ms as f64 / 1000.0;
    println!(
        "done  total={total}  succeeded={succeeded}  failed={failed}  timed-out={timed_out}  skipped={skipped}  dry={dry}  duration={total_secs:.1}s"
    );
}

// ─── Command probing ─────────────────────────────────────────────────────────

fn command_exists(name: &str) -> bool {
    resolve_command_path(name).is_some()
}

fn resolve_command_path(name: &str) -> Option<PathBuf> {
    if Path::new(name).components().count() > 1 {
        let p = Path::new(name);
        return if p.exists() { Some(p.to_path_buf()) } else { None };
    }

    let path = env::var_os("PATH")?;

    let extensions = if cfg!(windows) {
        env::var_os("PATHEXT")
            .map(|value| {
                env::split_paths(&value)
                    .filter_map(|path| path.into_os_string().into_string().ok())
                    .collect::<Vec<_>>()
            })
            .filter(|items| !items.is_empty())
            .unwrap_or_else(|| vec![".exe".to_string(), ".cmd".to_string(), ".bat".to_string(), ".ps1".to_string()])
    } else {
        vec!["".to_string()]
    };

    for dir in env::split_paths(&path) {
        if is_foreign_windows_mount(&dir) {
            continue;
        }
        if cfg!(windows) {
            for ext in &extensions {
                let candidate = dir.join(format!("{name}{ext}"));
                if candidate.is_file() {
                    return Some(candidate);
                }
            }
        } else if dir.join(name).is_file() {
            return Some(dir.join(name));
        }
    }

    None
}

/// Build a Command that can actually spawn on Windows. Bare `Command::new("yarn")`
/// fails when the target is yarn.cmd / composer.ps1 / gem.bat: CreateProcess does
/// no PATHEXT search, so the spawn errors and the task was silently marked
/// Skipped. Resolve the full path first; .bat/.cmd with a visible extension go
/// through std's safe cmd.exe handling, .ps1 needs an explicit PowerShell host.
fn new_task_command(program: &str) -> Command {
    if !cfg!(windows) {
        return Command::new(program);
    }
    let Some(path) = resolve_command_path(program) else {
        return Command::new(program);
    };
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase());
    match ext.as_deref() {
        Some("ps1") => {
            let shell = if resolve_command_path("pwsh").is_some() {
                "pwsh"
            } else {
                "powershell"
            };
            let mut c = Command::new(shell);
            c.args(["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File"])
                .arg(path);
            c
        }
        _ => Command::new(path),
    }
}

/// When running under WSL, skip automounted Windows drives (/mnt/<letter>/...)
/// so Windows-only tools aren't falsely detected as present.
#[cfg(unix)]
fn is_foreign_windows_mount(dir: &Path) -> bool {
    use std::sync::OnceLock;
    static IS_WSL: OnceLock<bool> = OnceLock::new();
    let is_wsl = *IS_WSL.get_or_init(|| {
        env::var_os("WSL_DISTRO_NAME").is_some()
            || env::var_os("WSL_INTEROP").is_some()
            || fs::read_to_string("/proc/version")
                .map(|v| v.to_ascii_lowercase().contains("microsoft"))
                .unwrap_or(false)
    });
    is_wsl && is_windows_drive_mount_path(dir)
}

#[cfg(windows)]
fn is_foreign_windows_mount(_dir: &Path) -> bool {
    false
}

#[allow(dead_code)]
fn is_windows_drive_mount_path(dir: &Path) -> bool {
    let mut comps = dir.components();
    matches!(comps.next(), Some(std::path::Component::RootDir))
        && comps.next().is_some_and(|c| c.as_os_str() == "mnt")
        && comps.next().is_some_and(|c| {
            let s = c.as_os_str().to_string_lossy();
            s.len() == 1 && s.chars().all(|ch| ch.is_ascii_alphabetic())
        })
}

// ─── Utilities ───────────────────────────────────────────────────────────────

fn matches_any(task: &Task, values: &BTreeSet<String>) -> bool {
    values.contains(&normalize(task.id))
        || values.contains(&normalize(task.category))
        || task.tags.iter().any(|tag| values.contains(&normalize(tag)))
}

fn normalize_set(values: &[String]) -> BTreeSet<String> {
    values.iter().map(|value| normalize(value)).collect()
}

fn normalize_slice(values: &[&str]) -> BTreeSet<String> {
    values.iter().map(|value| normalize(value)).collect()
}

fn normalize(value: &str) -> String {
    value.trim().to_ascii_lowercase()
}

fn shell_join(args: &[String]) -> String {
    args.iter()
        .map(|arg| {
            if arg.contains(' ') {
                format!("\"{}\"", arg.replace('"', "\\\""))
            } else {
                arg.clone()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

/// Same as `shell_join` but collapses multi-line arguments to a placeholder.
/// Inline `-c` script bodies are dozens of lines and bury the run/dry echo.
fn shell_join_brief(args: &[String]) -> String {
    let brief: Vec<String> = args
        .iter()
        .map(|arg| {
            if arg.contains('\n') {
                let lines = arg.lines().filter(|l| !l.trim().is_empty()).count();
                format!("<{lines}-line-script>")
            } else {
                arg.clone()
            }
        })
        .collect();
    shell_join(&brief)
}

#[allow(dead_code)]
fn tail<T>(items: Vec<T>, count: usize) -> Vec<T> {
    let len = items.len();
    items.into_iter().skip(len.saturating_sub(count)).collect()
}

/// Keep up to `max` lines: first half + last half when output is long.
/// Ensures both the initial version-table (head) and completion messages (tail) survive.
fn cap_output(lines: Vec<String>, max: usize) -> Vec<String> {
    if lines.len() <= max {
        return lines;
    }
    let head = max / 2;
    let tail_count = max - head;
    let mut result = lines[..head].to_vec();
    result.extend_from_slice(&lines[lines.len() - tail_count..]);
    result
}

fn now_string() -> String {
    let format = format_description!("[year]-[month]-[day]T[hour]:[minute]:[second]");
    OffsetDateTime::now_local()
        .unwrap_or_else(|_| OffsetDateTime::now_utc())
        .format(&format)
        .unwrap_or_else(|_| "unknown".to_string())
}

fn find_repo_root() -> Result<PathBuf> {
    let marker = |dir: &Path| {
        dir.join("update-config.json").exists() || dir.join("updatescript.ps1").exists()
    };

    // 1) Walk up from the current working directory.
    let mut dir = env::current_dir()?;
    loop {
        if marker(&dir) {
            return Ok(dir);
        }
        if !dir.pop() {
            break;
        }
    }

    // 2) Walk up from the executable's own directory. This makes config
    //    discovery work when the exe is invoked from an arbitrary CWD (e.g. the
    //    scheduled task runs from System32), as long as update-config.json sits
    //    next to the deployed binary.
    if let Ok(exe) = env::current_exe() {
        let mut dir = exe.parent().map(Path::to_path_buf);
        while let Some(d) = dir {
            if marker(&d) {
                return Ok(d);
            }
            dir = d.parent().map(Path::to_path_buf);
        }
    }

    env::current_dir().context("failed to resolve current directory")
}

// ─── State dir / process lock / scheduling ───────────────────────────────────

fn get_state_dir(cli: &Cli) -> PathBuf {
    if let Some(ref dir) = cli.state_dir {
        return dir.clone();
    }
    if let Some(local) = env::var_os("LOCALAPPDATA") {
        return PathBuf::from(local).join("Update-Everything");
    }
    env::temp_dir().join("Update-Everything")
}

struct ProcessLock {
    path: PathBuf,
}

impl ProcessLock {
    fn acquire(state_dir: &Path) -> Option<Self> {
        let path = state_dir.join("update-everything.lock");
        let _ = fs::create_dir_all(state_dir);

        // Try an atomic exclusive create first so two instances racing to start
        // can't both observe "no lock" and both write their own PID.
        for _ in 0..2 {
            match fs::OpenOptions::new()
                .write(true)
                .create_new(true)
                .open(&path)
            {
                Ok(mut file) => {
                    use std::io::Write as _;
                    let _ = file.write_all(std::process::id().to_string().as_bytes());
                    return Some(Self { path });
                }
                Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => {
                    if let Ok(text) = fs::read_to_string(&path)
                        && let Ok(pid) = text.trim().parse::<u32>()
                        && is_pid_running(pid)
                    {
                        eprintln!("warn: another instance is already running (PID {pid}); exiting");
                        std::process::exit(5);
                    }
                    // Stale lock (process gone or unparseable) — remove and retry the
                    // exclusive create on the next loop iteration.
                    let _ = fs::remove_file(&path);
                }
                Err(_) => {
                    // Can't create the lock file at all (e.g. permissions); fall back
                    // to running without single-instance protection.
                    return Some(Self { path });
                }
            }
        }
        Some(Self { path })
    }
}

impl Drop for ProcessLock {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

#[cfg(windows)]
fn is_pid_running(pid: u32) -> bool {
    Command::new("tasklist")
        .args(["/FI", &format!("PID eq {pid}"), "/NH", "/FO", "CSV"])
        .output()
        .map(|o| {
            let s = String::from_utf8_lossy(&o.stdout);
            // tasklist CSV: header + one row per match; if only header, no match
            s.lines().skip(1).any(|l| !l.trim().is_empty())
        })
        .unwrap_or(false)
}

#[cfg(unix)]
fn is_pid_running(pid: u32) -> bool {
    Path::new(&format!("/proc/{pid}")).exists()
}

#[cfg(windows)]
fn schedule_task(exe: &Path, schedule_time: Option<&str>) -> Result<()> {
    let task_name = "Update-Everything";
    let exe_str = exe.to_string_lossy().into_owned();

    let args: Vec<&str> = if let Some(time) = schedule_time {
        vec![
            "/Create", "/F", "/TN", task_name, "/TR", &exe_str, "/SC", "DAILY", "/ST", time,
        ]
    } else {
        vec![
            "/Create", "/F", "/TN", task_name, "/TR", &exe_str, "/SC", "ONLOGON",
        ]
    };

    let status = Command::new("schtasks")
        .args(&args)
        .status()
        .context("failed to run schtasks.exe")?;

    if !status.success() {
        anyhow::bail!("schtasks /Create failed (exit {:?})", status.code());
    }

    let trigger = schedule_time
        .map(|t| format!("daily at {t}"))
        .unwrap_or_else(|| "at logon".to_string());
    println!("Registered scheduled task '{task_name}' ({trigger})");
    println!("  exe: {exe_str}");
    Ok(())
}

#[cfg(unix)]
fn schedule_task(_exe: &Path, _schedule_time: Option<&str>) -> Result<()> {
    anyhow::bail!("--schedule is only supported on Windows; use cron on Unix");
}

// ─── What's Changed summary ───────────────────────────────────────────────────

/// True if a string looks like a version: contains a digit and at least one dot.
fn looks_like_version(s: &str) -> bool {
    s.chars().any(|c| c.is_ascii_digit()) && s.contains('.')
}

/// Strip spinner/progress junk chars that winget inserts into lines.
fn strip_progress(s: &str) -> String {
    s.chars()
        .filter(|c| c.is_ascii_graphic() || *c == ' ')
        .collect::<String>()
        .split_whitespace()
        .filter(|t| !matches!(*t, "-" | "\\" | "|" | "/"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn print_update_summary(results: &[TaskSummary]) {
    struct Entry {
        task: String,
        changes: Vec<String>,
    }

    let mut entries: Vec<Entry> = Vec::new();

    for r in results {
        if r.output_tail.is_empty() {
            continue;
        }
        let lines = &r.output_tail;
        let mut changes: Vec<String> = Vec::new();

        match r.id.as_str() {
            "npm" => {
                let mut i = 0;
                while i < lines.len() {
                    if lines[i].contains("Updating npm package:") {
                        let pkg = lines[i]
                            .split("Updating npm package:")
                            .nth(1)
                            .unwrap_or("")
                            .trim()
                            .trim_end_matches("@latest")
                            .to_string();
                        let count = lines
                            .get(i + 1)
                            .filter(|l| l.contains("changed") && l.contains("package"))
                            .and_then(|l| l.split_whitespace().nth(1))
                            .unwrap_or("?")
                            .to_string();
                        changes.push(format!("{pkg}  (+{count} pkg)"));
                    }
                    i += 1;
                }
            }
            "pnpm" => {
                for line in lines {
                    if line.contains("Switching")
                        && line.contains("from v")
                        && line.contains("to v")
                        && let Some(after_from) = line.split("from v").nth(1)
                    {
                        let tool = line
                            .split("Switching")
                            .nth(1)
                            .unwrap_or("")
                            .split_whitespace()
                            .next()
                            .unwrap_or("pnpm");
                        if let Some((old, rest)) = after_from.split_once(" to v") {
                            let new_ver = rest.trim_end_matches('.').trim_end_matches("..").trim();
                            changes.push(format!("{tool}: {old} → {new_ver}"));
                        }
                    }
                }
            }
            "winget" | "winget-batch" | "store-apps" => {
                // winget-batch: parse the version table (Name / Id / Version / Available)
                // winget: parse "[N/M] Upgrading: X" lines filtered for failures
                let mut in_table = false;
                let mut i = 0;
                while i < lines.len() {
                    let clean = strip_progress(&lines[i]);
                    // Detect table header
                    if !in_table
                        && clean.contains("Name")
                        && clean.contains("Id")
                        && clean.contains("Available")
                    {
                        in_table = true;
                        i += 1;
                        continue;
                    }
                    if in_table {
                        // Skip separator
                        if clean.trim().chars().all(|c| c == '-' || c == ' ') {
                            i += 1;
                            continue;
                        }
                        // Stop at end-of-table markers
                        if clean.trim().is_empty()
                            || clean.contains("package(s) are pinned")
                            || clean.contains("The following")
                            || clean.contains("upgrades available")
                            || clean.contains("upgrade available")
                        {
                            in_table = false;
                        } else {
                            // Split into tokens on 2+ spaces
                            let tokens: Vec<&str> = clean
                                .split("  ")
                                .map(|s| s.trim())
                                .filter(|s| !s.is_empty())
                                .collect();
                            if tokens.len() >= 3 {
                                let name = tokens[0];
                                let available = tokens[tokens.len() - 1];
                                let version = tokens[tokens.len() - 2];
                                if looks_like_version(available) || available.starts_with('<') {
                                    changes.push(format!("{name}: {version} → {available}"));
                                }
                            }
                        }
                    }
                    // Also capture "[N/M] Upgrading: X" lines from winget task (not already covered by table)
                    if lines[i].contains("] Upgrading:") && r.id == "winget" {
                        let pkg = lines[i]
                            .split("] Upgrading:")
                            .nth(1)
                            .unwrap_or("")
                            .trim()
                            .split(" (installed")
                            .next()
                            .unwrap_or("")
                            .trim()
                            .to_string();
                        let next = lines.get(i + 1).map_or("", |s| s.as_str());
                        if !next.contains("Not applicable:") && !next.contains("FAILED:") && !pkg.is_empty() {
                            // Only add if not already captured from table
                            if changes.iter().all(|c| !c.starts_with(pkg.as_str())) {
                                changes.push(pkg);
                            }
                        }
                    }
                    i += 1;
                }
            }
            "scoop" => {
                for line in lines {
                    // "Updating 'name' (old -> new)"
                    if (line.contains("Updating '") || line.contains("Updating "))
                        && line.contains("->")
                    {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "chocolatey" => {
                for line in lines {
                    // "Chocolatey upgraded N/M packages" — only if N > 0
                    if line.contains("upgraded") && line.contains("package") {
                        if !line.contains("upgraded 0/") {
                            changes.push(line.trim().to_string());
                        }
                    } else if line.contains(" to ") && looks_like_version(line.split(" to ").next().unwrap_or("").split_whitespace().last().unwrap_or("")) {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "cargo" => {
                for line in lines {
                    // "Updating crate-name v0.1 -> v0.2" or table "name  v0.1  v0.2  Yes"
                    if line.contains("->") && line.chars().any(|c| c.is_ascii_digit()) {
                        let t = line.trim();
                        if t.len() < 120 && !t.starts_with("Package") {
                            changes.push(t.to_string());
                            if changes.len() >= 10 { break; }
                        }
                    }
                }
            }
            "pip" | "uv-tools" | "uv-python" | "pipx" => {
                if r.id == "pipx" {
                    // "upgrading X..."
                    let upgraded: Vec<&str> = lines
                        .iter()
                        .filter(|l| l.starts_with("upgrading ") && l.ends_with("..."))
                        .map(|l| l.trim_start_matches("upgrading ").trim_end_matches("..."))
                        .collect();
                    if !upgraded.is_empty() {
                        changes.push(upgraded.join(", "));
                    }
                } else {
                    // "Successfully installed X-1.2.3 Y-4.5.6"
                    for line in lines {
                        if line.contains("Successfully installed") {
                            let pkgs = line
                                .trim()
                                .trim_start_matches("Successfully installed")
                                .trim();
                            if !pkgs.is_empty() {
                                changes.push(pkgs.to_string());
                                if changes.len() >= 5 { break; }
                            }
                        }
                    }
                    // uv-tools: "Updated X 1.0 -> 1.1"
                    for line in lines {
                        if line.contains("Updated") && line.contains("->") {
                            changes.push(line.trim().to_string());
                            if changes.len() >= 10 { break; }
                        }
                    }
                }
            }
            "poetry" => {
                let mut pkg_count = 0;
                let downgrades = lines.iter().filter(|l| l.contains("- Downgrading")).count();
                for line in lines {
                    if line.contains("Package operations:") {
                        let summary = line.trim().trim_start_matches("Package operations: ");
                        if !summary.contains("0 installs, 0 updates, 0 removals") {
                            if downgrades > 0 {
                                changes.push(format!("{summary} -- {downgrades} DOWNGRADED"));
                            } else {
                                changes.push(summary.to_string());
                            }
                        }
                    } else if pkg_count < 6
                        && (line.contains("- Installing")
                            || line.contains("- Updating")
                            || line.contains("- Downgrading"))
                        && line.contains('(')
                    {
                        changes.push(format!("  {}", line.trim().trim_start_matches("- ")));
                        pkg_count += 1;
                    }
                }
            }
            "mise" | "mise-upgrade" => {
                for line in lines {
                    // "mise python@3.11 -> python@3.12" or "Updated X from v1 to v2"
                    if (line.contains("→") || line.contains("->"))
                        && line.chars().any(|c| c.is_ascii_digit())
                    {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "gh-extensions" => {
                for line in lines {
                    if line.contains("upgraded") || (line.contains("Updated") && line.contains("→")) {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "ruby-gems" => {
                for line in lines {
                    // "Updated X from 1.0 to 2.0" or "Updating X (1.0 -> 2.0)"
                    if (line.contains("Updated") || line.contains("Updating"))
                        && (line.contains("->") || line.contains(" to "))
                        && line.chars().any(|c| c.is_ascii_digit())
                    {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "powershell-modules" => {
                for line in lines {
                    if (line.contains("Install") || line.contains("Update"))
                        && line.chars().any(|c| c.is_ascii_digit())
                    {
                        changes.push(line.trim().to_string());
                        if changes.len() >= 10 { break; }
                    }
                }
            }
            "dotnet-workloads" => {
                let cnt = lines
                    .iter()
                    .filter(|l| l.contains("Updated advertising manifest"))
                    .count();
                if cnt > 0 {
                    changes.push(format!("{cnt} manifests updated"));
                }
            }
            "appx-repair" => {
                for line in lines {
                    if line.contains("re-registered") {
                        changes.push(line.trim().to_string());
                    }
                }
            }
            _ => {
                // Generic: version arrows "X 1.2 -> 1.3" or "X v1 → v2"
                let mut count = 0;
                for line in lines {
                    if (line.contains("->") || line.contains('→'))
                        && line.chars().any(|c| c.is_ascii_digit())
                    {
                        let t = line.trim();
                        if t.len() < 120 && !t.is_empty() {
                            changes.push(t.to_string());
                            count += 1;
                            if count >= 5 { break; }
                        }
                    }
                }
            }
        }

        if !changes.is_empty() {
            entries.push(Entry {
                task: r.id.clone(),
                changes,
            });
        }
    }

    if entries.is_empty() {
        return;
    }

    println!();
    println!("What's Changed:");
    for entry in &entries {
        println!("  {:<18}  {}", entry.task, entry.changes[0]);
        for change in entry.changes.iter().skip(1) {
            println!("  {:<18}  {}", "", change);
        }
    }
}

// ─── Tests ───────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_trims_and_lowercases() {
        assert_eq!(normalize("  WinGet "), "winget");
        assert_eq!(normalize(""), "");
    }

    #[test]
    fn tail_keeps_last_n() {
        assert_eq!(tail(vec![1, 2, 3, 4], 2), vec![3, 4]);
        assert_eq!(tail(vec![1, 2], 5), vec![1, 2]);
        assert_eq!(tail(Vec::<i32>::new(), 3), Vec::<i32>::new());
    }

    #[test]
    fn shell_join_quotes_spaces() {
        let args = vec!["upgrade".to_string(), "My App".to_string()];
        assert_eq!(shell_join(&args), "upgrade \"My App\"");
    }

    #[test]
    fn windows_drive_mount_paths() {
        assert!(is_windows_drive_mount_path(Path::new("/mnt/c")));
        assert!(is_windows_drive_mount_path(Path::new(
            "/mnt/c/Users/yoshi/scoop/shims"
        )));
        assert!(is_windows_drive_mount_path(Path::new("/mnt/D/tools")));
        assert!(!is_windows_drive_mount_path(Path::new("/mnt")));
        assert!(!is_windows_drive_mount_path(Path::new("/mnt/wsl")));
        assert!(!is_windows_drive_mount_path(Path::new("/mnt/cd/x")));
        assert!(!is_windows_drive_mount_path(Path::new("/usr/bin")));
        assert!(!is_windows_drive_mount_path(Path::new("mnt/c")));
    }

    #[test]
    fn winget_args_omit_include_pinned() {
        // Skip packages are excluded via pinning (winget-pin-skip task), not --exclude,
        // and `--include-pinned` must stay absent so pinned packages are left untouched.
        let args = winget_upgrade_args(&["Foo.Bar".to_string(), "Baz.Qux".to_string()]);
        assert_eq!(args[0], "-NoProfile");
        assert_eq!(args[2], "-Command");
        let script = &args[3];
        assert!(!script.contains("--all"));
        assert!(!script.contains("--exclude"));
        assert!(!script.contains("--include-pinned"));
        assert!(script.contains("'Foo.Bar'"));
        assert!(script.contains("'Baz.Qux'"));
    }

    #[test]
    fn pip_args_embed_lowercased_skips() {
        let args = pip_upgrade_args(&["PyLint".to_string()]);
        assert_eq!(args[0], "-c");
        assert!(args[1].contains("\"pylint\""));
    }

    #[test]
    fn matches_any_by_id_category_and_tag() {
        let task = Task::new(
            "rustup",
            "systems-language",
            &["toolchain"],
            "rustup",
            &["update"],
        );
        let by_id = normalize_slice(&["RUSTUP"]);
        let by_cat = normalize_slice(&["systems-language"]);
        let by_tag = normalize_slice(&["toolchain"]);
        let none = normalize_slice(&["python"]);
        assert!(matches_any(&task, &by_id));
        assert!(matches_any(&task, &by_cat));
        assert!(matches_any(&task, &by_tag));
        assert!(!matches_any(&task, &none));
    }

    #[test]
    fn prev_summary_parses_both_casings() {
        let ps1_style = r#"{"Results":[{"Id":"winget","Status":"Succeeded"}]}"#;
        let rust_style = r#"{"results":[{"id":"winget","status":"Succeeded"}]}"#;
        for text in [ps1_style, rust_style] {
            let parsed: PrevSummary = serde_json::from_str(text).unwrap();
            assert_eq!(parsed.results.len(), 1);
            assert_eq!(parsed.results[0].id, "winget");
            assert_eq!(parsed.results[0].status, "Succeeded");
        }
    }

    #[test]
    fn config_parses_pascal_case() {
        let text = r#"{"WingetSkipPackages":["A.B"],"SkipManagers":["scoop"],"PipIgnoreHealthPackages":["pylint"]}"#;
        let config: Config = serde_json::from_str(text).unwrap();
        assert_eq!(config.winget_skip_packages, vec!["A.B"]);
        assert_eq!(config.skip_managers, vec!["scoop"]);
        assert_eq!(config.pip_ignore_health_packages, vec!["pylint"]);
    }

    #[test]
    fn with_skip_if_sets_reason() {
        let task = Task::new("test", "cat", &[], "cmd", &[]).with_skip_if(true, "test reason");
        assert_eq!(task.skip_reason.as_deref(), Some("test reason"));
    }

    #[test]
    fn with_skip_if_false_leaves_none() {
        let task = Task::new("test", "cat", &[], "cmd", &[]).with_skip_if(false, "test reason");
        assert!(task.skip_reason.is_none());
    }

    #[test]
    fn mark_missing_respects_existing_skip_reason() {
        let mut task = Task::new("test", "cat", &[], "nonexistent_binary_xyz", &[]);
        task.skip_reason = Some("opt-in".to_string());
        let task = mark_missing(task);
        assert_eq!(task.skip_reason.as_deref(), Some("opt-in"));
    }

    #[test]
    fn config_new_fields_default() {
        let config: Config = serde_json::from_str("{}").unwrap();
        assert!(config.npm_skip_packages.is_empty());
        assert!(config.windows_optional_features.is_empty());
        assert_eq!(config.temp_cleanup_days, 7);
        assert_eq!(config.log_retention_days, 14);
    }
}
