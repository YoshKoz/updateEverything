use std::{
    collections::{BTreeMap, BTreeSet, HashMap, VecDeque},
    env, fs,
    io::{BufRead, BufReader},
    path::{Path, PathBuf},
    process::{Command, ExitStatus, Stdio},
    sync::{Arc, Mutex, mpsc},
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

    /// Run up to N tasks in parallel (default 1 = serial)
    #[arg(long, default_value_t = 1)]
    jobs: usize,

    /// Exit non-zero if any task failed or timed out
    #[arg(long)]
    ci: bool,

    /// Skip tasks that succeeded within N hours in the last run (0 = disabled)
    #[arg(long, default_value_t = 0.0)]
    since_hours: f64,

    /// Path to the state directory containing last-run.json for --since-hours
    #[arg(long)]
    state_dir: Option<PathBuf>,
}

#[derive(Debug, Default, Deserialize)]
#[serde(rename_all = "PascalCase")]
struct Config {
    #[serde(default)]
    winget_skip_packages: Vec<String>,
    #[serde(default)]
    skip_managers: Vec<String>,
    #[serde(default)]
    pip_skip_packages: Vec<String>,
    /// Alias for pip_skip_packages — matches PS1 PipIgnoreHealthPackages field
    #[serde(default)]
    pip_ignore_health_packages: Vec<String>,
}

#[derive(Clone, Debug)]
struct Task {
    id: &'static str,
    category: &'static str,
    tags: &'static [&'static str],
    command: &'static str,
    args: Vec<String>,
    requires: &'static str,
    /// Named lock acquired before running in parallel mode ("" = no lock)
    resource: &'static str,
    skip_reason: Option<String>,
    /// Per-task timeout; overrides the global --task-timeout-sec when set
    timeout_override: Option<Duration>,
    /// Exit codes that should be treated as Succeeded (e.g. winget partial upgrades)
    acceptable_exit_codes: Vec<i32>,
    /// If true, TimedOut is treated as Succeeded (e.g. fire-and-forget GUI launchers)
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
    let start = Instant::now();
    let started_at = now_string();
    let repo_root = find_repo_root()?;
    let config_path = cli
        .config
        .clone()
        .unwrap_or_else(|| repo_root.join("update-config.json"));
    let config = load_config(&config_path)?;
    let tasks = build_tasks(&config);
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

    if let Some(path) = &cli.json_summary {
        write_summary(path, &summary)?;
        if !cli.quiet {
            println!("summary {}", path.display());
        }
    }

    print_summary(&summary);

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
                    shell_join(&task.args)
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

    if jobs <= 1 {
        for task in to_run {
            let r = run_task_streaming(&task, cli.quiet, timeout, false, &HashMap::new());
            results.push(r);
        }
    } else {
        // Build resource locks: tasks sharing a resource name run serially
        let mut resource_locks: HashMap<String, Arc<Mutex<()>>> = HashMap::new();
        for task in &to_run {
            if !task.resource.is_empty() {
                resource_locks
                    .entry(task.resource.to_string())
                    .or_insert_with(|| Arc::new(Mutex::new(())));
            }
        }
        let resource_locks = Arc::new(resource_locks);

        let queue: Arc<Mutex<VecDeque<Task>>> = Arc::new(Mutex::new(to_run.into_iter().collect()));
        let out: Arc<Mutex<Vec<TaskSummary>>> = Arc::new(Mutex::new(Vec::new()));

        let mut handles = vec![];
        for _ in 0..jobs {
            let queue = Arc::clone(&queue);
            let out = Arc::clone(&out);
            let locks = Arc::clone(&resource_locks);
            let quiet = cli.quiet;
            let h = thread::spawn(move || {
                loop {
                    let task = {
                        let mut q = queue.lock().unwrap();
                        match q.pop_front() {
                            Some(t) => t,
                            None => break,
                        }
                    };
                    let r = run_task_streaming(&task, quiet, timeout, true, &locks);
                    out.lock().unwrap().push(r);
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
            shell_join(&task.args)
        );
    }

    // Acquire resource lock for the duration of this task
    let _resource_guard = if !task.resource.is_empty() {
        resource_locks.get(task.resource).map(|m| m.lock().unwrap())
    } else {
        None
    };

    let timeout = task.timeout_override.unwrap_or(timeout);
    let start = Instant::now();
    let mut command = Command::new(task.command);
    command
        .args(&task.args)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped());
    #[cfg(unix)]
    {
        // Own process group, so a timeout can kill the task's whole process
        // tree (a plain kill() leaves grandchildren running and holding the
        // output pipes open).
        use std::os::unix::process::CommandExt;
        command.process_group(0);
    }
    let mut child = match command.spawn() {
        Ok(child) => child,
        Err(err) => {
            // Treat any spawn failure as Skipped rather than Failed — the command
            // exists on PATH (passed command_exists) but can't be launched (e.g. a
            // stale .cmd/.ps1 shim whose underlying tool was uninstalled, or a
            // .ps1 script that Windows won't execute directly). This keeps the
            // summary clean for tools the user hasn't fully set up.
            return make_summary(
                task,
                "Skipped",
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
        // On Windows a child that spawns a GUI sub-process (e.g.
        // `code --update-extensions` → vscode.exe) passes its inherited pipe
        // handles to the grandchild. The CLI parent exits (child.wait()
        // returns) but the GUI keeps the write-ends open, blocking join()
        // indefinitely. Use a bounded 3-second drain window; if the threads
        // haven't finished by then, detach them and move on.
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
    // On timeout (or drain timeout above) the reader threads may still be
    // blocked on pipes held open by surviving grandchildren; take what was
    // captured so far instead of joining.
    let mut all_lines = lines_out.lock().unwrap().clone();
    all_lines.extend(lines_err.lock().unwrap().iter().cloned());

    let code = exit_status.and_then(|s| s.code());
    let status = if timed_out && task.ok_on_timeout {
        "Succeeded".to_string()
    } else if timed_out {
        "TimedOut".to_string()
    } else if code.map_or(false, |c| task.acceptable_exit_codes.contains(&c)) {
        "Succeeded".to_string()
    } else {
        exit_status
            .map(status_name)
            .unwrap_or_else(|| "Failed".to_string())
    };

    if !quiet {
        println!(
            "done {:<22} {} ({:.1}s)",
            task.id,
            status,
            duration_ms as f64 / 1000.0
        );
    }

    make_summary(task, &status, duration_ms, code, tail(all_lines, 30))
}

/// Kill a timed-out task. On unix the child was started in its own process
/// group, so signal the whole group to take grandchildren down with it.
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

fn build_tasks(config: &Config) -> Vec<Task> {
    // Merge pip_skip and pip_ignore_health_packages (PS1 compat)
    let mut pip_skip = config.pip_skip_packages.clone();
    for pkg in &config.pip_ignore_health_packages {
        if !pip_skip.contains(pkg) {
            pip_skip.push(pkg.clone());
        }
    }

    let winget_upgrade_args = winget_upgrade_args(&config.winget_skip_packages);
    vec![
        Task::new(
            "winget-source",
            "package-manager",
            &["windows", "winget"],
            "winget",
            &["source", "update"],
        )
        .with_resource("winget"),
        // Kill portable-app processes that hold their own exe open — winget
        // can't overwrite a locked file even with --force. cmd.exe is always
        // present on Windows; the task is naturally skipped on Linux/macOS
        // (cmd not in PATH). Single `&` runs exit regardless of taskkill result.
        Task::new(
            "winget-pre",
            "package-manager",
            &["windows", "winget"],
            "cmd",
            &["/c", "taskkill /F /IM codex-x86_64-pc-windows-msvc.exe 2>nul & exit 0"],
        )
        .with_resource("winget"),
        Task::new_vec(
            "winget",
            "package-manager",
            &["windows", "winget"],
            "winget",
            winget_upgrade_args,
        )
        .with_resource("winget"),
        Task::new(
            "scoop",
            "package-manager",
            &["windows", "scoop"],
            "scoop",
            &["update", "*"],
        ),
        Task::new(
            "chocolatey",
            "package-manager",
            &["windows", "choco"],
            "choco",
            &["upgrade", "all", "-y"],
        ),
        // Linux (Arch/WSL) system packages; skipped on Windows where pacman
        // doesn't exist. Needs sudo, so run interactively (or with NOPASSWD).
        Task::new_with_requires(
            "pacman",
            "package-manager",
            &["linux", "arch"],
            "sudo",
            vec![
                "pacman".to_string(),
                "-Syu".to_string(),
                "--noconfirm".to_string(),
            ],
            "pacman",
        ),
        Task::new("npm", "javascript", &["node"], "npm", &["update", "-g"]),
        Task::new("pnpm", "javascript", &["node"], "pnpm", &["self-update"]),
        Task::new(
            "yarn",
            "javascript",
            &["node"],
            "yarn",
            &["global", "upgrade"],
        ),
        Task::new("bun", "javascript", &["node"], "bun", &["upgrade"]),
        Task::new("deno", "javascript", &["node"], "deno", &["upgrade"]),
        Task::new("rustup", "rust", &["rust"], "rustup", &["update"]),
        Task::new_with_requires(
            "cargo",
            "rust",
            &["rust"],
            "cargo",
            vec!["install-update".to_string(), "-a".to_string()],
            "cargo-install-update",
        ),
        Task::new("pipx", "python", &["python"], "pipx", &["upgrade-all"]),
        Task::new_vec(
            "pip",
            "python",
            &["python"],
            "python",
            pip_upgrade_args(&pip_skip),
        ),
        Task::new_vec(
            "uv",
            "python",
            &["python"],
            "python",
            vec![
                "-c".to_string(),
                // Skip self-update when uv is pip-managed (the pip task already
                // upgrades it). Only run self-update for standalone installs
                // (typically ~/.local/bin/uv or %LOCALAPPDATA%\uv\bin\uv.exe).
                [
                    "import shutil, subprocess, sys",
                    r#"p = (shutil.which("uv") or "").replace("\\", "/").lower()"#,
                    r#"if "/python" in p or "/scripts/" in p:"#,
                    r#"    print("uv is pip-managed; update handled by pip task")"#,
                    "    sys.exit(0)",
                    r#"r = subprocess.run(["uv", "self", "update"])"#,
                    "sys.exit(r.returncode)",
                ]
                .join("\n"),
            ],
        ),
        Task::new(
            "uv-tools",
            "python",
            &["python"],
            "uv",
            &["tool", "upgrade", "--all"],
        ),
        Task::new(
            "poetry",
            "python",
            &["python"],
            "poetry",
            &["self", "update"],
        ),
        Task::new("composer", "php", &["php"], "composer", &["self-update"]),
        Task::new("ruby-gems", "ruby", &["ruby"], "gem", &["update"]),
        Task::new("flutter", "flutter", &["flutter"], "flutter", &["upgrade"]),
        Task::new("juliaup", "julia", &["julia"], "juliaup", &["update"]),
        Task::new(
            "dotnet-workloads",
            "dotnet",
            &["dotnet"],
            "dotnet",
            &["workload", "update"],
        ),
        Task::new_with_requires(
            "dotnet-tools",
            "dotnet",
            &["dotnet"],
            "python",
            dotnet_tools_upgrade_args(),
            "dotnet",
        ),
        // Only update extensions when VSCode is already running. When it's
        // running, `code --update-extensions` talks to the existing instance via
        // IPC and exits in ~2 s. When it's not running, the command opens a
        // full GUI window and never exits on its own — skip instead.
        Task::new_with_requires(
            "vscode-extensions",
            "editor",
            &["vscode"],
            "python",
            vec![
                "-c".to_string(),
                [
                    "import shutil, subprocess, sys",
                    r#"if not shutil.which("code"):"#,
                    r#"    print("code not in PATH"); sys.exit(0)"#,
                    r#"r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq Code.exe", "/NH"],"#,
                    r#"                   capture_output=True, text=True)"#,
                    r#"if "Code.exe" not in r.stdout:"#,
                    r#"    print("VSCode not running; skipping extension update (run with VSCode open to update)")"#,
                    "    sys.exit(0)",
                    r#"r2 = subprocess.run(["code", "--update-extensions"])"#,
                    "sys.exit(r2.returncode)",
                ]
                .join("\n"),
            ],
            "code",
        ),
        // git lfs install refreshes global hooks; binary itself is managed by winget/scoop
        Task::new_with_requires(
            "git-lfs",
            "git",
            &["git"],
            "git",
            vec!["lfs".to_string(), "install".to_string()],
            "git-lfs",
        ),
        Task::new(
            "gh-extensions",
            "github",
            &["github"],
            "gh",
            &["extension", "upgrade", "--all"],
        ),
        Task::new("yt-dlp", "media", &["media"], "yt-dlp", &["-U"]),
        Task::new_vec(
            "mise",
            "version-manager",
            &["mise"],
            "python",
            // Skip self-update when mise is managed by winget/scoop/brew — those
            // package managers own the binary path and self-update will hang or
            // fail trying to overwrite a managed executable.
            vec![
                "-c".to_string(),
                [
                    "import shutil, subprocess, sys",
                    r#"p = (shutil.which("mise") or "").replace("\\", "/").lower()"#,
                    r#"if "winget" in p or "/microsoft/" in p or "/scoop/" in p or "/homebrew/" in p:"#,
                    r#"    print("mise is package-manager-managed; update handled by winget/scoop task")"#,
                    "    sys.exit(0)",
                    r#"r = subprocess.run(["mise", "self-update"])"#,
                    "sys.exit(r.returncode)",
                ]
                .join("\n"),
            ],
        ),
        Task::new(
            "mise-upgrade",
            "version-manager",
            &["mise"],
            "mise",
            &["upgrade"],
        ),
        Task::new("tldr", "dev-tools", &["dev-tools"], "tldr", &["--update"]),
        // ollama list shows installed models; use --pull-models flag (future) to upgrade them
        Task::new("ollama", "ai", &["ai"], "ollama", &["list"]),
    ]
    .into_iter()
    .map(mark_missing)
    .collect()
}

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
            skip_reason: None,
            timeout_override: None,
            acceptable_exit_codes: vec![],
            ok_on_timeout: false,
        }
    }

    fn with_timeout(mut self, secs: u64) -> Self {
        self.timeout_override = Some(Duration::from_secs(secs));
        self
    }

    fn with_acceptable_exit_codes(mut self, codes: &[i32]) -> Self {
        self.acceptable_exit_codes = codes.to_vec();
        self
    }

    fn with_ok_on_timeout(mut self) -> Self {
        self.ok_on_timeout = true;
        self
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
}

fn winget_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let mut args = vec![
        "upgrade",
        "--all",
        "--source",
        "winget",
        "--include-unknown",
        "--include-pinned",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--disable-interactivity",
        "--silent",
        "--force",
    ]
    .into_iter()
    .map(str::to_string)
    .collect::<Vec<_>>();

    for package in skip_packages {
        args.push("--exclude".to_string());
        args.push(package.clone());
    }

    args
}

/// Generates args for `python -c <script>` that upgrades pip itself then all outdated packages.
fn pip_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let skip_set = skip_packages
        .iter()
        .map(|s| format!("\"{}\"", s.to_lowercase().replace('"', "\\\"")))
        .collect::<Vec<_>>()
        .join(", ");

    // Python script (python -c accepts newlines): upgrade pip, then upgrade
    // all outdated packages. On PEP 668 externally-managed installs (Arch,
    // Debian, ...) pip must not touch system site-packages — report and bow
    // out so the OS package manager keeps ownership. Exits non-zero if any
    // package upgrade fails, so the task is not reported as Succeeded.
    let script = format!(
        r#"
import json, os, subprocess, sys, sysconfig
if os.path.exists(os.path.join(sysconfig.get_path("stdlib"), "EXTERNALLY-MANAGED")):
    print("pip: externally managed environment (PEP 668); system packages are owned by the OS package manager")
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
    rc = subprocess.run([sys.executable, "-m", "pip", "install", "-U", p], check=False).returncode
    print(("upgraded " if rc == 0 else "FAILED ") + p)
    if rc != 0:
        failed.append(p)
sys.exit(1 if failed else 0)
"#
    );

    vec!["-c".to_string(), script]
}

/// Generates args for `python -c <script>` that lists dotnet global tools and updates each.
fn dotnet_tools_upgrade_args() -> Vec<String> {
    let script = "import subprocess,json,sys;\
r=subprocess.run([\"dotnet\",\"tool\",\"list\",\"--global\"],capture_output=True,text=True);\
lines=r.stdout.strip().splitlines()[2:];\
tools=[l.split()[0] for l in lines if l.strip()];\
print(\"No global dotnet tools installed\") if not tools else [subprocess.run([\"dotnet\",\"tool\",\"update\",\"--global\",t],check=False) for t in tools]";

    vec!["-c".to_string(), script.to_string()]
}

fn mark_missing(mut task: Task) -> Task {
    if !command_exists(task.requires) {
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
    let fast_skip = normalize_slice(&[
        "chocolatey",
        "npm",
        "pnpm",
        "bun",
        "deno",
        "rustup",
        "cargo",
        "pip",
        "uv",
        "uv-tools",
        "juliaup",
        "vscode-extensions",
        "mise-upgrade",
        "tldr",
    ]);
    let ultra_skip = normalize_slice(&["winget", "winget-source", "scoop"]);

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
        return Ok(Config::default());
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

fn print_task_list(tasks: &[Task]) {
    println!("{:<24} {:<18} State", "Task", "Category");
    println!("{:<24} {:<18} -----", "----", "--------");
    for task in tasks {
        let state = task.skip_reason.as_deref().unwrap_or("planned");
        println!("{:<24} {:<18} {}", task.id, task.category, state);
    }
}

fn print_summary(summary: &RunSummary) {
    let mut counts = BTreeMap::<&str, usize>::new();
    for result in &summary.results {
        *counts.entry(&result.status).or_default() += 1;
    }

    println!();
    println!("{:<24} {:<12} {:<8} Exit", "Task", "Status", "Time(s)");
    println!("{:<24} {:<12} {:<8} ----", "----", "------", "-------");
    for r in &summary.results {
        let exit = r
            .exit_code
            .map(|c| c.to_string())
            .unwrap_or_else(|| "-".to_string());
        let secs = r.duration_ms as f64 / 1000.0;
        println!(
            "{:<24} {:<12} {:<8} {}",
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

fn command_exists(name: &str) -> bool {
    if Path::new(name).components().count() > 1 {
        return Path::new(name).exists();
    }

    let Some(path) = env::var_os("PATH") else {
        return false;
    };

    let extensions = if cfg!(windows) {
        env::var_os("PATHEXT")
            .map(|value| {
                env::split_paths(&value)
                    .filter_map(|path| path.into_os_string().into_string().ok())
                    .collect::<Vec<_>>()
            })
            .filter(|items| !items.is_empty())
            .unwrap_or_else(|| vec![".exe".to_string(), ".cmd".to_string(), ".bat".to_string()])
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
                    return true;
                }
            }
        } else if dir.join(name).is_file() {
            return true;
        }
    }

    false
}

/// WSL appends the Windows PATH to the Linux PATH, so Windows-side shims
/// (scoop, gem, code, ...) probe as "present" but cannot run in the Linux
/// environment (exec format error / CRLF scripts). When running under WSL,
/// ignore automounted Windows drives (/mnt/<drive-letter>/...) so those
/// tasks are reported as skipped instead of failing — the Windows build of
/// this tool is responsible for them.
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

/// True for paths under a WSL drvfs automount: /mnt/<single-drive-letter>[/...]
fn is_windows_drive_mount_path(dir: &Path) -> bool {
    let mut comps = dir.components();
    matches!(comps.next(), Some(std::path::Component::RootDir))
        && comps.next().is_some_and(|c| c.as_os_str() == "mnt")
        && comps.next().is_some_and(|c| {
            let s = c.as_os_str().to_string_lossy();
            s.len() == 1 && s.chars().all(|ch| ch.is_ascii_alphabetic())
        })
}

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

fn tail<T>(items: Vec<T>, count: usize) -> Vec<T> {
    let len = items.len();
    items.into_iter().skip(len.saturating_sub(count)).collect()
}

fn now_string() -> String {
    let format = format_description!("[year]-[month]-[day]T[hour]:[minute]:[second]");
    OffsetDateTime::now_local()
        .unwrap_or_else(|_| OffsetDateTime::now_utc())
        .format(&format)
        .unwrap_or_else(|_| "unknown".to_string())
}

fn find_repo_root() -> Result<PathBuf> {
    let mut dir = env::current_dir()?;
    loop {
        if dir.join("update-config.json").exists() || dir.join("updatescript.ps1").exists() {
            return Ok(dir);
        }
        if !dir.pop() {
            break;
        }
    }
    env::current_dir().context("failed to resolve current directory")
}

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
    fn winget_args_include_excludes() {
        let args = winget_upgrade_args(&["Foo.Bar".to_string(), "Baz.Qux".to_string()]);
        assert!(args.contains(&"--all".to_string()));
        let pairs: Vec<_> = args.windows(2).filter(|w| w[0] == "--exclude").collect();
        assert_eq!(pairs.len(), 2);
        assert_eq!(pairs[0][1], "Foo.Bar");
        assert_eq!(pairs[1][1], "Baz.Qux");
    }

    #[test]
    fn pip_args_embed_lowercased_skips() {
        let args = pip_upgrade_args(&["PyLint".to_string()]);
        assert_eq!(args[0], "-c");
        assert!(args[1].contains("\"pylint\""));
    }

    #[test]
    fn matches_any_by_id_category_and_tag() {
        let task = Task::new("rustup", "rust", &["toolchain"], "rustup", &["update"]);
        let by_id = normalize_slice(&["RUSTUP"]);
        let by_cat = normalize_slice(&["Rust"]);
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
}
