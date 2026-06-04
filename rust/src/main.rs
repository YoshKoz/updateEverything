use std::{
    collections::{BTreeMap, BTreeSet},
    env, fs,
    io::{BufRead, BufReader},
    path::{Path, PathBuf},
    process::{Command, ExitStatus, Stdio},
    sync::{Arc, Mutex},
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
}

#[derive(Clone, Debug)]
struct Task {
    id: &'static str,
    category: &'static str,
    tags: &'static [&'static str],
    command: &'static str,
    args: Vec<String>,
    requires: &'static str,
    skip_reason: Option<String>,
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

fn run_tasks(
    tasks: Vec<Task>,
    cli: &Cli,
    jobs: usize,
    timeout: Duration,
) -> Vec<TaskSummary> {
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

    let multi = jobs > 1;
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
            let r = run_task_streaming(&task, cli.quiet, timeout, false);
            results.push(r);
        }
    } else {
        let queue: Arc<Mutex<Vec<Task>>> = Arc::new(Mutex::new(to_run));
        let out: Arc<Mutex<Vec<TaskSummary>>> = Arc::new(Mutex::new(Vec::new()));

        let mut handles = vec![];
        let worker_count = jobs;
        for _ in 0..worker_count {
            let queue = Arc::clone(&queue);
            let out = Arc::clone(&out);
            let quiet = cli.quiet;
            let h = thread::spawn(move || loop {
                let task = {
                    let mut q = queue.lock().unwrap();
                    if q.is_empty() {
                        break;
                    }
                    q.remove(0)
                };
                let r = run_task_streaming(&task, quiet, timeout, multi);
                out.lock().unwrap().push(r);
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

fn run_task_streaming(task: &Task, quiet: bool, timeout: Duration, prefix_output: bool) -> TaskSummary {
    if !quiet {
        println!("run  {:<22} {} {}", task.id, task.command, shell_join(&task.args));
    }

    let start = Instant::now();
    let mut child = match Command::new(task.command)
        .args(&task.args)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
    {
        Ok(child) => child,
        Err(err) => {
            return make_summary(task, "Failed", start.elapsed().as_millis(), Some(127), vec![err.to_string()]);
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
                let _ = child.kill();
                break;
            }
            Ok(None) => thread::sleep(Duration::from_millis(200)),
            Err(err) => {
                return make_summary(task, "Failed", start.elapsed().as_millis(), None, vec![err.to_string()]);
            }
        }
    }

    let exit_status = child.wait().ok();
    let _ = stdout_handle.join();
    let _ = stderr_handle.join();

    let mut all_lines = Arc::try_unwrap(lines_out).unwrap().into_inner().unwrap();
    all_lines.extend(Arc::try_unwrap(lines_err).unwrap().into_inner().unwrap());

    let status = if timed_out {
        "TimedOut".to_string()
    } else {
        exit_status.map(status_name).unwrap_or_else(|| "Failed".to_string())
    };
    let code = exit_status.and_then(|s| s.code());
    let duration_ms = start.elapsed().as_millis();

    if !quiet {
        let secs = duration_ms as f64 / 1000.0;
        println!("done {:<22} {} ({:.1}s)", task.id, status, secs);
    }

    make_summary(task, &status, duration_ms, code, tail(all_lines, 30))
}

fn make_summary(task: &Task, status: &str, duration_ms: u128, exit_code: Option<i32>, output_tail: Vec<String>) -> TaskSummary {
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
    let winget_upgrade_args = winget_upgrade_args(&config.winget_skip_packages);
    vec![
        Task::new(
            "winget-source",
            "package-manager",
            &["windows", "winget"],
            "winget",
            &["source", "update"],
        ),
        Task::new_vec(
            "winget",
            "package-manager",
            &["windows", "winget"],
            "winget",
            winget_upgrade_args,
        ),
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
        Task::new(
            "cargo",
            "rust",
            &["rust"],
            "cargo",
            &["install-update", "-a"],
        ),
        Task::new("pipx", "python", &["python"], "pipx", &["upgrade-all"]),
        Task::new_vec(
            "pip",
            "python",
            &["python"],
            "python",
            pip_args(&config.pip_skip_packages),
        ),
        Task::new("uv", "python", &["python"], "uv", &["self", "update"]),
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
        Task::new(
            "vscode-extensions",
            "editor",
            &["vscode"],
            "code",
            &["--update-extensions"],
        ),
        Task::new("git-lfs", "git", &["git"], "git", &["lfs", "update"]),
        Task::new(
            "gh-extensions",
            "github",
            &["github"],
            "gh",
            &["extension", "upgrade", "--all"],
        ),
        Task::new("yt-dlp", "media", &["media"], "yt-dlp", &["-U"]),
        Task::new("mise", "version-manager", &["mise"], "mise", &["self-upgrade"]),
        Task::new("mise-upgrade", "version-manager", &["mise"], "mise", &["upgrade"]),
        Task::new("tldr", "dev-tools", &["dev-tools"], "tldr", &["--update"]),
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
            skip_reason: None,
        }
    }
}

fn winget_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let mut args = vec![
        "upgrade",
        "--all",
        "--include-unknown",
        "--include-pinned",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--disable-interactivity",
        "--silent",
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

fn pip_args(skip_packages: &[String]) -> Vec<String> {
    let mut args = vec!["-m", "pip", "list", "--outdated", "--format=json"]
        .into_iter()
        .map(str::to_string)
        .collect::<Vec<_>>();

    for package in skip_packages {
        args.push("--exclude".to_string());
        args.push(package.clone());
    }

    args
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
            v.push(PathBuf::from(local).join("Update-Everything").join("last-run.json"));
        }
        v.push(repo_root.join("staging").join("rust-run-summary.json"));
        v
    };

    for path in candidates {
        if !path.exists() {
            continue;
        }
        // Reject file if it's older than since_hours
        if cli.since_hours > 0.0 {
            if let Ok(meta) = fs::metadata(&path) {
                if let Ok(modified) = meta.modified() {
                    if let Ok(age) = modified.elapsed() {
                        if age.as_secs_f64() / 3600.0 > cli.since_hours {
                            continue;
                        }
                    }
                }
            }
        }
        if let Ok(text) = fs::read_to_string(&path) {
            if let Ok(summary) = serde_json::from_str::<PrevSummary>(&text) {
                return Some(summary);
            }
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
    println!("{:<24} {:<18} {}", "Task", "Category", "State");
    println!("{:<24} {:<18} {}", "----", "--------", "-----");
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
    println!(
        "{:<24} {:<12} {:<10} {}",
        "Task", "Status", "Time(ms)", "Exit"
    );
    println!(
        "{:<24} {:<12} {:<10} {}",
        "----", "------", "--------", "----"
    );
    for r in &summary.results {
        let exit = r
            .exit_code
            .map(|c| c.to_string())
            .unwrap_or_else(|| "-".to_string());
        println!(
            "{:<24} {:<12} {:<10} {}",
            r.id, r.status, r.duration_ms, exit
        );
    }
    println!();

    let total = summary.results.len();
    let succeeded = counts.get("Succeeded").copied().unwrap_or_default();
    let failed = counts.get("Failed").copied().unwrap_or_default();
    let timed_out = counts.get("TimedOut").copied().unwrap_or_default();
    let skipped = counts.get("Skipped").copied().unwrap_or_default();
    let dry = counts.get("DryRun").copied().unwrap_or_default();
    println!(
        "done  total={total}  succeeded={succeeded}  failed={failed}  timed-out={timed_out}  skipped={skipped}  dry={dry}  duration={}ms",
        summary.duration_ms
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
            .unwrap_or_else(|| {
                vec![
                    ".exe".to_string(),
                    ".cmd".to_string(),
                    ".bat".to_string(),
                ]
            })
    } else {
        vec!["".to_string()]
    };

    for dir in env::split_paths(&path) {
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
