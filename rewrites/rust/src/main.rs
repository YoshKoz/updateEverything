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
    #[serde(default = "default_cleanup_days")]
    temp_cleanup_days: u32,
    #[serde(default = "default_log_retention")]
    log_retention_days: u32,
}

fn default_cleanup_days() -> u32 { 7 }
fn default_log_retention() -> u32 { 14 }

#[derive(Clone, Debug)]
struct Task {
    id: &'static str,
    category: &'static str,
    tags: &'static [&'static str],
    command: &'static str,
    args: Vec<String>,
    requires: &'static str,
    resource: &'static str,
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
    let start = Instant::now();
    let started_at = now_string();
    let repo_root = find_repo_root()?;
    let config_path = cli
        .config
        .clone()
        .unwrap_or_else(|| repo_root.join("update-config.json"));
    let config = load_config(&config_path)?;
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
        use std::os::unix::process::CommandExt;
        command.process_group(0);
    }
    let mut child = match command.spawn() {
        Ok(child) => child,
        Err(err) => {
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

    let tasks = vec![
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
        Task::new_vec(
            "winget",
            "package-manager",
            &["windows", "winget"],
            "winget",
            winget_upgrade_args,
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_acceptable_exit_codes(&[-1978335189, -1978335212]),
        Task::new_vec(
            "winget-batch",
            "package-manager",
            &["windows", "winget"],
            "winget",
            vec![
                "upgrade".into(), "--all".into(), "--source".into(), "winget".into(),
                "--include-unknown".into(), "--accept-source-agreements".into(),
                "--disable-interactivity".into(),
            ],
        )
        .with_resource("winget")
        .with_timeout(cli.winget_timeout_sec)
        .with_acceptable_exit_codes(&[-1978335189, -1978335212]),
        Task::new_vec(
            "winget-pin-audit",
            "package-manager",
            &["windows", "winget"],
            "winget",
            vec!["pin".into(), "list".into()],
        )
        .with_acceptable_exit_codes(&[-1978335212]),
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
        .with_acceptable_exit_codes(&[-1978335189, -1978335212])
        .with_requires("winget")
        .with_skip_if(cli.skip_store_apps, "disabled by --skip-store-apps"),
        // ── scoop ────────────────────────────────────────────────────────────
        Task::new(
            "scoop",
            "package-manager",
            &["windows", "scoop"],
            "scoop",
            &["update", "*"],
        ),
        // ── chocolatey ───────────────────────────────────────────────────────
        Task::new_vec(
            "chocolatey",
            "package-manager",
            &["windows", "choco"],
            "choco",
            {
                let mut a = vec!["upgrade".into(), "all".into(), "-y".into(), "--no-progress".into()];
                for pkg in &config.chocolatey_skip_packages {
                    a.push("--except".into());
                    a.push(pkg.clone());
                }
                a
            },
        ),
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
                "try { Update-MpSignature -ErrorAction Stop; Write-Output 'Defender signatures updated.' } \
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
        Task::new("pipx", "python", &["python"], "pipx", &["upgrade-all"]),
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
        Task::new(
            "poetry",
            "python",
            &["python"],
            "poetry",
            &["self", "update"],
        )
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
        .with_acceptable_exit_codes(&[-1978335189, -1978335212])
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

fn winget_upgrade_args(skip_packages: &[String]) -> Vec<String> {
    let mut args: Vec<String> = vec![
        "upgrade", "--all", "--source", "winget", "--include-unknown", "--include-pinned",
        "--accept-package-agreements", "--accept-source-agreements",
        "--disable-interactivity", "--silent", "--force",
    ]
    .into_iter()
    .map(str::to_string)
    .collect();
    for package in skip_packages {
        args.push("--exclude".into());
        args.push(package.clone());
    }
    args
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
import shutil, subprocess, sys
p = (shutil.which("uv") or "").replace("\\", "/").lower()
if "/python" in p or "/scripts/" in p:
    print("uv is pip-managed; update handled by pip task")
    sys.exit(0)
r = subprocess.run(["uv", "self", "update"])
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
import json, subprocess, sys
skip = [{skip_json}]
r = subprocess.run(["npm", "ls", "-g", "--depth=0", "--json"], capture_output=True, text=True)
try:
    data = json.loads(r.stdout or "{{}}")
    pkgs = [k for k in data.get("dependencies", {{}}).keys() if k not in skip and not k.startswith("npm")]
except Exception:
    pkgs = []
if not pkgs:
    print("npm: no global packages to upgrade (or npm ls failed)")
    subprocess.run(["npm", "update", "-g"])
    sys.exit(0)
failed = []
for p in pkgs:
    rc = subprocess.run(["npm", "install", "-g", p]).returncode
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
benign = rc in (0, -1978335189, -1978335212) or \
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

fn wsl_distros_script() -> &'static str {
    "$prev=[Console]::OutputEncoding;\
     [Console]::OutputEncoding=[System.Text.Encoding]::Unicode;\
     $raw=wsl -l -q 2>$null;\
     [Console]::OutputEncoding=$prev;\
     $distros=@($raw|ForEach-Object{($_ -replace '\\0','').Trim()}|Where-Object{$_-and$_ -notmatch 'docker-desktop'});\
     if($distros.Count-eq 0){Write-Output 'No WSL distros found.';return};\
     $s='if command -v apt-get >/dev/null 2>&1; then sudo apt-get update -qq && sudo apt-get upgrade -y; \
     elif command -v pacman >/dev/null 2>&1; then sudo pacman -Syu --noconfirm; \
     elif command -v dnf >/dev/null 2>&1; then sudo dnf upgrade -y; \
     elif command -v yum >/dev/null 2>&1; then sudo yum update -y; \
     elif command -v zypper >/dev/null 2>&1; then sudo zypper update -y; \
     else echo no_known_package_manager; fi';\
     foreach($d in $distros){\
     Write-Output \"Updating WSL distro: $d\";\
     wsl -d $d -- sh -lc $s}"
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
        return "Write-Output 'No WindowsOptionalFeatures configured in update-config.json.'".into();
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
    "$pkgs=@(Get-AppxPackage -AllUsers -EA SilentlyContinue|\
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
        "chocolatey", "wsl-distros", "npm", "pnpm", "yarn", "bun", "deno",
        "rustup", "cargo", "go", "pip", "pip-health", "pipx", "uv", "uv-tools",
        "poetry", "composer", "ruby-gems", "flutter", "juliaup",
        "oh-my-posh", "yt-dlp", "volta", "fnm", "dotnet-tools",
        "dotnet-workloads", "vscode-extensions", "powershell-modules",
        "powershell-help", "uv-python", "ollama-models",
        "vcpkg", "conda", "gcloud", "az", "aws", "terraform", "pulumi",
        "kubectl", "helm", "hugo", "opentofu", "starship", "zoxide",
        "gitleaks", "trivy", "packer", "nvm", "devcontainer", "cross-manager",
        "mise-upgrade", "tldr",
    ]);

    // Matches PS1 UltraFastSkip
    let ultra_skip = normalize_slice(&[
        "windows-update", "store-apps", "wsl", "wsl-distros", "defender", "cleanup",
        "winget", "winget-source", "scoop",
    ]);

    // Profile-based skip presets
    let profile_skip: BTreeSet<String> = match cli.profile.as_deref() {
        Some("minimal") => normalize_slice(&[
            "vcpkg", "conda", "gcloud", "az", "aws", "terraform", "pulumi",
            "kubectl", "helm", "hugo", "opentofu", "starship", "gitleaks",
            "trivy", "packer", "nvm", "devcontainer",
        ]),
        Some("work") => BTreeSet::new(),
        Some("personal") => BTreeSet::new(),
        Some("gaming") => normalize_slice(&[
            "vcpkg", "conda", "gcloud", "az", "aws", "terraform", "pulumi",
            "kubectl", "helm", "hugo", "opentofu", "gitleaks", "trivy",
            "packer", "nvm", "devcontainer", "powershell-modules",
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
    for r in &summary.results {
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
        let task = Task::new("rustup", "systems-language", &["toolchain"], "rustup", &["update"]);
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
