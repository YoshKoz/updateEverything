const std = @import("std");
const Io = std.Io;
const Allocator = std.mem.Allocator;
const windows = std.os.windows;
const is_windows = @import("builtin").os.tag == .windows;

const version = "0.2.0";

// Zig 0.16 truncates a Windows exit status to u8 in Child.Term, which destroys
// the negative winget codes the task table matches on, so read the raw DWORD.
extern "kernel32" fn GetExitCodeProcess(hProcess: windows.HANDLE, lpExitCode: *u32) callconv(.winapi) windows.BOOL;
extern "kernel32" fn TerminateProcess(hProcess: windows.HANDLE, uExitCode: c_uint) callconv(.winapi) windows.BOOL;
extern "kernel32" fn GetCurrentProcessId() callconv(.winapi) u32;

const SystemTime = extern struct {
    year: u16,
    month: u16,
    day_of_week: u16,
    day: u16,
    hour: u16,
    minute: u16,
    second: u16,
    milliseconds: u16,
};
extern "kernel32" fn GetLocalTime(lpSystemTime: *SystemTime) callconv(.winapi) void;

const still_active: u32 = 259;
const reader_drain_ms: i64 = 3_000;

/// Task scripts print a line starting with this when they deliberately took no
/// action. Without it an intentional no-op reports as Succeeded, which reads as
/// "updated" in the summary.
const skip_prefix = "SKIPPED:";

// ─── Output ──────────────────────────────────────────────────────────────────

var out_mu: Io.Mutex = .init;

var color_on: bool = false;
var log_file: ?Io.File = null;
var log_pos: u64 = 0;

const A = struct {
    const reset = "\x1b[0m";
    const bold = "\x1b[1m";
    const dim = "\x1b[2m";
    const red = "\x1b[38;2;255;95;95m";
    const green = "\x1b[38;2;95;215;135m";
    const yellow = "\x1b[38;2;255;200;80m";
    const blue = "\x1b[38;2;110;170;255m";
    const magenta = "\x1b[38;2;215;135;255m";
    const cyan = "\x1b[38;2;95;215;255m";
    const grey = "\x1b[38;2;130;130;130m";
};

/// Nerd Font glyphs (escaped so the source file stays ASCII-safe).
const G = struct {
    const run = "\u{f04b}";
    const ok = "\u{f00c}";
    const fail = "\u{f00d}";
    const timeout = "\u{f0f2}";
    const skip = "\u{f056}";
    const dry = "\u{f06e}";
    const retry = "\u{f021}";
    const changed = "\u{f135}";
    const summary = "\u{f0ce}";
    const gear = "\u{f013}";
};

fn col(code: []const u8) []const u8 {
    return if (color_on) code else "";
}

fn statusColor(status: []const u8) []const u8 {
    if (std.mem.eql(u8, status, "Succeeded")) return col(A.green);
    if (std.mem.eql(u8, status, "Failed")) return col(A.red);
    if (std.mem.eql(u8, status, "TimedOut")) return col(A.yellow);
    if (std.mem.eql(u8, status, "Skipped")) return col(A.grey);
    if (std.mem.eql(u8, status, "DryRun")) return col(A.magenta);
    return "";
}

fn statusGlyph(status: []const u8) []const u8 {
    if (std.mem.eql(u8, status, "Succeeded")) return G.ok;
    if (std.mem.eql(u8, status, "Failed")) return G.fail;
    if (std.mem.eql(u8, status, "TimedOut")) return G.timeout;
    if (std.mem.eql(u8, status, "Skipped")) return G.skip;
    if (std.mem.eql(u8, status, "DryRun")) return G.dry;
    return " ";
}

/// Drop CSI sequences (ESC [ ... final-byte) so the log file stays plain text.
fn stripAnsi(src: []const u8, dst: []u8) []const u8 {
    var n: usize = 0;
    var i: usize = 0;
    while (i < src.len) : (i += 1) {
        if (src[i] == 0x1b and i + 1 < src.len and src[i + 1] == '[') {
            i += 2;
            while (i < src.len and (src[i] < 0x40 or src[i] > 0x7e)) : (i += 1) {}
            continue;
        }
        dst[n] = src[i];
        n += 1;
    }
    return dst[0..n];
}

fn emit(io: Io, file: Io.File, comptime fmt: []const u8, args: anytype) void {
    var buf: [16384]u8 = undefined;
    var wbuf: [4096]u8 = undefined;
    out_mu.lockUncancelable(io);
    defer out_mu.unlock(io);
    // Streaming mode: a positional writer restarts at offset 0 on every call,
    // which silently overwrote the log when stdout was redirected to a file.
    var w = file.writerStreaming(io, &wbuf);
    const text = std.fmt.bufPrint(&buf, fmt, args) catch {
        w.interface.print(fmt, args) catch {};
        w.interface.flush() catch {};
        return;
    };
    w.interface.writeAll(text) catch {};
    w.interface.flush() catch {};
    if (log_file) |lf| {
        var plain_buf: [16384]u8 = undefined;
        const plain = stripAnsi(text, &plain_buf);
        var lbuf: [4096]u8 = undefined;
        var lw = lf.writer(io, &lbuf);
        lw.pos = log_pos;
        lw.interface.writeAll(plain) catch {};
        lw.interface.flush() catch {};
        log_pos = lw.pos;
    }
}

fn p(io: Io, comptime fmt: []const u8, args: anytype) void {
    emit(io, Io.File.stdout(), fmt, args);
}

fn pe(io: Io, comptime fmt: []const u8, args: anytype) void {
    emit(io, Io.File.stderr(), fmt, args);
}

fn fatal(io: Io, comptime fmt: []const u8, args: anytype) noreturn {
    pe(io, "error: " ++ fmt ++ "\n", args);
    std.process.exit(2);
}

// ─── String utilities ────────────────────────────────────────────────────────

fn normalize(gpa: Allocator, value: []const u8) []const u8 {
    const trimmed = std.mem.trim(u8, value, " \t\r\n");
    const buf = gpa.alloc(u8, trimmed.len) catch @panic("oom");
    for (trimmed, 0..) |c, i| buf[i] = std.ascii.toLower(c);
    return buf;
}

const StringSet = struct {
    map: std.StringHashMap(void),

    fn init(gpa: Allocator) StringSet {
        return .{ .map = .init(gpa) };
    }

    fn add(self: *StringSet, value: []const u8) void {
        self.map.put(value, {}) catch @panic("oom");
    }

    fn addNormalized(self: *StringSet, gpa: Allocator, value: []const u8) void {
        self.add(normalize(gpa, value));
    }

    fn contains(self: *const StringSet, value: []const u8) bool {
        return self.map.contains(value);
    }

    fn isEmpty(self: *const StringSet) bool {
        return self.map.count() == 0;
    }

    fn fromSlices(gpa: Allocator, values: []const []const u8) StringSet {
        var set: StringSet = .init(gpa);
        for (values) |v| set.addNormalized(gpa, v);
        return set;
    }
};

fn concat(gpa: Allocator, parts: []const []const u8) []const u8 {
    return std.mem.concat(gpa, u8, parts) catch @panic("oom");
}

fn fmtAlloc(gpa: Allocator, comptime f: []const u8, args: anytype) []const u8 {
    return std.fmt.allocPrint(gpa, f, args) catch @panic("oom");
}

/// Replace every occurrence of `needle` with `replacement`.
fn replaceAll(gpa: Allocator, haystack: []const u8, needle: []const u8, replacement: []const u8) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    var rest = haystack;
    while (std.mem.indexOf(u8, rest, needle)) |idx| {
        list.appendSlice(gpa, rest[0..idx]) catch @panic("oom");
        list.appendSlice(gpa, replacement) catch @panic("oom");
        rest = rest[idx + needle.len ..];
    }
    list.appendSlice(gpa, rest) catch @panic("oom");
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

/// `'a'`-quoted, single-quote-doubled PowerShell literals joined by `sep`.
fn psList(gpa: Allocator, values: []const []const u8, sep: []const u8) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    for (values, 0..) |v, i| {
        if (i != 0) list.appendSlice(gpa, sep) catch @panic("oom");
        list.append(gpa, '\'') catch @panic("oom");
        list.appendSlice(gpa, replaceAll(gpa, v, "'", "''")) catch @panic("oom");
        list.append(gpa, '\'') catch @panic("oom");
    }
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

/// `"a", "b"` — double-quoted literals for embedding in a python list/set.
fn pyList(gpa: Allocator, values: []const []const u8, lower: bool) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    for (values, 0..) |v, i| {
        if (i != 0) list.appendSlice(gpa, ", ") catch @panic("oom");
        list.append(gpa, '"') catch @panic("oom");
        const text = if (lower) normalize(gpa, v) else v;
        list.appendSlice(gpa, replaceAll(gpa, text, "\"", "\\\"")) catch @panic("oom");
        list.append(gpa, '"') catch @panic("oom");
    }
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

fn jsonEscape(gpa: Allocator, value: []const u8) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    for (value) |c| switch (c) {
        '"' => list.appendSlice(gpa, "\\\"") catch @panic("oom"),
        '\\' => list.appendSlice(gpa, "\\\\") catch @panic("oom"),
        '\n' => list.appendSlice(gpa, "\\n") catch @panic("oom"),
        '\r' => list.appendSlice(gpa, "\\r") catch @panic("oom"),
        '\t' => list.appendSlice(gpa, "\\t") catch @panic("oom"),
        else => {
            if (c < 0x20) {
                list.appendSlice(gpa, fmtAlloc(gpa, "\\u{x:0>4}", .{c})) catch @panic("oom");
            } else {
                list.append(gpa, c) catch @panic("oom");
            }
        },
    };
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

fn shellJoin(gpa: Allocator, args: []const []const u8) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    for (args, 0..) |arg, i| {
        if (i != 0) list.append(gpa, ' ') catch @panic("oom");
        if (std.mem.indexOfScalar(u8, arg, ' ') != null) {
            list.append(gpa, '"') catch @panic("oom");
            list.appendSlice(gpa, replaceAll(gpa, arg, "\"", "\\\"")) catch @panic("oom");
            list.append(gpa, '"') catch @panic("oom");
        } else {
            list.appendSlice(gpa, arg) catch @panic("oom");
        }
    }
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

/// Same as `shellJoin` but collapses multi-line arguments to a placeholder.
/// Inline `-c` script bodies are dozens of lines and bury the run/dry echo.
fn shellJoinBrief(gpa: Allocator, args: []const []const u8) []const u8 {
    var brief: std.ArrayList([]const u8) = .empty;
    for (args) |arg| {
        if (std.mem.indexOfScalar(u8, arg, '\n') != null) {
            var count: usize = 0;
            var it = std.mem.splitScalar(u8, arg, '\n');
            while (it.next()) |line| {
                if (std.mem.trim(u8, line, " \t\r").len != 0) count += 1;
            }
            brief.append(gpa, fmtAlloc(gpa, "<{d}-line-script>", .{count})) catch @panic("oom");
        } else {
            brief.append(gpa, arg) catch @panic("oom");
        }
    }
    return shellJoin(gpa, brief.items);
}

/// Keep up to `max` lines: first half + last half when output is long.
/// Ensures both the initial version-table (head) and completion messages (tail) survive.
fn capOutput(gpa: Allocator, lines: []const []const u8, max: usize) [][]const u8 {
    if (lines.len <= max) return gpa.dupe([]const u8, lines) catch @panic("oom");
    const head = max / 2;
    const tail_count = max - head;
    var result = gpa.alloc([]const u8, max) catch @panic("oom");
    @memcpy(result[0..head], lines[0..head]);
    @memcpy(result[head..], lines[lines.len - tail_count ..]);
    return result;
}

fn nowString(gpa: Allocator) []const u8 {
    var st: SystemTime = undefined;
    GetLocalTime(&st);
    return fmtAlloc(gpa, "{d:0>4}-{d:0>2}-{d:0>2}T{d:0>2}:{d:0>2}:{d:0>2}", .{
        st.year, st.month, st.day, st.hour, st.minute, st.second,
    });
}

fn nowMs(io: Io) i64 {
    const ts = Io.Timestamp.now(io, .awake);
    return @intCast(@divTrunc(ts.nanoseconds, std.time.ns_per_ms));
}

fn sleepMs(io: Io, ms: u64) void {
    io.sleep(.{ .nanoseconds = @intCast(ms * std.time.ns_per_ms) }, .awake) catch {};
}

// ─── CLI ─────────────────────────────────────────────────────────────────────

const Cli = struct {
    dry_run: bool = false,
    list_tasks: bool = false,
    fast: bool = false,
    ultra_fast: bool = false,
    only: []const []const u8 = &.{},
    skip: []const []const u8 = &.{},
    config: ?[]const u8 = null,
    json_summary: ?[]const u8 = null,
    log_file: ?[]const u8 = null,
    color: ?bool = null,
    quiet: bool = false,
    task_timeout_sec: u64 = 1800,
    jobs: usize = 1,
    ci: bool = false,
    since_hours: f64 = 0,
    state_dir: ?[]const u8 = null,

    skip_windows_update: bool = false,
    skip_wsl: bool = false,
    skip_wsl_distros: bool = false,
    skip_defender: bool = false,
    skip_store_apps: bool = false,
    skip_powershell_modules: bool = false,
    skip_node: bool = false,
    skip_rust: bool = false,
    skip_go: bool = false,
    skip_flutter: bool = false,
    skip_ruby: bool = false,
    skip_composer: bool = false,
    skip_poetry: bool = false,
    skip_uv_tools: bool = false,
    skip_cleanup: bool = false,
    deep_clean: bool = false,
    skip_destructive: bool = false,
    skip_git_lfs: bool = false,
    skip_vcpkg: bool = false,
    skip_conda: bool = false,
    skip_cloud_tools: bool = false,
    skip_infra_tools: bool = false,
    skip_k8s_tools: bool = false,
    skip_starship: bool = false,
    skip_hugo: bool = false,
    skip_vscode_extensions: bool = false,
    skip_pip_health: bool = false,

    update_ollama_models: bool = false,
    update_powershell_help: bool = false,
    update_github_tools: bool = true,
    /// gh-notify-releases reports by default; installing from a release notification
    /// is opt-in because the repo->package mapping can only ever be a guess.
    notify_apply: bool = false,
    bypass_protection: bool = false,
    winget_timeout_sec: u64 = 600,
    ollama_timeout_sec: u64 = 600,
    retry_count: u32 = 3,
    profile: ?[]const u8 = null,
    show_skipped: bool = false,
    schedule: bool = false,
    schedule_time: ?[]const u8 = null,
};

const help_text =
    \\Zig updateEverything runner
    \\
    \\Usage: updateeverything [OPTIONS]
    \\
    \\Options:
    \\      --dry-run
    \\      --list-tasks
    \\      --fast
    \\      --ultra-fast
    \\      --only <ONLY>
    \\      --skip <SKIP>
    \\      --config <CONFIG>
    \\      --json-summary <JSON_SUMMARY>
    \\      --log-file <LOG_FILE>      plain-text copy of console output (appended)
    \\      --color / --no-color       [default: auto-detect terminal]
    \\      --quiet
    \\      --task-timeout-sec <TASK_TIMEOUT_SEC>   [default: 1800]
    \\      --jobs <JOBS>                           [default: 1]
    \\      --ci
    \\      --since-hours <SINCE_HOURS>             [default: 0]
    \\      --state-dir <STATE_DIR>
    \\      --skip-windows-update
    \\      --skip-wsl
    \\      --skip-wsl-distros
    \\      --skip-defender
    \\      --skip-store-apps
    \\      --skip-powershell-modules
    \\      --skip-node
    \\      --skip-rust
    \\      --skip-go
    \\      --skip-flutter
    \\      --skip-ruby
    \\      --skip-composer
    \\      --skip-poetry
    \\      --skip-uv-tools
    \\      --skip-cleanup
    \\      --deep-clean
    \\      --skip-destructive
    \\      --skip-git-lfs
    \\      --skip-vcpkg
    \\      --skip-conda
    \\      --skip-cloud-tools
    \\      --skip-infra-tools
    \\      --skip-k8s-tools
    \\      --skip-starship
    \\      --skip-hugo
    \\      --skip-vscode-extensions
    \\      --skip-pip-health
    \\      --update-ollama-models
    \\      --update-powershell-help
    \\      --update-github-tools   [default: on]
    \\      --notify-apply          install mapped packages found by gh-notify-releases [default: report only]
    \\      --bypass-protection
    \\      --winget-timeout-sec <WINGET_TIMEOUT_SEC>   [default: 600]
    \\      --ollama-timeout-sec <OLLAMA_TIMEOUT_SEC>   [default: 600]
    \\      --retry-count <RETRY_COUNT>                 [default: 3]
    \\      --profile <PROFILE>
    \\      --show-skipped
    \\      --schedule
    \\      --schedule-time <SCHEDULE_TIME>
    \\  -h, --help
    \\  -V, --version
    \\
;

fn splitCommas(gpa: Allocator, value: []const u8) []const []const u8 {
    var list: std.ArrayList([]const u8) = .empty;
    var it = std.mem.splitScalar(u8, value, ',');
    while (it.next()) |part| {
        if (part.len != 0) list.append(gpa, part) catch @panic("oom");
    }
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

fn parseCli(gpa: Allocator, io: Io, argv: []const [:0]const u8) Cli {
    var cli: Cli = .{};
    var i: usize = 1;
    while (i < argv.len) : (i += 1) {
        const raw: []const u8 = argv[i];
        if (std.mem.eql(u8, raw, "-h") or std.mem.eql(u8, raw, "--help")) {
            p(io, "{s}", .{help_text});
            std.process.exit(0);
        }
        if (std.mem.eql(u8, raw, "-V") or std.mem.eql(u8, raw, "--version")) {
            p(io, "updateeverything {s}\n", .{version});
            std.process.exit(0);
        }
        if (!std.mem.startsWith(u8, raw, "--")) fatal(io, "unexpected argument '{s}'", .{raw});

        var name = raw[2..];
        var inline_value: ?[]const u8 = null;
        if (std.mem.indexOfScalar(u8, name, '=')) |eq| {
            inline_value = name[eq + 1 ..];
            name = name[0..eq];
        }

        const Kind = enum { flag, string, list, u64v, usizev, u32v, f64v };
        const needs: Kind = blk: {
            inline for (.{
                .{ "only", Kind.list },               .{ "skip", Kind.list },
                .{ "config", Kind.string },           .{ "json-summary", Kind.string },
                .{ "log-file", Kind.string },         .{ "state-dir", Kind.string },
                .{ "profile", Kind.string },          .{ "schedule-time", Kind.string },
                .{ "task-timeout-sec", Kind.u64v },   .{ "winget-timeout-sec", Kind.u64v },
                .{ "ollama-timeout-sec", Kind.u64v }, .{ "jobs", Kind.usizev },
                .{ "retry-count", Kind.u32v },        .{ "since-hours", Kind.f64v },
            }) |entry| {
                if (std.mem.eql(u8, name, entry[0])) break :blk entry[1];
            }
            break :blk .flag;
        };

        var value: []const u8 = "";
        if (needs != .flag) {
            if (inline_value) |v| {
                value = v;
            } else {
                i += 1;
                if (i >= argv.len) fatal(io, "a value is required for '--{s}'", .{name});
                value = argv[i];
            }
        } else if (inline_value != null) {
            fatal(io, "'--{s}' takes no value", .{name});
        }

        if (std.mem.eql(u8, name, "dry-run")) {
            cli.dry_run = true;
        } else if (std.mem.eql(u8, name, "log-file")) {
            cli.log_file = value;
        } else if (std.mem.eql(u8, name, "color")) {
            cli.color = true;
        } else if (std.mem.eql(u8, name, "no-color")) {
            cli.color = false;
        } else if (std.mem.eql(u8, name, "list-tasks")) {
            cli.list_tasks = true;
        } else if (std.mem.eql(u8, name, "fast")) {
            cli.fast = true;
        } else if (std.mem.eql(u8, name, "ultra-fast")) {
            cli.ultra_fast = true;
        } else if (std.mem.eql(u8, name, "only")) {
            cli.only = splitCommas(gpa, value);
        } else if (std.mem.eql(u8, name, "skip")) {
            cli.skip = splitCommas(gpa, value);
        } else if (std.mem.eql(u8, name, "config")) {
            cli.config = value;
        } else if (std.mem.eql(u8, name, "json-summary")) {
            cli.json_summary = value;
        } else if (std.mem.eql(u8, name, "quiet")) {
            cli.quiet = true;
        } else if (std.mem.eql(u8, name, "task-timeout-sec")) {
            cli.task_timeout_sec = std.fmt.parseInt(u64, value, 10) catch fatal(io, "invalid value '{s}' for '--task-timeout-sec'", .{value});
        } else if (std.mem.eql(u8, name, "jobs")) {
            cli.jobs = std.fmt.parseInt(usize, value, 10) catch fatal(io, "invalid value '{s}' for '--jobs'", .{value});
        } else if (std.mem.eql(u8, name, "ci")) {
            cli.ci = true;
        } else if (std.mem.eql(u8, name, "since-hours")) {
            cli.since_hours = std.fmt.parseFloat(f64, value) catch fatal(io, "invalid value '{s}' for '--since-hours'", .{value});
        } else if (std.mem.eql(u8, name, "state-dir")) {
            cli.state_dir = value;
        } else if (std.mem.eql(u8, name, "skip-windows-update")) {
            cli.skip_windows_update = true;
        } else if (std.mem.eql(u8, name, "skip-wsl")) {
            cli.skip_wsl = true;
        } else if (std.mem.eql(u8, name, "skip-wsl-distros")) {
            cli.skip_wsl_distros = true;
        } else if (std.mem.eql(u8, name, "skip-defender")) {
            cli.skip_defender = true;
        } else if (std.mem.eql(u8, name, "skip-store-apps")) {
            cli.skip_store_apps = true;
        } else if (std.mem.eql(u8, name, "skip-powershell-modules")) {
            cli.skip_powershell_modules = true;
        } else if (std.mem.eql(u8, name, "skip-node")) {
            cli.skip_node = true;
        } else if (std.mem.eql(u8, name, "skip-rust")) {
            cli.skip_rust = true;
        } else if (std.mem.eql(u8, name, "skip-go")) {
            cli.skip_go = true;
        } else if (std.mem.eql(u8, name, "skip-flutter")) {
            cli.skip_flutter = true;
        } else if (std.mem.eql(u8, name, "skip-ruby")) {
            cli.skip_ruby = true;
        } else if (std.mem.eql(u8, name, "skip-composer")) {
            cli.skip_composer = true;
        } else if (std.mem.eql(u8, name, "skip-poetry")) {
            cli.skip_poetry = true;
        } else if (std.mem.eql(u8, name, "skip-uv-tools")) {
            cli.skip_uv_tools = true;
        } else if (std.mem.eql(u8, name, "skip-cleanup")) {
            cli.skip_cleanup = true;
        } else if (std.mem.eql(u8, name, "deep-clean")) {
            cli.deep_clean = true;
        } else if (std.mem.eql(u8, name, "skip-destructive")) {
            cli.skip_destructive = true;
        } else if (std.mem.eql(u8, name, "skip-git-lfs")) {
            cli.skip_git_lfs = true;
        } else if (std.mem.eql(u8, name, "skip-vcpkg")) {
            cli.skip_vcpkg = true;
        } else if (std.mem.eql(u8, name, "skip-conda")) {
            cli.skip_conda = true;
        } else if (std.mem.eql(u8, name, "skip-cloud-tools")) {
            cli.skip_cloud_tools = true;
        } else if (std.mem.eql(u8, name, "skip-infra-tools")) {
            cli.skip_infra_tools = true;
        } else if (std.mem.eql(u8, name, "skip-k8s-tools")) {
            cli.skip_k8s_tools = true;
        } else if (std.mem.eql(u8, name, "skip-starship")) {
            cli.skip_starship = true;
        } else if (std.mem.eql(u8, name, "skip-hugo")) {
            cli.skip_hugo = true;
        } else if (std.mem.eql(u8, name, "skip-vscode-extensions")) {
            cli.skip_vscode_extensions = true;
        } else if (std.mem.eql(u8, name, "skip-pip-health")) {
            cli.skip_pip_health = true;
        } else if (std.mem.eql(u8, name, "update-ollama-models")) {
            cli.update_ollama_models = true;
        } else if (std.mem.eql(u8, name, "update-powershell-help")) {
            cli.update_powershell_help = true;
        } else if (std.mem.eql(u8, name, "update-github-tools")) {
            cli.update_github_tools = true;
        } else if (std.mem.eql(u8, name, "notify-apply")) {
            cli.notify_apply = true;
        } else if (std.mem.eql(u8, name, "bypass-protection")) {
            cli.bypass_protection = true;
        } else if (std.mem.eql(u8, name, "winget-timeout-sec")) {
            cli.winget_timeout_sec = std.fmt.parseInt(u64, value, 10) catch fatal(io, "invalid value '{s}' for '--winget-timeout-sec'", .{value});
        } else if (std.mem.eql(u8, name, "ollama-timeout-sec")) {
            cli.ollama_timeout_sec = std.fmt.parseInt(u64, value, 10) catch fatal(io, "invalid value '{s}' for '--ollama-timeout-sec'", .{value});
        } else if (std.mem.eql(u8, name, "retry-count")) {
            cli.retry_count = std.fmt.parseInt(u32, value, 10) catch fatal(io, "invalid value '{s}' for '--retry-count'", .{value});
        } else if (std.mem.eql(u8, name, "profile")) {
            cli.profile = value;
        } else if (std.mem.eql(u8, name, "show-skipped")) {
            cli.show_skipped = true;
        } else if (std.mem.eql(u8, name, "schedule")) {
            cli.schedule = true;
        } else if (std.mem.eql(u8, name, "schedule-time")) {
            cli.schedule_time = value;
        } else {
            fatal(io, "unexpected argument '--{s}'", .{name});
        }
    }
    return cli;
}

// ─── Config ──────────────────────────────────────────────────────────────────

/// A manually-installed tool whose binaries come from a release/CI asset (no
/// package manager, no local git build). Providers: "github" (default),
/// "gitlab", "gitlab-artifact".
const GithubTool = struct {
    repo: []const u8 = "",
    install_dir: []const u8 = "",
    asset_regex: []const u8 = "",
    provider: ?[]const u8 = null,
    project_id: ?[]const u8 = null,
    git_ref: ?[]const u8 = null,
    job: ?[]const u8 = null,
    version_cmd: ?[]const u8 = null,
    version_regex: ?[]const u8 = null,
    id: ?[]const u8 = null,
    /// Pin to a named release tag (e.g. rolling "nightly"); version = published_at.
    tag: ?[]const u8 = null,
    /// How to order local vs latest: "semver", "build" (single integer), "date", or
    /// "auto" (default). Never falls back to stripping non-digits, which made
    /// `10649` compare as newer than `v0.3.0`.
    version_scheme: ?[]const u8 = null,
    /// Pick the newest release whose tag matches this regex instead of trusting
    /// `releases/latest`, which lags on repos that tag every build (llama.cpp).
    tag_regex: ?[]const u8 = null,
    /// pwsh snippets run before extract / after marker write (e.g. stop/start a service).
    pre_update: ?[]const u8 = null,
    post_update: ?[]const u8 = null,
};

const Config = struct {
    winget_skip_packages: []const []const u8 = &.{},
    skip_managers: []const []const u8 = &.{},
    pip_skip_packages: []const []const u8 = &.{},
    pip_ignore_health_packages: []const []const u8 = &.{},
    npm_skip_packages: []const []const u8 = &.{},
    chocolatey_skip_packages: []const []const u8 = &.{},
    store_app_skip_packages: []const []const u8 = &.{},
    gcloud_skip_components: []const []const u8 = &.{},
    az_skip_extensions: []const []const u8 = &.{},
    conda_skip_envs: []const []const u8 = &.{},
    vcpkg_skip_packages: []const []const u8 = &.{},
    windows_optional_features: []const []const u8 = &.{},
    /// Serialized back to compact JSON for the cross-manager python script.
    cross_manager_fallback: []const u8 = "{}",
    temp_cleanup_days: u32 = 7,
    log_retention_days: u32 = 14,
    github_tools: []const GithubTool = &.{},
    github_notification_ignore: []const []const u8 = &.{},
    github_notification_packages: []const NotifyPackage = &.{},
};

/// Explicit repo -> package mapping for gh-notify-releases. Without an entry a
/// release notification is only reported, never acted on.
const NotifyPackage = struct {
    repo: []const u8 = "",
    npm: ?[]const u8 = null,
    winget: ?[]const u8 = null,
};

fn jsonStringArray(gpa: Allocator, value: ?std.json.Value) []const []const u8 {
    const v = value orelse return &.{};
    if (v != .array) return &.{};
    var list: std.ArrayList([]const u8) = .empty;
    for (v.array.items) |item| {
        if (item == .string) list.append(gpa, gpa.dupe(u8, item.string) catch @panic("oom")) catch @panic("oom");
    }
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

fn jsonString(gpa: Allocator, value: ?std.json.Value) ?[]const u8 {
    const v = value orelse return null;
    if (v != .string) return null;
    return gpa.dupe(u8, v.string) catch @panic("oom");
}

fn jsonU32(value: ?std.json.Value, default: u32) u32 {
    const v = value orelse return default;
    return switch (v) {
        .integer => |n| if (n < 0) default else @intCast(n),
        else => default,
    };
}

/// Re-serialize a JSON value to compact text (used to embed the cross-manager map).
fn jsonToString(gpa: Allocator, value: std.json.Value) []const u8 {
    var list: std.ArrayList(u8) = .empty;
    writeJsonValue(gpa, &list, value);
    return list.toOwnedSlice(gpa) catch @panic("oom");
}

fn writeJsonValue(gpa: Allocator, list: *std.ArrayList(u8), value: std.json.Value) void {
    switch (value) {
        .null => list.appendSlice(gpa, "null") catch @panic("oom"),
        .bool => |b| list.appendSlice(gpa, if (b) "true" else "false") catch @panic("oom"),
        .integer => |n| list.appendSlice(gpa, fmtAlloc(gpa, "{d}", .{n})) catch @panic("oom"),
        .float => |f| list.appendSlice(gpa, fmtAlloc(gpa, "{d}", .{f})) catch @panic("oom"),
        .number_string => |s| list.appendSlice(gpa, s) catch @panic("oom"),
        .string => |s| {
            list.append(gpa, '"') catch @panic("oom");
            list.appendSlice(gpa, jsonEscape(gpa, s)) catch @panic("oom");
            list.append(gpa, '"') catch @panic("oom");
        },
        .array => |arr| {
            list.append(gpa, '[') catch @panic("oom");
            for (arr.items, 0..) |item, i| {
                if (i != 0) list.append(gpa, ',') catch @panic("oom");
                writeJsonValue(gpa, list, item);
            }
            list.append(gpa, ']') catch @panic("oom");
        },
        .object => |obj| {
            list.append(gpa, '{') catch @panic("oom");
            var it = obj.iterator();
            var first = true;
            while (it.next()) |entry| {
                if (!first) list.append(gpa, ',') catch @panic("oom");
                first = false;
                list.append(gpa, '"') catch @panic("oom");
                list.appendSlice(gpa, jsonEscape(gpa, entry.key_ptr.*)) catch @panic("oom");
                list.appendSlice(gpa, "\":") catch @panic("oom");
                writeJsonValue(gpa, list, entry.value_ptr.*);
            }
            list.append(gpa, '}') catch @panic("oom");
        },
    }
}

fn loadConfig(gpa: Allocator, io: Io, path: []const u8) Config {
    const text = Io.Dir.cwd().readFileAlloc(io, path, gpa, .limited(8 * 1024 * 1024)) catch return .{};
    const parsed = std.json.parseFromSlice(std.json.Value, gpa, text, .{}) catch |err| {
        fatal(io, "failed to parse config {s}: {t}", .{ path, err });
    };
    const root = parsed.value;
    if (root != .object) return .{};
    const obj = root.object;

    var config: Config = .{};
    config.winget_skip_packages = jsonStringArray(gpa, obj.get("WingetSkipPackages"));
    config.skip_managers = jsonStringArray(gpa, obj.get("SkipManagers"));
    config.pip_skip_packages = jsonStringArray(gpa, obj.get("PipSkipPackages"));
    config.pip_ignore_health_packages = jsonStringArray(gpa, obj.get("PipIgnoreHealthPackages"));
    config.npm_skip_packages = jsonStringArray(gpa, obj.get("NpmSkipPackages"));
    config.chocolatey_skip_packages = jsonStringArray(gpa, obj.get("ChocolateySkipPackages"));
    config.store_app_skip_packages = jsonStringArray(gpa, obj.get("StoreAppSkipPackages"));
    config.gcloud_skip_components = jsonStringArray(gpa, obj.get("GcloudSkipComponents"));
    config.az_skip_extensions = jsonStringArray(gpa, obj.get("AzSkipExtensions"));
    config.conda_skip_envs = jsonStringArray(gpa, obj.get("CondaSkipEnvs"));
    config.vcpkg_skip_packages = jsonStringArray(gpa, obj.get("VcpkgSkipPackages"));
    config.windows_optional_features = jsonStringArray(gpa, obj.get("WindowsOptionalFeatures"));
    config.temp_cleanup_days = jsonU32(obj.get("TempCleanupDays"), 7);
    config.log_retention_days = jsonU32(obj.get("LogRetentionDays"), 14);
    if (obj.get("CrossManagerFallback")) |v| {
        if (v == .object) config.cross_manager_fallback = jsonToString(gpa, v);
    }
    config.github_notification_ignore = jsonStringArray(gpa, obj.get("GithubNotificationIgnore"));
    if (obj.get("GithubNotificationPackages")) |v| {
        if (v == .array) {
            var pkgs: std.ArrayList(NotifyPackage) = .empty;
            for (v.array.items) |item| {
                if (item != .object) continue;
                pkgs.append(gpa, .{
                    .repo = jsonString(gpa, item.object.get("Repo")) orelse "",
                    .npm = jsonString(gpa, item.object.get("Npm")),
                    .winget = jsonString(gpa, item.object.get("Winget")),
                }) catch @panic("oom");
            }
            config.github_notification_packages = pkgs.toOwnedSlice(gpa) catch @panic("oom");
        }
    }
    if (obj.get("GithubTools")) |v| {
        if (v == .array) {
            var tools: std.ArrayList(GithubTool) = .empty;
            for (v.array.items) |item| {
                if (item != .object) continue;
                const o = item.object;
                tools.append(gpa, .{
                    .repo = jsonString(gpa, o.get("Repo")) orelse "",
                    .install_dir = jsonString(gpa, o.get("InstallDir")) orelse "",
                    .asset_regex = jsonString(gpa, o.get("AssetRegex")) orelse "",
                    .provider = jsonString(gpa, o.get("Provider")),
                    .project_id = jsonString(gpa, o.get("ProjectId")),
                    .git_ref = jsonString(gpa, o.get("Ref")),
                    .job = jsonString(gpa, o.get("Job")),
                    .version_cmd = jsonString(gpa, o.get("VersionCmd")),
                    .version_regex = jsonString(gpa, o.get("VersionRegex")),
                    .id = jsonString(gpa, o.get("Id")),
                    .tag = jsonString(gpa, o.get("Tag")),
                    .version_scheme = jsonString(gpa, o.get("VersionScheme")),
                    .tag_regex = jsonString(gpa, o.get("TagRegex")),
                    .pre_update = jsonString(gpa, o.get("PreUpdate")),
                    .post_update = jsonString(gpa, o.get("PostUpdate")),
                }) catch @panic("oom");
            }
            config.github_tools = tools.toOwnedSlice(gpa) catch @panic("oom");
        }
    }
    return config;
}

// ─── Task ────────────────────────────────────────────────────────────────────

const Task = struct {
    id: []const u8,
    category: []const u8,
    tags: []const []const u8,
    command: []const u8,
    args: []const []const u8,
    requires: []const u8,
    resource: []const u8 = "",
    /// Task ids that must finish (any status) before this one starts. Resources only
    /// guarantee mutual exclusion, not ordering, so tasks sharing a resource that must
    /// run in a specific sequence (e.g. pinning before `winget upgrade`) need this too.
    depends_on: []const []const u8 = &.{},
    skip_reason: ?[]const u8 = null,
    timeout_ms: ?u64 = null,
    acceptable_exit_codes: []const i32 = &.{},
    ok_on_timeout: bool = false,
    /// Substrings that mark a failure as permanent for this run. Retrying an
    /// upstream mirror mismatch or a paused service just burns the backoff.
    no_retry_patterns: []const []const u8 = &.{},

    fn skipIf(self: Task, condition: bool, reason: []const u8) Task {
        var task = self;
        if (condition and task.skip_reason == null) task.skip_reason = reason;
        return task;
    }
};

const winget_ok_codes = [_]i32{ -1978335188, -1978335189, -1978335212 };

const TaskSummary = struct {
    id: []const u8,
    category: []const u8,
    status: []const u8,
    duration_ms: u64,
    exit_code: ?i32,
    command: []const u8,
    args: []const []const u8,
    output_tail: [][]const u8,
};

// ─── Script/args builders ────────────────────────────────────────────────────

/// Prepended to python task scripts that need to free a locked executable.
/// Matches on the exact resolved path so an unrelated same-named binary
/// elsewhere on PATH is never touched.
const close_blockers_py =
    \\
    \\import os as _os, shutil as _sh, subprocess as _sp
    \\
    \\# pwsh starts noticeably faster than Windows PowerShell; 5.1 is only the fallback.
    \\_PS = _sh.which("pwsh") or "powershell"
    \\
    \\def close_locking_processes(bindir, names):
    \\    targets = [_os.path.join(bindir, n) for n in names]
    \\    quoted = ",".join("'" + t.replace("'", "''") + "'" for t in targets)
    \\    ps = (
    \\        "$t=@(" + quoted + ");"
    \\        "Get-Process -ErrorAction SilentlyContinue |"
    \\        " Where-Object { $t -contains $_.Path } |"
    \\        " ForEach-Object {"
    \\        "  Write-Output ('closed ' + $_.ProcessName + ' (pid ' + $_.Id + ')');"
    \\        "  Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }"
    \\    )
    \\    r = _sp.run(
    \\        [_PS, "-NoProfile", "-NonInteractive", "-Command", ps],
    \\        capture_output=True, text=True,
    \\    )
    \\    out = (r.stdout or "").strip()
    \\    if out:
    \\        print(out)
    \\    return bool(out)
    \\
;

/// Wrap a PowerShell command string as pwsh -NoProfile -NonInteractive -Command args.
fn pwshCmd(gpa: Allocator, script: []const u8) []const []const u8 {
    const args = gpa.alloc([]const u8, 4) catch @panic("oom");
    args[0] = "-NoProfile";
    args[1] = "-NonInteractive";
    args[2] = "-Command";
    args[3] = script;
    return args;
}

fn pyCmd(gpa: Allocator, script: []const u8) []const []const u8 {
    const args = gpa.alloc([]const u8, 2) catch @panic("oom");
    args[0] = "-c";
    args[1] = script;
    return args;
}

const cross_manager_pre =
    \\import json, shutil, subprocess, sys
    \\fallback = json.loads(
;
const cross_manager_post =
    \\)
    \\if not fallback:
    \\    print('No cross-manager fallback apps configured.')
    \\    sys.exit(0)
    \\choco = shutil.which('choco')
    \\scoop = shutil.which('scoop')
    \\if not choco and not scoop:
    \\    print('No alternate package managers (choco/scoop) available for fallback.')
    \\    sys.exit(0)
    \\failed = False
    \\for winget_id, alt in fallback.items():
    \\    print(f'Fallback check: {winget_id}')
    \\    if choco and alt.get('choco'):
    \\        result = subprocess.run(['choco', 'upgrade', alt['choco'], '-y', '--no-progress'], text=True)
    \\        failed = failed or result.returncode not in (0, 1)
    \\    if scoop and alt.get('scoop'):
    \\        result = subprocess.run(['scoop', 'update', alt['scoop']], text=True)
    \\        failed = failed or result.returncode != 0
    \\sys.exit(1 if failed else 0)
    \\
;

fn crossManagerArgs(gpa: Allocator, fallback_json: []const u8) []const []const u8 {
    const literal = concat(gpa, &.{ "\"", jsonEscape(gpa, fallback_json), "\"" });
    return pyCmd(gpa, concat(gpa, &.{ "\n", cross_manager_pre, literal, cross_manager_post }));
}

const gh_upgrade_script =
    \\import shutil, subprocess, sys
    \\path = (shutil.which('gh') or '').replace('\\', '/').lower()
    \\managed_markers = ['/scoop/apps/', '/scoop/shims/', '/chocolatey/lib/', '/microsoft/winget/packages/', '/windowsapps/']
    \\if any(marker in path for marker in managed_markers):
    \\    print(f'gh is managed by another package manager; handled elsewhere: {path}')
    \\    sys.exit(0)
    \\if not shutil.which('winget'):
    \\    print('winget missing; gh standalone update skipped.')
    \\    sys.exit(0)
    \\cmd = ['winget', 'upgrade', '--id', 'GitHub.cli', '--exact', '--include-unknown', '--disable-interactivity', '--accept-package-agreements', '--accept-source-agreements', '--silent']
    \\result = subprocess.run(cmd, capture_output=True, text=True)
    \\out = (result.stdout + result.stderr).strip()
    \\if out:
    \\    print(out)
    \\benign = ['No installed package found', 'No applicable update', 'No available upgrade', 'No newer package']
    \\if any(b in out for b in benign):
    \\    print('gh not tracked by winget (or already current); nothing to upgrade.')
    \\    sys.exit(0)
    \\sys.exit(result.returncode)
    \\
;

/// Shared PowerShell helpers (version compare) prepended to each managed-tool script.
const tool_script_prelude =
    \\$ErrorActionPreference = 'Stop'
    \\function Test-UpToDate($l, $r) {
    \\  if ($null -eq $l -or $null -eq $r) { return $false }
    \\  $lv = $null; $rv = $null
    \\  if ([version]::TryParse($l, [ref]$lv) -and [version]::TryParse($r, [ref]$rv)) {
    \\    return $lv -ge $rv
    \\  }
    \\  $ln = ($l -replace '\D', ''); $rn = ($r -replace '\D', '')
    \\  if ($ln -and $rn) { return [int64]$ln -ge [int64]$rn }
    \\  return $false
    \\}
    \\
;

const gitlab_release_body =
    \\$marker  = Join-Path $dir (".ue-" + ($repo -replace '[\\/:]', '_') + ".version")
    \\
    \\$local = if (Test-Path $marker) { (Get-Content $marker -Raw).Trim() } else { $null }
    \\$rel = Invoke-RestMethod -Uri "https://gitlab.com/api/v4/projects/$proj/releases/permalink/latest"
    \\$tag = $rel.tag_name
    \\Write-Host "$repo  local=$local  latest=$tag"
    \\if (Test-UpToDate $local $tag) { Write-Host 'up to date'; exit 0 }
    \\
    \\$dl = $rel.assets.links | Where-Object { $_.name -match $assetRe } | Select-Object -First 1
    \\if (-not $dl) { Write-Host "no asset matched /$assetRe/"; exit 1 }
    \\$url = if ($dl.direct_asset_url) { $dl.direct_asset_url } else { $dl.url }
    \\$tmp = Join-Path $env:TEMP $dl.name
    \\Write-Host "downloading $($dl.name)"
    \\Invoke-WebRequest -Uri $url -OutFile $tmp
    \\if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    \\if ($dl.name -match '\.zip$') { Expand-Archive -Path $tmp -DestinationPath $dir -Force }
    \\else { Copy-Item $tmp (Join-Path $dir $dl.name) -Force }
    \\Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    \\Set-Content -Path $marker -Value $tag -NoNewline
    \\Write-Host "updated $repo -> $tag"
    \\
;

/// GitLab release-asset provider: releases/permalink/latest, asset link by AssetRegex.
fn gitlabReleaseArgs(gpa: Allocator, tool: GithubTool) []const []const u8 {
    const script = concat(gpa, &.{
        tool_script_prelude,
        "$repo    = '",
        tool.repo,
        "'\n",
        "$dir     = '",
        tool.install_dir,
        "'\n",
        "$assetRe = '",
        tool.asset_regex,
        "'\n",
        "$proj    = '",
        tool.project_id orelse "",
        "'\n",
        gitlab_release_body,
    });
    return pwshCmd(gpa, script);
}

const gitlab_artifact_body =
    \\$marker = Join-Path $dir (".ue-" + ($repo -replace '[\\/:]', '_') + ".version")
    \\
    \\$pl = Invoke-RestMethod -Uri "https://gitlab.com/api/v4/projects/$proj/pipelines?ref=$ref&status=success&per_page=1"
    \\if (-not $pl) { Write-Host 'no successful pipeline found'; exit 1 }
    \\$latest = $pl[0].sha
    \\$local = if (Test-Path $marker) { (Get-Content $marker -Raw).Trim() } else { $null }
    \\Write-Host "$repo  local=$local  latest=$latest"
    \\if ($local -eq $latest) { Write-Host 'up to date'; exit 0 }
    \\
    \\$enc = [uri]::EscapeDataString($job)
    \\$url = "https://gitlab.com/api/v4/projects/$proj/jobs/artifacts/$ref/download?job=$enc"
    \\$tmp = Join-Path $env:TEMP ((($repo -replace '[\\/:]', '_')) + '-artifact.zip')
    \\Write-Host "downloading artifact (job=$job)"
    \\Invoke-WebRequest -Uri $url -OutFile $tmp
    \\if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    \\Expand-Archive -Path $tmp -DestinationPath $dir -Force
    \\Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    \\Set-Content -Path $marker -Value $latest -NoNewline
    \\Write-Host "updated $repo -> $latest"
    \\
;

/// GitLab CI-artifact provider: latest successful pipeline artifact for a Job.
/// Used by tools that publish builds as pipeline artifacts, not release assets.
fn gitlabArtifactArgs(gpa: Allocator, tool: GithubTool) []const []const u8 {
    const script = concat(gpa, &.{
        tool_script_prelude,
        "$repo = '",
        tool.repo,
        "'\n",
        "$dir  = '",
        tool.install_dir,
        "'\n",
        "$proj = '",
        tool.project_id orelse "",
        "'\n",
        "$ref  = '",
        tool.git_ref orelse "master",
        "'\n",
        "$job  = '",
        tool.job orelse "",
        "'\n",
        gitlab_artifact_body,
    });
    return pwshCmd(gpa, script);
}

const github_release_body =
    \\
    \\$marker = Join-Path $dir (".ue-" + ($repo -replace '[\\/:]', '_') + ".version")
    \\
    \\$local = $null
    \\if ($verCmd.Trim()) {
    \\  try {
    \\    Push-Location $dir
    \\    $out = & ([scriptblock]::Create($verCmd)) 2>&1 | Out-String
    \\    Pop-Location
    \\    if ($out -match $verRe) { $local = $matches[1] }
    \\  } catch { Write-Host "version probe failed: $_" }
    \\}
    \\# Fall back to the marker written by a previous run (tools with no queryable version).
    \\if ($null -eq $local -and (Test-Path $marker)) {
    \\  $local = (Get-Content $marker -Raw).Trim()
    \\}
    \\
    \\# Ordering key per scheme. Returns $null when the string does not fit the
    \\# scheme, which callers must treat as "cannot compare", never as up to date.
    \\function Get-VersionKey($v) {
    \\  if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    \\  $s = ([string]$v).Trim().TrimStart('vV')
    \\  switch ($scheme) {
    \\    'semver' {
    \\      if ($s -notmatch '^(\d+(?:\.\d+){0,3})') { return $null }
    \\      $parts = @($matches[1] -split '\.') + @('0','0','0','0')
    \\      return [version]::new([int]$parts[0], [int]$parts[1], [int]$parts[2], [int]$parts[3])
    \\    }
    \\    'build' {
    \\      # A single integer run, optionally prefixed (llama.cpp "b10711").
    \\      if ($s -notmatch '^[A-Za-z]*(\d+)$') { return $null }
    \\      return [int64]$matches[1]
    \\    }
    \\    'date' {
    \\      $d = ($s -replace '\D', '')
    \\      if ($d.Length -lt 8) { return $null }
    \\      return [int64]$d
    \\    }
    \\    default {
    \\      # auto: semver when it looks like semver, else a bare integer, else give up.
    \\      if ($s -match '^(\d+(?:\.\d+){1,3})') {
    \\        $parts = @($matches[1] -split '\.') + @('0','0','0','0')
    \\        return [version]::new([int]$parts[0], [int]$parts[1], [int]$parts[2], [int]$parts[3])
    \\      }
    \\      if ($s -match '^[A-Za-z]*(\d+)$') { return [int64]$matches[1] }
    \\      return $null
    \\    }
    \\  }
    \\}
    \\
    \\function Test-UpToDate($l, $r) {
    \\  $lk = Get-VersionKey $l
    \\  $rk = Get-VersionKey $r
    \\  if ($null -eq $lk -or $null -eq $rk) { return $false }
    \\  if ($lk.GetType() -ne $rk.GetType()) {
    \\    Write-Host "version scheme mismatch: local '$l' vs latest '$r'; treating as behind"
    \\    return $false
    \\  }
    \\  return $lk -ge $rk
    \\}
    \\
    \\$hdr = @{ 'User-Agent' = 'updateEverything' }
    \\if ($tagName) {
    \\  $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/tags/$tagName" -Headers $hdr
    \\  # Rolling tag: the name never changes, so publish time is the version.
    \\  $pa = $rel.published_at
    \\  $tag = if ($pa -is [datetime]) { $pa.ToUniversalTime().ToString('yyyyMMddHHmmss') } else { ([string]$pa) -replace '\D', '' }
    \\  if (-not $scheme) { $scheme = 'date' }
    \\} elseif ($tagRe) {
    \\  # releases/latest lags on repos that tag every build, so rank the matching
    \\  # tags ourselves and take the highest under the declared scheme.
    \\  $all = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases?per_page=50" -Headers $hdr
    \\  $cands = @($all | Where-Object { $_.tag_name -match $tagRe })
    \\  if (-not $cands) { Write-Host "no release tag matched /$tagRe/"; exit 1 }
    \\  $rel = $cands | Sort-Object -Property @{ Expression = { Get-VersionKey $_.tag_name } } -Descending | Select-Object -First 1
    \\  $tag = $rel.tag_name
    \\} else {
    \\  $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers $hdr
    \\  $tag = $rel.tag_name
    \\}
    \\$latest = $tag
    \\
    \\Write-Host "$repo  local=$local  latest=$tag"
    \\if (Test-UpToDate $local $latest) {
    \\  Write-Host "up to date"; exit 0
    \\}
    \\
    \\$dl = $rel.assets | Where-Object { $_.name -match $assetRe } | Select-Object -First 1
    \\if (-not $dl) { Write-Host "no asset matched /$assetRe/"; exit 1 }
    \\
    \\$tmp = Join-Path $env:TEMP $dl.name
    \\Write-Host "downloading $($dl.name)"
    \\Invoke-WebRequest -Uri $dl.browser_download_url -OutFile $tmp -Headers $hdr
    \\if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    \\if ($preCmd.Trim()) { Write-Host "pre-update: $preCmd"; & ([scriptblock]::Create($preCmd)) 2>&1 | Out-String | Write-Host }
    \\# Anything still running from InstallDir holds its DLLs and makes Expand-Archive fail.
    \\$dirFull = (Resolve-Path $dir).Path.TrimEnd('\') + '\'
    \\Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Path -and $_.Path.StartsWith($dirFull, [StringComparison]::OrdinalIgnoreCase) } | ForEach-Object {
    \\  Write-Host "stopping $($_.ProcessName) (pid $($_.Id)) - locks $dir"
    \\  Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    \\}
    \\if ($dl.name -match '\.zip$') {
    \\  Write-Host "extracting to $dir"
    \\  Expand-Archive -Path $tmp -DestinationPath $dir -Force
    \\} else {
    \\  Copy-Item $tmp (Join-Path $dir $dl.name) -Force
    \\}
    \\Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    \\Set-Content -Path $marker -Value $tag -NoNewline
    \\Write-Host "updated $repo -> $tag"
    \\if ($postCmd.Trim()) { Write-Host "post-update: $postCmd"; & ([scriptblock]::Create($postCmd)) 2>&1 | Out-String | Write-Host }
    \\
;

/// Build the pwsh args for a GitHub-release tool update: compare local version
/// to the latest release, download + extract the matching asset when behind.
fn githubReleaseInnerArgs(gpa: Allocator, tool: GithubTool) []const []const u8 {
    const script = concat(gpa, &.{
        "$ErrorActionPreference = 'Stop'\n",
        "$repo    = '",
        tool.repo,
        "'\n",
        "$dir     = '",
        tool.install_dir,
        "'\n",
        "$assetRe = '",
        tool.asset_regex,
        "'\n",
        "$verRe   = '",
        tool.version_regex orelse "(\\d+)",
        "'\n",
        "$tagName = '",
        tool.tag orelse "",
        "'\n",
        "$scheme  = '",
        tool.version_scheme orelse "",
        "'\n",
        "$tagRe   = '",
        tool.tag_regex orelse "",
        "'\n",
        "$preCmd  = @'\n",
        tool.pre_update orelse "",
        "\n'@\n",
        "$postCmd = @'\n",
        tool.post_update orelse "",
        "\n'@\n",
        "$verCmd  = @'\n",
        tool.version_cmd orelse "",
        "\n'@\n",
        github_release_body,
    });
    return pwshCmd(gpa, script);
}

/// Dispatch to the right provider builder for a managed tool.
fn githubReleaseArgs(gpa: Allocator, tool: GithubTool) []const []const u8 {
    const provider = tool.provider orelse "github";
    if (std.mem.eql(u8, provider, "gitlab")) return gitlabReleaseArgs(gpa, tool);
    if (std.mem.eql(u8, provider, "gitlab-artifact")) return gitlabArtifactArgs(gpa, tool);
    return githubReleaseInnerArgs(gpa, tool);
}

const github_notify_body =
    \\
    \\$raw = & gh api 'notifications?all=true&per_page=100' --jq '.[] | select(.subject.type=="Release") | .repository.full_name + "\t" + (.subject.title // "")' 2>&1
    \\if ($LASTEXITCODE -ne 0) {
    \\  if ("$raw" -match '(?i)(requires authentication|http 401|bad credentials)') {
    \\    Write-Host "SKIPPED: GitHub CLI is not authenticated"
    \\    exit 0
    \\  }
    \\  Write-Host "gh api failed: $raw"
    \\  exit 1
    \\}
    \\$known = @{ 'NousResearch/hermes-agent' = 'hermes task'; 'anthropics/claude-code' = 'self-updating, NpmSkipPackages' }
    \\$seen = [ordered]@{}
    \\foreach ($line in ($raw -split "`n")) {
    \\  $p = $line.Trim() -split "`t"
    \\  if ($p.Count -lt 2 -or -not $p[0]) { continue }
    \\  if (-not $seen.Contains($p[0])) { $seen[$p[0]] = $p[1] }
    \\}
    \\if ($seen.Count -eq 0) { Write-Host "no release notifications"; exit 0 }
    \\foreach ($repo in $seen.Keys) {
    \\  if ($repo -in $ignored) { continue }
    \\  $rel = $seen[$repo]
    \\  if ($managed -contains $repo) { Write-Host "managed   $repo  $rel  (GithubTools task)"; continue }
    \\  if ($known.Contains($repo)) { Write-Host "managed   $repo  $rel  ($($known[$repo]))"; continue }
    \\  # Only an explicit mapping may drive an install. Matching a repo name against
    \\  # `winget list` / `npm ls` output by substring upgraded the wrong package.
    \\  if (-not $pkgMap.Contains($repo)) {
    \\    Write-Host "UNMAPPED  $repo  $rel  https://github.com/$repo/releases"
    \\    continue
    \\  }
    \\  $m = $pkgMap[$repo]
    \\  if ($m.npm) {
    \\    if (-not $apply) { Write-Host "npm       $repo  $rel  -> npm i -g $($m.npm)@latest  (report only; use --notify-apply)"; continue }
    \\    Write-Host "npm       $repo  $rel  -> npm i -g $($m.npm)@latest"
    \\    & npm install -g "$($m.npm)@latest" 2>&1 | Out-String | Write-Host
    \\    continue
    \\  }
    \\  if ($m.winget) {
    \\    if (-not $apply) { Write-Host "winget    $repo  $rel  -> winget upgrade --id $($m.winget)  (report only; use --notify-apply)"; continue }
    \\    Write-Host "winget    $repo  $rel  -> winget upgrade --id $($m.winget)"
    \\    & winget upgrade --id $m.winget --exact --include-unknown --disable-interactivity --accept-package-agreements --accept-source-agreements 2>&1 | Out-String | Write-Host
    \\    continue
    \\  }
    \\  Write-Host "UNMAPPED  $repo  $rel  https://github.com/$repo/releases"
    \\}
    \\
;

/// Scan GitHub release notifications (subscribed repos) and report which are
/// covered by GithubTools, which winget tracks, and which are unmanaged.
fn githubNotifyArgs(
    gpa: Allocator,
    tools: []const GithubTool,
    ignored: []const []const u8,
    packages: []const NotifyPackage,
    apply: bool,
) []const []const u8 {
    var list: std.ArrayList(u8) = .empty;
    list.appendSlice(gpa, "$managed = @(") catch @panic("oom");
    for (tools, 0..) |tool, i| {
        if (i > 0) list.appendSlice(gpa, ",") catch @panic("oom");
        list.appendSlice(gpa, "'") catch @panic("oom");
        list.appendSlice(gpa, tool.repo) catch @panic("oom");
        list.appendSlice(gpa, "'") catch @panic("oom");
    }
    list.appendSlice(gpa, ")\n$ignored = @(") catch @panic("oom");
    list.appendSlice(gpa, psList(gpa, ignored, ", ")) catch @panic("oom");
    list.appendSlice(gpa, ")\n") catch @panic("oom");
    list.appendSlice(gpa, if (apply) "$apply = $True\n" else "$apply = $False\n") catch @panic("oom");
    list.appendSlice(gpa, "$pkgMap = [ordered]@{}\n") catch @panic("oom");
    for (packages) |pkg| {
        if (pkg.repo.len == 0) continue;
        list.appendSlice(gpa, concat(gpa, &.{
            "$pkgMap['",
            pkg.repo,
            "'] = @{ npm = ",
            if (pkg.npm) |v| concat(gpa, &.{ "'", v, "'" }) else "$null",
            "; winget = ",
            if (pkg.winget) |v| concat(gpa, &.{ "'", v, "'" }) else "$null",
            " }\n",
        })) catch @panic("oom");
    }
    const script = concat(gpa, &.{
        "$ErrorActionPreference = 'Stop'\n",
        list.items,
        github_notify_body,
    });
    return pwshCmd(gpa, script);
}

const winget_pin_skip_body =
    \\
    \\if (-not $pkgs) { Write-Output 'No winget skip-packages to pin.'; exit 0 }
    \\foreach ($p in $pkgs) {
    \\    winget pin add --id $p --exact --source winget --accept-source-agreements --disable-interactivity 2>&1 | Out-Null
    \\    Write-Output ("pinned: {0}" -f $p)
    \\}
    \\exit 0
;

/// Pin each skip package so `winget upgrade --all` (without --include-pinned)
/// leaves it untouched. Idempotent: pinning an already-pinned id just warns.
fn wingetPinSkipScript(gpa: Allocator, skip_packages: []const []const u8) []const u8 {
    return concat(gpa, &.{ "$pkgs = @(", psList(gpa, skip_packages, ", "), ")", winget_pin_skip_body });
}

const winget_per_package_body =
    \\)
    \\$raw = winget upgrade --include-unknown --disable-interactivity 2>&1 | Out-String
    \\$lines = $raw -split "`r?`n"
    \\$hdr = $null
    \\for ($i = 0; $i -lt $lines.Count; $i++) {
    \\    if ($lines[$i] -match '^Name\s+Id\s+Version') { $hdr = $i; break }
    \\}
    \\if ($null -eq $hdr) { Write-Output 'winget: no upgrade table found'; exit 0 }
    \\$idStart = $lines[$hdr].IndexOf('Id')
    \\$verStart = $lines[$hdr].IndexOf('Version')
    \\$pkgs = @()
    \\for ($i = $hdr + 1; $i -lt $lines.Count; $i++) {
    \\    $line = $lines[$i]
    \\    if ($line -match '^-{3,}$') { continue }
    \\    if ([string]::IsNullOrWhiteSpace($line)) { break }
    \\    # Trailer prose ("N upgrades available.", "The following packages ...") sits at the
    \\    # same offsets as the table, so slicing it yields junk ids like "le.".
    \\    if ($line -match '^\d+ (upgrades|package)' -or $line -match '^The following') { break }
    \\    if ($line.Length -le $idStart) { continue }
    \\    $len = [Math]::Min($verStart - $idStart, $line.Length - $idStart)
    \\    $id = $line.Substring($idStart, $len).Trim()
    \\    $name = $line.Substring(0, $idStart).Trim()
    \\    if ($name -and $id -match '^[A-Za-z0-9][A-Za-z0-9._+-]*\.[A-Za-z0-9][A-Za-z0-9._+-]*$' -and $id -notin $skip) {
    \\        $pkgs += [pscustomobject]@{ Id = $id; Name = $name }
    \\    }
    \\}
    \\$pkgs = $pkgs | Sort-Object Id -Unique
    \\Write-Output ("winget: {0} package(s) to upgrade" -f $pkgs.Count)
    \\
    \\# Registry install path for a display name, used to satisfy installers that demand
    \\# --location. Returns $null when the name matches zero or several entries.
    \\function Get-InstalledLocation($displayName) {
    \\    if (-not $displayName) { return $null }
    \\    $keys = @(
    \\        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    \\        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    \\        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    \\    )
    \\    $hits = @(Get-ItemProperty $keys -ErrorAction SilentlyContinue |
    \\        Where-Object { $_.DisplayName -eq $displayName -and $_.InstallLocation } |
    \\        Select-Object -ExpandProperty InstallLocation -Unique)
    \\    if ($hits.Count -eq 1 -and (Test-Path $hits[0])) { return $hits[0] }
    \\    return $null
    \\}
    \\
    \\# Codes winget returns when a listed upgrade cannot apply to this system; not our failure.
    \\$tolerated = @(-1978335188, -1978335212, -1978335215)
    \\# `upgrade` refuses these, but `install --force` over the top does the job: the manifest
    \\# does not apply to the installed package (-1978335189), or the new version uses a
    \\# different install technology and winget cannot uninstall the old one (-1978335090).
    \\$forceable = @(-1978335189, -1978335090)
    \\# The installer wants an explicit target directory (-1978335137). Retry with the path the
    \\# package is already installed to; give up if the registry does not identify one.
    \\$needsLocation = -1978335137
    \\# The de-elevated pass cannot touch machine-scope packages; the elevated pass covers them.
    \\$manual = @(-1978335226)
    \\$ok = @(); $forced = @(); $skipped = @(); $needsManual = @(); $failed = @()
    \\foreach ($pkg in $pkgs) {
    \\    $id = $pkg.Id
    \\    winget upgrade --id $id --exact --source winget --include-unknown --accept-package-agreements --accept-source-agreements --disable-interactivity --silent
    \\    $code = $LASTEXITCODE
    \\    if ($code -eq 0) { $ok += $id; continue }
    \\
    \\    if ($forceable -contains $code) {
    \\        Write-Output ("{0}: upgrade returned {1}; retrying as forced install" -f $id, $code)
    \\        winget install --id $id --exact --source winget --include-unknown --force --accept-package-agreements --accept-source-agreements --disable-interactivity --silent
    \\        $code2 = $LASTEXITCODE
    \\        if ($code2 -eq 0) { $forced += $id }
    \\        elseif ($tolerated -contains $code2) { $skipped += ("{0} ({1})" -f $id, $code2) }
    \\        else { $needsManual += ("{0} (upgrade {1}, forced install {2})" -f $id, $code, $code2) }
    \\        continue
    \\    }
    \\
    \\    if ($code -eq $needsLocation) {
    \\        $loc = Get-InstalledLocation $pkg.Name
    \\        if ($loc) {
    \\            Write-Output ("{0}: installer requires a location; retrying at {1}" -f $id, $loc)
    \\            winget install --id $id --exact --source winget --include-unknown --force --location "$loc" --accept-package-agreements --accept-source-agreements --disable-interactivity --silent
    \\            $code2 = $LASTEXITCODE
    \\            if ($code2 -eq 0) { $forced += $id }
    \\            else { $needsManual += ("{0} (needs --location, retry {1})" -f $id, $code2) }
    \\        } else {
    \\            $needsManual += ("{0} (needs --location, no install path found)" -f $id)
    \\        }
    \\        continue
    \\    }
    \\
    \\    if ($tolerated -contains $code) { $skipped += ("{0} ({1})" -f $id, $code) }
    \\    elseif ($manual -contains $code) { $needsManual += ("{0} ({1})" -f $id, $code) }
    \\    else { $failed += ("{0} ({1})" -f $id, $code) }
    \\}
    \\Write-Output ("upgraded ({0}): {1}" -f $ok.Count, ($ok -join ', '))
    \\Write-Output ("upgraded via forced install ({0}): {1}" -f $forced.Count, ($forced -join ', '))
    \\Write-Output ("not applicable ({0}): {1}" -f $skipped.Count, ($skipped -join ', '))
    \\Write-Output ("needs manual action ({0}): {1}" -f $needsManual.Count, ($needsManual -join ', '))
    \\Write-Output ("failed ({0}): {1}" -f $failed.Count, ($failed -join ', '))
    \\if ($deferFailures) { exit 0 }
    \\if ($failed.Count -gt 0) { exit 1 }
    \\exit 0
    \\
;

/// PowerShell that upgrades each outdated winget package individually.
///
/// `winget upgrade --all` aborts the whole batch without installing anything when any
/// listed package needs an install location: it prints the table, then exits
/// -1978335137 with zero installs. Per-package keeps one unusable entry from blocking
/// the rest. Packages in winget_skip_packages are excluded by pinning them (see the
/// winget-pin-skip task) and omitting `--include-pinned` here.
fn wingetPerPackageScript(gpa: Allocator, skip_packages: []const []const u8, defer_failures: bool) []const u8 {
    const defer_line = if (defer_failures) @as([]const u8, "$deferFailures = $True\n") else @as([]const u8, "$deferFailures = $False\n");
    return concat(gpa, &.{ defer_line, "\n$skip = @(", psList(gpa, skip_packages, ", "), winget_per_package_body });
}

fn wingetUpgradeArgs(gpa: Allocator, skip_packages: []const []const u8) []const []const u8 {
    return pwshCmd(gpa, wingetPerPackageScript(gpa, skip_packages, false));
}

const store_apps_body =
    \\)
    \\function Get-AppxSnapshot {
    \\    $map = @{}
    \\    foreach ($p in Get-AppxPackage) {
    \\        if ($p.Name -in $skip) { continue }
    \\        $map[$p.PackageFamilyName] = [pscustomobject]@{ Name = $p.Name; Version = $p.Version }
    \\    }
    \\    return $map
    \\}
    \\$before = Get-AppxSnapshot
    \\Write-Output ("store: {0} package(s) before scan" -f $before.Count)
    \\
    \\$ns = 'root\cimv2\mdm\dmmap'
    \\$cls = 'MDM_EnterpriseModernAppManagement_AppManagement01'
    \\$obj = Get-CimInstance -Namespace $ns -ClassName $cls -ErrorAction SilentlyContinue
    \\if (-not $obj) { Write-Output 'store: MDM app-management class unavailable; cannot trigger a scan.'; exit 0 }
    \\$res = Invoke-CimMethod -InputObject $obj -MethodName UpdateScanMethod -ErrorAction SilentlyContinue
    \\Write-Output ("store: UpdateScanMethod returned {0}" -f $res.ReturnValue)
    \\
    \\# Installs land asynchronously via AppX deployment; poll for changes rather than guess.
    \\$deadline = (Get-Date).AddSeconds(240)
    \\$changed = @{}
    \\$idle = 0
    \\while ((Get-Date) -lt $deadline) {
    \\    Start-Sleep 15
    \\    $now = Get-AppxSnapshot
    \\    foreach ($pfn in $now.Keys) {
    \\        $new = $now[$pfn]
    \\        $old = $before[$pfn]
    \\        if ($null -eq $old) { $changed[$pfn] = ("{0} (new {1})" -f $new.Name, $new.Version) }
    \\        elseif ($old.Version -ne $new.Version) { $changed[$pfn] = ("{0} {1} -> {2}" -f $new.Name, $old.Version, $new.Version) }
    \\    }
    \\    if (Get-Process -Name AppXSVC, AppInstaller, WinStore.App -ErrorAction SilentlyContinue) {
    \\        $idle = 0
    \\    } else {
    \\        if ($changed.Count -gt 0) { break }
    \\        # Nothing deploying and nothing moved: the scan had no work to do, so
    \\        # sitting out the rest of the window just burns run time.
    \\        $idle++
    \\        if ($idle -ge 3) { break }
    \\    }
    \\}
    \\if ($changed.Count -gt 0) {
    \\    Write-Output ("store: {0} app(s) updated" -f $changed.Count)
    \\    foreach ($v in $changed.Values) { Write-Output ("  " + $v) }
    \\} else {
    \\    Write-Output 'store: no app versions changed during the scan window.'
    \\    Write-Output 'store: this does NOT prove every Store app is current - the scan reports no per-app status,'
    \\    Write-Output 'store: and winget cannot enumerate msstore upgrades. Check Store > Library > Get updates to be sure.'
    \\}
    \\exit 0
    \\
;

/// PowerShell that asks the Store to update its apps, then reports what actually moved.
/// The MDM UpdateScanMethod returns 0 whether or not anything updates, so the only
/// trustworthy signal is an appx version snapshot taken before and after.
fn storeAppsScript(gpa: Allocator, skip_packages: []const []const u8) []const u8 {
    return concat(gpa, &.{ "\n$skip = @(", psList(gpa, skip_packages, ", "), store_apps_body });
}

const pip_upgrade_pre =
    \\
    \\import json, os, subprocess, sys, sysconfig
    \\if os.path.exists(os.path.join(sysconfig.get_path("stdlib"), "EXTERNALLY-MANAGED")):
    \\    print("pip: externally managed environment (PEP 668); skipping")
    \\    sys.exit(0)
    \\skip = {
;
const pip_upgrade_post =
    \\}
    \\health = subprocess.run([sys.executable, "-m", "pip", "check"], capture_output=True, text=True)
    \\if health.returncode != 0:
    \\    print("SKIPPED: pip dependency conflicts detected; upgrades deferred to preserve this shared environment")
    \\    for line in (health.stdout or health.stderr).strip().splitlines():
    \\        print("  " + line)
    \\    sys.exit(0)
    \\subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], check=False)
    \\r = subprocess.run([sys.executable, "-m", "pip", "list", "--outdated", "--format=json"], capture_output=True, text=True)
    \\pkgs = [p["name"] for p in json.loads(r.stdout or "[]") if p["name"].lower() not in skip]
    \\if not pkgs:
    \\    print("All pip packages up to date")
    \\    sys.exit(0)
    \\failed = []
    \\# One pip process per package costs ~3s of interpreter+resolver startup each.
    \\# Upgrade them in a single resolve, and only fall back to the per-package loop
    \\# when the batch fails, since one bad package aborts the whole batch.
    \\print(f"pip: upgrading {len(pkgs)} package(s) in one resolve")
    \\batch = subprocess.run(
    \\    [sys.executable, "-m", "pip", "install", "-U", "--upgrade-strategy", "only-if-needed", *pkgs],
    \\    check=False,
    \\).returncode
    \\if batch == 0:
    \\    for p in pkgs:
    \\        print("upgraded " + p)
    \\else:
    \\    print("pip: batch upgrade failed; retrying package by package")
    \\    for p in pkgs:
    \\        rc = subprocess.run([sys.executable, "-m", "pip", "install", "-U", "--upgrade-strategy", "only-if-needed", p], check=False).returncode
    \\        print(("upgraded " if rc == 0 else "FAILED ") + p)
    \\        if rc != 0:
    \\            failed.append(p)
    \\sys.exit(1 if failed else 0)
    \\
;

fn pipUpgradeArgs(gpa: Allocator, skip_packages: []const []const u8) []const []const u8 {
    return pyCmd(gpa, concat(gpa, &.{ pip_upgrade_pre, pyList(gpa, skip_packages, true), pip_upgrade_post }));
}

const python_venvs_script =
    \\
    \\import json, os, pathlib, subprocess, sys
    \\roots = [os.environ.get("WORKON_HOME"), str(pathlib.Path.home() / ".virtualenvs"), str(pathlib.Path.home() / ".venvs")]
    \\seen = set(); failed = []
    \\for root in filter(None, roots):
    \\    base = pathlib.Path(root)
    \\    if not base.is_dir(): continue
    \\    for cfg in base.rglob("pyvenv.cfg"):
    \\        exe = cfg.parent / "Scripts" / "python.exe"
    \\        if not exe.is_file() or str(exe).lower() in seen: continue
    \\        seen.add(str(exe).lower()); print(f"Updating venv: {cfg.parent}")
    \\        subprocess.run([str(exe), "-m", "pip", "install", "--upgrade", "pip"], check=False)
    \\        r = subprocess.run([str(exe), "-m", "pip", "list", "--outdated", "--format=json"], capture_output=True, text=True)
    \\        try: packages = [p["name"] for p in json.loads(r.stdout or "[]")]
    \\        except json.JSONDecodeError: packages = []
    \\        for package in packages:
    \\            if subprocess.run([str(exe), "-m", "pip", "install", "--upgrade", package]).returncode != 0: failed.append(f"{cfg.parent}:{package}")
    \\if not seen: print("SKIPPED: no virtual environments found in WORKON_HOME, ~/.virtualenvs, or ~/.venvs")
    \\else: print(f"python-venvs: scanned {len(seen)} environment(s)")
    \\sys.exit(1 if failed else 0)
    \\
;

const pip_health_pre =
    \\
    \\import subprocess, sys
    \\ignore = {
;
const pip_health_post =
    \\}
    \\r = subprocess.run([sys.executable, "-m", "pip", "check"], capture_output=True, text=True)
    \\lines = [l for l in r.stdout.strip().splitlines() if l.strip()]
    \\shown = [l for l in lines if not any(pkg in l.lower() for pkg in ignore)]
    \\if not lines:
    \\    print("pip check: no issues found")
    \\else:
    \\    if shown:
    \\        print("pip check found dependency conflicts (advisory, not failing the run):")
    \\        for l in shown:
    \\            print("  " + l)
    \\    ignored = len(lines) - len(shown)
    \\    if ignored:
    \\        print(f"pip check: {ignored} known/ignored conflict(s) suppressed")
    \\if r.stderr.strip():
    \\    print(r.stderr.strip())
    \\sys.exit(0)
    \\
;

/// Advisory only: pip check never updates anything, and dep conflicts are usually
/// pre-existing environment state. Report, but never fail the run.
fn pipHealthArgs(gpa: Allocator, ignore_packages: []const []const u8) []const []const u8 {
    return pyCmd(gpa, concat(gpa, &.{ pip_health_pre, pyList(gpa, ignore_packages, true), pip_health_post }));
}

const uv_self_update_script =
    \\
    \\import os, shutil, subprocess, sys, time
    \\
    \\uv = shutil.which("uv") or ""
    \\p = uv.replace("\\", "/").lower()
    \\if "/python" in p or "/scripts/" in p:
    \\    print("SKIPPED: uv is pip-managed; update handled by pip task.")
    \\    sys.exit(0)
    \\
    \\def update():
    \\    r = subprocess.run(["uv", "self", "update"], capture_output=True, text=True)
    \\    return r, (r.stdout or "") + (r.stderr or "")
    \\
    \\r, out = update()
    \\print(out.strip())
    \\# `uv self update` replaces uvx.exe alongside uv.exe. Anything hosting a uvx
    \\# tool (commonly an MCP server) holds that file open and the installer fails.
    \\# Close the holders by exact path and retry; whatever spawned them restarts them.
    \\if r.returncode != 0 and "being used by another process" in out:
    \\    closed = close_locking_processes(os.path.dirname(uv), ["uv.exe", "uvx.exe"])
    \\    if closed:
    \\        time.sleep(2)
    \\        r, out = update()
    \\        print(out.strip())
    \\    else:
    \\        print("SKIPPED: uv self-update blocked and no closable holder was found.")
    \\        sys.exit(0)
    \\sys.exit(r.returncode)
    \\
;

/// pipx owns poetry's venv when it installed it; `poetry self update` then
/// re-pins poetry's shared libraries down to its own declared bounds,
/// undoing pipx's upgrades on every run. Let pipx keep them current.
const poetry_self_update_script =
    \\
    \\import shutil, subprocess, sys
    \\if shutil.which("pipx"):
    \\    r = subprocess.run(["pipx", "list", "--short"], capture_output=True, text=True)
    \\    for line in (r.stdout or "").splitlines():
    \\        if line.split()[:1] == ["poetry"]:
    \\            print("SKIPPED: poetry is pipx-managed; upgrades handled by the pipx task.")
    \\            sys.exit(0)
    \\r = subprocess.run(
    \\    ["poetry", "self", "update", "--no-interaction"], capture_output=True, text=True
    \\)
    \\print(((r.stdout or "") + (r.stderr or "")).strip())
    \\sys.exit(r.returncode)
    \\
;

const uv_python_upgrade_script =
    \\
    \\import os, subprocess, sys
    \\# `python-downloads = "never"` (uv.toml or UV_PYTHON_DOWNLOADS) is a deliberate policy:
    \\# uv must not fetch interpreters. Installing them anyway is not this task's call, and
    \\# failing the run over a setting the user chose is noise. Report and skip.
    \\if os.environ.get("UV_PYTHON_DOWNLOADS", "").strip().lower() == "never":
    \\    print('uv python: downloads disabled (UV_PYTHON_DOWNLOADS=never); skipping interpreter upgrades')
    \\    sys.exit(0)
    \\r = subprocess.run(["uv", "python", "list", "--only-installed"], capture_output=True, text=True)
    \\if r.returncode != 0:
    \\    print("uv python list failed:", r.stderr.strip())
    \\    sys.exit(0)
    \\lines = [l.strip() for l in r.stdout.splitlines() if l.strip() and not l.startswith("cpython") or True]
    \\versions = []
    \\for line in r.stdout.splitlines():
    \\    parts = line.split()
    \\    if parts:
    \\        ver = parts[0]
    \\        if ver.startswith("cpython-") or ver.startswith("pypy-"):
    \\            ver_num = ver.split("-")[1] if "-" in ver else ver
    \\            if ver_num not in versions:
    \\                versions.append(ver_num)
    \\if not versions:
    \\    print("No uv-managed Python versions installed")
    \\    sys.exit(0)
    \\print(f"Upgrading {len(versions)} uv Python version(s): {', '.join(versions)}")
    \\r2 = subprocess.run(["uv", "python", "install"] + versions, capture_output=True, text=True)
    \\out = (r2.stdout or "") + (r2.stderr or "")
    \\print(out.strip())
    \\# Same policy, but configured in uv.toml rather than the environment.
    \\if r2.returncode != 0 and "downloads are not allowed" in out:
    \\    print('uv python: downloads disabled by uv config; skipping interpreter upgrades')
    \\    sys.exit(0)
    \\sys.exit(r2.returncode)
    \\
;

const npm_upgrade_pre =
    \\
    \\import json, os, shutil, subprocess, sys
    \\skip = [
;
const npm_upgrade_post =
    \\]
    \\# A half-downloaded browser dir makes puppeteer's postinstall fail forever instead
    \\# of re-downloading, so drop any version dir that has no .exe in it.
    \\cache = os.path.join(os.path.expanduser("~"), ".cache", "puppeteer")
    \\for browser in (os.listdir(cache) if os.path.isdir(cache) else []):
    \\    bdir = os.path.join(cache, browser)
    \\    if not os.path.isdir(bdir):
    \\        continue
    \\    for ver in os.listdir(bdir):
    \\        vdir = os.path.join(bdir, ver)
    \\        if not os.path.isdir(vdir):
    \\            continue
    \\        if not any(f.endswith(".exe") for _r, _d, fs in os.walk(vdir) for f in fs):
    \\            print("removing incomplete puppeteer browser: " + vdir)
    \\            shutil.rmtree(vdir, ignore_errors=True)
    \\# Bare "npm" is npm.cmd on Windows; subprocess needs the resolved path (or
    \\# shell=True) or CreateProcess fails with WinError 2.
    \\npm = shutil.which("npm")
    \\if not npm:
    \\    print("npm not found in PATH; skipping.")
    \\    sys.exit(0)
    \\# npm outdated exits 1 when anything is outdated, so parse stdout, not the code.
    \\r = subprocess.run([npm, "outdated", "-g", "--depth=0", "--json"], capture_output=True, text=True)
    \\try:
    \\    data = json.loads(r.stdout or "{}")
    \\    pkgs = [k for k in data.keys() if k not in skip and not k.startswith("npm")]
    \\except Exception:
    \\    pkgs = []
    \\if not pkgs:
    \\    print("npm: all global packages up to date")
    \\    sys.exit(0)
    \\failed = []
    \\for p in pkgs:
    \\    rc = subprocess.run([npm, "install", "-g", p]).returncode
    \\    print(("upgraded " if rc == 0 else "FAILED ") + p)
    \\    if rc != 0:
    \\        failed.append(p)
    \\sys.exit(1 if failed else 0)
    \\
;

fn npmUpgradeArgs(gpa: Allocator, skip_packages: []const []const u8) []const []const u8 {
    return pyCmd(gpa, concat(gpa, &.{ npm_upgrade_pre, pyList(gpa, skip_packages, false), npm_upgrade_post }));
}

const oh_my_posh_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("oh-my-posh") or "").replace("\\", "/").lower()
    \\managed = any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "windowsapps"])
    \\if managed:
    \\    print(f"oh-my-posh is managed by another package manager: {p}")
    \\    sys.exit(0)
    \\# Try winget first
    \\if shutil.which("winget"):
    \\    r = subprocess.run(
    \\        ["winget", "upgrade", "--id", "JanDeDobbeleer.OhMyPosh", "--exact",
    \\         "--include-unknown", "--disable-interactivity",
    \\         "--accept-package-agreements", "--accept-source-agreements", "--silent"],
    \\        capture_output=True, text=True
    \\    )
    \\    if r.returncode == 0 and "No installed package found" not in (r.stdout + r.stderr):
    \\        out = (r.stdout + r.stderr).strip()
    \\        if out:
    \\            print(out)
    \\        print("oh-my-posh checked via winget id JanDeDobbeleer.OhMyPosh.")
    \\        sys.exit(0)
    \\# Standalone upgrade
    \\r2 = subprocess.run(["oh-my-posh", "upgrade"], capture_output=True, text=True)
    \\import re
    \\out = re.sub(r'\x1b\][^\a]*(\a|\x1b\\)', '', r2.stdout + r2.stderr).strip()
    \\if out:
    \\    print(out)
    \\sys.exit(r2.returncode)
    \\
;

const vscode_extensions_script =
    \\
    \\import shutil, subprocess, sys
    \\if not shutil.which("code"):
    \\    print("code not in PATH")
    \\    sys.exit(0)
    \\# The CLI manages extensions without a running editor; requiring Code.exe
    \\# only left otherwise-updatable extensions behind.
    \\# Capture, don't inherit: code leaves helper processes holding the parent's
    \\# stdout pipe, so the runner would wait on EOF forever.
    \\try:
    \\    r2 = subprocess.run(["code", "--update-extensions"], capture_output=True, text=True, errors="ignore", timeout=300)
    \\except subprocess.TimeoutExpired:
    \\    print("code --update-extensions timed out after 300s")
    \\    sys.exit(1)
    \\out = (r2.stdout or "").strip()
    \\if not out:
    \\    print("SKIPPED: code --update-extensions produced no output; nothing verified")
    \\    sys.exit(r2.returncode)
    \\print(out)
    \\sys.exit(r2.returncode)
    \\
;

const yt_dlp_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("yt-dlp") or "").replace("\\", "/").lower()
    \\if "pipx/venvs" in p:
    \\    print("yt-dlp is managed by pipx; covered by the pipx task.")
    \\    sys.exit(0)
    \\if "/python" in p or "/scripts/" in p:
    \\    print("yt-dlp is managed by pip; covered by the pip task.")
    \\    sys.exit(0)
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "windowsapps"]):
    \\    print(f"yt-dlp is managed by another package manager: {p}")
    \\    sys.exit(0)
    \\r = subprocess.run(["yt-dlp", "-U"])
    \\sys.exit(r.returncode)
    \\
;

const gcloud_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("gcloud") or "").replace("\\", "/").lower()
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    \\    print("gcloud is managed by a package manager; skipping self-update.")
    \\    sys.exit(0)
    \\r = subprocess.run(["gcloud", "components", "update", "--quiet"])
    \\sys.exit(r.returncode)
    \\
;

const aws_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("aws") or "").replace("\\", "/").lower()
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "python", "scripts"]):
    \\    r = subprocess.run(["aws", "--version"], capture_output=True, text=True)
    \\    print(f"AWS CLI is managed ({r.stdout.strip()}); updating via package manager.")
    \\    sys.exit(0)
    \\print("Updating AWS CLI via pip...")
    \\r = subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "awscli"])
    \\sys.exit(r.returncode)
    \\
;

const terraform_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("terraform") or "").replace("\\", "/").lower()
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget", "tfenv"]):
    \\    print("Terraform is managed by a package manager; skipping.")
    \\    sys.exit(0)
    \\try:
    \\    r = subprocess.run(["terraform", "--version"], capture_output=True, text=True)
    \\    lines = r.stdout.strip().splitlines()
    \\    cur = lines[0] if lines else "unknown"
    \\    print(f"Current Terraform: {cur}")
    \\    import urllib.request, json as _json
    \\    with urllib.request.urlopen("https://api.github.com/repos/hashicorp/terraform/releases/latest", timeout=15) as resp:
    \\        data = _json.loads(resp.read())
    \\    latest = data["tag_name"].lstrip("v")
    \\    print(f"Latest Terraform: {latest}")
    \\    if cur and latest and latest in cur:
    \\        print("Terraform is current.")
    \\    else:
    \\        print(f"New version available: {latest}. Update via winget/scoop or download from https://releases.hashicorp.com/terraform/{latest}/")
    \\except Exception as e:
    \\    print(f"terraform check skipped: {e}")
    \\sys.exit(0)
    \\
;

const kubectl_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("kubectl") or "").replace("\\", "/").lower()
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    \\    print("kubectl is managed by a package manager; skipping.")
    \\    sys.exit(0)
    \\try:
    \\    r = subprocess.run(["kubectl", "version", "--client", "-o", "json"], capture_output=True, text=True)
    \\    import json
    \\    d = json.loads(r.stdout or "{}")
    \\    cur = d.get("clientVersion", {}).get("gitVersion", "unknown")
    \\    print(f"Current kubectl: {cur}")
    \\    import urllib.request
    \\    with urllib.request.urlopen("https://dl.k8s.io/release/stable.txt", timeout=15) as resp:
    \\        latest = resp.read().decode().strip()
    \\    print(f"Latest stable kubectl: {latest}")
    \\    if cur != "unknown" and latest in cur:
    \\        print("kubectl is current.")
    \\    else:
    \\        print(f"Update available. Install via: winget upgrade --id Kubernetes.kubectl")
    \\except Exception as e:
    \\    print(f"kubectl check skipped: {e}")
    \\sys.exit(0)
    \\
;

const starship_script =
    \\
    \\import shutil, subprocess, sys
    \\p = (shutil.which("starship") or "").replace("\\", "/").lower()
    \\if any(x in p for x in ["scoop/apps", "chocolatey/lib", "microsoft/winget"]):
    \\    print("starship is managed by a package manager; skipping.")
    \\    sys.exit(0)
    \\cmds = [["starship", "self", "update", "-y"], ["starship", "upgrade", "--yes"], ["starship", "self-update", "-y"]]
    \\for cmd in cmds:
    \\    try:
    \\        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    \\        err = (r.stderr or "").lower()
    \\        has_err = any(x in err for x in ["error", "unrecognized", "usage", "no such", "not found"])
    \\        if r.returncode == 0 and not has_err:
    \\            out = (r.stdout + r.stderr).strip()
    \\            if out:
    \\                print(out)
    \\            print("starship upgraded.")
    \\            sys.exit(0)
    \\    except Exception:
    \\        continue
    \\print("starship: no compatible self-update command found. Update via winget/scoop/choco.")
    \\sys.exit(0)
    \\
;

fn githubVersionCheckArgs(gpa: Allocator, bin: []const u8, ver_args: []const []const u8, repo: []const u8) []const []const u8 {
    const script = concat(gpa, &.{
        "\nimport subprocess, sys, urllib.request, json\ntry:\n    r = subprocess.run([\"",
        bin,
        "\", ",
        pyList(gpa, ver_args, false),
        "], capture_output=True, text=True, timeout=30)\n",
        "    cur = (r.stdout + r.stderr).strip().splitlines()[0] if (r.stdout + r.stderr).strip() else \"unknown\"\n",
        "    print(f\"Current ",
        bin,
        ": {cur}\")\n",
        "    with urllib.request.urlopen(f\"https://api.github.com/repos/",
        repo,
        "/releases/latest\", timeout=15) as resp:\n",
        "        data = json.loads(resp.read())\n",
        "    print(f\"Latest ",
        bin,
        ": {data['tag_name']}\")\n",
        "except Exception as e:\n",
        "    print(f\"",
        bin,
        " version check skipped: {e}\")\n",
        "sys.exit(0)\n",
    });
    return pyCmd(gpa, script);
}

const vcpkg_body =
    \\]
    \\print("Updating vcpkg baseline...")
    \\r1 = subprocess.run(["vcpkg", "update"])
    \\print("Upgrading vcpkg packages...")
    \\r2 = subprocess.run(["vcpkg", "upgrade", "--no-dry-run"])
    \\sys.exit(max(r1.returncode, r2.returncode))
    \\
;

fn vcpkgUpgradeArgs(gpa: Allocator, skip_packages: []const []const u8) []const []const u8 {
    return pyCmd(gpa, concat(gpa, &.{ "\nimport subprocess, sys\nskip = [", pyList(gpa, skip_packages, false), vcpkg_body }));
}

const conda_body =
    \\]
    \\print("Updating conda base environment...")
    \\r1 = subprocess.run(["conda", "update", "-n", "base", "conda", "-y"])
    \\r2 = subprocess.run(["conda", "update", "--all", "-y"])
    \\sys.exit(max(r1.returncode, r2.returncode))
    \\
;

fn condaUpgradeArgs(gpa: Allocator, skip_envs: []const []const u8) []const []const u8 {
    return pyCmd(gpa, concat(gpa, &.{ "\nimport subprocess, sys\nskip = [", pyList(gpa, skip_envs, false), conda_body }));
}

const ollama_body =
    \\try:
    \\    r = subprocess.run(["ollama", "list"], capture_output=True, text=True, timeout=min(TIMEOUT_SEC, 60))
    \\except Exception as e:
    \\    print(f"ollama list failed: {e}")
    \\    sys.exit(0)
    \\print(r.stdout.strip())
    \\
    \\def size_gb(parts):
    \\    # SIZE is two columns like "24 GB" / "5.2 GB" near the end
    \\    for i in range(len(parts) - 1):
    \\        if re.fullmatch(r"[0-9.]+", parts[i]) and parts[i + 1].upper() in ("GB", "MB", "KB", "B"):
    \\            v = float(parts[i])
    \\            u = parts[i + 1].upper()
    \\            return v if u == "GB" else v / 1024 if u == "MB" else v / 1048576 if u == "KB" else v / 1073741824
    \\    return None
    \\
    \\models = []
    \\for line in r.stdout.strip().splitlines()[1:]:
    \\    parts = line.split()
    \\    if parts:
    \\        models.append((parts[0], size_gb(parts)))
    \\if not models:
    \\    print("No Ollama models found.")
    \\    sys.exit(0)
    \\
    \\updated, unchanged, skipped = [], [], []
    \\for name, gb in models:
    \\    if gb is not None and gb > MAX_GB:
    \\        print(f"Skipping large model ({gb:.0f} GB > {MAX_GB:.0f} GB): {name}")
    \\        skipped.append(name)
    \\        continue
    \\    print(f"Pulling Ollama model: {name}")
    \\    try:
    \\        rc = subprocess.run(["ollama", "pull", name], timeout=per_model).returncode
    \\        (updated if rc == 0 else unchanged).append(name)
    \\    except subprocess.TimeoutExpired:
    \\        print(f"  timed out after {per_model}s; left unchanged: {name}")
    \\        unchanged.append(name)
    \\    except Exception as e:
    \\        print(f"  {name}: {e}")
    \\        unchanged.append(name)
    \\
    \\print(f"Ollama: {len(updated)} refreshed, {len(unchanged)} unchanged, {len(skipped)} skipped (too large).")
    \\if unchanged:
    \\    print(f"  unchanged (local/unavailable): {', '.join(unchanged)}")
    \\sys.exit(0)
    \\
;

/// Advisory + bounded. Locally-built models have no registry source, oversized
/// models are skipped, and a short per-model timeout caps any single hang. Never
/// fails the run; reports what changed and what was skipped.
fn ollamaModelsUpgradeArgs(gpa: Allocator, timeout_sec: u64) []const []const u8 {
    const timeout = fmtAlloc(gpa, "{d}", .{timeout_sec});
    const head = concat(gpa, &.{
        "\nimport re, subprocess, sys\n",
        "MAX_GB = 20.0                       # skip models larger than this\n",
        "per_model = min(",
        timeout,
        ", 300) # hard cap per pull\n",
    });
    return pyCmd(gpa, concat(gpa, &.{ head, replaceAll(gpa, ollama_body, "TIMEOUT_SEC", timeout) }));
}

const claude_script =
    \\
    \\import shutil, subprocess, sys
    \\if shutil.which("winget"):
    \\    r = subprocess.run(
    \\        ["winget", "upgrade", "--id", "Anthropic.Claude", "--exact",
    \\         "--include-unknown", "--disable-interactivity",
    \\         "--accept-package-agreements", "--accept-source-agreements", "--silent"],
    \\        capture_output=True, text=True
    \\    )
    \\    out = (r.stdout + r.stderr).strip()
    \\    if "No installed package found" not in out:
    \\        if out:
    \\            print(out)
    \\        print("claude checked via winget id Anthropic.Claude.")
    \\        sys.exit(0)
    \\r2 = subprocess.run(["claude", "update"])
    \\sys.exit(r2.returncode)
    \\
;

const codex_script =
    \\
    \\import shutil, subprocess, sys
    \\if not shutil.which("winget"):
    \\    print("winget not available; skipping codex upgrade.")
    \\    sys.exit(0)
    \\# Kill codex processes that lock the exe
    \\procs = subprocess.run(["tasklist", "/FI", "IMAGENAME eq codex-x86_64-pc-windows-msvc.exe", "/NH"],
    \\                        capture_output=True, text=True, errors="ignore")
    \\if "codex" in procs.stdout.lower():
    \\    subprocess.run(["taskkill", "/F", "/IM", "codex-x86_64-pc-windows-msvc.exe"], capture_output=True)
    \\    import time; time.sleep(2)
    \\r = subprocess.run(
    \\    ["winget", "upgrade", "--id", "OpenAI.Codex", "--exact", "--source", "winget",
    \\     "--include-unknown", "--disable-interactivity",
    \\     "--accept-package-agreements", "--accept-source-agreements", "--silent", "--force"],
    \\    capture_output=True, text=True
    \\)
    \\out = (r.stdout + r.stderr).strip()
    \\if out:
    \\    print(out)
    \\# winget exit codes arrive unsigned on Windows; normalize to signed int32
    \\rc = r.returncode
    \\if rc and rc > 0x7FFFFFFF:
    \\    rc -= 0x100000000
    \\low = out.lower()
    \\benign = rc in (0, -1978335188, -1978335189, -1978335212) or \
    \\    "no available upgrade" in low or "no newer package" in low or \
    \\    "no installed package found" in low
    \\if benign:
    \\    print("codex checked via winget id OpenAI.Codex.")
    \\    sys.exit(0)
    \\sys.exit(rc)
    \\
;

fn advisoryScript(gpa: Allocator, bin: []const u8, msg: []const u8) []const []const u8 {
    const script = concat(gpa, &.{
        "import shutil, sys\np = shutil.which(\"", bin,                  "\")\nif p:\n    print(f\"",
        bin,                                       ": {p}\")\nprint(\"", msg,
        "\")\nsys.exit(0)\n",
    });
    return pyCmd(gpa, script);
}

const cleanup_body =
    \\
    \\cutoff = time.time() - days * 86400
    \\temp_dirs = []
    \\t = os.environ.get("TEMP") or os.environ.get("TMP")
    \\if t:
    \\    temp_dirs.append(t)
    \\if platform.system() == "Windows":
    \\    temp_dirs.append(r"C:\Windows\Temp")
    \\
    \\for td in temp_dirs:
    \\    p = pathlib.Path(td)
    \\    if not p.is_dir():
    \\        continue
    \\    if skip_destructive:
    \\        print(f"Skipping temp cleanup (--skip-destructive): {td}")
    \\        continue
    \\    print(f"Cleaning temp files older than {days} day(s): {td}")
    \\    cleaned = 0
    \\    try:
    \\        items = list(p.iterdir())
    \\    except OSError as exc:
    \\        print(f"  Skipped {td}: {exc.strerror or exc}")
    \\        continue
    \\    for item in items:
    \\        try:
    \\            st = item.stat()
    \\            if st.st_mtime < cutoff and "WinGet" not in str(item):
    \\                if item.is_dir():
    \\                    import shutil
    \\                    shutil.rmtree(item, ignore_errors=True)
    \\                else:
    \\                    item.unlink(missing_ok=True)
    \\                cleaned += 1
    \\        except Exception:
    \\            pass
    \\    print(f"  Cleaned {cleaned} item(s)")
    \\
    \\if platform.system() == "Windows":
    \\    import subprocess
    \\    subprocess.run(["ipconfig", "/flushdns"], capture_output=True)
    \\    print("DNS cache flushed.")
    \\    if not skip_destructive:
    \\        import shutil as _sh
    \\        ps = _sh.which("pwsh") or "powershell"
    \\        subprocess.run([ps, "-NoProfile", "-Command", "Clear-RecycleBin -Force -ErrorAction SilentlyContinue"], capture_output=True)
    \\        print("Recycle bin emptied.")
    \\    if deep and not skip_destructive:
    \\        print("Running DISM component cleanup (this may take a while)...")
    \\        subprocess.run(["DISM.exe", "/Online", "/Cleanup-Image", "/StartComponentCleanup"])
    \\
    \\# Stale binary scan
    \\print("Checking for orphaned binaries in PATH...")
    \\path_dirs = [p for p in os.environ.get("PATH", "").split(os.pathsep) if p and os.path.isdir(p)]
    \\managed = ["scoop\\apps", "chocolatey\\lib", "Microsoft\\WinGet", "pipx\\venvs",
    \\           "Python\\Scripts", ".cargo\\bin", "node_modules\\.bin", ".dotnet\\tools",
    \\           "mise", "uv\\bin", "volta\\bin"]
    \\exclude_pfx = ["api-ms-win-", "ext-ms-win-", "concrt", "msvcp", "vcruntime", "msvcrt"]
    \\orphans = 0
    \\for d in path_dirs:
    \\    for exe in pathlib.Path(d).glob("*.exe"):
    \\        is_managed = any(m in str(exe) for m in managed)
    \\        is_sys = any(x in str(exe).lower() for x in ["\\system32", "\\system\\", "\\windows\\"])
    \\        is_excl = any(exe.name.startswith(p) for p in exclude_pfx)
    \\        if not is_managed and not is_sys and not is_excl:
    \\            try:
    \\                if os.path.getmtime(exe) < cutoff:
    \\                    orphans += 1
    \\            except Exception:
    \\                pass
    \\if orphans:
    \\    print(f"Stale binary scan: {orphans} orphaned .exe(s) older than {days} day(s) found in PATH.")
    \\else:
    \\    print("Stale binary scan: no orphaned binaries found.")
    \\sys.exit(0)
    \\
;

fn cleanupArgs(gpa: Allocator, days: u32, deep: bool, skip_destructive: bool) []const []const u8 {
    const head = concat(gpa, &.{
        "\nimport os, sys, time, pathlib, platform\n",
        "days = ",
        fmtAlloc(gpa, "{d}", .{days}),
        "\n",
        "deep = ",
        if (deep) "True" else "False",
        "\n",
        "skip_destructive = ",
        if (skip_destructive) "True" else "False",
        "\n",
    });
    return pyCmd(gpa, concat(gpa, &.{ head, cleanup_body }));
}

const self_update_script =
    \\
    \\import sys, urllib.request, json
    \\try:
    \\    print("Checking for script updates...")
    \\    with urllib.request.urlopen(
    \\        "https://api.github.com/repos/YoshKoz/updateEverything/releases/latest", timeout=15
    \\    ) as resp:
    \\        data = json.loads(resp.read())
    \\    latest = data["tag_name"].lstrip("v").strip()
    \\    current = "7.0.0"
    \\    print(f"Current: {current} | Latest: {latest}")
    \\    if latest and latest != current:
    \\        print(f"Update available: {current} → {latest}")
    \\        print(f"Download from: https://github.com/YoshKoz/updateEverything/releases/tag/v{latest}")
    \\    else:
    \\        print("Already at latest version.")
    \\except Exception as e:
    \\    msg = str(e)
    \\    if "404" in msg or "Not Found" in msg:
    \\        print("Self-update check skipped: no releases found on GitHub.")
    \\    else:
    \\        print(f"Self-update check skipped: {e}")
    \\sys.exit(0)
    \\
;

/// Git upgrades abort while any Git-shipped exe is running; Claude Code's
/// statusline respawns bash.exe every few seconds, so kill-loops and
/// /FORCECLOSEAPPLICATIONS both lose the race. Rename bash.exe away so
/// respawns fail harmlessly, upgrade, then clean up.
const winget_git_script =
    \\
    \\if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Output 'Git upgrade skipped: requires elevation.'; exit 0 }
    \\$gitDir = 'C:\Program Files\Git'
    \\if (-not (Test-Path (Join-Path $gitDir 'bin\bash.exe'))) { Write-Output 'Git for Windows not found; skipping.'; exit 0 }
    \\$list = winget list --id Git.Git --exact --upgrade-available --accept-source-agreements --disable-interactivity 2>&1 | Out-String
    \\if ($list -notmatch 'Git\.Git') { Write-Output 'Git: no upgrade available.'; exit 0 }
    \\Write-Output 'Git upgrade available. Stopping bash.exe and parking it (statusline respawns it)...'
    \\$bashPaths = @((Join-Path $gitDir 'usr\bin\bash.exe'), (Join-Path $gitDir 'bin\bash.exe'))
    \\Get-Process | Where-Object { $_.Path -in $bashPaths } | Stop-Process -Force -ErrorAction SilentlyContinue
    \\$renamed = @()
    \\foreach ($p in $bashPaths) {
    \\    if (Test-Path $p) {
    \\        try { Rename-Item $p ($p + '.ue-hold') -Force; $renamed += $p }
    \\        catch { Write-Output "rename failed: $p -- $($_.Exception.Message)" }
    \\    }
    \\}
    \\Get-Process | Where-Object { $_.Path -in $bashPaths } | Stop-Process -Force -ErrorAction SilentlyContinue
    \\winget upgrade --id Git.Git --exact --source winget --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
    \\$code = $LASTEXITCODE
    \\foreach ($p in $renamed) {
    \\    if (Test-Path $p) { Remove-Item ($p + '.ue-hold') -Force -ErrorAction SilentlyContinue }
    \\    else { Rename-Item ($p + '.ue-hold') $p -Force -ErrorAction SilentlyContinue }
    \\}
    \\if ($code -eq 0) { Write-Output 'Git upgraded.' } else { Write-Output "Git upgrade exit: $code" }
    \\exit $code
    \\
;

const winget_userscope_pre =
    \\
    \\if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Output 'Not elevated: user-scope packages already covered by main winget pass.'; exit 0 }
    \\$dir = Join-Path $env:TEMP ('ue-userscope-' + [guid]::NewGuid().ToString('N'))
    \\New-Item -ItemType Directory -Path $dir -Force | Out-Null
    \\$ps1File = Join-Path $dir 'run.ps1'
    \\$innerFile = Join-Path $dir 'upgrade.ps1'
    \\$log = Join-Path $dir 'run.log'
    \\# Inner script owns the exit code; run.ps1 only invokes it and records the result,
    \\# so the loop's own `exit` cannot swallow the DONE marker the outer wait polls for.
    \\$body = @'
    \\
;
const winget_userscope_post =
    \\
    \\'@
    \\Set-Content -Path $innerFile -Value $body -Encoding utf8
    \\@(
    \\    ('pwsh -NoProfile -ExecutionPolicy Bypass -File "' + $innerFile + '" *> "' + $log + '"'),
    \\    ('"exit: $LASTEXITCODE" | Add-Content -Path "' + $log + '"'),
    \\    ('"DONE" | Add-Content -Path "' + $log + '"')
    \\) | Set-Content -Path $ps1File -Encoding utf8
    \\runas /trustlevel:0x20000 "pwsh -NoProfile -ExecutionPolicy Bypass -File $ps1File" | Out-Null
    \\$deadline = (Get-Date).AddSeconds(900)
    \\while ((Get-Date) -lt $deadline) {
    \\    if ((Test-Path $log) -and (Select-String -Path $log -Pattern '^DONE' -Quiet -ErrorAction SilentlyContinue)) { break }
    \\    Start-Sleep 5
    \\}
    \\if (-not (Test-Path $log)) { Write-Output 'De-elevated winget produced no log (runas launch failed?).'; exit 1 }
    \\Get-Content $log | Select-Object -Last 40
    \\$m = Select-String -Path $log -Pattern '^exit: (-?\d+)' | Select-Object -Last 1
    \\$code = if ($m) { [int]$m.Matches[0].Groups[1].Value } else { 1 }
    \\Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
    \\if ($code -in 0, -1978335188, -1978335189, -1978335212) { exit 0 }
    \\exit $code
    \\
;

/// Elevated winget refuses to upgrade user-scope (zip/portable) packages, so
/// re-run the same per-package loop de-elevated via `runas /trustlevel:0x20000`.
fn wingetUserscopeScript(gpa: Allocator, skip_packages: []const []const u8) []const u8 {
    return concat(gpa, &.{ winget_userscope_pre, wingetPerPackageScript(gpa, skip_packages, true), winget_userscope_post });
}

const windows_update_script = "$s=New-Object -ComObject Microsoft.Update.Session;" ++
    "$r=$s.CreateUpdateSearcher().Search(\"IsInstalled=0 and Type='Software'\");" ++
    "if($r.Updates.Count-eq 0){Write-Output 'Windows: no updates available.';return};" ++
    "Write-Output \"Found $($r.Updates.Count) update(s). Downloading...\";" ++
    "$d=New-Object -ComObject Microsoft.Update.Downloader;" ++
    "$d.Updates=$r.Updates;$d.Download();" ++
    "$i=New-Object -ComObject Microsoft.Update.Installer;" ++
    "$i.Updates=$r.Updates;$ir=$i.Install();" ++
    "Write-Output \"Windows Update: $($r.Updates.Count) update(s) installed. Reboot required: $($ir.RebootRequired)\"";

const defender_script = "$mode = try { (Get-MpComputerStatus -ErrorAction Stop).AMRunningMode } catch { $null }; " ++
    "if ($mode -and $mode -ne 'Normal') { Write-Output \"Defender signature update skipped: AMRunningMode='$mode' (third-party AV active, Defender passive).\"; return }; " ++
    "try { Update-MpSignature -ErrorAction Stop; Write-Output 'Defender signatures updated.' } " ++
    "catch { Write-Output \"Defender update skipped: $($_.Exception.Message)\" }";

const powershell_help_script = "try { Update-Help -Force -ErrorAction Stop } " ++
    "catch { Write-Output \"PowerShell help not fully refreshed: $($_.Exception.Message)\" }";

/// winget refuses to upgrade MSYS2 ("cannot be upgraded using winget");
/// its own pacman is the supported path.
const msys2_script = "$bash=@('C:\\msys64\\usr\\bin\\bash.exe','C:\\tools\\msys64\\usr\\bin\\bash.exe'," ++
    "\"$env:SystemDrive\\msys64\\usr\\bin\\bash.exe\")|Where-Object{Test-Path $_}|Select-Object -First 1;" ++
    "if(-not $bash){Write-Output 'MSYS2 not installed; skipping.';return};" ++
    "Write-Output \"Updating MSYS2 via $bash\";" ++
    "& $bash -lc 'if [ -f /var/lib/pacman/db.lck ] && ! pgrep -x pacman >/dev/null 2>&1; then echo \"removing stale pacman lock\"; rm -f /var/lib/pacman/db.lck; fi; pacman -Syu --noconfirm';" ++
    "if($LASTEXITCODE-ne 0){Write-Output \"MSYS2 pacman exit $LASTEXITCODE\";exit $LASTEXITCODE}";

/// Every package manager is gated behind `sudo -n true` so a distro whose user
/// lacks passwordless sudo is skipped instead of hanging on a password prompt.
const wsl_distros_script =
    \\$prev=[Console]::OutputEncoding
    \\[Console]::OutputEncoding=[System.Text.Encoding]::Unicode
    \\$raw=wsl -l -q 2>$null
    \\[Console]::OutputEncoding=$prev
    \\$benign='(wsl2\.localhostForwarding setting has no effect|wsl: An internal error occurred\.|CreateInstance/CreateVm/ConfigureNetworking/0x8007054f|wsl: Failed to configure network|wsl: Failed to start the systemd user session)'
    \\$distros=@($raw|ForEach-Object{($_ -replace "`0",'').Trim()}|Where-Object{$_ -and $_ -notmatch 'docker-desktop' -and $_ -notmatch $benign}|Sort-Object -Unique)
    \\if($distros.Count -eq 0){Write-Output 'No WSL distros found.';return}
    \\$linuxScript=@'
    \\set -u
    \\
    \\resolve_any() {
    \\  for host in "$@"; do
    \\    if command -v getent >/dev/null 2>&1 && getent hosts "$host" >/dev/null 2>&1; then return 0; fi
    \\    if command -v nslookup >/dev/null 2>&1 && nslookup "$host" >/dev/null 2>&1; then return 0; fi
    \\    if command -v ping >/dev/null 2>&1 && ping -c 1 -W 2 "$host" >/dev/null 2>&1; then return 0; fi
    \\  done
    \\  return 1
    \\}
    \\
    \\if command -v apt-get >/dev/null 2>&1; then
    \\  if ! sudo -n true >/dev/null 2>&1; then
    \\    echo "Skipping apt-get: sudo requires a password"
    \\    exit 0
    \\  fi
    \\  if ! resolve_any archive.ubuntu.com security.ubuntu.com; then
    \\    echo "Skipping apt-get: WSL DNS/network is unavailable"
    \\    exit 0
    \\  fi
    \\  sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -o Acquire::Retries=2 update && sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y -o Dpkg::Options::=--force-confdef -o Dpkg::Options::=--force-confold full-upgrade && sudo -n env DEBIAN_FRONTEND=noninteractive apt-get -y autoremove
    \\elif command -v pacman >/dev/null 2>&1; then
    \\  if ! sudo -n true >/dev/null 2>&1; then
    \\    echo "Skipping pacman: sudo requires a password"
    \\    exit 0
    \\  fi
    \\  if ! resolve_any archlinux.org geo.mirror.pkgbuild.com; then
    \\    echo "Skipping pacman: WSL DNS/network is unavailable"
    \\    exit 0
    \\  fi
    \\  # Remove only a stale lock; a live pacman process keeps the lock intact.
    \\  if [ -e /var/lib/pacman/db.lck ] && ! pgrep -x pacman >/dev/null 2>&1; then
    \\    echo "removing stale pacman lock"
    \\    sudo -n rm -f /var/lib/pacman/db.lck
    \\  fi
    \\  pacman_lock_retries=0
    \\  while :; do
    \\    sudo -n pacman -Syu --noconfirm --needed
    \\    pacman_status=$?
    \\    if [ "$pacman_status" -eq 0 ]; then exit 0; fi
    \\    if [ ! -e /var/lib/pacman/db.lck ] || [ "$pacman_lock_retries" -ge 5 ]; then exit "$pacman_status"; fi
    \\    pacman_lock_retries=$((pacman_lock_retries + 1))
    \\    echo "pacman database locked; waiting 20 seconds (retry $pacman_lock_retries/5)"
    \\    sleep 20
    \\  done
    \\elif command -v dnf >/dev/null 2>&1; then
    \\  if ! sudo -n true >/dev/null 2>&1; then
    \\    echo "Skipping dnf: sudo requires a password"
    \\    exit 0
    \\  fi
    \\  sudo -n dnf -y upgrade
    \\elif command -v zypper >/dev/null 2>&1; then
    \\  if ! sudo -n true >/dev/null 2>&1; then
    \\    echo "Skipping zypper: sudo requires a password"
    \\    exit 0
    \\  fi
    \\  if ! resolve_any download.opensuse.org mirrors.opensuse.org; then
    \\    echo "Skipping zypper: WSL DNS/network is unavailable"
    \\    exit 0
    \\  fi
    \\  sudo -n zypper --non-interactive refresh && sudo -n zypper --non-interactive update
    \\else
    \\  echo "No supported Linux package manager found"
    \\fi
    \\'@
    \\$failed=@()
    \\foreach($d in $distros){
    \\  Write-Output "Updating WSL distro: $d"
    \\  wsl --distribution $d --exec sh -lc $linuxScript
    \\  if($LASTEXITCODE -ne 0){$failed+=$d;Write-Output "WSL distro failed: $d (exit $LASTEXITCODE)"}
    \\}
    \\if($failed.Count -gt 0){Write-Output ("failed WSL distros ({0}): {1}" -f $failed.Count,($failed -join ', '));exit 1}
    \\
;

const powershell_modules_script = "if(Get-Command Update-PSResource -EA SilentlyContinue){" ++
    "$mods=@(Get-PSResource -Name '*' -EA SilentlyContinue);" ++
    "if($mods.Count-eq 0){Write-Output 'No installed PSResource modules found.';return};" ++
    "$failed=@();" ++
    "foreach($m in ($mods|Sort-Object Name -Unique)){" ++
    "try{Update-PSResource -Name $m.Name -AcceptLicense -TrustRepository -EA Stop|Out-String|Write-Output}" ++
    "catch{Write-Output \"Not updated: $($m.Name) — $($_.Exception.Message)\";$failed+=$m.Name}};" ++
    "if($failed){Write-Output \"Left unchanged: $($failed -join ', ')\"}" ++
    "}elseif(Get-Command Update-Module -EA SilentlyContinue){" ++
    "$mods=@(Get-InstalledModule -EA SilentlyContinue);" ++
    "if($mods.Count-eq 0){Write-Output 'No installed PowerShellGet modules found.';return};" ++
    "$failed=@();" ++
    "foreach($m in $mods){" ++
    "try{Update-Module -Name $m.Name -Force -EA Stop}" ++
    "catch{Write-Output \"Not updated: $($m.Name) — $($_.Exception.Message)\";$failed+=$m.Name}};" ++
    "if($failed){Write-Output \"Left unchanged: $($failed -join ', ')\"}" ++
    "}else{Write-Output 'No PowerShell module updater found.'}";

fn windowsFeaturesScript(gpa: Allocator, features: []const []const u8) []const u8 {
    if (features.len == 0) {
        return "Write-Output 'No WindowsOptionalFeatures configured in update-config.json.'";
    }
    return concat(gpa, &.{
        "$features=@(",                               psList(gpa, features, ","),       ");",
        "$enabled=0;$skipped=0;",                     "foreach($f in $features){",      "$state=Get-WindowsOptionalFeature -Online -FeatureName $f -EA SilentlyContinue;",
        "if($state -and $state.State-ne 'Enabled'){", "Write-Output \"Enabling: $f\";", "Enable-WindowsOptionalFeature -Online -FeatureName $f -All -LimitAccess -EA SilentlyContinue|Out-Null;",
        "$enabled++",                                 "}else{$skipped++}}",             "Write-Output \"Windows Features: $enabled enabled, $skipped already present.\"",
    });
}

const appx_repair_script = "if(-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){" ++
    "Write-Output 'AppX repair skipped: requires elevation (Get-AppxPackage -AllUsers needs admin).';return};" ++
    "$pkgs=@(Get-AppxPackage -AllUsers -EA SilentlyContinue|" ++
    "Where-Object{$_.SignatureKind-eq 'Store'-and-not $_.IsFramework-and" ++
    "$_.Name-match '(Microsoft\\.WindowsStore|Microsoft\\.Store|Microsoft\\.WindowsCalculator|" ++
    "Microsoft\\.Windows\\.Photos|Microsoft\\.Windows\\.Camera|Microsoft\\.People|" ++
    "Microsoft\\.MSPaint|Microsoft\\.ScreenSketch|Microsoft\\.WindowsNotepad|" ++
    "Microsoft\\.WindowsTerminal)'});" ++
    "$repaired=0;" ++
    "foreach($p in $pkgs){" ++
    "try{$m=Get-AppxPackageManifest -Package $p -EA SilentlyContinue;" ++
    "if($m){Add-AppxPackage -Register -DisableDevelopmentMode -EA SilentlyContinue \"$($p.InstallLocation)\\AppxManifest.xml\" *>$null;$repaired++}}" ++
    "catch{Write-Output \"AppX repair failed for $($p.Name): $($_.Exception.Message)\"}};" ++
    "Write-Output \"AppX re-registration: $repaired package(s) re-registered.\"";

const mise_script = "import shutil, subprocess, sys\n" ++
    "p = (shutil.which(\"mise\") or \"\").replace(\"\\\\\", \"/\").lower()\n" ++
    "if \"winget\" in p or \"/microsoft/\" in p or \"/scoop/\" in p or \"/homebrew/\" in p:\n" ++
    "    print(\"mise is package-manager-managed; update handled by winget/scoop task\")\n" ++
    "    sys.exit(0)\n" ++
    "exe = shutil.which(\"mise\")\n" ++
    "if not exe:\n" ++
    "    print(\"mise not found on PATH; nothing to upgrade\")\n" ++
    "    sys.exit(0)\n" ++
    "r = subprocess.run([exe, \"self-upgrade\"])\n" ++
    "sys.exit(r.returncode)";

const chocolatey_pre = "if(-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){Write-Output 'Chocolatey skipped: requires elevation (run elevated to upgrade choco packages).';exit 0}; & choco upgrade all -y --no-progress";

/// choco requires elevation; when not elevated it prints a warning and blocks
/// ~20s on a "continue?" prompt. Skip cleanly instead.
fn chocolateyScript(gpa: Allocator, skip_packages: []const []const u8) []const u8 {
    var except: std.ArrayList(u8) = .empty;
    for (skip_packages) |pkg| {
        except.appendSlice(gpa, " --except '") catch @panic("oom");
        except.appendSlice(gpa, replaceAll(gpa, pkg, "'", "''")) catch @panic("oom");
        except.append(gpa, '\'') catch @panic("oom");
    }
    return concat(gpa, &.{ chocolatey_pre, except.items, "; exit $LASTEXITCODE" });
}

const scoop_script = "scoop update; if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }; scoop update *";

// ─── Task list ───────────────────────────────────────────────────────────────

fn mk(
    id: []const u8,
    category: []const u8,
    tags: []const []const u8,
    command: []const u8,
    args: []const []const u8,
) Task {
    return .{
        .id = id,
        .category = category,
        .tags = tags,
        .command = command,
        .args = args,
        .requires = command,
    };
}

const Opts = struct {
    requires: ?[]const u8 = null,
    resource: ?[]const u8 = null,
    timeout_sec: ?u64 = null,
    codes: ?[]const i32 = null,
    depends_on: ?[]const []const u8 = null,
    skip: ?[]const u8 = null,
    no_retry_patterns: ?[]const []const u8 = null,
};

fn with(base: Task, opts: Opts) Task {
    var task = base;
    if (opts.requires) |v| task.requires = v;
    if (opts.resource) |v| task.resource = v;
    if (opts.timeout_sec) |v| task.timeout_ms = v * 1000;
    if (opts.codes) |v| task.acceptable_exit_codes = v;
    if (opts.depends_on) |v| task.depends_on = v;
    if (opts.skip) |v| task.skip_reason = v;
    if (opts.no_retry_patterns) |v| task.no_retry_patterns = v;
    return task;
}

fn taskTable(gpa: Allocator, config: Config, cli: Cli) []Task {
    var pip_skip: std.ArrayList([]const u8) = .empty;
    pip_skip.appendSlice(gpa, config.pip_skip_packages) catch @panic("oom");
    for (config.pip_ignore_health_packages) |pkg| {
        var found = false;
        for (pip_skip.items) |existing| {
            if (std.mem.eql(u8, existing, pkg)) found = true;
        }
        if (!found) pip_skip.append(gpa, pkg) catch @panic("oom");
    }

    const node_skip: ?[]const u8 = if (cli.skip_node) "disabled by --skip-node" else null;
    const rust_skip: ?[]const u8 = if (cli.skip_rust) "disabled by --skip-rust" else null;
    const cloud_skip: ?[]const u8 = if (cli.skip_cloud_tools) "disabled by --skip-cloud-tools" else null;
    const infra_skip: ?[]const u8 = if (cli.skip_infra_tools) "disabled by --skip-infra-tools" else null;
    const k8s_skip: ?[]const u8 = if (cli.skip_k8s_tools) "disabled by --skip-k8s-tools" else null;
    const uv_tools_skip: ?[]const u8 = if (cli.skip_uv_tools) "disabled by --skip-uv-tools" else null;

    var tasks: std.ArrayList(Task) = .empty;
    const add = struct {
        fn f(list: *std.ArrayList(Task), a: Allocator, task: Task) void {
            list.append(a, task) catch @panic("oom");
        }
    }.f;

    // ── winget ───────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("winget-source", "package-manager", &.{ "windows", "winget" }, "winget", &.{ "source", "update" }), .{
        .resource = "winget",
        .timeout_sec = 300,
    }));
    // Kill portable-app processes that hold their own exe before winget upgrade
    add(&tasks, gpa, with(mk("winget-pre", "package-manager", &.{ "windows", "winget" }, "cmd", &.{
        "/c", "taskkill /F /IM codex-x86_64-pc-windows-msvc.exe 2>nul & exit 0",
    }), .{ .resource = "winget" }));
    add(&tasks, gpa, with(mk("winget-git", "package-manager", &.{ "windows", "winget", "git" }, "pwsh", pwshCmd(gpa, winget_git_script)), .{
        .resource = "winget",
        .timeout_sec = 900,
        .requires = "winget",
    }));
    add(&tasks, gpa, with(mk("winget-pin-skip", "package-manager", &.{ "windows", "winget" }, "pwsh", pwshCmd(gpa, wingetPinSkipScript(gpa, config.winget_skip_packages))), .{
        .resource = "winget",
        .timeout_sec = 300,
    }));
    add(&tasks, gpa, with(mk("winget", "package-manager", &.{ "windows", "winget" }, "pwsh", wingetUpgradeArgs(gpa, config.winget_skip_packages)), .{
        .resource = "winget",
        .timeout_sec = cli.winget_timeout_sec,
        .codes = &winget_ok_codes,
        .depends_on = &.{ "winget-pin-skip", "winget-git" },
    }));
    // A single per-package pass is authoritative; duplicate passes can replay stale entries.
    add(&tasks, gpa, with(mk("winget-pin-audit", "package-manager", &.{ "windows", "winget" }, "winget", &.{ "pin", "list" }), .{
        .codes = &winget_ok_codes,
    }));
    add(&tasks, gpa, with(mk("cross-manager", "package-manager", &.{ "windows", "winget", "scoop", "choco" }, "python", crossManagerArgs(gpa, config.cross_manager_fallback)), .{
        .resource = "package-manager",
    }));
    // ── store apps ───────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("store-apps", "system", &.{ "windows", "store" }, "pwsh", pwshCmd(gpa, storeAppsScript(gpa, config.store_app_skip_packages))), .{
        .resource = "winget",
        .timeout_sec = cli.winget_timeout_sec,
        .codes = &winget_ok_codes,
        .requires = "winget",
        .skip = if (cli.skip_store_apps) "disabled by --skip-store-apps" else null,
    }));
    // ── scoop ────────────────────────────────────────────────────────────────
    // scoop is a .ps1 shim — must be invoked via pwsh
    add(&tasks, gpa, with(mk("scoop", "package-manager", &.{ "windows", "scoop" }, "pwsh", pwshCmd(gpa, scoop_script)), .{
        .requires = "scoop",
    }));
    // ── chocolatey ───────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("chocolatey", "package-manager", &.{ "windows", "choco" }, "pwsh", pwshCmd(gpa, chocolateyScript(gpa, config.chocolatey_skip_packages))), .{
        .requires = "choco",
    }));
    // ── Windows Update ───────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("windows-update", "system", &.{"windows"}, "pwsh", pwshCmd(gpa, windows_update_script)), .{
        .resource = "windows-update",
        .timeout_sec = 7200,
        .requires = "pwsh",
        .skip = if (cli.skip_windows_update) "disabled by --skip-windows-update" else null,
    }));
    // ── Defender ─────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("defender", "system", &.{ "windows", "security" }, "pwsh", pwshCmd(gpa, defender_script)), .{
        .resource = "defender",
        .requires = "pwsh",
        .skip = if (cli.skip_defender) "disabled by --skip-defender" else null,
    }));
    // ── WSL ──────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("wsl", "system", &.{ "windows", "linux" }, "wsl", &.{"--update"}), .{
        .codes = &[_]i32{-1},
        .skip = if (cli.skip_wsl) "disabled by --skip-wsl" else null,
    }));
    add(&tasks, gpa, with(mk("wsl-distros", "system", &.{ "windows", "linux" }, "pwsh", pwshCmd(gpa, wsl_distros_script)), .{
        .resource = "wsl",
        .timeout_sec = 3600,
        .requires = "wsl",
        // A mirror mid-sync serves the same short Packages index to every retry.
        .no_retry_patterns = &.{
            "File has unexpected size",
            "Mirror sync in progress",
        },
        .skip = if (cli.skip_wsl or cli.skip_wsl_distros) "disabled by WSL skip flag" else null,
    }));
    // ── Windows Features ─────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("windows-features", "system", &.{ "windows", "system" }, "pwsh", pwshCmd(gpa, windowsFeaturesScript(gpa, config.windows_optional_features))), .{
        .requires = "pwsh",
    }));
    // ── AppX repair ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("appx-repair", "system", &.{ "windows", "store" }, "pwsh", pwshCmd(gpa, appx_repair_script)), .{
        .requires = "pwsh",
    }));
    // ── Linux packages (Arch/WSL) ────────────────────────────────────────────
    add(&tasks, gpa, with(mk("pacman", "package-manager", &.{ "linux", "arch" }, "sudo", &.{ "pacman", "-Syu", "--noconfirm" }), .{
        // On Windows, PATH resolves pacman to the MSYS2 install; running it here
        // races the dedicated msys2 task on the same package DB.
        .requires = "pacman",
        .skip = if (is_windows) "Linux-only; native MSYS2 is handled by the msys2 task" else null,
    }));
    // ── MSYS2 (native Windows) ───────────────────────────────────────────────
    add(&tasks, gpa, with(mk("msys2", "package-manager", &.{ "windows", "msys2" }, "pwsh", pwshCmd(gpa, msys2_script)), .{
        .resource = "msys2",
        .timeout_sec = 1800,
        .requires = "pwsh",
    }));
    // ── JavaScript ───────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("npm", "javascript", &.{"node"}, "python", npmUpgradeArgs(gpa, config.npm_skip_packages)), .{
        .requires = "npm",
        .skip = node_skip,
    }));
    add(&tasks, gpa, with(mk("pnpm", "javascript", &.{"node"}, "pnpm", &.{"self-update"}), .{ .skip = node_skip }));
    add(&tasks, gpa, with(mk("yarn", "javascript", &.{"node"}, "yarn", &.{ "global", "upgrade" }), .{ .skip = node_skip }));
    add(&tasks, gpa, with(mk("bun", "javascript", &.{"node"}, "bun", &.{"upgrade"}), .{ .skip = node_skip }));
    add(&tasks, gpa, with(mk("deno", "javascript", &.{"node"}, "deno", &.{"upgrade"}), .{ .skip = node_skip }));
    add(&tasks, gpa, with(mk("volta", "javascript", &.{"node"}, "python", advisoryScript(gpa, "volta", "Volta does not provide a stable non-interactive self-update. Install/update via winget/scoop/chocolatey for automatic coverage.")), .{
        .requires = "volta",
        .skip = node_skip,
    }));
    add(&tasks, gpa, with(mk("fnm", "javascript", &.{"node"}, "python", advisoryScript(gpa, "fnm", "fnm does not provide a self-update command. Install/update via winget/scoop/chocolatey for automatic coverage.")), .{
        .requires = "fnm",
        .skip = node_skip,
    }));
    add(&tasks, gpa, with(mk("nvm", "javascript", &.{"node"}, "python", advisoryScript(gpa, "nvm", "nvm (for Windows): update via winget upgrade --id CoreyButler.NVMforWindows")), .{
        .requires = "nvm",
        .skip = node_skip,
    }));
    // ── Python ───────────────────────────────────────────────────────────────
    add(&tasks, gpa, mk("pip", "python", &.{"python"}, "python", pipUpgradeArgs(gpa, pip_skip.items)));
    add(&tasks, gpa, mk("python-venvs", "python", &.{ "python", "venv" }, "python", pyCmd(gpa, python_venvs_script)));
    add(&tasks, gpa, with(mk("pip-health", "python", &.{ "python", "health" }, "python", pipHealthArgs(gpa, config.pip_ignore_health_packages)), .{
        .skip = if (cli.skip_pip_health) "disabled by --skip-pip-health" else null,
    }));
    // pipx defaults to the uv backend; force pip so it works when uv is not installed.
    add(&tasks, gpa, mk("pipx", "python", &.{"python"}, "pipx", &.{ "upgrade-all", "--backend", "pip" }));
    add(&tasks, gpa, with(mk("hermes", "agents", &.{ "python", "agents" }, "hermes", &.{ "update", "--yes", "--no-backup" }), .{
        .timeout_sec = 900,
        // A service stuck in a paused/pending SCM state needs manual attention;
        // the gateway ownership probe fails the same way on every attempt.
        .no_retry_patterns = &.{
            "has indeterminate status",
            "SCM service enumeration failed",
        },
    }));
    add(&tasks, gpa, mk("uv", "python", &.{"python"}, "python", pyCmd(gpa, concat(gpa, &.{ close_blockers_py, uv_self_update_script }))));
    // Must not overlap the `uv` task: that one has winget replace the uv shim, and a
    // spawn landing in that window dies with FileNotFound.
    add(&tasks, gpa, with(mk("uv-tools", "python", &.{"python"}, "uv", &.{ "tool", "upgrade", "--all" }), .{
        .depends_on = &.{"uv"},
        .skip = uv_tools_skip,
    }));
    add(&tasks, gpa, with(mk("uv-python", "python", &.{ "python", "uv" }, "python", pyCmd(gpa, uv_python_upgrade_script)), .{
        .requires = "uv",
        .skip = uv_tools_skip,
    }));
    add(&tasks, gpa, with(mk("poetry", "python", &.{"python"}, "python", pyCmd(gpa, poetry_self_update_script)), .{
        .requires = "poetry",
        .skip = if (cli.skip_poetry) "disabled by --skip-poetry" else null,
    }));
    add(&tasks, gpa, with(mk("conda", "python", &.{"python"}, "python", condaUpgradeArgs(gpa, config.conda_skip_envs)), .{
        .requires = "conda",
        .resource = "conda",
        .skip = if (cli.skip_conda) "disabled by --skip-conda" else null,
    }));
    // ── Rust ─────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("rustup", "systems-language", &.{"rust"}, "rustup", &.{"update"}), .{
        .resource = "rust",
        .skip = rust_skip,
    }));
    add(&tasks, gpa, with(mk("cargo", "systems-language", &.{"rust"}, "cargo", &.{ "install-update", "-a" }), .{
        .requires = "cargo-install-update",
        .resource = "rust",
        .skip = rust_skip,
    }));
    // ── Go ───────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("go", "systems-language", &.{"go"}, "go", &.{ "install", "golang.org/x/tools/gopls@latest" }), .{
        .skip = if (cli.skip_go) "disabled by --skip-go" else null,
    }));
    // ── PHP ──────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("composer", "runtime", &.{"php"}, "composer", &.{ "self-update", "--no-interaction" }), .{
        .skip = if (cli.skip_composer) "disabled by --skip-composer" else null,
    }));
    // ── Ruby ─────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("ruby-gems", "runtime", &.{"ruby"}, "gem", &.{"update"}), .{
        .skip = if (cli.skip_ruby) "disabled by --skip-ruby" else null,
    }));
    // ── Flutter ──────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("flutter", "systems-language", &.{"flutter"}, "flutter", &.{"upgrade"}), .{
        .skip = if (cli.skip_flutter) "disabled by --skip-flutter" else null,
    }));
    // ── Julia ────────────────────────────────────────────────────────────────
    add(&tasks, gpa, mk("juliaup", "systems-language", &.{"julia"}, "juliaup", &.{"update"}));
    add(&tasks, gpa, with(mk("gh", "dev-tools", &.{ "github", "dev-tools" }, "python", pyCmd(gpa, gh_upgrade_script)), .{
        .requires = "gh",
        .resource = "winget",
        .codes = &winget_ok_codes,
    }));
    // ── .NET ─────────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("dotnet-workloads", "dotnet", &.{"dotnet"}, "dotnet", &.{ "workload", "update" }), .{
        .timeout_sec = 3600,
    }));
    add(&tasks, gpa, mk("dotnet-tools", "dotnet", &.{"dotnet"}, "dotnet", &.{ "tool", "update", "--global", "--all" }));
    // ── PowerShell ───────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("powershell7", "system", &.{ "windows", "powershell" }, "winget", &.{
        "upgrade",                     "--id",                       "Microsoft.PowerShell",
        "--exact",                     "--include-unknown",          "--disable-interactivity",
        "--accept-package-agreements", "--accept-source-agreements", "--silent",
    }), .{
        .timeout_sec = 300,
        .codes = &winget_ok_codes,
        .requires = "winget",
    }));
    add(&tasks, gpa, with(mk("powershell-modules", "powershell", &.{"powershell"}, "pwsh", pwshCmd(gpa, powershell_modules_script)), .{
        .resource = "powershell-gallery",
        .requires = "pwsh",
        .skip = if (cli.skip_powershell_modules) "disabled by --skip-powershell-modules" else null,
    }));
    add(&tasks, gpa, with(mk("powershell-help", "powershell", &.{"powershell"}, "pwsh", pwshCmd(gpa, powershell_help_script)), .{
        .resource = "powershell-gallery",
        .requires = "pwsh",
        .skip = if (!cli.update_powershell_help) "opt-in via --update-powershell-help" else null,
    }));
    // ── Editor ───────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("vscode-extensions", "editor", &.{"vscode"}, "python", pyCmd(gpa, vscode_extensions_script)), .{
        .requires = "code",
        // The inner python already caps at 300s; this is the outer backstop so a
        // wedged `code` helper cannot hold the whole run for the 1800s default.
        .timeout_sec = 360,
        .skip = if (cli.skip_vscode_extensions) "disabled by --skip-vscode-extensions" else null,
    }));
    // ── Dev tools ────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("git-lfs", "dev-tools", &.{"git"}, "git", &.{ "lfs", "install", "--skip-repo" }), .{
        .requires = "git-lfs",
        .skip = if (cli.skip_git_lfs) "disabled by --skip-git-lfs" else null,
    }));
    add(&tasks, gpa, mk("gh-extensions", "dev-tools", &.{"github"}, "gh", &.{ "extension", "upgrade", "--all" }));
    add(&tasks, gpa, with(mk("devcontainer", "dev-tools", &.{"dev"}, "python", githubVersionCheckArgs(gpa, "devcontainer", &.{"--version"}, "devcontainers/cli")), .{
        .requires = "devcontainer",
    }));
    // ── Version managers ─────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("mise", "version-manager", &.{"mise"}, "python", pyCmd(gpa, mise_script)), .{
        .requires = "mise",
    }));
    add(&tasks, gpa, mk("mise-upgrade", "version-manager", &.{"mise"}, "mise", &.{"upgrade"}));
    // ── Media tools ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("yt-dlp", "media-tools", &.{"media"}, "python", pyCmd(gpa, yt_dlp_script)), .{
        .requires = "yt-dlp",
    }));
    // ── Shell tools ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("oh-my-posh", "shell", &.{"shell"}, "python", pyCmd(gpa, oh_my_posh_script)), .{
        .requires = "oh-my-posh",
    }));
    add(&tasks, gpa, with(mk("starship", "shell", &.{"shell"}, "python", pyCmd(gpa, starship_script)), .{
        .requires = "starship",
        .skip = if (cli.skip_starship) "disabled by --skip-starship" else null,
    }));
    add(&tasks, gpa, with(mk("zoxide", "shell", &.{"shell"}, "python", advisoryScript(gpa, "zoxide", "zoxide: update via your package manager (winget/scoop/choco/brew).")), .{
        .requires = "zoxide",
    }));
    add(&tasks, gpa, mk("tldr", "dev-tools", &.{"dev-tools"}, "tldr", &.{"--update"}));
    // ── Cloud tools ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("gcloud", "cloud", &.{"cloud"}, "python", pyCmd(gpa, gcloud_script)), .{
        .resource = "gcloud",
        .timeout_sec = 600,
        .requires = "gcloud",
        .skip = cloud_skip,
    }));
    add(&tasks, gpa, with(mk("az", "cloud", &.{"cloud"}, "az", &.{ "upgrade", "--all", "-y" }), .{
        .resource = "az",
        .timeout_sec = 600,
        .skip = cloud_skip,
    }));
    add(&tasks, gpa, with(mk("aws", "cloud", &.{"cloud"}, "python", pyCmd(gpa, aws_script)), .{
        .requires = "aws",
        .resource = "aws",
        .skip = cloud_skip,
    }));
    // ── Infrastructure tools ─────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("terraform", "infrastructure", &.{"infrastructure"}, "python", pyCmd(gpa, terraform_script)), .{
        .resource = "terraform",
        .requires = "terraform",
        .skip = infra_skip,
    }));
    add(&tasks, gpa, with(mk("pulumi", "infrastructure", &.{"infrastructure"}, "pulumi", &.{"upgrade"}), .{
        .timeout_sec = 600,
        .resource = "pulumi",
        .skip = infra_skip,
    }));
    add(&tasks, gpa, with(mk("opentofu", "infrastructure", &.{"infrastructure"}, "python", githubVersionCheckArgs(gpa, "tofu", &.{"--version"}, "opentofu/opentofu")), .{
        .requires = "tofu",
        .skip = infra_skip,
    }));
    add(&tasks, gpa, with(mk("packer", "infrastructure", &.{"infrastructure"}, "python", githubVersionCheckArgs(gpa, "packer", &.{"--version"}, "hashicorp/packer")), .{
        .requires = "packer",
        .skip = infra_skip,
    }));
    // ── Kubernetes tools ─────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("kubectl", "infrastructure", &.{"kubernetes"}, "python", pyCmd(gpa, kubectl_script)), .{
        .resource = "kubectl",
        .requires = "kubectl",
        .skip = k8s_skip,
    }));
    add(&tasks, gpa, with(mk("helm", "infrastructure", &.{"kubernetes"}, "helm", &.{ "repo", "update" }), .{
        .timeout_sec = 120,
        .resource = "helm",
        .skip = k8s_skip,
    }));
    // ── Security tools ───────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("gitleaks", "security", &.{"security"}, "python", githubVersionCheckArgs(gpa, "gitleaks", &.{"version"}, "gitleaks/gitleaks")), .{
        .requires = "gitleaks",
    }));
    add(&tasks, gpa, with(mk("trivy", "security", &.{"security"}, "trivy", &.{"update"}), .{
        .timeout_sec = 300,
        .resource = "trivy",
    }));
    // ── Static site ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("hugo", "dev-tools", &.{"static-site"}, "python", githubVersionCheckArgs(gpa, "hugo", &.{"version"}, "gohugoio/hugo")), .{
        .requires = "hugo",
        .skip = if (cli.skip_hugo) "disabled by --skip-hugo" else null,
    }));
    // ── Package managers (C++) ───────────────────────────────────────────────
    add(&tasks, gpa, with(mk("vcpkg", "package-manager", &.{"cpp"}, "python", vcpkgUpgradeArgs(gpa, config.vcpkg_skip_packages)), .{
        .resource = "vcpkg",
        .requires = "vcpkg",
        .skip = if (cli.skip_vcpkg) "disabled by --skip-vcpkg" else null,
    }));
    // ── AI tools ─────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("claude", "ai-tools", &.{ "ai", "claude" }, "python", pyCmd(gpa, claude_script)), .{
        .timeout_sec = 300,
        .requires = "claude",
    }));
    add(&tasks, gpa, with(mk("codex", "ai-tools", &.{ "ai", "codex" }, "python", pyCmd(gpa, codex_script)), .{
        .timeout_sec = 300,
        .requires = "codex",
    }));
    add(&tasks, gpa, with(mk("ollama-models", "ai", &.{"ai"}, "python", ollamaModelsUpgradeArgs(gpa, cli.ollama_timeout_sec)), .{
        .timeout_sec = 7200,
        .resource = "ollama",
        .requires = "ollama",
        .skip = if (!cli.update_ollama_models) "use --update-ollama-models to refresh local models" else null,
    }));
    // ── Docker ───────────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("docker-prune", "maintenance", &.{ "docker", "maintenance" }, "docker", &.{ "system", "prune", "-f", "--volumes" }), .{
        .timeout_sec = 300,
        .skip = if (!cli.deep_clean) "opt-in: requires --deep-clean and docker running" else null,
    }));
    // ── Maintenance ──────────────────────────────────────────────────────────
    add(&tasks, gpa, with(mk("cleanup", "maintenance", &.{"maintenance"}, "python", cleanupArgs(gpa, config.temp_cleanup_days, cli.deep_clean, cli.skip_destructive)), .{
        .timeout_sec = 3600,
        .skip = if (cli.skip_cleanup) "disabled by --skip-cleanup" else null,
    }));
    add(&tasks, gpa, mk("self-update", "maintenance", &.{"self"}, "python", pyCmd(gpa, self_update_script)));

    // ── GitHub-release tools (manually-installed, config-driven, opt-in) ──────
    for (config.github_tools) |tool| {
        const name = tool.id orelse blk: {
            var it = std.mem.splitBackwardsScalar(u8, tool.repo, '/');
            break :blk it.first();
        };
        const id = concat(gpa, &.{ "gh-", name });
        add(&tasks, gpa, with(mk(id, "github-tools", &.{ "github", "tools" }, "pwsh", githubReleaseArgs(gpa, tool)), .{
            .resource = "github-tools",
            .timeout_sec = 900,
            .skip = if (!cli.update_github_tools) "opt-in: use --update-github-tools" else null,
        }));
    }
    add(&tasks, gpa, with(mk("gh-notify-releases", "github-tools", &.{ "github", "tools", "report" }, "pwsh", githubNotifyArgs(gpa, config.github_tools, config.github_notification_ignore, config.github_notification_packages, cli.notify_apply)), .{
        .requires = "gh",
        .timeout_sec = 180,
        .skip = if (!cli.update_github_tools) "opt-in: use --update-github-tools" else null,
    }));

    return tasks.toOwnedSlice(gpa) catch @panic("oom");
}

/// Probing PATH is the only part of task construction that needs Io, so it lives
/// here and `taskTable` stays pure and unit-testable.
fn buildTasks(gpa: Allocator, io: Io, config: Config, cli: Cli) []Task {
    const tasks = taskTable(gpa, config, cli);
    // Only mark missing if no explicit skip_reason is already set (e.g. disabled by flag)
    for (tasks) |*task| {
        if (task.skip_reason == null and !commandExists(gpa, io, task.requires)) {
            task.skip_reason = fmtAlloc(gpa, "missing command: {s}", .{task.requires});
        }
    }
    return tasks;
}

fn findTask(tasks: []const Task, id: []const u8) ?Task {
    for (tasks) |task| {
        if (std.mem.eql(u8, task.id, id)) return task;
    }
    return null;
}

fn dependsOn(task: Task, id: []const u8) bool {
    for (task.depends_on) |dep| {
        if (std.mem.eql(u8, dep, id)) return true;
    }
    return false;
}

test "gh-notify-releases reports without installing unless asked" {
    const gpa = std.testing.allocator;
    var arena = std.heap.ArenaAllocator.init(gpa);
    defer arena.deinit();
    const a = arena.allocator();

    const packages = [_]NotifyPackage{.{ .repo = "google-gemini/gemini-cli", .npm = "@google/gemini-cli" }};

    const off = githubNotifyArgs(a, &.{}, &.{}, &packages, false);
    const off_script = off[off.len - 1];
    try std.testing.expect(std.mem.indexOf(u8, off_script, "$apply = $False") != null);
    try std.testing.expect(std.mem.indexOf(u8, off_script, "$pkgMap['google-gemini/gemini-cli']") != null);

    const on = githubNotifyArgs(a, &.{}, &.{}, &packages, true);
    try std.testing.expect(std.mem.indexOf(u8, on[on.len - 1], "$apply = $True") != null);

    // The substring guesswork that upgraded packages by repo-name match is gone.
    try std.testing.expect(std.mem.indexOf(u8, off_script, "$wl.Contains($name)") == null);
    try std.testing.expect(std.mem.indexOf(u8, off_script, "$npm.Contains(") == null);
}

test "uv-tools waits for the uv self-update" {
    const gpa = std.testing.allocator;
    var arena = std.heap.ArenaAllocator.init(gpa);
    defer arena.deinit();

    const tasks = taskTable(arena.allocator(), .{}, .{});
    const uv_tools = findTask(tasks, "uv-tools") orelse return error.TaskMissing;
    // Without this the two race: winget replaces the uv shim while uv-tools spawns it.
    try std.testing.expect(dependsOn(uv_tools, "uv"));
    try std.testing.expect(findTask(tasks, "uv") != null);
}

// ─── Task filtering ──────────────────────────────────────────────────────────

// Matches PS1 FastModeSkip
const fast_skip_ids = [_][]const u8{
    "chocolatey",       "wsl-distros",       "npm",                "pnpm",            "yarn",
    "bun",              "deno",              "rustup",             "cargo",           "go",
    "pip",              "pip-health",        "pipx",               "uv",              "uv-tools",
    "poetry",           "composer",          "ruby-gems",          "flutter",         "juliaup",
    "oh-my-posh",       "yt-dlp",            "volta",              "fnm",             "dotnet-tools",
    "dotnet-workloads", "vscode-extensions", "powershell-modules", "powershell-help", "uv-python",
    "ollama-models",    "vcpkg",             "conda",              "gcloud",          "az",
    "aws",              "terraform",         "pulumi",             "kubectl",         "helm",
    "hugo",             "opentofu",          "starship",           "zoxide",          "gitleaks",
    "trivy",            "packer",            "nvm",                "devcontainer",    "cross-manager",
    "mise-upgrade",     "tldr",
};

// Matches PS1 UltraFastSkip
const ultra_skip_ids = [_][]const u8{
    "windows-update", "store-apps", "wsl",           "wsl-distros",      "defender",
    "cleanup",        "winget",     "winget-source", "winget-userscope", "scoop",
};

const minimal_profile_skip = [_][]const u8{
    "vcpkg", "conda",    "gcloud",   "az",       "aws",   "terraform", "pulumi", "kubectl",      "helm",
    "hugo",  "opentofu", "starship", "gitleaks", "trivy", "packer",    "nvm",    "devcontainer",
};

const gaming_profile_skip = [_][]const u8{
    "vcpkg", "conda",    "gcloud",   "az",    "aws",    "terraform", "pulumi",       "kubectl",            "helm",
    "hugo",  "opentofu", "gitleaks", "trivy", "packer", "nvm",       "devcontainer", "powershell-modules",
};

fn matchesAny(gpa: Allocator, task: Task, values: *const StringSet) bool {
    if (values.contains(normalize(gpa, task.id))) return true;
    if (values.contains(normalize(gpa, task.category))) return true;
    for (task.tags) |tag| {
        if (values.contains(normalize(gpa, tag))) return true;
    }
    return false;
}

fn filterTasks(gpa: Allocator, tasks: []Task, cli: Cli, config: Config, prev_succeeded: *const StringSet) []Task {
    var only = StringSet.fromSlices(gpa, cli.only);
    var skip = StringSet.fromSlices(gpa, cli.skip);
    var config_skip = StringSet.fromSlices(gpa, config.skip_managers);
    var fast_skip = StringSet.fromSlices(gpa, &fast_skip_ids);
    var ultra_skip = StringSet.fromSlices(gpa, &ultra_skip_ids);

    var profile_skip: StringSet = .init(gpa);
    if (cli.profile) |profile| {
        if (std.mem.eql(u8, profile, "minimal")) {
            profile_skip = StringSet.fromSlices(gpa, &minimal_profile_skip);
        } else if (std.mem.eql(u8, profile, "gaming")) {
            profile_skip = StringSet.fromSlices(gpa, &gaming_profile_skip);
        }
    }

    var selected: std.ArrayList(Task) = .empty;
    for (tasks) |task| {
        if (!only.isEmpty() and !matchesAny(gpa, task, &only)) continue;
        var t = task;
        if (t.skip_reason == null) {
            const id = normalize(gpa, t.id);
            if (skip.contains(id) or config_skip.contains(id)) {
                t.skip_reason = "filtered by skip";
            } else if (profile_skip.contains(id)) {
                t.skip_reason = fmtAlloc(gpa, "skipped by --profile {s}", .{cli.profile orelse ""});
            } else if (cli.fast and fast_skip.contains(id)) {
                t.skip_reason = "filtered by fast mode";
            } else if (cli.ultra_fast and (fast_skip.contains(id) or ultra_skip.contains(id))) {
                t.skip_reason = "filtered by ultra-fast mode";
            } else if (!prev_succeeded.isEmpty() and prev_succeeded.contains(id)) {
                t.skip_reason = fmtAlloc(gpa, "succeeded within last {d:.0}h (--since-hours)", .{cli.since_hours});
            }
        }
        selected.append(gpa, t) catch @panic("oom");
    }
    return selected.toOwnedSlice(gpa) catch @panic("oom");
}

// ─── Command probing ─────────────────────────────────────────────────────────

fn pathExists(io: Io, path: []const u8) bool {
    Io.Dir.cwd().access(io, path, .{}) catch return false;
    return true;
}

/// PATH/PATHEXT are split once and lookups memoized: the task table probes ~85
/// commands and re-resolves them at spawn time, and every miss costs a full
/// dirs × extensions sweep of stat calls.
var path_dirs: [][]const u8 = &.{};
var path_exts: [][]const u8 = &.{};
var resolve_cache: std.StringHashMap(?[]const u8) = undefined;
var resolve_mu: Io.Mutex = .init;

/// Filename (lowercased) → first PATH directory index that holds it, plus the
/// full path. Listing each PATH directory once is far cheaper than stat-ing
/// dirs × extensions candidates for every command the task table probes.
const PathHit = struct { dir_index: usize, path: []const u8 };
var path_index: std.StringHashMap(PathHit) = undefined;

fn buildPathIndex(gpa: Allocator, io: Io) void {
    path_index = .init(gpa);
    for (path_dirs, 0..) |dir_path, dir_index| {
        var dir = Io.Dir.cwd().openDir(io, dir_path, .{ .iterate = true }) catch continue;
        defer dir.close(io);
        var it = dir.iterate();
        while (it.next(io) catch null) |entry| {
            if (entry.kind == .directory) continue;
            const key = normalize(gpa, entry.name);
            const existing = path_index.get(key);
            if (existing != null and existing.?.dir_index <= dir_index) continue;
            const full = std.fs.path.join(gpa, &.{ dir_path, entry.name }) catch continue;
            path_index.put(key, .{ .dir_index = dir_index, .path = full }) catch @panic("oom");
        }
    }
}

fn initPathCache(gpa: Allocator) void {
    resolve_cache = .init(gpa);

    var dirs: std.ArrayList([]const u8) = .empty;
    if (envAlloc(gpa, "PATH")) |path_var| {
        var it = std.mem.splitScalar(u8, path_var, if (is_windows) ';' else ':');
        while (it.next()) |dir| {
            if (dir.len != 0) dirs.append(gpa, dir) catch @panic("oom");
        }
    }
    path_dirs = dirs.toOwnedSlice(gpa) catch @panic("oom");

    var exts: std.ArrayList([]const u8) = .empty;
    if (is_windows) {
        if (envAlloc(gpa, "PATHEXT")) |value| {
            var it = std.mem.splitScalar(u8, value, ';');
            while (it.next()) |ext| {
                if (ext.len != 0) exts.append(gpa, ext) catch @panic("oom");
            }
        }
        if (exts.items.len == 0) {
            exts.appendSlice(gpa, &.{ ".exe", ".cmd", ".bat", ".ps1" }) catch @panic("oom");
        }
    } else {
        exts.append(gpa, "") catch @panic("oom");
    }
    path_exts = exts.toOwnedSlice(gpa) catch @panic("oom");
}

/// Mirrors the Rust probe order: the earliest PATH directory wins, and within a
/// directory the PATHEXT order decides.
fn resolveUncached(gpa: Allocator, name: []const u8) ?[]const u8 {
    var best: ?PathHit = null;
    for (path_exts) |ext| {
        const key = normalize(gpa, concat(gpa, &.{ name, ext }));
        const hit = path_index.get(key) orelse continue;
        if (best == null or hit.dir_index < best.?.dir_index) best = hit;
    }
    return if (best) |hit| hit.path else null;
}

fn resolveCommandPath(gpa: Allocator, io: Io, name: []const u8) ?[]const u8 {
    if (std.mem.indexOfAny(u8, name, "/\\") != null) {
        return if (pathExists(io, name)) name else null;
    }

    resolve_mu.lockUncancelable(io);
    if (resolve_cache.get(name)) |hit| {
        resolve_mu.unlock(io);
        return hit;
    }
    resolve_mu.unlock(io);

    const resolved = resolveUncached(gpa, name);

    resolve_mu.lockUncancelable(io);
    defer resolve_mu.unlock(io);
    resolve_cache.put(gpa.dupe(u8, name) catch @panic("oom"), resolved) catch @panic("oom");
    return resolved;
}

fn commandExists(gpa: Allocator, io: Io, name: []const u8) bool {
    return resolveCommandPath(gpa, io, name) != null;
}

/// Bare `spawn("yarn")` fails when the target is yarn.cmd / composer.ps1 /
/// gem.bat: CreateProcess does no PATHEXT search, so the spawn errors and the
/// task was silently marked Skipped. Resolve the full path first; .ps1 needs an
/// explicit PowerShell host.
fn taskArgv(gpa: Allocator, io: Io, command: []const u8, args: []const []const u8) []const []const u8 {
    var argv: std.ArrayList([]const u8) = .empty;
    const resolved = resolveCommandPath(gpa, io, command);
    if (resolved) |path| {
        const ext = std.fs.path.extension(path);
        if (std.ascii.eqlIgnoreCase(ext, ".ps1")) {
            const shell: []const u8 = if (resolveCommandPath(gpa, io, "pwsh") != null) "pwsh" else "powershell";
            argv.appendSlice(gpa, &.{ shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", path }) catch @panic("oom");
        } else {
            argv.append(gpa, path) catch @panic("oom");
        }
    } else {
        argv.append(gpa, command) catch @panic("oom");
    }
    argv.appendSlice(gpa, args) catch @panic("oom");
    return argv.toOwnedSlice(gpa) catch @panic("oom");
}

// ─── Task execution ──────────────────────────────────────────────────────────

const LineSink = struct {
    gpa: Allocator,
    io: Io,
    mu: Io.Mutex = .init,
    lines: std.ArrayList([]const u8) = .empty,
    quiet: bool,
    prefix: bool,
    id: []const u8,
    is_err: bool,
    file: Io.File,
    done: std.atomic.Value(bool) = .init(false),
};

fn readerThread(sink: *LineSink) void {
    var buf: [64 * 1024]u8 = undefined;
    var reader = sink.file.reader(sink.io, &buf);
    const stream = &reader.interface;
    while (stream.takeDelimiter('\n') catch null) |raw| {
        const line = std.mem.trimEnd(u8, raw, "\r\n");
        const owned = sink.gpa.dupe(u8, line) catch @panic("oom");
        if (!sink.quiet) {
            const tone = if (sink.is_err) col(A.yellow) else col(A.grey);
            if (sink.prefix) {
                if (sink.is_err) pe(sink.io, "{s}[{s}]{s} {s}{s}{s}\n", .{ col(A.blue), sink.id, col(A.reset), tone, owned, col(A.reset) }) else p(sink.io, "{s}[{s}]{s} {s}{s}{s}\n", .{ col(A.blue), sink.id, col(A.reset), tone, owned, col(A.reset) });
            } else {
                if (sink.is_err) pe(sink.io, "  {s}{s}{s}\n", .{ tone, owned, col(A.reset) }) else p(sink.io, "  {s}{s}{s}\n", .{ tone, owned, col(A.reset) });
            }
        }
        sink.mu.lockUncancelable(sink.io);
        sink.lines.append(sink.gpa, owned) catch @panic("oom");
        sink.mu.unlock(sink.io);
    }
    sink.done.store(true, .release);
}

const Watchdog = struct {
    io: Io,
    handle: windows.HANDLE,
    deadline_ms: i64,
    done: std.atomic.Value(bool) = .init(false),
    timed_out: std.atomic.Value(bool) = .init(false),
};

fn watchdogThread(w: *Watchdog) void {
    while (!w.done.load(.acquire)) {
        if (nowMs(w.io) >= w.deadline_ms) {
            _ = TerminateProcess(w.handle, 1);
            w.timed_out.store(true, .release);
            return;
        }
        sleepMs(w.io, 50);
    }
}

fn makeSummary(
    gpa: Allocator,
    task: Task,
    status: []const u8,
    duration_ms: u64,
    exit_code: ?i32,
    output_tail: [][]const u8,
) TaskSummary {
    _ = gpa;
    return .{
        .id = task.id,
        .category = task.category,
        .status = status,
        .duration_ms = duration_ms,
        .exit_code = exit_code,
        .command = task.command,
        .args = task.args,
        .output_tail = output_tail,
    };
}

const Outcome = struct {
    timed_out: bool,
    ok_on_timeout: bool,
    code: ?i32,
    acceptable_exit_codes: []const i32,
    lines: []const []const u8,
};

fn classifyStatus(o: Outcome) []const u8 {
    var status: []const u8 = undefined;
    if (o.timed_out and o.ok_on_timeout) {
        status = "Succeeded";
    } else if (o.timed_out) {
        status = "TimedOut";
    } else if (o.code != null and containsCode(o.acceptable_exit_codes, o.code.?)) {
        status = "Succeeded";
    } else if (o.code != null and o.code.? == 0) {
        status = "Succeeded";
    } else {
        status = "Failed";
    }
    if (!std.mem.eql(u8, status, "Succeeded")) return status;
    for (o.lines) |line| {
        const trimmed = std.mem.trimStart(u8, line, " \t");
        if (std.mem.startsWith(u8, trimmed, skip_prefix)) return "Skipped";
        // winget helper scripts keep running after an individual package needs
        // intervention, then exit 0. Surface that in the task status.
        if ((std.mem.startsWith(u8, trimmed, "failed (") or
            std.mem.startsWith(u8, trimmed, "needs manual action (")) and
            std.mem.indexOf(u8, trimmed, "(0)") == null)
        {
            return "Failed";
        }
    }
    return status;
}

test "classifyStatus" {
    const ok: Outcome = .{
        .timed_out = false,
        .ok_on_timeout = false,
        .code = 0,
        .acceptable_exit_codes = &.{},
        .lines = &.{},
    };
    const expect = std.testing.expectEqualStrings;

    try expect("Succeeded", classifyStatus(ok));

    var o = ok;
    o.code = 1;
    try expect("Failed", classifyStatus(o));

    o = ok;
    o.code = -1978335189;
    o.acceptable_exit_codes = &.{-1978335189};
    try expect("Succeeded", classifyStatus(o));

    o = ok;
    o.timed_out = true;
    try expect("TimedOut", classifyStatus(o));
    o.ok_on_timeout = true;
    try expect("Succeeded", classifyStatus(o));

    o = ok;
    o.lines = &.{"  needs manual action (1): TheStellaTeam.Stella"};
    try expect("Failed", classifyStatus(o));

    o.lines = &.{"needs manual action (0)"};
    try expect("Succeeded", classifyStatus(o));

    o.lines = &.{"failed (2): something"};
    try expect("Failed", classifyStatus(o));

    o.lines = &.{skip_prefix ++ " nothing to do"};
    try expect("Skipped", classifyStatus(o));

    // A non-zero exit is not rescued by a clean-looking log line.
    o = ok;
    o.code = 1;
    o.lines = &.{skip_prefix ++ " nothing to do"};
    try expect("Failed", classifyStatus(o));
}

fn runTaskStreaming(
    gpa: Allocator,
    io: Io,
    task: Task,
    quiet: bool,
    default_timeout_ms: u64,
    prefix_output: bool,
    resource_locks: *std.StringHashMap(*Io.Mutex),
) TaskSummary {
    if (!quiet) {
        p(io, "{s}{s} run {s} {s}{s: <22}{s} {s}{s} {s}{s}\n", .{ col(A.cyan), G.run, col(A.reset), col(A.bold), task.id, col(A.reset), col(A.grey), task.command, shellJoinBrief(gpa, task.args), col(A.reset) });
    }

    var held: ?*Io.Mutex = null;
    if (task.resource.len != 0) {
        if (resource_locks.get(task.resource)) |mutex| {
            mutex.lockUncancelable(io);
            held = mutex;
        }
    }
    defer if (held) |mutex| mutex.unlock(io);

    const timeout_ms = task.timeout_ms orelse default_timeout_ms;
    const start = nowMs(io);
    const argv = taskArgv(gpa, io, task.command, task.args);

    var child = std.process.spawn(io, .{
        .argv = argv,
        .stdout = .pipe,
        .stderr = .pipe,
    }) catch |err| {
        pe(io, "{s}{s} fail{s} {s: <22} spawn failed: {t}\n", .{ col(A.red), G.fail, col(A.reset), task.id, err });
        const lines = gpa.alloc([]const u8, 1) catch @panic("oom");
        lines[0] = fmtAlloc(gpa, "spawn failed: {t}", .{err});
        return makeSummary(gpa, task, "Failed", @intCast(nowMs(io) - start), null, lines);
    };

    // Heap, not stack: a reader blocked on a pipe a grandchild still holds gets
    // detached below and must not outlive its sink.
    const out_sink = gpa.create(LineSink) catch @panic("oom");
    out_sink.* = .{
        .gpa = gpa,
        .io = io,
        .quiet = quiet,
        .prefix = prefix_output,
        .id = task.id,
        .is_err = false,
        .file = child.stdout.?,
    };
    const err_sink = gpa.create(LineSink) catch @panic("oom");
    err_sink.* = .{
        .gpa = gpa,
        .io = io,
        .quiet = quiet,
        .prefix = prefix_output,
        .id = task.id,
        .is_err = true,
        .file = child.stderr.?,
    };

    var watchdog: Watchdog = .{
        .io = io,
        .handle = child.id.?,
        .deadline_ms = start + @as(i64, @intCast(timeout_ms)),
    };

    const out_thread = std.Thread.spawn(.{}, readerThread, .{out_sink}) catch @panic("thread spawn failed");
    const err_thread = std.Thread.spawn(.{}, readerThread, .{err_sink}) catch @panic("thread spawn failed");
    const dog_thread = std.Thread.spawn(.{}, watchdogThread, .{&watchdog}) catch @panic("thread spawn failed");

    // Wait on the process, never on pipe EOF: a grandchild that inherited the
    // pipe (VS Code helpers, service hosts) holds it open long after the child
    // exits, and joining the readers first hangs the task until its timeout.
    var raw_code: u32 = still_active;
    while (true) {
        if (GetExitCodeProcess(child.id.?, &raw_code) == .FALSE) break;
        if (raw_code != still_active) break;
        sleepMs(io, 5);
    }

    watchdog.done.store(true, .release);
    dog_thread.join();
    const timed_out = watchdog.timed_out.load(.acquire);

    const drain_deadline = nowMs(io) + reader_drain_ms;
    while (nowMs(io) < drain_deadline and
        !(out_sink.done.load(.acquire) and err_sink.done.load(.acquire)))
    {
        sleepMs(io, 10);
    }

    if (out_sink.done.load(.acquire) and err_sink.done.load(.acquire)) {
        out_thread.join();
        err_thread.join();
        _ = child.wait(io) catch {};
    } else {
        // Sinks and pipe handles leak on purpose: closing a pipe under a
        // blocking read is not safe, and the run is short-lived.
        out_thread.detach();
        err_thread.detach();
    }
    const duration_ms: u64 = @intCast(nowMs(io) - start);

    var all_lines: std.ArrayList([]const u8) = .empty;
    out_sink.mu.lockUncancelable(io);
    all_lines.appendSlice(gpa, out_sink.lines.items) catch @panic("oom");
    out_sink.mu.unlock(io);
    err_sink.mu.lockUncancelable(io);
    all_lines.appendSlice(gpa, err_sink.lines.items) catch @panic("oom");
    err_sink.mu.unlock(io);

    const code: ?i32 = if (raw_code == still_active) null else @as(i32, @bitCast(raw_code));

    const status = classifyStatus(.{
        .timed_out = timed_out,
        .ok_on_timeout = task.ok_on_timeout,
        .code = code,
        .acceptable_exit_codes = task.acceptable_exit_codes,
        .lines = all_lines.items,
    });

    if (!quiet) {
        p(io, "{s}{s} done{s} {s: <22} {s}{s}{s} {s}({d:.1}s){s}\n", .{ statusColor(status), statusGlyph(status), col(A.reset), task.id, statusColor(status), status, col(A.reset), col(A.grey), @as(f64, @floatFromInt(duration_ms)) / 1000.0, col(A.reset) });
    }

    // Normalize whitelisted exit codes to 0 in the summary so the report column
    // doesn't show an alarming raw code (e.g. winget -1978335189 "no applicable
    // upgrade") for a task that succeeded.
    const reported_code: ?i32 = if (code != null and containsCode(task.acceptable_exit_codes, code.?)) 0 else code;
    return makeSummary(gpa, task, status, duration_ms, reported_code, capOutput(gpa, all_lines.items, 300));
}

fn containsCode(codes: []const i32, code: i32) bool {
    for (codes) |c| {
        if (c == code) return true;
    }
    return false;
}

/// A dependency that was filtered out of this run never blocks; only a scheduled
/// one that has not finished yet does.
fn depsMet(runnable: *const std.StringHashMap(void), finished: *const std.StringHashMap(void), task: Task) bool {
    for (task.depends_on) |dep| {
        if (!runnable.contains(dep)) continue;
        if (!finished.contains(dep)) return false;
    }
    return true;
}

test "depsMet blocks until dependency finished" {
    const gpa = std.testing.allocator;
    var runnable: std.StringHashMap(void) = .init(gpa);
    defer runnable.deinit();
    var finished: std.StringHashMap(void) = .init(gpa);
    defer finished.deinit();

    const uv_tools: Task = .{
        .id = "uv-tools",
        .category = "python",
        .tags = &.{},
        .command = "uv",
        .args = &.{},
        .requires = "uv",
        .depends_on = &.{"uv"},
    };

    // `uv` is scheduled but still running: uv-tools must wait, otherwise it spawns
    // while winget is mid-reinstall of the uv shim and dies with FileNotFound.
    try runnable.put("uv", {});
    try std.testing.expect(!depsMet(&runnable, &finished, uv_tools));

    try finished.put("uv", {});
    try std.testing.expect(depsMet(&runnable, &finished, uv_tools));

    // `uv` filtered out of this run (--skip uv): nothing to wait for.
    var empty: std.StringHashMap(void) = .init(gpa);
    defer empty.deinit();
    try std.testing.expect(depsMet(&empty, &empty, uv_tools));
}

/// Returns the pattern that marks this failure as permanent, or null to retry.
fn matchNoRetry(patterns: []const []const u8, lines: []const []const u8) ?[]const u8 {
    for (patterns) |pattern| {
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, pattern) != null) return pattern;
        }
    }
    return null;
}

fn sameOutput(a: []const []const u8, b: []const []const u8) bool {
    if (a.len == 0 or a.len != b.len) return false;
    for (a, b) |x, y| {
        if (!std.mem.eql(u8, x, y)) return false;
    }
    return true;
}

test "matchNoRetry" {
    const patterns = [_][]const u8{"has indeterminate status"};
    const lines = [_][]const u8{ "boot", "RuntimeError: SCM service QingpingLogger has indeterminate status: paused" };
    try std.testing.expectEqualStrings("has indeterminate status", matchNoRetry(&patterns, &lines).?);
    try std.testing.expect(matchNoRetry(&patterns, &.{"unrelated"}) == null);
    try std.testing.expect(matchNoRetry(&.{}, &lines) == null);
}

test "sameOutput" {
    const a = [_][]const u8{ "one", "two" };
    const b = [_][]const u8{ "one", "two" };
    const c = [_][]const u8{ "one", "three" };
    try std.testing.expect(sameOutput(&a, &b));
    try std.testing.expect(!sameOutput(&a, &c));
    // No output is not evidence of a repeated failure.
    try std.testing.expect(!sameOutput(&.{}, &.{}));
}

const Runner = struct {
    gpa: Allocator,
    io: Io,
    tasks: []const Task,
    next: usize = 0,
    queue_mu: Io.Mutex = .init,
    results: std.ArrayList(TaskSummary) = .empty,
    results_mu: Io.Mutex = .init,
    finished: std.StringHashMap(void),
    runnable: std.StringHashMap(void),
    resource_locks: std.StringHashMap(*Io.Mutex),
    resource_busy: std.StringHashMap(void),
    quiet: bool,
    timeout_ms: u64,
    retry_count: u32,
    prefix_output: bool,
    taken: std.ArrayList(bool) = .empty,

    /// Take the first queued task whose scheduled dependencies have all finished.
    /// Returns null when the queue is drained.
    fn take(self: *Runner) ?Task {
        while (true) {
            self.queue_mu.lockUncancelable(self.io);
            if (self.next >= self.tasks.len) {
                self.queue_mu.unlock(self.io);
                return null;
            }
            var index: ?usize = null;
            var i = self.next;
            while (i < self.tasks.len) : (i += 1) {
                if (self.taken.items[i]) continue;
                const t = self.tasks[i];
                if (t.resource.len != 0 and self.resource_busy.contains(t.resource)) continue;
                if (self.dependenciesMet(t)) {
                    index = i;
                    break;
                }
            }
            if (index) |idx| {
                self.taken.items[idx] = true;
                if (self.tasks[idx].resource.len != 0) self.resource_busy.put(self.tasks[idx].resource, {}) catch @panic("oom");
                while (self.next < self.tasks.len and self.taken.items[self.next]) self.next += 1;
                self.queue_mu.unlock(self.io);
                return self.tasks[idx];
            }
            // Everything left is blocked on an in-flight dependency; wait and re-check.
            var pending = false;
            var j = self.next;
            while (j < self.tasks.len) : (j += 1) {
                if (!self.taken.items[j]) pending = true;
            }
            self.queue_mu.unlock(self.io);
            if (!pending) return null;
            sleepMs(self.io, 200);
        }
    }

    fn dependenciesMet(self: *Runner, task: Task) bool {
        return depsMet(&self.runnable, &self.finished, task);
    }

    fn runOne(self: *Runner, task: Task) void {
        var result = runTaskStreaming(self.gpa, self.io, task, self.quiet, self.timeout_ms, self.prefix_output, &self.resource_locks);
        var attempt: u32 = 1;
        while (attempt <= self.retry_count) : (attempt += 1) {
            if (!std.mem.eql(u8, result.status, "Failed") and !std.mem.eql(u8, result.status, "TimedOut")) break;
            if (matchNoRetry(task.no_retry_patterns, result.output_tail)) |hit| {
                pe(self.io, "{s}{s} noretry{s} {s} {s}(permanent: {s}){s}\n", .{ col(A.yellow), G.skip, col(A.reset), task.id, col(A.grey), hit, col(A.reset) });
                break;
            }
            const previous = result.output_tail;
            const shift: u6 = @intCast(@min(attempt, 4));
            const delay_sec: u64 = 3 * (@as(u64, 1) << shift);
            pe(self.io, "{s}{s} retry {d}/{d}{s} {s} {s}(waiting {d}s){s}\n", .{ col(A.yellow), G.retry, attempt, self.retry_count, col(A.reset), task.id, col(A.grey), delay_sec, col(A.reset) });
            sleepMs(self.io, delay_sec * 1000);
            result = runTaskStreaming(self.gpa, self.io, task, self.quiet, self.timeout_ms, self.prefix_output, &self.resource_locks);
            // A byte-identical failure means the cause is not transient; a third
            // identical run only adds noise and backoff.
            if ((std.mem.eql(u8, result.status, "Failed") or std.mem.eql(u8, result.status, "TimedOut")) and
                sameOutput(previous, result.output_tail))
            {
                pe(self.io, "{s}{s} noretry{s} {s} {s}(identical failure output){s}\n", .{ col(A.yellow), G.skip, col(A.reset), task.id, col(A.grey), col(A.reset) });
                break;
            }
        }
        self.results_mu.lockUncancelable(self.io);
        self.results.append(self.gpa, result) catch @panic("oom");
        self.results_mu.unlock(self.io);
        self.queue_mu.lockUncancelable(self.io);
        if (task.resource.len != 0) _ = self.resource_busy.remove(task.resource);
        self.finished.put(task.id, {}) catch @panic("oom");
        self.queue_mu.unlock(self.io);
    }

    fn workerLoop(self: *Runner) void {
        while (self.take()) |task| self.runOne(task);
    }
};

fn runTasks(gpa: Allocator, io: Io, tasks: []const Task, cli: Cli, jobs: usize, timeout_ms: u64) []TaskSummary {
    var results: std.ArrayList(TaskSummary) = .empty;

    if (cli.dry_run) {
        for (tasks) |task| {
            if (task.skip_reason) |reason| {
                if (!cli.quiet) p(io, "{s}{s} skip{s} {s: <22} {s}{s}{s}\n", .{ col(A.grey), G.skip, col(A.reset), task.id, col(A.grey), reason, col(A.reset) });
                results.append(gpa, makeSummary(gpa, task, "Skipped", 0, null, &.{})) catch @panic("oom");
            } else {
                p(io, "{s}{s} dry {s} {s: <22} {s}{s} {s}{s}\n", .{ col(A.cyan), G.dry, col(A.reset), task.id, col(A.grey), task.command, shellJoinBrief(gpa, task.args), col(A.reset) });
                results.append(gpa, makeSummary(gpa, task, "DryRun", 0, null, &.{})) catch @panic("oom");
            }
        }
        return results.toOwnedSlice(gpa) catch @panic("oom");
    }

    var to_run: std.ArrayList(Task) = .empty;
    for (tasks) |task| {
        if (task.skip_reason) |reason| {
            if (!cli.quiet) p(io, "{s}{s} skip{s} {s: <22} {s}{s}{s}\n", .{ col(A.grey), G.skip, col(A.reset), task.id, col(A.grey), reason, col(A.reset) });
            results.append(gpa, makeSummary(gpa, task, "Skipped", 0, null, &.{})) catch @panic("oom");
        } else {
            to_run.append(gpa, task) catch @panic("oom");
        }
    }

    var runner: Runner = .{
        .gpa = gpa,
        .io = io,
        .tasks = to_run.items,
        .finished = .init(gpa),
        .runnable = .init(gpa),
        .resource_locks = .init(gpa),
        .resource_busy = .init(gpa),
        .quiet = cli.quiet,
        .timeout_ms = timeout_ms,
        .retry_count = cli.retry_count,
        .prefix_output = jobs > 1,
    };
    for (to_run.items) |task| {
        runner.taken.append(gpa, false) catch @panic("oom");
        runner.runnable.put(task.id, {}) catch @panic("oom");
        if (task.resource.len != 0 and !runner.resource_locks.contains(task.resource)) {
            const mutex = gpa.create(Io.Mutex) catch @panic("oom");
            mutex.* = .init;
            runner.resource_locks.put(task.resource, mutex) catch @panic("oom");
        }
    }

    if (jobs <= 1) {
        runner.workerLoop();
    } else {
        var threads: std.ArrayList(std.Thread) = .empty;
        var i: usize = 0;
        while (i < jobs) : (i += 1) {
            const thread = std.Thread.spawn(.{}, Runner.workerLoop, .{&runner}) catch break;
            threads.append(gpa, thread) catch @panic("oom");
        }
        for (threads.items) |thread| thread.join();
    }

    results.appendSlice(gpa, runner.results.items) catch @panic("oom");
    return results.toOwnedSlice(gpa) catch @panic("oom");
}

// ─── Output ──────────────────────────────────────────────────────────────────

fn printTaskList(gpa: Allocator, io: Io, tasks: []const Task) void {
    _ = gpa;
    p(io, "{s: <26} {s: <20} State\n", .{ "Task", "Category" });
    p(io, "{s: <26} {s: <20} -----\n", .{ "----", "--------" });
    for (tasks) |task| {
        p(io, "{s: <26} {s: <20} {s}\n", .{ task.id, task.category, task.skip_reason orelse "planned" });
    }
}

fn countStatus(results: []const TaskSummary, status: []const u8) usize {
    var count: usize = 0;
    for (results) |r| {
        if (std.mem.eql(u8, r.status, status)) count += 1;
    }
    return count;
}

fn printSummary(gpa: Allocator, io: Io, results: []const TaskSummary, total_duration_ms: u64) void {
    p(io, "\n", .{});
    p(io, "{s}{s}  {s: <26} {s: <12} {s: <8} Exit{s}\n", .{ col(A.bold), G.summary, "Task", "Status", "Time(s)", col(A.reset) });
    p(io, "{s}   {s: <26} {s: <12} {s: <8} ----{s}\n", .{ col(A.grey), "----", "------", "-------", col(A.reset) });
    // Skipped rows are almost all "tool not installed"; they drown out the tasks
    // that actually ran. Counted below, just not listed.
    for (results) |r| {
        if (std.mem.eql(u8, r.status, "Skipped")) continue;
        const exit = if (r.exit_code) |code| fmtAlloc(gpa, "{d}", .{code}) else "-";
        const secs = fmtAlloc(gpa, "{d:.1}", .{@as(f64, @floatFromInt(r.duration_ms)) / 1000.0});
        p(io, "{s}{s}{s}  {s: <26} {s}{s: <12}{s} {s: <8} {s}\n", .{ statusColor(r.status), statusGlyph(r.status), col(A.reset), r.id, statusColor(r.status), r.status, col(A.reset), secs, exit });
    }
    p(io, "\n", .{});

    const failed = countStatus(results, "Failed");
    const timed_out = countStatus(results, "TimedOut");
    const tone = if (failed + timed_out > 0) col(A.red) else col(A.green);
    p(io, "{s}{s} done{s}  total={d}  {s}succeeded={d}{s}  {s}failed={d}{s}  {s}timed-out={d}{s}  {s}skipped={d}  dry={d}{s}  duration={d:.1}s\n", .{
        tone,
        G.gear,
        col(A.reset),
        results.len,
        col(A.green),
        countStatus(results, "Succeeded"),
        col(A.reset),
        if (failed > 0) col(A.red) else col(A.grey),
        failed,
        col(A.reset),
        if (timed_out > 0) col(A.yellow) else col(A.grey),
        timed_out,
        col(A.reset),
        col(A.grey),
        countStatus(results, "Skipped"),
        countStatus(results, "DryRun"),
        col(A.reset),
        @as(f64, @floatFromInt(total_duration_ms)) / 1000.0,
    });
}

fn writeSummaryJson(
    gpa: Allocator,
    io: Io,
    path: []const u8,
    started_at: []const u8,
    duration_ms: u64,
    dry_run: bool,
    results: []const TaskSummary,
) !void {
    var buf: std.ArrayList(u8) = .empty;
    const put = struct {
        fn f(list: *std.ArrayList(u8), a: Allocator, text: []const u8) void {
            list.appendSlice(a, text) catch @panic("oom");
        }
    }.f;

    put(&buf, gpa, fmtAlloc(gpa, "{{\n  \"version\": \"{s}\",\n  \"started_at\": \"{s}\",\n  \"duration_ms\": {d},\n  \"dry_run\": {s},\n  \"results\": [\n", .{
        version, started_at, duration_ms, if (dry_run) "true" else "false",
    }));
    for (results, 0..) |r, i| {
        put(&buf, gpa, fmtAlloc(gpa, "    {{\n      \"id\": \"{s}\",\n      \"category\": \"{s}\",\n      \"status\": \"{s}\",\n      \"duration_ms\": {d},\n", .{
            jsonEscape(gpa, r.id), jsonEscape(gpa, r.category), jsonEscape(gpa, r.status), r.duration_ms,
        }));
        if (r.exit_code) |code| {
            put(&buf, gpa, fmtAlloc(gpa, "      \"exit_code\": {d},\n", .{code}));
        } else {
            put(&buf, gpa, "      \"exit_code\": null,\n");
        }
        put(&buf, gpa, fmtAlloc(gpa, "      \"command\": \"{s}\",\n      \"args\": [", .{jsonEscape(gpa, r.command)}));
        for (r.args, 0..) |arg, j| {
            if (j != 0) put(&buf, gpa, ", ");
            put(&buf, gpa, fmtAlloc(gpa, "\"{s}\"", .{jsonEscape(gpa, arg)}));
        }
        put(&buf, gpa, "],\n      \"output_tail\": [");
        for (r.output_tail, 0..) |line, j| {
            if (j != 0) put(&buf, gpa, ", ");
            put(&buf, gpa, fmtAlloc(gpa, "\"{s}\"", .{jsonEscape(gpa, line)}));
        }
        put(&buf, gpa, fmtAlloc(gpa, "]\n    }}{s}\n", .{if (i + 1 == results.len) "" else ","}));
    }
    put(&buf, gpa, "  ]\n}\n");

    if (std.fs.path.dirname(path)) |parent| {
        Io.Dir.cwd().createDirPath(io, parent) catch {};
    }
    try Io.Dir.cwd().writeFile(io, .{ .sub_path = path, .data = buf.items });
}

// ─── What's Changed summary ──────────────────────────────────────────────────

/// True if a string looks like a version: contains a digit and at least one dot.
fn looksLikeVersion(s: []const u8) bool {
    var has_digit = false;
    for (s) |c| {
        if (std.ascii.isDigit(c)) has_digit = true;
    }
    return has_digit and std.mem.indexOfScalar(u8, s, '.') != null;
}

fn hasDigit(s: []const u8) bool {
    for (s) |c| {
        if (std.ascii.isDigit(c)) return true;
    }
    return false;
}

/// Strip spinner/progress junk chars that winget inserts into lines.
fn stripProgress(gpa: Allocator, s: []const u8) []const u8 {
    var filtered: std.ArrayList(u8) = .empty;
    for (s) |c| {
        if (std.ascii.isPrint(c)) filtered.append(gpa, c) catch @panic("oom");
    }
    var tokens: std.ArrayList([]const u8) = .empty;
    var it = std.mem.tokenizeAny(u8, filtered.items, " \t");
    while (it.next()) |token| {
        if (std.mem.eql(u8, token, "-") or std.mem.eql(u8, token, "\\") or
            std.mem.eql(u8, token, "|") or std.mem.eql(u8, token, "/")) continue;
        tokens.append(gpa, token) catch @panic("oom");
    }
    return std.mem.join(gpa, " ", tokens.items) catch @panic("oom");
}

fn afterMarker(line: []const u8, marker: []const u8) ?[]const u8 {
    const idx = std.mem.indexOf(u8, line, marker) orelse return null;
    return line[idx + marker.len ..];
}

fn changesFor(gpa: Allocator, r: TaskSummary) [][]const u8 {
    const lines = r.output_tail;
    var changes: std.ArrayList([]const u8) = .empty;

    if (std.mem.eql(u8, r.id, "npm")) {
        for (lines) |line| {
            if (afterMarker(line, "upgraded ")) |pkg| {
                const t = std.mem.trim(u8, pkg, " \t");
                if (t.len != 0) changes.append(gpa, t) catch @panic("oom");
            }
        }
    } else if (std.mem.eql(u8, r.id, "pnpm")) {
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "Switching") == null) continue;
            const after_from = afterMarker(line, "from v") orelse continue;
            if (std.mem.indexOf(u8, line, "to v") == null) continue;
            const after_switching = afterMarker(line, "Switching").?;
            var name_it = std.mem.tokenizeAny(u8, after_switching, " \t");
            const tool = name_it.next() orelse "pnpm";
            const split = std.mem.indexOf(u8, after_from, " to v") orelse continue;
            const old = std.mem.trim(u8, after_from[0..split], " \t");
            const new_ver = std.mem.trim(u8, std.mem.trimEnd(u8, after_from[split + 5 ..], "."), " \t");
            if (!std.mem.eql(u8, old, new_ver)) {
                changes.append(gpa, fmtAlloc(gpa, "{s}: {s} → {s}", .{ tool, old, new_ver })) catch @panic("oom");
            }
        }
    } else if (std.mem.eql(u8, r.id, "winget") or std.mem.eql(u8, r.id, "winget-batch") or std.mem.eql(u8, r.id, "winget-userscope") or std.mem.eql(u8, r.id, "store-apps")) {
        // The per-package loop prints "upgraded (N): a, b" and
        // "upgraded via forced install (N): c, d" summary lines. Split them into
        // individual package ids so the report lists each upgraded package.
        for (lines) |line| {
            const clean = stripProgress(gpa, line);
            if (std.mem.startsWith(u8, clean, "upgraded via forced install (")) {
                if (afterMarker(clean, "): ")) |list| {
                    var it = std.mem.splitSequence(u8, list, ", ");
                    while (it.next()) |pkg| {
                        const t = std.mem.trim(u8, pkg, " \t");
                        if (t.len != 0) changes.append(gpa, fmtAlloc(gpa, "{s} (forced)", .{t})) catch @panic("oom");
                    }
                }
            } else if (std.mem.startsWith(u8, clean, "upgraded (")) {
                if (afterMarker(clean, "): ")) |list| {
                    var it = std.mem.splitSequence(u8, list, ", ");
                    while (it.next()) |pkg| {
                        const t = std.mem.trim(u8, pkg, " \t");
                        if (t.len != 0) changes.append(gpa, t) catch @panic("oom");
                    }
                }
            }
        }
        // Fallback: legacy table parsing for the pre-per-package winget-batch
        // output and "[N/M] Upgrading:" lines.
        var in_table = false;
        for (lines, 0..) |line, i| {
            const clean = stripProgress(gpa, line);
            if (!in_table and std.mem.indexOf(u8, clean, "Name") != null and
                std.mem.indexOf(u8, clean, "Id") != null and std.mem.indexOf(u8, clean, "Available") != null)
            {
                in_table = true;
                continue;
            }
            if (in_table) {
                const trimmed = std.mem.trim(u8, clean, " \t");
                var only_dashes = trimmed.len != 0;
                for (trimmed) |c| {
                    if (c != '-' and c != ' ') only_dashes = false;
                }
                if (only_dashes) continue;
                if (trimmed.len == 0 or
                    std.mem.indexOf(u8, clean, "package(s) are pinned") != null or
                    std.mem.indexOf(u8, clean, "The following") != null or
                    std.mem.indexOf(u8, clean, "upgrades available") != null or
                    std.mem.indexOf(u8, clean, "upgrade available") != null)
                {
                    in_table = false;
                } else {
                    var tokens: std.ArrayList([]const u8) = .empty;
                    var it = std.mem.splitSequence(u8, clean, "  ");
                    while (it.next()) |part| {
                        const t = std.mem.trim(u8, part, " \t");
                        if (t.len != 0) tokens.append(gpa, t) catch @panic("oom");
                    }
                    if (tokens.items.len >= 3) {
                        const name = tokens.items[0];
                        const available = tokens.items[tokens.items.len - 1];
                        const ver = tokens.items[tokens.items.len - 2];
                        if (looksLikeVersion(available) or std.mem.startsWith(u8, available, "<")) {
                            changes.append(gpa, fmtAlloc(gpa, "{s}: {s} → {s}", .{ name, ver, available })) catch @panic("oom");
                        }
                    }
                }
            }
            // Also capture "[N/M] Upgrading: X" lines from the winget task
            if (std.mem.indexOf(u8, line, "] Upgrading:") != null and std.mem.eql(u8, r.id, "winget")) {
                const rest = afterMarker(line, "] Upgrading:").?;
                var pkg = std.mem.trim(u8, rest, " \t");
                if (std.mem.indexOf(u8, pkg, " (installed")) |cut| pkg = std.mem.trim(u8, pkg[0..cut], " \t");
                const next: []const u8 = if (i + 1 < lines.len) lines[i + 1] else "";
                if (std.mem.indexOf(u8, next, "Not applicable:") == null and
                    std.mem.indexOf(u8, next, "FAILED:") == null and pkg.len != 0)
                {
                    var already = false;
                    for (changes.items) |c| {
                        if (std.mem.startsWith(u8, c, pkg)) already = true;
                    }
                    if (!already) changes.append(gpa, pkg) catch @panic("oom");
                }
            }
        }
    } else if (std.mem.eql(u8, r.id, "scoop")) {
        for (lines) |line| {
            // "Updating 'name' (old -> new)"
            if (std.mem.indexOf(u8, line, "Updating ") != null and std.mem.indexOf(u8, line, "->") != null) {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                if (changes.items.len >= 10) break;
            }
        }
    } else if (std.mem.eql(u8, r.id, "chocolatey")) {
        for (lines) |line| {
            // "Chocolatey upgraded N/M packages" — only if N > 0
            if (std.mem.indexOf(u8, line, "upgraded") != null and std.mem.indexOf(u8, line, "package") != null) {
                if (std.mem.indexOf(u8, line, "upgraded 0/") == null) {
                    changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                }
            } else if (std.mem.indexOf(u8, line, " to ")) |cut| {
                var head = std.mem.tokenizeAny(u8, line[0..cut], " \t");
                var last: []const u8 = "";
                while (head.next()) |token| last = token;
                if (looksLikeVersion(last)) {
                    changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                    if (changes.items.len >= 10) break;
                }
            }
        }
    } else if (std.mem.eql(u8, r.id, "cargo")) {
        for (lines) |line| {
            // "Updating crate-name v0.1 -> v0.2" or table "name  v0.1  v0.2  Yes"
            if (std.mem.indexOf(u8, line, "->") != null and hasDigit(line)) {
                const t = std.mem.trim(u8, line, " \t");
                if (t.len < 120 and !std.mem.startsWith(u8, t, "Package")) {
                    changes.append(gpa, t) catch @panic("oom");
                    if (changes.items.len >= 10) break;
                }
            }
        }
    } else if (std.mem.eql(u8, r.id, "pip") or std.mem.eql(u8, r.id, "uv-tools") or
        std.mem.eql(u8, r.id, "uv-python") or std.mem.eql(u8, r.id, "pipx"))
    {
        if (std.mem.eql(u8, r.id, "pipx")) {
            // "upgrading X..."
            var upgraded: std.ArrayList([]const u8) = .empty;
            for (lines) |line| {
                if (std.mem.startsWith(u8, line, "upgrading ") and std.mem.endsWith(u8, line, "...")) {
                    upgraded.append(gpa, line["upgrading ".len .. line.len - 3]) catch @panic("oom");
                }
            }
            if (upgraded.items.len != 0) {
                changes.append(gpa, std.mem.join(gpa, ", ", upgraded.items) catch @panic("oom")) catch @panic("oom");
            }
        } else {
            // "Successfully installed X-1.2.3 Y-4.5.6"
            for (lines) |line| {
                if (afterMarker(line, "Successfully installed")) |rest| {
                    const pkgs = std.mem.trim(u8, rest, " \t");
                    if (pkgs.len != 0) {
                        changes.append(gpa, pkgs) catch @panic("oom");
                        if (changes.items.len >= 5) break;
                    }
                }
            }
            // uv-tools: "Updated X 1.0 -> 1.1"
            for (lines) |line| {
                if (std.mem.indexOf(u8, line, "Updated") != null and std.mem.indexOf(u8, line, "->") != null) {
                    changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                    if (changes.items.len >= 10) break;
                }
            }
        }
    } else if (std.mem.eql(u8, r.id, "poetry")) {
        var pkg_count: usize = 0;
        var downgrades: usize = 0;
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "- Downgrading") != null) downgrades += 1;
        }
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "Package operations:") != null) {
                const trimmed = std.mem.trim(u8, line, " \t");
                const summary = if (afterMarker(trimmed, "Package operations: ")) |rest| rest else trimmed;
                if (std.mem.indexOf(u8, summary, "0 installs, 0 updates, 0 removals") == null) {
                    if (downgrades > 0) {
                        changes.append(gpa, fmtAlloc(gpa, "{s} -- {d} DOWNGRADED", .{ summary, downgrades })) catch @panic("oom");
                    } else {
                        changes.append(gpa, summary) catch @panic("oom");
                    }
                }
            } else if (pkg_count < 6 and
                (std.mem.indexOf(u8, line, "- Installing") != null or
                    std.mem.indexOf(u8, line, "- Updating") != null or
                    std.mem.indexOf(u8, line, "- Downgrading") != null) and
                std.mem.indexOfScalar(u8, line, '(') != null)
            {
                const trimmed = std.mem.trim(u8, line, " \t");
                const body = if (std.mem.startsWith(u8, trimmed, "- ")) trimmed[2..] else trimmed;
                changes.append(gpa, fmtAlloc(gpa, "  {s}", .{body})) catch @panic("oom");
                pkg_count += 1;
            }
        }
    } else if (std.mem.eql(u8, r.id, "mise") or std.mem.eql(u8, r.id, "mise-upgrade")) {
        for (lines) |line| {
            if ((std.mem.indexOf(u8, line, "→") != null or std.mem.indexOf(u8, line, "->") != null) and hasDigit(line)) {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                if (changes.items.len >= 10) break;
            }
        }
    } else if (std.mem.eql(u8, r.id, "gh-extensions")) {
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "upgraded") != null or
                (std.mem.indexOf(u8, line, "Updated") != null and std.mem.indexOf(u8, line, "→") != null))
            {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                if (changes.items.len >= 10) break;
            }
        }
    } else if (std.mem.eql(u8, r.id, "ruby-gems")) {
        for (lines) |line| {
            // "Updated X from 1.0 to 2.0" or "Updating X (1.0 -> 2.0)"
            if ((std.mem.indexOf(u8, line, "Updated") != null or std.mem.indexOf(u8, line, "Updating") != null) and
                (std.mem.indexOf(u8, line, "->") != null or std.mem.indexOf(u8, line, " to ") != null) and hasDigit(line))
            {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                if (changes.items.len >= 10) break;
            }
        }
    } else if (std.mem.eql(u8, r.id, "powershell-modules")) {
        for (lines) |line| {
            if ((std.mem.indexOf(u8, line, "Install") != null or std.mem.indexOf(u8, line, "Update") != null) and hasDigit(line)) {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
                if (changes.items.len >= 10) break;
            }
        }
    } else if (std.mem.eql(u8, r.id, "dotnet-workloads")) {
        var count: usize = 0;
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "Updated advertising manifest") != null) count += 1;
        }
        if (count > 0) changes.append(gpa, fmtAlloc(gpa, "{d} manifests updated", .{count})) catch @panic("oom");
    } else if (std.mem.eql(u8, r.id, "appx-repair")) {
        for (lines) |line| {
            if (std.mem.indexOf(u8, line, "re-registered") != null) {
                changes.append(gpa, std.mem.trim(u8, line, " \t")) catch @panic("oom");
            }
        }
    } else if (std.mem.eql(u8, r.id, "wsl-distros")) {
        var distro: []const u8 = "";
        for (lines) |line| {
            const clean = stripProgress(gpa, line);
            if (afterMarker(clean, "Updating WSL distro: ")) |d| {
                distro = std.mem.trim(u8, d, " \t");
            } else if (std.mem.indexOf(u8, clean, " upgraded, ") != null and
                !std.mem.startsWith(u8, clean, "0 upgraded,"))
            {
                const name = if (distro.len != 0) distro else "WSL";
                changes.append(gpa, fmtAlloc(gpa, "{s}: {s}", .{ name, clean })) catch @panic("oom");
            }
        }
    } else {
        // Generic: version arrows "X 1.2 -> 1.3" or "X v1 → v2"
        var count: usize = 0;
        for (lines) |line| {
            if ((std.mem.indexOf(u8, line, "->") != null or std.mem.indexOf(u8, line, "→") != null) and hasDigit(line)) {
                const t = std.mem.trim(u8, line, " \t");
                if (t.len < 120 and t.len != 0) {
                    changes.append(gpa, t) catch @panic("oom");
                    count += 1;
                    if (count >= 5) break;
                }
            }
        }
    }

    return changes.toOwnedSlice(gpa) catch @panic("oom");
}

fn printUpdateSummary(gpa: Allocator, io: Io, results: []const TaskSummary) void {
    const Entry = struct { task: []const u8, changes: [][]const u8 };
    var entries: std.ArrayList(Entry) = .empty;

    for (results) |r| {
        if (r.output_tail.len == 0) continue;
        const changes = changesFor(gpa, r);
        if (changes.len != 0) entries.append(gpa, .{ .task = r.id, .changes = changes }) catch @panic("oom");
    }
    if (entries.items.len == 0) return;

    p(io, "\n", .{});
    p(io, "{s}{s}{s} What's Changed:{s}\n", .{ col(A.magenta), G.changed, col(A.bold), col(A.reset) });
    for (entries.items) |entry| {
        p(io, "  {s}{s: <18}{s}  {s}\n", .{ col(A.cyan), entry.task, col(A.reset), entry.changes[0] });
        for (entry.changes[1..]) |change| {
            p(io, "  {s: <18}  {s}\n", .{ "", change });
        }
    }
}

// ─── State dir / previous run / process lock / scheduling ────────────────────

var env_handle: std.process.Environ = undefined;

fn envAlloc(gpa: Allocator, key: []const u8) ?[]const u8 {
    return std.process.Environ.getAlloc(env_handle, gpa, key) catch null;
}

fn joinPath(gpa: Allocator, parts: []const []const u8) []const u8 {
    return std.fs.path.join(gpa, parts) catch @panic("oom");
}

fn getStateDir(gpa: Allocator, cli: Cli) []const u8 {
    if (cli.state_dir) |dir| return dir;
    if (envAlloc(gpa, "LOCALAPPDATA")) |local| return joinPath(gpa, &.{ local, "Update-Everything" });
    if (envAlloc(gpa, "TEMP")) |tmp| return joinPath(gpa, &.{ tmp, "Update-Everything" });
    return "Update-Everything";
}

fn findRepoRoot(gpa: Allocator, io: Io) []const u8 {
    const marker = struct {
        fn hit(a: Allocator, i: Io, dir: []const u8) bool {
            return pathExists(i, joinPath(a, &.{ dir, "update-config.json" })) or
                pathExists(i, joinPath(a, &.{ dir, "updatescript.ps1" }));
        }
    }.hit;

    // 1) Walk up from the current working directory.
    const cwd = std.process.currentPathAlloc(io, gpa) catch "";
    var dir: []const u8 = cwd;
    while (dir.len != 0) {
        if (marker(gpa, io, dir)) return dir;
        dir = std.fs.path.dirname(dir) orelse break;
    }

    // 2) Walk up from the executable's own directory, so config discovery works
    //    when the exe is invoked from an arbitrary CWD (e.g. the scheduled task
    //    runs from System32).
    if (std.process.executablePathAlloc(io, gpa)) |exe| {
        var exe_dir = std.fs.path.dirname(exe);
        while (exe_dir) |d| {
            if (marker(gpa, io, d)) return d;
            exe_dir = std.fs.path.dirname(d);
        }
    } else |_| {}

    return cwd;
}

fn loadPrevSucceeded(gpa: Allocator, io: Io, cli: Cli, repo_root: []const u8) StringSet {
    var set: StringSet = .init(gpa);
    if (cli.since_hours <= 0) return set;

    var candidates: std.ArrayList([]const u8) = .empty;
    if (cli.state_dir) |dir| candidates.append(gpa, joinPath(gpa, &.{ dir, "last-run.json" })) catch @panic("oom");
    if (envAlloc(gpa, "LOCALAPPDATA")) |local| {
        candidates.append(gpa, joinPath(gpa, &.{ local, "Update-Everything", "last-run.json" })) catch @panic("oom");
    }
    candidates.append(gpa, joinPath(gpa, &.{ repo_root, "staging", "rust-run-summary.json" })) catch @panic("oom");

    for (candidates.items) |path| {
        const stat = Io.Dir.cwd().statFile(io, path, .{}) catch continue;
        const age_hours = @as(f64, @floatFromInt(nowUnixMs(io) - @divTrunc(stat.mtime.nanoseconds, std.time.ns_per_ms))) / 3_600_000.0;
        if (age_hours > cli.since_hours) continue;
        const text = Io.Dir.cwd().readFileAlloc(io, path, gpa, .limited(64 * 1024 * 1024)) catch continue;
        const parsed = std.json.parseFromSlice(std.json.Value, gpa, text, .{}) catch continue;
        if (parsed.value != .object) continue;
        const results = parsed.value.object.get("Results") orelse parsed.value.object.get("results") orelse continue;
        if (results != .array) continue;
        for (results.array.items) |item| {
            if (item != .object) continue;
            const id = item.object.get("Id") orelse item.object.get("id") orelse continue;
            const status = item.object.get("Status") orelse item.object.get("status") orelse continue;
            if (id != .string or status != .string) continue;
            if (std.mem.eql(u8, status.string, "Succeeded")) set.addNormalized(gpa, id.string);
        }
        return set;
    }
    return set;
}

fn nowUnixMs(io: Io) i64 {
    const ts = Io.Timestamp.now(io, .real);
    return @intCast(@divTrunc(ts.nanoseconds, std.time.ns_per_ms));
}

/// Single-instance guard. Atomic exclusive create so two instances racing to
/// start can't both observe "no lock" and both write their own PID.
fn acquireLock(gpa: Allocator, io: Io, state_dir: []const u8) ?[]const u8 {
    const path = joinPath(gpa, &.{ state_dir, "update-everything.lock" });
    Io.Dir.cwd().createDirPath(io, state_dir) catch {};

    var attempt: usize = 0;
    while (attempt < 2) : (attempt += 1) {
        if (Io.Dir.cwd().createFile(io, path, .{ .exclusive = true })) |file| {
            var buf: [64]u8 = undefined;
            const pid_text = std.fmt.bufPrint(&buf, "{d}", .{GetCurrentProcessId()}) catch "0";
            file.writeStreamingAll(io, pid_text) catch {};
            file.close(io);
            return path;
        } else |err| switch (err) {
            error.PathAlreadyExists => {
                const text = Io.Dir.cwd().readFileAlloc(io, path, gpa, .limited(64)) catch "";
                const pid = std.fmt.parseInt(u32, std.mem.trim(u8, text, " \t\r\n"), 10) catch 0;
                if (pid != 0 and isPidRunning(gpa, io, pid)) {
                    pe(io, "warn: another instance is already running (PID {d}); exiting\n", .{pid});
                    std.process.exit(5);
                }
                // Stale lock (process gone or unparseable) — remove and retry.
                Io.Dir.cwd().deleteFile(io, path) catch {};
            },
            // Can't create the lock file at all (e.g. permissions); run without
            // single-instance protection.
            else => return null,
        }
    }
    return null;
}

fn isPidRunning(gpa: Allocator, io: Io, pid: u32) bool {
    const filter = fmtAlloc(gpa, "PID eq {d}", .{pid});
    const result = std.process.run(gpa, io, .{
        .argv = &.{ "tasklist", "/FI", filter, "/NH", "/FO", "CSV" },
    }) catch return false;
    // tasklist CSV: header + one row per match; if only header, no match
    var it = std.mem.splitScalar(u8, result.stdout, '\n');
    _ = it.next();
    while (it.next()) |line| {
        if (std.mem.trim(u8, line, " \t\r").len != 0) return true;
    }
    return false;
}

fn scheduleTask(gpa: Allocator, io: Io, exe: []const u8, schedule_time: ?[]const u8) void {
    const task_name = "Update-Everything";
    const argv: []const []const u8 = if (schedule_time) |time|
        &.{ "schtasks", "/Create", "/F", "/TN", task_name, "/TR", exe, "/SC", "DAILY", "/ST", time }
    else
        &.{ "schtasks", "/Create", "/F", "/TN", task_name, "/TR", exe, "/SC", "ONLOGON" };

    const result = std.process.run(gpa, io, .{ .argv = argv }) catch |err| {
        fatal(io, "failed to run schtasks.exe: {t}", .{err});
    };
    if (result.term != .exited or result.term.exited != 0) {
        fatal(io, "schtasks /Create failed ({any})", .{result.term});
    }
    const trigger = if (schedule_time) |t| fmtAlloc(gpa, "daily at {s}", .{t}) else "at logon";
    p(io, "Registered scheduled task '{s}' ({s})\n", .{ task_name, trigger });
    p(io, "  exe: {s}\n", .{exe});
}

// ─── main ────────────────────────────────────────────────────────────────────

pub fn main(init: std.process.Init) !void {
    if (!is_windows) @compileError("this port targets Windows: the task table, timeout kill and exit-code probe are Win32-only");

    const gpa = init.arena.allocator();
    const io = init.io;
    env_handle = init.minimal.environ;
    initPathCache(gpa);
    buildPathIndex(gpa, io);

    const argv = try init.minimal.args.toSlice(gpa);
    const cli = parseCli(gpa, io, argv);

    color_on = cli.color orelse blk: {
        const out = Io.File.stdout();
        out.enableAnsiEscapeCodes(io) catch {};
        break :blk out.supportsAnsiEscapeCodes(io) catch false;
    };
    if (cli.log_file) |path| {
        // .read needed: length() stats the handle, which fails on a write-only handle.
        if (Io.Dir.cwd().createFile(io, path, .{ .truncate = false, .read = true })) |f| {
            log_file = f;
            log_pos = f.length(io) catch 0;
        } else |err| pe(io, "warn: could not open --log-file {s}: {t}\n", .{ path, err });
    }

    // Handle --schedule before anything else
    if (cli.schedule) {
        const exe = try std.process.executablePathAlloc(io, gpa);
        scheduleTask(gpa, io, exe, cli.schedule_time);
        return;
    }

    const start = nowMs(io);
    const started_at = nowString(gpa);
    const repo_root = findRepoRoot(gpa, io);
    const state_dir = getStateDir(gpa, cli);
    const config_path = cli.config orelse joinPath(gpa, &.{ repo_root, "update-config.json" });
    const config = loadConfig(gpa, io, config_path);

    // Prevent concurrent runs (non-fatal: if the lock fails we still continue)
    const lock_path = acquireLock(gpa, io, state_dir);
    defer if (lock_path) |path| Io.Dir.cwd().deleteFile(io, path) catch {};

    const tasks = buildTasks(gpa, io, config, cli);
    var prev_succeeded = loadPrevSucceeded(gpa, io, cli, repo_root);
    const selected = filterTasks(gpa, tasks, cli, config, &prev_succeeded);

    if (cli.list_tasks) {
        printTaskList(gpa, io, selected);
        return;
    }

    const jobs = @max(cli.jobs, 1);
    const results = runTasks(gpa, io, selected, cli, jobs, cli.task_timeout_sec * 1000);
    const duration_ms: u64 = @intCast(nowMs(io) - start);

    // Auto-save last-run.json to the state dir (enables --since-hours next run)
    if (!cli.dry_run) {
        const auto_path = joinPath(gpa, &.{ state_dir, "last-run.json" });
        writeSummaryJson(gpa, io, auto_path, started_at, duration_ms, cli.dry_run, results) catch |err| {
            pe(io, "warn: could not save {s}: {t}\n", .{ auto_path, err });
        };
    }

    // Also save to the explicit --json-summary path if given
    if (cli.json_summary) |path| {
        try writeSummaryJson(gpa, io, path, started_at, duration_ms, cli.dry_run, results);
        if (!cli.quiet) p(io, "summary {s}\n", .{path});
    }

    printSummary(gpa, io, results, duration_ms);

    if (!cli.dry_run) printUpdateSummary(gpa, io, results);

    if (cli.ci) {
        for (results) |r| {
            if (std.mem.eql(u8, r.status, "Failed") or std.mem.eql(u8, r.status, "TimedOut")) {
                std.process.exit(1);
            }
        }
    }
}
