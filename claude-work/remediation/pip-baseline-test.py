"""Exercises the pip conflict-baseline branch that main.zig embeds.

The embedded script needs a real `pip check`, so the decision logic is replayed
here against the actual conflict lines from the 2026-09-01 run.
Run: python claude-work/remediation/pip-baseline-test.py
"""
import json
import os
import sys
import tempfile

REAL_CONFLICTS = [
    "sglang 0.5.10 requires apache-tvm-ffi, which is not installed.",
    "datasets 5.0.1 has requirement fsspec[http]<=2026.6.0,>=2023.1.0, but you have fsspec 2026.7.0.",
    "vllm 0.28.0 has requirement lark==1.2.2, but you have lark 1.3.1.",
]


def decide(baseline_dir, returncode, stdout):
    """Mirror of the embedded branch: returns (blocked, new_conflicts, known_count)."""
    baseline_path = os.path.join(baseline_dir, "pip-conflicts.json")
    current = sorted(l.strip() for l in stdout.splitlines() if l.strip())
    try:
        with open(baseline_path, encoding="utf-8") as fh:
            known = set(json.load(fh))
    except (OSError, ValueError):
        known = set()
    new_conflicts = [l for l in current if l not in known]
    blocked = returncode != 0 and bool(new_conflicts)
    os.makedirs(baseline_dir, exist_ok=True)
    with open(baseline_path, "w", encoding="utf-8") as fh:
        json.dump(current, fh, indent=1)
    return blocked, new_conflicts, len(current) - len(new_conflicts)


def main():
    failures = 0
    with tempfile.TemporaryDirectory() as tmp:
        out = "\n".join(REAL_CONFLICTS)

        # First run: conflicts unknown -> block, and record them.
        blocked, new, known = decide(tmp, 1, out)
        ok = blocked and len(new) == 3 and known == 0
        print(f"{'PASS' if ok else 'FAIL'}  first run blocks on 3 unseen conflicts (new={len(new)})")
        failures += not ok

        # Second run, same conflicts: this is where the old code skipped forever.
        blocked, new, known = decide(tmp, 1, out)
        ok = not blocked and not new and known == 3
        print(f"{'PASS' if ok else 'FAIL'}  unchanged conflicts no longer block (known={known})")
        failures += not ok

        # A genuinely new conflict must still stop the upgrade.
        out2 = out + "\nnumpy 2.5.2 has requirement foo==1.0, but you have foo 2.0."
        blocked, new, known = decide(tmp, 1, out2)
        ok = blocked and len(new) == 1 and known == 3
        print(f"{'PASS' if ok else 'FAIL'}  a new conflict still blocks (new={len(new)}, known={known})")
        failures += not ok

        # Clean environment: nothing to block on.
        blocked, new, known = decide(tmp, 0, "")
        ok = not blocked and not new
        print(f"{'PASS' if ok else 'FAIL'}  clean pip check does not block")
        failures += not ok

    print("FAILURES:", failures)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
