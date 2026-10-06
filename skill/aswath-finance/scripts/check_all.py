"""Run every engine's --selftest (each replays cases taken from Damodaran's own spreadsheets) and the
reference audit (every numeric claim vs its cited post). Exit code 1 if anything fails.

Usage: python check_all.py
"""
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def run(args: list[str]) -> tuple[bool, str]:
    p = subprocess.run([sys.executable, *args], cwd=HERE, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    last = (p.stdout.strip().splitlines() or [p.stderr.strip()[-200:]])[-1]
    return p.returncode == 0, last


def main() -> int:
    maintenance = {"check_all.py", "audit.py", "calibrate.py", "branch_test.py", "build_post_index.py"}
    engines = sorted(f.name for f in HERE.glob("*.py") if "--selftest" in f.read_text(encoding="utf-8")
                     and f.name not in maintenance)
    bad = 0
    for name in engines:
        ok, last = run([name, "--selftest"])
        bad += not ok
        print(f"{'ok ' if ok else 'BAD'} {name:24s} {last}")
    ok, last = run(["audit.py"])
    print(f"{'ok ' if ok else 'BAD'} {'audit.py':24s} {last}")
    return 1 if bad or not ok else 0


if __name__ == "__main__":
    sys.exit(main())
