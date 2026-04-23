#!/usr/bin/env python3

from __future__ import annotations

import shlex
import subprocess
import sys
from pathlib import Path

WORKSPACE_ROOT = Path(__file__).resolve().parents[2]


def run_checked(command: list[str], *, cwd: Path) -> None:
    display = shlex.join(command)
    print(f"$ {display}")
    subprocess.run(command, cwd=str(cwd), check=True)

def main() -> int:
    run_checked(
        [
            "python3",
            "-m",
            "pytest",
            "-q",
            "tests/test_notebook_session1_pivot.py",
            "tests/test_notebook_engine_publisher_polish.py",
            "tests/test_notebook_quality_benchmark.py",
        ],
        cwd=WORKSPACE_ROOT,
    )
    run_checked(
        [
            "python3",
            "-m",
            "py_compile",
            "notebook_engine.py",
            "notebook_engine_app.py",
            "notebook_folder_runner.py",
            "validation-tools/src/benchmark_notebook_quality.py",
            "notebook-engine/enhancement/src/premium_polish.py",
            "notebook-engine/enhancement/src/run_enhancement_inbox.py",
        ],
        cwd=WORKSPACE_ROOT,
    )
    run_checked(
        ["python3", "validation-tools/src/benchmark_notebook_quality.py"],
        cwd=WORKSPACE_ROOT,
    )

    print("Notebook engine verification passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
