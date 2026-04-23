#!/usr/bin/env python3

from __future__ import annotations

import shlex
import subprocess
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_DIR = WORKSPACE_ROOT / "codex-lesson-plan-generator"
SAMPLE_DECK = GENERATOR_DIR / "examples" / "sample_input_slides.pptx"
LAUNCHER = WORKSPACE_ROOT / "Generate Lesson Plan.command"


def run_checked(command: list[str], *, cwd: Path, env: dict[str, str] | None = None) -> None:
    display = shlex.join(command)
    print(f"$ {display}")
    subprocess.run(command, cwd=str(cwd), env=env, check=True)

def main() -> int:
    if not SAMPLE_DECK.exists():
        raise SystemExit(f"Sample deck not found: {SAMPLE_DECK}")
    if not LAUNCHER.exists():
        raise SystemExit(f"Launcher not found: {LAUNCHER}")

    run_checked(["python3", "-m", "pytest", "-q", "tests"], cwd=GENERATOR_DIR)
    run_checked(
        ["python3", "-m", "pytest", "-q", "tests/test_lesson_plan_quality_benchmark.py"],
        cwd=WORKSPACE_ROOT,
    )
    run_checked(
        [
            "python3",
            "-m",
            "py_compile",
            "run.py",
            "src/build_lesson_plan.py",
            "src/extract_slides.py",
            "src/render_docx.py",
            "src/utils.py",
            "src/validate_plan.py",
            "../validation-tools/src/benchmark_lesson_plan_quality.py",
        ],
        cwd=GENERATOR_DIR,
    )
    run_checked(
        ["python3", "validation-tools/src/benchmark_lesson_plan_quality.py"],
        cwd=WORKSPACE_ROOT,
    )

    print("Lesson plan engine verification passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
