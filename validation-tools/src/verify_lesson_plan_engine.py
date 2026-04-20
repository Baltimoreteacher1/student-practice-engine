#!/usr/bin/env python3

from __future__ import annotations

import os
import shlex
import subprocess
import tempfile
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
GENERATOR_DIR = WORKSPACE_ROOT / "codex-lesson-plan-generator"
SAMPLE_DECK = GENERATOR_DIR / "examples" / "sample_input_slides.pptx"
LAUNCHER = WORKSPACE_ROOT / "Generate Lesson Plan.command"


def run_checked(command: list[str], *, cwd: Path, env: dict[str, str] | None = None) -> None:
    display = shlex.join(command)
    print(f"$ {display}")
    subprocess.run(command, cwd=str(cwd), env=env, check=True)


def require_file(path: Path) -> None:
    if not path.exists():
        raise SystemExit(f"Missing expected verification artifact: {path}")


def main() -> int:
    if not SAMPLE_DECK.exists():
        raise SystemExit(f"Sample deck not found: {SAMPLE_DECK}")
    if not LAUNCHER.exists():
        raise SystemExit(f"Launcher not found: {LAUNCHER}")

    run_checked(["python3", "-m", "pytest", "-q", "tests"], cwd=GENERATOR_DIR)
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
        ],
        cwd=GENERATOR_DIR,
    )

    with tempfile.TemporaryDirectory(prefix="lesson_plan_engine_verify_") as temp_dir:
        verification_root = Path(temp_dir)
        inbox_dir = verification_root / "INBOX"
        output_dir = verification_root / "OUTPUT"
        inbox_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        env = os.environ.copy()
        env["LESSON_PLAN_NO_OPEN"] = "1"
        env["EDUWONDERLAB_WATCH_DIR"] = str(inbox_dir)
        env["EDUWONDERLAB_OUTPUT_DIR"] = str(output_dir)
        run_checked(
            [str(LAUNCHER), "--deck", str(SAMPLE_DECK)],
            cwd=WORKSPACE_ROOT,
            env=env,
        )

        for relative_path in (
            Path("lesson_plan.json"),
            Path("lesson_plan.md"),
            Path("lesson_plan.docx"),
            Path("validation_report.md"),
            Path("extracted/raw_slide_text.json"),
            Path("extracted/normalized_lesson.json"),
            Path("extracted/source_fidelity_map.json"),
        ):
            require_file(output_dir / relative_path)

    print("Lesson plan engine verification passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
