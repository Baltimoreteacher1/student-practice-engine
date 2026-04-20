from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from polish_notebook_pptx import polish_notebook_pptx, write_pptx_polish_report
from premium_polish import run_notebook_enhancement


REPO_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_BASE_DIR = Path(
    os.environ.get("EDUWONDERLAB_NOTEBOOK_ENHANCEMENT_DIR")
    or Path(__file__).resolve().parents[1]
)
DEFAULT_INBOX_DIR = DEFAULT_BASE_DIR / "INBOX"
DEFAULT_OUTPUT_DIR = DEFAULT_BASE_DIR / "OUTPUT"
DEFAULT_ARCHIVE_DIR = DEFAULT_BASE_DIR / "ARCHIVE"
NOTEBOOK_ENGINE_PATH = REPO_ROOT / "notebook_engine.py"


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")


def discover_bundle_dirs(inbox_dir: Path) -> list[Path]:
    if not inbox_dir.exists():
        return []
    direct_plan = inbox_dir / "notebook_plan.json"
    if direct_plan.exists():
        return [inbox_dir.resolve()]
    bundles: list[Path] = []
    for child in sorted(inbox_dir.iterdir()):
        if child.is_dir() and (child / "notebook_plan.json").exists():
            bundles.append(child.resolve())
    return bundles


def discover_pptx_files(inbox_dir: Path) -> list[Path]:
    if not inbox_dir.exists():
        return []
    return sorted(path.resolve() for path in inbox_dir.glob("*.pptx") if path.is_file())


def render_polished_bundle(output_job_dir: Path) -> Path | None:
    plan_path = output_job_dir / "notebook_plan.json"
    deck_path = output_job_dir / "source_deck.json"
    if not plan_path.exists() or not deck_path.exists():
        return None
    rendered_dir = output_job_dir / "rendered"
    command = [
        sys.executable,
        str(NOTEBOOK_ENGINE_PATH),
        "render",
        str(plan_path),
        "--deck",
        str(deck_path),
        "--output-dir",
        str(rendered_dir),
    ]
    subprocess.run(command, check=True)
    return rendered_dir


def process_enhancement_inbox(
    inbox_dir: Path,
    output_dir: Path,
    *,
    render_outputs: bool = True,
) -> dict[str, Any]:
    inbox_dir = inbox_dir.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    bundles = discover_bundle_dirs(inbox_dir)
    pptx_files = discover_pptx_files(inbox_dir)
    results: list[dict[str, Any]] = []
    for bundle_dir in bundles:
        output_job_dir = output_dir / bundle_dir.name
        try:
            report = run_notebook_enhancement(bundle_dir, output_job_dir)
            premium_passed = bool(report.get("finalReport", {}).get("passed", False))
            result: dict[str, Any] = {
                "job": bundle_dir.name,
                "bundle_dir": str(bundle_dir),
                "output_dir": str(output_job_dir),
                "status": "polished",
                "premium_passed": premium_passed,
                "manual_review_required": not premium_passed,
                "quality_tier": report.get("qualityTier", report.get("finalReport", {}).get("qualityTier", "fail")),
                "repairs_applied": len(report.get("finalReport", {}).get("repairsApplied", [])),
            }
            if render_outputs:
                rendered_dir = render_polished_bundle(output_job_dir)
                if rendered_dir is not None:
                    result["rendered_dir"] = str(rendered_dir)
                    result["status"] = "rendered"
            results.append(result)
        except Exception as exc:  # noqa: BLE001
            results.append(
                {
                    "job": bundle_dir.name,
                    "bundle_dir": str(bundle_dir),
                    "output_dir": str(output_job_dir),
                    "status": "failed",
                    "premium_passed": False,
                    "quality_tier": "fail",
                    "repairs_applied": 0,
                    "error": str(exc),
                    "manual_review_required": True,
                }
            )
    for pptx_path in pptx_files:
        output_job_dir = output_dir / pptx_path.stem
        output_job_dir.mkdir(parents=True, exist_ok=True)
        polished_name = f"{pptx_path.stem} - Polished.pptx"
        polished_path = output_job_dir / polished_name
        try:
            report = polish_notebook_pptx(pptx_path, polished_path)
            write_pptx_polish_report(output_job_dir / "pptx_polish_report.json", report)
            quality_tier = str(report.get("qualityTier", "enhanced"))
            manual_review_required = quality_tier != "premium" or bool(report.get("warnings"))
            results.append(
                {
                    "job": pptx_path.stem,
                    "bundle_dir": str(pptx_path),
                    "output_dir": str(output_job_dir),
                    "rendered_dir": str(output_job_dir),
                    "status": "polished_pptx",
                    "premium_passed": quality_tier == "premium",
                    "manual_review_required": manual_review_required,
                    "quality_tier": quality_tier,
                    "repairs_applied": int(report["stats"]["fontLifts"]) + int(report["stats"]["marginRepairs"]) + int(report["stats"]["wrapRepairs"]),
                    "polished_pptx": str(polished_path),
                }
            )
        except Exception as exc:  # noqa: BLE001
            results.append(
                {
                    "job": pptx_path.stem,
                    "bundle_dir": str(pptx_path),
                    "output_dir": str(output_job_dir),
                    "status": "failed",
                    "premium_passed": False,
                    "manual_review_required": True,
                    "quality_tier": "fail",
                    "repairs_applied": 0,
                    "error": str(exc),
                }
            )
    has_failures = any(result.get("status") == "failed" for result in results)
    needs_review = any(result.get("manual_review_required") for result in results if result.get("status") != "failed")
    summary = {
        "status": "completed_with_errors"
        if has_failures
        else "completed_with_review_needed"
        if needs_review
        else "completed"
        if results
        else "no_files",
        "processed_count": len(results),
        "inbox_dir": str(inbox_dir),
        "output_dir": str(output_dir),
        "results": results,
    }
    write_json(output_dir / "enhancement_inbox_last_run.json", summary)
    return summary


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Process notebook enhancement bundles from an inbox into polished outputs."
    )
    parser.add_argument("--inbox-dir", default=str(DEFAULT_INBOX_DIR), help="Directory containing bundle folders to polish")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Directory for polished outputs")
    parser.add_argument("--skip-render", action="store_true", help="Polish bundles but skip PPTX rerendering")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    summary = process_enhancement_inbox(
        Path(args.inbox_dir),
        Path(args.output_dir),
        render_outputs=not args.skip_render,
    )
    print(f"Run status: {summary['status']}")
    print(f"Processed bundles: {summary['processed_count']}")
    print(f"Output folder: {summary['output_dir']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
