#!/usr/bin/env python3
"""Process downloaded PPTX decks from a folder into compact Session 1 notebooks."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from notebook_engine import (
    DEFAULT_MODEL,
    PREMIUM_NOTEBOOK_GUIDANCE,
    choose_api_key,
    enforce_runtime_quality_guidance,
    normalize_whitespace,
    normalize_saved_guidance,
    slugify,
)
from notebook_engine_app import generate_job, load_config
from notebook_launchers import (
    DEFAULT_ARCHIVE_DIR,
    DEFAULT_INBOX_DIR,
    DEFAULT_OUTPUT_DIR,
    ensure_inbox_launcher,
)
RUN_LOG_FILENAME = "notebook_inbox_last_run.log"
RUN_SUMMARY_FILENAME = "notebook_inbox_last_run.json"
LOCKED_FOLDER_GUIDANCE = PREMIUM_NOTEBOOK_GUIDANCE


def ensure_pipeline_dirs(*paths: Path) -> None:
    for path in paths:
        path.mkdir(parents=True, exist_ok=True)


def log_line(message: str, log_path: Path | None = None) -> None:
    print(message, flush=True)
    if log_path is None:
        return
    log_path.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{timestamp}] {message}\n")


def write_run_summary(summary_path: Path, payload: dict[str, Any]) -> None:
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def list_pptx_files(folder: Path) -> list[Path]:
    return sorted(
        [
            path
            for path in folder.iterdir()
            if path.is_file() and path.suffix.lower() == ".pptx" and not path.name.startswith("~$")
        ],
        key=lambda item: (item.stat().st_mtime, item.name.lower()),
    )


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 2
    while True:
        candidate = parent / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def unique_dir(path: Path) -> Path:
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = path.parent / f"{path.name}-{index}"
        if not candidate.exists():
            return candidate
        index += 1


def merged_guidance(config_guidance: str, extra_guidance: str) -> str:
    locked = normalize_saved_guidance(LOCKED_FOLDER_GUIDANCE)
    saved = normalize_saved_guidance(config_guidance)
    if saved == locked:
        saved = ""
    extra = normalize_whitespace(extra_guidance)
    return enforce_runtime_quality_guidance(
        normalize_whitespace(" ".join(part for part in [locked, saved, extra] if part))
    )


def copy_job_outputs(job: dict[str, Any], output_dir: Path) -> dict[str, Any]:
    job_id = job["job_id"]
    source_filename = job["source_filename"]
    from notebook_engine_app import RUNS_DIR

    job_dir = RUNS_DIR / job_id
    deck_stem = Path(source_filename).stem
    export_dir = unique_dir(output_dir / f"{slugify(deck_stem)[:48]}-notebooks")
    export_dir.mkdir(parents=True, exist_ok=True)

    session1_src = job_dir / job["relative_files"]["session1"]
    plan_src = job_dir / job["relative_files"]["plan"]
    deck_src = job_dir / job["relative_files"]["deck"]
    quality_src = job_dir / job["relative_files"]["quality_report"]

    session1_dest = export_dir / f"{deck_stem} - Session 1 Student Notebook.pptx"
    plan_dest = export_dir / "notebook_plan.json"
    deck_dest = export_dir / "source_deck.json"
    quality_dest = export_dir / "quality_report.json"
    manifest_dest = export_dir / "job.json"

    shutil.copy2(session1_src, session1_dest)
    shutil.copy2(plan_src, plan_dest)
    shutil.copy2(deck_src, deck_dest)
    shutil.copy2(quality_src, quality_dest)
    manifest_dest.write_text(json.dumps(job, indent=2), encoding="utf-8")

    copied = {
        "export_dir": export_dir,
        "session1": session1_dest,
        "plan": plan_dest,
        "deck": deck_dest,
        "quality_report": quality_dest,
        "manifest": manifest_dest,
    }

    session1_output = unique_path(output_dir / session1_dest.name)
    shutil.copy2(session1_dest, session1_output)
    copied["session1_output"] = session1_output

    session2_relative = job["relative_files"].get("session2")
    if session2_relative:
        session2_src = job_dir / session2_relative
        session2_dest = export_dir / f"{deck_stem} - Session 2 Student Notebook.pptx"
        shutil.copy2(session2_src, session2_dest)
        copied["session2"] = session2_dest
        session2_output = unique_path(output_dir / session2_dest.name)
        shutil.copy2(session2_dest, session2_output)
        copied["session2_output"] = session2_output

    html_relative = job["relative_files"].get("html_notebook")
    if html_relative:
        html_src = job_dir / html_relative
        html_dest = export_dir / f"{deck_stem} - Interactive Student Notebook.html"
        shutil.copy2(html_src, html_dest)
        copied["html_notebook"] = html_dest
        html_output = unique_path(output_dir / html_dest.name)
        shutil.copy2(html_dest, html_output)
        copied["html_notebook_output"] = html_output

    return copied


def archive_source(source_path: Path, archive_dir: Path) -> Path:
    archived = unique_path(archive_dir / source_path.name)
    shutil.move(str(source_path), str(archived))
    return archived


def process_deck(
    source_path: Path,
    *,
    output_dir: Path,
    archive_dir: Path,
    model: str,
    offline: bool,
    guidance: str,
    api_key: str,
    log_path: Path | None = None,
) -> dict[str, Any]:
    log_line(f"Processing: {source_path}", log_path)
    uploaded_bytes = source_path.read_bytes()
    try:
        job = generate_job(
            uploaded_filename=source_path.name,
            uploaded_bytes=uploaded_bytes,
            model=model,
            offline=offline,
            guidance=guidance,
            api_key=api_key,
        )
    except Exception as exc:
        mode_label = "offline" if offline else "live"
        raise RuntimeError(
            f"{mode_label.title()} notebook generation failed for {source_path.name}: {exc}"
        ) from exc
    copied = copy_job_outputs(job, output_dir)
    log_line(
        f"  Session 1 output: {copied.get('session1_output', copied['session1'])}",
        log_path,
    )
    log_line(f"  Notebook package: {copied['export_dir']}", log_path)
    if copied.get("session2"):
        log_line(
            f"  Session 2 output: {copied.get('session2_output', copied['session2'])}",
            log_path,
        )
    if copied.get("html_notebook"):
        log_line(
            "  HTML Web App output: "
            f"{copied.get('html_notebook_output', copied['html_notebook'])}",
            log_path,
        )
    log_line(f"  Quality report: {copied['quality_report']}", log_path)
    if not job.get("quality_passed", False):
        raise RuntimeError("Generated notebook failed the production quality validation gate.")
    try:
        archived_path = archive_source(source_path, archive_dir)
        log_line(f"  Archived source: {archived_path}", log_path)
    except FileNotFoundError:
        archived_path = source_path
        log_line(f"  Source deleted from inbox before archiving.", log_path)
    return {
        "job": job,
        "copied": copied,
        "archived_source": archived_path,
    }


def process_inbox(
    *,
    input_dir: Path,
    output_dir: Path,
    archive_dir: Path,
    model: str,
    offline: bool,
    guidance: str,
    api_key: str,
    log_path: Path | None = None,
    limit: int = 0,
) -> list[dict[str, Any]]:
    ensure_pipeline_dirs(input_dir, output_dir, archive_dir)
    decks = list_pptx_files(input_dir)
    if limit > 0:
        decks = decks[:limit]
    results: list[dict[str, Any]] = []
    for deck in decks:
        results.append(
            process_deck(
                deck,
                output_dir=output_dir,
                archive_dir=archive_dir,
                model=model,
                offline=offline,
                guidance=guidance,
                api_key=api_key,
                log_path=log_path,
            )
        )
    return results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Turn PPTX decks in a folder into compact Session 1 notebooks.")
    parser.add_argument("--input-dir", default=str(DEFAULT_INBOX_DIR), help="Folder to scan for .pptx lesson decks")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="Folder where notebook exports will be copied")
    parser.add_argument("--archive-dir", default=str(DEFAULT_ARCHIVE_DIR), help="Folder where processed source decks will be moved")
    parser.add_argument("--model", default="", help="Optional model override")
    parser.add_argument("--guidance", default="", help="Optional extra guidance appended to the locked defaults")
    parser.add_argument("--offline", action="store_true", help="Use the local heuristic planner instead of the API")
    parser.add_argument("--watch", action="store_true", help="Keep polling the input folder for new .pptx files")
    parser.add_argument("--interval", type=int, default=20, help="Polling interval in seconds when --watch is enabled")
    parser.add_argument("--limit", type=int, default=0, help="Maximum number of decks to process in one pass")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    input_dir = Path(args.input_dir).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    archive_dir = Path(args.archive_dir).expanduser().resolve()
    ensure_pipeline_dirs(input_dir, output_dir, archive_dir)
    ensure_inbox_launcher(inbox_dir=input_dir)
    log_path = output_dir / RUN_LOG_FILENAME
    summary_path = output_dir / RUN_SUMMARY_FILENAME
    log_path.write_text("", encoding="utf-8")
    config = load_config()
    model = normalize_whitespace(args.model) or config.get("model") or DEFAULT_MODEL
    guidance = merged_guidance(config.get("default_guidance", ""), args.guidance)
    api_key = choose_api_key(os.getenv("OPENAI_API_KEY", ""), config.get("api_key", ""))

    if not args.offline and not api_key:
        log_line("No OpenAI API key is saved. Add one in notebook_engine_app_data/config.json or run with --offline.", log_path)
        log_line(f"Input folder: {input_dir}", log_path)
        log_line(f"Output folder: {output_dir}", log_path)
        write_run_summary(
            summary_path,
            {
                "status": "failed",
                "reason": "missing_api_key",
                "input_dir": str(input_dir),
                "output_dir": str(output_dir),
                "archive_dir": str(archive_dir),
                "log_path": str(log_path),
                "timestamp": datetime.now().isoformat(timespec="seconds"),
            },
        )
        return 1

    log_line(f"Input folder:   {input_dir}", log_path)
    log_line(f"Output folder:  {output_dir}", log_path)
    log_line(f"Archive folder: {archive_dir}", log_path)
    log_line(f"Model:          {model}", log_path)
    log_line(f"Offline mode:   {args.offline}", log_path)

    try:
        while True:
            results = process_inbox(
                input_dir=input_dir,
                output_dir=output_dir,
                archive_dir=archive_dir,
                model=model,
                offline=args.offline,
                guidance=guidance,
                api_key=api_key,
                log_path=log_path,
                limit=args.limit,
            )
            if results:
                log_line(f"Processed {len(results)} deck(s).", log_path)
                write_run_summary(
                    summary_path,
                    {
                        "status": "completed",
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "input_dir": str(input_dir),
                        "output_dir": str(output_dir),
                        "archive_dir": str(archive_dir),
                        "model": model,
                        "offline_requested": args.offline,
                        "log_path": str(log_path),
                        "processed_count": len(results),
                        "results": [
                            {
                                "source_filename": result["job"]["source_filename"],
                                "fallback_mode": result["job"].get("fallback_mode", ""),
                                "session1": str(
                                    result["copied"].get(
                                        "session1_output", result["copied"]["session1"]
                                    )
                                ),
                                "session1_output": str(
                                    result["copied"].get(
                                        "session1_output", result["copied"]["session1"]
                                    )
                                ),
                                "session1_package": str(result["copied"]["session1"]),
                                "quality_report": str(result["copied"]["quality_report"]),
                                "export_dir": str(result["copied"]["export_dir"]),
                                "archived_source": str(result["archived_source"]),
                                **(
                                    {
                                        "session2": str(
                                            result["copied"].get(
                                                "session2_output",
                                                result["copied"]["session2"],
                                            )
                                        ),
                                        "session2_output": str(
                                            result["copied"].get(
                                                "session2_output",
                                                result["copied"]["session2"],
                                            )
                                        ),
                                        "session2_package": str(
                                            result["copied"]["session2"]
                                        ),
                                    }
                                    if result["copied"].get("session2")
                                    else {}
                                ),
                                **(
                                    {
                                        "html_notebook": str(
                                            result["copied"].get(
                                                "html_notebook_output",
                                                result["copied"]["html_notebook"],
                                            )
                                        ),
                                        "html_notebook_output": str(
                                            result["copied"].get(
                                                "html_notebook_output",
                                                result["copied"]["html_notebook"],
                                            )
                                        ),
                                        "html_notebook_package": str(
                                            result["copied"]["html_notebook"]
                                        ),
                                    }
                                    if result["copied"].get("html_notebook")
                                    else {}
                                ),
                            }
                            for result in results
                        ],
                    },
                )
            else:
                log_line("No .pptx files found in the input folder.", log_path)
                write_run_summary(
                    summary_path,
                    {
                        "status": "no_files",
                        "timestamp": datetime.now().isoformat(timespec="seconds"),
                        "input_dir": str(input_dir),
                        "output_dir": str(output_dir),
                        "archive_dir": str(archive_dir),
                        "model": model,
                        "offline_requested": args.offline,
                        "log_path": str(log_path),
                        "processed_count": 0,
                        "results": [],
                    },
                )

            if not args.watch:
                break

            time.sleep(max(5, args.interval))
        return 0
    except Exception as exc:
        log_line(f"Run failed: {exc}", log_path)
        log_line(traceback.format_exc().rstrip(), log_path)
        write_run_summary(
            summary_path,
            {
                "status": "failed",
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "input_dir": str(input_dir),
                "output_dir": str(output_dir),
                "archive_dir": str(archive_dir),
                "model": model,
                "offline_requested": args.offline,
                "log_path": str(log_path),
                "error": str(exc),
            },
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
