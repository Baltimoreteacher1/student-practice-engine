#!/usr/bin/env python3

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

CONFIG = {
    "default_config_path": "config/generator_config.json",
    "default_school_defaults_path": "config/school_defaults.json",
    "supported_trigger_examples": [
        "Ready 4.7.2026",
        "Ready 04.07.2026",
        "Ready 2026-04-07",
        "LP: 4.7.2026",
    ],
}

ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = ROOT.parent
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from apply_supports import apply_supports  # noqa: E402
from build_lesson_plan import build_lesson_plan  # noqa: E402
from detect_lesson_type import detect_lesson_type  # noqa: E402
from extract_slides import extract_slide_deck  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from render_docx import ensure_template_docx, render_docx, render_markdown  # noqa: E402
from utils import (  # noqa: E402
    LessonPlanError,
    ensure_directory,
    load_effective_config,
    parse_trigger_text,
    read_optional_agenda,
    resolve_run_date,
    select_slide_deck,
    write_json,
)
from validate_plan import build_validation_payload, write_validation_report  # noqa: E402


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate a deterministic lesson plan from a teacher slide deck.")
    parser.add_argument("--deck", default="", help="Optional explicit path to a .pptx deck.")
    parser.add_argument("--config", default=CONFIG["default_config_path"], help="Path to generator_config.json")
    parser.add_argument(
        "--school-defaults",
        default=CONFIG["default_school_defaults_path"],
        help="Path to school_defaults.json",
    )
    parser.add_argument("--trigger", default="", help="Optional Ready/LP trigger text such as 'Ready 2026-04-07'.")
    return parser


def validateConfig(config: dict[str, object]) -> None:
    required_roots = ("input", "extracted", "output", "templates")
    for key in required_roots:
        if key not in config or not isinstance(config[key], dict):
            raise LessonPlanError(f"Config is missing required block '{key}'.")

    input_config = config["input"]
    output_config = config["output"]
    templates_config = config["templates"]
    required_paths = {
        "input.slides_dir": input_config.get("slides_dir"),
        "output.json": output_config.get("json"),
        "output.markdown": output_config.get("markdown"),
        "output.docx": output_config.get("docx"),
        "output.validation_report": output_config.get("validation_report"),
        "templates.docx": templates_config.get("docx"),
        "templates.markdown": templates_config.get("markdown"),
    }
    for label, value in required_paths.items():
        if not value:
            raise LessonPlanError(f"Config value '{label}' must not be blank.")


def selfTest(config: dict[str, object]) -> None:
    expected_duration = int(config.get("lesson_duration_minutes", 55))
    if expected_duration <= 0:
        raise LessonPlanError("lesson_duration_minutes must be greater than zero.")
    phase_timing = dict(config.get("phase_timing", {}))
    if phase_timing:
        total = sum(int(value) for value in phase_timing.values())
        if total != expected_duration:
            raise LessonPlanError(
                f"Configured phase_timing totals {total} minutes, expected {expected_duration}."
            )


def applyEnvironmentOverrides(config: dict[str, object]) -> dict[str, object]:
    watch_dir = os.environ.get("EDUWONDERLAB_WATCH_DIR", "").strip()
    output_dir = os.environ.get("EDUWONDERLAB_OUTPUT_DIR", "").strip()
    if watch_dir:
        config["input"]["slides_dir"] = watch_dir
        config["input"]["agenda_dir"] = str(Path(watch_dir) / "agenda")
        config["input"]["notes_dir"] = str(Path(watch_dir) / "notes")
    if output_dir:
        output_root = Path(output_dir)
        config["extracted"]["raw_slide_text"] = str(output_root / "extracted" / "raw_slide_text.json")
        config["extracted"]["normalized_lesson"] = str(output_root / "extracted" / "normalized_lesson.json")
        config["extracted"]["source_fidelity_map"] = str(output_root / "extracted" / "source_fidelity_map.json")
        config["output"]["json"] = str(output_root / "lesson_plan.json")
        config["output"]["markdown"] = str(output_root / "lesson_plan.md")
        config["output"]["docx"] = str(output_root / "lesson_plan.docx")
        config["output"]["validation_report"] = str(output_root / "validation_report.md")
    return config


def main() -> int:
    parser = build_parser()
    args, extras = parser.parse_known_args()

    try:
        config = load_effective_config(ROOT / args.config, ROOT / args.school_defaults)
        config = applyEnvironmentOverrides(config)
        validateConfig(config)
        selfTest(config)

        input_slides_dir = ROOT / config["input"]["slides_dir"]
        input_agenda_dir = ROOT / config["input"]["agenda_dir"]
        extracted_config = config["extracted"]
        output_config = config["output"]
        template_config = config["templates"]

        raw_output_path = ROOT / extracted_config["raw_slide_text"]
        normalized_output_path = ROOT / extracted_config["normalized_lesson"]
        fidelity_output_path = ROOT / extracted_config["source_fidelity_map"]
        lesson_json_path = ROOT / output_config["json"]
        lesson_md_path = ROOT / output_config["markdown"]
        lesson_docx_path = ROOT / output_config["docx"]
        validation_report_path = ROOT / output_config["validation_report"]
        template_docx_path = ROOT / template_config["docx"]
        template_md_path = ROOT / template_config["markdown"]

        for path in (
            input_slides_dir,
            input_agenda_dir,
            raw_output_path.parent,
            normalized_output_path.parent,
            fidelity_output_path.parent,
            lesson_json_path.parent,
            template_docx_path.parent,
        ):
            ensure_directory(path)

        trigger_text = args.trigger or " ".join(extras).strip()
        trigger_context = parse_trigger_text(trigger_text)
        run_date = resolve_run_date(trigger_context, config)

        deck_path = select_slide_deck(input_slides_dir, args.deck, WORKSPACE_ROOT)
        agenda_items = read_optional_agenda(input_agenda_dir)
        ensure_template_docx(template_docx_path, config)

        raw_deck = extract_slide_deck(deck_path, raw_output_path, base_dir=WORKSPACE_ROOT)
        lesson_extract = run_lesson_extract(raw_deck, config)
        write_json(normalized_output_path, lesson_extract)

        lesson_type = detect_lesson_type(raw_deck)
        lesson_plan = build_lesson_plan(
            lesson_extract=lesson_extract,
            raw_deck=raw_deck,
            lesson_type=lesson_type,
            config=config,
            fidelity_output_path=fidelity_output_path,
            agenda_items=agenda_items,
            run_date=run_date,
            requested_session_numbers=trigger_context.get("requested_session_numbers", []),
        )
        lesson_plan = apply_supports(lesson_plan, config)
        write_json(lesson_json_path, lesson_plan)

        validation_payload = build_validation_payload(
            lesson_plan=lesson_plan,
            raw_deck=raw_deck,
            config=config,
            output_file_status={
                "json": lesson_json_path.exists(),
                "markdown": False,
                "docx_count": 0,
                "validation_report": False,
            },
        )
        if not validation_payload["passed"]:
            write_validation_report(validation_payload, validation_report_path)
            raise LessonPlanError(
                "Validation failed before rendering. Review output/validation_report.md for details."
            )

        render_markdown(lesson_plan, template_md_path, lesson_md_path)
        rendered_docx_paths = render_docx(lesson_plan, template_docx_path, lesson_docx_path, config)

        validation_payload["output_file_status"] = {
            "json": lesson_json_path.exists(),
            "markdown": lesson_md_path.exists(),
            "docx_count": len(rendered_docx_paths),
            "validation_report": True,
        }
        validation_payload["passed"] = all(check["passed"] for check in validation_payload["checks"])
        write_validation_report(validation_payload, validation_report_path)

        print("Lesson plan generation complete.")
        print(f"Trigger date: {run_date}")
        print(f"Source deck: {deck_path}")
        print(f"JSON output: {lesson_json_path}")
        print(f"Markdown output: {lesson_md_path}")
        for path in rendered_docx_paths:
            print(f"DOCX output: {path}")
        print(f"Validation report: {validation_report_path}")
        return 0
    except LessonPlanError as exc:
        print(f"Lesson plan generation failed: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - defensive CLI guard
        print(f"Unexpected failure: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
