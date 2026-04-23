#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
from pathlib import Path
from typing import Any

WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
EXAMPLES_DIR = WORKSPACE_ROOT / "validation-tools" / "examples"

sys.path.insert(0, str(WORKSPACE_ROOT))

from notebook_engine import (
    build_rendered_quality_report,
    enforce_plan_requirements,
    extract_pptx_slide_texts,
    render_plan,
    template_role_signature,
)


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def discover_case_paths(selected_cases: list[str]) -> list[Path]:
    available_paths = sorted(EXAMPLES_DIR.glob("notebook*_case.json"))
    if not available_paths:
        raise SystemExit(f"No notebook benchmark cases were found in {EXAMPLES_DIR}")
    if not selected_cases:
        return available_paths

    by_stem = {path.stem: path for path in available_paths}
    by_name = {path.name: path for path in available_paths}
    matched: list[Path] = []
    missing: list[str] = []
    for name in selected_cases:
        path = by_stem.get(name) or by_name.get(name)
        if path is None:
            missing.append(name)
            continue
        matched.append(path)
    if missing:
        available = ", ".join(path.stem for path in available_paths)
        raise SystemExit(f"Unknown notebook benchmark case(s): {', '.join(missing)}. Available: {available}")
    return matched


def append_failure(failures: list[str], case_name: str, message: str) -> None:
    failures.append(f"{case_name}: {message}")


def resolve_workspace_path(path_text: str) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return WORKSPACE_ROOT / path


def normalize_slide_texts(slide_texts: list[str]) -> list[str]:
    return [re.sub(r"\s+", " ", text).strip() for text in slide_texts]


def check_numeric_expectation(
    *,
    actual: Any,
    expectation: dict[str, Any],
    label: str,
    failures: list[str],
    case_name: str,
) -> None:
    if not isinstance(actual, (int, float)):
        append_failure(failures, case_name, f"{label} was not numeric: {actual!r}")
        return
    if "equals" in expectation and actual != expectation["equals"]:
        append_failure(failures, case_name, f"{label} expected {expectation['equals']}, got {actual}")
    if "min" in expectation and actual < expectation["min"]:
        append_failure(failures, case_name, f"{label} expected at least {expectation['min']}, got {actual}")
    if "max" in expectation and actual > expectation["max"]:
        append_failure(failures, case_name, f"{label} expected at most {expectation['max']}, got {actual}")


def check_sequence_expectation(
    *,
    actual: Any,
    expectation: dict[str, Any],
    label: str,
    failures: list[str],
    case_name: str,
) -> None:
    if not isinstance(actual, list):
        append_failure(failures, case_name, f"{label} was not a list: {actual!r}")
        return
    if "equals" in expectation and actual != expectation["equals"]:
        append_failure(failures, case_name, f"{label} expected {expectation['equals']!r}, got {actual!r}")
    if "includes" in expectation:
        missing = [item for item in expectation["includes"] if item not in actual]
        if missing:
            append_failure(failures, case_name, f"{label} missing expected entries: {missing!r}")
    if "max_len" in expectation and len(actual) > expectation["max_len"]:
        append_failure(failures, case_name, f"{label} expected length <= {expectation['max_len']}, got {len(actual)}")
    if "min_len" in expectation and len(actual) < expectation["min_len"]:
        append_failure(failures, case_name, f"{label} expected length >= {expectation['min_len']}, got {len(actual)}")


def validate_plan(case_name: str, plan: dict[str, Any], expectations: dict[str, Any]) -> list[str]:
    failures: list[str] = []
    session_1 = plan.get("session_1")
    if not isinstance(session_1, dict):
        append_failure(failures, case_name, "plan did not include session_1")
        return failures

    expected_roles = expectations.get("template_roles")
    if expected_roles is not None:
        actual_roles = [list(item) for item in template_role_signature(session_1)]
        if actual_roles != expected_roles:
            append_failure(failures, case_name, f"template roles expected {expected_roles!r}, got {actual_roles!r}")

    slides = session_1.get("slides", [])
    if not isinstance(slides, list):
        append_failure(failures, case_name, "session_1.slides was not a list")
        return failures

    for slide_expectation in expectations.get("slides", []):
        index = slide_expectation["index"]
        if index >= len(slides):
            append_failure(failures, case_name, f"expected slide index {index}, but only found {len(slides)} slides")
            continue
        slide = slides[index]
        primary_text = str(slide.get("primary_text", "")).lower()
        for needle in slide_expectation.get("primary_text_includes", []):
            if needle.lower() not in primary_text:
                append_failure(failures, case_name, f"slide {index} primary_text missing {needle!r}")

        source_slide_numbers = list(slide.get("source_slide_numbers", []))
        forbidden_source_numbers = set(slide_expectation.get("source_slide_numbers_excludes", []))
        overlap = sorted(forbidden_source_numbers.intersection(source_slide_numbers))
        if overlap:
            append_failure(failures, case_name, f"slide {index} used forbidden source slide numbers {overlap!r}")

        required_source_numbers = set(slide_expectation.get("source_slide_numbers_includes", []))
        missing_source_numbers = sorted(required_source_numbers.difference(source_slide_numbers))
        if missing_source_numbers:
            append_failure(failures, case_name, f"slide {index} was missing source slide numbers {missing_source_numbers!r}")

        vocabulary_words = {
            str(item.get("word", "")).lower()
            for item in slide.get("vocabulary", [])
            if isinstance(item, dict)
        }
        missing_vocab = [word for word in slide_expectation.get("vocabulary_words_include", []) if word.lower() not in vocabulary_words]
        if missing_vocab:
            append_failure(failures, case_name, f"slide {index} vocabulary missing {missing_vocab!r}")
        forbidden_vocab = [word for word in slide_expectation.get("vocabulary_words_exclude", []) if word.lower() in vocabulary_words]
        if forbidden_vocab:
            append_failure(failures, case_name, f"slide {index} vocabulary included forbidden words {forbidden_vocab!r}")

    return failures


def validate_report(case_name: str, report: dict[str, Any], expectations: dict[str, Any]) -> list[str]:
    failures: list[str] = []
    if "passed" in expectations and bool(report.get("passed")) != bool(expectations["passed"]):
        append_failure(failures, case_name, f"report passed expected {expectations['passed']}, got {report.get('passed')}")

    sessions = report.get("sessions", {})
    if not isinstance(sessions, dict):
        append_failure(failures, case_name, "report.sessions was not a mapping")
        return failures

    required_sessions = expectations.get("required_sessions")
    if required_sessions is not None and set(sessions.keys()) != set(required_sessions):
        append_failure(failures, case_name, f"report sessions expected {required_sessions!r}, got {sorted(sessions.keys())!r}")

    for session_label, session_expectations in expectations.get("sessions", {}).items():
        session_report = sessions.get(session_label)
        if not isinstance(session_report, dict):
            append_failure(failures, case_name, f"report was missing session {session_label!r}")
            continue

        for metric_name in ("slide_count", "problem_slide_count", "activity_slide_count", "engagement_slide_count", "flagship_activity_count"):
            metric_expectation = session_expectations.get(metric_name)
            if metric_expectation is not None:
                check_numeric_expectation(
                    actual=session_report.get(metric_name),
                    expectation=metric_expectation,
                    label=f"{session_label}.{metric_name}",
                    failures=failures,
                    case_name=case_name,
                )

        for sequence_name in ("engagement_modes", "long_problem_stack_pages", "copyedit_flags", "issues"):
            sequence_expectation = session_expectations.get(sequence_name)
            if sequence_expectation is not None:
                check_sequence_expectation(
                    actual=session_report.get(sequence_name),
                    expectation=sequence_expectation,
                    label=f"{session_label}.{sequence_name}",
                    failures=failures,
                    case_name=case_name,
                )

        problem_checks_expectation = session_expectations.get("problem_checks")
        if problem_checks_expectation is not None:
            problem_checks = session_report.get("problem_checks")
            check_sequence_expectation(
                actual=problem_checks,
                expectation=problem_checks_expectation,
                label=f"{session_label}.problem_checks",
                failures=failures,
                case_name=case_name,
            )
            if isinstance(problem_checks, list) and problem_checks_expectation.get("all_full_source_problem_preserved"):
                missing_pages = [
                    check.get("page")
                    for check in problem_checks
                    if not isinstance(check, dict) or not check.get("full_source_problem_preserved", False)
                ]
                if missing_pages:
                    append_failure(
                        failures,
                        case_name,
                        f"{session_label}.problem_checks had unpreserved source problems on pages {missing_pages!r}",
                    )

    return failures


def validate_gold_standard(
    case_name: str,
    plan: dict[str, Any],
    report: dict[str, Any],
    outputs: dict[str, Path],
    expectations: dict[str, Any],
) -> list[str]:
    failures: list[str] = []

    plan_path_text = expectations.get("plan_path")
    if plan_path_text:
        expected_plan = load_json(resolve_workspace_path(plan_path_text))
        if plan != expected_plan:
            append_failure(failures, case_name, f"generated plan did not match gold standard plan at {plan_path_text}")

    report_path_text = expectations.get("report_path")
    if report_path_text:
        expected_report = load_json(resolve_workspace_path(report_path_text))
        if report != expected_report:
            append_failure(failures, case_name, f"generated quality report did not match gold standard report at {report_path_text}")

    session1_pptx_path_text = expectations.get("session1_pptx_path")
    if session1_pptx_path_text:
        generated_session1 = outputs.get("session1")
        if generated_session1 is None:
            append_failure(failures, case_name, "generated outputs did not include session1 for gold standard comparison")
        else:
            expected_slide_texts = normalize_slide_texts(
                extract_pptx_slide_texts(resolve_workspace_path(session1_pptx_path_text))
            )
            actual_slide_texts = normalize_slide_texts(extract_pptx_slide_texts(generated_session1))
            if actual_slide_texts != expected_slide_texts:
                append_failure(
                    failures,
                    case_name,
                    f"generated Session 1 PPTX text did not match gold standard PPTX at {session1_pptx_path_text}",
                )

    return failures


def run_case(case_path: Path) -> tuple[str, dict[str, Any]]:
    case = load_json(case_path)
    case_name = str(case.get("name") or case_path.stem)
    deck = case["deck"]
    plan_seed = case["plan_seed"]
    plan = enforce_plan_requirements(plan_seed, deck)
    plan_failures = validate_plan(case_name, plan, case.get("plan_expectations", {}))
    report_failures: list[str] = []
    gold_standard_failures: list[str] = []

    with tempfile.TemporaryDirectory(prefix=f"{case_name}_") as temp_dir:
        output_dir = Path(temp_dir)
        outputs = render_plan(plan, deck, output_dir)
        report = build_rendered_quality_report(plan, outputs)
        report_failures = validate_report(case_name, report, case.get("report_expectations", {}))
        gold_standard_failures = validate_gold_standard(case_name, plan, report, outputs, case.get("gold_standard", {}))

    failures = [*plan_failures, *report_failures, *gold_standard_failures]
    if failures:
        raise SystemExit("Notebook quality benchmark failed:\n- " + "\n- ".join(failures))
    return case_name, report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run deterministic notebook quality benchmark cases.")
    parser.add_argument(
        "--case",
        action="append",
        default=[],
        help="Case stem or filename from validation-tools/examples. May be passed more than once.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    case_paths = discover_case_paths(args.case)
    for case_path in case_paths:
        case_name, report = run_case(case_path)
        session = report["sessions"]["Session 1"]
        print(
            "PASS",
            case_name,
            f"(slides={session['slide_count']}, activities={session['activity_slide_count']}, engagement={session['engagement_slide_count']})",
        )
    print(f"Notebook quality benchmark passed for {len(case_paths)} case(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
