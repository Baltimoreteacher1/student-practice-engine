from __future__ import annotations

from pathlib import Path
from typing import Any

from apply_supports import APPROVED_SUPPORT_MAPPING
from build_lesson_plan import SESSION_SECTION_KEYS
from utils import LessonPlanError, parse_standards, validate_against_schema, write_text


ROOT = Path(__file__).resolve().parents[1]
LESSON_PLAN_SCHEMA_PATH = ROOT / "schemas" / "lesson_plan_schema.json"
QA_SCHEMA_PATH = ROOT / "schemas" / "qa_schema.json"

PHASE_KEYS = [
    "opening_warm_up_launch",
    "mini_lesson_modeling_concept_development",
    "guided_practice_collaborative_learning",
    "independent_practice_application_stations",
    "closure_exit_ticket_assessment",
]
FORBIDDEN_TOKENS = ("chatgpt", "ai-sounding", "placeholder", "tbd", "todo", "draft", "brainstorm")


def build_validation_payload(
    lesson_plan: dict[str, Any],
    raw_deck: dict[str, Any],
    config: dict[str, Any],
    output_file_status: dict[str, Any],
) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []

    schema_errors = validate_against_schema(lesson_plan, LESSON_PLAN_SCHEMA_PATH)
    checks.append(
        {
            "name": "lesson_plan_schema",
            "passed": not schema_errors,
            "details": "Lesson plan matched schema." if not schema_errors else "; ".join(schema_errors[:6]),
        }
    )

    missing_sections = find_missing_required_sections(lesson_plan)
    checks.append(
        {
            "name": "required_sections",
            "passed": not missing_sections,
            "details": "All required session sections are present."
            if not missing_sections
            else f"Missing sections: {', '.join(missing_sections)}",
        }
    )

    standards_ok, standards_details = validate_standards_source(lesson_plan, raw_deck)
    checks.append({"name": "standards_source_rule", "passed": standards_ok, "details": standards_details})

    source_grounding_ok, source_grounding_details = validate_source_grounding(lesson_plan)
    checks.append(
        {
            "name": "source_grounding",
            "passed": source_grounding_ok,
            "details": source_grounding_details,
        }
    )

    timing_ok, timing_payload = validate_timing(lesson_plan, config)
    checks.append(
        {
            "name": "timing_coherence",
            "passed": timing_ok,
            "details": timing_payload["details"],
        }
    )

    session_ok, session_details = validate_session_labels(lesson_plan, raw_deck)
    checks.append({"name": "session_labels", "passed": session_ok, "details": session_details})

    duplicate_ok, duplicate_details = validate_no_duplicate_sections(lesson_plan)
    checks.append({"name": "no_duplicate_sections", "passed": duplicate_ok, "details": duplicate_details})

    teacher_ready_ok, teacher_ready_details = validate_teacher_facing_output(lesson_plan)
    checks.append(
        {
            "name": "teacher_ready_language",
            "passed": teacher_ready_ok,
            "details": teacher_ready_details,
        }
    )

    supports_ok, supports_details = validate_supports(lesson_plan)
    checks.append({"name": "supports", "passed": supports_ok, "details": supports_details})

    appendix_ok, appendix_details = validate_appendix(lesson_plan, raw_deck)
    checks.append({"name": "source_fidelity_appendix", "passed": appendix_ok, "details": appendix_details})

    payload = {
        "passed": all(check["passed"] for check in checks),
        "checks": checks,
        "warnings": [],
        "missing_sections": missing_sections,
        "timing": timing_payload,
        "standards_source_validation": {
            "passed": standards_ok,
            "details": standards_details,
        },
        "output_file_status": output_file_status,
    }

    qa_errors = validate_against_schema(payload, QA_SCHEMA_PATH)
    if qa_errors:
        raise LessonPlanError("QA payload failed schema validation: " + "; ".join(qa_errors[:5]))
    return payload


def find_missing_required_sections(lesson_plan: dict[str, Any]) -> list[str]:
    missing: list[str] = []
    for session in lesson_plan.get("sessions", []):
        for key in SESSION_SECTION_KEYS:
            value = session.get(key)
            if value in (None, "", []):
                missing.append(f"{session['session_label']}: {key}")
    return missing


def validate_standards_source(lesson_plan: dict[str, Any], raw_deck: dict[str, Any]) -> tuple[bool, str]:
    candidate_numbers = set(raw_deck.get("learning_target_candidate_numbers", []))
    for session in lesson_plan.get("sessions", []):
        section = session.get("standards_and_learning_targets", {})
        source = section.get("standards_source", {})
        source_numbers = set(source.get("slide_numbers", []))
        source_lines = source.get("source_lines", [])
        standards = set(section.get("standards", []))
        status = section.get("standards_status", "")

        if source_numbers and not source_numbers.issubset(candidate_numbers):
            return False, f"{session['session_label']} uses standards source slides outside Learning Target extraction."
        if standards:
            if not source_lines:
                return False, f"{session['session_label']} lists standards but does not record source lines from a Learning Target slide."
            extracted = set(parse_standards(" ".join(source_lines)))
            if not standards.issubset(extracted):
                return False, f"{session['session_label']} includes a standard that was not found on the Learning Target slide."
        else:
            if status != "Not explicitly listed in source slides.":
                return False, f"{session['session_label']} must mark missing standards as not explicitly listed in source slides."
    return True, "Standards are limited to Learning Target slide extraction or explicitly marked as missing from the source slides."


def validate_source_grounding(lesson_plan: dict[str, Any]) -> tuple[bool, str]:
    for session in lesson_plan.get("sessions", []):
        for key in PHASE_KEYS:
            phase = session.get(key, {})
            if not phase.get("source_slide_numbers"):
                return False, f"{session['session_label']} is missing source slide numbers for {key}."
            if not phase.get("source_excerpt"):
                return False, f"{session['session_label']} is missing a source excerpt for {key}."
    return True, "All major lesson phases are grounded in source slide references."


def validate_timing(lesson_plan: dict[str, Any], config: dict[str, Any]) -> tuple[bool, dict[str, Any]]:
    expected = int(config.get("lesson_duration_minutes", 55))
    session_payloads = []
    passed = True
    for session in lesson_plan.get("sessions", []):
        total = 0
        missing_phase = False
        for key in PHASE_KEYS:
            if key not in session:
                missing_phase = True
                passed = False
                continue
            total += int(session[key]["time_minutes"])
        session_passed = (total == expected) and not missing_phase
        passed = passed and session_passed
        session_payloads.append(
            {
                "session_label": session["session_label"],
                "planned_total_minutes": total,
                "passed": session_passed,
            }
        )
    details = "; ".join(
        f"{item['session_label']}: {item['planned_total_minutes']} minutes"
        for item in session_payloads
    ) or "No session timing payloads were generated."
    if passed:
        details = f"Each session totals {expected} minutes. {details}"
    else:
        details = f"Expected {expected} minutes per session. {details}"
    return passed, {
        "expected_duration_minutes": expected,
        "planned_total_minutes": sum(item["planned_total_minutes"] for item in session_payloads),
        "passed": passed,
        "session_totals": session_payloads,
        "details": details,
    }


def validate_session_labels(lesson_plan: dict[str, Any], raw_deck: dict[str, Any]) -> tuple[bool, str]:
    extracted_labels = []
    for slide in raw_deck.get("slides", []):
        title = str(slide.get("title", "")).strip().lower()
        if title in {"session 1", "session 2"}:
            extracted_labels.append(title.title())
    if not extracted_labels:
        return True, "Source deck did not include explicit session labels."

    plan_labels = [session["session_label"] for session in lesson_plan.get("sessions", [])]
    if not set(plan_labels).issubset(set(extracted_labels)):
        return False, "Rendered session labels do not match the explicit session labels in the source deck."
    return True, "Rendered session labels match the source deck."


def validate_no_duplicate_sections(lesson_plan: dict[str, Any]) -> tuple[bool, str]:
    for session in lesson_plan.get("sessions", []):
        if len(set(SESSION_SECTION_KEYS)) != len(SESSION_SECTION_KEYS):
            return False, "Section key list contains duplicates."
        if len(session.keys() & set(SESSION_SECTION_KEYS)) != len(SESSION_SECTION_KEYS):
            return False, f"{session['session_label']} is missing one or more unique section keys."
    return True, "Each session contains the exact ten required section buckets without duplication."


def validate_teacher_facing_output(lesson_plan: dict[str, Any]) -> tuple[bool, str]:
    teacher_facing_payload = {"sessions": lesson_plan.get("sessions", [])}
    blob = serialize_payload_strings(teacher_facing_payload).lower()
    for token in FORBIDDEN_TOKENS:
        if token in blob:
            return False, f"Teacher-facing output contains forbidden token '{token}'."
    return True, "Teacher-facing output does not contain placeholders, AI/meta language, or draft framing."


def validate_supports(lesson_plan: dict[str, Any]) -> tuple[bool, str]:
    for session in lesson_plan.get("sessions", []):
        supports = session.get("differentiation_sped_esol_supports_and_teacher_notes", {})
        for item in supports.get("sped", []):
            approved = set(APPROVED_SUPPORT_MAPPING.get(item["student"], []))
            if not set(item["supports"]).issubset(approved):
                return False, f"{session['session_label']} includes an unapproved SPED support for {item['student']}."
        for esol_support in supports.get("esol", []):
            if len(esol_support) > 180:
                return False, f"{session['session_label']} contains an ESOL support that is too long to be classroom-usable."
    return True, "SPED supports use the approved mapping and ESOL supports stay concise."


def validate_appendix(lesson_plan: dict[str, Any], raw_deck: dict[str, Any]) -> tuple[bool, str]:
    appendix = lesson_plan.get("appendix", [])
    source_slides = raw_deck.get("slides", [])
    if len(appendix) != len(source_slides):
        return False, "Appendix does not include every slide from the source deck."
    appendix_numbers = {entry.get("slide_number") for entry in appendix}
    source_numbers = {slide.get("slide_number") for slide in source_slides}
    if appendix_numbers != source_numbers:
        return False, "Appendix slide numbers do not match the source deck."
    return True, "Appendix covers every source slide."


def serialize_payload_strings(payload: Any) -> str:
    if isinstance(payload, str):
        return payload
    if isinstance(payload, dict):
        return " ".join(serialize_payload_strings(value) for value in payload.values())
    if isinstance(payload, list):
        return " ".join(serialize_payload_strings(item) for item in payload)
    return str(payload)


def write_validation_report(payload: dict[str, Any], output_path: Path) -> None:
    status = "PASS" if payload["passed"] else "FAIL"
    missing_section_lines = [f"- {name}" for name in payload["missing_sections"]] or ["- None"]
    lines = [
        "# Validation Report",
        "",
        f"**Overall Status:** {status}",
        "",
        "## Checks",
    ]
    for check in payload["checks"]:
        marker = "PASS" if check["passed"] else "FAIL"
        lines.append(f"- **{marker}** `{check['name']}`: {check['details']}")

    lines.extend(
        [
            "",
            "## Missing Sections",
            *missing_section_lines,
            "",
            "## Timing",
            f"- Expected duration per session: {payload['timing']['expected_duration_minutes']} minutes",
            f"- Combined planned total: {payload['timing']['planned_total_minutes']} minutes",
            f"- Passed: {payload['timing']['passed']}",
        ]
    )
    for item in payload["timing"].get("session_totals", []):
        lines.append(
            f"- {item['session_label']}: {item['planned_total_minutes']} minutes (passed: {item['passed']})"
        )
    lines.extend(
        [
            "",
            "## Standards Source Validation",
            f"- Passed: {payload['standards_source_validation']['passed']}",
            f"- Details: {payload['standards_source_validation']['details']}",
            "",
            "## Output File Status",
            f"- JSON: {payload['output_file_status']['json']}",
            f"- Markdown: {payload['output_file_status']['markdown']}",
            f"- DOCX count: {payload['output_file_status']['docx_count']}",
            f"- Validation report: {payload['output_file_status']['validation_report']}",
        ]
    )
    write_text(output_path, "\n".join(lines) + "\n")
