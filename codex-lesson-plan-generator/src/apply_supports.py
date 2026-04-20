from __future__ import annotations

from typing import Any

from utils import LessonPlanError, unique_preserve


APPROVED_SUPPORT_MAPPING: dict[str, list[str]] = {
    "Profile A": ["sentence starters", "word banks", "expressive language support"],
    "Profile B": ["reduced distractions", "manipulatives", "extended time"],
    "Profile C": ["breaks", "chunking", "graphic organizers", "word banks"],
    "Profile D": ["graphic organizers", "text-to-speech", "visuals", "repetition"],
    "Profile E": ["preteach vocabulary", "highlight tool", "attention strategies", "small group"],
    "Profile F": ["clarified/repeated directions", "text-to-speech", "visuals", "immediate feedback"],
    "Profile G": ["simplified language", "word banks", "graphic organizers", "calculator", "small group"],
}

PHASE_KEYS = [
    "opening_warm_up_launch",
    "mini_lesson_modeling_concept_development",
    "guided_practice_collaborative_learning",
    "independent_practice_application_stations",
    "closure_exit_ticket_assessment",
]


def apply_supports(lesson_plan: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    active_students = unique_preserve(config.get("active_student_supports", []))
    invalid_students = [student for student in active_students if student not in APPROVED_SUPPORT_MAPPING]
    if invalid_students:
        joined = ", ".join(invalid_students)
        raise LessonPlanError(f"Unsupported support profile ids in active_student_supports: {joined}")

    for session in lesson_plan.get("sessions", []):
        sped_supports = [
            {
                "student": student,
                "supports": APPROVED_SUPPORT_MAPPING[student],
            }
            for student in active_students
        ]
        esol_supports = build_esol_supports(session, config)
        session["differentiation_sped_esol_supports_and_teacher_notes"]["sped"] = sped_supports
        session["differentiation_sped_esol_supports_and_teacher_notes"]["esol"] = esol_supports
        session["differentiation_sped_esol_supports_and_teacher_notes"][
            "implementation_note"
        ] = "Embed these concise supports during the named lesson phases while keeping the source task intact."

        for phase_key in PHASE_KEYS:
            if phase_key not in session:
                continue
            session[phase_key]["embedded_supports"] = build_embedded_supports(session[phase_key], sped_supports, esol_supports)

    return lesson_plan


def build_esol_supports(session: dict[str, Any], config: dict[str, Any]) -> list[str]:
    if not config.get("enable_esol_supports", True):
        return []

    supports: list[str] = []
    standards_section = session.get("standards_and_learning_targets", {})
    support_section = session.get("differentiation_sped_esol_supports_and_teacher_notes", {})
    has_vocabulary = bool(support_section.get("teacher_notes"))
    if standards_section.get("learning_targets") or has_vocabulary:
        supports.append("Preview student-friendly vocabulary definitions and point to the matching visuals during modeling.")
    if session.get("opening_warm_up_launch", {}).get("focus_tasks") or session.get("guided_practice_collaborative_learning", {}).get("focus_tasks"):
        supports.append("Provide sentence starters for partner talk, discussion, and closure explanations.")
    if session.get("independent_practice_application_stations", {}).get("focus_tasks"):
        supports.append("Highlight key words in the task directions and repeat directions before independent work.")

    limit = max(int(config.get("esol_support_limit", 3)), 0)
    return unique_preserve(supports)[:limit]


def build_embedded_supports(
    phase: dict[str, Any],
    sped_supports: list[dict[str, Any]],
    esol_supports: list[str],
) -> list[str]:
    phase_title = phase.get("section_title", "").lower()
    embedded: list[str] = []

    flattened_supports = {support for item in sped_supports for support in item["supports"]}
    if {"sentence starters", "word banks", "expressive language support"} & flattened_supports:
        if any(keyword in phase_title for keyword in ("opening", "guided", "closure")):
            embedded.append("Offer sentence starters or a word bank while students explain their thinking.")
    if {"graphic organizers", "visuals", "manipulatives", "highlight tool"} & flattened_supports:
        if any(keyword in phase_title for keyword in ("mini-lesson", "guided", "independent")):
            embedded.append("Keep the visual model or graphic organizer visible while students work through the task.")
    if {"clarified/repeated directions", "text-to-speech", "immediate feedback", "chunking", "breaks"} & flattened_supports:
        if any(keyword in phase_title for keyword in ("independent", "closure", "opening")):
            embedded.append("Repeat or chunk directions as needed and check in quickly before students continue.")
    if {"small group", "reduced distractions", "extended time"} & flattened_supports:
        if any(keyword in phase_title for keyword in ("guided", "independent")):
            embedded.append("Use a brief small-group or low-distraction check-in for students who need extra processing time.")

    if esol_supports:
        if "partner talk" in " ".join(esol_supports).lower() and any(keyword in phase_title for keyword in ("opening", "guided", "closure")):
            embedded.append("Prompt students to rehearse an oral explanation with a partner before sharing aloud.")
        if "highlight key words" in " ".join(esol_supports).lower() and "independent" in phase_title:
            embedded.append("Point out the key words in the problem statement before independent work begins.")

    return unique_preserve(embedded)
