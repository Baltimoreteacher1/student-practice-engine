from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable

from utils import (
    LessonPlanError,
    clean_line,
    format_session_docx_name,
    join_slide_numbers,
    truncate,
    unique_preserve,
    write_json,
)


DEFAULT_PHASE_TIMING = {
    "opening": 10,
    "mini_lesson": 18,
    "guided_practice": 10,
    "independent_practice": 10,
    "closure": 7,
}

ASSESSMENT_PHASE_TIMING = {
    "opening": 5,
    "mini_lesson": 8,
    "guided_practice": 7,
    "independent_practice": 25,
    "closure": 10,
}

SESSION_SECTION_KEYS = [
    "lesson_information",
    "standards_and_learning_targets",
    "lesson_objective_and_student_success_criteria",
    "materials_and_preparation",
    "opening_warm_up_launch",
    "mini_lesson_modeling_concept_development",
    "guided_practice_collaborative_learning",
    "independent_practice_application_stations",
    "closure_exit_ticket_assessment",
    "differentiation_sped_esol_supports_and_teacher_notes",
]


def build_lesson_plan(
    lesson_extract: dict[str, Any],
    raw_deck: dict[str, Any],
    lesson_type: str,
    config: dict[str, Any],
    fidelity_output_path: Path,
    agenda_items: list[str] | None = None,
    run_date: str = "",
    requested_session_numbers: list[int] | None = None,
) -> dict[str, Any]:
    agenda_items = agenda_items or []
    requested_session_numbers = requested_session_numbers or []
    session_extracts = list(lesson_extract.get("sessions", []))
    if not session_extracts:
        raise LessonPlanError("LESSON_EXTRACT did not yield any usable sessions.")

    if requested_session_numbers:
        session_extracts = [
            session
            for session in session_extracts
            if int(session.get("session_number", 0)) in requested_session_numbers
        ]
        if not session_extracts:
            requested = ", ".join(str(number) for number in requested_session_numbers)
            raise LessonPlanError(f"Requested session(s) {requested} were not found in the source deck.")

    session_plans = [
        build_session_plan(
            session_extract=session_extract,
            lesson_extract=lesson_extract,
            lesson_type=lesson_type,
            config=config,
            agenda_items=agenda_items,
            run_date=run_date,
        )
        for session_extract in session_extracts
    ]
    configured_docx = clean_line(str((((config.get("output") or {})).get("docx")) or ""))
    if len(session_plans) == 1 and configured_docx:
        session_plans[0]["output_filename"] = Path(configured_docx).name

    appendix = build_source_fidelity_appendix(raw_deck, session_plans)
    write_json(fidelity_output_path, appendix)

    return {
        "date": run_date,
        "lesson_type": lesson_type,
        "source_deck": {
            "source_file": raw_deck.get("source_file", ""),
            "source_filename": raw_deck.get("source_filename", ""),
            "deck_title": lesson_extract.get("deck_title", ""),
            "lesson_topic": lesson_extract.get("lesson_topic", ""),
            "selected_session_numbers": [session["session_number"] for session in session_plans],
        },
        "sessions": session_plans,
        "appendix": appendix,
        "metadata": {
            "grade": str(config.get("default_grade", "")),
            "subject": str(config.get("default_subject", "")),
            "agenda_used": bool(agenda_items),
            "requested_session_numbers": requested_session_numbers,
            "generated_docx_filenames": [session["output_filename"] for session in session_plans],
        },
    }


def build_session_plan(
    session_extract: dict[str, Any],
    lesson_extract: dict[str, Any],
    lesson_type: str,
    config: dict[str, Any],
    agenda_items: list[str],
    run_date: str,
) -> dict[str, Any]:
    session_label = session_extract["session_label"]
    timing = resolve_phase_timing(session_extract, config)
    materials = infer_materials(session_extract, config, agenda_items)
    visuals = session_extract.get("required_visuals", [])
    lesson_title = clean_line(session_extract.get("session_title") or lesson_extract.get("deck_title", "Lesson Plan"))
    success_criteria = derive_success_criteria(session_extract)
    output_filename = format_session_docx_name(run_date, session_label)

    return {
        "session_label": session_label,
        "session_number": int(session_extract["session_number"]),
        "output_filename": output_filename,
        "lesson_information": {
            "date": run_date,
            "lesson_title": lesson_title,
            "session_number": int(session_extract["session_number"]),
            "session_label": session_label,
            "course_or_grade": build_course_grade_label(config),
            "estimated_duration_minutes": int(config.get("lesson_duration_minutes", 55)),
            "lesson_type": lesson_type,
            "source_slide_numbers": session_extract.get("source_slide_numbers", []),
        },
        "standards_and_learning_targets": {
            "standards": list(session_extract.get("standards", [])),
            "standards_status": session_extract.get("standards_status", ""),
            "standards_source": dict(session_extract.get("standards_source", {})),
            "learning_targets": list(session_extract.get("learning_targets", [])),
            "i_can_statements": [
                target
                for target in session_extract.get("learning_targets", [])
                if clean_line(target).lower().startswith("i can")
            ],
        },
        "lesson_objective_and_student_success_criteria": {
            "lesson_objective": build_lesson_objective(session_extract),
            "student_success_criteria": success_criteria,
        },
        "materials_and_preparation": {
            "materials": materials,
            "preparation_notes": build_preparation_notes(session_extract, agenda_items),
            "required_visuals": visuals,
        },
        "opening_warm_up_launch": build_phase_section(
            section_title="Opening / Warm-Up / Launch",
            time_minutes=timing["opening"],
            source_group=session_extract.get("opening_source", {}),
            focus_tasks=merge_lists(
                session_extract.get("opening_source", {}).get("lines", []),
                session_extract.get("be_curious_prompts", []),
                session_extract.get("mindset_prompts", []),
            ),
            teacher_actions=build_opening_teacher_actions(session_extract),
            student_actions=build_opening_student_actions(session_extract),
            evidence_of_learning=build_opening_evidence(session_extract),
            misconceptions_to_monitor=infer_misconceptions(session_extract),
        ),
        "mini_lesson_modeling_concept_development": build_phase_section(
            section_title="Mini-Lesson / Modeling / Concept Development",
            time_minutes=timing["mini_lesson"],
            source_group=session_extract.get("modeling_source", {}),
            focus_tasks=merge_lists(
                session_extract.get("modeling_source", {}).get("lines", []),
                session_extract.get("reasoning_tasks", []),
                [item["definition"] for item in session_extract.get("vocabulary_terms", [])],
            ),
            teacher_actions=build_modeling_teacher_actions(session_extract),
            student_actions=build_modeling_student_actions(session_extract),
            evidence_of_learning=build_modeling_evidence(session_extract),
            misconceptions_to_monitor=infer_misconceptions(session_extract),
        ),
        "guided_practice_collaborative_learning": build_phase_section(
            section_title="Guided Practice / Collaborative Learning",
            time_minutes=timing["guided_practice"],
            source_group=preferred_group(
                session_extract.get("guided_source", {}),
                session_extract.get("collaborative_source", {}),
            ),
            focus_tasks=merge_lists(
                session_extract.get("collaborative_tasks", []),
                session_extract.get("guided_practice", []),
                session_extract.get("checks_for_understanding", []),
            ),
            teacher_actions=build_guided_teacher_actions(session_extract),
            student_actions=build_guided_student_actions(session_extract),
            evidence_of_learning=build_guided_evidence(session_extract),
            misconceptions_to_monitor=infer_misconceptions(session_extract),
        ),
        "independent_practice_application_stations": build_phase_section(
            section_title="Independent Practice / Application / Stations",
            time_minutes=timing["independent_practice"],
            source_group=session_extract.get("independent_source", {}),
            focus_tasks=merge_lists(
                session_extract.get("independent_practice", []),
                session_extract.get("lets_explore_more_tasks", []),
                session_extract.get("apply_transfer_tasks", []),
                session_extract.get("reveal_math_workbook_references", []),
            ),
            teacher_actions=build_independent_teacher_actions(session_extract),
            student_actions=build_independent_student_actions(session_extract),
            evidence_of_learning=build_independent_evidence(session_extract),
            misconceptions_to_monitor=infer_misconceptions(session_extract),
        ),
        "closure_exit_ticket_assessment": build_phase_section(
            section_title="Closure / Exit Ticket / Assessment",
            time_minutes=timing["closure"],
            source_group=session_extract.get("closure_source", {}),
            focus_tasks=merge_lists(
                session_extract.get("summary_closure_language", []),
                session_extract.get("apply_transfer_tasks", []),
                session_extract.get("checks_for_understanding", []),
                session_extract.get("homework_follow_up_tasks", []),
            ),
            teacher_actions=build_closure_teacher_actions(session_extract),
            student_actions=build_closure_student_actions(session_extract),
            evidence_of_learning=build_closure_evidence(session_extract),
            misconceptions_to_monitor=infer_misconceptions(session_extract),
        ),
        "differentiation_sped_esol_supports_and_teacher_notes": {
            "implementation_note": "Embed supports during the source-aligned lesson flow without changing the core task.",
            "sped": [],
            "esol": [],
            "teacher_notes": build_teacher_notes(session_extract),
            "precision_monitoring": build_precision_monitoring(session_extract),
        },
    }


def resolve_phase_timing(session_extract: dict[str, Any], config: dict[str, Any]) -> dict[str, int]:
    full_text = " ".join(
        merge_lists(
            session_extract.get("opening_source", {}).get("lines", []),
            session_extract.get("summary_closure_language", []),
            session_extract.get("independent_practice", []),
        )
    )
    lowered = full_text.lower()
    timing = dict(DEFAULT_PHASE_TIMING)
    timing.update({key: int(value) for key, value in config.get("phase_timing", {}).items()})
    if any(marker in lowered for marker in ("quiz", "test", "assessment")):
        timing = dict(ASSESSMENT_PHASE_TIMING)
    expected = int(config.get("lesson_duration_minutes", 55))
    actual = sum(timing.values())
    if actual != expected:
        raise LessonPlanError(
            f"Phase timing totals {actual} minutes for {session_extract['session_label']}, expected {expected}."
        )
    return timing


def build_course_grade_label(config: dict[str, Any]) -> str:
    grade = clean_line(str(config.get("default_grade", "")))
    subject = clean_line(str(config.get("default_subject", "")))
    if grade and subject:
        return f"Grade {grade} {subject}"
    return grade or subject or "Course/grade not listed in current config"


def build_lesson_objective(session_extract: dict[str, Any]) -> str:
    learning_targets = session_extract.get("learning_targets", [])
    if learning_targets:
        objective = clean_line(learning_targets[0])
        for prefix in ("I can ", "We will "):
            if objective.startswith(prefix):
                objective = objective[len(prefix) :]
        return ensure_sentence(objective)

    reasoning_tasks = session_extract.get("reasoning_tasks", [])
    if reasoning_tasks:
        return ensure_sentence(reasoning_tasks[0].rstrip("?"))

    source_lines = session_extract.get("modeling_source", {}).get("lines", [])
    if source_lines:
        return ensure_sentence(source_lines[0].rstrip("?"))

    return "Use the source slides to meet the lesson goal stated in class."


def derive_success_criteria(session_extract: dict[str, Any]) -> list[str]:
    criteria = list(session_extract.get("success_criteria", []))
    if criteria:
        return criteria

    derived: list[str] = []
    reasoning_tasks = session_extract.get("reasoning_tasks", [])
    for task in reasoning_tasks:
        lowered = task.lower()
        if "base" in lowered or "height" in lowered:
            derived.append("Identify the base and height named in the source diagram.")
        if "decompose" in lowered or "compose" in lowered:
            derived.append("Explain how the source shape is decomposed or composed to support the area strategy.")
        if "formula" in lowered or "area" in lowered:
            derived.append("Use the source relationship or formula to solve and justify the problem.")
        if "missing" in lowered or "unknown dimension" in lowered:
            derived.append("Use the given area and one dimension to determine the missing dimension from the source task.")
    if not derived:
        source_lines = session_extract.get("modeling_source", {}).get("lines", [])
        derived = [
            f"Complete the source task: {truncate(source_lines[0], 120)}"
            for source_lines in [source_lines]
            if source_lines
        ]
    return unique_preserve(derived)[:4]


def infer_materials(session_extract: dict[str, Any], config: dict[str, Any], agenda_items: list[str]) -> list[str]:
    materials = list(config.get("materials_defaults", ["Teacher slide deck"]))
    source_blob = " ".join(
        merge_lists(
            session_extract.get("opening_source", {}).get("lines", []),
            session_extract.get("modeling_source", {}).get("lines", []),
            session_extract.get("guided_practice", []),
            session_extract.get("independent_practice", []),
            session_extract.get("required_visuals", []),
        )
    ).lower()
    keyword_to_material = {
        "graph paper": "Graph paper",
        "grid paper": "Graph paper or grid paper",
        "grid": "Projected grid or lined visual from the source deck",
        "ruler": "Ruler",
        "calculator": "Calculator",
        "tile": "Projected tile or shape diagram from the source deck",
        "workspace": "Reveal Math Workspace task from the source deck",
        "workbook": "Reveal Math workbook task referenced in the source slides",
    }
    for keyword, label in keyword_to_material.items():
        if keyword in source_blob:
            materials.append(label)
    if agenda_items:
        materials.append("Posted lesson agenda")
    return unique_preserve(materials)


def build_preparation_notes(session_extract: dict[str, Any], agenda_items: list[str]) -> list[str]:
    notes = [
        (
            f"Queue the slide deck to {session_extract['session_label']} "
            f"(slides {session_extract['slide_range']['start']}-{session_extract['slide_range']['end']})."
        ),
    ]
    if session_extract.get("required_visuals"):
        notes.append(f"Project the source visual(s) used in class: {truncate(session_extract['required_visuals'][0], 140)}")
    if session_extract.get("reveal_math_workbook_references"):
        notes.append("Have the Reveal Math Workspace or workbook task ready exactly as referenced in the source slides.")
    if agenda_items:
        notes.append(f"Keep the posted agenda visible: {truncate('; '.join(agenda_items), 140)}")
    return notes


def build_phase_section(
    section_title: str,
    time_minutes: int,
    source_group: dict[str, Any],
    focus_tasks: list[str],
    teacher_actions: list[str],
    student_actions: list[str],
    evidence_of_learning: list[str],
    misconceptions_to_monitor: list[str],
) -> dict[str, Any]:
    lines = source_group.get("lines", [])
    excerpt = source_group.get("source_excerpt") or truncate(" ".join(lines), 220)
    slide_numbers = source_group.get("slide_numbers", [])
    if not slide_numbers:
        raise LessonPlanError(f"Section '{section_title}' is missing grounded source slide numbers.")
    return {
        "section_title": section_title,
        "time_minutes": time_minutes,
        "focus_tasks": unique_preserve(focus_tasks)[:5],
        "teacher_actions": unique_preserve(teacher_actions)[:4],
        "student_actions": unique_preserve(student_actions)[:4],
        "evidence_of_learning": unique_preserve(evidence_of_learning)[:4],
        "misconceptions_to_monitor": unique_preserve(misconceptions_to_monitor)[:3],
        "embedded_supports": [],
        "source_slide_numbers": slide_numbers,
        "source_excerpt": excerpt,
    }


def build_opening_teacher_actions(session_extract: dict[str, Any]) -> list[str]:
    prompts = merge_lists(session_extract.get("be_curious_prompts", []), session_extract.get("opening_source", {}).get("lines", []))
    moves = []
    if prompts:
        moves.append(
            f"Launch the actual opening from slides {session_extract['opening_source']['slide_label']}: {truncate(prompts[0], 150)}"
        )
    if session_extract.get("be_curious_prompts"):
        moves.append("Collect notice-wonder thinking before naming the mathematical focus of the lesson.")
    if session_extract.get("mindset_prompts"):
        moves.append(f"Use the mindset prompt to frame discussion: {truncate(session_extract['mindset_prompts'][0], 140)}")
    return moves or ["Launch the source opener and connect student observations to the lesson focus."]


def build_opening_student_actions(session_extract: dict[str, Any]) -> list[str]:
    actions = [
        "Complete the opening task and record an initial strategy, estimate, or observation.",
        "Share one idea with a partner or the class using evidence from the source slide.",
    ]
    if session_extract.get("be_curious_prompts"):
        actions.append("Respond to the notice-wonder prompt before transitioning into the main task.")
    return actions


def build_opening_evidence(session_extract: dict[str, Any]) -> list[str]:
    evidence = [
        "Students refer to the actual figure, prompt, or question from the source deck when sharing their first idea.",
    ]
    if session_extract.get("mindset_prompts"):
        evidence.append("Students explain or compare ideas rather than giving only a quick answer.")
    return evidence


def build_modeling_teacher_actions(session_extract: dict[str, Any]) -> list[str]:
    lines = session_extract.get("modeling_source", {}).get("lines", [])
    reasoning = session_extract.get("reasoning_tasks", [])
    moves = []
    if lines:
        moves.append(f"Model the source sequence from slides {session_extract['modeling_source']['slide_label']}: {truncate(lines[0], 150)}")
    if reasoning:
        moves.append(f"Press for the exact mathematical relationship named in the deck: {truncate(reasoning[0], 150)}")
    if session_extract.get("vocabulary_terms"):
        vocab = session_extract["vocabulary_terms"][0]
        moves.append(f"Name and use the lesson vocabulary precisely: {vocab['term']} - {truncate(vocab['definition'], 110)}")
    return moves or ["Model the source strategy directly from the deck and make the reasoning visible."]


def build_modeling_student_actions(session_extract: dict[str, Any]) -> list[str]:
    actions = [
        "Track the modeled reasoning and annotate the diagram, relationship, or formula shown in the source slides.",
        "Explain how the source representation supports the strategy used in the lesson.",
    ]
    if session_extract.get("reasoning_tasks"):
        actions.append(f"Respond to the reasoning prompt: {truncate(session_extract['reasoning_tasks'][0], 140)}")
    return actions


def build_modeling_evidence(session_extract: dict[str, Any]) -> list[str]:
    evidence = [
        "Students use the same decomposition, composition, formula, or measurement language shown in the source deck.",
    ]
    if session_extract.get("vocabulary_terms"):
        evidence.append("Students use the key vocabulary accurately while explaining their work.")
    return evidence


def build_guided_teacher_actions(session_extract: dict[str, Any]) -> list[str]:
    lines = merge_lists(session_extract.get("collaborative_tasks", []), session_extract.get("guided_practice", []))
    moves = []
    if lines:
        moves.append(f"Coach the collaborative task from the source deck: {truncate(lines[0], 150)}")
    moves.append("Pause for checks for understanding before students move into more independent work.")
    if session_extract.get("checks_for_understanding"):
        moves.append(f"Use the source check question: {truncate(session_extract['checks_for_understanding'][0], 140)}")
    return moves


def build_guided_student_actions(session_extract: dict[str, Any]) -> list[str]:
    actions = [
        "Work with a partner or table group on the exact guided prompt from the slides.",
        "Compare strategies and revise work when feedback or a classmate's idea changes the approach.",
    ]
    return actions


def build_guided_evidence(session_extract: dict[str, Any]) -> list[str]:
    evidence = [
        "Students explain why the strategy works, not just what answer they got.",
    ]
    if session_extract.get("checks_for_understanding"):
        evidence.append("Students respond to the source discussion or check-for-understanding prompt using complete reasoning.")
    return evidence


def build_independent_teacher_actions(session_extract: dict[str, Any]) -> list[str]:
    lines = merge_lists(
        session_extract.get("independent_practice", []),
        session_extract.get("lets_explore_more_tasks", []),
        session_extract.get("apply_transfer_tasks", []),
    )
    moves = []
    if lines:
        moves.append(f"Assign the actual independent/application task: {truncate(lines[0], 150)}")
    if session_extract.get("reveal_math_workbook_references"):
        moves.append("Use the Reveal Math-only task exactly as referenced in the deck instead of inventing a new station rotation.")
    moves.append("Circulate and confer while monitoring whether students can transfer the modeled reasoning independently.")
    return moves


def build_independent_student_actions(session_extract: dict[str, Any]) -> list[str]:
    actions = [
        "Solve the source task independently and show all reasoning needed to justify the answer.",
        "Label dimensions, units, or key relationships exactly as the lesson expects.",
    ]
    if session_extract.get("apply_transfer_tasks"):
        actions.append("Complete the transfer/application problem if it appears in the session slides.")
    return actions


def build_independent_evidence(session_extract: dict[str, Any]) -> list[str]:
    evidence = [
        "Students set up the problem correctly and maintain precision with labels and units.",
    ]
    if session_extract.get("reveal_math_workbook_references"):
        evidence.append("Workbook or Workspace responses match the reasoning emphasized during modeling and guided practice.")
    return evidence


def build_closure_teacher_actions(session_extract: dict[str, Any]) -> list[str]:
    lines = merge_lists(
        session_extract.get("summary_closure_language", []),
        session_extract.get("apply_transfer_tasks", []),
        session_extract.get("checks_for_understanding", []),
    )
    moves = []
    if lines:
        moves.append(f"Use the session close from the source slides: {truncate(lines[0], 150)}")
    moves.append("Listen for precise mathematical language before dismissing the class or assigning follow-up work.")
    if session_extract.get("homework_follow_up_tasks"):
        moves.append(f"Assign the follow-up only if it appears in the deck: {truncate(session_extract['homework_follow_up_tasks'][0], 140)}")
    return moves


def build_closure_student_actions(session_extract: dict[str, Any]) -> list[str]:
    actions = [
        "Complete the closing check, summary, or exit-style prompt from the source deck.",
        "Explain the strategy or conclusion using the lesson vocabulary.",
    ]
    return actions


def build_closure_evidence(session_extract: dict[str, Any]) -> list[str]:
    evidence = [
        "Students independently show whether they can apply the session's mathematical focus without teacher prompting.",
    ]
    if session_extract.get("apply_transfer_tasks"):
        evidence.append("Students can extend the strategy to the transfer task when one is present in the source slides.")
    return evidence


def infer_misconceptions(session_extract: dict[str, Any]) -> list[str]:
    source_text = " ".join(
        merge_lists(
            session_extract.get("reasoning_tasks", []),
            session_extract.get("opening_source", {}).get("lines", []),
            session_extract.get("independent_practice", []),
        )
    ).lower()
    notes = []
    if "base" in source_text or "height" in source_text:
        notes.append("Watch for students choosing a slanted side instead of the perpendicular height.")
    if "decompose" in source_text or "compose" in source_text:
        notes.append("Watch for students changing the size of the figure instead of only rearranging pieces.")
    if "diagonal" in source_text or "perpendicular" in source_text:
        notes.append("Watch for confusion between side lengths and diagonals, and recheck perpendicular relationships from the source deck.")
    if "missing" in source_text or "unknown dimension" in source_text:
        notes.append("Watch for students multiplying immediately instead of using the known area to solve for the missing dimension.")
    if "formula" in source_text or "area" in source_text:
        notes.append("Watch for missing square units or a formula applied to the wrong dimensions.")
    explicit = [
        line
        for line in session_extract.get("checks_for_understanding", [])
        if "misconception" in line.lower()
    ]
    notes.extend(explicit)
    return unique_preserve(notes)[:3]


def build_teacher_notes(session_extract: dict[str, Any]) -> list[str]:
    notes = list(session_extract.get("speaker_note_teaching_moves", []))
    notes.extend(build_precision_monitoring(session_extract))
    if session_extract.get("checks_for_understanding"):
        notes.append(f"Monitor the source check prompt closely: {truncate(session_extract['checks_for_understanding'][0], 120)}")
    return unique_preserve(notes)[:6]


def build_precision_monitoring(session_extract: dict[str, Any]) -> list[str]:
    source_text = " ".join(session_extract.get("reasoning_tasks", [])).lower()
    notes = []
    if "base" in source_text or "height" in source_text:
        notes.append("Monitor whether students name the base and perpendicular height precisely.")
    if "decompose" in source_text or "compose" in source_text:
        notes.append("Listen for decomposition and composition language that matches the source slides.")
    if "diagonal" in source_text or "perpendicular" in source_text:
        notes.append("Revoice diagonal and perpendicular vocabulary exactly as it appears in the source lesson.")
    if session_extract.get("vocabulary_terms"):
        terms = ", ".join(item["term"] for item in session_extract["vocabulary_terms"][:3])
        notes.append(f"Keep these vocabulary terms visible during discussion: {terms}.")
    return unique_preserve(notes)[:4]


def build_source_fidelity_appendix(raw_deck: dict[str, Any], session_plans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    section_map = build_appendix_section_map(session_plans)
    appendix = []
    for slide in raw_deck.get("slides", []):
        used_in = [
            label
            for label, numbers in section_map.items()
            if slide.get("slide_number") in numbers
        ]
        appendix.append(
            {
                "slide_number": slide.get("slide_number"),
                "title": clean_line(slide.get("title", "")),
                "source_excerpt": truncate(slide.get("full_text", "") or "No visible text extracted.", 220),
                "used_in_sections": used_in,
                "mapping_note": "Directly mapped to the listed session section(s)." if used_in else "Retained for source fidelity context only.",
            }
        )
    return appendix


def build_appendix_section_map(session_plans: list[dict[str, Any]]) -> dict[str, list[int]]:
    mapping: dict[str, list[int]] = {}
    for session in session_plans:
        session_label = session["session_label"]
        for key in SESSION_SECTION_KEYS:
            section = session.get(key, {})
            if isinstance(section, dict):
                source_numbers = list(section.get("source_slide_numbers", []))
                if not source_numbers and "standards_source" in section:
                    source_numbers = list(section["standards_source"].get("slide_numbers", []))
                mapping[f"{session_label}: {key}"] = source_numbers
    return mapping


def preferred_group(primary: dict[str, Any], secondary: dict[str, Any]) -> dict[str, Any]:
    if primary.get("slide_numbers"):
        return primary
    return secondary


def merge_lists(*groups: Iterable[str]) -> list[str]:
    merged: list[str] = []
    for group in groups:
        merged.extend(str(item) for item in group if str(item).strip())
    return unique_preserve(merged)


def ensure_sentence(text: str) -> str:
    cleaned = clean_line(text)
    if not cleaned:
        return ""
    if cleaned.endswith((".", "!", "?")):
        return cleaned
    return f"{cleaned}."
