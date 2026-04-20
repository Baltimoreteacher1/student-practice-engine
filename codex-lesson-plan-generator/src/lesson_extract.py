from __future__ import annotations

import re
from typing import Any, Iterable

from utils import (
    LessonPlanError,
    clean_line,
    clean_source_line,
    join_slide_numbers,
    parse_standards,
    truncate,
    unique_preserve,
)


SESSION_LABEL_PATTERN = re.compile(r"^(session|day)\s*([12])$", re.IGNORECASE)
DEFINITION_PATTERN = re.compile(
    r"^(?:a|an|the)\s+([A-Za-z][A-Za-z \-/]+?)\s+(?:is|are)\s+(.+)$",
    re.IGNORECASE,
)
COLON_DEFINITION_PATTERN = re.compile(r"^([A-Za-z][A-Za-z \-/]{1,50})\s*:\s*(.+)$")

OPENING_MARKERS = (
    "do now",
    "this or that",
    "is it reasonable",
    "be curious",
    "notice",
    "wonder",
    "warm-up",
    "launch",
)
COLLABORATIVE_MARKERS = ("collaborate and connect", "turn and talk", "partner", "discussion", "connect")
GUIDED_MARKERS = ("guided practice", "we do", "practice together", "check for understanding")
INDEPENDENT_MARKERS = (
    "independent practice",
    "let’s explore more",
    "let's explore more",
    "workspace",
    "apply",
    "on your own",
    "station",
)
CLOSURE_MARKERS = ("summarize", "summary", "exit ticket", "closure", "apply", "final check")
HOMEWORK_MARKERS = ("homework", "follow-up", "next steps")
VOCABULARY_MARKERS = ("vocabulary", "term", "define", "diagonal", "perpendicular", "base", "height")
REASONING_MARKERS = (
    "decompose",
    "compose",
    "formula",
    "missing",
    "unknown dimension",
    "base",
    "height",
    "diagonal",
    "perpendicular",
    "justify",
    "explain",
)
VISUAL_MARKERS = (
    "grid",
    "figure",
    "diagram",
    "rectangle",
    "triangle",
    "parallelogram",
    "rhombus",
    "diagonal",
    "height",
    "base",
    "graph paper",
    "tile",
)
WORKBOOK_MARKERS = ("workspace", "workbook", "reveal", "page", "station", "let’s explore more", "let's explore more")


def run_lesson_extract(raw_deck: dict[str, Any], config: dict[str, Any]) -> dict[str, Any]:
    slides = raw_deck.get("slides", [])
    if not slides:
        raise LessonPlanError("The extracted deck contains no slides.")

    session_ranges = split_sessions(slides)
    learning_target_bundle = extract_learning_target_bundle(raw_deck)
    deck_title = infer_deck_title(slides, session_ranges, learning_target_bundle)
    lesson_topic = infer_lesson_topic(deck_title, learning_target_bundle)
    session_extracts = [
        build_session_extract(
            session_info=session_info,
            raw_deck=raw_deck,
            deck_title=deck_title,
            global_learning_target_bundle=learning_target_bundle,
        )
        for session_info in session_ranges
    ]

    return {
        "deck_title": deck_title,
        "lesson_topic": lesson_topic,
        "source_file": raw_deck.get("source_file", ""),
        "source_filename": raw_deck.get("source_filename", ""),
        "slide_count": raw_deck.get("slide_count", len(slides)),
        "learning_target_bundle": learning_target_bundle,
        "session_count": len(session_extracts),
        "sessions": session_extracts,
        "metadata": {
            "default_grade": str(config.get("default_grade", "")),
            "default_subject": str(config.get("default_subject", "")),
            "learning_target_candidate_numbers": raw_deck.get("learning_target_candidate_numbers", []),
        },
    }


def extract_learning_target_bundle(raw_deck: dict[str, Any]) -> dict[str, Any]:
    slides = raw_deck.get("slides", [])
    candidate_numbers = raw_deck.get("learning_target_candidate_numbers", [])
    candidate_slides = [slide for slide in slides if slide.get("slide_number") in candidate_numbers]
    standards_source_lines: list[str] = []
    standards: list[str] = []
    learning_targets: list[str] = []
    success_criteria: list[str] = []

    for slide in candidate_slides:
        lines = [clean_source_line(line) for line in slide.get("text_items", []) if clean_source_line(line)]
        capture_success_criteria = False
        for line in lines:
            lowered = line.lower()
            detected_standards = parse_standards(line)
            if detected_standards:
                standards_source_lines.append(line)
                standards.extend(detected_standards)
                continue
            if "success criteria" in lowered:
                capture_success_criteria = True
                continue
            if lowered.startswith(("i can", "we will")):
                if capture_success_criteria:
                    success_criteria.append(line)
                else:
                    learning_targets.append(line)
                continue
            if "learning target" in lowered and len(line.split()) > 2:
                learning_targets.append(line)
                continue
            if capture_success_criteria and len(line) > 5:
                success_criteria.append(line.lstrip("-• "))

    learning_targets = unique_preserve(learning_targets, cleaner=clean_source_line)
    success_criteria = unique_preserve(success_criteria, cleaner=clean_source_line)
    standards_source_lines = unique_preserve(standards_source_lines, cleaner=clean_source_line)
    standards = unique_preserve(standards)

    status = ""
    if candidate_slides and not standards:
        status = "Not explicitly listed in source slides."

    return {
        "slide_numbers": [slide.get("slide_number") for slide in candidate_slides],
        "standards": standards,
        "standards_source_lines": standards_source_lines,
        "standards_status": status,
        "learning_targets": learning_targets,
        "success_criteria": success_criteria,
    }


def split_sessions(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    markers: list[dict[str, Any]] = []
    for index, slide in enumerate(slides):
        label = detect_session_label(slide)
        if not label:
            continue
        markers.append(
            {
                "label": label["label"],
                "number": label["number"],
                "start_index": index,
            }
        )

    if not markers:
        return [
            {
                "label": "Session 1",
                "number": 1,
                "start_index": 0,
                "end_index": len(slides) - 1,
                "slides": list(slides),
            }
        ]

    sessions: list[dict[str, Any]] = []
    for index, marker in enumerate(markers):
        end_index = markers[index + 1]["start_index"] - 1 if index + 1 < len(markers) else len(slides) - 1
        sessions.append(
            {
                "label": marker["label"],
                "number": marker["number"],
                "start_index": marker["start_index"],
                "end_index": end_index,
                "slides": slides[marker["start_index"] : end_index + 1],
            }
        )
    return sessions


def detect_session_label(slide: dict[str, Any]) -> dict[str, Any] | None:
    candidates = [slide.get("title", "")] + list(slide.get("text_items", [])[:3])
    for candidate in candidates:
        match = SESSION_LABEL_PATTERN.match(clean_line(candidate))
        if not match:
            continue
        number = int(match.group(2))
        return {
            "label": f"Session {number}",
            "number": number,
        }
    return None


def infer_deck_title(
    slides: list[dict[str, Any]],
    session_ranges: list[dict[str, Any]],
    learning_target_bundle: dict[str, Any],
) -> str:
    for session_info in session_ranges:
        cover_title = infer_session_cover_title(session_info)
        if cover_title:
            return cover_title
    first_title = clean_line(slides[0].get("title", "")) if slides else ""
    if first_title and not SESSION_LABEL_PATTERN.match(first_title):
        return first_title
    learning_targets = learning_target_bundle.get("learning_targets", [])
    if learning_targets:
        return clean_line(learning_targets[0].replace("I can ", "").replace("We will ", ""))
    return "Lesson Plan"


def infer_session_cover_title(session_info: dict[str, Any]) -> str:
    slides = session_info.get("slides", [])
    if not slides:
        return ""
    cover_slide = slides[0]
    lines = [clean_line(line) for line in cover_slide.get("text_items", []) if clean_line(line)]
    session_label = session_info.get("label", "")
    title_lines = [line for line in lines if line.lower() != session_label.lower()]
    return clean_line(" ".join(title_lines[:3]))


def infer_lesson_topic(deck_title: str, learning_target_bundle: dict[str, Any]) -> str:
    if deck_title:
        return deck_title
    learning_targets = learning_target_bundle.get("learning_targets", [])
    if learning_targets:
        return clean_line(learning_targets[0].replace("I can ", ""))
    return "Lesson topic not explicitly listed in source slides."


def build_session_extract(
    session_info: dict[str, Any],
    raw_deck: dict[str, Any],
    deck_title: str,
    global_learning_target_bundle: dict[str, Any],
) -> dict[str, Any]:
    session_slides = session_info.get("slides", [])
    if not session_slides:
        raise LessonPlanError(f"{session_info['label']} does not contain any slides.")

    session_numbers = {slide.get("slide_number") for slide in session_slides}
    local_lt_numbers = [
        number
        for number in raw_deck.get("learning_target_candidate_numbers", [])
        if number in session_numbers
    ]
    if local_lt_numbers:
        local_bundle = extract_learning_target_bundle(
            {
                "slides": session_slides,
                "learning_target_candidate_numbers": local_lt_numbers,
            }
        )
    else:
        local_bundle = dict(global_learning_target_bundle)

    session_cover = infer_session_cover_title(session_info)
    session_title = session_cover or deck_title
    instructional_slides = trim_session_cover(session_slides)

    opening_slides = match_slides(instructional_slides, OPENING_MARKERS) or instructional_slides[: min(4, len(instructional_slides))]
    collaborative_slides = match_slides(instructional_slides, COLLABORATIVE_MARKERS)
    guided_slides = match_slides(instructional_slides, GUIDED_MARKERS) or collaborative_slides
    independent_slides = match_slides(instructional_slides, INDEPENDENT_MARKERS)
    closure_slides = match_slides(instructional_slides, CLOSURE_MARKERS)

    excluded_numbers = {
        *slide_numbers(opening_slides),
        *slide_numbers(collaborative_slides),
        *slide_numbers(guided_slides),
        *slide_numbers(independent_slides),
        *slide_numbers(closure_slides),
        *local_bundle.get("slide_numbers", []),
    }
    modeling_slides = [
        slide
        for slide in instructional_slides
        if slide.get("slide_number") not in excluded_numbers
    ]
    if not modeling_slides:
        modeling_slides = instructional_slides[min(len(opening_slides), len(instructional_slides)) :][:4]
    if not guided_slides:
        guided_slides = collaborative_slides or modeling_slides[-2:] or instructional_slides[-2:]
    if not independent_slides:
        independent_slides = instructional_slides[-2:] if len(instructional_slides) >= 2 else instructional_slides
    if not closure_slides:
        closure_slides = independent_slides[-1:] or instructional_slides[-1:]

    be_curious_prompts = extract_lines_with_keywords(opening_slides, ("notice", "wonder", "be curious"))
    mindset_prompts = extract_mindset_prompts(opening_slides)
    collaborative_tasks = summarize_group(collaborative_slides)
    lets_explore_more_tasks = extract_lines_with_keywords(independent_slides, ("let’s explore more", "let's explore more", "workspace"))
    summary_closure_language = extract_lines_with_keywords(closure_slides, ("summarize", "apply", "exit", "closure"))
    apply_transfer_tasks = extract_lines_with_keywords(instructional_slides, ("apply", "transfer"))
    homework_follow_up_tasks = extract_lines_with_keywords(instructional_slides, HOMEWORK_MARKERS)

    return {
        "session_label": session_info["label"],
        "session_number": session_info["number"],
        "session_title": session_title,
        "slide_range": {
            "start": session_slides[0].get("slide_number"),
            "end": session_slides[-1].get("slide_number"),
        },
        "source_slide_numbers": slide_numbers(session_slides),
        "learning_targets": local_bundle.get("learning_targets", []),
        "success_criteria": local_bundle.get("success_criteria", []),
        "standards": local_bundle.get("standards", []),
        "standards_status": local_bundle.get("standards_status", ""),
        "standards_source": {
            "slide_numbers": local_bundle.get("slide_numbers", []),
            "source_lines": local_bundle.get("standards_source_lines", []),
        },
        "opening_source": build_extract_group(opening_slides),
        "modeling_source": build_extract_group(modeling_slides),
        "collaborative_source": build_extract_group(collaborative_slides),
        "guided_source": build_extract_group(guided_slides),
        "independent_source": build_extract_group(independent_slides),
        "closure_source": build_extract_group(closure_slides),
        "be_curious_prompts": be_curious_prompts,
        "mindset_prompts": mindset_prompts,
        "vocabulary_terms": find_vocabulary_terms(instructional_slides),
        "worked_examples": extract_worked_examples(modeling_slides + guided_slides),
        "reveal_discussion_prompts": extract_worked_examples(collaborative_slides + closure_slides),
        "collaborative_tasks": collaborative_tasks["lines"],
        "guided_practice": summarize_group(guided_slides)["lines"],
        "independent_practice": summarize_group(independent_slides)["lines"],
        "lets_explore_more_tasks": lets_explore_more_tasks,
        "summary_closure_language": summary_closure_language,
        "apply_transfer_tasks": apply_transfer_tasks,
        "homework_follow_up_tasks": homework_follow_up_tasks,
        "required_visuals": find_required_visuals(instructional_slides),
        "speaker_note_teaching_moves": extract_speaker_note_moves(instructional_slides),
        "checks_for_understanding": find_checks_for_understanding(
            collaborative_slides + guided_slides + independent_slides + closure_slides
        ),
        "reasoning_tasks": find_reasoning_tasks(instructional_slides),
        "reveal_math_workbook_references": find_workbook_references(instructional_slides),
        "source_summary": (
            f"{session_info['label']} spans slides "
            f"{session_slides[0]['slide_number']}-{session_slides[-1]['slide_number']}."
        ),
    }


def trim_session_cover(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not slides:
        return []
    first_slide = slides[0]
    label = detect_session_label(first_slide)
    if label:
        return slides[1:]
    return slides


def match_slides(slides: list[dict[str, Any]], markers: Iterable[str]) -> list[dict[str, Any]]:
    return [slide for slide in slides if slide_matches(slide, markers)]


def slide_matches(slide: dict[str, Any], markers: Iterable[str]) -> bool:
    blob = build_slide_blob(slide)
    return any(marker in blob for marker in markers)


def build_slide_blob(slide: dict[str, Any]) -> str:
    parts = [slide.get("title", "")]
    parts.extend(slide.get("text_items", []))
    parts.extend(slide.get("speaker_notes", []))
    return clean_source_line(" ".join(parts)).lower()


def slide_numbers(slides: list[dict[str, Any]]) -> list[int]:
    return [int(slide.get("slide_number")) for slide in slides if slide.get("slide_number")]


def build_extract_group(slides: list[dict[str, Any]]) -> dict[str, Any]:
    lines = summarize_group(slides)["lines"]
    slide_number_list = slide_numbers(slides)
    return {
        "slide_numbers": slide_number_list,
        "lines": lines,
        "source_excerpt": truncate(" ".join(lines), 240) if lines else "",
        "slide_label": join_slide_numbers(slide_number_list),
    }


def summarize_group(slides: list[dict[str, Any]], limit: int = 6) -> dict[str, Any]:
    lines: list[str] = []
    for slide in slides:
        lines.extend(extract_content_lines(slide, include_notes=False))
    return {
        "slide_numbers": slide_numbers(slides),
        "lines": unique_preserve(lines, cleaner=clean_source_line)[:limit],
    }


def extract_content_lines(slide: dict[str, Any], *, include_notes: bool) -> list[str]:
    title = clean_line(slide.get("title", ""))
    lines = []
    for line in slide.get("text_items", []):
        cleaned = clean_source_line(line)
        if not cleaned or cleaned.lower() == title.lower():
            continue
        lines.append(cleaned)
    if include_notes:
        lines.extend(clean_source_line(line) for line in slide.get("speaker_notes", []))
    return [line for line in lines if line]


def extract_lines_with_keywords(slides: list[dict[str, Any]], keywords: Iterable[str], limit: int = 6) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in extract_content_lines(slide, include_notes=False):
            lowered = line.lower()
            if any(keyword in lowered for keyword in keywords):
                results.append(line)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]


def extract_mindset_prompts(slides: list[dict[str, Any]]) -> list[str]:
    prompts: list[str] = []
    for slide in slides:
        lines = [clean_source_line(line) for line in slide.get("text_items", []) if clean_source_line(line)]
        capture_next = False
        for line in lines:
            lowered = line.lower()
            if "mindset" in lowered:
                capture_next = True
                continue
            if capture_next and len(line) > 5:
                prompts.append(line)
                capture_next = False
    return unique_preserve(prompts, cleaner=clean_source_line)


def find_vocabulary_terms(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for slide in slides:
        for line in extract_content_lines(slide, include_notes=False):
            lowered = line.lower()
            if not any(marker in lowered for marker in VOCABULARY_MARKERS):
                continue
            term = ""
            definition = ""
            match = DEFINITION_PATTERN.match(line)
            if match:
                term = clean_line(match.group(1))
                definition = clean_source_line(match.group(2))
            else:
                colon_match = COLON_DEFINITION_PATTERN.match(line)
                if colon_match:
                    term = clean_line(colon_match.group(1))
                    definition = clean_source_line(colon_match.group(2))
            if not term or not definition:
                continue
            key = (term.lower(), definition.lower())
            if key in seen:
                continue
            seen.add(key)
            results.append(
                {
                    "term": term,
                    "definition": definition,
                    "slide_numbers": [slide.get("slide_number")],
                }
            )
    return results


def extract_worked_examples(slides: list[dict[str, Any]], limit: int = 5) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for slide in slides:
        lines = extract_content_lines(slide, include_notes=False)
        if not lines:
            continue
        questions = [line for line in lines if line.endswith("?")]
        reveals = [line for line in lines if line not in questions][:2]
        if not questions and not reveals:
            continue
        results.append(
            {
                "prompt": questions[0] if questions else lines[0],
                "reveal": reveals[0] if reveals else "",
                "slide_numbers": [slide.get("slide_number")],
            }
        )
        if len(results) >= limit:
            break
    return results


def find_required_visuals(slides: list[dict[str, Any]], limit: int = 8) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in extract_content_lines(slide, include_notes=False):
            lowered = line.lower()
            if any(marker in lowered for marker in VISUAL_MARKERS):
                results.append(line)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]


def extract_speaker_note_moves(slides: list[dict[str, Any]], limit: int = 6) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in slide.get("speaker_notes", []):
            cleaned = clean_source_line(line)
            if cleaned:
                results.append(cleaned)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]


def find_checks_for_understanding(slides: list[dict[str, Any]], limit: int = 6) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in extract_content_lines(slide, include_notes=False):
            lowered = line.lower()
            if line.endswith("?") or "check for understanding" in lowered or "summarize" in lowered:
                results.append(line)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]


def find_reasoning_tasks(slides: list[dict[str, Any]], limit: int = 8) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in extract_content_lines(slide, include_notes=False):
            lowered = line.lower()
            if any(marker in lowered for marker in REASONING_MARKERS):
                results.append(line)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]


def find_workbook_references(slides: list[dict[str, Any]], limit: int = 6) -> list[str]:
    results: list[str] = []
    for slide in slides:
        for line in slide.get("text_items", []):
            cleaned = clean_source_line(line)
            if any(marker in cleaned.lower() for marker in WORKBOOK_MARKERS):
                results.append(cleaned)
    return unique_preserve(results, cleaner=clean_source_line)[:limit]
