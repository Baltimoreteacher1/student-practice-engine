from __future__ import annotations

import argparse
import json
import re
import shutil
from collections import Counter
from copy import deepcopy
from pathlib import Path
from typing import Any


QUALITY_CATEGORY_KEYS = (
    "sourceFidelity",
    "lessonAdaptation",
    "activityQuality",
    "supportIntegration",
    "visualHierarchy",
    "layoutDiscipline",
    "typographyReadability",
    "editability",
    "toneWording",
    "benchmarkFinish",
)

GENERIC_FILLER_MARKERS = (
    "complete the activity",
    "do the work",
    "use the space below",
    "show what you know",
    "practice page",
    "fun activity",
    "student activity",
    "example a",
    "example b",
    "card 1",
    "card 2",
    "category a",
    "category b",
    "sort the cards",
    "move the pieces",
    "use the pieces",
)

PLACEHOLDER_MARKERS = (
    "placeholder",
    "tbd",
    "todo",
    "[insert",
    "lorem ipsum",
    "sample text",
    "your lesson here",
)

SUPPORT_REQUIRED_KINDS = {
    "practice",
    "worked_example",
    "challenge",
    "exit_ticket",
    "guided_notes",
    "independent_practice",
    "collaborative_practice",
}

GENERIC_TITLES = {
    "practice",
    "independent practice",
    "collaborative practice",
    "reflection",
    "exit ticket",
    "challenge",
}


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_key(value: Any) -> str:
    return normalize_text(value).lower()


def truncate_text(text: str, limit: int) -> str:
    text = normalize_text(text)
    if len(text) <= limit:
        return text
    trimmed = text[: limit - 3].rstrip(" ,;:")
    return trimmed + "..."


def category_payload(status: str, score: int, notes: list[str] | None = None) -> dict[str, Any]:
    return {
        "status": status,
        "score": score,
        "notes": notes or [],
    }


def load_notebook_package(bundle_dir: Path) -> dict[str, Any]:
    plan_path = bundle_dir / "notebook_plan.json"
    if not plan_path.exists():
        raise RuntimeError(f"Notebook enhancement requires notebook_plan.json in {bundle_dir}")
    package: dict[str, Any] = {
        "bundleDir": str(bundle_dir.resolve()),
        "plan": read_json(plan_path),
        "planPath": str(plan_path.resolve()),
    }
    deck_path = bundle_dir / "source_deck.json"
    if deck_path.exists():
        package["deck"] = read_json(deck_path)
        package["deckPath"] = str(deck_path.resolve())
    quality_report_path = bundle_dir / "quality_report.json"
    if quality_report_path.exists():
        package["qualityReport"] = read_json(quality_report_path)
        package["qualityReportPath"] = str(quality_report_path.resolve())
    pptx_paths = sorted(bundle_dir.glob("*.pptx"))
    if pptx_paths:
        package["pptxPaths"] = [str(path.resolve()) for path in pptx_paths]
    return package


def iter_session_slides(plan: dict[str, Any]) -> list[tuple[str, int, dict[str, Any]]]:
    rows: list[tuple[str, int, dict[str, Any]]] = []
    for session_key in ("session_1", "session_2"):
        session = plan.get(session_key, {})
        slides = session.get("slides", []) if isinstance(session, dict) else []
        for index, slide in enumerate(slides):
            if isinstance(slide, dict):
                rows.append((session_key, index, slide))
    return rows


def slide_text_parts(slide: dict[str, Any]) -> list[str]:
    parts: list[str] = []
    for field_name in (
        "section",
        "title",
        "subtitle",
        "primary_text",
        "secondary_text",
        "response_prompt",
        "activity_name",
        "activity_instructions",
        "answer_check",
        "image_caption",
    ):
        value = normalize_text(slide.get(field_name, ""))
        if value:
            parts.append(value)
    for list_name in ("bullets", "tasks", "sentence_starters", "movable_pieces"):
        for item in slide.get(list_name, []):
            text = normalize_text(item)
            if text:
                parts.append(text)
    for vocab in slide.get("vocabulary", []):
        if isinstance(vocab, dict):
            for value in vocab.values():
                text = normalize_text(value)
                if text:
                    parts.append(text)
        else:
            text = normalize_text(vocab)
            if text:
                parts.append(text)
    return parts


def slide_text_blob(slide: dict[str, Any]) -> str:
    return normalize_text(" ".join(slide_text_parts(slide)))


def source_records_for_slide(deck: dict[str, Any] | None, slide: dict[str, Any]) -> list[dict[str, Any]]:
    if not isinstance(deck, dict):
        return []
    source_numbers = []
    for raw in slide.get("source_slide_numbers", []) or []:
        if isinstance(raw, int):
            source_numbers.append(raw)
    if not source_numbers:
        for raw in slide.get("sourceAnchors", []) or []:
            if isinstance(raw, int):
                source_numbers.append(raw)
    if not source_numbers:
        return []
    records: list[dict[str, Any]] = []
    for record in deck.get("slides", []):
        if not isinstance(record, dict):
            continue
        if record.get("slide_number") in source_numbers:
            records.append(record)
    return records


def source_text_for_slide(deck: dict[str, Any] | None, slide: dict[str, Any]) -> str:
    parts: list[str] = []
    for record in source_records_for_slide(deck, slide):
        for field_name in ("title", "text", "notes"):
            text = normalize_text(record.get(field_name, ""))
            if text:
                parts.append(text)
        for list_name in ("text_items", "problem_texts"):
            for item in record.get(list_name, []):
                text = normalize_text(item)
                if text:
                    parts.append(text)
    return normalize_text(" ".join(parts))


def lesson_focus_phrase(source_text: str, fallback: str = "the lesson idea") -> str:
    candidates = re.split(r"[.?!;:]", source_text)
    for candidate in candidates:
        cleaned = normalize_text(candidate)
        if len(cleaned.split()) >= 3:
            words = cleaned.split()[:7]
            return " ".join(words)
    return fallback


def has_placeholder_wording(text: str) -> bool:
    lowered = normalize_key(text)
    return any(marker in lowered for marker in PLACEHOLDER_MARKERS)


def has_generic_filler(text: str) -> bool:
    lowered = normalize_key(text)
    return any(re.search(rf"(?<![a-z0-9]){re.escape(marker)}(?![a-z0-9])", lowered) for marker in GENERIC_FILLER_MARKERS)


def slide_has_supports(slide: dict[str, Any]) -> bool:
    return bool(slide.get("sentence_starters") or slide.get("vocabulary"))


def slide_is_dense(slide: dict[str, Any]) -> bool:
    text_length = len(slide_text_blob(slide))
    return (
        text_length > 420
        or len(slide.get("bullets", [])) > 4
        or len(slide.get("tasks", [])) > 4
        or len(slide.get("sentence_starters", [])) > 4
    )


def slide_is_bland(slide: dict[str, Any], source_text: str) -> bool:
    title = normalize_key(slide.get("title", ""))
    has_structure = bool(slide.get("tasks") or slide.get("response_prompt") or slide.get("activity_name"))
    has_support = slide_has_supports(slide)
    source_anchor = bool(source_text)
    return bool(
        source_anchor
        and title in GENERIC_TITLES
        and not has_structure
        and not has_support
    )


def slide_text_length(slide: dict[str, Any]) -> int:
    return len(slide_text_blob(slide))


def slide_is_sparse(slide: dict[str, Any]) -> bool:
    text_length = slide_text_length(slide)
    return bool(
        text_length < 160
        or (text_length < 220 and not slide.get("tasks") and not slide_has_supports(slide))
        or (text_length < 260 and not slide.get("bullets") and not slide.get("response_prompt"))
    )


def token_set(text: str) -> set[str]:
    return {
        token
        for token in re.findall(r"[a-z0-9]+", normalize_key(text))
        if len(token) > 3
    }


def source_overlap_count(slide: dict[str, Any], source_text: str) -> int:
    if not source_text:
        return 0
    return len(token_set(slide_text_blob(slide)) & token_set(source_text))


def slide_is_soft_bland(slide: dict[str, Any], source_text: str) -> bool:
    return bool(source_text and (slide_is_bland(slide, source_text) or slide_is_sparse(slide) or source_overlap_count(slide, source_text) < 2))


def repeated_title_count(plan: dict[str, Any]) -> int:
    titles = [normalize_key(slide.get("title", "")) for _session, _index, slide in iter_session_slides(plan)]
    counts = Counter(title for title in titles if title)
    return sum(count - 1 for count in counts.values() if count > 1)


def collect_plan_vocabulary(plan: dict[str, Any]) -> list[str]:
    words: list[str] = []
    for _session, _index, slide in iter_session_slides(plan):
        for vocab in slide.get("vocabulary", []):
            if isinstance(vocab, dict):
                word = normalize_text(vocab.get("word", ""))
            else:
                word = normalize_text(vocab)
            if word:
                words.append(word)
    seen: set[str] = set()
    ordered: list[str] = []
    for word in words:
        key = normalize_key(word)
        if key not in seen:
            seen.add(key)
            ordered.append(word)
    return ordered


def derive_source_prompt(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> str:
    source_text = source_text_for_slide(deck, slide)
    focus = lesson_focus_phrase(source_text, fallback=normalize_text(slide.get("title", "")) or "the lesson idea")
    if vocabulary_bank:
        return f"Use {vocabulary_bank[0]} to explain how {focus.lower()}."
    return f"Use the source model, numbers, or representation to explain how {focus.lower()}."


def derive_sentence_starters(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> list[str]:
    source_text = source_text_for_slide(deck, slide)
    focus = lesson_focus_phrase(source_text, fallback=normalize_text(slide.get("title", "")) or "the lesson idea")
    starters = [
        f"I used {focus.lower()} to notice ___.",
        "My evidence is ___.",
    ]
    if vocabulary_bank:
        starters.insert(0, f"One important word is {vocabulary_bank[0]} because ___.")
    return starters[:3]


def derive_task_list(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> list[str]:
    source_text = source_text_for_slide(deck, slide)
    focus = lesson_focus_phrase(source_text, fallback=normalize_text(slide.get("title", "")) or "the lesson idea")
    tasks = [
        f"Work with {focus.lower()} from the source lesson.",
        "Record your reasoning in the workspace.",
        "Justify one choice or conclusion with evidence.",
    ]
    if vocabulary_bank:
        tasks[1] = f"Use {vocabulary_bank[0]} in your explanation."
    return tasks


def tighten_list(items: list[Any], limit: int) -> list[str]:
    tightened: list[str] = []
    for item in items[:limit]:
        text = truncate_text(normalize_text(item), 110)
        if text:
            tightened.append(text)
    return tightened


def score_notebook_quality_by_category(notebook_package: dict[str, Any]) -> dict[str, Any]:
    plan = notebook_package["plan"]
    deck = notebook_package.get("deck")
    slides = iter_session_slides(plan)
    total = max(len(slides), 1)
    generic_count = 0
    placeholder_count = 0
    dense_count = 0
    bland_count = 0
    sparse_count = 0
    source_mismatch_count = 0
    untethered_count = 0
    support_gap_count = 0
    weak_direction_count = 0
    flat_hierarchy_count = 0
    for _session, _index, slide in slides:
        source_text = source_text_for_slide(deck, slide)
        text_blob = slide_text_blob(slide)
        if has_placeholder_wording(text_blob):
            placeholder_count += 1
        if has_generic_filler(text_blob):
            generic_count += 1
        if slide_is_dense(slide):
            dense_count += 1
        if slide_is_bland(slide, source_text):
            bland_count += 1
        if slide_is_sparse(slide):
            sparse_count += 1
        if source_text and source_overlap_count(slide, source_text) < 2 and not has_generic_filler(text_blob) and not has_placeholder_wording(text_blob):
            source_mismatch_count += 1
        if slide.get("kind") in SUPPORT_REQUIRED_KINDS and not slide_has_supports(slide):
            support_gap_count += 1
        if slide.get("kind") in SUPPORT_REQUIRED_KINDS and not source_text:
            untethered_count += 1
        if has_generic_filler(normalize_text(slide.get("activity_instructions", ""))) or has_generic_filler(normalize_text(slide.get("response_prompt", ""))):
            weak_direction_count += 1
        if not normalize_text(slide.get("title", "")) or not (
            normalize_text(slide.get("subtitle", "")) or normalize_text(slide.get("primary_text", ""))
        ):
            flat_hierarchy_count += 1

    repeated_titles = repeated_title_count(plan)
    categories = {
        "sourceFidelity": category_payload(
            "fail" if untethered_count or source_mismatch_count > max(1, total // 2) else "weak" if source_mismatch_count or repeated_titles or bland_count else "pass",
            0 if untethered_count or source_mismatch_count > max(1, total // 2) else 1 if source_mismatch_count or repeated_titles or bland_count else 2,
            []
            if not (untethered_count or source_mismatch_count)
            else [f"{untethered_count + source_mismatch_count} slide(s) need tighter source anchoring."],
        ),
        "lessonAdaptation": category_payload(
            "fail" if generic_count or placeholder_count else "weak" if bland_count or sparse_count else "pass",
            0 if generic_count or placeholder_count else 1 if bland_count or sparse_count else 2,
            []
            if not (generic_count or bland_count or sparse_count)
            else [f"{generic_count + bland_count + sparse_count} slide(s) still feel too generic or thin."],
        ),
        "activityQuality": category_payload(
            "fail" if generic_count or untethered_count or placeholder_count else "weak" if bland_count or sparse_count else "pass",
            0 if generic_count or untethered_count or placeholder_count else 1 if bland_count or sparse_count else 2,
            []
            if not (bland_count or sparse_count)
            else [f"{bland_count + sparse_count} slide(s) are technically correct but still too flat or underbuilt."],
        ),
        "supportIntegration": category_payload(
            "fail" if support_gap_count > max(1, total // 2) else "weak" if support_gap_count or sparse_count else "pass",
            0 if support_gap_count > max(1, total // 2) else 1 if support_gap_count or sparse_count else 2,
            [] if not (support_gap_count or sparse_count) else [f"{support_gap_count + sparse_count} slide(s) need integrated supports or vocabulary cues."],
        ),
        "visualHierarchy": category_payload(
            "fail" if flat_hierarchy_count > max(1, total // 2) else "weak" if flat_hierarchy_count or bland_count or sparse_count else "pass",
            0 if flat_hierarchy_count > max(1, total // 2) else 1 if flat_hierarchy_count or bland_count or sparse_count else 2,
            [] if not (flat_hierarchy_count or bland_count or sparse_count) else [f"{flat_hierarchy_count + bland_count + sparse_count} slide(s) need clearer focal hierarchy."],
        ),
        "layoutDiscipline": category_payload(
            "fail" if dense_count > max(1, total // 2) else "weak" if dense_count or repeated_titles else "pass",
            0 if dense_count > max(1, total // 2) else 1 if dense_count or repeated_titles else 2,
            [] if not dense_count else [f"{dense_count} slide(s) are too dense for clean notebook layout."],
        ),
        "typographyReadability": category_payload(
            "fail" if dense_count > max(1, total // 2) else "weak" if dense_count or sparse_count else "pass",
            0 if dense_count > max(1, total // 2) else 1 if dense_count or sparse_count else 2,
            [] if not (dense_count or sparse_count) else [f"{dense_count + sparse_count} slide(s) need lighter text load before rendering."],
        ),
        "editability": category_payload(
            "pass",
            2,
            ["Enhancement stays plan-first and leaves final rendering to the editable notebook engine."],
        ),
        "toneWording": category_payload(
            "fail" if placeholder_count or weak_direction_count or generic_count else "weak" if bland_count else "pass",
            0 if placeholder_count or weak_direction_count or generic_count else 1 if bland_count else 2,
            [] if not (placeholder_count or weak_direction_count or generic_count) else [f"{placeholder_count + weak_direction_count + generic_count} wording issue(s) need cleanup."],
        ),
        "benchmarkFinish": category_payload(
            "fail" if generic_count or placeholder_count or untethered_count else "weak" if bland_count or dense_count or repeated_titles or sparse_count else "pass",
            0 if generic_count or placeholder_count or untethered_count else 1 if bland_count or dense_count or repeated_titles or sparse_count else 2,
            []
            if not (bland_count or repeated_titles or sparse_count)
            else [f"{bland_count or repeated_titles or sparse_count} slide-level premium-finish gap(s) detected."],
        ),
    }
    return categories


def detect_notebook_hard_fails(notebook_package: dict[str, Any], categories: dict[str, Any] | None = None) -> list[str]:
    plan = notebook_package["plan"]
    deck = notebook_package.get("deck")
    hard_fails: list[str] = []
    slides = iter_session_slides(plan)
    for session_key, index, slide in slides:
        text_blob = slide_text_blob(slide)
        source_text = source_text_for_slide(deck, slide)
        label = f"{session_key} slide {index + 1}"
        if has_placeholder_wording(text_blob):
            hard_fails.append(f"{label}: placeholder wording remains")
        if has_generic_filler(text_blob):
            hard_fails.append(f"{label}: generic filler text remains")
        if slide_is_dense(slide):
            hard_fails.append(f"{label}: text density is too high for premium readability")
        if slide.get("kind") in SUPPORT_REQUIRED_KINDS and not source_text:
            hard_fails.append(f"{label}: activity or practice page is not clearly tied to source lesson content")
    if categories:
        for key, payload in categories.items():
            if payload.get("status") == "fail":
                hard_fails.append(f"{key}: premium quality category failed")
    deduped: list[str] = []
    seen: set[str] = set()
    for item in hard_fails:
        if item not in seen:
            seen.add(item)
            deduped.append(item)
    return deduped


def repair_slide_title(slide: dict[str, Any], deck: dict[str, Any] | None) -> str | None:
    current_text = normalize_text(slide.get("title", ""))
    current = normalize_key(current_text)
    needs_repair = not current or current in GENERIC_TITLES or has_placeholder_wording(current_text)
    if not needs_repair:
        return None
    fallback_title = (
        normalize_text(slide.get("section", ""))
        or normalize_text(str(slide.get("kind", "")).replace("_", " ")).title()
        or "Notebook Page"
    )
    source_text = source_text_for_slide(deck, slide)
    focus_source = source_text or slide_text_blob(slide)
    focus = lesson_focus_phrase(focus_source, fallback=fallback_title)
    base_title = current_text if current_text and current not in GENERIC_TITLES and not has_placeholder_wording(current_text) else fallback_title
    if normalize_key(base_title) == normalize_key(focus):
        new_title = truncate_text(base_title, 52)
    else:
        new_title = truncate_text(f"{base_title}: {focus}", 52)
    slide["title"] = new_title
    return f"retitled slide to anchor it to source content ({new_title})"


def repair_slide_supports(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> str | None:
    if slide.get("kind") not in SUPPORT_REQUIRED_KINDS or slide_has_supports(slide):
        return None
    starters = derive_sentence_starters(slide, deck, vocabulary_bank)
    if not starters:
        return None
    slide["sentence_starters"] = starters
    return "added integrated sentence starters for access support"


def repair_slide_copy(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> list[str]:
    repairs: list[str] = []
    if has_generic_filler(normalize_text(slide.get("activity_instructions", ""))):
        slide["activity_instructions"] = derive_source_prompt(slide, deck, vocabulary_bank)
        repairs.append("replaced generic activity directions with source-derived directions")
    if has_generic_filler(normalize_text(slide.get("response_prompt", ""))) or not normalize_text(slide.get("response_prompt", "")):
        slide["response_prompt"] = derive_source_prompt(slide, deck, vocabulary_bank)
        repairs.append("replaced weak response prompt with source-derived prompt")
    if not slide.get("tasks"):
        slide["tasks"] = derive_task_list(slide, deck, vocabulary_bank)
        repairs.append("added structured tasks so the page feels authored instead of flat")
    return repairs


def repair_slide_premium_finish(slide: dict[str, Any], deck: dict[str, Any] | None, vocabulary_bank: list[str]) -> list[str]:
    repairs: list[str] = []
    source_text = source_text_for_slide(deck, slide)
    if not source_text:
        return repairs
    if slide_is_soft_bland(slide, source_text):
        focus = lesson_focus_phrase(source_text, fallback=normalize_text(slide.get("title", "")) or "the lesson idea")
        if not normalize_text(slide.get("primary_text", "")) or has_generic_filler(normalize_text(slide.get("primary_text", ""))):
            slide["primary_text"] = truncate_text(f"Use {focus.lower()} to show how the lesson idea works.", 150)
            repairs.append("reframed the page with lesson-specific primary copy")
        if not normalize_text(slide.get("secondary_text", "")) or has_generic_filler(normalize_text(slide.get("secondary_text", ""))):
            if vocabulary_bank:
                slide["secondary_text"] = truncate_text(f"Use {vocabulary_bank[0]} and the source example to explain your thinking.", 120)
            else:
                slide["secondary_text"] = truncate_text("Use the source example or model to explain your thinking.", 120)
            repairs.append("added a clearer premium support line")
        if not slide.get("tasks") or slide_is_sparse(slide):
            tasks = derive_task_list(slide, deck, vocabulary_bank)
            if slide.get("tasks"):
                tasks = tighten_list(list(slide.get("tasks", [])) + tasks, 3)
            slide["tasks"] = tasks
            repairs.append("expanded the activity structure so the page feels authored")
        if not slide.get("response_prompt") or has_generic_filler(normalize_text(slide.get("response_prompt", ""))):
            slide["response_prompt"] = derive_source_prompt(slide, deck, vocabulary_bank)
            repairs.append("tightened the response prompt around lesson evidence")
    return repairs


def repair_slide_density(slide: dict[str, Any]) -> list[str]:
    repairs: list[str] = []
    if slide_is_dense(slide):
        if slide.get("bullets"):
            slide["bullets"] = tighten_list(slide.get("bullets", []), 3)
            repairs.append("tightened bullet count for cleaner spacing")
        if slide.get("tasks"):
            slide["tasks"] = tighten_list(slide.get("tasks", []), 3)
            repairs.append("tightened task list for readability")
        if slide.get("sentence_starters"):
            slide["sentence_starters"] = tighten_list(slide.get("sentence_starters", []), 3)
            repairs.append("trimmed support copy to protect readability")
        if normalize_text(slide.get("secondary_text", "")):
            slide["secondary_text"] = truncate_text(normalize_text(slide.get("secondary_text", "")), 140)
            repairs.append("shortened secondary text to reduce visual crowding")
    return repairs


def repair_notebook_quality_issues(notebook_package: dict[str, Any], qa_report: dict[str, Any]) -> dict[str, Any]:
    repaired = deepcopy(notebook_package)
    plan = repaired["plan"]
    deck = repaired.get("deck")
    vocabulary_bank = collect_plan_vocabulary(plan)
    repairs_applied: list[str] = []
    for session_key, index, slide in iter_session_slides(plan):
        label = f"{session_key} slide {index + 1}"
        slide_repairs: list[str] = []
        title_repair = repair_slide_title(slide, deck)
        if title_repair:
            slide_repairs.append(f"{label}: {title_repair}")
        support_repair = repair_slide_supports(slide, deck, vocabulary_bank)
        if support_repair:
            slide_repairs.append(f"{label}: {support_repair}")
        for item in repair_slide_premium_finish(slide, deck, vocabulary_bank):
            slide_repairs.append(f"{label}: {item}")
        for item in repair_slide_copy(slide, deck, vocabulary_bank):
            slide_repairs.append(f"{label}: {item}")
        for item in repair_slide_density(slide):
            slide_repairs.append(f"{label}: {item}")
        if slide_repairs:
            repairs_applied.extend(slide_repairs)
            slide["enhancement_notes"] = tighten_list(slide.get("enhancement_notes", []) + slide_repairs[-4:], 4)
    categories = score_notebook_quality_by_category(repaired)
    hard_fails = detect_notebook_hard_fails(repaired, categories)
    repaired_report = {
        "passed": not hard_fails and all(item["status"] == "pass" for item in categories.values()),
        "categories": categories,
        "hardFails": hard_fails,
        "repairsApplied": repairs_applied,
        "warnings": [],
        "qualityTier": "fail"
        if hard_fails or any(item["status"] == "fail" for item in categories.values())
        else "weak"
        if any(item["status"] == "weak" for item in categories.values())
        else "pass",
    }
    if not repaired.get("deck"):
        repaired_report["warnings"].append("source_deck.json was not available, so source-aware repairs were limited")
    repaired["qaReport"] = repaired_report
    return repaired


def validate_notebook_ready_for_export(notebook_package: dict[str, Any]) -> bool:
    qa_report = notebook_package.get("qaReport") or evaluate_notebook_premium_quality(notebook_package)
    return bool(qa_report.get("passed", False))


def evaluate_notebook_premium_quality(notebook_package: dict[str, Any]) -> dict[str, Any]:
    categories = score_notebook_quality_by_category(notebook_package)
    hard_fails = detect_notebook_hard_fails(notebook_package, categories)
    warnings: list[str] = []
    if not notebook_package.get("deck"):
        warnings.append("source_deck.json missing; premium review can still run, but source-anchored repair is limited")
    return {
        "passed": not hard_fails and all(item["status"] == "pass" for item in categories.values()),
        "categories": categories,
        "hardFails": hard_fails,
        "repairsApplied": [],
        "warnings": warnings,
        "qualityTier": "fail"
        if hard_fails or any(item["status"] == "fail" for item in categories.values())
        else "weak"
        if any(item["status"] == "weak" for item in categories.values())
        else "pass",
    }


def copy_reference_files(bundle_dir: Path, output_dir: Path) -> None:
    for path in bundle_dir.iterdir():
        if path.is_file() and path.suffix.lower() in {".json", ".pptx"}:
            shutil.copy2(path, output_dir / path.name)


def run_notebook_enhancement(bundle_dir: Path, output_dir: Path) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    package = load_notebook_package(bundle_dir)
    initial_report = evaluate_notebook_premium_quality(package)
    repaired = repair_notebook_quality_issues(package, initial_report)
    final_report = repaired["qaReport"]
    copy_reference_files(bundle_dir, output_dir)
    write_json(output_dir / "notebook_plan.json", repaired["plan"])
    if repaired.get("deck"):
        write_json(output_dir / "source_deck.json", repaired["deck"])
    report_payload = {
        "bundleDir": package["bundleDir"],
        "initialReport": initial_report,
        "finalReport": final_report,
        "readyForRender": validate_notebook_ready_for_export(repaired),
        "qualityTier": final_report.get("qualityTier", "fail"),
    }
    write_json(output_dir / "enhancement_report.json", report_payload)
    if repaired.get("deck"):
        command_text = (
            f"python3 notebook_engine.py render {output_dir / 'notebook_plan.json'} "
            f"--deck {output_dir / 'source_deck.json'} --output-dir {output_dir / 'rendered'}"
        )
        (output_dir / "rerender_command.txt").write_text(command_text + "\n", encoding="utf-8")
    return report_payload


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Audit and polish an existing notebook bundle without changing the core notebook generator."
    )
    parser.add_argument("bundle_dir", help="Directory containing notebook_plan.json and optional bundle artifacts")
    parser.add_argument("--output-dir", required=True, help="Directory for the polished bundle")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    report = run_notebook_enhancement(Path(args.bundle_dir), Path(args.output_dir))
    final_report = report["finalReport"]
    print(f"enhancement_report: {Path(args.output_dir).resolve() / 'enhancement_report.json'}")
    print(f"premium_passed: {final_report.get('passed', False)}")
    print(f"repairs_applied: {len(final_report.get('repairsApplied', []))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
