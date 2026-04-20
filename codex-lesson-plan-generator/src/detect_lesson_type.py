from __future__ import annotations

from collections import defaultdict
from typing import Any

from utils import normalize_whitespace


LESSON_TYPE_RULES = {
    "direct_instruction": (
        "learning target",
        "objective",
        "model",
        "mini-lesson",
        "i can",
        "success criteria",
    ),
    "review": (
        "review",
        "spiral",
        "revisit",
        "summarize",
        "what do you remember",
    ),
    "guided_practice": (
        "guided practice",
        "collaborate",
        "connect",
        "turn and talk",
        "we do",
        "check for understanding",
    ),
    "assessment_review": (
        "assessment",
        "quiz review",
        "benchmark",
        "test review",
        "reteach after assessment",
    ),
    "intervention": (
        "intervention",
        "reteach",
        "small group",
        "scaffold",
        "targeted support",
    ),
}


def detect_lesson_type(raw_deck: dict[str, Any]) -> str:
    scores: dict[str, int] = defaultdict(int)
    slides = raw_deck.get("slides", [])
    joined_text = " ".join(
        normalize_whitespace(f"{slide.get('title', '')} {slide.get('full_text', '')}").lower()
        for slide in slides
    )

    for lesson_type, markers in LESSON_TYPE_RULES.items():
        for marker in markers:
            scores[lesson_type] += joined_text.count(marker)

    if raw_deck.get("learning_target_candidate_numbers"):
        scores["direct_instruction"] += 2

    if any("guided practice" in slide.get("full_text", "").lower() for slide in slides):
        scores["guided_practice"] += 2
    if any("exit ticket" in slide.get("full_text", "").lower() for slide in slides):
        scores["assessment_review"] += 1
    if any("independent practice" in slide.get("full_text", "").lower() for slide in slides):
        scores["direct_instruction"] += 1

    ranked = sorted(scores.items(), key=lambda item: (-item[1], item[0]))
    top_label, top_score = ranked[0] if ranked else ("mixed_lesson", 0)
    second_score = ranked[1][1] if len(ranked) > 1 else 0

    if top_score <= 0:
        return "mixed_lesson"
    if top_score - second_score <= 1 and second_score > 0:
        return "mixed_lesson"
    return top_label
