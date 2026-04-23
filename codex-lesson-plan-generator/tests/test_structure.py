from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from apply_supports import apply_supports  # noqa: E402
from build_lesson_plan import SESSION_SECTION_KEYS, build_lesson_objective, build_lesson_plan  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from validate_plan import build_validation_payload  # noqa: E402


def load_fixture() -> dict:
    fixture_path = ROOT / "tests" / "fixtures" / "sample_extracted_raw_slide_text.json"
    return json.loads(fixture_path.read_text(encoding="utf-8"))


def sample_config() -> dict:
    return {
        "lesson_duration_minutes": 55,
        "date_override": "2026-04-07",
        "teacher_name": "Neft.Alba",
        "default_grade": "6",
        "default_subject": "Mathematics",
        "materials_defaults": ["Teacher slide deck"],
        "active_student_supports": ["Profile A"],
        "enable_esol_supports": True,
        "esol_support_limit": 2,
        "output": {"docx": "output/lesson_plan.docx"},
        "docx_style": {},
    }


def multi_session_fixture() -> dict:
    return {
        "source_file": "/tmp/example-multi.pptx",
        "source_filename": "example-multi.pptx",
        "slide_count": 10,
        "learning_target_candidate_numbers": [3, 10],
        "slides": [
            {"slide_number": 1, "title": "Session 1", "text_items": ["Session 1", "Determine the Area of Parallelograms"], "full_text": "Session 1 Determine the Area of Parallelograms", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 2, "title": "Be Curious", "text_items": ["Be Curious", "What do you notice?", "What do you wonder?", "Mindset", "How can different ideas help you understand math better?"], "full_text": "Be Curious What do you notice? What do you wonder? Mindset How can different ideas help you understand math better?", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 3, "title": "Learning Target", "text_items": ["Learning Target", "Standard: 6.G.A.1", "I can determine the area of a parallelogram by composing it into a rectangle."], "full_text": "Learning Target Standard 6.G.A.1 I can determine the area of a parallelogram by composing it into a rectangle.", "speaker_notes": [], "is_learning_target_candidate": True},
            {"slide_number": 4, "title": "Parallelograms", "text_items": ["How can you decompose the parallelogram into other shapes that you know?", "We can compose the pieces into a rectangle."], "full_text": "How can you decompose the parallelogram into other shapes that you know? We can compose the pieces into a rectangle.", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 5, "title": "Collaborate and Connect", "text_items": ["How can you use grid paper to confirm the area of the parallelogram found using the formula?"], "full_text": "How can you use grid paper to confirm the area of the parallelogram found using the formula?", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 6, "title": "Let's Explore More", "text_items": ["What is the missing dimension of the parallelogram shown?", "Workspace"], "full_text": "What is the missing dimension of the parallelogram shown? Workspace", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 7, "title": "Session 2", "text_items": ["Session 2", "Determine the Area of Rhombuses"], "full_text": "Session 2 Determine the Area of Rhombuses", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 8, "title": "Be Curious", "text_items": ["Be Curious", "What do you notice?", "What do you wonder?"], "full_text": "Be Curious What do you notice? What do you wonder?", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 9, "title": "Tiling a Backsplash", "text_items": ["A diagonal is a line segment that connects two nonconsecutive vertices of a polygon.", "The diagonals of a rhombus are perpendicular.", "How can you compose the rhombus into a rectangle?"], "full_text": "A diagonal is a line segment that connects two nonconsecutive vertices of a polygon. The diagonals of a rhombus are perpendicular. How can you compose the rhombus into a rectangle?", "speaker_notes": [], "is_learning_target_candidate": False},
            {"slide_number": 10, "title": "Learning Targets", "text_items": ["Learning Targets", "I can find the area of rhombuses by using a formula.", "I can attend to precision to find a missing dimension by using the area formula."], "full_text": "Learning Targets I can find the area of rhombuses by using a formula. I can attend to precision to find a missing dimension by using the area formula.", "speaker_notes": [], "is_learning_target_candidate": True}
        ]
    }


class StructureTests(unittest.TestCase):
    def test_required_session_sections_exist(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        lesson_extract = run_lesson_extract(raw_deck, config)
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            lesson_plan = build_lesson_plan(
                lesson_extract=lesson_extract,
                raw_deck=raw_deck,
                lesson_type="direct_instruction",
                config=config,
                fidelity_output_path=tmp_path / "fidelity.json",
                agenda_items=[],
                run_date="2026-04-07",
                requested_session_numbers=[],
            )
            lesson_plan = apply_supports(lesson_plan, config)

        self.assertIn("sessions", lesson_plan)
        self.assertEqual(len(lesson_plan["sessions"]), 1)
        self.assertEqual(lesson_plan["sessions"][0]["output_filename"], "lesson_plan.docx")
        for key in SESSION_SECTION_KEYS:
            self.assertIn(key, lesson_plan["sessions"][0])

    def test_multisession_deck_defaults_to_both_sessions(self) -> None:
        raw_deck = multi_session_fixture()
        config = sample_config()
        lesson_extract = run_lesson_extract(raw_deck, config)
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            lesson_plan = build_lesson_plan(
                lesson_extract=lesson_extract,
                raw_deck=raw_deck,
                lesson_type="direct_instruction",
                config=config,
                fidelity_output_path=tmp_path / "fidelity.json",
                agenda_items=[],
                run_date="2026-04-07",
                requested_session_numbers=[],
            )
        self.assertEqual(len(lesson_plan["sessions"]), 2)
        self.assertEqual(lesson_plan["sessions"][0]["session_label"], "Session 1")
        self.assertEqual(lesson_plan["sessions"][1]["session_label"], "Session 2")
        self.assertTrue(lesson_plan["sessions"][0]["output_filename"].endswith("Session_1.docx"))
        self.assertTrue(lesson_plan["sessions"][1]["output_filename"].endswith("Session_2.docx"))

    def test_validation_fails_when_required_section_is_missing(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        lesson_extract = run_lesson_extract(raw_deck, config)
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            lesson_plan = build_lesson_plan(
                lesson_extract=lesson_extract,
                raw_deck=raw_deck,
                lesson_type="direct_instruction",
                config=config,
                fidelity_output_path=tmp_path / "fidelity.json",
                agenda_items=[],
                run_date="2026-04-07",
                requested_session_numbers=[],
            )
        lesson_plan["sessions"][0].pop("closure_exit_ticket_assessment")
        payload = build_validation_payload(
            lesson_plan=lesson_plan,
            raw_deck=raw_deck,
            config=config,
            output_file_status={
                "json": True,
                "markdown": False,
                "docx_count": 0,
                "validation_report": False
            }
        )
        self.assertFalse(payload["passed"])
        self.assertIn("Session 1: closure_exit_ticket_assessment", payload["missing_sections"])

    def test_lesson_objective_prefers_best_matching_learning_target(self) -> None:
        session_extract = {
            "learning_targets": [
                "I can find the area of a regular polygon by decomposing the figure into triangles.",
                "I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes.",
            ],
            "reasoning_tasks": [
                "Use structure to decompose the composite figure into shapes with known area.",
                "Explain how the irregular flag can be composed into a trapezoid.",
            ],
            "opening_source": {"lines": ["What do you notice about the trapezoid?"]},
            "modeling_source": {"lines": ["The Ohio Burgee is an irregular shape. What is the area of the flag?"]},
            "guided_practice": ["How can you use your prior knowledge of area to determine the area of the irregular shape?"],
            "independent_practice": ["Find the area of the composite figure and explain your method."],
            "checks_for_understanding": [],
            "session_title": "Apply Area Concepts to Solve Problems",
        }

        self.assertEqual(
            build_lesson_objective(session_extract),
            "make use of structure to find the area of a composite figure by decomposing the figure into other shapes.",
        )


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
