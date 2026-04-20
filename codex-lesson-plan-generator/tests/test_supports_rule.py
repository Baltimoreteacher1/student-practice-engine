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

from apply_supports import APPROVED_SUPPORT_MAPPING, apply_supports  # noqa: E402
from build_lesson_plan import build_lesson_plan  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from utils import LessonPlanError  # noqa: E402


def load_fixture() -> dict:
    fixture_path = ROOT / "tests" / "fixtures" / "sample_extracted_raw_slide_text.json"
    return json.loads(fixture_path.read_text(encoding="utf-8"))


def sample_config() -> dict:
    return {
        "lesson_duration_minutes": 55,
        "date_override": "2026-04-07",
        "default_grade": "6",
        "default_subject": "Mathematics",
        "materials_defaults": ["Teacher slide deck"]
    }


class SupportsRuleTests(unittest.TestCase):
    def test_support_mappings_stay_within_approved_list(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        config["active_student_supports"] = ["Profile A", "Profile G"]
        config["enable_esol_supports"] = True
        config["esol_support_limit"] = 2
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

        supports = lesson_plan["sessions"][0]["differentiation_sped_esol_supports_and_teacher_notes"]["sped"]
        for item in supports:
            approved = set(APPROVED_SUPPORT_MAPPING[item["student"]])
            self.assertTrue(set(item["supports"]).issubset(approved))
        self.assertTrue(
            lesson_plan["sessions"][0]["opening_warm_up_launch"]["embedded_supports"]
            or lesson_plan["sessions"][0]["guided_practice_collaborative_learning"]["embedded_supports"]
        )

    def test_unknown_student_initials_fail_loudly(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        config["active_student_supports"] = ["UNKNOWN"]
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
        with self.assertRaises(LessonPlanError):
            apply_supports(lesson_plan, config)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
