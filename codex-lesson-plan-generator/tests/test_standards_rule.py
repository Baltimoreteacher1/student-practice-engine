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

from build_lesson_plan import build_lesson_plan  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from validate_plan import validate_standards_source  # noqa: E402


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


class StandardsRuleTests(unittest.TestCase):
    def test_standards_are_pulled_from_learning_target_slide(self) -> None:
        raw_deck = load_fixture()
        lesson_extract = run_lesson_extract(raw_deck, sample_config())
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            lesson_plan = build_lesson_plan(
                lesson_extract=lesson_extract,
                raw_deck=raw_deck,
                lesson_type="direct_instruction",
                config=sample_config(),
                fidelity_output_path=tmp_path / "fidelity.json",
                agenda_items=[],
                run_date="2026-04-07",
                requested_session_numbers=[],
            )
        section = lesson_plan["sessions"][0]["standards_and_learning_targets"]
        self.assertEqual(section["standards"], ["6.G.A.1"])
        self.assertEqual(section["standards_source"]["slide_numbers"], [3])

    def test_validation_rejects_non_learning_target_standard_source(self) -> None:
        raw_deck = load_fixture()
        lesson_extract = run_lesson_extract(raw_deck, sample_config())
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            lesson_plan = build_lesson_plan(
                lesson_extract=lesson_extract,
                raw_deck=raw_deck,
                lesson_type="direct_instruction",
                config=sample_config(),
                fidelity_output_path=tmp_path / "fidelity.json",
                agenda_items=[],
                run_date="2026-04-07",
                requested_session_numbers=[],
            )
        lesson_plan["sessions"][0]["standards_and_learning_targets"]["standards_source"]["slide_numbers"] = [1]
        passed, details = validate_standards_source(lesson_plan, raw_deck)
        self.assertFalse(passed)
        self.assertIn("Learning Target", details)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
