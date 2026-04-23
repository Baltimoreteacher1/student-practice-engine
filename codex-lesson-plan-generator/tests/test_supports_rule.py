from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from apply_supports import APPROVED_SUPPORT_MAPPING, apply_supports  # noqa: E402
from build_lesson_plan import build_lesson_plan  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from utils import LessonPlanError  # noqa: E402
from validate_plan import build_validation_payload  # noqa: E402


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
            approved = set(APPROVED_SUPPORT_MAPPING[item["profile"]])
            self.assertTrue(set(item["supports"]).issubset(approved))
        self.assertTrue(
            lesson_plan["sessions"][0]["opening_warm_up_launch"]["embedded_supports"]
            or lesson_plan["sessions"][0]["guided_practice_collaborative_learning"]["embedded_supports"]
        )

    def test_custom_student_support_assignments_preserve_initials_and_matrix_text(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        config["student_support_assignments"] = [
            {
                "label": "M.E.",
                "profile": "Profile A",
                "supports": ["oral responses", "sentence starters", "scribed exit if needed"],
                "matrix_supports": "Oral responses, sentence starters, scribed exit if needed",
            }
        ]
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

        support = lesson_plan["sessions"][0]["differentiation_sped_esol_supports_and_teacher_notes"]["sped"][0]
        self.assertEqual(support["student"], "M.E.")
        self.assertEqual(support["profile"], "Profile A")
        self.assertIn("oral responses", support["supports"])
        self.assertEqual(support["matrix_supports"], "Oral responses, sentence starters, scribed exit if needed")

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

    def test_validation_fails_when_locked_iep_line_or_matrix_contract_disappears(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        config["student_support_assignments"] = [
            {
                "label": "M.E.",
                "profile": "Profile A",
                "supports": ["oral responses", "sentence starters", "scribed exit if needed"],
                "matrix_supports": "Oral responses, sentence starters, scribed exit if needed",
            }
        ]
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

        with patch(
            "validate_plan._build_compact_session_view",
            return_value={
                "iep_students_line": "",
                "accommodations_matrix_rows": [],
                "procedure_rows": [],
            },
        ):
            payload = build_validation_payload(
                lesson_plan=lesson_plan,
                raw_deck=raw_deck,
                config=config,
                output_file_status={
                    "json": True,
                    "markdown": False,
                    "docx_count": 0,
                    "validation_report": False,
                },
            )

        supports_check = next(check for check in payload["checks"] if check["name"] == "supports")
        self.assertFalse(payload["passed"])
        self.assertFalse(supports_check["passed"])
        self.assertIn("IEP Students line", supports_check["details"])

    def test_validation_fails_when_small_group_profiles_have_no_small_group_lesson_move(self) -> None:
        raw_deck = load_fixture()
        config = sample_config()
        config["student_support_assignments"] = [
            {
                "label": "A.M.M.",
                "profile": "Profile G",
                "supports": ["calculator", "forced choice options", "extended time"],
                "matrix_supports": "Calculator, forced choice options, extended time",
            }
        ]
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

        with patch(
            "validate_plan._build_compact_session_view",
            return_value={
                "iep_students_line": "IEP Students: A.M.M.",
                "accommodations_matrix_rows": [
                    {"student": "A.M.M.", "supports": "Calculator, forced choice options, extended time"}
                ],
                "procedure_rows": [
                    {
                        "phase_time": "Collaborative Practice\n(10 min)",
                        "sped_supports": ["A.M.M.: small group, calculator"],
                        "teacher_moves": ["Keep partner talk anchored to the discussion task."],
                        "student_moves": ["Discuss the collaborative prompt."],
                    },
                    {
                        "phase_time": "Independent Practice\n(10 min)",
                        "sped_supports": ["A.M.M.: small group, extended time"],
                        "teacher_moves": ["Assign the independent prompt."],
                        "student_moves": ["Complete the independent prompt."],
                    },
                ],
            },
        ):
            payload = build_validation_payload(
                lesson_plan=lesson_plan,
                raw_deck=raw_deck,
                config=config,
                output_file_status={
                    "json": True,
                    "markdown": False,
                    "docx_count": 0,
                    "validation_report": False,
                },
            )

        supports_check = next(check for check in payload["checks"] if check["name"] == "supports")
        self.assertFalse(payload["passed"])
        self.assertFalse(supports_check["passed"])
        self.assertIn("small-group lesson move", supports_check["details"])


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
