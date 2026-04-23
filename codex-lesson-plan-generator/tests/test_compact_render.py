from __future__ import annotations

import copy
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
from build_lesson_plan import build_lesson_plan  # noqa: E402
from lesson_extract import run_lesson_extract  # noqa: E402
from render_docx import _build_compact_session_view  # noqa: E402


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
        "student_support_assignments": [
            {
                "label": "M.E.",
                "profile": "Profile A",
                "supports": ["oral responses", "sentence starters", "scribed exit if needed"],
                "matrix_supports": "Oral responses, sentence starters, scribed exit if needed",
            },
            {
                "label": "J.C.",
                "profile": "Profile A",
                "supports": ["sentence starters", "para support", "movement breaks", "oral options"],
                "matrix_supports": "Sentence starters, para support, movement breaks, oral options",
            },
        ],
        "active_student_supports": [],
        "enable_esol_supports": True,
        "esol_support_limit": 3,
        "output": {"docx": "output/lesson_plan.docx"},
        "docx_style": {},
    }


def build_sample_session(config_overrides: dict | None = None) -> tuple[dict, dict]:
    raw_deck = load_fixture()
    config = copy.deepcopy(sample_config())
    if config_overrides:
        config.update(copy.deepcopy(config_overrides))
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
    return lesson_plan["sessions"][0], config


def build_composite_area_session() -> tuple[dict, dict]:
    session, config = build_sample_session()
    session = copy.deepcopy(session)
    config = dict(config)

    session["lesson_information"]["lesson_title"] = "Apply Area Concepts to Solve Problems"
    session["standards_and_learning_targets"]["standards"] = []
    session["standards_and_learning_targets"]["standards_status"] = "Not explicitly listed in source slides."
    session["standards_and_learning_targets"]["standards_source"] = {"slide_numbers": [], "source_lines": []}
    session["standards_and_learning_targets"]["learning_targets"] = [
        "I can find the area of a regular polygon by decomposing the figure into triangles.",
        "I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes.",
    ]
    session["standards_and_learning_targets"]["i_can_statements"] = list(
        session["standards_and_learning_targets"]["learning_targets"]
    )
    session["lesson_objective_and_student_success_criteria"][
        "lesson_objective"
    ] = "make use of structure to find the area of a composite figure by decomposing the figure into other shapes."

    def rewrite_phase(
        key: str,
        *,
        focus: list[str],
        teacher: list[str],
        student: list[str],
        evidence: list[str],
        excerpt: str,
    ) -> None:
        phase = session[key]
        phase["focus_tasks"] = focus
        phase["teacher_actions"] = teacher
        phase["student_actions"] = student
        phase["evidence_of_learning"] = evidence
        phase["source_excerpt"] = excerpt

    rewrite_phase(
        "opening_warm_up_launch",
        focus=["What do you notice about the trapezoid?", "How are they the same?"],
        teacher=["Display the trapezoid comparison prompt."],
        student=["Share one noticing and one wondering about the figures."],
        evidence=["Students compare the figures using area language."],
        excerpt="What do you notice about the trapezoid? How are they the same?",
    )
    rewrite_phase(
        "mini_lesson_modeling_concept_development",
        focus=[
            "The Ohio Burgee is an irregular shape. What is the area of the flag?",
            "Let's compose the flag into a trapezoid and other familiar shapes.",
        ],
        teacher=["Model how the composite figure can be decomposed into familiar shapes."],
        student=["Explain how the composite figure can be broken apart to find area."],
        evidence=["Students justify how the decomposition supports the solution."],
        excerpt="The Ohio Burgee is an irregular shape. What is the area of the flag? Let's compose the flag into a trapezoid.",
    )
    rewrite_phase(
        "guided_practice_collaborative_learning",
        focus=[
            "How can you use your prior knowledge of area to determine the area of the irregular shape?",
            "What is another way you could decompose the flag?",
        ],
        teacher=["Prompt partners to compare two decomposition strategies for the composite figure."],
        student=["Discuss how the irregular shape can be decomposed into shapes with known area."],
        evidence=["Students justify which decomposition is most efficient."],
        excerpt="How can you use your prior knowledge of area to determine the area of the irregular shape?",
    )
    rewrite_phase(
        "independent_practice_application_stations",
        focus=["Find the area of the composite figure and explain your method."],
        teacher=["Circulate while students solve the flag problem independently."],
        student=["Write an explanation of the decomposition used to find the area."],
        evidence=["Independent work shows accurate decomposition and area reasoning."],
        excerpt="Find the area of the composite figure and explain your method.",
    )
    rewrite_phase(
        "closure_exit_ticket_assessment",
        focus=["What is another way you could determine the area of the flag?"],
        teacher=["Collect one final written explanation about the composite figure strategy."],
        student=["Explain which decomposition strategy is most efficient and why."],
        evidence=["Students defend a strategy for the irregular shape."],
        excerpt="What is another way you could determine the area of the flag?",
    )
    session["reference_render_context"]["be_curious_prompts"] = ["What do you notice about the trapezoid?"]
    session["reference_render_context"]["vocabulary_terms"] = []
    session["reference_render_context"]["reasoning_tasks"] = [
        "Use structure to decompose the composite figure into shapes with known area.",
        "Explain how the irregular flag can be composed into a trapezoid.",
    ]
    session["reference_render_context"]["guided_practice"] = [
        "How can you use your prior knowledge of area to determine the area of the irregular shape?"
    ]
    session["reference_render_context"]["independent_practice"] = [
        "Find the area of the composite figure and explain your method."
    ]
    return session, config


class CompactRenderTests(unittest.TestCase):
    def test_reference_context_preserves_opening_prompts(self) -> None:
        session, _ = build_sample_session()
        context = session["reference_render_context"]

        self.assertIn("What shapes do you notice?", context["be_curious_prompts"])
        self.assertGreaterEqual(len(context["be_curious_prompts"]), 1)

    def test_compact_view_matches_reference_contract(self) -> None:
        session, config = build_sample_session()
        view = _build_compact_session_view(session, config)

        self.assertEqual(view["teacher_name"], "Neft.Alba")
        self.assertTrue(view["title_line"].startswith("Neft.Alba |"))
        self.assertEqual(view["do_now"], "What shapes do you notice? Estimate which shape has the greater area.")
        self.assertEqual(
            view["content_objective"],
            "I can determine the area of a triangle using a related rectangle or parallelogram",
        )
        self.assertEqual(
            view["language_objective"],
            "I will explain how the area of a triangle is related to a rectangle or parallelogram using sentence frames with the terms area, triangle, and rectangle.",
        )
        self.assertEqual(view["iep_students_line"], "IEP Students: M.E. | J.C.")
        self.assertTrue(view["standards_cells"][0].startswith("5.MD.C.5:"))
        self.assertTrue(view["standards_cells"][1].startswith("6.G.A.1:"))
        self.assertTrue(view["standards_cells"][2].startswith("7.G.B.6:"))
        self.assertEqual(len(view["procedure_rows"]), 8)
        self.assertEqual(len(view["vocabulary_rows"]), 5)
        self.assertEqual(len(view["accommodations_matrix_rows"]), 2)
        self.assertEqual(view["accommodations_matrix_rows"][0]["student"], "M.E.")
        self.assertGreaterEqual(len(view["procedure_rows"][3]["teacher_moves"]), 4)
        self.assertGreaterEqual(len(view["procedure_rows"][6]["student_moves"]), 3)
        self.assertGreaterEqual(len(view["procedure_rows"][7]["teacher_moves"]), 3)

    def test_compact_view_keeps_all_initials_visible_and_adds_source_grounded_small_group_moves(self) -> None:
        session, config = build_sample_session(
            {
                "student_support_assignments": [
                    {
                        "label": "M.E.",
                        "profile": "Profile A",
                        "supports": ["oral responses", "sentence starters", "scribed exit if needed"],
                        "matrix_supports": "Oral responses, sentence starters, scribed exit if needed",
                    },
                    {
                        "label": "S.D.L.P.",
                        "profile": "Profile F",
                        "supports": ["read aloud all text and questions"],
                        "matrix_supports": "Read aloud all text and questions",
                    },
                    {
                        "label": "R.B.",
                        "profile": "Profile C",
                        "supports": ["extended time", "fidget tool", "movement breaks"],
                        "matrix_supports": "Extended time, fidget tool, movement breaks",
                    },
                    {
                        "label": "A.M.M.",
                        "profile": "Profile G",
                        "supports": ["calculator", "forced choice options", "extended time"],
                        "matrix_supports": "Calculator, forced choice options, extended time",
                    },
                    {
                        "label": "R.M.G.",
                        "profile": "Profile E",
                        "supports": ["pre-circled/highlighted options", "forced choice", "pre-drawn number lines"],
                        "matrix_supports": "Pre-circled/highlighted options, forced choice, pre-drawn number lines",
                    },
                ]
            }
        )
        view = _build_compact_session_view(session, config)
        collaborative_row = next(
            row for row in view["procedure_rows"] if row["phase_time"].startswith("Collaborative Practice")
        )
        independent_row = next(
            row for row in view["procedure_rows"] if row["phase_time"].startswith("Independent Practice")
        )

        self.assertEqual(
            view["iep_students_line"],
            "IEP Students: M.E. | S.D.L.P. | R.B. | A.M.M. | R.M.G.",
        )
        self.assertEqual(len(collaborative_row["sped_supports"]), 5)
        self.assertEqual(len(independent_row["sped_supports"]), 5)
        self.assertTrue(any(line.startswith("R.M.G.:") for line in collaborative_row["sped_supports"]))
        self.assertTrue(any("small group" in line.lower() for line in collaborative_row["sped_supports"]))

        collaborative_teacher_blob = " ".join(collaborative_row["teacher_moves"]).lower()
        independent_teacher_blob = " ".join(independent_row["teacher_moves"]).lower()
        collaborative_student_blob = " ".join(collaborative_row["student_moves"]).lower()
        independent_student_blob = " ".join(independent_row["student_moves"]).lower()
        self.assertIn("small group", collaborative_teacher_blob)
        self.assertIn("same slide/book discussion", collaborative_teacher_blob)
        self.assertIn("doubling the triangle creates a parallelogram", collaborative_teacher_blob)
        self.assertIn("small group", independent_teacher_blob)
        self.assertIn("same slide/book task", independent_teacher_blob)
        self.assertIn("find the area of triangle b", independent_teacher_blob)
        self.assertIn("small group", collaborative_student_blob)
        self.assertIn("small group", independent_student_blob)

    def test_composite_area_view_prefers_matching_target_and_richer_vocabulary(self) -> None:
        session, config = build_composite_area_session()
        view = _build_compact_session_view(session, config)

        self.assertEqual(
            view["content_objective"],
            "I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes",
        )
        self.assertTrue(view["standards_cells"][1].startswith("Source learning target: I can make use of structure"))
        self.assertIn("composite or irregular figure", view["language_objective"])
        self.assertGreaterEqual(len(view["vocabulary_rows"]), 3)
        self.assertIn("Composite Figure", {row["term"] for row in view["vocabulary_rows"]})

    def test_compact_view_respects_selected_session_objective_when_targets_are_inherited(self) -> None:
        session, config = build_composite_area_session()
        session["lesson_objective_and_student_success_criteria"][
            "lesson_objective"
        ] = "make use of structure to find the area of a composite figure by decomposing the figure into other shapes."

        view = _build_compact_session_view(session, config)

        self.assertEqual(
            view["content_objective"],
            "I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes",
        )
        self.assertTrue(
            view["standards_cells"][1].startswith(
                "Source learning target: I can make use of structure to find the area of a composite figure"
            )
        )


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
