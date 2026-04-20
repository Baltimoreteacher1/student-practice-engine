import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from notebook_engine import (
    build_rendered_quality_report,
    enforce_plan_requirements,
    render_plan,
    template_role_signature,
)
from notebook_engine_app import build_result_links
from notebook_folder_runner import copy_job_outputs


def sample_deck() -> dict:
    return {
        "lesson_title": "Determine the Volume of Rectangular Prisms",
        "summary": "Students use dimensions and the volume formula to solve the first prism problem.",
        "keyword_candidates": ["volume", "rectangular prism", "dimensions"],
        "standards": ["6.G.A.2"],
        "slides": [
            {
                "slide_number": 1,
                "title": "Learning Target",
                "text": "We will determine the volume of rectangular prisms.",
                "text_items": [
                    "We will determine the volume of rectangular prisms.",
                    "Students will explain how the dimensions and formula show the volume.",
                ],
                "problem_texts": [],
                "notes": "",
                "image_count": 0,
            },
            {
                "slide_number": 2,
                "title": "Be Curious",
                "text": "What do you notice about the rectangular prism model?",
                "text_items": [
                    "What do you notice about the rectangular prism model?",
                    "What do you wonder about the prism?",
                ],
                "problem_texts": ["What do you notice about the rectangular prism model?"],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 3,
                "title": "Vocabulary",
                "text": "rectangular prism, volume, dimensions",
                "text_items": ["Rectangular prism", "Volume", "Dimensions"],
                "problem_texts": [],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 4,
                "title": "First Problem",
                "text": "A rectangular prism has length 4 units, width 3 units, and height 2 units. What is its volume?",
                "text_items": ["Write V = l x w x h."],
                "problem_texts": [
                    "A rectangular prism has length 4 units, width 3 units, and height 2 units. What is its volume?"
                ],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 5,
                "title": "Your Turn",
                "text": "Solve the prism problem and explain how you know your answer is correct.",
                "text_items": ["Use the formula, show the work, and label the answer."],
                "problem_texts": [
                    "A rectangular prism has length 6 units, width 2 units, and height 3 units. What is its volume?"
                ],
                "notes": "",
                "image_count": 0,
            },
        ],
    }


def minimal_plan() -> dict:
    return {
        "lesson_title": "Determine the Volume of Rectangular Prisms",
        "subject": "Math",
        "grade_level": "6",
        "standards": ["6.G.A.2"],
        "topic_summary": "Use dimensions and the volume formula to solve the first prism problem.",
        "session_1": {
            "session_title": "Session 1 Student Notebook",
            "session_subtitle": "",
            "slides": [],
        },
    }


class NotebookSession1PivotTests(unittest.TestCase):
    def test_enforce_plan_requirements_builds_six_slide_session1_only_plan(self) -> None:
        plan = enforce_plan_requirements(minimal_plan(), sample_deck())

        self.assertNotIn("session_2", plan)
        slides = plan["session_1"]["slides"]
        self.assertEqual(len(slides), 6)
        self.assertEqual(
            template_role_signature(plan["session_1"]),
            [
                ("learning_target", "learning_objectives"),
                ("be_curious", "prior_session_review"),
                ("vocabulary", "vocabulary_table"),
                ("worked_example", "guided_practice"),
                ("practice", "interactive_activity"),
                ("practice", "best_fit_review"),
            ],
        )

    def test_render_plan_and_quality_report_only_emit_session1_outputs(self) -> None:
        plan = enforce_plan_requirements(minimal_plan(), sample_deck())

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_dir = Path(tmp_dir)
            outputs = render_plan(plan, sample_deck(), output_dir)
            report = build_rendered_quality_report(plan, outputs)

            self.assertEqual(set(outputs.keys()), {"session1"})
            self.assertTrue(outputs["session1"].exists())
            self.assertEqual(set(report["sessions"].keys()), {"Session 1"})

    def test_build_result_links_omits_session2_when_not_rendered(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            app_home = Path(tmp_dir)
            job_dir = app_home / "runs" / "job-1"
            job_dir.mkdir(parents=True)
            session1 = job_dir / "Session 1 - Student Notebook.pptx"
            plan_path = job_dir / "notebook_plan.json"
            deck_path = job_dir / "source_deck.json"
            session1.write_bytes(b"pptx")
            plan_path.write_text("{}", encoding="utf-8")
            deck_path.write_text("{}", encoding="utf-8")

            with patch("notebook_engine_app.APP_HOME", app_home):
                links = build_result_links(
                    job_dir,
                    {
                        "outputs": {"session1": session1},
                        "plan_path": plan_path,
                        "deck_path": deck_path,
                    },
                )

            self.assertIn("session1", links)
            self.assertNotIn("session2", links)

    def test_copy_job_outputs_handles_jobs_without_session2(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            runs_dir = Path(tmp_dir) / "runs"
            output_dir = Path(tmp_dir) / "output"
            job_dir = runs_dir / "job-1"
            job_dir.mkdir(parents=True)
            files = {
                "session1": job_dir / "Session 1 - Student Notebook.pptx",
                "plan": job_dir / "notebook_plan.json",
                "deck": job_dir / "source_deck.json",
                "quality_report": job_dir / "quality_report.json",
            }
            for path in files.values():
                if path.suffix == ".pptx":
                    path.write_bytes(b"pptx")
                else:
                    path.write_text("{}", encoding="utf-8")

            job = {
                "job_id": "job-1",
                "source_filename": "lesson-deck.pptx",
                "relative_files": {
                    "session1": files["session1"].name,
                    "plan": files["plan"].name,
                    "deck": files["deck"].name,
                    "quality_report": files["quality_report"].name,
                },
            }

            with patch("notebook_engine_app.RUNS_DIR", runs_dir):
                copied = copy_job_outputs(job, output_dir)

            self.assertIn("session1", copied)
            self.assertNotIn("session2", copied)
            self.assertTrue(copied["session1"].exists())


if __name__ == "__main__":
    unittest.main()
