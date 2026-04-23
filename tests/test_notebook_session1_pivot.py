import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from notebook_engine import (
    apply_publisher_copyedit,
    build_rendered_quality_report,
    build_responses_payload,
    ensure_peer_discussion_support,
    enforce_plan_requirements,
    notebook_plan_schema,
    render_plan,
    template_role_signature,
)
from notebook_engine_app import build_result_links
from notebook_folder_runner import copy_job_outputs


def sample_deck() -> dict:
    return {
        "lesson_title": "Determine the Volume of Rectangular Prisms",
        "source_filename": "volume-prisms.pptx",
        "slide_count": 5,
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


def regular_polygon_deck() -> dict:
    return {
        "lesson_title": "Apply Area Concepts to Solve Problems",
        "source_filename": "lesson-5-4.pptx",
        "slide_count": 7,
        "summary": "Students find the area of a regular polygon by decomposing an octagon into congruent triangles.",
        "keyword_candidates": ["area", "regular polygon", "octagon", "decompose"],
        "standards": ["6.G.A.1"],
        "slides": [
            {
                "slide_number": 1,
                "title": "Session 1",
                "text": "Session 1 Apply Area Concepts to Solve Problems",
                "text_items": ["Session 1"],
                "problem_texts": [],
                "notes": "",
                "image_count": 0,
            },
            {
                "slide_number": 2,
                "title": "Be Curious",
                "text": "Why does confidence in yourself and your abilities help you be successful in math?",
                "text_items": [
                    "What do you notice? What do you wonder?",
                    "Why does confidence in yourself and your abilities help you be successful in math?",
                ],
                "problem_texts": ["Why does confidence in yourself and your abilities help you be successful in math?"],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 3,
                "title": "Stop Sign",
                "text": "How can you decompose the octagon into shapes that you can find the area of?",
                "text_items": ["How can you decompose the octagon into shapes that you can find the area of?"],
                "problem_texts": ["How can you decompose the octagon into shapes that you can find the area of?"],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 4,
                "title": "Stop Sign",
                "text": "Let's decompose the octagon into 8 congruent triangles. Because each side has the same length, the octagon is a regular polygon. You can decompose a regular polygon into triangles to determine the area.",
                "text_items": [
                    "Let's decompose the octagon into 8 congruent triangles. Because each side has the same length, the octagon is a regular polygon. You can decompose a regular polygon into triangles to determine the area."
                ],
                "problem_texts": ["Why is the height of one triangle 18 inches?"],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 5,
                "title": "Stop Sign",
                "text": "The octagon can be decomposed into 8 congruent triangles. Each side is the base of a triangle. How can you determine the area of the octagon using the triangles?",
                "text_items": [
                    "The octagon can be decomposed into 8 congruent triangles.",
                    "Each side is the base of a triangle.",
                ],
                "problem_texts": ["How can you determine the area of the octagon using the triangles?"],
                "notes": "",
                "image_count": 1,
            },
            {
                "slide_number": 6,
                "title": "Learning Targets",
                "text": "I can find the area of a regular polygon by decomposing the figure into triangles. I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes.",
                "text_items": [
                    "I can find the area of a regular polygon by decomposing the figure into triangles.",
                    "I can make use of structure to find the area of a composite figure by decomposing the figure into other shapes.",
                ],
                "problem_texts": [],
                "notes": "",
                "image_count": 0,
            },
            {
                "slide_number": 7,
                "title": "Session 2",
                "text": "Session 2 Apply Area Concepts to Solve Problems",
                "text_items": ["Session 2"],
                "problem_texts": [],
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


def regular_polygon_plan() -> dict:
    return {
        "lesson_title": "Apply Area Concepts to Solve Problems",
        "subject": "Math",
        "grade_level": "6",
        "standards": ["6.G.A.1"],
        "topic_summary": "Use decomposition and congruent triangles to find the area of a regular polygon.",
        "session_1": {
            "session_title": "Session 1 Student Notebook",
            "session_subtitle": "",
            "slides": [],
        },
    }


class NotebookSession1PivotTests(unittest.TestCase):
    def test_session1_schema_matches_openai_required_key_rule(self) -> None:
        schema = notebook_plan_schema()

        self.assertEqual(set(schema["properties"].keys()), set(schema["required"]))
        self.assertIn("session_1", schema["properties"])
        self.assertNotIn("session_2", schema["properties"])

    def test_responses_payload_uses_session1_only_schema(self) -> None:
        payload = build_responses_payload(sample_deck(), model="gpt-5.4-mini", prompt_images=[])
        schema = payload["text"]["format"]["schema"]

        self.assertEqual(payload["text"]["format"]["name"], "session1_notebook_plan")
        self.assertEqual(set(schema["properties"].keys()), set(schema["required"]))
        self.assertNotIn("session_2", schema["properties"])

    def test_interactive_activity_regains_discussion_support_after_copyedit_pass(self) -> None:
        plan = enforce_plan_requirements(minimal_plan(), sample_deck())
        interactive_slide = plan["session_1"]["slides"][4]

        interactive_slide["discussion_questions"] = []
        interactive_slide["partner_prompt"] = ""

        repaired = apply_publisher_copyedit(plan, sample_deck())
        ensure_peer_discussion_support(repaired["session_1"]["slides"][4])

        questions = repaired["session_1"]["slides"][4]["discussion_questions"]
        self.assertGreaterEqual(len(questions), 2)
        self.assertTrue(all(question.endswith("?") for question in questions[:2]))
        self.assertTrue(repaired["session_1"]["slides"][4]["partner_prompt"])

    def test_exact_template_prefers_source_math_vocabulary_over_launch_words(self) -> None:
        plan = enforce_plan_requirements(regular_polygon_plan(), regular_polygon_deck())

        slides = plan["session_1"]["slides"]
        self.assertIn("regular polygon", slides[0]["primary_text"].lower())
        self.assertNotIn(2, slides[1]["source_slide_numbers"])

        vocab_words = {item["word"].lower() for item in slides[2]["vocabulary"]}
        self.assertIn("regular polygon", vocab_words)
        self.assertNotIn("confidence", vocab_words)
        self.assertNotIn("yourself", vocab_words)

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
