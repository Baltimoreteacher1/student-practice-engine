import subprocess
import unittest
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
BENCHMARK_SCRIPT = WORKSPACE_ROOT / "validation-tools" / "src" / "benchmark_lesson_plan_quality.py"


class LessonPlanQualityBenchmarkTests(unittest.TestCase):
    def test_lesson_plan_gold_standard_passes(self) -> None:
        result = subprocess.run(
            ["python3", str(BENCHMARK_SCRIPT)],
            cwd=str(WORKSPACE_ROOT),
            check=True,
            capture_output=True,
            text=True,
        )

        self.assertIn("PASS lesson_plan_sample_triangle_launch", result.stdout)
        self.assertIn("Lesson plan quality benchmark passed.", result.stdout)


if __name__ == "__main__":
    unittest.main()
