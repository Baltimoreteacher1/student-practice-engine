import subprocess
import unittest
from pathlib import Path


WORKSPACE_ROOT = Path(__file__).resolve().parents[1]
BENCHMARK_SCRIPT = WORKSPACE_ROOT / "validation-tools" / "src" / "benchmark_notebook_quality.py"


class NotebookQualityBenchmarkTests(unittest.TestCase):
    def test_benchmark_cases_pass(self) -> None:
        result = subprocess.run(
            ["python3", str(BENCHMARK_SCRIPT)],
            cwd=str(WORKSPACE_ROOT),
            check=True,
            capture_output=True,
            text=True,
        )

        self.assertIn("Notebook quality benchmark passed for", result.stdout)
        self.assertIn("notebook_regular_polygon_fidelity_case", result.stdout)
        self.assertIn("notebook_session1_volume_case", result.stdout)


if __name__ == "__main__":
    unittest.main()
