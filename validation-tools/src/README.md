# Source Notes

Place durable validators, QA helpers, extraction utilities, and audit scripts here when they are stable enough to maintain.

Prefer shared tools here when the same check should apply across lesson plans, notebooks, and Apps Script outputs.

Current canonical lesson-plan verifier:

- `verify_lesson_plan_engine.py` - runs tests, Python compile checks, the desktop launcher, and artifact existence checks for the sample lesson-plan deck.

Current canonical lesson-plan benchmark:

- `benchmark_lesson_plan_quality.py` - runs the lesson-plan launcher against the canonical sample deck and compares the generated JSON, Markdown, validation report, and DOCX text against the promoted gold-standard bundle.

Current canonical notebook benchmark:

- `benchmark_notebook_quality.py` - renders the deterministic notebook benchmark cases in `../examples/`, validates plan-level fidelity expectations, and checks the rendered quality report against locked thresholds.
