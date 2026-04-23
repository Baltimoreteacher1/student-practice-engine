# Lesson Plan Engine

Canonical workspace home for lesson-plan generation.

Current implementation is preserved in `../codex-lesson-plan-generator/`, but the active default runtime now reads from `INBOX/` and writes outputs to `OUTPUT/`.

Use this folder for durable workflow docs, clean inbox/output/archive handling, examples, gold standards, and future migration work.

Canonical runtime: `../codex-lesson-plan-generator/run.py`

Desktop launcher: `../Generate Lesson Plan.command`

Canonical verification:

```bash
python3 ../validation-tools/src/verify_lesson_plan_engine.py
```

Deterministic gold-standard benchmark:

```bash
python3 ../validation-tools/src/benchmark_lesson_plan_quality.py
```
