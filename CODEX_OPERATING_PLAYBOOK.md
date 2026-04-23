# Codex Operating Playbook

Use this playbook for repeatable, production-safe work in this workspace.

## 1. Start with routing

Before editing anything, decide which workflow you are actually touching:

- lesson plans:
  - contract docs: `lesson-plan-engine/`
  - live implementation: `codex-lesson-plan-generator/`
- notebooks:
  - contract docs: `notebook-engine/`
  - live local implementation: `notebook_engine.py`, `notebook_engine_app.py`, `notebook_folder_runner.py`
  - hosted implementation: `flagship-notebook-generator/`
- notebook polish on an existing bundle:
  - `notebook-engine/enhancement/`
- validation:
  - `validation-tools/`

If the task is durable, update the canonical folder docs as part of the same pass.

## 2. Default execution loop

For medium or large tasks, work in this order:

1. Goal
2. Context
3. Constraints
4. Done when

Then execute:

1. inspect the relevant code and docs
2. plan the work
3. implement the smallest safe change
4. verify with the smallest relevant checks
5. review the diff for regressions
6. summarize what changed, what was verified, and what risk remains

## 3. Notebook contract reminders

- Default to the compressed 6-slide Session 1 notebook unless the user explicitly asks for a different structure.
- Preserve source wording for the guided/source problem when available.
- Keep vocabulary anchored to source lesson language, not generic launch or mindset language.
- Treat notebook quality failures as real failures: weak hierarchy, placeholder copy, generic activities, answer leakage, or poor editability all require repair before finalizing.

## 4. Root surface discipline

The workspace root still contains a few preserved live entrypoints. Treat these as the only normal root-level execution paths:

- `Generate Lesson Plan.command`
- `Launch Notebook Inbox.command`
- `process_notebook_inbox.command`
- `launch_notebook_engine.command`
- `notebook_engine.py`
- `notebook_engine_app.py`
- `notebook_folder_runner.py`
- `notebook_launchers.py`
- root `tests/`

Do not treat preserved local one-off scripts or preview renders as reusable production tooling. Those belong in `shared-assets/archive/` until they are intentionally cleaned up and promoted.

## 5. Archive rule

When a root-level helper is no longer part of the active workflow but still needs to be preserved:

1. move it out of the workspace root
2. place it under `shared-assets/archive/`
3. document why it was preserved
4. keep it out of the normal production path

Archive first. Promote later only after cleanup, parameterization, and validation.

## 6. Validation floor

Use the smallest relevant validation that proves the touched path:

- lesson plans:
  - `python3 validation-tools/src/verify_lesson_plan_engine.py`
- notebooks:
  - `python3 validation-tools/src/verify_notebook_engine.py`
- focused notebook regressions:
  - `python3 -m pytest -q tests/test_notebook_session1_pivot.py tests/test_notebook_engine_publisher_polish.py`

If a required check cannot run, say exactly what failed and what manual verification path remains.

## 7. Final review checklist

Before finalizing:

- confirm the user-facing workflow is still obvious
- confirm docs point to real code paths
- confirm no broken file references remain
- confirm no root clutter is masquerading as an active runtime
- confirm verification passed or explain the blocker clearly
