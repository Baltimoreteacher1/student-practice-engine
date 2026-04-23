# Education Workflow Workspace

This workspace is now organized around a small set of canonical folders for repeated Codex work, while preserving the existing lesson-plan and notebook systems already in use.

## Canonical folders

```text
.
├─ lesson-plan-engine/
├─ notebook-engine/
├─ shared-assets/
├─ quiz-form-builder/
└─ validation-tools/
```

Use those folders for new durable specs, inbox/output/archive flows, examples, gold standards, logs, and migration notes.

For the current active work-area map and recommended day-to-day workflow, use [WORKSPACE_RUNBOOK.md](WORKSPACE_RUNBOOK.md).

For the shortest reliable operating path through lesson plans, notebooks, and enhancement, use [OPERATING_CHECKLIST.md](OPERATING_CHECKLIST.md).

For the shared execution and review contract used across this workspace, use [CODEX_OPERATING_PLAYBOOK.md](CODEX_OPERATING_PLAYBOOK.md).

Default runtime decisions:

- lesson plans: `codex-lesson-plan-generator/` remains the active implementation
- notebooks: the preserved local root notebook runtime is the primary default path
- hosted notebook flow: `flagship-notebook-generator/` is secondary and used when web/API behavior is the actual target

## What each folder is for

- `lesson-plan-engine/`: canonical home for lesson-plan workflow specs, inbox/output/archive contracts, examples, and migration notes.
- `notebook-engine/`: canonical home for student notebook workflow specs, operational folders, and migration notes.
- `shared-assets/`: reusable templates, reference inputs, reference outputs, and shared gold standards.
- `quiz-form-builder/`: Google Apps Script quiz/form builders and related examples.
- `validation-tools/`: validation, extraction, QA, and audit tooling notes.

## Existing systems preserved in place

The current working implementations were preserved to avoid breaking active paths:

- `codex-lesson-plan-generator/`: existing lesson-plan generator repo
- `flagship-notebook-generator/`: existing notebook web app and backend
- root notebook engine files such as `notebook_engine.py`, `notebook_engine_app.py`, `notebook_folder_runner.py`
- selected legacy smoke folders and sample outputs at workspace root
- preserved local one-off scripts, scratch prompts, and preview renders in `shared-assets/archive/`

## How to use this workspace

1. Start in the canonical folder for the workflow you are working on.
2. Read that folder's `AGENTS.md`, `SPEC.md`, `TASKS.md`, and `README.md` when they exist.
3. Reuse preserved implementations before creating new code.
4. Put raw inputs in `INBOX/`, final outputs in `OUTPUT/`, completed source bundles in `ARCHIVE/`, and verification notes in `logs/`.
5. Promote clean examples and approved final references into `examples/` and `gold-standards/`.

## Active root entrypoints

Treat the workspace root as a thin runtime surface, not a general script dump. The normal root-level entrypoints are:

- `Generate Lesson Plan.command`
- `Launch Notebook Inbox.command`
- `process_notebook_inbox.command`
- `launch_notebook_engine.command`
- `notebook_engine.py`
- `notebook_engine_app.py`
- `notebook_folder_runner.py`
- `notebook_launchers.py`
- root `tests/`

Historical one-off helpers are preserved under `shared-assets/archive/` and should not be treated as active workflow tooling without cleanup.

## Validation shortcuts

- lesson plans: `python3 validation-tools/src/verify_lesson_plan_engine.py`
- notebooks: `python3 validation-tools/src/verify_notebook_engine.py`

## Practical rule

If a change is exploratory or legacy-specific, preserve first and document the next migration step. If a change is durable, place it in the canonical structure so future Codex runs can find it quickly.
