# Workspace Runbook

Use this file to decide where to work before editing anything.

## Active work areas

### Lesson plan generation

- durable workspace home: `lesson-plan-engine/`
- active code/work area today: `codex-lesson-plan-generator/`
- launcher path in use: `Generate Lesson Plan.command`

Reason: the preserved lesson-plan generator already has a coherent `src/`, config, templates, tests, and output contract. The canonical folder should guide work, but the preserved generator is the real implementation until migration is deliberate.

### Student notebook generation

- durable workspace home: `notebook-engine/`
- primary local runtime today:
  - `notebook_engine.py`
  - `notebook_engine_app.py`
  - `notebook_folder_runner.py`
  - `notebook_launchers.py`
  - `tests/test_notebook_engine_publisher_polish.py`
- secondary hosted runtime today:
  - `flagship-notebook-generator/backend/`
  - `flagship-notebook-generator/frontend/`
  - `flagship-notebook-generator/functions/`

Reason: the notebook workflow has two real paths. Use the root Python files for local generation, QA, inbox automation, and default production work. Use `flagship-notebook-generator/` only when the task is about the hosted web product.

### Quiz and form builders

- durable workspace home: `quiz-form-builder/`
- active code/work areas today:
  - `apps_script_export_area_triangles/`
  - root `.gs` files such as `Code.gs`, `NotebookExtractors_Precision.gs`, `Notebookgrids.gs`, `notebook_generator.gs`, `participant_workbook_session1.gs`

Reason: the Apps Script sources are still preserved where they have been used. The canonical folder is the place to document, consolidate, and later migrate the stable versions.

### Validation and QA tools

- durable workspace home: `validation-tools/`
- active code/work areas today:
  - `codex-lesson-plan-generator/tests/`
  - root `tests/`
  - notebook smoke and inspect folders at workspace root when reproducing render issues

Reason: validation already exists, but it is split between the lesson-plan repo, root notebook tests, and many preserved smoke folders. Treat `validation-tools/` as the consolidation target, not yet the sole runtime.

## Root surface rule

The workspace root is intentionally limited to active launchers, core runtimes, shared docs, and a small number of preserved outputs. Do not leave new one-off scripts at the root after a task is complete.

Preserve local-only helpers in `shared-assets/archive/` unless they are being promoted into a canonical workflow area.

## Best daily workflow

### Debug code

1. Start from the canonical folder docs for the workflow.
2. Jump immediately to the active code/work area named above.
3. Reproduce with the smallest existing test, smoke case, or launcher path.
4. Fix in the real implementation, not in the canonical placeholder folder.
5. Record durable notes or contract changes back in the canonical folder.

### Build a new script

1. Decide whether it is lesson-plan, notebook, quiz/form, or validation work.
2. Put the durable spec or task note in the canonical folder first.
3. If the script supports an existing engine, add it next to that engine’s real implementation.
4. If it is shared by multiple workflows, place it in `validation-tools/src/` or the future shared area only after reuse is proven.

### Improve the notebook pipeline

1. Default to the local runtime first.
2. For local generation or inbox flow, work in the root notebook Python files.
3. For hosted upload flow, work in `flagship-notebook-generator/`.
4. Reuse preserved smoke cases and the root notebook test before inventing new scaffolding.
5. Update `notebook-engine/SPEC.md`, `notebook-engine/TASKS.md`, or logs when the contract changes.

### Local notebook build

1. Use `python3 notebook_engine.py run <deck.pptx>` for a direct run.
2. Use `python3 notebook_engine.py run <deck.pptx> --offline` for a smoke pass without the API.
3. Use `python3 notebook_engine_app.py` for the local upload workflow.
4. Use `python3 notebook_folder_runner.py` for inbox-style recurring processing.

### Local notebook debugging

1. Start with the smallest failing path: direct CLI, app upload, or folder runner.
2. Prefer `extract`, `plan`, and `render` separately when isolating failures.
3. Use the root notebook test and the smallest preserved smoke case that still reproduces the issue.
4. Fix the root runtime first unless the bug exists only in the hosted path.

### Hosted notebook fallback workflow

1. Use `flagship-notebook-generator/backend/` for API and artifact bundling issues.
2. Use `flagship-notebook-generator/frontend/` and `functions/` for browser and proxy behavior.
3. Treat the hosted runtime as a secondary path that should follow the local runtime contract where practical.

### Create or improve a lesson-plan generator

1. Start in `lesson-plan-engine/` for the workflow contract.
2. Implement in `codex-lesson-plan-generator/`.
3. Preserve the fixed lesson-plan structure and standards-source rule.
4. Validate with the existing lesson-plan tests before broadening scope.
5. Promote examples, gold standards, and migration notes into the canonical folder as they become durable.

### Lesson-plan implementation workflow

1. Read `lesson-plan-engine/` docs first.
2. Implement in `codex-lesson-plan-generator/`.
3. Preserve the standards-source rule, fixed structure, and validation report contract.
4. Reflect any durable workflow change back into the canonical lesson-plan docs.

### Quiz and form implementation workflow

1. Read `quiz-form-builder/` docs first.
2. Inspect preserved Apps Script sources before changing names or entrypoints.
3. Preserve working Apps Script function names unless replacement is deliberate and documented.
4. Promote only verified, non-duplicated script sets into the canonical folder.

### Validation workflow

1. Start in `validation-tools/` for the workflow contract.
2. Use generator-local tests first when the change is isolated.
3. Use preserved smoke folders only when a smaller regression does not reproduce the issue.
4. Promote shared checks into `validation-tools/` only after reuse is proven.

## What not to do

- Do not start new durable work from the workspace root unless you are explicitly modifying a preserved active runtime that already lives there.
- Do not treat smoke folders as the source of truth.
- Do not migrate working code into the canonical folders without documenting parity and verification.
- Do not create duplicate notebook runtimes when the change belongs to either the local path or the hosted path.
- Do not confuse preserved archive material in `shared-assets/archive/` with supported production tooling.
