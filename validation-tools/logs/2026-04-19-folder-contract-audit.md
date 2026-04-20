# EduWonderLab Folder Contract Audit

Date: 2026-04-19
Workspace: `/Users/joelneft/.codex/workspaces/default`
Companion code audit: `validation-tools/logs/2026-04-19-eduwonderlab-pipeline-audit.md`

## Scope

This extension audit covered the workspace folder layout itself, with emphasis on lesson-plan input and output handling:

- top-level folder classification
- canonical folder contract drift
- input/output portability
- checked-in generated artifact hygiene
- privacy and local-path leakage in extracted and output files
- launcher and script defaults that still assumed one machine

## Outcome

No CRITICAL or HIGH folder-level issues remain in the active audited pipeline after fixes.

The remaining risks are LOW and operational rather than release-blocking:

1. broader canonical migration into `lesson-plan-engine/` and `notebook-engine/` is still incomplete, even though the active folder contracts are now documented and enforced

## Final State Addendum

After the original folder audit:

- `run.py` was designated as the canonical lesson-plan generator
- `Generate Lesson Plan.command` was switched to the canonical Python runtime
- the legacy Node exporter, `run.js`, was retired
- `validation-tools/src/verify_lesson_plan_engine.py` was added as the canonical end-to-end lesson-plan verification script
- root-level smoke and `tmp_*` directories were moved into:
  - `notebook-engine/ARCHIVE/root-smoke-runs/`
  - `notebook-engine/ARCHIVE/root-temp-runs/`
- root-level artifact folders were moved into:
  - `notebook-engine/ARCHIVE/root-artifact-folders/`
- adjacent non-pipeline projects were moved into:
  - `notebook-engine/ARCHIVE/root-adjacent-projects/`
- sample canonical artifacts were promoted into:
  - `lesson-plan-engine/examples/`
  - `lesson-plan-engine/gold-standards/`
- the `localstorage-file` Node warning was confirmed to be external to the repo and became moot once the legacy Node exporter was removed

## 2026-04-20 Closeout Addendum

- Added `validation-tools/src/verify_lesson_plan_engine.py` as the canonical end-to-end verifier for the preserved lesson-plan runtime.
- Verified the end-to-end lesson-plan path on 2026-04-20 with:
  - `python3 validation-tools/src/verify_lesson_plan_engine.py`
- Archived these adjacent root projects into `notebook-engine/ARCHIVE/root-adjacent-projects/`:
  - `architect-mode-skill`
  - `axiom-adventure-game`
  - `axiom-playable-game`
  - `lesson-to-game-studio`
  - `mcap-game-blueprint`
  - `student-notebook`

## CRITICAL And HIGH Issues Fixed

### 1. Lesson-plan extracted and output artifacts leaked machine-local source paths

- Files:
  - `codex-lesson-plan-generator/src/utils.py:93-100`
  - `codex-lesson-plan-generator/src/extract_slides.py:117-167`
  - `codex-lesson-plan-generator/run.py:133-138`
  - `codex-lesson-plan-generator/tests/test_trigger_flow.py:39-58`
  - `codex-lesson-plan-generator/extracted/raw_slide_text.json:1-3`
  - `codex-lesson-plan-generator/extracted/normalized_lesson.json:1-5`
  - `codex-lesson-plan-generator/output/lesson_plan.json:1-6`
- Severity: HIGH
- Problem:
  - The extractor serialized `pptx_path.resolve()`, which wrote absolute machine-local source paths into checked-in extracted artifacts and generated lesson-plan JSON output.
  - That leaks local environment details and makes artifacts look non-portable.
- Exact fix:
  - Added `sanitize_artifact_path()` to emit a repo-relative path when the deck is inside the workspace and a basename otherwise.
  - Updated `extract_slide_deck()` to use that sanitizer.
  - Passed `ROOT` into the extractor from `run.py`.
  - Added regression tests.
  - Regenerated the checked-in extracted and output JSON artifacts.

### 2. Lesson-plan folder docs still told operators to use one personal machine path

- Files:
  - `codex-lesson-plan-generator/README.md:86-139`
  - `codex-lesson-plan-generator/CLAUDE.md:50-55`
  - `codex-lesson-plan-generator/config/locked_rules.md:49-50`
- Severity: HIGH
- Problem:
  - The lesson-plan folder docs still referenced `/Users/joelneft/...` paths and one named maintainer.
  - That is exactly the kind of folder-contract drift that causes misruns and false assumptions across machines.
- Exact fix:
  - Rewrote the README to use repo-local folders such as `Lesson Plan Inbox/`, `output/`, and `extracted/`.
  - Replaced the absolute `run.js` path in `CLAUDE.md` with a repo-local reference.
  - Replaced `Joel explicitly changes the model` with `the workspace maintainer explicitly changes the model`.

### 3. Notebook enhancement folder launch contract was still machine-local

- Files:
  - `notebook-engine/enhancement/Run Notebook Enhancement.command:4-7`
  - `notebook-engine/enhancement/src/run_enhancement_inbox.py:15-22`
- Severity: HIGH
- Problem:
  - The launcher pointed to a hardcoded absolute Python runner path.
  - The script itself defaulted to `~/Documents/Chatgpt Notebook and Lesson plans/Notebook Enhancement`, which is not a portable workspace contract.
- Exact fix:
  - Made the launcher resolve the runner from `BASE_DIR`.
  - Changed the Python script default base directory to the enhancement folder itself, with optional env override via `EDUWONDERLAB_NOTEBOOK_ENHANCEMENT_DIR`.

## Patch-Style Diff Snippets

### A. Sanitized lesson-plan source paths

```diff
-        "source_file": str(pptx_path.resolve()),
+        "source_file": sanitize_artifact_path(pptx_path, base_dir),
```

```diff
-raw_deck = extract_slide_deck(deck_path, raw_output_path)
+raw_deck = extract_slide_deck(deck_path, raw_output_path, base_dir=ROOT)
```

```diff
-  "source_file": "/Users/joelneft/.codex/workspaces/default/codex-lesson-plan-generator/examples/sample_input_slides.pptx",
+  "source_file": "examples/sample_input_slides.pptx",
```

### B. Portable lesson-plan folder docs

```diff
-1. Place one `.pptx` teacher slide deck in `/Users/joelneft/Documents/Chatgpt Notebook and Lesson plans/Lesson Plan Inbox/`.
+1. Place one `.pptx` teacher slide deck in `Lesson Plan Inbox/`, or pass an explicit `--deck` path.
```

```diff
-/Users/joelneft/Documents/Chatgpt Notebook and Lesson plans/Lesson Plan Output/lesson_plan.json
+output/lesson_plan.json
```

```diff
- Main entrypoint: [run.js](/Users/joelneft/.codex/workspaces/default/codex-lesson-plan-generator/run.js)
+ Main entrypoint: `run.js`
```

### C. Portable enhancement folder contract

```diff
-RUNNER="/Users/joelneft/.codex/workspaces/default/notebook-engine/enhancement/src/run_enhancement_inbox.py"
+RUNNER="$BASE_DIR/src/run_enhancement_inbox.py"
```

```diff
-DEFAULT_BASE_DIR = Path.home() / "Documents" / "Chatgpt Notebook and Lesson plans" / "Notebook Enhancement"
+DEFAULT_BASE_DIR = Path(
+    os.environ.get("EDUWONDERLAB_NOTEBOOK_ENHANCEMENT_DIR")
+    or Path(__file__).resolve().parents[1]
+)
```

## Lesson-Plan Input And Output Deep Dive

### Lesson-plan canonical folders

| Folder | Status | Notes |
| --- | --- | --- |
| `lesson-plan-engine/INBOX` | Pass | Present and empty. Canonical durable input location exists. |
| `lesson-plan-engine/OUTPUT` | Pass | Canonical output location is active and verified. |
| `lesson-plan-engine/ARCHIVE` | Pass | Present and empty. Canonical durable archive location exists. |
| `lesson-plan-engine/examples` | Pass | Contains promoted scrubbed sample input and sample output references. |
| `lesson-plan-engine/gold-standards` | Pass | Contains promoted sample lesson-plan DOCX reference. |

### Preserved lesson-plan generator runtime folders

| Folder | Status | Notes |
| --- | --- | --- |
| `codex-lesson-plan-generator/Lesson Plan Inbox` | Pass | Legacy inbox marker retained only as guidance toward the canonical inbox. |
| `codex-lesson-plan-generator/inputs` | Pass | Duplicate sample input was removed from the preserved generator. |
| `codex-lesson-plan-generator/extracted` | Pass | Legacy checked-in extracted artifacts were removed after the canonical output migration. |
| `codex-lesson-plan-generator/output` | Pass | Legacy checked-in generated outputs were removed after the canonical output migration. |
| `codex-lesson-plan-generator/examples` | Pass | Sample deck and sample output live here as explicit examples. |

### Current lesson-plan folder contents

```text
lesson-plan-engine/INBOX                -> empty
lesson-plan-engine/OUTPUT               -> canonical generated outputs plus extracted artifacts
lesson-plan-engine/ARCHIVE              -> empty
lesson-plan-engine/examples             -> sample_input_slides.pptx, sample_lesson_plan.json, sample_lesson_plan.md, sample_validation_report.md
lesson-plan-engine/gold-standards       -> sample_lesson_plan.docx
codex-lesson-plan-generator/Lesson Plan Inbox -> DROP_PPTX_HERE.txt
codex-lesson-plan-generator/inputs      -> empty
codex-lesson-plan-generator/extracted   -> empty
codex-lesson-plan-generator/output      -> empty
```

## Top-Level Folder Inventory

Legend:

- Fixed: active folder had a CRITICAL or HIGH issue and it was resolved
- Pass: audited, no release-blocking issue found
- Backlog: audited, only MEDIUM or LOW follow-up remains
- Inventory: generated, temporary, or adjacent material; not treated as shipping pipeline source

### Core workflow folders

| Folder | Files | Status | Notes |
| --- | ---: | --- | --- |
| `.git` | 1456 | Inventory | Repository metadata. Not part of shipping logic. |
| `.pytest_cache` | 5 | Inventory | Test cache only. |
| `__pycache__` | 21 | Inventory | Python cache only. |
| `apps_script_export_area_triangles` | 4 | Pass | Earlier answer-leak fix already applied in companion code audit. |
| `codex-lesson-plan-generator` | 64 | Fixed | Folder contract, docs, and artifact hygiene cleaned. |
| `flagship-notebook-generator` | 18 | Pass | No new folder-level blocker found in this pass. |
| `lesson-plan-engine` | 8 | Backlog | Canonical folder exists, but active runtime still lives elsewhere. |
| `notebook-engine` | 23 | Backlog | Enhancement folder fixed; broader canonical migration still incomplete. |
| `notebook_engine_app_data` | 3430 | Inventory | Generated app data and run state. Keep out of release bundles. |
| `quiz-form-builder` | 7 | Pass | Canonical docs-only home, no blocker found. |
| `shared-assets` | 1 | Pass | Clean shared-asset root. |
| `tests` | 17 | Pass | Test folder healthy in current scope. |
| `tools` | 1 | Pass | No blocker found. |
| `validation-tools` | 8 | Pass | Correct canonical home for audit logs and validation docs. |

### Adjacent product, experiment, and preserved artifact folders

| Folder | Files | Status | Notes |
| --- | ---: | --- | --- |
| `architect-mode-skill` | 2 | Inventory | Skill workspace, not EduWonderLab pipeline runtime. |
| `area-concepts-notebook-ai` | 54 | Inventory | Generated artifact bundle. |
| `area-concepts-notebook-final` | 59 | Inventory | Generated artifact bundle. |
| `area-concepts-notebook-locked` | 57 | Inventory | Generated artifact bundle. |
| `area-concepts-notebook-run` | 57 | Inventory | Generated artifact bundle. |
| `axiom-adventure-game` | 4 | Inventory | Adjacent project. |
| `axiom-playable-game` | 4 | Inventory | Adjacent project. |
| `case_qlook` | 1 | Inventory | Quick-look artifact. |
| `editable-lesson-presentation-determine-the-volume-of-rectangular-prisms-1-notebook-build` | 55 | Inventory | Generated notebook build artifact. |
| `hybrid-notebook-final-20260410` | 68 | Inventory | Generated notebook artifact. |
| `lesson-to-game-studio` | 21197 | Inventory | Large adjacent app with `node_modules` and `.next`; exclude from EduWonderLab release audits. |
| `mcap-game-blueprint` | 2 | Inventory | Adjacent project. |
| `neft_contact_sheets` | 4 | Inventory | Media artifact folder. |
| `neft_contact_sheets_final` | 4 | Inventory | Media artifact folder. |
| `pptx_inspect_media` | 9 | Inventory | Inspection output. |
| `pptx_preview` | 14 | Inventory | Preview output. |
| `pptx_preview_v16` | 0 | Inventory | Empty preview folder. |
| `pptx_qlook` | 1 | Inventory | Quick-look artifact. |
| `ready_png_export` | 0 | Inventory | Empty export folder. |
| `ready_qlook` | 1 | Inventory | Quick-look artifact. |
| `ready_qlook2` | 1 | Inventory | Quick-look artifact. |
| `student-notebook` | 3 | Inventory | Skill/reference folder, not the active notebook runtime. |
| `triangle-notebook-lock-check` | 57 | Inventory | Validation artifact folder. |
| `unit9_qlook` | 1 | Inventory | Quick-look artifact. |

### Smoke and validation run folders

| Folder | Files | Status | Notes |
| --- | ---: | --- | --- |
| `copyedit-lock-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-activity-fix-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-engine-html-smoke` | 8 | Inventory | Smoke output folder. |
| `notebook-engine-pptx-only-smoke` | 7 | Inventory | Smoke output folder. |
| `notebook-engine-premium-smoke` | 7 | Inventory | Smoke output folder. |
| `notebook-engine-rigorous-smoke` | 7 | Inventory | Smoke output folder. |
| `notebook-engine-smoke` | 7 | Inventory | Smoke output folder. |
| `notebook-engine-standards-smoke` | 8 | Inventory | Smoke output folder. |
| `notebook-exact-template-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-lock-smoke` | 72 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke` | 3 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v10` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v11` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v12` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v13` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v14` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v15` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v16` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v2` | 3 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v3` | 3 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v4` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v5` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v6` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v7` | 57 | Inventory | Smoke output folder. |
| `notebook-flagship-upgrade-smoke-v9` | 57 | Inventory | Smoke output folder. |
| `notebook-inspect-no-review` | 56 | Inventory | Smoke output folder. |
| `notebook-launcher-smoke` | 3 | Inventory | Smoke output folder. |
| `notebook-model-lock-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-objective-fix-smoke` | 72 | Inventory | Smoke output folder. |
| `notebook-open-workspace-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-plan-inspect` | 55 | Inventory | Smoke output folder. |
| `notebook-practice-tips-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-premium-lock-smoke` | 83 | Inventory | Smoke output folder. |
| `notebook-problem-boxes-smoke` | 57 | Inventory | Smoke output folder. |
| `notebook-quality-fix-smoke` | 72 | Inventory | Smoke output folder. |
| `notebook-template-baseline-smoke` | 54 | Inventory | Smoke output folder. |
| `notebook-template-baseline-smoke-v2` | 56 | Inventory | Smoke output folder. |
| `notebook-template-baseline-smoke-v3` | 56 | Inventory | Smoke output folder. |
| `notebook-template-baseline-smoke-v4` | 56 | Inventory | Smoke output folder. |
| `notebook-template-inspect-no-review` | 56 | Inventory | Smoke output folder. |
| `notebook-template-plan-inspect` | 55 | Inventory | Smoke output folder. |
| `notebook-template-plan-inspect-v2` | 55 | Inventory | Smoke output folder. |
| `notebook-tpt-workbook-smoke` | 55 | Inventory | Smoke output folder. |
| `notebook-tpt-workbook-smoke-v2` | 56 | Inventory | Smoke output folder. |
| `notebook-tpt-workbook-smoke-v3` | 56 | Inventory | Smoke output folder. |
| `notebook-tpt-workbook-smoke-v4` | 57 | Inventory | Smoke output folder. |
| `problem-fidelity-lock-smoke` | 57 | Inventory | Smoke output folder. |

### Temporary and diagnostic folders

| Folder | Files | Status | Notes |
| --- | ---: | --- | --- |
| `tmp_area_triangles_images` | 20 | Inventory | Temp diagnostic output. |
| `tmp_core_inspect` | 54 | Inventory | Temp diagnostic output. |
| `tmp_deck_inspect` | 54 | Inventory | Temp diagnostic output. |
| `tmp_edu_nb_patch` | 2 | Inventory | Temp patch workspace. |
| `tmp_inbox_2_4_case` | 78 | Inventory | Temp validation bundle. |
| `tmp_inbox_2_4_case_v2` | 226 | Inventory | Temp validation bundle. |
| `tmp_inbox_2_4_case_v3` | 302 | Inventory | Temp validation bundle. |
| `tmp_inbox_2_4_case_v4` | 151 | Inventory | Temp validation bundle. |
| `tmp_inbox_runner_case` | 10 | Inventory | Temp validation bundle. |
| `tmp_inbox_runner_case_v2` | 10 | Inventory | Temp validation bundle. |
| `tmp_inbox_runner_case_v3` | 10 | Inventory | Temp validation bundle. |
| `tmp_inbox_runner_case_v4` | 10 | Inventory | Temp validation bundle. |
| `tmp_notebook_42_validate` | 10 | Inventory | Temp validation bundle. |
| `tmp_notebook_activity_bank_validate` | 3 | Inventory | Temp validation bundle. |
| `tmp_notebook_embed_drag_boxes` | 52 | Inventory | Temp render output. |
| `tmp_notebook_engagement_validate` | 3 | Inventory | Temp validation bundle. |
| `tmp_notebook_global_lock_42` | 52 | Inventory | Temp render output. |
| `tmp_notebook_global_lock_verify` | 5 | Inventory | Temp validation bundle. |
| `tmp_notebook_pairing_verify_run` | 52 | Inventory | Temp render output. |
| `tmp_notebook_pairing_verify_run_2` | 52 | Inventory | Temp render output. |
| `tmp_notebook_peer_discussion` | 3 | Inventory | Temp validation bundle. |
| `tmp_notebook_publisher_polish` | 3 | Inventory | Temp validation bundle. |
| `tmp_notebook_quality_check` | 56 | Inventory | Temp quality-check output. |
| `tmp_notebook_quality_lock` | 54 | Inventory | Temp quality-check output. |
| `tmp_notebook_quality_lock_regen` | 4 | Inventory | Temp quality-check output. |
| `tmp_notebook_runner_validate` | 10 | Inventory | Temp validation bundle. |
| `tmp_notebook_universal_lock_bad_guidance` | 52 | Inventory | Temp render output. |
| `tmp_notebook_universal_render` | 3 | Inventory | Temp render output. |
| `tmp_render_inspect` | 55 | Inventory | Temp render output. |
| `tmp_render_inspect2` | 54 | Inventory | Temp render output. |
| `tmp_render_inspect3` | 55 | Inventory | Temp render output. |
| `tmp_render_vocab_check` | 55 | Inventory | Temp render output. |
| `tmp_render_vocab_check2` | 55 | Inventory | Temp render output. |
| `tmp_render_vocab_check3` | 55 | Inventory | Temp render output. |
| `tmp_render_vocab_check4` | 55 | Inventory | Temp render output. |
| `tmp_text_size_rerender` | 3 | Inventory | Temp validation bundle. |
| `tmp_text_size_rerender_after_floor` | 3 | Inventory | Temp validation bundle. |
| `tmp_vocab_debug` | 54 | Inventory | Temp debug output. |
| `tmp_vocab_debug2` | 54 | Inventory | Temp debug output. |
| `tmp_vocab_debug3` | 54 | Inventory | Temp debug output. |

## Prioritized Remaining Backlog

### CRITICAL

None remaining in active folder contracts.

### HIGH

None remaining in active folder contracts.

### MEDIUM

1. Workspace clutter is high.
   - There are dozens of `notebook-*-smoke` and `tmp_*` folders plus large adjacent projects.
   - Exact fix: archive or prune historical smoke runs and temporary diagnostics before packaging or broad workspace-wide automation.

2. The preserved lesson-plan generator still contains two executable runtimes.
   - `codex-lesson-plan-generator/run.py`
   - `codex-lesson-plan-generator/run.js`
   - Both now honor the canonical `lesson-plan-engine/INBOX` and `lesson-plan-engine/OUTPUT` contract, but they still produce different lesson-plan artifacts for different workflows.
   - Exact fix: decide whether both remain supported or one becomes the single canonical runtime.

### LOW

1. Historical log files still mention old Documents-facing paths.
   - Example: `notebook-engine/enhancement/logs/RUN_LOG.md`
   - This is harmless if logs stay internal, but they should not be treated as current operator docs.

2. Canonical lesson-plan folders would benefit from one scrubbed reference bundle.
   - `lesson-plan-engine/examples/`
   - `lesson-plan-engine/gold-standards/`

## Remediation Addendum

After the initial folder audit, the following remaining contract issues were fixed:

- `codex-lesson-plan-generator/config/generator_config.json` now points the active Python runtime at `../lesson-plan-engine/INBOX` and `../lesson-plan-engine/OUTPUT`.
- `Generate Lesson Plan.command` now defaults to the same canonical folders and exports those values into the Node runtime.
- `codex-lesson-plan-generator/run.js` now defaults to the canonical folders instead of `~/EduWonderLab/...`.
- `codex-lesson-plan-generator/src/build_lesson_plan.py` now makes single-session runs honor the configured `lesson_plan.docx` filename.
- `codex-lesson-plan-generator/run.js` now writes multi-session files as `YYYY-MM-DD_Lesson_Plan_Session_1.docx` and `YYYY-MM-DD_Lesson_Plan_Session_2.docx`, matching the locked prompt contract.
- Legacy generated files were removed from:
  - `codex-lesson-plan-generator/output/`
  - `codex-lesson-plan-generator/extracted/`
- The duplicate sample deck was removed from:
  - `codex-lesson-plan-generator/inputs/slides/`

Current verified canonical output bundle:

- `lesson-plan-engine/OUTPUT/lesson_plan.json`
- `lesson-plan-engine/OUTPUT/lesson_plan.md`
- `lesson-plan-engine/OUTPUT/lesson_plan.docx`
- `lesson-plan-engine/OUTPUT/validation_report.md`
- `lesson-plan-engine/OUTPUT/extracted/*`
- `lesson-plan-engine/OUTPUT/2026-04-19_Lesson_Plan_Session_1.docx`
- `lesson-plan-engine/OUTPUT/2026-04-19_Lesson_Plan_Session_2.docx`

## Verification

### Search-based checks

```bash
rg -n '/Users/joelneft|/Users/|Chatgpt Notebook and Lesson plans|Joel explicitly changes the model' \
  codex-lesson-plan-generator notebook-engine/enhancement lesson-plan-engine notebook-engine \
  quiz-form-builder shared-assets flagship-notebook-generator validation-tools \
  --glob '!**/__pycache__/**' --glob '!**/*.pyc' --glob '!**/logs/**'
```

Status: no remaining matches in active folder docs and active folder runtime files

### Lesson-plan tests

```bash
python3 -m pytest -q tests
```

Working directory: `codex-lesson-plan-generator/`

Status: passed, 11 tests

### Syntax checks

```bash
python3 -m py_compile \
  notebook-engine/enhancement/src/run_enhancement_inbox.py \
  codex-lesson-plan-generator/run.py \
  codex-lesson-plan-generator/src/extract_slides.py \
  codex-lesson-plan-generator/src/utils.py
```

Status: passed

### Lesson-plan regeneration

```bash
python3 run.py --deck examples/sample_input_slides.pptx
```

Working directory: `codex-lesson-plan-generator/`

Status: passed

Generated or refreshed:

- `codex-lesson-plan-generator/extracted/raw_slide_text.json`
- `codex-lesson-plan-generator/extracted/normalized_lesson.json`
- `codex-lesson-plan-generator/output/lesson_plan.json`
- `codex-lesson-plan-generator/output/lesson_plan.md`
- `codex-lesson-plan-generator/output/validation_report.md`
- `codex-lesson-plan-generator/output/2026-04-19_Lesson_Plan_Session_1.docx`
