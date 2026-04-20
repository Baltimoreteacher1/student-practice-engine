# EduWonderLab Pipeline Audit

Date: 2026-04-19
Workspace: `/Users/joelneft/.codex/workspaces/default`
Standard: pre-production review for a 10,000-teacher release

## Scope

This audit covered the active EduWonderLab lesson-plan and notebook pipeline, with special attention to:

- notebook generation contracts
- exact-template compressed notebook behavior
- launcher portability
- silent failure paths
- support-profile privacy
- source-fidelity gates
- student-facing answer leakage
- active verification paths and tests

Generated outputs, smoke artifacts, and historical one-off scripts were inventory-scanned and classified, but only active human-authored pipeline code was treated as production code for CRITICAL and HIGH remediation.

## Method

1. Read the active pipeline files end to end.
2. Ran targeted repo-wide searches for:
   - hardcoded machine-local paths
   - stale initials and teacher names
   - silent fallbacks
   - exact-template slide contract drift
   - broken image assumptions
3. Patched every CRITICAL and HIGH issue found in active pipeline code.
4. Re-ran syntax and test verification.

## CRITICAL And HIGH Issues Fixed

### 1. Compressed notebook contract was not enforced in the active engines

- File: `notebook_engine.py`
- Lines: `373-379`, `2431-2465`, `6488-6522`, `8425-8886`, `14220-14299`, `16020-16060`, `17618-17725`
- Severity: HIGH
- Problem:
  - The active exact-template notebook flow still reflected older multi-slide practice/review assumptions instead of the locked 6-slide compressed structure.
  - Validation and design-review gates still expected older guided/independent artifacts, which could cause false passes, false failures, or structurally wrong notebooks.
- Exact fix:
  - Replaced the exact-template sequence with the locked 6-slide order:
    - `learning_objectives`
    - `prior_session_review`
    - `vocabulary_table`
    - `guided_practice`
    - `interactive_activity`
    - `best_fit_review`
  - Rebuilt exact session assembly to emit that sequence only.
  - Updated objective, vocabulary, guided, interactive, and review renderers to match the compressed spec.
  - Updated review gates so exact-template notebooks are judged against the compressed contract instead of legacy slide expectations.

### 2. Hosted flagship notebook engine had the same compressed-template drift

- File: `flagship-notebook-generator/backend/notebook_engine.py`
- Lines: `372-375`, `675-690`, `2459-2463`, `5640-5665`, `7770-8204`, `12916-13395`, `14775-14805`, `15346-15357`, `16362-16525`
- Severity: HIGH
- Problem:
  - The hosted engine mirrored the same outdated exact-template assumptions, so the local and hosted generators could diverge.
- Exact fix:
  - Mirrored the 6-slide exact-template contract, renderer behavior, review gates, and image-lookup hardening from the local engine.

### 3. Folder runner silently degraded from live generation to offline generation

- File: `notebook_folder_runner.py`
- Lines: `167-180`
- Severity: CRITICAL
- Problem:
  - When live generation failed, the runner quietly fell back to offline mode.
  - That masked failures and violated the fail-loud output discipline in `AGENTS.md`.
- Exact fix:
  - Removed the silent fallback behavior.
  - `process_deck()` now raises a clear `RuntimeError` with the mode and underlying exception if generation fails.

### 4. Launchers used machine-specific workspace assumptions

- Files:
  - `Generate Lesson Plan.command`
  - `Launch Notebook Inbox.command`
- Lines:
  - `Generate Lesson Plan.command:4-9`
  - `Launch Notebook Inbox.command:4-7`
- Severity: HIGH
- Problem:
  - Launcher path resolution assumed a local layout instead of resolving from the script directory or environment.
  - That would break on other machines.
- Exact fix:
  - Switched both launchers to:
    - `SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"`
    - `WORKSPACE_ROOT="${EDUWONDERLAB_WORKSPACE_ROOT:-$SCRIPT_DIR}"`

### 5. Student-facing Apps Script rendered visible answers on worked-example slides

- Files:
  - `Code.gs`
  - `apps_script_export_area_triangles/Code.gs`
- Lines:
  - `Code.gs:362`
  - `apps_script_export_area_triangles/Code.gs:362`
- Severity: CRITICAL
- Problem:
  - `renderWorkedExampleSlide_()` displayed `slideSpec.answer` in a visible "Check" box.
  - That could leak answers directly to students.
- Exact fix:
  - Replaced visible answer rendering with a generic teacher-safe prompt:
    - `slideSpec.checkPrompt || 'Write the final labeled answer only after you finish the steps above.'`

### 6. Lesson-plan generator leaked teacher identity and student profile initials

- Files:
  - `codex-lesson-plan-generator/run.js`
  - `codex-lesson-plan-generator/src/apply_supports.py`
  - `codex-lesson-plan-generator/config/generator_config.json`
  - `codex-lesson-plan-generator/rules/support_rules.md`
  - `codex-lesson-plan-generator/tests/test_structure.py`
  - `codex-lesson-plan-generator/tests/test_supports_rule.py`
  - `codex-lesson-plan-generator/output/lesson_plan.md`
  - `codex-lesson-plan-generator/output/lesson_plan.json`
- Lines:
  - `run.js:78-103`, `122`, `1450-1462`, `1770-1803`
  - `apply_supports.py:8-15`, `32`
  - `generator_config.json:24-27`
  - `support_rules.md:5-11`
- Severity: CRITICAL
- Problem:
  - The pipeline used real-looking initials and a teacher-specific name in active config and output paths.
  - That created privacy risk and could surface student-coded support groups to end users.
- Exact fix:
  - Replaced initials with neutral support profile ids:
    - `Profile A` through `Profile G`
  - Replaced hardcoded teacher name default with:
    - `process.env.EDUWONDERLAB_TEACHER_NAME || "Teacher Name"`
  - Updated support logic, config, rules, tests, and checked-in generated outputs to use the neutral profile ids.

### 7. Active lesson-plan docs referenced a personal local DOCX path

- Files:
  - `codex-lesson-plan-generator/CLAUDE.md`
  - `codex-lesson-plan-generator/config/locked_rules.md`
  - `codex-lesson-plan-generator/rules/structure_rules.md`
- Severity: HIGH
- Problem:
  - These active docs referenced `/Users/joelneft/...` sample input paths.
  - That encourages brittle operator behavior and fails portability review.
- Exact fix:
  - Replaced personal paths with generic workspace-configured lesson-plan sample wording.

### 8. Image lookup assumed every source slide carried an `images` array

- Files:
  - `notebook_engine.py`
  - `flagship-notebook-generator/backend/notebook_engine.py`
- Lines:
  - `notebook_engine.py:16602-16613`
  - `flagship-notebook-generator/backend/notebook_engine.py:15346-15357`
- Severity: HIGH
- Problem:
  - `build_image_lookup()` indexed `slide["images"]` directly.
  - Sparse or partially structured slide payloads could crash instead of failing clearly upstream.
- Exact fix:
  - Hardened lookup construction to use `slide.get("images") or []` and skip invalid slide numbers safely.

### 9. Product docs still described the wrong notebook structure

- File: `NOTEBOOK_ENGINE_APP.md`
- Severity: HIGH
- Problem:
  - The app docs still described `10 to 14 slides per session`, which no longer matched the default compressed notebook contract.
- Exact fix:
  - Updated docs to describe the locked compressed 6-slide Session 1 structure by default, with Session 2 only when explicitly requested.

### 10. Regression test expected the old session structure

- File: `tests/test_notebook_session1_pivot.py`
- Severity: HIGH
- Problem:
  - The test still validated the old session structure, which would let the new compressed contract drift again or fail for the wrong reason.
- Exact fix:
  - Updated the test to assert the 6-slide Session 1 sequence.

## Patch-Style Diff Snippets For Fixed Issues

These are the effective before/after diffs for the CRITICAL and HIGH fixes. `git diff` is not reliable in this workspace because most project files are untracked, so the patch snippets below are the audit record.

### A. Exact-template slide contract

```diff
-EXACT_ESOL_TEMPLATE_SEQUENCE = [
-    "learning_objectives",
-    "prior_session_review",
-    "vocabulary_table",
-    "guided_practice",
-    "together_practice",
-    "independent_practice",
-]
+EXACT_ESOL_TEMPLATE_SEQUENCE = [
+    "learning_objectives",
+    "prior_session_review",
+    "vocabulary_table",
+    "guided_practice",
+    "interactive_activity",
+    "best_fit_review",
+]
```

```diff
-# legacy exact-template review expected together/independent practice pages
+# exact-template review now accepts the locked compressed pattern:
+# guided problem + interactive activity + best-fit interactive review
```

### B. Folder runner fail-loud behavior

```diff
 try:
     output_path = generate_premium_notebook(...)
 except Exception as exc:
-    if mode == "live":
-        output_path = generate_offline_notebook(...)
-        summary["fallback_mode"] = "offline"
-    else:
-        raise
+    mode_label = "live" if mode == "live" else "offline"
+    raise RuntimeError(
+        f"{mode_label.title()} notebook generation failed for "
+        f"{ppt_path.name}: {exc}"
+    ) from exc
```

### C. Portable launcher workspace root

```diff
-WORKSPACE_ROOT="/Users/joelneft/.codex/workspaces/default"
+SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
+WORKSPACE_ROOT="${EDUWONDERLAB_WORKSPACE_ROOT:-$SCRIPT_DIR}"
```

### D. Student answer leak removed from Apps Script

```diff
-addPromptBox_(slide, x, y, w, h, 'Check', slideSpec.answer || '', 16);
+addPromptBox_(
+  slide,
+  x,
+  y,
+  w,
+  h,
+  'Check Your Work',
+  slideSpec.checkPrompt || 'Write the final labeled answer only after you finish the steps above.',
+  16
+);
```

### E. Neutral support profiles and portable teacher name

```diff
-TEACHER_NAME: normalizeWhitespace(process.env.EDUWONDERLAB_TEACHER_NAME || "Joel Neft"),
+TEACHER_NAME: normalizeWhitespace(process.env.EDUWONDERLAB_TEACHER_NAME || "Teacher Name"),
```

```diff
-"M.E.": "...",
-"A.M.M.": "...",
+"Profile A": "...",
+"Profile G": "...",
```

```diff
-raise ValueError(f"Unsupported student initials in supports mapping: {sorted(invalid)}")
+raise ValueError(f"Unsupported support profile ids in supports mapping: {sorted(invalid)}")
```

### F. Hardened image lookup

```diff
-for image in slide["images"]:
+for image in slide.get("images") or []:
```

### G. Product docs and regression test aligned to compressed structure

```diff
-Session 1 and Session 2 decks typically produce 10 to 14 slides each.
+By default, the notebook engine produces the locked compressed 6-slide Session 1 notebook.
+Generate Session 2 only when the user explicitly requests it.
```

```diff
-self.assertGreaterEqual(len(slides), 10)
-self.assertIn("together_practice", roles)
-self.assertIn("independent_practice", roles)
+self.assertEqual(
+    roles,
+    [
+        "learning_objectives",
+        "prior_session_review",
+        "vocabulary_table",
+        "guided_practice",
+        "interactive_activity",
+        "best_fit_review",
+    ],
+)
```

## Prioritized Remaining Backlog

### CRITICAL

None remaining in the active audited pipeline.

### HIGH

None remaining in the active audited pipeline.

### MEDIUM

1. Legacy one-off scripts still contain hardcoded local paths and should be archived or parameterized before reuse.
   - `finalize_unit9_test_review_v5.py:24-25`
   - `publisher_polish_neft_ppt.py:9-10`
   - `fix_requested_slides_v4.py:9-10`
   - `build_two_session_notebooks.py:12-13`
   - `upgrade_notebooks_visuals_wida.py:12-13`
   - `upgrade_unit9_test_review.py:10-11`
   - Exact fix if promoted back to active use: convert to CLI args or env-driven paths.

2. `notebook_engine_app.py` still has a few silent recovery paths that should emit explicit warnings or fail earlier for operator clarity.
   - `83-88`: invalid settings JSON recovery path can suppress the original parse issue
   - `218-221`: malformed job manifest is skipped quietly
   - `459-462`: invalid CLI port value is ignored quietly

3. Checked-in generated binary artifact may still contain old support initials.
   - `codex-lesson-plan-generator/output/lesson_plan.docx`
   - Exact fix: regenerate or remove before release packaging.

4. Temporary generated PptxGenJS scripts in `tmp_edu_nb_patch/` violate the preferred locked rectangle/text pattern, but they are temp artifacts rather than active pipeline code.
   - Exact fix: delete or move out of workspace before packaging.

### LOW

1. `notebook_folder_runner.py:315` still carries a display field for `fallback_mode` even though fallback behavior was removed.
2. `notebook-engine/enhancement/src/run_enhancement_inbox.py:15` defaults to `~/Documents/...`; allow env override for cleaner portability.
3. Canonical migration remains incomplete for some legacy root scripts and docs. Track in `SPEC.md` and `TASKS.md` if those workflows are being kept.

## File-By-File Audit Ledger

Legend:

- Fixed: CRITICAL or HIGH issue found and resolved
- Pass: audited, no release-blocking issue found
- Backlog: audited, only MEDIUM or LOW follow-up remains
- Inventory: generated, binary, or historical artifact scanned but not treated as active production code

### Root workflow and docs

| File | Status | Notes |
| --- | --- | --- |
| `AGENTS.md` | Pass | Contract used for audit execution. |
| `CODEX_OPERATING_PLAYBOOK.md` | Pass | Reference playbook; no direct execution issues. |
| `README.md` | Pass | No active release blocker found. |
| `NOTEBOOK_ENGINE_APP.md` | Fixed | Updated to compressed 6-slide default. |
| `notebook_engine.py` | Fixed | Exact-template contract, renderers, reviews, image lookup. |
| `notebook_engine_app.py` | Backlog | No CRITICAL/HIGH issue found; medium silent-recovery cases remain. |
| `notebook_folder_runner.py` | Fixed | Removed silent live -> offline fallback. |
| `notebook_launchers.py` | Pass | No release blocker found in active use. |
| `Generate Lesson Plan.command` | Fixed | Portable workspace-root discovery. |
| `Launch Notebook Inbox.command` | Fixed | Portable workspace-root discovery. |
| `process_notebook_inbox.command` | Pass | No release blocker found. |
| `Code.gs` | Fixed | Removed student-visible answer leak. |
| `apps_script_export_area_triangles/Code.gs` | Fixed | Removed student-visible answer leak. |
| `tests/test_notebook_session1_pivot.py` | Fixed | Now enforces compressed sequence. |

### Canonical workspace folders

| File Or Folder | Status | Notes |
| --- | --- | --- |
| `lesson-plan-engine/` | Pass | Canonical folder present; no active release blocker surfaced here during code audit. |
| `notebook-engine/` | Pass | Canonical notebook area present. |
| `notebook-engine/enhancement/src/run_enhancement_inbox.py` | Backlog | Portable enough, but env override would improve durability. |
| `quiz-form-builder/` | Pass | No active release blocker surfaced in this audit path. |
| `validation-tools/` | Pass | Correct place for this audit log. |
| `shared-assets/` | Pass | No code-path release blocker found. |

### Hosted notebook generator

| File | Status | Notes |
| --- | --- | --- |
| `flagship-notebook-generator/backend/notebook_engine.py` | Fixed | Mirrored exact-template and hardening fixes. |
| `flagship-notebook-generator/backend/server.py` | Pass | No CRITICAL/HIGH issue found in current audit scope. |
| `flagship-notebook-generator/frontend/` | Pass | No active release blocker surfaced in current audit scope. |
| `flagship-notebook-generator/package.json` | Pass | No release blocker found. |
| `flagship-notebook-generator/next.config.*` | Pass | No release blocker found. |

### Lesson-plan generator

| File | Status | Notes |
| --- | --- | --- |
| `codex-lesson-plan-generator/run.js` | Fixed | Neutral profile ids, portable teacher name default. |
| `codex-lesson-plan-generator/run.py` | Pass | No CRITICAL/HIGH issue found. |
| `codex-lesson-plan-generator/src/apply_supports.py` | Fixed | Neutral profile ids, clearer validation error. |
| `codex-lesson-plan-generator/config/generator_config.json` | Fixed | Active support profiles updated. |
| `codex-lesson-plan-generator/rules/support_rules.md` | Fixed | Neutral profile ids. |
| `codex-lesson-plan-generator/CLAUDE.md` | Fixed | Removed personal sample-input path. |
| `codex-lesson-plan-generator/config/locked_rules.md` | Fixed | Removed personal sample-input path. |
| `codex-lesson-plan-generator/rules/structure_rules.md` | Fixed | Removed personal sample-input path. |
| `codex-lesson-plan-generator/tests/test_structure.py` | Fixed | Updated support profile expectations. |
| `codex-lesson-plan-generator/tests/test_supports_rule.py` | Fixed | Updated support profile expectations. |
| `codex-lesson-plan-generator/output/lesson_plan.md` | Fixed | Removed checked-in initials leak. |
| `codex-lesson-plan-generator/output/lesson_plan.json` | Fixed | Removed checked-in initials leak. |
| `codex-lesson-plan-generator/output/lesson_plan.docx` | Backlog | Binary artifact should be regenerated or removed. |

### Legacy and preserved scripts

| File | Status | Notes |
| --- | --- | --- |
| `finalize_unit9_test_review_v5.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |
| `publisher_polish_neft_ppt.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |
| `fix_requested_slides_v4.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |
| `build_two_session_notebooks.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |
| `upgrade_notebooks_visuals_wida.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |
| `upgrade_unit9_test_review.py` | Backlog | Hardcoded paths; preserved script, not active pipeline. |

### Temporary and generated artifacts

| File Or Folder | Status | Notes |
| --- | --- | --- |
| `tmp_edu_nb_patch/` | Inventory | Generated temp scripts; not active production code. |
| smoke output folders under root | Inventory | Inventory-scanned only. |
| binary `.pptx`, `.docx`, export folders | Inventory | Treated as outputs, not source code. |

## Verification

### Syntax and compile checks

```bash
python3 -m py_compile notebook_engine.py notebook_folder_runner.py notebook_engine_app.py flagship-notebook-generator/backend/notebook_engine.py
```

Status: passed

### Python unit tests

```bash
python3 -m unittest discover tests
```

Status: passed, 58 tests

### Lesson-plan generator tests

```bash
python3 -m pytest -q codex-lesson-plan-generator/tests
```

Status: passed, 9 tests

### Additional targeted checks

```bash
node -c codex-lesson-plan-generator/run.js
python3 -m json.tool codex-lesson-plan-generator/config/generator_config.json > /tmp/generator_config.pretty.json
```

Status: passed

### Targeted audit scans

```bash
rg -n "M\.E\.|R\.B\.|J\.C\.|R\.M\.G\.|J\.L\.|S\.D\.L\.P\.|A\.M\.M\.|Neft\.Alba|Unsupported student initials" codex-lesson-plan-generator
```

Status: no remaining matches in active lesson-plan generator source and checked-in text outputs

```bash
rg -n '/Users/joelneft|/Users/' notebook_engine.py notebook_folder_runner.py notebook_engine_app.py flagship-notebook-generator/backend/notebook_engine.py Code.gs apps_script_export_area_triangles/Code.gs codex-lesson-plan-generator/run.js codex-lesson-plan-generator/config/generator_config.json "Generate Lesson Plan.command" "Launch Notebook Inbox.command" NOTEBOOK_ENGINE_APP.md
```

Status: no remaining machine-local path matches in audited active pipeline files

## Residual Risk

The active audited pipeline no longer has known CRITICAL or HIGH blockers from this review. Remaining risk is concentrated in:

- preserved one-off scripts that should not be treated as reusable tooling without cleanup
- a stale generated binary output that should be regenerated or removed
- a few medium-grade silent recovery paths in `notebook_engine_app.py`

