# AGENTS.md

## Purpose
This repository contains EduWonderLab workflows for generating lesson plans, student notebooks, and related classroom artifacts from source materials.

Treat source files as the source of truth. Preserve instructional intent, sequencing, examples, vocabulary, and artifact fidelity unless explicitly asked to redesign.
Uploaded/source lesson slides, existing engine files, and current project files are source-of-truth inputs. Preserve math content, problem order, vocabulary, standards, objectives, examples, and lesson intent. Do not invent, simplify, omit, or reorder source content unless explicitly asked.

## Core operating rules
- Prefer safe, minimal, high-confidence edits over broad rewrites.
- Keep outputs final, production-ready, and immediately usable.
- Preserve the existing workflow, folder structure, and pipeline architecture unless explicitly asked to change them.
- Do not change notebook/lesson engine architecture, renderer structure, file layout, schemas, function names, or existing stable workflows unless the task explicitly requests an engine update.
- For extractor or fix requests, edit only the requested layer.
- Do not invent lesson content, examples, objectives, vocabulary, assessments, teacher moves, or student tasks not supported by source materials.
- When ambiguity is minor, choose the highest-utility interpretation that best preserves source fidelity and workflow stability.
- Repair before rebuilding when the current structure is substantially correct.
- Inspect existing files and nearby scripts before making changes so edits match the repo’s current architecture.

## Repository layout
- `Lesson Plan Inbox/` = source files for lesson-plan generation
- `Lesson Plan Output/` = final lesson-plan artifacts
- `Notebook Inbox/` = source files for notebook generation
- `Notebook Output/` = final notebook artifacts
- `Notebook Archive/` = archived notebook artifacts
- `Notebook Enhancement/` = enhancement prompts, helper files, upgrade assets

If additional workflow folders, scripts, templates, or validators exist, inspect them before editing.

## Default execution pattern
For significant work, follow this sequence:
extract -> map -> build -> audit -> repair -> final QA

Do not jump directly to final output when source extraction or mapping is needed.

For complex, high-risk, or multi-step work:
1. make a short plan first
2. inspect the relevant files before editing
3. keep diffs scoped and targeted
4. validate before claiming success
5. summarize changes, risks, and validation clearly at the end

## Source fidelity rules
- Preserve lesson intent, pacing logic, problem progression, and instructional sequence from source materials.
- Preserve editable output formats whenever applicable.
- Prefer source fidelity and usability over decorative improvement.
- Match the established benchmark style when revising an existing artifact family.
- Do not silently remove sections, objectives, vocabulary, examples, or activity structures unless clearly unsupported or explicitly requested.
- Do not replace a working repo-specific pattern with a generic pattern unless the change is necessary to fix a real defect.

## Editing and scope rules
- Keep changes localized when possible.
- Do not rename files, move folders, or change interfaces unless required for correctness or explicitly requested.
- Do not delete existing artifacts, archives, or helper assets unless explicitly asked or clearly fixing a duplicate/broken output.
- Preserve backward compatibility when working on shared scripts unless a breaking change is explicitly requested.
- Prefer the smallest reliable fix over the most ambitious rewrite.

## Destructive and external-impact boundary
High autonomy is allowed for normal edits, refactors, artifact generation, and tests. Still ask before:
- deleting large folders
- wiping generated artifacts
- changing authentication, credentials, or secrets
- installing global tools
- pushing to remote repositories
- changing unrelated files
- modifying stable engine architecture

## Code and automation rules
- Check platform and API correctness before finalizing code.
- Keep code paste-ready, syntax-safe, and consistent with the existing codebase.
- Avoid unsupported, brittle, or speculative API calls.
- Prefer targeted fixes unless a rebuild is clearly safer and more reliable.
- If a script depends on folders, files, triggers, libraries, services, environment variables, or manual setup, state that clearly.
- If tests, lint, build, or validation scripts already exist, use them instead of guessing.

### Google Apps Script host rules
When writing or editing Apps Script, verify the correct host/service:
- `FormsApp` for Google Forms
- `SpreadsheetApp` for Google Sheets
- `DocumentApp` for Google Docs
- `SlidesApp` for Google Slides
- `DriveApp` for Drive operations

If the requested script would fail in the wrong Apps Script environment, warn clearly before finalizing.

For Google Apps Script:
- validate API assumptions before writing code
- avoid unsupported or speculative methods
- include `validateEnvironment_` and `selfTest_` when useful
- prefer stable `FormsApp`, `SlidesApp`, `DocumentApp`, and `SpreadsheetApp` patterns
- do not use questionable methods such as `item.asQuizItem()` unless confirmed valid
- make requested quiz point values and required questions explicit

## Artifact quality rules

### Lesson plans
- Keep them teacher-facing, practical, polished, and ready to teach from.
- Preserve lesson flow, objective alignment, examples, checks for understanding, and exit task logic.
- Keep vocabulary, activities, and assessment aligned to the source lesson.
- Avoid vague filler language; keep directions and teacher moves concrete and usable.

### Student notebooks
- Keep them editable, readable, visually clean, and instructionally aligned.
- Maintain strong hierarchy, adequate spacing, consistent formatting, and student-friendly readability.
- Preserve source problems and information while improving clarity and polish.
- Avoid overcrowding, tiny text, inconsistent slide logic, and decorative elements that reduce usability.

## Validation and definition of done
Before finishing:
- when code is changed, verify syntax/build passes when a relevant check is available
- when generating PPTX/DOCX/HTML, confirm files exist at expected paths
- when editability is requested, confirm outputs remain in editable formats
- check for placeholder text
- check for missing sections or dangling references
- check for broken formatting, obvious overflow, malformed filenames, or layout issues
- check math notation, units, formulas, and labels for consistency with the source
- check for source-fidelity drift
- check source content was preserved
- check internal consistency across titles, objectives, vocabulary, examples, directions, and activity names
- check for obvious runtime risks in code or automation files
- run existing tests, lint, build, or validation commands when they exist
- if no formal checks exist, perform a concise manual validation pass and state exactly what was checked
- do not claim success without stating what was verified

## Preferred final response shape
For meaningful work, end with:
- files changed or created
- commands run
- pass/fail results
- unresolved issues, risks, or assumptions
- exact next action if manual user action is required

## Agent and documentation tool rules
- Use subagents only when the task is complex enough to benefit from parallel review, such as codebase audits, source-fidelity audits, bug hunts across many files, or comparing implementation against requirements.
- Do not spawn subagents for simple single-file edits.
- When a task depends on current external developer documentation, use configured MCP documentation tools if available before guessing.

## Skills boundary
Use `AGENTS.md` for durable repo behavior, constraints, and quality standards.

When a workflow is specialized, repeatable, or multi-step, prefer a Skill rather than adding more bulk here. Examples include:
- pipeline orchestration
- Apps Script host validation
- artifact fidelity auditing
- benchmark polish passes
- repo repair/debug workflows

## Stability rule
When in doubt, preserve:
1. source fidelity
2. editability
3. working architecture
4. validation honesty
5. clear final reporting
