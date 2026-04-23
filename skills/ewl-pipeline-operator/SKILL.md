---
name: ewl-pipeline-operator
description: Use this skill for EduWonderLab notebook pipelines, lesson-plan pipelines, extraction-to-output workflows, artifact generation, audit-repair passes, and source-fidelity-preserving production work.
---

# EduWonderLab Pipeline Operator

## Purpose
Use this skill when working on EduWonderLab pipelines that generate lesson plans, student notebooks, or related classroom artifacts from source materials.

This skill is for repeatable production work where the correct pattern is:
extract -> map -> build -> audit -> repair -> final QA

Use this skill for:
- notebook pipeline runs or repairs
- lesson-plan pipeline runs or repairs
- extraction and mapping workflows
- source-to-artifact generation
- artifact cleanup with source fidelity
- QA and repair passes
- “make this production-ready” requests tied to the repo workflow

Do not use this skill for unrelated one-off coding tasks unless they directly affect the EduWonderLab production pipeline.

## Core operating rules
- Treat source materials as the source of truth.
- Preserve instructional intent, sequencing, examples, vocabulary, and task structure unless explicitly asked to redesign.
- Prefer safe, minimal, high-confidence edits over broad rewrites.
- Preserve the existing pipeline architecture unless a real defect requires structural change.
- Repair before rebuilding when the structure is substantially correct.
- Keep outputs final, production-ready, and immediately usable.

## Required execution pattern
For significant work, follow this sequence:

1. **Extract**
   - Identify the source files, scripts, templates, and outputs involved.
   - Inspect the existing pipeline before editing.
   - Confirm what the real inputs and expected outputs are.

2. **Map**
   - Build a compact internal map of source -> intermediate logic -> final artifact.
   - Identify where lesson fidelity, formatting, editability, or runtime behavior could break.
   - Preserve one-to-one traceability when possible.

3. **Build**
   - Make the smallest reliable change that moves the pipeline forward.
   - Match the repo’s current architecture and conventions.
   - Avoid introducing generic patterns that overwrite working repo-specific behavior.

4. **Audit**
   - Check source fidelity.
   - Check artifact completeness.
   - Check formatting/readability/editability.
   - Check for missing sections, placeholders, dangling references, or naming drift.
   - Check for runtime risk if code/scripts were changed.

5. **Repair**
   - Fix the highest-impact defects first.
   - Prefer targeted repairs over cosmetic churn.
   - Keep changes scoped and explainable.

6. **Final QA**
   - Do not claim success without saying what was actually verified.
   - Report risks, assumptions, and validation honestly.

## Source fidelity standards
Always preserve unless explicitly asked to change:
- lesson intent
- pacing logic
- problem progression
- objectives
- vocabulary
- activity structure
- assessment logic
- session distinctions
- editability expectations

Never invent:
- unsupported lesson content
- unsupported examples
- unsupported objectives
- unsupported teacher moves
- unsupported assessment content

## Artifact-specific checks

### For lesson plans
Verify:
- objectives align to source
- activity flow matches source lesson logic
- vocabulary and assessment stay aligned
- directions are practical and teacher-facing
- no vague filler replaces concrete lesson moves

### For student notebooks
Verify:
- editable structure is preserved
- slide/page logic is coherent
- text is readable
- spacing and hierarchy are strong
- source problems and information remain intact
- polish does not distort lesson meaning

## Code-change discipline
When code or scripts are involved:
- inspect the existing script before editing
- preserve working interfaces unless a real fix requires change
- prefer minimal diffs
- avoid speculative rewrites
- use existing tests/validators when present
- if no tests exist, do a concise manual validation pass

## Final response shape
For meaningful work, end with:
- what changed
- what was validated
- key risks or assumptions
- any required follow-up items for full reliability

## Refusal boundary
Do not expand scope into unrelated refactors, architectural redesigns, or folder reorganizations unless the user explicitly requests them or a real blocking defect requires them.
