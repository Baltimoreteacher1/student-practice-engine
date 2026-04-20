# EduWonderLab - Canonical Lesson Plan Contract

Reads PPTX from: `../lesson-plan-engine/INBOX/`

Writes lesson-plan artifacts to: `../lesson-plan-engine/OUTPUT/`

Do not change folder paths without updating the canonical lesson-plan engine docs and launcher together.

## Canonical runtime

- Primary runtime: `python3 run.py`
- The desktop launcher must invoke the primary Python runtime.

## Canonical outputs

Every successful canonical run must produce:

- `lesson_plan.json`
- `lesson_plan.md`
- `lesson_plan.docx` for a single-session run
- `validation_report.md`
- extracted artifacts in `OUTPUT/extracted/`

If more than one session is selected, the canonical runtime may emit dated session DOCX files instead of a single `lesson_plan.docx`.

## Canonical section structure

Each rendered session plan must include this exact 10-section structure in this exact order:

1. `Lesson Information`
2. `Standards and Learning Targets`
3. `Lesson Objective and Student Success Criteria`
4. `Materials and Preparation`
5. `Opening / Warm-Up / Launch`
6. `Mini-Lesson / Modeling / Concept Development`
7. `Guided Practice / Collaborative Learning`
8. `Independent Practice / Application / Stations`
9. `Closure / Exit Ticket / Assessment`
10. `Differentiation, SPED/ESOL Supports, and Teacher Notes`

## Source-of-truth rules

- The lesson PPTX is the instructional source of truth.
- Standards must come only from Learning Target slide extraction.
- Adapt content to the lesson slides without changing the canonical session structure.
- If required information is missing, fail clearly or mark the missing item with `[INFERRED - VERIFY]` rather than inventing unsupported content.

## Runtime contract

- The canonical launcher must preserve the `lesson-plan-engine/INBOX` -> `lesson-plan-engine/OUTPUT` flow.
- The canonical runtime must generate JSON, Markdown, DOCX, and validation output together.
- Terminal output must report the source deck and generated artifact paths.
