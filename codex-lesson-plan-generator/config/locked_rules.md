# Locked Lesson-Plan Rules

## Canonical runtime rule

- `run.py` is the canonical lesson-plan generator.

## Canonical folder rule

- Read lesson decks from `../lesson-plan-engine/INBOX/` unless `--deck` is provided.
- Write canonical artifacts to `../lesson-plan-engine/OUTPUT/`.
- Keep extracted artifacts under `../lesson-plan-engine/OUTPUT/extracted/`.

## Source rules

- The uploaded lesson slides are the instructional source of truth.
- Preserve source-slide fidelity.
- Do not invent lesson content beyond what is needed for instructional coherence.
- If content is missing in the deck, keep the structure and mark the item with `[INFERRED - VERIFY]`.

## Standards rule

- Standards must be pulled directly from the Learning Target slide only.
- Standards may not be inferred from other slides.

## Locked structure rule

Every canonical lesson plan session must include this exact structure in this exact order:

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

## Output rule

- Final outputs must be polished and teacher-ready.
- Canonical runs must generate JSON, Markdown, DOCX, and validation output together.
