# Internal Project Brief

## Project purpose

Build a deterministic local lesson-plan generator whose canonical runtime is `run.py`, using the canonical `lesson-plan-engine/INBOX` and `lesson-plan-engine/OUTPUT` folders.

## Canonical run order

1. locate the most relevant lesson deck from the canonical inbox, unless `--deck` is provided
2. extract source slide text and notes
3. build a structured lesson extract
4. map the extract into session-specific lesson-plan data
5. apply approved SPED and ESOL supports
6. validate against the locked rules and schema
7. render JSON, Markdown, DOCX, and validation output together

## Locked workflow behavior

- Keep the canonical `lesson-plan-engine/INBOX` and `lesson-plan-engine/OUTPUT` contract stable.
- Use the teacher slide deck as the source of truth.
- Pull standards only from Learning Target slides.
- Preserve the source lesson flow, worked examples, practice, and vocabulary.
- Fail clearly when the deck cannot support a usable lesson plan.
- Generate all canonical artifacts together instead of partial output.

## Required lesson extract fields

Every lesson extract must capture, at minimum:

- deck title
- lesson topic
- session labels
- Learning Target slide text
- standards text exactly as shown
- source-aligned opening, modeled, guided, independent, and closure tasks
- vocabulary terms and definitions when present
- checks for understanding
- required visuals or diagrams
- source slide references for each major section

## Final outputs

- Teacher-facing output: DOCX lesson-plan files in `lesson-plan-engine/OUTPUT`
- Internal output: lesson-plan JSON package, Markdown package, source-fidelity map, extracted artifacts, validation report
