# Notebook Engine Publisher Audit

## Review Lens

This audit reviews the notebook system as if a product design, applied pedagogy, and systems engineering team were examining whether the output feels like professional curriculum publishing rather than a competent slide generator.

## Main Critiques

### Design Critiques

1. The layouts were structurally organized, but many pages still looked generated rather than art-directed.
   Symptoms: heavy borders, repeated card bands, overly blank covers, and cramped micro-activities in narrow panels.

2. The visual system was too mechanical.
   Symptoms: the same border treatment appeared on almost every surface, compact boards clipped titles, and the name/date footer looked like utility UI instead of publishing-grade page furniture.

3. The output sometimes preserved source fidelity without preserving editorial judgment.
   Symptoms: long source problems survived, but some companion copy stayed generic or awkward, which weakened the sense of premium authorship.

4. Vocabulary support could surface context words instead of lesson words.
   Symptoms: phrases like "Teaching Experience" or action labels like "Describe Data" could outrank more instructionally useful terms.

5. Topic labeling could drift.
   Symptoms: statistics lessons could inherit generic or incorrect practice subtitles such as "ratio problem" from broad keyword heuristics.

### Code Critiques

1. `notebook_engine.py` is doing too many jobs at once.
   Extraction, planning, rendering, HTML output, quality review, and CLI orchestration all live in one large file, which makes design work slower and regression riskier.

2. The deployed backend copy had drifted from the local engine.
   The backend engine in `flagship-notebook-generator/backend/` no longer matched the main renderer, so polish fixes were at risk of landing in one delivery path but not the other.

3. The quality gate was stronger on fidelity than on editorial quality.
   It could confirm that source problems survived while still allowing cramped activity headers, low-value vocabulary, and generic topic phrasing through.

4. A lot of visual behavior depended on local magic numbers instead of a small, explicit design system.
   This made it harder to improve overall polish with confidence.

## Updates Applied

1. Introduced a shared publisher style version marker in the render quality report.

2. Softened the shared card and header system.
   Changes: lighter borders, slimmer accent bars, calmer footer treatment, and a cleaner page header rhythm.

3. Rebuilt the cover support furniture.
   Changes: tighter cover spacing, better subtitle hierarchy, cleaner lesson-language row, and a more professional student-record footer.

4. Improved compact activity rendering.
   Changes: narrow activity boards now switch to a micro layout so titles and draggable pieces stop clipping.

5. Fixed statistics-topic inference and profile selection.
   Changes: median / dot-plot lessons now classify as data analysis rather than accidentally falling into ratio-like phrasing.

6. Tightened vocabulary filtering.
   Changes: low-value context phrases and lead-verb phrases are more aggressively rejected so the support layer favors lesson language over scene language.

7. Made problem workspaces more purposeful.
   Changes: practice, guided, and exit panels now emphasize solve, check, and explain moves instead of leaving generic blank mini-boxes.

8. Added a stricter publisher-rigor pass.
   Changes: student-facing copy now rewrites internal generator jargon, upgrades weak assessment phrasing, shortens problem subtitles so they do not clip into half-sentences, and renames interactive-practice sections with publishing-safe language.

9. Expanded validation and regression coverage for the publisher-polish fixes.
   Changes: plan review now checks session-level copy, render quality reports now track rendered copy issues, and regression tests cover jargon cleanup plus cleaner problem-subtitle fallbacks.

10. Added a structured peer-discussion layer to the notebook contract.
   Changes: worked-example, practice, challenge, turn-and-teach, and reflection pages now receive specific partner-talk prompts and question sets, while layout review still blocks crowded or placeholder discussion copy.

## Next Backlog

1. Split the engine into modules for extraction, planning, render components, and quality review.

2. Add image-based visual regression checks for representative slides rather than relying only on text and shape heuristics.

3. Move shared design tokens into one explicit theme object so new slide types inherit polish automatically.

4. Expand the quality report to score editorial copy, vocabulary usefulness, and compact-layout safety in addition to source fidelity.
