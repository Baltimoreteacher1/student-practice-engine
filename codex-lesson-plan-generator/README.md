# codex-lesson-plan-generator

## 1. Project purpose

`codex-lesson-plan-generator` is a deterministic, source-first local project that turns a teacher slide deck into a polished classroom lesson plan.

The implementation stays preserved in this folder, while the default live runtime now uses the canonical lesson-plan workspace folders in `../lesson-plan-engine/`.

Canonical runtime: `python3 run.py`

The generator is built for reliability and auditability:

- the slide deck is the source of truth
- standards are pulled only from the Learning Target slide
- the output always follows one fixed flagship lesson-plan structure
- supports are limited to approved SPED and ESOL rules
- validation runs before final reporting
- every output includes a slide-by-slide source fidelity appendix

The project is designed to be extendable later for batch processing, school-specific defaults, or a larger notebook-generator ecosystem without changing the core pipeline shape.

## 2. Folder structure

```text
codex-lesson-plan-generator/
├─ README.md
├─ AGENTS.md
├─ PROMPT.md
├─ requirements.txt
├─ run.py
├─ config/
│  ├─ generator_config.json
│  ├─ school_defaults.json
│  └─ locked_rules.md
├─ schemas/
│  ├─ lesson_plan_schema.json
│  └─ qa_schema.json
├─ templates/
│  ├─ lesson_plan_template.docx
│  └─ lesson_plan_template.md
├─ rules/
│  ├─ structure_rules.md
│  ├─ support_rules.md
│  └─ style_rules.md
├─ src/
│  ├─ extract_slides.py
│  ├─ detect_lesson_type.py
│  ├─ build_lesson_plan.py
│  ├─ apply_supports.py
│  ├─ validate_plan.py
│  ├─ render_docx.py
│  └─ utils.py
├─ tests/
│  ├─ test_structure.py
│  ├─ test_standards_rule.py
│  ├─ test_supports_rule.py
│  └─ fixtures/
│     └─ sample_extracted_raw_slide_text.json
└─ examples/
   ├─ sample_input_slides.pptx
   └─ sample_output_lesson_plan.docx
```

Default runtime folders:

```text
lesson-plan-engine/
├─ INBOX/
├─ OUTPUT/
│  ├─ lesson_plan.json
│  ├─ lesson_plan.md
│  ├─ lesson_plan.docx
│  ├─ validation_report.md
│  └─ extracted/
└─ ARCHIVE/
```

## 3. Installation

From a fresh local environment:

```bash
cd codex-lesson-plan-generator
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## 4. How to run

1. Place one `.pptx` teacher slide deck in `../lesson-plan-engine/INBOX/`, or pass an explicit `--deck` path.
2. Optionally place one agenda file in `../lesson-plan-engine/INBOX/agenda/`.
3. Run the canonical Python pipeline from this project folder:

```bash
python run.py
```

Optional explicit deck selection:

```bash
python run.py --deck examples/sample_input_slides.pptx
```

Optional config override:

```bash
python run.py --config config/generator_config.json --school-defaults config/school_defaults.json
```

## 5. Input expectations

The generator expects a teacher-facing PowerPoint deck with visible instructional text.

Best results come from decks that include:

- a clear title slide
- a Learning Target or Objective slide
- visible standard codes on the Learning Target slide
- worked examples, practice, discussion prompts, or exit-ticket prompts

Important input rules:

- by default, drop one deck into `../lesson-plan-engine/INBOX/`
- only one deck should be present in the configured slide-input folder unless `--deck` is used
- standards must appear directly on the Learning Target slide
- if the deck has no Learning Target slide or no standards on that slide, the run fails clearly

Optional agenda files can be `.txt`, `.md`, or `.json` and are used only to tighten the lesson overview and materials when helpful.

## 6. Output files

Successful runs generate:

- `../lesson-plan-engine/OUTPUT/lesson_plan.json`
- `../lesson-plan-engine/OUTPUT/lesson_plan.md`
- one or more DOCX files in `../lesson-plan-engine/OUTPUT/`
- `../lesson-plan-engine/OUTPUT/validation_report.md`

Intermediate extracted artifacts are also written to:

- `../lesson-plan-engine/OUTPUT/extracted/raw_slide_text.json`
- `../lesson-plan-engine/OUTPUT/extracted/normalized_lesson.json`
- `../lesson-plan-engine/OUTPUT/extracted/source_fidelity_map.json`

## 7. Validation checks

The pipeline validates:

- required lesson-plan structure
- standards source rule
- required section presence
- timing coherence
- appendix coverage
- support mapping compliance
- output file generation status

The validation report is written even when warnings are present. The run stops when a hard validation failure occurs.

## 7.1 Development checks

For contributor validation, run:

```bash
python3 -m pytest tests -q
```

## 8. How to adapt the template

Update the following files for local customization:

- `config/generator_config.json`
  - output paths
  - date override
  - active student supports
  - lesson duration override
- `config/school_defaults.json`
  - school name
  - default grade and subject
  - timing distribution
  - ESOL toggle
- `rules/*.md`
  - locked instructional and style expectations
- `templates/lesson_plan_template.md`
  - Markdown layout

The DOCX renderer uses `templates/lesson_plan_template.docx` as a style-safe base. If the file is missing, the renderer rebuilds a clean fallback template automatically.

## 9. Future extension ideas

- batch processing for multiple decks
- school-wide roster integration for automatic support selection
- agenda and notes fusion with stronger source ranking
- richer standards parsing for non-CCSS formats
- HTML export or LMS-ready export
- notebook-generator interoperability with shared extraction and validation layers
