# Notebook Engine

`notebook_engine.py` is the reusable local path for generating production-ready student notebook bundles from an uploaded `.pptx` deck.

## What It Does

- Extracts text, notes, and embedded images from a source PowerPoint deck
- Saves a structured source file at `source_deck.json`
- Uses the OpenAI Responses API with strict JSON schema output to build `notebook_plan.json`
- Pulls from `activity_library.txt` so the planner can choose lesson-fit interactive notebook structures
- Renders:
  - `Session 1 - Student Notebook.pptx` on the default compressed notebook path
  - `Session 2 - Student Notebook.pptx` when the workflow explicitly requests a second session

## Why This Path

- The existing builder in this workspace is high quality, but it is lesson-specific and hardcoded to one source deck.
- PowerPoint parsing is reliable locally with `python-pptx`.
- OpenAI docs support structured outputs and image inputs through the Responses API, which is a good fit for turning extracted lesson content into a predictable notebook plan.
- OpenAI docs note that non-PDF files such as `.pptx` are processed as text-only inputs, without embedded images or charts, so extracting PPTX content locally and attaching selected images is the safer first engine.
- The activity library lets the model preserve the original lesson problems while upgrading slides into more premium, interactive TpT-style notebook experiences.

## Save Your API Key

Create a local `.env` file in this workspace:

```env
OPENAI_API_KEY=your_key_here
OPENAI_MODEL=gpt-5.4
```

`OPENAI_MODEL` is optional. The script defaults to `gpt-5.4`.

## Default Build Contract

Unless the prompt explicitly asks for a different structure, the default notebook path is the compressed 6-slide Session 1 sequence:

1. `Objectives + Session Map`
2. `Be Curious`
3. `Vocabulary + Reference Tool`
4. `Guided Problem`
5. `Interactive Activity`
6. `Best-Fit Interactive Review`

## Main Command

```bash
python3 notebook_engine.py run "/absolute/path/to/source-deck.pptx"
```

Optional:

```bash
python3 notebook_engine.py run "/absolute/path/to/source-deck.pptx" --output-dir "/absolute/path/to/output"
```

## Offline Smoke Test

This does not call the API. It uses a heuristic local plan so you can verify the pipeline end to end:

```bash
python3 notebook_engine.py run "/absolute/path/to/source-deck.pptx" --offline
```

## Separate Stages

```bash
python3 notebook_engine.py extract "/absolute/path/to/source-deck.pptx"
python3 notebook_engine.py plan "/absolute/path/to/source-deck.pptx"
python3 notebook_engine.py render "/absolute/path/to/output/notebook_plan.json" --deck "/absolute/path/to/output/source_deck.json"
```

## Current Limits

- Best with `.pptx` decks that contain real text and embedded images
- Does not yet ingest Google Slides URLs directly
- Does not yet rasterize arbitrary slide shapes into whole-slide images
- Image-only decks will need a later OCR or slide-export upgrade
- Activity selection is now guided by the local library and lesson extraction, but the renderer still uses family-based templates rather than one bespoke layout per activity name
