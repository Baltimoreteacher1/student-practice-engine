# Notebook Engine App

`notebook_engine_app.py` is the local upload app for the notebook engine, and `index.html` is the actual frontend file it serves.

## What It Gives You

- A local browser page for uploading `.pptx` lesson decks
- A real `index.html` frontend instead of HTML embedded in Python
- Saved API key and default model settings
- Locked notebook defaults for a compressed 6-slide Session 1 notebook: Objectives + Session Map, Be Curious, Vocabulary + Reference Tool, Guided Problem, Interactive Activity, Best-Fit Interactive Review
- A bundled `activity_database.json` plus `activity_library.txt` for premium interactive notebook moves
- One-click generation of:
  - `Session 1 - Student Notebook.pptx`
  - `Session 2 - Student Notebook.pptx` when the workflow explicitly requests a second notebook session
- A folder runner for processing downloaded decks in batches or from an inbox folder
- A per-job folder that also keeps:
  - the uploaded source deck
  - `source_deck.json`
  - `notebook_plan.json`

## Start The App

```bash
python3 notebook_engine_app.py
```

Then open:

```text
http://127.0.0.1:8765
```

## Double-Click Launch

You can also use:

```text
launch_notebook_engine.command
```

on macOS. The launcher starts the local server and opens the app in your browser automatically.

## Fast Folder Workflow

If you want a lower-friction download-to-notebooks workflow, use:

```bash
python3 notebook_folder_runner.py
```

By default it uses these folders:

```text
~/Documents/Chatgpt Notebook and Lesson plans/Notebook Inbox
~/Documents/Chatgpt Notebook and Lesson plans/Notebook Output
~/Documents/Chatgpt Notebook and Lesson plans/Notebook Archive
```

The first time you run the app or folder runner, it also drops a clickable macOS launcher into:

```text
~/Documents/Chatgpt Notebook and Lesson plans/Notebook Inbox/Launch Notebook Inbox.command
```

Drop one or more `.pptx` files into `Notebook Inbox`, then double-click that launcher or run the command manually. It will:

- generate the default compressed Session 1 notebook for each deck
- generate Session 2 only when the workflow or prompt explicitly requests a second session
- copy the finished notebooks into `Notebook Output/<deck-name>-notebooks/`
- move the original source deck into `Notebook Archive`
- write `Notebook Output/notebook_inbox_last_run.log`
- write `Notebook Output/notebook_inbox_last_run.json`

You can also double-click:

```text
process_notebook_inbox.command
```

Optional watch mode:

```bash
python3 notebook_folder_runner.py --watch
```

That keeps polling the inbox folder for new `.pptx` files.

## Where Files Go

The app stores its files here:

```text
notebook_engine_app_data/
```

Inside that folder:

- `config.json` stores the saved API key, model, and default guidance
- `runs/<job-id>/` stores each notebook generation job and its outputs
- `activity_database.json` stores the structured interactive activity database
- `activity_library.txt` stores the text fallback activity list

## Frontend File

The browser UI lives in:

```text
index.html
```

The Python app serves that file at `http://127.0.0.1:8765/`.

## Notes

- Best with `.pptx` decks that contain actual text and embedded images
- Offline mode is only for smoke-testing the pipeline without the API
- The folder runner now ignores placeholder API keys from the shell and falls back to the saved config key
- The config loader auto-recovers malformed `config.json` files when possible
- The app is fully local; it does not depend on Flask, FastAPI, or Streamlit
