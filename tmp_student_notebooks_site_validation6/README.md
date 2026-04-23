# Student Notebooks Site

This repository powers a GitHub Pages site for daily HTML student notebooks and interactive lesson experiences, and it also acts as the processing workspace for lesson source decks. The site publishes from `/docs` on the `main` branch, with one current notebook in `/docs/latest/` and dated links on the homepage.

## Folder structure
- `/slides-inbox/` stores the newest lesson source slides.
- `/notebook-output/` stores generated notebook slide decks.
- `/html-incoming/` stores generated HTML lesson files waiting to publish.
- `/incoming/` remains available as a backward-compatible manual HTML drop zone.
- `/docs/index.html` is the GitHub Pages homepage.
- `/docs/latest/index.html` is the current published notebook.
- `/docs/archive/YYYY-MM-DD/index.html` stores previous published notebooks by date.
- `/Process Lesson Source.command` runs the source-to-notebook-and-HTML workflow on macOS.
- `/scripts/process-lesson-source.mjs` finds the newest source deck, launches Codex to build both outputs, and publishes the HTML result.
- `/scripts/publish-notebook.mjs` publishes a notebook, archives the previous latest page when needed, and rebuilds the homepage.

## Process a new lesson source
1. Drop the lesson slides into `/slides-inbox/`.
2. Click `Process Lesson Source.command`, or run the processing script from Terminal.
3. The workflow creates a notebook slide deck in `/notebook-output/`, creates an HTML lesson in `/html-incoming/`, publishes that HTML lesson to `/docs/latest/`, archives the prior live lesson if needed, and updates the open PR.

Shortest daily command:

```bash
npm run process:lesson-source -- 2026-04-23
```

Exact command with an explicit source file:

```bash
npm run process:lesson-source -- slides-inbox/your-lesson-source.pptx 2026-04-23
```

## Publish an HTML file directly
1. Drop the ready-made HTML file into `/html-incoming/`, or use `/incoming/` for older manual workflows.
2. Run the publish command with just the publish date, or include an explicit HTML path.

Shortest direct publish command:

```bash
npm run publish:notebook -- 2026-04-23
```

Exact command with an explicit HTML path:

```bash
npm run publish:notebook -- html-incoming/your-lesson.html 2026-04-23
```

## Daily Codex prompt
Use this short prompt:

```text
Process today’s lesson from /slides-inbox/ with date YYYY-MM-DD, create the notebook slides and HTML lesson, publish the HTML, and open or update the PR.
```

## Slides-to-site Codex prompt
Use this short prompt:

```text
Turn these uploaded lesson slides into a polished notebook slide deck and a premium single-file HTML lesson, save them in /notebook-output/ and /html-incoming/, publish the HTML with date YYYY-MM-DD, and open or update the PR.
```
