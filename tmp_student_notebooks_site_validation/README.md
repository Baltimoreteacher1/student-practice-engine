# Student Notebooks Site

This repository powers a GitHub Pages site for daily HTML student notebooks and interactive lesson experiences. The site publishes from `/docs` on the `main` branch, with one current notebook in `/docs/latest/` and dated links on the homepage.

## Folder structure
- `/incoming/` stores the next HTML notebook or lesson file waiting to be published.
- `/docs/index.html` is the GitHub Pages homepage.
- `/docs/latest/index.html` is the current published notebook.
- `/docs/archive/YYYY-MM-DD/index.html` stores previous published notebooks by date.
- `/scripts/publish-notebook.mjs` publishes a notebook, archives the previous latest page when needed, and rebuilds the homepage.

## Add a new notebook
1. Drop the new single-file HTML notebook into `/incoming/`.
2. Run the publish command with the file path and the publish date.
3. Open a PR with the generated changes.

Exact command:

```bash
npm run publish:notebook -- incoming/your-notebook.html 2026-04-23
```

## Add a new lesson HTML file from uploaded slides
1. Upload the lesson slides to Codex.
2. Ask Codex to build a premium single-file HTML lesson experience from the slides.
3. Have Codex save the finished HTML file into `/incoming/`.
4. Publish it with the same command:

```bash
npm run publish:notebook -- incoming/your-lesson.html 2026-04-23
```

## Daily Codex prompt
Use this short prompt:

```text
Publish today’s notebook from /incoming/ with date YYYY-MM-DD and open a PR.
```

## Slides-to-site Codex prompt
Use this short prompt:

```text
Turn these uploaded lesson slides into a premium single-file HTML lesson, save it in /incoming/, publish it with date YYYY-MM-DD, and open a PR.
```
