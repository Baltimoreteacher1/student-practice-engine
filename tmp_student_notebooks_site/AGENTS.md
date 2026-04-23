# AGENTS.md

## Repository purpose
This repository is the permanent GitHub Pages site and lesson-processing workspace for daily student notebook slide decks and HTML lesson experiences.

## Lesson source workflow
- When asked to process a lesson source, use the newest slide deck in `/slides-inbox/` unless the user specifies a different file.
- Create the notebook slide output in `/notebook-output/`.
- Create the generated HTML lesson in `/html-incoming/`.
- Keep the notebook and HTML outputs aligned to the same source deck.
- Use updated styles and polish, but do not drift from the source lesson.
- After generating the HTML lesson, publish it to the site and update the open pull request.

## Publish workflow
- When asked to publish a notebook or HTML lesson, use the newest HTML file in `/html-incoming/` first, then fall back to `/incoming/` only if needed, unless the user specifies a different file.
- Before replacing `/docs/latest/index.html`, archive the current latest page to `/docs/archive/YYYY-MM-DD/index.html`, but only if `/docs/latest/index.html` is a real notebook and not the placeholder.
- Then copy the new notebook to `/docs/latest/index.html`.
- Update `/docs/index.html` so the latest link works and the archive list includes the new date.
- Preserve any Pages-related files such as `CNAME` if present.
- Keep the notebook HTML intact unless a small path or metadata fix is required for publishing.
- Open a pull request with a concise summary of what changed.

## Lesson build workflow
- When building a new lesson experience, use the Math Mission Interactive Builder rules below.
- Treat uploaded lesson slides as the source of truth.
- Produce premium single-file HTML lesson experiences by default unless the user explicitly requests a different format.
- Produce a matching notebook slide deck when the request is to process a lesson source through the repo workflow.

## Math Mission Interactive Builder

### Mission
Convert uploaded lesson slides into premium single-file HTML student lesson experiences.

### Source of truth
- Uploaded lesson slides are the source of truth.
- Preserve lesson intent, examples, vocabulary, practice flow, and exit ticket.
- Do not hard-code topic content unless the slides themselves define it.

### Output default
- Single self-contained HTML file.
- Internal CSS and JS only.
- Browser-openable without a build step.
- Include local autosave and backend-ready submission hooks.

### Design default
- Navy, teal, and warm cream palette.
- Rounded card layout.
- Strong hierarchy.
- High readability.
- Premium classroom look.
- Minimal clutter.

### Lesson structure
- Map slide content into interactive sections such as Launch, Vocabulary, Learn, Try, Talk, Detect, Level Up, Reflect, and Exit.
- Follow the slides when they suggest a different structure.

### Student work persistence
- Always include student name.
- Always include class or period.
- Always include local autosave.
- Always include save status.
- Always include final submit.
- Always include a backend-ready payload structure for Google Apps Script or Google Sheets.

### Quality bar
- No flat worksheet feel.
- No image-only output.
- No missing sections.
- No unreadable layouts.
- No invented content that conflicts with the slides.
