# Flagship Notebook Generator

A GitHub-ready project for turning uploaded lesson slides into two premium student notebook decks:

- `Session 1 - Student Notebook.pptx`
- `Session 2 - Student Notebook.pptx`

The frontend is designed for Cloudflare Pages. The PPTX generation happens in a Python backend that returns one ZIP bundle per run.

## Why This Architecture

This project keeps the browser app and the notebook generator separate on purpose:

- Cloudflare Pages is a strong fit for the static frontend.
- The generator itself still depends on Python and `python-pptx`, which is a better fit for a Python runtime.
- The backend is stateless, so you do not need a database just to generate and download notebook bundles.

OpenAI reference:

- As of March 16, 2026, OpenAI’s models page recommends `gpt-5.4` as the starting point for complex reasoning and coding workflows: [OpenAI Models](https://platform.openai.com/docs/models)
- The planner uses strict schema-shaped output, aligned with OpenAI’s structured outputs guidance: [Structured Outputs](https://platform.openai.com/docs/guides/structured-outputs)

Cloudflare reference:

- Cloudflare Pages Functions let you proxy `/api/*` routes from the Pages site to a separate backend: [Pages Functions Routing](https://developers.cloudflare.com/pages/functions/routing/)
- Pages environment variables can be used for that proxy setup: [Pages Bindings](https://developers.cloudflare.com/pages/functions/bindings/)

## Repo Layout

```text
flagship-notebook-generator/
  backend/
    activity_library.txt
    notebook_engine.py
    server.py
    requirements.txt
    Dockerfile
  frontend/
    index.html
    config.js
    _headers
  functions/
    api/[[path]].js
  .env.example
  .gitignore
  README.md
```

## What The User Gets

Every successful generation returns a ZIP file containing:

- `outputs/Session 1 - Student Notebook.pptx`
- `outputs/Session 2 - Student Notebook.pptx`
- `diagnostics/notebook_plan.json`
- `diagnostics/source_deck.json`
- `input/<uploaded deck>.pptx`
- `assets/` extracted source images when available

## Backend Setup

The backend is a simple Python HTTP server with one main generation endpoint:

- `GET /api/health`
- `POST /api/generate`

### Environment Variables

Create a `.env` file from `.env.example` or set these in your host:

```env
OPENAI_API_KEY=
OPENAI_MODEL=gpt-5.4
ALLOW_ORIGIN=*
PORT=8787
```

Notes:

- `OPENAI_API_KEY` is optional if users supply their own key from the frontend.
- `ALLOW_ORIGIN=*` is fine for development. In production, set it to your Cloudflare Pages domain if you call the backend directly from the browser.

### Run Locally

From the `backend/` folder:

```bash
python3 -m pip install -r requirements.txt
python3 server.py
```

That starts the API on `http://127.0.0.1:8787`.

### Deploy The Backend

This backend can be deployed anywhere that can run Python 3.11 plus `python-pptx`, including:

- Render
- Railway
- Fly.io
- Google Cloud Run
- a VPS

If your host supports Docker, point it at `backend/Dockerfile`.

If your host uses a start command instead, use:

```bash
python server.py
```

## Frontend Setup

The frontend is a static site intended for Cloudflare Pages.

### Option A: Same-Origin Pages Proxy

Use the included `functions/api/[[path]].js` proxy.

In Cloudflare Pages:

1. Connect the GitHub repo.
2. Set the project root to the repo root.
3. Leave the build command blank.
4. Set the build output directory to `frontend`.
5. Add a Pages environment variable named `BACKEND_BASE_URL`.
   Example: `https://your-backend-domain.com`

With that setup, the frontend can call same-origin `/api/*` routes and you can leave `frontend/config.js` as:

```js
window.API_BASE = "";
```

### Option B: Direct Browser-to-Backend

If you do not want the Cloudflare proxy, edit `frontend/config.js`:

```js
window.API_BASE = "https://your-backend-domain.com";
```

In that setup, make sure the backend `ALLOW_ORIGIN` value includes your Pages domain.

## Local Frontend Test

Serve the frontend from the repo root or from the `frontend/` directory. For example:

```bash
python3 -m http.server 8080 --directory frontend
```

Then open:

```text
http://127.0.0.1:8080
```

If you are not using the Cloudflare proxy locally, set `frontend/config.js` to:

```js
window.API_BASE = "http://127.0.0.1:8787";
```

## GitHub Upload Checklist

Push this entire folder to GitHub:

- `backend/`
- `frontend/`
- `functions/`
- `.env.example`
- `.gitignore`
- `README.md`

Do not commit:

- your real `.env`
- generated notebook bundles
- uploaded lesson decks

## Notes

- Best results come from `.pptx` decks with real text and embedded images.
- Image-only decks will still need an OCR or slide-rasterization upgrade later.
- The frontend stores the user’s API key in browser `localStorage`, not in a backend database.
