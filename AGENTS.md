# AGENTS.md

## Project Boundary
- Work only inside `/Users/joelneft/Documents/student-practice-engine`.
- Do not use `/Users/joelneft/Documents/mcap-grade-6-practice`.
- Do not use `/Users/joelneft/.codex/workspaces/default` for this project.

## Project Purpose
This is the reusable React + Vite student practice activity engine. It should stay as one repo and one Cloudflare Pages site where new classroom activities are added as pages/routes inside the existing app.

## Activity Rules
- Add new activities inside this app; do not create a new project for each activity.
- Keep activity metadata in `src/data/activityCatalog.js`.
- Keep question banks in `src/data/activities/`.
- Keep question banks teacher-editable and beginner-readable.
- Keep the app structure simple enough for a non-developer teacher to understand.
- Preserve existing working behavior unless a change is needed for the multi-page activity engine.

## Routing And Deployment
- Student routes are handled by the React app:
  - `/`
  - `/units/:unitSlug`
  - `/activities/:activitySlug`
- Cloudflare direct links are handled by `public/_redirects`.
- Use one Cloudflare Pages static deployment.
- Do not use VitePress.
- Do not use Wrangler Worker deployment.
- Do not add a database, authentication, or teacher dashboard unless explicitly requested later.

## Required Cloudflare Pages Settings
- Framework preset: Vite
- Build command: `npm run build`
- Build output directory: `dist`
- Root directory: blank/default or `/`
- Production branch: `main`

## Validation
- Always run `npm run build` before finishing.
- Do not manually edit the `dist` folder.
- Use `git --no-pager` or `GIT_PAGER=cat` for Git output.
