# Project Instructions: MCAP Grade 6 Practice Engine

This repository is the only workspace for student activity, HTML, React, Vite, and Cloudflare Pages practice activities.

Project path:
/Users/joelneft/Documents/mcap-grade-6-practice

Core Rules:
- Work only inside this project folder.
- Do not use /Users/joelneft/.codex/workspaces/default.
- Do not create or edit files outside this project unless I explicitly ask.
- Keep the app beginner-readable and teacher-friendly.
- Keep the structure simple.
- Do not add a database, backend, authentication, or teacher dashboard unless I explicitly request it.
- Use React + Vite patterns.
- Keep question banks easy for a non-developer teacher to edit.
- Use localStorage for student progress unless told otherwise.
- Keep student-facing directions clear, visual, accessible, and ESOL-friendly.
- Build classroom-ready activities that are serious, engaging, and easy for students to use.

Preferred Folder Structure:
- src/components
- src/data
- src/utils
- src/styles if helpful
- public only for static assets

Required Checks Before Major Changes:
- pwd
- ls
- npm run build

Required Checks After Changes:
- npm run build
- npm run preview when helpful

Cloudflare Pages Settings:
- Build command: npm run build
- Build output directory: dist
- Production branch: main

When Adding a New Activity:
- Use the existing reusable practice engine.
- Add or update a question bank in src/data.
- Reuse existing components when possible.
- Keep questions organized by topic and section.
- Include a final score/report screen.
- Make sure progress saving still works.
- Avoid overcomplicating the app.

When You Finish:
Always summarize:
1. What changed.
2. Which files matter most.
3. How I edit questions.
4. Exact commands I should run next.
5. Any known limitations.
