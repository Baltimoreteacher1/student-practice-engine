# Student Practice Engine

## 1. What This Project Is
This is a React + Vite classroom activity hub for student practice pages. It is designed for one reusable Cloudflare Pages site where each unit and activity gets a clean student link.

## 2. How To Run Locally
```bash
cd /Users/joelneft/Documents/student-practice-engine
npm install
npm run dev
```

Open the local Vite URL shown in Terminal.

## 3. How To Add A New Activity
1. Add a new question-bank file in `src/data/activities/`.
2. Export it from `src/data/activities/index.js`.
3. Add a matching catalog entry in `src/data/activityCatalog.js`.
4. Use a simple slug, such as `ratios-review`, for the student URL.

Each catalog item needs:
- `slug`
- `title`
- `unit`
- `unitSlug`
- `grade`
- `activityType`
- `estimatedMinutes`
- `description`
- `standardsOrSkills`
- `questionBankId`
- `status`

## 4. How To Edit Questions
Edit the activity question bank in `src/data/activities/`.

Questions currently support:
- `multipleChoice`
- `shortAnswer`

Keep each question teacher-readable:
- `id`
- `type`
- `prompt`
- `choices` for multiple choice
- `correctAnswer`
- `explanation`
- `skill`
- `points`

## 5. How To Test Before Deploying
Run:
```bash
npm run build
```

Optional local preview:
```bash
npm run preview
```

## 6. How To Deploy To Cloudflare Pages
Use GitHub-connected Cloudflare Pages. After changes are committed and pushed to `main`, Cloudflare Pages should rebuild automatically.

Common command flow:
```bash
cd /Users/joelneft/Documents/student-practice-engine
npm install
npm run dev
npm run build
git add .
git commit -m "Update student practice engine"
git push
```

## 7. Cloudflare Settings
Cloudflare reminder:
- Framework preset: Vite
- Build command: npm run build
- Build output directory: dist
- Root directory: blank/default or /
- Production branch: main

Do not use VitePress. Do not use `npx wrangler deploy`. Do not add a Worker setup for this app.

## 8. Student Link Pattern
Examples:
- `/`
- `/units/statistics`
- `/units/expressions`
- `/activities/mean-median-mode-review`
- `/activities/exponents-warmup`
- `/activities/geometry-area-review`

Direct links are supported by `public/_redirects`:
```txt
/* /index.html 200
```

## 9. Common Troubleshooting
If a direct activity link shows a Cloudflare 404, check that `public/_redirects` exists and was included in the latest build.

If an activity card appears but the activity page is not found, check that `questionBankId` in `src/data/activityCatalog.js` matches the key exported from `src/data/activities/index.js`.

If student progress seems mixed between activities, check that storage keys include the activity slug in `src/utils/storage.js`.

If deployment does not update, confirm the latest commit was pushed to `main` and Cloudflare Pages is connected to this repository.
