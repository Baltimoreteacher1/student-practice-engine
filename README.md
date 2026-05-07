# Grade 6 MCAP Math Practice App

This is a simple React + Vite student practice app for Grade 6 MCAP-style math review.

Students can enter their name, choose a section, answer multiple-choice questions, get instant feedback, save progress in the browser, and view a final score report.

No database, login system, or teacher dashboard is included yet. Progress is saved with `localStorage`, which means it stays on the same device and browser.

## What Students Can Do

- Type their name
- Choose one of five math sections
- Answer MCAP-style multiple-choice questions
- See instant feedback after each answer
- Return later on the same device and continue
- View a final score report
- Reset their progress

## Important Files

- `src/App.jsx` - controls the activity flow
- `src/main.jsx` - starts the React app
- `src/data/questionBanks.js` - main five-section question bank
- `src/data/meanMedianMode.js` - example reusable topic bank
- `src/components/ActivityShell.jsx` - student name screen, level map, and progress overview
- `src/components/QuestionCard.jsx` - question, feedback, and next button
- `src/components/ProgressBar.jsx` - simple progress display
- `src/components/ScoreReport.jsx` - final score report
- `src/utils/storage.js` - localStorage progress saving and reset
- `src/utils/scoring.js` - score calculations
- `src/styles/app.css` - visual design
- `public/_redirects` - helps Cloudflare Pages handle refreshes

## Run Locally

1. Open Terminal in this project folder:

   ```bash
   cd /Users/joelneft/.codex/workspaces/default
   ```

2. Install the project:

   ```bash
   npm install
   ```

3. Start the local app:

   ```bash
   npm run dev
   ```

4. Open the URL that Vite shows. It is usually:

   ```text
   http://localhost:5173/
   ```

## Edit Questions

Open this file:

```text
src/data/questionBanks.js
```

Each section contains a `questions` list. A question looks like this:

```js
{
  id: "ratios-1",
  skill: "Equivalent ratios",
  prompt: "Which ratio is equivalent to 3:5?",
  choices: ["6:10", "9:10", "12:15", "15:20"],
  answerIndex: 0,
  explanation: "3:5 is equivalent to 6:10 because both terms are multiplied by 2.",
}
```

Teacher notes:

- Keep every `id` unique.
- `choices` are the answer options students see.
- `answerIndex` starts counting at 0.
- Use `0` for choice A, `1` for choice B, `2` for choice C, and `3` for choice D.
- `explanation` is shown after the student answers.
- Keep the commas, brackets, and quotation marks in place.

## Swap Or Add Topic Banks

`src/data/meanMedianMode.js` shows a reusable topic bank. It is imported into `src/data/questionBanks.js` and added to Statistics like this:

```js
import { meanMedianModeQuestions } from "./meanMedianMode.js";

// inside the Statistics questions list
...meanMedianModeQuestions,
```

To add a new topic later, make another file in `src/data`, export an array of questions, import it in `questionBanks.js`, and spread it into the section where you want it.

## Build for Production

Run:

```bash
npm run build
```

This creates a production-ready folder named:

```text
dist
```

## Preview the Production Build

Run:

```bash
npm run preview
```

## Deploy to Cloudflare Pages

1. Put this project in a GitHub repository.
2. In Cloudflare, go to Workers & Pages.
3. Create a new Pages project.
4. Connect your GitHub repository.
5. Use these settings:

   ```text
   Framework preset: Vite
   Build command: npm run build
   Output directory: dist
   Root directory: leave blank unless this app is inside a subfolder
   ```

6. Deploy.

No environment variables are needed.

## Reset Saved Student Progress

Use the in-app `Reset` or `Reset Practice` button. For testing, you can also clear this site's browser data.
