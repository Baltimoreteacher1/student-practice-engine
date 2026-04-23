import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const PLACEHOLDER_MARKER = "<!-- NOTEBOOK_STATUS: placeholder -->";
const PUBLISHED_MARKER = "NOTEBOOK_STATUS: published";
const DATE_MARKER = "NOTEBOOK_DATE:";

const usage = [
  "Usage:",
  "  npm run publish:notebook -- <input-file> <YYYY-MM-DD>",
  "  npm run publish:notebook -- <YYYY-MM-DD>"
].join("\n");

const [, , firstArg, secondArg] = process.argv;
const usingImplicitInput = Boolean(firstArg && !secondArg);
const inputArg = usingImplicitInput ? null : firstArg;
const dateArg = usingImplicitInput ? firstArg : secondArg;

if (!dateArg) {
  fail(`${usage}\nMissing required arguments.`);
}

if (!isValidDate(dateArg)) {
  fail(`Invalid date "${dateArg}". Expected YYYY-MM-DD.`);
}

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const docsDir = path.join(repoRoot, "docs");
const latestPath = path.join(docsDir, "latest", "index.html");
const archiveRoot = path.join(docsDir, "archive");
const inputSearchDirs = [
  path.join(repoRoot, "html-incoming"),
  path.join(repoRoot, "incoming")
];
const inputPath = inputArg
  ? path.resolve(repoRoot, inputArg)
  : await findNewestIncomingHtml(inputSearchDirs);

await ensureDirectory(path.dirname(latestPath));
await ensureDirectory(archiveRoot);

const inputHtml = await readRequiredFile(inputPath);
const latestHtml = await readIfExists(latestPath);

if (latestHtml && isRealNotebook(latestHtml)) {
  const previousDate = extractNotebookDate(latestHtml);

  if (!previousDate) {
    fail("The current latest notebook is missing NOTEBOOK_DATE metadata, so it cannot be archived safely.");
  }

  const archivePath = path.join(archiveRoot, previousDate, "index.html");
  await ensureDirectory(path.dirname(archivePath));
  await fs.writeFile(archivePath, latestHtml, "utf8");
}

const publishedHtml = wrapPublishedNotebook(inputHtml, dateArg);
await fs.writeFile(latestPath, publishedHtml, "utf8");

const archiveDates = await listArchiveDates(archiveRoot);
const homepage = buildHomepage({
  archiveDates,
  latestDate: dateArg
});

await fs.writeFile(path.join(docsDir, "index.html"), homepage, "utf8");

process.stdout.write(
  [
    `Published notebook: ${path.relative(repoRoot, inputPath)}`,
    `Latest page: docs/latest/index.html`,
    `Homepage rebuilt with ${archiveDates.length + 1} dated link(s).`
  ].join("\n")
);

function fail(message) {
  process.stderr.write(`${message}\n`);
  process.exit(1);
}

function isValidDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }

  const [year, month, day] = value.split("-").map(Number);
  const candidate = new Date(Date.UTC(year, month - 1, day));

  return candidate.getUTCFullYear() === year
    && candidate.getUTCMonth() === month - 1
    && candidate.getUTCDate() === day;
}

async function ensureDirectory(targetPath) {
  await fs.mkdir(targetPath, { recursive: true });
}

async function readRequiredFile(targetPath) {
  try {
    const stat = await fs.stat(targetPath);

    if (!stat.isFile()) {
      fail(`Input path is not a file: ${targetPath}`);
    }

    return await fs.readFile(targetPath, "utf8");
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      fail(`Input file does not exist: ${targetPath}`);
    }

    throw error;
  }
}

async function readIfExists(targetPath) {
  try {
    return await fs.readFile(targetPath, "utf8");
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      return null;
    }

    throw error;
  }
}

async function findNewestIncomingHtml(directoryPaths) {
  for (const directoryPath of directoryPaths) {
    const entries = await fs.readdir(directoryPath, { withFileTypes: true }).catch((error) => {
      if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
        return [];
      }

      throw error;
    });

    const htmlFiles = await Promise.all(entries
      .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".html"))
      .map(async (entry) => {
        const fullPath = path.join(directoryPath, entry.name);
        const stat = await fs.stat(fullPath);
        return {
          fullPath,
          modifiedMs: stat.mtimeMs
        };
      }));

    if (htmlFiles.length > 0) {
      htmlFiles.sort((left, right) => right.modifiedMs - left.modifiedMs || left.fullPath.localeCompare(right.fullPath));
      return htmlFiles[0].fullPath;
    }
  }

  fail(`No HTML files were found in ${directoryPaths.join(" or ")}`);
}

function isRealNotebook(html) {
  return !html.includes(PLACEHOLDER_MARKER) && html.includes(PUBLISHED_MARKER);
}

function extractNotebookDate(html) {
  const match = html.match(/NOTEBOOK_DATE:\s*(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : null;
}

function wrapPublishedNotebook(html, date) {
  const cleaned = html.replace(/<!--\s*NOTEBOOK_STATUS:\s*(placeholder|published)\s*-->\n?/g, "")
    .replace(/<!--\s*NOTEBOOK_DATE:\s*\d{4}-\d{2}-\d{2}\s*-->\n?/g, "");

  return `<!-- NOTEBOOK_STATUS: published -->\n<!-- ${DATE_MARKER} ${date} -->\n${cleaned}`;
}

async function listArchiveDates(rootPath) {
  const entries = await fs.readdir(rootPath, { withFileTypes: true }).catch(() => []);

  return entries
    .filter((entry) => entry.isDirectory() && /^\d{4}-\d{2}-\d{2}$/.test(entry.name))
    .map((entry) => entry.name)
    .sort((left, right) => right.localeCompare(left));
}

function buildHomepage({ archiveDates, latestDate }) {
  const mergedDates = Array.from(new Set([latestDate, ...archiveDates])).sort((left, right) => right.localeCompare(left));

  const archiveHtml = mergedDates.length > 0
    ? mergedDates.map((date) => {
        const isLatest = date === latestDate;
        const href = isLatest ? "/latest/" : `/archive/${date}/`;
        const note = isLatest ? "Current live notebook" : "Archived notebook";

        return [
          '<li class="archive-item">',
          `  <a href="${href}">`,
          `    <span class="archive-label">${date}</span>`,
          `    <span class="archive-note">${note}</span>`,
          "  </a>",
          "</li>"
        ].join("\n");
      }).join("\n")
    : '<li class="empty">No notebook archives yet.</li>';

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Student Notebooks</title>
  <meta name="description" content="Daily HTML student notebooks and lesson experiences.">
  <style>
    :root {
      --navy: #102542;
      --navy-soft: #183a63;
      --teal: #1f8a8a;
      --cream: #f7f1e3;
      --paper: rgba(255, 255, 255, 0.92);
      --ink: #22303c;
      --muted: #5f6c78;
      --line: rgba(16, 37, 66, 0.12);
      --shadow: 0 24px 56px rgba(16, 37, 66, 0.12);
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      color: var(--ink);
      font-family: Georgia, "Times New Roman", serif;
      background:
        radial-gradient(circle at top left, rgba(31, 138, 138, 0.22), transparent 28%),
        radial-gradient(circle at bottom right, rgba(16, 37, 66, 0.12), transparent 24%),
        linear-gradient(180deg, #fffefb 0%, var(--cream) 100%);
    }

    .shell {
      width: min(1120px, calc(100% - 32px));
      margin: 0 auto;
      padding: 40px 0 64px;
    }

    .hero {
      background: linear-gradient(140deg, rgba(16, 37, 66, 0.96), rgba(24, 58, 99, 0.92));
      color: white;
      border-radius: 32px;
      padding: 40px 28px;
      box-shadow: var(--shadow);
      position: relative;
      overflow: hidden;
    }

    .hero::after {
      content: "";
      position: absolute;
      inset: auto -10% -18% auto;
      width: 320px;
      height: 320px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(31, 138, 138, 0.38), transparent 68%);
    }

    .eyebrow {
      margin: 0 0 14px;
      letter-spacing: 0.12em;
      text-transform: uppercase;
      font: 600 0.82rem/1.2 "Trebuchet MS", Arial, sans-serif;
      color: #9be4da;
    }

    h1 {
      margin: 0;
      max-width: 12ch;
      font-size: clamp(2.4rem, 7vw, 4.8rem);
      line-height: 0.98;
    }

    .hero p {
      max-width: 640px;
      font-size: 1.08rem;
      line-height: 1.75;
      color: rgba(255, 255, 255, 0.88);
    }

    .grid {
      display: grid;
      gap: 20px;
      margin-top: 24px;
    }

    @media (min-width: 860px) {
      .grid {
        grid-template-columns: minmax(0, 1.1fr) minmax(300px, 0.9fr);
      }
    }

    .card {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 28px;
      padding: 26px;
      box-shadow: var(--shadow);
      backdrop-filter: blur(10px);
    }

    h2 {
      margin-top: 0;
      margin-bottom: 12px;
      color: var(--navy);
      font-size: 1.5rem;
    }

    p, li {
      font-size: 1rem;
      line-height: 1.7;
    }

    .actions {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-top: 20px;
    }

    .button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 48px;
      padding: 0 18px;
      border-radius: 999px;
      text-decoration: none;
      font-weight: 700;
      font-family: "Trebuchet MS", Arial, sans-serif;
    }

    .button-primary {
      color: white;
      background: linear-gradient(135deg, var(--teal), #157171);
    }

    .button-secondary {
      color: var(--navy);
      background: rgba(16, 37, 66, 0.08);
    }

    .archive-list {
      list-style: none;
      padding: 0;
      margin: 18px 0 0;
      display: grid;
      gap: 12px;
    }

    .archive-item a {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 14px 16px;
      border-radius: 18px;
      color: var(--navy);
      text-decoration: none;
      background: rgba(255, 255, 255, 0.78);
      border: 1px solid rgba(16, 37, 66, 0.08);
    }

    .archive-item a:hover {
      transform: translateY(-1px);
      background: white;
    }

    .archive-label {
      font-weight: 700;
    }

    .archive-note {
      color: var(--muted);
      font-size: 0.92rem;
      font-family: "Trebuchet MS", Arial, sans-serif;
    }
  </style>
</head>
<body>
  <div class="shell">
    <section class="hero">
      <p class="eyebrow">EduWonderLab</p>
      <h1>Student notebooks, ready every day.</h1>
      <p>This GitHub Pages site keeps one current lesson experience live at all times and preserves each published notebook by date for easy classroom reuse.</p>
      <div class="actions">
        <a class="button button-primary" href="/latest/">Open Latest Notebook</a>
        <a class="button button-secondary" href="#archive">Browse Archive</a>
      </div>
    </section>

    <section class="grid">
      <article class="card">
        <h2>Latest lesson experience</h2>
        <p>The latest published notebook always lives at <code>/latest/</code>. This keeps daily sharing simple even as the archive grows over time.</p>
        <p><strong>Current published date:</strong> ${latestDate}</p>
      </article>

      <article class="card" id="archive">
        <h2>Archive</h2>
        <p>Published notebook links appear here by date, with the current live page always pinned to <code>/latest/</code>.</p>
        <ul class="archive-list">
${indent(archiveHtml, 10)}
        </ul>
      </article>
    </section>
  </div>
</body>
</html>`;
}

function indent(value, spaces) {
  const prefix = " ".repeat(spaces);
  return value.split("\n").map((line) => `${prefix}${line}`).join("\n");
}
