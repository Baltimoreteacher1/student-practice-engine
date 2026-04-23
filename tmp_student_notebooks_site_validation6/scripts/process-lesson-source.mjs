import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const usage = [
  "Usage:",
  "  npm run process:lesson-source -- <YYYY-MM-DD>",
  "  npm run process:lesson-source -- <source-file> <YYYY-MM-DD>"
].join("\n");

const [, , firstArg, secondArg] = process.argv;
const sourceArg = secondArg ? firstArg : null;
const dateArg = secondArg ?? firstArg;

if (!dateArg) {
  fail(`${usage}\nMissing required arguments.`);
}

if (!isValidDate(dateArg)) {
  fail(`Invalid date "${dateArg}". Expected YYYY-MM-DD.`);
}

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const slidesInboxDir = path.join(repoRoot, "slides-inbox");
const notebookOutputDir = path.join(repoRoot, "notebook-output");
const htmlIncomingDir = path.join(repoRoot, "html-incoming");
const latestPath = path.join(repoRoot, "docs", "latest", "index.html");

await ensureDirectory(slidesInboxDir);
await ensureDirectory(notebookOutputDir);
await ensureDirectory(htmlIncomingDir);

const sourcePath = sourceArg
  ? path.resolve(repoRoot, sourceArg)
  : await findNewestSlideSource(slidesInboxDir);

const sourceStat = await fs.stat(sourcePath).catch((error) => {
  if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
    fail(`Source slide file does not exist: ${sourcePath}`);
  }

  throw error;
});

if (!sourceStat.isFile()) {
  fail(`Source path is not a file: ${sourcePath}`);
}

const sourceName = path.basename(sourcePath, path.extname(sourcePath));
const slug = slugify(sourceName);
const notebookOutputPath = path.join(notebookOutputDir, `${dateArg}-${slug}-student-notebook.pptx`);
const htmlOutputPath = path.join(htmlIncomingDir, `${dateArg}-${slug}.html`);

const prompt = [
  `Process the lesson source at ${sourcePath}.`,
  "",
  "Goals:",
  `- Create a polished student notebook slide deck at ${notebookOutputPath}`,
  `- Create a premium single-file HTML lesson experience at ${htmlOutputPath}`,
  `- Publish the HTML lesson with: npm run publish:notebook -- ${path.relative(repoRoot, htmlOutputPath)} ${dateArg}`,
  "- Open or update the repository PR with a concise summary of the changes",
  "",
  "Requirements:",
  "- Treat the uploaded lesson slides as the source of truth",
  "- Preserve lesson intent, examples, vocabulary, practice flow, and exit ticket",
  "- Keep the notebook and HTML outputs aligned to each other",
  "- Use the repo instructions in AGENTS.md",
  "- Keep the HTML self-contained with internal CSS and JS only",
  "- Keep the notebook output editable and polished for classroom use",
  "- Save both output files at the exact paths above",
  "- Do not stop after planning; complete the edits, publishing step, and PR update"
].join("\n");

const codexArgs = [
  "exec",
  "--cd",
  repoRoot,
  "--sandbox",
  "danger-full-access",
  "--skip-git-repo-check",
  prompt
];

const exitCode = await runCommand("codex", codexArgs, { cwd: repoRoot });

if (exitCode !== 0) {
  process.exit(exitCode);
}

await assertExists(notebookOutputPath, "Notebook output");
await assertExists(htmlOutputPath, "HTML output");
await assertPublishedLatest(latestPath, dateArg);

process.stdout.write(
  [
    `Processed source: ${path.relative(repoRoot, sourcePath)}`,
    `Notebook output: ${path.relative(repoRoot, notebookOutputPath)}`,
    `HTML output: ${path.relative(repoRoot, htmlOutputPath)}`
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

async function findNewestSlideSource(directoryPath) {
  const allowedExtensions = new Set([".pptx", ".ppt", ".pdf"]);
  const entries = await fs.readdir(directoryPath, { withFileTypes: true }).catch((error) => {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      fail(`Slides inbox does not exist: ${directoryPath}`);
    }

    throw error;
  });

  const sourceFiles = await Promise.all(entries
    .filter((entry) => entry.isFile() && allowedExtensions.has(path.extname(entry.name).toLowerCase()))
    .map(async (entry) => {
      const fullPath = path.join(directoryPath, entry.name);
      const stat = await fs.stat(fullPath);
      return {
        fullPath,
        modifiedMs: stat.mtimeMs
      };
    }));

  if (sourceFiles.length === 0) {
    fail(`No slide source files were found in ${directoryPath}`);
  }

  sourceFiles.sort((left, right) => right.modifiedMs - left.modifiedMs || left.fullPath.localeCompare(right.fullPath));
  return sourceFiles[0].fullPath;
}

function slugify(value) {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-")
    || "lesson";
}

function runCommand(command, args, options) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      ...options,
      stdio: "inherit"
    });

    child.on("error", reject);
    child.on("close", (code) => resolve(code ?? 1));
  });
}

async function assertExists(targetPath, label) {
  try {
    const stat = await fs.stat(targetPath);

    if (!stat.isFile()) {
      fail(`${label} path is not a file: ${targetPath}`);
    }
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      fail(`${label} was not created: ${targetPath}`);
    }

    throw error;
  }
}

async function assertPublishedLatest(targetPath, date) {
  const html = await fs.readFile(targetPath, "utf8").catch((error) => {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      fail(`Published latest page was not created: ${targetPath}`);
    }

    throw error;
  });

  if (!html.includes(`NOTEBOOK_DATE: ${date}`)) {
    fail(`Published latest page does not include date ${date}: ${targetPath}`);
  }
}
