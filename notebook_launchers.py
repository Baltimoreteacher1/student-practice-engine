#!/usr/bin/env python3
"""Shared notebook Documents-folder paths and launcher helpers."""

from __future__ import annotations

import shlex
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DOCUMENTS_DIR = Path.home() / "Documents"
WORKFLOW_HOME = DOCUMENTS_DIR / "Chatgpt Notebook and Lesson plans"
DEFAULT_INBOX_DIR = WORKFLOW_HOME / "Notebook Inbox"
DEFAULT_OUTPUT_DIR = WORKFLOW_HOME / "Notebook Output"
DEFAULT_ARCHIVE_DIR = WORKFLOW_HOME / "Notebook Archive"
LESSON_PLAN_HOME = Path.home() / "EduWonderLab"
DEFAULT_LESSON_PLAN_INBOX_DIR = LESSON_PLAN_HOME / "watch_lessonplans"
DEFAULT_LESSON_PLAN_OUTPUT_DIR = LESSON_PLAN_HOME / "output_lessonplans"
INBOX_LAUNCHER_NAME = "Launch Notebook Inbox.command"
LESSON_PLAN_LAUNCHER_NAME = "Generate Lesson Plan.command"
LESSON_PLAN_HELP_NAME = "DROP_PPTX_HERE.txt"


def build_inbox_launcher_script(workspace_root: Path | None = None) -> str:
    root = (workspace_root or ROOT).resolve()
    root_quoted = shlex.quote(str(root))
    return f"""#!/bin/zsh
set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE_ROOT="$SCRIPT_DIR"

if [[ ! -f "$WORKSPACE_ROOT/notebook_folder_runner.py" ]]; then
  WORKSPACE_ROOT={root_quoted}
fi

RUNNER="$WORKSPACE_ROOT/notebook_folder_runner.py"
ARGS=("$@")
OUTPUT_DIR="$HOME/Documents/Chatgpt Notebook and Lesson plans/Notebook Output"

index=1
while (( index <= ${{#ARGS[@]}} )); do
  arg="${{ARGS[index]}}"
  case "$arg" in
    --output-dir)
      if (( index < ${{#ARGS[@]}} )); then
        OUTPUT_DIR="${{ARGS[index + 1]}}"
      fi
      (( index += 2 ))
      ;;
    --output-dir=*)
      OUTPUT_DIR="${{arg#--output-dir=}}"
      (( index += 1 ))
      ;;
    *)
      (( index += 1 ))
      ;;
  esac
done

LOG_PATH="$OUTPUT_DIR/notebook_inbox_last_run.log"
SUMMARY_PATH="$OUTPUT_DIR/notebook_inbox_last_run.json"

mkdir -p "$OUTPUT_DIR"
export PYTHONUNBUFFERED=1

if [[ -n "${{OPENAI_API_KEY:-}}" ]]; then
  case "${{OPENAI_API_KEY}}" in
    PASTE_*|*YOUR_API_KEY*|OPENAI_API_KEY)
      unset OPENAI_API_KEY
      ;;
  esac
fi

if [[ -z "${{SSL_CERT_FILE:-}}" ]]; then
  if [[ -f /etc/ssl/cert.pem ]]; then
    export SSL_CERT_FILE=/etc/ssl/cert.pem
  elif [[ -f /private/etc/ssl/cert.pem ]]; then
    export SSL_CERT_FILE=/private/etc/ssl/cert.pem
  fi
fi

echo "Running notebook inbox processor..."
echo "Log: $LOG_PATH"
python3 "$RUNNER" "${{ARGS[@]}}"
STATUS=$?
echo
echo "Latest log: $LOG_PATH"
echo "Latest summary: $SUMMARY_PATH"

if [[ -f "$SUMMARY_PATH" ]]; then
  python3 - "$SUMMARY_PATH" <<'PY'
import json
import sys
from pathlib import Path

summary_path = Path(sys.argv[1])
try:
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
except Exception:
    raise SystemExit(0)

status = payload.get("status", "unknown")
processed_count = int(payload.get("processed_count") or 0)
results = payload.get("results") or []
output_dir = payload.get("output_dir", "")

print(f"Run status: {{status}}")
print(f"Processed decks: {{processed_count}}")

if status == "completed":
    if processed_count == 1 and results and results[0].get("export_dir"):
        print(f"Notebook folder: {{results[0]['export_dir']}}")
    elif output_dir:
        print(f"Notebook folder: {{output_dir}}")
elif status == "no_files":
    print("No PPTX files were found in the inbox.")
elif status == "failed":
    reason = payload.get("reason") or payload.get("error") or "unknown_error"
    print(f"Failure reason: {{reason}}")
PY
fi

if [[ $STATUS -eq 0 && -f "$SUMMARY_PATH" ]]; then
  OPEN_TARGET="$(python3 - "$SUMMARY_PATH" <<'PY'
import json
import sys
from pathlib import Path

summary_path = Path(sys.argv[1])
try:
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
except Exception:
    raise SystemExit(0)

if payload.get("status") != "completed":
    raise SystemExit(0)

results = payload.get("results") or []
if len(results) == 1 and results[0].get("export_dir"):
    print(results[0]["export_dir"])
elif payload.get("output_dir"):
    print(payload["output_dir"])
PY
)"
  if [[ -n "$OPEN_TARGET" ]]; then
    echo "Opening: $OPEN_TARGET"
    if [[ -z "${{NOTEBOOK_NO_OPEN:-}}" ]]; then
      open "$OPEN_TARGET" >/dev/null 2>&1 || true
    fi
  fi
fi

exit $STATUS
"""


def ensure_inbox_launcher(
    *,
    workspace_root: Path | None = None,
    inbox_dir: Path = DEFAULT_INBOX_DIR,
) -> Path:
    target_dir = inbox_dir.expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    launcher_path = target_dir / INBOX_LAUNCHER_NAME
    launcher_path.write_text(build_inbox_launcher_script(workspace_root), encoding="utf-8")
    launcher_path.chmod(0o755)
    return launcher_path


def build_lesson_plan_launcher_script(workspace_root: Path | None = None) -> str:
    root = (workspace_root or ROOT).resolve()
    root_quoted = shlex.quote(str(root))
    watch_dir_quoted = shlex.quote(str(DEFAULT_LESSON_PLAN_INBOX_DIR))
    output_dir_quoted = shlex.quote(str(DEFAULT_LESSON_PLAN_OUTPUT_DIR))
    return f"""#!/bin/zsh
set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE_ROOT={root_quoted}
PROJECT_ROOT="$WORKSPACE_ROOT/codex-lesson-plan-generator"
WATCH_DIR="${{EDUWONDERLAB_WATCH_DIR:-{watch_dir_quoted}}}"
OUTPUT_DIR="${{EDUWONDERLAB_OUTPUT_DIR:-{output_dir_quoted}}}"

if [[ ! -f "$PROJECT_ROOT/run.js" ]]; then
  echo "Lesson plan generator was not found at: $PROJECT_ROOT"
  exit 1
fi

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js is required to run the lesson plan generator."
  exit 1
fi

if ! command -v npm >/dev/null 2>&1; then
  echo "npm is required to install lesson plan generator dependencies."
  exit 1
fi

mkdir -p "$WATCH_DIR"
mkdir -p "$OUTPUT_DIR"

cd "$PROJECT_ROOT"
if ! node -e "require.resolve('docx'); require.resolve('jszip')" >/dev/null 2>&1; then
  echo "Installing lesson plan dependencies..."
  npm install >/dev/null 2>&1 || npm install docx >/dev/null 2>&1
fi

echo "Running lesson plan generator..."
echo "Input folder: $WATCH_DIR"
echo "Lesson plan output: $OUTPUT_DIR"
node run.js "$@"
STATUS=$?
echo
echo "Lesson plan output: $OUTPUT_DIR"

if [[ $STATUS -eq 0 ]]; then
  if [[ -z "${{LESSON_PLAN_NO_OPEN:-}}" ]]; then
    open "$OUTPUT_DIR" >/dev/null 2>&1 || true
  fi
fi

exit $STATUS
"""


def build_lesson_plan_help_text() -> str:
    return """Drop one teacher slide deck (.pptx) into this folder, then double-click:

Generate Lesson Plan.command

If you want to run from Terminal instead:

node run.js

The generator reads from:

~/EduWonderLab/watch_lessonplans

Finished lesson plans are saved in:

~/EduWonderLab/output_lessonplans
"""


def ensure_lesson_plan_assets(
    *,
    workspace_root: Path | None = None,
    lesson_plan_dir: Path = DEFAULT_LESSON_PLAN_INBOX_DIR,
) -> dict[str, Path]:
    target_dir = lesson_plan_dir.expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    DEFAULT_LESSON_PLAN_OUTPUT_DIR.expanduser().resolve().mkdir(parents=True, exist_ok=True)

    launcher_path = target_dir / LESSON_PLAN_LAUNCHER_NAME
    launcher_path.write_text(build_lesson_plan_launcher_script(workspace_root), encoding="utf-8")
    launcher_path.chmod(0o755)

    help_path = target_dir / LESSON_PLAN_HELP_NAME
    help_path.write_text(build_lesson_plan_help_text(), encoding="utf-8")

    return {"launcher": launcher_path, "help": help_path}
