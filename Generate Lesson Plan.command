#!/bin/zsh
set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE_ROOT="${EDUWONDERLAB_WORKSPACE_ROOT:-$SCRIPT_DIR}"
WATCH_DIR="${EDUWONDERLAB_WATCH_DIR:-$WORKSPACE_ROOT/lesson-plan-engine/INBOX}"
OUTPUT_DIR="${EDUWONDERLAB_OUTPUT_DIR:-$WORKSPACE_ROOT/lesson-plan-engine/OUTPUT}"

PROJECT_ROOT="$WORKSPACE_ROOT/codex-lesson-plan-generator"

if [[ ! -f "$PROJECT_ROOT/run.py" ]]; then
  echo "Lesson plan generator was not found at: $PROJECT_ROOT"
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 is required to run the lesson plan generator."
  exit 1
fi

if ! python3 - <<'PY' >/dev/null 2>&1
import importlib
for name in ("pptx", "docx", "jsonschema"):
    importlib.import_module(name)
PY
then
  echo "Python lesson plan dependencies are missing. Run:"
  echo "  pip install -r $PROJECT_ROOT/requirements.txt"
  exit 1
fi

mkdir -p "$WATCH_DIR"
mkdir -p "$OUTPUT_DIR"
mkdir -p "$WORKSPACE_ROOT/lesson-plan-engine/ARCHIVE"
export EDUWONDERLAB_WATCH_DIR="$WATCH_DIR"
export EDUWONDERLAB_OUTPUT_DIR="$OUTPUT_DIR"

cd "$PROJECT_ROOT"

echo "Running lesson plan generator..."
echo "Input folder: $WATCH_DIR"
echo "Lesson plan output: $OUTPUT_DIR"
python3 run.py "$@"
STATUS=$?
echo
echo "Lesson plan output: $OUTPUT_DIR"

if [[ $STATUS -eq 0 ]]; then
  if [[ -z "${LESSON_PLAN_NO_OPEN:-}" ]]; then
    open "$OUTPUT_DIR" >/dev/null 2>&1 || true
  fi
fi

exit $STATUS
