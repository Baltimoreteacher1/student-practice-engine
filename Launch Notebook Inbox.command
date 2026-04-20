#!/bin/zsh
set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE_ROOT="${EDUWONDERLAB_WORKSPACE_ROOT:-$SCRIPT_DIR}"

RUNNER="$WORKSPACE_ROOT/notebook_folder_runner.py"
ARGS=("$@")
OUTPUT_DIR="$HOME/Documents/Chatgpt Notebook and Lesson plans/Notebook Output"

index=1
while (( index <= ${#ARGS[@]} )); do
  arg="${ARGS[index]}"
  case "$arg" in
    --output-dir)
      if (( index < ${#ARGS[@]} )); then
        OUTPUT_DIR="${ARGS[index + 1]}"
      fi
      (( index += 2 ))
      ;;
    --output-dir=*)
      OUTPUT_DIR="${arg#--output-dir=}"
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

if [[ -n "${OPENAI_API_KEY:-}" ]]; then
  case "${OPENAI_API_KEY}" in
    PASTE_*|*YOUR_API_KEY*|OPENAI_API_KEY)
      unset OPENAI_API_KEY
      ;;
  esac
fi

if [[ -z "${SSL_CERT_FILE:-}" ]]; then
  if [[ -f /etc/ssl/cert.pem ]]; then
    export SSL_CERT_FILE=/etc/ssl/cert.pem
  elif [[ -f /private/etc/ssl/cert.pem ]]; then
    export SSL_CERT_FILE=/private/etc/ssl/cert.pem
  fi
fi

echo "Running notebook inbox processor..."
echo "Log: $LOG_PATH"
python3 "$RUNNER" "${ARGS[@]}"
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

print(f"Run status: {status}")
print(f"Processed decks: {processed_count}")

if status == "completed":
    if processed_count == 1 and results and results[0].get("export_dir"):
        print(f"Notebook folder: {results[0]['export_dir']}")
    elif output_dir:
        print(f"Notebook folder: {output_dir}")
elif status == "no_files":
    print("No PPTX files were found in the inbox.")
elif status == "failed":
    reason = payload.get("reason") or payload.get("error") or "unknown_error"
    print(f"Failure reason: {reason}")
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
    if [[ -z "${NOTEBOOK_NO_OPEN:-}" ]]; then
      open "$OPEN_TARGET" >/dev/null 2>&1 || true
    fi
  fi
fi

exit $STATUS
