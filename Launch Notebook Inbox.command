#!/bin/zsh
set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORKSPACE_ROOT="${EDUWONDERLAB_WORKSPACE_ROOT:-$SCRIPT_DIR}"

RUNNER="$WORKSPACE_ROOT/notebook_folder_runner.py"
ARGS=("$@")
RUN_ARGS=("${ARGS[@]}")
INPUT_DIR="$HOME/Documents/Chatgpt/Notebook/Notebook Inbox"
OUTPUT_DIR="$HOME/Documents/Chatgpt/Notebook/Notebook Output"
ARCHIVE_DIR="$HOME/Documents/Chatgpt/Notebook/Notebook Archive"
HAS_INPUT_DIR=0
HAS_OUTPUT_DIR=0
HAS_ARCHIVE_DIR=0

index=1
while (( index <= ${#ARGS[@]} )); do
  arg="${ARGS[index]}"
  case "$arg" in
    --input-dir)
      HAS_INPUT_DIR=1
      if (( index < ${#ARGS[@]} )); then
        INPUT_DIR="${ARGS[index + 1]}"
      fi
      (( index += 2 ))
      ;;
    --input-dir=*)
      HAS_INPUT_DIR=1
      INPUT_DIR="${arg#--input-dir=}"
      (( index += 1 ))
      ;;
    --output-dir)
      HAS_OUTPUT_DIR=1
      if (( index < ${#ARGS[@]} )); then
        OUTPUT_DIR="${ARGS[index + 1]}"
      fi
      (( index += 2 ))
      ;;
    --output-dir=*)
      HAS_OUTPUT_DIR=1
      OUTPUT_DIR="${arg#--output-dir=}"
      (( index += 1 ))
      ;;
    --archive-dir)
      HAS_ARCHIVE_DIR=1
      if (( index < ${#ARGS[@]} )); then
        ARCHIVE_DIR="${ARGS[index + 1]}"
      fi
      (( index += 2 ))
      ;;
    --archive-dir=*)
      HAS_ARCHIVE_DIR=1
      ARCHIVE_DIR="${arg#--archive-dir=}"
      (( index += 1 ))
      ;;
    *)
      (( index += 1 ))
      ;;
  esac
done

if [[ "$HAS_INPUT_DIR" -eq 0 ]]; then
  RUN_ARGS+=("--input-dir" "$INPUT_DIR")
fi

if [[ "$HAS_OUTPUT_DIR" -eq 0 ]]; then
  RUN_ARGS+=("--output-dir" "$OUTPUT_DIR")
fi

if [[ "$HAS_ARCHIVE_DIR" -eq 0 ]]; then
  RUN_ARGS+=("--archive-dir" "$ARCHIVE_DIR")
fi

LOG_PATH="$OUTPUT_DIR/notebook_inbox_last_run.log"
SUMMARY_PATH="$OUTPUT_DIR/notebook_inbox_last_run.json"

mkdir -p "$INPUT_DIR" "$OUTPUT_DIR" "$ARCHIVE_DIR"
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
echo "Inbox: $INPUT_DIR"
echo "Output: $OUTPUT_DIR"
echo "Log: $LOG_PATH"
python3 "$RUNNER" "${RUN_ARGS[@]}"
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
    for result in results:
        notebook_path = result.get("session1_output") or result.get("session1")
        if notebook_path:
            print(f"Notebook file: {notebook_path}")
    if output_dir:
        print(f"Notebook output folder: {output_dir}")
    elif processed_count == 1 and results and results[0].get("export_dir"):
        print(f"Notebook package: {results[0]['export_dir']}")
elif status == "no_files":
    print("No PPTX files were found in the inbox.")
elif status == "failed":
    reason = payload.get("reason") or payload.get("error") or "unknown_error"
    reason_text = str(reason)
    if "insufficient_quota" in reason_text or "OpenAI API error 429" in reason_text:
        print("Failure reason: OpenAI API quota was exceeded. No notebook was generated.")
    else:
        print(f"Failure reason: {reason_text}")
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
if payload.get("output_dir"):
    print(payload["output_dir"])
elif len(results) == 1 and results[0].get("export_dir"):
    print(results[0]["export_dir"])
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
