#!/bin/zsh
set -u

BASE_DIR="$(cd "$(dirname "$0")" && pwd)"
RUNNER="$BASE_DIR/src/run_enhancement_inbox.py"
INBOX_DIR="$BASE_DIR/INBOX"
OUTPUT_DIR="$BASE_DIR/OUTPUT"
SUMMARY_PATH="$OUTPUT_DIR/enhancement_inbox_last_run.json"

mkdir -p "$INBOX_DIR" "$OUTPUT_DIR" "$BASE_DIR/ARCHIVE"
export PYTHONUNBUFFERED=1

echo "Running notebook enhancement inbox..."
echo "Inbox: $INBOX_DIR"
echo "Output: $OUTPUT_DIR"
python3 "$RUNNER" --inbox-dir "$INBOX_DIR" --output-dir "$OUTPUT_DIR"
STATUS=$?
echo
echo "Summary: $SUMMARY_PATH"

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

print(f"Run status: {payload.get('status', 'unknown')}")
print(f"Processed bundles: {int(payload.get('processed_count') or 0)}")

results = payload.get("results") or []
if len(results) == 1:
    rendered_dir = results[0].get("rendered_dir")
    output_dir = results[0].get("output_dir")
    if rendered_dir:
        print(f"Rendered output: {rendered_dir}")
    elif output_dir:
        print(f"Polished bundle: {output_dir}")
elif payload.get("output_dir"):
    print(f"Output folder: {payload['output_dir']}")
PY
fi

if [[ $STATUS -eq 0 ]]; then
  OPEN_TARGET="$OUTPUT_DIR"
  if [[ -f "$SUMMARY_PATH" ]]; then
    OPEN_TARGET="$(python3 - "$SUMMARY_PATH" <<'PY'
import json
import sys
from pathlib import Path

summary_path = Path(sys.argv[1])
try:
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
except Exception:
    raise SystemExit(0)

results = payload.get("results") or []
if len(results) == 1:
    print(results[0].get("rendered_dir") or results[0].get("output_dir") or payload.get("output_dir") or "")
else:
    print(payload.get("output_dir") or "")
PY
)"
  fi
  if [[ -n "$OPEN_TARGET" ]]; then
    echo "Opening: $OPEN_TARGET"
    open "$OPEN_TARGET" >/dev/null 2>&1 || true
  fi
fi

exit $STATUS
