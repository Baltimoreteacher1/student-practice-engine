#!/bin/zsh
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
BACKEND_DIR="$PROJECT_DIR/backend"
DOWNLOADS_DIR="$HOME/Downloads"
HTML_FILE="$DOWNLOADS_DIR/Flagship Notebook Generator.html"
CONFIG_FILE="$DOWNLOADS_DIR/config.js"
LOG_FILE="/tmp/flagship-notebook-generator.log"

if [[ ! -f "$HTML_FILE" ]]; then
  HTML_FILE="$PROJECT_DIR/frontend/index.html"
fi

if [[ ! -f "$CONFIG_FILE" ]]; then
  printf 'window.API_BASE = "http://127.0.0.1:8787";\n' > "$CONFIG_FILE"
fi

if ! curl -s http://127.0.0.1:8787/api/health >/dev/null 2>&1; then
  cd "$BACKEND_DIR"
  nohup python3 server.py > "$LOG_FILE" 2>&1 &
  sleep 2
fi

open "$HTML_FILE"

echo "Flagship Notebook Generator opened."
echo "Backend log: $LOG_FILE"
