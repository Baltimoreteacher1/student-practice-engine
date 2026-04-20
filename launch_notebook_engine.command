#!/bin/zsh
cd "$(dirname "$0")"
PORT="${NOTEBOOK_ENGINE_PORT:-8765}"

python3 notebook_engine_app.py "$PORT" &
SERVER_PID=$!

cleanup() {
  kill "$SERVER_PID" 2>/dev/null
}

trap cleanup EXIT INT TERM

sleep 2
open "http://127.0.0.1:${PORT}"

wait "$SERVER_PID"
