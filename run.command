#!/bin/zsh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

exec "$SCRIPT_DIR/venv/bin/python" "$SCRIPT_DIR/app.py"
