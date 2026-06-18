#!/usr/bin/env bash
set -euo pipefail

PYTHON="/Users/norielgecolea/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"

if [ ! -x "$PYTHON" ]; then
  PYTHON="python3"
fi

exec "$PYTHON" server.py
