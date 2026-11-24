#!/usr/bin/env bash
cd "$(dirname "$0")"
PY=${PYTHON:-python3}
exec $PY scripts/serve_dashboard.py --port "${PORT:-8899}" --host "${HOST:-127.0.0.1}"
