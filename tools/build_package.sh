#!/usr/bin/env bash
# Rebuild entire case package — Participants/ + Facilitators/ only
set -e
cd "$(dirname "$0")/.."
if [ ! -d .venv ]; then
  python3 -m venv .venv
  .venv/bin/pip install -q -r tools/requirements.txt
fi
.venv/bin/python tools/rebuild_package.py
