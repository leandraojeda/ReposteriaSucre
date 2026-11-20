#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  python3 -m venv .venv
fi
.venv/bin/python -m pip install -q -r requirements.txt
echo "Principal: http://localhost:5000 | Segunda PC: http://IP-DE-ESTA-PC:5000"
exec .venv/bin/python app.py
