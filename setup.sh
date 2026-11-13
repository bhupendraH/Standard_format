#!/usr/bin/env bash
# Installs Python deps for the cloud run. Run from the repo: bash setup.sh
set -euo pipefail
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt 2>/dev/null \
  || python3 -m pip install -q --break-system-packages -r requirements.txt
