#!/usr/bin/env bash
# Installs what a cloud run needs. Run from the repo: bash setup.sh
set -euo pipefail
cd "$(dirname "$0")"
# --ignore-installed: the image has Ubuntu's own cryptography/cffi (built for a
# different Python version). Without it, pip treats them as "already satisfied"
# and the azure import crashes with "No module named '_cffi_backend'". This
# installs PyPI copies into /usr/local/..., which comes first on sys.path.
python3 -m pip install -q --ignore-installed -r requirements.txt 2>/dev/null \
  || python3 -m pip install -q --ignore-installed --break-system-packages -r requirements.txt
# Fail here, loudly, if the import still doesn't work.
python3 -c "import azure.storage.blob, requests, openpyxl, pandas"

# LibreOffice recalculates the workbook's formulas for the zero-errors check.
# Best effort: if it cannot be installed the run continues and says so.
if ! command -v soffice >/dev/null 2>&1; then
  SUDO=""
  [ "$(id -u)" -ne 0 ] && command -v sudo >/dev/null 2>&1 && SUDO="sudo -n"
  { timeout 300 $SUDO apt-get update -qq \
      && DEBIAN_FRONTEND=noninteractive timeout 600 $SUDO apt-get install -y -qq \
           --no-install-recommends libreoffice-calc; } >/dev/null 2>&1 || true
fi
if command -v soffice >/dev/null 2>&1; then
  echo "SETUP OK (LibreOffice available)"
else
  echo "SETUP OK (LibreOffice NOT available)"
fi
