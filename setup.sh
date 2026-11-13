#!/usr/bin/env bash
# Installs Python deps for the cloud run. Run from the repo: bash setup.sh
set -euo pipefail
cd "$(dirname "$0")"
# --ignore-installed: the image has Ubuntu's own cryptography/cffi (built for a
# different Python version). Without it, pip treats them as "already satisfied"
# and the azure import crashes with "No module named '_cffi_backend'". This
# installs PyPI copies into /usr/local/..., which comes first on sys.path.
python3 -m pip install -q --ignore-installed -r requirements.txt 2>/dev/null \
  || python3 -m pip install -q --ignore-installed --break-system-packages -r requirements.txt
# Fail here, loudly, if the import still doesn't work.
python3 -c "import azure.storage.blob"
