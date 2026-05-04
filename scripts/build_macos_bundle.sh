#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

if [[ ! -d ".venv" ]]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

pyinstaller \
  --noconfirm \
  --clean \
  --onedir \
  --windowed \
  --name "PDFtoPPTXReference" \
  --add-data "app/static:app/static" \
  launch_app.py

echo ""
echo "Build complete."
echo "Send the folder at dist/PDFtoPPTXReference to the Mac user."
