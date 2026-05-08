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

MAC_APP_MACOS_DIR="dist/PDFtoPPTXReference.app/Contents/MacOS"
if [[ -f ".env.local" ]]; then
  cp ".env.local" "$MAC_APP_MACOS_DIR/.env.local"
  echo "Bundled project-local .env.local into the app package."
fi
if [[ -f ".env.local.example" ]]; then
  cp ".env.local.example" "$MAC_APP_MACOS_DIR/.env.local.example"
fi

echo ""
echo "Build complete."
echo "Send dist/PDFtoPPTXReference.app to the Mac user, or zip the dist folder if you want to include sidecar files beside the app."
