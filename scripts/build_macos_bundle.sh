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
MAC_APP_BUNDLE="dist/PDFtoPPTXReference.app"
INCLUDE_LOCAL_ENV="${PDF_PPTX_BUNDLE_LOCAL_ENV:-1}"
BUILD_NO_AI="${PDF_PPTX_BUILD_NO_AI:-0}"
if [[ "$BUILD_NO_AI" == "1" || "$BUILD_NO_AI" == "true" ]]; then
  cat > "$MAC_APP_MACOS_DIR/.env.local" <<'EOF'
PDF_PPTX_ENABLE_AI_QC=false
EOF
  echo "Bundled no-AI .env.local into the app package."
elif [[ "$INCLUDE_LOCAL_ENV" != "0" && "$INCLUDE_LOCAL_ENV" != "false" && -f ".env.local" ]]; then
  cp ".env.local" "$MAC_APP_MACOS_DIR/.env.local"
  echo "Bundled project-local .env.local into the app package."
fi
if [[ -f ".env.local.example" ]]; then
  cp ".env.local.example" "$MAC_APP_MACOS_DIR/.env.local.example"
fi

codesign --force --deep --sign - "$MAC_APP_BUNDLE"
echo "Re-signed app bundle after bundling env files."

ADDINS_DIR="dist/PowerPoint Add-ins"
rm -rf "$ADDINS_DIR"
if [[ "$BUILD_NO_AI" != "1" && "$BUILD_NO_AI" != "true" ]]; then
  mkdir -p "$ADDINS_DIR"

  copy_optional_addin() {
    local source_path="$1"
    local target_name="$2"
    if [[ -f "$source_path" ]]; then
      cp "$source_path" "$ADDINS_DIR/$target_name"
      echo "Bundled PowerPoint add-in asset: $target_name"
    fi
  }

  copy_optional_addin "${PDF_PPTX_IMAGEEXPORT_PPAM:-$HOME/Desktop/ImageExport.ppam}" "ImageExport.ppam"
  copy_optional_addin "${PDF_PPTX_REFERENCE_HELPER_PPAM:-$HOME/Desktop/PDF-PPT_Helper.ppam}" "PDF-PPT_Helper.ppam"
  copy_optional_addin "${PDF_PPTX_REFERENCE_HELPER_PPTM:-$HOME/Desktop/PDF-PPT_Helper.pptm}" "PDF-PPT_Helper.pptm"

  if [[ -f "support/powerpoint/SlideExport.bas" ]]; then
    cp "support/powerpoint/SlideExport.bas" "$ADDINS_DIR/SlideExport.bas"
  fi
  if [[ -f "support/powerpoint/README.md" ]]; then
    cp "support/powerpoint/README.md" "$ADDINS_DIR/PowerPoint-Addins-README.md"
  fi
else
  echo "Skipping PowerPoint add-in bundling for no-AI build."
fi

echo ""
echo "Build complete."
echo "Send dist/PDFtoPPTXReference.app to the Mac user, or zip the dist folder if you want to include sidecar files beside the app."
