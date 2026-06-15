#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

"$ROOT_DIR/scripts/build_macos_bundle.sh"

BUILD_NO_AI="${PDF_PPTX_BUILD_NO_AI:-0}"
DELIVERABLE_SUFFIX=""
if [[ "$BUILD_NO_AI" == "1" || "$BUILD_NO_AI" == "true" ]]; then
  DELIVERABLE_SUFFIX="-no-ai"
fi

STAGING_DIR="$(mktemp -d "${TMPDIR:-/tmp}/pdf-to-pptx-dmg.XXXXXX")"
cleanup() {
  rm -rf "$STAGING_DIR"
}
trap cleanup EXIT

cp -R "dist/PDFtoPPTXReference.app" "$STAGING_DIR/"
cp "RECIPIENT_MAC_README.txt" "$STAGING_DIR/"
if [[ -d "dist/PowerPoint Add-ins" ]]; then
  cp -R "dist/PowerPoint Add-ins" "$STAGING_DIR/"
fi
ln -s /Applications "$STAGING_DIR/Applications"

mkdir -p deliverables
rm -f "deliverables/PDFtoPPTXReference-macOS${DELIVERABLE_SUFFIX}.dmg"

hdiutil create \
  -volname "PDFtoPPTXReference" \
  -srcfolder "$STAGING_DIR" \
  -ov \
  -format UDZO \
  "deliverables/PDFtoPPTXReference-macOS${DELIVERABLE_SUFFIX}.dmg"

echo ""
echo "DMG complete."
echo "Send deliverables/PDFtoPPTXReference-macOS${DELIVERABLE_SUFFIX}.dmg to the Mac user."
