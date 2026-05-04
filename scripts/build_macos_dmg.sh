#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

"$ROOT_DIR/scripts/build_macos_bundle.sh"

STAGING_DIR="$(mktemp -d "${TMPDIR:-/tmp}/pdf-to-pptx-dmg.XXXXXX")"
cleanup() {
  rm -rf "$STAGING_DIR"
}
trap cleanup EXIT

cp -R "dist/PDFtoPPTXReference.app" "$STAGING_DIR/"
cp "RECIPIENT_MAC_README.txt" "$STAGING_DIR/"
ln -s /Applications "$STAGING_DIR/Applications"

mkdir -p deliverables
rm -f "deliverables/PDFtoPPTXReference-macOS.dmg"

hdiutil create \
  -volname "PDFtoPPTXReference" \
  -srcfolder "$STAGING_DIR" \
  -ov \
  -format UDZO \
  "deliverables/PDFtoPPTXReference-macOS.dmg"

echo ""
echo "DMG complete."
echo "Send deliverables/PDFtoPPTXReference-macOS.dmg to the Mac user."
