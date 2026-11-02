#!/bin/bash
# Build TDForestEdu.app + TDForestEdu.dmg trên macOS
# Cách chạy:  chmod +x build_mac.sh && ./build_mac.sh
set -e
cd "$(dirname "$0")"
APP=TDForestEdu
VER=1.0.0
ARCH=$(uname -m)

echo "==> Tạo môi trường ảo Python"
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "==> Đóng gói .app bằng PyInstaller"
rm -rf build dist
pyinstaller --noconfirm --clean TDForestEdu_mac.spec

echo "==> Ký ad-hoc (cho phép chạy trên Mac Apple Silicon)"
codesign --force --deep --sign - "dist/$APP.app"

echo "==> Tạo file .dmg"
STAGE=dist/dmg_stage
rm -rf "$STAGE"; mkdir -p "$STAGE"
cp -R "dist/$APP.app" "$STAGE/"
ln -s /Applications "$STAGE/Applications"   # để người dùng kéo-thả vào Applications
DMG="dist/${APP}_${VER}_mac_${ARCH}.dmg"
rm -f "$DMG"
hdiutil create -volname "$APP" -srcfolder "$STAGE" -ov -format UDZO "$DMG"
rm -rf "$STAGE"

echo ""
echo "XONG! File cài đặt: $DMG"
