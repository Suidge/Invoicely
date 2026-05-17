#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
APP_NAME="Invoicely发票整理"
TMP_BASE="/private/tmp/invoicely_packaging"
VENV_DIR="$TMP_BASE/.venv"
DIST_TMP="$TMP_BASE/dist"
BUILD_TMP="$TMP_BASE/build"
SPEC_TMP="$TMP_BASE/spec"
CACHE_TMP="$TMP_BASE/cache"

rm -rf "$TMP_BASE"
mkdir -p "$DIST_TMP" "$BUILD_TMP" "$SPEC_TMP" "$CACHE_TMP"

python3 -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip
python -m pip install pyinstaller
python -m pip install -r "$ROOT_DIR/requirements.txt"

PYINSTALLER_CONFIG_DIR="$CACHE_TMP" \
MPLCONFIGDIR="$CACHE_TMP/mpl" \
XDG_CACHE_HOME="$CACHE_TMP/xdg" \
pyinstaller --noconfirm --clean --windowed --onedir \
  --name "$APP_NAME" \
  --distpath "$DIST_TMP" \
  --workpath "$BUILD_TMP" \
  --specpath "$SPEC_TMP" \
  --paths "$ROOT_DIR/src" \
  --icon "$ROOT_DIR/assets/app.icns" \
  --exclude-module matplotlib \
  --exclude-module matplotlib.pyplot \
  --exclude-module gradio \
  --exclude-module gradio_client \
  --exclude-module tkinter \
  "$ROOT_DIR/src/invoicely/invoice_sorter_native.py"

/usr/libexec/PlistBuddy -c "Set :CFBundleIdentifier local.invoicely.native" "$DIST_TMP/$APP_NAME.app/Contents/Info.plist"
codesign --force --deep --sign - "$DIST_TMP/$APP_NAME.app"
codesign --verify --deep --strict "$DIST_TMP/$APP_NAME.app"

mkdir -p "$ROOT_DIR/dist"
rm -rf "$ROOT_DIR/dist/$APP_NAME.app"
ditto --noextattr --norsrc "$DIST_TMP/$APP_NAME.app" "$ROOT_DIR/dist/$APP_NAME.app"

rm -rf "$TMP_BASE"
echo "完成：$ROOT_DIR/dist/$APP_NAME.app"
