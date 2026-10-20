#!/usr/bin/env bash
set -euo pipefail

# Config
SRC_PNG="icon.png"
ICONSET="icon.iconset"
ICNS="icon.icns"
APP_NAME="HOA Payment History Converter"
SPEC="payment_history_converter.spec"

# Check for source PNG
if [[ ! -f "$SRC_PNG" ]]; then
  echo "❌ $SRC_PNG not found. Place your source PNG in this directory."
  exit 1
fi

echo "📦 Creating .icns from $SRC_PNG..."

# Clean up any previous builds
rm -rf "$ICONSET" "$ICNS" build dist

# Create iconset directory and generate all required sizes
mkdir "$ICONSET"

sips -z 16 16     "$SRC_PNG" --out "$ICONSET/icon_16x16.png"
sips -z 32 32     "$SRC_PNG" --out "$ICONSET/icon_16x16@2x.png"
sips -z 32 32     "$SRC_PNG" --out "$ICONSET/icon_32x32.png"
sips -z 64 64     "$SRC_PNG" --out "$ICONSET/icon_32x32@2x.png"
sips -z 128 128   "$SRC_PNG" --out "$ICONSET/icon_128x128.png"
sips -z 256 256   "$SRC_PNG" --out "$ICONSET/icon_128x128@2x.png"
sips -z 256 256   "$SRC_PNG" --out "$ICONSET/icon_256x256.png"
sips -z 512 512   "$SRC_PNG" --out "$ICONSET/icon_256x256@2x.png"
sips -z 512 512   "$SRC_PNG" --out "$ICONSET/icon_512x512.png"
sips -z 1024 1024 "$SRC_PNG" --out "$ICONSET/icon_512x512@2x.png"

iconutil -c icns "$ICONSET" -o "$ICNS"
echo "✅ Created $ICNS"

# Build app with PyInstaller
echo "🚀 Building $APP_NAME.app..."
python -m PyInstaller --noconfirm "$SPEC"

echo "🎉 Build complete!"
echo "Check the dist/ folder for '$APP_NAME.app'."
