#!/bin/zsh
set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="DOCX EPUB Converter"
APP_AUTHOR="Marek Zettel"
APP_DIR="$HOME/Desktop/$APP_NAME.app"
PYTHON_BIN="$PROJECT_DIR/.venv/bin/python"

cd "$PROJECT_DIR"

if [ ! -x "$PYTHON_BIN" ]; then
    echo "Tworzenie srodowiska Python..."
    python3 -m venv .venv
fi

echo "Instalowanie zaleznosci..."
"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r requirements.txt

echo "Tworzenie aplikacji na pulpicie..."
rm -rf "$APP_DIR"
mkdir -p "$APP_DIR/Contents/MacOS" "$APP_DIR/Contents/Resources"

cat > "$APP_DIR/Contents/MacOS/launch" <<EOF
#!/bin/zsh
cd "$PROJECT_DIR"
"$PYTHON_BIN" "$PROJECT_DIR/run_converter.py"
EOF
chmod +x "$APP_DIR/Contents/MacOS/launch"

cat > "$APP_DIR/Contents/Info.plist" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>$APP_NAME</string>
    <key>CFBundleDisplayName</key>
    <string>$APP_NAME</string>
    <key>NSHumanReadableCopyright</key>
    <string>Autor: $APP_AUTHOR</string>
    <key>CFBundleExecutable</key>
    <string>launch</string>
    <key>CFBundleIdentifier</key>
    <string>pl.cyfrowyprzyjaciel.docxepubconverter</string>
    <key>CFBundleIconFile</key>
    <string>app_icon</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
</dict>
</plist>
EOF

if [ -f "$PROJECT_DIR/assets/app_icon.icns" ]; then
    cp "$PROJECT_DIR/assets/app_icon.icns" "$APP_DIR/Contents/Resources/app_icon.icns"
fi

echo "Gotowe. Aplikacja utworzona:"
echo "$APP_DIR"
