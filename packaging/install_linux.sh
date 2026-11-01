#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
APP_SOURCE="$SCRIPT_DIR"
if [[ ! -x "$APP_SOURCE/LAYAAgenticStudio" ]]; then
  printf 'Could not find the packaged app beside this installer.\n' >&2
  exit 1
fi

INSTALL_ROOT="${XDG_DATA_HOME:-$HOME/.local/share}/laya-agentic-studio/app"
mkdir -p "$INSTALL_ROOT"
find "$APP_SOURCE" -mindepth 1 -maxdepth 1 ! -name "install_linux.sh" \
  -exec cp -a {} "$INSTALL_ROOT/" \;
mkdir -p "$HOME/.local/bin" "$HOME/.local/share/applications"
cat > "$HOME/.local/bin/laya-agentic-studio" <<EOF
#!/usr/bin/env sh
exec "$INSTALL_ROOT/LAYAAgenticStudio" "\$@"
EOF
chmod +x "$HOME/.local/bin/laya-agentic-studio"
cat > "$HOME/.local/share/applications/laya-agentic-studio.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=LAYA Agentic Studio
Exec=$HOME/.local/bin/laya-agentic-studio
Terminal=false
Categories=Office;Utility;
EOF
printf 'Installed to %s\nLaunch with: %s\n' "$INSTALL_ROOT" "$HOME/.local/bin/laya-agentic-studio"
