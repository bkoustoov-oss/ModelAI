#!/bin/bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "$0")" && pwd)"
APP_SOURCE="$SCRIPT_DIR/LAYAAgenticStudio.app"
if [[ ! -d "$APP_SOURCE" ]]; then
  printf 'Could not find LAYAAgenticStudio.app beside this installer.\n' >&2
  exit 1
fi

DESTINATION="$HOME/Applications"
mkdir -p "$DESTINATION"
rm -rf "$DESTINATION/LAYAAgenticStudio.app"
ditto "$APP_SOURCE" "$DESTINATION/LAYAAgenticStudio.app"
open "$DESTINATION/LAYAAgenticStudio.app"
printf 'Installed to %s\n' "$DESTINATION/LAYAAgenticStudio.app"
