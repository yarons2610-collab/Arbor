#!/bin/bash
# Installs the twice-daily export job. Re-run this after changing
# tools/arbor_export.py, since launchd runs an installed copy: macOS (TCC)
# refuses background jobs any access to ~/Desktop, where this repo lives.
set -euo pipefail
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DEST="$HOME/.local/share/arbor"
OUT="$HOME/Arbor exports"
PLIST="$HOME/Library/LaunchAgents/com.arbor.export.plist"

mkdir -p "$DEST" "$OUT"
cp "$SRC/arbor_export.py" "$SRC/arbor_export.sh" "$DEST/"
chmod +x "$DEST/arbor_export.sh"

launchctl bootout "gui/$UID/com.arbor.export" 2>/dev/null || true
launchctl bootstrap "gui/$UID" "$PLIST"
echo "Installed. Exports: $OUT"
