#!/bin/bash
# Twice-daily Arbor export, run by launchd (com.arbor.export).
# Writes a styled .xlsx plus the raw JSON it came from, so the folder is both a
# readable report and a backup that depends on nothing else — not a browser,
# not GitHub, not Arbor.
set -uo pipefail

GIST="ce3d7c7936209f0f0322af9531f2a539"
OUT="$HOME/Arbor exports"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="/opt/anaconda3/bin/python3"
GH="$HOME/.local/bin/gh"
LOG="$OUT/export.log"

mkdir -p "$OUT"
printf '\n[%s] starting\n' "$(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"

if "$PY" "$HERE/arbor_export.py" --gist "$GIST" --out "$OUT" --gh "$GH" >> "$LOG" 2>&1; then
  printf '[%s] ok\n' "$(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
else
  printf '[%s] FAILED — previous exports left untouched\n' "$(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
  tail -n 400 "$LOG" > "$LOG.tmp" && mv "$LOG.tmp" "$LOG"
  exit 1
fi
tail -n 400 "$LOG" > "$LOG.tmp" && mv "$LOG.tmp" "$LOG"
