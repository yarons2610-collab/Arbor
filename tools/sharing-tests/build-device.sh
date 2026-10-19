#!/bin/bash
# usage: mk.sh <dir> <git-rev|WORKTREE>
HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
OUT="${OUT:-$HERE/.devices}"
D="$OUT/$1"; mkdir -p "$D"
if [ "$2" = "WORKTREE" ]; then cp "$REPO/index.html" "$D/index.html"; else (cd "$REPO" && git show "$2:index.html") > "$D/index.html"; fi
for x in vendor icons; do [ -e "$D/$x" ] || ln -s "$REPO/$x" "$D/$x"; done
cp "$REPO/manifest.json" "$D/" 2>/dev/null
python3 - "$D/index.html" "$HERE/shim.js" <<'PYEOF'
import sys
p,shim=sys.argv[1],open(sys.argv[2]).read()
s=open(p).read(); anchor='<meta charset="utf-8">'
i=s.index(anchor)+len(anchor)
open(p,"w").write(s[:i]+"\n<script>"+shim+"</script>\n"+s[i:])
PYEOF
cp "$HERE/suite.js" "$OUT/lib.js" 2>/dev/null
echo "$1 <- $2 : shim=$(grep -c 'window.__dev=dev' "$D/index.html")"
