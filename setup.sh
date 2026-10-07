#!/usr/bin/env bash
# One-time setup for the Shift Projects slide builder.
# Safe to re-run. macOS / Linux (on Windows use Git Bash or WSL).
set -u
GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'
ok(){ echo -e "${GREEN}✓${NC} $*"; }; warn(){ echo -e "${YELLOW}!${NC} $*"; }; bad(){ echo -e "${RED}✗${NC} $*"; }

echo "Shift Projects slide builder: setup check"
echo

# 1. Where am I? The skill must live at ~/.claude/skills/frontend-slides
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ "$HERE" != "$HOME/.claude/skills/frontend-slides" ]; then
  warn "This folder is $HERE. Claude Code only finds the skill at ~/.claude/skills/frontend-slides."
fi

# 2. Node.js (needed for PDF export)
for d in "$HOME/.nvm/versions/node/$(ls "$HOME/.nvm/versions/node" 2>/dev/null | sort -V | tail -1)/bin" /opt/homebrew/bin /usr/local/bin; do
  [ -d "$d" ] && case ":$PATH:" in *":$d:"*) ;; *) PATH="$d:$PATH";; esac
done
if command -v node >/dev/null 2>&1; then ok "Node.js $(node --version)"; HAVE_NODE=1
else bad "Node.js not found (needed for PDF export). macOS: brew install node, or https://nodejs.org"; HAVE_NODE=0; fi

# 3. Python + python-pptx (needed for the Google Slides template and PowerPoint conversion)
PY="$(command -v python3 || true)"
if [ -n "$PY" ]; then
  ok "Python $($PY --version 2>&1 | cut -d' ' -f2)"
  if $PY -c "import pptx" 2>/dev/null; then ok "python-pptx installed"
  else
    echo "  installing python-pptx..."
    if $PY -m pip install --user --quiet python-pptx 2>/dev/null || $PY -m pip install --user --quiet --break-system-packages python-pptx 2>/dev/null; then ok "python-pptx installed"
    else bad "Could not install python-pptx. Run: $PY -m pip install --user python-pptx"; fi
  fi
else bad "Python 3 not found (needed for the Google Slides template). macOS: brew install python"; fi

# 4. Pre-download the PDF-export browser so the first export is quick (about 150 MB, one time)
if [ "$HAVE_NODE" = 1 ]; then
  echo "  downloading Chromium for PDF export (one time)..."
  TMP="$(mktemp -d)"; ( cd "$TMP" && npm init -y >/dev/null 2>&1 && npm install playwright >/dev/null 2>&1 && npx playwright install chromium >/dev/null 2>&1 ) \
    && ok "PDF export browser ready" || warn "Chromium download failed; it will retry on first PDF export"
  rm -rf "$TMP"
fi

echo
ok "Done. Open Claude Code and type:  /frontend-slides"
echo "   (say 'Shift Projects deck about ...' and it uses the Shift brand automatically)"
