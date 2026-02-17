#!/usr/bin/env bash
# Warhol Setup Script
# Run this to set up the project environment from scratch.
# Usage: bash scripts/setup.sh

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
cd "$PROJECT_ROOT"

echo "=============================="
echo "  Warhol Setup"
echo "=============================="
echo ""

# 1. Check Python version
echo "→ Checking Python..."
if ! command -v python3 &> /dev/null; then
    echo "✗ Python 3 is not installed. Install Python 3.11+ and try again."
    exit 1
fi

PY_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
PY_MAJOR=$(python3 -c 'import sys; print(sys.version_info.major)')
PY_MINOR=$(python3 -c 'import sys; print(sys.version_info.minor)')

if [ "$PY_MAJOR" -lt 3 ] || { [ "$PY_MAJOR" -eq 3 ] && [ "$PY_MINOR" -lt 11 ]; }; then
    echo "✗ Python $PY_VERSION found, but 3.11+ is required."
    exit 1
fi
echo "  ✓ Python $PY_VERSION"

# 2. Create virtual environment
echo "→ Creating virtual environment..."
if [ -d ".venv" ]; then
    echo "  .venv already exists, reusing it"
else
    python3 -m venv .venv
    echo "  ✓ Created .venv"
fi

# 3. Activate and install
echo "→ Installing dependencies..."
source .venv/bin/activate
pip install --upgrade pip --quiet
pip install -e . --quiet
echo "  ✓ Dependencies installed"

# 4. Ensure directory structure
echo "→ Ensuring directory structure..."
for dir in workspace output inputs templates; do
    mkdir -p "$dir"
    if [ ! -f "$dir/.gitkeep" ]; then
        touch "$dir/.gitkeep"
    fi
done
echo "  ✓ Directories ready"

# 5. Verify installation
echo "→ Verifying installation..."
python3 -c "
import pptx, pydantic, bs4, yaml
print('  ✓ python-pptx', pptx.__version__)
print('  ✓ pydantic', pydantic.__version__)
print('  ✓ beautifulsoup4')
print('  ✓ pyyaml')
" 2>&1 || {
    echo "✗ Package verification failed. Try: pip install -e ."
    exit 1
}

echo ""
echo "=============================="
echo "  ✓ Warhol is ready!"
echo "=============================="
echo ""
echo "Next steps:"
echo "  1. Drop a .pptx template into templates/ (optional)"
echo "  2. Drop your input document into inputs/"
echo "  3. Ask Cursor: \"Generate a presentation from inputs/your-file.pdf\""
echo ""
