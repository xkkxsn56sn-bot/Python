#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$ROOT_DIR/.venv"
REQ_FILE="$ROOT_DIR/requirements.txt"

log() {
  printf "[setup] %s\n" "$1"
}

warn() {
  printf "[setup][warning] %s\n" "$1"
}

log "Project root: $ROOT_DIR"

if [[ ! -d "$VENV_DIR" ]]; then
  log "Creating virtual environment in .venv"
  python3 -m venv "$VENV_DIR"
else
  log "Virtual environment already exists"
fi

# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

log "Upgrading pip tooling"
python -m pip install --upgrade pip setuptools wheel

if [[ -f "$REQ_FILE" ]]; then
  log "Installing Python dependencies from requirements.txt"
  pip install -r "$REQ_FILE"
else
  warn "requirements.txt not found, skipping pip install"
fi

if ! command -v pandoc >/dev/null 2>&1; then
  warn "pandoc not found (needed by DOCX/Markdown conversion scripts)"
fi

if ! command -v libreoffice >/dev/null 2>&1 && ! command -v soffice >/dev/null 2>&1; then
  warn "libreoffice/soffice not found (needed by WPS conversion scripts)"
fi

if ! command -v xelatex >/dev/null 2>&1; then
  warn "xelatex not found (needed for some PDF conversion flows)"
fi

log "Done. Activate environment with: source .venv/bin/activate"
log "List scripts with: python main.py list"
