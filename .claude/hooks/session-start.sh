#!/bin/bash
set -euo pipefail

# Only run in remote cloud sessions
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

pip install --quiet pymupdf pypdf pdfplumber reportlab
