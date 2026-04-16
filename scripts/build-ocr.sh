#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
SRC="$REPO_ROOT/vendor/ocr/ocr.swift"
OUT="$REPO_ROOT/vendor/ocr/ocr-bin"

if ! command -v swiftc >/dev/null 2>&1; then
    echo "[build-ocr] swiftc not found; skipping OCR build. Image attachments will save to disk but not extract text."
    exit 0
fi

if [ ! -f "$SRC" ]; then
    echo "[build-ocr] source missing at $SRC; skipping."
    exit 0
fi

echo "[build-ocr] compiling $SRC -> $OUT"
swiftc -O "$SRC" -o "$OUT"
echo "[build-ocr] done."
