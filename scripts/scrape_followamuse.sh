#!/usr/bin/env bash
set -euo pipefail

URL="https://followamuse.nl/apps/carnaval-der-dieren/"
DEST="ios/CarnavalDerDierenApp/CarnavalDerDierenApp/Resources/WebMirror"

rm -rf "$DEST"
mkdir -p "$DEST"

wget \
  --mirror \
  --convert-links \
  --adjust-extension \
  --page-requisites \
  --no-parent \
  --directory-prefix "$DEST" \
  "$URL"

# Flatten first downloaded html page to index.html in WebMirror root.
FIRST_HTML=$(find "$DEST" -name '*.html' | head -n 1)
if [[ -n "${FIRST_HTML:-}" ]]; then
  cp "$FIRST_HTML" "$DEST/index.html"
fi

echo "Scraped app into: $DEST"
