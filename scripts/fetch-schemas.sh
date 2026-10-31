#!/usr/bin/env bash
# Downloads the ECMA-376 (5th ed.) Transitional XML schemas used by the test
# suite to validate every generated part with xmllint. Output: .cache/ecma/xsd
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DEST="$ROOT/.cache/ecma"
URL="https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip"

if [[ -f "$DEST/xsd/sml.xsd" ]]; then
  echo "schemas already present in $DEST/xsd"
  exit 0
fi

mkdir -p "$DEST/xsd"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

echo "downloading ECMA-376 Part 4..."
curl -fsSL -o "$TMP/part4.zip" "$URL"
unzip -oq "$TMP/part4.zip" OfficeOpenXML-XMLSchema-Transitional.zip -d "$TMP"
unzip -oq "$TMP/OfficeOpenXML-XMLSchema-Transitional.zip" -d "$DEST/xsd"
echo "schemas installed in $DEST/xsd"
