#!/usr/bin/env bash
# Renders .xlsx files to PNG through LibreOffice (xlsx -> pdf -> png) so the
# visual result can be inspected. Exits non-zero if LibreOffice fails to
# convert any file (doubles as a LibreOffice compatibility smoke test).
#
# Usage: scripts/render.sh [file.xlsx ...]   (default: out/showcase/*.xlsx)
set -uo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$ROOT/out/render"
PROFILE="$(mktemp -d)"
trap 'rm -rf "$PROFILE"' EXIT

command -v soffice >/dev/null || { echo "soffice (LibreOffice) not found" >&2; exit 2; }
command -v pdftoppm >/dev/null || { echo "pdftoppm (poppler) not found" >&2; exit 2; }

if [[ $# -eq 0 ]]; then
  set -- "$ROOT"/out/showcase/*.xlsx
fi

mkdir -p "$OUT/pdf"
failed=0
for f in "$@"; do
  name="$(basename "${f%.*}")"
  rm -f "$OUT/pdf/$name.pdf" "$OUT/$name"-*.png
  soffice -env:UserInstallation="file://$PROFILE" --headless --convert-to pdf --outdir "$OUT/pdf" "$f" >/dev/null 2>&1
  if [[ ! -s "$OUT/pdf/$name.pdf" ]]; then
    echo "FAIL  $name (LibreOffice could not convert)"
    failed=1
    continue
  fi
  pdftoppm -png -r 70 "$OUT/pdf/$name.pdf" "$OUT/$name"
  pages=$(ls "$OUT/$name"-*.png 2>/dev/null | wc -l)
  echo "ok    $name ($pages page(s))"
done
echo "PNGs in $OUT"
exit $failed
