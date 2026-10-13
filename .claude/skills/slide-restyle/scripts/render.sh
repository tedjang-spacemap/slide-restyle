#!/usr/bin/env bash
# Render a .pptx to PDF + per-slide PNGs with LibreOffice headless, then report font substitution.
#
#   render.sh <deck.pptx> <out_dir> [dpi]
#
# Output: <out_dir>/<stem>.pdf, <out_dir>/slide-N.png (pdftoppm zero-pads by page count),
#         <out_dir>/fonts.txt (pdffonts output) and an embedded-font summary on stdout.
#
# Fonts: headless LibreOffice uses its bundled fontconfig, which on macOS sees none of the system
# fonts. We generate a fonts.conf that adds the macOS font dirs plus project font dirs
# ($SLIDE_RESTYLE_FONT_DIRS, colon-separated; default ./.fonts) and maps fonts that cannot be
# installed on macOS (Malgun Gothic) to the closest local Korean font.
set -euo pipefail

deck="${1:?usage: render.sh <deck.pptx> <out_dir> [dpi]}"
out="${2:?usage: render.sh <deck.pptx> <out_dir> [dpi]}"
dpi="${3:-110}"

soffice="$(command -v soffice || true)"
[[ -z "$soffice" && -x /Applications/LibreOffice.app/Contents/MacOS/soffice ]] && soffice=/Applications/LibreOffice.app/Contents/MacOS/soffice
[[ -z "$soffice" ]] && { echo "soffice not found (install LibreOffice)" >&2; exit 2; }

mkdir -p "$out"
out="$(cd "$out" && pwd)"
stem="$(basename "${deck%.*}")"

# Isolated profile: avoids clashing with a running LibreOffice and the first-run hang.
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

{
  echo '<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "fonts.dtd"><fontconfig>'
  IFS=: read -ra extra <<< "${SLIDE_RESTYLE_FONT_DIRS:-$PWD/.fonts}"
  for d in "${extra[@]}"; do [[ -d "$d" ]] && echo "  <dir>$(cd "$d" && pwd)</dir>"; done
  for d in /System/Library/Fonts /System/Library/AssetsV2 /Library/Fonts "$HOME/Library/Fonts"; do
    [[ -d "$d" ]] && echo "  <dir>$d</dir>"
  done
  echo "  <cachedir>$tmp/fc-cache</cachedir>"
  for f in "Malgun Gothic" "맑은 고딕"; do
    echo "  <alias binding=\"same\"><family>$f</family><prefer><family>Apple SD Gothic Neo</family></prefer></alias>"
  done
  echo '</fontconfig>'
} > "$tmp/fonts.conf"

FONTCONFIG_FILE="$tmp/fonts.conf" "$soffice" -env:UserInstallation="file://$tmp/profile" \
  --headless --norestore --convert-to pdf --outdir "$out" "$deck" >/dev/null 2>&1

pdf="$out/$stem.pdf"
[[ -f "$pdf" ]] || { echo "conversion failed: $deck" >&2; exit 1; }

rm -f "$out"/slide-*.png
pdftoppm -png -r "$dpi" "$pdf" "$out/slide"

pdffonts "$pdf" > "$out/fonts.txt" 2>/dev/null || true
echo "pdf: $pdf"
echo "pngs: $(ls "$out"/slide-*.png | wc -l | tr -d ' ')"
echo "embedded fonts:"
tail -n +3 "$out/fonts.txt" | awk '{print $1}' | sed 's/^[A-Z]\{6\}+//' | sort -u | while read -r f; do
  case "$f" in
    Liberation*|LinuxLibertine*|DejaVu*|OpenSymbol*|FrankRuhl*) echo "  $f   <- LibreOffice fallback (requested font missing)";;
    *) echo "  $f";;
  esac
done
