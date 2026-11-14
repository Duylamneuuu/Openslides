#!/usr/bin/env bash
# Render a .pptx to one PNG per slide plus a contact sheet.
# Needs LibreOffice (soffice) and pdftoppm (poppler).
# Usage: scripts/render_preview.sh deck.pptx [output_dir]

set -euo pipefail

if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo "usage: scripts/render_preview.sh deck.pptx [output_dir]" >&2
  exit 2
fi

INPUT="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
if [[ ! -f "$INPUT" ]]; then
  echo "error: file not found: $1" >&2
  exit 2
fi

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
if [[ $# -eq 2 ]]; then
  OUT="$2"
  mkdir -p "$OUT"
  OUT="$(cd "$OUT" && pwd)"
else
  OUT="$(dirname "$INPUT")/$(basename "$INPUT" .pptx)-preview"
  mkdir -p "$OUT"
fi

if [[ -d "$ROOT/fonts" ]]; then
  FONT_DIR="${HOME}/.local/share/fonts/openslides"
  mkdir -p "$FONT_DIR"
  cp -f "$ROOT"/fonts/*.ttf "$FONT_DIR"/ 2>/dev/null || true
  if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f "$FONT_DIR" >/dev/null 2>&1 || true
  fi
fi

SOFFICE=""
for candidate in soffice libreoffice; do
  if command -v "$candidate" >/dev/null 2>&1; then
    SOFFICE="$candidate"
    break
  fi
done
if [[ -z "$SOFFICE" ]]; then
  echo "error: LibreOffice is not installed (need soffice or libreoffice)." >&2
  echo "Install it with: sudo apt-get install -y libreoffice-impress poppler-utils" >&2
  exit 1
fi
if ! command -v pdftoppm >/dev/null 2>&1; then
  echo "error: pdftoppm is not installed (poppler-utils)." >&2
  exit 1
fi

PROFILE="$(mktemp -d)"
cleanup() { rm -rf "$PROFILE"; }
trap cleanup EXIT

# 96 dpi on a 13.333in slide is about 1280px wide.
PDF="$OUT/deck.pdf"
rm -f "$PDF" "$OUT"/slide-*.png "$OUT"/contact-sheet.png

echo "Converting to PDF…"
if ! timeout 180 "$SOFFICE" --headless --norestore --nolockcheck \
  -env:UserInstallation="file://${PROFILE}" \
  --convert-to pdf --outdir "$OUT" "$INPUT" >/tmp/openslides-soffice.log 2>&1; then
  echo "error: LibreOffice failed to convert. Log:" >&2
  cat /tmp/openslides-soffice.log >&2 || true
  exit 1
fi

# LibreOffice names the PDF after the input file.
PRODUCED="$OUT/$(basename "$INPUT" .pptx).pdf"
if [[ ! -f "$PRODUCED" ]]; then
  echo "error: LibreOffice did not write a PDF. Log:" >&2
  cat /tmp/openslides-soffice.log >&2 || true
  exit 1
fi
if [[ "$PRODUCED" != "$PDF" ]]; then
  mv -f "$PRODUCED" "$PDF"
fi

pdftoppm -png -r 96 "$PDF" "$OUT/slide"
python3 - "$OUT" << 'PY'
import sys
from pathlib import Path

out = Path(sys.argv[1])
frames = sorted(out.glob("slide-*.png"), key=lambda path: path.name)
# pdftoppm writes slide-1.png. Rename to slide-01.png so they sort.
renamed = []
for path in frames:
    stem = path.stem
    if stem.startswith("slide-") and stem.split("-", 1)[1].isdigit():
        number = int(stem.split("-", 1)[1])
        target = out / f"slide-{number:02d}.png"
        if path != target:
            path.rename(target)
        renamed.append(target)
    else:
        renamed.append(path)
renamed = sorted(out.glob("slide-*.png"))
if not renamed:
    raise SystemExit("error: pdftoppm produced no PNGs")

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    print(f"Wrote {len(renamed)} PNGs in {out}")
    print("warning: Pillow is not installed, so there is no contact sheet", file=sys.stderr)
    raise SystemExit(0)

font = None
for face in (
    Path.home() / ".local/share/fonts/openslides/BeVietnamPro-SemiBold.ttf",
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
):
    if face.is_file():
        font = ImageFont.truetype(str(face), 22)
        break
if font is None:
    font = ImageFont.load_default()

thumbs = []
for path in renamed:
    image = Image.open(path).convert("RGB")
    image.thumbnail((480, 270))
    thumbs.append(image)

cols = 4 if len(thumbs) > 4 else max(1, len(thumbs))
rows = (len(thumbs) + cols - 1) // cols
pad = 16
label_h = 28
cell_w = max(thumb.width for thumb in thumbs)
cell_h = max(thumb.height for thumb in thumbs)
sheet = Image.new("RGB", (cols * cell_w + (cols + 1) * pad, rows * (cell_h + label_h) + (rows + 1) * pad), "#F4F0E8")
draw = ImageDraw.Draw(sheet)
for index, thumb in enumerate(thumbs):
    row, col = divmod(index, cols)
    x = pad + col * (cell_w + pad) + (cell_w - thumb.width) // 2
    y = pad + row * (cell_h + label_h + pad)
    sheet.paste(thumb, (x, y))
    draw.text((pad + col * (cell_w + pad), y + cell_h + 2), f"{index + 1:02d}", fill="#1C1915", font=font)
sheet.save(out / "contact-sheet.png", optimize=True)
print(f"Wrote {len(renamed)} slides and a contact sheet in {out}")
PY
rm -f "$PDF"
