"""Render one PDF page at high DPI and crop it into legible tiles.

Tiles are kept under ~1500px on the long edge so the Claude Read tool does
not downscale them, preserving handwriting detail for review.

Usage:
    python _crop_page.py --pdf PATH --page N [--dpi 300] [--cols 3] [--rows 2]
                         [--outdir _crops]
Writes <outdir>/pNN_r<row>c<col>.png and prints the paths.
"""

import argparse
import sys
from pathlib import Path

import pypdfium2 as pdfium


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--page", type=int, required=True, help="1-indexed page")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--cols", type=int, default=3)
    ap.add_argument("--rows", type=int, default=2)
    ap.add_argument("--overlap", type=int, default=60, help="px overlap between tiles")
    ap.add_argument("--outdir", default="_crops")
    ap.add_argument("--box", default=None,
                    help="fractional box l,t,r,b (0-1) to crop a single region instead of a grid")
    args = ap.parse_args()

    pdf = pdfium.PdfDocument(str(args.pdf))
    page = pdf[args.page - 1]
    pil = page.render(scale=args.dpi / 72).to_pil()
    W, H = pil.size

    out = Path(args.outdir)
    out.mkdir(parents=True, exist_ok=True)

    if args.box:
        l, t, r, b = (float(x) for x in args.box.split(","))
        crop = pil.crop((int(l * W), int(t * H), int(r * W), int(b * H)))
        p = out / f"p{args.page:02d}_box.png"
        crop.save(p, format="PNG", optimize=True)
        print(p)
        return 0

    tile_w = W // args.cols
    tile_h = H // args.rows
    ov = args.overlap
    for r in range(args.rows):
        for c in range(args.cols):
            left = max(0, c * tile_w - ov)
            upper = max(0, r * tile_h - ov)
            right = min(W, (c + 1) * tile_w + ov)
            lower = min(H, (r + 1) * tile_h + ov)
            crop = pil.crop((left, upper, right, lower))
            p = out / f"p{args.page:02d}_r{r}c{c}.png"
            crop.save(p, format="PNG", optimize=True)
            print(p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
