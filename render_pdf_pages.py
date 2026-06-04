"""
Render each page of a PDF to a PNG file, for use by the Claude Code
interactive flow (Claude reads the PNGs and produces JSON).

The PNGs land in a subdirectory named after the PDF (without extension),
so successive runs with different PDFs don't collide:

    pages/SDR-22.05.2026/p01.png
    pages/SDR-29.05.2026/p01.png

Usage:
    python render_pdf_pages.py --pdf path/to/sdr.pdf [--outdir pages] [--dpi 150]
"""

import argparse
import sys
from pathlib import Path

import pypdfium2 as pdfium

from _common import safe_dir_name


def main() -> int:
    ap = argparse.ArgumentParser(description="Render each page of a PDF to a PNG.")
    ap.add_argument("--pdf", required=True, help="Input PDF path")
    ap.add_argument(
        "--outdir",
        default="pages",
        help="Base output directory (default: pages). Renders go to <outdir>/<pdf-stem>/.",
    )
    ap.add_argument(
        "--name",
        default=None,
        help="Override subdir name (default: PDF filename without extension)",
    )
    ap.add_argument("--dpi", type=int, default=150, help="Rendering DPI (default: 150)")
    args = ap.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"ERROR: {pdf_path} not found.", file=sys.stderr)
        return 2

    subdir_name = safe_dir_name(args.name or pdf_path.stem)
    out_dir = Path(args.outdir) / subdir_name
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(pdf_path))
    n = len(pdf)
    print(f"PDF: {pdf_path} ({n} pages)")
    print(f"Output: {out_dir.resolve()}  DPI: {args.dpi}")

    width = max(2, len(str(n)))
    for i in range(n):
        page = pdf[i]
        pil = page.render(scale=args.dpi / 72).to_pil()
        out_path = out_dir / f"p{i + 1:0{width}d}.png"
        pil.save(out_path, format="PNG", optimize=True)
        print(f"  wrote {out_path}")

    print(f"Done. {n} page(s) written.")
    print()
    print("Next:")
    print(f"  Flow A (Gemini):   python extract.py --pdf \"{pdf_path}\"")
    print(f"  Flow B (Claude):   ask Claude Code to process pages from {out_dir}/")
    print(f"  Then build xlsx:   python build_xlsx_from_json.py --all output/{subdir_name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
