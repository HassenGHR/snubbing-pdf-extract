"""
Gemini batch flow: PDF -> per-page FR-SH xlsx.

Renders each PDF page, sends it to Gemini for structured extraction, and
hands the JSON to _common.build_fr_sh_xlsx to produce a parser-compatible
xlsx. Also drops the raw JSON and the rendered PNG next to each xlsx so the
operator can spot-check.

For the alternative *interactive* flow that uses Claude Code itself as the
agent (no external API key), see render_pdf_pages.py + build_xlsx_from_json.py
and the README.
"""

import argparse
import base64
import glob as _glob
import io
import json
import os
import re
import sys
import time as _time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import zipfile

import pypdfium2 as pdfium
from dotenv import load_dotenv
from google import genai
from google.genai import types

from _common import SYSTEM_PROMPT, build_fr_sh_xlsx, safe_dir_name, slugify


# ---------------------------------------------------------------------------
# PDF rendering
# ---------------------------------------------------------------------------

def render_page_png(pdf: pdfium.PdfDocument, page_index: int, dpi: int) -> bytes:
    page = pdf[page_index]
    pil = page.render(scale=dpi / 72).to_pil()
    buf = io.BytesIO()
    pil.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Gemini call
# ---------------------------------------------------------------------------

def extract_with_gemini(
    client: "genai.Client",
    model: str,
    image_bytes: bytes,
    mime_type: str = "image/png",
) -> Dict[str, Any]:
    resp = client.models.generate_content(
        model=model,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            "Extract this report as JSON.",
        ],
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_PROMPT,
            response_mime_type="application/json",
            temperature=0.0,
        ),
    )

    text = (resp.text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)

    usage_dict: Dict[str, int] = {}
    usage = getattr(resp, "usage_metadata", None)
    if usage is not None:
        for attr, key in (
            ("prompt_token_count", "prompt_tokens"),
            ("candidates_token_count", "completion_tokens"),
            ("total_token_count", "total_tokens"),
            ("cached_content_token_count", "cached_tokens"),
        ):
            v = getattr(usage, attr, None)
            if v is not None:
                usage_dict[key] = v

    return {"data": json.loads(text), "usage": usage_dict}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Convert each page of an image-based SDR PDF (or a folder of "
        "already-rendered images) into a per-page FR-SH xlsx.",
    )
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--pdf", help="Input PDF path")
    src.add_argument(
        "--images",
        help="Directory of page images (png/jpg) or a glob (e.g. 'BYL01/*.png'). "
        "Use this when you already have rendered pages instead of a PDF.",
    )
    ap.add_argument(
        "--outdir",
        default="output",
        help="Base output directory (default: output). Files go to <outdir>/<pdf-stem>/.",
    )
    ap.add_argument(
        "--name",
        default=None,
        help="Override subdir name (default: PDF filename without extension)",
    )
    ap.add_argument(
        "--model",
        default="gemini-2.5-flash",
        help="Gemini model id (e.g. gemini-2.5-flash, gemini-2.5-pro)",
    )
    ap.add_argument("--dpi", type=int, default=150, help="PDF render DPI (default: 150)")
    ap.add_argument(
        "--pages",
        default="",
        help="Comma-separated 1-indexed pages to process (default: all)",
    )
    ap.add_argument(
        "--max-retries",
        type=int,
        default=2,
        help="API call retries on failure (default: 2)",
    )
    args = ap.parse_args()

    load_dotenv()
    if not os.environ.get("GEMINI_API_KEY"):
        load_dotenv(".env.example")

    if not os.environ.get("GEMINI_API_KEY"):
        print("ERROR: GEMINI_API_KEY is not set. Put it in .env or export it.", file=sys.stderr)
        return 2

    client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

    ok = 0
    fail = 0
    failures: List[str] = []

    # Build a uniform list of (label, image_bytes, mime_type, stem_base).
    sources: List[Tuple[str, bytes, str, str]] = []
    if args.pdf:
        pdf_path = Path(args.pdf)
        if not pdf_path.exists():
            print(f"ERROR: {pdf_path} not found.", file=sys.stderr)
            return 2
        subdir_name = safe_dir_name(args.name or pdf_path.stem)
        out_dir = Path(args.outdir) / subdir_name
        out_dir.mkdir(parents=True, exist_ok=True)

        pdf = pdfium.PdfDocument(str(pdf_path))
        n_pages = len(pdf)
        if args.pages.strip():
            page_indices: List[int] = []
            for tok in args.pages.split(","):
                tok = tok.strip()
                if not tok:
                    continue
                page_indices.append(int(tok) - 1)
        else:
            page_indices = list(range(n_pages))

        for page_idx in page_indices:
            page_num = page_idx + 1
            label = f"page {page_num:02d}"
            try:
                img_bytes = render_page_png(pdf, page_idx, dpi=args.dpi)
            except Exception as e:  # noqa: BLE001
                print(f"[?/{n_pages}] {label}: render FAIL: {e}")
                fail += 1
                failures.append(f"p{page_num}: render: {e}")
                continue
            sources.append((label, img_bytes, "image/png", f"p{page_num:02d}"))
        source_desc = f"PDF: {pdf_path} ({n_pages} pages)"
    else:  # --images: already-rendered page images
        img_dir = Path(args.images)
        if img_dir.is_dir():
            files = []
            for ext in ("*.png", "*.jpg", "*.jpeg", "*.PNG", "*.JPG", "*.JPEG"):
                files.extend(img_dir.glob(ext))
            files = sorted(set(files))
        else:
            files = sorted(set(Path(p) for p in _glob.glob(args.images)))

        if not files:
            print(f"ERROR: no images found for {args.images!r}.", file=sys.stderr)
            return 2

        subdir_name = safe_dir_name(args.name or img_dir.name)
        out_dir = Path(args.outdir) / subdir_name
        out_dir.mkdir(parents=True, exist_ok=True)

        for f in files:
            mime = "image/png" if f.suffix.lower() == ".png" else "image/jpeg"
            try:
                img_bytes = f.read_bytes()
            except Exception as e:  # noqa: BLE001
                print(f"{f.name}: read FAIL: {e}")
                fail += 1
                failures.append(f"{f.name}: read: {e}")
                continue
            sources.append((f.name, img_bytes, mime, f.stem))
        source_desc = f"Images: {args.images} ({len(files)} files)"

    print(source_desc)
    print(f"Model: {args.model}")
    print(f"Output: {out_dir.resolve()}")
    print(f"Processing {len(sources)} source(s)")
    print("-" * 60)

    totals = {
        "prompt_tokens": 0,
        "completion_tokens": 0,
        "total_tokens": 0,
        "cached_tokens": 0,
    }

    for i, (label, img_bytes, mime, stem_base) in enumerate(sources, 1):
        prefix = f"[{i}/{len(sources)}] {label}"
        last_err: Optional[Exception] = None
        result: Optional[Dict[str, Any]] = None
        for attempt in range(args.max_retries + 1):
            try:
                result = extract_with_gemini(client, args.model, img_bytes, mime)
                break
            except json.JSONDecodeError as e:
                last_err = e
                print(f"{prefix}: JSON parse error on attempt {attempt + 1}: {e}")
            except Exception as e:  # noqa: BLE001
                last_err = e
                print(f"{prefix}: API error on attempt {attempt + 1}: {e}")
            _time.sleep(1.5 * (attempt + 1))

        if result is None:
            print(f"{prefix}: extraction FAIL after retries: {last_err}")
            fail += 1
            failures.append(f"{label}: {last_err}")
            continue

        data = result["data"]
        usage = result.get("usage", {})
        for k, v in usage.items():
            totals[k] = totals.get(k, 0) + (v or 0)

        date_slug = data.get("date") or "unknown-date"
        well_slug = slugify(data.get("well_name"), "unknown")
        stem = f"{stem_base}_{date_slug}_{well_slug}"
        xlsx_path = out_dir / f"{stem}.xlsx"
        json_path = out_dir / f"{stem}.json"
        png_path = out_dir / f"{stem}.png"

        try:
            build_fr_sh_xlsx(data, xlsx_path)
            json_path.write_text(
                json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8"
            )
            png_path.write_bytes(img_bytes)
        except Exception as e:  # noqa: BLE001
            print(f"{prefix}: write FAIL: {e}")
            fail += 1
            failures.append(f"{label}: write: {e}")
            continue

        n_acts = len(data.get("activities") or [])
        print(f"{prefix}: OK -> {xlsx_path.name}  (acts={n_acts})")
        ok += 1

        # Create zip containing this Excel file (include rig/unit name in filename)
        rig_slug = slugify(data.get("unit_name"), "unknown-rig")
        zip_path = out_dir / f"{stem}_{rig_slug}.zip"
        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(xlsx_path, arcname=xlsx_path.name)
            print(f"Created zip archive: {zip_path}")
        except Exception as e:
            print(f"Failed to create zip archive for {xlsx_path.name}: {e}")

    print("-" * 60)
    print(f"Done. OK={ok}  FAIL={fail}")
    if any(totals.values()):
        print("Token usage:", totals)
    if failures:
        print("Failures:")
        for f in failures:
            print(f"  - {f}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
