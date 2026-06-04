"""
Build FR-SH xlsx files from JSON payloads matching the schema documented in
_common.SYSTEM_PROMPT.

Modes:
  --json PATH         convert one JSON file
  --all [DIR]         convert every *.json in DIR (default: output) to a
                      same-basename .xlsx alongside it
  (no flag)           read a single JSON from stdin

Examples:
    python build_xlsx_from_json.py --json output\p01.json --xlsx output\p01.xlsx
    python build_xlsx_from_json.py --json output\p01.json    # auto-named
    type output\p01.json | python build_xlsx_from_json.py --xlsx output\p01.xlsx
    python build_xlsx_from_json.py --all                     # batch ./output
    python build_xlsx_from_json.py --all output --pattern 'p*.json'
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, Tuple

from _common import build_fr_sh_xlsx, slugify


def _load_json(path: Path) -> Dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    if not text.strip():
        raise ValueError("empty file")
    return json.loads(text)


def _convert_one(data: Dict[str, Any], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    build_fr_sh_xlsx(data, out_path)


def _run_all(indir: Path, pattern: str) -> Tuple[int, int]:
    """Convert every JSON matching *pattern* in *indir*. Returns (ok, fail)."""
    if not indir.exists():
        print(f"ERROR: {indir} not found.", file=sys.stderr)
        return 0, 1

    json_files = sorted(indir.glob(pattern))
    if not json_files:
        print(f"No JSON files matching {pattern!r} in {indir}.", file=sys.stderr)
        return 0, 0

    print(f"Batch mode: {len(json_files)} JSON file(s) in {indir.resolve()}")
    print("-" * 60)

    ok = 0
    fail = 0
    for jp in json_files:
        xp = jp.with_suffix(".xlsx")
        try:
            data = _load_json(jp)
            _convert_one(data, xp)
            print(f"  {jp.name:40s} -> {xp.name}")
            ok += 1
        except Exception as e:  # noqa: BLE001
            print(f"  {jp.name:40s} FAIL: {e}")
            fail += 1

    print("-" * 60)
    print(f"Done. OK={ok}  FAIL={fail}")
    return ok, fail


def main() -> int:
    ap = argparse.ArgumentParser(description="Build FR-SH xlsx files from JSON.")
    ap.add_argument("--json", dest="json_path", help="single JSON input file")
    ap.add_argument("--xlsx", dest="xlsx_path", help="output xlsx path (only with --json)")
    ap.add_argument(
        "--outdir",
        default="output",
        help="output dir when --xlsx not given (default: output)",
    )
    ap.add_argument(
        "--all",
        dest="all_dir",
        nargs="?",
        const="output",
        default=None,
        metavar="DIR",
        help="batch-convert every *.json in DIR (default: output) to same-basename .xlsx",
    )
    ap.add_argument(
        "--pattern",
        default="*.json",
        help="glob used in --all mode (default: *.json)",
    )
    args = ap.parse_args()

    # --- Batch mode ---
    if args.all_dir is not None:
        if args.json_path or args.xlsx_path:
            print("ERROR: --all is exclusive with --json / --xlsx.", file=sys.stderr)
            return 2
        _, fail = _run_all(Path(args.all_dir), args.pattern)
        return 0 if fail == 0 else 1

    # --- Single-file mode ---
    if args.json_path:
        try:
            data = _load_json(Path(args.json_path))
        except (json.JSONDecodeError, ValueError, FileNotFoundError) as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return 2
    else:
        text = sys.stdin.read()
        if not text.strip():
            print("ERROR: empty JSON input.", file=sys.stderr)
            return 2
        try:
            data = json.loads(text)
        except json.JSONDecodeError as e:
            print(f"ERROR: invalid JSON: {e}", file=sys.stderr)
            return 2

    if args.xlsx_path:
        out_path = Path(args.xlsx_path)
    else:
        out_dir = Path(args.outdir)
        date_slug = data.get("date") or "unknown-date"
        well_slug = slugify(data.get("well_name"), "unknown")
        out_path = out_dir / f"{date_slug}_{well_slug}.xlsx"

    _convert_one(data, out_path)
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
