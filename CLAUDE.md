# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Converts scanned (image-based) Sonatrach SDR PDFs into per-page FR-SH Excel files that the **workover-api** backend ingests via `/v1/snubbing_excel_import/upload-auto`. Each output `.xlsx` has values placed in the exact cells the backend's FR-SH parser reads.

## Commands

Activate the virtualenv first for every command:

```powershell
.\venv\Scripts\Activate.ps1
```

**Flow A — Gemini batch (full PDF → xlsx in one shot):**
```powershell
python extract.py --pdf SDR-22.05.2026.pdf
python extract.py --pdf SDR-22.05.2026.pdf --pages 1,3,5   # spot-check first
python extract.py --pdf SDR-22.05.2026.pdf --model gemini-2.5-pro --dpi 200
```

**Flow B — Claude Code interactive:**
```powershell
# Step 1: render PDF pages to PNGs
python render_pdf_pages.py --pdf SDR-22.05.2026.pdf

# Step 3: convert reviewed JSONs to xlsx (after Claude writes the JSON files)
python build_xlsx_from_json.py --all output\SDR-22.05.2026
python build_xlsx_from_json.py --json output\SDR-22.05.2026\p01.json  # single page
```

**Regenerate xlsx from an edited JSON:**
```powershell
python build_xlsx_from_json.py --json output\SDR-22.05.2026\p01_2026-05-22_OMJZ742.json --xlsx output\SDR-22.05.2026\p01_2026-05-22_OMJZ742.xlsx
```

## Architecture

```
_common.py               ← canonical JSON schema (SYSTEM_PROMPT) + build_fr_sh_xlsx()
extract.py               ← Flow A: PDF → Gemini → JSON → xlsx
render_pdf_pages.py      ← Flow B step 1: PDF → per-page PNGs in pages/<pdf-stem>/
build_xlsx_from_json.py  ← Flow B step 3 + any one-off JSON → xlsx conversion
```

**Data flow:** PDF → `render_page_png()` → Gemini (with `SYSTEM_PROMPT`) → JSON → `build_fr_sh_xlsx()` → `.xlsx`

Output lands in `output/<pdf-stem>/` as a triplet: `pNN_<date>_<well>.xlsx`, `.json`, `.png`.  
Rendered PNGs land in `pages/<pdf-stem>/`.

## Critical field mappings (counterintuitive)

`current_status` ← **"Situation au rapport"** cell (terse state at report time)  
`day_summary` ← **"Résumé:"** block under the operations table (fuller day summary)

These are intentionally swapped from what the field names suggest. The SYSTEM_PROMPT documents this explicitly.

**Supervisors:** extract only the SH side — `superintendent` (Superintendant/SH Superintendent column) and `supervisor` (Superviseur SH column). Always skip: Chef de poste/C.P, Chef de chantier/C.C, ENSP C/P, ENSP C.Ch, stagiaire, and any contractor column (B.YL Supervisor, HALL Supervisor, Senior Oper).

**remark field:** The `REMARQUES`/`Notes` panel content **plus** any below-operations lines prefixed `N.B:` / `NB:` / `Note:` / `Notes:` (strip prefix and leading `*`, join with ` / `). Do NOT include bare logistics lines (Ambulance, FOURGON, GRUE, etc.) that lack these prefixes.

## Excel cell map (key cells the backend reads)

| Cell | Field |
|---|---|
| O3 | date (as `datetime` object) |
| B5 | well\_name |
| D5 | field\_name |
| F5 | unit\_name |
| A10 | operation\_type |
| A16 | BOP test date (embedded in label string) |
| A18–A33 / B18–B33 | activity start / end times |
| C18–C33 | activity descriptions |
| J18–J33 | activity durations |
| L32 | equipment |
| C35–C40 | day\_summary (first line prefixed "Résumé:") |
| K41 | remark |
| B45 | current\_status |
| B46 | next\_day\_summary |
| B47 | program |
| N48 | superintendent |
| P48 | supervisor |

Sheet name is `"RAP SH  "` (two trailing spaces — deliberate).  
Activities are capped at rows 18–33 (max 16 rows). Equipment is truncated to 200 chars on DB import.

## Schema source of truth

`_common.py::SYSTEM_PROMPT` is the canonical schema definition used by both Gemini (Flow A) and Claude Code (Flow B). `build_fr_sh_xlsx()` in the same file is the single Excel-writing function. If the workover-api parser changes which cells it reads, update `build_fr_sh_xlsx()` there.

## Environment

Requires `GEMINI_API_KEY` in `.env` (Flow A only). Copy `.env.example` → `.env` and fill in the key.
