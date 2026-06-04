# SDR PDF → FR-SH Excel Converter

Convert scanned (image-based) snubbing daily-report PDFs into per-page Excel
files that the **workover-api** backend can ingest directly.

Each output `.xlsx` is placed at the exact cell positions the backend's FR-SH
parser reads (B5 = well, O3 = date, A10 = operation type, K32/L32 =
équipement, K38/K41 = remarques, L48/P48 = supervisors, etc.), so once an
xlsx is uploaded, no re-mapping is needed.

---

## Why this exists

The Sonatrach SDR reports arrive as scanned PDFs (just images — no extractable
text). The workover-api `/v1/snubbing_excel_import/upload-auto` endpoint
already handles Excel files but cannot read image PDFs directly. This tool
bridges the gap:

```
   sdr.pdf (20 image pages)
        │
        ▼   (this project)
  ┌─────────────────────────┐
  │   PDF → image OCR/VLM   │
  │   structured JSON       │
  │   FR-SH Excel writer    │
  └─────────────────────────┘
        │
        ▼
   20 per-page .xlsx files  ──upload──▶  workover-api  ──▶  database
```

---

## Two ways to extract: pick one (or use both)

| | **Flow A — Gemini batch** | **Flow B — Claude Code interactive** |
|---|---|---|
| Who does the OCR/extraction | Google Gemini 2.5 Flash (vision) | Claude Code itself (the CLI you're using) |
| External API key needed | Yes — `GEMINI_API_KEY` (free tier) | None |
| Cost | Free tier covers 1500 req/day | Counts against your Claude Code plan |
| Speed (20 pages) | ~3 minutes | ~30+ minutes (interactive review per page) |
| Review style | Run all → review → upload | Review each page as it's extracted |
| Best for | Bulk processing, easy pages | Hard pages, manual oversight |

You don't have to choose. Use Flow A for the batch, fall back to Flow B for
any pages Flow A got wrong. Both produce identical xlsx output.

---

## Project structure

```
sdr_pdf_to_excel/
├── _common.py                ← shared JSON schema + Excel builder
├── extract.py                ← Flow A: PDF → per-page xlsx (Gemini)
├── render_pdf_pages.py       ← Flow B step 1: PDF → per-page PNGs
├── build_xlsx_from_json.py   ← Flow B step 3: JSON → xlsx (one or many)
├── requirements.txt
├── .env.example              ← put your GEMINI_API_KEY here (or copy to .env)
├── README.md                 ← (this file)
├── venv/                     ← isolated Python virtualenv
├── pages/                    ← rendered page images, one subdir per PDF
│   └── SDR-22.05.2026/
│       ├── p01.png
│       └── ...
└── output/                   ← extracted xlsx/json, one subdir per PDF
    └── SDR-22.05.2026/
        ├── p01_2026-05-22_OMJZ742.xlsx
        ├── p01_2026-05-22_OMJZ742.json
        ├── p01_2026-05-22_OMJZ742.png
        └── ...
```

---

## Prerequisites

- **Python 3.10+** (other versions probably work but unverified).
- **Windows / macOS / Linux** (developed on Windows 11).
- For Flow A: a **Gemini API key** (free) — see [Getting a Gemini key](#getting-a-gemini-api-key) below.
- For Flow B: Claude Code installed and authenticated.

This project lives **outside** the workover-api repo. It has its own
virtualenv and dependencies; nothing here touches the workover-api venv.

---

## One-time setup

```powershell
cd H:\new\sdr_pdf_to_excel
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
# edit .env, paste GEMINI_API_KEY=... (only needed for Flow A)
```

(macOS / Linux: `source venv/bin/activate` instead of the `Activate.ps1`.)

### Getting a Gemini API key

1. Visit https://aistudio.google.com/apikey
2. Sign in with any Google account.
3. Click **Create API key** → copy the value (starts with `AIzaSy…`).
4. Paste into `.env`:
   ```
   GEMINI_API_KEY=AIzaSy...
   ```

Free tier: 1500 requests/day on `gemini-2.5-flash`. Easily covers hundreds of
pages. No credit card required.

---

## Flow A — Gemini batch

For the typical "I have 20 reports, get me all the Excel files" workflow.

```powershell
.\venv\Scripts\Activate.ps1
python extract.py --pdf ..\workover-api\SDR-22.05.2026.pdf
```

What happens:
1. PDF is split into pages and each page is rendered to a PNG in-memory.
2. Each PNG is sent to Gemini 2.5 Flash with a strict JSON-schema prompt.
3. Gemini returns structured JSON (date, well, activities, supervisors, …).
4. The JSON is written to an `.xlsx` matching the FR-SH template layout.
5. The raw `.json` and rendered `.png` are saved next to the `.xlsx` so you
   can compare them during review.

Output:
```
output/SDR-22.05.2026/p01_2026-05-22_OMJZ742.xlsx
output/SDR-22.05.2026/p01_2026-05-22_OMJZ742.json
output/SDR-22.05.2026/p01_2026-05-22_OMJZ742.png
output/SDR-22.05.2026/p02_*.xlsx
...
```

The subdir `SDR-22.05.2026` is derived from the PDF filename — successive
runs with different PDFs don't collide.

### Flow A options

| Flag | Default | Purpose |
|---|---|---|
| `--pdf PATH` | required | Input PDF |
| `--outdir DIR` | `output` | Base output dir; files go to `<outdir>/<pdf-stem>/` |
| `--name FOO` | (PDF stem) | Override the per-PDF subdir name |
| `--pages 1,5,10` | (all) | Only process these 1-indexed pages |
| `--model gemini-2.5-pro` | `gemini-2.5-flash` | Switch to a more capable (slower) model |
| `--dpi 200` | 150 | Bump rendering resolution for hard scans |
| `--max-retries 3` | 2 | Retries on API/JSON-parse errors |

---

## Flow B — Claude Code interactive

For when you want to oversee every page or Flow A struggles on certain pages.

### Step 1 — render the PDF to PNGs

```powershell
.\venv\Scripts\Activate.ps1
python render_pdf_pages.py --pdf ..\workover-api\SDR-22.05.2026.pdf
```

Produces `pages/SDR-22.05.2026/p01.png … p20.png`. The script prints the
exact commands to run for steps 2 and 3.

### Step 2 — let Claude Code read each page

Start Claude Code in this folder:

```powershell
claude
```

Then paste a prompt like:

> Read `_common.py` to learn the JSON schema, then for each
> `pages/SDR-22.05.2026/pNN.png`, read the image, produce the JSON for that
> page, save it to `output/SDR-22.05.2026/pNN.json`, and show me each
> extraction before writing so I can correct it.

Substitute `SDR-22.05.2026` with whatever subdir step 1 printed. The agent
will:
1. `Read` `p01.png` — vision-enabled, sees the form.
2. Show you the JSON it inferred.
3. You approve or correct.
4. Saves `p01.json`. Moves to `p02.png`. Repeat.

### Step 3 — convert JSON files to xlsx

Once all the JSONs are reviewed:

```powershell
python build_xlsx_from_json.py --all output\SDR-22.05.2026
```

`--all` walks the directory, converts every `*.json` to a same-basename
`.xlsx` next to it.

Other usage modes:

```powershell
# Single file, explicit output path
python build_xlsx_from_json.py --json output\SDR-22.05.2026\p01.json --xlsx output\SDR-22.05.2026\p01.xlsx

# Single file, auto-named <date>_<well>.xlsx
python build_xlsx_from_json.py --json output\SDR-22.05.2026\p01.json

# Stdin
type output\SDR-22.05.2026\p01.json | python build_xlsx_from_json.py --xlsx output\SDR-22.05.2026\p01.xlsx

# Custom glob in batch mode
python build_xlsx_from_json.py --all output\SDR-22.05.2026 --pattern 'p1*.json'
```

---

## Reviewing extractions

For each report you'll have three files side by side:

| File | Purpose |
|---|---|
| `pNN…png` | The rendered page — your "source of truth" for the review |
| `pNN…json` | Raw structured data extracted. Edit here if you want the xlsx regenerated. |
| `pNN…xlsx` | The Excel that will be uploaded |

Two ways to fix mistakes:

**A. Edit the JSON, regenerate the xlsx.**
```powershell
notepad output\SDR-22.05.2026\p01_2026-05-22_OMJZ742.json
python build_xlsx_from_json.py --json output\SDR-22.05.2026\p01_2026-05-22_OMJZ742.json --xlsx output\SDR-22.05.2026\p01_2026-05-22_OMJZ742.xlsx
```

**B. Edit the xlsx directly in Excel.**
The backend parser reads cell values, not formatting. Any change you make in
Excel that ends up in the right cell will be picked up on import.

---

## Uploading to the backend

Once an `.xlsx` is reviewed and correct:

1. Start workover-api locally:
   ```powershell
   cd H:\new\workover-api
   .\venv\Scripts\Activate.ps1
   uvicorn app.main:app --port 8000 --host 0.0.0.0 --reload
   ```
2. POST the xlsx to `http://localhost:8000/v1/snubbing_excel_import/upload-auto`
   (or use the existing front-end uploader).
3. The backend's auto-resolver picks the matching `wellbore_operations` row
   from rig + well + date.

For 20 reports you can script the upload (PowerShell / curl loop) but
typically you upload one at a time so you can react to any per-file errors.

---

## JSON schema reference

The full schema and field rules live in `_common.py` as `SYSTEM_PROMPT`. The
shape is:

```json
{
  "date": "YYYY-MM-DD",
  "well_name": "MD 336",
  "field_name": "HMD",
  "unit_name": "HRS 225",
  "operation_type": "Démontage appareil.",
  "supervisors": ["SAYOUDI", "S.BAADOUD"],
  "vehicle_mat": "",
  "activities": [
    {"start_time": "06:00", "end_time": "06:30", "description": "Safety meeting…", "duration": 0.5}
  ],
  "current_status": "Démontage appareil",
  "day_summary": "Démontage appareil SNB @ 100% …",
  "next_day_summary": "Opération CTU",
  "program": "Changement CCE 1\"660 + N.F au CTU.",
  "remark": "Prévoir OCT 4\"1/16 + vanne 3\"1/8.",
  "bop_test_date": "2026-04-13",
  "equipment": "Bride taraudée 2\"1/16 + Bride taraudée 3\"1/8 + …"
}
```

| Field | Type | Source on the form | Notes |
|---|---|---|---|
| `date` | `YYYY-MM-DD` or `null` | O3 (top-right) | Form shows `DD/MM/YYYY`; output ISO. |
| `well_name` | string | PUITS | e.g. `"MD 336"` |
| `field_name` | string | CHAMPS | usually `"HMD"` |
| `unit_name` | string | APPAREIL | e.g. `"HRS 225"` |
| `operation_type` | string | TYPE OPÉRATION | free-form French |
| `supervisors` | array | Signature block | French: `[Superviseur SH/DP or Superintendant, SUPERVISEUR]`. English/BYL/HALL: `[SUPERINTENDANT, SH SUPERVISOR]` (drop contractor + stagiaire). Always skip Chef de poste/chantier (C.P/C.C). |
| `vehicle_mat` | string | Véhicule / Mat | Plate. Empty if absent. |
| `activities` | array | DE / A / Opérations / Heures table | `duration` is decimal hours. |
| `current_status` | string | **Situation au rapport** | Mapped opposite to the field name — terse current state. |
| `day_summary` | string | **Résumé** block (under operations) | Mapped opposite to the field name — the fuller day summary. |
| `next_day_summary` | string | Demain / Next day | `""` if no such row. |
| `program` | string | Programme prévu / Program | |
| `remark` | string | REMARQUES / Notes panel | **Plus** below-ops `N.B:` / `NB:` / `Note:` / `Notes:` lines (prefix + leading `*` stripped, joined with ` / `). |
| `bop_test_date` | `YYYY-MM-DD` or `null` | "Dernier test des BOP's" | |
| `equipment` | string | Equipements SHDP | Often long; truncated to 200 chars on DB import. |

---

## Command-line reference

### `extract.py` (Flow A)
```
python extract.py --pdf PATH [--outdir DIR] [--name FOO]
                  [--pages 1,3,5] [--model MODEL] [--dpi N]
                  [--max-retries N]
```

### `render_pdf_pages.py` (Flow B step 1)
```
python render_pdf_pages.py --pdf PATH [--outdir DIR] [--name FOO] [--dpi N]
```

### `build_xlsx_from_json.py` (Flow B step 3 and any one-off conversion)
```
# Single JSON file
python build_xlsx_from_json.py --json PATH [--xlsx PATH]
python build_xlsx_from_json.py --json PATH [--outdir DIR]    # auto-named

# stdin
<json> | python build_xlsx_from_json.py [--xlsx PATH | --outdir DIR]

# Batch a directory
python build_xlsx_from_json.py --all [DIR] [--pattern '*.json']
```

---

## Troubleshooting

### `ERROR: GEMINI_API_KEY is not set`
- You haven't created `.env` from `.env.example`, OR
- You did but pasted the key with surrounding quotes or whitespace.
- Verify: open `.env`, ensure it reads exactly `GEMINI_API_KEY=AIzaSy...` on
  one line with no quotes.

### `429 / RESOURCE_EXHAUSTED` from Gemini
- You hit the free-tier rate limit (typically 15 req/min for Flash).
- The script auto-retries with backoff. If it keeps failing, wait a minute or
  switch to `--model gemini-2.5-pro` (different quota bucket) or process
  fewer pages at a time with `--pages`.

### The generated xlsx has empty `day_summary` even though the JSON had a value
- Known parser quirk in workover-api: the FR-SH parser's `_is_fr_label`
  helper treats any cell containing `"mat"` as a "label" cell (false-matches
  on "matériels", "matériaux", etc.). Workaround: edit the cell in Excel to
  rephrase, or accept the loss for that field.

### Activities with empty `start_time` / `end_time` get dropped or merged
- Intentional in the backend parser — rows with no times are treated as
  continuation lines of the previous activity. Per the schema, `N.B:` / `Note:`
  lines below the operations table are extracted into `remark` (prefix stripped),
  not kept as activity rows.

### Long `equipment` strings get cut off
- `General.equipment` is `CharField(max_length=200)` in the DB. Anything past
  200 characters is truncated on import. Currently a known limitation; not
  fixable without a DB migration.

### `vehicle_mat` from Gemini is "1049" but the xlsx has no vehicle cell
- The backend parser only accepts Algerian-style plate format (three digit
  groups). If Gemini reads a fleet number or partial value, the script
  intentionally skips writing it so the parser doesn't reject the file
  outright. Fill the cell by hand in Excel if you need it.

### Re-rendering a PDF doesn't pick up changes
- `pages/<pdf-stem>/` is the cache. If you replaced the PDF with a new
  version under the same filename, delete that subdir and re-render:
  ```powershell
  Remove-Item pages\SDR-22.05.2026 -Recurse -Force
  python render_pdf_pages.py --pdf ..\workover-api\SDR-22.05.2026.pdf
  ```

---

## Known limitations

- Hand-written French on poorly-scanned pages will produce errors no matter
  which flow you pick. Plan to review every output.
- The "Dernier test des BOP's" date is parsed only when written inline in the
  same cell as the label, in `DD-MM-YYYY` (or `/`-separated) format.
- The backend parser's "Demain → Program fallback" means an empty `Demain`
  cell causes `next_day_summary` to be filled with the `program` text.
- `_common.py` is the canonical source for the JSON schema. If the workover-api
  parser changes which cells it reads, update `build_fr_sh_xlsx()` there.

---

## Tips

- Run `python extract.py --pages 1` (or `--pages 1,5,10`) on a new PDF first
  to verify quality before kicking off the full batch.
- Keep the rendered `.png` next to each `.xlsx` during review — comparing the
  two side by side is much faster than re-opening the PDF.
- For repeated runs on the same set of PDFs, the per-PDF subdirs serve as
  history; you can re-upload any older xlsx by pointing the backend at it.
