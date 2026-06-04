"""
Shared bits used by both flows:
  * extract.py            (Gemini batch flow)
  * build_xlsx_from_json.py (Claude Code interactive flow)

If you change the JSON schema, update both SYSTEM_PROMPT and build_fr_sh_xlsx.
"""

import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Dict, Optional

import openpyxl


# ---------------------------------------------------------------------------
# JSON schema documented for both the LLM (Gemini) and a human agent (Claude
# Code / a person) building the JSON by hand.
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """You are an expert at reading scanned French snubbing daily reports
(Rapport Journalier Snubbing) from SONATRACH Division Production HMD.

Your task: extract structured data from a single page image of a daily report.

Respond with ONLY a JSON object (no markdown fences, no commentary) using this exact schema:

{
  "date": "YYYY-MM-DD" or null,
  "well_name": "string",
  "field_name": "string",
  "unit_name": "string",
  "operation_type": "string",
  "superintendent": "string",
  "supervisor": "string",
  "vehicle_mat": "string",
  "activities": [
    {"start_time": "HH:MM", "end_time": "HH:MM", "description": "string", "duration": <number>}
  ],
  "current_status": "string",
  "day_summary": "string",
  "next_day_summary": "string",
  "program": "string",
  "remark": "string",
  "bop_test_date": "YYYY-MM-DD" or null,
  "equipment": "string"
}

Field guide:
- date: top-right "DATE" field, form shows DD/MM/YYYY -> output as YYYY-MM-DD.
- well_name: PUITS field (e.g. "MD 336").
- field_name: CHAMPS field (e.g. "HMD").
- unit_name: APPAREIL field (e.g. "HRS 225", "HRL 8").
- operation_type: TYPE OPÉRATION value, top-left, French text.
- superintendent: the SH SUPERINTENDENT column value. The title looks like
  "Superintendant", "S intendant SH", or "SH/Superintendent". "" if that column is
  absent or its cell is blank.
- supervisor: the SH SUPERVISOR column value. The title looks like "SUPERVISEUR" /
  "SUPRVISEUR", "Superviseur SH", "Superviseur SH/DP", or "SH/SUPERVISOR" (the Sonatrach
  side). "" if absent or blank.
  superintendent and supervisor are DIFFERENT people in DIFFERENT cells -- NEVER merge them.
  ALWAYS skip these columns entirely: "Chef de poste" / "C.P", "Chef de chantier" / "C.C",
  "ENSP C/P", "ENSP C.Ch", "stagiaire", and the CONTRACTOR supervisor column
  (B.YL Supervisor, HALL Supervisor, Senior Oper). Keep ONLY the SH side.
  A single cell may itself hold two names joined with "/" or "+"; keep them together.
- vehicle_mat: "Véhicule" plate number, or "" if absent.
- activities: the DE / A / Opérations / Heures table. One row -> one object.
  - duration is decimal hours. If form shows "11:30" or "11H30", convert to 11.5.
- current_status: "Situation au rapport" / "Situation" value (terse current well state at
  report time).
- day_summary: "Résumé" / "Resume" text block found UNDER the operations table (the fuller
  summary of the day's work).
  NOTE: current_status and day_summary are intentionally mapped OPPOSITE to what the field
  names might suggest -- current_status = Situation cell, day_summary = Résumé block.
- next_day_summary: "Demain" / "À demain" / "Next day" value ("" if the form has no such row).
- program: "Programme prévu" / "Programme" / "Program" value.
- remark: "REMARQUES" / "Notes" panel content (Puits/Notes or Well/Pump), PLUS any
  below-operations annotation lines prefixed "N.B:" / "NB:" / "Note:" / "Notes:". Strip the
  prefix (and any leading "*") -- keep only the note text -- and join pieces with " / ".
  Do NOT include loose below-ops lines lacking such a prefix (Reste heures d'entretien/
  inspection, bare Ambulance / détecteur / line-de-vie lines, vehicle/logistics lists like
  FOURGON / GRUE).
- bop_test_date: date inside "Dernier test des BOP's ... Le: DD-MM-YYYY"
  -> output as YYYY-MM-DD.
- equipment: "Equipements SHDP" value (right-side panel). Often a long string
  of components separated by " + ".

Rules:
- Preserve French spelling and accents (é à ç è).
- If a field is empty / illegible, return "" for strings, [] for lists,
  null for dates.
- Times in HH:MM 24-hour format.
- Do NOT invent data. Empty is better than wrong.
- Output PURE JSON, no prose, no markdown fences.
"""


# ---------------------------------------------------------------------------
# Excel building
# ---------------------------------------------------------------------------

def _parse_hhmm(s: Optional[str]) -> Optional[time]:
    if not s:
        return None
    s = str(s).strip()
    m = re.match(r"^(\d{1,2})[:.hH](\d{2})(?:[:.](\d{2}))?$", s)
    if not m:
        return None
    h, mn = int(m.group(1)), int(m.group(2))
    if 0 <= h < 24 and 0 <= mn < 60:
        return time(h, mn)
    return None


def _duration_to_time(d: Optional[float]) -> Optional[time]:
    if d is None:
        return None
    try:
        d = float(d)
    except (TypeError, ValueError):
        return None
    if d <= 0:
        return None
    h = int(d)
    mn = int(round((d - h) * 60))
    if mn == 60:
        h += 1
        mn = 0
    if h >= 24:
        return None
    return time(h, mn)


def _parse_iso_date(s: Optional[str]) -> Optional[date]:
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def slugify(s: Optional[str], fallback: str = "unknown") -> str:
    if not s:
        return fallback
    out = re.sub(r"[^A-Za-z0-9_-]+", "_", str(s)).strip("_")
    return out or fallback


def safe_dir_name(s: str, fallback: str = "unknown") -> str:
    """Sanitize a string for use as a directory name on Windows/POSIX.

    Keeps letters, digits, dot, dash, underscore, and space. Replaces anything
    Windows forbids (<>:"/\\|?*) with '_'. Trims trailing dots / whitespace.
    """
    if not s:
        return fallback
    out = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", str(s))
    out = out.strip().rstrip(". ")
    return out or fallback


def build_fr_sh_xlsx(data: Dict[str, Any], out_path: Path) -> None:
    """Write a workbook whose 'RAP SH  ' sheet matches FR-SH parser expectations.

    Cell map (only what the parser reads):
      O3              date (datetime, parser checks isinstance(datetime))
      A5/B5           PUITS: / well_name
      C5/D5           CHAMPS: / field_name
      E5/F5           APPAREIL: / unit_name
      A7              "TYPE OPÉRATION" label
      A10             operation_type
      A16             "Dernier test des BOP's (X) Le: DD-MM-YYYY" with date inline
      A17/B17/C17/J17 DE / A / Opérations / Heures (activity header)
      A18+/B18+/C18+/J18+ activity rows (start, end, description, hours)
      K32/L32         "Equipements SHDP" / equipment value
      C35..C40        "Résumé: <line>" then continuation (from day_summary)
      K38             "REMARQUES:" (no space before colon -- matches parser cleanly)
      K41             remark text
      A45/B45         "Situation" / current_status
      A46/B46         "Demain" / next_day_summary
      A47/B47         "Programme" / program
      N47/N48         "Superintendant" / superintendent name
      P47/P48         "Superviseur SH" / supervisor name
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RAP SH  "  # trailing spaces deliberate; parser strips on lookup

    def s(r: int, c: int, v: Any) -> None:
        ws.cell(row=r, column=c).value = v

    # --- header / date ---
    s(1, 4, "RAPPORT JOURNALIER SNUBBING")
    s(1, 15, "DATE")
    d = _parse_iso_date(data.get("date"))
    if d is not None:
        s(3, 15, datetime(d.year, d.month, d.day))

    # --- PUITS / CHAMPS / APPAREIL row ---
    s(5, 1, "PUITS:")
    s(5, 2, data.get("well_name") or "")
    s(5, 3, "CHAMPS:")
    s(5, 4, data.get("field_name") or "")
    s(5, 5, "APPAREIL:")
    s(5, 6, data.get("unit_name") or "")

    # --- Operation type ---
    s(7, 1, " TYPE OPÉRATION")
    s(10, 1, data.get("operation_type") or "")

    # --- BOP test date (embedded in label string) ---
    bop = _parse_iso_date(data.get("bop_test_date"))
    if bop is not None:
        s(16, 1, f"Dernier test des BOP's Le: {bop.strftime('%d-%m-%Y')}")

    # --- Activities ---
    s(17, 1, "DE")
    s(17, 2, "A")
    s(17, 3, "Opérations")
    s(17, 10, "Heures")

    for i, act in enumerate(data.get("activities") or []):
        r = 18 + i
        if r > 33:
            break
        s(r, 1, _parse_hhmm(act.get("start_time")))
        s(r, 2, _parse_hhmm(act.get("end_time")))
        s(r, 3, act.get("description") or "")
        s(r, 10, _duration_to_time(act.get("duration")))

    # --- Equipment ---
    s(32, 11, "Equipements SHDP")
    s(32, 12, data.get("equipment") or "")

    # --- Résumé at C35, continuation C36..C40 (from day_summary; see SYSTEM_PROMPT) ---
    resume_text = (data.get("day_summary") or "").strip()
    if resume_text:
        lines = [ln for ln in resume_text.splitlines() if ln.strip()]
        first = lines[0] if lines else ""
        s(35, 3, f"Résumé: {first}".strip())
        for j, line in enumerate(lines[1:5], 1):
            s(35 + j, 3, line)

    # --- REMARQUES panel ---
    s(38, 11, "REMARQUES:")
    remark = (data.get("remark") or "").strip()
    if remark:
        s(41, 11, remark)

    # --- Situation / Demain / Programme ---
    s(45, 1, "Situation")
    s(45, 2, data.get("current_status") or "")
    s(46, 1, "Demain")
    s(46, 2, data.get("next_day_summary") or "")
    s(47, 1, "Programme")
    s(47, 2, data.get("program") or "")

    # --- Signature block ---
    # superintendent -> "Superintendant" column (N), supervisor -> "Superviseur
    # SH" column (P). The backend parser classifies these labels and reads the
    # name in the row below. Chef / contractor columns are intentionally absent.
    s(47, 14, "Superintendant")
    s(47, 16, "Superviseur SH")
    superintendent = (data.get("superintendent") or "").strip()
    supervisor = (data.get("supervisor") or "").strip()
    if superintendent:
        s(48, 14, superintendent)
    if supervisor:
        s(48, 16, supervisor)

    # --- Vehicle (only when value looks like an Algerian plate) ---
    veh = (data.get("vehicle_mat") or "").strip()
    if veh and re.search(r"\d{3}.{0,3}\d{2,3}.{0,3}\d{2,3}", veh):
        s(49, 16, "Véhicule")
        s(50, 16, veh)

    wb.save(out_path)
