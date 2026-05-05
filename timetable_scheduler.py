"""
timetable_scheduler.py
======================
Reads Planning for Timetable.xlsx and writes Timetable_Output.xlsx.

Logic (v2 — Assumption A):
  1. Class list answer already has fixed Day / Time / Venue from last term.
  2. System reads those fixed assignments and re-assigns Lecturer 1/2/3
     based on Teacher load table.
  3. Output: updated Class list answer + 5 daily timetable sheets.

Each new Term:
  - Update Teacher load table if teachers have changed.
  - Upload the file and run / click Generate.
"""

import os
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

import openpyxl

# ─── Paths ────────────────────────────────────────────────────────────────────

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE  = os.path.join(SCRIPT_DIR, "Planning for Timetable.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Timetable_Output.xlsx")

# ─── Sheet names ──────────────────────────────────────────────────────────────

DAY_TO_SHEET = {
    "Monday":    "Mon(Term1) ",   # trailing space is in the original file
    "Tuesday":   "Tue(Term1)",
    "Wednesday": "Wed(Term1)",
    "Thursday":  "Thu(Term1)",
    "Friday":    "Fri(Term1)",
}

# ─── Venue alias map ──────────────────────────────────────────────────────────
# Class list answer uses shorthand names; daily timetable sheets use full names.

_VENUE_ALIAS = {
    "WT - WT":  "WT - WT1",
    "KT - A":   "KT - KT1",
    "KT - B":   "KT - KT2",
    "KT - C":   "KT - KT3",
    "TM - A":   "TM - TM1",
    "TM - B":   "TM - TM2",
    "TM - C":   "TM - TM3",
    "TW - A":   "TW - TW1",
    "TW - B":   "TW - TW2",
    "TW - C":   "TW - TW3",
    "ST - A":   "ST - ST1",
    "ST - B":   "ST - ST2",
    "ST - C":   "ST - ST3",
    "TK - TK":  "TKO - TKO",
}

def _resolve_venue(venue: str) -> str:
    return _VENUE_ALIAS.get(str(venue).strip(), str(venue).strip())

# ─── Per-sheet dynamic column positions ───────────────────────────────────────

_MARKER_TO_SLOT = {
    "0900": "0900 - 1100",  900:  "0900 - 1100",
    "1100": "1100 - 1300",  1100: "1100 - 1300",
    "1400": "1400 - 1600",  1400: "1400 - 1600",
    "1600": "1600 - 1800",  1600: "1600 - 1800",
}

def _slot_cols_for_sheet(ws) -> dict:
    """Return {slot_key: start_col_1based} by reading row 1 of the sheet."""
    row1  = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    result = {}
    for i, val in enumerate(row1):
        slot = _MARKER_TO_SLOT.get(val)
        if slot:
            result[slot] = i + 1
    return result

# ─── Subject name normalisation ───────────────────────────────────────────────

_SUBJECT_NAME_MAP = {
    "MathsPlus":               "Maths Plus",
    "Career and Life Learning": "Career and Life Planning",
}

# ─── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class TeacherEntry:
    name:   str
    quotas: dict = field(default_factory=dict)  # {subject_name: max_lec1_classes}

@dataclass
class ClassAssignment:
    code:         str
    subject_name: str
    subject_cn:   str
    loading:      int
    students:     int
    day:          str
    time1:        str
    time2:        Optional[str]
    venue:        str               # as written in Class list answer
    venue_room:   str               # resolved to daily-sheet room name
    lecturer1:    Optional[str] = None
    lecturer2:    Optional[str] = None
    lecturer3:    Optional[str] = None

# ─── Phase 1 — Read inputs ────────────────────────────────────────────────────

def read_class_assignments(wb) -> list:
    """
    Read the fixed schedule (Day/Time/Venue) from Class list answer.
    Only rows with a course code AND a day are included.
    """
    ws = wb["Class list answer"]
    assignments = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[1]
        day  = row[8]
        if not code or not day:
            continue
        venue     = str(row[11]).strip() if row[11] else ""
        loading   = 4  # default; will match from Class list if available
        students  = 0

        assignments.append(ClassAssignment(
            code         = str(code).strip(),
            subject_name = str(row[2]).strip() if row[2] else "",
            subject_cn   = str(row[3]).strip() if row[3] else "",
            loading      = loading,
            students     = students,
            day          = str(day).strip(),
            time1        = str(row[9]).strip() if row[9] else "",
            time2        = str(row[10]).strip() if row[10] else None,
            venue        = venue,
            venue_room   = _resolve_venue(venue),
        ))
    return assignments


def _enrich_from_class_list(wb, assignments: list):
    """Fill in loading and student count from the Class list sheet."""
    ws = wb["Class list"]
    info = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[0]
        if not code:
            continue
        loading  = int(row[6]) if isinstance(row[6], (int, float)) else 4
        students = int(row[7]) if isinstance(row[7], (int, float)) else 0
        info[str(code).strip()] = (loading, students)

    for ca in assignments:
        if ca.code in info:
            ca.loading, ca.students = info[ca.code]


def read_teachers(wb) -> list:
    """Read Teacher load table. Combine both sections (summed quotas)."""
    ws          = wb["Teacher load table with subject"]
    raw_headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    headers     = [_SUBJECT_NAME_MAP.get(h, h) for h in raw_headers]
    subj_cols   = headers[2:]

    teachers: dict = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[1]
        if not name or not str(name).strip():
            continue
        name = str(name).strip()
        if name not in teachers:
            teachers[name] = TeacherEntry(name=name)
        for i, subj in enumerate(subj_cols):
            if not subj:
                continue
            val = row[i + 2] if (i + 2) < len(row) else None
            if isinstance(val, (int, float)) and val > 0:
                capped = min(int(val), 20)
                teachers[name].quotas[subj] = (
                    teachers[name].quotas.get(subj, 0) + capped
                )
    return list(teachers.values())

# ─── Phase 2 — Assign lecturers ───────────────────────────────────────────────

def assign_lecturers(assignments: list, teachers: list) -> tuple:
    """
    Assign Lecturer 1 (primary, respects quota) and
    Lecturer 2/3 (backups, just different people — no quota check).

    Returns (assignments_with_lecturers, unassigned_codes).
    """
    # Build subject -> teachers sorted by quota desc
    by_subject: dict = defaultdict(list)
    for t in teachers:
        for subj, quota in t.quotas.items():
            if quota > 0:
                by_subject[subj].append(t)
    for subj in by_subject:
        by_subject[subj].sort(key=lambda t: t.quotas.get(subj, 0), reverse=True)

    # Remaining Lec1 quota per (teacher, subject)
    remaining: dict = {}
    for t in teachers:
        for subj, quota in t.quotas.items():
            remaining[(t.name, subj)] = quota

    # Group assignments by subject
    grouped: dict = defaultdict(list)
    for ca in assignments:
        grouped[ca.subject_name].append(ca)

    unassigned: list = []

    for subj, group in grouped.items():
        eligible = by_subject.get(subj, [])
        if not eligible:
            unassigned.extend(ca.code for ca in group)
            continue

        p_idx = 0   # round-robin pointer for Lec1

        for ca in group:
            # ── Lecturer 1 (primary, respects quota) ──
            lec1 = None
            for offset in range(len(eligible)):
                t = eligible[(p_idx + offset) % len(eligible)]
                if remaining.get((t.name, subj), 0) > 0:
                    lec1 = t.name
                    remaining[(t.name, subj)] -= 1
                    p_idx = (p_idx + offset + 1) % len(eligible)
                    break
            if lec1 is None:
                # Quota exhausted — overflow to first teacher
                lec1 = eligible[p_idx % len(eligible)].name
                p_idx = (p_idx + 1) % len(eligible)
                unassigned.append(ca.code)

            ca.lecturer1 = lec1

            # ── Lecturer 2 & 3 (backups — any other eligible teacher) ──
            others = [t.name for t in eligible if t.name != lec1]
            ca.lecturer2 = others[0] if len(others) > 0 else None
            ca.lecturer3 = others[1] if len(others) > 1 else None

    return assignments, unassigned

# ─── Phase 3 — Write output ───────────────────────────────────────────────────

def write_output_wb(wb, assignments: list):
    """
    Write Lec1/2/3 into Class list answer, then regenerate 5 daily sheets.
    Preserves all other existing data (other subjects, non-scheduled slots).
    """
    # ── 3a. Update Class list answer (only Lec1/2/3 columns) ────────────────
    ws     = wb["Class list answer"]
    lec_by_code = {ca.code: ca for ca in assignments}

    for row in ws.iter_rows(min_row=2):
        code = row[1].value   # col B
        if not code or str(code).strip() not in lec_by_code:
            continue
        ca         = lec_by_code[str(code).strip()]
        row[4].value = ca.lecturer1   # col E
        row[5].value = ca.lecturer2   # col F
        row[6].value = ca.lecturer3   # col G

    # ── 3b. Build slot lookup: (day, venue_room, slot) -> ClassAssignment ───
    slot_lookup: dict = {}
    for ca in assignments:
        if ca.day and ca.time1 and ca.venue_room:
            slot_lookup[(ca.day, ca.venue_room, ca.time1)] = ca
            if ca.time2:
                slot_lookup[(ca.day, ca.venue_room, ca.time2)] = ca

    # ── 3c. Regenerate 5 daily timetable sheets ──────────────────────────────
    for day, sheet_name in DAY_TO_SHEET.items():
        if sheet_name not in wb.sheetnames:
            print(f"  WARNING: sheet '{sheet_name}' not found — skipping {day}")
            continue
        ws = wb[sheet_name]

        # Build room -> row number map from col A (rows 3-37 only)
        room_row_map: dict = {}
        for row in ws.iter_rows(min_row=3, max_row=37, max_col=1):
            val = row[0].value
            if val and isinstance(val, str) and " - " in val:
                room_row_map[val] = row[0].row

        # Write each time slot
        for slot, start_col in _slot_cols_for_sheet(ws).items():
            for room_code, row_num in room_row_map.items():
                ca = slot_lookup.get((day, room_code, slot))
                if ca:
                    ws.cell(row_num, start_col,     ca.code)
                    ws.cell(row_num, start_col + 1, ca.students)
                    ws.cell(row_num, start_col + 2, ca.subject_cn)
                    ws.cell(row_num, start_col + 3, ca.lecturer1)
                    # Time col (start_col+4): leave blank for standard slots


def write_output(input_path: str, output_path: str, assignments: list):
    """CLI helper: open template, fill, save to disk."""
    wb = openpyxl.load_workbook(input_path)
    write_output_wb(wb, assignments)
    wb.save(output_path)
    print(f"\n  Saved: {output_path}")

# ─── Validation ───────────────────────────────────────────────────────────────

def collect_stats(assignments: list, teachers: list, unassigned: list) -> dict:
    """Return validation results as a plain dict for the web API."""
    total    = len(assignments)
    assigned = sum(1 for ca in assignments if ca.lecturer1)
    lec2_cov = sum(1 for ca in assignments if ca.lecturer2)
    lec3_cov = sum(1 for ca in assignments if ca.lecturer3)

    # Teacher loading summary
    lec1_count: dict = defaultdict(int)
    for ca in assignments:
        if ca.lecturer1:
            lec1_count[ca.lecturer1] += 1

    quota_map: dict = {}
    for t in teachers:
        total_q = sum(t.quotas.values())
        if total_q > 0:
            quota_map[t.name] = total_q

    loading_ok   = [n for n, cnt in lec1_count.items()
                    if abs(cnt - quota_map.get(n, cnt)) <= 1]
    loading_over = [f"{n}: assigned {cnt}, quota {quota_map.get(n,'?')}"
                    for n, cnt in lec1_count.items()
                    if cnt > quota_map.get(n, cnt) + 1]

    return {
        "scheduled":     assigned,
        "total_classes": total,
        "unassigned":    unassigned,
        "violations":    loading_over,
        "lec2_coverage": lec2_cov,
        "lec3_coverage": lec3_cov,
        "preferred_pct": 100,  # times are fixed, not re-scheduled
        "centre_dist":   {},   # not applicable under assumption A
    }


def validate_and_report(assignments: list, teachers: list, unassigned: list):
    stats = collect_stats(assignments, teachers, unassigned)
    total = stats["total_classes"]
    sched = stats["scheduled"]

    sep = "=" * 62
    print(f"\n{sep}")
    print("  TIMETABLE SCHEDULER — VALIDATION REPORT")
    print(sep)
    print(f"\n  Classes with Lecturer 1 assigned : {sched}/{total}")
    print(f"  Lecturer 2 coverage              : {stats['lec2_coverage']}/{total}")
    print(f"  Lecturer 3 coverage              : {stats['lec3_coverage']}/{total}")

    if unassigned:
        print(f"\n  WARNING — no eligible teacher for ({len(unassigned)}):")
        for code in sorted(unassigned):
            print(f"    - {code}")

    if stats["violations"]:
        print("\n  Loading exceeded:")
        for msg in stats["violations"]:
            print(f"    - {msg}")
    else:
        print("\n  Loading: OK — all within quota")

    status = "GO" if not stats["violations"] and not unassigned else "REVIEW NEEDED"
    print(f"\n  Overall status: {status}")
    print(sep)

# ─── Web API entry point ──────────────────────────────────────────────────────

def run_from_bytes(excel_bytes: bytes) -> tuple:
    """
    Process an Excel file given as raw bytes.
    Returns (output_excel_bytes, stats_dict).
    """
    from io import BytesIO

    wb_read     = openpyxl.load_workbook(BytesIO(excel_bytes), data_only=True)
    assignments = read_class_assignments(wb_read)
    _enrich_from_class_list(wb_read, assignments)
    teachers    = read_teachers(wb_read)

    assignments, unassigned = assign_lecturers(assignments, teachers)

    wb_out = openpyxl.load_workbook(BytesIO(excel_bytes))
    write_output_wb(wb_out, assignments)

    buf = BytesIO()
    wb_out.save(buf)
    output_bytes = buf.getvalue()

    stats = collect_stats(assignments, teachers, unassigned)
    return output_bytes, stats

# ─── CLI entry point ──────────────────────────────────────────────────────────

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: Input file not found:\n  {INPUT_FILE}")
        sys.exit(1)

    print(f"Input : {INPUT_FILE}")
    print(f"Output: {OUTPUT_FILE}\n")

    wb_read = openpyxl.load_workbook(INPUT_FILE, data_only=True)

    print("[1/4] Reading fixed schedule from Class list answer...")
    assignments = read_class_assignments(wb_read)
    _enrich_from_class_list(wb_read, assignments)
    print(f"      {len(assignments)} classes with fixed Day/Time/Venue")

    print("[2/4] Reading teacher pool...")
    teachers = read_teachers(wb_read)
    print(f"      {len(teachers)} teachers")

    print("[3/4] Assigning Lecturer 1 / 2 / 3...")
    assignments, unassigned = assign_lecturers(assignments, teachers)
    assigned = sum(1 for ca in assignments if ca.lecturer1)
    print(f"      {assigned}/{len(assignments)} classes assigned Lecturer 1")

    print("[4/4] Writing output...")
    write_output(INPUT_FILE, OUTPUT_FILE, assignments)

    validate_and_report(assignments, teachers, unassigned)


if __name__ == "__main__":
    main()
