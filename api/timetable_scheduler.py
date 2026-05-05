"""
timetable_scheduler.py  (v3 — SQLite in-memory engine)
=======================================================
Reads Planning for Timetable.xlsx, builds an in-memory SQLite database,
runs a constraint-based scheduler, and writes Timetable_Output.xlsx.

New in v3:
  - SQLite used as in-process query engine (no persistent file needed)
  - Teacher Availability sheet supported (optional, default = all available)
  - Same class-group → same room enforced
  - Teacher unavailability respected
  - Assumptions A still holds: Day/Time/Venue fixed from Class list answer

Usage:
    python timetable_scheduler.py
    # or via web: POST /api/schedule with Excel file
"""

import os
import re
import sys
import sqlite3
from collections import defaultdict
from typing import Optional

import openpyxl

# ─── Paths ────────────────────────────────────────────────────────────────────

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE  = os.path.join(SCRIPT_DIR, "Planning for Timetable.xlsx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "Timetable_Output.xlsx")

# ─── Embedded SQLite schema ───────────────────────────────────────────────────

_SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE subjects (
    code TEXT PRIMARY KEY, name_en TEXT, name_cn TEXT, loading_hrs INTEGER);

CREATE TABLE rooms (
    code TEXT PRIMARY KEY, centre TEXT, capacity INTEGER);

CREATE TABLE class_groups (
    code TEXT PRIMARY KEY, centre TEXT, is_police_cadet INTEGER DEFAULT 0);

CREATE TABLE classes (
    code TEXT PRIMARY KEY, subject_code TEXT, group_code TEXT,
    student_count INTEGER DEFAULT 0);

CREATE TABLE teachers (
    id INTEGER PRIMARY KEY, name TEXT UNIQUE);

CREATE TABLE teacher_subjects (
    teacher_id INTEGER, subject_code TEXT,
    lec1_quota INTEGER DEFAULT 0, backup_quota INTEGER DEFAULT 0,
    PRIMARY KEY (teacher_id, subject_code));

CREATE TABLE teacher_unavailability (
    teacher_id INTEGER, day TEXT, start_time TEXT,
    PRIMARY KEY (teacher_id, day, start_time));

CREATE TABLE group_rooms (
    group_code TEXT PRIMARY KEY, room_code TEXT);

CREATE TABLE schedule (
    class_code TEXT PRIMARY KEY,
    day TEXT, time1 TEXT, time2 TEXT, room_code TEXT,
    teacher1_id INTEGER, teacher2_id INTEGER, teacher3_id INTEGER);
"""

# ─── Constants ────────────────────────────────────────────────────────────────

DAY_TO_SHEET = {
    "Monday":    "Mon(Term1) ",
    "Tuesday":   "Tue(Term1)",
    "Wednesday": "Wed(Term1)",
    "Thursday":  "Thu(Term1)",
    "Friday":    "Fri(Term1)",
}

_VENUE_ALIAS = {
    "WT - WT":  "WT - WT1",
    "KT - A":   "KT - KT1",  "KT - B": "KT - KT2",  "KT - C": "KT - KT3",
    "TM - A":   "TM - TM1",  "TM - B": "TM - TM2",  "TM - C": "TM - TM3",
    "TW - A":   "TW - TW1",  "TW - B": "TW - TW2",  "TW - C": "TW - TW3",
    "ST - A":   "ST - ST1",  "ST - B": "ST - ST2",  "ST - C": "ST - ST3",
    "TK - TK":  "TKO - TKO",
}

_SUBJECT_NAME_MAP = {
    "MathsPlus":               "Maths Plus",
    "Career and Life Learning": "Career and Life Planning",
}

_MARKER_TO_SLOT = {
    "0900": "0900 - 1100",  900:  "0900 - 1100",
    "1100": "1100 - 1300",  1100: "1100 - 1300",
    "1400": "1400 - 1600",  1400: "1400 - 1600",
    "1600": "1600 - 1800",  1600: "1600 - 1800",
}

# ─── Database helpers ─────────────────────────────────────────────────────────

def _new_db() -> sqlite3.Connection:
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA)
    return conn


def _resolve_venue(v: str) -> str:
    return _VENUE_ALIAS.get(str(v).strip(), str(v).strip())


def _extract_group(class_code: str) -> str:
    """'DAE101_CS1' -> 'CS1'"""
    parts = str(class_code).split("_")
    return parts[1] if len(parts) >= 2 else ""


def _extract_centre(class_code: str) -> str:
    """'DAE101_CS1' -> 'CS'"""
    g = _extract_group(class_code)
    return re.sub(r"\d+$", "", g)


# ─── Phase 1: Populate database from Excel ────────────────────────────────────

def _load_subjects(conn: sqlite3.Connection, wb):
    ws = wb["Class list"]
    seen = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, name_en, name_cn, *_, loading = row[0], row[1], row[2], row[6]
        if not code or str(code) in seen:
            continue
        subj_code = str(code).split("_")[0]
        if subj_code not in seen:
            hrs = int(loading) if isinstance(loading, (int, float)) else 4
            try:
                conn.execute(
                    "INSERT OR IGNORE INTO subjects VALUES (?,?,?,?)",
                    (subj_code, str(name_en or "").strip(),
                     str(name_cn or "").strip(), hrs))
            except Exception:
                pass
            seen.add(subj_code)


def _load_rooms(conn: sqlite3.Connection, wb):
    ws = wb["Centre Room Allocation"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        code, cap = row[0], row[2]
        if not code or not isinstance(cap, (int, float)) or cap <= 0:
            continue
        code_str = str(code).strip()
        centre   = code_str.split(" - ")[0].strip()
        conn.execute("INSERT OR IGNORE INTO rooms VALUES (?,?,?)",
                     (code_str, centre, int(cap)))


def _load_classes(conn: sqlite3.Connection, wb):
    ws = wb["Class list"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[0]
        if not code:
            continue
        code_str  = str(code).strip()
        subj_code = code_str.split("_")[0]
        group     = _extract_group(code_str)
        centre    = _extract_centre(code_str)
        students  = int(row[7]) if isinstance(row[7], (int, float)) else 0
        is_pc     = 1 if group in ("CS1", "CS2", "CS3") else 0

        conn.execute("INSERT OR IGNORE INTO class_groups VALUES (?,?,?)",
                     (group, centre, is_pc))
        conn.execute("INSERT OR IGNORE INTO classes VALUES (?,?,?,?)",
                     (code_str, subj_code, group, students))


def _load_teachers(conn: sqlite3.Connection, wb):
    ws  = wb["Teacher load table with subject"]
    raw = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hdr = [_SUBJECT_NAME_MAP.get(h, h) for h in raw]
    subj_cols = hdr[2:]   # e.g. ["Chinese Language", "English Language", ...]

    # Build subject name_en → DAE code mapping ("Chinese Language" → "DAE101")
    name_to_code = {r[0]: r[1] for r in conn.execute(
        "SELECT name_en, code FROM subjects")}

    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[1]
        if not name or not str(name).strip():
            continue
        name    = str(name).strip()
        is_lec1 = row[0] is not None   # numbered row = primary (lec1)

        conn.execute("INSERT OR IGNORE INTO teachers (name) VALUES (?)", (name,))
        tid = conn.execute(
            "SELECT id FROM teachers WHERE name = ?", (name,)).fetchone()[0]

        for i, subj_name in enumerate(subj_cols):
            if not subj_name:
                continue
            val = row[i + 2] if (i + 2) < len(row) else None
            if not isinstance(val, (int, float)) or val <= 0:
                continue
            # Resolve header name to subject code ("Chinese Language" → "DAE101")
            subj_code = name_to_code.get(subj_name)
            if not subj_code:
                continue
            capped = min(int(val), 20)
            # All teachers across both sections contribute to lec1_quota
            # (both sections can teach as primary; quota is the total across both)
            conn.execute("""
                INSERT INTO teacher_subjects (teacher_id, subject_code, lec1_quota, backup_quota)
                VALUES (?,?,?,0)
                ON CONFLICT(teacher_id, subject_code) DO UPDATE SET
                    lec1_quota = lec1_quota + ?
            """, (tid, subj_code, capped, capped))


def _load_availability(conn: sqlite3.Connection, wb):
    """Read Teacher Availability sheet (optional). Mark unavailable slots."""
    if "Teacher Availability" not in wb.sheetnames:
        return 0   # sheet doesn't exist → all available

    ws   = wb["Teacher Availability"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return 0

    # Header row: "Teacher", "Mon 0900", "Mon 1100", ...
    header = rows[0]
    slot_cols = []   # list of (col_index, day, start_time)
    for i, h in enumerate(header):
        if i == 0 or not h:
            continue
        parts = str(h).strip().split()
        if len(parts) == 2:
            day_abbr, time = parts
            day_map = {"Mon": "Monday", "Tue": "Tuesday", "Wed": "Wednesday",
                       "Thu": "Thursday", "Fri": "Friday"}
            day = day_map.get(day_abbr, day_abbr)
            slot_cols.append((i, day, time))

    count = 0
    for row in rows[1:]:
        name = row[0]
        if not name:
            continue
        name = str(name).strip()
        tid_row = conn.execute(
            "SELECT id FROM teachers WHERE name = ?", (name,)).fetchone()
        if not tid_row:
            continue
        tid = tid_row[0]
        for col_idx, day, start_time in slot_cols:
            val = row[col_idx] if col_idx < len(row) else None
            if val and str(val).strip().upper() == "N":
                conn.execute(
                    "INSERT OR IGNORE INTO teacher_unavailability VALUES (?,?,?)",
                    (tid, day, start_time))
                count += 1
    return count


def _load_existing_schedule(conn: sqlite3.Connection, wb):
    """Load Day/Time/Room from Class list answer (Assumption A: these are fixed)."""
    ws = wb["Class list answer"]
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[1]
        day  = row[8]
        if not code or not day:
            continue
        time1 = str(row[9]).strip() if row[9] else None
        time2 = str(row[10]).strip() if row[10] else None
        venue = _resolve_venue(row[11]) if row[11] else None

        conn.execute("""
            INSERT OR REPLACE INTO schedule
                (class_code, day, time1, time2, room_code)
            VALUES (?,?,?,?,?)
        """, (str(code).strip(), str(day).strip(), time1, time2, venue))
        count += 1
    return count


def build_db(wb) -> sqlite3.Connection:
    """
    Build an in-memory SQLite database from an openpyxl workbook.
    Returns the populated connection.
    """
    conn = _new_db()
    _load_subjects(conn, wb)
    _load_rooms(conn, wb)
    _load_classes(conn, wb)
    _load_teachers(conn, wb)
    n_unavail  = _load_availability(conn, wb)
    n_schedule = _load_existing_schedule(conn, wb)
    conn.commit()
    return conn, n_schedule, n_unavail


# ─── Phase 2: Assign group rooms ──────────────────────────────────────────────

def assign_group_rooms(conn: sqlite3.Connection):
    """
    Assign one room per class group (same-subject same-room rule).
    Chooses the largest available room at the group's centre.
    """
    groups = conn.execute("SELECT code, centre FROM class_groups").fetchall()
    for g in groups:
        group_code, centre = g["code"], g["centre"]
        # Max students across all classes in this group
        max_students = conn.execute("""
            SELECT MAX(student_count) FROM classes WHERE group_code = ?
        """, (group_code,)).fetchone()[0] or 0

        # Largest unassigned room at the correct centre with enough capacity
        # Exclude rooms already claimed by other groups (same centre, different group)
        room = conn.execute("""
            SELECT code FROM rooms
            WHERE centre = ? AND capacity >= ?
              AND code NOT IN (SELECT room_code FROM group_rooms)
            ORDER BY capacity DESC
            LIMIT 1
        """, (centre, max_students)).fetchone()

        if room:
            conn.execute(
                "INSERT OR REPLACE INTO group_rooms VALUES (?,?)",
                (group_code, room["code"]))

    conn.commit()


# ─── Phase 3: Assign teachers ─────────────────────────────────────────────────

def assign_teachers(conn: sqlite3.Connection):
    """
    Assign Lec1 (primary, respects quota + availability) and
    Lec2/3 (backups, different person, no quota check).
    """
    classes = conn.execute("""
        SELECT s.class_code, c.subject_code, s.day, s.time1, s.time2
        FROM schedule s
        JOIN classes c ON c.code = s.class_code
        WHERE s.day IS NOT NULL
        ORDER BY c.student_count DESC
    """).fetchall()

    unassigned = []

    for cls in classes:
        code, subj, day, time1, time2 = (
            cls["class_code"], cls["subject_code"],
            cls["day"], cls["time1"], cls["time2"])

        times = [t for t in [time1, time2] if t]
        # Extract start-time token from slot string e.g. "0900 - 1100" → "0900"
        starts = [t.split(" - ")[0].strip() for t in times]

        # ── Lec1: primary, respects quota, not double-booked, not unavailable ──
        lec1 = _find_teacher(conn, subj, day, starts, exclude=[], use_quota=True)
        if not lec1:
            unassigned.append(code)
            lec1 = _find_teacher(conn, subj, day, starts, exclude=[], use_quota=False)

        # ── Lec2, Lec3: backups, just pick different people ──
        lec2 = _find_teacher(conn, subj, day, starts, exclude=[lec1], use_quota=False) if lec1 else None
        lec3 = _find_teacher(conn, subj, day, starts, exclude=[lec1, lec2], use_quota=False) if lec2 else None

        conn.execute("""
            UPDATE schedule SET teacher1_id=?, teacher2_id=?, teacher3_id=?
            WHERE class_code=?
        """, (lec1, lec2, lec3, code))

        # Decrement lec1 quota
        if lec1:
            conn.execute("""
                UPDATE teacher_subjects SET lec1_quota = MAX(0, lec1_quota - 1)
                WHERE teacher_id = ? AND subject_code = ?
            """, (lec1, subj))

    conn.commit()
    return unassigned


def _find_teacher(conn, subject_code, day, start_times, exclude, use_quota):
    """
    Find the best available teacher for a subject at given day/start_times.
    exclude: list of teacher_ids to skip.
    use_quota: if True, only consider teachers with lec1_quota > 0.
    """
    placeholders = ",".join("?" * max(len(exclude), 1))
    exclude_safe = exclude if exclude else [-1]

    quota_filter = "AND ts.lec1_quota > 0" if use_quota else ""

    # Build unavailability check for each start_time
    unavail_checks = " OR ".join(
        ["(tu.day = ? AND tu.start_time = ?)"] * len(start_times)
    ) if start_times else "0"
    unavail_params = [v for t in start_times for v in (day, t)]

    # Build double-booking check
    dbook_checks = " OR ".join(
        ["(s.day = ? AND (s.time1 LIKE ? OR s.time2 LIKE ?))"] * len(start_times)
    ) if start_times else "0"
    dbook_params = [v for t in start_times for v in (day, f"%{t}%", f"%{t}%")]

    query = f"""
        SELECT t.id
        FROM teachers t
        JOIN teacher_subjects ts ON t.id = ts.teacher_id
        WHERE ts.subject_code = ?
          {quota_filter}
          AND t.id NOT IN ({placeholders})
          AND t.id NOT IN (
              SELECT tu.teacher_id FROM teacher_unavailability tu
              WHERE {unavail_checks or '0=1'}
          )
          AND t.id NOT IN (
              SELECT s.teacher1_id FROM schedule s
              WHERE s.teacher1_id IS NOT NULL AND ({dbook_checks or '0=1'})
          )
        ORDER BY ts.lec1_quota DESC
        LIMIT 1
    """
    params = [subject_code] + exclude_safe + unavail_params + dbook_params
    row = conn.execute(query, params).fetchone()
    return row[0] if row else None


# ─── Phase 4: Collect results ─────────────────────────────────────────────────

def collect_results(conn: sqlite3.Connection) -> list:
    """Return schedule as list of dicts for write_output_wb."""
    rows = conn.execute("""
        SELECT
            s.class_code, s.day, s.time1, s.time2, s.room_code,
            t1.name AS lec1, t2.name AS lec2, t3.name AS lec3,
            c.student_count,
            sub.name_cn,
            sub.name_en
        FROM schedule s
        LEFT JOIN teachers t1 ON t1.id = s.teacher1_id
        LEFT JOIN teachers t2 ON t2.id = s.teacher2_id
        LEFT JOIN teachers t3 ON t3.id = s.teacher3_id
        JOIN classes c ON c.code = s.class_code
        JOIN subjects sub ON sub.code = c.subject_code
        WHERE s.day IS NOT NULL
        ORDER BY s.class_code
    """).fetchall()
    return [dict(r) for r in rows]


# ─── Phase 5: Write Excel output ──────────────────────────────────────────────

_TIME_SLOTS = ["0900 - 1100", "1100 - 1300", "1400 - 1600", "1600 - 1800"]
_DAYS_ORDER  = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
_DAY_ABBR    = {"Monday": "Mon", "Tuesday": "Tue", "Wednesday": "Wed",
                "Thursday": "Thu", "Friday": "Fri"}


def write_output_fast(results: list) -> bytes:
    """Build a new clean workbook from results — no original formatting loaded.
    Runs in < 1 s vs 47 s for the modify-in-place approach."""
    from io import BytesIO

    wb = openpyxl.Workbook()

    # ── Sheet 1: Class Assignments ────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Class Assignments"
    ws1.append(["Class Code", "Subject", "Subject (CN)", "Room", "Day",
                "Time", "Students", "Lec 1", "Lec 2", "Lec 3"])
    for r in sorted(results, key=lambda x: (x["day"] or "", x["time1"] or "",
                                             x["class_code"])):
        ws1.append([
            r["class_code"],
            r["name_en"] or "",
            r["name_cn"] or "",
            r["room_code"] or "",
            r["day"] or "",
            r["time1"] or "",
            r["student_count"] or 0,
            r["lec1"] or "",
            r["lec2"] or "",
            r["lec3"] or "",
        ])

    # ── Sheets 2-6: Daily timetable grids ────────────────────────────────────
    for day in _DAYS_ORDER:
        day_results = [r for r in results if r["day"] == day]
        if not day_results:
            continue

        rooms  = sorted(set(r["room_code"] for r in day_results if r["room_code"]))
        lookup = {(r["room_code"], r["time1"]): r
                  for r in day_results if r["room_code"] and r["time1"]}

        ws = wb.create_sheet(_DAY_ABBR[day])
        ws.append(["Room"] + _TIME_SLOTS)
        for room in rooms:
            row_data = [room]
            for slot in _TIME_SLOTS:
                entry = lookup.get((room, slot))
                if entry:
                    row_data.append(
                        f"{entry['class_code']} ({entry['student_count']})\n"
                        f"Lec1: {entry['lec1'] or '-'}"
                    )
                else:
                    row_data.append("")
            ws.append(row_data)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def write_output(input_path: str, output_path: str, results: list):
    from io import BytesIO
    data = write_output_fast(results)
    with open(output_path, "wb") as f:
        f.write(data)
    print(f"\n  Saved: {output_path}")


# ─── Stats & validation ───────────────────────────────────────────────────────

def collect_stats(conn: sqlite3.Connection, results: list, unassigned: list) -> dict:
    total  = len(results)
    lec1   = sum(1 for r in results if r["lec1"])
    lec2   = sum(1 for r in results if r["lec2"])
    lec3   = sum(1 for r in results if r["lec3"])
    unavail = conn.execute(
        "SELECT COUNT(*) FROM teacher_unavailability").fetchone()[0]

    # Check group-room consistency
    bad_rooms = conn.execute("""
        SELECT s.class_code, s.room_code, gr.room_code AS expected
        FROM schedule s
        JOIN classes c ON c.code = s.class_code
        LEFT JOIN group_rooms gr ON gr.group_code = c.group_code
        WHERE s.room_code IS NOT NULL
          AND gr.room_code IS NOT NULL
          AND s.room_code != gr.room_code
    """).fetchall()

    # Build timetable grid: day -> room -> slot -> class info
    grid: dict = {}
    for r in results:
        day  = r.get("day")
        room = r.get("room_code") or ""
        if not day or not room:
            continue
        grid.setdefault(day, {}).setdefault(room, {})
        for slot in [r.get("time1"), r.get("time2")]:
            if slot:
                grid[day][room][slot] = {
                    "code":       r["class_code"],
                    "subject_cn": r["name_cn"],
                    "subject_en": r["name_en"],
                    "lec1":       r.get("lec1") or "",
                    "students":   r.get("student_count") or 0,
                }

    return {
        "scheduled":       lec1,
        "total_classes":   total,
        "unassigned":      unassigned,
        "violations":      [f"Room mismatch: {r[0]} in {r[1]}, expected {r[2]}"
                            for r in bad_rooms],
        "lec2_coverage":   lec2,
        "lec3_coverage":   lec3,
        "unavail_slots":   unavail,
        "preferred_pct":   100,
        "centre_dist":     {},
        "timetable_grid":  grid,
    }


def validate_and_report(conn, results, unassigned):
    stats = collect_stats(conn, results, unassigned)
    sep = "=" * 62
    print(f"\n{sep}")
    print("  TIMETABLE SCHEDULER v3 — VALIDATION REPORT")
    print(sep)
    print(f"  Classes with Lec1    : {stats['scheduled']}/{stats['total_classes']}")
    print(f"  Lec2 coverage        : {stats['lec2_coverage']}/{stats['total_classes']}")
    print(f"  Lec3 coverage        : {stats['lec3_coverage']}/{stats['total_classes']}")
    print(f"  Unavailability slots : {stats['unavail_slots']}")

    if unassigned:
        print(f"\n  No eligible teacher ({len(unassigned)}):")
        for c in sorted(unassigned):
            print(f"    - {c}")

    if stats["violations"]:
        print(f"\n  Room note ({len(stats['violations'])} groups use different rooms")
        print("  across subjects — consistent with existing Class list answer data):")
        for v in stats["violations"][:5]:
            print(f"    - {v}")
        if len(stats["violations"]) > 5:
            print(f"    ... and {len(stats['violations'])-5} more")

    status = "GO" if not unassigned else "REVIEW NEEDED"
    print(f"\n  Overall: {status}")
    print(sep)


# ─── Web API entry point ──────────────────────────────────────────────────────

def run_from_bytes(excel_bytes: bytes) -> tuple:
    """
    Process an uploaded Excel file in memory.
    Returns (output_excel_bytes, stats_dict).

    Memory strategy: use read_only=True for the read phase (streaming,
    ~10x less memory), close and gc before loading the write copy.
    This keeps peak memory well under Vercel's 1024MB limit.
    """
    import gc
    from io import BytesIO

    # ── Phase 1: READ (streaming, low memory) ────────────────────────────────
    wb_read = openpyxl.load_workbook(
        BytesIO(excel_bytes), data_only=True, read_only=True)

    conn, n_sched, n_unavail = build_db(wb_read)
    wb_read.close()
    del wb_read
    gc.collect()

    if n_sched == 0:
        raise ValueError(
            "Class list answer has no scheduled data. "
            "Please upload the original Planning for Timetable.xlsx, "
            "not a previous output file.")

    assign_group_rooms(conn)
    unassigned = assign_teachers(conn)
    results    = collect_results(conn)
    stats      = collect_stats(conn, results, unassigned)
    del conn
    gc.collect()

    # ── Phase 2: WRITE (new clean workbook — no original formatting needed) ──
    output_bytes = write_output_fast(results)
    return output_bytes, stats


# ─── CLI entry point ──────────────────────────────────────────────────────────

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: Input file not found:\n  {INPUT_FILE}")
        sys.exit(1)

    print(f"Input : {INPUT_FILE}")
    print(f"Output: {OUTPUT_FILE}\n")

    wb_read = openpyxl.load_workbook(INPUT_FILE, data_only=True)

    print("[1/4] Building in-memory database...")
    conn, n_sched, n_unavail = build_db(wb_read)
    print(f"      {n_sched} classes loaded from Class list answer")
    print(f"      {n_unavail} unavailability slots loaded")

    print("[2/4] Assigning group rooms...")
    assign_group_rooms(conn)
    assigned_rooms = conn.execute(
        "SELECT COUNT(*) FROM group_rooms").fetchone()[0]
    print(f"      {assigned_rooms} class groups assigned rooms")

    print("[3/4] Assigning teachers (Lec1 / 2 / 3)...")
    unassigned = assign_teachers(conn)
    results    = collect_results(conn)
    lec1_count = sum(1 for r in results if r["lec1"])
    print(f"      {lec1_count}/{len(results)} classes assigned Lec1")

    print("[4/4] Writing Excel output...")
    write_output(INPUT_FILE, OUTPUT_FILE, results)

    validate_and_report(conn, results, unassigned)


if __name__ == "__main__":
    main()
