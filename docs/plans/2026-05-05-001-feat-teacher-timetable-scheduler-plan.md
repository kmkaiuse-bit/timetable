---
title: Semi-Automated Teacher Timetable Scheduler
type: feat
status: completed
date: 2026-05-05
---

# Semi-Automated Teacher Timetable Scheduler

## Overview

A Python script that reads the existing `Planning for Timetable.xlsx`, automatically assigns teachers and schedules all ~108 classes (day, time slots, room), then writes a new Excel file in **exactly** the same format as the original — ready for manual fine-tuning in Excel.

Each new Term, the user updates the input Excel data and re-runs one command:

```bash
python timetable_scheduler.py
```

Output: `Timetable_Output.xlsx` — a filled copy of the input file with `Class list answer` and all 5 daily timetable sheets populated.

---

## Problem Statement

Scheduling ~108 classes across 37 teachers, 30 rooms, and 5 days is done entirely by hand today. The main pain points are:

- **Room overflow**: manually checking that classroom capacity ≥ student count across all sessions
- **Loading imbalance**: ensuring each teacher's class count stays within their target loading
- **Repetition**: every Term requires redoing the same calculations from scratch

The doc file (`Output of weekly timetable.docx`) documents the rules but there is no system enforcing them.

---

## Proposed Solution

A two-phase Python script:

1. **Teacher Assignment** — read `Teacher load table with subject`, match each class to eligible teachers respecting their per-subject loading quota
2. **Greedy Scheduler** — for each class, find the best valid `(day, time_slot_1, time_slot_2, room)` combination using priority-ordered constraint checking

Output is written back into the existing Excel template using `openpyxl` (open template → fill cells → save as new file), preserving all existing formatting and COUNTIF formulas.

OR-Tools CP-SAT is noted as an upgrade path if solution quality proves insufficient (see Risks).

---

## Technical Approach

### File Structure

```
09_timetable/
├── Planning for Timetable.xlsx        # input (unchanged)
├── Output of weekly timetable.docx    # reference rules doc (unchanged)
├── timetable_scheduler.py             # NEW — main script
├── requirements.txt                   # NEW — ortools, openpyxl
└── Timetable_Output.xlsx              # NEW — generated each run
```

---

### Data Models

```python
# timetable_scheduler.py

@dataclass
class ClassEntry:
    code: str           # "DAE101_CS1"
    subject_code: str   # "DAE101"
    subject_name: str   # "Chinese Language"
    subject_cn: str     # "中國語文"
    loading: int        # 2 or 4 (hours)
    students: int
    centre: str         # "CS" extracted from code suffix

@dataclass
class TeacherEntry:
    name: str
    subject_quotas: dict  # {"Chinese Language": 4, ...}  max classes to assign

@dataclass
class RoomEntry:
    code: str           # "CSW - C1"
    capacity: int       # use "No of Seats" column
    centre: str         # "CSW" prefix

@dataclass
class ScheduledClass:
    entry: ClassEntry
    day: str            # "Monday"
    slot1: str          # "0900 - 1100"
    slot2: str | None   # "1100 - 1300" or None for 2hr Maths
    room: str
    lecturer1: str
    lecturer2: str | None
    lecturer3: str | None
```

---

### Constraint Rules

**Hard constraints (must not be violated):**

| Rule | Description |
|------|-------------|
| H1 | A teacher cannot be assigned to two classes at the same time on the same day |
| H2 | A room cannot host two classes at the same time on the same day |
| H3 | Room capacity ≥ student count for the assigned class |

**Soft constraints (optimise, not enforce — scored and prioritised):**

| Priority | Rule | Description |
|----------|------|-------------|
| S1 | Loading balance | Each teacher's assigned classes ≤ their quota from Teacher load table |
| S2 | Preferred days | Main subjects (Chinese/English/Maths/CLP/MathsPlus) prefer Mon, Tue, Thu |
| S3 | Centre grouping | Classes from the same centre prefix (e.g., all WT_*) prefer the same day |

**Special rule — CS1/CS2/CS3 (Police Cadet Program):**
These three classes only take Chinese, English, Maths, CLP (no MathsPlus). They may be scheduled on any day including Wed/Fri.

---

### Time Slot Model

```python
STANDARD_SLOTS = [
    "0900 - 1100",
    "1100 - 1300",
    "1400 - 1600",
    "1600 - 1800",
]

# Valid two-slot combinations for 4hr subjects (loading=4)
# Ordered by preference: consecutive morning > consecutive afternoon > split
SLOT_PAIRS_4HR = [
    ("0900 - 1100", "1100 - 1300"),   # morning block (preferred)
    ("1400 - 1600", "1600 - 1800"),   # afternoon block
    ("0900 - 1100", "1400 - 1600"),   # morning + early afternoon (split)
    ("1100 - 1300", "1400 - 1600"),   # late morning + afternoon (split)
    ("0900 - 1100", "1600 - 1800"),   # morning + late afternoon (split)
    ("1100 - 1300", "1600 - 1800"),   # late morning + late afternoon (split)
]

# Single-slot for 2hr subjects (loading=2)
SLOTS_2HR = ["0900 - 1100", "1100 - 1300", "1400 - 1600", "1600 - 1800"]

PREFERRED_DAYS = ["Monday", "Tuesday", "Thursday"]
ALL_DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
```

---

### Scheduling Algorithm — Greedy with Priority Ordering

```
FUNCTION assign_teachers(classes, teachers):
    # Group classes by subject
    # For each subject, sort teachers by quota descending
    # Assign classes to teachers round-robin until quota reached
    # Return: classes with lecturer1 assigned

FUNCTION schedule(classes_with_teachers, rooms):
    # Sort classes by: students DESC, loading DESC (4hr before 2hr)
    # Build state: teacher_busy[teacher][day][slot] = bool
    #              room_busy[room][day][slot] = bool

    scheduled = []
    unscheduled = []

    FOR class IN sorted_classes:
        assigned = False
        FOR day IN preferred_days_for(class):         # soft S2
            FOR slot_combo IN valid_slots(class.loading):
                FOR room IN rooms_sorted_by_fit(class.students):  # smallest room that fits
                    IF room.capacity >= class.students:
                        AND NOT room_busy[room][day][slot_combo]:
                        AND NOT teacher_busy[class.lecturer1][day][slot_combo]:
                            assign(class, day, slot_combo, room)
                            assigned = True
                            BREAK
            IF assigned: BREAK
        IF NOT assigned:
            unscheduled.append(class)  # report for manual resolution

    RETURN scheduled, unscheduled
```

If `unscheduled` is non-empty, the script prints a warning listing unresolved classes for manual attention.

---

### Excel Output Strategy

The output file **is** the input file with cells filled in — not a file built from scratch.

```python
def write_output(input_path, output_path, schedule):
    wb = openpyxl.load_workbook(input_path)  # loads all formatting, formulas, styles

    # 1. Fill "Class list answer" sheet
    ws = wb["Class list answer"]
    # Clear existing data rows (keep headers)
    # Write one row per scheduled class: code, name, lecturer1/2/3, day, time1, time2, venue

    # 2. Fill 5 daily timetable sheets
    for day in ALL_DAYS:
        sheet_name = DAY_TO_SHEET[day]  # e.g., "Mon(Term1) "
        ws = wb[sheet_name]
        # For each room row in the sheet:
        #   find the scheduled class for that room on that day
        #   fill Code, Class Size, Subject, Lecturer cells for each time slot
        # Leave COUNTIF formula rows untouched (they reference the filled cells above)

    wb.save(output_path)
```

The COUNTIF formulas in the bottom section of each daily sheet auto-update when Excel opens the file, because they reference the cells filled above — no formula re-writing needed.

---

### Sheet-to-Day Mapping

```python
DAY_TO_SHEET = {
    "Monday":    "Mon(Term1) ",   # note trailing space in original
    "Tuesday":   "Tue(Term1)",
    "Wednesday": "Wed(Term1)",
    "Thursday":  "Thu(Term1)",
    "Friday":    "Fri(Term1)",
}
```

---

### Implementation Phases

#### Phase 1 — Data Reader (`read_input.py` or module in `timetable_scheduler.py`)

- Read `Class list` sheet → list of `ClassEntry`
- Read `Teacher load table with subject` → dict of `TeacherEntry`
  - Handle the two-section structure (numbered + unnumbered rows)
  - Combine quotas: sum both sections for same teacher name
- Read `Centre Room Allocation` → dict of `RoomEntry`
- Read `Class list answer` (existing answer) for validation/comparison only
- **Deliverable:** All inputs parsed without errors, printed as summary

**Files:** `timetable_scheduler.py` (functions: `read_classes()`, `read_teachers()`, `read_rooms()`)

---

#### Phase 2 — Teacher Assigner

- For each subject, collect eligible teachers (those with quota > 0 for that subject)
- Sort by quota descending (higher-quota teachers get assigned first)
- Assign Lecturer1 to each class round-robin within quota
- Flag any classes where no eligible teacher is available
- **Deliverable:** All 108 classes have a Lecturer1 assigned; assignment summary printed

**Files:** `timetable_scheduler.py` (function: `assign_teachers(classes, teachers)`)

---

#### Phase 3 — Greedy Scheduler

- Implement greedy algorithm described above
- Track `teacher_busy` and `room_busy` state dictionaries
- Apply soft constraint ordering (preferred days → preferred rooms)
- Collect and print any unscheduled classes at end
- **Deliverable:** All (or nearly all) classes have `(day, slot1, slot2, room)` assigned

**Files:** `timetable_scheduler.py` (function: `schedule_classes(classes, rooms)`)

---

#### Phase 4 — Excel Writer

- Open `Planning for Timetable.xlsx` as template
- Fill `Class list answer` sheet rows
- Fill all 5 daily timetable sheets (Code, Class Size, Subject, Lecturer columns per time group)
- Save as `Timetable_Output.xlsx`
- **Deliverable:** Output file opens in Excel with all sheets filled, COUNTIF formulas computing correctly

**Files:** `timetable_scheduler.py` (function: `write_output(wb, schedule, output_path)`)

---

#### Phase 5 — Validation & Reporting

- Script prints a summary after running:
  - Total classes scheduled vs unscheduled
  - Any teacher with >1 class at same time (should be 0)
  - Any room with >1 class at same time (should be 0)
  - Any room where capacity < students (should be 0)
  - Teachers whose assigned count differs from quota by more than 1
- **Deliverable:** Clean console output with GO/NO-GO on hard constraints

**Files:** `timetable_scheduler.py` (function: `validate_and_report(schedule)`)

---

## Acceptance Criteria

- [ ] Script runs with `python timetable_scheduler.py` and completes without error
- [ ] `Timetable_Output.xlsx` is created in the same directory
- [ ] `Class list answer` sheet is fully populated (all 108 rows have Day, Time, Venue, Lecturer1)
- [ ] All 5 daily timetable sheets (`Mon`–`Fri`) are populated with correct class/room/teacher data
- [ ] Zero hard constraint violations: no teacher double-booked, no room double-booked, no room overflow
- [ ] COUNTIF formula rows in daily sheets compute correct teacher counts (verified by opening in Excel)
- [ ] Main subjects (Chinese/English/Maths/CLP) are scheduled on Mon/Tue/Thu in ≥80% of cases
- [ ] Script outputs a validation report to console listing any unresolved classes
- [ ] `requirements.txt` exists with `openpyxl` and `ortools`
- [ ] A non-technical user can run the script by following a 3-step README

---

## Dependencies & Risks

| Item | Detail |
|------|--------|
| **openpyxl** | Already installed. Template-modification approach preserves all formatting. |
| **ortools** | `pip install ortools` — single wheel on Windows 11, no compiler needed. Only needed if greedy produces poor results (upgrade path). |
| **Teacher load table data quality** | The table has duplicate rows for some teachers and one obvious data error (Mr. Mercury Lee CLP=60). Phase 1 must handle these gracefully. |
| **Non-standard time slots** | Existing `Class list answer` uses some non-standard start times (8:30, 10:30, 13:00, etc.). MVP uses standard slots only; user adjusts edge cases manually. |
| **Greedy may leave some classes unscheduled** | If the problem is too tightly constrained, greedy gets stuck. Mitigation: relax soft constraints on retry; upgrade to OR-Tools CP-SAT if needed. |
| **Excel sheet name trailing space** | `Mon(Term1) ` has a trailing space in the original file — must be matched exactly when accessing by name. |
| **Co-teaching (Lecturer2, Lecturer3)** | MVP assigns Lecturer1 only. Net teachers (Lecturer2/3) are left blank or copied from existing answer manually. This is a known limitation. |

---

## Sources & References

- Research: OR-Tools CP-SAT recommended over `python-constraint` for 100+ variable timetabling (single `pip install`, pre-compiled Windows wheels, CP-SAT nurse-scheduling pattern maps directly to this problem)
- openpyxl best practice: open template → fill → save (preserves all formatting/styles without re-specification)
- Input file: `Planning for Timetable.xlsx` — all 9 sheets analysed
- Rules doc: `Output of weekly timetable.docx` — scheduling constraints source
- OR-Tools docs: https://developers.google.com/optimization/reference/python/sat/python/cp_model
- openpyxl docs: https://openpyxl.readthedocs.io/en/stable/
