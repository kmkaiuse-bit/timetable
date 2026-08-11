# PRD — HKIT Timetable Scheduler

**Status:** Active · **Engine version:** V6 (V6.4 + 2026-08 updates)
**Last updated:** 2026-08-11
**Owner:** HKIT scheduling (kmkaiuse-bit/timetable)
**Related docs:** `MASTER_PLAN.md` (constraint status of record), `docs/pending-updates-2026-08.md`

> This PRD was first created on 2026-08-11. No PRD file existed in the repo before this;
> constraint decisions had lived in `MASTER_PLAN.md` and the meeting docs under `docs/meetings/`.

---

## 1. Purpose

Automate HKIT's class timetable scheduling. Given an Excel workbook describing classes,
teachers, and rooms, the system produces a complete timetable — assigning each class a
Day, Time, Room and Lecturer — that satisfies all hard constraints and optimises soft
ones, then surfaces every class it could not place with a diagnosed reason.

## 2. Users & stakeholders

- **Scheduler (primary user):** uploads the Excel, generates the timetable, reviews
  issues/warnings, hand-edits, downloads the Excel output.
- **Jo:** owns the scheduling rules and data decisions (capacity, combine groups, Net
  teachers). Open questions are tracked as J1–J10 (§10).

## 3. Architecture

```
Excel (.xlsx)
  └─ api/timetable_scheduler.py   Python engine, SQLite in-memory
        ├─ Phase 0  pre-schedule cadet classes (Mon/Wed/Fri, fixed rooms)
        ├─ Phase A  schedule Core classes (cc_group empty)
        ├─ Phase B  schedule CC Combine classes (cc_group set)
        └─ Teacher assignment (5-pass)
  └─ api/index.py                 Flask API (Vercel serverless entry point)
  └─ api/v4.html                  Web UI (served for all routes except /v3)
```

**Deploy:** Vercel; `api/index.py` handles all routes. The `api/` layout must not change.

## 4. Inputs

Excel workbook (`data/input/…xlsx`) with sheets including **Class list**, teachers, rooms,
availability. Key Class-list columns: class code (e.g. `DAE101_TS1`), subject/lecturer
names, Loading (hrs), Student No, and optional **Venue / Time / Date** and **CC Group**.

**Auto-assign gate:** the engine auto-schedules only when the Class-list Day/Time answers
are all empty (`n_sched == 0`). If any rows are pre-filled, the existing schedule is used
as-is and Phases 0/A/B (and the day rules, and CC combine) are skipped. The production
file `Planning for Timetable.xlsx` is pre-filled; the clean/template files auto-assign.

## 5. Functional requirements

### 5.1 Day rules  *(updated 2026-08)*

- **DAE subjects are scheduled on Monday, Tuesday, Thursday only** — never Wednesday or
  Friday. Implemented as `_DAY_PRIORITY = [Monday, Tuesday, Thursday]`.
- **Cadet classes keep Monday / Wednesday / Friday** as an explicit exception, via
  `_CADET_DAY_PRIORITY` in Phase 0 (independent of `_DAY_PRIORITY`).

### 5.2 Hard constraints

| ID | Rule |
|----|------|
| H1 | No teacher double-booking in a slot |
| H2 | No room double-booking in a slot |
| H3 | Students attend at their enrolled campus (TS→TM toggle `_TS_TM_ENABLED`, pending J1) |
| H4 | Teacher weekly loading — **soft cap 6, warn only, never block** (see 5.5) |
| H5 | CC Combine members share subject + language |
| H6 | TKO classes start ≥ 10:00 |
| H7 | Teacher qualified for the subject |
| H8 | Teacher at one centre per day; cross-centre allowed if travel ≤ 90 min (warn), else hard block |
| H9 | Students at one centre per day (exceptions: CSW↔SSP, TW↔CSW, TW↔SSP) |

### 5.3 Soft constraints

S1a unavailability · S1b centre preference · S2 no 天地堂 gap (students) · S3 same-subject
day rotation · S4 CC prefers centre with most students.

### 5.4 CC Combine  *(Phase B)*

Classes sharing a **CC Group** label are combined into one scheduled session.
- L1: find a day where no member has a Core class at a different centre.
- L2: if L1 fails, try the next-best centre; else report the group unplaced.
- Combine groups are input data (CC Group column), not code. Jo's DAE106 groups captured
  2026-08 (§10, `scripts/add_dae106_combine.py`).

### 5.5 Teacher assignment (5-pass) & loading

Five-pass Lec1 search relaxing, in order: quota → centre pref (S1b warn) → H4 cap (H4
soft warn) → unavailability (S1 warn). **Loading may exceed 6**; when it does the engine
warns and the UI flags the teacher red — it never blocks or refuses the class. Backups
Lec2/Lec3 do not count toward loading. A single global threshold of 6 applies (no
per-teacher caps — no data source, not required).

## 6. UI requirements  *(updated 2026-08)*

- **Four views:** Day, Week·by-room, Week·by-teacher, Month/semester calendar.
- **Full-width layout:** main container widened to 1440px so the timetable uses the whole
  page width.
- **Subject code in every cell:** each class cell shows the subject code (e.g. `DAE101`)
  as a bold badge across Day / Week·room / Week·teacher; enlarged fonts and cells.
- **Keep all five weekdays:** every view always renders Mon–Fri, keeping Wed/Fri
  columns/tabs even when empty (cadet classes and manual entries can still use them).
- Teacher Loading panel colours any teacher above 6 sessions red (超限). Issues panel and
  warnings surface unplaced classes and soft-constraint relaxations.

## 7. Outputs

Excel workbook (`Timetable_V4_Output.xlsx`) with the timetable, an Issues sheet, semester
calendar sheets (15 weeks with HK holidays), and teacher loading. Web UI mirrors these.

## 8. Known limitations & capacity  *(measured 2026-08, auto-assign clean input)*

- Restricting DAE to 3 days concentrates demand and **worsens room shortage** (TS / FL /
  TM already tight): on clean input, scheduled 111 → 94, SLOT_ROOM issues 1 → 14.
- With Jo's combine data added, **16 classes remain unplaceable** (114 total, 98 placed):
  - **7 core** — DAE101_TS4 and all FL (粉嶺): DAE103_FL1/2/3, DAE106_FL1/2/3.
  - **9 combine** — 4 of 5 DAE106 combine groups fail (`_find_cc_day` L2 exhausted under
    3 days); only KT3+KT4 places. Combining does reduce core room pressure (SLOT_ROOM
    14 → 7) because combined classes share a room.
- Root cause of both is the no-Wed/Fri rule (§5.1) plus FL room capacity. This is a
  **capacity decision for Jo** (add FL rooms / accept unplaced / revise combine rules).
  Wed/Fri use for DAE is explicitly ruled out by the requirement.

## 9. Configurable toggles (`api/timetable_scheduler.py`)

`_DAY_PRIORITY` (Mon/Tue/Thu) · `_CADET_DAY_PRIORITY` (Mon/Wed/Fri) ·
`_TEACHER_WEEKLY_SESSION_CAP` (6) · `_TS_TM_ENABLED` (J1) · `_CC_MAX_TRAVEL_MIN` (J2) ·
`_H8_MAX_TRAVEL_MIN` (90) · `_ENG_NET_MIN_BLOCKS` (1, J9).

## 10. Open questions for Jo

Carried from `MASTER_PLAN.md`: J1 (TS→TM), J2 (CC travel threshold), J3 (cadet rooms),
J5 (cost), J6 (teacher list), J8 (room overflow FL/SW/TKO/ST/TM), J9 (Net block = 20 hrs),
J10 (KT1/WT2 Net coverage). J4 answered (Net teachers).

New (2026-08, combine course):
1. `DAE106_TM3` and `DAE106_KT4` were missing from the Class list; appended with Jo's
   student counts (8, 9) and cloned subject/loading — confirm details.
2. Student counts on Jo's combine sheet differ from the workbook (e.g. CS6 10 vs 33) —
   which snapshot is authoritative? (Existing counts not overwritten.)
3. 4 of 5 combine groups cannot be placed under the no-Wed/Fri rule — accept, or handle
   the affected DAE106 combine classes manually?

## 11. Version history

| Date | Version | What |
|------|---------|------|
| 2026-05 | V3–V5 | Pre-filled → V4 auto-assign → constraint refinements, CC two-phase |
| 2026-05-23 → 06-09 | V6–V6.3 | Soft H4/H8, H6/H9 relaxations, reason codes, Issues panel, J1/J8 assumptions |
| 2026-07-08 | V6.4 | English auto-assign fixes; true V4 on clean input (111/111) |
| 2026-08-11 | 2026-08 | DAE Mon/Tue/Thu (cadet Mon/Wed/Fri); H4 soft-cap verified; UI (wide container, subject code, keep Wed/Fri); DAE106 combine groups captured |
