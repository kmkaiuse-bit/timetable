---
title: "feat: Scheduling Constraints V5 — H4/H6/H8/H9 + CC Combine + 天地堂"
type: feat
status: active
date: 2026-05-12
---

# feat: Scheduling Constraints V5

Implement the full agreed constraint set into `api/timetable_scheduler.py`, including teacher/student single-centre-per-day rules, TKO special time slots, weekly loading cap, no-gap-class rule, and the CC Combine two-phase algorithm.

---

## Overview

The current V4 auto-scheduler implements H1, H2, H3 (partial), H7, and S3. Five constraints agreed in the May 2026 meetings are not yet implemented: **H4, H6, H8, H9** (all Hard) and **S2** (Soft). CC Combine logic (H5, S4, L1, L2) is entirely absent and requires a new scheduling phase.

This plan delivers all of the above in six incremental phases, each independently testable.

---

## Constraints Being Implemented

### Hard Constraints (new)

| ID | Rule | Impact |
|----|------|--------|
| H4 | Teacher max **6 sessions/week** (1 session = 2 hrs) | Blocks over-assignment of popular teachers |
| H6 | TKO centre: only `10:00–14:00` or `15:00–19:00` | Fixes TKO air-con cost rule + latent centre-alias bug |
| H8 | Teacher cannot go to **2 different centres same day** | Hard block in teacher assignment |
| H9 | Student group cannot go to **2 different centres same day** | Hard block in slot selection |

### Soft Constraints (new/improved)

| ID | Rule | Impact |
|----|------|--------|
| S1 | Teacher unavailability: skip first, **warn if forced** | Upgrade from silent skip to warn-on-force |
| S2 | No **天地堂** (gap >2 hrs in same AM or PM half-day) | For both teachers and class groups |

### CC Combine (entirely new)

| ID | Rule |
|----|------|
| H5 | Combine only if same subject code **and** same language (中/英) |
| S4 | Combine location = centre with **most enrolled students** (tie-break: proximity to SSP/CSW) |
| L1 | Schedule Core classes first; find a day where all CC students have no Core at a different centre |
| L2 | Fallback: try next-best centre if L1 finds no valid day; ERROR if all fail |

---

## Technical Approach

### Architecture

Single-file approach is preserved — all changes go into `api/timetable_scheduler.py`. No new files.

New constants, helper functions, and schema changes are added incrementally. The pipeline order becomes:

```
run_v4_from_bytes()
  └─ build_db()                        # unchanged
  └─ auto_assign_schedule()            # modified: +H6, +H9, +S2, +sort combines last
  └─ assign_teachers()                 # modified: +H4, +H8, +S1-warn
  └─ assign_combine_classes()          # NEW: CC Combine L1/L2
  └─ collect_results()                 # unchanged
  └─ collect_stats()                   # modified: +violation counts
  └─ write_output_fast()               # unchanged
```

### Key Data Structures (new, in-memory only)

```python
# Tracks which centre each group is committed to per day (for H9)
group_day_centre: dict[tuple[str, str], str]  # (group_code, day) -> centre

# Tracks which centre each teacher is committed to per day (for H8)
teacher_day_centre: dict[tuple[int, str], str]  # (teacher_id, day) -> centre

# Tracks total sessions assigned to each teacher this week (for H4)
teacher_weekly_sessions: dict[int, int]  # teacher_id -> session count

# Combine group registry (for H5/S4/L1/L2)
combine_groups: dict[str, list[str]]  # combine_id -> [class_code, ...]
```

---

## Implementation Phases

### Phase 1 — Fix TKO + Implement H6 ✦ Low Risk

**Files:** `api/timetable_scheduler.py`

**Changes:**

1. Add `'TK': 'TKO'` to `_GROUP_CENTRE_ALIAS` (line ~100).
   - Fixes the latent bug where `_pick_room()` looks for centre `TK` but all TKO rooms are stored as centre `TKO`.

2. Add TKO-specific time blocks constant:
   ```python
   _TKO_TIME_BLOCKS = [
       ("1000 - 1200", "1200 - 1400"),  # AM block
       ("1500 - 1700", "1700 - 1900"),  # PM block
   ]
   _TKO_SINGLE_SLOTS = ["1000 - 1200", "1200 - 1400", "1500 - 1700", "1700 - 1900"]
   ```

3. In `auto_assign_schedule()` (line ~480): detect if the class's centre is `TKO` and substitute `_TKO_TIME_BLOCKS` / `_TKO_SINGLE_SLOTS` for the standard ones.
   ```python
   centre = _group_to_room_centre(group_code)
   time_blocks = _TKO_TIME_BLOCKS if centre == "TKO" else _TIME_BLOCKS
   single_slots = _TKO_SINGLE_SLOTS if centre == "TKO" else _SINGLE_SLOTS
   ```

4. Add `teacher_unavailability` rows for TKO teachers covering the 0900 slot, if not already marked, since TKO teachers must not be scheduled at 0900. (Or simply let H6 room filter handle it.)

**Acceptance Criteria:**
- [ ] TK1, TK2 classes are assigned to `TKO - TKO` room (not random fallback)
- [ ] All TKO classes use only `1000`, `1200`, `1500`, or `1700` start times
- [ ] No TKO class starts at `0900` or `1400`
- [ ] All other centres unaffected

---

### Phase 2 — Implement H4: Teacher Weekly Loading Cap

**Files:** `api/timetable_scheduler.py`

**Changes:**

1. Add helper `_teacher_session_count(conn, teacher_id) -> int`:
   ```python
   def _teacher_session_count(conn, teacher_id):
       row = conn.execute("""
           SELECT COUNT(*) FROM schedule
           WHERE teacher1_id = ?
       """, (teacher_id,)).fetchone()
       return row[0] if row else 0
   ```
   Each row in `schedule` counts as 1 session (time1) or 2 sessions if `time2` is also non-null (4-hour class = 2 × 2-hr sessions).

   Refined version:
   ```python
   def _teacher_session_count(conn, teacher_id):
       row = conn.execute("""
           SELECT
               SUM(CASE WHEN time2 IS NOT NULL THEN 2 ELSE 1 END)
           FROM schedule WHERE teacher1_id = ?
       """, (teacher_id,)).fetchone()
       return row[0] or 0
   ```

2. In `_find_teacher()` (line ~585): add parameter `weekly_cap=6` and filter out teachers at or above cap:
   ```python
   # Add to the exclusion logic:
   AND t.id NOT IN (
       SELECT teacher1_id FROM schedule
       WHERE teacher1_id IS NOT NULL
       GROUP BY teacher1_id
       HAVING SUM(CASE WHEN time2 IS NOT NULL THEN 2 ELSE 1 END) >= 6
   )
   ```

3. When `use_quota=False` fallback is triggered (forced assignment), allow cap relaxation but add to unassigned warnings list with reason `"H4: loading cap exceeded"`.

**Acceptance Criteria:**
- [ ] No teacher has more than 6 sessions assigned in the final output
- [ ] Stats output shows teachers at cap (for debugging)
- [ ] Forced-over-cap assignments appear in warning list

---

### Phase 3 — Implement H8 + H9: No Multi-Centre Per Day

**Files:** `api/timetable_scheduler.py`

This is the most structurally important phase.

#### H9 — Student (class group) single centre per day

Add check in `auto_assign_schedule()` after `_room_free()` passes but before committing:

```python
# After picking (day, slot, room):
room_centre = conn.execute(
    "SELECT centre FROM rooms WHERE code=?", (room_code,)
).fetchone()[0]

# Check H9: group not already at a different centre today
existing_centre = group_day_centre.get((group_code, day))
if existing_centre and existing_centre != room_centre:
    continue  # reject this day/room combination

# If accepted, record
group_day_centre[(group_code, day)] = room_centre
```

The `group_day_centre` dict is initialised once before the main loop and passed through.

#### H8 — Teacher single centre per day

Add check in `_find_teacher()`: pass the current class's `room_centre` and filter teachers who are already committed to a different centre today:

```python
# New parameter: current_centre: str
# Add to SQL WHERE clause:
AND t.id NOT IN (
    SELECT s.teacher1_id
    FROM schedule s
    JOIN rooms r ON r.code = s.room_code
    WHERE s.day = ?
      AND s.teacher1_id IS NOT NULL
      AND r.centre != ?
)
# params: (..., day, current_centre)
```

Update all callers of `_find_teacher()` to pass `current_centre`.

**Acceptance Criteria:**
- [ ] No teacher has classes at two different centres on the same day in the output
- [ ] No class group appears at two different centres on the same day in the output
- [ ] Unassignable classes (H8 forces no valid teacher) appear in ERROR list with reason

---

### Phase 4 — Implement S1 Improvement: Warn When Forced into Unavailable Slot

**Files:** `api/timetable_scheduler.py`

**Changes:**

1. Modify `_find_teacher()` to return `(teacher_id, forced: bool)`:
   - First pass: with unavailability filter (current behaviour). If found → `(id, False)`.
   - Second pass (only if first returns None): without unavailability filter. If found → `(id, True)`.
   - If both pass fail → `(None, False)`.

2. In `assign_teachers()`: if `forced=True`, add to a new `forced_warnings` list:
   ```python
   {"class": class_code, "teacher": teacher_name, "reason": "S1: assigned to unavailable slot"}
   ```

3. Include `forced_warnings` in `collect_stats()` output and display in UI.

**Acceptance Criteria:**
- [ ] Teachers assigned to their marked-unavailable slots appear in the Warning section of stats
- [ ] System still assigns them (does not block) — S1 is Soft
- [ ] UI shows count of S1 warnings

---

### Phase 5 — Implement S2: No 天地堂

**Files:** `api/timetable_scheduler.py`

**Definition:** Within the AM block (before 14:00) or PM block (14:00+), if two classes for the same group are more than 2 hours apart (with a gap in between), that is 天地堂. Lunch break does not count.

**Helper function:**

```python
_SLOT_ORDER = {
    "0900 - 1100": 0, "1000 - 1200": 0,  # TKO AM slot 1
    "1100 - 1300": 1, "1200 - 1400": 1,  # TKO AM slot 2
    "1400 - 1600": 2, "1500 - 1700": 2,  # TKO PM slot 1
    "1600 - 1800": 3, "1700 - 1900": 3,  # TKO PM slot 2
}
_AM_SLOTS = {0, 1}
_PM_SLOTS = {2, 3}

def _has_tiandi_violation(conn, group_code, day, new_slot):
    """Returns True if scheduling group_code at new_slot on day creates a gap >2hr in same half-day."""
    existing = conn.execute(
        "SELECT time1 FROM schedule WHERE group_code=? AND day=?",
        (group_code, day)
    ).fetchall()
    # Need to join classes table to get group_code... add group_code to schedule or derive from class_code
```

**Schema note:** `schedule` does not currently store `group_code` directly. It stores `class_code` (e.g. `DAE101_CS1`). Add a derived lookup via `classes` table.

Implementation:
1. Add `group_code` column to `schedule` table in `_new_db()` schema.
2. Populate it when inserting into `schedule` in `auto_assign_schedule()`.
3. Implement `_has_tiandi_violation(conn, group_code, day, new_slot)` using `_SLOT_ORDER` to check for non-adjacent slots in the same half-day.
4. Call before committing a (day, slot) assignment. If violation detected, try next slot.
5. If all slots on all days create 天地堂, assign anyway and add to `tiandi_warnings`.

**Acceptance Criteria:**
- [ ] No class group has a gap >2 hours within AM block or PM block in the final output
- [ ] 天地堂 warnings appear in stats if unavoidable
- [ ] Applies to both standard and TKO time slots

---

### Phase 6 — CC Combine (H5, S4, L1, L2) ✦ Most Complex

**Files:** `api/timetable_scheduler.py`, Excel template

This requires a new input mechanism (to declare which classes are CC Combine) and a new scheduling phase.

#### 6a. Data Model Changes

Add to `_SCHEMA`:

```sql
CREATE TABLE IF NOT EXISTS combine_classes (
    combine_id   TEXT NOT NULL,   -- e.g. "CC_MKT101_EN"
    class_code   TEXT NOT NULL,   -- e.g. "MKT101_CS1"
    language     TEXT NOT NULL,   -- "CN" or "EN"
    is_combine   INTEGER DEFAULT 1,
    PRIMARY KEY (combine_id, class_code)
);
```

Add `is_combine INTEGER DEFAULT 0` column to `classes` table.

#### 6b. New Excel Input

Add a new sheet `CC Combine Groups` with columns:
- A: `Combine ID` (e.g. `CC_MKT101_EN`)
- B: `Class Code` (e.g. `MKT101_CS1`)
- C: `Language` (`CN` or `EN`)

Add `_load_combine_groups(conn, wb)` to `build_db()`.

H5 validation: if two classes share same subject code but different language, they cannot be in the same combine_id — raise a data error.

#### 6c. New Scheduling Phase: `assign_combine_classes(conn)`

```
For each combine_id:
  1. Get all class_codes in this combine group
  2. Count students per centre across all classes → pick centre with most (S4)
     Tie-break: highest distance rank (SSP/CSW=9 > WT=8 > ... > FL=1)
  3. METHOD A (L1):
     For each candidate day (Mon→Tue→Thu→Wed→Fri):
       For each class in combine group:
         Get group_code → look up group_day_centre[(group_code, day)]
         If committed centre != combine_centre AND committed centre exists → day invalid
       If day is valid for ALL groups:
         Pick time slot at combine_centre, check room free + teacher available
         → Schedule all combine classes together at (day, slot, combine_centre_room)
         → Update group_day_centre for all involved groups
         → DONE, break
  4. METHOD B (L2) — if Method A failed:
     For each centre in distance rank order (descending):
       Repeat Method A search with this centre as combine_centre
       → If found: schedule + add "S4 fallback" warning
     → If all fail: mark as ERROR, add to unassigned list with reason
```

#### 6d. `assign_teachers()` integration

CC Combine classes need a teacher too. Since all combine classes share one room/day/time, assign one teacher to all of them (or one per class if teacher capacity allows). Apply H8 (teacher single centre per day) as normal.

#### 6e. Modify `auto_assign_schedule()`

Sort classes so `is_combine = 0` classes are processed first, `is_combine = 1` classes are skipped (handled by `assign_combine_classes()` instead).

**Acceptance Criteria:**
- [ ] CC Combine classes are scheduled at the centre with the most students (S4)
- [ ] No student group goes to two centres on the same day due to CC Combine (H9 preserved)
- [ ] Method A is attempted first; Method B fallback is logged
- [ ] All-fail cases appear as ERROR in UI
- [ ] Combine classes with different languages are never merged (H5)

---

## Centre Distance Ranking (for S4 tie-break)

```python
_CENTRE_RANK = {
    "SSP": 9, "CSW": 9,  # core
    "WT":  8,
    "TS":  7, "TW": 7,
    "ST":  6,
    "SW":  5,
    "KT":  4,
    "TKO": 3,
    "TM":  2,
    "FL":  1,
}
```

---

## Warning System Updates

Add to `collect_stats()` and UI display:

| Warning Type | Trigger | Severity |
|---|---|---|
| `H4_OVERLOAD` | Teacher assigned >6 sessions | 🔴 ERROR |
| `H8_MULTI_CENTRE` | Teacher at 2 centres same day (should never fire if H8 implemented correctly — safety net) | 🔴 ERROR |
| `H9_MULTI_CENTRE` | Group at 2 centres same day (safety net) | 🔴 ERROR |
| `S1_FORCED` | Teacher assigned to their unavailable slot | 🟡 WARNING |
| `S2_TIANDI` | Group or teacher has gap >2hr in same half-day | 🟡 WARNING |
| `S4_FALLBACK` | CC Combine used Method B location, not the top-student centre | 🟡 WARNING |
| `CC_NO_SOLUTION` | CC Combine has no valid day/centre | 🔴 ERROR |

---

## Acceptance Criteria (Full Feature)

### Hard Constraints
- [ ] H4: No teacher has >6 sessions in any output
- [ ] H6: All TKO classes use only 1000/1200/1500/1700 start times
- [ ] H8: No teacher appears at 2 different centres on the same day
- [ ] H9: No class group appears at 2 different centres on the same day

### Soft Constraints
- [ ] S1: Teachers forced into unavailable slots appear in Warning section
- [ ] S2: No group or teacher has a >2hr gap within the same AM or PM half-day

### CC Combine
- [ ] H5: CC Combine only merges same subject + same language
- [ ] S4: Combine location defaults to centre with most students
- [ ] L1: Core classes scheduled first; combine placed on a conflict-free day
- [ ] L2: Fallback centre tried if L1 fails; ERROR if all fail

### Regression
- [ ] All previously passing classes (0–7 unassigned) remain assigned
- [ ] TKO classes now correctly use TKO room (not random fallback)
- [ ] Output Excel format unchanged

---

## Dependencies & Risks

| Risk | Impact | Mitigation |
|---|---|---|
| TKO time slots not in `teacher_unavailability` | TKO teachers may be blocked | Add `1000`/`1500` slots to Teacher Availability sheet; or derive from centre |
| CC Combine requires new Excel sheet | If sheet absent, phase 6 is skipped gracefully | `_load_combine_groups()` returns early if sheet not found |
| H9 may increase unassigned count | Groups with many subjects on same day may conflict | Accept and report; human review via ERROR list |
| H8 + popular teachers (e.g. Mercury Lee) | With centre restriction, may not find valid teachers | H4 cap reduces over-assignment; ERROR list flags it |
| Adding `group_code` to `schedule` schema | May break existing output queries | Test `collect_results()` JOIN after schema change |

---

## Implementation Order

```
Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5 → Phase 6
  H6/fix    H4        H8+H9     S1        S2        CC Combine
  (1 day)  (half)    (1 day)   (half)    (1 day)   (2–3 days)
```

Each phase is independently testable against the existing Excel template. Regression test: upload current `Planning for Timetable V4 Template.xlsx` and verify unassigned count does not increase after each phase.

---

## Files to Modify

| File | Changes |
|---|---|
| `api/timetable_scheduler.py` | All constraint logic — all phases |
| `api/v4.html` | UI: add Warning section display for S1/S2/S4/CC errors |
| `Planning for Timetable V4 Template.xlsx` | Phase 6: add `CC Combine Groups` sheet; Phase 1: verify TKO rows |

---

## Sources & References

### Internal
- Current scheduler: `api/timetable_scheduler.py:1–906`
- Constraint spec: `scheduling-algorithm.html`
- Meeting doc: `scheduling-meeting.docx` / `generate_meeting_doc.py`
- Prior auto-scheduler plan: `docs/plans/2026-05-05-004-feat-full-auto-scheduler-v4-plan.md`
- DB design plan: `docs/plans/2026-05-05-003-feat-sqlite-database-timetable-system-plan.md`

### Key Constants (current locations in timetable_scheduler.py)
- `_GROUP_CENTRE_ALIAS`: line ~100 (add `'TK': 'TKO'`)
- `_TIME_BLOCKS`, `_SINGLE_SLOTS`: lines ~108–113 (add TKO variants)
- `auto_assign_schedule()`: lines 443–530 (H6, H9, S2 checks)
- `_find_teacher()`: lines 585–628 (H8, H4 SQL filter)
- `assign_teachers()`: lines 535–582 (H4 counter, S1 warn)
- `_new_db()`: lines 117–121 (schema: add `group_code` to schedule, new combine tables)
