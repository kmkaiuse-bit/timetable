---
title: V6 Actionable Constraint Improvements
type: feat
status: active
date: 2026-05-25
---

# V6 Actionable Constraint Improvements

## Overview

Five improvements that can be implemented now without waiting for Jo's pending answers (J1–J7). Each item is self-contained. Execution order matters: Item 1 (_TRAVEL_TIME matrix) is a prerequisite for Items 2 and 5.

---

## Item 1 — Add `_TRAVEL_TIME` Matrix (Prerequisite)

**Why:** No travel time data exists in code. `scheduling-algorithm.html` documents the matrix but it was never ported to Python. Required by H8 (Item 2) and CC distance filter (Item 5).

**What:** Add constant to `api/timetable_scheduler.py` after `_CENTRE_RANK` (~line 164):

```python
# Estimated MTR travel time in minutes between centres (symmetric)
_TRAVEL_TIME: dict[str, dict[str, int]] = {
    "SSP": {"SSP":  0, "CSW":  8, "WT": 12, "KT": 25, "TKO": 35, "ST": 30, "FL": 55, "SW": 20, "TW": 15, "TS": 20, "TM": 50},
    "CSW": {"SSP":  8, "CSW":  0, "WT": 10, "KT": 22, "TKO": 30, "ST": 28, "FL": 50, "SW": 18, "TW": 15, "TS": 22, "TM": 48},
    "WT":  {"SSP": 12, "CSW": 10, "WT":  0, "KT": 20, "TKO": 25, "ST": 25, "FL": 45, "SW": 15, "TW": 20, "TS": 25, "TM": 45},
    "KT":  {"SSP": 25, "CSW": 22, "WT": 20, "KT":  0, "TKO": 12, "ST": 35, "FL": 40, "SW": 30, "TW": 35, "TS": 35, "TM": 55},
    "TKO": {"SSP": 35, "CSW": 30, "WT": 25, "KT": 12, "TKO":  0, "ST": 45, "FL": 50, "SW": 40, "TW": 45, "TS": 40, "TM": 65},
    "ST":  {"SSP": 30, "CSW": 28, "WT": 25, "KT": 35, "TKO": 45, "ST":  0, "FL": 35, "SW": 20, "TW": 30, "TS": 30, "TM": 45},
    "FL":  {"SSP": 55, "CSW": 50, "WT": 45, "KT": 40, "TKO": 50, "ST": 35, "FL":  0, "SW": 30, "TW": 45, "TS": 45, "TM": 80},
    "SW":  {"SSP": 20, "CSW": 18, "WT": 15, "KT": 30, "TKO": 40, "ST": 20, "FL": 30, "SW":  0, "TW": 25, "TS": 25, "TM": 50},
    "TW":  {"SSP": 15, "CSW": 15, "WT": 20, "KT": 35, "TKO": 45, "ST": 30, "FL": 45, "SW": 25, "TW":  0, "TS": 10, "TM": 35},
    "TS":  {"SSP": 20, "CSW": 22, "WT": 25, "KT": 35, "TKO": 40, "ST": 30, "FL": 45, "SW": 25, "TW": 10, "TS":  0, "TM": 30},
    "TM":  {"SSP": 50, "CSW": 48, "WT": 45, "KT": 55, "TKO": 65, "ST": 45, "FL": 80, "SW": 50, "TW": 35, "TS": 30, "TM":  0},
}
```

**Add helper function:**

```python
def _travel_mins(centre_a: str, centre_b: str) -> int:
    """Return travel time in minutes between two centres. Returns 0 for same centre."""
    if centre_a == centre_b:
        return 0
    return _TRAVEL_TIME.get(centre_a, {}).get(centre_b, 999)
```

**Acceptance Criteria:**
- [ ] `_TRAVEL_TIME` dict added after `_CENTRE_RANK` in `api/timetable_scheduler.py`
- [ ] `_travel_mins("FL", "TM")` returns 80
- [ ] `_travel_mins("SSP", "SSP")` returns 0
- [ ] `_travel_mins("TW", "TS")` returns 10

---

## Item 2 — H8: Change from Hard Block to Travel-Time Soft Gate

**Why:** Jo confirmed H8 becomes Soft. Teachers can work two centres same day if travel time ≤ 90 min. Only >90 min is hard blocked. SSP↔CSW (8 min) should always be allowed.

**Constant to add** (after `_TEACHER_WEEKLY_SESSION_CAP`):

```python
_H8_MAX_TRAVEL_MIN = 90  # teachers allowed to cross centres if travel ≤ this
```

**Change in `_find_teacher()`** (~line 1005–1021):

Replace current `h8_filter` SQL block (which hard-blocks ALL cross-centre) with Python post-filter:

```python
# H8: check teacher's already-committed centre on this day
def _get_teacher_day_centre(conn, teacher_id: int, day: str) -> Optional[str]:
    row = conn.execute("""
        SELECT r.centre FROM schedule s
        JOIN rooms r ON r.code = s.room_code
        WHERE s.teacher1_id = ? AND s.day = ?
        LIMIT 1
    """, (teacher_id, day)).fetchone()
    return row[0] if row else None
```

In `_find_teacher()`, after getting candidate list, add:

```python
# H8 travel-time gate (never relaxed)
if current_centre:
    filtered = []
    for t in candidates:
        committed = _get_teacher_day_centre(conn, t["id"], day)
        if committed is None or committed == current_centre:
            filtered.append(t)
        elif _travel_mins(committed, current_centre) <= _H8_MAX_TRAVEL_MIN:
            filtered.append(t)  # allowed — will warn at assignment
        # else: >90 min → hard excluded
    candidates = filtered
```

**Add warning** when assigning a teacher who crosses centres:

```python
# In assign_teachers(), when a teacher is assigned:
committed = _get_teacher_day_centre(conn, teacher_id, day)
if committed and committed != current_centre:
    warnings.append({
        "class": class_code,
        "teacher": teacher_name,
        "day": day,
        "reason": f"H8 (soft): cross-centre {committed}↔{current_centre} "
                  f"({_travel_mins(committed, current_centre)} min travel)"
    })
```

**Bug fix — H8 for Lec2/Lec3** (~line 953–956):

Currently `current_centre` is not passed when finding backup teachers. Fix:

```python
lec2 = _find_teacher(conn, subj, day, starts,
                     exclude=[lec1], current_centre=current_centre,
                     use_quota=False, ignore_h4_cap=True) if lec1 else None
lec3 = _find_teacher(conn, subj, day, starts,
                     exclude=[lec1, lec2], current_centre=current_centre,
                     use_quota=False) if lec2 else None
```

**Acceptance Criteria:**
- [ ] `_H8_MAX_TRAVEL_MIN = 90` constant added
- [ ] `_get_teacher_day_centre()` helper added
- [ ] Old `h8_filter` SQL block replaced with Python travel-time filter in `_find_teacher()`
- [ ] SSP↔CSW (8 min) teacher assignment produces warning but is allowed
- [ ] FL↔TM (80 min) teacher assignment is allowed with warning (≤90)
- [ ] Hypothetical centre pair >90 min is hard blocked
- [ ] H8 now enforced for Lec2/Lec3 as well
- [ ] Warning format: `"H8 (soft): cross-centre SSP↔CSW (8 min travel)"`

---

## Item 3 — H3: Add TS→TM Toggle

**Why:** Jo said TS students can go to TM (special exception). Exact activation date pending (J1). Add the logic now with a toggle so it can be turned on without code changes later.

**Constant to add:**

```python
_TS_TM_ENABLED = False  # set True when Jo confirms TS→TM rule is active
```

**Change in `_group_to_room_centre()`** (~line 470):

Current logic maps group code → single centre. Add TS→TM fallback:

```python
def _allowed_centres_for_group(group: str) -> list[str]:
    """Return allowed room centres for a group (primary first)."""
    primary = _group_to_room_centre(group)
    if _TS_TM_ENABLED and primary == "TS":
        return ["TS", "TM"]
    return [primary]
```

**Change in `_pick_room()`** (~line 475–509):

Replace single-centre room lookup with multi-centre attempt:

```python
for centre in _allowed_centres_for_group(group):
    room = conn.execute("""
        SELECT code FROM rooms
        WHERE centre = ? AND capacity >= ?
        ORDER BY capacity ASC LIMIT 1
    """, (centre, min_capacity)).fetchone()
    if room:
        return room[0]
return None  # no room found at any allowed centre
```

**Acceptance Criteria:**
- [ ] `_TS_TM_ENABLED = False` constant added
- [ ] `_allowed_centres_for_group()` helper added
- [ ] `_pick_room()` iterates allowed centres list
- [ ] With `_TS_TM_ENABLED = False`: TS groups only get TS rooms (no change from current)
- [ ] With `_TS_TM_ENABLED = True`: TS groups try TS first, fall back to TM if TS room is full

---

## Item 4 — S3: Phase 0 Cadet Class Framework

**Why:** `is_police_cadet` flag exists in schema but is never used. Cadet classes (subject DAE256, single-letter group codes like _E, _F) need to be pre-scheduled on Mon/Wed/Fri in their fixed rooms before Phase A runs. Otherwise Phase A occupies those rooms and cadet classes get displaced.

**Constants to add:**

```python
_CADET_DAYS = {"Monday", "Wednesday", "Friday"}
_CADET_SUBJECT_PREFIX = "DAE256"  # pending J3 confirmation — may need expanding
```

**Detection helper:**

```python
def _is_cadet_class(class_code: str) -> bool:
    """Identify cadet classes by single-letter group code (e.g. DAE256_E)."""
    parts = class_code.split("_")
    if len(parts) < 2:
        return False
    group = parts[-1]
    return len(group) == 1 and group.isalpha()
```

**Phase 0 function:**

```python
def phase0_schedule_cadets(conn: sqlite3.Connection) -> tuple[int, list[str]]:
    """
    Phase 0: Pre-schedule cadet classes (single-letter group code) on Mon/Wed/Fri.
    Must run before auto_assign_schedule() so their rooms are blocked first.
    Returns (scheduled_count, error_list).
    """
    cadet_classes = [
        row for row in conn.execute(
            "SELECT code, group_code, subject_code, student_count FROM classes"
        ).fetchall()
        if _is_cadet_class(row["code"])
    ]

    if not cadet_classes:
        return 0, []

    scheduled = 0
    errors = []

    for cls in cadet_classes:
        centre = conn.execute(
            "SELECT centre FROM class_groups WHERE code = ?",
            (cls["group_code"],)
        ).fetchone()
        if not centre:
            errors.append(f"{cls['code']}: no centre found for group {cls['group_code']}")
            continue

        room = conn.execute("""
            SELECT code FROM rooms
            WHERE centre = ? AND capacity >= ?
            ORDER BY capacity ASC LIMIT 1
        """, (centre[0], cls["student_count"] or 0)).fetchone()
        if not room:
            errors.append(f"{cls['code']}: no room at {centre[0]} with capacity >= {cls['student_count']}")
            continue

        # Try each cadet day until a free slot is found
        assigned = False
        subj = cls["subject_code"]
        loading = conn.execute(
            "SELECT loading_hrs FROM subjects WHERE code = ?", (subj,)
        ).fetchone()
        loading_hrs = (loading[0] if loading else 4) or 4
        need_two = loading_hrs >= 4
        time_blocks = _TKO_TIME_BLOCKS if centre[0] == "TKO" else _TIME_BLOCKS
        single_slots = _TKO_SINGLE_SLOTS if centre[0] == "TKO" else _SINGLE_SLOTS

        for day in sorted(_CADET_DAYS, key=lambda d: _DAY_PRIORITY.index(d)):
            if need_two:
                for s1, s2 in time_blocks:
                    if _room_free(conn, room[0], day, s1) and _room_free(conn, room[0], day, s2):
                        conn.execute(
                            "INSERT INTO schedule VALUES (?,?,?,?,NULL,NULL,NULL)",
                            (cls["code"], cls["group_code"], room[0], day, s1)
                        )
                        conn.execute(
                            "INSERT INTO schedule VALUES (?,?,?,?,NULL,NULL,NULL)",
                            (cls["code"] + "_2", cls["group_code"], room[0], day, s2)
                        )
                        scheduled += 1
                        assigned = True
                        break
            else:
                for s in single_slots:
                    if _room_free(conn, room[0], day, s):
                        conn.execute(
                            "INSERT INTO schedule VALUES (?,?,?,?,NULL,NULL,NULL)",
                            (cls["code"], cls["group_code"], room[0], day, s)
                        )
                        scheduled += 1
                        assigned = True
                        break
            if assigned:
                break

        if not assigned:
            errors.append(f"{cls['code']}: could not find free slot on Mon/Wed/Fri at {centre[0]}")

    return scheduled, errors
```

**Wire into `run_v4_from_bytes()`** (~line 1317):

```python
# Phase 0: pre-schedule cadet classes
cadet_scheduled, cadet_errors = phase0_schedule_cadets(conn)
for e in cadet_errors:
    stats["warnings"].append({"class": e, "teacher": "", "day": "", "reason": "Cadet Phase0 ERROR"})

# Phase A: Core classes
auto_assign_schedule(conn, cc_only=False)
# Phase B: CC Combine
cc_assigned, cc_errors = cc_assign_schedule(conn)
```

**Acceptance Criteria:**
- [ ] `_is_cadet_class("DAE256_E")` returns True
- [ ] `_is_cadet_class("DAE256_CS1")` returns False
- [ ] `_is_cadet_class("DAE101_TW2")` returns False
- [ ] `phase0_schedule_cadets()` added and wired as first step in `run_v4_from_bytes()`
- [ ] With current Excel data (no single-letter groups): Phase 0 runs, finds 0 cadet classes, proceeds normally
- [ ] When DAE256_E data is provided: class is pre-scheduled on first available Mon/Wed/Fri slot

---

## Item 5 — CC Combine: Add Configurable Distance Filter

**Why:** Jo mentioned CC Combine has "距離限制" but exact threshold pending (J2). Add infrastructure now with `None` default (disabled), ready to activate when Jo confirms the number.

**Constant to add:**

```python
_CC_MAX_TRAVEL_MIN: Optional[int] = None  # None = no limit; set to Jo's confirmed value (J2)
```

**Change in `cc_assign_schedule()`** (~line 801):

After getting `centre_order` from `_select_cc_centre()`, filter by max travel time if enabled:

```python
if _CC_MAX_TRAVEL_MIN is not None:
    # Get home centres of all CC group classes
    home_centres = set()
    for code in codes:
        row = conn.execute("""
            SELECT cg.centre FROM classes c
            JOIN class_groups cg ON cg.code = c.group_code
            WHERE c.code = ?
        """, (code,)).fetchone()
        if row:
            home_centres.add(row[0])

    centre_order = [
        c for c in centre_order
        if all(_travel_mins(home, c) <= _CC_MAX_TRAVEL_MIN for home in home_centres)
    ]
```

**Acceptance Criteria:**
- [ ] `_CC_MAX_TRAVEL_MIN = None` constant added
- [ ] With `None`: CC Combine behaves identically to current (no change)
- [ ] With `_CC_MAX_TRAVEL_MIN = 40`: centres too far from any participating group's home are excluded from `centre_order`
- [ ] Setting `_CC_MAX_TRAVEL_MIN = 40` and running with FL + TM groups should exclude SSP/CSW as CC centre options

---

## Implementation Order

1. **Item 1** — `_TRAVEL_TIME` + `_travel_mins()` (no tests needed, just data)
2. **Item 2** — H8 travel-time soft gate + Lec2/Lec3 fix
3. **Item 3** — H3 TS→TM toggle
4. **Item 4** — S3 Phase 0 cadet framework
5. **Item 5** — CC distance filter

Each item can be verified individually. Items 2 and 5 depend on Item 1.

---

## Out of Scope (Pending Jo)

| Item | Blocked by |
|------|-----------|
| Activate TS→TM rule | J1 (when does it start?) |
| Set CC distance threshold | J2 (what is the limit?) |
| Full cadet class details | J3 (room codes, exact detection) |
| Teacher type logic | J4 |
| Cost optimization | J5 |
| Teacher auto-assignment | J6 (name list) |
| Full CC grouping | J7 (grouping list) |
