---
title: Full Coverage & Diagnostics — Timetable Scheduler V6.3
type: feat
status: active
date: 2026-06-09
origin: docs/brainstorms/2026-06-09-full-coverage-diagnostics-requirements.md
---

# Full Coverage & Diagnostics — Timetable Scheduler V6.3

## Overview

Improve the scheduler so it covers as many classes as possible, and for any class it cannot schedule, give Jo a precise, actionable explanation of why. Two interim assumptions replace pending Jo questions: J1 (TS→TM enabled) and J8 (cross-centre fallback for room overflow). All assumptions are recorded in the output so they can be reversed once Jo responds.

*Key decisions from origin document:*
- J1 assumption: `_TS_TM_ENABLED = True` (reverse with one constant)
- J8 assumption: cross-centre fallback via `_TRAVEL_TIME` matrix ≤ 90 min (see [origin](../brainstorms/2026-06-09-full-coverage-diagnostics-requirements.md))
- Diagnostics: UI Issues panel (open by default) + Excel "Issues" sheet
- Category taxonomy: `DATA_PROBLEM` / `ALGORITHM_ISSUE` / `ASSUMPTION`

---

## Problem Statement

After the B3 fix (removed H3 room fallback), 19+ classes fail to schedule and the failures are either invisible (not shown in UI) or described with a single generic string. Jo cannot act on "no free slot across all days" — it does not say whether to add a teacher, split a class, or relocate to another centre. Additionally:
- `stats["unscheduled_rooms"]` is never rendered in the UI (pre-existing gap)
- Teacher failures (`stats["unassigned"]`) have no reason codes
- J1 and J8 block scheduling unnecessarily until Jo's written approval arrives

---

## Category Mapping (R5 — resolved from SpecFlow Q3)

| Failure Type | Reason Code | Category | Suggested Action |
|---|---|---|---|
| Room capacity too small for class | `NO_ROOM_CAPACITY` | DATA_PROBLEM | Increase room capacity at [centre] or split class |
| Room exists but fully booked on all days | `SLOT_ROOM` | ALGORITHM_ISSUE | Scheduler conflict — retry with fewer H constraints |
| No qualified teacher for subject | `TEACHER_NONE` | DATA_PROBLEM | Add teacher with [subject] qualification to Teachers sheet |
| All teachers at weekly session cap | `TEACHER_CAPACITY` | DATA_PROBLEM | Reduce other assignments or add another [subject] teacher |
| All teachers unavailable on day | `TEACHER_AVAIL` | DATA_PROBLEM | Update teacher availability in Teachers sheet |
| H8 travel blocks all candidates | `TEACHER_H8` | ALGORITHM_ISSUE | No teacher can travel to [centre] from prior commitment |
| All days blocked by H9 student conflict | `SLOT_H9` | ALGORITHM_ISSUE | Student group has conflicting class on same day |
| TKO ≥10:00 constraint eliminates all slots | `SLOT_H6` | ALGORITHM_ISSUE | No valid TKO slot available (H6 constraint) |
| TS→TM assignment applied (J1 assumption) | `H3_TS_TM` | ASSUMPTION | Confirm J1 with Jo: TS→TM cross-centre is now enabled |
| Cross-centre room fallback applied (J8) | `H3_FALLBACK` | ASSUMPTION | Confirm J8 with Jo: class moved to [target_centre] |
| English Net teacher shortfall | `NET_SHORTFALL` | DATA_PROBLEM | Add more Net teachers to Net Teachers sheet |
| Cadet Phase 0 error | `CADET_ERROR` | ALGORITHM_ISSUE | Cadet class could not be placed on Mon/Wed/Fri |

---

## Technical Approach

### Architecture

Seven backend changes across two files, one new stats key, and one new Excel sheet. Frontend adds one panel following the existing loading-panel pattern.

**Critical architectural constraint (SpecFlow I6):** The DB is closed (`del conn`) before `write_output_fast()` is called. All subject/centre/student_count enrichment must happen *before* `del conn` in `run_v4_from_bytes()`.

**Backward-compatibility:** Keep `stats["unassigned"]` as `list[str]` for frontend compat. Add new `stats["issues"]` as the unified list of dicts. Do not remove `stats["unscheduled_rooms"]`.

### Implementation Phases

---

#### Phase 1: Foundation — `_TS_TM_ENABLED` and `_pick_room()` refactor

**File:** `api/timetable_scheduler.py`

**Task 1.1 — Enable J1 assumption (line 159)**
```python
# Before
_TS_TM_ENABLED = False

# After
_TS_TM_ENABLED = True   # J1 assumption: TS→TM enabled (confirm with Jo)
```

**Task 1.2 — Refactor `_pick_room()` return type (lines 550–579)**

Change signature: returns `(room_code: str | None, centre_used: str | None)`.
- Normal success: `(room_code, home_centre)`
- Cross-centre success: `(room_code, other_centre)` — triggers ASSUMPTION record
- Full failure: `(None, None)`

```python
def _pick_room(conn, group_code, students):
    """
    Returns (room_code, centre_used).
    centre_used differs from home centre when cross-centre fallback fires.
    Returns (None, None) if no room found anywhere within travel limit.
    """
    centre  = _group_to_room_centre(group_code)
    allowed = _allowed_centres_for_group(group_code)  # includes TM for TS when J1 enabled

    # Pass 1: exact room for group code
    row = conn.execute("""
        SELECT room_code, centre FROM rooms
        WHERE room_code = ? AND capacity >= ?
        LIMIT 1
    """, (f"{centre} - {centre}{_group_num(group_code)}", students)).fetchone()
    if row:
        return row[0], row[1]

    # Pass 2: largest room at allowed centres
    for c in allowed:
        row = conn.execute("""
            SELECT room_code FROM rooms WHERE centre=? AND capacity >= ?
            ORDER BY capacity ASC LIMIT 1
        """, (c, students)).fetchone()
        if row:
            return row[0], c

    # Pass 3: J8 assumption — try nearest centres within H8 travel limit
    nearby = sorted(
        [(c, _travel_mins(centre, c)) for c in _TRAVEL_TIME.get(centre, {})
         if c not in allowed and _travel_mins(centre, c) <= _H8_MAX_TRAVEL_MIN],
        key=lambda x: x[1]   # closest first
    )
    for fb_centre, _ in nearby:
        row = conn.execute("""
            SELECT room_code FROM rooms WHERE centre=? AND capacity >= ?
            ORDER BY capacity ASC LIMIT 1
        """, (fb_centre, students)).fetchone()
        if row:
            return row[0], fb_centre   # ASSUMPTION path

    return None, None   # truly no room anywhere
```

**Task 1.3 — Update callers of `_pick_room()`**

Three call sites to update:

*In `auto_assign_schedule()` (~line 706):*
```python
# Before
room = _pick_room(conn, group, students)
if not room:
    unscheduled.append({"code": code, "reason": f"no room at {centre}..."})

# After
room, room_centre = _pick_room(conn, group, students)
if not room:
    unscheduled.append({
        "code": code,
        "reason_code": "NO_ROOM_CAPACITY",
        "reason": f"no room at {centre} with capacity >= {students}",
        "category": "DATA_PROBLEM",
    })
    continue
if room_centre != centre:
    h3_exceptions.append({
        "code": code, "from_centre": centre, "to_centre": room_centre,
        "reason_code": "H3_TS_TM" if centre == "TS" else "H3_FALLBACK",
    })
```

*In `phase0_schedule_cadets()` (~line 1615):*
```python
# Before
room = _pick_room(conn, cls["code"], cls["student_count"] or 0)

# After
room, _room_centre = _pick_room(conn, cls["code"], cls["student_count"] or 0)
```

**Task 1.4 — Update `auto_assign_schedule()` return signature**
```python
# Before
return count, unscheduled

# After
return count, unscheduled, h3_exceptions   # h3_exceptions: list[dict]
```

Update the caller in `run_v4_from_bytes()`:
```python
_core_count, core_unscheduled, core_h3 = auto_assign_schedule(conn, cc_only=False)
```

---

#### Phase 2: Slot Failure Diagnostics (R3)

**File:** `api/timetable_scheduler.py` — `auto_assign_schedule()` inner loop

**Task 2.1 — Track per-day blocking reason**

Add a `slot_blocks` counter dict before the day loop:
```python
slot_blocks = {"SLOT_ROOM": 0, "SLOT_H9": 0, "SLOT_H6": 0, "SLOT_TEACHER": 0}
```

Inside the day loop, increment the relevant counter at each `continue`:
- H9 check fires → `slot_blocks["SLOT_H9"] += 1`
- H6 check fires (TKO before 10:00) → `slot_blocks["SLOT_H6"] += 1`
- Room booked / no room slot available → `slot_blocks["SLOT_ROOM"] += 1`
- Teacher not found on this slot → `slot_blocks["SLOT_TEACHER"] += 1`

**Task 2.2 — Emit prioritised reason code on failure**

Priority rule (resolves SpecFlow Q1 and Q2):
1. `NO_ROOM_CAPACITY` if room was None from `_pick_room()` (handled in Task 1.3)
2. `SLOT_H9` if all days blocked by H9
3. `SLOT_H6` if dominant (TKO class, no valid time)
4. `SLOT_ROOM` if room was present but occupied all available slots
5. `SLOT_TEACHER` as default (room existed, teacher failed)

```python
# At the "no free slot" fallback point (~line 790):
if not placed:
    if slot_blocks["SLOT_H9"] >= len(days_tried) and slot_blocks["SLOT_H9"] > 0:
        rc, cat = "SLOT_H9", "ALGORITHM_ISSUE"
    elif slot_blocks["SLOT_H6"] > 0 and group.startswith("TK"):
        rc, cat = "SLOT_H6", "ALGORITHM_ISSUE"
    elif slot_blocks["SLOT_ROOM"] >= slot_blocks["SLOT_TEACHER"]:
        rc, cat = "SLOT_ROOM", "ALGORITHM_ISSUE"
    else:
        rc, cat = "SLOT_TEACHER", "ALGORITHM_ISSUE"

    unscheduled.append({
        "code": code,
        "reason_code": rc,
        "reason": f"no free slot: {rc} was primary blocker ({slot_blocks})",
        "category": cat,
    })
```

---

#### Phase 3: Teacher Failure Diagnostics (R4)

**File:** `api/timetable_scheduler.py` — `assign_teachers()`

**Task 3.1 — Post-failure diagnostic query**

After all 5 passes fail and `lec1` is still None, run a lightweight query to determine the TEACHER_* code. This query runs per unassigned class (SpecFlow Q2 resolved: yes, needed).

```python
if not lec1:
    # Diagnostic: determine WHY all passes failed
    qual_count = conn.execute("""
        SELECT COUNT(DISTINCT t.id) FROM teachers t
        JOIN teacher_subjects ts ON ts.teacher_id = t.id
        WHERE ts.subject_code = ?
    """, (subj,)).fetchone()[0]

    if qual_count == 0:
        t_reason_code = "TEACHER_NONE"
        t_category    = "DATA_PROBLEM"
    else:
        # Check if all qualified teachers are over cap
        avail_count = conn.execute("""
            SELECT COUNT(DISTINCT t.id) FROM teachers t
            JOIN teacher_subjects ts ON ts.teacher_id = t.id
            LEFT JOIN schedule s ON s.teacher1_id = t.id
            WHERE ts.subject_code = ?
            GROUP BY t.id
            HAVING COUNT(s.id) < ?
        """, (subj, _TEACHER_WEEKLY_SESSION_CAP)).fetchone()
        if not avail_count:
            t_reason_code = "TEACHER_CAPACITY"
            t_category    = "DATA_PROBLEM"
        else:
            # Teachers exist and have cap space — likely H8 or availability
            t_reason_code = "TEACHER_AVAIL"
            t_category    = "DATA_PROBLEM"

    unassigned_detail.append({
        "code":        code,
        "subj":        subj,
        "centre":      group[:2] if len(group) >= 2 else group,
        "students":    students,
        "reason_code": t_reason_code,
        "reason":      f"teacher assignment failed: {t_reason_code}",
        "category":    t_category,
    })
    unassigned.append(code)   # keep list[str] for backward compat
```

**Task 3.2 — Return `unassigned_detail` from `assign_teachers()`**

Change return signature:
```python
# Before
return unassigned, warnings

# After
return unassigned, warnings, unassigned_detail   # unassigned_detail: list[dict]
```

Update caller in `run_v4_from_bytes()`:
```python
unassigned, s1_warnings, unassigned_detail = assign_teachers(conn)
```

---

#### Phase 4: Unified Issues + Assumptions Stats (R5, R8)

**File:** `api/timetable_scheduler.py` — `run_v4_from_bytes()`

**Task 4.1 — Enrich all failure items before `del conn`**

Before the `del conn` line (~line 1712), join unscheduled and unassigned_detail with the DB to attach subject/centre/students:

```python
# Enrich unscheduled_rooms items
enriched_unscheduled = []
for item in unscheduled_rooms:
    code = item["code"]
    row = conn.execute("""
        SELECT c.student_count, sub.code as subj_code, cg.centre
        FROM classes c
        JOIN class_groups cg ON cg.id = c.group_id
        JOIN subjects sub    ON sub.id  = c.subject_id
        WHERE c.code = ?
    """, (code,)).fetchone()
    enriched_unscheduled.append({
        **item,
        "students": row["student_count"] if row else 0,
        "subj":     row["subj_code"]     if row else "",
        "centre":   row["centre"]        if row else "",
    })

unscheduled_rooms = enriched_unscheduled
```

**Task 4.2 — Build `stats["issues"]` unified list**

After enrichment, combine all failure types into one list:
```python
all_issues = []

# Slot / room failures
for item in unscheduled_rooms:
    all_issues.append({
        "code":        item["code"],
        "subj":        item.get("subj", ""),
        "centre":      item.get("centre", ""),
        "students":    item.get("students", 0),
        "reason_code": item.get("reason_code", "NO_ROOM_CAPACITY"),
        "reason":      item["reason"],
        "category":    item.get("category", "DATA_PROBLEM"),
        "suggested":   _SUGGESTED_ACTIONS.get(item.get("reason_code", ""), "Contact Jo"),
    })

# Teacher failures
for item in unassigned_detail:
    all_issues.append({
        "code":        item["code"],
        "subj":        item["subj"],
        "centre":      item["centre"],
        "students":    item.get("students", 0),
        "reason_code": item["reason_code"],
        "reason":      item["reason"],
        "category":    item["category"],
        "suggested":   _SUGGESTED_ACTIONS.get(item["reason_code"], "Contact Jo"),
    })

stats["issues"]      = all_issues
stats["assumptions"] = [
    {"question": "J1", "assumption": "TS→TM cross-centre enabled",
     "affected": [e["from_centre"] + "→" + e["to_centre"] + ":" + e["code"]
                  for e in core_h3 if e["reason_code"] == "H3_TS_TM"]},
    {"question": "J8", "assumption": "Cross-centre fallback for room overflow",
     "affected": [e["from_centre"] + "→" + e["to_centre"] + ":" + e["code"]
                  for e in core_h3 if e["reason_code"] == "H3_FALLBACK"]},
]
```

**Task 4.3 — Define `_SUGGESTED_ACTIONS` constant (near top of file)**

```python
_SUGGESTED_ACTIONS = {
    "NO_ROOM_CAPACITY": "Increase room capacity at this centre, or split the class",
    "SLOT_ROOM":        "Room exists but fully booked — check for scheduling conflicts",
    "SLOT_TEACHER":     "Room available but no teacher free — check teacher loading panel",
    "SLOT_H9":          "Student group has conflicting class on same day",
    "SLOT_H6":          "No valid TKO time slot (class must start ≥ 10:00)",
    "TEACHER_NONE":     "Add a teacher with this subject qualification to the Teachers sheet",
    "TEACHER_CAPACITY": "All qualified teachers at weekly cap — add more or reduce other sessions",
    "TEACHER_AVAIL":    "All qualified teachers unavailable on the scheduled day",
    "TEACHER_H8":       "No teacher can reach this centre within 90 min of prior commitment",
    "H3_TS_TM":         "J1 assumption applied — confirm TS→TM with Jo",
    "H3_FALLBACK":      "J8 assumption applied — confirm cross-centre with Jo",
    "NET_SHORTFALL":    "Add more Net teachers to the Net Teachers sheet",
    "CADET_ERROR":      "Cadet class could not be placed on Mon/Wed/Fri",
}
```

---

#### Phase 5: Excel "Issues" Sheet (R7)

**File:** `api/timetable_scheduler.py` — `write_output_fast()`

**Task 5.1 — Add "Issues" sheet parameter and sheet creation**

Change signature: `write_output_fast(results, english_weekly=None, issues=None, assumptions=None)`

```python
if issues:
    ws_issues = wb.create_sheet("Issues")
    headers = ["Class Code", "Centre", "Subject", "Students",
               "Category", "Reason Code", "Reason", "Suggested Action"]
    ws_issues.append(headers)
    # Style header row bold
    for cell in ws_issues[1]:
        cell.font = Font(bold=True)
    # Group and sort: DATA_PROBLEM first, then ALGORITHM_ISSUE, then ASSUMPTION
    for item in sorted(issues, key=lambda x: ["DATA_PROBLEM","ALGORITHM_ISSUE","ASSUMPTION"].index(x.get("category","DATA_PROBLEM"))):
        ws_issues.append([
            item.get("code", ""),
            item.get("centre", ""),
            item.get("subj", ""),
            item.get("students", 0),
            item.get("category", ""),
            item.get("reason_code", ""),
            item.get("reason", ""),
            item.get("suggested", ""),
        ])

    if assumptions:
        ws_issues.append([])
        ws_issues.append(["--- ASSUMPTIONS (pending Jo confirmation) ---"])
        ws_issues.append(["Question", "Assumption", "Affected Classes"])
        for a in assumptions:
            ws_issues.append([
                a.get("question", ""),
                a.get("assumption", ""),
                ", ".join(a.get("affected", [])),
            ])
```

**Task 5.2 — Update caller**

```python
output_bytes = write_output_fast(
    results,
    english_weekly=english_weekly,
    issues=stats.get("issues"),
    assumptions=stats.get("assumptions"),
)
```

---

#### Phase 6: UI Issues Panel (R6)

**File:** `api/v4.html`

**Task 6.1 — CSS (add after loading-panel CSS, ~line 86)**

```css
/* Issues panel */
.issues-panel { margin-bottom: 16px; }
.issues-panel summary {
    font-size: .88rem; font-weight: 700; cursor: pointer;
    padding: 8px 12px; border-radius: 8px 8px 0 0;
    background: #fef3c7; border: 1px solid #fde68a; color: #92400e;
    list-style: none; display: flex; align-items: center; gap: 8px;
}
.issues-panel[open] summary { border-radius: 8px 8px 0 0; border-bottom: none; }
.issues-inner { background: #fff; border: 1px solid #fde68a; border-top: none;
    border-radius: 0 0 8px 8px; padding: 12px 14px; }
.issue-group-header { font-size: .8rem; font-weight: 700; margin: 10px 0 4px;
    padding: 3px 6px; border-radius: 4px; }
.issue-group-header.data    { background: #fee2e2; color: #991b1b; }
.issue-group-header.algo    { background: #e0e7ff; color: #3730a3; }
.issue-group-header.assume  { background: #d1fae5; color: #065f46; }
.issue-row { display: flex; gap: 8px; align-items: baseline;
    padding: 3px 0; border-bottom: 1px solid #f3f4f6; font-size: .82rem; }
.issue-code { font-weight: 600; min-width: 120px; }
.issue-reason { color: #374151; flex: 1; }
.issue-suggest { color: #6b7280; font-size: .78rem; }
.issues-badge { background: #fcd34d; color: #92400e; border-radius: 999px;
    padding: 1px 8px; font-size: .78rem; font-weight: 700; }
```

**Task 6.2 — HTML element (inside `.results.card`, after loading panel, ~line 244)**

```html
<details class="issues-panel" id="issuesPanel" style="display:none" open>
  <summary>
    ⚠ Issues &amp; Assumptions
    <span class="issues-badge" id="issuesBadge">0</span>
  </summary>
  <div class="issues-inner" id="issuesInner"></div>
</details>
```

Note: `open` attribute makes it expanded by default per R6 requirement.

**Task 6.3 — JS render in `renderResults()` (after loading panel render block)**

```javascript
// Issues panel
const issues = stats.issues || [];
const issuesPanel = document.getElementById('issuesPanel');
const issuesBadge = document.getElementById('issuesBadge');
if (issues.length > 0) {
    issuesPanel.style.display = '';
    // Count unique class codes
    const uniqueCodes = new Set(issues.map(i => i.code));
    issuesBadge.textContent = uniqueCodes.size;

    const groups = { DATA_PROBLEM: [], ALGORITHM_ISSUE: [], ASSUMPTION: [] };
    issues.forEach(i => {
        const cat = i.category || 'DATA_PROBLEM';
        if (groups[cat]) groups[cat].push(i);
    });

    const catLabel = {
        DATA_PROBLEM:    { cls: 'data',   label: '🔴 Data Problems — needs input data update' },
        ALGORITHM_ISSUE: { cls: 'algo',   label: '🔵 Scheduling Conflicts — may resolve automatically' },
        ASSUMPTION:      { cls: 'assume', label: '🟡 Assumptions Applied — confirm with Jo' },
    };

    let html = '';
    for (const [cat, items] of Object.entries(groups)) {
        if (!items.length) continue;
        const { cls, label } = catLabel[cat];
        html += `<div class="issue-group-header ${cls}">${label} (${items.length})</div>`;
        items.forEach(i => {
            html += `<div class="issue-row">
                <span class="issue-code">${i.code}</span>
                <span class="issue-reason">${i.reason}</span>
                <span class="issue-suggest">${i.suggested || ''}</span>
            </div>`;
        });
    }

    // Append assumption details from stats.assumptions
    const assumptions = stats.assumptions || [];
    if (assumptions.length > 0) {
        html += `<div class="issue-group-header assume">📋 Assumption Details</div>`;
        assumptions.forEach(a => {
            const n = (a.affected || []).length;
            html += `<div class="issue-row">
                <span class="issue-code">${a.question}</span>
                <span class="issue-reason">${a.assumption} (${n} classes affected)</span>
            </div>`;
        });
    }

    document.getElementById('issuesInner').innerHTML = html;
} else {
    issuesPanel.style.display = 'none';
}
```

---

### System-Wide Impact

#### Interaction Graph

```
_pick_room()
  → returns (room_code, centre_used)
  → if centre_used ≠ home_centre: auto_assign_schedule() records h3_exceptions

auto_assign_schedule()
  → returns (count, unscheduled, h3_exceptions)
  → run_v4_from_bytes() unpacks 3-tuple

assign_teachers()
  → returns (unassigned, warnings, unassigned_detail)
  → run_v4_from_bytes() unpacks 3-tuple

run_v4_from_bytes()
  → enriches unscheduled + unassigned_detail BEFORE del conn
  → builds stats["issues"] + stats["assumptions"]
  → passes to write_output_fast(issues=..., assumptions=...)
  → passes full stats to index.py (verbatim)

index.py
  → no changes needed (stats passed verbatim)

v4.html renderResults()
  → reads stats.issues → Issues panel
  → reads stats.assumptions → Assumption details
  → stats.unassigned (list[str]) still read by existing code (unchanged)
```

#### Error & Failure Propagation

- If `_pick_room()` raises an unexpected exception, it propagates to `auto_assign_schedule()` which has no try/except — will surface as HTTP 500. No change to this behaviour.
- Post-failure diagnostic query in `assign_teachers()` uses `fetchone()` — returns `None` safely if no rows, defaulting to `TEACHER_AVAIL`.
- If `stats["issues"]` is missing (old code path / v3 route), `v4.html` uses `|| []` defensively.

#### State Lifecycle Risks

- `stats["unassigned"]` remains `list[str]` — no breaking change to existing frontend code
- `h3_exceptions` is a local list; if `auto_assign_schedule()` is called twice (CC mode), exceptions are separately tracked per call
- DB enrichment loop (`enriched_unscheduled`) is read-only — no risk of partial DB mutation

#### API Surface Parity

- `/api/schedule/v4` → gains `stats.issues` and `stats.assumptions` fields (additive, non-breaking)
- `/api/schedule` (v3 route) → unaffected; does not call `run_v4_from_bytes()`

#### Integration Test Scenarios

1. Upload Excel with TS class → verify `stats.assumptions` has J1 entry, TS class scheduled at TM room
2. Upload Excel with FL class (students > FL room capacity) → verify cross-centre fallback fires, `stats.assumptions` has J8 entry
3. Upload Excel with subject that has no teachers → verify Issues panel shows `TEACHER_NONE` under DATA_PROBLEM
4. Upload Excel with TKO class where all slots pre-10:00 → verify `SLOT_H6` under ALGORITHM_ISSUE
5. Download Excel → verify "Issues" sheet present with same data as UI panel

---

## Acceptance Criteria

### Functional

- [ ] `_TS_TM_ENABLED = True` set; TS classes that previously failed now schedule at TM (H3 exception warning recorded)
- [ ] Classes at FL/SW/TKO/ST/TM that previously failed `NO_ROOM_CAPACITY` now try cross-centre fallback; if fallback succeeds, class is scheduled with ASSUMPTION record
- [ ] `auto_assign_schedule()` emits one of `{NO_ROOM_CAPACITY, SLOT_ROOM, SLOT_TEACHER, SLOT_H9, SLOT_H6}` per failure — no bare strings
- [ ] `assign_teachers()` emits one of `{TEACHER_NONE, TEACHER_CAPACITY, TEACHER_AVAIL, TEACHER_H8}` per unassigned class
- [ ] `stats["issues"]` contains all failures with category, reason_code, suggested action
- [ ] `stats["assumptions"]` lists J1 and J8 with affected class codes
- [ ] UI Issues panel renders on results page, open by default, grouped by category, badge shows unique class count
- [ ] Output Excel includes "Issues" sheet with correct columns and an "Assumptions" section

### Non-Functional

- [ ] `stats["unassigned"]` remains `list[str]` — no regression in existing unassigned display
- [ ] v3 route (`/api/schedule`) unaffected by changes
- [ ] Scheduler still runs in < 30 seconds for full Excel input (no significant performance regression from new diagnostic queries)

### Quality Gates

- [ ] Manual test: upload `data/input/Planning for Timetable.xlsx` → Issues panel shows, no Python traceback
- [ ] Manual check: TS classes now appear in schedule output (previously unscheduled due to J1)
- [ ] Manual check: output Excel has "Issues" sheet

---

## Success Metrics

- Zero classes silently dropped (all failures visible in UI and Excel)
- TS classes now scheduled (J1 assumption active)
- FL/SW/TKO/ST/TM room-overflow classes schedule via cross-centre fallback (J8 assumption active)
- Every item in Issues panel has a "Suggested Action" Jo can act on directly

---

## Dependencies & Risks

| Item | Status | Impact |
|---|---|---|
| J1 (TS→TM) | Assumed True | Reverse: set `_TS_TM_ENABLED = False` |
| J8 (cross-centre) | Assumed allowed | Reverse: remove Phase 3 fallback from `_pick_room()` |
| `_TRAVEL_TIME` accuracy | Assumed correct | If wrong, cross-centre fallback may pair wrong centres |
| `teacher_subjects` table schema | Must verify during impl | Diagnostic query in Task 3.1 must match actual DB schema |

---

## Implementation Order

1. Phase 1 (Foundation) — most risk, must do first
2. Phase 2 (R3 slot codes) — depends on Phase 1 inner loop access
3. Phase 3 (R4 teacher codes) — independent, can be done in parallel with Phase 2
4. Phase 4 (stats unification) — depends on Phases 1–3
5. Phase 5 (Excel sheet) — depends on Phase 4 API
6. Phase 6 (UI) — depends on Phase 4 API, can be done alongside Phase 5
7. Run Flask server and test manually

Total estimate: 1 coding session (~3–4 hours)

---

## Sources & References

### Origin
- **Origin document:** [docs/brainstorms/2026-06-09-full-coverage-diagnostics-requirements.md](../brainstorms/2026-06-09-full-coverage-diagnostics-requirements.md)
  Key decisions carried forward: J1/J8 assumption strategy, DATA_PROBLEM/ALGORITHM_ISSUE/ASSUMPTION taxonomy, UI panel + Excel sheet dual output

### Internal References

- `_pick_room()`: `api/timetable_scheduler.py:550`
- `auto_assign_schedule()`: `api/timetable_scheduler.py:652`
- `assign_teachers()`: `api/timetable_scheduler.py:974`
- `_find_teacher()`: `api/timetable_scheduler.py:1105`
- `write_output_fast()`: `api/timetable_scheduler.py:1352`
- `run_v4_from_bytes()`: `api/timetable_scheduler.py:1663`
- `collect_stats()`: `api/timetable_scheduler.py:1442`
- `_TRAVEL_TIME` matrix: `api/timetable_scheduler.py:187`
- `_TS_TM_ENABLED` toggle: `api/timetable_scheduler.py:159`
- Loading panel pattern (HTML): `api/v4.html:241`
- Loading panel JS render: `api/v4.html:606`
- `index.py` stats passthrough: `api/index.py:95`

### Related Work

- B3 fix (H3 fallback removal): `docs/plans/2026-06-02-change-log.md`
- B6 room overflow issue: `MASTER_PLAN.md`
- J1 (TS→TM) and J8 (room overflow): `MASTER_PLAN.md` Pending Jo section
