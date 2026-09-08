---
date: 2026-06-09
topic: full-coverage-diagnostics
---

# Full Coverage & Diagnostics

## Problem Frame

The HKIT timetable scheduler V6.2 currently fails to schedule some classes silently or with vague error messages. Two categories of problem exist:

1. **Algorithm gaps** — constraints or assumptions that prevent scheduling but could be resolved with code changes or interim assumptions
2. **Data problems** — room capacity too small, teacher list incomplete, etc. — these need Jo to update the input data

Jo needs clear, actionable output to know exactly what to fix in the data. Planners need maximum coverage so the output timetable is as complete as possible.

## Requirements

### Algorithm Improvements

- **R1.** Enable TS→TM cross-centre scheduling (`_TS_TM_ENABLED = True`).
  - *Assumption replacing J1:* TS students may attend class at TM centre.
  - Record each TS→TM assignment with an H3-exception warning in stats.

- **R2.** When a class's home centre has no room meeting capacity, try the nearest valid centre (cross-centre fallback) before marking unscheduled.
  - *Assumption replacing J8:* Cross-centre allowed when home centre has zero suitable rooms.
  - Use the existing `_TRAVEL_TIME` matrix to rank fallback centres by proximity.
  - Record each fallback with an H3-exception warning including `from_centre → to_centre`.
  - This replaces the current hard-fail behaviour introduced by the B3 fix.

- **R3.** When a class fails "no free slot", diagnose and record the specific blocking constraint:
  - `SLOT_ROOM`: All slots on all days have no free room of sufficient size at target centre
  - `SLOT_TEACHER`: Room available but all qualified teachers fully booked on valid days
  - `SLOT_H6`: TKO class — all pre-10:00 slots eliminated, not enough remaining slots
  - `SLOT_H9`: Student centre conflict eliminates all candidate days
  - This replaces the generic `"no free slot across all days"` message.

- **R4.** When `assign_teachers()` fails to assign a teacher to a scheduled class, record why:
  - `TEACHER_NONE`: No teacher in the system has the required subject qualification
  - `TEACHER_CAPACITY`: Qualified teachers exist but all exceed the weekly session cap (H4)
  - `TEACHER_AVAIL`: Qualified teachers exist but all unavailable on the scheduled day (S1)
  - `TEACHER_H8`: Qualified teachers exist but all would violate H8 travel constraint
  - This replaces the current silent "unassigned" marker.

### Diagnostic Output

- **R5.** Unscheduled and unassigned items must be categorised as:
  - `DATA_PROBLEM` — room capacity, no qualified teacher: *Jo must update input data*
  - `ALGORITHM_ISSUE` — slot exhaustion, scheduling conflict: *may be fixable in code*
  - `ASSUMPTION` — J1/J8/other assumption applied: *confirm with Jo before release*
  - Each item includes: class code, centre, subject, student count, failure reason code, human-readable message, and category.

- **R6.** UI: Add an **"Unscheduled / Issues" collapsible panel** (similar to loading panel) showing:
  - Grouped by category (DATA_PROBLEM / ALGORITHM_ISSUE / ASSUMPTION)
  - Per-item: class code + reason + suggested action
  - Count badge on the panel header (e.g. "⚠ 12 issues")
  - Shown automatically if any issues exist (not hidden by default)

- **R7.** Excel output: Add an **"Issues" sheet** to the output workbook containing:
  - Columns: Class Code | Centre | Subject | Students | Category | Reason | Suggested Action
  - One row per unscheduled class or unassigned teacher slot
  - Separate section for H3 exceptions (J1/J8 assumptions)
  - Usable as a handoff document for Jo

### Assumption Logging

- **R8.** All assumptions made in place of pending Jo answers must be recorded in `stats["assumptions"]`:
  - `{"question": "J1", "assumption": "TS→TM enabled", "affected_classes": [...]}`
  - `{"question": "J8", "assumption": "cross-centre fallback for FL/SW/TKO/ST/TM", "affected_classes": [...]}`
  - Displayed in UI and written to a dedicated "Assumptions" section in the Issues sheet.

## Success Criteria

- Zero unscheduled classes due to R1/R2 (i.e. the 19 previously-failing classes now schedule with warnings)
- Every failure has a reason code from R3/R4 — no more generic "no free slot" or silent unassigned
- UI Issues panel lists all problems on results page
- Excel output includes Issues sheet and Assumptions section
- Assumptions are flagged clearly so they can be reversed when Jo confirms

## Scope Boundaries

- **Not in scope:** Backtracking / re-scheduling previously placed classes to free up slots for a stuck class (too complex for now)
- **Not in scope:** J2 (CC travel threshold), J3 (cadet room codes), J5 (cost), J6/J7 (teacher list / CC groups) — these remain blocked on Jo's data
- **Not in scope:** Changing the 5-pass teacher assignment logic itself; only adding diagnostic output to it
- **Not in scope:** Net teacher shortfall (J9) — remains a data problem; already warned

## Key Decisions

- **Cross-centre fallback (R2):** Use `_TRAVEL_TIME` matrix to pick nearest centre, not arbitrary. Cap at H8 limit (90 min) to avoid extreme cross-centre.
- **Assumption J1:** `_TS_TM_ENABLED` set to `True` in code; toggling it back to `False` must remain easy (single constant).
- **Diagnostic granularity (R3/R4):** Reason codes are machine-readable (for UI grouping) + human-readable string (for Excel / Jo).
- **UI panel:** Always visible when issues > 0; don't require user to manually expand (unlike loading panel).

## Dependencies / Assumptions

- J1 (TS→TM): Assumed `True`. Reverse by setting `_TS_TM_ENABLED = False`.
- J8 (cross-centre): Assumed allowed. Reverse by removing R2 fallback logic.
- Existing `_TRAVEL_TIME` matrix is accurate for centre-to-centre distances.
- `assign_teachers()` currently returns `(unassigned_list, warnings_list)` — R4 requires adding a reason field to `unassigned_list`.

## Outstanding Questions

### Resolve Before Planning
*(none — all blocking decisions made via assumptions above)*

### Deferred to Planning
- [Affects R2][Technical] Which specific centres should be eligible as cross-centre fallbacks for each home centre? (Use `_TRAVEL_TIME` matrix ≤ 90 min as the filter.)
- [Affects R3][Technical] How to detect `SLOT_H9` efficiently — requires querying student schedule conflicts per candidate day.
- [Affects R4][Technical] Exact pass in `assign_teachers()` where each reason code should be recorded.

## Next Steps

→ `/ce:plan` for structured implementation planning
