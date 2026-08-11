# Timetable Updates — 2026-08

**Branch:** `feat/aug2026-timetable-updates` (off `master`; not merged — rollback = discard branch)
**Author:** implementation session 2026-08-11
**Core files:** `api/timetable_scheduler.py` (engine), `api/v4.html` (UI)

> Note: this document was created during the 2026-08 session. The `docs/PRD-timetable-scheduler.md`
> and `docs/timetable-constraints.md` files referenced verbally did not exist in the repo at that
> time; the constraint reference of record is `MASTER_PLAN.md` (Constraint Status table).

---

## Update 1 — DAE subjects: Mon / Tue / Thu only (no Wed / Fri)

**Requirement:** DAE curriculum subjects must not be scheduled on Wednesday or Friday.
Cadet classes keep their existing Mon / Wed / Fri exception.

**Change:** `api/timetable_scheduler.py`
- `_DAY_PRIORITY = ["Monday", "Tuesday", "Thursday"]` (was 5 days).
- New `_CADET_DAY_PRIORITY = ["Monday", "Wednesday", "Friday"]`.
- Phase 0 (`phase0_schedule_cadets`) now iterates `_CADET_DAY_PRIORITY` instead of
  `sorted(_CADET_DAYS, key=_DAY_PRIORITY.index)` — the old form would raise
  `ValueError` once Wed/Fri left `_DAY_PRIORITY`. **This is the key trap.**

**Scope:** only the v4 auto-assign path (empty Class-list answer, `n_sched == 0`).
Pre-filled v3 inputs (e.g. `Planning for Timetable.xlsx`, `n_sched = 108`) are used
as-is and are **not** affected by `_DAY_PRIORITY`.

**Measured impact** (auto-assign on `Planning for Timetable (clean).xlsx`):

| Metric | Before (5-day) | After (3-day) |
|---|---|---|
| Scheduled classes | 111 | **94** (−17) |
| Wed / Fri classes | 30 / 25 | **0 / 0** ✅ |
| Issues total | 9 | 27 |
| SLOT_ROOM (room shortage) | 1 | **14** |
| TEACHER_CAPACITY | 0 | 4 |

**Consequence (expected):** removing 2 of 5 days concentrates demand onto Mon/Tue/Thu
and sharply worsens room shortage (TS / FL / TM already tight). 17 classes become
unplaceable. This is a **capacity decision for Jo** — add rooms, accept fewer classes,
or allow a partial Wed/Fri exception.

---

## Update 2 — Teacher loading may exceed 6; warn only (never block)

**Verification result: already implemented — no code change made.**

- Soft cap constant: `_TEACHER_WEEKLY_SESSION_CAP = 6` (`api/timetable_scheduler.py`).
- `assign_teachers()` five-pass search: passes 1–2 prefer teachers under the cap;
  pass 4 relaxes it (`ignore_h4_cap=True`), assigns anyway, and emits the warning
  *"H4 (soft): teacher exceeds preferred weekly loading of 6 sessions"*. Never blocks.
- UI (`api/v4.html`) Teacher Loading panel flags any teacher above 6 with a red
  **超限** badge; the H4 warning also appears in the warnings list.

Confirmed on real data: one teacher currently trips the H4 soft warning, and the
loading panel colours them red — i.e. "exceed is allowed, and it reminds you."

**Per-teacher individual caps: not implemented and not recommended** — there is no
per-teacher cap data source in the input, and the requirement describes a single
global advisory threshold of 6, which is exactly the current behaviour.

---

## Update 3 — UI: bigger timetable + subject code in every cell

**Change:** `api/v4.html`
- **Wider overall container:** `.container` max-width `980px → 1440px` so the whole
  body container (and the timetable inside it) uses the full page width.
- **Enlarged all four views:** day grid (taller rows, wider time columns, larger
  fonts), week-by-room, week-by-teacher, and month / semester calendar.
- **Subject code per cell:** the subject code (e.g. `DAE101`) now shows as a distinct
  bold badge in every cell across day / week-by-room / week-by-teacher. The day view
  splits the class code into subject-code + group suffix. The month view is a date
  calendar (no class cells), so it is enlarged but carries no subject code.
- **Keep Wed / Fri in the UI:** all four views now always render the full Mon–Fri
  week, keeping Wed/Fri columns/tabs even when empty (guarded against `undefined`
  day buckets). DAE subjects no longer land there, but cadet classes and manual
  entries still can, and the empty columns keep the week structure legible.
- Help text updated: DAE = Mon/Tue/Thu, cadet = Mon/Wed/Fri; loading cap is advisory.

---

## Combine course (DAE106) — Jo's groupings, 2026-08

Captured into the input workbook's **"CC Group"** column (the scheduler detects any
header containing both "cc" and "group" — `_load_classes()`), via
`scripts/add_dae106_combine.py` → `data/input/Planning for Timetable (CC combine).xlsx`.

| CC Group | Classes | Target centre (Jo) |
|---|---|---|
| CC-DAE106-1 | DAE106_CS6 + DAE106_WT2 | CSW |
| CC-DAE106-2 | DAE106_TW1 + DAE106_TW3 | TW |
| CC-DAE106-3 | DAE106_TS1 + DAE106_TS2 | TM |
| CC-DAE106-4 | DAE106_TS3 + DAE106_TS4 + DAE106_TM3 | TM |
| CC-DAE106-5 | DAE106_KT3 + DAE106_KT4 | KT |

**Items for Jo to confirm:**
1. **Two classes were missing** from the Class list (`DAE106_TM3`, `DAE106_KT4`).
   The script appended them (student counts 8 and 9 from Jo's sheet; subject
   names/loading cloned from a sibling DAE106 row). Confirm their details.
2. **Student-count discrepancy:** Jo's combine sheet numbers differ from the workbook
   for several classes (e.g. CS6 10 vs 33). Existing counts were **not** overwritten.
   Confirm which snapshot is authoritative.
3. **Combine only runs in auto-assign mode** (`n_sched == 0`); the CC column has no
   effect in a pre-filled v3 file.
4. **Update 1 interaction:** restricting to 3 days worsens combine placement — with
   5 days the engine placed 7 combine instances (2 errors); with 3 days only 2 (4
   groups "L2 exhausted"). Combine courses may need Wed/Fri or a relaxed rule.

---

## Rollback

All work is on branch `feat/aug2026-timetable-updates` in an isolated git worktree;
`master` and the main repo path are untouched. Rollback = do not merge / delete the
branch. The generated `(CC combine).xlsx` is a copy; the original inputs are unchanged.
