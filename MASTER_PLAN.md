田# HKIT Timetable Scheduler — Master Plan

**Last updated:** 2026-05-26
**Current version:** V6
**Branch:** `feat/v6-constraint-improvements`

---

## Project Goal

Automate HKIT's class timetable scheduling. Given an Excel input (classes, teachers, rooms), the system outputs a complete timetable that satisfies all hard constraints and optimises for soft constraints.

---

## System Overview

```
Excel Input (.xlsx)
    └─ api/timetable_scheduler.py  ← Python scheduling engine (SQLite in-memory)
           │
           ├─ Phase 0: Pre-schedule cadet classes (Mon/Wed/Fri fixed)
           ├─ Phase A: Schedule Core classes (H1–H9, S1–S4)
           ├─ Phase B: Schedule CC Combine classes (L1/L2/H5/S4)
           └─ Teacher Assignment: 5-pass system
                    │
                    └─ api/index.py  ← Flask API (Vercel serverless)
                             │
                             └─ api/v4.html  ← Web UI
```

**Deploy:** Vercel · `api/index.py` handles all routes

---

## Constraint Status

### Hard Constraints

| ID | Rule | Status | Notes |
|----|------|--------|-------|
| H1 | No teacher double-booking same slot | ✅ Implemented | Always enforced |
| H2 | No room double-booking same slot | ✅ Implemented | Always enforced |
| H3 | Students attend class at enrolled campus | ✅ Partial | Enforced via room selection; TS→TM toggle added (off by default, pending J1) |
| H4 | Teacher max 6 sessions/week | ✅ Soft cap | Warn only (changed from Hard in V6) |
| H5 | CC Combine: same subject + language | ✅ Implemented | Checked in Phase B |
| H6 | TKO: class start ≥ 10:00 | ✅ Implemented | V6: any slot ≥10:00 allowed |
| H7 | Teacher qualified for subject | ✅ Implemented | Always enforced, all passes |
| H8 | Teacher at one centre per day | ✅ Soft gate | V6: cross-centre allowed if travel ≤ 90 min (warn); >90 min = hard block |
| H9 | Students at one centre per day | ✅ Implemented | Exceptions: CSW↔SSP, TW↔CSW, TW↔SSP |

### Soft Constraints

| ID | Rule | Status | Notes |
|----|------|--------|-------|
| S1a | Teacher unavailability respected | ✅ Implemented | Pass 5 relaxes with warning |
| S1b | Teacher centre preference respected | ✅ Implemented | Pass 3 relaxes with warning |
| S2 | No 天地堂 gap for students | ✅ Implemented | Students only (not teachers) |
| S3 | Day rotation for same-subject classes | ✅ Implemented | `subj_day_offset` counter |
| S3 | Cadet class fixed Mon/Wed/Fri rooms | ⚠️ Framework only | Phase 0 added; awaiting J3 for room codes + detection details |
| S4 | CC Combine: prefer centre with most students | ✅ Implemented | Tie-break by `_CENTRE_RANK` |

### CC Combine Logic

| ID | Rule | Status | Notes |
|----|------|--------|-------|
| L1 | Find day where no student has Core class at different centre | ✅ Implemented | |
| L2 | If L1 fails, try next-best centre | ✅ Implemented | |
| — | CC distance filter | ⚠️ Ready, disabled | `_CC_MAX_TRAVEL_MIN = None`; activate when Jo confirms threshold (J2) |

---

## Teacher Assignment (5-Pass)

| Pass | What's relaxed | Warning |
|------|----------------|---------|
| 1 | Nothing (fully strict) | — |
| 2 | Subject quota | — |
| 3 | Centre preference (S1b) | `S1b: teacher assigned to non-preferred centre` |
| 4 | H4 weekly cap | `H4 (soft): teacher exceeds 6 sessions` |
| 5 | Unavailability (S1) | `S1: teacher assigned to unavailable slot` |

H8 never relaxed. Cross-centre with travel >90 min = hard block in all passes.

---

## Pending — Waiting for Jo

| # | Category | Question | Impact |
|---|----------|----------|--------|
| J1 | H3 | TS→TM exception — confirm assumption | **Assumed True** in V6.3 — reverse by setting `_TS_TM_ENABLED = False` |
| J2 | CC | CC Combine max travel time threshold (minutes)? | Set `_CC_MAX_TRAVEL_MIN = X` |
| J3 | S3 | Cadet class: exact room codes + Excel detection method? | Complete Phase 0 implementation |
| J4 | Teacher | English/Local/Net teacher types — scheduling difference? | ✅ Answered 2026-06-02: Net = Barrett + Chris Hon; ≥20 hrs/term/class; 30min travel; CS1-3+CS7 exempt |
| J5 | Cost | Include teacher price in scheduling? | New optimisation pass |
| J6 | Teacher | Teacher assignment list — when available? | Teacher auto-assignment can't run without this |
| J7 | CC | CC grouping full list — when available? | CC Combine can't run without this |
| J8 | Rooms | FL/SW/TKO/ST/TM centres: classes with students > room capacity — split/relocate/accept? | **Assumed cross-centre fallback** in V6.3 (nearest centre ≤90 min); Issues panel shows all H3 exceptions |
| J9 | English | Net requirement: is 1 five-week block = 20 Net hrs (so ≥1 block/term satisfies "≥20 Net hrs/term")? | **Assumed True** 2026-07-08 — `_ENG_NET_MIN_BLOCKS` set to 1; eliminated all false shortfall warnings |
| J10 | English | KT1 (觀塘) + WT2 can't get a Net teacher under full auto-assign — both land on the crowded Tuesday 14:00 slot and are >30 min travel from where the 2 Net teachers are committed. Reschedule these 2 classes to a lighter slot, add a 3rd Net teacher positioned for KT/WT, or accept no Net coverage? | Surfaced in Issues panel (🔵 排班衝突) with diagnostic — **awaiting Jo's feedback**. Adding a 3rd Net teacher proven ineffective (test 2026-07-08). |

---

## Known Bugs / Tech Debt

| # | Description | Status |
|---|-------------|--------|
| B1 | H8 not enforced for Lec2/Lec3 backup teachers | ✅ Fixed (2026-05-25) |
| B2 | `is_police_cadet` DB flag set but never read | ⚠️ Superseded by `_is_cadet_class()` — flag can be removed later |
| B3 | H3 fallback in `_pick_room()` could assign any room | ✅ Fixed (2026-06-02) — removed; now returns None |
| B4 | `phase0_schedule_cadets` inserted fake `_slot2` row | ✅ Fixed (2026-06-02) — uses time1+time2 properly |
| B5 | `auto_assign_schedule` silently dropped unscheduled classes | ✅ Fixed (2026-06-02) — returns `(count, unscheduled_list)` |
| B6 | FL/SW/TKO/ST/TM centres have classes exceeding room capacity | ⚠️ Exposed by B3 fix — needs Jo input (J8) |

---

## Development Progression

### Done

| Date | Version | What |
|------|---------|------|
| 2026-05 | V3 | Day/Time/Room pre-filled → assign teachers only |
| 2026-05 | V4 | Auto-assign Day/Time/Room/Teachers from scratch |
| 2026-05 | V5 | Constraint refinements, CC Combine two-phase |
| 2026-05-23 | V6 | H4 soft cap, H6 TKO relax, H9 exceptions, S1b centre prefs, S2 students-only |
| 2026-05-25 | V6.1 | `_TRAVEL_TIME` matrix, H8 soft gate, H3 toggle, S3 Phase 0, CC distance filter |
| 2026-06-02 | V6.2 | Bug fixes B3/B4/B5, MRV ordering, English Net teacher weekly assignment |
| 2026-06-09 | V6.3 | J1/J8 assumptions, reason codes, Issues panel + Excel sheet |
| 2026-07-08 | V6.4 | English auto-assign fixes: removed hard-coded `_ENG_WEEKLY_PREASSIGNED`, honorific-tolerant teacher name matching, Mode-2 availability check, `_ENG_NET_MIN_BLOCKS` 2→1 (J9). Data-quality fixes via `scripts/build_clean_input.py` (English availability, Mr./Ms. unification, cleared English Weekly, **cleared Class list answer → true V4 auto-assign**). **True V4 result on clean input: 111/111 scheduled, 0 unassigned (beats Jo's V3 answer 105/107), 0 availability conflicts, 4 Net shortfalls (KT1, WT2 × 2 terms — genuine 2-Net-teacher supply limit). NOTE: earlier "0 conflict" figures were a V3-mode artifact — the real input had Class list answer pre-filled, so run_v4 was replaying Jo's days, not auto-scheduling.** English-only trial (`— English.xlsx`): 27/27 incl. added CS7 (Oct extra, student count assumed 35) |

### Next (unblocked)

| Priority | Item | Blocked by |
|----------|------|-----------|
| High | Get Jo's answers (J2–J9) | Jo |
| High | Fix Net teacher shortfall (24 warnings) | Jo — more Net teachers or revised rules |
| Medium | Confirm J1 assumption (TS→TM currently enabled) | Jo confirmation |
| Medium | Confirm J8 assumption (cross-centre fallback currently enabled) | Jo confirmation |
| Medium | Set CC distance threshold | J2 |
| Medium | Complete S3 cadet class | J3 |
| Low | Cost optimisation | J5 |

### Blocked (needs data)

| Item | Blocked by |
|------|-----------|
| Teacher auto-assignment end-to-end | J6 (name list) |
| CC Combine end-to-end | J7 (grouping list) |
| Room overflow resolution | J8 (Jo decision on FL/SW/TKO/ST/TM capacity) |

---

## Key Files

| File | Purpose |
|------|---------|
| `api/timetable_scheduler.py` | Core scheduling engine — single source of truth |
| `api/index.py` | Flask API + route handler |
| `api/v4.html` | Web UI (V4/V6) |
| `api/template_v4.xlsx` | Excel input template for users |
| `data/input/Planning for Timetable.xlsx` | Current real data |
| `docs/algorithm/scheduling-algorithm.html` | Algorithm documentation (V6) |
| `docs/meetings/scheduling-meeting-v2.docx` | V6 constraint decisions (source of truth for rules) |
| `docs/questions/questions-for-jo.docx` | Outstanding questions tracker |
| `docs/plans/` | Individual feature implementation plans |

---

## Configurable Toggles (in `timetable_scheduler.py`)

| Constant | Default | Activate when |
|----------|---------|---------------|
| `_TS_TM_ENABLED` | `False` | Jo confirms TS→TM rule active (J1) |
| `_CC_MAX_TRAVEL_MIN` | `None` | Jo confirms CC distance threshold (J2) |
| `_H8_MAX_TRAVEL_MIN` | `90` | Adjust if Jo changes travel limit |
| `_TEACHER_WEEKLY_SESSION_CAP` | `6` | Adjust if cap changes |
| `_DAY_PRIORITY` | `[Mon, Tue, Thu]` | DAE subjects — no Wed/Fri (2026-08). Cadets use `_CADET_DAY_PRIORITY` |
| `_ENG_NET_MIN_BLOCKS` | `1` | Set to 2 if Jo confirms 1 block = 10 Net hrs (needs 2/term) — see J9 |

---

## 2026-08 Updates (branch `feat/aug2026-timetable-updates`, not merged)

See `docs/pending-updates-2026-08.md` for full detail, impact numbers, and Jo follow-ups.

1. **DAE Mon/Tue/Thu only** — `_DAY_PRIORITY` restricted; cadets keep Mon/Wed/Fri via
   new `_CADET_DAY_PRIORITY`. Auto-assign only. Worsens room shortage
   (scheduled 111→94, SLOT_ROOM 1→14 on clean input) — capacity decision for Jo.
2. **Teacher loading soft cap** — verified already correct (H4 warns above 6, never
   blocks; UI flags red). No code change; per-teacher caps not added (no data source).
3. **UI** — container widened 980→1440px, all four views enlarged, subject code shown
   in every cell, Wed/Fri columns kept even when empty.
4. **Combine course (DAE106)** — Jo's 5 groups captured into a "CC Group" column via
   `scripts/add_dae106_combine.py` → `data/input/Planning for Timetable (CC combine).xlsx`.
   Two classes (DAE106_TM3, DAE106_KT4) were missing and appended; student-count
   discrepancy vs Jo's sheet left unresolved — both need Jo's confirmation.
