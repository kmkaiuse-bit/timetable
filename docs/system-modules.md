# System Module Map

One line per module/function so future changes can be targeted. **🎯 = class-scheduling
("分班") logic — change these when a scheduling rule changes.** Line numbers are a guide;
search the symbol name if they drift.

---

## Deploy / entry points

| File | One-liner |
|------|-----------|
| `api/index.py` | Flask app + Vercel entry; serves `v4.html`, handles upload → calls `run_v4_from_bytes`, returns JSON + Excel. |
| `vercel.json` | Vercel build: `@vercel/python` on `api/index.py`, all routes → it. Do not move `api/`. |
| `requirements.txt` | Runtime deps: `openpyxl`, `flask`. |

---

## `api/timetable_scheduler.py` — engine (single source of truth)

### 🎯 Scheduling rules & toggles (top of file — change here first)

| Symbol | Line~ | One-liner |
|--------|-------|-----------|
| 🎯 `_DAY_PRIORITY` | 124 | **Days DAE subjects may use — currently Mon/Tue/Thu.** Add/remove a day here. |
| 🎯 `_CADET_DAY_PRIORITY` | 2006 | **Days cadet classes may use — Mon/Wed/Fri.** Cadet exception, independent of above. |
| 🎯 `_CADET_DAYS` | 2002 | Set of cadet days (membership test). Keep in sync with `_CADET_DAY_PRIORITY`. |
| 🎯 `_TIME_BLOCKS` / `_SINGLE_SLOTS` | 129/134 | Standard 4-hr blocks and 2-hr slots for non-TKO centres. |
| 🎯 `_TKO_TIME_BLOCKS` / `_TKO_SINGLE_SLOTS` | 137/142 | TKO-only slots (start ≥ 10:00, H6). |
| 🎯 `_TEACHER_WEEKLY_SESSION_CAP` | 148 | Soft loading cap (6). Warn-only threshold; raising/lowering changes the warning point. |
| 🎯 `_TS_TM_ENABLED` | 170 | H3 TS→TM cross-centre toggle (J1). |
| 🎯 `_CC_MAX_TRAVEL_MIN` | 191 | CC-combine max travel filter (J2); `None` = disabled. |
| 🎯 `_H8_MAX_TRAVEL_MIN` | 167 | Cross-centre travel cap (90 min): ≤ warn, > block. |
| 🎯 `_H9_ALLOWED_CROSS_CENTRE` | 194 | Student same-day cross-centre exceptions (CSW↔SSP, TW↔CSW/SSP). |
| `_ENG_*` (Net) | 151–163 | English/Net teacher rules for DAE102 (min blocks, travel, exempt groups). |
| `_TRAVEL_TIME` / `_NEAREST_CENTRES` | 216/232 | Centre-to-centre travel matrix + fallback ordering. |
| `_GROUP_CENTRE_ALIAS` / `_CENTRE_RANK` | 117/203 | Group→centre aliases; centre ranking for S4 tie-break. |

### Input loading (Excel → SQLite)

| Function | Line~ | One-liner |
|----------|-------|-----------|
| `build_db` | 575 | Orchestrates loading all sheets into the in-memory DB; returns `(conn, n_sched, n_unavail)`. |
| `_load_classes` | 348 | Reads the **Class list** sheet, incl. optional **CC Group** column → `classes.cc_group`. 🎯 combine input. |
| `_load_subjects` / `_load_rooms` / `_load_teachers` | 297/317/419 | Load subjects (loading hrs), rooms (capacity/centre), teachers (+ subject quotas). |
| `_load_availability` / `_load_centre_preferences` / `_load_net_teachers` | 459/526/552 | Teacher unavailability (S1), centre avoid-prefs (S1b), Net flags. |
| `_load_existing_schedule` | 502 | Loads pre-filled Day/Time (v3 mode); its count drives the auto-assign gate (`n_sched`). |

### 🎯 Scheduling phases (the core "分班")

| Function | Line~ | One-liner |
|----------|-------|-----------|
| 🎯 `phase0_schedule_cadets` | 2018 | **Phase 0** — place cadet classes on Mon/Wed/Fri (uses `_CADET_DAY_PRIORITY`), fixed rooms first. |
| 🎯 `auto_assign_schedule` | 761 | **Phase A** — place Core classes (cc_group empty): picks day (`_DAY_PRIORITY`), slot, room; enforces H3/H6/H9/S2/S3. The main scheduling loop. |
| 🎯 `cc_assign_schedule` | 995 | **Phase B** — place CC-combine groups as one session (L1/L2/H5/S4). |
| 🎯 `_find_cc_day` | 968 | Finds a combine day with no member's cross-centre conflict (uses `_DAY_PRIORITY`; the 3-day limit bites here). |
| `_select_cc_centre` | 948 | Orders candidate centres for a combine group by student count (S4). |
| `_pick_room` / `_allowed_centres_for_group` | 644/634 | Room selection incl. cross-centre fallback within travel cap (H3/J8). |
| `_build_teacher_capacity` | 700 | Pre-computes max concurrent classes per (subject, day, start) for Phase A. |
| `_has_tiandi_violation` | 741 | S2 天地堂 gap check (students only). |
| `_room_free` | 691 | Room-slot availability check (H2). |

### 🎯 Teacher assignment

| Function | Line~ | One-liner |
|----------|-------|-----------|
| 🎯 `assign_teachers` | 1123 | 5-pass Lec1 + backups; relaxes quota→centre→**H4 cap (soft warn)**→availability. Never blocks on loading. |
| `_find_teacher` | 1285 | The per-pass SQL query (qualification, double-booking, H4 filter, availability, centre pref). |
| `assign_english_weekly` | 1432 | DAE102 Net weekly assignment across 3 five-week blocks (Net hours, travel). |

### Output & stats

| Function | Line~ | One-liner |
|----------|-------|-----------|
| `run_v4_from_bytes` | 2098 | **Top-level v4 entry**: build_db → (if `n_sched==0`) Phase 0/A/B → teachers → English → collect → Excel bytes + stats. |
| `collect_results` / `collect_stats` | 1630/1868 | Build the results list and the stats/issues/warnings/loading payload for the UI. |
| `write_output_fast` | 1706 | Write the output workbook incl. 15-week semester calendar + Issues sheet. |
| `generate_class_dates` / `_HK_HOLIDAYS` | 1692/1662 | Semester date generation and HK holiday set. |

---

## `api/v4.html` — Web UI

| Section / function | One-liner |
|--------------------|-----------|
| `.container` CSS (~line 32) | 🎯 Overall page width (1440px). Widen/narrow the whole layout here. |
| `.tt-table` / `.cell-*` CSS (~151–179) | Day-view table + cell sizing/subject-code badge styling. |
| `.week-table` / `.wc-*` CSS (~122–132) | Week-view (room & teacher) table + cell styling. |
| `.cal-table` CSS (~137–143) | Month/semester calendar styling. |
| `DAY_ORDER` (~898) | 🎯 The five weekdays shown in every view (all views now always render Mon–Fri). |
| `setView` (~940) | View switcher: day / week(room) / teacher / month. |
| `renderGrid` / `showDayGrid` (~918/1074) | Day view — tabs + per-room timetable; renders subject-code + group per cell. |
| `renderWeekOverview` (~994) | Week·by-room — rooms × (day×slot); subject code + group + lecturer per cell. |
| `renderTeacherWeek` (~955) | Week·by-teacher — teachers × (day×slot); subject code + group per cell; flags double-booking. |
| `renderMonthCalendar` (~1033) | Month — 15-week semester calendar with HK holidays (no class cells). |
| `rebuildGrid` (~1062) | Builds the `_grid[day][room][slot]` structure from `state.results`. |
| Teacher Loading panel (~744) | Renders per-teacher sessions/hours; red 超限 badge above 6 (H4 reminder). |
| Issues panel (~773+) | Renders unplaced classes + assumptions with reason-code translations. |

---

## Where to change common things

- **Which days a subject type may use** → `_DAY_PRIORITY` (DAE) / `_CADET_DAY_PRIORITY` (cadet).
- **Time slots / TKO windows** → `_TIME_BLOCKS`, `_SINGLE_SLOTS`, `_TKO_*`.
- **Teacher loading threshold** → `_TEACHER_WEEKLY_SESSION_CAP` (+ pass 4 in `assign_teachers`).
- **Combine groups** → input Excel "CC Group" column (data), logic in `cc_assign_schedule` / `_find_cc_day`.
- **Room / cross-centre rules** → `_pick_room`, `_H9_ALLOWED_CROSS_CENTRE`, `_H8_MAX_TRAVEL_MIN`, `_TRAVEL_TIME`.
- **UI size / views / subject code** → `api/v4.html` CSS + the four `render*` functions.
