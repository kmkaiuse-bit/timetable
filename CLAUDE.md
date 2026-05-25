# HKIT Timetable Scheduler — Claude Instructions

## Session Start (read these first, every time)

1. **`MASTER_PLAN.md`** — overall project status, constraint implementation state, pending Jo questions (J1–J7), configurable toggles
2. **`docs/plans/`** — find the most recent checklist or plan file (highest date) to know what was left off last session
3. **`docs/questions/questions-for-jo.docx`** — outstanding questions for Jo; check if any have been answered since last session

## Project Context

- **Language:** Python 3.13 + Flask + SQLite in-memory
- **Deploy:** Vercel serverless — `api/index.py` is the entry point, all routes go there
- **Core file:** `api/timetable_scheduler.py` — single source of truth for the scheduling algorithm
- **UI:** `api/v4.html` — served for all routes except `/v3`
- **Input data:** `data/input/Planning for Timetable.xlsx`

## Key Rules

- `api/` folder structure must not change — Vercel requires `api/index.py` at that path
- Algorithm version is V6. Do not revert constraint logic without checking `MASTER_PLAN.md` first
- Before implementing any constraint change, check if it is blocked by a pending Jo answer (J1–J7)
- Configurable toggles are at the top of `api/timetable_scheduler.py` (~line 142–155); check their current values before assuming behaviour

## Folder Structure

```
api/          ← Vercel backend (DO NOT reorganise)
data/input/   ← Excel input files
data/output/  ← Generated timetable outputs (gitignored)
db/           ← SQLite schema reference
docs/
  algorithm/  ← scheduling-algorithm.html (V6 docs)
  meetings/   ← meeting notes and decision records
  plans/      ← dated implementation plans and checklists
  questions/  ← questions-for-jo tracker
scripts/      ← utility scripts (generate_meeting_doc etc.)
MASTER_PLAN.md  ← start here every session
```

## Conventions

- Plan files: `docs/plans/YYYY-MM-DD-NNN-<type>-<name>-plan.md`
- Checklist files: `docs/plans/YYYY-MM-DD-checklist.md`
- Feature branches: `feat/<short-description>`
- Commit style: conventional commits (`feat:`, `fix:`, `chore:`, `docs:`)
