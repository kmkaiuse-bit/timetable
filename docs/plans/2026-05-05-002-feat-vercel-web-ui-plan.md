---
title: Vercel Web UI for Timetable Scheduler
type: feat
status: completed
date: 2026-05-05
origin: docs/plans/2026-05-05-001-feat-teacher-timetable-scheduler-plan.md
---

# Vercel Web UI for Timetable Scheduler

## Overview

Wrap the existing `timetable_scheduler.py` in a browser-accessible web app deployable via GitHub → Vercel. Teachers upload their Excel file, click one button, see a validation report, and download the output — no Python, no command line.

```
GitHub push → Vercel auto-deploy → https://your-app.vercel.app
```

---

## Architecture

```
09_timetable/
├── api/
│   └── index.py               # Flask app (2 routes: GET / and POST /api/schedule)
├── public/
│   └── index.html             # Single-page UI (vanilla JS, no framework)
├── timetable_scheduler.py     # MODIFIED — add run_from_bytes() for in-memory use
├── requirements.txt           # MODIFIED — add flask>=3.0
└── vercel.json                # NEW — routing config
```

**Data flow (per request):**
```
Browser uploads Excel bytes
  → POST /api/schedule
    → run_from_bytes(bytes)         # scheduler runs in ~2s
      → returns (output_bytes, stats_dict)
    → JSON response: {file: base64, stats: {...}}
  → Browser shows stats
  → Browser creates blob download link
```

No filesystem writes, no state between requests — fully stateless serverless.

---

## Implementation Phases

### Phase 1 — Refactor `timetable_scheduler.py` for BytesIO

**Goal:** Add `run_from_bytes(excel_bytes: bytes) -> tuple[bytes, dict]` so the scheduler can be called without touching the filesystem.

**Changes:**

```python
# NEW function — BytesIO-based entry point for web use
def run_from_bytes(excel_bytes: bytes) -> tuple[bytes, dict]:
    """
    Accepts raw Excel file bytes.
    Returns (output_excel_bytes, stats_dict).
    """
    from io import BytesIO

    # Read with data_only for values
    wb_read = openpyxl.load_workbook(BytesIO(excel_bytes), data_only=True)
    classes  = read_classes(wb_read)
    teachers = read_teachers(wb_read)
    rooms    = read_rooms(wb_read)

    assignments, _ = assign_teachers(classes, teachers)
    scheduled, unscheduled = schedule_classes(classes, assignments, rooms)

    # Write output to memory (template = fresh load of same bytes)
    wb_out = openpyxl.load_workbook(BytesIO(excel_bytes))
    write_output_wb(wb_out, scheduled)   # refactor write_output to accept wb + scheduled only
    buf = BytesIO()
    wb_out.save(buf)
    output_bytes = buf.getvalue()

    stats = collect_stats(scheduled, rooms, unscheduled)
    return output_bytes, stats

# NEW function — collect stats as dict (extracted from validate_and_report)
def collect_stats(scheduled, rooms, unscheduled) -> dict:
    room_cap = {r.code: r.capacity for r in rooms}
    violations = []
    teacher_slots = defaultdict(list)
    room_slots    = defaultdict(list)

    for sc in scheduled:
        slots = [sc.slot1] + ([sc.slot2] if sc.slot2 else [])
        for s in slots:
            teacher_slots[sc.lecturer1].append((sc.day, s, sc.entry.code))
            room_slots[sc.room].append((sc.day, s, sc.entry.code))
        cap = room_cap.get(sc.room, 0)
        if cap < sc.entry.students:
            violations.append(f"Overflow: {sc.entry.code} in {sc.room}")

    for teacher, entries in teacher_slots.items():
        seen = {}
        for day, slot, code in entries:
            k = (day, slot)
            if k in seen:
                violations.append(f"Teacher clash: {teacher} — {seen[k]} & {code} on {day} {slot}")
            else:
                seen[k] = code

    for room, entries in room_slots.items():
        seen = {}
        for day, slot, code in entries:
            k = (day, slot)
            if k in seen:
                violations.append(f"Room clash: {room} — {seen[k]} & {code} on {day} {slot}")
            else:
                seen[k] = code

    pref  = sum(1 for sc in scheduled if sc.day in PREFERRED_DAYS)
    total = len(scheduled)

    centre_dist = defaultdict(lambda: defaultdict(int))
    for sc in scheduled:
        centre_dist[sc.entry.centre][sc.day] += 1

    return {
        "scheduled":      total,
        "total_classes":  total + len(unscheduled),
        "unscheduled":    unscheduled,
        "violations":     violations,
        "preferred_pct":  round(pref / total * 100) if total else 0,
        "centre_dist":    {c: dict(d) for c, d in centre_dist.items()},
    }
```

Also refactor `write_output` to split into:
- `write_output(input_path, output_path, scheduled)` — existing CLI function (unchanged)
- `write_output_wb(wb, scheduled)` — operates on an already-loaded workbook (used by `run_from_bytes`)

**Deliverable:** `run_from_bytes()` callable from any Python context, including Flask.

---

### Phase 2 — Flask API (`api/index.py`)

**Two routes only:**

```python
# api/index.py
import base64, os, sys
from flask import Flask, request, jsonify, send_from_directory
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from timetable_scheduler import run_from_bytes

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), '..', 'public'))

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_ui(path):
    if path and os.path.exists(os.path.join(app.static_folder, path)):
        return send_from_directory(app.static_folder, path)
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/api/schedule', methods=['POST'])
def schedule():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'No file uploaded'}), 400

    filename = f.filename or ''
    if not filename.endswith('.xlsx'):
        return jsonify({'error': 'Please upload an .xlsx file'}), 400

    try:
        excel_bytes = f.read()
        output_bytes, stats = run_from_bytes(excel_bytes)
        return jsonify({
            'stats': stats,
            'file':  base64.b64encode(output_bytes).decode('utf-8'),
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
```

**Deliverable:** `POST /api/schedule` accepts multipart Excel upload, returns JSON with stats + base64 output.

---

### Phase 3 — Web UI (`public/index.html`)

Single HTML file. No build step, no npm, no framework.

**Layout:**
```
┌─────────────────────────────────────────┐
│  課程時間表排班系統                        │
│  Timetable Scheduler                    │
├─────────────────────────────────────────┤
│  ┌───────────────────────────────────┐  │
│  │  拖放 Excel 檔案至此                │  │
│  │  或按此選擇檔案                      │  │
│  │  [Planning for Timetable.xlsx]     │  │
│  └───────────────────────────────────┘  │
│                                         │
│        [生成時間表 Generate]             │
├─────────────────────────────────────────┤
│  ✓ 已排班: 112 / 112                    │  ← results panel
│  ✓ 衝突: 0                             │    (hidden until done)
│  ✓ 優先日期使用率: 99%                   │
│                                         │
│  Centre  Mon  Tue  Thu  ...             │
│  CS      16    8    1                   │
│  WT       8    0    0                   │
│  ...                                    │
│                                         │
│        [下載 Timetable_Output.xlsx]     │
└─────────────────────────────────────────┘
```

**Key JS logic:**

```javascript
// Form submit → fetch POST → show stats → create blob download
async function generate() {
    const file = document.getElementById('fileInput').files[0];
    const fd   = new FormData();
    fd.append('file', file);

    showLoading(true);
    const res  = await fetch('/api/schedule', { method: 'POST', body: fd });
    const data = await res.json();
    showLoading(false);

    if (data.error) { showError(data.error); return; }

    showStats(data.stats);

    // Create download from base64
    const bytes  = Uint8Array.from(atob(data.file), c => c.charCodeAt(0));
    const blob   = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url    = URL.createObjectURL(blob);
    const anchor = document.getElementById('downloadBtn');
    anchor.href     = url;
    anchor.download = 'Timetable_Output.xlsx';
    anchor.style.display = 'block';
}
```

**Deliverable:** Clean, functional single-page UI. Works on Chrome/Edge desktop. No external CSS/JS dependencies (everything inline).

---

### Phase 4 — Vercel Config + requirements.txt

**`vercel.json`:**
```json
{
  "rewrites": [
    { "source": "/api/(.*)", "destination": "/api/index" },
    { "source": "/(.*)",     "destination": "/api/index" }
  ],
  "functions": {
    "api/index.py": { "maxDuration": 30 }
  }
}
```

**`requirements.txt`** (root level):
```
openpyxl>=3.1.0
flask>=3.0
```

**Deliverable:** Vercel can build and route correctly on first deploy.

---

### Phase 5 — GitHub init + Vercel deploy

```bash
git init
git add .
git commit -m "feat: timetable scheduler with Vercel web UI"
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

Then on vercel.com: Import repo → Framework Preset: Other → Deploy.

**Deliverable:** Live URL accessible from browser.

---

## Acceptance Criteria

- [ ] `python timetable_scheduler.py` still works (CLI unchanged)
- [ ] `run_from_bytes(bytes)` returns valid output bytes + stats dict without touching filesystem
- [ ] `POST /api/schedule` with a valid `.xlsx` returns 200 JSON `{stats, file}`
- [ ] `POST /api/schedule` with no file returns 400 with `{error}`
- [ ] Browser page loads at `/`
- [ ] Uploading `Planning for Timetable.xlsx` → Generate → shows 112/112 scheduled, 0 violations
- [ ] Download button produces a valid `.xlsx` that opens in Excel
- [ ] Vercel deploy succeeds from GitHub push (no manual build steps)
- [ ] Page works on Chrome and Edge (desktop)

---

## Dependencies & Risks

| Item | Detail |
|------|--------|
| `flask>=3.0` | Adds ~2 MB to the function bundle. Well within Vercel's 250 MB limit. |
| Vercel 10s timeout (Hobby) | Scheduler runs <2s — safe. Set `maxDuration: 30` as buffer. |
| Base64 payload size | Output Excel ~70 KB → base64 ~95 KB → well within Vercel's 4.5 MB response limit. |
| `sys.path.insert` in `api/index.py` | Needed to import `timetable_scheduler` from parent dir. Standard pattern for Vercel Python. |
| `write_output` refactor | Must not break the existing CLI `main()` function. Keep both variants. |
| openpyxl BytesIO double-load | Template must be loaded from bytes a **second time** (fresh `BytesIO`) because `write_output_wb` mutates the workbook. The read pass (data_only=True) is separate. |

---

## Future Improvements

| Item | Notes |
|------|-------|
| **Lecturer 2 / 3 auto-assignment** | Rules unclear — needs clarification on whether it's by class size, subject type, or fixed teacher pairing. Kept as manual step for now. UI shows explicit reminder to fill after download. |
| **Non-standard start times** | Some classes use 8:30, 10:30, 13:00 slots. MVP uses standard slots only. Could add a configurable time-slot list. |
| **Teacher-centre affinity** | Currently teachers are assigned round-robin. Could improve by preferring teachers who historically teach at a specific centre. |

## Sources & References

- Origin plan: `docs/plans/2026-05-05-001-feat-teacher-timetable-scheduler-plan.md`
- Existing scheduler: `timetable_scheduler.py` (functions: `run_from_bytes`, `write_output_wb` to be added)
- Vercel Python runtime: `functions` + `rewrites` config (2026 recommended over legacy `builds`/`routes`)
- Flask `send_from_directory` for static HTML serving on serverless
- openpyxl BytesIO: `load_workbook(BytesIO(b))` and `wb.save(BytesIO())` — no temp files needed
