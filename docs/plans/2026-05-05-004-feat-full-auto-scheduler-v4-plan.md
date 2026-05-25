---
title: "feat: Full auto-scheduler v4 — auto-assign Day/Time/Room when Class list answer is empty"
type: feat
status: active
date: 2026-05-05
---

# Full Auto-Scheduler v4

## Problem Statement

**Current system (v3 — Assumption A):**
- Reads Day / Time / Room from "Class list answer" (manually pre-filled by admin)
- Only auto-assigns Lec1/2/3
- 如果 Class list answer 係空嘅，系統無法排班

**New capability (v4):**
- "Class list answer" 可以係空的
- 系統根據 Class list + Teacher list + Teacher Availability + Centre Room Allocation 自動決定：
  - 每個 class 上哪一天（Day）
  - 上什麼時間（Time slot）
  - 用哪個課室（Room）
  - 由哪位老師教（Lec1/2/3）

---

## 現有數據（不需要改動 Excel 結構）

| Sheet | 提供什麼 |
|---|---|
| Class list | class code, subject, loading hours, student count |
| Teacher load table with subject | 老師 → 科目 → quota |
| Teacher Availability | 老師唔得閒的 time slots（N = 唔得閒）|
| Centre Room Allocation | 課室 code, 座位數 |
| Class list answer | **目標：由系統自動填寫** |

---

## 關鍵推論

### 1. 課時結構（Sessions per class）

Loading = 4 小時 → 一個 session = 2 個連續 2h slots（同一天）：
- Morning block：`0900 - 1100` + `1100 - 1300`
- Afternoon block：`1400 - 1600` + `1600 - 1800`

Loading = 2 小時 → 一個 session = 1 個 2h slot

### 2. 中心（Centre）推論

Class code 格式：`DAE101_TW2` → group = `TW2` → centre = `TW`

| Group prefix | 優先課室中心 | 課室前綴 |
|---|---|---|
| TW | TW | TW - TW* |
| KT | KT | KT - KT* |
| TM | TM | TM - TM* |
| ST | ST | ST - ST* |
| TS | TS | TS - TS* |
| TKO | TKO | TKO - TKO* |
| CS | CSW（警察少年之家）| CSW - C* |

**Fallback 規則（適用所有 group）：**
若本中心課室在該時段已被其他班佔用，使用任何有足夠座位的空置課室。
CS 在現有數據出現 WT - WT 正是因為 CSW 課室被別人用了。

### 3. 課室選擇

- 同 group 同一個課室（same group → same room）
- 課室容量 ≥ 學生人數
- 優先較大的課室（現有邏輯）

### 4. Day 優先順序（現有設定）

Monday > Tuesday > Thursday > Wednesday > Friday

---

## 算法設計

### Phase 0：偵測模式

```python
n_sched = count rows in "Class list answer" with Day filled
if n_sched == 0:
    run v4 full auto-assign
else:
    run v3 (existing logic — read fixed assignments)
```

### Phase 1：讀取資料（現有）

與 v3 相同：load subjects, rooms, classes, teachers, availability

### Phase 2（新）：自動分配 Day / Time / Room

```
for each class in order (sort by: student_count DESC, class_code ASC):
    centre = derive_centre(class_code)
    room   = pick_room(centre, student_count, already_assigned_rooms)
    slots  = required_slots(loading_hours)  # 1 slot or 2 consecutive

    for each preferred_day in [Mon, Tue, Thu, Wed, Fri]:
        for each block in [Morning (0900+1100), Afternoon (1400+1600)]:
            if room is free at this day+block:
              AND at least one eligible teacher exists for subject at this day+block:
                assign(class, day, block, room)
                break

    if no slot found:
        add to unassigned
```

### Phase 3：分配老師（現有 v3 邏輯）

與現在完全相同。

---

## 實施步驟

### Step 1：`derive_room_from_code()` 函數

Group prefix → room centre 的對應：大部分直接一致，只有一個例外：

```python
# CS groups → CSW centre（長沙灣分校，room code 用 CSW 前綴）
_GROUP_CENTRE_ALIAS = {"CS": "CSW"}

def group_to_centre(group_code: str) -> str:
    prefix = re.sub(r'\d+$', '', group_code)  # "TW2" → "TW"
    return _GROUP_CENTRE_ALIAS.get(prefix, prefix)

def derive_preferred_room(class_code, student_count, conn) -> Optional[str]:
    """Try group's home centre first; fallback to any room with enough seats."""
    group = _extract_group(class_code)              # "TW2"
    centre = group_to_centre(group)                 # "TW"
    room_num = re.search(r'\d+$', group)            # "2"
    if room_num:
        candidate = f"{centre} - {centre}{room_num.group()}"  # "TW - TW2"
        row = conn.execute(
            "SELECT code FROM rooms WHERE code = ?", (candidate,)).fetchone()
        if row:
            return row[0]
    # Fallback: largest room at home centre with enough capacity
    row = conn.execute(
        "SELECT code FROM rooms WHERE centre = ? AND capacity >= ? ORDER BY capacity DESC LIMIT 1",
        (centre, student_count)).fetchone()
    return row[0] if row else None
```

### Step 2：`auto_assign_schedule()` 函數

新增函數，替代 `_load_existing_schedule()`。

### Step 3：`build_db()` 修改

```python
n_sched = _load_existing_schedule(conn, wb)
if n_sched == 0:
    n_sched = auto_assign_schedule(conn, wb)  # v4 path
```

### Step 4：`collect_results()` / output 不變

Phase 3（老師分配）和 output 邏輯完全不需要改動。

---

## Edge Cases / 已知限制

| 情況 | 處理方式 |
|---|---|
| CS groups 課室映射不明確 | 需要確認 CS → CSW 或 WT 的規則 |
| 課室已滿（所有 day+time 都有課）| 加入 unassigned list，顯示提示 |
| 老師全部唔得閒某個時段 | 找不到老師 → 先排房間，老師 = 空，顯示 warning |
| Loading = 6h 或其他奇數 | 目前只支援 2h 和 4h，需要釐清業務邏輯 |
| 同一科不同 group 是否同一天？ | 目前不強制，可加 preference |

---

## 驗收標準

- [ ] 上載只有 "Teacher Availability" 填寫的 Excel → 系統自動排出完整時間表
- [ ] 每個 class 都有 Day + Time + Room + Lec1（或顯示 unscheduled 原因）
- [ ] Day 優先順序正確（Mon > Tue > Thu > Wed > Fri）
- [ ] 同一 group 同一課室
- [ ] 老師不 double-booked
- [ ] 現有 Excel（Class list answer 已填）仍然走 v3 邏輯，結果不變

---

## 工作量估計

| Component | 複雜度 |
|---|---|
| `derive_room_from_code()` | Low — pattern matching |
| `auto_assign_schedule()` | Medium — nested loop with 2 constraints |
| CS group mapping clarification | Need user input |
| Test with real Excel | Low — reuse existing test |

**總計：半天工作量**（不計 CS group 確認時間）
