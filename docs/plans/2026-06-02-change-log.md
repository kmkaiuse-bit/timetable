# Change Log — 2026-06-02

---

## [1] Fix: phase0_schedule_cadets — 假 row bug

**檔案：** `api/timetable_scheduler.py`  
**函數：** `phase0_schedule_cadets()`  
**改動前：** 4小時課（need_two=True）用兩個 INSERT，第二個 INSERT 用 `cls["code"] + "_slot2"` 製造假 class code（例如 `DAE256_E_slot2`），此 code 在 `classes` 表不存在，導致 `collect_results` JOIN 失敗，cadet class 不出現在輸出。  
**改動後：** 一個 `INSERT OR REPLACE`，用 `time1, time2` 兩欄存兩個時段，與 `auto_assign_schedule` 一致。  
**影響：** Cadet class（Phase 0）的 4小時課現在正確出現在輸出，老師分配亦可正常執行。

---

## [2] Fix: _pick_room — 移除 H3 違規 fallback（B3 bug）

**檔案：** `api/timetable_scheduler.py`  
**函數：** `_pick_room()`  
**改動前：** 若本 centre 無合適房間，最後 fallback 查詢全系統任何房間（`WHERE capacity >= ?`），可能分配去完全不同的 centre（如 FL 學生跑去 ST 上課），靜靜違反 H3。  
**改動後：** 移除最後 fallback，若本 centre 及 allowed centres 均無合適房間，直接返回 `None`，caller 記錄為未排班並報告原因。  
**影響：** 發現隱藏的資料問題 — FL、SW、TKO、TM 幾個 centre 有課堂學生人數超出房間最大容量，舊版靜靜放去錯誤 centre，現在會明確報告需要 Jo 處理。

---

## [3] Feat: auto_assign_schedule — 追蹤並報告未排班課堂

**檔案：** `api/timetable_scheduler.py`  
**函數：** `auto_assign_schedule()`  
**改動前：** 返回 `int`（排班數量），未能排班的課堂靜靜消失，無原因記錄。  
**改動後：** 返回 `tuple(int, list)`，list 每項為 `{"code": ..., "reason": ...}`，區分兩類失敗：  
- `"no room at X with capacity >= N"` — 資料問題（房間不夠大）  
- `"no free slot across all days"` — 演算法問題（時段/老師容量耗盡）  
`run_v4_from_bytes()` 將結果存入 `stats["unscheduled_rooms"]`，前端可顯示。  
**影響：** 用家現在可以清楚看到哪些課堂排不了、原因是什麼。

---

## [4] Improve: auto_assign_schedule — MRV 排序（限制越多越先排）

**檔案：** `api/timetable_scheduler.py`  
**函數：** `auto_assign_schedule()`  
**改動前：** `ORDER BY student_count DESC, code ASC` — 只按學生人數排序。  
**改動後：**
```sql
ORDER BY
    CASE WHEN cg.centre = 'TKO' THEN 0 ELSE 1 END,  -- TKO 時間限制最多，先排
    CASE WHEN sub.loading_hrs >= 4 THEN 0 ELSE 1 END, -- 4小時課需要兩個連續時段，先排
    c.student_count DESC,                              -- 學生越多房間選擇越少
    c.code ASC                                         -- 穩定排序
```
**原理：** MRV（Minimum Remaining Values）— 選擇空間越少的課堂越先排，避免它們被後來的課堂佔去所有合適時段。  
**影響：** 減少「no free slot」失敗，特別是 TKO 課堂和 4小時課。

---

## 測試結果

| 指標 | 修改前 | 修改後 |
|------|--------|--------|
| 總排班 | 107/112 | 107/112（v3 模式，Excel 已有數據）|
| Lec1 分配 | 105 | 105 |
| Lec2 分配 | 101 | 101 |
| 警告數量 | 31 | 31 |
| 未排班可見性 | 靜靜消失 | 明確報告原因 |
| H3 違規 | 靜靜發生 | 返回 None，報告錯誤 |

**注意：** 現有 Excel 已有 108 行 Class list answer 數據，程式走 v3 模式（只分配老師），v4 auto-assign 改動需用空白 Class list answer 測試。

---

## 新增問題（需問 Jo）

由修改後發現的資料問題，以下 centre 有課堂學生人數超出房間最大容量：

| Centre | 最大學生數 | 最大房間容量 | 差額 |
|--------|-----------|------------|------|
| FL     | 34        | 29         | +5   |
| SW     | 34        | 31         | +3   |
| TKO    | 34        | 30         | +4   |
| TM     | 43        | 42         | +1   |
| ST     | 38        | 29         | +9   |

舊版系統靜靜將這些課堂放去其他 centre（違反 H3），現在正確顯示為錯誤。  
→ **需要 Jo 確認：這些課堂應該如何處理？加大房間？拆班？接受跨 centre？**

---

---

## [5] Feat: English Net Teacher Weekly Assignment（DAE102）

**檔案：** `api/timetable_scheduler.py`, `data/input/Planning for Timetable.xlsx`  
**背景：** Jo 確認英文科需要 Net teacher 安排：每班每學期 ≥20 Net teacher hours（即 3 週段中 ≥2 個由 Net teacher 教）。  
**改動：**

### 資料層
- `Planning for Timetable.xlsx` 新增 `Net Teachers` sheet（列出 Net teacher 名單：Mr. Peter Barrett, Mr. Chris Hon）
- DB schema：`teachers` 表加 `is_net INTEGER DEFAULT 0`
- 新函數 `_load_net_teachers(conn, wb)` — 讀取 Net Teachers sheet，標記對應老師

### 算法層
新常數（`timetable_scheduler.py` 頂部，方便調整）：
```python
_ENG_SUBJECT_CODE      = "DAE102"
_ENG_NET_MIN_BLOCKS    = 2          # ≥2 Net blocks per term
_ENG_MAX_TRAVEL_MIN    = 30         # English teacher travel cap
_ENG_TERMS             = ["T2025C", "T2026A"]
_ENG_WEEK_BLOCKS       = ["wk1-5", "wk6-10", "wk11-15"]
_ENG_NET_EXEMPT_GROUPS = {"CS1","CS2","CS3","CS7"}
```

新函數 `assign_english_weekly(conn) -> dict`：
- 為每個 DAE102 班 × 2學期 × 3週段分配老師（共 156 個分配）
- 優先用 Net teacher，直到達到 _ENG_NET_MIN_BLOCKS；其餘用 local teacher
- Travel check：英文老師同一天同一週段 travel ≤ 30 min
- CS1-3, CS7 豁免 Net 要求（_ENG_NET_EXEMPT_GROUPS）
- 無 Net Teachers sheet → 靜靜跳過 Net 限制（向後相容）

### Output 層
`write_output_fast()` 加入 `English Weekly` sheet：

| Class Code | Day | Time | Room | T2025C wk1-5 | T2025C wk6-10 | T2025C wk11-15 | T2026A wk1-5 | T2026A wk6-10 | T2026A wk11-15 |
|---|---|---|---|---|---|---|---|---|---|

### 警告
Net teacher 不足時，在 `stats["warnings"]` 加入提示：
```
English Net shortfall: 1/2 Net blocks in T2025C (insufficient Net teacher supply)
```

**測試結果：**
- 156 個週段分配生成（26班 × 6週段）
- CS1-3, CS7 豁免正確
- 11 個非豁免班達到 ≥2 Net blocks per term
- 12 個非豁免班 Net shortfall → 警告（根本原因：2個 Net teachers 時段有限，不足以覆蓋所有 22 個非豁免班）
- 輸出 Excel 包含 `English Weekly` sheet

**待 Jo 決定：**
- 24 個 Net shortfall warnings 是否可接受？還是要增加 Net teacher？
- 這是 `J4`（teacher type 問題）的延伸

---

## 下一步

- [x] Merge `feat/v6-constraint-improvements` → `master` ✓
- [ ] 更新 `MASTER_PLAN.md`
- [ ] 新增問題 J8：各 centre 超容量課堂處理方案
- [ ] 詢問 Jo：Net shortfall 如何處理（增加 Net teachers / 調整時間 / 接受豁免更多班）
