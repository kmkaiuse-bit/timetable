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

## 下一步

- [ ] Merge `feat/v6-constraint-improvements` → `master`
- [ ] 更新 `MASTER_PLAN.md` 加入 B4/B5 bug 修復記錄
- [ ] 新增問題 J8：各 centre 超容量課堂處理方案
