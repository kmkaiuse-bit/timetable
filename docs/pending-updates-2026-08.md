---
title: 待處理系統 Update（Jo 開會，2026-08）
status: 待實作（另一 session）
---

# 三個系統 Update（2026-08）

> 給新 session：實作以下 3 項。每項附現有狀態、要改的位置、及要注意的陷阱。
> 核心檔案：`api/timetable_scheduler.py`（排程引擎）、`api/v4.html`（UI）。

---

## Update 1：DAE 科目不排星期三、五（只用 Mon/Tue/Thu）

**要求：** 目前所有 DAE 科目（DAE101–108，即一般 Core 班）**星期三、星期五都不要排班**，只用星期一、二、四。

**現狀：**
- `_DAY_PRIORITY = ["Monday", "Tuesday", "Thursday", "Wednesday", "Friday"]`（scheduler.py:120）—— Phase A 一般班用足 5 日。

**要改：**
- 把一般班的可排日子改為 **只 Mon/Tue/Thu**（改 `_DAY_PRIORITY` 或在 Phase A 過濾 Wed/Fri）。

**⚠️ 陷阱（必看）：**
1. **Cadet（學警）班固定在 Wed/Fri，必須保留。** `_CADET_DAYS = {"Monday","Wednesday","Friday"}`（scheduler.py:1998），由 Phase 0 處理，跟 `_DAY_PRIORITY` 分開。改 `_DAY_PRIORITY` 只影響 Phase A，理論上不動 cadet —— **但要驗證 cadet Phase 0 不依賴 `_DAY_PRIORITY`**。
   - 另注意：`_CADET_DAYS` 現含 Monday，但 Jo 說 cadet 是「星期三、五」。此不一致要向 Jo 確認。
2. **容量大減：** 由 5 日縮到 3 日 = 少 40% 時段。TS/FL/TM 本已課室不足（見 [timetable-constraints](timetable-constraints.md) 及 issues），改完後**未排到的班會大增**。實作後要重新評估、並提醒用戶此影響。
3. CC Combine 的 `_find_cc_day`、S3 輪流分佈等亦用 `_DAY_PRIORITY`，改動會連帶影響 —— 一併檢查。

---

## Update 2：教師 Loading 可超過 6，只在多於 6 時提醒

**要求：** 教師每週 loading 不一定上限 6，可以加；只有**多於 6 時要提醒**（不阻擋）。

**現狀（可能已達成，先驗證）：**
- `_TEACHER_WEEKLY_SESSION_CAP = 6`（scheduler.py:144）。
- H4 在 V6 已改為**軟上限**：超過 6 只發警告（`H4 (soft): teacher exceeds 6 sessions`），不阻擋排班。

**要做：**
1. 先驗證現有行為是否已符合「可超過、只提醒」—— 很可能已達成，只需確認。
2. 若 Jo 想**每位教師有自己的上限**（而非全域 6），則改為從 input Excel 讀每人 loading，警告門檻按個人設定。（限制文件已記「以 Jo Excel 填的 loading 為準」。）

---

## Update 3：UI 時間表格放大 ＋ 格中顯示 subject code

**要求：** 時間表要**大一點**，每格中要**顯示 subject code**。

**現狀：**
- `api/v4.html` 的 `showDayGrid`（約 line 943）每格顯示：class code（如 `DAE101_CS1`）、時間、subject_cn、lec1、人數。
- 週·課室、週·教師檢視的格（`renderWeekOverview` / `renderTeacherWeek`）較細，只顯示 class 短碼 ＋ 老師。

**要做：**
1. **放大格：** 調 CSS —— `.tt-table`、`.week-table` 的 font-size、儲存格 padding、min-width、`.cell-*` 尺寸。
2. **顯示 subject code：** 在格中明確加上科目代號（如 `DAE101`）。注意 class code 已含前綴（`DAE101_CS1`），可考慮把 subject code 獨立一行或加粗，令它更顯眼。
3. 四個檢視（日／週·課室／週·教師／月）都要一致處理。

---

## 建議次序
1. **Update 3（UI）** —— 純前端、風險最低、即見效果。
2. **Update 2** —— 多數已達成，先驗證。
3. **Update 1** —— 影響最大（容量減、連帶 cadet／CC／S3），最後做並充分測試。

*相關背景：`docs/timetable-constraints.md`（Jo 確認的限制）、`docs/PRD-timetable-scheduler.md`。系統已部署在 master。*
