# 2026-05-26 Tomorrow's Checklist

## 1. 發問題給 Jo（優先）

> 文件：`docs/questions/questions-for-jo.docx`

- [ ] **J1** — H3：TS→TM 例外，這學期適用還是 next year？
- [ ] **J2** — CC：CC Combine 兩個 centre 之間最大 travel time 是多少分鐘？
- [ ] **J3** — S3：學警班課室代號是什麼？Excel 裡怎麼標記？（現在資料顯示 CSW-C3，Mon/Wed/Fri）
- [ ] **J4** — 老師：English / Local / Net teacher 排班有分別嗎？
- [ ] **J5** — 成本：排班需要考慮老師收費（price）嗎？
- [ ] **J6** — 老師：老師分配名單什麼時候可以提供？
- [ ] **J7** — CC：CC grouping 完整名單什麼時候可以提供？

---

## 2. Code

- [ ] Merge `feat/v6-constraint-improvements` → `master`
- [ ] 用真實 Excel（`data/input/Planning for Timetable.xlsx`）跑一次 scheduler，確認沒有 regression

---

## 3. 收到 Jo 答案後立即做

| Jo 答案 | 動作 | 位置 |
|---------|------|------|
| J1 確認今學期 | `_TS_TM_ENABLED = False` → `True` | `api/timetable_scheduler.py` line ~148 |
| J2 確認數字 | `_CC_MAX_TRAVEL_MIN = None` → `X` | `api/timetable_scheduler.py` line ~151 |
| J3 確認課室 | 完成 S3 Phase 0 cadet class | `phase0_schedule_cadets()` in scheduler |
| J6 名單到 | 上傳到 Excel，跑老師自動分配 | `data/input/` |
| J7 名單到 | 填入 CC group 欄，跑 CC Combine | `data/input/` |

---

## 4. 參考

**今天完成（2026-05-25）：**
- V6.1：`_TRAVEL_TIME` matrix、H8 soft gate（≤90 min warn, >90 min block）
- H3 TS→TM toggle（default off）
- S3 Phase 0 cadet class 框架
- CC Combine distance filter（default disabled）
- Folder 重組
- `MASTER_PLAN.md` 建立

**整體進度：** → 見 `MASTER_PLAN.md`
