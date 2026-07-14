# 2026-07-08-001 — 英文老師自動排班 + 資料修復計劃

**狀態:核心完成 (2026-07-08)。** clean input 結果:0 英文可用衝突(Jo 手排 12)、0 Net shortfall、26 個 DAE102 班全在 Mon/Tue/Thu。剩 Phase 3.4(template 同步)、Phase 6.3(commit)。

**目標:** 令整個 timetable app 合理運行 — Jo 自己上傳一份 input Excel,系統**全自動**排出 Day/Time/Room + 老師(含英文科),結果質素對標 Jo 手排版。

**方向(2026-07-08 對齊):**
- 用家 = Jo 本人用網頁,全自動 V4
- `English Teacher Arrangement.xlsx` = 約束來源 + 手排基準,**非照抄定案**
- 修復策略 = **資料 + 程式兩邊都修**

---

## Phase 1 審計實況(2026-07-08 完成)

真實 input `data/input/Planning for Timetable.xlsx` 已有齊英文老師資料,但**資料壞掉**,這才是 app 跑不合理的根因:

| # | 問題 | 後果 |
|---|------|------|
| **A** | `Teacher Availability` sheet 8 位英文老師全空白 = 系統當佢地 Mon–Fri 全可用;真實只 Mon/Tue/Thu | 自動排會排去 Wed/Fri,錯 |
| **B** | `Teacher load table` 用 `Mr. Cherry Ip`(quota 6)、`Mr. Lee Kit Wan`(quota 2);arrangement/Net/English Weekly 用 `Ms.` | 跨表 join 斷裂,同人拆兩個 record,Net/可用檢查靜默失效 |
| **C** | `English Weekly` sheet 已預填 Jo 答案 | 觸發 preassigned 模式,自動排從未執行 |
| D | `_ENG_WEEKLY_PREASSIGNED` 150 行硬編碼(抄 Jo 答案)| 走錯方向,應移除 |
| E | `_ENG_NET_MIN_BLOCKS = 2` 應為 1(1 block=20 小時,檔案 Hours=課數×20 證實)| 44 假 shortfall |

**資料鏈參考:**
- 資格+quota:`_load_teachers`(scheduler.py:522)按全名 join,numbered row=lec1,兩區塊 quota 相加
- 可用時間:`_load_availability`(:564)讀 `Teacher Availability`,格式 `Mon 0900` 等 4 時段×5 日,`N`=不可用
- Net:`_load_net_teachers`(:660)讀 `Net Teachers` sheet 標 is_net
- 英文分配:`assign_english_weekly`(:1538)Mode1 preassigned / Mode2 auto

---

## Phase 2 — 程式容錯(P0)

- [ ] 2.1 姓名正規化:新增 `_normalize_teacher_name()`,join 時去掉 `Mr./Ms./Mrs./Dr.` 前綴做 match key(顯示仍保留原名);令 `Mr. Cherry Ip` 與 `Ms. Cherry Ip` 視為同一人
- [ ] 2.2 `English Weekly` 預填時發**明確警告**(Issues panel):「偵測到英文預填分配,已覆寫自動排;清空此 sheet 以啟用全自動」— 不再靜默
- [ ] 2.3 移除 `_ENG_WEEKLY_PREASSIGNED`(:159);Mode 2 auto 成為預設路徑
- [ ] 2.4 `_ENG_NET_MIN_BLOCKS` 2 → 1;記入 MASTER_PLAN toggles,J9 標 Assumed

## Phase 3 — 資料修復:產出正確測試 input(P0)

寫 script `scripts/build_clean_input.py`,由現有 input + arrangement 檔生成 `data/input/Planning for Timetable (clean).xlsx`:

- [ ] 3.1 `Teacher Availability`:按 arrangement 31–40 行填 8 英文老師可用時間(AM=0900+1100,PM=1400+1600,其餘 slot 標 `N`)
  - Lee Kit Wan: Thu AM · Elise Ye: Tue PM+Thu PM · Barrett: Mon/Tue/Thu 全 · Cherry Ip: Mon/Tue/Thu 全 · Ray Leung: Tue 全 · Sasha Cheung: Mon/Tue/Thu 全 · Ivan Yuen: Mon/Thu 全 · Chris Hon: Mon/Tue/Thu 全
- [ ] 3.2 姓名統一:`Mr. Cherry Ip`→`Ms. Cherry Ip`、`Mr. Lee Kit Wan`→`Ms. Lee Kit Wan`(全 sheet)
- [ ] 3.3 清空 `English Weekly` 資料列(保留 header)
- [ ] 3.4 template_v4.xlsx 同步:加 `Net Teachers` sheet(現無)、可用時間填寫範例、附說明

## Phase 4 — DAE102 自動排協調(P0)

- [ ] 4.1 確認自動排 DAE102 時 Day 候選受老師可用日限制(資料驅動,不寫死)
- [ ] 4.2 同時段 DAE102 班數 ≤ 該時段可用英文老師數
- [ ] 4.3 驗證 27 班全排到 + 全配到老師

## Phase 5 — 基準對照(驗收核心)

- [ ] 5.1 script 讀 arrangement as 基準,對比系統輸出
- [ ] 5.2 指標:可用衝突(Jo 12 → 系統 **0**)、Net 達標 23/23、未排班 0、負載極差 ≤ Jo(680-160)
- [ ] 5.3 對照報告 = demo 給 Jo 的證據

## Phase 6 — 端到端 + 收尾

- [ ] 6.1 clean input 跑 `run_v4_from_bytes`,UI 驗證(grid / Issues / English Weekly 輸出)
- [ ] 6.2 其他科目無回歸
- [ ] 6.3 更新 MASTER_PLAN;刪 `_temp_read.py`;整理未提交檔;conventional commits

---

## 待 Jo 確認(不阻擋)

| # | 問題 | 假設 |
|---|------|------|
| J9 | 1 block = 20 Net 小時,每學期 1 block 達標? | 是(門檻 1)|
| 新 | Cherry Ip / Lee Kit Wan 正確稱謂 Ms.? | 是 |
| 新 | 英文老師可用時間以 arrangement 31–40 行為準? | 是 |

## 風險
- 只 2 個 Net 老師,若 Jo 要 2 blocks/學期,數學上不可行 → Issues panel 需附證據
- 姓名正規化去前綴後若有真正同姓不同人會誤併 → 目前資料無此情況,但 log 出被併的名
- 系統自動排結果會與 Jo 手排不同 → Phase 5 對照報告化解
