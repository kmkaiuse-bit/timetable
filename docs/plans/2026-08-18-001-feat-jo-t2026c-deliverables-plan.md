# T2026C 交付計劃 — Jo 的六項任務

**日期:** 2026-08-18
**來源:** Jo 的兩封 email (2026-08-17 / 2026-08-18) + 三份附件
**狀態:** 需求已確認，待實作

---

## 附件清單

| 檔案 | 用途 |
|------|------|
| `T2026C_WeeklyTimetable_V4 0817.xlsx` | 主檔。`Mon(Term1)`…`Sat(Term1)` = 已排好的週時間表；`Classes` = 175 科班清單；`Classroom Code` = 房間對照；`ClassTeacher26-27` = 英文 Net 編排目標 |
| `HKIT DAE Calendar_V1.xlsx` | `DAE Calendar 2026-2027` 欄 AA–AM = 15 個上課日的日期（按 Day + AM/PM） |
| `DailyTimetable_T2025C sample.xlsx` | Task 2 輸出格式範本 |
| `ClassTimetable_T2025C.xlsx` | Task 6 輸出格式範本（每班一個 sheet） |
| `DAE_TeachingSchedule.xlsx` | Task 5 輸出格式範本（每位教師一個 sheet），2026-08-18 收到 |

---

## 六項任務

| # | 任務 | 狀態 |
|---|------|------|
| 1 | `Classes` sheet 填 Day / Time / Location | ✅ 已完成 — 168/175（7 個議定排除） |
| 2 | 匯出 Daily Timetable Excel | ✅ 已完成 — 4,494 行 |
| 3 | 英文 Net teacher 編排 + 工時表 | ✅ 已完成 — 168 個 block 全數指派，五位有預算的教師鐘數完全命中 |
| 4 | 按 Task 3 結果重出 Daily Timetable 的英文行 | ✅ 已完成 — 併入模組 B，每 5 堂換一次 Lecturer |
| 5 | Teacher view timetable（10 位教師） | ✅ 已完成 — 儲存格寫法採 Kevin 式，待 Jo 確認 |
| 6 | Class view timetable（32 班） | ✅ 已完成 |

### 輸出檔（`data/output/`）

| 檔案 | 對應 |
|------|------|
| `Classes_DayTimeLocation_T2026C.xlsx` | Task 1（含 `_Issues`） |
| `DailyTimetable_T2026C.xlsx` | Task 2 + 4 |
| `EnglishNetAssignment_T2026C.xlsx` | Task 3 |
| `DAE_TeachingSchedule_T2026C.xlsx` | Task 5 |
| `ClassTimetable_T2026C.xlsx` | Task 6 |

執行方式：`python scripts/build_t2026c_outputs.py <weekly>.xlsx <calendar>.xlsx`

### 六項需 Jo 覆核的衝突

程式已自動濾走 combine class（同房同教師），剩餘六項無法用 combine 解釋：

| 類別 | 內容 |
|------|------|
| 課室推斷可疑 | 週二 `ST - ST3`：`DAE101_ST1`（Mercury Lee）撞 `DAE102_FL1`（Elise Ye） |
| 課室推斷可疑 | 週二 `TW - TW1`：`DAE106_TW2`（Newman Lo）撞 `DAE102_TW1`（Ray Leung） |
| 教師分身 | 週五 Philip Chan：`DAE205_A/B` @SSP-303 vs `DAE229_A` @SSP-201 |
| 教師分身 | 週三 Philip Chan：`DAE243_C` @SSP-104 vs `DAE243_A` @SSP-201 |
| 學生撞堂 | 週二 CS5：`DAE103_CS5` 1600-1800 撞 `DAE108_CS5` 1400-1800 |

前兩項源自空白課室格向上填充，兩格原本都無底色，缺乏分組證據，最可能是 Jo 漏填課室。

### Task 3 的結構性發現

28 個英文班只落在 **6 個不同的 day+time 時段**上。Loveleen 的 36 個 block ÷ 6 個 block 期 = 每期 6 班，
剛好等於時段數，因此她**每個五週期都必須在全部六個時段各教一班**，毫無走位空間。
任何一個英文班改動日子或時間，整個 Net 編排都要重算，且未必仍然可行。此點已寫入給 Jo 的回覆。

---

## Jo 已確認的決定

| 項目 | 決定 |
|------|------|
| Task 1 性質 | **不重新排班**，只從 `Term1` 網格抽取 |
| CSW 房間對照 | `CSW - C1` 至 `CSW - C6` **各自保留原碼**（`Classroom Code` sheet 的 fill-down 是錯的，須 override） |
| 未排班的 6 個警察學員班 | `DAE256_G/H`、`DAE258_G/H`、`DAE260_G/H` 由警察學院另行安排，Jo 收到後再提供 → **本次排除** |
| `DAE270` | 待 AASFP 確認師資；如 T2026C 無法安排則順延至 T2027A → **本次排除** |
| 空白課室格 | 屬 combine class，**沿用上方最近有房名的一行**（room + teacher 共用）。Jo 用底色標示同組 |
| 節數 | 15 個上課日 × 2 節 = **30 個 topic**；Loading=2 的科目則 15 個 |
| Programme 欄 | `DAE - FT2026C` |
| 每班英文總時數 | 30 週 = **6 個 block × 20 小時 = 120 小時** |
| CS1 / CS2 特例 | Net 只需 **20 小時**（1 個 block），Mr. Ivan Yuen 各佔 100 小時 |
| Net 教師 | Mr. Peter Barrett（360 小時）、Ms. Loveleen Kaur（**我方假設 720**，Jo 給 680，其原始區間為 640–720） |

---

## 已驗證的數據事實

- `Classes` 共 175 行；`Term1` 網格覆蓋 173 個 code。扣除上述 7 個排除項後 **覆蓋率 100%**。
- `Term2` 那批 sheet 屬 26/27 下學期（含 DAE105、DAE108），與 T2026C 無關，**不得讀取**。
- 空白課室格共 27 個。以「向上填充」推斷後，**26 個都與 Jo 的底色分組一致**；唯一例外 `Cherry HD` 不是 DAE 科目，可忽略。
- 111 個時間格為空，須由 row 1 的欄頭時間（0900 / 1100 / 1400 / 1600）補上。
- 時間格式有 `0900-1100`、`0900 - 1100`、`0900 -1100` 三種，另有 typo `0900-1101`（應為 1100）。須正規化為 `HHMM - HHMM`。
- 網格內 Lecturer 欄是 XLOOKUP 公式。**應直接由 `Classes!E` 依 code 查表，不要讀網格的快取值。**
- Sheet 名 `Mon(Term1) ` 尾部有一個空格。
- ⚠️ **Calendar 的 AM 與 PM 日期不一定相同**（例：Fri AM Topic 4 = 2026-09-25，Fri PM Topic 4 = 2026-10-02）。必須按首節開始時間判斷 AM（<1300）或 PM（≥1400）後再查表。

---

## Net 鐘數：採用 Loveleen = 720 作為工作假設

Jo 最新數字（Peter 360 + Loveleen 680 = 1,040）較需求 1,080 少 40 小時。
Jo 第一封信原本給的區間是 **640–720**，因此取上限 720 屬其授權範圍內，
**決定以 720 開工，並在回覆中明確聲明此假設**。

```
需求  28 班 × 2 個 Net block                    = 56 個 block
      CS1、CS2 各減 1 個                         = -2
                                          小計  = 54 個 block (1,080 hrs)
供給  Loveleen 720/20 = 36  +  Peter 360/20 = 18 = 54 個 block (1,080 hrs)
                                          剩餘  =  0
```

⚠️ **兩點必須記住：**

1. **零彈性。** 每學期 3 個 block 期、兩學期共 6 個。Loveleen 每個 block 期都要同時
   帶 6 班、Peter 3 班，完全沒有走位空間。任何一個時段衝突都會直接變成缺口。
2. **真正的瓶頸是車程而非鐘數。** `MASTER_PLAN.md` J10 記錄：KT1 與 WT2 排不到 Net
   teacher，因兩班都落在週二 14:00 的擁擠時段，且與兩位 Net teacher 已承諾的地點相距
   超過 30 分鐘。2026-07-08 曾測試加入第三位 Net teacher，**證實無效**。
   把帳面補平不等於排得出。

**因此模組 C 的輸出規則：** 不得為了湊數而放寬 30 分鐘車程限制。
凡是仍然拿不到 Net block 的班別，一律連同原因如實列入 `_Issues`，交由 Jo 決定。

## 實作設計

```
Mon..Sat(Term1) 網格 ─┐
Classroom Code       ─┼─► [A] grid_to_classes ──► Classes!K:M + _Issues
                      │            │
DAE Calendar AA-AM   ─┼────────────┴─► [B] daily_expand ──► DailyTimetable_T2026C.xlsx
                      │                       ▲
ClassTeacher26-27    ─┴─► [C] english_net_assign ──► ClassTeacher26-27!M:R + 工時表
                                               │
                      ┌────────────────────────┘
                      ├─► [E] teacher_view ──► DAE_TeachingSchedule_T2026C.xlsx（10 位教師）
                      └─► [F] class_view   ──► ClassTimetable_T2026C.xlsx（32 班）
```

模組 A、B、E、F 為決定性運算（無 AI，可重複驗證）；只有模組 C 需要排程演算法。
模組 D（Task 4）= 模組 C 完成後重跑模組 B，DAE102 各行的 Lecturer 按 topic 分段取值（1–10 / 11–20 / 21–30）。

### 模組 A — `grid_to_classes`
1. 只讀 regex `^(Mon|Tue|Wed|Thu|Fri|Sat)\(Term1\)\s*$` 的 sheet。
2. 以 row 2 的 `Code` 欄位定位每個 block，每 block 讀 5 欄。
3. 時間空白 → 取 row 1 欄頭時間；然後正規化。
4. 課室空白 → 向上填充；同時記入 `_Issues`。
5. 同班連續 block 合併（`0900-1100` + `1100-1300` → `0900 - 1300`）。
6. 寫入 `Classes!K/L/M`；M 欄經 `Classroom Code` 對照，但 **CSW 系列保留原碼**。

### 模組 B — `daily_expand`
- 節數以 `Classes!H (Loading)` 判定：4 → 2 節/日；2 → 1 節/日（不要寫死科目 code）。
- Topic 編號 = `(第幾堂 − 1) × 每日節數 + 該日第幾節`。
- 輸出欄位跟足 2025 範本：`Raw / Code / Subject(En) / Subject(Ch) / Mode / Topics / Date / Day / Time / Venue / Lecturer / Programme / Remark / Year / Month / Date`。
- Mode 固定 `D`；Programme 固定 `DAE - FT2026C`；G 欄寫 `=TEXT(DATE(N{r},O{r},P{r}),"YYYY-MM-DD")`。
- 預估約 4,700 行。

### 模組 C — `english_net_assign`
- 決策變數：28 班 × 6 個 block = 168 個 block。
- 硬約束：① 每班每學期 1 個 Net block（CS1、CS2 全期共 1 個）；② 同一教師同一 block 期內不可撞相同 day+time；③ 同日跨中心車程 ≤ 30 分鐘；④ 須具英文資格。
- 固定指派：Ivan Yuen — CS1 100 小時、CS2 100 小時。
- 目標：先填滿 FT（Cherry Ip 420、Hailey Wong 660、Loveleen Kaur 720 — 見上方假設），餘額才派 PT；同班盡量少換教師。
- 輸出：`ClassTeacher26-27!M:R` + 更新後工時表。

### 模組 E — `teacher_view`
- 格式依 `DAE_TeachingSchedule.xlsx`：每位教師一個 sheet，sheet 名用**名字**（`Jo`、`Kevin`、`ChrisHon`）。
- A2/B2/C2 = `Teacher` / 教師名 / `Teaching Schedule T2026C`（C2:H2 合併）。
- Row 3 = `From | To | Monday…Saturday`；Row 4 起為 **0830–1930** 的 30 分鐘格（class view 只到 1830）。
- 授課格按時長合併儲存格。
- ⚠️ **範本內 11 個 sheet 的儲存格寫法完全不一致**（教師自行填寫所致）：
  `DAE103_KT3
(KT)`（Kevin）、`DAE105_FL2
ST`（Jo）、`CSW
DAE101`（Chris）、
  `數學 (WT2)`（Edward）、`KT1`（Cherry）、`DAE102 - CS4`（ChrisHon）。
  **建議統一採用 Kevin 的寫法 `<科目code>
(<中心>)`**，待 Jo 確認。
- 只填授課格。Office / Lunch / Travel / Meeting / ECA 等非授課項目由教師自行補（Jo 明言此檔是
  「before it is filled in by teachers」的版本）。
- 10 位教師：Jo Hugh、Kevin Ho、Chris Chau、Man Li、Alan Ho、Eddie Chueng、Edward Siu、
  Cherry Ip、Loveleen Kaur、Hailey Wong。
- ⚠️ 名單與範本不符：**Loveleen Kaur、Hailey Wong 範本中沒有 sheet**（需新建）；
  範本中的 Mike、ChrisHon、Mercury(PT) **不在 Jo 的名單內**（不產出）。
- 模組 E 與模組 F 結構相同，應共用同一個 grid renderer，只是 pivot 不同。

### 模組 F — `class_view`
- 格式依 `ClassTimetable_T2025C.xlsx`：每班一個 sheet，A2/B2 = `Class` / 班名，row 3 為 `From | To | Monday…Saturday`，row 4–23 為 0830–1830 的 30 分鐘格。
- 儲存格內容 = `科目中文名\n中心代號`，按時長合併儲存格。
- 32 班：CS1–CS7、WT1–WT2、SW1、TW1–TW3、TM1–TM5、TS1–TS4、KT1–KT4、TK1、ST1–ST2、FL1–FL3。
- TM4、TM5 只會有英文科資料；TS1–TS4 由 Jo 手動補充。
- 註：Jo 信中寫「30 classes」，但列出的清單有 32 個，須確認。

---

## 落地方式

`api/index.py` 新增 endpoint（`/api/export-daily`、`/api/assign-english-net`、`/api/export-views`），
`api/v4.html` 加對應按鈕。不改動 `api/` 目錄結構（Vercel 部署要求）。

---

## 待 Jo 回覆

1. 確認 Loveleen 由 680 調至 720（我方已按此開工）
2. Teacher view 儲存格寫法是否統一為 `<科目code>
(<中心>)`
3. Class view 是 30 班還是 32 班
4. Teacher view 名單：Loveleen Kaur、Hailey Wong 需新建 sheet；Mike、ChrisHon、Mercury(PT) 是否確定不出
