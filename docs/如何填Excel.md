# 如何填寫排班 Excel（輸入資料）

> 這份說明你要在輸入 Excel 填什麼。以 V4 範本為基礎，各分頁已備好。
> **最重要：「Class list answer」分頁要完全留空**，系統才會自動排班。

---

## 分頁一覽（哪些要填）

| 分頁 | 是否必填 | 填什麼 |
|---|---|---|
| Class list | ✅ 必填 | 每個班別 |
| Teacher load table with subject | ✅ 必填 | 每個老師每一科能教幾多班 |
| Centre Room Allocation | ✅ 必填 | 每間課室的座位數 |
| Teacher Availability | 選填 | 老師不可用的時間 |
| **Class list answer** | ⭐ **留空** | 系統自動填 |
| Net Teachers | 選填（英文科） | Net 老師名單 |
| English Weekly | 選填 | 英文老師每週安排 |

---

## 1. Class list（班別清單）— 必填

每行一個班。欄位：

| 欄 | 填什麼 | 例 |
|---|---|---|
| Subject Code | 班代碼 = **科目_組別** | `DAE101_CS1` |
| Subject Name (English) | 科目英文名 | `Chinese Language` |
| Subject Name (Chinese) | 科目中文名 | `中國語文` |
| Lecturer 1 / 2 / 3 | **留空**（系統自動派老師）；想指定才填 | |
| **Loading** | 一班每週小時：**4**＝4 小時、**2**＝2 小時 | `4` |
| Student No | 學生人數（**要準，決定用多大課室**） | `31` |
| Venue / Time / Date (optional) | **留空**（系統自動排） | |
| CC Group（如有這欄） | 要合併的班填**相同標籤**；不合併留空 | `CC-DAE106-1` |

- **班代碼格式：** `科目_組別`。組別＝中心＋組號，例 `CS1`＝長沙灣第 1 組、`TS3`＝沙田第 3 組。
- **Cadet 班**：組別用單一英文字母（例 `DAE256_E`），系統自動當 Cadet，排一/三/五。

---

## 2. Teacher load table with subject（老師能教什麼）— 必填 ⭐

**這就是「teacher loading」。** 每行一個老師，每一科填一個**數字 = 他這一科能教幾多班**。

| No. | Teachers | Chinese Language | English Language | Mathematics | … |
|---|---|---|---|---|---|
| Original | Ms. Jo Hugh | | | **1** | |

- **數字** = 該老師該科可教的班數（例：Mathematics 填 `1` = 可教 1 班數學）。**留空或 0 = 不教該科**。
- 每一科**用一欄**；欄的標題要用科目英文名（與 Class list 的 Subject Name 對得上）。
- 「No.」欄有填（例 `Original` 或編號）= 該老師是**主教**（Lec1）。
- 這張表決定「邊個老師可以教邊科、最多教幾多班」。**沒有在這裡出現的老師，系統不會派他上堂。**

> 注意：這裡的「幾多班」是**能力配額**，跟 Class list 的「Loading（小時）」不同，也跟「每週上限 6 班」不同。

---

## 3. Centre Room Allocation（課室）— 必填

| 欄 | 填什麼 | 例 |
|---|---|---|
| Classroom Code | 課室代碼 = **中心 - 房號**（中心由「-」前面決定） | `CSW - C1` |
| No of Seats | 實際座位 | `48` |
| Max no of seats | 認可上限座位（**系統優先用這個**；沒有才用 No of Seats） | `48` |

- 系統**不會**把班排入座位少於人數的課室。
- 課室代碼的中心前綴要正確（`CSW`、`TS`、`FL`…），系統靠它分中心。

---

## 4. Teacher Availability（老師不可用時間）— 選填

- 第一欄是老師名，其後每欄是一個時段（`Mon 0900`、`Mon 1100`… `Thu 1600`…）。
- 老師**不可用**的格填 **`N`**；可用就**留空**。全部留空 = 任何時間都可以。

---

## 5. Class list answer（排班結果）— ⭐ 留空！

- V4 自動排班時**完全留空**。系統會自動填 Day / Time / Venue / Lecturer。
- **有殘留資料的話，系統會當「已排好」照抄，不會自動排** —— 這是最常見的錯。

---

## 6. Net Teachers（英文 Net 老師）— 選填（英文科用）

- 每行一個 Net 老師姓名（要與 Teacher load table 的名字一致）。

## 7. English Weekly（英文每週安排）— 選填

- 留空 = 系統自動分配英文老師；只在想手動鎖定某班某段時才填。

---

## 合併班（CC Group）

- 在 Class list 的 **CC Group** 欄，要合併的班填**相同標籤**（例 `CC-DAE102-1`）。
- 必須**同科**；合併後**總人數要坐得下一間課室**（只合併小班）。

---

## 常見錯誤

| 錯誤 | 後果 | 正確做法 |
|---|---|---|
| Class list answer 有殘留資料 | 不自動排，只照抄 | **清空**這頁 |
| Student No 留空或填 0 | 課室大小算錯 | 填**準確**人數 |
| 老師沒在 Teacher load table 出現 | 系統不會派他上堂 | 在表內填上他能教的科目班數 |
| Teacher load table 填 0／留空 | 該老師不教該科 | 想他教就填班數（如 1、2） |
| 課室代碼中心前綴錯 | 認錯中心 | 用 `中心 - 房號`（如 `CSW - C1`） |
| CC Group 標籤不完全相同 | 不會合併 | **一字不差**填相同標籤 |

---

*規則背景見 `排班規則與排法-完整版.md`。*
