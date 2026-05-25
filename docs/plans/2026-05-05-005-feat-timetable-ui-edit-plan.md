---
title: "feat: Timetable UI inline edit — click-to-edit + drag-and-drop + time display fix"
type: feat
status: active
date: 2026-05-05
---

# Timetable UI Inline Edit

## Problem Statement

排班結果生成後，用戶需要在 web UI 直接調整：
- 修改老師（Lec1/2/3）
- 移動某個 class 到不同 Day / Time / Room

現在只能下載 Excel 手動修改，改完無法在 UI 確認效果。

此外，現有 timetable grid 有兩個顯示問題：
1. **Column header 只顯示 start time**（"0900"），沒有 end time，容易混淆
2. **時間 hardcoded**，不支援特殊時間（如 0830、1030）

---

## 功能範圍

### Phase 0 — Time display fix（最快，先做）
- Column header 顯示完整時段：`0900-1100` 而非 `0900`
- Dynamic columns：從實際數據掃描所有時段，支援任意特殊時間
- 4 小時班用 `colspan=2` 橫跨兩個 time column，視覺更清晰

### Phase 1 — Click-to-edit（簡單，優先做）
點擊 timetable grid 任何 cell 的 Lec1 → inline 編輯 → 按 Enter 儲存

### Phase 2 — Drag-and-drop（進階）
拖動 class card → 放到另一個空置 time slot → 更新 Day / Time / Room

---

## 技術方案

### Phase 0 實作：Time display fix

**問題根源：** `SLOT_ORDER` 和 `SLOT_LBL` 硬編碼了 4 個標準時段，無法處理 0830、1030 等特殊時間。

**Dynamic column 生成：**

```js
// 從 grid data 掃描所有實際出現的時段，排序後作為 columns
function getSlotOrder(dayData) {
  const slots = new Set();
  Object.values(dayData).forEach(roomSlots => {
    Object.keys(roomSlots).forEach(slot => slots.add(slot));
  });
  return [...slots].sort();  // "0830 - 1030", "0900 - 1100", ... 按字母序 = 時間序
}
```

**Column header 顯示完整時段：**
```js
// "0900 - 1100" → "0900-1100"（縮短連字號節省空間）
function fmtSlot(slot) { return slot.replace(' - ', '-'); }
```

**4 小時班 colspan（time1 + time2 橫跨兩格）：**
```js
// 建立 cell 時：若 entry.time2 存在且 time2 是下一個相鄰 slot → colspan=2，跳過該 column
const slotOrder = getSlotOrder(dayData);
// 渲染時追蹤 skip set，避免在 time2 column 再畫一個空格
```

示意圖效果：
```
Room    │ 0830-1030 │ 0900-1100       │ 1030-1230 │ 1100-1300 │ ...
────────┼───────────┼─────────────────┼───────────┼───────────┼───
TW-TW1  │           │ DAE101 (4h) ────┤           │           │
TW-TW2  │ DAE102 ───┤                 │ ──────────┤           │
```

**`stats.timetable_grid` 對應改動（Python）：**

目前 grid 每個 entry 只存 `time1` 的 key。要支援 colspan，Python 需在 entry 裡保留 `time2`（已有）。JS 端利用 `entry.time2` 判斷是否 colspan，無需改 server。

---

### 狀態管理（純 client-side）

排班結果儲存為 JS `state.results[]`（陣列），每個 entry：
```js
{
  class_code, day, time1, time2, room_code,
  lec1, lec2, lec3, name_cn, name_en, student_count
}
```
所有編輯只改這個陣列，timetable grid 重新 render。

**不需要 server round-trip** 直到用戶按下載。

---

### Phase 1 實作：Click-to-edit Lec1

**UI 互動：**
1. 每個 cell 的 Lec1 行加 `contenteditable="true"` 或換成 `<input>`
2. `blur` / `Enter` 時 → 更新 `state.results` → 重新 render grid
3. 已修改的 cell 加 `edited` class（黃色 highlight）

```js
// 點擊後變 input
cell.addEventListener('click', e => {
  const lecEl = e.target.closest('.cell-lec');
  if (!lecEl) return;
  const code = lecEl.closest('td').dataset.code;
  const entry = state.results.find(r => r.class_code === code);
  // show inline input, on blur → entry.lec1 = input.value; rerenderGrid()
});
```

**Server 不需改動**（Phase 1）。

---

### Phase 2 實作：Drag-and-drop

**拖動源（Draggable）：** 每個有 class 的 `<td>` 設 `draggable="true"`

**放置目標（Drop target）：** 空置的 `<td>`（即 cell 無 class）

**HTML5 Drag & Drop API：**
```js
// source cell
td.setAttribute('draggable', 'true');
td.addEventListener('dragstart', e => {
  e.dataTransfer.setData('class_code', entry.class_code);
});

// target cell (empty slot)
emptyTd.addEventListener('dragover', e => e.preventDefault());
emptyTd.addEventListener('drop', e => {
  const code = e.dataTransfer.getData('class_code');
  const { targetDay, targetRoom, targetSlot } = emptyTd.dataset;
  moveClass(code, targetDay, targetRoom, targetSlot);
});
```

**`moveClass()` 邏輯：**
```js
function moveClass(class_code, newDay, newRoom, newSlot) {
  const entry = state.results.find(r => r.class_code === class_code);
  entry.day    = newDay;
  entry.room_code = newRoom;
  entry.time1  = newSlot;
  entry.time2  = nextSlot(newSlot);  // 4h class → 自動填 time2
  rerenderGrid();
}
```

**衝突處理：**
- Drop 前檢查目標 slot 是否已有 class
- 若有：顯示 tooltip "此時間已有 [class_code]"，拒絕 drop

---

### Phase 3 實作：下載含修改的 Excel

新增 `POST /api/download` endpoint：
- Request body: `{ results: [...] }` (JSON)
- Server: `write_output_fast(results)` → return Excel bytes

```python
@app.route("/api/download", methods=["POST"])
def download_edited():
    data = request.get_json()
    results = data.get("results", [])
    output_bytes = write_output_fast(results)
    return send_file(
        BytesIO(output_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Timetable_Edited.xlsx"
    )
```

下載按鈕改為：
```js
async function downloadEdited() {
  const res = await fetch('/api/download', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ results: state.results })
  });
  const blob = await res.blob();
  // trigger download
}
```

---

## UI 改動清單

| 元素 | 改動 |
|---|---|
| `SLOT_ORDER` / `SLOT_LBL` | 移除 hardcode，改為 `getSlotOrder(dayData)` 動態生成 |
| Column header | 顯示 `0900-1100` 完整時段（`fmtSlot(slot)`） |
| 4h class `<td>` | 加 `colspan=2`，time2 column 跳過 |
| `showDayGrid()` | 每個 `<td>` 加 `data-code`, `data-day`, `data-room`, `data-slot` |
| `.cell-lec` | 加 `contenteditable` 或換 `<input>` |
| 空置 `<td>` | 加 `data-day`, `data-room`, `data-slot`（drop target 用） |
| 下載按鈕 | 改為呼叫 `/api/download` with current state |
| 新增 "已修改" badge | 顯示有幾個 class 被手動改動 |

---

## 驗收標準

- [ ] Column header 顯示完整時段（`0900-1100`，非 `0900`）
- [ ] 特殊時間（0830、1030 等）自動出現為獨立 column
- [ ] 4 小時班橫跨兩個 time column（colspan=2）
- [ ] 點擊 Lec1 → inline input → Enter → grid 更新，老師名稱改變
- [ ] 修改過的 cell 有視覺標記（黃色邊框）
- [ ] 拖動 class → 放到空白 slot → grid 更新，class 移到新位置
- [ ] 拖到已佔用 slot → 拒絕，顯示提示
- [ ] 下載 Excel → 包含所有修改（新老師 + 新位置）
- [ ] 重新生成排班 → 修改清除（回到系統排班結果）

---

## 實作次序建議

1. **Phase 0**（time display fix）— 2-3 小時，純前端，視覺改善最大
2. **Phase 3**（`/api/download` endpoint）— 1 小時，server 改動少
3. **Phase 1**（click-to-edit Lec1）— 半天，純 client-side
4. **Phase 2**（drag-drop）— 1 天，需處理 edge cases

**總計：約 2.5 天工作量**
