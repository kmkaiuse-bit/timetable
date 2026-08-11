# 明日示範腳本

**目的：** 示範「同事用日常說話提出更改 → AI agent 改設定 → 即時跑數回報 → 每步有 git 記錄 → 負責人日後檢視」。

**示範檔：** `data/demo/demo_start.xlsx`（已放入項目，基準為 24 班排不到）。

**事前準備：** 在項目資料夾開一個終端機（能執行 `python`），確認可跑：
```
python scripts/verify.py "data/demo/demo_start.xlsx"
```

---

## 第 1 幕：現況（跑數）

**做：**
```
python scripts/verify.py "data/demo/demo_start.xlsx"
```
**預期輸出（重點）：**
```
DAE 上課日   : ['Monday', 'Tuesday', 'Thursday']
排到         : 102
排不到       : 24
超額（應為 0）: 0
排不到 — 各中心： CSW 6, FL 6, TS 6, WT 2, TM 2, KT 2
```
**講：** 「這是目前排班結果 —— 24 班排不到,全部因課室不足,沒有超額。」

---

## 第 2 幕：同事提出更改（設定）

**同事說（日常語言）：** 「試下如果 DAE 科目可以用星期三,會多排到幾多?」

**AI agent 做的事（示範者可代 agent 執行）：**

1. 改 `config/rules.json` 一行 —— 把 `dae_days` 改成含星期三：
   ```json
   "dae_days": ["Monday", "Tuesday", "Wednesday", "Thursday"],
   ```
2. 再跑數：
   ```
   python scripts/verify.py "data/demo/demo_start.xlsx"
   ```
   **預期：**
   ```
   DAE 上課日   : ['Monday', 'Tuesday', 'Wednesday', 'Thursday']
   排到         : 111      ← 由 102 升
   排不到       : 15       ← 由 24 跌
   超額（應為 0）: 0
   ```
3. **記錄這次更改（git）：**
   ```
   git add config/rules.json
   git commit -m "試:DAE 上課日加星期三 → 排不到 24→15,超額 0"
   ```

**講：** 「一句話、改一行,系統即時話你多排 9 班。而且這次嘗試已被記錄。」

---

## 第 3 幕：這條規則其實不准 → 還原（也要有記錄）

**同事說：** 「星期三其實不能用,還原吧。」

**做：** 把 `config/rules.json` 的 `dae_days` 改回 `["Monday", "Tuesday", "Thursday"]`,再：
```
git add config/rules.json
git commit -m "還原:DAE 上課日回到 Mon/Tue/Thu（星期三不可用）"
```
**講：** 「連『試完還原』都有記錄 —— 負責人看得到試過什麼、為何回退。」

---

## 第 4 幕：負責人檢視記錄（重點）

**做：**
```
git log --oneline -5
```
**預期（示意）：**
```
xxxxxxx 還原:DAE 上課日回到 Mon/Tue/Thu（星期三不可用）
xxxxxxx 試:DAE 上課日加星期三 → 排不到 24→15,超額 0
...
```
看某條改了什麼：
```
git show HEAD~1
```
**講：** 「負責人日後就是這樣逐條看同事改了什麼、影響多少,把好的合併入完整程式。同事全程只是對話,沒碰過 git。」

---

## 第 5 幕（可選）：交接（不用 GitHub）

**做：**
```
git bundle create my-changes.bundle master
```
**講：** 「同事把 `my-changes.bundle` 一個檔傳給負責人（電郵/雲端/USB 均可）,負責人匯入就看到全部更改。全程不需要 GitHub。」

---

## 收尾要點（對觀眾講）

1. 同事只需 **對話 + 自己的 Excel**;git、跑數都由 agent 隱藏。
2. **每次更改都有記錄**（改了什麼 + 影響數據），負責人可逐項檢視、選擇性合併。
3. **有護欄**：agent 只改設定/資料、不改演算法;每次強制跑數,不准靜默 —— 保護原本的設計。

---

## 備援（示範時如出錯）

- 若 `python` 找不到,改用 `py scripts/verify.py "..."`。
- 若數字不同,確認 `config/rules.json` 的 `dae_days` 目前值(第 1 幕應為 Mon/Tue/Thu)。
- 第 2 幕改完記得第 3 幕還原,否則 `config/rules.json` 會停在含星期三的狀態。
