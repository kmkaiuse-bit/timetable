# AI Agent 協作排班 — 設計藍本

**日期：** 2026-08-11
**狀態：** 構想 + 已有可運作雛型（見文末「已建立」）

## 一、構想（來自負責人）

負責排班的同事在**負責人原本的設計框架內**做優化（只改參數與資料，不重新設計），
改由 **AI agent 操作**。要求：

1. 同事只需與 AI 對話，不需懂電腦、不需用 git 或 GitHub。
2. 每一次更改都要**像 git 一樣有記錄**（改了什麼、影響如何）。
3. 負責人日後能逐項檢視，把好的更改**合併成完整的程式**，由負責人把關。
4. 同事**只在自己本機**操作 AI agent。

## 二、角色分工

| 角色 | 做什麼 | 接觸 git / GitHub？ |
|---|---|---|
| **排班同事** | 與本機 AI 對話：「把 TS1、TS2 合併去 TM」。看自己的 Excel。 | 否 —— 完全無感 |
| **AI agent（本機）** | 改設定/資料 → 跑數 → **每項更改自動本機 commit**（白話訊息 + 數據）→ 完成後打包交接檔。 | agent 代做，同事看不到 |
| **負責人** | 匯入交接檔，檢視每條 commit 的差異與影響 → 好的合併入 master。 | 是（只在負責人這邊） |

即現有「分支 → review → 合併 master」模式，只是同事那一端由 agent 隱藏了所有 git 操作。

## 三、架構

```
同事本機
  ├─ 項目資料夾（引擎 + 他們的 Excel）        ← 由負責人派發
  ├─ 本機 AI agent（如 Claude Code）
  ├─ config/rules.json      ← 規則的文字檔（agent 改；git 可讀差異）
  ├─ 本機 git 歷史          ← agent 每次更改自動 commit（同事無感）
  └─ CHANGELOG（白話）      ← agent 累積，供人閱讀
        │
        │  完成後 agent 產生「交接檔」（git bundle 或 patch + CHANGELOG）
        ▼
負責人本機 → 匯入 → 檢視差異/影響 → 合併入 master（→ 可選 push GitHub）
```

**要點：** git 是本機工具，不需 GitHub。交接用 `git bundle`（一個檔含完整歷史），
以電郵/雲端/USB 傳送即可。GitHub 只在負責人一端。

## 四、可改的範圍（護欄）

- **可改：** 設定（`config/rules.json`：上課日、老師上限…）+ Excel 資料（班、人數、CC Group…）。
- **不可改：** 引擎邏輯（演算法）—— 保護負責人原本的設計。
- **每次更改強制跑數回報**，不准靜默更改（例：曾靠此抓到「33 人塞 25 座房」的錯）。

## 五、「像 git 一樣的記錄」如何達到

1. **規則放文字檔**（`config/rules.json`）→ git 顯示清楚的一行差異，例如
   `上課日：[Mon,Tue,Thu] → [Mon,Tue,Wed,Thu]`。
2. **Excel 資料的更改** → agent 在 commit 訊息用白話寫摘要（例：「填入 CC-DAE102-1：TS1+TS2+TM4」）。
3. 每條 commit 都含：**白話說明 + 具體差異 + 影響數據**（排到/排不到/超額）。

## 六、已建立（可運作雛型）

- `config/rules.json` — 規則文字檔（`dae_days` / `cadet_days` / `teacher_weekly_cap`）。
- 引擎 `load_rules_config()` — 啟動時讀取並覆蓋；無檔則用預設（行為不變）。
- `scripts/verify.py` — 一個指令輸出指標報告（見 `docs/demo-walkthrough.md`）。
- 同事版說明書 `docs/排班規則-同事版.md`、技術版 `docs/system-modules.md`。

## 七、待建立

1. 把更多「規則行為」參數化進 `config/rules.json`（如揀房策略、combine 搜尋）。
2. **agent 操作守則**檔（只可改設定/資料、每次必 commit + 跑數 + 回報、完成產生交接檔）。
3. **CHANGELOG 自動累積** + **交接打包**（`git bundle`）的固定做法。
4. 派發給同事的**本機安裝包**（資料夾 + agent 設定）。
