# 你不理財，才不理你 — CLAUDE.md

## 專案概述
個人股票記帳 Web App，使用 Google Apps Script (GAS) + Google Sheets 作為後端與資料庫，Vue 3 + Tailwind CSS 作為前端介面。

## 檔案結構

| 檔案 | 說明 |
|------|------|
| `4.0.js` | 主要 GAS 後端：安裝精靈、Gmail 擷取、庫存重建（FIFO）、股利、已實現損益、DCA |
| `WebAPI.js` | Web App 入口與 API 函式：登入驗證、儀表板、資料清單、設定、Gemini 市場分析 |
| `Index.html` | 前端單頁應用（Vue 3 + Tailwind），含所有頁面與互動邏輯 |
| `appsscript.json` | GAS 設定檔，通常不需要動 |
| `.clasp.json` | clasp 設定，含 Script ID，**不要刪除、不要公開** |

> GAS 專案裡所有 `.gs` 檔案共用同一個命名空間，**不可有重複的函式名稱**。
> **函式歸屬原則：`api_*` 函式只放 `WebAPI.js`；核心排程邏輯放 `4.0.js`。不跨檔重複定義。**

## 開發工作流程

**所有程式碼修改只在本地進行，GAS 編輯器當唯讀。**

```powershell
cd D:\Claude\Danveloper
clasp push   # 推送程式碼到 GAS
clasp deploy --deploymentId AKfycbzrFTtWxBH1aisKKkXWihYFittWQwUGldnjJTo3YE-jXonP_RoRhuoFsTKznW1Qtumw --description "說明"
```

改完記得清除瀏覽器快取才會生效。

其他指令：
- `clasp pull` — **若曾在 GAS 編輯器直接修改，push 前必須先 pull 確認差異**
- `clasp deployments` — 列出所有部署版本

## GAS 部署設定
- 執行身分：**我**（腳本擁有者）
- 存取權限：**任何人**（無需登入即可開啟 Web App）
- 類型：Web App
- **正式 Deployment ID**：`AKfycbzrFTtWxBH1aisKKkXWihYFittWQwUGldnjJTo3YE-jXonP_RoRhuoFsTKznW1Qtumw`
  - 這是 `index.html` 的 `DEFAULT_GAS_URL` 所指向的部署，**每次 deploy 都要用這個 ID**

## 前端架構
- `index.html`（小寫）是 **GitHub Pages 版本**，用 `fetch()` 呼叫 GAS `doPost` 端點
- `Index.html`（大寫）若存在是 GAS HtmlService 版本（`google.script.run`），**兩者架構不同，不可混用**
- App 正式入口：`https://danveloper99.github.io/finance-dashboard/`
- 前端呼叫後端的流程：`callGAS(action, args)` → fetch POST → GAS `doPost` → ALLOWED set 驗證 → 執行對應函式

## 試算表工作表結構
工作表名稱定義於「設定」分頁，詳細說明請見 App 內「系統說明」頁面。

| 設定 Key | 預設名稱 | 說明 |
|----------|----------|------|
| `SHEET_TRADES` | 交易紀錄 | Gmail 自動擷取，手動勿改 |
| `SHEET_HOLD` | 庫存紀錄 | 每日自動重建，手動勿改 |
| `SHEET_OPENING` | 期初庫存 | 手動填入一次 |
| `SHEET_DIV` | 股利狀況 | FinMind API 自動追加 |
| `SHEET_REALIZED` | 已實現損益 | 增量自動追加（清空可強制完整重建） |
| `SHEET_DCA` | 定期定額 | 自動追加 |
| `ALERT_LOG_SHEET` | 錯誤通知紀錄 | 自動記錄 |

## 架構重點

### 認證機制
- 密碼儲存於 Script Properties（`APP_PASSWORD`）
- Token 使用 **SHA-256 + 隨機 Salt** 產生（`makeToken_`），不可反推密碼
- Salt 儲存於 `TOKEN_SALT`（Script Properties）

### API Key 安全
- Gemini API Key 儲存於 Script Properties（`GEMINI_API_KEY`），前端看不到
- Cloud Run URL 儲存於「設定」工作表

### 已實現損益（增量模式）
- 預設為增量更新，只追加新賣出記錄
- 想強制完整重建：手動清空《已實現損益》工作表後再執行 `rebuildRealizedPnL_FIFO_SAFE`

### 股利（增量模式）
- 已是增量設計，`existingDivKeys` 去重，不會重複寫入
- 分批處理（每次 `DIV_SYMBOLS_PER_RUN` 檔），大量股票透過 `runDividendsFullCycle_SAFE` 自動接力

### 設定頁面
- 前端設定頁由 `api_getSettingsSchema` 的 schema 動態渲染
- Schema 全部使用 `group` 結構（含 `icon` 屬性）
- `api_saveSettings` 使用 batch write（一次讀取、記憶體更新、一次寫回）

## 編碼偏好
- **語言**：繁體中文回應，程式碼變數/函式名可英文
- **風格**：直接修改檔案，說明修改了哪些項目與原因
- **原則**：vibe coding，先求功能正確，不過度工程化
- **GAS 特性**：注意 6 分鐘執行時限；使用 `LockService` 避免並行寫入衝突

## 分享給其他用戶（線上載入架構）
- **機密原則**：每人用自己的 Google 帳號複製「範本試算表」並自己部署後端，作者看不到任何人的資料。不要提議「主控試算表 + 使用者對應表」或要求共用試算表給作者的方案。
- 範本試算表的 Apps Script 只有 `template/Code.js`（殼程式，shellVersion 4）：執行時從 GitHub Pages 下載 `releases/vN.js`（`4.0.js` + `WebAPI.js` 打包），用 `new Function` 執行。**殼程式之後不再修改**。
- 殼程式依 `release.json` 選版本：`current` 發布滿 1 天才採用，之前用 `previous`；設定頁「立即套用最新版」會寫入 `FIN_SKIP_DELAY` 跳過等待。程式碼快取在 CacheService（6 小時），下載失敗會退回 `FIN_LAST_GOOD_VERSION`。
- 線上載入模式下 ScriptApp / Properties / Lock 都是使用者自己的；程式裡用 `LOADER_`（`isLoaderMode_()`）判斷。
- **殼程式只有固定的排程函式名稱**（見 template/Code.js）。之後新增排程一律用 `newTaskTrigger_('函式名')`，會掛在殼程式的 `finTrigger`，不需要改殼程式。
- `template/appsscript.json` 的 oauthScopes **不要隨意增加**：新增權限會讓朋友的排程在重新授權前全部失敗。
- 每份後端只有一組帳號：Script Properties 的 `APP_USER` / `APP_PASSWORD`，由前端「首次設定」呼叫 `api_setupAccount` 建立。
- 前端必須相容舊版後端（朋友的版本最多落後 1 天，舊的程式庫版殼程式可能更舊）。
- **範本試算表不可用「複製主控試算表」產生**（會帶走個人資料）。範本一律用全新空白試算表 + `clasp -P .clasp.template.json push --force`。
- 舊版「程式庫殼程式」（FinLib，shellVersion 1～3）仍相容：程式庫的 Script Properties / Lock 是所有使用者共用的，所以程式碼一律用 `getProps_()` / `getScriptLock_()`，不可直接呼叫 `PropertiesService.getScriptProperties()` 或 `LockService.getScriptLock()`。

| 檔案 / clasp 設定檔 | 對象 |
|------|------|
| `.clasp.json` | 你自己的正式專案（綁定主控試算表，直接部署完整程式） |
| `.clasp.template.json` | 範本試算表的殼程式（`rootDir: template`） |
| `.clasp.lib.json` | 舊版程式庫 FinLib（給還沒換殼程式的舊使用者） |
| `releases/vN.js`、`release.json` | 線上載入用的程式包與版本資訊（GitHub Pages） |
| `tools/build-release.js` | 打包工具 |

### 發布新版本
1. `WebAPI.js` 的 `APP_VERSION` +1，`index.html` 的 `LATEST_BACKEND_VERSION` 改成一樣的數字
2. 跑測試後 `node tools/build-release.js "這次更新的說明"`（產生 `releases/vN.js`、更新 `release.json`；朋友 1 天後自動換新版）
   - 緊急修正或使用者說「立即發布」：加 `--now`（不留 previous，朋友約 10 分鐘內換新版）
3. 正式專案 `clasp push` + `clasp deploy --deploymentId ...`
4. （選用）舊程式庫使用者：`clasp -P .clasp.lib.json push --force` + `clasp -P .clasp.lib.json version "說明"`
5. commit（含 `releases/`、`release.json`）並 push → GitHub Pages；朋友在 1 天後自動更新
