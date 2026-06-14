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

## 未來規劃（分享給其他用戶）
目前架構適合「每人複製一份試算表」的方式分享。
前置作業：清空個人資料建立範本、Cloud Run 服務處理方案、友善化 `installWizard_Init`。
**暫不考慮**單一 Web App 多用戶架構（改動量過大）。
