/**
 * ============================================================
 * 你不理財，才不理你 — 殼程式（範本試算表專用）
 * 所有功能都在程式庫 FinLib 裡，這裡只負責把「網頁請求」和「排程」轉給程式庫。
 * 資料、密碼、Gemini Key 都存在你自己的試算表與 Apps Script，App 作者看不到。
 *
 * 更新到新版本：
 *   1. 左側「程式庫」→ 點 FinLib → 版本選最新 → 儲存
 *   2. 右上「部署 → 管理部署作業」→ 鉛筆圖示 → 版本選「新版本」→ 部署（網址不會變）
 * ============================================================
 */

function lib_() {
  // 把「你自己這個專案」的服務交給程式庫：
  // 觸發器建在你的專案；密碼、Gemini Key 等設定存在你自己的指令碼屬性（不會和其他人共用）
  FinLib.bindEnv({
    scriptApp:   ScriptApp,
    props:       PropertiesService.getScriptProperties(),
    lockService: LockService,
  });
  return FinLib;
}

/** Web App 入口（前端 App 呼叫的就是這個） */
function doPost(e) { return lib_().doPost(e); }

/* ===== 首次設定：依序執行 ===== */
/** 1. 建立《設定》分頁（注意：會清空既有的設定值，只在第一次執行） */
function installWizard_Init() { return lib_().installWizard_Init(); }
/** 2. 填好《設定》後執行：建立交易、庫存、股利等工作表 */
function installWizard_Apply() { return lib_().installWizard_Apply(); }
/** 3.（選用）建立每日自動排程：Gmail 擷取、更新庫存、損益、股利 */
function setupAllSuggestedTriggers_SAFE() { return lib_().setupAllSuggestedTriggers_SAFE(); }
/** 移除所有自動排程 */
function removeAllSuggestedTriggers_SAFE() { return lib_().removeAllSuggestedTriggers_SAFE(); }

/* ===== 排程會呼叫的函式（名稱必須和程式庫一致，請勿改名） ===== */
function ingestFromGmail_Plaintext_SAFE() { return lib_().ingestFromGmail_Plaintext_SAFE(); }
function rebuildAll_B_SAFE()              { return lib_().rebuildAll_B_SAFE(); }
function dailyDataMaintenance_SAFE()      { return lib_().dailyDataMaintenance_SAFE(); }
function rebuildRealizedPnL_FIFO_SAFE()   { return lib_().rebuildRealizedPnL_FIFO_SAFE(); }
function appendDCAFromHoldings_SAFE()     { return lib_().appendDCAFromHoldings_SAFE(); }
function runDividendsFullCycle_SAFE()     { return lib_().runDividendsFullCycle_SAFE(); }
function wealthReminder()                 { return lib_().wealthReminder(); }
