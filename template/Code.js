/**
 * ============================================================
 * 你不理財，才不理你 — 範本殼程式（線上載入版）
 *
 * 所有功能的程式碼都放在 App 作者的 GitHub Pages，這裡負責下載（有快取）並執行。
 * 你的資料、密碼、Gemini Key 都只存在你自己的試算表與 Apps Script，作者看不到。
 *
 * 自動更新：作者發布新版本「滿 1 天」後自動生效，不需要做任何事。
 * ⚠ 這個檔案不需要、也請不要修改。
 * ============================================================
 */

var FIN_BASE_URL      = 'https://danveloper99.github.io/finance-dashboard/';
var FIN_SHELL_VERSION = 4;
var FIN_DELAY_MS      = 24 * 3600 * 1000; // 新版發布後等 1 天才採用
var FIN_CHUNK         = 30000;            // CacheService 每格上限 100KB，中文字最多 3 bytes
var FIN_ = null;                          // 本次執行已載入的程式

/** 取得已載入的程式（每次執行只載入一次） */
function fin_() {
  if (FIN_) return FIN_;
  const props = PropertiesService.getScriptProperties();
  const rel = finPickVersion_(props);
  let code = null, version = rel.version;
  try {
    code = finLoadCode_(version);
  } catch (e) {
    // 下載失敗：退回上次成功載入的版本（快取裡還有的話）
    const last = Number(props.getProperty('FIN_LAST_GOOD_VERSION')) || 0;
    if (!last || last === version) throw e;
    code = finLoadCode_(last);
    version = last;
  }
  const lib = new Function(code)();
  if (!lib || lib.__version !== version || typeof lib.doPost !== 'function') throw new Error('程式碼載入失敗（v' + version + '）');
  props.setProperty('FIN_LAST_GOOD_VERSION', String(version));
  if (typeof lib.setLoaderInfo === 'function')
    lib.setLoaderInfo({ shellVersion: FIN_SHELL_VERSION, version: version, latest: rel.latest, readyAt: rel.readyAt });
  FIN_ = lib;
  return lib;
}

/** 讀 release.json 決定要用哪一版：最新版發布滿 1 天才用，否則用上一版 */
function finPickVersion_(props) {
  const cache = CacheService.getScriptCache();
  let txt = cache.get('fin_release');
  if (!txt) {
    try {
      const r = UrlFetchApp.fetch(FIN_BASE_URL + 'release.json?t=' + Date.now(), { muteHttpExceptions: true });
      if (r.getResponseCode() === 200) {
        txt = r.getContentText('utf-8');
        cache.put('fin_release', txt, 600);
        props.setProperty('FIN_RELEASE_BACKUP', txt);
      }
    } catch (e) {}
    if (!txt) txt = props.getProperty('FIN_RELEASE_BACKUP'); // GitHub 連不上：用上次讀到的
  }
  if (!txt) throw new Error('無法取得版本資訊（release.json）');
  const rel = JSON.parse(txt);
  const cur = rel.current || {};
  const latest = Number(cur.version) || 0;
  const readyAt = (Date.parse(cur.releasedAt) || 0) + FIN_DELAY_MS;
  const skip = Number(props.getProperty('FIN_SKIP_DELAY')) === latest; // 設定頁按了「立即套用最新版」
  const useLatest = Date.now() >= readyAt || skip || !rel.previous;
  return { version: useLatest ? latest : Number(rel.previous.version), latest: latest, readyAt: readyAt };
}

/** 下載某一版的程式碼（快取 6 小時，分段存放） */
function finLoadCode_(version) {
  const cache = CacheService.getScriptCache();
  const key = 'fin_code_v' + version;
  const n = Number(cache.get(key + '_n')) || 0;
  if (n) {
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const parts = cache.getAll(keys);
    if (keys.every(k => parts[k] != null)) return keys.map(k => parts[k]).join('');
  }
  const r = UrlFetchApp.fetch(FIN_BASE_URL + 'releases/v' + version + '.js?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('無法下載程式碼 v' + version + '（' + r.getResponseCode() + '）');
  const code = r.getContentText('utf-8');
  if (code.indexOf('// @@FIN_BUNDLE v' + version + '\n') !== 0) throw new Error('下載的程式碼不正確（v' + version + '）');
  const put = {};
  let i = 0;
  for (let p = 0; p < code.length; p += FIN_CHUNK) put[key + '_' + (i++)] = code.slice(p, p + FIN_CHUNK);
  put[key + '_n'] = String(i);
  try { cache.putAll(put, 21600); } catch (e) {}
  return code;
}

/* ===== Web App 入口（前端 App 呼叫的就是這個） ===== */
function doPost(e) { const f = fin_(); return f.doPost.call(f, e); }

/* ===== 首次設定：依序執行 ===== */
/** 1. 建立《設定》分頁（注意：會清空既有的設定值，只在第一次執行） */
function installWizard_Init()  { return fin_().installWizard_Init(); }
/** 2. 填好《設定》後執行：建立交易、庫存、股利等工作表 */
function installWizard_Apply() { return fin_().installWizard_Apply(); }
/** 3. 建立每日自動排程：Gmail 擷取、更新庫存、損益、股利、月報 */
function setupAllSuggestedTriggers_SAFE()  { return fin_().setupAllSuggestedTriggers_SAFE(); }
/** 移除所有自動排程 */
function removeAllSuggestedTriggers_SAFE() { return fin_().removeAllSuggestedTriggers_SAFE(); }

/* ===== 排程會呼叫的函式（名稱必須固定，請勿改名） ===== */
function ingestFromGmail_Plaintext_SAFE() { return fin_().ingestFromGmail_Plaintext_SAFE(); }
function rebuildAll_B_SAFE()              { return fin_().rebuildAll_B_SAFE(); }
function dailyDataMaintenance_SAFE()      { return fin_().dailyDataMaintenance_SAFE(); }
function rebuildRealizedPnL_FIFO_SAFE()   { return fin_().rebuildRealizedPnL_FIFO_SAFE(); }
function appendDCAFromHoldings_SAFE()     { return fin_().appendDCAFromHoldings_SAFE(); }
function runDividendsFullCycle_SAFE()     { return fin_().runDividendsFullCycle_SAFE(); }
function updateDividendsFromFinMind_SAFE(){ return fin_().updateDividendsFromFinMind_SAFE(); }
function wealthReminder()                 { return fin_().wealthReminder(); }
function monthlyReport_SAFE()             { return fin_().monthlyReport_SAFE(); }
/** 通用排程入口：之後新增的排程都會用這個，殼程式不必再修改 */
function finTrigger(e)                    { const f = fin_(); return f.runTriggerTask.call(f, e); }
