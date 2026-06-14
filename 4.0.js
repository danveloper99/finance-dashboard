/***** ===========================================
 * (1) Gmail 擷取 → 交易紀錄（中文表頭 + 費/稅/淨額）
 * (2) 庫存紀錄（B 模式一鍵：內嵌 先沖期初賣單 → 再重建）
 * (3) 股利狀況（FinMind TaiwanStockDividendResult；APPEND ONLY）
 * SAFE：所有可排程入口皆提供 *_SAFE() 包裝
 * 修訂重點（2025-12-30）：
 * - 庫存 FIFO 邏輯升級：支援「同股票、不同券商」獨立計算成本
 * - 庫存表新增「證券商」欄位
 * - 設定頁新增券商參數
 * - 補回所有遺失的 SAFE 函式與工具
 * ============================================ */

/***** =======================
 * Part 1/4
 * 安裝精靈 + 共用工具 + SAFE 包裝
 * ======================== */

/** 產生〈設定〉分頁（保留使用者自訂說明 + 新增 DCA 中文名稱欄位） */
function installWizard_Init() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定') || ss.insertSheet('設定');
  sh.clear();

  // 表頭：鍵 / 值 / 說明
  sh.getRange(1,1,1,3).setValues([[ '鍵', '值', '說明' ]]).setFontWeight('bold');

  // === 1. 基礎設定 (完全保留您的說明文字) ===
  const rows = [
    ['TZ', 'Asia/Taipei', '時區（IANA 格式），例：Asia/Taipei、America/Los_Angeles。影響日期格式與排程時區。'],
    ['GMAIL_LABEL_HTML', '', '【HTML信件】單純內文解析的標籤名稱（無加密PDF）。'],
    ['GMAIL_LABEL_PDF',  '',  '【PDF信件】含有加密 PDF 附件的標籤名稱（需 Cloud Run 解鎖）。'],
    ['GMAIL_QUERY_DAYS', '7', '往回撈幾天的郵件。整數，例：7。'],
    
    // --- 🆕 券商設定 ---
    ['BROKER_DEFAULT_NAME', '國泰證券', '【券商】預設券商名稱（無特定關鍵字時使用）。'],
    ['FEE_DISCOUNT', '0.28', '手續費折數（0~1 之間的小數）。手續費 = ROUNDDOWN(成交金額 × 1.425‰ × 折數)，最低 1 元。例：0.28 = 2.8 折。'],
    ['BROKER_2_KEYWORD', '統一', '【券商】第二券商判定關鍵字（如：統一）。若無可留空。'],
    ['BROKER_2_NAME', '統一證券', '【券商】第二券商顯示名稱。'],
    ['BROKER_2_DISCOUNT', '0.6', '【券商】第二券商手續費折數 (0~1)。'],
    // ------------------

    ['ID_NUMBER', '', '【必要】身分證字號（用於 PDF 解鎖密碼）。'],
    ['CLOUD_RUN_URL', '', '【必要】Google Cloud Run 服務網址（https://...）。'],
    ['SHEET_TRADES', '交易紀錄', '交易紀錄分頁名稱（可自訂，但需與其它設定一致）。'],
    ['SHEET_HOLD', '庫存紀錄', '庫存紀錄分頁名稱。'],
    ['SHEET_OPENING', '期初庫存', '期初庫存分頁名稱（表頭與庫存相同）。'],
    ['SHEET_DIV', '股利狀況', '股利狀況分頁名稱（FinMind 追加資料的輸出表）。'],
    ['SHEET_REALIZED', '已實現損益', '已實現損益（FIFO）輸出分頁名稱。'],
    ['SHEET_DCA', '定期定額', '定期定額分頁名稱（自庫存複製 DCA 匯總）。'],
    ['SHEET_GUIDE', '安裝指引', '安裝指引分頁名稱（說明與檢查清單）。'],
    ['ALERT_ENABLED', 'TRUE', '是否啟用錯誤即時通知（TRUE/FALSE）。FALSE 仍會寫入錯誤紀錄分頁。'],
    ['ALERT_TO', '', '錯誤通知收件人 Email。建議填你自己的信箱。多位可用逗號分隔。'],
    ['ALERT_SUBJECT_PREFIX', '【自動化錯誤】', '錯誤通知主旨前綴。'],
    ['ALERT_LOG_SHEET', '錯誤通知紀錄', '錯誤紀錄分頁名稱。'],

    ['HOLDINGS_START_DATE', '', '庫存計算起算日（YYYY/MM/DD）。留空＝不限制。只納入此日（含）之後的交易。'],
    ['DIV_START_DATE', '', '股利計算起算日（YYYY/MM/DD）。留空＝不限制。僅計算此日（含）之後除權息。'],
    ['DIV_YEAR_FROM', String(new Date().getFullYear() - 8), '股利抓取的起始年度（含），通常抓近 8 年即可。'],
    ['DIV_STOCK_BONUS_TO_TRADES', 'TRUE', '是否把「股票股利入帳」寫回〈交易紀錄〉（TRUE/FALSE）。'],
    ['DIV_STOCK_BONUS_ROUNDING', 'FLOOR', '配股取整規則：FLOOR(無條件捨去) / ROUND(四捨五入) / CEIL(無條件進位)。'],
    ['DIV_CASH_FEE_PER_PAYOUT', '10', '每筆現金股利入帳之固定手續費（元）。無手續費請填 0。'],
    ['DIV_THROTTLE_MS', '1000', '抓取 FinMind 每檔之間的延遲（毫秒）。太小可能被限流。'],
    ['FINMIND_TOKEN', '', 'FinMind API Token。無 Token 可能遇到配額限制或 401。請至 FinMind 取得並貼上。'],
    ['DIV_SYMBOLS_PER_RUN', '10', '股利分批每輪處理檔數（建議 8~12）'],
    ['DIV_CURSOR_RESET_IF_STALE_DAYS', '3', '距上次執行超過 N 天即重置游標（建議 3）'],
  ];

  // === 2. 定期定額設定 (使用迴圈生成，確保文字與您的一致，並插入新欄位) ===
  for (let i = 1; i <= 10; i++) {
    const symbolDesc = (i === 1) ? `定期定額#${i} 股票代碼（例：00878）` : `定期定額#${i} 股票代碼`;
    rows.push([`DCA_${i}_NAME`,   '', `定期定額#${i} PDF中文名稱 (如: 國泰永續高股息)`]);
    rows.push([`DCA_${i}_SYMBOL`, '', symbolDesc]);
    rows.push([`DCA_${i}_START`,  '', `定期定額#${i} 開始日期（YYYY/MM/DD）`]);
    rows.push([`DCA_${i}_END`,    '', `定期定額#${i} 結束日期（留空＝至今）`]);
  }

  sh.getRange(2,1,rows.length,3).setValues(rows);

  // 驗證
  const boolRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE'], true).build();
  const roundingRule = SpreadsheetApp.newDataValidation().requireValueInList(['FLOOR','ROUND','CEIL'], true).build();
  const numberRule01 = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).build();

  const rowIndexByKey = Object.fromEntries(rows.map((r,i)=>[r[0], i+2]));
  ['ALERT_ENABLED','DIV_STOCK_BONUS_TO_TRADES'].forEach(k=>{
    if (rowIndexByKey[k]) sh.getRange(rowIndexByKey[k], 2).setDataValidation(boolRule);
  });
  if (rowIndexByKey['DIV_STOCK_BONUS_ROUNDING']) sh.getRange(rowIndexByKey['DIV_STOCK_BONUS_ROUNDING'], 2).setDataValidation(roundingRule);
  if (rowIndexByKey['FEE_DISCOUNT']) sh.getRange(rowIndexByKey['FEE_DISCOUNT'], 2).setDataValidation(numberRule01);

  // 版面
  sh.setFrozenRows(1);
  sh.setColumnWidths(1,1,160); 
  sh.setColumnWidths(2,1,240); 
  sh.setColumnWidths(3,1,560); 
  const lastRow = 1 + rows.length;
  try {
    sh.getRange(2,2,lastRow-1,1).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW).setWrap(false);
  } catch (e) {
    sh.getRange(2,2,lastRow-1,1).setWrap(false);
  }
  sh.getRange(2,3,lastRow-1,1).setWrap(true).setVerticalAlignment('top');

  SpreadsheetApp.flush();
  Logger.log('✅ 已重建〈設定〉分頁。');
}

/** 讀取〈設定〉分頁為物件 */
function getCfg_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定');
  if (!sh) throw new Error('找不到〈設定〉，請先執行 installWizard_Init()');

  const lastRow = sh.getLastRow();
  const raw = (lastRow>1) ? sh.getRange(2,1,lastRow-1,2).getValues() : [];
  const cfg = {};
  raw.forEach(([k,v]) => {
    if (k) cfg[String(k).trim()] = (v==null ? '' : String(v).trim());
  });
  return cfg;
}

/** 安裝精靈：依〈設定〉建立必要分頁 */
function installWizard_Apply() {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeader_(C.SHEET_TRADES, [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','證券商','備註'
  ]);
  ensureSheetWithHeader_(C.SHEET_OPENING, [
    '股票代碼','股票名稱','買進日期','買入價','持有股數',
    '買入成本 (單純買入價*股數)','手續費','證券商'
  ]);
  ensureSheetWithHeader_(C.SHEET_HOLD, [
    '股票代碼','股票名稱','買進日期','買入價','持有股數',
    '買入成本 (單純買入價*股數)','手續費','證券商','現價','即時市值'
  ]);
  ensureSheetWithHeader_(C.SHEET_DIV, [
    '股票代碼','股票名稱','股利所屬年度','除息日','現金股利發放日',
    '現金股利 (元/股)','除息日持股數',
    '現金總股息','實際領取金額 (扣除每筆手續費10元)',
    '平均成交價','個人殖利率 (現金/買入價)','備註'
  ]);
  ensureSheetWithHeader_(C.SHEET_REALIZED, [
    '股票代碼','股票名稱',
    '買進日期','股數','買進單價','買進成本',
    '賣出日期','賣出單價','賣出總金額',
    '買進手續費','賣出手續費','交易稅','淨獲利','股利','含息報酬','含息報酬率(%)','持有天數','每天獲利金額'
  ]);
  ensureSheetWithHeader_(C.SHEET_DCA || '定期定額', [
    '買入日期','股票代碼','股票名稱','成交價','股數',
    '買入成本 (=成交價*股數)','手續費','總成本 (=買入成本+手續費)'
  ]);
  ensureSheetWithHeader_(C.ALERT_LOG_SHEET || '錯誤通知紀錄', [
    '時間','入口函式','錯誤訊息','堆疊','附加資訊','試算表名稱','URL'
  ]);

  Logger.log('✅ 基礎表格已建立；填好〈設定〉即可開始使用 SAFE 入口或建立排程。');
}

/** 建議排程（可選） */
function setupAllSuggestedTriggers_SAFE() {
  const C = getCfg_();
  removeAllSuggestedTriggers_SAFE();
  const tz = C.TZ || 'Asia/Taipei';

  ScriptApp.newTrigger('ingestFromGmail_Plaintext_SAFE').timeBased().everyDays(1).atHour(18).nearMinute(0).inTimezone(tz).create();
  ScriptApp.newTrigger('rebuildAll_B_SAFE').timeBased().everyDays(1).atHour(18).nearMinute(30).inTimezone(tz).create();
  ScriptApp.newTrigger('rebuildRealizedPnL_FIFO_SAFE').timeBased().everyDays(1).atHour(19).nearMinute(0).inTimezone(tz).create();
  ScriptApp.newTrigger('appendDCAFromHoldings_SAFE').timeBased().everyDays(1).atHour(12).nearMinute(0).inTimezone(tz).create();
  ScriptApp.newTrigger('runDividendsFullCycle_SAFE').timeBased().everyDays(1).atHour(11).nearMinute(0).inTimezone(tz).create();

  Logger.log('✅ 已建立新排程。');
}

/** 解除建議排程 */
function removeAllSuggestedTriggers_SAFE() {
  const names = new Set([
    'ingestFromGmail_Plaintext_SAFE',
    'rebuildAll_B_SAFE',
    'rebuildRealizedPnL_FIFO_SAFE',
    'appendDCAFromHoldings_SAFE',
    'runDividendsFullCycle_SAFE',
    'updateDividendsFromFinMind_SAFE'
  ]);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (names.has(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  Logger.log('🧹 已移除建議排程。');
}

/* ===== 共用小工具 ===== */
function ensureSheetWithHeader_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(header);
  return sh;
}
function toYMDslash_(s) {
  if (!s) return '';
  s = String(s).trim();
  let m = s.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (m) return `${m[1]}/${('0'+m[2]).slice(-2)}/${('0'+m[3]).slice(-2)}`;
  m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return `${m[1]}/${('0'+m[2]).slice(-2)}/${('0'+m[3]).slice(-2)}`;
  m = s.match(/(\d{2,3})[\/年](\d{1,2})[\/月](\d{1,2})/); // 民國年
  if (m && Number(m[1])<200) return `${Number(m[1])+1911}/${('0'+m[2]).slice(-2)}/${('0'+m[3]).slice(-2)}`;
  try {
    const d = new Date(s);
    if (!isNaN(d)) {
      const yyyy=d.getFullYear(), mm=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
      return `${yyyy}/${mm}/${dd}`;
    }
  } catch(e){}
  return s;
}
function normTime_(t){
  const s = String(t||'').trim();
  const m = s.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
  if (!m) return s || '00:00:00';
  const HH = ('0'+m[1]).slice(-2);
  const MM = ('0'+m[2]).slice(-2);
  const SS = ('0'+(m[3]||'00')).slice(-2);
  return `${HH}:${MM}:${SS}`;
}
function round2_(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }
function round4_(x){ return Math.round((x + Number.EPSILON) * 10000) / 10000; }
function num_(s){ return Number(String(s).replace(/,/g,'')); }
function escapeHtml_(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function getFeeDiscount_() {
  try {
    const C = getCfg_ ? getCfg_() : {};
    let d = Number((C && C.FEE_DISCOUNT) != null ? C.FEE_DISCOUNT : 0.28);
    if (!isFinite(d)) d = 0.28;
    return Math.max(0, Math.min(1, d));
  } catch (e) {
    return 0.28;
  }
}

function calcFee_(amount) {
  if (!amount || amount <= 0) return 0;
  const discount = (typeof getFeeDiscount_ === 'function') ? getFeeDiscount_() : 0.28;
  const raw = amount * 1.425 / 1000 * discount;
  return Math.max(1, Math.floor(raw));
}

function calcTax_(sideText, amount){
  const s = String(sideText || '').trim();
  if (!amount || amount <= 0) return 0;
  if (s === '現賣') return Math.round(amount * 0.003);
  if (s === '沖賣') return Math.round(amount * 0.0015);
  return 0;
}

function calcNet_(sideText, amount, fee, tax){
  const s = String(sideText || '');
  if (s.includes('買')) return -Number(amount||0) - Number(fee||0) - Number(tax||0);
  if (s.includes('賣')) return  Number(amount||0) - Number(fee||0) - Number(tax||0);
  return - Number(fee||0) - Number(tax||0);
}

function normDate_(d){ return toYMDslash_(String(d||'').trim()); }
function normOrderNo_(s) {
  return String(s || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
}

function isStockBonusSide_(sideText) {
  if (!sideText) return false;
  const s = String(sideText).replace(/\s+/g, '').replace(/[()（）\[\]【】]/g,'');
  return (
    s.includes('股票股利') ||
    s.includes('配股入帳') ||
    (s.includes('股票股利') && s.includes('入帳')) ||
    s === '配股' || s === '股票股利'
  );
}
function isDayLoopSide_(sideText){
  const s = String(sideText||'').trim();
  return (s.includes('沖買') || s.includes('沖賣'));
}

/* ===== 錯誤通知 ===== */
function getAlertCfg_() {
  const C = getCfg_();
  return {
    ENABLED: (String(C.ALERT_ENABLED||'TRUE').toUpperCase()==='TRUE'),
    TO: C.ALERT_TO || '',
    CC: '',
    SUBJECT_PREFIX: C.ALERT_SUBJECT_PREFIX || '【自動化錯誤】',
    TZ: C.TZ || 'Asia/Taipei',
    LOG_SHEET: C.ALERT_LOG_SHEET || '錯誤通知紀錄'
  };
}
function onError_(entryName, err, extra) {
  try {
    const A = getAlertCfg_();
    const now = Utilities.formatDate(new Date(), A.TZ, 'yyyy/MM/dd HH:mm:ss');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(A.LOG_SHEET) || ss.insertSheet(A.LOG_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['時間','入口函式','錯誤訊息','堆疊','附加資訊','試算表名稱','URL']);
    }
    sh.appendRow([
      now, entryName, String(err && err.message || err),
      String(err && err.stack || ''), extra ? JSON.stringify(extra) : '',
      ss.getName(), ss.getUrl()
    ]);

    if (A.ENABLED && A.TO) {
      const html = [
        `<p>Hi，系統偵測到錯誤：</p>`,
        `<ul><li><b>時間：</b>${now}</li>`,
        `<li><b>入口函式：</b>${escapeHtml_(entryName)}</li>`,
        `<li><b>試算表：</b>${escapeHtml_(ss.getName())}</li>`,
        `<li><b>連結：</b><a href="${ss.getUrl()}" target="_blank">${ss.getUrl()}</a></li></ul>`,
        `<p><b>錯誤訊息：</b></p><pre>${escapeHtml_(String(err && err.message || err))}</pre>`,
        `<p><b>堆疊：</b></p><pre>${escapeHtml_(String(err && err.stack || ''))}</pre>`,
        extra ? `<p><b>附加資訊：</b></p><pre>${escapeHtml_(JSON.stringify(extra,null,2))}</pre>` : ''
      ].join('\n');
      MailApp.sendEmail({ to: A.TO, subject: A.SUBJECT_PREFIX + entryName + ' 錯誤', htmlBody: html });
    }
  } catch(e2){ Logger.log('onError_ failed: %s', e2); }
}
function runWithAlert_(fn, entryName, extra) {
  try { return fn(); } catch (err) { onError_(entryName, err, extra); throw err; }
}

/* ===== SAFE 入口 ===== */
function ingestFromGmail_Plaintext_SAFE(){ return runWithAlert_(ingestFromGmail_Plaintext,'ingestFromGmail_Plaintext'); }
function rebuildAll_B_SAFE(){ return runWithAlert_(rebuildAll_B,'rebuildAll_B'); }
function updateDividendsFromFinMind_SAFE(){ return runWithAlert_(updateDividendsFromFinMind,'updateDividendsFromFinMind'); }
function rebuildRealizedPnL_FIFO_SAFE(){ return runWithAlert_(rebuildRealizedPnL_FIFO,'rebuildRealizedPnL_FIFO'); }
function appendDCAFromHoldings_SAFE(){ return runWithAlert_(appendDCAFromHoldings, 'appendDCAFromHoldings'); }
function runDividendsFullCycle_SAFE() { return runWithAlert_(runDividendsFullCycle_, 'runDividendsFullCycle'); }
function rebuildDCADividends_SAFE() { return runWithAlert_(appendDCAFromHoldings, 'appendDCAFromHoldings'); }

/* ===== 測試用 ===== */
function testErrorAlert_SendSample() {
  onError_('testErrorAlert_SendSample', new Error('這是一封測試錯誤通知'), { hint: '僅測試郵件與紀錄' });
}

/***** =======================
 * Part 2/4
 * Gmail 擷取
 * ======================== */

function makeKey_NoOrder_(r) {
  return [r[0], normTime_(r[1]||''), r[2], String(r[4]||'').trim(), r[5], r[6]].join('|');
}
function makeKey_Order_(dateStr, orderNo, symbol) {
  return [String(dateStr||'').trim(), normOrderNo_(orderNo), String(symbol||'').trim()].join('|');
}

function consolidateByOrderNo_(rows) {
  const map = new Map();
  const noOrderList = [];

  for (const r of rows) {
    const [date, time, code, name, type, qty, price, amt, orderNo, fee, tax, net] = r;
    const normalizedOrderNo = normOrderNo_(orderNo);

    if (normalizedOrderNo === '') {
      noOrderList.push(r);
      continue;
    }

    const key = makeKey_Order_(date, orderNo, code);
    const timeN = normTime_(time || '');

    if (!map.has(key)) {
      map.set(key, {
        date: date,
        timeMax: timeN || '',
        sym: String(code || '').trim(),
        name: name || '',
        side: String(type || '').trim(),
        qtySum: Number(qty || 0),
        amtSum: Number(amt || 0),
        feeSum: Number(fee || 0), 
        taxSum: Number(tax || 0),
        orderNo: orderNo
      });
    } else {
      const o = map.get(key);
      if (timeN > (o.timeMax || '')) o.timeMax = timeN;
      o.qtySum += Number(qty || 0);
      o.amtSum += Number(amt || 0);
      o.feeSum += Number(fee || 0);
      o.taxSum += Number(tax || 0);
    }
  }

  const merged = [];
  for (const o of map.values()) {
    const px = o.qtySum !== 0 ? round2_(o.amtSum / o.qtySum) : 0;
    const finalFee = o.feeSum;
    const finalTax = o.taxSum;
    const finalNet = calcNet_(o.side, o.amtSum, finalFee, finalTax);

    merged.push([
      o.date, o.timeMax, o.sym, o.name, o.side,
      o.qtySum, px, o.amtSum, o.orderNo, finalFee, finalTax, finalNet, '合併訂單'
    ]);
  }
  return merged.concat(noOrderList);
}

function callCloudRunToUnlock_(blob, password) {
  const C = getCfg_();
  const apiUrl = C.CLOUD_RUN_URL;
  if (!apiUrl) throw new Error("未設定 CLOUD_RUN_URL");

  const payload = { file_content: Utilities.base64Encode(blob.getBytes()), password: password };
  const options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };

  // 冷啟動重試：最多 3 次，每次間隔 15 秒
  const MAX_RETRY = 3;
  const RETRY_WAIT_MS = 15000;

  for (let attempt = 1; attempt <= MAX_RETRY; attempt++) {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) {
      const json = JSON.parse(text);
      if (json.status !== 'success') throw new Error(`PDF 解鎖失敗: ${json.message}`);
      return json.data;
    }

    if (code === 503) {
      Logger.log(`⏳ Cloud Run 冷啟動中，第 ${attempt}/${MAX_RETRY} 次，等待 ${RETRY_WAIT_MS/1000} 秒後重試...`);
      if (attempt < MAX_RETRY) Utilities.sleep(RETRY_WAIT_MS);
      continue;
    }

    // 其他錯誤直接拋出，不重試
    throw new Error(`Cloud Run Error (${code}): ${text}`);
  }

  throw new Error(`Cloud Run 服務無法連線，已重試 ${MAX_RETRY} 次，請確認服務是否正常運作。`);
}

function parseCloudRunData_(rawData, tz) {
  if (!rawData || !rawData.length) return [];
  
  let defaultDate = '';
  const dateRegex = /(\d{3,4})[\/年](\d{1,2})[\/月](\d{1,2})/;
  
  for (let i = 0; i < Math.min(rawData.length, 20); i++) {
    const rowStr = rawData[i].join('');
    const m = rowStr.match(dateRegex);
    if (m) {
      let y = parseInt(m[1]);
      if (y < 1911) y += 1911; 
      defaultDate = `${y}/${('0'+m[2]).slice(-2)}/${('0'+m[3]).slice(-2)}`;
      break;
    }
  }

  const colMap = {
    date: ['成交日期', '交易日期'], time: ['成交時間', '時間'],
    sym:  ['股票代碼', '股票代號', '股號', '商品代碼', '商品名稱'],
    name: ['股票名稱', '商品名稱', '名稱'], side: ['類別', '交易別', '買賣別', '種類'],
    qty:  ['成交股數', '股數', '數量'], price:['單價', '成交單價', '成交價'],
    amt:  ['成交金額', '價金', '金額'], ord:  ['委託書號', '委託單號', '書號']
  };

  const stopKeywords = [
    '庫存明細', '集保明細', '當日收盤價', '集保股數', '集保市值', '理財資訊', '個股分析', '我要下單'
  ];

  const out = [];
  let headerIdx = -1;
  let headers = [];

  for (let i = 0; i < rawData.length; i++) {
    const rowStr = rawData[i].join('').replace(/\s/g, '');
    const hasSym  = colMap.sym.some(k => rowStr.includes(k));
    const hasSide = colMap.side.some(k => rowStr.includes(k));
    const hasPrc  = colMap.price.some(k => rowStr.includes(k));

    if (hasSym && hasSide && hasPrc) {
      headerIdx = i;
      headers = rawData[i].map(s => String(s).trim());
      break;
    }
  }

  if (headerIdx === -1) return [];

  const idx = {};
  for (const [key, keywords] of Object.entries(colMap)) {
    idx[key] = headers.findIndex(h => keywords.some(k => h.includes(k)));
  }

  for (let i = headerIdx + 1; i < rawData.length; i++) {
    const row = rawData[i];
    const rowStrSimple = row.join('').replace(/\s/g, '');

    if (stopKeywords.some(kw => rowStrSimple.includes(kw))) break;
    if (row.length < headers.length - 3) continue;
    if (rowStrSimple.includes('合計') || rowStrSimple.includes('總計')) continue;

    const sideRaw = (idx.side >= 0) ? String(row[idx.side]).trim() : '';
    if (!sideRaw.includes('買') && !sideRaw.includes('賣') && !sideRaw.includes('沖')) continue;

    let dateStr = defaultDate; 
    if (idx.date >= 0 && row[idx.date]) {
        const d = toYMDslash_(row[idx.date]);
        if (d.length >= 5) dateStr = d;
    }
    if (!dateStr) continue;

    const timeStr = (idx.time >= 0 && row[idx.time]) ? normTime_(row[idx.time]) : '00:00:00';
    let symRaw = (idx.sym >= 0) ? String(row[idx.sym]).trim() : '';
    let nameRaw = (idx.name >= 0) ? String(row[idx.name]).trim() : '';
    if (!symRaw && nameRaw) symRaw = nameRaw; 
    if (!nameRaw && symRaw) nameRaw = symRaw; 
    const sym = symRaw.replace(/\s+/g, ''); 
    const name = nameRaw.replace(/\s+/g, '');

    const qty     = (idx.qty >= 0) ? num_(row[idx.qty]) : 0;
    const price   = (idx.price >= 0) ? num_(row[idx.price]) : 0;
    const amount  = (idx.amt >= 0) ? num_(row[idx.amt]) : Math.round(qty * price);
    const ord     = (idx.ord >= 0) ? String(row[idx.ord]).trim() : '';

    if (!qty && !amount) continue; 

    const feeCol = headers.findIndex(h => h.includes('手續費'));
    const taxCol = headers.findIndex(h => h.includes('交易稅'));
    const fee = (feeCol >= 0) ? num_(row[feeCol]) : calcFee_(amount);
    const tax = (taxCol >= 0) ? num_(row[taxCol]) : calcTax_(sideRaw, amount);
    const net = calcNet_(sideRaw, amount, fee, tax);

    out.push([dateStr, timeStr, sym, name, sideRaw, qty, price, amount, ord, fee, tax, net, 'CloudRun']);
  }
  return out;
}

function ingestFromGmail_Plaintext() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000))
    throw new Error('ingestFromGmail_Plaintext: 另一個執行緒正在執行，請稍後再試');
  try {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 寫入目標：待確認交易（staging）
  const shStaging = ensureStagingSheet_();

  // 去重來源：同時檢查「交易紀錄」+ 「待確認交易」，避免重複送進 staging
  const shTrades = ensureSheetWithHeader_(C.SHEET_TRADES, [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','証券商','備註'
  ]);

  const nameToCodeMap = new Map();
  const dcaWhitelist = new Set();

  for (let i=1; i<=10; i++) {
    const sym = String(C[`DCA_${i}_SYMBOL`]||'').trim();
    const nm  = String(C[`DCA_${i}_NAME`]||'').trim().replace(/\s+/g,'');
    if (sym) {
      dcaWhitelist.add(sym);
      if (nm) nameToCodeMap.set(nm, sym);
    }
  }

  const learnFromSheet = (sheetName) => {
    const s = ss.getSheetByName(sheetName);
    if (!s || s.getLastRow() < 2) return;
    const d = s.getRange(2, 1, s.getLastRow()-1, 4).getValues();
    d.forEach(r => {
      const c = String(r[2]||'').trim();
      const n = String(r[3]||'').trim().replace(/\s+/g,'');
      if (c && n && !nameToCodeMap.has(n)) nameToCodeMap.set(n, c);
    });
  };
  learnFromSheet(C.SHEET_TRADES);
  learnFromSheet(C.SHEET_OPENING);
  learnFromSheet(C.SHEET_HOLD);
  // 從待確認交易補學：只採信代碼為數字的項目，避免名稱誤解析的髒資料污染 map
  const stagingShTmp = ss.getSheetByName(C.SHEET_STAGING || '待確認交易');
  if (stagingShTmp && stagingShTmp.getLastRow() > 1) {
    stagingShTmp.getRange(2, 1, stagingShTmp.getLastRow()-1, 4).getValues().forEach(r => {
      const c = String(r[2]||'').trim(), n = String(r[3]||'').trim().replace(/\s+/g,'');
      if (c && n && !isNaN(Number(c)) && !nameToCodeMap.has(n)) nameToCodeMap.set(n, c);
    });
  }

  const existedOrder = new Set();
  const existedNoOrd = new Set();

  // 去重用的代碼正規化
  const normSymForDedup_ = (s) => {
    const t = String(s||'').trim();
    if (!t || !isNaN(Number(t))) return t;
    // 完全比對
    if (nameToCodeMap.has(t)) return nameToCodeMap.get(t);
    // 去掉尾端 * 等特殊字元後再比對（如「國巨*」→「國巨」）
    const stripped = t.replace(/[*＊·•！]+$/, '').trim();
    if (stripped && stripped !== t && nameToCodeMap.has(stripped)) return nameToCodeMap.get(stripped);
    // 開頭是 4~6 位數字就提取為代碼（如「2327國巨*」→「2327」）
    const numPrefix = t.match(/^(\d{4,6})/);
    if (numPrefix) return numPrefix[1];
    return t;
  };
  // 去重用的類別正規化：現買/沖買/集買 → 買；現賣/沖賣 → 賣
  // 避免同一筆交易因不同信件的類別標記不同（如「現買」vs「沖買」）而被當成兩筆
  const normTypeForDedup_ = (type) => String(type||'').includes('賣') ? '賣' : '買';

  const loadDedup_ = (sh) => {
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    for (const r of data) {
      const sym  = normSymForDedup_(r[2]);
      const ord  = normOrderNo_(r[8]);
      const date = String(r[0]||'').trim();
      if (ord) existedOrder.add(makeKey_Order_(date, ord, sym));
      else {
        const typeNorm = normTypeForDedup_(r[4]);
        existedNoOrd.add([date, normTime_(r[1]||''), sym, typeNorm, r[5], r[6]].join('|'));
      }
    }
  };
  loadDedup_(shTrades);
  loadDedup_(shStaging);

  const since = new Date();
  since.setDate(since.getDate() - Number(C.GMAIL_QUERY_DAYS || 7));
  const tz = C.TZ || 'Asia/Taipei';
  const after = Utilities.formatDate(since, tz, 'yyyy/MM/dd');

  let parsed = [];
  const logMsg = [];
  // HTML 信件中出現過的委託單號，PDF 遇到相同單號時跳過，避免一般交易重複
  const htmlOrderSet = new Set();

  // 1. 先處理 HTML 信件，建立已知委託單號集合
  const labelHtml = (C.GMAIL_LABEL_HTML || '').trim();
  if (labelHtml) {
    const qHtml = `label:${labelHtml} after:${after}`;
    const threadsHtml = GmailApp.search(qHtml, 0, 30);
    threadsHtml.forEach(t => t.getMessages().forEach(m => {
      const html = m.getBody();
      const body = m.getPlainBody();
      let rows = parseBrokerMailHTMLTable_CN_(html, m, tz);
      if (!rows.length) rows = parseBrokerMailPlainTable_CN_(body, m, tz);
      rows.forEach(r => { const ord = normOrderNo_(r[8]); if (ord) htmlOrderSet.add(ord); });
      if (rows.length > 0) parsed = parsed.concat(rows);
    }));
  }

  // 2. 再處理 PDF 附件，跳過委託單號已在 HTML 出現過的（一般交易），其餘全保留（定期定額等）
  const labelPdf = (C.GMAIL_LABEL_PDF || '').trim();
  if (labelPdf) {
    const qPdf = `label:${labelPdf} after:${after} has:attachment`;
    const threadsPdf = GmailApp.search(qPdf, 0, 30);
    threadsPdf.forEach(t => t.getMessages().forEach(m => {
      const atts = m.getAttachments();
      atts.forEach(att => {
        if (att.getContentType() === "application/pdf" || att.getName().toLowerCase().endsWith(".pdf")) {
          try {
            const rawTable = callCloudRunToUnlock_(att, C.ID_NUMBER);
            const rows = parseCloudRunData_(rawTable, tz);
            const keptRows = [];
            rows.forEach(r => {
              let sym = String(r[2]||'').trim();
              // 嘗試 nameToCodeMap 對應（如 DCA_i_NAME 設定）
              if (isNaN(Number(sym)) && nameToCodeMap.has(sym)) {
                sym = nameToCodeMap.get(sym); r[2] = sym;
              }
              // 退而求其次：從基金名稱提取 4~6 位數字代碼（如「國泰永續高股息00919」→「00919」）
              if (isNaN(Number(sym))) {
                const codeMatch = sym.match(/(\d{4,6})/);
                if (codeMatch) { sym = codeMatch[1]; r[2] = sym; }
              }
              const ord = normOrderNo_(r[8]);
              if (ord && htmlOrderSet.has(ord)) {
                Logger.log(`[PDF skip] 委託單號="${ord}" 已在 HTML 信件中，跳過`);
                return;
              }
              r[12] = 'CloudRun_SIP';
              keptRows.push(r);
            });
            if (keptRows.length > 0) {
              parsed = parsed.concat(keptRows);
              logMsg.push(`[PDF] ${m.getSubject()} (${keptRows.length}筆)`);
            }
          } catch (e) {
            Logger.log(`PDF失敗: ${e.message}`);
          }
        }
      });
    }));
  }

  // 合併前先正規化代碼，確保同委託單號的「國巨*」與「2327」能在 consolidate 時被合併
  parsed.forEach(r => {
    const sym = normSymForDedup_(String(r[2]||'').trim());
    if (sym !== String(r[2]||'').trim()) r[2] = sym;
  });

  const consolidated = consolidateByOrderNo_(parsed);
  const toWrite = [];

  for (const r of consolidated) {
    // 正規化代碼：若為非數字名稱且有對應，直接修正 r[2]，讓 staging 也拿到正確代碼
    const rawSym = String(r[2]||'').trim();
    if (rawSym && isNaN(Number(rawSym)) && nameToCodeMap.has(rawSym)) {
      r[2] = nameToCodeMap.get(rawSym);
    }
    const sym = String(r[2]||'').trim();
    const ord = normOrderNo_(r[8]);
    if (ord) {
      const k = makeKey_Order_(r[0], ord, sym);
      if (existedOrder.has(k)) continue;
      existedOrder.add(k);
      toWrite.push(r);
    } else {
      const typeNorm = normTypeForDedup_(r[4]);
      const k = [String(r[0]||'').trim(), normTime_(r[1]||''), sym, typeNorm, r[5], r[6]].join('|');
      if (existedNoOrd.has(k)) continue;
      existedNoOrd.add(k);
      toWrite.push(r);
    }
  }

  if (toWrite.length) {
    // 轉換為 staging 格式（15 欄）：前 12 欄不變，插入空的証券商，保留備註，加確認狀態
    const stagingRows = toWrite.map(r => [
      ...r.slice(0, 12),  // 成交日期..淨收付金額
      '',                  // 証券商（email 無法得知）
      r[12] || '',         // 備註（CloudRun 等標記）
      '待確認',            // 確認狀態
    ]);
    const startRow = shStaging.getLastRow() + 1;
    const numRows  = stagingRows.length;
    shStaging.getRange(startRow, 3, numRows, 1).setNumberFormat('@'); // 股票代碼
    shStaging.getRange(startRow, 9, numRows, 1).setNumberFormat('@'); // 委託單號
    shStaging.getRange(startRow, 1, numRows, stagingRows[0].length).setValues(stagingRows);
  }
  Logger.log(`總計新增：${toWrite.length} 筆交易至待確認`);
  } finally {
    lock.releaseLock();
  }
}

function parseBrokerMailHTMLTable_CN_(html, msg, tz) {
  if (!html) return [];
  const out = [];
  const rawDate = (() => {
    const m = html.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
    if (!m) return '';
    return `${m[1]}-${('0'+m[2]).slice(-2)}-${('0'+m[3]).slice(-2)}`;
  })() || Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd');
  const dateStr = toYMDslash_(rawDate);

  const tables = html.match(/<table[\s\S]*?<\/table>/gi) || [];
  const wantHeaders = ['成交時間','委託單號','股號','股票代號','股票名稱','類別','買賣別','股數','單價','成交價','價金','成交金額'];

  for (const tbl of tables) {
    const trs = tbl.match(/<tr[\s\S]*?<\/tr>/gi) || [];
    if (!trs.length) continue;
    const headerCells = extractCellsText_(trs[0]);
    if (!headerCells.length) continue;
    const hits = wantHeaders.filter(h => headerCells.join('|').indexOf(h) >= 0).length;
    if (hits < 6) continue;

    const idx = {
      time:   indexOfLike_(headerCells, ['成交時間','時間']),
      orderNo:indexOfLike_(headerCells, ['委託單號','委託書號','委託單號碼']),
      symbol: indexOfLike_(headerCells, ['股號','股票代號','商品代碼']),
      name:   indexOfLike_(headerCells, ['股票名稱','商品名稱','名稱']),
      side:   indexOfLike_(headerCells, ['成交類別','類別','買賣別','方向']),
      qty:    indexOfLike_(headerCells, ['股數','數量','成交股數']),
      price:  indexOfLike_(headerCells, ['成交價','單價']),
      amount: indexOfLike_(headerCells, ['成交金額','價金'])
    };

    for (let i=1;i<trs.length;i++){
      const cells = extractCellsText_(trs[i]).map(s => s.replace(/&nbsp;/g,'').trim());
      if (!cells.length) continue;
      const time    = normTime_(safePick_(cells, idx.time));
      const orderNo = safePick_(cells, idx.orderNo);
      const symbol  = safePick_(cells, idx.symbol);
      const name    = safePick_(cells, idx.name);
      const sideRaw = keepSide_(safePick_(cells, idx.side));
      const qty     = num_(safePick_(cells, idx.qty));
      const price   = num_(safePick_(cells, idx.price));
      const amount  = num_(safePick_(cells, idx.amount));

      if (!symbol || !time || !sideRaw || !qty || !price) continue;
      const fee = calcFee_(amount);
      const tax = calcTax_(sideRaw, amount);
      const net = calcNet_(sideRaw, amount, fee, tax);
      out.push([ dateStr, time, symbol, name, sideRaw, qty, price, amount, orderNo, fee, tax, net, '' ]);
    }
    if (out.length) break;
  }
  return out;
}

function parseBrokerMailPlainTable_CN_(text, msg, tz) {
  if (!text) return [];
  const out = [];
  const rawDate = (() => {
    const m = text.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
    if (!m) return '';
    return `${m[1]}-${('0'+m[2]).slice(-2)}-${('0'+m[3]).slice(-2)}`;
  })() || Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd');
  const dateStr = toYMDslash_(rawDate);

  const headerRe = /成交時間\s+委託單號\s+(?:股號|股票代號)\s+股票名稱\s+(?:成交類別|類別|買賣別)\s+股數\s+(?:成交價|單價)\s+(?:成交金額|價金)(?:\s+\S+)?/;
  const idx = text.search(headerRe);
  if (idx < 0) return out;

  const tail = text.slice(idx).split(/\r?\n/).slice(1);
  const lineRe = /^\s*([0-9:]{5,8})\s+(\S+)\s+(\d{3,6})\s+(\S+)\s+(\S+)\s+([\d,]+)\s+([\d,\.]+)\s+([\d,\.]+)(?:\s+\S+)?/;

  for (const line of tail) {
    const m = line.match(lineRe);
    if (!m) { if (/^\s*$/.test(line)) break; continue; }
    const time    = normTime_(m[1]);
    const orderNo = m[2];
    const symbol  = m[3];
    const name    = m[4];
    const sideRaw = keepSide_(m[5]);
    const qty     = num_(m[6]);
    const price   = num_(m[7]);
    const amount  = num_(m[8]);
    const fee = calcFee_(amount);
    const tax = calcTax_(sideRaw, amount);
    const net = calcNet_(sideRaw, amount, fee, tax);
    out.push([ dateStr, time, symbol, name, sideRaw, qty, price, amount, orderNo, fee, tax, net, '' ]);
  }
  return out;
}

function extractCellsText_(trHtml) {
  const reg = /<(?:td|th)[^>]*>([\s\S]*?)<\/(?:td|th)>/gi;
  const cells = []; let m;
  while ((m = reg.exec(trHtml)) !== null) cells.push(cleanHtmlText_(m[1]));
  return cells;
}
function cleanHtmlText_(s) {
  return String(s||'').replace(/<br\s*\/?>/gi,' ').replace(/<[^>]*>/g,'').replace(/&nbsp;/g,' ').trim();
}
function indexOfLike_(arr, keys){ for (const k of keys){ const i=arr.findIndex(x=>x.indexOf(k)>=0); if (i>=0) return i; } return -1; }
function safePick_(arr,i){ return (i>=0 && i<arr.length)?arr[i]:''; }
function keepSide_(s){ return (s==null)?'':String(s).trim(); }

/***** =======================
 * Part 3/4
 * 庫存紀錄
 * ======================== */

function rebuildAll_B() {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const headersOpen = ['股票代碼','股票名稱','買進日期','買入價','持有股數','買入成本 (單純買入價*股數)','手續費','證券商'];
  const headersHold = ['股票代碼','股票名稱','買進日期','買入價','持有股數','買入成本 (單純買入價*股數)','手續費','證券商','現價','即時市值'];
  
  const shOpen = ensureSheetWithHeader_(C.SHEET_OPENING, headersOpen);
  const shH     = ensureSheetWithHeader_(C.SHEET_HOLD, headersHold);
  const shT     = ensureSheetWithHeader_(C.SHEET_TRADES, [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','證券商','備註'
  ]);

  // === 輔助函式：標準化時間戳記 (確保排序絕對正確) ===
  const getTs_ = (d, t) => {
    // 處理日期
    let dateStr = "";
    if (d instanceof Date) {
      dateStr = Utilities.formatDate(d, 'GMT+8', 'yyyy/MM/dd');
    } else {
      dateStr = String(d).replace(/\-/g, '/').trim();
    }
    // 處理時間
    let timeStr = "00:00:00";
    if (t instanceof Date) {
      timeStr = Utilities.formatDate(t, 'GMT+8', 'HH:mm:ss');
    } else if (t && String(t).trim()) {
      timeStr = String(t).trim();
    }
    return new Date(`${dateStr} ${timeStr}`).getTime();
  };
  // ===============================================

  // 1. 優先處理：用賣單沖銷期初庫存 (針對持有中但尚未在此系統記錄買入的)
  (function applySalesToOpening_inline() {
    const openRows = readTableAsObjects_(shOpen);
    const openMap = new Map();
    openRows.forEach(o=>{
      const sym = String(o['股票代碼']).trim(); if (!sym) return;
      // Key 加上券商，確保分券商沖銷
      const key = sym + '|' + String(o['證券商']||'').trim();
      openMap.set(key, {
        name: String(o['股票名稱']||'').trim(),
        buyDate: toYMDslash_(o['買進日期']||''),
        buyPrice: Number(o['買入價']||0),
        qty: Number(o['持有股數']||0),
        fee: Number(o['手續費']||0),
        broker: String(o['證券商']||'')
      });
    });
    
    if (openMap.size > 0) {
        const start = (C.HOLDINGS_START_DATE||'').trim();
        const sales = readTableAsObjects_(shT)
          .filter(r => String(r['成交類別']||'').includes('賣'))
          .filter(r => !start || String(r['成交日期']||'') >= start)
          .map(r => {
             // 這裡也要加上 timestamp 方便排序
             r._ts = getTs_(r['成交日期'], r['成交時間']);
             return r;
          })
          .sort((a,b)=> a._ts - b._ts); // 絕對時間排序

        sales.forEach(r=>{
          const sym = String(r['股票代碼']||'').trim();
          const broker = String(r['證券商']||'').trim();
          const sellQty = Number(r['股數']||0);
          
          if (!sym || sellQty<=0) return;
          
          // 嘗試找對應券商的期初庫存
          const key = sym + '|' + broker;
          const o = openMap.get(key);
          
          if (!o || o.qty<=0) return;
          
          const use = Math.min(o.qty, sellQty);
          const prop = (o.qty>0) ? (use/o.qty) : 0;
          o.qty -= use;
          o.fee = Math.round(o.fee * (1 - prop));
        });

        shOpen.clear(); shOpen.appendRow(headersOpen);
        const out=[];
        openMap.forEach((v,k)=>{
          // 只有當數量 > 0 才寫回，否則代表已完全沖銷
          if(Math.round(v.qty) > 0) {
            const buyCost = v.buyPrice * v.qty;
            out.push([k.split('|')[0], v.name, toYMDslash_(v.buyDate)||'', round2_(v.buyPrice), Math.round(v.qty), round2_(buyCost), Math.round(v.fee), v.broker]);
          }
        });
        // 確保格式正確
        if (out.length) {
            shOpen.getRange(2,1,out.length,out[0].length).setValues(out);
            shOpen.getRange(2,1,out.length,1).setNumberFormat('@'); // 代碼設為文字
        }
    }
  })();

  // 2. 重建目前庫存 (FIFO)
  const start = (C.HOLDINGS_START_DATE||'').trim();
  
  // 讀取所有交易並加上 Timestamp
  const txAll = readTableAsObjects_(shT)
    .filter(r => !start || String(r['成交日期']||'') >= start)
    .map(r => ({
      sym:   String(r['股票代碼']||'').trim(),
      name:  String(r['股票名稱']||'').trim(),
      date:  toYMDslash_(r['成交日期']||''),
      time:  normTime_(String(r['成交時間']||'')),
      ts:    getTs_(r['成交日期'], r['成交時間']), // 加入 Timestamp
      side:  String(r['成交類別']||''),
      qty:   Number(r['股數']||0),
      price: Number(r['成交價']||0),
      amt:   Number(r['成交金額']||0),
      fee:   Number(r['手續費']||0) || R4_calcFee_(Number(r['成交金額']||0)),
      tax:   Number(r['交易稅']||0)  || R4_calcTax_(String(r['成交類別']||''), Number(r['成交金額']||0)),
      broker: String(r['證券商']||'')
    }))
    .filter(x => x.sym && x.date && x.qty>0);

  // === 關鍵修正：依照 Timestamp 排序，保證先買後賣 ===
  txAll.sort((a,b)=> {
    if (a.sym !== b.sym) return a.sym.localeCompare(b.sym);
    return a.ts - b.ts;
  });

  const lotsByKey = new Map();

  // 載入期初庫存 (已扣除完畢的剩餘量)
  readTableAsObjects_(ensureSheetWithHeader_(C.SHEET_OPENING, [])).forEach(r=>{
    const sym = String(r['股票代碼']||'').trim(); if (!sym) return;
    const qty = Math.round(Number(r['持有股數']||0)); if (qty<=0) return;
    const name = String(r['股票名稱']||'').trim();
    const buyDate = toYMDslash_(r['買進日期']||'') || '1900/01/01';
    const buyPrice = Number(r['買入價']||0);
    const buyFeePerShare = (qty>0) ? (Number(r['手續費']||0)/qty) : 0;
    const broker = String(r['證券商']||'').trim(); 
    
    const key = `${sym}|${broker}`;
    if (!lotsByKey.has(key)) lotsByKey.set(key, []);
    lotsByKey.get(key).push({ sym, name, buyDate, buyPrice, buyFeePerShare, remainQty: qty, broker });
  });

  // 處理交易紀錄
  txAll.forEach(x=>{
    const key = `${x.sym}|${x.broker}`;
    
    // 忽略當沖
    if (isDayLoopSide_(x.side)) return;

    if (isStockBonusSide_(x.side)) {
      // 股票股利 (忽略成本)
      if (!lotsByKey.has(key)) lotsByKey.set(key, []);
      lotsByKey.get(key).push({ sym: x.sym, name:x.name, buyDate:x.date, buyPrice:0, buyFeePerShare:0, remainQty: Math.floor(x.qty), broker: x.broker });
      return;
    }
    
    if (x.side.includes('買') && x.price>=0) { // 允許 0 元買入 (手動補配股)
      if (!lotsByKey.has(key)) lotsByKey.set(key, []);
      const perShareFee = (x.qty>0) ? (x.fee / x.qty) : 0;
      lotsByKey.get(key).push({ sym: x.sym, name:x.name, buyDate:x.date, buyPrice:x.price, buyFeePerShare: perShareFee, remainQty: Math.round(x.qty), broker: x.broker });
    }
    else if (x.side.includes('賣')) {
      // 賣出扣抵 (FIFO)
      let remain = Math.round(x.qty);
      const q = lotsByKey.get(key) || [];
      
      // 清除已耗盡的批次
      while (q.length && q[0].remainQty <= 0.001) q.shift();

      if (q.length === 0) {
        // 這裡就是報錯的地方，但因為我們已經修正了排序，理論上不會再發生
        Logger.log(`⚠ [Warning] ${x.date} ${x.sym} (${x.broker}) 賣出 ${remain} 股時庫存不足 (已忽略短缺部分)`);
      } else {
        while (remain > 0 && q.length > 0) {
           const lot = q[0];
           const part = Math.min(remain, lot.remainQty);
           lot.remainQty -= part;
           remain -= part;
           if (lot.remainQty <= 0.001) q.shift();
        }
      }
    }
  });

  // 3. 匯總結果
  const grouped = new Map(); 
  for (const [key, lots] of lotsByKey.entries()) {
    lots.filter(l => l.remainQty > 0.1).forEach(l => {
      const outKey = `${l.sym}|${l.buyDate}|${l.broker}`;
      if (!grouped.has(outKey)) grouped.set(outKey, {
        sym: l.sym || key.split('|')[0], name: l.name || '', buyDate: l.buyDate,
        sumQty: 0, sumPxQty: 0, sumFee: 0, broker: l.broker || ''
      });
      const g = grouped.get(outKey);
      g.sumQty += Math.round(l.remainQty);
      g.sumPxQty += (l.buyPrice || 0) * Math.round(l.remainQty);
      g.sumFee += (l.buyFeePerShare || 0) * Math.round(l.remainQty);
      if (!g.name && l.name) g.name = l.name;
    });
  }

  const out = Array.from(grouped.values())
    .sort((a,b) => (String(a.sym||'').localeCompare(String(b.sym||'')) || String(a.buyDate||'').localeCompare(String(b.buyDate||''))))
    .map(g => {
      const avgBuy = g.sumQty > 0 ? round2_(g.sumPxQty / g.sumQty) : 0;
      return [g.sym, g.name, g.buyDate || '', avgBuy, Math.round(g.sumQty), round2_(avgBuy * Math.round(g.sumQty)), Math.round(g.sumFee), g.broker || '', '', ''];
    });

  // 4. 寫入庫存紀錄
  shH.clear(); shH.appendRow(headersHold);
  if (out.length) {
      shH.getRange(2, 1, out.length, out[0].length).setValues(out);
      shH.getRange(2, 1, out.length, 1).setNumberFormat('@'); // 代碼補零保護
      
      const formulas = [];
      for (let r = 2; r <= out.length + 1; r++) {
         const f_price = `=IFERROR(GOOGLEFINANCE("TPE:" & A${r}), IFERROR(IMPORTXML("https://finance.yahoo.com/quote/" & A${r} & ".TWO", "//fin-streamer[@data-field='regularMarketPrice']"), 0))`;
         const f_market = `=E${r} * I${r}`; 
         formulas.push([f_price, f_market]);
      }
      shH.getRange(2, 9, formulas.length, 2).setFormulas(formulas);
      shH.getRange(2, 4, out.length, 7).setNumberFormat("#,##0.00");
    }
    Logger.log(`Holdings updated: ${out.length} rows`);
}
function readTableAsObjects_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow<2) return [];
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||'').trim());
  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  return data.map(row => {
    const o={}; headers.forEach((h,i)=>o[h]=row[i]); return o;
  });
}

/***** =======================
 * Part 4/4 股利與其他
 * ======================== */

function fetchStockNameOnce_(sym, token) {
  if (!sym) return '';
  try {
    const props = PropertiesService.getScriptProperties();
    const KEY = 'NAME_CACHE__' + String(sym).trim();
    const cached = props.getProperty(KEY);
    if (cached) return cached;
    const url = 'https://api.finmindtrade.com/api/v4/data';
    const headers = token ? { 'Authorization': 'Bearer ' + token } : {};
    const qs = 'dataset=TaiwanStockInfo&data_id=' + encodeURIComponent(sym);
    const resp = UrlFetchApp.fetch(url + '?' + qs, { muteHttpExceptions: true, followRedirects: true, headers });
    if (resp.getResponseCode() !== 200) return '';
    const json = JSON.parse(resp.getContentText('utf-8'));
    const rows = Array.isArray(json.data) ? json.data : [];
    const hit = rows.find(x => x && (x.stock_name || x.company_name || x.name));
    const nm = hit ? String(hit.stock_name || hit.company_name || hit.name || '').trim() : '';
    if (nm) props.setProperty(KEY, nm);
    return nm || '';
  } catch (e) { return ''; }
}

function buildNameCache_(shHold, shOpen, shT) {
  const cache = new Map();
  const push = (sym, nm) => {
    const s = String(sym || '').trim(); const n = String(nm || '').trim();
    if (s && n && !cache.has(s)) cache.set(s, n);
  };
  readTableAsObjects_(shHold).forEach(r => push(r['股票代碼'], r['股票名稱']));
  readTableAsObjects_(shOpen).forEach(r => push(r['股票代碼'], r['股票名稱']));
  readTableAsObjects_(shT).forEach(r => push(r['股票代碼'], r['股票名稱']));
  return cache;
}

function updateDividendsFromFinMind() {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shHold = ensureSheetWithHeader_(C.SHEET_HOLD, []);
  const shOpen = ensureSheetWithHeader_(C.SHEET_OPENING, []);
  const shT     = ensureSheetWithHeader_(C.SHEET_TRADES, []);
  
  const headers = [
    '股票代碼','股票名稱','股利所屬年度','除息日','現金股利發放日',
    '現金股利 (元/股)','除息日持股數','現金總股息',
    '實際領取金額 (扣除每筆手續費10元)','平均成交價','個人殖利率 (現金/買入價)','備註'
  ];
  const shOut = ensureSheetWithHeader_(C.SHEET_DIV, headers);

  // 1. 建立現有資料快取 (防止重複) - 修正比對邏輯
  const existingDivKeys = new Set();
  const lastRowDiv = shOut.getLastRow();
  if (lastRowDiv > 1) {
    const divData = shOut.getRange(2, 1, lastRowDiv - 1, 4).getValues();
    divData.forEach(r => {
      let s = String(r[0]||'').trim();
      // ★ 修正比對邏輯：3位數補成5碼，2位數補成4碼
      if (/^\d{1,3}$/.test(s)) {
         const n = Number(s);
         s = (n < 100) ? ('0000' + n).slice(-4) : ('00000' + n).slice(-5);
      }
      const d = toYMDslash_(r[3]||'');
      if (s && d) existingDivKeys.add(`${s}|${d}`);
    });
  }

  const nameBySym = buildNameCache_(shHold, shOpen, shT);
  const openingQty = new Map();
  readTableAsObjects_(shOpen).forEach(o => {
    const sym = String(o['股票代碼'] || '').trim(); if (sym) openingQty.set(sym, Number(o['持有股數'] || 0));
  });

  const start = (C.DIV_START_DATE || '').trim();
  const tradeRows = readTableAsObjects_(shT);
  const txBySym = new Map();
  tradeRows.forEach(r => {
    const sym = String(r['股票代碼'] || '').trim(); if (!sym) return;
    const side= String(r['成交類別'] || '');
    const date= toYMDslash_(String(r['成交日期'] || '').trim());
    const qty = Number(r['股數'] || 0);
    if (!date || !(qty > 0)) return;
    if (start && date < start) return;
    if (side.includes('沖買') || side.includes('沖賣')) return;
    const delta = (isStockBonusSide_(side) || side.includes('買')) ? qty : (side.includes('賣') ? -qty : 0);
    if (delta === 0) return;
    if (!txBySym.has(sym)) txBySym.set(sym, []);
    txBySym.get(sym).push({ date, deltaQty: delta });
  });
  for (const [sym, arr] of txBySym.entries()) arr.sort((a, b) => a.date.localeCompare(b.date));

  const tradeSyms = Array.from(new Set(tradeRows.map(r => String(r['股票代碼'] || '').trim()).filter(Boolean)));
  const allSymbols = Array.from(new Set([...Array.from(openingQty.keys()), ...Array.from(nameBySym.keys()), ...tradeSyms])).filter(Boolean);
  const yearFrom = String(C.DIV_YEAR_FROM || (new Date().getFullYear() - 8));
  const hadPosOrBuy = new Set();
  tradeRows.forEach(r => {
    const d = toYMDslash_(String(r['成交日期'] || '')); const y = d ? d.slice(0, 4) : '';
    const sym = String(r['股票代碼'] || '').trim(); const side= String(r['成交類別'] || '');
    if (!sym || !y || y < yearFrom) return;
    if (side.includes('買') || isStockBonusSide_(side)) hadPosOrBuy.add(sym);
  });
  const filteredSymbols = allSymbols.filter(sym => {
    const openQ = Number(openingQty.get(sym) || 0);
    if (openQ > 0) return true;
    return hadPosOrBuy.has(sym);
  }).sort();

  const perRun = Math.max(1, Number(C.DIV_SYMBOLS_PER_RUN || 10));
  const props = PropertiesService.getScriptProperties();
  const CUR_KEY = 'DIV_CURSOR'; const TS_KEY  = 'DIV_CURSOR_TS';
  const staleDays = Math.max(1, Number(C.DIV_CURSOR_RESET_IF_STALE_DAYS || 3));
  const lastTs = Number(props.getProperty(TS_KEY) || 0);
  const inCycle = PropertiesService.getScriptProperties().getProperty('DIV_CYCLE_ACTIVE') === '1';
  if (!inCycle && lastTs && (Date.now() - lastTs) > staleDays * 86400000) props.deleteProperty(CUR_KEY);

  let cur = Number(props.getProperty(CUR_KEY) || 0);
  if (cur >= filteredSymbols.length) cur = 0;
  
  if (props.getProperty(CUR_KEY) === null && shOut.getLastRow() === 0) {
      shOut.appendRow(headers);
  }

  const slice = filteredSymbols.slice(cur, cur + perRun);
  if (!slice.length) { Logger.log('DIV: 無可處理之代號'); return; }

  const rowsToAppend = [];

  for (const sym of slice) {
    const list = fetchFinMind_Dividends_(sym, C.FINMIND_TOKEN, C.DIV_YEAR_FROM);
    list.sort((a, b) => (a.exDate || '').localeCompare(b.exDate || ''));
    
    const openQ = openingQty.get(sym) || 0;
    const txs   = txBySym.get(sym) || [];
    
    const qtyOnDateBase = (function (openQty, txs) {
      const cache = new Map();
      return function (dateStr) {
        if (!dateStr) return Math.max(0, Math.round(openQty));
        if (cache.has(dateStr)) return cache.get(dateStr);
        let qty = openQty;
        for (const t of (txs || [])) { if (t.date <= dateStr) qty += t.deltaQty; else break; }
        qty = Math.max(0, Math.round(qty));
        cache.set(dateStr, qty);
        return qty;
      };
    })(openQ, txs);

    let bonusCarry = 0; 

    list.forEach(rec => {
      const exDate = rec.exDate; if (!exDate) return;
      
      // ★★★ 修正點：比對前先標準化 CheckSym (2碼->4碼, 3碼->5碼) ★★★
      let checkSym = sym;
      if (/^\d{1,3}$/.test(checkSym)) {
         const n = Number(checkSym);
         checkSym = (n < 100) ? ('0000' + n).slice(-4) : ('00000' + n).slice(-5);
      }
      
      if (existingDivKeys.has(`${checkSym}|${exDate}`)) return;

      const cashPerShare = Number(rec.cashPerShare || 0);
      const baseQty = qtyOnDateBase(exDate);
      const qtyAtEx = Math.max(0, Math.round(baseQty)); 
      if (qtyAtEx <= 0) return;

      const cashTotal     = round2_(qtyAtEx * cashPerShare);
      const feePerPayout = Number(C.DIV_CASH_FEE_PER_PAYOUT || 10);
      const actualCash    = cashTotal > 0 ? Math.max(0, round2_(cashTotal - feePerPayout)) : 0;

      let avgBuyAtEx = 0, myYield = '';
      if (cashPerShare > 0) {
        avgBuyAtEx = avgBuyOnDate_(sym, exDate, readTableAsObjects_(shOpen), tradeRows);
        myYield    = (avgBuyAtEx > 0) ? round4_(cashPerShare / avgBuyAtEx) : '';
      }

      // ★★★ 修正點：寫入前標準化 FixedSym (2碼->4碼, 3碼->5碼) ★★★
      let fixedSym = sym;
      if (/^\d{1,3}$/.test(fixedSym)) {
         const n = Number(fixedSym);
         fixedSym = (n < 100) ? ('0000' + n).slice(-4) : ('00000' + n).slice(-5);
      }

      rowsToAppend.push([
        fixedSym,  // 使用修正後的 5 碼代碼
        nameBySym.get(sym) || rec.stockName || '', 
        rec.year || '', 
        exDate, 
        rec.payDate || '',
        cashPerShare,
        qtyAtEx,      
        cashTotal,    
        actualCash,   
        (avgBuyAtEx > 0 ? round2_(avgBuyAtEx) : ''), 
        myYield,      
        ''            
      ]);
    });
    Utilities.sleep(Number(C.DIV_THROTTLE_MS || 1000));
  }

  if (rowsToAppend.length) {
    const startRow = shOut.getLastRow() + 1;
    // 1. 先設定「股票代碼」欄位為純文字
    shOut.getRange(startRow, 1, rowsToAppend.length, 1).setNumberFormat('@');
    
    // 2. 寫入資料
    shOut.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    
    // 3. 設定其他數值格式
    shOut.getRange(startRow, 6, rowsToAppend.length, 1).setNumberFormat('0.00'); // 現金股利
    shOut.getRange(startRow, 8, rowsToAppend.length, 2).setNumberFormat('0.00'); // 總額/實領
    shOut.getRange(startRow, 11, rowsToAppend.length, 1).setNumberFormat('0.00%'); // 殖利率
  }
  
  let next = Number(props.getProperty(CUR_KEY) || 0) + slice.length;
  props.setProperty(CUR_KEY, String(next >= filteredSymbols.length ? 0 : next));
  props.setProperty(TS_KEY,  String(Date.now()));
  Logger.log(`DIV batch done`);
}

function resetDividendBatchCursor() {
  PropertiesService.getScriptProperties().deleteProperty('DIV_CURSOR');
  Logger.log('已重置股利分批游標。');
}

function appendStockBonusToTrades_(rows, C) {
  if (!rows || !rows.length) return;
  const sh = ensureSheetWithHeader_(C.SHEET_TRADES || '交易紀錄', []);
  const grouped = new Map(); 
  rows.forEach(r => {
    const key = [String(r.date).trim(), String(r.sym).trim()].join('|');
    const cur = grouped.get(key) || { date:r.date, sym:r.sym, name:r.name||'', qty:0 };
    cur.qty += Math.floor(Number(r.qty||0));
    grouped.set(key, cur);
  });
  const toUpsert = Array.from(grouped.values()).filter(v => v.qty > 0).map(v => {
      const ord = 'SB' + v.sym + v.date.replace(/\//g,''); 
      return [v.date, '00:00:00', v.sym, v.name, '股票股利', Math.floor(v.qty), 0, 0, ord, 0, 0, 0, '', '由股利寫回'];
  });
  if (!toUpsert.length) return;
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const data = (lastRow > 1) ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
  const existing = new Map(); 
  for (let r=0; r<data.length; r++) {
      if (String(data[r][4]) === '股票股利') existing.set(String(data[r][8]).trim(), r + 2);
  }
  const toAppend = [];
  toUpsert.forEach(row => {
    const hit = existing.get(row[8]);
    if (hit) sh.getRange(hit, 1, 1, row.length).setValues([row]); 
    else toAppend.push(row);
  });
  if (toAppend.length) sh.getRange(sh.getLastRow()+1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
}

function fetchFinMind_Dividends_(stockId, token, yearFromStr) {
  const url = 'https://api.finmindtrade.com/api/v4/data';
  const headers = token ? { 'Authorization': 'Bearer ' + token } : {};
  const qs = ['dataset=TaiwanStockDividend','data_id='+stockId,'start_date='+(new Date().getFullYear()-12)+'-01-01'].join('&');
  const resp = UrlFetchApp.fetch(url+'?'+qs, { muteHttpExceptions:true, headers });
  if (resp.getResponseCode() !== 200) return [];
  const json = JSON.parse(resp.getContentText('utf-8'));
  const rows = Array.isArray(json.data) ? json.data : [];
  const norm = rows.map(normalizeFinMindDividendRow_).filter(Boolean);
  const byKey = new Map();

  norm.forEach(r=>{
    // 計算現金 (盈餘+公積)
    const cash = r.cash_from_earnings + r.cash_from_statutory;
    
    // === 強制忽略股票股利 (設為 0) ===
    const stockRaw = 0; 

    // 若該次只有配股(無現金)且被我們歸零了，就直接略過不處理
    // 除非有現金發放日(代表可能有現金)
    const exDate = (cash>0 ? r.exDate_cash : '') || r.exDate_stock || '';
    if (!exDate && cash<=0) return;

    const key = [r.stockId, exDate].join('|');
    
    if (!byKey.has(key)) byKey.set(key, { 
      stockId:r.stockId, 
      stockName:r.stockName, 
      exDate, 
      payDate:r.payDate_cash, 
      cashPerShare:0, 
      stockPerShareRaw:0, // 這裡也強制 0
      yearText:r.yearText 
    });
    
    const cur = byKey.get(key);
    cur.cashPerShare += cash; 
    cur.stockPerShareRaw += 0; // 這裡也強制 +0
  });

  const yFrom = yearFromStr ? String(yearFromStr) : null;
  
  return Array.from(byKey.values()).map(v=>{
    // 年份邏輯：優先用發放日，沒有才用除息日
    let dateForYear = v.payDate || v.exDate;
    let year = dateForYear ? dateForYear.slice(0,4) : (v.yearText ? String(v.yearText).match(/(\d{4})/)[1] : '');

    return { 
      stockId:v.stockId, 
      stockName:v.stockName, 
      year, 
      exDate:v.exDate, 
      payDate:v.payDate, 
      cashPerShare:round2_(v.cashPerShare), 
      stockPerShareRaw:0, // 強制回傳 0
      stockPerShare:0     // 強制回傳 0
    };
  }).filter(v => !yFrom || v.year >= yFrom).sort((a,b)=> (a.exDate||'').localeCompare(b.exDate||''));
}

function normalizeFinMindDividendRow_(o) {
  if (!o) return null;
  const pickNum = (k) => { const n=Number(String(o[k]||'').replace(/[^\d.\-]/g,'')); return isFinite(n)?n:0; };
  return {
    stockId: o.stock_id, stockName: o.stock_name || o.company_name, yearText: o.year,
    exDate_cash: toYMDslash_(o.CashExDividendTradingDate), exDate_stock: toYMDslash_(o.StockExDividendTradingDate),
    payDate_cash: toYMDslash_(o.CashDividendPaymentDate),
    cash_from_earnings: pickNum('CashEarningsDistribution'), cash_from_statutory: pickNum('CashStatutorySurplus'),
    stock_from_earnings: pickNum('StockEarningsDistribution'), stock_from_statutory: pickNum('StockStatutorySurplus')
  };
}

function avgBuyOnDate_(sym, cutoffDate, openRows, tradeRows) {
  const lots = [];
  openRows.forEach(r=>{
    if (String(r['股票代碼']).trim()===sym) lots.push({ buyDate:toYMDslash_(r['買進日期'])||'1900/01/01', price:Number(r['買入價']), qty:Math.round(Number(r['持有股數'])) });
  });
  const tx = tradeRows.filter(r => String(r['股票代碼']).trim()===sym && toYMDslash_(r['成交日期'])<=cutoffDate)
    .map(r => ({ date:toYMDslash_(r['成交日期']), side:String(r['成交類別']), qty:Math.round(Number(r['股數'])), px:Number(r['成交價']) }));
  
  for (const x of tx) {
    if (x.side.includes('沖')) continue;
    if (isStockBonusSide_(x.side)) lots.push({ buyDate:x.date, price:0, qty:x.qty });
    else if (x.side.includes('買')) lots.push({ buyDate:x.date, price:x.px, qty:x.qty });
    else if (x.side.includes('賣')) {
      let remain = x.qty; lots.sort((a,b)=>a.buyDate.localeCompare(b.buyDate));
      for (let i=0; i<lots.length && remain>0; i++){
        const take = Math.min(lots[i].qty, remain);
        lots[i].qty -= take; remain -= take;
      }
    }
  }
  const qtySum = lots.reduce((s,l)=>s+Math.max(0,l.qty),0);
  return qtySum>0 ? round2_(lots.reduce((s,l)=>s+l.price*Math.max(0,l.qty),0)/qtySum) : 0;
}

/**
 * 已實現損益（增量模式）
 * - 若《已實現損益》已有資料，只計算最新賣出日之後的新賣單並追加
 * - 若工作表為空，執行完整重建
 * - 想強制完整重建：手動清空《已實現損益》工作表後再執行
 */
function rebuildRealizedPnL_FIFO() {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOpen = ensureSheetWithHeader_(C.SHEET_OPENING, []);
  const shT    = ensureSheetWithHeader_(C.SHEET_TRADES, []);
  const HEADERS = ['股票代碼','股票名稱','買進日期','股數','買進單價','買進成本','賣出日期','賣出單價','賣出總金額','買進手續費','賣出手續費','交易稅','淨獲利','股利','含息報酬','含息報酬率(%)','持有天數','每天獲利金額'];
  const out = ensureSheetWithHeader_(C.SHEET_REALIZED, HEADERS);
  if (out.getLastRow() === 0) out.appendRow(HEADERS);

  // 偵測舊格式（無「股利」欄），自動清空資料並更新表頭，強制完整重建
  if (out.getLastRow() >= 1) {
    const existingHdrs = out.getRange(1, 1, 1, out.getLastColumn()).getValues()[0].map(String);
    if (!existingHdrs.includes('股利')) {
      if (out.getLastRow() > 1) out.getRange(2, 1, out.getLastRow() - 1, out.getLastColumn()).clearContent();
      out.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      Logger.log('已實現損益：偵測到舊格式，已清空並更新表頭，執行完整重建');
    }
  }

  // 載入股利資料：code -> [{exDate, cashPerShare}]
  const divMap = new Map();
  const shDiv = ss.getSheetByName(C.SHEET_DIV || '股利狀況');
  if (shDiv && shDiv.getLastRow() > 1) {
    readTableAsObjects_(shDiv).forEach(r => {
      const sym    = String(r['股票代碼'] || '').trim();
      const exDate = toYMDslash_(r['除息日'] || '');
      const cash   = Number(r['現金股利 (元/股)'] || 0);
      if (sym && exDate && cash > 0) {
        if (!divMap.has(sym)) divMap.set(sym, []);
        divMap.get(sym).push({ exDate, cash });
      }
    });
  }

  // 找出已處理的最新賣出日期
  let lastDate = '';
  if (out.getLastRow() > 1) {
    const existingDates = out.getRange(2, 7, out.getLastRow() - 1, 1)
      .getValues().flat().map(v => toYMDslash_(String(v || ''))).filter(Boolean);
    if (existingDates.length) lastDate = existingDates.reduce((a, b) => a > b ? a : b);
  }

  const allTx = readTableAsObjects_(shT).map(r => ({
    sym: String(r['股票代碼']).trim(), name: String(r['股票名稱']).trim(),
    date: toYMDslash_(r['成交日期']), time: normTime_(r['成交時間']),
    price: Number(r['成交價']), qty: Number(r['股數']), amount: Number(r['成交金額']),
    fee: Number(r['手續費']), tax: Number(r['交易稅']), side: String(r['成交類別'])
  })).filter(x => x.sym && x.date && x.qty > 0);

  // 從期初庫存 + 所有買進建立 queue
  const buyQueues = buildBuyQueuesFromOpeningAndTrades_R4_(shOpen, shT);

  const sells = allTx.filter(x => x.side.includes('賣') && !isDayLoopSide_(x.side))
    .sort((a, b) => a.sym.localeCompare(b.sym) || a.date.localeCompare(b.date) || a.time.localeCompare(b.time));

  // 快進：消耗所有「已處理日期以前」的賣出，讓 queue 到達正確狀態
  if (lastDate) {
    let j = 0;
    while (j < sells.length) {
      const sym = sells[j].sym;
      const group = [];
      while (j < sells.length && sells[j].sym === sym) { group.push(sells[j]); j++; }
      const q = buyQueues.get(sym) || [];
      group.filter(s => s.date <= lastDate).forEach(sell => {
        let remain = sell.qty;
        while (remain > 0) {
          while (q.length && q[0].remainQty <= 0) q.shift();
          if (!q.length) break;
          const part = Math.min(remain, q[0].remainQty);
          q[0].remainQty -= part;
          remain -= part;
        }
      });
    }
  }

  // 只計算新賣出（lastDate 之後）
  const newSells = lastDate ? sells.filter(s => s.date > lastDate) : sells;
  if (!newSells.length) {
    Logger.log(`已實現損益：無新賣出資料（最新已處理：${lastDate || '尚無'}）`);
    return;
  }

  const rows = [];
  let i = 0;
  while (i < newSells.length) {
    const sym = newSells[i].sym;
    const group = [];
    while (i < newSells.length && newSells[i].sym === sym) { group.push(newSells[i]); i++; }
    const q = buyQueues.get(sym) || [];
    group.forEach(sell => {
      let remain = sell.qty;
      let sellFeeLeft = round2_(sell.fee || R4_calcFee_(sell.amount));
      let sellTaxLeft = round2_(sell.tax || R4_calcTax_(sell.side, sell.amount));
      while (remain > 0) {
        while (q.length && q[0].remainQty <= 0) q.shift();
        if (!q.length) break;
        const lot = q[0];
        const part = Math.min(remain, lot.remainQty);
        const buyCost    = round2_(lot.buyPrice * part);
        const buyFeePart = round2_(lot.buyFeePerShare * part);
        const sellGross  = round2_(sell.price * part);
        let sFee = round2_(sellFeeLeft * (part / sell.qty));
        let sTax = round2_(sellTaxLeft * (part / sell.qty));
        if (part === remain) { sFee = sellFeeLeft; sTax = sellTaxLeft; }
        sellFeeLeft -= sFee; sellTaxLeft -= sTax;
        const pnl  = round2_(sellGross - buyCost - buyFeePart - sFee - sTax);
        const days = Math.max(1, daysBetween_R4_(lot.buyDate, sell.date));
        // 計算持有期間股利：除息日在 [buyDate, sellDate] 之間的每股現金股利 × 本筆股數
        let divInPeriod = 0;
        (divMap.get(sym) || []).forEach(d => {
          if (d.exDate >= lot.buyDate && d.exDate <= sell.date) divInPeriod += d.cash * part;
        });
        divInPeriod = round2_(divInPeriod);
        const totalReturn    = round2_(pnl + divInPeriod);
        const totalReturnPct = buyCost > 0 ? round2_((totalReturn / buyCost) * 100) : '';
        rows.push([sym, lot.name || sell.name, lot.buyDate, part, lot.buyPrice, buyCost,
                   sell.date, sell.price, sellGross, buyFeePart, sFee, sTax, pnl,
                   divInPeriod, totalReturn, totalReturnPct, days, round2_(pnl / days)]);
        lot.remainQty -= part; remain -= part;
      }
    });
  }

  if (rows.length) {
    const startRow = out.getLastRow() + 1;
    out.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    out.getRange(startRow, 16, rows.length, 1).setNumberFormat('0.0'); // 含息報酬率(%)
  }
  Logger.log(`已實現損益：新增 ${rows.length} 筆（增量，前次最新賣出：${lastDate || '（首次完整）'}）`);
}

function buildBuyQueuesFromOpeningAndTrades_R4_(shOpen, shT) {
  const map = new Map();
  
  readTableAsObjects_(shOpen).forEach(r => {
    const sym = String(r['股票代碼']).trim();
    if (sym) {
      if (!map.has(sym)) map.set(sym, []);
      const qty = Number(r['持有股數']);
      map.get(sym).push({
        name: r['股票名稱'],
        buyDate: toYMDslash_(r['買進日期']) || '1900/01/01',
        buyPrice: Number(r['買入價']),
        buyFeePerShare: Number(r['手續費']) / qty,
        remainQty: Math.round(qty)
      });
    }
  });

  readTableAsObjects_(shT)
    .filter(r => {
      const side = String(r['成交類別']);
      // ✅ 加入配股判斷
      return (side.includes('買') || isStockBonusSide_(side)) && !isDayLoopSide_(side);
    })
    .forEach(x => {
      const sym = String(x['股票代碼']).trim();
      if (sym) {
        if (!map.has(sym)) map.set(sym, []);
        const qty = Number(x['股數']);
        const isBonu = isStockBonusSide_(String(x['成交類別']));
        map.get(sym).push({
          name: x['股票名稱'],
          buyDate: toYMDslash_(x['成交日期']),
          buyPrice: isBonu ? 0 : Number(x['成交價']),       // 配股成本為 0
          buyFeePerShare: isBonu ? 0 : Number(x['手續費']) / qty, // 配股手續費為 0
          remainQty: Math.round(qty)
        });
      }
    });

  for (const arr of map.values())
    arr.sort((a, b) => a.buyDate.localeCompare(b.buyDate));
  
  return map;
}

/** ===== 定期定額：從〈庫存紀錄〉複製並計算殖利率 =====
 * 邏輯升級：
 * 1. 讀取〈股利狀況〉建立快取 (代碼 -> [ {除息日, 現金股利}, ... ])
 * 2. 針對每一筆 DCA 買進，計算「買進後累計領到的總股息」
 * 3. 殖利率 = 累計總股息 / 總成本
 */
function appendDCAFromHoldings() {
  const C = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 準備分頁
  const shHold = ensureSheetWithHeader_(C.SHEET_HOLD, []);
  // 🆕 修改表頭：加入「累計已領股息」與「個人殖利率(%)」
  const shOut = ensureSheetWithHeader_(C.SHEET_DCA || '定期定額', [
    '買入日期','股票代碼','股票名稱','成交價','股數',
    '買入成本','手續費','總成本','累計已領股息','個人殖利率(%)'
  ]);

  // 2. 讀取並整理〈股利狀況〉資料 (用來查配息)
  const divMap = new Map(); // Key: 股票代碼, Value: Array of { exDate, cash }
  const shDiv = ss.getSheetByName(C.SHEET_DIV || '股利狀況');
  if (shDiv && shDiv.getLastRow() > 1) {
    const divRows = readTableAsObjects_(shDiv);
    divRows.forEach(r => {
      const sym = String(r['股票代碼']||'').trim();
      const exDate = toYMDslash_(r['除息日']||'');
      const cash = Number(r['現金股利 (元/股)']||0);
      
      if (sym && exDate && cash > 0) {
        if (!divMap.has(sym)) divMap.set(sym, []);
        divMap.get(sym).push({ exDate, cash });
      }
    });
  }

  // 3. 收集 DCA 設定 (DCA_1 ~ DCA_10)
  const dcaConfigs = [];
  for (let i=1; i<=10; i++){
    const sym   = String(C[`DCA_${i}_SYMBOL`]||'').trim();
    const start = toYMDslash_(C[`DCA_${i}_START`]||'');
    const end   = toYMDslash_(C[`DCA_${i}_END`]||''); // 留空＝至今
    if (sym && start) dcaConfigs.push({ sym, start, end });
  }
  if (dcaConfigs.length === 0) {
    Logger.log('DCA：〈設定〉沒有可用的 DCA_* 組合');
    return;
  }

  // 4. 讀取庫存紀錄 (來源)
  const holdRows = readTableAsObjects_(shHold).map(r => ({
    date: toYMDslash_(r['買進日期']||''),
    sym:  String(r['股票代碼']||'').trim(),
    name: String(r['股票名稱']||'').trim(),
    px:   Number(r['買入價']||0),
    qty:  Number(r['持有股數']||0),
    fee:  Number(r['手續費']||0)
  })).filter(r => r.sym && r.date && r.qty > 0);

  // 5. 既有資料去重 (避免重複寫入)
  const existed = new Set();
  if (shOut.getLastRow() > 1) {
    const data = shOut.getRange(2, 1, shOut.getLastRow()-1, 5).getValues(); // 只讀前5欄做key
    data.forEach(r => {
      // Key: 日期|代碼|單價|股數
      const key = [toYMDslash_(r[0]), String(r[1]).trim(), round2_(Number(r[3]||0)), Math.round(Number(r[4]||0))].join('|');
      existed.add(key);
    });
  }

  const toAppend = [];
  const inRange = (d, s, e) => (!s && !e) ? true : (s && !e) ? (d >= s) : (d >= s && d <= e);

  // 6. 遍歷並計算
  dcaConfigs.forEach(cfg => {
    holdRows
      .filter(r => r.sym === cfg.sym && inRange(r.date, cfg.start, cfg.end || null))
      .forEach(r => {
        const px  = round2_(r.px);
        const qty = Math.round(r.qty);
        const key = [r.date, r.sym, px, qty].join('|');

        if (existed.has(key)) return; // 若已存在則跳過
        existed.add(key); // 標記本次已處理

        // 成本計算
        const buyCost = round2_(px * qty);
        const totalCost = round2_(buyCost + Math.round(r.fee));

        // ★ 核心邏輯：計算這筆 DCA 領了多少股息
        let totalReceivedDiv = 0;
        const divs = divMap.get(r.sym) || [];
        divs.forEach(d => {
          // 如果「除息日」 >= 「買入日期」，代表這筆庫存有參與到除息
          if (d.exDate >= r.date) {
            totalReceivedDiv += (d.cash * qty);
          }
        });
        totalReceivedDiv = Math.round(totalReceivedDiv);

        // ★ 核心邏輯：殖利率 = 領到的總股息 / 總成本
        const yieldPct = (totalCost > 0) ? round2_((totalReceivedDiv / totalCost) * 100) : 0;

        toAppend.push([
          r.date, 
          r.sym, 
          r.name, 
          px, 
          qty,
          buyCost, 
          Math.round(r.fee), 
          totalCost,
          totalReceivedDiv, // 🆕 累計已領股息
          yieldPct          // 🆕 個人殖利率(%)
        ]);
      });
  });

  // 7. 寫入
  if (toAppend.length) {
    const startRow = shOut.getLastRow() + 1;
    shOut.getRange(startRow, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    
    // 設定格式 (第10欄是殖利率)
    shOut.getRange(startRow, 10, toAppend.length, 1).setNumberFormat("0.00");
    // 設定格式 (第4,6,7,8,9欄是金額)
    shOut.getRange(startRow, 4, toAppend.length, 6).setNumberFormat("#,##0");
  }
  
  Logger.log(`DCA：追加 ${toAppend.length} 筆 (含殖利率計算)`);
}
function runDividendsFullCycle_() {
  const props = PropertiesService.getScriptProperties();
  const CUR_KEY = 'DIV_CURSOR';
  if (props.getProperty(CUR_KEY) === null || props.getProperty(CUR_KEY) === '0') props.deleteProperty(CUR_KEY);
  props.setProperty('DIV_CYCLE_ACTIVE', '1');
  const startMs = Date.now();
  try {
    while (true) {
      updateDividendsFromFinMind_SAFE();
      const after = props.getProperty(CUR_KEY);
      if (after === '0' || after === null) return;
      if (Date.now() - startMs > 280000) {
        ScriptApp.newTrigger('runDividendsFullCycle_SAFE').timeBased().at(new Date(Date.now()+60000)).create();
        return;
      }
      Utilities.sleep(500);
    }
  } finally { props.deleteProperty('DIV_CYCLE_ACTIVE'); }
}

function fixStockCodes_OneTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['交易紀錄', '庫存紀錄', '期初庫存'].forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const rng = sh.getRange(2, 3, sh.getLastRow() - 1, 1);
    rng.setNumberFormat('@');
    const values = rng.getValues().map(r => {
      const val = r[0];
      if (!isNaN(val) && val !== '') {
        const num = Number(val); const str = String(val).trim();
        if (num < 100 && str.length < 4) return ["'" + ('0000' + num).slice(-4)];
        if (num >= 100 && num < 1000 && str.length < 5) return ["'" + ('00000' + num).slice(-5)];
        return [String(val)];
      }
      return [val];
    });
    rng.setValues(values);
  });
  Logger.log('🎉 代碼修復完成');
}

function tool_AuditBrokerInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const C = getCfg_();
  let sh = ss.getSheetByName('【對帳】券商庫存比對');
  if (sh) sh.clear(); else sh = ss.insertSheet('【對帳】券商庫存比對');
  
  const holdRows = readTableAsObjects_(ss.getSheetByName(C.SHEET_HOLD));
  const map = new Map();
  holdRows.forEach(r => {
    let sym = String(r['股票代碼']).trim();
    if(/^\d{1,3}$/.test(sym)) sym = sym.padStart(4,'0');
    const qty = Number(r['持有股數']);
    if (qty > 0) {
        if(!map.has(sym)) map.set(sym, {name:r['股票名稱'], qty:0});
        map.get(sym).qty += qty;
    }
  });
  const out = Array.from(map.keys()).sort().map(k => [k, map.get(k).name, map.get(k).qty, '', '', '']);
  sh.getRange(1,1,1,6).setValues([['股票代碼', '股票名稱', '系統總股數(A)', '券商實際股數(B)', '差異', '狀態']]).setFontWeight('bold').setBackground('#dad7cd');
  if (out.length) {
      sh.getRange(2,1,out.length,1).setNumberFormat('@');
      sh.getRange(2,1,out.length,6).setValues(out);
      sh.getRange(2,5,out.length,1).setFormulaR1C1('=IF(ISNUMBER(R[0]C[-1]), R[0]C[-1] - R[0]C[-2], "")');
      sh.getRange(2,6,out.length,1).setFormulaR1C1('=IF(NOT(ISNUMBER(R[0]C[-2])), "請輸入", IF(R[0]C[-1]=0, "✅", "❌"))');
  }
}

function cleanDuplicateDividends_OneTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('股利狀況');
  if (!sh || sh.getLastRow()<2) return;
  const data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const seen = new Set(); const kept = [];
  data.forEach(r => {
      const key = `${r[0]}|${toYMDslash_(r[3])}`;
      if(!seen.has(key)) { seen.add(key); kept.push(r); }
  });
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  if(kept.length) sh.getRange(2,1,kept.length,kept[0].length).setValues(kept);
  Logger.log(`清理完成，保留 ${kept.length} 筆`);
}
/* =========================================
   補強漏掉的 R4 小工具 (請貼在檔案最下方)
   ========================================= */

function daysBetween_R4_(d1, d2) { 
  const a = new Date(d1);
  const b = new Date(d2); 
  return Math.floor((b - a) / (24 * 3600 * 1000)); 
}

function R4_calcFee_(amount) {
  if (!amount || amount <= 0) return 0;
  const discount = (typeof getFeeDiscount_ === 'function') ? getFeeDiscount_() : 0.28;
  const raw = amount * 1.425 / 1000 * discount;
  return Math.max(1, Math.floor(raw));
}

function R4_calcTax_(sideText, amount) { 
  const s = String(sideText || '').trim(); 
  if (s.includes('現賣') || s === '賣') return Math.round(amount * 0.003); 
  if (s.includes('沖賣')) return Math.round(amount * 0.0015); 
  return 0; 
}


/**
 * [庫存追蹤器 V2] 強制時間排序修正版
 * 解決日期格式不一致導致的排序錯亂問題
 */
function debug_TraceStock_V2() {
  // ▼▼▼ 請在這裡輸入要檢查的股票代碼 ▼▼▼
  const TARGET_STOCK = '2330'; 
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOpen = ss.getSheetByName('期初庫存');
  const shTrade = ss.getSheetByName('交易紀錄');
  const logs = [];

  logs.push(`🔍 開始追蹤股票：【${TARGET_STOCK}】 (時間邏輯修正版)`);

  // 讀取資料通用函式
  const getRows = (sh) => {
    if (!sh || sh.getLastRow() < 2) return [];
    const raw = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    return raw.map(r => {
      const o = {};
      headers.forEach((h, i) => o[String(h).trim()] = r[i]);
      return o;
    });
  };

  const openRows = getRows(shOpen);
  const tradeRows = getRows(shTrade);

  // === 1. 建立期初庫存 ===
  const inventory = new Map(); // Key: Broker, Value: Qty
  
  openRows.forEach(r => {
    const sym = String(r['股票代碼']).trim();
    if (sym === TARGET_STOCK) {
      const qty = Number(r['持有股數']);
      const broker = String(r['證券商'] || '預設').trim();
      if (!inventory.has(broker)) inventory.set(broker, 0);
      inventory.set(broker, inventory.get(broker) + qty);
      logs.push(`[期初] 發現庫存：${qty} 股 | 券商：[${broker}]`);
    }
  });

  // === 2. 整理與標準化交易紀錄 ===
  // 輔助函式：將各種格式的日期轉為標準 Date 物件
  const parseDateTime = (d, t) => {
    let dateObj;
    if (d instanceof Date) {
      dateObj = new Date(d);
    } else {
      // 嘗試解析字串 "2023/01/01" 或 "2023-01-01"
      dateObj = new Date(String(d).replace(/\-/g, '/'));
    }
    
    // 處理時間
    let timeStr = "00:00:00";
    if (t instanceof Date) {
      // 如果時間欄位是 Date 物件 (Google Sheet 常見情況)，提取時分秒
      timeStr = Utilities.formatDate(t, 'GMT+8', 'HH:mm:ss');
    } else if (t) {
      timeStr = String(t).trim();
    }
    
    // 合併
    const dateStr = Utilities.formatDate(dateObj, 'GMT+8', 'yyyy/MM/dd');
    return {
      fullTime: new Date(`${dateStr} ${timeStr}`).getTime(), // 轉成毫秒數，絕對準確
      displayDate: dateStr,
      displayTime: timeStr
    };
  };

  const txs = tradeRows.filter(r => String(r['股票代碼']).trim() === TARGET_STOCK)
    .map(r => {
      const dt = parseDateTime(r['成交日期'], r['成交時間']);
      return {
        ts: dt.fullTime,          // 排序用的毫秒數
        dateStr: dt.displayDate,  // 顯示用的日期
        timeStr: dt.displayTime,  // 顯示用的時間
        type: String(r['成交類別']).trim(),
        qty: Number(r['股數']),
        broker: String(r['證券商'] || '預設').trim(),
        rowNum: r['__ROW_NUM__'] // 若有需要除錯行號
      };
    });

  // === 3. 強力排序 (依照毫秒數) ===
  txs.sort((a, b) => a.ts - b.ts);

  // === 4. 逐筆模擬 ===
  logs.push(`\n--- 開始模擬交易流程 (依時間戳記排序) ---`);
  
  txs.forEach(tx => {
    const key = tx.broker;
    let currentQty = inventory.get(key) || 0;
    
    const isBuy = tx.type.includes('買') || tx.type.includes('股利') || tx.type.includes('配股');
    const isSell = tx.type.includes('賣');

    if (tx.type.includes('沖')) {
       // logs.push(`⏭️ [跳過] ${tx.dateStr} (當沖)`);
       return;
    }

    if (isBuy) {
      currentQty += tx.qty;
      inventory.set(key, currentQty);
      logs.push(`➕ [買入] ${tx.dateStr} ${tx.timeStr} | +${tx.qty} | 券商:[${key}] | 結餘: ${currentQty}`);
    } 
    else if (isSell) {
      if (currentQty < tx.qty) {
        logs.push(`❌❌❌ [賣出失敗] ${tx.dateStr} ${tx.timeStr} | 要賣 ${tx.qty} | 券商:[${key}] | 庫存剩: ${currentQty}`);
        logs.push(`    ⚠️ 庫存不足！請檢查這一天之前的買入紀錄是否正確？`);
        // 檢查是否有別家券商有貨
        let otherHas = false;
        inventory.forEach((q, k) => {
          if (k !== key && q > 0) {
             logs.push(`    💡 提示：券商 [${k}] 還有 ${q} 股，但無法跨券商扣抵。`);
             otherHas = true;
          }
        });
        if(!otherHas) logs.push(`    💡 提示：所有券商都沒有庫存了。`);
        
        // 強制扣到負數繼續模擬
        currentQty -= tx.qty;
        inventory.set(key, currentQty);
      } else {
        currentQty -= tx.qty;
        inventory.set(key, currentQty);
        logs.push(`➖ [賣出] ${tx.dateStr} ${tx.timeStr} | -${tx.qty} | 券商:[${key}] | 結餘: ${currentQty}`);
      }
    } 
    else {
      logs.push(`❓ [無視] ${tx.dateStr} 類別:[${tx.type}]`);
    }
  });

  // 輸出結果
  const result = logs.join('\n');
  Logger.log(result);
  SpreadsheetApp.getUi().showModalDialog(
    SpreadsheetApp.createHtmlOutput(`<textarea style="width:100%; height:400px; font-family:monospace;">${result}</textarea>`).setWidth(600).setHeight(500),
    '庫存追蹤報告 V2'
  );
}


/**
 * [工具] 強力清除〈股利狀況〉重複資料
 * 邏輯：以「股票代碼 + 除息日」為唯一鍵值 (Key)
 * 修正：強化代碼 (補零/轉字串) 與 日期格式 (統一轉 yyyy/MM/dd) 的比對能力
 */
function cleanDuplicateDividends_Safe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('股利狀況');
  
  if (!sh || sh.getLastRow() < 2) {
    Logger.log('❌ 找不到〈股利狀況〉分頁或無資料。');
    return;
  }

  // 1. 讀取所有資料
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const range = sh.getRange(2, 1, lastRow - 1, lastCol);
  const data = range.getValues();
  
  const seen = new Set();
  const kept = [];
  let duplicateCount = 0;

  // 2. 輔助函式：標準化代碼 (去除空白, 轉字串, 若是純數字補滿4位)
  const normSym = (v) => {
    let s = String(v || '').trim();
    if (/^\d{1,3}$/.test(s)) s = ('0000' + s).slice(-4); // 50 -> 0050
    return s;
  };

  // 3. 遍歷並篩選
  data.forEach(row => {
    // 欄位索引：0=代碼, 3=除息日
    const sym = normSym(row[0]);
    const dateRaw = row[3];
    
    // 如果代碼或除息日是空的，視為無效行，暫時保留或略過 (這裡選擇保留以免誤刪手動資料)
    if (!sym || !dateRaw) {
      kept.push(row);
      return;
    }

    // 標準化日期
    const dateStr = toYMDslash_(dateRaw);
    
    // 產生唯一 Key: "00878|2023/08/16"
    const key = `${sym}|${dateStr}`;

    if (seen.has(key)) {
      duplicateCount++;
      // 發現重複！跳過此行，不加入 kept 陣列
    } else {
      seen.add(key);
      kept.push(row);
    }
  });

  // 4. 如果有發現重複，才執行寫入
  if (duplicateCount > 0) {
    // 清空舊資料
    range.clearContent();
    
    // 寫入去重後的資料
    if (kept.length > 0) {
      sh.getRange(2, 1, kept.length, kept[0].length).setValues(kept);
      
      // 順便修復格式：股票代碼欄位設為純文字
      sh.getRange(2, 1, kept.length, 1).setNumberFormat('@');
    }
    
    const msg = `✅ 清理完成！移除了 ${duplicateCount} 筆重複的股利資料，保留 ${kept.length} 筆。`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
  } else {
    const msg = '🎉 檢查完畢，沒有發現重複資料。';
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
  }
}

/**
 * [工具] 全域股票代碼修復 (修正版：區分 4碼 與 5碼 ETF)
 * 適用範圍：交易紀錄、庫存紀錄、期初庫存、股利狀況、已實現損益、定期定額
 */
function fixStockCodes_Global() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const targets = [
    { name: '交易紀錄', col: 3 },
    { name: '庫存紀錄', col: 1 },
    { name: '期初庫存', col: 1 },
    { name: '股利狀況', col: 1 },
    { name: '已實現損益', col: 1 },
    { name: '定期定額', col: 2 }
  ];

  let totalFixed = 0;

  targets.forEach(t => {
    const sh = ss.getSheetByName(t.name);
    if (!sh || sh.getLastRow() < 2) return;

    const lastRow = sh.getLastRow();
    const range = sh.getRange(2, t.col, lastRow - 1, 1);
    const values = range.getValues();
    let hasChange = false;

    const newValues = values.map(r => {
      let val = String(r[0]).trim();
      
      // 檢查是否為 1~3 位數的純數字
      if (/^\d{1,3}$/.test(val)) {
        const num = Number(val);
        
        // ★ 修正邏輯：
        // 1. 如果是 2 位數或更少 (如 50, 56, 8) -> 視為 00xx (4碼)
        if (num < 100) {
           val = ('0000' + num).slice(-4);
        } 
        // 2. 如果是 3 位數 (如 878, 929, 940) -> 視為 00xxx (5碼)
        else {
           val = ('00000' + num).slice(-5);
        }

        hasChange = true;
        totalFixed++;
        return [val]; 
      }
      
      // ★ 二次檢查：如果是錯誤的 0878 (4碼且0開頭，但應該是5碼的ETF)
      // 邏輯：如果是 "0" 開頭，且長度為 4，且轉數字後大於等於 100 (例如 "0878" -> 878)
      // 這代表它應該要是 00878
      if (val.length === 4 && val.startsWith('0') && Number(val) >= 100) {
         val = '0' + val; // 0878 -> 00878
         hasChange = true;
         totalFixed++;
         return [val];
      }

      return [val];
    });

    if (hasChange) {
      range.setNumberFormat('@'); // 強制純文字
      range.setValues(newValues);
      Logger.log(`✅ 已修復 [${t.name}] 的股票代碼格式`);
    }
  });

  const msg = totalFixed > 0 
    ? `🎉 修復完成！共修正了 ${totalFixed} 筆代碼 (含 00878 修正)。`
    : `👍 檢查完畢，代碼格式皆正確。`;
    
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

/* =========================================
   Part 5: 系統初始化與註冊 (分發專用)
   ========================================= */

/**
 * API: 檢查系統是否已初始化 (是否有帳號存在)
 * 回傳: { initialized: boolean }
 */
function api_checkSystemStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('帳號管理'); // 或是您原本設定的 User Sheet 名稱
  
  // 如果分頁不存在，或只有標題列(沒有內容)，視為未初始化
  if (!sh || sh.getLastRow() < 2) {
    return { initialized: false };
  }
  return { initialized: true };
}

/**
 * API: 註冊第一個管理員 (只在系統未初始化時允許執行)
 */
function api_registerFirstUser(id, pwd) {
  const status = api_checkSystemStatus();
  if (status.initialized) {
    return { ok: false, msg: '系統已初始化，禁止註冊。請直接登入。' };
  }

  if (!id || !pwd) return { ok: false, msg: '帳號密碼不能為空' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('帳號管理');
  
  // 如果分頁不存在，自動建立
  if (!sh) {
    sh = ss.insertSheet('帳號管理');
    // 建立表頭: ID, Password, Name, Role
    sh.appendRow(['ID', 'Password', 'Name', 'Role']);
  }

  // 寫入第一個使用者 (Admin)
  // 欄位順序: ID, Password, Name, Role
  sh.appendRow([id, pwd, 'Admin', 'admin']);

  return { ok: true, msg: '初始化成功！請使用新帳號登入。' };
}


/* =========================================
   Part 6: 手動工具
   ========================================= */

// api_getSettingsSchema 與 api_saveSettings 已移至 WebAPI.js，此處不重複定義

/** @deprecated 已移至 WebAPI.js */
function api_getSettingsSchema_UNUSED_() {
  const C = getCfg_(); // 讀取目前的設定值
  
  // 定義前端顯示的結構 (Schema)
  const schema = [
    // --- 第一組：核心與安全 ---
    {
      header: '核心與安全',
      group: '基本資料',
      items: [
        { key: 'ID_NUMBER', label: '身分證字號 (解鎖PDF用)', placeholder: 'A123456789' },
        { key: 'CLOUD_RUN_URL', label: 'Cloud Run 解鎖服務網址', placeholder: 'https://...' },
        { key: 'TZ', label: '系統時區', placeholder: 'Asia/Taipei', disabled: true }
      ]
    },

    // --- 第二組：Gmail 自動擷取 ---
    {
      header: 'Gmail 擷取設定',
      group: '信件標籤與範圍',
      items: [
        { key: 'GMAIL_LABEL_PDF', label: 'PDF 電子對帳單標籤', placeholder: 'Stock_PDF' },
        { key: 'GMAIL_LABEL_HTML', label: 'HTML 成交回報標籤', placeholder: 'Stock_Text' },
        { key: 'GMAIL_QUERY_DAYS', label: '每次往回抓取天數', type: 'number', placeholder: '7' }
      ]
    },

    // --- 第三組：券商與費率 ---
    {
      header: '券商與費率',
      group: '預設券商',
      items: [
        { key: 'BROKER_DEFAULT_NAME', label: '主要券商名稱', placeholder: '國泰證券' },
        { key: 'FEE_DISCOUNT', label: '主要手續費折數 (0~1)', type: 'number', placeholder: '0.28' }
      ]
    },
    {
      group: '第二券商 (選填)',
      items: [
        { key: 'BROKER_2_KEYWORD', label: '判定關鍵字 (如: 統一)', placeholder: '統一' },
        { key: 'BROKER_2_NAME', label: '顯示名稱', placeholder: '統一證券' },
        { key: 'BROKER_2_DISCOUNT', label: '手續費折數', type: 'number', placeholder: '0.6' }
      ]
    },

    // --- 第四組：股利設定 ---
    {
      header: '股利與 FinMind',
      group: '股利參數',
      items: [
        { key: 'FINMIND_TOKEN', label: 'FinMind API Token', placeholder: '選填，增加抓取額度' },
        { key: 'DIV_YEAR_FROM', label: '抓取起始年份', type: 'number', placeholder: '2015' },
        { key: 'DIV_CASH_FEE_PER_PAYOUT', label: '匯費 (每筆扣除)', type: 'number', placeholder: '10' },
        { key: 'DIV_STOCK_BONUS_TO_TRADES', label: '配股是否回寫交易紀錄', type: 'select', options: ['TRUE', 'FALSE'] }
      ]
    },
    {
      group: '進階控制',
      items: [
        { key: 'DIV_THROTTLE_MS', label: 'API 間隔 (毫秒)', type: 'number', placeholder: '1000' },
        { key: 'DIV_SYMBOLS_PER_RUN', label: '每次更新檔數', type: 'number', placeholder: '10' }
      ]
    },

    // --- 第五組：通知設定 ---
    {
      header: '系統通知',
      group: 'Email 通知',
      items: [
        { key: 'ALERT_ENABLED', label: '啟用錯誤通知', type: 'select', options: ['TRUE', 'FALSE'] },
        { key: 'ALERT_TO', label: '接收通知的 Email', placeholder: 'your@email.com' }
      ]
    },
    
    // --- 第六組：資料表名稱 (進階) ---
    {
      header: '資料表名稱 (進階)',
      group: '分頁名稱設定 (修改請謹慎)',
      items: [
        { key: 'SHEET_TRADES', label: '交易紀錄分頁' },
        { key: 'SHEET_HOLD', label: '庫存紀錄分頁' },
        { key: 'SHEET_DIV', label: '股利狀況分頁' },
        { key: 'SHEET_REALIZED', label: '已實現損益分頁' },
        { key: 'SHEET_DCA', label: '定期定額分頁' }
      ]
    }
  ];

  // --- 第七組：DCA 設定 (動態生成 1~10) ---
  const dcaItems = [];
  for (let i = 1; i <= 5; i++) { 
    dcaItems.push({ key: `DCA_${i}_NAME`, label: `[DCA ${i}] 股票名稱`, placeholder: '如: 國泰永續高股息' });
    dcaItems.push({ key: `DCA_${i}_SYMBOL`, label: `[DCA ${i}] 股票代碼`, placeholder: '00878' });
    dcaItems.push({ key: `DCA_${i}_START`, label: `[DCA ${i}] 開始日期`, type: 'text', placeholder: 'YYYY/MM/DD' });
  }
  
  schema.push({
    header: '定期定額設定 (DCA)',
    group: '投資標的 (前5組)',
    items: dcaItems
  });

  return { ok: true, schema: schema, values: C };
}

/** @deprecated 已移至 WebAPI.js */
function api_saveSettings_UNUSED_(newSettings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('設定');
  if (!sh) return { ok: false, msg: '找不到設定分頁' };

  const lastRow = sh.getLastRow();
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues(); 

  const newData = data.map(row => {
    const key = String(row[0]).trim();
    if (newSettings.hasOwnProperty(key)) {
      return [row[0], String(newSettings[key])];
    }
    return row;
  });

  sh.getRange(2, 1, newData.length, 2).setValues(newData);
  return { ok: true, msg: '設定已儲存！部分設定可能需要重新執行排程才能生效。' };
}



/**
 * [手動工具] 依指定日期區間擷取 Gmail 交易紀錄 → 寫入獨立分頁
 * 使用方式：修改下方三個參數後，在編輯器執行此函式
 */
function ingestFromGmail_ByDateRange() {
  // ▼▼▼ 修改這三個參數 ▼▼▼
  const START_DATE  = '2026/01/01';   // 起始日（含）YYYY/MM/DD
  const END_DATE    = '2026/03/31';   // 結束日（含）YYYY/MM/DD
  const OUTPUT_SHEET = '查詢結果';    // 輸出分頁名稱（不存在會自動建立）
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  const HEADERS = [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','備註'
  ];

  const C  = getCfg_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = C.TZ || 'Asia/Taipei';

  // --- 準備輸出分頁（每次清空重建）---
  let shOut = ss.getSheetByName(OUTPUT_SHEET);
  if (shOut) {
    shOut.clear();
  } else {
    shOut = ss.insertSheet(OUTPUT_SHEET);
  }
  shOut.appendRow(HEADERS);
  shOut.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#cfe2f3');
  shOut.setFrozenRows(1);

  // --- Gmail 搜尋日期參數 ---
  // Gmail 的 before 不含當天，所以要 +1 天
  const afterStr = START_DATE.replace(/\//g, '/');
  const beforeStr = (() => {
    const d = new Date(END_DATE.replace(/\//g, '-'));
    d.setDate(d.getDate() + 1);
    const y = d.getFullYear();
    const m = ('0' + (d.getMonth() + 1)).slice(-2);
    const dd = ('0' + d.getDate()).slice(-2);
    return `${y}/${m}/${dd}`;
  })();

  Logger.log(`📅 擷取範圍：${START_DATE} ～ ${END_DATE} → 輸出至〈${OUTPUT_SHEET}〉`);

  // --- 建立名稱對照表 ---
  const nameToCodeMap = new Map();
  const dcaWhitelist  = new Set();
  for (let i = 1; i <= 10; i++) {
    const sym = String(C[`DCA_${i}_SYMBOL`] || '').trim();
    const nm  = String(C[`DCA_${i}_NAME`]   || '').trim().replace(/\s+/g, '');
    if (sym) { dcaWhitelist.add(sym); if (nm) nameToCodeMap.set(nm, sym); }
  }
  const learnFromSheet = (sheetName) => {
    const s = ss.getSheetByName(sheetName);
    if (!s || s.getLastRow() < 2) return;
    s.getRange(2, 1, s.getLastRow() - 1, 4).getValues().forEach(r => {
      const c = String(r[2] || '').trim();
      const n = String(r[3] || '').trim().replace(/\s+/g, '');
      if (c && n && !nameToCodeMap.has(n)) nameToCodeMap.set(n, c);
    });
  };
  learnFromSheet(C.SHEET_TRADES);
  learnFromSheet(C.SHEET_OPENING);
  learnFromSheet(C.SHEET_HOLD);

  let parsed = [];

// --- 抓 PDF 信件 ---
  const labelPdf = (C.GMAIL_LABEL_PDF || '').trim();
  if (labelPdf) {
    const qPdf = `label:${labelPdf} after:${afterStr} before:${beforeStr} has:attachment`;
    Logger.log(`[PDF] 搜尋：${qPdf}`);
    const threads = GmailApp.search(qPdf, 0, 50);
    Logger.log(`[PDF] 找到 ${threads.length} 個對話串`);

    // 先把所有 PDF 附件收集起來
    const allPdfs = [];
    threads.forEach(t => t.getMessages().forEach(m => {
      m.getAttachments().forEach(att => {
        if (att.getContentType() === 'application/pdf' || att.getName().toLowerCase().endsWith('.pdf')) {
          allPdfs.push({ att, subject: m.getSubject() });
        }
      });
    }));

    Logger.log(`[PDF] 共找到 ${allPdfs.length} 個 PDF，開始逐一處理...`);

    // 逐一處理，每個之間等待
    for (let i = 0; i < allPdfs.length; i++) {
      const { att, subject } = allPdfs[i];

      if (i > 0) {
        Logger.log(`⏳ 等待 10 秒 (${i}/${allPdfs.length})...`);
        Utilities.sleep(10000);
      }

      try {
        Logger.log(`[PDF] 處理第 ${i+1}/${allPdfs.length}：${subject}`);
        const rawTable = callCloudRunToUnlock_(att, C.ID_NUMBER);
        const rows = parseCloudRunData_(rawTable, tz);
        const kept = [];
        rows.forEach(r => {
          let sym = r[2];
          if (isNaN(Number(sym)) && nameToCodeMap.has(sym)) { sym = nameToCodeMap.get(sym); r[2] = sym; }
          kept.push(r);
        });
        if (kept.length) {
          parsed = parsed.concat(kept);
          Logger.log(`[PDF] ✅ ${subject} → ${kept.length} 筆`);
        }
      } catch (e) {
        Logger.log(`[PDF] ❌ 失敗 (${subject})：${e.message}`);
      }
    }
  }

  // --- 抓 HTML 信件 ---
  const labelHtml = (C.GMAIL_LABEL_HTML || '').trim();
  if (labelHtml) {
    const qHtml = `label:${labelHtml} after:${afterStr} before:${beforeStr}`;
    Logger.log(`[HTML] 搜尋：${qHtml}`);
    const threads = GmailApp.search(qHtml, 0, 50);
    Logger.log(`[HTML] 找到 ${threads.length} 個對話串`);

    threads.forEach(t => t.getMessages().forEach(m => {
      const html = m.getBody();
      const body = m.getPlainBody();
      let rows = parseBrokerMailHTMLTable_CN_(html, m, tz);
      if (!rows.length) rows = parseBrokerMailPlainTable_CN_(body, m, tz);
      if (rows.length) parsed = parsed.concat(rows);
    }));
  }

  // --- 合併同委託單號 ---
  const consolidated = consolidateByOrderNo_(parsed);

  // --- 去重（同一次查詢內不重複）---
  const seenOrder = new Set();
  const seenNoOrd = new Set();
  const toWrite   = [];

  for (const r of consolidated) {
    const sym = String(r[2] || '').trim();
    const ord = normOrderNo_(r[8]);
    if (ord) {
      const k = makeKey_Order_(r[0], ord, sym);
      if (seenOrder.has(k)) continue;
      seenOrder.add(k); toWrite.push(r);
    } else {
      const k = makeKey_NoOrder_(r);
      if (seenNoOrd.has(k)) continue;
      seenNoOrd.add(k); toWrite.push(r);
    }
  }

  // --- 寫入輸出分頁 ---
  if (toWrite.length) {
    const startRow = shOut.getLastRow() + 1;
    shOut.getRange(startRow, 3, toWrite.length, 1).setNumberFormat('@'); // 股票代碼純文字
    shOut.getRange(startRow, 9, toWrite.length, 1).setNumberFormat('@'); // 委託單號純文字
    shOut.getRange(startRow, 1, toWrite.length, toWrite[0].length).setValues(toWrite);
    shOut.autoResizeColumns(1, HEADERS.length);
  }

  // --- 在分頁第一行加上查詢說明 ---
  // 插入一列說明在標題上方
  shOut.insertRowBefore(1);
  shOut.getRange(1, 1, 1, 4).setValues([[
    `查詢區間：${START_DATE} ～ ${END_DATE}`,
    `共 ${toWrite.length} 筆`,
    `執行時間：${Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')}`,
    ''
  ]]);
  shOut.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#f9cb9c')
    .setFontWeight('bold');
  shOut.setFrozenRows(2); // 說明列 + 表頭都凍結

  const msg = `✅ 完成！${START_DATE} ～ ${END_DATE}，共 ${toWrite.length} 筆 → 已寫入〈${OUTPUT_SHEET}〉`;
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}