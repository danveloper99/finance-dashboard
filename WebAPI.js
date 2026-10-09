/**
 * ============================================================
 * WebAPI.gs - 你不理財，才不理你後端 (完整版 v2.0)
 * 包含：登入、儀表板、資料清單、設定、市場分析模組
 * ============================================================
 */

/** Web App 入口 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('你不理財，才不理你')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/* ============================================================
   帳號與認證
   ============================================================ */

/**
 * 產生 HMAC-SHA256 Token（密碼 + 隨機 Salt）
 * Token 無法反推密碼，比 base64 安全
 */
function makeToken_(pwd) {
  const props = getProps_();
  let salt = props.getProperty('TOKEN_SALT');
  if (!salt) {
    salt = Utilities.base64Encode(
      Utilities.newBlob(String(Date.now()) + String(Math.random())).getBytes()
    );
    props.setProperty('TOKEN_SALT', salt);
  }
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pwd + salt
  );
  return Utilities.base64Encode(digest);
}

/**
 * 後端版本號：每次發布新的程式庫版本時 +1，並同步修改 index.html 的 LATEST_BACKEND_VERSION。
 * 前端會用它判斷朋友的後端是否過舊、需要更新程式庫版本。
 */
var APP_VERSION = 10;

/**
 * 帳號：每份後端（每個人用自己 Google 帳號部署的 GAS）只有一組帳號
 * - APP_USER / APP_PASSWORD 存在 Script Properties，只有部署者本人看得到
 * - 舊版沒有 APP_USER 時，帳號沿用《設定》的 ID_NUMBER
 */
function getAppUser_() {
  const u = getProps_().getProperty('APP_USER');
  return String(u || getCfg_()['ID_NUMBER'] || '').trim();
}

/** 這個後端是否已設定帳號（前端用來判斷要顯示「登入」還是「首次設定」） */
function api_getSetupStatus() {
  return { ok: true, configured: !!getProps_().getProperty('APP_PASSWORD') };
}

/** 首次設定帳號密碼：只有尚未設定過時可以呼叫 */
function api_setupAccount(userInput, pwdInput) {
  const user = String(userInput || '').trim();
  const pwd  = String(pwdInput || '');
  if (!/^[A-Za-z0-9_.@-]{3,40}$/.test(user)) return { ok: false, msg: '帳號限 3~40 個英數字（可含 _ . @ -）' };
  if (pwd.length < 6) return { ok: false, msg: '密碼至少 6 個字元' };

  const lock = getScriptLock_();
  if (!lock.tryLock(10000)) return { ok: false, msg: '系統忙碌中，請稍後再試' };
  try {
    const props = getProps_();
    if (props.getProperty('APP_PASSWORD')) return { ok: false, msg: '這個後端已經設定過帳號，請直接登入' };
    props.setProperties({ APP_USER: user, APP_PASSWORD: pwd });
    return { ok: true, token: makeToken_(pwd), user, version: APP_VERSION, autoUpdate: isLoaderMode_() };
  } finally {
    lock.releaseLock();
  }
}

/** 登入驗證，回傳 Token */
function api_login(idInput, pwdInput) {
  const id  = String(idInput  || '').trim();
  const pwd = String(pwdInput || '');
  if (!id || !pwd) return { ok: false, msg: '請輸入帳號與密碼' };

  const storedPwd = getProps_().getProperty('APP_PASSWORD');
  if (!storedPwd) return { ok: false, needSetup: true, msg: '這個後端還沒設定帳號，請先到「首次設定」' };

  const storedUser = getAppUser_();
  if (storedUser && id.toUpperCase() === storedUser.toUpperCase() && pwd === storedPwd)
    return { ok: true, token: makeToken_(storedPwd), user: storedUser, version: APP_VERSION, autoUpdate: isLoaderMode_() };
  return { ok: false, msg: '帳號或密碼錯誤' };
}

/** 驗證 Token；改密碼後舊 Token 自動失效 */
function resolveAuth_(token) {
  const storedPwd = getProps_().getProperty('APP_PASSWORD');
  return !!(token && storedPwd && token === makeToken_(storedPwd));
}

/** 自動登入 Token 驗證 */
function api_auth_token(tokenInput) {
  return resolveAuth_(tokenInput) ? { ok: true, user: getAppUser_(), version: APP_VERSION, autoUpdate: isLoaderMode_() } : { ok: false };
}

/** 修改密碼（需登入），回傳新 Token */
function api_changePassword(oldPwd, newPwd) {
  const props     = getProps_();
  const storedPwd = props.getProperty('APP_PASSWORD');
  if (storedPwd && String(oldPwd) !== String(storedPwd))
    return { ok: false, msg: '舊密碼不正確' };
  if (String(newPwd || '').length < 6) return { ok: false, msg: '新密碼至少 6 個字元' };
  props.setProperty('APP_PASSWORD', String(newPwd));
  return { ok: true, msg: '密碼已更新', token: makeToken_(String(newPwd)) };
}

/* ============================================================
   儀表板
   ============================================================ */

function api_getDashboard() {
  const C  = getCfg_();
  const ss = getSS_();

  // 1. 庫存：計算成本與市值
  const shHold = ss.getSheetByName(C.SHEET_HOLD);
  const holds  = readSheetAsObjects_(shHold);
  let totalCost = 0, totalMarketVal = 0;
  holds.forEach(r => {
    const qty   = Number(r['持有股數'] || 0);
    const cost  = Number(r['買入成本 (單純買入價*股數)'] || 0);
    let price   = Number(r['現價']); if (isNaN(price)) price = 0;
    totalCost      += cost;
    totalMarketVal += price > 0 ? qty * price : cost;
  });

  // 2. 已實現損益
  const shReal = ss.getSheetByName(C.SHEET_REALIZED);
  const reals  = readSheetAsObjects_(shReal);
  let totalRealized = 0;
  const pnlByYear   = {};
  reals.forEach(r => {
    const p = Number(r['淨獲利'] || 0);
    totalRealized += p;
    const d = String(r['賣出日期'] || '');
    const y = d.length >= 4 ? d.substring(0, 4) : '未知';
    if (y !== '未知') pnlByYear[y] = (pnlByYear[y] || 0) + p;
  });

  // 3. 股利
  const shDiv  = ss.getSheetByName(C.SHEET_DIV);
  const divs   = readSheetAsObjects_(shDiv);
  let totalDiv = 0;
  const divByYear  = {};
  divs.forEach(r => {
    const d    = Number(r['實際領取金額 (扣除每筆手續費10元)'] || 0);
    totalDiv  += d;
    const yStr = String(r['股利所屬年度'] || r['除息日'] || '');
    const y    = yStr.length >= 4 ? yStr.substring(0, 4) : '未知';
    if (y !== '未知') divByYear[y] = (divByYear[y] || 0) + d;
  });

  // 4. 歷年含息總報酬
  const totalByYear = {};
  new Set([...Object.keys(pnlByYear), ...Object.keys(divByYear)]).forEach(y => {
    totalByYear[y] = (pnlByYear[y] || 0) + (divByYear[y] || 0);
  });

  return {
    ok: true,
    summary: {
      investedCost:     totalCost,
      marketValue:      totalMarketVal,
      realizedProfit:   totalRealized,
      totalDividend:    totalDiv,
      unrealizedProfit: totalMarketVal - totalCost,
      totalProfit:      totalRealized + totalDiv + (totalMarketVal - totalCost),
    },
    charts: { pnlByYear, divByYear, totalByYear },
  };
}

/* ============================================================
   資料清單
   ============================================================ */

function api_getDataList(type) {
  const C  = getCfg_();
  const ss = getSS_();
  const sheetMap = {
    holdings:  C.SHEET_HOLD,
    dividends: C.SHEET_DIV,
    realized:  C.SHEET_REALIZED,
    dca:       C.SHEET_DCA,
  };
  const sh = ss.getSheetByName(sheetMap[type]);
  if (!sh) return { ok: true, data: [] };

  const config = type === 'holdings' ? {
    defaultName:     C.BROKER_DEFAULT_NAME   || '預設券商',
    defaultDiscount: Number(C.FEE_DISCOUNT   || 0.28),
    broker2Key:      C.BROKER_2_KEYWORD      || '',
    broker2Name:     C.BROKER_2_NAME         || '',
    broker2Discount: Number(C.BROKER_2_DISCOUNT || 0.28),
  } : {};

  return { ok: true, data: readSheetAsObjects_(sh), config };
}

/* ============================================================
   設定頁面
   ============================================================ */

function api_getSettingsSchema() {
  const C = getCfg_();
  const schema = [
    { group: '基本設定', icon: 'ph-gear',
      desc: '設定時區與手續費折數，影響所有損益計算的基礎。「解鎖對帳單密碼」填入 PDF 對帳單的解鎖密碼（通常為身分證字號），由 Cloud Run 服務使用。',
      items: [
      { key: 'ID_NUMBER',    label: '解鎖對帳單密碼',     type: 'text',   placeholder: '如: A123456789（身分證字號）' },
      { key: 'TZ',           label: '時區',               type: 'text',   placeholder: 'Asia/Taipei' },
      { key: 'FEE_DISCOUNT', label: '手續費折數 (0~1)',   type: 'number', placeholder: '0.28' },
    ]},
    { group: 'Gmail 擷取', icon: 'ph-envelope',
      desc: '設定 Gmail 往回搜尋天數與郵件分類標籤（請先在 Gmail 建立對應標籤並套用至成交回報信件）。Cloud Run 網址為 PDF 解鎖服務，留空則跳過 PDF 解析。',
      items: [
      { key: 'GMAIL_QUERY_DAYS', label: '往回搜尋天數',  type: 'number', placeholder: '7' },
      { key: 'GMAIL_LABEL_PDF',  label: 'PDF 郵件標籤',  type: 'text',   placeholder: '如: 對帳單' },
      { key: 'GMAIL_LABEL_HTML', label: 'HTML 郵件標籤', type: 'text',   placeholder: '如: 成交回報' },
      { key: 'CLOUD_RUN_URL',    label: 'Cloud Run 網址', type: 'text',  placeholder: 'https://...' },
    ]},
    { group: '券商設定', icon: 'ph-buildings',
      desc: '設定主要券商名稱與手續費折數。若有第二家券商，填入其 Email 關鍵字，系統會自動依關鍵字區分兩家券商的成本計算。',
      items: [
      { key: 'BROKER_DEFAULT_NAME', label: '主要券商名稱',   type: 'text',   placeholder: '國泰證券' },
      { key: 'BROKER_2_KEYWORD',    label: '第二券商關鍵字', type: 'text',   placeholder: '統一（Email 中出現的關鍵字）' },
      { key: 'BROKER_2_NAME',       label: '第二券商名稱',   type: 'text',   placeholder: '統一證券' },
      { key: 'BROKER_2_DISCOUNT',   label: '第二券商折數',   type: 'number', placeholder: '0.28' },
    ]},
    { group: '庫存與損益起算日', icon: 'ph-calendar',
      desc: '若只需統計特定日期後的交易（例如從某年度開始），填入起算日；留空則計算全部歷史紀錄。',
      items: [
      { key: 'HOLDINGS_START_DATE', label: '庫存計算起算日', type: 'date', placeholder: '' },
      { key: 'DIV_START_DATE',      label: '股利計算起算日', type: 'date', placeholder: '' },
    ]},
    { group: '股利設定', icon: 'ph-plant',
      desc: '需在 FinMind 官網申請免費 Token 才能自動抓取配息資料。配股可選擇是否寫回交易紀錄（影響成本計算），以及零股的取整方式。',
      items: [
      { key: 'FINMIND_TOKEN',             label: 'FinMind Token',       type: 'text',   placeholder: '前往 finmindtrade.com 申請' },
      { key: 'DIV_YEAR_FROM',             label: '股利起始年份',        type: 'number', placeholder: String(new Date().getFullYear() - 8) },
      { key: 'DIV_CASH_FEE_PER_PAYOUT',   label: '每筆股利手續費 (元)', type: 'number', placeholder: '10' },
      { key: 'DIV_STOCK_BONUS_TO_TRADES', label: '配股寫回交易紀錄',   type: 'select', options: ['TRUE', 'FALSE'] },
      { key: 'DIV_STOCK_BONUS_ROUNDING',  label: '配股取整規則',       type: 'select', options: ['FLOOR', 'ROUND', 'CEIL'] },
    ]},
    { group: '通知設定', icon: 'ph-bell',
      desc: '啟用後，系統執行發生錯誤時會自動寄信通知到指定信箱，方便排查問題。建議填入本人的 Gmail 地址。',
      items: [
      { key: 'ALERT_TO',      label: '通知 Email',    type: 'text',   placeholder: 'your@gmail.com' },
      { key: 'ALERT_ENABLED', label: '啟用錯誤通知', type: 'select', options: ['TRUE', 'FALSE'] },
      { key: 'MONTHLY_REPORT_ENABLED', label: '每月月報（1 號早上寄上個月摘要）', type: 'select', options: ['TRUE', 'FALSE'] },
    ]},
  ];

  const dcaDesc = '設定一組定期定額標的。填入 PDF 信件中出現的關鍵字（用於識別此標的）、股票代碼與投資期間，系統會自動彙整累計股數、平均成本與殖利率。';
  for (let i = 1; i <= 10; i++) {
    schema.push({ group: `定期定額 #${i}`, icon: 'ph-calendar-check', desc: dcaDesc, items: [
      { key: `DCA_${i}_NAME`,   label: 'PDF 關鍵字',       type: 'text', placeholder: '如: 國泰永續高股息' },
      { key: `DCA_${i}_SYMBOL`, label: '股票代碼',         type: 'text', placeholder: '如: 00878' },
      { key: `DCA_${i}_START`,  label: '開始日', type: 'date', placeholder: '' },
      { key: `DCA_${i}_END`,    label: '結束日 (留空至今)', type: 'date', placeholder: '' },
    ]});
  }

  return { ok: true, schema, values: C };
}

function api_saveSettings(newValues) {
  const ss = getSS_();
  const sh = ss.getSheetByName('設定');
  if (!sh) return { ok: false, msg: '找不到設定頁' };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, msg: '設定頁無資料' };

  // 一次讀取全部 key-value，在記憶體更新，再一次 batch 寫回（比逐格 setValue 快 10x）
  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const keyMap = new Map();
  data.forEach(([k], i) => { if (k) keyMap.set(String(k).trim(), i); });

  let changed = false;
  const NEW_KEYS = new Set(['MONTHLY_REPORT_ENABLED', 'AUTO_UPDATE']); // schema 新增、舊試算表還沒有的 key
  const toAppend = [];
  Object.entries(newValues).forEach(([key, val]) => {
    if (keyMap.has(key)) { data[keyMap.get(key)][1] = val; changed = true; }
    else if (NEW_KEYS.has(key)) toAppend.push([key, val]);
  });
  if (toAppend.length) sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 2).setValues(toAppend);

  if (changed) {
    // 先將 DCA 日期欄位設為純文字格式，避免 Google Sheets 自動轉換日期
    const dcaDateKeys = new Set();
    for (let i = 1; i <= 10; i++) { dcaDateKeys.add('DCA_' + i + '_START'); dcaDateKeys.add('DCA_' + i + '_END'); }
    data.forEach(([k], i) => {
      if (k && dcaDateKeys.has(String(k).trim())) sh.getRange(i + 2, 2).setNumberFormat('@');
    });
    sh.getRange(2, 2, data.length, 1).setValues(data.map(r => [r[1]]));
  }
  return { ok: true, msg: '設定已儲存' };
}

/* ============================================================
   市場分析模組（Gemini 版）
   ============================================================ */

/**
 * 讀取《庫存紀錄》，依股票代碼彙總
 * 現價直接使用工作表的 GOOGLEFINANCE 公式值
 */
function api_getHoldingsForAnalysis() {
  const C  = getCfg_();
  const ss = getSS_();
  const sh = ss.getSheetByName(C.SHEET_HOLD || '庫存紀錄');
  if (!sh || sh.getLastRow() < 2) return { ok: false, msg: '找不到庫存紀錄或無資料' };

  const rows = readSheetAsObjects_(sh);
  const map  = new Map();

  // 讀取股利資料，累計每檔股票已領現金股息
  const divTotals = new Map();
  const divSh = ss.getSheetByName(C.SHEET_DIV || '股利狀況');
  if (divSh && divSh.getLastRow() > 1) {
    const divRows = readSheetAsObjects_(divSh);
    divRows.forEach(r => {
      const code = String(r['股票代碼'] || '').trim();
      if (!code) return;
      const amt = Number(r['實際領取金額 (扣除每筆手續費10元)'] || 0);
      if (!isNaN(amt) && amt > 0) divTotals.set(code, (divTotals.get(code) || 0) + amt);
    });
  }

  rows.forEach(r => {
    const code = String(r['股票代碼'] || '').trim();
    if (!code) return;
    const name     = String(r['股票名稱'] || '').trim();
    const qty      = Number(r['持有股數']  || 0);
    const avgPrice = Number(r['買入價']    || 0);
    const cost     = Number(r['買入成本 (單純買入價*股數)'] || 0) || avgPrice * qty;
    let   curPrice = parseFloat(r['現價']);
    if (isNaN(curPrice) || curPrice <= 0) curPrice = 0;

    if (!map.has(code)) {
      map.set(code, { code, name, totalQty: 0, totalCost: 0, totalPrincipal: 0, curPrice: 0 });
    }
    const item = map.get(code);
    item.totalQty       += qty;
    item.totalCost      += cost;
    item.totalPrincipal += avgPrice * qty;
    if (curPrice > 0) item.curPrice = curPrice;
  });

  const data = Array.from(map.values())
    .filter(item => item.totalQty > 0.1)
    .sort((a, b) => a.code.localeCompare(b.code))
    .map(item => {
      const avgCost = item.totalQty > 0 ? item.totalPrincipal / item.totalQty : 0;
      const curVal  = item.curPrice > 0 ? item.curPrice * item.totalQty : null;
      const capPnl  = curVal != null ? curVal - item.totalCost : null;
      const pct     = capPnl != null && item.totalCost > 0
                      ? Math.round(capPnl / item.totalCost * 1000) / 10 : null;
      const totalDiv = Math.round(divTotals.get(item.code) || 0);
      const totalReturn = capPnl != null ? capPnl + totalDiv : (totalDiv > 0 ? totalDiv : null);
      const totalReturnPct = totalReturn != null && item.totalCost > 0
        ? Math.round(totalReturn / item.totalCost * 1000) / 10 : null;
      return {
        code:           item.code,
        name:           item.name,
        qty:            Math.round(item.totalQty),
        avgCost:        Math.round(avgCost * 100) / 100,
        totalCost:      Math.round(item.totalCost),
        curPrice:       item.curPrice || null,
        curVal:         curVal  != null ? Math.round(curVal)  : null,
        capPnl:         capPnl  != null ? Math.round(capPnl)  : null,
        pct,
        totalDiv,
        totalReturn:    totalReturn != null ? Math.round(totalReturn) : null,
        totalReturnPct,
      };
    });

  return { ok: true, data };
}

/**
 * 透過 GAS 後端呼叫 Gemini API
 * API Key 存於 Script Properties，前端完全看不到
 */
/**
 * images: [{mimeType:'image/jpeg', base64:'...'}] (選填)
 * 支援 Gemini multimodal：同時傳入文字 + 圖片
 */
function api_callGemini(prompt, images) {
  const key = getProps_().getProperty('GEMINI_API_KEY') || '';
  if (!key) return { ok: false, msg: '尚未設定 Gemini API Key，請至「市場分析」頁儲存。' };

  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + key;

    // 組合 parts：先放文字，再附加圖片
    const parts = [{ text: prompt }];
    if (Array.isArray(images)) {
      images.forEach(img => {
        if (img && img.mimeType && img.base64) {
          parts.push({ inline_data: { mime_type: img.mimeType, data: img.base64 } });
        }
      });
    }

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [
          { role: 'user', parts: parts }
        ],
        generationConfig: {
          maxOutputTokens: 8192,
          temperature: 0.7,
        },
      }),
      muteHttpExceptions: true,
    });

    const code = resp.getResponseCode();
    const body = JSON.parse(resp.getContentText('utf-8'));

    if (code !== 200)
      return { ok: false, msg: `API 錯誤 (${code}): ${body.error?.message || ''}` };

    const text = body.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('') || '';

    const finishReason = body.candidates?.[0]?.finishReason || '';
    Logger.log('Gemini finishReason: ' + finishReason);
    Logger.log('Gemini response length: ' + text.length);

    if (!text) return { ok: false, msg: 'Gemini 回傳空白內容，請稍後再試' };

    return { ok: true, text, finishReason };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

/**
 * 儲存分析結果到《市場分析紀錄》工作表
 */
function api_saveAnalysis(text) {
  if (!text) return { ok: false };
  const ss = getSS_();
  let sh   = ss.getSheetByName('市場分析紀錄');
  if (!sh) {
    sh = ss.insertSheet('市場分析紀錄');
    sh.appendRow(['時間', '分析內容']);
    sh.getRange(1, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground('#344e41')
      .setFontColor('white');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 900);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  sh.appendRow([now, text]);
  sh.getRange(sh.getLastRow(), 2).setWrap(true);
  return { ok: true, ts: now };
}

/**
 * 讀取《市場分析紀錄》歷史，最新在前，最多 30 筆
 */
function api_getAnalysisHistory() {
  const ss = getSS_();
  const sh = ss.getSheetByName('市場分析紀錄');
  if (!sh || sh.getLastRow() < 2) return { ok: true, data: [] };

  const raw = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return {
    ok: true,
    data: raw
      .map(r => ({
        ts:   r[0] instanceof Date
              ? Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm')
              : String(r[0] || ''),
        text: String(r[1] || ''),
      }))
      .filter(r => r.text)
      .reverse()
      .slice(0, 30),
  };
}


/* ============================================================
   資產追蹤模組
   ============================================================ */

// 帳戶 key 清單（順序對應 Sheet 欄位）
var WEALTH_KEYS_ = [
  'ctbc_twd_saving','ctbc_twd_fixed',
  'ctbc_usd_saving','ctbc_usd_fixed',
  'ctbc_cny_saving','ctbc_cny_fixed',
  'cathay_twd','cathay_usd',
  'taishin_twd','taishin_jpy',
  'richart_twd','richart_fund',
  'chang_twd','land_twd','post_twd',
  'uni_stock','cathay_stock','cathay_us_stock',
  'rate_usd','rate_jpy','rate_cny'
];

var WEALTH_SHEET_HEADERS_ = ['記錄日期','期間',
  '中信台幣活存','中信台幣定存','中信美金活存','中信美金定存','中信人民幣活存','中信人民幣定存',
  '國泰台幣活存','國泰美金活存',
  '台新台幣活存','台新日幣活存',
  'Richart活存','Richart基金',
  '彰銀台幣','合庫台幣','郵局台幣',
  '統一證券','國泰證券','國泰證券（美）',
  '匯率USD/TWD','匯率JPY/TWD','匯率CNY/TWD',
  '台幣合計','備註'
];

function ensureWealthSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName('資產快照');
  if (!sh) {
    sh = ss.insertSheet('資產快照');
    sh.getRange(1,1,1,WEALTH_SHEET_HEADERS_.length).setValues([WEALTH_SHEET_HEADERS_])
      .setFontWeight('bold').setBackground('#344e41').setFontColor('white');
    sh.setFrozenRows(1);
  } else {
    const lastCol = sh.getLastColumn();
    const col20Val = lastCol >= 20 ? sh.getRange(1, 20).getValue() : '';
    const col21Val = lastCol >= 21 ? sh.getRange(1, 21).getValue() : '';
    if (col20Val === '國泰美股') {
      if (col21Val === '國泰證券（美）') {
        // 舊版誤插了「國泰美股」空欄，刪除後「國泰證券（美）」自動回到 col20
        sh.deleteColumn(20);
      } else {
        // 「國泰美股」存在但後面沒有「國泰證券（美）」，直接改名
        sh.getRange(1, 20).setValue('國泰證券（美）')
          .setFontWeight('bold').setBackground('#344e41').setFontColor('white');
      }
    } else if (col20Val !== '國泰證券（美）') {
      // 兩個欄都不存在，補插入
      sh.insertColumnAfter(19);
      sh.getRange(1, 20).setValue('國泰證券（美）')
        .setFontWeight('bold').setBackground('#344e41').setFontColor('white');
    }
    // col20Val === '國泰證券（美）' → 已正確，不需動
  }
  return sh;
}

function api_saveWealthSnapshot(data) {
  try {
    const sh = ensureWealthSheet_();
    const now = Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd HH:mm');
    const row = [now, data.period || ''];
    WEALTH_KEYS_.forEach(k => row.push(Number(data[k]) || 0));

    // 後端重新計算台幣合計，避免前端傳值有誤
    const rU = Number(data.rate_usd) || 32;
    const rJ = Number(data.rate_jpy) || 0.22;
    const rC = Number(data.rate_cny) || 4.4;
    const TWD_KEYS = ['ctbc_twd_saving','ctbc_twd_fixed','cathay_twd','taishin_twd',
                      'richart_twd','richart_fund','chang_twd','land_twd','post_twd',
                      'uni_stock','cathay_stock'];
    const USD_KEYS = ['ctbc_usd_saving','ctbc_usd_fixed','cathay_usd','cathay_us_stock'];
    const JPY_KEYS = ['taishin_jpy'];
    const CNY_KEYS = ['ctbc_cny_saving','ctbc_cny_fixed'];
    let total = 0;
    TWD_KEYS.forEach(k => total += Number(data[k])||0);
    USD_KEYS.forEach(k => total += (Number(data[k])||0) * rU);
    JPY_KEYS.forEach(k => total += (Number(data[k])||0) * rJ);
    CNY_KEYS.forEach(k => total += (Number(data[k])||0) * rC);
    total = Math.round(total);

    row.push(total, data.note || '');
    sh.appendRow(row);
    return { ok: true, msg: '快照已儲存', total: total };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_getWealthHistory() {
  try {
    const sh = ensureWealthSheet_(); // 確保欄位結構為最新（修正舊版欄位錯位問題）
    if (sh.getLastRow() < 2) return { ok: true, data: [], last: null };
    const rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    const data = rows.map(row => {
      const obj = {
        date:   row[0] instanceof Date ? Utilities.formatDate(row[0],'Asia/Taipei','yyyy-MM-dd') : String(row[0]||'').slice(0,10),
        period: String(row[1]||''),
      };
      WEALTH_KEYS_.forEach((k,i) => { obj[k] = Number(row[i+2])||0; });
      obj.totalTWD  = Number(row[WEALTH_KEYS_.length+2])||0;
      obj.note      = String(row[WEALTH_KEYS_.length+3]||'');
      return obj;
    });
    return { ok: true, data, last: data[data.length-1] || null };
  } catch(e) { return { ok: false, msg: e.message, data: [], last: null }; }
}

/* --- 第二份資產紀錄（自訂帳戶清單）---
 * 名稱與帳戶清單存在《設定》：WEALTH_BOOK_A_NAME / WEALTH_BOOK_B_NAME / WEALTH_B_ACCOUNTS（JSON）
 * 快照存在《資產快照2》，一個帳戶一列（改名、刪除帳戶都不影響舊紀錄）
 */
var WEALTH_B_SHEET_    = '資產快照2';
var WEALTH_B_HEADERS_  = ['記錄日期','期間','帳戶ID','帳戶名稱','幣別','類型','金額','匯率','台幣金額','備註'];
var WEALTH_CURRENCIES_ = ['TWD','USD','JPY','CNY'];

/** 寫入目前試算表《設定》的某個 key（沒有該列就新增） */
function setCfgValue_(key, value) {
  const sh = getSS_().getSheetByName('設定');
  if (!sh) throw new Error('找不到設定頁');
  const lastRow = sh.getLastRow();
  const keys = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  const idx = keys.indexOf(key);
  const row = idx >= 0 ? idx + 2 : lastRow + 1;
  sh.getRange(row, 2).setNumberFormat('@');
  sh.getRange(row, 1, 1, 2).setValues([[key, value]]);
}

function readWealthBooksCfg_() {
  const C = getCfg_();
  let accounts = [];
  try { accounts = JSON.parse(C.WEALTH_B_ACCOUNTS || '[]'); } catch (e) { accounts = []; }
  return {
    nameA:    C.WEALTH_BOOK_A_NAME || '我的資產',
    nameB:    C.WEALTH_BOOK_B_NAME || '第二份紀錄',
    accounts: Array.isArray(accounts) ? accounts : [],
  };
}

function api_getWealthBooks() {
  return Object.assign({ ok: true }, readWealthBooksCfg_());
}

/** 儲存兩份紀錄的名稱 + 第二份的帳戶清單 */
function api_saveWealthBooks(cfg) {
  try {
    cfg = cfg || {};
    const clean = s => String(s || '').trim().slice(0, 30);
    const seen = new Set();
    const accounts = (Array.isArray(cfg.accounts) ? cfg.accounts : [])
      .map(a => ({
        id:       String((a && a.id) || '').replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 20),
        name:     clean(a && a.name),
        currency: WEALTH_CURRENCIES_.includes(a && a.currency) ? a.currency : 'TWD',
        type:     (a && a.type) === 'invest' ? 'invest' : 'bank',
      }))
      .filter(a => a.id && a.name && !seen.has(a.id) && seen.add(a.id));
    if (accounts.length > 50) return { ok: false, msg: '帳戶最多 50 個' };
    const nameA = clean(cfg.nameA) || '我的資產';
    const nameB = clean(cfg.nameB) || '第二份紀錄';
    setCfgValue_('WEALTH_BOOK_A_NAME', nameA);
    setCfgValue_('WEALTH_BOOK_B_NAME', nameB);
    setCfgValue_('WEALTH_B_ACCOUNTS', JSON.stringify(accounts));
    return { ok: true, nameA, nameB, accounts };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function ensureWealthBSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(WEALTH_B_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(WEALTH_B_SHEET_);
    sh.getRange(1, 1, 1, WEALTH_B_HEADERS_.length).setValues([WEALTH_B_HEADERS_])
      .setFontWeight('bold').setBackground('#344e41').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** data: { values:{帳戶ID:原幣金額}, rate_usd, rate_jpy, rate_cny, period, note } */
function api_saveWealthSnapshotB(data) {
  try {
    data = data || {};
    const cfg = readWealthBooksCfg_();
    if (!cfg.accounts.length) return { ok: false, msg: '請先在「管理帳戶」新增帳戶' };
    const rates = { TWD: 1, USD: Number(data.rate_usd) || 32, JPY: Number(data.rate_jpy) || 0.22, CNY: Number(data.rate_cny) || 4.4 };
    const vals  = data.values || {};
    const now   = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    let total = 0;
    const rows = cfg.accounts.map(a => {
      const amt  = Number(vals[a.id]) || 0;
      const rate = rates[a.currency] || 1;
      const twd  = Math.round(amt * rate);
      total += twd;
      return [now, data.period || '', a.id, a.name, a.currency, a.type === 'invest' ? '投資' : '存款', amt, rate, twd, data.note || ''];
    });
    const sh = ensureWealthBSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, WEALTH_B_HEADERS_.length).setValues(rows);
    return { ok: true, msg: '快照已儲存', total };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/** 依「記錄日期」把多列組回一筆快照 */
function api_getWealthHistoryB() {
  try {
    const sh = getSS_().getSheetByName(WEALTH_B_SHEET_);
    if (!sh || sh.getLastRow() < 2) return { ok: true, data: [], last: null };
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, WEALTH_B_HEADERS_.length).getValues();
    const map = new Map();
    rows.forEach(r => {
      const key = r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : String(r[0] || '').trim();
      const id  = String(r[2] || '').trim();
      if (!key || !id) return;
      if (!map.has(key)) map.set(key, { date: key, period: String(r[1] || ''), note: String(r[9] || ''), values: {}, twd: {}, accounts: {}, totalTWD: 0 });
      const s = map.get(key);
      const twd = Number(r[8]) || 0;
      s.values[id]   = Number(r[6]) || 0;
      s.twd[id]      = twd;
      s.accounts[id] = { name: String(r[3] || ''), currency: String(r[4] || 'TWD'), type: r[5] === '投資' ? 'invest' : 'bank' };
      s.totalTWD    += twd;
    });
    const data = [...map.values()].sort((a, b) => a.date.localeCompare(b.date));
    return { ok: true, data, last: data[data.length - 1] || null };
  } catch (e) { return { ok: false, msg: e.message, data: [], last: null }; }
}

/* ============================================================
   資產追蹤 v2：兩份紀錄都用「機構（銀行／證券商）→ 帳戶」的自訂結構
   - 設定：《設定》WEALTH_CONFIG（JSON：{ books: { A: {name, institutions}, B: {...} } }）
   - 快照：《資產快照明細》一個帳戶一列（改名、刪除都不影響舊紀錄）
   - 第一次讀取時自動把舊版《資產快照》《資產快照2》轉過來；舊工作表保留當備份、不再寫入
   ============================================================ */
var WEALTH_V2_SHEET_   = '資產快照明細';
var WEALTH_V2_HEADERS_ = ['記錄日期','紀錄','機構ID','機構','機構類別','帳戶ID','帳戶','幣別','類型','金額','匯率','台幣金額','備註'];
var WEALTH_COLORS_     = ['#4a7a9b','#4a8a5e','#b85c3a','#c4883a','#7a6545','#5a7a6a','#9a4a4a','#6a4a8a','#3a7a7a','#8a6a3a'];
var WEALTH_MAX_BOOKS_  = 10; // 資產紀錄最多幾份（預設 1 份，用「新增紀錄」加）
var WEALTH_DEFAULT_NAMES_ = { A: '我的資產', B: '第二份紀錄' };
function validBookId_(id) { return /^[A-Za-z0-9_-]{1,20}$/.test(String(id || '')); }
// 舊版第一份紀錄寫死的機構與帳戶（只用於轉換舊資料；新使用者一開始是空白）
var WEALTH_LEGACY_A_ = [
  { id: 'ctbc',        name: '中國信託',          kind: 'bank',   color: '#4a7a9b', accounts: [['ctbc_twd_saving','台幣活存','TWD'],['ctbc_twd_fixed','台幣定存','TWD'],['ctbc_usd_saving','美金活存','USD'],['ctbc_usd_fixed','美金定存','USD'],['ctbc_cny_saving','人民幣活存','CNY'],['ctbc_cny_fixed','人民幣定存','CNY']] },
  { id: 'cathay_bank', name: '國泰世華',          kind: 'bank',   color: '#4a8a5e', accounts: [['cathay_twd','台幣活存','TWD'],['cathay_usd','美金活存','USD']] },
  { id: 'taishin',     name: '台新銀行',          kind: 'bank',   color: '#b85c3a', accounts: [['taishin_twd','台幣活存','TWD'],['taishin_jpy','日幣活存','JPY']] },
  { id: 'richart',     name: 'Richart（台新數位）', kind: 'bank', color: '#c4883a', accounts: [['richart_twd','台幣活存','TWD'],['richart_fund','基金市值','TWD']] },
  { id: 'chang',       name: '彰銀',              kind: 'bank',   color: '#7a6545', accounts: [['chang_twd','台幣活存','TWD']] },
  { id: 'land',        name: '合庫',              kind: 'bank',   color: '#5a7a6a', accounts: [['land_twd','台幣活存','TWD']] },
  { id: 'post',        name: '郵局',              kind: 'bank',   color: '#9a4a4a', accounts: [['post_twd','台幣活存','TWD']] },
  { id: 'uni_sec',     name: '統一證券',          kind: 'broker', color: '#6a4a8a', accounts: [['uni_stock','台股總市值','TWD']] },
  { id: 'cathay_sec',  name: '國泰證券',          kind: 'broker', color: '#3a7a7a', accounts: [['cathay_stock','台股總市值','TWD'],['cathay_us_stock','美股總市值','USD']] },
];
var WEALTH_LEGACY_INVEST_ = ['uni_stock', 'cathay_stock', 'cathay_us_stock', 'richart_fund'];

function wealthRate_(rates, cur) {
  if (cur === 'TWD') return 1;
  return Number(rates && rates[cur]) || ({ USD: 32, JPY: 0.22, CNY: 4.4 })[cur] || 1;
}
function wealthDateKey_(v, tz) {
  return v instanceof Date ? Utilities.formatDate(v, tz || 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : String(v || '').trim();
}

/** 清理前端送來的一份紀錄設定 */
function sanitizeWealthBook_(book, fallbackName) {
  const clean = s => String(s == null ? '' : s).trim().slice(0, 30);
  const toId  = s => String(s == null ? '' : s).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 30);
  const seenI = new Set(), seenA = new Set();
  const insts = (Array.isArray(book && book.institutions) ? book.institutions : []).map((it, i) => {
    it = it || {};
    const kind = it.kind === 'broker' ? 'broker' : 'bank';
    return {
      id: toId(it.id), name: clean(it.name), kind,
      color: /^#[0-9a-fA-F]{6}$/.test(String(it.color || '')) ? it.color : WEALTH_COLORS_[i % WEALTH_COLORS_.length],
      accounts: (Array.isArray(it.accounts) ? it.accounts : []).map(a => {
        a = a || {};
        return { id: toId(a.id), name: clean(a.name),
                 currency: WEALTH_CURRENCIES_.includes(a.currency) ? a.currency : 'TWD',
                 type: a.type === 'invest' || a.type === 'bank' ? a.type : (kind === 'broker' ? 'invest' : 'bank') };
      }).filter(a => a.id && a.name && !seenA.has(a.id) && seenA.add(a.id)),
    };
  }).filter(it => it.id && it.name && !seenI.has(it.id) && seenI.add(it.id));
  return { name: clean(book && book.name) || fallbackName, institutions: insts };
}

function readWealthConfig_() {
  const raw = getCfg_().WEALTH_CONFIG;
  if (!raw) return null;
  try { const c = JSON.parse(raw); return c && c.books ? c : null; } catch (e) { return null; }
}

function ensureWealthV2Sheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(WEALTH_V2_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(WEALTH_V2_SHEET_);
    sh.getRange(1, 1, 1, WEALTH_V2_HEADERS_.length).setValues([WEALTH_V2_HEADERS_])
      .setFontWeight('bold').setBackground('#344e41').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 第一次使用 v2：建立設定，並把舊工作表的快照轉成《資產快照明細》 */
function migrateWealthToV2_() {
  const lock = getScriptLock_();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');
  try {
    const existing = readWealthConfig_();
    if (existing) return existing;
    const C = getCfg_(), ss = getSS_(), tz = C.TZ || 'Asia/Taipei';
    const out = [];
    const push = (date, book, inst, acc, amt, rate, twd, note) =>
      out.push([date, book, inst.id, inst.name, inst.kind === 'broker' ? '證券商' : '銀行', acc.id, acc.name, acc.currency,
                acc.type === 'invest' ? '投資' : '存款', amt, rate, twd, note || '']);

    // 第一份：舊《資產快照》有資料才帶入原本的機構清單
    let instA = [];
    const shA = ss.getSheetByName('資產快照');
    if (shA && shA.getLastRow() > 1) {
      instA = WEALTH_LEGACY_A_.map(it => ({ id: it.id, name: it.name, kind: it.kind, color: it.color,
        accounts: it.accounts.map(([id, name, currency]) => ({ id, name, currency, type: WEALTH_LEGACY_INVEST_.includes(id) ? 'invest' : 'bank' })) }));
      // ★ 只讀不寫：依表頭名稱找欄位，完全不修改舊工作表（不呼叫會調整欄位的 ensureWealthSheet_）
      const lastCol = shA.getLastColumn();
      const hdr = shA.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
      const colOf = label => hdr.indexOf(label);
      const colByKey = {};
      WEALTH_KEYS_.forEach((k, i) => {
        let c = colOf(WEALTH_SHEET_HEADERS_[i + 2]);
        if (c < 0 && k === 'cathay_us_stock') c = colOf('國泰美股'); // 更早期的欄位名稱
        colByKey[k] = c;
      });
      const noteCol = colOf('備註');
      const vals = shA.getRange(2, 1, shA.getLastRow() - 1, lastCol).getValues();
      vals.forEach(row => {
        const date = wealthDateKey_(row[0], tz);
        if (!date) return;
        const o = {};
        WEALTH_KEYS_.forEach(k => { o[k] = colByKey[k] >= 0 ? (Number(row[colByKey[k]]) || 0) : 0; });
        const rates = { USD: o.rate_usd, JPY: o.rate_jpy, CNY: o.rate_cny };
        const note = noteCol >= 0 ? String(row[noteCol] || '') : '';
        instA.forEach(inst => inst.accounts.forEach(acc => {
          const amt = o[acc.id] || 0, rate = wealthRate_(rates, acc.currency);
          push(date, 'A', inst, acc, amt, rate, Math.round(amt * rate), note);
        }));
      });
    }

    // 第二份：舊版是單層帳戶，依「存款／投資」分成兩個機構
    const oldB = readWealthBooksCfg_();
    const instOf = type => type === 'invest'
      ? { id: 'b_invest', name: '投資帳戶', kind: 'broker', color: WEALTH_COLORS_[1] }
      : { id: 'b_bank',   name: '存款帳戶', kind: 'bank',   color: WEALTH_COLORS_[0] };
    const instB = [];
    ['bank', 'invest'].forEach(type => {
      const accs = oldB.accounts.filter(a => (a.type === 'invest' ? 'invest' : 'bank') === type)
        .map(a => ({ id: a.id, name: a.name, currency: a.currency, type }));
      if (accs.length) instB.push(Object.assign(instOf(type), { accounts: accs }));
    });
    (api_getWealthHistoryB().data || []).forEach(snap => {
      Object.keys(snap.values || {}).forEach(id => {
        const a = (snap.accounts || {})[id] || {};
        const acc = { id, name: a.name || id, currency: a.currency || 'TWD', type: a.type === 'invest' ? 'invest' : 'bank' };
        const amt = Number(snap.values[id]) || 0, twd = Number((snap.twd || {})[id]) || 0;
        push(snap.date, 'B', instOf(acc.type), acc, amt, amt ? Math.round(twd / amt * 10000) / 10000 : wealthRate_({}, acc.currency), twd, snap.note);
      });
    });

    if (out.length) {
      const sh = ensureWealthV2Sheet_();
      sh.getRange(sh.getLastRow() + 1, 1, out.length, WEALTH_V2_HEADERS_.length).setValues(out);
    }
    // 預設只有一份紀錄；舊版第二份有帳戶或快照才保留
    const cfg = { order: ['A'], books: { A: { name: C.WEALTH_BOOK_A_NAME || WEALTH_DEFAULT_NAMES_.A, institutions: instA } } };
    if (instB.length || out.some(r => r[1] === 'B')) {
      cfg.order.push('B');
      cfg.books.B = { name: oldB.nameB || WEALTH_DEFAULT_NAMES_.B, institutions: instB };
    }
    setCfgValue_('WEALTH_CONFIG', JSON.stringify(cfg));
    Logger.log(`資產追蹤 v2：已轉換 ${out.length} 列舊快照`);
    return cfg;
  } finally { lock.releaseLock(); }
}

function getWealthConfig_() { return readWealthConfig_() || migrateWealthToV2_(); }

/** 《資產快照明細》依「紀錄 + 記錄日期」組回一筆筆快照 */
function readWealthV2History_() {
  const res = {};
  const sh = getSS_().getSheetByName(WEALTH_V2_SHEET_);
  if (!sh || sh.getLastRow() < 2) return res;
  const tz = getCfg_().TZ || 'Asia/Taipei';
  const map = new Map();
  sh.getRange(2, 1, sh.getLastRow() - 1, WEALTH_V2_HEADERS_.length).getValues().forEach(r => {
    const date = wealthDateKey_(r[0], tz), book = String(r[1] || '').trim(), id = String(r[5] || '').trim();
    if (!date || !validBookId_(book) || !id) return;
    const key = book + '|' + date;
    if (!map.has(key)) map.set(key, { book, date, note: String(r[12] || ''), values: {}, twd: {}, accounts: {}, rates: {}, totalTWD: 0 });
    const s = map.get(key), cur = String(r[7] || 'TWD'), twd = Number(r[11]) || 0;
    s.values[id] = Number(r[9]) || 0;
    s.twd[id] = twd;
    s.accounts[id] = { name: String(r[6] || ''), instId: String(r[2] || ''), instName: String(r[3] || ''),
                       kind: r[4] === '證券商' ? 'broker' : 'bank', currency: cur, type: r[8] === '投資' ? 'invest' : 'bank' };
    if (cur !== 'TWD' && Number(r[10])) s.rates[cur] = Number(r[10]);
    s.totalTWD += twd;
  });
  [...map.values()].sort((a, b) => a.date.localeCompare(b.date)).forEach(s => { const b = s.book; delete s.book; (res[b] = res[b] || []).push(s); });
  return res;
}

/** 資產追蹤頁一次拿齊：各份紀錄的設定（依 order 排列）+ 歷史快照 */
function api_getWealth() {
  try {
    const cfg = getWealthConfig_();
    const hist = readWealthV2History_();
    if (!Array.isArray(cfg.order)) {
      // v9 的設定固定有 A、B 兩份：改成「預設一份」，空白的其他紀錄直接拿掉
      cfg.order = Object.keys(cfg.books).filter(id =>
        id === 'A' || (cfg.books[id].institutions || []).length || (hist[id] || []).length);
      if (!cfg.order.length) cfg.order = [Object.keys(cfg.books)[0] || 'A'];
      Object.keys(cfg.books).forEach(id => { if (!cfg.order.includes(id)) delete cfg.books[id]; });
      if (!cfg.books[cfg.order[0]]) cfg.books[cfg.order[0]] = { name: WEALTH_DEFAULT_NAMES_.A, institutions: [] };
      setCfgValue_('WEALTH_CONFIG', JSON.stringify(cfg));
    }
    const history = {};
    cfg.order.forEach(id => { history[id] = hist[id] || []; });
    return { ok: true, order: cfg.order, books: cfg.books, history };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function wealthOrder_(cfg) { return Array.isArray(cfg.order) ? cfg.order : Object.keys(cfg.books); }

/** 刪除一份紀錄（設定拿掉；《資產快照明細》裡的歷史資料保留） */
function api_deleteWealthBook(bookId) {
  try {
    const cfg = getWealthConfig_();
    const order = wealthOrder_(cfg);
    if (!order.includes(bookId)) return { ok: false, msg: '找不到這份紀錄' };
    if (order.length <= 1) return { ok: false, msg: '至少要保留一份紀錄' };
    cfg.order = order.filter(id => id !== bookId);
    delete cfg.books[bookId];
    setCfgValue_('WEALTH_CONFIG', JSON.stringify(cfg));
    return { ok: true, order: cfg.order, books: cfg.books };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/** 儲存某一份紀錄的名稱與機構／帳戶設定 */
function api_saveWealthConfig(bookId, book) {
  try {
    if (!validBookId_(bookId)) return { ok: false, msg: '紀錄代號錯誤' };
    const cfg = getWealthConfig_();
    const order = wealthOrder_(cfg);
    const isNew = !order.includes(bookId);
    if (isNew && order.length >= WEALTH_MAX_BOOKS_) return { ok: false, msg: `資產紀錄最多 ${WEALTH_MAX_BOOKS_} 份` };
    const clean = sanitizeWealthBook_(book, WEALTH_DEFAULT_NAMES_[bookId] || '新的紀錄');
    const nAcc = clean.institutions.reduce((s, it) => s + it.accounts.length, 0);
    if (clean.institutions.length > 30 || nAcc > 100) return { ok: false, msg: '機構最多 30 個、帳戶最多 100 個' };
    cfg.books[bookId] = clean;
    cfg.order = isNew ? order.concat(bookId) : order;
    setCfgValue_('WEALTH_CONFIG', JSON.stringify(cfg));
    return { ok: true, order: cfg.order, books: cfg.books };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/** data: { values:{帳戶ID:原幣金額}, rate_usd, rate_jpy, rate_cny, note } */
function api_saveWealthSnapshotV2(bookId, data) {
  try {
    data = data || {};
    const book = validBookId_(bookId) ? getWealthConfig_().books[bookId] : null;
    if (!book) return { ok: false, msg: '找不到這份紀錄' };
    const pairs = [];
    (book.institutions || []).forEach(inst => (inst.accounts || []).forEach(acc => pairs.push([inst, acc])));
    if (!pairs.length) return { ok: false, msg: '請先在「管理」新增銀行或證券商與帳戶' };
    const rates = { USD: data.rate_usd, JPY: data.rate_jpy, CNY: data.rate_cny };
    const date = Utilities.formatDate(new Date(), getCfg_().TZ || 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    const vals = data.values || {};
    let total = 0;
    const rows = pairs.map(([inst, acc]) => {
      const amt = Number(vals[acc.id]) || 0, rate = wealthRate_(rates, acc.currency), twd = Math.round(amt * rate);
      total += twd;
      return [date, bookId, inst.id, inst.name, inst.kind === 'broker' ? '證券商' : '銀行', acc.id, acc.name, acc.currency,
              acc.type === 'invest' ? '投資' : '存款', amt, rate, twd, data.note || ''];
    });
    const sh = ensureWealthV2Sheet_();
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, WEALTH_V2_HEADERS_.length).setValues(rows);
    return { ok: true, msg: '快照已儲存', total };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* --- 月報 --- */
/** 立刻寄一份月報。ym 例：'2026/09'；沒給就寄上個月 */
function api_sendMonthlyReport(ym) {
  try {
    let y, m;
    const mt = String(ym || '').match(/^(\d{4})[\/-](\d{1,2})$/);
    if (mt) { y = Number(mt[1]); m = Number(mt[2]); }
    else {
      const tz = getCfg_().TZ || 'Asia/Taipei', now = new Date();
      y = Number(Utilities.formatDate(now, tz, 'yyyy')); m = Number(Utilities.formatDate(now, tz, 'M')) - 1;
      if (m === 0) { y -= 1; m = 12; }
    }
    if (m < 1 || m > 12) return { ok: false, msg: '月份格式錯誤' };
    const r = sendMonthlyReport_(y, m);
    return { ok: true, msg: `已寄出 ${y} 年 ${m} 月月報到 ${r.to}` };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* --- 自動更新（線上載入模式） --- */
function api_getUpdateStatus() {
  if (!LOADER_) return { ok: true, supported: false };
  const tz = getCfg_().TZ || 'Asia/Taipei', L = LOADER_;
  const pending = L.latest > L.version;
  return { ok: true, supported: true, enabled: true, status: {
    status: pending ? 'waiting' : 'latest',
    msg: pending
      ? `目前 v${L.version}，新版 v${L.latest} 將於 ${Utilities.formatDate(new Date(L.readyAt), tz, 'MM/dd HH:mm')} 後自動生效`
      : `已是最新版本 v${L.version}`,
    at: Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm') } };
}
/** 設定頁「立即套用最新版」：不等 1 天（殼程式讀 FIN_SKIP_DELAY） */
function api_runAutoUpdate() {
  if (!LOADER_) return { ok: false, status: 'unsupported', msg: '這個後端不是範本殼程式，不需要自動更新' };
  const tz = getCfg_().TZ || 'Asia/Taipei', L = LOADER_;
  const at = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm');
  if (!(L.latest > L.version)) return { ok: true, status: 'latest', msg: `已是最新版本 v${L.version}`, at };
  getProps_().setProperty('FIN_SKIP_DELAY', String(L.latest));
  return { ok: true, status: 'updated', msg: `已套用最新版 v${L.latest}，重新整理頁面後生效`, at };
}

/* --- 每季/每半年 Email 提醒 --- */
function wealthReminder() {
  const props = getProps_();
  const type  = props.getProperty('WEALTH_REMINDER_TYPE') || 'quarterly';
  const month = new Date().getMonth() + 1;
  const active = type === 'halfyear' ? [1,7] : [1,4,7,10];
  if (!active.includes(month)) return;

  const year   = new Date().getFullYear();
  const qLabel = month<=3?'Q1':month<=6?'Q2':month<=9?'Q3':'Q4';
  const label  = type === 'halfyear' ? (month<=6?'上半年':'下半年') : qLabel;
  const url    = getScriptApp_().getService().getUrl();
  const msg    = `【資產記帳提醒】${year} ${label} 到了！\n💰 記得記錄這期的總資產\n\n開啟 App：${url}`;
  try { const C=getCfg_(); if(C['ALERT_TO']) MailApp.sendEmail(C['ALERT_TO'],'【資產記帳提醒】'+year+' '+label, msg); } catch(e){}
}

function api_setupWealthTrigger(type) {
  // type: 'quarterly' | 'halfyear' | 'none'
  try {
    getScriptApp_().getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'wealthReminder')
      .forEach(t => getScriptApp_().deleteTrigger(t));
    if (type !== 'none') {
      getScriptApp_().newTrigger('wealthReminder').timeBased().onMonthDay(1).atHour(9).create();
      getProps_().setProperty('WEALTH_REMINDER_TYPE', type);
    } else {
      getProps_().deleteProperty('WEALTH_REMINDER_TYPE');
    }
    const label = type==='quarterly'?'每季（1/4/7/10月）':type==='halfyear'?'每半年（1/7月）':'已關閉';
    return { ok: true, msg: '提醒設定：' + label };
  } catch(e) { return { ok: false, msg: e.message }; }
}

/* ============================================================
   股票代碼查名稱
   ============================================================ */

function api_lookupStockName(code) {
  if (!code) return { ok: true, name: '' };
  const C   = getCfg_();
  const ss  = getSS_();
  const key = String(code).trim();
  const sheetNames = [C.SHEET_HOLD, C.SHEET_TRADES, C.SHEET_OPENING];
  for (const shName of sheetNames) {
    const sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2) continue;
    const rows = readSheetAsObjects_(sh);
    const found = rows.find(r => String(r['股票代碼'] || '').trim().replace(/^'+/, '') === key);
    if (found && found['股票名稱']) return { ok: true, name: String(found['股票名稱']).trim() };
  }
  return { ok: true, name: '' };
}

/* ============================================================
   手動新增交易
   ============================================================ */

function api_addManualTrade(trade) {
  const C  = getCfg_();
  const ss = getSS_();
  const sh = ensureSheetWithHeader_(C.SHEET_TRADES || '交易紀錄', [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','備註','證券商'
  ]);

  const type   = String(trade.type   || '現買').trim();
  const qty    = parseFloat(trade.qty)   || 0;
  const price  = parseFloat(trade.price) || 0;
  const amount = qty * price;

  const feeOverride = (trade.fee !== '' && trade.fee != null) ? parseFloat(trade.fee) : null;
  const taxOverride = (trade.tax !== '' && trade.tax != null) ? parseFloat(trade.tax) : null;

  const feeCalc = Math.max(20, Math.round(amount * 0.001425 * (Number(C.FEE_DISCOUNT) || 0.28)));
  const fee = feeOverride !== null ? feeOverride : feeCalc;

  const taxRate = type === '沖賣' ? 0.0015 : (type.includes('賣') ? 0.003 : 0);
  const taxCalc = Math.round(amount * taxRate);
  const tax = taxOverride !== null ? taxOverride : taxCalc;

  const net = type.includes('賣') ? (amount - fee - tax) : -(amount + fee);

  const dateStr = String(trade.date || '').replace(/-/g, '/');

  // 按欄位名稱寫入，避免受現有試算表欄位順序影響
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
    : [];
  const rowData = new Array(Math.max(lastCol, 14)).fill('');
  const setCol  = (name, val) => {
    const i = headers.indexOf(name);
    if (i >= 0) rowData[i] = val;
  };
  setCol('成交日期',   dateStr);
  setCol('成交時間',   trade.time   || '');
  setCol('股票代碼',   String(trade.code || '').trim());
  setCol('股票名稱',   trade.name   || '');
  setCol('成交類別',   type);
  setCol('股數',       qty);
  setCol('成交價',     price);
  setCol('成交金額',   amount);
  setCol('委託單號',   '');
  setCol('手續費',     fee);
  setCol('交易稅',     tax);
  setCol('淨收付金額', net);
  setCol('證券商',     trade.broker || '');
  setCol('備註',       trade.note   || '手動新增');

  // 用 setValues 取代 appendRow，並在寫入前把股票代碼欄設為文字格式
  // 防止 Google Sheets 把 "0056" 自動轉為數字 56
  const newRow     = sh.getLastRow() + 1;
  const codeColIdx = headers.indexOf('股票代碼');
  if (codeColIdx >= 0) {
    sh.getRange(newRow, codeColIdx + 1).setNumberFormat('@');
  }
  sh.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

  return { ok: true, msg: `${String(trade.code).trim()} ${type} ${qty}股 已新增至交易紀錄` };
}

/* ============================================================
   共用工具
   ============================================================ */

/** 讀取 Sheet 轉為物件陣列 */
function readSheetAsObjects_(sh) {
  if (!sh) return [];
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const data    = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(row => {
    const o = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) val = Utilities.formatDate(val, 'Asia/Taipei', 'yyyy/MM/dd');
      o[h] = val;
    });
    return o;
  });
}

/** Gemini Key 狀態（對應前端呼叫） */
function api_getGeminiKeyStatus() {
  const k = getProps_().getProperty('GEMINI_API_KEY') || '';
  return { ok: true, hasKey: !!k, masked: k ? k.slice(0, 8) + '...' : '' };
}

/** 儲存 Gemini Key（對應前端呼叫） */
function api_saveGeminiKey(key) {
  if (!key) return { ok: false, msg: 'Key 不能為空' };
  getProps_().setProperty('GEMINI_API_KEY', key.trim());
  return { ok: true };
}

/* ============================================================
   待確認交易
   ============================================================ */

function ensureStagingSheet_() {
  const C  = getCfg_();
  const ss = getSS_();
  const name = C.SHEET_STAGING || '待確認交易';
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['成交日期','成交時間','股票代碼','股票名稱','成交類別',
                  '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','證券商','備註','確認狀態']);
  }
  return sh;
}

function api_getPendingTrades() {
  const sh   = ensureStagingSheet_();
  const rows = readSheetAsObjects_(sh);
  const pending = rows
    .map((r, i) => ({ ...r, _rowIndex: i + 2 }))
    .filter(r => !r['確認狀態'] || r['確認狀態'] === '待確認');
  return { ok: true, rows: pending };
}

function api_confirmTrades(confirmedRows) {
  const C  = getCfg_();
  const shStaging = ensureStagingSheet_();
  const shTrades  = ensureSheetWithHeader_(C.SHEET_TRADES || '交易紀錄', [
    '成交日期','成交時間','股票代碼','股票名稱','成交類別',
    '股數','成交價','成交金額','委託單號','手續費','交易稅','淨收付金額','備註','證券商'
  ]);

  // 一次讀取所有 staging 資料（含狀態欄），避免逐列讀取
  const lastRow = shStaging.getLastRow();
  if (lastRow < 2) return { ok: true, msg: '0 筆交易已確認寫入' };
  const allStaging = shStaging.getRange(2, 1, lastRow - 1, 15).getValues();

  // 整批收集要寫入的列，再一次 setValues 並預設代碼欄為文字格式
  const toWrite   = [];
  const confirmed = [];
  for (const row of (confirmedRows || [])) {
    const idx = row._rowIndex - 2;
    const sr  = allStaging[idx];
    if (!sr) continue;
    const status = String(sr[14] || '');
    if (status === '已刪除' || status === '已確認') continue;

    const side   = String(row['成交類別'] || sr[4]);
    const amount = Number(sr[7]) || 0;
    const fee    = Number(sr[9]) || 0;
    const tax    = (row['交易稅'] !== undefined && row['交易稅'] !== null)
                   ? Number(row['交易稅']) : Number(sr[10]) || 0;
    const net    = side.includes('賣') ? amount - fee - tax : -(amount + fee);

    toWrite.push([
      sr[0], sr[1],
      String(sr[2]),   // 股票代碼：保持字串，避免 0056/00919 被截掉前導零
      sr[3], side,
      Number(sr[5]), Number(sr[6]), amount,
      sr[8] || '',
      fee, tax, net,
      sr[13] || '', sr[12] || '',  // TRADES 欄位順序：備註(staging[13]), 証券商(staging[12])
    ]);
    confirmed.push(row._rowIndex);
  }

  if (toWrite.length) {
    const startRow = shTrades.getLastRow() + 1;
    const n = toWrite.length;
    // 先把代碼欄（col 3）與委託單號欄（col 9）設為文字格式，防止前導零被截
    shTrades.getRange(startRow, 3, n, 1).setNumberFormat('@');
    shTrades.getRange(startRow, 9, n, 1).setNumberFormat('@');
    shTrades.getRange(startRow, 1, n, toWrite[0].length).setValues(toWrite);
    confirmed.forEach(idx => shStaging.getRange(idx, 15).setValue('已確認'));
  }
  return { ok: true, msg: `${toWrite.length} 筆交易已確認寫入` };
}

function api_deletePendingTrade(rowIndex) {
  ensureStagingSheet_().getRange(rowIndex, 15).setValue('已刪除');
  return { ok: true };
}

function api_batchDeletePendingTrades(rowIndices) {
  if (!Array.isArray(rowIndices) || !rowIndices.length) return { ok: true, count: 0 };
  const sh = ensureStagingSheet_();
  rowIndices.forEach(idx => sh.getRange(idx, 15).setValue('已刪除'));
  return { ok: true, count: rowIndices.length };
}

function api_updatePendingTrade(rowIndex, fields) {
  const sh = ensureStagingSheet_();
  if (!rowIndex || rowIndex < 2 || rowIndex > sh.getLastRow())
    return { ok: false, msg: '無效的列索引' };
  const status = String(sh.getRange(rowIndex, 15).getValue());
  if (status === '已刪除' || status === '已確認')
    return { ok: false, msg: `此交易已${status}，無法修改` };

  const existing = sh.getRange(rowIndex, 1, 1, 14).getValues()[0];
  const side   = String(fields['成交類別'] || existing[4]);
  const qty    = Number(fields['股數'])    || Number(existing[5]) || 0;
  const price  = Number(fields['成交價'])  || Number(existing[6]) || 0;
  const amount = Math.round(qty * price * 100) / 100;
  const fee    = fields['手續費'] !== undefined ? Number(fields['手續費']) : Number(existing[9]);
  const tax    = fields['交易稅'] !== undefined ? Number(fields['交易稅']) : Number(existing[10]);
  const net    = side.includes('賣') ? amount - fee - tax : -(amount + fee);

  sh.getRange(rowIndex, 3).setNumberFormat('@'); // 股票代碼保持文字格式
  sh.getRange(rowIndex, 1, 1, 14).setValues([[
    fields['成交日期'] || existing[0],
    existing[1],
    String(fields['股票代碼'] !== undefined ? fields['股票代碼'] : existing[2]),
    fields['股票名稱'] !== undefined ? fields['股票名稱'] : existing[3],
    side,
    qty, price, amount,
    existing[8],
    fee, tax, net,
    fields['証券商'] !== undefined ? fields['証券商'] : existing[12],
    fields['備註']   !== undefined ? fields['備註']   : existing[13],
  ]]);
  return { ok: true, msg: '已更新', data: { '成交金額': amount, '手續費': fee, '交易稅': tax, '淨收付金額': net } };
}

/* ============================================================
   doPost — 統一 API 入口
   ============================================================ */

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    const args   = body.args || [];

    // 資產追蹤舊版 API（《資產快照》《資產快照2》）已停用：避免舊頁面把快照寫進舊工作表、新頁面看不到
    const ALLOWED = new Set([
      'api_getDashboard', 'api_getDataList', 'api_getSettingsSchema',
      'api_saveSettings', 'api_getHoldingsForAnalysis', 'api_callGemini',
      'api_saveAnalysis', 'api_getAnalysisHistory', 'api_getGeminiKeyStatus',
      'api_saveGeminiKey',
      'api_setupWealthTrigger', 'api_sendMonthlyReport', 'api_getUpdateStatus', 'api_runAutoUpdate',
      'api_getWealth', 'api_saveWealthConfig', 'api_saveWealthSnapshotV2', 'api_deleteWealthBook',
      'ingestFromGmail_Plaintext_SAFE', 'rebuildAll_B_SAFE',
      'api_runDividendsUpdate', 'appendDCAFromHoldings_SAFE',
      'rebuildRealizedPnL_FIFO_SAFE', 'rebuildDCADividends_SAFE',
      'api_getPendingTrades', 'api_confirmTrades', 'api_deletePendingTrade', 'api_batchDeletePendingTrades',
      'api_updatePendingTrade',
      'api_addManualTrade', 'api_lookupStockName',
      'api_login', 'api_auth_token', 'api_getSetupStatus', 'api_setupAccount', 'api_changePassword',
    ]);
    // 不需登入即可呼叫
    const PUBLIC = new Set(['api_login', 'api_auth_token', 'api_getSetupStatus', 'api_setupAccount']);

    if (!ALLOWED.has(action)) {
      output.setContent(JSON.stringify({ ok: false, msg: '不允許的 action: ' + action }));
      return output;
    }
    if (!PUBLIC.has(action) && !resolveAuth_(body.token)) {
      output.setContent(JSON.stringify({ ok: false, authError: true, msg: '登入已失效，請重新登入' }));
      return output;
    }
    const fn = this[action];
    if (typeof fn !== 'function') {
      output.setContent(JSON.stringify({ ok: false, msg: '找不到函數: ' + action }));
      return output;
    }
    const result = fn.apply(this, args);
    output.setContent(JSON.stringify(result ?? { ok: true }));
  } catch(err) {
    output.setContent(JSON.stringify({ ok: false, msg: err.message || String(err) }));
  }
  return output;
}

function api_runDividendsUpdate() {
  getProps_().deleteProperty('DIV_CURSOR');
  getScriptApp_().newTrigger('runDividendsFullCycle_SAFE').timeBased().at(new Date(Date.now() + 3000)).create();
  return { ok: true };
}

