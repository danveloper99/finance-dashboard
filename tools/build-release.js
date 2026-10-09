/**
 * 打包發布：把 4.0.js + WebAPI.js 打包成 releases/v{APP_VERSION}.js，並更新 release.json
 * 用法：node tools/build-release.js "這次更新的說明"
 *
 * 範本殼程式（template/Code.js）會從 GitHub Pages 下載這個檔案，用 new Function 執行，
 * 拿到回傳的物件（所有頂層函式）後再呼叫 doPost / 排程函式。
 * 新版在 releasedAt 滿 1 天後才會被殼程式採用，之前繼續用 previous。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const read = f => fs.readFileSync(path.join(ROOT, f), 'utf8').replace(/\r\n/g, '\n');
const sources = ['4.0.js', 'WebAPI.js'].map(f => ({ f, code: read(f) }));

const verMatch = read('WebAPI.js').match(/var APP_VERSION = (\d+);/);
if (!verMatch) throw new Error('WebAPI.js 找不到 APP_VERSION');
const version = Number(verMatch[1]);

// 所有頂層函式都匯出（doPost 用 this[action] 找 API，排程用名稱呼叫）
const names = [];
sources.forEach(({ code }) => {
  for (const m of code.matchAll(/^function\s+([A-Za-z0-9_$]+)\s*\(/gm)) names.push(m[1]);
});
const dup = names.filter((n, i) => names.indexOf(n) !== i);
if (dup.length) throw new Error('函式名稱重複：' + [...new Set(dup)].join(', '));
['doPost', 'setLoaderInfo', 'runTriggerTask'].forEach(n => { if (!names.includes(n)) throw new Error('缺少必要函式：' + n); });

const bundle = [
  `// @@FIN_BUNDLE v${version}`,
  `// 你不理財，才不理你 — 程式包 v${version}（由 tools/build-release.js 產生，請勿手動修改）`,
  ...sources.map(({ f, code }) => `// ===== ${f} =====\n${code}`),
  `return { __version: ${version},\n  ${names.map(n => `${n}: ${n}`).join(',\n  ')}\n};`,
  '',
].join('\n');

// 用 Function 試跑一次，確定語法正確
new Function(bundle);

fs.mkdirSync(path.join(ROOT, 'releases'), { recursive: true });
const out = path.join(ROOT, 'releases', `v${version}.js`);
fs.writeFileSync(out, bundle);

const relPath = path.join(ROOT, 'release.json');
const rel = fs.existsSync(relPath) ? JSON.parse(fs.readFileSync(relPath, 'utf8')) : {};
const now = new Date(Date.now() + 8 * 3600e3).toISOString().replace(/\.\d+Z$/, '+08:00');
const prev = rel.current && rel.current.version !== version ? { version: rel.current.version } : (rel.previous || null);
const next = {
  current: { version, releasedAt: now, notes: process.argv[2] || '' },
  previous: prev,
  libVersion: rel.libVersion, // 舊版「程式庫殼程式」用，保留
};
fs.writeFileSync(relPath, JSON.stringify(next, null, 2) + '\n');
console.log(`✅ releases/v${version}.js（${(bundle.length / 1024).toFixed(0)} KB、${names.length} 個函式）`);
console.log(`✅ release.json：current v${version}${prev ? `、previous v${prev.version}` : ''}`);
