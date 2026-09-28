#!/usr/bin/env node
// check-i18n.js — 驗證 order-bridge.user.js 四語 DICT key 完全對齊 + 所有 t('key') 都有翻譯
// 用法: node dev/check-i18n.js [path-to-user.js]
// 以「行首 6 空白 + name: {」當頂層 locale 邊界,不受字串內容干擾。
const fs = require('fs');
const path = require('path');
const file = process.argv[2] || path.join(__dirname, '..', 'order-bridge.user.js');
const src = fs.readFileSync(file, 'utf8');

const names = ['zh-TW', 'zh-CN', 'en', 'ja'];
const lines = src.split('\n');

// 找 DICT 區塊
let start = -1;
lines.forEach((ln, i) => { if (start < 0 && /^\s*var DICT = \{/.test(ln)) start = i; });
if (start < 0) { console.error('找不到 var DICT'); process.exit(1); }

const dicts = {}; // name -> {key: value}
let cur = null;
for (let i = start; i < lines.length; i++) {
  const ln = lines[i];
  const top = ln.match(/^      ('(zh-TW|zh-CN)'|en|ja):\s*\{\s*$/);
  if (top) { cur = (top[1].startsWith("'") ? top[1].slice(1, -1) : top[1]); dicts[cur] = {}; continue; }
  if (cur && /^\s*\};\s*$/.test(ln)) break; // DICT 結束
  if (cur) {
    const kv = ln.match(/^        '([a-zA-Z][a-zA-Z0-9.]*)':\s*'(.*)'\s*,?\s*$/);
    if (kv) {
      if (dicts[cur][kv[1]] !== undefined) { console.error(`❌ ${cur}: 重複 key '${kv[1]}' (行 ${i + 1})`); process.exit(1); }
      dicts[cur][kv[1]] = kv[2];
    }
  }
}

let fail = 0;
const all = new Set();
Object.values(dicts).forEach((d) => Object.keys(d).forEach((k) => all.add(k)));

for (const n of names) {
  const d = dicts[n] || {};
  const keys = Object.keys(d);
  const missing = [...all].filter((k) => !(k in d));
  if (missing.length) { fail++; console.error(`❌ ${n}: ${keys.length} keys, 缺 [${missing.join(', ')}]`); }
  else console.log(`✅ ${n}: ${keys.length} keys, 與其他語言完全對齊`);
}

// 所有 t('key') 呼叫都要有翻譯(去掉 // 行註解,避免吃到說明文字)
const used = new Set();
const re2 = /(?<![A-Za-z0-9_$])t\(\s*'([a-zA-Z][a-zA-Z0-9.]*)'/g;
lines.forEach((ln) => {
  const c = ln.indexOf('//');
  const body = c > -1 ? ln.slice(0, c) : ln;
  re2.lastIndex = 0;
  let m2;
  while ((m2 = re2.exec(body))) used.add(m2[1]);
});
const noDict = [...used].filter((k) => !(k in dicts['zh-TW']));
if (noDict.length) { fail++; console.error(`❌ 有 t('key') 呼叫但 DICT 缺 key: [${noDict.join(', ')}]`); }
else console.log(`✅ 所有 t() 呼叫都有翻譯(使用 ${used.size}/${all.size} 個 key)`);

const unused = [...all].filter((k) => !used.has(k));
if (unused.length) console.log(`ℹ️  未使用的 DICT key: [${unused.join(', ')}]`);

process.exit(fail ? 1 : 0);
