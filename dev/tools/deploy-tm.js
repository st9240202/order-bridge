#!/usr/bin/env node
// deploy-tm.js — 把 user.js 寫進本機 Tampermonkey(透過 CDP 開 TM dashboard 頁寫 extdb)
// 用法: node dev/tools/deploy-tm.js [path-to-user.js]
// 環境: Chrome 需以 --user-data-dir=~/dev/chrome-real --remote-debugging-port=9333 啟動
const { chromium } = require('playwright-core');
const fs = require('fs');
const path = require('path');

const TM_EXT = 'dhdgffkkebhmkfjojejmpbldmpobfkfo';
const CDP = process.env.CDP || 'http://localhost:9333';
const file = process.argv[2] || path.join(__dirname, '..', '..', 'order-bridge.user.js');
const src = fs.readFileSync(file, 'utf8');
const m = src.match(/\/\/\s*@name\s+(.*)/);
if (!m) { console.error('找不到 @name'); process.exit(1); }
const name = m[1].trim();
const vm = src.match(/\/\/\s*@version\s+([\d.]+)/);
const version = vm ? vm[1] : '0.0.0';

function updateFn() {
  return (arg) => {
    const { uuid, newName, newVersion, code } = arg;
    return new Promise((res) => {
      chrome.storage.local.get(null, (all) => {
        const meta = all['!extdb.@meta#' + uuid] && all['!extdb.@meta#' + uuid].value;
        if (!meta) { res({ error: 'meta missing — 腳本可能未安裝或 uuid 錯' }); return; }
        const header = code.match(/\/\/ ==UserScript==[\s\S]*?\/\/ ==\/UserScript==\s*/)[0];
        const updatedMeta = Object.assign({}, meta, { name: newName, version: newVersion, header, lastModified: Date.now(), user_modified: Date.now() });
        chrome.storage.local.set({
          ['!extdb.@source#' + uuid]: { origin: 'tampermonkey', value: code, ts: Date.now() },
          ['!extdb.@meta#' + uuid]: { origin: 'tampermonkey', value: updatedMeta, ts: Date.now() },
          ['!extdb.@uid#' + uuid]: { origin: 'tampermonkey', value: newName, ts: Date.now() }
        }, () => res({ ok: true, name: newName, version: newVersion, srcLen: code.length }));
      });
    });
  };
}

(async () => {
  const browser = await chromium.connectOverCDP(CDP);
  const ctx = browser.contexts()[0];
  const page = await ctx.newPage();
  await page.goto(`chrome-extension://${TM_EXT}/options.html`);
  await page.waitForTimeout(2500);

  const found = await page.evaluate((arg) => {
    const { name } = arg;
    return new Promise((res) => {
      chrome.storage.local.get(null, (o) => {
        const hits = [];
        for (const k of Object.keys(o)) {
          const m = k.match(/^!extdb\.@meta#(.+)$/);
          if (!m) continue;
          const meta = o[k] && o[k].value;
          if (meta && meta.name === name) hits.push({ uuid: m[1], version: meta.version, enabled: meta.enabled });
        }
        res(hits);
      });
    });
  }, { name });

  if (!found.length) { console.error('未找到名為「' + name + '」的腳本'); await browser.close(); process.exit(1); }
  if (found.length > 1) { console.error('找到多個同名腳本:', found); await browser.close(); process.exit(1); }
  const uuid = found[0].uuid;
  console.log('找到:', JSON.stringify({ uuid, 目前版本: found[0].version, 啟用: found[0].enabled }));
  if (found[0].version === version) console.log('(版本相同,仍會重寫以同步代碼)');

  const out = await page.evaluate(updateFn(), { uuid, newName: name, newVersion: version, code: src });
  console.log(out.ok ? `✅ 已更新 TM: ${name} → v${version} (${out.srcLen} bytes)` : `❌ ${out.error}`);
  console.log('⚠️ 記得 reload 受影響的頁面才生效');
  await page.close();
  await browser.close();
})().catch((e) => { console.error('ERR:', e.message); process.exit(1); });
