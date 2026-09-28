#!/usr/bin/env node
// e2e-taobao.js — 完整流程: 切待收貨 → FAB → 確認視窗 → 確定 → 斷言 localStorage
// 用法: node dev/tools/e2e-taobao.js [waitSeconds=25]
// 前置: 頁面已在 buyertrade.taobao.com(任意分頁)、TM 已載入最新腳本(部署後要 reload)
const { chromium } = require('playwright-core');
const CDP = process.env.CDP || 'http://localhost:9333';
const wait = Number(process.argv[2] || 25);

(async () => {
  const browser = await chromium.connectOverCDP(CDP);
  const tb = browser.contexts()[0].pages().find((p) => p.url().includes('buyertrade.taobao.com'));
  if (!tb) { console.error('找不到淘寶頁'); process.exit(1); }
  const logs = [];
  tb.on('console', (m) => logs.push(m.type() + ': ' + m.text().substring(0, 250)));
  tb.on('pageerror', (e) => logs.push('PAGEERROR: ' + String(e).substring(0, 250)));
  await tb.bringToFront();

  const pre = await tb.evaluate(() => (localStorage.getItem('ob_bridge_v1') || ''));
  console.log('before bridge:', pre.substring(0, 120));

  // 1. 切到「待收貨」分頁
  await tb.evaluate(() => {
    const el = Array.from(document.querySelectorAll('span')).find((e) => /^待收货/.test(e.textContent.replace(/\s/g, '')) && e.textContent.replace(/\s/g, '').length < 12);
    if (el) el.click();
  });
  await new Promise((r) => setTimeout(r, 3500));

  // 2. 點 FAB
  const fab = await tb.$('#ob-src-fab');
  if (!fab) { console.error('找不到 #ob-src-fab'); process.exit(1); }
  await fab.click();
  console.log('FAB clicked, waiting ' + wait + 's...');
  await new Promise((r) => setTimeout(r, wait * 1000));

  const mid = await tb.evaluate(() => ({
    confirm: (() => { const m = document.getElementById('ob-src-confirm'); return m ? { display: getComputedStyle(m).display, text: m.textContent.replace(/\s+/g, ' ').substring(0, 160) } : null; })(),
    toast: (document.getElementById('ob-toast') || {}).textContent || '',
  }));
  console.log('confirm:', JSON.stringify(mid));
  console.log('toast :', mid.toast);

  // 3. 有確認視窗 → 點確定
  if (mid.confirm && mid.confirm.display !== 'none') {
    await tb.evaluate(() => { const ok = document.getElementById('ob-src-cok'); if (ok) ok.click(); });
    console.log('clicked 確定');
    await new Promise((r) => setTimeout(r, 4000));
  }

  const fin = await tb.evaluate(() => {
    let n = 0;
    try { const s = JSON.parse(localStorage.getItem('ob_bridge_v1') || 'null'); n = s && s.records ? s.records.length : 0; } catch (e) { }
    return {
      bridge: (localStorage.getItem('ob_bridge_v1') || '').substring(0, 400),
      recCount: n,
      toast: (document.getElementById('ob-toast') || {}).textContent || '',
    };
  });
  console.log('after bridge:', fin.bridge);
  console.log('after toast :', fin.toast);
  console.log('recCount    :', fin.recCount);
  console.log('===== CONSOLE =====');
  logs.forEach((l) => console.log(' ', l));
  if (!logs.length) console.log('  (empty)');

  const ok = fin.recCount > 0;
  console.log(ok ? '✅ PASS: bridge 有記錄' : '❌ FAIL: bridge 沒有記錄(看上面 toast/console)');
  await browser.close();
  process.exit(ok ? 0 : 1);
})().catch((e) => { console.error('ERR:', e.message); process.exit(1); });
