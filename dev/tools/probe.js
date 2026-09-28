#!/usr/bin/env node
// probe.js — 單頁狀態探針:DOM、localStorage、window.__OB、readyState
// 用法: node dev/tools/probe.js <url-substring>
// 例:   node dev/tools/probe.js buyertrade.taobao.com
const { chromium } = require('playwright-core');
const CDP = process.env.CDP || 'http://localhost:9333';
const want = process.argv[2] || 'taobao.com';

(async () => {
  const browser = await chromium.connectOverCDP(CDP);
  const page = browser.contexts()[0].pages().find((p) => p.url().includes(want));
  if (!page) { console.error('找不到 page 含「' + want + '」'); process.exit(1); }
  console.log('PAGE:', page.url());
  const st = await page.evaluate(() => {
    const g = (id) => { const e = document.getElementById(id); return e ? { display: getComputedStyle(e).display, text: (e.textContent || '').trim().substring(0, 80) } : null; };
    return {
      readyState: document.readyState,
      __OB: !!window.__OB,
      __OB_I18N: !!window.__OB_I18N,
      dom: {
        obv: g('obv'), 'ob-src-fab': g('ob-src-fab'), 'ob-src-mgr': g('ob-src-mgr'),
        'ob-fab-dock': g('ob-fab-dock'), 'ob-toast': g('ob-toast'), 'obv-mask': g('obv-mask'),
        'ob-src-confirm': g('ob-src-confirm'), 'ob-src-confirm-mask': g('ob-src-confirm-mask'),
      },
      ls: Object.keys(localStorage).filter((k) => k.startsWith('ob_')).reduce((a, k) => (a[k] = (localStorage.getItem(k) || '').substring(0, 300), a), {}),
      sess: Object.keys(sessionStorage).filter((k) => k.startsWith('__OB')).reduce((a, k) => (a[k] = sessionStorage.getItem(k), a), {}),
    };
  });
  console.log(JSON.stringify(st, null, 2));
  await browser.close();
})().catch((e) => { console.error('ERR:', e.message); process.exit(1); });
