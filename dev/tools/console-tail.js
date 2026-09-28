#!/usr/bin/env node
// console-tail.js — 掛 console/pageerror 監聽 N 秒(抓用戶腳本訊息最可靠的方式)
// 用法: node dev/tools/console-tail.js <url-substring> [seconds=15]
// 注意: 監聽掛上後才開始收 — 要在「觸发动作前」跑,或在腳本內先 bringToFront 再 click
const { chromium } = require('playwright-core');
const CDP = process.env.CDP || 'http://localhost:9333';
const want = process.argv[2] || 'taobao.com';
const secs = Number(process.argv[3] || 15);

(async () => {
  const browser = await chromium.connectOverCDP(CDP);
  const page = browser.contexts()[0].pages().find((p) => p.url().includes(want));
  if (!page) { console.error('找不到 page 含「' + want + '」'); process.exit(1); }
  const logs = [];
  page.on('console', (m) => logs.push(m.type() + ': ' + m.text().substring(0, 300)));
  page.on('pageerror', (e) => logs.push('PAGEERROR: ' + String(e).substring(0, 300)));
  await page.bringToFront();
  console.log('listening ' + secs + 's on ' + page.url());
  await new Promise((r) => setTimeout(r, secs * 1000));
  console.log('===== CONSOLE =====');
  logs.forEach((l) => console.log(' ', l));
  if (!logs.length) console.log('  (empty)');
  await browser.close();
})().catch((e) => { console.error('ERR:', e.message); process.exit(1); });
