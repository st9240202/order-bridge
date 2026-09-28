#!/usr/bin/env node
// capture-blob.js — main world 掛 createObjectURL hook,把頁面下載的 blob 存成檔
// 用途: 拿淘寶「導出訂單」產生的 xlsx(腳本會攔截,不落地)
// 用法: node dev/tools/capture-blob.js <url-substring> <out.xlsx> [seconds=15]
// 流程: 先跑此腳本裝 hook 並等待 — 但 hook 裝在「此腳本 attach 之後」的 page 上才有效;
//       因此正確用法是: reload 頁面 → 立即跑本腳本 → 腳本內會自動點 #ob-src-fab
const { chromium } = require('playwright-core');
const fs = require('fs');
const CDP = process.env.CDP || 'http://localhost:9333';
const want = process.argv[2] || 'taobao.com';
const out = process.argv[3] || '/tmp/captured.xlsx';
const secs = Number(process.argv[4] || 15);

(async () => {
  const browser = await chromium.connectOverCDP(CDP);
  const page = browser.contexts()[0].pages().find((p) => p.url().includes(want));
  if (!page) { console.error('找不到 page 含「' + want + '」'); process.exit(1); }
  await page.bringToFront();
  await page.evaluate(() => {
    window.__B64 = null;
    const co = URL.createObjectURL.bind(URL);
    URL.createObjectURL = function (b) {
      const u = co(b);
      try {
        if (b && b.size > 5000 && b.type === 'application/octet-stream') {
          b.arrayBuffer().then((buf) => {
            const bytes = new Uint8Array(buf);
            let bin = '';
            for (let i = 0; i < bytes.length; i += 8192) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
            window.__B64 = btoa(bin);
          });
        }
      } catch (e) { }
      return u;
    };
  });
  const fab = await page.$('#ob-src-fab');
  if (!fab) { console.error('找不到 #ob-src-fab(腳本未載入?)'); process.exit(1); }
  await fab.click();
  console.log('FAB clicked, waiting ' + secs + 's for blob capture...');
  await new Promise((r) => setTimeout(r, secs * 1000));
  const b64 = await page.evaluate(() => window.__B64);
  if (!b64) { console.log('NO BLOB CAPTURED'); await browser.close(); process.exit(1); }
  fs.writeFileSync(out, Buffer.from(b64, 'base64'));
  console.log('saved', out, '(' + Math.round(b64.length * 3 / 4) + ' bytes)');
  await browser.close();
})().catch((e) => { console.error('ERR:', e.message); process.exit(1); });
