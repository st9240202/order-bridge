# 開發指南(Agent / 開發者手冊)

本專案用 Tampermonkey 5.5 + Chrome + CDP/Playwright 做本機測試。
本文件是「接手開發的 agent 該知道的Everything」;踩坑歷史見 [LESSONS.md](LESSONS.md)。

## 1. 環境

| 項目 | 值 |
|---|---|
| Chrome | 136+ 必須用**非預設 user-data-dir** 才能開 remote debugging(本機用 APFS clone `~/dev/chrome-real`) |
| CDP port | `http://localhost:9333` |
| 啟動方式 | `open -na "Google Chrome" --args --user-data-dir=$HOME/dev/chrome-real --remote-debugging-port=9333` |
| 使用者 profile | **Default**(`st9240202`)— 不是 "Profile 1" |
| TM 版本 | 5.5.x(CWS,ext id `dhdgffkkebhmkfjojejmpbldmpobfkfo`) |
| TM 選單 icon | Chrome toolbar 的 Tampermonkey icon(要 pin 出來) |
| 測試腳本 UUID | Order Bridge = `f69d83d0-a8be-4c20-baf7-72ddebe37e0f` |
| Node | 26(內建 WebSocket;`playwright-core` 用 npx cache:`NODE_PATH=$HOME/.npm/_npx/*/node_modules`) |

## 2. 部署新腳本到本機 TM

```bash
node dev/tools/deploy-tm.js [path-to-user.js]   # 預設 ../order-bridge.user.js
```

原理:開 `chrome-extension://dhdgffkkebhmkfjojejmpbldmpobfkfo/options.html`(TM dashboard),
用 `page.evaluate` 寫 `chrome.storage.local`:

- `!extdb.@source#<uuid>` = `{ origin:'tampermonkey', value: <code>, ts }`
- `!extdb.@meta#<uuid>` = 原 meta + 新 `version`/`header`/`lastModified`/`user_modified`
- `!extdb.@uid#<uuid>` = 顯示名稱

**部署後要 reload 受影響的頁面**才生效(TM 熱更新對已載入頁面不保證重注入)。

## 3. 測試工具(`dev/tools/`)

| 腳本 | 用途 |
|---|---|
| `deploy-tm.js` | 把 user.js 寫進本機 TM(見上) |
| `probe.js` | 單頁狀態探針:DOM 元素、localStorage、`window.__OB`、`readyState`、console 近期內容 |
| `console-tail.js` | 掛 `page.on('console'/'pageerror')` 聽 N 秒 — **抓用戶腳本訊息最可靠的方式** |
| `capture-blob.js` | 掛 main-world `createObjectURL` hook,把頁面下載的 blob 存成檔(拿淘寶導出的 xlsx) |
| `e2e-taobao.js` | 完整流程:切待收貨 → 點 FAB → 等確認視窗 → 點確定 → 斷言 localStorage 有資料 |

共通模式:

```js
const { chromium } = require('playwright-core');
const browser = await chromium.connectOverCDP(process.env.CDP || 'http://localhost:9333');
const ctx = browser.contexts()[0];
const page = ctx.pages().find(p => p.url().includes('buyertrade.taobao.com'));
```

## 4. 關鍵陷阱(詳見 LESSONS.md,這裡是速查)

1. **兩個 world**:userscript 跑在 TM 的 world(TM 5.5 在 Chrome 111+ 用 isolated world,但
   透過 `<script id="__gmscript__">` 注入,所以 `window.__OB` / `#ob-src-fab` 等頂層屬性
   **在 main world 的 `page.evaluate` 也能看到**)。真正看不到的是頁面 main world 的
   變數(如淘寶自己的 `window.XLSX`)— 要進 main world 用 `evaluateOnNewDocument` /
   CDP `Runtime.evaluate{worldName}`。反過來說,main world 掛的 hook(如 `grabfile.js` 的
   `createObjectURL` wrapper)能抓到腳本/頁面建立的 blob,但**讀不到腳本 closure 內的狀態**
   (如 `xlsxLib`、parse 結果)→ 腳本內部狀態要在腳本自己 world 的 console 或 DOM 上看。
2. **blob 攔截**:腳本 hook `URL.createObjectURL` 來攔截淘寶的 xlsx 下載。前提是淘寶的
   下載程式碼**沒有提前把 `URL.createObjectURL` 存成局部變數**(2026-06 實測:每次下載
   都即時 lookup,所以 hook 有效;若未來失效,改用 `a[download].click` 的 click 事件
   + `a.href` 的 blob: URL fetch 回來)。
3. **TM 5.5 GM 寫入延遲一輪**:`GM_setValue` 後 `GM_getValue`(另一 page/重新載入)可能
   拿到舊值 → 腳本用 `_memCache`(同頁真相)+ `freshGet`(跨頁 sync 取 max ts)處理;
   測試時「存了但別頁看不到」先等 1~2 秒或 reload 再驗證。
4. **`GM_getValue` 回傳 `o{...}` 字串**:TM 5.5 的 GM 值在 extdb 存成 `o` 前綴字串
   (`o` = object),`asStore()` 要先去前綴再 `JSON.parse`。
5. **`chrome.storage.local` 不直接可用**:TM 5.5 在 Chrome 111+ 把 content world 改成
   isolated world,`chrome.*` 不注入 → 跨站儲存必須走 GM_setValue(主)+ localStorage(備援)。
   注意:**localStorage 是全域的**(同一個 origin 所有 world 共用),所以腳本在 LS 寫的
   `ob_bridge_v1` 在 `page.evaluate` 看得到 — 它常成為唯一的「地面真相」驗證點。
6. **XLSX 用 closure 變數**(`var xlsxLib`),不要依賴 `window.XLSX` — 頁面自己的 XLSX
   在 main world,隔離 world 看不到。
7. **CSP 嚴格的頁面**:YouTube 等用 TrustedHTML,直接 `innerHTML=` 會被擋 → `OB.utils.setHTML`
   (先試 innerHTML,失敗轉 `elFromHTML` 純 DOM 解析)。
8. **SPA 會清掉注入的 DOM**:GitHub 等 SPA 會把 `<body>` 下自定義節點移除 → 每次 `open()`
   前 `ensureDom()` 重新掛回 body。
9. **淘寶表頭會變**:新版表頭帶括號附註(「物流單號(當前僅支持未完結訂單)」),匹配一律用
   `indexOf` 子串,不用 `startsWith`。**只有「待收貨」分頁的導出檔才有物流欄位**。
10. **GM_notify 靠不住**:使用者常看不到桌面通知;所有重要提示都要同時寫**頁內 toast**。
11. **`page.evaluate(fn, arg)` 只吃一個 arg** — 多個參數打包成 object 傳。
12. **CDP 事件監聽要趁早**:`Runtime.contextCreated` 在 reload 後才來,腳本要在
    `bringToFront`/navigate **之前**就 `attach` 好;否則拿不到 world 資訊。
13. **i18n 四語對齊**:任何新 `t('key')` 要在 zh-TW / zh-CN / en / ja **四個 DICT 都加**;
    CI 用 `dev/check-i18n.js`(正規表示式抽 key 集合比對)。
14. **GreasyFork**:API 只收密碼驗證(GitHub 第三方登入要另設密碼);只能上傳版本,
    **不能用 API 建立新腳本**;header 要 `@license` + 四語 `@description` 才不被警告。

## 5. 開發循環(改完腳本後)

```
1. node --check order-bridge.user.js                 # 語法
2. node dev/check-i18n.js                            # 四語 key 對齊
3. node dev/tools/deploy-tm.js                       # 寫進本機 TM
4. 相關頁面 reload(CDP: page.reload())
5. node dev/tools/e2e-taobao.js                      # 端到端
6. git commit + push(GitHub raw 是 @downloadURL,使用者端即可檢查更新)
7. 發 GreasyFork(手動上傳新版本,或打 tag 觸發 workflow)
```

## 6. 常用地點

- 腳本:`/Users/joey/Documents/project/order-bridge/order-bridge.user.js`
  (symlink:`/Users/joey/dev/scripts/order-bridge.user.js`)
- 測試資料:淘寶頁 `localStorage['ob_bridge_v1']`(buyertrade.taobao.com 域)
- 歷史一次性測試腳本:`/Users/joey/tm55-*.js`(**不要引用,已過期**;要復用就複製進 `dev/tools/`)
- 截圖目錄:`/Users/joey/.playwright-mcp/`
