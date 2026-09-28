# 踩坑大全(按主題歸類的血淚史)

本專案 v1.0 → v3.2.25 開發期間踩過的所有坑。每條格式:**現象 → 根因 → 解法**。
新坑照格式補在下面,對應章節。

## 1. Tampermonkey / 世界(world)相關

### 1.1 抓不到用戶腳本的 console 訊息
- **現象**:`page.on('console')` 一直空的,以為腳本沒跑。
- **根因**:監聽掛太晚(reload/bringToFront 之後才 `on`),訊息已流走;CDP 直連的
  `Runtime.consoleAPICalled` 也有同樣問題 — `Runtime.contextCreated` 事件要在
  navigate **之前** attach 才收得到。
- **解法**:先 `page.on('console')` + `page.on('pageerror')`,再動作(見 `dev/tools/console-tail.js`)。
  Playwright 的 console 事件能跨 world 抓到所有 console 輸出,比 CDP 直連省事。

### 1.2 `page.evaluate` 看到/看不到的邊界
- **現象**:`window.__OB` 在 `page.evaluate` 看得到,但腳本 closure 內的變數
  (`xlsxLib`、parse 結果)看不到;main world 的 `window.XLSX`(淘寶自己的)也看不到。
- **根因**:TM 5.5 把 script 注入 `<head>` 的 `<script id="__gmscript__">`(main-world
  相容),所以頂層 `window.*` 屬性兩邊都看得到;但 IIFE closure 變數和 `var XLSX`
  這種 block 內的 binding 不掛 window。
- **解法**:要讀腳本內部狀態 → (a) 讓腳本把狀態掛 `window.__OB`(本腳本已做);
  (b) 讓腳本把狀態畫到 DOM(toast/視窗)再從 DOM 讀;別期待能進 closure。

### 1.3 GM_setValue 寫入延遲一輪
- **現象**:淘寶頁存了暫存,聖天頁(新開)看不到,或看到舊值。
- **根因**:TM 5.5 的 GM 值寫入 extdb(`chrome.storage.local` 的 `!extdb.@*#<uuid>`)
  是批次/延遲 flush,寫後立刻讀可能拿舊值;跨 page 更慘(另一 page 的 `GM_getValue`
  要等 SW sync)。
- **解法**(已內建在腳本):`_memCache` 當同頁真相、`storeSet` 三寫(mem+GM+LS)、
  `freshGet` 跨頁 sync 時讀 GM/LS 取 **max ts** 的那份、viewer `open()` 後 800ms
  再 `freshGet` 一次(等 GM flush)。測試時:跨頁驗證一律 reload 後再斷言。

### 1.4 GM_getValue 回傳 `o{...}` 字串
- **現象**:`JSON.parse(GM_getValue(...))` 偶爾炸。
- **根因**:TM 5.5 extdb 把 object 存成 `o` 前綴字串(`o` = object,`a` = array,
  `s` = string),讀取端要解包。
- **解法**:`asStore()` 先 `v.charAt(0)==='o'` 去前綴再 parse。

### 1.5 chrome.storage 在 content world 不注入
- **現象**:腳本內 `chrome.storage` undefined(Chrome 111+ / TM 5.x)。
- **根因**:isolated world 沒有 `chrome.*`。
- **解法**:跨站狀態一律 GM_setValue;同站備援用 `localStorage`(LS 所有 world 共用)。
  要讀/寫 extdb 本身 → 走 TM dashboard/options 頁(extension world)evaluate。

### 1.6 部署後頁面沒更新
- **根因**:TM 熱更新對**已載入**頁面不重新注入 userscript。
- **解法**:部署後對所有目標 page 做 `page.reload()`。本機測試流程固定:deploy → reload → 驗證。

## 2. 淘寶導出相關

### 2.1 「待發貨」分頁導出沒有物流欄位
- **現象**:按導出後確認視窗不彈(或提示找不到欄位)。
- **根因**:未發貨訂單沒有單號,該分頁的 xlsx 11 欄裡**沒有**物流公司/物流單號。
- **解法**:只有「待收貨」分頁的導出檔有 13 欄(含物流)。v3.2.25 起:
  (a) 表頭匹配改 `indexOf` 子串(新表頭帶「(當前僅支持未完結訂單)」附註,`startsWith` 會失配);
  (b) 錯誤訊息直接寫「請切待收貨分頁導出」;(c) notify 改頁內 toast。

### 2.2 抓不到下載的 xlsx
- **解法**:main world 掛 `URL.createObjectURL` hook,過濾 `size>5000 &&
  type==='application/octet-stream'` 的 blob,`arrayBuffer` → base64 → 存檔
  (見 `dev/tools/capture-blob.js`)。2026-06 實測淘寶每次下載都即時
  `URL.createObjectURL(blob)`,hook 有效。
- **注意**:hook 要在**點擊導出前**裝好;`window.__B64` 是單值,多次下載會互相覆蓋。

### 2.3 導出要連點兩下(導出訂單 → 下載訂單)
- 淘寶新版是兩步:先點「導出訂單」產生任務,幾秒後再點「下載訂單」才出 xlsx。
  腳本/測試都要等第一下的回響(成功 toast)再點第二下。

## 3. 聖天集運 API

### 3.1 端點
- `POST /Member/AddBillcode`(新預報,form 表單)
- `POST /Member/EditPackByID`(更新/補品名)
- **欄位名有官方 typo,不能修**:`Goods_meno`(不是 memo)、`Immediatrly`
- `EditPackByID` 要帶 `goods_id=0`(整數 0 不是空字串,否則回失敗)
- `AddBillcode` **每次最多 10 筆**,超出要自動分組
- 提交完頁面要 reload 才會顯示新包裹(不主動回傳結果)

## 4. DOM / UI 相關

### 4.1 所有 close button 要有 `data-x`
- viewer 用單一 event delegation 綁 click,close 鈕統一 `data-x="..."`;
  漏掉屬性 = 按鈕點了沒反應(v3.2.16 的 bug)。

### 4.2 動態新增的 `<tr>` 要有 `data-r`
- `recItems(ri)` 用 `querySelectorAll('.obv-row[data-r="ri"]')` 讀回編輯值;
  `addRowHtml` 忘了放 `data-r` → 新增列永遠讀不到(v3.2.21 的 bug)。

### 4.3 SPA 會吃掉注入的 DOM
- GitHub 等 SPA 重新 render 會把 `<body>` 自定義節點清掉 → `open()` 時
  `ensureDom()` 檢查節點還在不在,不在就重新 appendChild。

### 4.4 CSP / TrustedHTML
- YouTube 等 CSP 頁面 `innerHTML =` 拋 TrustedHTML 錯誤 → `setHTML()` 先 try
  innerHTML,catch 後走 `elFromHTML()`(純 DOM createElement 解析,不吃 innerHTML)。

### 4.5 mask 擋住底下的 FAB
- 確認視窗的 mask 是 `position:fixed; inset:0`,沒關掉前點 FAB 會被擋。
  測試腳本要在兩步之間先點掉視窗(或 `display:none`)。

## 5. 測試/工具鏈

### 5.1 Playwright
- `page.evaluate(fn, arg)` **只吃一個 arg**,多個打包成 object。
- `connectOverCDP('http://localhost:9333')` → `browser.contexts()[0]` 是 Default profile。
- `NODE_PATH=$HOME/.npm/_npx/*/node_modules`(npx cache 裡的 playwright-core)。
- `browser.close()` 會斷掉 CDP 連線 — 每個測試腳本用完就 close,別 reuse。
- 截圖存 `/Users/joey/.playwright-mcp/`。

### 5.2 Chrome / CDP
- Chrome 136+ 要 `--user-data-dir` 非預設才開得出 `--remote-debugging-port`
  (本機用 APFS clone `~/dev/chrome-real`,啟動參見 DEV.md §1)。
- service worker(TM 的核心)要用 `Target.attachToTarget{flatten}` +
  `Runtime.evaluate` 進 SW world 才能 `chrome.storage.local.get` 讀 extdb。
  Node 26 有內建 `WebSocket`,不用裝 ws。

### 5.3 Tampermonkey
- 本機 TM ext id:`dhdgffkkebhmkfjojejmpbldmpobfkfo`(CWS,5.5.x)。
- dashboard 頁:`chrome-extension://dhdgffkkebhmkfjojejmpbldmpobfkfo/options.html`。
- extdb key 格式:`!extdb.@source#<uuid>` / `!extdb.@meta#<uuid>` /
  `!extdb.@uid#<uuid>`,value = `{origin, value, ts}`。
- 更新腳本 = 覆寫 source + 改 meta.header/version(見 `dev/tools/deploy-tm.js`)。
- TM 5.5 `chrome.storage.local` 有 unlimitedStorage,GM 值不會被容量限制砍。

## 6. 發布

### 6.1 GreasyFork
- **API 只收密碼**,不吃 token/OAuth → GitHub 第三方登入的帳號要到
  「帳號設定」另設一個密碼(可與登入方式並存)。
- API 能上傳**已有腳本**的新版本(`POST /scripts/<id>/versions`),
  **不能建立新腳本** → 第一次一定要手動網頁上傳(貼 raw URL)。
- 上傳頁入口不好找:登入後右上角 avatar 選單(不是 URL 直連;
  `/scripts/upload` 302 到 locale 頁後 404 是正常的)。
- 免警告要件:`@license`(用 MIT)+ 四語 `@description` 都要有內容。
- 發布後記下 script id → 填進 GitHub Actions secrets 就能自動發版。

### 6.2 GitHub
- repo:`st9240202/order-bridge`(public),`@downloadURL` 指 raw URL → 使用者端自動更新。
- 本地 TM 測試用 CDP 直接寫 extdb(`deploy-tm.js`),**不要**依賴 TM 檢查更新
  (快,且不動使用者的「手動修改」狀態)。
