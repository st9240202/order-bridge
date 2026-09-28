# Order Bridge(訂單跨站導入橋)

Tampermonkey userscript:把電商訂單導出檔(xlsx)自動解析並導入集運站的包裹預報表,免去手動逐單輸入品名與運單號。

目前支援:**淘寶 → 聖天集運**。架構為通用跨站橋,擴充新來源/新集運站只需加一個適配器定義。

## 功能

- **淘寶一鍵導出**:在「已買到的寶貝」頁點橘色 FAB,自動完成「導出訂單 → 下載 → 解析」全程,零檔案對話框
- **解析後勾選確認**:彈出選擇視窗,勾選要暫存的單號才存入(可全選/個別勾選)
- **跨站暫存**:淘寶存的資料,到聖天集運網頁直接讀取(GM_setValue 跨站通道)
- **聖天自動導入**:自動對照「查看到貨情況」現有包裹,三種情況自動分流:
  - 🟢 新單號 → 自動提交包裹預報(`AddBillcode`,每 10 筆自動分組)
  - 🟡 已到貨且有品名 → 可選更新品名(`EditPackByID`)
  - 🟠 已到貨但無品名 → 自動補品名
- **提交前預覽**:所有操作先列出預覽表,確認后才提交
- **全站暫存管理**:任何網頁 `Ctrl+Shift+B`(或 TM 選單)開啟中央檢視視窗 — 勾選、複製 JSON 分享、貼上 JSON 匯入、匯出 xlsx、手動編輯/新增/刪除
- **四語介面**:繁體中文 / 簡體中文 / 英文 / 日本語,自動偵測瀏覽器語言,可手動切換並記憶

## 安裝

**方式 A — GreasyFork(推薦)**

到 [GreasyFork](https://greasyfork.org) 搜尋 **Order Bridge** 安裝。

**方式 B — GitHub raw**

複製下列網址到瀏覽器,Tampermonkey 會彈出安裝視窗:

```
https://raw.githubusercontent.com/st9240202/order-bridge/main/order-bridge.user.js
```

腳本會自動檢查更新(或到 TM dashboard 手動「檢查更新」)。

## 使用流程

### 第 1 步:淘寶導出並暫存

1. 開啟 [已買到的寶貝](https://buyertrade.taobao.com/trade/itemlist/list_bought_items.htm)
2. **切到「待收貨」分頁**(重要!見下方[常見問題])
3. 點右側橘色「📦 導出」按鈕 — 腳本會自動點「導出訂單」→「下載訂單」
4. 彈出選擇視窗,列出解析出的單號與品名 — 勾選要暫存的項目
5. 按「確定」,顯示「✅ 已暫存(N 個新單號)」

### 第 2 步:聖天集運導入

1. 開啟 [聖天集運 - 查看到貨情況](http://member.stjh168.com/Member/MyPack)
2. 點「⬇ 導入」按鈕
3. 預覽表顯示對照結果(新預報/更新/補品名),勾選要提交的項
4. 按「開始提交」— 自動逐組提交,完成後頁面重新載入顯示結果

### 全站暫存管理

任何網頁按 `Ctrl+Shift+B`(或 Tampermonkey 選單 →「📋 Order Bridge 暫存清單」):

- 查看所有暫存記錄(來源、時間、單號數)
- 勾選單號 → **複製 JSON(勾選)** → 傳給朋友
- **貼上 JSON 匯入** → 合併進自己的暫存區
- 每筆記錄**匯出 xlsx** / 全部匯出 / 全部清除
- 直接**手動編輯**單號、品名、快遞公司,或新增/刪除整筆記錄

### JSON 分享格式

```json
{"ob":"order-bridge/1","items":[{"billcode":"79036183259778","goods":"XXX滑雪板包","company":"中通快递"}]}
```

## 常見問題

**Q: 按了「📦 導出」但沒有跳出選擇視窗 / 提示找不到欄位?**

A: 幾乎都是因為在「**待發貨**」分頁導出 — 未發貨的訂單還沒有物流單號,那個分頁導出的 xlsx 沒有「物流公司/物流單號」欄位。請切到「**待收貨**」分頁再按一次「📦 導出」。

**Q: 為什麼有的單號被標記為「已存在(略過)」?**

A: 該運單號在暫存區已經有記錄了(之前匯入過),不會重複存入。

**Q: 淘寶的「導出訂單」會下載檔案嗎?**

A: 不會 — 腳本在後台攔截下載直接解析,不會存檔也不彈下載視窗。

**Q: 暫存資料存多久?**

A: 7 天自動清理,最多保留 20 筆記錄。導入聖天完成後建議用「導入並刪除」清掉。

**Q: 介面語言不對?**

A: 檢視視窗 / 導入面板右上角有語言下拉,選完會記住,所有網頁生效。

## 技術架構

```
OB.Sources  (輸入端適配器,每網站一個)
   taobao: { match, parseBuffer(buf)→{items,errors}, install(api) }
        │  blob 攔截(createObjectURL)→ xlsx 解析 → 勾選確認 → 合併存入
        ▼
OB.Bridge   (跨站暫存:GM_setValue 主通道 + localStorage 備援)
   { activeId, records: [{ id, name, ts, items[], errors[] }] }
   記憶體快取(同頁真相)+ 三寫 + freshGet 取最新 + generation guard 防髒寫
   同一來源重複匯出 → 合併(新單號加進同一筆,重複單號略過)
   parse 失敗 → 只通知,不存入
        │
        ▼
OB.Sites    (輸出端適配器,每集運站一個)
   stjh: { readExisting(), submitNew(), submitEdit() }
        │
        ▼
OB.UI       (面板 / FAB / 預覽 / 全站檢視視窗 / 四語 i18n)
```

**標準 Item**(bridge 的合約,站點適配器只認這個格式):

```js
{ billcode: '運單號', company: '快遞公司', goods: '品名(≤60字)', rawNames: ['原始商品名', ...] }
```

**擴充新來源(如 eBay/Amazon)**:在 `OB.Sources` 加一個物件
`{ id, name, match(href), parseBuffer(buf)→{items,errors}, install(api) }`
並在 header 加一列 `@match`,站點端不用改。

**擴充新集運站**:在 `OB.Sites` 加 `{ id, name, match, readExisting(), submitNew(), submitEdit() }`。

### 多語言(i18n)

`OB.i18n` 模組管理四語字典(`zh-TW` / `zh-CN` / `en` / `ja`,各 138 keys):

- 自動偵測:優先讀 GM 值 `ob_lang`(手動選過就記憶),否則依 `navigator.language`
- 執行期切換:`window.__OB_I18N.setLang('en'|'zh-TW'|'zh-CN'|'ja')`
- 技術:`OB.utils.setHTML` 對 CSP 嚴格的頁面(如 YouTube 的 TrustedHTML)自動退回純 DOM 解析,確保檢視視窗在**任何網頁**都能開啟

## 開發

開發者/agent 請先讀 [dev/DEV.md](dev/DEV.md)(環境與工具)和 [dev/LESSONS.md](dev/LESSONS.md)(踩坑大全)。

```bash
git clone https://github.com/st9240202/order-bridge.git
cd order-bridge

# 改 script 後:
node --check order-bridge.user.js     # 語法
node dev/check-i18n.js                # 四語 i18n 對齊檢查
node dev/tools/deploy-tm.js           # 部署到本機 TM(需 CDP)
node dev/tools/e2e-taobao.js          # 淘寶端到端測試(詳見 DEV.md §5 開發循環)
git add -A && git commit -m "..." && git push

# 使用者端更新:TM dashboard 按「檢查更新」,或瀏覽器開 raw 連結
```

## License

[MIT](LICENSE) — 自由使用、修改、再發布。
