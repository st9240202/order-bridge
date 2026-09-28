# Order Bridge(訂單跨站導入橋)

Tampermonkey userscript:把電商訂單導出檔(xlsx)自動解析並導入集運站的包裹預報表,免去手動逐單輸入品名。

## 目前支援

| 角色 | 站點 | 說明 |
|---|---|---|
| 來源 (Source) | 淘寶 `buyertrade.taobao.com` | 一鍵「導出並暫存」:後台攔截導出的 xlsx,解析後存入跨站暫存 |
| 站點 (Site) | 聖天集運 `member.stjh168.com` | 自動讀取暫存,對照現有包裹:🟢 新預報 / 🟡 更新品名 / 🟠 補品名 |

> ⚠️ **請在「待收貨」分頁導出**：「待發貨」訂單還沒有物流單號，導出的 xlsx 沒有「物流公司/物流單號」欄位，無法導入。腳本會以頁面 toast 提示你切分頁。

## 安裝

**方式 A — GreasyFork(推薦)**:在 [GreasyFork](https://greasyfork.org) 搜尋「Order Bridge」直接安裝,更新由 GreasyFork 分發。

**方式 B — GitHub raw**:在 Tampermonkey dashboard 安裝(開 raw 檔會自動彈安裝視窗):

```
https://raw.githubusercontent.com/st9240202/order-bridge/main/order-bridge.user.js
```

腳本 header 有 `@downloadURL` 指向 GitHub raw,裝好後 TM 會自動檢查更新(或 dashboard 手動「檢查更新」)。

## 開發流程

```bash
# 改 script
nano order-bridge.user.js        # 或任何編輯器

# 推上去
git add -A && git commit -m "..." && git push

# 更新 Tampermonkey:
# 方式 1: 在 TM dashboard 該 script 按「檢查更新」(有 @downloadURL 就會自動拉最新)
# 方式 2: 瀏覽器開 raw 連結 → 彈出更新視窗
```

## 架構

```
OB.Sources  (輸入端適配器,每網站一個)
   taobao: { match, parseBuffer(buf)→{items,errors}, install(api) }
        │  匯出時 parse 成標準 items,合併存入
        ▼
OB.Bridge   (跨站暫存:GM_setValue 主通道 + localStorage 備援)
   { activeId, records: [{ id, name, ts, source, items[], errors[] }] }
   同一 source 合併成一筆;parse 失敗不存入
        │
        ▼
OB.Sites    (輸出端適配器,每集運站一個)
   stjh: { readExisting(), submitNew(), submitEdit() }
        │
        ▼
OB.UI       (面板/FAB/預覽/暫存管理)
```

**標準 Item**(bridge 的合約,站點適配器只認這個格式):

```js
{ billcode: '運單號', company: '快遞公司', goods: '品名(≤60字)', rawNames: ['原始商品名', ...] }
```

**擴充新來源(如 eBay/Amazon)**:在 `OB.Sources` 加一個物件
`{ id, name, match(href), parseBuffer(buf)→{items,errors}, install(api) }`
並在 header 加一列 `@match`。站點端不用改。

## 版本

- **v3.2.25** — 修復淘寶新版導出解析:表頭匹配改子串(相容「物流單號(當前僅支持未完結訂單)」附註);notify 改用頁內 toast(4 秒)不再只靠 TM 桌面通知;提示訊息直接指引「切待收貨分頁導出」
- **v3.2.24** — 修正 EN locale 的中文殘留(vw.* 表頭/按鈕);新增日文 (ja) locale;四語各 137 keys 對齊;header 補 `@description:zh-TW/zh-CN/ja/en` + `@license MIT`(GreasyFork 發布要件)
- **v3.2.23** — 補齊 st.lblCache / st.delAfter / st.delAfterDone 三語翻譯
- **v3.2.22** — 修復檢視器「複製 JSON」按鈕無 handler 的問題
- **v3.2.21** — 檢視器可手動編輯/刪除/新增項目與記錄;移除 source 欄位;「導入並刪除」;儲存層重構(記憶體快取 + GM/LS 雙寫雙讀 + generation guard 防髒寫)
- **v3.2.19** — 統一暫存管理(單一中央視窗)
- v3.1.0 — GM_setValue 跨站儲存(TM 5.5 content world 下 chrome.storage 不注入)
- v3.0.0 — 通用化 OB.Sources / OB.Sites 架構
- v2.x — 單一 script(淘寶+聖天)直連版

## 多語言(i18n)

支援 **繁體中文(預設)/ 簡體中文 / 英文 / 日本語**,透過 `OB.i18n` 模組:

- **自動偵測**:優先讀 GM 值 `ob_lang`(手動選擇過就記住),否則依 `navigator.language` 判斷(zh-CN→簡體,其餘 zh→繁體,ja→日文,其他→英文)
- **切換**:導入面板 / 檢視視窗 / 管理面板右上角有語言下拉,選完立即重渲染並記憶(GM_setValue)
- **全域**:語言是 GM 值,一次設定後所有網頁(聖天/淘寶/任何檢視頁)都用同一語言
- 任何頁面都能:TM 選單「📋 Order Bridge 暫存清單」或 `Ctrl+Shift+B` 開啟檢視視窗,右上下拉切語言
- 手動 API:`window.__OB_I18N.setLang('zh-TW'|'zh-CN'|'en'|'ja')`

> 技術:`OB.utils.setHTML(el, html)` 對 CSP 嚴格的頁面(如 YouTube 的 TrustedHTML)會自動退回純 DOM 解析(`elFromHTML`),確保檢視視窗在**任何網頁**都能開啟。

## JSON 分享(複製/貼上匯入)

檢視視窗(任何網頁 `Ctrl+Shift+B` 或 TM 選單)內:

- 每筆記錄可勾「全選」或勾個別單號 → **複製 JSON(勾選)** → 傳給對方
  格式:`{"ob":"order-bridge/1","items":[{"billcode","goods","company"}]}`
- 接收方開檢視視窗 → **貼上 JSON 匯入** → 貼上文字 → 自動合併進暫存區(可再導入聖天)

## 統一暫存管理(v3.2.19)

原本淘寶/聖天各有一個「暫存資料」側欄、加上全站的檢視視窗 — 三套介面。
現已**統一成一個中央視窗**:

- 淘寶「暫存」FAB、聖天「暫存」FAB、TM 選單、`Ctrl+Shift+B` → **全部開啟同一個視窗**
- 功能:勾選/全選 → 複製 JSON(勾選)・貼上 JSON 匯入・每筆記錄匯出 xlsx・全部匯出・全部清除
- 兩套舊的側欄 admin 面板程式碼已移除

## v3.2.21 — 複製/貼上 UI 升級

- **複製 JSON 後有明確回饋**:視窗頂端浮出一個圓角 toast(「已複製到剪貼簿 · N」),2.2 秒後自動淡出。不再用 Tampermonkey 的角落通知(在任意網頁上也可能看不到)。
- **貼上 JSON 改成自製對話框**:點「貼上 JSON 匯入」會開一個 640px 的中央彈窗 — 標題列 + 等寬字體 textarea(可自由貼上/修改)+「取消 / 匯入」按鈕。解析失敗、無有效項目等錯誤直接顯示在彈窗底部(紅字),不用瀏覽器原生 prompt/alert。
- 技術順帶修兩坑:
  - Tampermonkey 沙箱下 `document.getElementById` 找不到 userscript 建立的節點 → 所有 viewer 節點改用閉包引用;
  - SPA 頁面(如 GitHub)會把注入節點清掉 → `open()` 時自動重新掛回 body。

## v3.2.22–v3.2.25 更新記錄

### v3.2.25 — 淘寶新版導出解析修復(重要)
- **根因**:淘寶「待發貨」分頁導出的 xlsx **沒有物流公司/物流單號欄位**(未發貨訂單無單號),加上新版表頭帶括號附註(如「物流單號(當前僅支持未完結訂單)」),舊的 `startsWith` 匹配失敗
- 表頭匹配改**子串**(`indexOf`),附註表頭也能匹配
- `notify()` 優先寫**頁內 toast**(4 秒顯示),TM 桌面通知降為備援 — 任何提示都不會再「消失」
- 解析失敗訊息直接寫出解法:「請切待收貨分頁導出」
- 實測:待收貨分頁 FAB → 下載 → 確認視窗(10 筆)→ 確定 → 成功存入

### v3.2.24 — i18n 大修 + 日文
- 修復 EN locale 16 個 `vw.*` 鍵被中文值覆蓋的 bug(表頭/按鈕曾顯示中文)
- 新增完整**日文 (ja)** locale;`detect()` 支援 `navigator.language=ja`;下拉選單加「日本語」
- 四語各 137 keys 完全對齊;header 補 `@description` 四語 + `@license MIT` + repo 新增 `LICENSE`

### v3.2.23 — 翻譯補齊
- 補上 `st.lblCache` / `st.delAfter` / `st.delAfterDone` 三語翻譯

### v3.2.22 — 複製 JSON 修復
- 修復檢視器「複製 JSON(勾選)」按鈕無 handler 的問題;複製成功/失敗有 toast 回饋

### v3.2.21 — 檢視器全面升級
- 檢視視窗可**手動編輯**(單號/品名/快遞)、**刪除列/記錄**、**新增空記錄**
- 「導入並刪除」:導入聖天後自動清除 bridge
- 儲存層重構:記憶體快取(同頁真相)+ GM/LS 三寫 + `freshGet` 取最新(處理 TM 5.5 GM 寫入延遲)+ generation guard 防重疊重寫
