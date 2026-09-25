# Order Bridge(訂單跨站導入橋)

Tampermonkey userscript:把電商訂單導出檔(xlsx)自動解析並導入集運站的包裹預報表,免去手動逐單輸入品名。

## 目前支援

| 角色 | 站點 | 說明 |
|---|---|---|
| 來源 (Source) | 淘寶 `buyertrade.taobao.com` | 一鍵「導出並暫存」:後台攔截導出的 xlsx,解析後存入跨站暫存 |
| 站點 (Site) | 聖天集運 `member.stjh168.com` | 自動讀取暫存,對照現有包裹:🟢 新預報 / 🟡 更新品名 / 🟠 補品名 |

## 安裝

在 Tampermonkey dashboard 安裝(開 raw 檔會自動彈安裝視窗):

```
https://raw.githubusercontent.com/st9240202/order-bridge/main/order-bridge.user.js
```

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

- **v3.2.x** — 合併式暫存、items 預覽選單、parse 失敗不存入
- v3.1.0 — GM_setValue 跨站儲存(TM 5.5 content world 下 chrome.storage 不注入)
- v3.0.0 — 通用化 OB.Sources / OB.Sites 架構
- v2.x — 單一 script(淘寶+聖天)直連版
