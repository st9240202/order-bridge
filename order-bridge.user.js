// ==UserScript==
// @name         訂單跨站導入橋 (Order Bridge)
// @name:zh-TW   訂單跨站導入橋 (Order Bridge)
// @name:zh-CN   订单跨站导入桥 (Order Bridge)
// @name:en      Order Bridge
// @namespace    https://tampermonkey.net/
// @version      3.2.17
// @match        https://buyertrade.taobao.com/trade/itemlist/*
// @match        http://member.stjh168.com/Member/MyPack
// @match        *://*/*
// @description  通用訂單 xlsx 跨站橋:輸入端 OB.Sources(暫存/管理)+ 輸出端 OB.Sites(適配器)。現含:淘寶 → 聖天集運。擴充新站點只需加一個 Source/Site 定義。
// @homepage     https://github.com/st9240202/order-bridge
// @source       https://raw.githubusercontent.com/st9240202/order-bridge/main/order-bridge.user.js
// @downloadURL  https://raw.githubusercontent.com/st9240202/order-bridge/main/order-bridge.user.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_registerMenuCommand
// @grant        GM_registerMenuCommandClose
// @grant        GM_notify
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // =====================================================================
  // OB 命名空間(window.__OB 可除錯/擴充):
  //   OB.utils   共用工具(DOM/編碼/通知/SheetJS 載入)
  //   OB.Bridge  跨站暫存(多筆記錄 + 來源標記 + 舊版遷移 + 自動清理)
  //   OB.Parser  xlsx → 標準 Item[](輸入端方言解析;每個 Source 可自訂)
  //   OB.Sources 輸入端適配器(每個來源網站一個;提供 UI/觸發方式)
  //   OB.Sites   輸出端適配器(每個集運站一個;定義讀現有/提交新/提交改)
  //   OB.UI      輸出端面板/FAB/預覽/暫存管理選單(只認 site 介面)
  //
  // 標準 Item: { billcode, company, goods, rawNames[] }
  // 暫存記錄:  { id, name, ts, source, items[], errors[] }  (bridge 100% 只存 parse 後的 items;舊版記錄可能帶 b64,載入時由對應 Source 現 parse)
  //
  // 擴充新來源(如 eBay/Amazon):在 OB.Sources 加一個物件
  //   { id, name, match(href), parseBuffer(buf) → {items, errors}, install(api) }
  //   並加一列 @match。站點適配器只認標準 Item,不需改。
  // =====================================================================
  var OB = window.__OB = window.__OB || {};
  var xlsxLib = null; // SheetJS 快取(避免 window 污染 TM sandbox)

  // ===================== OB.i18n(多語言) =====================
  // 預設語言:TM 設定值 'ob_lang' ∈ zh-TW/zh-CN/en;否則 navigator.language(zh-CN→簡體,其餘→繁中)
  // 執行期切換:window.__OB_I18N.setLang('en'|'zh-CN'|'zh-TW');面板內也有語言下拉
  OB.i18n = (function () {
    var DICT = {
      'zh-TW': {
        'lang': '繁體中文',
        'notify.title': '訂單導入橋',
        'notify.sheetjsFail': '❌ 無法載入 SheetJS(CDN 全失敗):{m}',
        'utils.loadFailed': '載入後找不到 XLSX',
        'prev.noData': '(無資料)',
        'prev.more': '… 其餘 {n} 個單號',
        'parse.noData': '檔案沒有資料',
        'parse.noCols': '找不到「商品名稱/物流單號」欄位,請確認檔案格式',
        'parse.rowBad': '第{n}行單號異常:{b}',
        'parse.sheetjs': 'SheetJS 未載入',
        'src.parseFail': '⚠ 解析失敗({f}:{m}),未存入',
        'src.noItems': '⚠ 未解析到任何單號({f},{e}),未存入',
        'src.cancel': '⚪ 已取消,未存入',
        'src.saved': '✅ 已暫存({a} 個新單號{d},來源:{s})',
        'src.savedDup': ',{d} 個已存在(略過)',
        'src.saveFail': '❌ 暫存失敗:{m}',
        'src.saveFailFile': '❌ 暫存失敗({f}:{m})',
        'src.confirmTitle': '📦 選擇要暫存的單號',
        'src.selectAll': '全選',
        'src.cancelBtn': '取消',
        'src.okBtn': '確認暫存',
        'src.empty': '尚無資料。觸發本站點導出或手動下載訂單後會自動出現。',
        'src.adminTitle': '暫存資料',
        'src.source': '來源:{s}',
        'src.nItems': '{n} 個單號',
        'src.dlBtn': '匯出 xlsx(備份)',
        'src.clearBtn': '全部清除',
        'src.clearAsk': '確定清除所有暫存資料?',
        'lblExport': '導出',
        'lblCache': '暫存',
        'tb.busy': '⏳',
        'tb.exportBtnNotFound': '找不到「導出訂單」按鈕',
        'tb.waitTimeout': '等待「下載訂單」超時',
        'tb.captureFail': '未捕獲到下載檔案,請手動下載或改在管理選單匯入',
        'ui.importTo': '導入訂單 → {s} 預報',
        'ui.auto': '自動(暫存+現有)',
        'ui.pickFile': '選擇檔案(備用)',
        'ui.reload': '重新讀取暫存',
        'ui.idle': '尚未讀取資料',
        'ui.editEmpty': '已無品名 → 補品名',
        'ui.editFill': '已有品名 → 覆寫更新',
        'ui.autoRefresh': '提交完成後自動重新整理頁面',
        'ui.colBill': '單號',
        'ui.colCompany': '物流公司',
        'ui.colGoods': '品名(goods,≤60字)',
        'ui.colAction': '處理方式',
        'ui.colResult': '結果',
        'ui.submit': '開始提交',
        'ui.reanalyze': '重新分析',
        'ui.bNew': '新預報',
        'ui.bAddName': '補品名',
        'ui.bUpdateName': '更新品名',
        'ui.bNeedCheck': '需勾選',
        'ui.existing': '現有:{g}',
        'ui.noParsed': '⚠ 請先讀取檔案(點「自動(暫存+現有)」或選擇檔案)',
        'ui.plan': '📋 共 {n} 個單號 → <b>{a}</b> 新預報 / <b>{b}</b> 需補品名 / <b>{c}</b> 可更新品名。勾選要處理的列(預設全選),再按「開始提交」。',
        'ui.noneChecked': '⚠ 沒有勾選任何要處理的項',
        'ui.submitting': '⏳ 提交中…',
        'ui.okNew': '✓ 已預報',
        'ui.okEdit': '✓ 已更新',
        'ui.fail': '✗ {m}',
        'ui.done': '⚠ 完成:成功 {a} / 失敗 {b}',
        'ui.doneAll': '✅ 全部成功({n} 筆)。可到「查看到貨情況」核對。',
        'ui.willRefresh': ' <span style="color:#999">(3 秒後自動重新整理…)</span>',
        'ui.notifyDone': '導入完成:成功 {a},失敗 {b}{r}',
        'ui.refreshNote': ',3 秒後重新整理',
        'ui.noBridge': '⚠ 暫存區沒有檔案,請先在來源站點導出',
        'ui.err': '❌ {m}',
        'ui.autoFail': '❌ 自動載入失敗:{m}',
        'ui.readingOld': '⏳ 解析舊版暫存(原始檔)…',
        'ui.oldParseFail': '舊版暫存解析失敗:{m}',
        'ui.noDataInRec': '暫存記錄無資料',
        'ui.readOk': '✅ 已讀取「{f}」({s},解析出 <b>{n}</b> 個單號{e})',
        'ui.readErrs': '<br><span class="ob-err">⚠ {e}</span>',
        'ui.fileParsing': '⏳ 解析 {f}…',
        'ui.fileOk': '✅ {f} → <b>{n}</b> 個單號{e}',
        'ui.fileParseFail': '❌ 解析失敗:{m}',
        'ui.srcManual': '手動匯入',
        'st.empty': '尚無暫存資料。請先在來源站點導出訂單。',
        'st.titleImport': '導入訂單 → {s}',
        'st.titleCache': '暫存管理',
        'st.lblImport': '導入',
        'vw.menuTitle': '📋 檢視 Order Bridge 暫存清單',
        'vw.title': '📦 Order Bridge 暫存記錄',
        'vw.exportAll': '全部匯出 xlsx',
        'vw.clear': '全部清除',
        'vw.empty': '暫存區無資料(在淘寶等來源站點導出訂單後會自動出現)',
        'vw.active': '· 導入時使用',
        'vw.nItems': ' · {n} 個單號',
        'vw.clearAsk': '確定清除所有暫存資料?',
        'vw.loading': '載入中…',
        'vw.titleFull': '📋 Order Bridge 暫存清單',
        'vw.emptyNote': '目前沒有暫存資料。請先在淘寶 buyertrade 訂單頁用本工具導出。',
        'vw.activeTag': '<span style="color:#1890ff">(目前導入用)</span>',
        'vw.noItems': '(無項目)',
        'vw.readFail': '讀取失敗:{m}',
        'ui.langLabel': '語言',
        'site.fail': '失敗',
        'site.badResp': '回應異常',
        'site.netErr': '網路錯誤:{m}',
        'api.menuTitle': '📋 檢視 Order Bridge 暫存清單'
      },
      'zh-CN': {
        'lang': '简体中文',
        'notify.title': '订单导入桥',
        'notify.sheetjsFail': '❌ 无法加载 SheetJS(CDN 全失败):{m}',
        'utils.loadFailed': '加载后找不到 XLSX',
        'prev.noData': '(无数据)',
        'prev.more': '… 其余 {n} 个单号',
        'parse.noData': '文件没有数据',
        'parse.noCols': '找不到“商品名称/物流单号”字段,请确认文件格式',
        'parse.rowBad': '第{n}行单号异常:{b}',
        'parse.sheetjs': 'SheetJS 未加载',
        'src.parseFail': '⚠ 解析失败({f}:{m}),未存入',
        'src.noItems': '⚠ 未解析到任何单号({f},{e}),未存入',
        'src.cancel': '⚪ 已取消,未存入',
        'src.saved': '✅ 已暂存({a} 个新单号{d},来源:{s})',
        'src.savedDup': ',{d} 个已存在(略过)',
        'src.saveFail': '❌ 暂存失败:{m}',
        'src.saveFailFile': '❌ 暂存失败({f}:{m})',
        'src.confirmTitle': '📦 选择要暂存的单号',
        'src.selectAll': '全选',
        'src.cancelBtn': '取消',
        'src.okBtn': '确认暂存',
        'src.empty': '暂无数据。触发本站点导出或手动下载订单后会自动出现。',
        'src.adminTitle': '暂存数据',
        'src.source': '来源:{s}',
        'src.nItems': '{n} 个单号',
        'src.dlBtn': '导出 xlsx(备份)',
        'src.clearBtn': '全部清除',
        'src.clearAsk': '确定清除所有暂存数据?',
        'lblExport': '导出',
        'lblCache': '暂存',
        'tb.busy': '⏳',
        'tb.exportBtnNotFound': '找不到“导出订单”按钮',
        'tb.waitTimeout': '等待“下载订单”超时',
        'tb.captureFail': '未捕获到下载文件,请手动下载或改在管理菜单导入',
        'ui.importTo': '导入订单 → {s} 预报',
        'ui.auto': '自动(暂存+现有)',
        'ui.pickFile': '选择文件(备用)',
        'ui.reload': '重新读取暂存',
        'ui.idle': '尚未读取数据',
        'ui.editEmpty': '已无品名 → 补品名',
        'ui.editFill': '已有品名 → 覆写更新',
        'ui.autoRefresh': '提交完成后自动刷新页面',
        'ui.colBill': '单号',
        'ui.colCompany': '物流公司',
        'ui.colGoods': '品名(goods,≤60字)',
        'ui.colAction': '处理方式',
        'ui.colResult': '结果',
        'ui.submit': '开始提交',
        'ui.reanalyze': '重新分析',
        'ui.bNew': '新预报',
        'ui.bAddName': '补品名',
        'ui.bUpdateName': '更新品名',
        'ui.bNeedCheck': '需勾选',
        'ui.existing': '现有:{g}',
        'ui.noParsed': '⚠ 请先读取文件(点“自动(暂存+现有)”或选择文件)',
        'ui.plan': '📋 共 {n} 个单号 → <b>{a}</b> 新预报 / <b>{b}</b> 需补品名 / <b>{c}</b> 可更新品名。勾选要处理的行(默认全选),再按“开始提交”。',
        'ui.noneChecked': '⚠ 没有勾选任何要处理的项',
        'ui.submitting': '⏳ 提交中…',
        'ui.okNew': '✓ 已预报',
        'ui.okEdit': '✓ 已更新',
        'ui.fail': '✗ {m}',
        'ui.done': '⚠ 完成:成功 {a} / 失败 {b}',
        'ui.doneAll': '✅ 全部成功({n} 笔)。可到“查看到货情况”核对。',
        'ui.willRefresh': ' <span style="color:#999">(3 秒后自动刷新…)</span>',
        'ui.notifyDone': '导入完成:成功 {a},失败 {b}{r}',
        'ui.refreshNote': ',3 秒后刷新',
        'ui.noBridge': '⚠ 暂存区没有文件,请先在来源站点导出',
        'ui.err': '❌ {m}',
        'ui.autoFail': '❌ 自动加载失败:{m}',
        'ui.readingOld': '⏳ 解析旧版暂存(原始文件)…',
        'ui.oldParseFail': '旧版暂存解析失败:{m}',
        'ui.noDataInRec': '暂存记录无数据',
        'ui.readOk': '✅ 已读取“{f}”({s},解析出 <b>{n}</b> 个单号{e})',
        'ui.readErrs': '<br><span class="ob-err">⚠ {e}</span>',
        'ui.fileParsing': '⏳ 解析 {f}…',
        'ui.fileOk': '✅ {f} → <b>{n}</b> 个单号{e}',
        'ui.fileParseFail': '❌ 解析失败:{m}',
        'ui.srcManual': '手动导入',
        'st.empty': '暂无暂存数据。请先在来源站点导出订单。',
        'st.titleImport': '导入订单 → {s}',
        'st.titleCache': '暂存管理',
        'st.lblImport': '导入',
        'vw.menuTitle': '📋 查看 Order Bridge 暂存清单',
        'vw.title': '📦 Order Bridge 暂存记录',
        'vw.exportAll': '全部导出 xlsx',
        'vw.clear': '全部清除',
        'vw.empty': '暂存区无数据(在淘宝等来源站点导出订单后会自动出现)',
        'vw.active': '· 导入时使用',
        'vw.nItems': ' · {n} 个单号',
        'vw.clearAsk': '确定清除所有暂存数据?',
        'vw.loading': '加载中…',
        'vw.titleFull': '📋 Order Bridge 暂存清单',
        'vw.emptyNote': '目前没有暂存数据。请先在淘宝 buyertrade 订单页用本工具导出。',
        'vw.activeTag': '<span style="color:#1890ff">(目前导入用)</span>',
        'vw.noItems': '(无项目)',
        'vw.readFail': '读取失败:{m}',
        'ui.langLabel': '语言',
        'site.fail': '失败',
        'site.badResp': '响应异常',
        'site.netErr': '网络错误:{m}',
        'api.menuTitle': '📋 查看 Order Bridge 暂存清单'
      },
      en: {
        'lang': 'English',
        'notify.title': 'Order Bridge',
        'notify.sheetjsFail': '❌ SheetJS failed to load (all CDNs): {m}',
        'utils.loadFailed': 'XLSX not found after load',
        'prev.noData': '(no data)',
        'prev.more': '… {n} more tracking no.',
        'parse.noData': 'File has no data',
        'parse.noCols': '“product name / tracking no.” columns not found; check file format',
        'parse.rowBad': 'Row {n}: invalid tracking number {b}',
        'parse.sheetjs': 'SheetJS not loaded',
        'src.parseFail': '⚠ Parse failed ({f}: {m}) — not stored',
        'src.noItems': '⚠ No tracking numbers parsed ({f}, {e}) — not stored',
        'src.cancel': '⚪ Cancelled — not stored',
        'src.saved': '✅ Staged ({a} new tracking no.{d}, source: {s})',
        'src.savedDup': ', {d} already existed (skipped)',
        'src.saveFail': '❌ Stage failed: {m}',
        'src.saveFailFile': '❌ Stage failed ({f}: {m})',
        'src.confirmTitle': '📦 Select tracking numbers to stage',
        'src.selectAll': 'All',
        'src.cancelBtn': 'Cancel',
        'src.okBtn': 'Confirm',
        'src.empty': 'No data yet. Trigger this site’s export or download an order file and it will appear here.',
        'src.adminTitle': 'Staged data',
        'src.source': 'source: {s}',
        'src.nItems': '{n} tracking no.',
        'src.dlBtn': 'Export xlsx (backup)',
        'src.clearBtn': 'Clear all',
        'src.clearAsk': 'Clear all staged data?',
        'lblExport': 'Export',
        'lblCache': 'Stage',
        'tb.busy': '⏳',
        'tb.exportBtnNotFound': '“Export orders” button not found',
        'tb.waitTimeout': 'Timeout waiting for “Download orders”',
        'tb.captureFail': 'Download not captured — download manually or import via the admin panel',
        'ui.importTo': 'Import orders → {s} pre-declaration',
        'ui.auto': 'Auto (staged + existing)',
        'ui.pickFile': 'Pick file (fallback)',
        'ui.reload': 'Re-read staged data',
        'ui.idle': 'No data loaded yet',
        'ui.editEmpty': 'Arrived w/o name → add name',
        'ui.editFill': 'Arrived w/ name → overwrite',
        'ui.autoRefresh': 'Auto-refresh page after submit',
        'ui.colBill': 'Tracking no.',
        'ui.colCompany': 'Carrier',
        'ui.colGoods': 'Product name (goods, ≤60 chars)',
        'ui.colAction': 'Action',
        'ui.colResult': 'Result',
        'ui.submit': 'Start submit',
        'ui.reanalyze': 'Re-analyze',
        'ui.bNew': 'New',
        'ui.bAddName': 'Add name',
        'ui.bUpdateName': 'Update name',
        'ui.bNeedCheck': 'Check me',
        'ui.existing': 'Current: {g}',
        'ui.noParsed': '⚠ Load a file first (click “Auto” or pick a file)',
        'ui.plan': '📋 {n} tracking no. → <b>{a}</b> new / <b>{b}</b> need name / <b>{c}</b> can update. Check rows to process (all by default), then “Start submit”.',
        'ui.noneChecked': '⚠ Nothing checked to process',
        'ui.submitting': '⏳ Submitting…',
        'ui.okNew': '✓ Declared',
        'ui.okEdit': '✓ Updated',
        'ui.fail': '✗ {m}',
        'ui.done': '⚠ Done: {a} ok / {b} failed',
        'ui.doneAll': '✅ All {n} succeeded. Check “Arrival status” to verify.',
        'ui.willRefresh': ' <span style="color:#999">(auto-refresh in 3s…)</span>',
        'ui.notifyDone': 'Import done: {a} ok, {b} failed{r}',
        'ui.refreshNote': ', refreshing in 3s',
        'ui.noBridge': '⚠ Bridge is empty — export orders on the source site first',
        'ui.err': '❌ {m}',
        'ui.autoFail': '❌ Auto-load failed: {m}',
        'ui.readingOld': '⏳ Parsing legacy staged file…',
        'ui.oldParseFail': 'Legacy staged file failed to parse: {m}',
        'ui.noDataInRec': 'Staged record has no data',
        'ui.readOk': '✅ Loaded “{f}” ({s}, <b>{n}</b> tracking no.{e})',
        'ui.readErrs': '<br><span class="ob-err">⚠ {e}</span>',
        'ui.fileParsing': '⏳ Parsing {f}…',
        'ui.fileOk': '✅ {f} → <b>{n}</b> tracking no.{e}',
        'ui.fileParseFail': '❌ Parse failed: {m}',
        'ui.srcManual': 'manual import',
        'st.empty': 'Nothing staged. Export orders on the source site first.',
        'st.titleImport': 'Import orders → {s}',
        'st.titleCache': 'Staging manager',
        'st.lblImport': 'Import',
        'vw.menuTitle': '📋 View Order Bridge staged records',
        'vw.title': '📦 Order Bridge staged records',
        'vw.exportAll': 'Export all (xlsx)',
        'vw.clear': 'Clear all',
        'vw.empty': 'Bridge is empty (export orders on a source site and they will appear here)',
        'vw.active': '· used for import',
        'vw.nItems': ' · {n} tracking no.',
        'vw.clearAsk': 'Clear all staged data?',
        'vw.loading': 'Loading…',
        'vw.titleFull': '📋 Order Bridge staged list',
        'vw.emptyNote': 'Nothing staged yet. Export orders on Taobao buyertrade using this tool first.',
        'vw.activeTag': '<span style="color:#1890ff">(in use for import)</span>',
        'vw.noItems': '(no items)',
        'vw.readFail': 'Read failed: {m}',
        'ui.langLabel': 'Language',
        'site.fail': 'Failed',
        'site.badResp': 'Bad response',
        'site.netErr': 'Network error: {m}',
        'api.menuTitle': '📋 View Order Bridge staged records'
      }
    };
    function detect() {
      try {
        var v = (typeof GM_getValue === 'function') ? GM_getValue('ob_lang') : null;
        if (v === 'zh-TW' || v === 'zh-CN' || v === 'en') return v;
      } catch (e) { }
      var loc = (typeof navigator !== 'undefined' && navigator.language) || '';
      return /^zh[-_]cn/i.test(loc) ? 'zh-CN' : 'zh-TW';
    }
    var lang = detect();
    function setLang(l) {
      if (!DICT[l]) return;
      lang = l;
      try { if (typeof GM_setValue === 'function') GM_setValue('ob_lang', l); } catch (e) { }
      try { document.dispatchEvent(new CustomEvent('ob-lang-change', { detail: { lang: l } })); } catch (e) { }
    }
    // t('key', {k: v}) — 命名佔位 {k};無參亦可 t('key')
    function t(key, args) {
      var s = (DICT[lang] && DICT[lang][key]) || DICT['zh-TW'][key] || key;
      if (args) for (var k in args) s = s.replace(new RegExp('\\{' + k + '\\}', 'g'), args[k]);
      return s;
    }
    // 語言切換下拉(HTML);變更時 setLang → 各 UI 自動重渲染
    function langHtml() {
      return '<span id="ob-lang-wrap" title="語言 / Language" style="display:inline-block;margin-left:10px;vertical-align:middle">' +
        '<select id="ob-lang-sel" style="padding:2px 6px;border:1px solid #d9d9d9;border-radius:4px;font-size:12px;background:#fff;cursor:pointer">' +
        ['zh-TW', 'zh-CN', 'en'].map(function (l) {
          return '<option value="' + l + '"' + (l === lang ? ' selected' : '') + '>' + DICT[l]['lang'] + '</option>';
        }).join('') + '</select></span>';
    }
    function bindLangSel(root) {
      var sel = (root || document).querySelector && (root || document).querySelector('#ob-lang-sel');
      if (!sel || sel.__obBound) return;
      sel.__obBound = true;
      sel.addEventListener('click', function (e) { e.stopPropagation(); }); // 免點到遮罩關閉視窗
      sel.addEventListener('change', function () { setLang(this.value); });
    }
    return { lang: lang, setLang: setLang, t: t, DICT: DICT, langHtml: langHtml, bindLangSel: bindLangSel };
  })();
  var t = OB.i18n.t;
  window.__OB_I18N = OB.i18n;

  // ===================== OB.utils =====================
  function $(sel, root) { return (root || document).querySelector(sel); }
  function esc(s) {
    var d = document.createElement('div');
    d.textContent = s == null ? '' : String(s);
    return d.innerHTML;
  }
  function billcodeValid(s) { return typeof s === 'string' && /^[0-9A-Za-z]{8,20}$/.test(s); }
  function notify(text) {
    try { if (typeof GM_notify === 'function') { GM_notify({ title: t('notify.title'), text: text, priority: 2 }); return; } } catch (e) { }
    console.log('[OrderBridge]' + text);
  }
  function bufToB64(buf) {
    var bytes = new Uint8Array(buf), bin = '', CH = 0x8000;
    for (var i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(bin);
  }
  function b64ToBuf(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out.buffer;
  }
  // 暫存記錄匯出為 xlsx:優先用原始 b64;否則由 items 重建工作簿
  function downloadRecord(rec) {
    var finish = function (buf) {
      var blob = new Blob([buf], { type: 'application/octet-stream' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = rec.name || 'orders.xlsx';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
    };
    if (rec.b64) { finish(b64ToBuf(rec.b64)); return; }
    if (rec.items && rec.items.length) {
      loadSheetJs().then(function (XLSX) {
        var rows = rec.items.map(function (it) {
          return { '商品名称': it.rawNames ? it.rawNames.join('、') : it.goods, '物流公司': it.company, '物流单号': it.billcode };
        });
        var ws = XLSX.utils.json_to_sheet(rows);
        var out = XLSX.write(ws, { bookType: 'xlsx', type: 'array' });
        finish(out.buffer || out);
      });
    }
  }
  function loadSheetJs() {
    if (xlsxLib) return Promise.resolve(xlsxLib);
    if (window.XLSX && typeof window.XLSX.read === 'function') { xlsxLib = window.XLSX; return Promise.resolve(xlsxLib); }
    var urls = [
      'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
    ];
    var i = 0;
    function tryNext() {
      return fetch(urls[i++], { mode: 'cors' }).then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        return r.text();
      }).then(function (code) {
        var f = new Function(code + '\nreturn window.XLSX;');
        var X = f();
        if (!X || typeof X.read !== 'function') throw new Error(t('utils.loadFailed'));
        xlsxLib = X;
        return X;
      }).catch(function (e) {
        if (i >= urls.length) throw new Error(t('notify.sheetjsFail', { m: e.message }));
        return tryNext();
      });
    }
    return tryNext();
  }
  function fmtTs(ts) {
    var d = new Date(ts);
    return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
  }
  function itemsPreviewHtml(items) {
    items = items || [];
    if (!items.length) return '<div class="more">' + t('prev.noData') + '</div>';
    var MAX = 50;
    var shown = items.slice(0, MAX);
    var html = '<table>' + shown.map(function (it) {
      return '<tr><td>' + esc(it.billcode) + '</td><td>' + esc(it.goods || '') + (it.company ? ' <span style="color:#999">(' + esc(it.company) + ')</span>' : '') + '</td></tr>';
    }).join('') + '</table>' + (items.length > MAX ? '<div class="more">' + t('prev.more', { n: items.length - MAX }) + '</div>' : '');
    return html;
  }
  OB.utils = { $: $, esc: esc, billcodeValid: billcodeValid, notify: notify, bufToB64: bufToB64, b64ToBuf: b64ToBuf, downloadRecord: downloadRecord, loadSheetJs: loadSheetJs, fmtTs: fmtTs, itemsPreviewHtml: itemsPreviewHtml };

  // ===================== 方言解析(xlsx → Item[]) =====================
  // 每個 Source 可帶自己的 parseBuffer(buf);預設/淘寶方言如下(簡/繁表頭皆可)
  // 保留為 OB.Parser 作為各 Source 的共用工具
  var GOODS_NAME_MAX = 60;
  function parseWorkbook(wb) {
    var XLSX = xlsxLib;
    if (!XLSX) throw new Error(t('parse.sheetjs'));
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows.length) return { items: [], errors: [t('parse.noData')] };
    var header = (rows[0] || []).map(function (h) { return String(h).trim(); });
    function colIdx(prefixes) {
      for (var i = 0; i < header.length; i++)
        for (var p = 0; p < prefixes.length; p++)
          if (header[i].indexOf(prefixes[p]) === 0) return i;
      return -1;
    }
    var cName = colIdx(['商品名称', '商品名稱']);
    var cCompany = colIdx(['物流公司']);
    var cBill = colIdx(['物流单号', '物流單號']);
    if (cName < 0 || cBill < 0) return { items: [], errors: [t('parse.noCols')] };

    var groups = {}, errors = [];
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var bill = String(row[cBill] == null ? '' : row[cBill]).trim();
      var name = String(row[cName] == null ? '' : row[cName]).trim();
      if (!bill) continue;
      if (!billcodeValid(bill)) { errors.push(t('parse.rowBad', { n: r + 1, b: bill })); continue; }
      if (!name) continue;
      var company = cCompany >= 0 ? String(row[cCompany] || '').trim() : '';
      if (!groups[bill]) groups[bill] = { billcode: bill, company: company, names: [] };
      var g = groups[bill];
      if (company && !g.company) g.company = company;
      if (g.names.indexOf(name) === -1) g.names.push(name);
    }
    var items = Object.keys(groups).map(function (k) {
      var g = groups[k];
      return { billcode: g.billcode, company: g.company || '', goods: g.names.join('、').slice(0, GOODS_NAME_MAX), rawNames: g.names };
    });
    return { items: items, errors: errors };
  }
  OB.Parser = { parseWorkbook: parseWorkbook, GOODS_NAME_MAX: GOODS_NAME_MAX };

  // ===================== OB.Bridge(跨站暫存) =====================
  // store = { activeId, records: [ {id,name,size,ts,source,items[],errors[]};舊版記錄可能帶 b64 ] }(新→舊)
  // 優先級:GM_setValue(跨站主通道) → localStorage(同站備援)
  // 遷移:ob_bridge_v1 ← stjh_bridge_v2 ← stjh_tb_xlsx(舊格式含 b64,載入時轉 items)
  var STORE_KEY = 'ob_bridge_v1';
  var LEGACY_KEYS = ['stjh_bridge_v2', 'stjh_tb_xlsx'];
  var MAX_AGE = 7 * 864e5;
  var MAX_COUNT = 20;

  function migrateLegacy(leg) {
    // 舊格式可能是單筆記錄(帶 b64)或 store({records:[...]})
    if (!leg) return null;
    var rawRecs = leg.records && leg.records.length ? leg.records : [leg];
    var recs = [];
    rawRecs.forEach(function (r) {
      if (!r) return;
      if (!r.b64 && !(r.items && r.items.length)) return; // 無法使用的記錄
      var rec = { id: String(r.ts || Date.now()), name: r.name || 'orders.xlsx', size: r.size || 0, ts: r.ts || Date.now(), source: r.source || r.items ? r.source || 'taobao' : 'taobao' };
      if (r.b64) rec.b64 = r.b64;
      else { rec.items = r.items; rec.errors = r.errors; }
      recs.push(rec);
    });
    if (!recs.length) return null;
    return { activeId: recs[0].id, records: recs };
  }
  function storeGet() {
    // 優先級: GM(跨站主通道) → localStorage(同站備援)
    // TM 5.5 下 @grant chrome.storage.local 在 content world 不一定注入,故不使用
    return new Promise(function (resolve) {
      function fromGm() {
        try {
          if (typeof GM_getValue !== 'function') return ls();
          var v = GM_getValue(STORE_KEY);
          if (v && v.records && v.records.length) return resolve(v);
          for (var j = 0; j < LEGACY_KEYS.length; j++) {
            var m2 = migrateLegacy(GM_getValue(LEGACY_KEYS[j]));
            if (m2) return resolve(m2);
          }
          if (v) return resolve(v); // 空 store 也視為存在(避免重複遷移)
        } catch (e) { }
        ls();
      }
      function ls() {
        try {
          var s = localStorage.getItem(STORE_KEY);
          if (s) { var p = JSON.parse(s); if (p && p.records && p.records.length) return resolve(p); }
          for (var k = 0; k < LEGACY_KEYS.length; k++) {
            var sl = localStorage.getItem(LEGACY_KEYS[k]);
            var m3 = sl ? migrateLegacy(JSON.parse(sl)) : null;
            if (m3) return resolve(m3);
          }
          if (s) { var p2 = JSON.parse(s); if (p2 && p2.records) return resolve(p2); }
        } catch (e) { }
        resolve(null);
      }
      fromGm();
    });
  }
  function storeSet(store) {
    return new Promise(function (resolve) {
      var done = false;
      function fin() { if (!done) { done = true; resolve(); } }
      try { if (typeof GM_setValue === 'function') GM_setValue(STORE_KEY, store); } catch (e) { }
      try { localStorage.setItem(STORE_KEY, JSON.stringify(store)); } catch (e) { }
      // 清除 legacy keys(僅在新 store 確有資料時,避免空 store 誤刪舊資料)
      if (store.records && store.records.length) {
        for (var i = 0; i < LEGACY_KEYS.length; i++) {
          try { if (typeof GM_setValue === 'function') GM_setValue(LEGACY_KEYS[i], null); } catch (e) { }
          try { localStorage.removeItem(LEGACY_KEYS[i]); } catch (e) { }
        }
      }
      fin();
    });
  }
  function prune(records) {
    var now = Date.now();
    records = records.filter(function (r) { return now - (r.ts || 0) <= MAX_AGE; });
    if (records.length > MAX_COUNT) records = records.slice(0, MAX_COUNT);
    return records;
  }
  function isDup(store, rec) {
    if (!store) return false;
    var fp = recFp(rec);
    return store.records.some(function (r) {
      return Date.now() - r.ts < 5 * 60 * 1000 && recFp(r) === fp;
    });
  }
  function recFp(rec) {
    if (rec.b64) return rec.size + ':' + rec.b64.slice(0, 64);
    return JSON.stringify(rec.items || []);
  }
  // bridge 100% 只存 items
  OB.Bridge = {
    STORE_KEY: STORE_KEY,
    MAX_AGE: MAX_AGE,
    getAll: storeGet,
    save: storeSet,
    // 合併式存入:同一 source 永遠只有一筆記錄;新單號併入,重複單號忽略
    // → { store, added, dup, created }
    async merge(items, opts) {
      items = items || [];
      opts = opts || {};
      var source = opts.source || location.hostname || 'unknown';
      var store = (await storeGet()) || { activeId: '', records: [] };
      store.records = prune(store.records || []);
      var existing = store.records.find(function (r) { return r.source === source; }) || null;
      var added = 0, dup = 0;
      var map = {};
      if (existing && existing.items) {
        existing.items.forEach(function (it) { if (it && it.billcode) map[it.billcode] = it; });
      }
      items.forEach(function (it) {
        if (!it || !it.billcode) return;
        if (map[it.billcode]) { dup++; return; }
        map[it.billcode] = it;
        added++;
      });
      if (!added) {
        // 無新單號:不碰 store(避免無意義 ts 更新)
        if (existing) store.activeId = existing.id;
        return { store: store, added: added, dup: dup, created: false };
      }
      if (existing) {
        existing.items = Object.keys(map).map(function (k) { return map[k]; });
        existing.ts = Date.now();
        if (opts.name) existing.name = opts.name;
        store.activeId = existing.id;
      } else {
        var rec = { id: String(Date.now()), name: opts.name || 'orders', ts: Date.now(), source: source, items: Object.keys(map).map(function (k) { return map[k]; }) };
        store.records.unshift(rec);
        store.activeId = rec.id;
      }
      store.records = prune(store.records);
      await storeSet(store);
      return { store: store, added: added, dup: dup, created: !existing };
    },
    async add(rec) {
      var store = (await storeGet()) || { activeId: '', records: [] };
      rec.id = rec.id || String(Date.now());
      rec.source = rec.source || location.hostname || 'unknown';
      store.records.unshift(rec);
      store.records = prune(store.records);
      store.activeId = rec.id;
      await storeSet(store);
      return store;
    },
    async remove(id) {
      var store = (await storeGet()) || { activeId: '', records: [] };
      store.records = store.records.filter(function (r) { return r.id !== id; });
      if (store.activeId === id) store.activeId = store.records[0] ? store.records[0].id : '';
      await storeSet(store);
      return store;
    },
    async rename(id, name) {
      var store = (await storeGet()) || { activeId: '', records: [] };
      var r = store.records.find(function (x) { return x.id === id; });
      if (r) { r.name = name; await storeSet(store); }
      return store;
    },
    async setActive(id) {
      var store = (await storeGet()) || { activeId: '', records: [] };
      store.activeId = id;
      var r = store.records.find(function (x) { return x.id === id; });
      if (r) r.used = true;
      await storeSet(store);
      return store;
    },
    async clear() { await storeSet({ activeId: '', records: [] }); }
  };

  // ===================== OB.Sources(輸入端適配器) =====================
  // 每個來源網站一個:{ id, name, match(href), install(api) }
  // api = { notify, loadSheetJs, parse }
  OB.Sources = {
    taobao: {
      id: 'taobao',
      name: '淘寶',
      match: function (href) { return href.indexOf('buyertrade.taobao.com') > -1; },
      // 淘寶訂單導出 = xlsx(商品名称/物流单号/物流公司)
      parseBuffer: function (buf) {
        var XLSX = xlsxLib;
        if (!XLSX) throw new Error(t('parse.sheetjs'));
        var wb = XLSX.read(buf, { type: 'array' });
        return parseWorkbook(wb);
      },
      install: function (api) {
        var self = this;
        // ---------- blob 攔截(通用):createObjectURL + a[download].click ----------
        var blobRef = {};
        var lastCap = null;
        function onCapture(buf, name) {
          var base = { name: name, size: buf.byteLength, ts: Date.now(), source: self.id };
          api.loadSheetJs().then(function () {
            var res;
            try { res = self.parseBuffer(buf); }
            catch (e) {
              // parse 失敗:只通知不存入(不污染暫存區)
              api.notify(t('src.parseFail', { f: name, m: e.message }));
              return Promise.resolve();
            }
            if (!res.items.length) {
              api.notify(t('src.noItems', { f: name, e: res.errors[0] || t('parse.noData') }));
              return Promise.resolve();
            }
            lastCap = { name: base.name, ts: base.ts, source: base.source, items: res.items, errors: res.errors };
            // 彈出選擇視窗:勾選要暫存的項,確認後才合併存入
            return new Promise(function (resolve) {
              showConfirm(res.items, res.errors, function (selItems) {
                if (!selItems.length) { api.notify(t('src.cancel')); resolve(); return; }
                OB.Bridge.merge(selItems, { source: base.source, name: name }).then(function (m) {
                  var msg = t('src.saved', { a: m.added, d: m.dup ? t('src.savedDup', { d: m.dup }) : '', s: base.source });
                  api.notify(msg);
                  resolve();
                }).catch(function (e) { api.notify(t('src.saveFail', { m: e.message })); resolve(); });
              });
            });
          }).catch(function (e) { api.notify(t('src.saveFailFile', { f: name, m: e.message })); });
        }
        function fromBlob(b, name, guard) {
          function save(buf) {
            if (guard && lastCap && lastCap.ts > Date.now() - 5000) return;
            onCapture(buf, name);
          }
          if (b && typeof b.arrayBuffer === 'function') b.arrayBuffer().then(save);
          else { var r = new FileReader(); r.onload = function () { save(new Uint8Array(r.result).buffer); }; r.readAsArrayBuffer(b); }
        }
        // 後台接走:攔截到的下載直接 preventDefault(不存檔/不彈視窗)
        // 想照常下載:DevTools 執行 window.__OB_CAPTURE_SILENT = false
        var silent = true;
        try {
          Object.defineProperty(window, '__OB_CAPTURE_SILENT', {
            configurable: true,
            get: function () { return silent; },
            set: function (v) { silent = !!v; }
          });
        } catch (e) { if (window.__OB_CAPTURE_SILENT === false) silent = false; }
        var suppressedCount = 0;

        var origCreate = URL.createObjectURL.bind(URL);
        URL.createObjectURL = function (obj) {
          var u = origCreate(obj);
          try {
            if (obj && typeof obj.size === 'number' && obj.size > 1024 && obj.size < 30 * 1024 * 1024) {
              blobRef[u] = obj;
              if (obj.type && (obj.type === 'application/octet-stream' || obj.type.indexOf('spreadsheet') > -1)) {
                setTimeout(function () { fromBlob(obj, 'orders.xlsx', true); }, 200);
              }
            }
          } catch (e) { }
          return u;
        };
        function looksLikeExport(a, blob) {
          var type = (blob && blob.type) || '';
          if (type.indexOf('spreadsheet') > -1 || type === 'application/octet-stream') return true;
          var name = (a.getAttribute('download') || '').toLowerCase();
          return name.indexOf('订单') > -1 || name.indexOf('訂單') > -1 || name.indexOf('.xls') > -1;
        }
        document.addEventListener('click', function (e) {
          var a = e.target && e.target.closest ? e.target.closest('a[href^="blob:"]') : null;
          if (!a) return;
          var name = a.getAttribute('download') || 'orders.xlsx';
          var blob = blobRef[a.href];
          if (blob && looksLikeExport(a, blob)) {
            fromBlob(blob, name, false);
            if (silent) { e.preventDefault(); suppressedCount++; }
            return;
          }
          fetch(a.href).then(function (r) { return r.arrayBuffer(); }).then(function (buf) { onCapture(buf, name); }).catch(function () { });
        }, true);
        OB.Sources.taobao.suppressed = function () { return suppressedCount; };

        // ---------- FAB 群組:觸發導出 + 暫存管理 ----------
        var st = document.createElement('style');
        st.textContent = [
          '#ob-src-fabs{position:fixed;right:24px;bottom:140px;z-index:99999;display:flex;gap:8px}',
          '#ob-src-fabs button{border:none;border-radius:24px;padding:12px 16px;font-size:14px;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.3);font-family:inherit;color:#fff;display:flex;align-items:center;justify-content:center;gap:6px}',
          '#ob-src-fab{background:#ff5000}#ob-src-fab:hover{background:#ff6a26}#ob-src-fab:disabled{background:#bbb;cursor:default}',
          '#ob-src-mgr{background:#555}#ob-src-mgr:hover{background:#777}',
          /* 融入淘寶右側工具列(#tb-toolkit-new):48x48 透明磁磚,icon+label 垂直 */
          '#ob-src-fabs.docked{position:static;right:auto;bottom:auto;gap:8px;z-index:auto;display:contents}',
          '#ob-src-fabs.docked button{width:48px;height:48px;border-radius:10px;padding:2px;margin:8px 0 0;flex-direction:column;gap:1px;box-shadow:none;color:#333;background:transparent;font-size:10px}',
          '#ob-src-fabs.docked button .ico{font-size:18px;line-height:1.3}',
          '#ob-src-fabs.docked button .lbl{font-size:10px;line-height:1.2;white-space:nowrap}',
          '#ob-src-fabs.docked #ob-src-fab{background:#fff3ec;color:#ff5000}#ob-src-fabs.docked #ob-src-fab:hover{background:#ffe8d9}',
          '#ob-src-fabs.docked #ob-src-mgr:hover{background:#f2f2f2}',
          '#ob-src-admin{position:fixed;right:24px;bottom:196px;z-index:99999;background:#fff;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.3);width:430px;max-height:60vh;display:none;font-family:inherit;overflow:auto;text-align:left}',
          '#ob-src-admin h4{margin:0;padding:10px 14px;border-bottom:1px solid #eee;font-size:14px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;background:#fff}',
          '.ob-rec{padding:8px 14px;border-bottom:1px solid #f0f0f0;font-size:12px}',
          '.ob-rec input{width:56%;box-sizing:border-box;padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px}',
          '.ob-rec .meta{color:#999;margin:3px 0}',
          '.ob-rec button{padding:2px 10px;font-size:12px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:6px}',
          '.ob-prev{margin-top:6px;border:1px solid #eee;border-radius:4px;max-height:180px;overflow:auto;font-size:11px;background:#fafafa}',
          '.ob-prev table{width:100%;border-collapse:collapse}',
          '.ob-prev td{padding:3px 6px;border-top:1px solid #f0f0f0;vertical-align:top}',
          '.ob-prev td:first-child{white-space:nowrap;color:#555;font-family:monospace}',
          '.ob-prev .more{padding:4px 6px;color:#999;font-style:italic}',
          '#ob-src-confirm{position:fixed;left:50%;top:8%;transform:translateX(-50%);z-index:100000;background:#fff;border-radius:10px;box-shadow:0 10px 40px rgba(0,0,0,.35);width:520px;max-width:92vw;display:none;font-family:inherit;text-align:left}',
          '#ob-src-confirm-mask{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:99998;display:none}',
          '#ob-src-confirm h4{margin:0;padding:12px 16px;border-bottom:1px solid #eee;font-size:15px;display:flex;justify-content:space-between;align-items:center}',
          '#ob-src-confirm .cbody{max-height:52vh;overflow:auto;padding:0 16px}',
          '#ob-src-confirm table{width:100%;border-collapse:collapse;font-size:12px}',
          '#ob-src-confirm td{padding:5px 4px;border-bottom:1px solid #f2f2f2;vertical-align:top}',
          '#ob-src-confirm td:first-child{width:28px}',
          '#ob-src-confirm .cbill{white-space:nowrap;font-family:monospace;color:#555}',
          '#ob-src-confirm .cerr{color:#cf1322;font-size:12px;padding:6px 0}',
          '#ob-src-confirm .cfoot{padding:10px 16px;border-top:1px solid #eee;text-align:right;display:flex;gap:8px;justify-content:flex-end;align-items:center}',
          '#ob-src-confirm .cfoot label{margin-right:auto;font-size:12px;color:#666;cursor:pointer}',
          '#ob-src-confirm button{padding:7px 16px;border:1px solid #d9d9d9;border-radius:5px;background:#fff;cursor:pointer;font-size:13px}',
          '#ob-src-confirm .c-ok{background:#1890ff;color:#fff;border-color:#1890ff}',
          '#ob-src-admin.docked{right:70px;left:auto;top:10vh;bottom:auto;width:580px;max-width:92vw;max-height:80vh}',
          '#ob-src-admin.docked *{font-size:13px}',
          '#ob-src-admin.docked button{font-size:13px;font-family:inherit;color:#333;padding:7px 18px;line-height:1.4}',
          '#ob-src-admin.docked .ob-prev{font-size:12px;max-height:46vh}',
          '#ob-src-admin.docked .ob-prev td{padding:6px 8px}',
          '#ob-src-admin.docked button{font-size:13px!important;color:#333!important;font-family:inherit}'
        ].join('');
        document.head.appendChild(st);

        // ---------- 匯出選擇視窗:解析後勾選要暫存的項 ----------
        var confirmMask = document.createElement('div');
        confirmMask.id = 'ob-src-confirm-mask';
        document.body.appendChild(confirmMask);
        var confirmBox = document.createElement('div');
        confirmBox.id = 'ob-src-confirm';
        document.body.appendChild(confirmBox);
        function showConfirm(items, errors, cb) {
          items = items || [];
          var html =
            '<h4><span>' + t('src.confirmTitle') + '</span>' +
            '<button id="ob-src-cx" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>' +
            '<div class="cbody">' +
            '<table><tbody>' +
            items.map(function (it, i) {
              return '<tr><td><input type="checkbox" class="ob-c-sel" data-i="' + i + '" checked></td>' +
                '<td class="cbill">' + OB.utils.esc(it.billcode) + '</td>' +
                '<td>' + OB.utils.esc(it.goods || '') + (it.company ? ' <span style="color:#999">(' + OB.utils.esc(it.company) + ')</span>' : '') + '</td></tr>';
            }).join('') +
            '</tbody></table>' +
            (errors && errors.length ? '<div class="cerr">⚠ ' + errors.map(OB.utils.esc).join('<br>') + '</div>' : '') +
            '</div>' +
            '<div class="cfoot"><label><input type="checkbox" id="ob-src-call" checked> ' + t('src.selectAll') + '</label>' +
            '<button id="ob-src-ccancel">' + t('src.cancelBtn') + '</button>' +
            '<button class="c-ok" id="ob-src-cok">' + t('src.okBtn') + '</button></div>';
          confirmBox.innerHTML = html;
          confirmMask.style.display = 'block';
          confirmBox.style.display = 'block';
          function close() { confirmMask.style.display = 'none'; confirmBox.style.display = 'none'; }
          function collect() {
            return confirmBox.querySelectorAll('input.ob-c-sel').length === items.length
              ? items
              : items.filter(function (_, i) {
                  var c = confirmBox.querySelector('input.ob-c-sel[data-i="' + i + '"]');
                  return c && c.checked;
                });
          }
          function done(list) { close(); cb(list); }
          $('#ob-src-cx').onclick = function () { done([]); };
          confirmMask.onclick = function () { done([]); };
          $('#ob-src-ccancel').onclick = function () { done([]); };
          $('#ob-src-cok').onclick = function () { done(collect()); };
          $('#ob-src-call').onchange = function () {
            var all = this.checked;
            confirmBox.querySelectorAll('input.ob-c-sel').forEach(function (c) { c.checked = all; });
          };
        }

        var wrap = document.createElement('div');
        wrap.id = 'ob-src-fabs';
        var fab = document.createElement('button');
        fab.id = 'ob-src-fab';
        fab.innerHTML = '<span class="ico">📦</span><span class="lbl">' + t('lblExport') + '</span>';
        var mgrBtn = document.createElement('button');
        mgrBtn.id = 'ob-src-mgr';
        mgrBtn.innerHTML = '<span class="ico">🗂</span><span class="lbl">' + t('lblCache') + '</span>';
        wrap.appendChild(fab);
        wrap.appendChild(mgrBtn);
        function setLbl(b, t2) { var l = b.querySelector('.lbl'); if (l) l.textContent = t2; else b.textContent = t2; }
        // 掛進淘寶右側工具列(#tb-toolkit-new);找不到就退回漂浮 FAB(在 admin 建立後呼叫)。
        // 工具列可能晚於腳本執行才出現,用輪詢補掛(最多 ~20s)。
        function mount() {
          if (!admin) { if (!wrap.parentElement || wrap.parentElement === document.body) document.body.appendChild(wrap); return false; }
          var tk = document.querySelector('#tb-toolkit-new .tb-toolkit-list-new') || document.querySelector('#tb-toolkit-new');
          if (tk) {
            wrap.classList.add('docked'); admin.classList.add('docked');
            wrap.style.display = 'contents'; // 容器消失 → 兩按鈕各自佔工具列一列(各佔一行)
            var lbls = wrap.querySelectorAll('.lbl');
            for (var i = 0; i < lbls.length; i++) lbls[i].style.fontSize = '10px'; // 免被工具列 .lbl 小字體規則影響
            admin.querySelectorAll('button').forEach(function (b) { if (b.style.fontSize) b.style.fontSize = ''; });
            tk.appendChild(wrap); return true;
          }
          if (!wrap.parentElement || wrap.parentElement === document.body) document.body.appendChild(wrap);
          return false;
        }
        mount();
        var mTry = 0;
        var mTimer = setInterval(function () {
          mTry++;
          if (wrap.parentElement !== document.body) { clearInterval(mTimer); return; } // 已掛進工具列
          if (mount()) { clearInterval(mTimer); return; }
          if (mTry >= 20) clearInterval(mTimer); // 20s 後放棄,維持漂浮模式
        }, 1000);

        // 淘寶專屬:點 FAB 自動操作「導出訂單」流程(其他 Source 可換成自己的觸發邏輯)
        fab.onclick = function () {
          fab.disabled = true;
          setLbl(fab, '⏳');
          function fail(msg) { fab.disabled = false; setLbl(fab, t('lblExport')); api.notify('❌ ' + msg); }
          try {
            Array.prototype.slice.call(document.querySelectorAll('.ant-tooltip button, .ant-popover button, [class*=tooltip] button'))
              .forEach(function (b) { if (b.textContent.indexOf('知道了') > -1) b.click(); });
          } catch (e) { }
          setTimeout(function () {
            var btn = document.querySelector('[class*=exportBtn]');
            if (!btn) {
              var cands = Array.prototype.slice.call(document.querySelectorAll('div,span,button,a'));
              btn = cands.find(function (el) {
                var t2 = (el.textContent || '').replace(/\s/g, '');
                return el.childElementCount <= 2 && (t2.indexOf('导出订单') === 0 || t2.indexOf('導出訂單') === 0) && t2.length < 12;
              });
            }
            if (!btn) { fail(t('tb.exportBtnNotFound')); return; }
            btn.click();
            var tries = 0;
            var timer = setInterval(function () {
              tries++;
              var dl = Array.prototype.slice.call(document.querySelectorAll('.ant-modal button, .ant-modal a, .ant-modal [class*=btn]'))
                .find(function (b) {
                  var t2 = (b.textContent || '').replace(/\s/g, '');
                  return (t2.indexOf('下载订单') > -1 || t2.indexOf('下載訂單') > -1) && (b.className || '').indexOf('disabled') === -1;
                });
              if (dl) {
                clearInterval(timer);
                dl.click();
                var w = 0;
                var t2 = setInterval(function () {
                  w++;
                  if (lastCap && lastCap.ts > Date.now() - 15000) { clearInterval(t2); fab.disabled = false; setLbl(fab, t('lblExport')); }
                  else if (w >= 10) { clearInterval(t2); fail(t('tb.captureFail')); }
                }, 1000);
              } else if (tries >= 20) { clearInterval(timer); fail(t('tb.waitTimeout')); }
            }, 1000);
          }, 400);
        };

        // ---------- 暫存管理選單(來源側:查看/編輯/匯出/刪除/清除) ----------
        var admin = document.createElement('div');
        admin.id = 'ob-src-admin';
        document.body.appendChild(admin);

        function renderAdmin(store) {
          store = store || { activeId: '', records: [] };
          var r = store.records[0];
          var title = '<h4><span>' + t('src.adminTitle') + '</span>' +
            '<button id="ob-src-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>';
          if (!r) {
            admin.innerHTML = title +
              '<div class="ob-rec" style="color:#999">' + t('src.empty') + '</div>';
          } else {
            admin.innerHTML = title +
              '<div class="meta" style="padding:8px 14px 0;color:#999;font-size:12px">' +
              OB.utils.fmtTs(r.ts) + ' · ' + t('src.source', { s: r.source || '?' }) +
              (r.items ? ' · ' + t('src.nItems', { n: r.items.length }) : '') + '</div>' +
              (r.items ? '<div class="ob-prev" style="display:block;margin:8px 14px">' + OB.utils.itemsPreviewHtml(r.items) + '</div>' : '');
            var foot = document.createElement('div');
            foot.style.cssText = 'padding:8px 14px;text-align:right;border-top:1px solid #eee';
            foot.innerHTML =
              '<button id="ob-src-admin-dl" style="padding:7px 18px;font-size:13px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:8px">' + t('src.dlBtn') + '</button>' +
              '<button id="ob-src-admin-clear" style="padding:7px 18px;font-size:13px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer">' + t('src.clearBtn') + '</button>';
            admin.appendChild(foot);
          }
          admin.style.display = 'block';
          $('#ob-src-admin-x', admin).onclick = function () { admin.style.display = 'none'; };
          var dlBtn = $('#ob-src-admin-dl');
          if (dlBtn) dlBtn.onclick = function () { if (r) OB.utils.downloadRecord(r); };
          var clBtn = $('#ob-src-admin-clear');
          if (clBtn) clBtn.onclick = function () {
            if (!confirm(t('src.clearAsk'))) return;
            OB.Bridge.clear().then(function () { renderAdmin({ activeId: '', records: [] }); });
          };
        }
        mgrBtn.onclick = function () {
          if (admin.style.display === 'block') { admin.style.display = 'none'; return; }
          OB.Bridge.getAll().then(renderAdmin);
        };
        mount();
        // 語言切換:更新按鈕標籤;若 admin/confirm 開著就重渲染
        document.addEventListener('ob-lang-change', function () {
          var l1 = fab.querySelector('.lbl'); if (l1 && fab.disabled !== true) l1.textContent = t('lblExport');
          var l2 = mgrBtn.querySelector('.lbl'); if (l2) l2.textContent = t('lblCache');
          if (admin.style.display === 'block') OB.Bridge.getAll().then(renderAdmin);
          if (confirmBox.style.display === 'block') confirmBox.style.display = 'none', confirmMask.style.display = 'none';
        });
        }
    }
  };
  function pickSource() {
    var href = location.href;
    for (var k in OB.Sources) {
      var s = OB.Sources[k];
      if (s && s.match && s.match(href)) return s;
    }
    return null;
  }

  // ===================== OB.Sites(輸出端適配器) =====================
  // 每個集運站一個:
  //   { id, name, match(href),
  //     readExisting() → { billcode: {id, goods, number} },
  //     submitNew(chunk, onEachDone(it, ok, msg)),           // 新預報
  //     submitEdit(item, existing) → Promise<{ok, msg}> }    // 補/改品名
  OB.Sites = {
    stjh: {
      id: 'stjh',
      name: '聖天集運',
      match: function (href) {
        return href.indexOf('member.stjh168.com') > -1 && href.indexOf('/Member/MyPack') > -1;
      },
      // 讀取「查看到貨情況」現有記錄
      readExisting: function () {
        var table = document.querySelector('.brp_HouseNameTab_item table');
        if (!table) return {};
        var map = {};
        Array.prototype.slice.call(table.querySelectorAll('tbody tr')).forEach(function (tr) {
          var billEl = tr.querySelector('.kd_billcode');
          var bill = billEl ? billEl.textContent.trim() : '';
          if (!bill) {
            var tds = Array.prototype.slice.call(tr.querySelectorAll('td'));
            for (var i = 0; i < tds.length; i++) {
              var m2 = tds[i].textContent.trim().match(/^[0-9A-Za-z]{8,20}$/);
              if (m2) { bill = m2[0]; break; }
            }
          }
          if (!billcodeValid(bill)) return;
          var goodsEl = tr.querySelector('td.goodsname');
          var goods = goodsEl ? goodsEl.textContent.trim() : '';
          var numEl = tr.querySelector('td.goodsNumber');
          var num = '1';
          if (numEl) { var n = parseInt(numEl.textContent, 10); if (!isNaN(n)) num = String(n); }
          var idEl = tr.querySelector('p[data-id], a[data-id], span[data-id], button[data-id], input[data-id]');
          var id = idEl ? (idEl.getAttribute('data-id') || '') : '';
          if (!id) id = tr.getAttribute('data-id') || '';
          map[bill] = { id: id, goods: goods, number: num };
        });
        return map;
      },
      // 新預報:POST /Member/AddBillcode(欄名 typo 為官方格式;每 10 筆一組)
      submitNew: function (chunk, onEachDone) {
        var datas = chunk.map(function (it) {
          return { kdBillcode: it.billcode, Number: 1, Goods: it.goods, Goodsid: 0, Goods_meno: '', Immediatrly: 0, GoodsType: '' };
        });
        fetch('/Member/AddBillcode', {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
          body: 'datas=' + encodeURIComponent(JSON.stringify(datas))
        })
          .then(function (r) { return r.text(); })
          .then(function (text) {
            var list = null;
            try { list = JSON.parse(text); } catch (e) { }
            if (Array.isArray(list) && list.length === chunk.length) {
              chunk.forEach(function (it, i) {
                var d = list[i];
                var ok = !!(d && d.State);
                onEachDone(it, ok, ok ? '' : (d ? (d.MsgText || t('site.fail')) : t('site.badResp')));
              });
            } else {
              var msg = String(text).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 150);
              chunk.forEach(function (it) { onEachDone(it, false, msg || t('site.badResp')); });
            }
          })
          .catch(function (e) { chunk.forEach(function (it) { onEachDone(it, false, t('site.netErr', { m: e.message })); }); });
      },
      // 補/改品名:POST /Member/EditPackByID(goods_id 必須為 0)
      submitEdit: function (item, existing) {
        return fetch('/Member/EditPackByID', {
          method: 'POST',
          credentials: 'same-origin',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
          body: [
            'id=' + encodeURIComponent(existing.id),
            'goods_id=0',
            'goods=' + encodeURIComponent(item.goods),
            'goods_memo=',
            'kd_billcode=' + encodeURIComponent(item.billcode),
            'immediately=0',
            'goodstype=',
            'Number=' + encodeURIComponent(existing.number || '1')
          ].join('&')
        }).then(function (r) {
          return r.text().then(function (text) {
            try {
              var d = JSON.parse(text);
              var ok = !!(d && d.State);
              return { ok: ok, msg: d ? (d.MsgText || '') : '' };
            } catch (e) {
              return { ok: text.indexOf('成功') > -1, msg: text.slice(0, 120) };
            }
          });
        });
      }
    }
  };
  function pickSite() {
    var href = location.href;
    for (var k in OB.Sites) {
      var a = OB.Sites[k];
      if (a && a.match && a.match(href)) return a;
    }
    return null;
  }

  // ===================== OB.UI(輸出端面板/FAB/預覽/管理選單) =====================
  // 自訂按鈕/外觀:每個 Site 可宣告 site.uiStyle(CSS 字串),會追加在預設樣式之後,
  // 可覆寫 #ob-fab / #ob-fab-mgr / #ob-panel 等任何 id。例:
  //   OB.Sites.example = { ..., uiStyle: "#ob-fab{background:#ff5000;bottom:40px}" }
  // 或用 OB.UI.setStyle(css) 做全域覆寫(供其他腳本/進階使用者)。
  OB.UI = {
    _extraStyle: '',
    setStyle: function (css) { OB.UI._extraStyle = css || ''; return OB.UI._extraStyle; },

    // ---------- 檢視模式:非來源/非目標網頁也能從 TM 選單看暫存清單 ----------
    installViewer: function () {
      var VIEWER_CSS =
        '#obv-mask{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:99997;display:none}' +
        '#obv{position:fixed;left:50%;top:5vh;transform:translateX(-50%);width:min(92vw,900px);max-height:90vh;overflow:auto;background:#fff;z-index:99998;box-shadow:0 10px 40px rgba(0,0,0,.35);padding:16px 20px;box-sizing:border-box;border-radius:10px;font-family:inherit;font-size:13px;text-align:left}' +
        '#obv h3{margin:0 0 10px;font-size:16px;display:flex;justify-content:space-between;align-items:center}' +
        '#obv .x{cursor:pointer;font-size:20px;color:#999}' +
        '#obv .meta{color:#999;font-size:12px;margin:4px 0}' +
        '#obv table{width:100%;border-collapse:collapse;font-size:12px}' +
        '#obv td{padding:4px 8px;border-top:1px solid #f0f0f0;vertical-align:top}' +
        '#obv td.c-bill{white-space:nowrap;color:#555;font-family:monospace}' +
        '#obv .btns{margin-top:10px;text-align:right}' +
        '#obv .btns button{padding:6px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;font-size:13px;margin-left:8px}' +
        '#obv .btns .danger{color:#cf1322;border-color:#ffa39e}';

      function ensureDom() {
        if (document.getElementById('obv')) return document.getElementById('obv');
        var st = document.createElement('style');
        st.textContent = VIEWER_CSS + (OB.UI._extraStyle || '');
        document.head.appendChild(st);
        var mask = document.createElement('div'); mask.id = 'obv-mask';
        var box = document.createElement('div'); box.id = 'obv';
        mask.onclick = function () { mask.style.display = 'none'; box.style.display = 'none'; };
        document.body.appendChild(mask); document.body.appendChild(box);
        return box;
      }

      function close() { var b = document.getElementById('obv'); if (b) b.style.display = 'none'; var m = document.getElementById('obv-mask'); if (m) m.style.display = 'none'; }
      function open() {
        var box = ensureDom();
        document.getElementById('obv-mask').style.display = 'block';
        box.innerHTML = '<h3><span>' + t('vw.titleFull') + '</span><span class="x" data-x>✕</span></h3><div class="meta">' + t('vw.loading') + '</div>' + OB.i18n.langHtml();
        box.style.display = 'block';
        box.querySelector('[data-x]').onclick = close;
        OB.i18n.bindLangSel(box);

        OB.Bridge.getAll().then(function (st) {
          st = st || { records: [] };
          var recs = st.records || [];
          var html = '<h3><span>' + t('vw.titleFull') + '</span><span class="x" data-x>✕</span></h3>' + OB.i18n.langHtml();
          if (!recs.length) {
            html += '<div class="meta">' + t('vw.emptyNote') + '</div>';
            box.innerHTML = html;
            box.querySelector('[data-x]').onclick = close;
            OB.i18n.bindLangSel(box);
            return;
          }
          recs.forEach(function (r) {
            html += '<div style="margin-bottom:14px">' +
              '<div class="meta"><b>' + OB.utils.esc(r.name || r.id) + (st.activeId === r.id ? ' ' + t('vw.activeTag') : '') + ' · ' + t('src.source', { s: OB.utils.esc(r.source || '?') }) +
              ' · ' + OB.utils.fmtTs(r.ts) + (r.items ? ' ' + t('vw.nItems', { n: r.items.length }) : '') + '</div>' +
              (r.items && r.items.length ? '<div style="margin:6px 0;border:1px solid #eee;border-radius:4px;max-height:260px;overflow:auto">' + OB.utils.itemsPreviewHtml(r.items) + '</div>' : '<div class="meta">' + t('vw.noItems') + '</div>') +
              '</div>';
          });
          html += '<div class="btns">' +
            '<button id="obv-export">' + t('vw.exportAll') + '</button>' +
            '<button id="obv-clear" class="danger">' + t('vw.clear') + '</button></div>';
          box.innerHTML = html;
          box.querySelector('[data-x]').onclick = close;
          OB.i18n.bindLangSel(box);
          var active = null;
          if (st.activeId) active = recs.find(function (r) { return r.id === st.activeId; });
          if (!active && recs.length) active = recs[0];
          document.getElementById('obv-export').onclick = function () { if (active) OB.utils.downloadRecord(active); };
          document.getElementById('obv-clear').onclick = function () {
            if (!confirm(t('vw.clearAsk'))) return;
            OB.Bridge.clear().then(function () { open(); });
          };
          return;
        }).catch(function (e) {
          box.innerHTML = '<h3><span>' + t('vw.titleFull') + '</span><span class="x" data-x>✕</span></h3><div class="meta" style="color:#cf1322">' + t('vw.readFail', { m: OB.utils.esc(e.message) }) + '</div>' + OB.i18n.langHtml();
          box.querySelector('[data-x]').onclick = close;
          OB.i18n.bindLangSel(box);
        });
      }

      if (typeof GM_registerMenuCommand === 'function') {
        var cmd = GM_registerMenuCommand(t('api.menuTitle'), function () { open(); });
        if (typeof GM_registerMenuCommandClose === 'function' && cmd) {
          GM_registerMenuCommandClose(cmd, close);
        }
      }
      // 語言切換:若 modal 開著就重渲染
      document.addEventListener('ob-lang-change', function () {
        var b = document.getElementById('obv');
        if (b && b.style.display !== 'none') open();
      });
      // 備用快捷鍵:Ctrl+Shift+B 開啟檢視 modal(方便在任意網頁快速查看)
      document.addEventListener('keydown', function (e) {
        if (e && e.ctrlKey && e.shiftKey && (e.key === 'B' || e.key === 'b')) {
          e.preventDefault(); e.stopPropagation(); open();
        }
      }, true);
    },
    install: function (site) {
      var MAX_ROWS_PER_SUBMIT = 10;
      var emptyState = function () { return { items: [], existing: {}, parsed: false }; };
      var state = emptyState();
      var store = null; // 暫存記錄快取

      var STYLE = [
        '#ob-fab-dock{position:fixed;right:22px;bottom:110px;z-index:99999;display:flex;flex-direction:column;gap:6px;padding:8px;background:rgba(255,255,255,.96);backdrop-filter:blur(6px);border:1px solid rgba(0,0,0,.06);border-radius:16px;box-shadow:0 6px 24px rgba(0,0,0,.16)}',
        '#ob-fab,#ob-fab-mgr{position:static;align-self:center;display:flex;flex-direction:column;align-items:center;gap:2px;border:none;background:transparent;cursor:pointer;padding:8px 10px;border-radius:12px;font-family:inherit;line-height:1.1;transition:background .15s}',
        '#ob-fab .ob-fico,#ob-fab-mgr .ob-fico{font-size:24px;line-height:1}',
        '#ob-fab .ob-flbl,#ob-fab-mgr .ob-flbl{font-size:11px;font-weight:600}',
        '#ob-fab{color:#1890ff}#ob-fab:hover{background:#e6f4ff}',
        '#ob-fab-mgr{color:#737373}#ob-fab-mgr:hover{background:#f5f5f5}',
        '@media (max-width:768px){#ob-fab-dock{padding:6px;gap:2px}#ob-fab,#ob-fab-mgr{padding:6px 8px}#ob-fab .ob-fico,#ob-fab-mgr .ob-fico{font-size:20px}#ob-fab .ob-flbl,#ob-fab-mgr .ob-flbl{font-size:10px}}',
        '#ob-panel-mask{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:99997;display:none}',
        '#ob-panel{position:fixed;left:50%;top:4vh;transform:translateX(-50%);width:min(96vw,1280px);max-height:92vh;overflow:auto;background:#fff;z-index:99998;box-shadow:0 10px 40px rgba(0,0,0,.35);padding:18px 22px;box-sizing:border-box;border-radius:10px}',
        '#ob-panel h3{margin:0 0 12px;font-size:17px}',
        '#ob-panel .ob-close{position:sticky;float:right;top:0;cursor:pointer;font-size:22px;color:#999}',
        '#ob-status{padding:8px 0;font-size:13px;color:#555}',
        '#ob-bridge-status{padding:6px 0;font-size:13px;color:#1890ff}',
        '#ob-bridge-btns{padding:6px 0}',
        '#ob-bridge-btns button,#ob-panel button{margin:4px 6px 0 0;padding:6px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;font-size:13px}',
        '#ob-panel .primary{background:#1890ff;color:#fff;border-color:#1890ff}',
        '#ob-panel .primary:disabled{background:#91d5ff;cursor:default}',
        '.ob-badge{display:inline-block;padding:2px 10px;border-radius:10px;font-size:12px;margin-right:6px}',
        '.ob-b-new{background:#e6f7e6;color:#389e0d}.ob-b-edit{background:#fff7e6;color:#d46b08}.ob-b-editnew{background:#fff1f0;color:#cf1322}',
        '#ob-table{width:100%;border-collapse:collapse;font-size:12px;margin:8px 0}',
        '#ob-table th,#ob-table td{border:1px solid #eee;padding:5px 8px;text-align:left;vertical-align:top;word-break:break-all}',
        '#ob-table th{background:#fafafa}',
        '#ob-table input{width:100%;box-sizing:border-box;padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px}',
        '#ob-table .c-bill{width:110px}#ob-table .c-goods{width:200px}',
        '.ob-err{color:#cf1322;font-size:12px;padding:4px 0}',
        '.ob-prev{margin-top:6px;border:1px solid #eee;border-radius:4px;max-height:180px;overflow:auto;font-size:11px;background:#fafafa}',
        '.ob-prev table{width:100%;border-collapse:collapse}',
        '.ob-prev td{padding:3px 6px;border-top:1px solid #f0f0f0;vertical-align:top}',
        '.ob-prev td:first-child{white-space:nowrap;color:#555;font-family:monospace}',
        '.ob-prev .more{padding:4px 6px;color:#999;font-style:italic}',
        '.ob-ok{color:#389e0d;font-size:12px;padding:4px 0}',
        '#ob-admin{position:fixed;right:24px;bottom:256px;z-index:99999;background:#fff;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.3);width:430px;max-height:60vh;display:none;font-family:inherit;overflow:auto;text-align:left}',
        '#ob-admin h4{margin:0;padding:10px 14px;border-bottom:1px solid #eee;font-size:14px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;background:#fff}',
        '#ob-admin .ob-rec{padding:8px 14px;border-bottom:1px solid #f0f0f0;font-size:12px}',
        '#ob-admin .ob-rec input{width:56%;box-sizing:border-box;padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px}',
        '#ob-admin .ob-rec .meta{color:#999;margin:3px 0}',
        '#ob-admin .ob-rec button{padding:2px 10px;font-size:12px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:6px}'
      ].join('');
      var st = document.createElement('style');
      st.textContent = STYLE + (site.uiStyle || '') + OB.UI._extraStyle;
      document.head.appendChild(st);

      var fabDock = document.createElement('div');
      fabDock.id = 'ob-fab-dock';
      var fab = document.createElement('button');
      fab.id = 'ob-fab';
      fab.title = t('st.titleImport', { s: site.name });
      fab.innerHTML = '<span class="ob-fico">📦</span><span class="ob-flbl">' + t('st.lblImport') + '</span>';
      var fabMgr = document.createElement('button');
      fabMgr.id = 'ob-fab-mgr';
      fabMgr.title = t('st.titleCache');
      fabMgr.innerHTML = '<span class="ob-fico">🗂</span><span class="ob-flbl">' + t('st.lblCache') + '</span>';
      fabDock.appendChild(fab);
      fabDock.appendChild(fabMgr);
      document.body.appendChild(fabDock);

      // ---------- 預覽面板(中央 modal) ----------
      var panelMask = document.createElement('div');
      panelMask.id = 'ob-panel-mask';
      document.body.appendChild(panelMask);
      var panel = document.createElement('div');
      panel.id = 'ob-panel';
      panel.style.display = 'none';
      function openPanel() { panel.style.display = 'block'; panelMask.style.display = 'block'; }
      function closePanel() { panel.style.display = 'none'; panelMask.style.display = 'none'; }
      panelMask.onclick = closePanel;
      panel.innerHTML =
        '<span class="ob-close" data-ob="close">✕</span>' +
        '<h3>' + esc(t('ui.importTo', { s: esc(site.name) })) + '</h3>' +
        '<div id="ob-bridge-status"></div>' +
        '<div id="ob-bridge-btns"><button data-ob="auto">' + t('ui.auto') + '</button><button data-ob="pick-file">' + t('ui.pickFile') + '</button><button data-ob="reload-bridge">' + t('ui.reload') + '</button></div>' +
        '<div id="ob-status">' + t('ui.idle') + '</div>' +
        '<div><label><input type="checkbox" data-ob="edit-empty" checked> ' + t('ui.editEmpty') + '</label>' +
        '<label style="margin-left:12px"><input type="checkbox" data-ob="edit-fill"> ' + t('ui.editFill') + '</label></div>' +
        '<div style="margin-top:6px"><label><input type="checkbox" data-ob="auto-refresh" checked> ' + t('ui.autoRefresh') + '</label></div>' +
        '<table id="ob-table" style="display:none"><thead><tr>' +
        '<th style="width:30px"><input type="checkbox" id="ob-check-all" checked></th>' +
        '<th class="c-bill">' + t('ui.colBill') + '</th><th>' + t('ui.colCompany') + '</th><th class="c-goods">' + t('ui.colGoods') + '</th><th>' + t('ui.colAction') + '</th><th>' + t('ui.colResult') + '</th>' +
        '</tr></thead><tbody id="ob-tbody"></tbody></table>' +
        '<div id="ob-result"></div>' +
        '<div style="margin-top:12px"><button class="primary" data-ob="submit" disabled>' + t('ui.submit') + '</button>' +
        '<button data-ob="refresh">' + t('ui.reanalyze') + '</button></div>' +
        OB.i18n.langHtml();
      document.body.appendChild(panelMask);
      document.body.appendChild(panel);
      OB.i18n.bindLangSel(panel);
      function openPanel() { panel.style.display = 'block'; panelMask.style.display = 'block'; }
      function closePanel() { panel.style.display = 'none'; panelMask.style.display = 'none'; }

      function $(sel, root) { return OB.utils.$(sel, root || panel); }
      function esc(s) { return OB.utils.esc(s); }
      function setStatus(html) { $('#ob-status').innerHTML = html; }
      function setBridgeStatus(html) { $('#ob-bridge-status').innerHTML = html; }

      // ---------- 暫存管理選單 ----------
      var admin = document.createElement('div');
      admin.id = 'ob-admin';
      document.body.appendChild(admin);

      function refreshStore(then) {
        return OB.Bridge.getAll().then(function (s) { store = s || { activeId: '', records: [] }; if (then) then(store); });
      }
      function renderAdmin() {
        refreshStore(function (st) {
          var r = (st.records || [])[0];
          var title = '<h4><span>' + t('st.adminTitle') + '</span>' +
            '<button id="ob-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>';
          if (!r) {
            admin.innerHTML = title +
              '<div class="ob-rec" style="color:#999">' + t('st.empty') + '</div>';
          } else {
            admin.innerHTML = title +
              '<div class="meta" style="padding:8px 14px 0;color:#999;font-size:12px">' +
              OB.utils.fmtTs(r.ts) + ' · ' + t('src.source', { s: r.source || '?' }) +
              (r.items ? ' · ' + t('src.nItems', { n: r.items.length }) : '') + '</div>' +
              (r.items ? '<div class="ob-prev" style="display:block;margin:8px 14px">' + OB.utils.itemsPreviewHtml(r.items) + '</div>' : '');
            var foot2 = document.createElement('div');
            foot2.style.cssText = 'padding:8px 14px;text-align:right;border-top:1px solid #eee';
            foot2.innerHTML =
              '<button id="ob-admin-dl" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:8px">' + t('src.dlBtn') + '</button>' +
              '<button id="ob-admin-clear" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer">' + t('src.clearBtn') + '</button>';
            admin.appendChild(foot2);
            $('#ob-admin-dl', admin).onclick = function () { OB.utils.downloadRecord(r); };
            $('#ob-admin-clear', admin).onclick = function () {
              if (!confirm(t('src.clearAsk'))) return;
              OB.Bridge.clear().then(function () { renderAdmin(); refreshBridgeUI(); });
            };
          }
          admin.style.display = 'block';
          $('#ob-admin-x', admin).onclick = function () { admin.style.display = 'none'; };
        });
      }
      fabMgr.onclick = function () {
        if (admin.style.display === 'block') { admin.style.display = 'none'; return; }
        renderAdmin();
      };

      // ---------- 暫存 → 解析 ----------
      async function autoLoadBridge() {
        await refreshStore();
        var rec = null, how = null;
        if (store && store.activeId) {
          rec = store.records.find(function (r) { return r.id === store.activeId; });
          if (rec) how = 'activeId';
        }
        if (!rec && store && store.records.length) { rec = store.records[0]; how = 'latest'; }
        if (!rec) return null;
        var items, errors;
        if (rec.items) { items = rec.items; errors = rec.errors || []; }
        else if (rec.b64) {
          setStatus(t('ui.readingOld'));
          var buf = OB.utils.b64ToBuf(rec.b64);
          await OB.utils.loadSheetJs();
          var pbuf = (OB.Sources[rec.source] && OB.Sources[rec.source].parseBuffer) || OB.Sources.taobao.parseBuffer;
          var res;
          try { res = pbuf(buf); items = res.items; errors = res.errors; }
          catch (e) { items = []; errors = [t('ui.oldParseFail', { m: e.message })]; }
          if (items.length) {
            rec.items = items; rec.errors = errors; delete rec.b64;
            await OB.Bridge.save(store); // 一次性遷移:之後直接存 items
          }
        }
        else { items = []; errors = [t('ui.noDataInRec')]; }
        state.items = items;
        state.parsed = true;
        state.rec = rec;
        await OB.Bridge.setActive(rec.id);
        setStatus(t('ui.readOk', {
          f: esc(rec.name),
          s: t('src.source', { s: esc(rec.source || '?') }),
          n: items.length,
          e: errors.length ? t('ui.readErrs', { e: errors.map(esc).join('<br>') }) : ''
        }));
        return rec;
      }
      function refreshBridgeUI() {
        refreshStore(); // 只在背景更新資料;啟動時不顯示任何字
      }

      // ---------- 檔案選擇(手動備用) ----------
      function pickFile() {
        var inp = document.createElement('input');
        inp.type = 'file';
        inp.accept = '.xlsx,.xls';
        inp.onchange = async function () {
          var f = inp.files[0];
          if (!f) return;
          setStatus(t('ui.fileParsing', { f: esc(f.name) }));
          try {
            var buf = await f.arrayBuffer();
            await OB.utils.loadSheetJs();
            var res = OB.Sources.taobao.parseBuffer(buf); // 手動匯入固定為淘寶 xlsx 格式
            state.items = res.items;
            state.parsed = true;
            state.rec = { id: 'file', name: f.name, size: f.size, ts: Date.now(), source: t('ui.srcManual'), items: res.items, errors: res.errors };
            setStatus(t('ui.fileOk', {
              f: esc(f.name), n: res.items.length,
              e: res.errors.length ? t('ui.readErrs', { e: res.errors.map(esc).join('<br>') }) : ''
            }));
          } catch (e) {
            setStatus(t('ui.fileParseFail', { m: esc(e.message) }));
          }
        };
        inp.click();
      }

      // ---------- 分類 + 預覽 ----------
      function classify() {
        var ex = state.existing || {};
        return state.items.map(function (it) {
          var e = ex[it.billcode];
          if (!e) return { item: it, existing: null, action: 'new' };
          if (!e.goods) return { item: it, existing: e, action: 'edit-new' };
          return { item: it, existing: e, action: 'edit-fill' };
        });
      }
      function runAnalysis() {
        if (!state.parsed) { setStatus(t('ui.noParsed')); return; }
        state.existing = site.readExisting();
        var rows = classify();
        var nNew = rows.filter(function (r) { return r.action === 'new'; }).length;
        var nEN = rows.filter(function (r) { return r.action === 'edit-new'; }).length;
        var nEF = rows.filter(function (r) { return r.action === 'edit-fill'; }).length;
        var editEmptyOn = $('[data-ob="edit-empty"]').checked;
        var editFillOn = $('[data-ob="edit-fill"]').checked;
        var tbody = $('#ob-tbody');
        tbody.innerHTML = '';
        var count = 0;
        rows.forEach(function (row, idx) {
          var doEdit = (row.action === 'edit-new' && editEmptyOn) || (row.action === 'edit-fill' && editFillOn);
          var badge = '';
          if (row.action === 'new') { badge = '<span class="ob-badge ob-b-new">' + t('ui.bNew') + '</span>'; count++; }
          else if (row.action === 'edit-new') { badge = '<span class="ob-badge ' + (doEdit ? 'ob-b-edit' : 'ob-b-editnew') + '">' + (doEdit ? t('ui.bAddName') : t('ui.bNeedCheck')) + '</span>'; if (doEdit) count++; }
          else { badge = '<span class="ob-badge ' + (doEdit ? 'ob-b-edit' : 'ob-b-editnew') + '">' + (doEdit ? t('ui.bUpdateName') : t('ui.bNeedCheck')) + '</span>'; if (doEdit) count++; }
          var tr = document.createElement('tr');
          tr.setAttribute('data-idx', idx);
          var checked = (row.action === 'new') ? 'checked' : (doEdit ? 'checked' : '');
          tr.innerHTML =
            '<td><input type="checkbox" data-f="sel" ' + checked + '></td>' +
            '<td class="c-bill">' + esc(row.item.billcode) + '</td>' +
            '<td>' + esc(row.item.company || '') + '</td>' +
            '<td class="c-goods"><input data-f="goods" value="' + esc(row.item.goods) + '"></td>' +
            '<td>' + badge +
            (row.existing && row.existing.goods ? '<div style="color:#999;font-size:11px">' + t('ui.existing', { g: esc(row.existing.goods) }) + '</div>' : '') +
            '</td><td class="c-res">-</td>';
          tbody.appendChild(tr);
        });
        $('#ob-table').style.display = 'block';
        $('[data-ob="submit"]').disabled = !count;
        state.pending = rows;
        state.planCount = count;
        $('#ob-check-all').checked = true;
        setStatus(t('ui.plan', { n: rows.length, a: nNew, b: nEN, c: nEF }));
      }

      // ---------- 提交 ----------
      function readGoodsFromRow(tr) {
        var inp = tr.querySelector('input[data-f="goods"]');
        return inp ? inp.value.trim() : '';
      }
      function submit() {
        var doEditEmpty = $('[data-ob="edit-empty"]').checked;
        var doEditFill = $('[data-ob="edit-fill"]').checked;
        var tbodyRows = Array.prototype.slice.call($('#ob-tbody').querySelectorAll('tr'));
        var jobs = [];
        state.pending.forEach(function (row, idx) {
          var tr = tbodyRows[idx];
          var selChk = tr.querySelector('input[data-f="sel"]');
          if (selChk && !selChk.checked) return; // 未勾選 → 跳過
          var goods = readGoodsFromRow(tr);
          if (!goods) goods = row.item.goods;
          var act = 'skip';
          if (row.action === 'new') act = 'new';
          else if (row.action === 'edit-new' && doEditEmpty) act = 'edit';
          else if (row.action === 'edit-fill' && doEditFill) act = 'edit';
          if (act !== 'skip') jobs.push({ act: act, row: row, tr: tr, goods: goods });
        });
        if (!jobs.length) { setStatus(t('ui.noneChecked')); return; }
        $('[data-ob="submit"]').disabled = true;
        var resultDiv = $('#ob-result');
        resultDiv.innerHTML = '<div class="ob-ok">' + t('ui.submitting') + '</div>';
        var okN = 0, failN = 0;
        // 1) 新預報:分組(每 10 筆)
        var newJobs = jobs.filter(function (j) { return j.act === 'new'; });
        var editJobs = jobs.filter(function (j) { return j.act === 'edit'; });
        var chunks = [];
        for (var i = 0; i < newJobs.length; i += MAX_ROWS_PER_SUBMIT) chunks.push(newJobs.slice(i, i + MAX_ROWS_PER_SUBMIT));
        function markJob(j, ok, msg) {
          var resTd = j.tr.querySelector('.c-res');
          if (ok) { okN++; resTd.innerHTML = '<span class="ob-ok">' + esc(j.act === 'new' ? t('ui.okNew') : t('ui.okEdit')) + '</span>'; }
          else { failN++; resTd.innerHTML = '<span class="ob-err">' + esc(t('ui.fail', { m: msg || t('site.fail') })) + '</span>'; }
        }
        function done() {
          $('[data-ob="submit"]').disabled = false;
          var ar = $('[data-ob="auto-refresh"]');
          var willRefresh = ar && ar.checked;
          resultDiv.innerHTML = '<div class="' + (failN ? 'ob-err' : 'ob-ok') + '">' +
            (failN ? t('ui.done', { a: okN, b: failN }) : t('ui.doneAll', { n: okN })) +
            (willRefresh ? t('ui.willRefresh') : '') + '</div>';
          OB.utils.notify(t('ui.notifyDone', { a: okN, b: failN, r: willRefresh ? t('ui.refreshNote') : '' }));
          if (willRefresh) {
            try { sessionStorage.setItem('__OB_REOPEN', '1'); } catch (e) { }
            setTimeout(function () { location.reload(); }, 3000);
          }
        }
        var pend = newJobs.length + editJobs.length;
        if (!pend) { done(); return; }
        var settled = 0;
        function oneSettled() {
          settled++;
          if (settled === pend) done();
        }
        chunks.forEach(function (chunk) {
          site.submitNew(chunk.map(function (j) {
            return { billcode: j.row.item.billcode, company: j.row.item.company, goods: j.goods };
          }), function (it, ok, msg) {
            var j = chunk.find(function (x) { return x.row.item.billcode === it.billcode; });
            if (j) markJob(j, ok, msg);
            oneSettled();
          });
        });
        editJobs.forEach(function (j) {
          site.submitEdit({ billcode: j.row.item.billcode, goods: j.goods }, j.row.existing)
            .then(function (r) { markJob(j, !!(r && r.ok), r && r.msg); })
            .catch(function (e) { markJob(j, false, e.message); })
            .then(oneSettled);
        });
      }

      // ---------- 事件 ----------
      panel.addEventListener('click', function (e) {
        var el = e.target;
        while (el && el !== panel && !el.getAttribute) el = el.parentNode;
        if (!el) return;
        var ob = el.getAttribute && el.getAttribute('data-ob');
        if (!ob) return;
        if (ob === 'close') closePanel();
        else if (ob === 'reload-bridge') {
          autoLoadBridge().then(function (rec) {
            if (!rec) { setStatus(t('ui.noBridge')); return; }
            runAnalysis();
          }).catch(function (e) { setStatus(t('ui.err', { m: esc(e.message) })); });
        }
        else if (ob === 'auto') {
          autoLoadBridge().then(function (rec) {
            if (!rec) { setStatus(t('ui.noBridge')); return; }
            runAnalysis();
          }).catch(function (e) { setStatus(t('ui.err', { m: esc(e.message) })); });
        }
        else if (ob === 'pick-file') pickFile();
        else if (ob === 'refresh') runAnalysis();
        else if (ob === 'submit') submit();
      });
      $('[data-ob="edit-empty"]').addEventListener('change', function () { if (state.parsed) runAnalysis(); });
      $('[data-ob="edit-fill"]').addEventListener('change', function () { if (state.parsed) runAnalysis(); });
      $('#ob-check-all').addEventListener('change', function () {
        var all = this.checked;
        $('#ob-tbody').querySelectorAll('input[data-f="sel"]').forEach(function (c) { c.checked = all; });
      });

      fab.onclick = function () {
        openPanel();
        state = emptyState();
        refreshBridgeUI();
        autoLoadBridge().then(function (rec) {
          if (!rec) { setStatus(t('ui.noBridge')); return; }
          runAnalysis();
        }).catch(function (e) { setStatus(t('ui.autoFail', { m: esc(e.message) })); });
      };
      // 語言切換:重建面板靜態文字 + 依狀態重渲染;保留勾選狀態
      document.addEventListener('ob-lang-change', function () {
        fab.title = t('st.titleImport', { s: site.name });
        fab.querySelector('.ob-flbl').textContent = t('st.lblImport');
        fabMgr.title = t('st.titleCache');
        fabMgr.querySelector('.ob-flbl').textContent = t('st.lblCache');
        if (panel.style.display !== 'none') {
          var checkedRows = [];
          $('#ob-tbody').querySelectorAll('tr').forEach(function (tr) {
            var c = tr.querySelector('input[data-f="sel"]');
            checkedRows.push(!!(c && c.checked));
          });
          var editE = $('[data-ob="edit-empty"]').checked, editF = $('[data-ob="edit-fill"]').checked;
          panel.innerHTML =
            '<span class="ob-close" data-ob="close">✕</span>' +
            '<h3>' + esc(t('ui.importTo', { s: esc(site.name) })) + '</h3>' +
            '<div id="ob-bridge-status"></div>' +
            '<div id="ob-bridge-btns"><button data-ob="auto">' + t('ui.auto') + '</button><button data-ob="pick-file">' + t('ui.pickFile') + '</button><button data-ob="reload-bridge">' + t('ui.reload') + '</button></div>' +
            '<div id="ob-status">' + (state.parsed ? '<span style="color:#999">…</span>' : t('ui.idle')) + '</div>' +
            '<div><label><input type="checkbox" data-ob="edit-empty"' + (editE ? ' checked' : '') + '> ' + t('ui.editEmpty') + '</label>' +
            '<label style="margin-left:12px"><input type="checkbox" data-ob="edit-fill"' + (editF ? ' checked' : '') + '> ' + t('ui.editFill') + '</label></div>' +
            '<div style="margin-top:6px"><label><input type="checkbox" data-ob="auto-refresh" checked> ' + t('ui.autoRefresh') + '</label></div>' +
            '<table id="ob-table"' + (state.parsed ? '' : ' style="display:none"') + '><thead><tr>' +
            '<th style="width:30px"><input type="checkbox" id="ob-check-all" checked></th>' +
            '<th class="c-bill">' + t('ui.colBill') + '</th><th>' + t('ui.colCompany') + '</th><th class="c-goods">' + t('ui.colGoods') + '</th><th>' + t('ui.colAction') + '</th><th>' + t('ui.colResult') + '</th>' +
            '</tr></thead><tbody id="ob-tbody"></tbody></table>' +
            '<div id="ob-result"></div>' +
            '<div style="margin-top:12px"><button class="primary" data-ob="submit" disabled>' + t('ui.submit') + '</button>' +
            '<button data-ob="refresh">' + t('ui.reanalyze') + '</button></div>' +
            OB.i18n.langHtml();
          OB.i18n.bindLangSel(panel);
          if (state.parsed && state.rec) {
            autoLoadBridge().then(function (rec) {
              if (!rec) { setStatus(t('ui.noBridge')); return; }
              runAnalysis();
              if (checkedRows.length) {
                $('#ob-tbody').querySelectorAll('tr').forEach(function (tr, i) {
                  var c = tr.querySelector('input[data-f="sel"]');
                  if (c && checkedRows[i] === false) c.checked = false;
                });
              }
            });
          }
        }
        if (admin.style.display === 'block') renderAdmin();
      });
      panel.style.display = 'none';
      refreshBridgeUI();
    }
  };

  // ===================== 路由 =====================
  (function boot() {
    var src = pickSource();
    if (src) {
      src.install({
        notify: notify,
        loadSheetJs: loadSheetJs,
        parseBuffer: (src.parseBuffer) || null,
        parse: (src.parseBuffer) ? function (buf) { return src.parseBuffer(buf); } : null
      });
      return;
    }
    var site = pickSite();
    if (site) {
      OB.UI.install(site);
      // 重新整理後自動重新打開導入面板(提交完觸發的 reload)
      try {
        if (sessionStorage.getItem('__OB_REOPEN')) {
          sessionStorage.removeItem('__OB_REOPEN');
          var f = document.getElementById('ob-fab');
          if (f) setTimeout(function () { f.click(); }, 600);
        }
      } catch (e) { }
    } else {
      OB.UI.installViewer();
    }
  })();
})();
