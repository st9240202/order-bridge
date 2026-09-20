// ==UserScript==
// @name         訂單跨站導入橋 (Order Bridge)
// @namespace    https://tampermonkey.net/
// @version      3.2.3
// @match        https://buyertrade.taobao.com/trade/itemlist/*
// @match        http://member.stjh168.com/Member/MyPack
// @description  通用訂單 xlsx 跨站橋:輸入端 OB.Sources(暫存/管理)+ 輸出端 OB.Sites(適配器)。現含:淘寶 → 聖天集運。擴充新站點只需加一個 Source/Site 定義。
// @grant        GM_setValue
// @grant        GM_getValue
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

  // ===================== OB.utils =====================
  function $(sel, root) { return (root || document).querySelector(sel); }
  function esc(s) {
    var d = document.createElement('div');
    d.textContent = s == null ? '' : String(s);
    return d.innerHTML;
  }
  function billcodeValid(s) { return typeof s === 'string' && /^[0-9A-Za-z]{8,20}$/.test(s); }
  function notify(text) {
    try { if (typeof GM_notify === 'function') { GM_notify({ title: '訂單導入橋', text: text, priority: 2 }); return; } } catch (e) { }
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
        if (!X || typeof X.read !== 'function') throw new Error('載入後找不到 XLSX');
        xlsxLib = X;
        return X;
      }).catch(function (e) {
        if (i >= urls.length) throw new Error('無法載入 SheetJS(CDN 全失敗):' + e.message);
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
    if (!items.length) return '<div class="more">(無資料)</div>';
    var MAX = 50;
    var shown = items.slice(0, MAX);
    var html = '<table>' + shown.map(function (it) {
      return '<tr><td>' + esc(it.billcode) + '</td><td>' + esc(it.goods || '') + (it.company ? ' <span style="color:#999">(' + esc(it.company) + ')</span>' : '') + '</td></tr>';
    }).join('') + '</table>' + (items.length > MAX ? '<div class="more">… 其餘 ' + (items.length - MAX) + ' 個單號</div>' : '');
    return html;
  }
  OB.utils = { $: $, esc: esc, billcodeValid: billcodeValid, notify: notify, bufToB64: bufToB64, b64ToBuf: b64ToBuf, downloadRecord: downloadRecord, loadSheetJs: loadSheetJs, fmtTs: fmtTs, itemsPreviewHtml: itemsPreviewHtml };

  // ===================== 方言解析(xlsx → Item[]) =====================
  // 每個 Source 可帶自己的 parseBuffer(buf);預設/淘寶方言如下(簡/繁表頭皆可)
  // 保留為 OB.Parser 作為各 Source 的共用工具
  var GOODS_NAME_MAX = 60;
  function parseWorkbook(wb) {
    var XLSX = xlsxLib;
    if (!XLSX) throw new Error('SheetJS 未載入');
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows.length) return { items: [], errors: ['檔案沒有資料'] };
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
    if (cName < 0 || cBill < 0) return { items: [], errors: ['找不到「商品名稱/物流單號」欄位,請確認檔案格式'] };

    var groups = {}, errors = [];
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var bill = String(row[cBill] == null ? '' : row[cBill]).trim();
      var name = String(row[cName] == null ? '' : row[cName]).trim();
      if (!bill) continue;
      if (!billcodeValid(bill)) { errors.push('第' + (r + 1) + '行單號異常:' + bill); continue; }
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
        if (!XLSX) throw new Error('SheetJS 未載入');
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
              api.notify('⚠ 解析失敗(' + name + ':' + e.message + '),未存入');
              return Promise.resolve();
            }
            if (!res.items.length) {
              api.notify('⚠ 未解析到任何單號(' + name + ',' + (res.errors[0] || '檔案無資料') + '),未存入');
              return Promise.resolve();
            }
            lastCap = { name: base.name, ts: base.ts, source: base.source, items: res.items, errors: res.errors };
            return OB.Bridge.merge(res.items, { source: base.source, name: name }).then(function (m) {
              var msg = '✅ 已暫存(' + m.added + ' 個新單號';
              if (m.dup) msg += ',' + m.dup + ' 個已存在(略過)';
              msg += ',來源:' + base.source + ')';
              api.notify(msg);
            });
          }).catch(function (e) { api.notify('❌ 暫存失敗(' + name + ':' + e.message + ')'); });
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
          '#ob-src-fabs button{border:none;border-radius:24px;padding:12px 16px;font-size:14px;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.3);font-family:inherit;color:#fff}',
          '#ob-src-fab{background:#ff5000}#ob-src-fab:hover{background:#ff6a26}#ob-src-fab:disabled{background:#bbb;cursor:default}',
          '#ob-src-mgr{background:#555}#ob-src-mgr:hover{background:#777}',
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
          '.ob-prev .more{padding:4px 6px;color:#999;font-style:italic}'
        ].join('');
        document.head.appendChild(st);

        var wrap = document.createElement('div');
        wrap.id = 'ob-src-fabs';
        var fab = document.createElement('button');
        fab.id = 'ob-src-fab';
        fab.textContent = '📦 導出並暫存';
        var mgrBtn = document.createElement('button');
        mgrBtn.id = 'ob-src-mgr';
        mgrBtn.textContent = '🗂 暫存';
        wrap.appendChild(fab);
        wrap.appendChild(mgrBtn);
        document.body.appendChild(wrap);

        // 淘寶專屬:點 FAB 自動操作「導出訂單」流程(其他 Source 可換成自己的觸發邏輯)
        fab.onclick = function () {
          fab.disabled = true;
          fab.textContent = '⏳ 導出中…';
          function fail(msg) { fab.disabled = false; fab.textContent = '📦 導出並暫存'; api.notify('❌ ' + msg); }
          try {
            Array.prototype.slice.call(document.querySelectorAll('.ant-tooltip button, .ant-popover button, [class*=tooltip] button'))
              .forEach(function (b) { if (b.textContent.indexOf('知道了') > -1) b.click(); });
          } catch (e) { }
          setTimeout(function () {
            var btn = document.querySelector('[class*=exportBtn]');
            if (!btn) {
              var cands = Array.prototype.slice.call(document.querySelectorAll('div,span,button,a'));
              btn = cands.find(function (el) {
                var t = (el.textContent || '').replace(/\s/g, '');
                return el.childElementCount <= 2 && t.indexOf('导出订单') === 0 && t.length < 12;
              });
            }
            if (!btn) { fail('找不到「導出訂單」按鈕'); return; }
            btn.click();
            var tries = 0;
            var timer = setInterval(function () {
              tries++;
              var dl = Array.prototype.slice.call(document.querySelectorAll('.ant-modal button, .ant-modal a, .ant-modal [class*=btn]'))
                .find(function (b) {
                  var t = (b.textContent || '').replace(/\s/g, '');
                  return t.indexOf('下载订单') > -1 && (b.className || '').indexOf('disabled') === -1;
                });
              if (dl) {
                clearInterval(timer);
                dl.click();
                var w = 0;
                var t2 = setInterval(function () {
                  w++;
                  if (lastCap && lastCap.ts > Date.now() - 15000) { clearInterval(t2); fab.disabled = false; fab.textContent = '📦 導出並暫存'; }
                  else if (w >= 10) { clearInterval(t2); fail('未捕獲到下載檔案,請手動下載或改在管理選單匯入'); }
                }, 1000);
              } else if (tries >= 20) { clearInterval(timer); fail('等待「下載訂單」超時'); }
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
          if (!r) {
            admin.innerHTML =
              '<h4><span>暫存資料</span>' +
              '<button id="ob-src-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>' +
              '<div class="ob-rec" style="color:#999">尚無資料。觸發本站點導出或手動下載訂單後會自動出現。</div>';
          } else {
            admin.innerHTML =
              '<h4><span>暫存資料</span>' +
              '<button id="ob-src-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>' +
              '<div class="meta" style="padding:8px 14px 0;color:#999;font-size:12px">' +
              OB.utils.fmtTs(r.ts) + ' · 來源:' + (r.source || '?') +
              (r.items ? ' · ' + r.items.length + ' 個單號' : '') + '</div>' +
              (r.items ? '<div class="ob-prev" style="display:block;margin:8px 14px">' + OB.utils.itemsPreviewHtml(r.items) + '</div>' : '');
            var foot = document.createElement('div');
            foot.style.cssText = 'padding:8px 14px;text-align:right;border-top:1px solid #eee';
            foot.innerHTML =
              '<button id="ob-src-admin-dl" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:8px">匯出 xlsx(備份)</button>' +
              '<button id="ob-src-admin-clear" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer">全部清除</button>';
            admin.appendChild(foot);
          }
          admin.style.display = 'block';
          $('#ob-src-admin-x', admin).onclick = function () { admin.style.display = 'none'; };
          var dlBtn = $('#ob-src-admin-dl');
          if (dlBtn) dlBtn.onclick = function () { if (r) OB.utils.downloadRecord(r); };
          var clBtn = $('#ob-src-admin-clear');
          if (clBtn) clBtn.onclick = function () {
            if (!confirm('確定清除所有暫存資料?')) return;
            OB.Bridge.clear().then(function () { renderAdmin({ activeId: '', records: [] }); });
          };
        }
        mgrBtn.onclick = function () {
          if (admin.style.display === 'block') { admin.style.display = 'none'; return; }
          OB.Bridge.getAll().then(renderAdmin);
        };      }
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
                onEachDone(it, ok, ok ? '' : (d ? (d.MsgText || '失敗') : '回應異常'));
              });
            } else {
              var msg = String(text).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 150);
              chunk.forEach(function (it) { onEachDone(it, false, msg || '回應異常'); });
            }
          })
          .catch(function (e) { chunk.forEach(function (it) { onEachDone(it, false, '網路錯誤:' + e.message); }); });
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
  OB.UI = {
    install: function (site) {
      var MAX_ROWS_PER_SUBMIT = 10;
      var emptyState = function () { return { items: [], existing: {}, parsed: false }; };
      var state = emptyState();
      var store = null; // 暫存記錄快取

      var STYLE = [
        '#ob-fab{position:fixed;right:24px;bottom:120px;z-index:99999;border:none;border-radius:50%;width:60px;height:60px;background:#1890ff;color:#fff;font-size:26px;cursor:pointer;box-shadow:0 4px 12px rgba(24,144,255,.4);transition:transform .15s}',
        '#ob-fab:hover{transform:scale(1.08)}',
        '#ob-fab-mgr{position:fixed;right:24px;bottom:196px;z-index:99999;border:none;border-radius:50%;width:44px;height:44px;background:#595959;color:#fff;font-size:18px;cursor:pointer;box-shadow:0 3px 8px rgba(0,0,0,.3)}',
        '#ob-panel{position:fixed;top:0;right:0;width:560px;max-height:100vh;overflow:auto;background:#fff;z-index:99998;box-shadow:-4px 0 20px rgba(0,0,0,.2);padding:20px;box-sizing:border-box}',
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
      st.textContent = STYLE;
      document.head.appendChild(st);

      var fab = document.createElement('button');
      fab.id = 'ob-fab';
      fab.title = '導入訂單 → ' + site.name;
      fab.textContent = '📦';
      var fabMgr = document.createElement('button');
      fabMgr.id = 'ob-fab-mgr';
      fabMgr.title = '暫存管理';
      fabMgr.textContent = '🗂';
      document.body.appendChild(fab);
      document.body.appendChild(fabMgr);

      // ---------- 預覽面板 ----------
      var panel = document.createElement('div');
      panel.id = 'ob-panel';
      panel.style.display = 'none';
      panel.innerHTML =
        '<span class="ob-close" data-ob="close">✕</span>' +
        '<h3>導入訂單 → ' + esc(site.name) + ' 預報</h3>' +
        '<div id="ob-bridge-status"></div>' +
        '<div id="ob-bridge-btns"><button data-ob="auto">自動(暫存+現有)</button><button data-ob="pick-file">選擇檔案(備用)</button><button data-ob="reload-bridge">重新讀取暫存</button></div>' +
        '<div id="ob-status">尚未讀取資料</div>' +
        '<div><label><input type="checkbox" data-ob="edit-empty" checked> 已無品名 → 補品名</label>' +
        '<label style="margin-left:12px"><input type="checkbox" data-ob="edit-fill"> 已有品名 → 覆寫更新</label></div>' +
        '<table id="ob-table" style="display:none"><thead><tr>' +
        '<th class="c-bill">單號</th><th>物流公司</th><th class="c-goods">品名(goods,≤60字)</th><th>處理方式</th><th>結果</th>' +
        '</tr></thead><tbody id="ob-tbody"></tbody></table>' +
        '<div id="ob-result"></div>' +
        '<div style="margin-top:12px"><button class="primary" data-ob="submit" disabled>開始提交</button>' +
        '<button data-ob="refresh">重新分析</button></div>';
      document.body.appendChild(panel);

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
          if (!r) {
            admin.innerHTML =
              '<h4><span>暫存資料</span>' +
              '<button id="ob-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>' +
              '<div class="ob-rec" style="color:#999">尚無暫存資料。請先在來源站點導出訂單。</div>';
          } else {
            admin.innerHTML =
              '<h4><span>暫存資料</span>' +
              '<button id="ob-admin-x" style="border:none;background:none;font-size:18px;cursor:pointer">✕</button></h4>' +
              '<div class="meta" style="padding:8px 14px 0;color:#999;font-size:12px">' +
              OB.utils.fmtTs(r.ts) + ' · 來源:' + (r.source || '?') +
              (r.items ? ' · ' + r.items.length + ' 個單號' : '') + '</div>' +
              (r.items ? '<div class="ob-prev" style="display:block;margin:8px 14px">' + OB.utils.itemsPreviewHtml(r.items) + '</div>' : '');
            var foot2 = document.createElement('div');
            foot2.style.cssText = 'padding:8px 14px;text-align:right;border-top:1px solid #eee';
            foot2.innerHTML =
              '<button id="ob-admin-dl" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer;margin-right:8px">匯出 xlsx(備份)</button>' +
              '<button id="ob-admin-clear" style="padding:5px 14px;border:1px solid #d9d9d9;border-radius:4px;background:#fff;cursor:pointer">全部清除</button>';
            admin.appendChild(foot2);
            $('#ob-admin-dl', admin).onclick = function () { OB.utils.downloadRecord(r); };
            $('#ob-admin-clear', admin).onclick = function () {
              if (!confirm('確定清除所有暫存資料?')) return;
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
          setStatus('⏳ 解析舊版暫存(原始檔)…');
          var buf = OB.utils.b64ToBuf(rec.b64);
          await OB.utils.loadSheetJs();
          var pbuf = (OB.Sources[rec.source] && OB.Sources[rec.source].parseBuffer) || OB.Sources.taobao.parseBuffer;
          var res;
          try { res = pbuf(buf); items = res.items; errors = res.errors; }
          catch (e) { items = []; errors = ['舊版暫存解析失敗:' + e.message]; }
          if (items.length) {
            rec.items = items; rec.errors = errors; delete rec.b64;
            await OB.Bridge.save(store); // 一次性遷移:之後直接存 items
          }
        }
        else { items = []; errors = ['暫存記錄無資料']; }
        state.items = items;
        state.parsed = true;
        state.rec = rec;
        await OB.Bridge.setActive(rec.id);
        setStatus('✅ 已讀取「' + esc(rec.name) + '」(來源:' + (rec.source || '?') +
          ',解析出 <b>' + items.length + '</b> 個單號' +
          (errors.length ? '<br><span class="ob-err">⚠ ' + errors.map(esc).join('<br>') + '</span>' : ''));
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
          setStatus('⏳ 解析 ' + f.name + '…');
          try {
            var buf = await f.arrayBuffer();
            await OB.utils.loadSheetJs();
            var res = OB.Sources.taobao.parseBuffer(buf); // 手動匯入固定為淘寶 xlsx 格式
            state.items = res.items;
            state.parsed = true;
            state.rec = { id: 'file', name: f.name, size: f.size, ts: Date.now(), source: '手動匯入', items: res.items, errors: res.errors };
            setStatus('✅ ' + f.name + ' → <b>' + res.items.length + '</b> 個單號' +
              (res.errors.length ? '<br><span class="ob-err">⚠ ' + res.errors.map(esc).join('<br>') + '</span>' : ''));
          } catch (e) {
            setStatus('❌ 解析失敗:' + esc(e.message));
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
        if (!state.parsed) { setStatus('⚠ 請先讀取檔案(點「自動(暫存+現有)」或選擇檔案)'); return; }
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
          if (row.action === 'new') { badge = '<span class="ob-badge ob-b-new">新預報</span>'; count++; }
          else if (row.action === 'edit-new') { badge = '<span class="ob-badge ' + (doEdit ? 'ob-b-edit' : 'ob-b-editnew') + '">' + (doEdit ? '補品名' : '需勾選') + '</span>'; if (doEdit) count++; }
          else { badge = '<span class="ob-badge ' + (doEdit ? 'ob-b-edit' : 'ob-b-editnew') + '">' + (doEdit ? '更新品名' : '需勾選') + '</span>'; if (doEdit) count++; }
          var tr = document.createElement('tr');
          tr.setAttribute('data-idx', idx);
          tr.innerHTML =
            '<td class="c-bill">' + esc(row.item.billcode) + '</td>' +
            '<td>' + esc(row.item.company || '') + '</td>' +
            '<td class="c-goods"><input data-f="goods" value="' + esc(row.item.goods) + '"></td>' +
            '<td>' + badge +
            (row.existing && row.existing.goods ? '<div style="color:#999;font-size:11px">現有:' + esc(row.existing.goods) + '</div>' : '') +
            '</td><td class="c-res">-</td>';
          tbody.appendChild(tr);
        });
        $('#ob-table').style.display = 'block';
        $('[data-ob="submit"]').disabled = !count;
        state.pending = rows;
        state.planCount = count;
        setStatus('📋 共 ' + rows.length + ' 個單號 → <b>' + nNew + '</b> 新預報 / <b>' + nEN + '</b> 需補品名 / <b>' + nEF + '</b> 可更新品名。' +
          '本次將處理 <b>' + count + '</b> 筆。');
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
          var goods = readGoodsFromRow(tr);
          if (!goods) goods = row.item.goods;
          var act = 'skip';
          if (row.action === 'new') act = 'new';
          else if (row.action === 'edit-new' && doEditEmpty) act = 'edit';
          else if (row.action === 'edit-fill' && doEditFill) act = 'edit';
          if (act !== 'skip') jobs.push({ act: act, row: row, tr: tr, goods: goods });
        });
        if (!jobs.length) { setStatus('⚠ 沒有要處理的項'); return; }
        $('[data-ob="submit"]').disabled = true;
        var resultDiv = $('#ob-result');
        resultDiv.innerHTML = '<div class="ob-ok">⏳ 提交中…</div>';
        var okN = 0, failN = 0;
        // 1) 新預報:分組(每 10 筆)
        var newJobs = jobs.filter(function (j) { return j.act === 'new'; });
        var editJobs = jobs.filter(function (j) { return j.act === 'edit'; });
        var chunks = [];
        for (var i = 0; i < newJobs.length; i += MAX_ROWS_PER_SUBMIT) chunks.push(newJobs.slice(i, i + MAX_ROWS_PER_SUBMIT));
        function markJob(j, ok, msg) {
          var resTd = j.tr.querySelector('.c-res');
          if (ok) { okN++; resTd.innerHTML = '<span class="ob-ok">✓ ' + esc(j.act === 'new' ? '已預報' : '已更新') + '</span>'; }
          else { failN++; resTd.innerHTML = '<span class="ob-err">✗ ' + esc(msg || '失敗') + '</span>'; }
        }
        function done() {
          $('[data-ob="submit"]').disabled = false;
          resultDiv.innerHTML = '<div class="' + (failN ? 'ob-err' : 'ob-ok') + '">' +
            (failN ? '⚠ 完成:成功 ' + okN + ' / 失敗 ' + failN : '✅ 全部成功(' + okN + ' 筆)。可到「查看到貨情況」核對。') + '</div>';
          OB.utils.notify('導入完成:成功 ' + okN + ',失敗 ' + failN);
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
        if (ob === 'close') panel.style.display = 'none';
        else if (ob === 'reload-bridge') {
          autoLoadBridge().then(function (rec) {
            if (!rec) { setStatus('⚠ 暫存區沒有檔案,請先在來源站點導出'); return; }
            runAnalysis();
          }).catch(function (e) { setStatus('❌ ' + esc(e.message)); });
        }
        else if (ob === 'auto') {
          autoLoadBridge().then(function (rec) {
            if (!rec) { setStatus('⚠ 暫存區沒有檔案,請先在來源站點導出'); return; }
            runAnalysis();
          }).catch(function (e) { setStatus('❌ ' + esc(e.message)); });
        }
        else if (ob === 'pick-file') pickFile();
        else if (ob === 'refresh') runAnalysis();
        else if (ob === 'submit') submit();
      });
      $('[data-ob="edit-empty"]').addEventListener('change', function () { if (state.parsed) runAnalysis(); });
      $('[data-ob="edit-fill"]').addEventListener('change', function () { if (state.parsed) runAnalysis(); });

      fab.onclick = function () {
        panel.style.display = 'block';
        state = emptyState();
        refreshBridgeUI();
        autoLoadBridge().then(function (rec) {
          if (!rec) { setStatus('⚠ 暫存區沒有檔案,請先在來源站點導出'); return; }
          runAnalysis();
        }).catch(function (e) { setStatus('❌ 自動載入失敗:' + esc(e.message)); });
      };
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
    if (site) OB.UI.install(site);
  })();
})();
