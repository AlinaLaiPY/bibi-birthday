/**
 * 台中長照接送速查 — Google 試算表後端
 * ------------------------------------------------------------
 * 用法：在 Google 試算表 → 擴充功能 → Apps Script，把這整份貼進去，
 *      改下面的 TOKEN，再「部署 → 新增部署作業 → 網頁應用程式」。
 *      執行身分：我自己　／　誰可以存取：所有人
 *      部署後複製那個 /exec 網址，貼進網頁的「設定」。
 *
 * 資料表會自動建立，不用自己開：
 *   「紀錄」    ─ 每家單位的狀態、星號、筆記（家人共用同一份）
 *   「使用紀錄」─ 誰、什麼時候、搜尋了什麼、打給誰
 */

/** 家人共用的通關碼。跟網頁設定裡填的要一樣。請改成自己的。 */
var TOKEN = 'bibi2026';

/** 使用紀錄最多留幾列，超過就從最舊的刪。 */
var LOG_KEEP = 3000;

var REC_SHEET = '紀錄';
var LOG_SHEET = '使用紀錄';
var REC_HEAD = ['序號', '單位名稱', '狀態', '星號', '筆記', '最後更新者', '最後更新時間', '_ts'];
var LOG_HEAD = ['時間', '使用者', '動作', '內容', '序號', '單位名稱'];

var ST_TEXT = { none: '未聯絡', called: '已聯絡', yes: '可接送', hold: '待確認', no: '不接' };

/* ============================ 入口 ============================ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.t !== TOKEN) throw new Error('通關碼不正確');
    switch (p.op) {
      case 'ping': out = { ok: true, name: ss().getName(), now: Date.now() }; break;
      case 'pull': out = opPull(p); break;
      case 'save': out = opSave(p); break;
      case 'log':  out = opLog(p);  break;
      default: throw new Error('不認得的動作：' + p.op);
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return reply(out, p.callback);
}

/** 同時支援 POST（保留彈性，目前網頁走 GET/JSONP） */
function doPost(e) {
  return doGet(e);
}

function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================ 動作 ============================ */

/** 抓全部紀錄 + 最近的使用紀錄 */
function opPull(p) {
  var sh = recSheet();
  var rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, REC_HEAD.length).getValues()
    : [];
  var records = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] && r[0] !== 0) continue;
    records.push({
      id: Number(r[0]),
      st: String(r[2] || 'none'),
      star: r[3] ? 1 : 0,
      note: String(r[4] || ''),
      by: String(r[5] || ''),
      ts: Number(r[7]) || 0
    });
  }
  var limit = Math.min(Number(p.n) || 60, 300);
  return { ok: true, records: records, log: readLog(limit), now: Date.now() };
}

/** 寫入或更新一家單位的紀錄（後寫的贏，用 _ts 比大小） */
function opSave(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recSheet();
    var id = Number(p.id);
    if (!id) throw new Error('缺少序號');
    var ts = Number(p.ts) || Date.now();
    var ids = sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      : [];
    var row = 0;
    for (var i = 0; i < ids.length; i++) {
      if (Number(ids[i][0]) === id) { row = i + 2; break; }
    }
    if (row) {
      var old = Number(sh.getRange(row, 8).getValue()) || 0;
      if (old > ts) {
        // 試算表上的比較新，不覆蓋，把新的回傳給手機
        var cur = sh.getRange(row, 1, 1, REC_HEAD.length).getValues()[0];
        return {
          ok: true, stale: true,
          record: { id: id, st: String(cur[2] || 'none'), star: cur[3] ? 1 : 0,
                    note: String(cur[4] || ''), by: String(cur[5] || ''), ts: old }
        };
      }
    } else {
      row = sh.getLastRow() + 1;
    }
    var st = String(p.st || 'none');
    sh.getRange(row, 1, 1, REC_HEAD.length).setValues([[
      id,
      String(p.nm || ''),
      st,
      p.star === '1' || p.star === 1 || p.star === true ? '★' : '',
      String(p.note || ''),
      String(p.user || ''),
      new Date(ts),
      ts
    ]]);
    sh.getRange(row, 7).setNumberFormat('yyyy/mm/dd hh:mm');
    return { ok: true, id: id, ts: ts };
  } finally {
    lock.releaseLock();
  }
}

/** 寫使用紀錄，可一次送多筆：d = JSON 陣列 */
function opLog(p) {
  var items = [];
  if (p.d) {
    try { items = JSON.parse(p.d); } catch (err) { throw new Error('使用紀錄格式錯誤'); }
  } else if (p.act) {
    items = [{ act: p.act, txt: p.txt, id: p.id, nm: p.nm, at: Number(p.at) || Date.now() }];
  }
  if (!items.length) return { ok: true, n: 0 };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = logSheet();
    var user = String(p.user || '');
    var rows = items.slice(0, 40).map(function (it) {
      return [
        new Date(Number(it.at) || Date.now()),
        String(it.user || user),
        String(it.act || ''),
        String(it.txt || ''),
        it.id ? Number(it.id) : '',
        String(it.nm || '')
      ];
    });
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, LOG_HEAD.length).setValues(rows);
    sh.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    trimLog(sh);
    return { ok: true, n: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function readLog(limit) {
  var sh = logSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = Math.min(limit, last - 1);
  var rows = sh.getRange(last - n + 1, 1, n, LOG_HEAD.length).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    var at = r[0] instanceof Date ? r[0].getTime() : Number(r[0]) || 0;
    out.push({ at: at, user: String(r[1] || ''), act: String(r[2] || ''),
               txt: String(r[3] || ''), id: r[4] || '', nm: String(r[5] || '') });
  }
  return out;
}

function trimLog(sh) {
  var extra = sh.getLastRow() - 1 - LOG_KEEP;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* ============================ 工具 ============================ */

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheetWithHead(name, head) {
  var s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#DCEDE7');
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).createFilter();
  }
  return s;
}

function recSheet() {
  var s = sheetWithHead(REC_SHEET, REC_HEAD);
  if (s.getMaxColumns() >= 8) s.hideColumns(8); // _ts 是給程式看的
  return s;
}

function logSheet() { return sheetWithHead(LOG_SHEET, LOG_HEAD); }

/** 在試算表裡加一個選單，方便測試 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('長照接送')
    .addItem('建立／檢查資料表', 'setup')
    .addItem('狀態代碼說明', 'showLegend')
    .addToUi();
}

function setup() {
  recSheet(); logSheet();
  SpreadsheetApp.getUi().alert('資料表已就緒：「' + REC_SHEET + '」與「' + LOG_SHEET + '」。');
}

function showLegend() {
  var t = Object.keys(ST_TEXT).map(function (k) { return k + ' = ' + ST_TEXT[k]; }).join('\n');
  SpreadsheetApp.getUi().alert('「狀態」欄位代碼\n\n' + t);
}
