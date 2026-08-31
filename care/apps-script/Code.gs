/**
 * 台中長照接送速查 — Google 試算表後端
 * ------------------------------------------------------------
 * 貼進 Apps Script → 改下面的 TOKEN 與 SS_ID → 部署為網頁應用程式
 * （執行身分：我自己　／　誰可以存取：所有人）
 *
 * 自動建立三張工作表：
 *   「紀錄」  ─ 每家單位的狀態、星號、筆記（全家共用一份）
 *   「使用紀錄」─ 誰、哪台裝置、什麼時候、做了什麼
 *   「裝置」  ─ 有哪幾支手機在用、各用了幾次
 */

/** 家人共用的通關碼。跟網頁設定裡填的要一樣。 */
var TOKEN = '1234';

/**
 * 試算表 ID。
 * 從「試算表 → 擴充功能 → Apps Script」開的可留空；
 * 獨立專案請貼上網址中間那段：
 * https://docs.google.com/spreadsheets/d/【這一段】/edit
 */
var SS_ID = '';

/**
 * ── 版本控制（跟版系統）───────────────────────────────
 * 改好新版網頁、推上 GitHub 後，把 APP_VERSION 和 MIN_VERSION
 * 一起 +1，再「部署 → 管理部署作業 → 編輯 → 新版本」。
 * 家人手機下次連線就會被擋下來，強制更新到新版才能繼續用。
 */
var APP_VERSION = 2;   // 目前最新的網頁版本
var MIN_VERSION = 2;   // 低於這個版本一律強制更新

/** 使用紀錄最多留幾列，超過就從最舊的刪。 */
var LOG_KEEP = 3000;

var REC_SHEET = '紀錄', LOG_SHEET = '使用紀錄', DEV_SHEET = '裝置';
var REC_HEAD = ['序號', '單位名稱', '狀態', '星號', '筆記', '最後更新者', '裝置', '最後更新時間', '_ts'];
var LOG_HEAD = ['時間', '使用者', '裝置', '動作', '內容', '序號', '單位名稱'];
var DEV_HEAD = ['裝置代碼', '暱稱', '首次使用', '最後使用', '使用次數'];

var ST_TEXT = { none: '未聯絡', called: '已聯絡', yes: '可接送', hold: '待確認', no: '不接' };

/* ============================ 入口 ============================ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.t !== TOKEN) throw new Error('通關碼不正確');
    var v = Number(p.v) || 0;
    if (v < MIN_VERSION) {
      return reply({
        ok: false, needUpdate: true, minVersion: MIN_VERSION, appVersion: APP_VERSION,
        error: '網頁版本過舊（你的 v' + v + '，需要 v' + MIN_VERSION + '），請重新整理更新'
      }, p.callback);
    }
    touchDevice(p.dev, p.u);
    switch (p.op) {
      case 'ping': out = { ok: true, name: ss().getName(), now: Date.now() }; break;
      case 'pull': out = opPull(p); break;
      case 'save': out = opSave(p); break;
      case 'log':  out = opLog(p);  break;
      default: throw new Error('不認得的動作：' + p.op);
    }
    out.minVersion = MIN_VERSION;
    out.appVersion = APP_VERSION;
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return reply(out, p.callback);
}

function doPost(e) { return doGet(e); }

function reply(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ 動作 ============================ */

function opPull(p) {
  var sh = recSheet();
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, REC_HEAD.length).getValues() : [];
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
      dev: String(r[6] || ''),
      ts: Number(r[8]) || 0
    });
  }
  return { ok: true, records: records, log: readLog(Math.min(Number(p.n) || 60, 300)), now: Date.now() };
}

function opSave(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recSheet();
    var id = Number(p.id);
    if (!id) throw new Error('缺少序號');
    var ts = Number(p.ts) || Date.now();
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
    var row = 0;
    for (var i = 0; i < ids.length; i++) {
      if (Number(ids[i][0]) === id) { row = i + 2; break; }
    }
    if (row) {
      var old = Number(sh.getRange(row, 9).getValue()) || 0;
      if (old > ts) {
        var cur = sh.getRange(row, 1, 1, REC_HEAD.length).getValues()[0];
        return {
          ok: true, stale: true,
          record: { id: id, st: String(cur[2] || 'none'), star: cur[3] ? 1 : 0, note: String(cur[4] || ''),
                    by: String(cur[5] || ''), dev: String(cur[6] || ''), ts: old }
        };
      }
    } else {
      row = sh.getLastRow() + 1;
    }
    sh.getRange(row, 1, 1, REC_HEAD.length).setValues([[
      id, String(p.nm || ''), String(p.st || 'none'),
      (p.star === '1' || p.star === 1 || p.star === true) ? '★' : '',
      String(p.note || ''), String(p.u || ''), String(p.dev || ''), new Date(ts), ts
    ]]);
    sh.getRange(row, 8).setNumberFormat('yyyy/mm/dd hh:mm');
    return { ok: true, id: id, ts: ts };
  } finally {
    lock.releaseLock();
  }
}

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
    var rows = items.slice(0, 40).map(function (it) {
      return [
        new Date(Number(it.at) || Date.now()),
        String(it.user || p.u || ''),
        String(it.dev || p.dev || ''),
        String(it.act || ''), String(it.txt || ''),
        it.id ? Number(it.id) : '', String(it.nm || '')
      ];
    });
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, rows.length, LOG_HEAD.length).setValues(rows);
    sh.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
    var extra = sh.getLastRow() - 1 - LOG_KEEP;
    if (extra > 0) sh.deleteRows(2, extra);
    return { ok: true, n: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function readLog(limit) {
  var sh = logSheet(), last = sh.getLastRow();
  if (last < 2) return [];
  var n = Math.min(limit, last - 1);
  var rows = sh.getRange(last - n + 1, 1, n, LOG_HEAD.length).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var r = rows[i];
    out.push({
      at: r[0] instanceof Date ? r[0].getTime() : Number(r[0]) || 0,
      user: String(r[1] || ''), dev: String(r[2] || ''), act: String(r[3] || ''),
      txt: String(r[4] || ''), id: r[5] || '', nm: String(r[6] || '')
    });
  }
  return out;
}

/** 記下這台裝置：第一次見到就新增一列，之後更新最後使用時間與次數 */
function touchDevice(devId, nick) {
  if (!devId) return;
  var sh = devSheet();
  var last = sh.getLastRow();
  var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var row = 0;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(devId)) { row = i + 2; break; }
  }
  var now = new Date();
  if (row) {
    if (nick) sh.getRange(row, 2).setValue(nick);
    sh.getRange(row, 4).setValue(now).setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange(row, 5).setValue((Number(sh.getRange(row, 5).getValue()) || 0) + 1);
  } else {
    row = last + 1;
    sh.getRange(row, 1, 1, DEV_HEAD.length).setValues([[String(devId), String(nick || ''), now, now, 1]]);
    sh.getRange(row, 3, 1, 2).setNumberFormat('yyyy/mm/dd hh:mm');
  }
}

/* ============================ 工具 ============================ */

function ss() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('找不到試算表：這是獨立的 Apps Script 專案，請在程式碼上方的 SS_ID 填入試算表 ID');
  return active;
}

/**
 * 取得工作表，並確保標題列正確。
 * 空表直接覆寫標題；已有資料則只在後面補上缺少的欄位，不動既有資料。
 */
function sheetWithHead(name, head) {
  var book = ss(), s = book.getSheetByName(name);
  if (!s) {
    s = book.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#DCEDE7');
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).createFilter();
    return s;
  }
  if (s.getMaxColumns() < head.length) s.insertColumnsAfter(s.getMaxColumns(), head.length - s.getMaxColumns());
  var cur = s.getRange(1, 1, 1, head.length).getValues()[0];
  var same = true;
  for (var i = 0; i < head.length; i++) if (String(cur[i]) !== head[i]) { same = false; break; }
  if (!same) {
    if (s.getLastRow() <= 1) {
      s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#DCEDE7');
      s.setFrozenRows(1);
    } else {
      for (var j = 0; j < head.length; j++) {
        if (String(cur[j]) !== head[j] && !String(cur[j])) s.getRange(1, j + 1).setValue(head[j]).setFontWeight('bold');
      }
    }
  }
  return s;
}

function recSheet() {
  var s = sheetWithHead(REC_SHEET, REC_HEAD);
  if (s.getMaxColumns() >= 9) s.hideColumns(9); // _ts 是給程式用的
  return s;
}
function logSheet() { return sheetWithHead(LOG_SHEET, LOG_HEAD); }
function devSheet() { return sheetWithHead(DEV_SHEET, DEV_HEAD); }

/** 按上方「執行」選這個，確認有沒有接到試算表 */
function 測試設定() {
  var s = ss();
  recSheet(); logSheet(); devSheet();
  Logger.log('已接上試算表：' + s.getName());
  Logger.log('資料表就緒：' + REC_SHEET + '、' + LOG_SHEET + '、' + DEV_SHEET);
  Logger.log('目前版本設定：APP_VERSION=' + APP_VERSION + '，MIN_VERSION=' + MIN_VERSION);
  return s.getName();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('長照接送')
    .addItem('建立／檢查資料表', '測試設定')
    .addItem('狀態代碼說明', 'showLegend')
    .addToUi();
}

function showLegend() {
  var t = Object.keys(ST_TEXT).map(function (k) { return k + ' = ' + ST_TEXT[k]; }).join('\n');
  SpreadsheetApp.getUi().alert('「狀態」欄位代碼\n\n' + t);
}
