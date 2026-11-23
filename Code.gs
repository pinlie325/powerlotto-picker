/**
 * 威力彩選號機：雲端選號紀錄（Google Apps Script）
 * 用法：在 Google 試算表 → 擴充功能 → Apps Script，貼上整段程式，
 *       把下面的 SECRET 改成你自己的同步密碼，再部署成「網頁應用程式」。
 */
const SECRET = '請改成你自己的密碼';
const SHEET_NAME = '選號紀錄';
const HEADER = ['id', '儲存時間', '號1', '號2', '號3', '號4', '號5', '號6', '第二區', '熱度'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A:B').setNumberFormat('@'); // id、時間存成純文字，避免被轉成日期
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function list_() {
  const v = sheet_().getDataRange().getValues();
  v.shift();
  const tz = Session.getScriptTimeZone();
  return v.filter(r => r[0] !== '').map(r => ({
    id: String(r[0]),
    t: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd HH:mm') : String(r[1]),
    n: r.slice(2, 8).map(Number),
    z: Number(r[8]),
    h: r[9] === '' ? null : Number(r[9])
  }));
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return out_({ ok: true, msg: '威力彩選號紀錄雲端已啟用' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: '資料格式錯誤' }); }
  if (b.key !== SECRET) return out_({ ok: false, error: '同步密碼錯誤' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (b.action === 'add') {
      const have = new Set(list_().map(i => i.id));
      (b.items || []).forEach(i => {
        if (have.has(String(i.id))) return;
        const n = (i.n || []).map(Number);
        if (n.length !== 6) return;
        sh.appendRow([String(i.id), String(i.t || '')].concat(n, [Number(i.z), i.h == null ? '' : Number(i.h)]));
      });
    } else if (b.action === 'delete') {
      const ids = new Set((b.ids || []).map(String));
      const v = sh.getDataRange().getValues();
      for (let r = v.length - 1; r >= 1; r--) {
        if (ids.has(String(v[r][0]))) sh.deleteRow(r + 1);
      }
    } else if (b.action !== 'list') {
      return out_({ ok: false, error: '未知的動作' });
    }
    return out_({ ok: true, items: list_() });
  } finally {
    lock.releaseLock();
  }
}
