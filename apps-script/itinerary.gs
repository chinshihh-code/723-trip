/**
 * 805 寒假班遊｜行程討論後台（獨立專案，不影響原本的報名後台）
 * 資料寫在報名試算表的「行程」分頁（第一次儲存時自動建立）。
 *
 * 部署方式：script.google.com → 新專案 → 貼上本檔 → 部署 → 新增部署作業
 *   類型：網頁應用程式｜執行身分：我｜誰可以存取：所有人
 * 把部署後的網址（…/exec）交給 Claude 填進網頁。
 */
const SHEET_ID = '1Z98JaVjBIy1NLi7MDcpORcnAdQelcWUetKz_U3B8m44';
const SHEET_NAME = '行程';
const ROOM_SHEET = '房號';
const PASSWORD = '805trip';   // ← 編輯密碼，請改成你們自己的（改完要「管理部署作業」→ 編輯 → 新版本）

function doGet() {
  return out(Object.assign(readAll(), { rooms: readRooms() }));
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: '資料格式錯誤' }); }

  if (b.action === 'check') return out({ ok: b.pw === PASSWORD });

  if (b.action === 'save') {
    if (b.pw !== PASSWORD) return out({ ok: false, error: '密碼錯誤' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const cur = readAll();
      if (String(b.version) !== String(cur.version)) {
        return out({ ok: false, conflict: true, error: '有其他人剛更新過行程，已載入最新版本，請再改一次。', data: cur });
      }
      writeAll(b.days || []);
      return out(Object.assign({ ok: true }, readAll()));
    } finally {
      lock.releaseLock();
    }
  }
  if (b.action === 'saveRooms') {
    if (b.pw !== PASSWORD) return out({ ok: false, error: '密碼錯誤' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rooms = readRooms();
      Object.keys(b.rooms || {}).forEach(function (k) {
        const v = String(b.rooms[k] || '').trim();
        if (v) rooms[k] = v; else delete rooms[k];
      });
      writeRooms(rooms);
      return out({ ok: true, rooms: readRooms() });
    } finally {
      lock.releaseLock();
    }
  }
  return out({ ok: false, error: '未知的動作' });
}

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function readAll() {
  const props = PropertiesService.getScriptProperties();
  const version = props.getProperty('v') || '0';
  const sh = sheet_();
  if (sh.getLastRow() < 2) return { days: null, version: version, updated: props.getProperty('t') || '' };
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues();
  const days = [], idx = {};
  vals.forEach(function (r) {
    const date = String(r[0]).trim();
    if (!date) return;
    if (!(date in idx)) { idx[date] = days.length; days.push({ date: date, items: [] }); }
    if (r[1] || r[2] || r[3] || r[4]) days[idx[date]].items.push({ time: r[1], title: r[2], place: r[3], content: r[4] });
  });
  return { days: days, version: version, updated: props.getProperty('t') || '' };
}

function writeAll(days) {
  const sh = sheet_();
  const rows = [['日期', '時間', '行程', '地點', '內容']];
  days.forEach(function (d) {
    const date = String(d.date || '').trim();
    if (!date) return;
    if (!d.items || !d.items.length) rows.push([date, '', '', '', '']);
    (d.items || []).forEach(function (it) {
      rows.push([date, it.time || '', it.title || '', it.place || '', it.content || '']);
    });
  });
  sh.clear();
  const rg = sh.getRange(1, 1, rows.length, 5);
  rg.setNumberFormat('@');            // 全部當文字，避免日期被自動轉換
  rg.setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('v', String(Number(props.getProperty('v') || '0') + 1));
  props.setProperty('t', Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm'));
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ===== 房號（座號 → 房號），寫在「房號」分頁，也可以直接在試算表改 ===== */
function roomSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(ROOM_SHEET) || ss.insertSheet(ROOM_SHEET);
}
function readRooms() {
  const sh = roomSheet_(), map = {};
  if (sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) {
    const seat = String(r[0]).trim(), no = String(r[1]).trim();
    if (seat && no) map[seat] = no;
  });
  return map;
}
function writeRooms(map) {
  const sh = roomSheet_();
  const rows = [['座號', '房號']];
  Object.keys(map).sort(function (a, b) { return Number(a) - Number(b); }).forEach(function (k) { rows.push([k, map[k]]); });
  sh.clear();
  const rg = sh.getRange(1, 1, rows.length, 2);
  rg.setNumberFormat('@');
  rg.setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold');
}
