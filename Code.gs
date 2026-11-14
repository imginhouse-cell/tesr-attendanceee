/**
 * ระบบเช็คชื่อออนไลน์: Google Apps Script backend
 * บันทึกการเช็คชื่อทุกครั้งลง Google Sheets (ชีต Attendance) พร้อมเก็บรูปลายเซ็นไว้ใน Google Drive
 *
 * ติดตั้ง: ดู README.md ในโปรเจกต์
 * รหัสผู้ดูแลตั้งค่าที่ Project Settings › Script Properties › ADMIN_KEY (ไม่ได้อยู่ในโค้ด จึงขึ้น GitHub ได้ปลอดภัย)
 */

const SAVE_SIGNATURE_TO_DRIVE = true;               // เก็บไฟล์ลายเซ็น .png ไว้ใน Drive แล้วใส่ลิงก์ในชีต
const SIGNATURE_FOLDER = 'Attendance Signatures';   // ชื่อโฟลเดอร์ใน Drive

const R_HEAD = ['id', 'ชื่อ-นามสกุล', 'หน่วยงาน', 'ชื่อภาษาอังกฤษ'];
const A_HEAD = ['วันที่', 'id', 'ชื่อ-นามสกุล', 'หน่วยงาน', 'เวลาเช็คอิน', 'สถานะ', 'สาย (นาที)', 'บันทึกเมื่อ (ISO)', 'ลิงก์ลายเซ็น', 'ลายเซ็น (ข้อมูลภาพ)'];
const S_HEAD = ['key', 'value'];

/** รัน 1 ครั้งหลังวางโค้ด: สร้างชีต, โฟลเดอร์ลายเซ็น และขออนุญาตสิทธิ์ */
function setup() {
  sheet_('Roster', R_HEAD);
  sheet_('Attendance', A_HEAD);
  const s = sheet_('Settings', S_HEAD);
  if (s.getLastRow() < 2) s.getRange(2, 1, 3, 2).setValues([['title', 'Claude AI'], ['start', '09:00'], ['grace', 15]]);
  if (SAVE_SIGNATURE_TO_DRIVE) folder_();
  if (!adminKey_()) Logger.log('⚠️ ยังไม่ได้ตั้ง ADMIN_KEY ใน Project Settings › Script Properties');
  Logger.log('ติดตั้งเรียบร้อย ไป Deploy เป็น Web app ต่อได้เลย');
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'state') return out_(state_(String(p.date || today_())));
    return out_({ ok: true, hello: 'TESR attendance API', admin_key_set: !!adminKey_() });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'checkin') return out_(checkin_(b));
    if (b.action === 'uncheck') { auth_(b); return out_(uncheck_(b)); }
    if (b.action === 'saveClass') { auth_(b); return out_(saveClass_(b.cls || {})); }
    return out_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- actions ---------- */

function state_(date) {
  const cls = settings_();
  const rs = sheet_('Roster', R_HEAD).getDataRange().getValues().slice(1);
  cls.people = rs.filter(r => r[0] && r[1]).map(r => ({ id: String(r[0]), name: String(r[1]), org: String(r[2] || ''), en: String(r[3] || '') }));
  const checks = {};
  sheet_('Attendance', A_HEAD).getDataRange().getValues().slice(1).forEach(r => {
    if (fmtDate_(r[0]) === date) checks[String(r[1])] = { t: String(r[7]), name: String(r[2]), org: String(r[3]), sig: String(r[9] || '') };
  });
  return { ok: true, cls, checks };
}

function checkin_(b) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(b.date || '')) || !b.pid) throw new Error('missing_fields');
  const s = sheet_('Attendance', A_HEAD);
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (fmtDate_(v[i][0]) === b.date && String(v[i][1]) === String(b.pid)) return { ok: true, already: true };
  }
  const r = b.rec || {};
  const sig = /^data:image\/png;base64,/.test(String(r.sig || '')) ? String(r.sig) : '';
  let link = '';
  if (SAVE_SIGNATURE_TO_DRIVE && sig) {
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(sig.split(',')[1]), 'image/png', `${b.date}_${safe_(r.name || b.pid)}.png`);
      link = folder_().createFile(blob).getUrl();
    } catch (err) { link = ''; } // ถ้าบันทึกไฟล์ไม่ได้ ยังเช็คชื่อได้ตามปกติ
  }
  s.appendRow([b.date, String(b.pid), r.name || '', r.org || '', r.time || '', r.status || '',
    r.late === '' || r.late == null ? '' : Number(r.late), r.t || new Date().toISOString(), link, sig.slice(0, 49000)]);
  return { ok: true };
}

function uncheck_(b) {
  const s = sheet_('Attendance', A_HEAD);
  const v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (fmtDate_(v[i][0]) === b.date && String(v[i][1]) === String(b.pid)) s.deleteRow(i + 1);
  }
  return { ok: true };
}

function saveClass_(cls) {
  const st = sheet_('Settings', S_HEAD);
  st.getRange(2, 1, Math.max(st.getLastRow(), 4), 2).clearContent();
  st.getRange(2, 1, 3, 2).setValues([['title', cls.title || 'Claude AI'], ['start', cls.start || '09:00'], ['grace', Number(cls.grace) || 0]]);
  const rs = sheet_('Roster', R_HEAD);
  if (rs.getLastRow() > 1) rs.getRange(2, 1, rs.getLastRow() - 1, R_HEAD.length).clearContent();
  const people = (cls.people || []).filter(p => p && p.name);
  if (people.length) rs.getRange(2, 1, people.length, R_HEAD.length).setValues(people.map(p => [String(p.id), p.name, p.org || '', p.en || '']));
  return { ok: true, count: people.length };
}

/* ---------- helpers ---------- */

function adminKey_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || ''; }

function auth_(b) {
  const k = adminKey_();
  if (!k || String(b.key || '') !== k) throw new Error('bad_key');
}

function settings_() {
  const o = { title: 'Claude AI', start: '09:00', grace: 15 };
  sheet_('Settings', S_HEAD).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0])] = r[1]; });
  if (o.start instanceof Date) o.start = Utilities.formatDate(o.start, Session.getScriptTimeZone(), 'HH:mm');
  o.start = String(o.start); o.title = String(o.title); o.grace = Number(o.grace) || 0;
  return o;
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#eaf8f0');
    s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('@');
    if (name === 'Attendance') {
      ['B:B', 'E:E', 'H:H'].forEach(c => s.getRange(c).setNumberFormat('@'));
      s.setColumnWidth(3, 200); s.setColumnWidth(4, 220); s.setColumnWidth(10, 120);
    }
    if (name === 'Settings') s.getRange('B:B').setNumberFormat('@');
  }
  return s;
}

function folder_() {
  const it = DriveApp.getFoldersByName(SIGNATURE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SIGNATURE_FOLDER);
}

function safe_(s) { return String(s).replace(/[\\/:*?"<>|]+/g, '').slice(0, 60); }

function fmtDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}

function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
