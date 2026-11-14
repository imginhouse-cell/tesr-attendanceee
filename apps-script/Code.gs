/**
 * ระบบเช็คชื่อคลาส Claude AI — ส่วนหลังบ้านบน Google Sheet (Google Apps Script)
 * วิธีติดตั้งดูในไฟล์ README.md
 */
const ADMIN_PIN = '1234';           // ← เปลี่ยนเป็นรหัสผู้สอนของคุณ (ใช้นำเข้า/ลบรายชื่อ, ยกเลิกเช็คชื่อ, ตั้งค่า)
const TIMEZONE  = 'Asia/Bangkok';

const SHEETS = {
  students: { name: 'รายชื่อ',  cols: ['id', 'code', 'name', 'org', 'nickname', 'email', 'order', 'updatedAt'] },
  checkins: { name: 'เช็คชื่อ', cols: ['date', 'studentId', 'code', 'name', 'org', 'timeText', 'status', 'lateMinutes', 'time', 'signature', 'createdAt'] },
  config:   { name: 'ตั้งค่า',   cols: ['key', 'value'] },
};

function sheet_(key) {
  const ss = SpreadsheetApp.getActive();
  const def = SHEETS[key];
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold').setBackground('#d1fae5');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, def.cols.length).setNumberFormat('@'); // เก็บเป็นข้อความ ไม่ให้แปลงวันที่/เวลาเอง
  }
  return sh;
}

function rows_(key) {
  const sh = sheet_(key);
  const values = sh.getDataRange().getValues();
  const head = values.shift() || [];
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((c, j) => { o[c] = r[j]; });
    return o;
  });
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function upsert_(key, keyCol, keyVal, obj) {
  const sh = sheet_(key);
  const cols = SHEETS[key].cols;
  const found = rows_(key).find(r => String(r[keyCol]) === String(keyVal));
  const line = cols.map(c => (obj[c] === undefined || obj[c] === null) ? '' : obj[c]);
  if (found) sh.getRange(found._row, 1, 1, cols.length).setValues([line]);
  else sh.appendRow(line);
}

function deleteWhere_(key, test) {
  const sh = sheet_(key);
  rows_(key).filter(test).map(r => r._row).sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
}

/** อ่านข้อมูลทั้งหมดของวันที่ระบุ */
function doGet(e) {
  const date = dateKey_((e && e.parameter && e.parameter.date) || '');
  const students = rows_('students').filter(r => r.id).map(r => ({
    id: String(r.id), code: String(r.code || ''), name: String(r.name || ''), org: String(r.org || ''),
    nickname: String(r.nickname || ''), email: String(r.email || ''), order: Number(r.order) || 0,
  }));
  const checkins = {};
  rows_('checkins').forEach(r => {
    if (dateKey_(r.date) !== date || !r.studentId) return;
    checkins[String(r.studentId)] = {
      studentId: String(r.studentId), name: String(r.name || ''), code: String(r.code || ''),
      time: Number(r.time), sig: String(r.signature || ''),
    };
  });
  const config = {};
  rows_('config').forEach(r => { if (r.key) config[r.key] = r.key === 'grace' ? Number(r.value) : String(r.value); });
  return json_({ ok: true, students, checkins, config });
}

/** บันทึกข้อมูล */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const isAdmin = String(b.pin || '') === String(ADMIN_PIN);
    const adminOnly = ['verify', 'putStudent', 'putStudents', 'delStudent', 'delCheckin', 'putConfig'];
    if (adminOnly.indexOf(b.action) >= 0 && !isAdmin) return json_({ ok: false, error: 'pin' });

    switch (b.action) {
      case 'verify':
        return json_({ ok: true });
      case 'putStudent':
        upsert_('students', 'id', b.id, Object.assign({}, b.data, { id: b.id }));
        return json_({ ok: true });
      case 'putStudents':
        (b.items || []).forEach(it => upsert_('students', 'id', it.id, Object.assign({}, it.data, { id: it.id })));
        return json_({ ok: true });
      case 'delStudent':
        deleteWhere_('students', r => String(r.id) === String(b.id));
        return json_({ ok: true });
      case 'checkin': {
        const date = dateKey_(b.date);
        if (rows_('checkins').some(r => dateKey_(r.date) === date && String(r.studentId) === String(b.id)))
          return json_({ ok: false, error: 'exists' });
        const d = b.data || {};
        const line = SHEETS.checkins.cols.map(c =>
          c === 'date' ? date : c === 'studentId' ? b.id : c === 'signature' ? (d.sig || '') : (d[c] === undefined || d[c] === null ? '' : d[c]));
        sheet_('checkins').appendRow(line);
        return json_({ ok: true });
      }
      case 'delCheckin':
        deleteWhere_('checkins', r => dateKey_(r.date) === dateKey_(b.date) && String(r.studentId) === String(b.id));
        return json_({ ok: true });
      case 'putConfig':
        Object.keys(b.data || {}).forEach(k => upsert_('config', 'key', k, { key: k, value: b.data[k] }));
        return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** กด Run ฟังก์ชันนี้ครั้งแรกเพื่อสร้างแท็บและอนุญาตสิทธิ์ */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
}
