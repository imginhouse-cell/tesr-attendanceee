/**
 * ระบบเช็คชื่อคลาส Claude AI — ส่วนหลังบ้านบน Google Sheet (Google Apps Script)
 * วิธีติดตั้งดูในไฟล์ README.md
 *
 * ชีตที่ใช้: "สำเนาของ บทที่ 6 · เฉลยสำหรับผู้สอน (ชีต + Apps Script)"
 * https://docs.google.com/spreadsheets/d/1n9Tr5oyydEUM3tEOpJuJFrPFQ1Ka31T6dDMJ_qv9weY/edit
 *
 * แท็บในชีต
 *   Roster     — รายชื่อผู้เรียน: id, ชื่อ-นามสกุล, หน่วยงาน, ชื่อภาษาอังกฤษ
 *   Attendance — การเช็คชื่อแบบละเอียด (เวลา สถานะ ลายเซ็น)
 *   Data       — บันทึกแบบอ่านง่าย: timestamp, name, note (มา / สาย)
 *   Settings   — title, start (เวลาเริ่มเรียน), grace (นาทีที่ยังไม่นับสาย)
 */
const ADMIN_PIN = '1234';           // ← เปลี่ยนเป็นรหัสผู้สอนของคุณ (ใช้นำเข้า/ลบรายชื่อ, ยกเลิกเช็คชื่อ, ตั้งค่า)
const TIMEZONE  = 'Asia/Bangkok';
const SPREADSHEET_ID = '1n9Tr5oyydEUM3tEOpJuJFrPFQ1Ka31T6dDMJ_qv9weY';
const LOG_TIME_FORMAT = 'dd/MM/yyyy HH:mm';

// ชื่อแท็บ + หัวคอลัมน์ในชีต (ซ้าย) ↔ ชื่อฟิลด์ที่หน้าเว็บใช้ (ขวา)
const SHEETS = {
  roster: {
    name: 'Roster',
    cols: [['id', 'id'], ['ชื่อ-นามสกุล', 'name'], ['หน่วยงาน', 'org'], ['ชื่อภาษาอังกฤษ', 'nickname']],
  },
  attendance: {
    name: 'Attendance',
    cols: [['วันที่', 'date'], ['id', 'studentId'], ['ชื่อ-นามสกุล', 'name'], ['หน่วยงาน', 'org'],
           ['เวลาเช็คอิน', 'timeText'], ['สถานะ', 'status'], ['สาย (นาที)', 'lateMinutes'],
           ['บันทึกเมื่อ (ISO)', 'iso'], ['ลิงก์ลายเซ็น', 'sigLink'], ['ลายเซ็น (ข้อมูลภาพ)', 'sig']],
  },
  log: {
    name: 'Data',
    cols: [['timestamp', 'timestamp'], ['name', 'name'], ['note', 'note']],
  },
  settings: {
    name: 'Settings',
    cols: [['key', 'key'], ['value', 'value']],
  },
};
// key ในแท็บ Settings ↔ ฟิลด์ที่หน้าเว็บใช้
const SETTING_KEYS = { title: 'title', start: 'startTime', grace: 'grace' };

/* ---------------- ตัวช่วยอ่าน/เขียนชีต ---------------- */
function ss_() {
  try { return SpreadsheetApp.openById(SPREADSHEET_ID); }
  catch (e) { return SpreadsheetApp.getActive(); }
}

function sheet_(key) {
  const def = SHEETS[key];
  const ss = ss_();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(c => c[0])]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** อ่านทุกแถวเป็น object ตามชื่อฟิลด์ของเว็บ (อ้างหัวคอลัมน์จริงในชีต จึงสลับลำดับคอลัมน์ได้) */
function rows_(key) {
  const def = SHEETS[key];
  const values = sheet_(key).getDataRange().getValues();
  const head = (values.shift() || []).map(h => String(h).trim());
  const idx = def.cols.map(([h, f]) => [f, head.indexOf(h)]);
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    idx.forEach(([f, j]) => { o[f] = j >= 0 ? r[j] : ''; });
    return o;
  });
}

/** เรียงค่าให้ตรงกับหัวคอลัมน์จริงของชีต */
function line_(key, obj) {
  const def = SHEETS[key];
  const sh = sheet_(key);
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  const map = {};
  def.cols.forEach(([h, f]) => { map[h] = f; });
  return head.map(h => {
    const f = map[h];
    const v = f ? obj[f] : '';
    return (v === undefined || v === null) ? '' : v;
  });
}

function writeRow_(key, row, obj) {
  const sh = sheet_(key);
  const line = line_(key, obj);
  const rg = sh.getRange(row, 1, 1, line.length);
  rg.setNumberFormat('@');          // เก็บเป็นข้อความ ไม่ให้ชีตแปลงวันที่/เวลาเอง
  rg.setValues([line.map(v => typeof v === 'number' ? String(v) : v)]);
}

function append_(key, obj) {
  writeRow_(key, sheet_(key).getLastRow() + 1, obj);
}

function upsert_(key, field, val, obj) {
  const found = rows_(key).find(r => String(r[field]).trim() === String(val).trim());
  if (found) writeRow_(key, found._row, obj);
  else append_(key, obj);
}

function deleteWhere_(key, test) {
  const sh = sheet_(key);
  rows_(key).filter(test).map(r => r._row).sort((a, b) => b - a).forEach(n => sh.deleteRow(n));
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);           // รองรับ dd/MM/yyyy ด้วย
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s.slice(0, 10);
}

function hm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TIMEZONE, 'HH:mm');
  return String(v || '').slice(0, 5);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- แท็บ Data ---------------- */
function appendLog_(d) {
  const ms = Number(d.time) || Date.now();
  append_('log', {
    timestamp: Utilities.formatDate(new Date(ms), TIMEZONE, LOG_TIME_FORMAT),
    name: String(d.name || ''),
    note: (d.status === 'มาสาย' || Number(d.lateMinutes) > 0) ? 'สาย' : 'มา',
  });
}

function removeLog_(date, name) {
  deleteWhere_('log', r => dateKey_(r.timestamp) === date && String(r.name).trim() === String(name).trim());
}

/* ---------------- API ที่หน้าเว็บเรียก ---------------- */
/** อ่านข้อมูลทั้งหมดของวันที่ระบุ */
function doGet(e) {
  const date = dateKey_((e && e.parameter && e.parameter.date) || '');
  const students = rows_('roster').filter(r => String(r.id).trim() && String(r.name).trim()).map(r => ({
    id: String(r.id).trim(), code: '', name: String(r.name).trim(), org: String(r.org || ''),
    nickname: String(r.nickname || ''), email: '', order: r._row,
  }));
  const checkins = {};
  rows_('attendance').forEach(r => {
    if (dateKey_(r.date) !== date || !String(r.studentId).trim()) return;
    let time = Date.parse(String(r.iso || ''));
    if (isNaN(time)) time = Date.parse(date + 'T' + (hm_(r.timeText) || '00:00') + ':00+07:00');
    checkins[String(r.studentId).trim()] = {
      studentId: String(r.studentId).trim(), name: String(r.name || ''), code: '',
      time: time, sig: String(r.sig || ''),
    };
  });
  const config = {};
  rows_('settings').forEach(r => {
    const f = SETTING_KEYS[String(r.key).trim()];
    if (!f) return;
    config[f] = f === 'grace' ? Number(r.value) : (f === 'startTime' ? hm_(r.value) : String(r.value));
  });
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
        upsert_('roster', 'id', b.id, Object.assign({}, b.data, { id: b.id }));
        return json_({ ok: true });
      case 'putStudents':
        (b.items || []).forEach(it => upsert_('roster', 'id', it.id, Object.assign({}, it.data, { id: it.id })));
        return json_({ ok: true });
      case 'delStudent':
        deleteWhere_('roster', r => String(r.id).trim() === String(b.id));
        return json_({ ok: true });
      case 'checkin': {
        const date = dateKey_(b.date);
        if (rows_('attendance').some(r => dateKey_(r.date) === date && String(r.studentId).trim() === String(b.id)))
          return json_({ ok: false, error: 'exists' });
        const d = b.data || {};
        const ms = Number(d.time) || Date.now();
        append_('attendance', {
          date: date, studentId: b.id, name: d.name || '', org: d.org || '',
          timeText: d.timeText || Utilities.formatDate(new Date(ms), TIMEZONE, 'HH:mm'),
          status: d.status || '', lateMinutes: Number(d.lateMinutes) || 0,
          iso: new Date(ms).toISOString(), sigLink: '', sig: d.sig || '',
        });
        appendLog_(Object.assign({}, d, { time: ms }));
        return json_({ ok: true });
      }
      case 'delCheckin': {
        const date = dateKey_(b.date);
        const test = r => dateKey_(r.date) === date && String(r.studentId).trim() === String(b.id);
        const gone = rows_('attendance').filter(test);
        deleteWhere_('attendance', test);
        gone.forEach(r => removeLog_(date, r.name));
        return json_({ ok: true });
      }
      case 'putConfig': {
        const rev = {};
        Object.keys(SETTING_KEYS).forEach(k => { rev[SETTING_KEYS[k]] = k; });
        Object.keys(b.data || {}).forEach(f => {
          const k = rev[f];
          if (k) upsert_('settings', 'key', k, { key: k, value: b.data[f] });
        });
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** กด Run ฟังก์ชันนี้ครั้งแรกเพื่อตรวจแท็บและอนุญาตสิทธิ์ */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
}

/** (ไม่บังคับ) ทดสอบเขียนแถวตัวอย่างลงแท็บ Data แล้วลบออก */
function testLog() {
  appendLog_({ name: 'ทดสอบระบบ', time: Date.now(), status: 'มาเรียน' });
  removeLog_(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'), 'ทดสอบระบบ');
}
