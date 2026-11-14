/*
 * ตั้งค่าเว็บเช็คชื่อ: แก้ไฟล์นี้ไฟล์เดียว
 *
 * SHEET_API : URL ของ Google Apps Script Web App (ลงท้ายด้วย /exec)
 *             ดูวิธีสร้างได้ใน README.md หัวข้อ "ขั้นที่ 1"
 *             ถ้าเว้นว่าง หน้าเว็บจะเก็บข้อมูลไว้ในเครื่องที่เปิดเท่านั้น (ใช้ทดสอบได้)
 * POLL_MS   : ความถี่ในการดึงข้อมูลล่าสุดจาก Google Sheets (มิลลิวินาที)
 *
 * URL นี้ไม่ใช่ความลับ ส่วนรหัสผู้ดูแล (ADMIN_KEY) อยู่ใน Script Properties ของ Apps Script ห้ามใส่ไว้ในไฟล์นี้
 */
window.ATTENDANCE_CONFIG = {
  SHEET_API: '',
  POLL_MS: 15000
};
