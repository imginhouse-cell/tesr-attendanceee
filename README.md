# เว็บเช็คชื่อคลาส Claude AI

เว็บเช็คชื่อออนไลน์ ใช้ได้ทั้งมือถือ แท็บเล็ต และคอม ผู้เรียนค้นหาชื่อ เซ็นชื่อ แล้วกดยืนยัน ระบบบันทึกเวลาลง **Google Sheets** ให้ทันที

- โฮสต์ฟรีบน **GitHub Pages** (เป็นไฟล์ static ไม่ต้องมีเซิร์ฟเวอร์)
- หลังบ้านคือ **Google Apps Script** ที่ผูกกับ Google Sheets ของคุณ
- ค้นหาได้ทั้งชื่อไทย ชื่ออังกฤษ และบริษัท เซ็นได้ด้วยนิ้ว ปากกา หรือเมาส์ และระบุ "ตรงเวลา / สาย X นาที" ให้อัตโนมัติ
- นำเข้ารายชื่อจาก CSV / Excel / ไฟล์ export ของ Google Form ได้ (แยกผู้เข้าเรียนคนที่ 2–5 ให้เอง)
- ส่งออกเป็น CSV หรือใบเซ็นชื่อพร้อมลายเซ็นสำหรับพิมพ์ได้

## โครงสร้างไฟล์

```
index.html                หน้าเว็บทั้งหมด
config.js                 ใส่ URL ของ Apps Script ที่นี่ (ไฟล์เดียวที่ต้องแก้)
apps-script/Code.gs       โค้ดหลังบ้านสำหรับ Google Sheets
apps-script/appsscript.json  ไฟล์ตั้งค่า Apps Script (เขตเวลา Asia/Bangkok, Web app)
docs/sample-roster.csv    ตัวอย่างไฟล์รายชื่อ (ชื่อสมมติ)
.nojekyll                 ให้ GitHub Pages เสิร์ฟไฟล์ตรงๆ
```

---

## ขั้นที่ 1: สร้างหลังบ้าน Google Sheets

1. สร้าง Google Sheets ใหม่ (เช่น ชื่อ "เช็คชื่อ Claude AI")
2. เมนู **ส่วนขยาย › Apps Script**
3. ลบโค้ดเดิมทิ้ง แล้ววางเนื้อหาจาก `apps-script/Code.gs`
4. ไอคอนเฟือง **Project Settings**:
   - ติ๊ก "Show appsscript.json" แล้ววางเนื้อหาจาก `apps-script/appsscript.json` (ตั้งเขตเวลาเป็นกรุงเทพฯ)
   - ส่วน **Script Properties** › Add property › ชื่อ `ADMIN_KEY` › ค่าเป็นรหัสผู้ดูแลที่คุณตั้งเอง
5. กลับไปหน้า Editor เลือกฟังก์ชัน `setup` แล้วกด **Run** และกดอนุญาตสิทธิ์ (Sheets + Drive สำหรับเก็บรูปลายเซ็น)
6. **Deploy › New deployment** › เลือกประเภท **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - กด Deploy แล้วคัดลอก **Web app URL** (ลงท้าย `/exec`)
7. ทดสอบโดยเปิด URL นั้นในเบราว์เซอร์ ควรเห็น `{"ok":true,"hello":"TESR attendance API","admin_key_set":true}`

## ขั้นที่ 2: ใส่ URL ในเว็บ

เปิด `config.js` แล้ววาง URL:

```js
window.ATTENDANCE_CONFIG = {
  SHEET_API: 'https://script.google.com/macros/s/xxxxxxxx/exec',
  POLL_MS: 15000
};
```

## ขั้นที่ 3: ขึ้น GitHub Pages

1. สร้าง repository ใหม่ แล้วอัปโหลดไฟล์ทั้งหมดในโฟลเดอร์นี้
2. **Settings › Pages** › Source: *Deploy from a branch* › Branch: `main` / `(root)` › Save
3. รอประมาณ 1 นาที จะได้ลิงก์ `https://<username>.github.io/<repo>/`

## ขั้นที่ 4: ใส่รายชื่อและใช้งาน

1. ผู้สอนเปิด `https://<username>.github.io/<repo>/#admin`
2. ปุ่ม **ตั้งค่า** › ใส่ **รหัสผู้ดูแล** (ADMIN_KEY) › บันทึก (เครื่องนี้จะจำรหัสไว้)
3. ปุ่ม **นำเข้า CSV/Excel** › เลือกไฟล์รายชื่อ › **แทนที่รายชื่อเดิม** (รายชื่อจะไปอยู่ในชีต Roster)
4. ส่งลิงก์ปกติ (ไม่มี `#admin`) ให้ผู้เรียน หรือทำเป็น QR Code ติดหน้าห้อง

## ข้อมูลที่บันทึกใน Google Sheets

| ชีต | เนื้อหา |
|---|---|
| **Attendance** | 1 แถวต่อ 1 การเช็คชื่อ: วันที่, id, ชื่อ, หน่วยงาน, เวลาเช็คอิน, สถานะ (ตรงเวลา/สาย), สาย (นาที), เวลาแบบ ISO, ลิงก์รูปลายเซ็นใน Drive, ข้อมูลภาพลายเซ็น |
| **Roster** | รายชื่อผู้เรียน (แก้ผ่านหน้าเว็บโหมด `#admin`) |
| **Settings** | ชื่อคลาส, เวลาเริ่มเรียน, เวลาผ่อนผันสาย |

รูปลายเซ็นทั้งหมดอยู่ในโฟลเดอร์ Drive ชื่อ **Attendance Signatures** ถ้าไม่ต้องการ ให้ตั้ง `SAVE_SIGNATURE_TO_DRIVE = false` ใน `Code.gs`

## หมายเหตุ

- **อย่าใส่ไฟล์รายชื่อจริงลง repo สาธารณะ** เพราะมีชื่อและบริษัทของผู้เรียน ให้นำเข้าผ่านหน้าเว็บแทน (`.gitignore` กันโฟลเดอร์ `private/` ไว้ให้แล้ว)
- ถ้าแก้ `Code.gs` ภายหลัง ต้อง **Deploy › Manage deployments › แก้ไข (ไอคอนดินสอ) › Version: New version** URL จะยังเป็นอันเดิม
- ผู้เรียนแต่ละคนเช็คชื่อได้ครั้งเดียวต่อวัน ถ้าต้องแก้ ผู้สอนกด "ยกเลิกการเช็คชื่อ" ในโหมด `#admin`
- ถ้า `SHEET_API` ว่างไว้ เว็บยังใช้งานได้แต่เก็บข้อมูลในเครื่องที่เปิดเครื่องเดียว (เหมาะกับการทดสอบ)
