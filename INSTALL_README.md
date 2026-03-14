# 📦 PP5 Smart — ระบบบันทึกผลการเรียน ปพ.5
## คำแนะนำการติดตั้งสำหรับโรงเรียนใหม่

---

## ขั้นตอนการติดตั้ง

### ขั้นที่ 1: สร้าง Google Apps Script Project

1. เปิด https://script.google.com
2. ล็อกอินด้วย Google Account ของโรงเรียน
3. กดปุ่ม **"+ New project"**
4. ตั้งชื่อ เช่น "PP5 Smart - [ชื่อโรงเรียน]"

### ขั้นที่ 2: อัปโหลดไฟล์โค้ด

1. ในหน้า GAS Editor ลบไฟล์ `Code.gs` เดิมที่มีอยู่
2. กดปุ่ม **"+"** (เพิ่มไฟล์) → เลือก **Script** หรือ **HTML**
3. อัปโหลดไฟล์ทั้งหมดจากโฟลเดอร์นี้:
   - ไฟล์ `.js` → เลือก "Script" แล้ววางโค้ด (ชื่อไฟล์ไม่ต้องใส่ .js)
   - ไฟล์ `.html` → เลือก "HTML" แล้ววางโค้ด (ชื่อไฟล์ไม่ต้องใส่ .html)
4. ไฟล์ `appsscript.json` → ไปที่ Project Settings → เปิด "Show appsscript.json" → วางเนื้อหาทับ

**หมายเหตุ**: ไม่ต้องอัปโหลดไฟล์ต่อไปนี้:
- `INSTALL_README.md` (ไฟล์นี้)
- `.clasp.json`
- `สำรองโค้ด.html` (โค้ดเก่าสำรอง ไม่จำเป็น)

### ขั้นที่ 3: Deploy Web App

1. กดปุ่ม **"Deploy"** → **"New deployment"**
2. กดไอคอนเฟือง → เลือก **"Web app"**
3. ตั้งค่า:
   - **Description**: PP5 Smart
   - **Execute as**: Me
   - **Who has access**: Anyone
4. กด **"Deploy"**
5. กด **"Authorize access"** → อนุญาตสิทธิ์ทั้งหมด
6. **คัดลอก URL** ที่ได้ เก็บไว้ (นี่คือลิงก์เข้าระบบ)

### ขั้นที่ 4: เปิด Setup Wizard

1. เปิด URL ที่คัดลอกมา แล้วต่อท้ายด้วย `?page=setup_wizard`
   ตัวอย่าง: `https://script.google.com/macros/s/XXXXX/exec?page=setup_wizard`
2. กรอกข้อมูลโรงเรียน:
   - ชื่อโรงเรียน
   - ที่อยู่
   - ชื่อผู้อำนวยการ
   - ปีการศึกษา
   - ชื่อผู้ใช้ Admin + รหัสผ่าน
3. กดปุ่ม **"🚀 สร้างระบบใหม่"**
4. รอจนเสร็จ — ระบบจะสร้าง Spreadsheet และชีตทั้งหมดให้อัตโนมัติ

### ขั้นที่ 5: เริ่มใช้งาน

1. เปิด URL หลัก (ไม่ต้องใส่ `?page=`)
2. ล็อกอินด้วย Admin ที่ตั้งไว้
3. เริ่มใช้งานได้เลย!

---

## รายการไฟล์ทั้งหมด

### ไฟล์ Script (.js) — สร้างเป็น Script ใน GAS Editor
- Code.js (ไฟล์หลัก)
- Installer.js (ตัวติดตั้ง)
- Attendance.js
- Assessments.js
- Dashboard.js
- Export Csv.js
- Export_PDF.js
- Filter.js
- get_student_scores.js
- globalSettings.js
- HolidayService.js
- logoManagement.js
- OnePageReport.js
- pp5book.js
- pp5book_sheets.js
- pp5cover.js
- pp6.js
- ReportGenerator.js
- SchoolMISSync.js
- Scoring.js
- settings_unified.js
- sheet_template.js
- Students.js
- subject.js
- subjectManagement.js
- Teachers.js
- Users.js
- Utilities.js
- utils.js
- validation.js

### ไฟล์ HTML (.html) — สร้างเป็น HTML ใน GAS Editor
- (ไฟล์ .html ทั้งหมดในโฟลเดอร์)

### ไฟล์ตั้งค่า
- appsscript.json (วางใน Project Settings)

---

## ปัญหาที่พบบ่อย

### Q: กด Deploy แล้วขึ้น "Authorization required"
A: กด "Authorize access" แล้วเลือก Google Account → กด "Allow"

### Q: เปิด URL แล้วขึ้นหน้าว่าง
A: ตรวจสอบว่าอัปโหลดไฟล์ครบทุกไฟล์ โดยเฉพาะ Code.js

### Q: Setup Wizard ขึ้น error
A: ตรวจสอบว่าไฟล์ Installer.js และ settings_unified.js อัปโหลดแล้ว

---

**พัฒนาโดย**: ระบบ PP5 Smart
**เวอร์ชัน**: 2.x
