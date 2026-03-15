# 🧭 PP5 Smart — Architecture Overview

เอกสารนี้ตั้งใจให้ “คน/AI ที่เพิ่งเข้ามา” เข้าใจระบบได้เร็ว และรู้ว่าจะเริ่มแก้ตรงไหนโดยไม่กระทบทั้งระบบ

## ภาพรวมระบบ
- Platform: Google Apps Script (Web App) + HTML/JS (Frontend)
- Data store หลัก: Google Spreadsheet (ชีตหลายตัว)
- รูปแบบการทำงาน: SPA ฝั่งแอดมิน/ครู + Portal แยกสำหรับผู้ปกครอง

## จุดเข้า (Entry Points)
### Web App Routing
- `doGet(e)` เป็น router หลัก: [Code.js](Code.js)
  - `?page=setup_wizard` → [setup_wizard.html](setup_wizard.html) (ผ่าน `serveSetupWizard`)
  - `?page=parent_portal` → [parent_portal.html](parent_portal.html)
  - `?page=dashboard|app` หรือไม่มี page → [spa_main_menu.html](spa_main_menu.html)
  - `?dev=1` หรือ deployment `/dev` → `serveDevBypass()` (บังคับเป็น admin สำหรับทดสอบ)

### การโหลดหน้าใน SPA
- SPA หลักอยู่ที่: [spa_main_menu.html](spa_main_menu.html)
- หน้าในระบบเป็นไฟล์ `.html` แยก และโหลดผ่าน `google.script.run.getPageContent(pageName, token)` (token-based) ใน [Code.js](Code.js)

## Authentication / Authorization
### Session (ฝั่ง server)
- ใช้ `PropertiesService.getUserProperties()` เก็บ `username` ของผู้ล็อกอิน
- อ่าน session ด้วย `getLoginSession()` / `getCurrentUser()` ใน [Users.js](Users.js)

### Token (ฝั่ง client)
- ใช้ token เซ็นด้วย HMAC: `generateAuthToken()` / `verifyAuthToken()` ใน [Code.js](Code.js)
- SPA เก็บ `authToken` ใน `localStorage` แล้วส่งไปกับบาง API (เช่น โหลดหน้า, เช็คอัปเดต, ส่งออก)

### บทบาท (Role)
- role ถูกอ่านจากชีต `Users` (คอลัมน์ `role`) แล้วใช้กำหนดเมนู/สิทธิ์
- แนวคิดสำคัญ: UI ซ่อนเมนูได้ แต่ต้อง “บล็อกฝั่ง server” ซ้ำเสมอสำหรับงาน admin

## โมดูลหลักในระบบ (Server Files)
- Setup / Multi-tenant: [Installer.js](Installer.js)
- Users & auth helpers: [Users.js](Users.js), [Code.js](Code.js)
- Students CRUD/import/search: [Students.js](Students.js)
- Subjects/score config: [Scoring.js](Scoring.js), [subjectManagement.js](subjectManagement.js), [subject.js](subject.js)
- Attendance: [Attendance.js](Attendance.js), [HolidayService.js](HolidayService.js)
- Assessments: [Assessments.js](Assessments.js)
- Reporting/PDF:
  - รายงานหน้าเดียว: [OnePageReport.js](OnePageReport.js)
  - รายงานรวม/อื่น ๆ: [ReportGenerator.js](ReportGenerator.js), [Export_PDF.js](Export_PDF.js), [pp5book.js](pp5book.js), [pp5cover.js](pp5cover.js)
- Parent portal server APIs: [get_student_scores.js](get_student_scores.js)
- System update (เช็คเวอร์ชันจาก GitHub): [AutoUpdate.js](AutoUpdate.js)

## Data Model (Spreadsheet)
- แหล่งจริงของ schema อยู่ที่ `setupSheetHeaders_()` ใน [Installer.js](Installer.js)
- เอกสาร schema แบบย่อยและข้อควรระวัง: [SHEETS_SCHEMA.md](SHEETS_SCHEMA.md)

## Flow สำคัญ (แผนที่ “ไหล” ของข้อมูล)
### 1) ติดตั้งโรงเรียนใหม่ (Setup Wizard)
- ผู้ใช้เปิด `?page=setup_wizard` → กรอกข้อมูล → `setupNewSchool(formData)` ใน [Installer.js](Installer.js)
- ระบบสร้าง spreadsheet ใหม่/หรือ copy template และสร้างชีต + header
- บันทึก ScriptProperties สำคัญ: `SPREADSHEET_ID`, `SCRIPT_ID`, `SCHOOL_NAME`, `PDF_OUTPUT_FOLDER_ID` (ตามขั้นตอนติดตั้ง)

### 2) บันทึกคะแนน → คลังคะแนน
- UI สร้าง/บันทึกตารางคะแนนผ่าน [Scoring.js](Scoring.js)
- ระบบสรุปลง `SCORES_WAREHOUSE` เพื่อใช้ Dashboard/รายงาน

### 3) สร้าง PDF รายงาน
- Parent portal เรียก `generateParentReportCard(studentId)` → ใช้ `generateOnePageReportPdf()` ใน [OnePageReport.js](OnePageReport.js)
- ไฟล์ PDF ถูกสร้างใน Drive และคืน URL เพื่อเปิดดู

### 4) ส่งออกให้โรงเรียนอื่น
- เมนูส่งออกใน SPA เรียก `exportProjectAsZip(token)` ใน [Installer.js](Installer.js)
- วิธีส่งออกเป็น “ลิงก์ทำสำเนา Spreadsheet” (Container-bound script ไปด้วย)
- รุ่นปัจจุบันทำ “sanitized export” (ไม่พ่วงข้อมูลนักเรียน/ผู้ใช้/ข้อมูลส่วนบุคคล)

## จุดที่ควรระวังเมื่อแก้ระบบ
- เปลี่ยนชื่อชีต/คอลัมน์: กระทบหลายไฟล์ → อ้างอิง schema ที่ [Installer.js](Installer.js) และ [SHEETS_SCHEMA.md](SHEETS_SCHEMA.md)
- Deployment `/exec` ต้อง redeploy ทุกครั้งเมื่อเปลี่ยนหน้า/ไฟล์สำคัญ (ต่างจาก `/dev`)
- สิทธิ์ admin ต้อง enforce ฝั่ง server (อย่าเชื่อเฉพาะเมนูฝั่ง client)

