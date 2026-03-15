# 🚀 PP5 Smart — Onboarding (60 นาทีแรก)

เอกสารนี้ออกแบบให้คน/AI ที่เพิ่งเข้ามา “เปิด repo แล้วเริ่มทำงานได้ภายใน 60 นาที” โดยไม่ต้องเดาสุ่ม

## 0) สิ่งที่ต้องรู้ก่อน (2 นาที)
- ระบบเป็น Google Apps Script Web App + Spreadsheet เป็นฐานข้อมูลหลัก
- มี 2 UI หลัก:
  - SPA สำหรับครู/แอดมิน: [spa_main_menu.html](spa_main_menu.html)
  - Portal ผู้ปกครอง: [parent_portal.html](parent_portal.html)
- Router หลักอยู่ที่ `doGet(e)` ใน [Code.js](Code.js)

เอกสารที่ควรเปิดคู่กัน:
- ภาพรวมสถาปัตย์: [ARCHITECTURE.md](ARCHITECTURE.md)
- schema ชีต: [SHEETS_SCHEMA.md](SHEETS_SCHEMA.md)
- งานดูแลระบบ/deploy: [RUNBOOK.md](RUNBOOK.md)

## 1) หา “จุดเข้า” ให้เจอ (10 นาที)
### 1.1 Routing ของเว็บ
- อ่าน [Code.js](Code.js) เฉพาะส่วน `doGet(e)` และ page serving:
  - `serveMainApp()` → SPA
  - `serveParentPortal_()` → parent portal
  - `serveSetupWizard()` → setup wizard

### 1.2 โครงสร้าง SPA
- เปิด [spa_main_menu.html](spa_main_menu.html)
  - ดูเมนู (data-page / onclick)
  - ดูการโหลดหน้า (เรียก `getPageContent(pageName, token)` ใน [Code.js](Code.js))
  - ดู role-based menu (admin-only/teacher-only/master-only)

### 1.3 โครงสร้าง Parent Portal
- เปิด [parent_portal.html](parent_portal.html)
  - ดูฟังก์ชัน `getCurrentUser()` (ฝั่ง client เรียก server)
  - ดู action “ใบรายงาน” → `generateParentReportCard(studentId)`
  - ฝั่ง server อยู่ใน [get_student_scores.js](get_student_scores.js)

## 2) เข้าใจ “ข้อมูล” ก่อนแก้ (10 นาที)
อ่าน schema ที่สร้างใน installer:
- `setupSheetHeaders_()` ใน [Installer.js](Installer.js)

ชีตที่ต้องจำให้แม่น (ขั้นต่ำ):
- `global_settings` (key/value + marker ติดตั้ง)
- `Users` (username/role)
- `Students` (student_id/grade/class_no)
- `รายวิชา` (template)
- `SCORES_WAREHOUSE` (คลังคะแนน)

ดูสรุปแบบอ่านง่าย: [SHEETS_SCHEMA.md](SHEETS_SCHEMA.md)

## 3) เข้าใจ auth ให้ชัด (10 นาที)
ระบบมี 2 แนวทางคู่กัน:
- Session (server): `getLoginSession()` / `getCurrentUser()` ใน [Users.js](Users.js)
- Token (client): `generateAuthToken()` / `verifyAuthToken()` ใน [Code.js](Code.js)

กฎสำคัญสำหรับงานใหม่:
- “ซ่อนเมนู” ใน UI ไม่พอ ต้องเช็คสิทธิ์ฝั่ง server ด้วยเสมอ
- `/dev` อาจ bypass บางส่วนเพื่อทดสอบ (ดู `serveDevBypass()` ใน [Code.js](Code.js))

## 4) เช็คแผนที่ API (10 นาที)
เปิด: [api_inventory.md](api_inventory.md)
- ดู page → function → serverFile
- ถ้าจะเพิ่มฟีเจอร์ใหม่ ให้เพิ่มรายการในตารางนี้ด้วย (ช่วยคนรุ่นถัดไป)

## 5) เลือกงานแรกแบบ “ปลอดภัย” (15 นาที)
งานที่เหมาะเป็นงานแรก:
- ปรับ UI/ข้อความ/สี (ไม่แตะ schema)
- เพิ่มปุ่ม/สถานะโหลดในหน้าเดิม (เช่น spinner)
- เพิ่ม validation ฝั่ง client (ไม่แตะฐานข้อมูล)

งานที่ไม่ควรเริ่มทันทีถ้ายังไม่เข้าใจ:
- เปลี่ยนชื่อชีต/คอลัมน์
- ปรับ logic คำนวณคะแนน/GPA
- ปรับ flow Setup Wizard / Multi-tenant

## 6) Checklist ก่อนส่ง PR/ก่อน deploy (3 นาที)
- ตรวจว่าแก้ทั้ง “client + server” ครบถ้าเป็นเรื่องสิทธิ์
- ถ้าแก้ไฟล์ที่ผู้ใช้เห็น (HTML/JS): ต้อง redeploy `/exec`
- ถ้าแก้ schema: อัปเดต [SHEETS_SCHEMA.md](SHEETS_SCHEMA.md) + ทดสอบ flow หลัก

## Quick Map (สรุปแหล่งโค้ดตามงาน)
- Login/session/role: [Users.js](Users.js), [Code.js](Code.js)
- Setup Wizard/ติดตั้ง: [Installer.js](Installer.js), [setup_wizard.html](setup_wizard.html)
- Student CRUD/import: [Students.js](Students.js)
- Scoring/Warehouse: [Scoring.js](Scoring.js)
- Reports/PDF: [OnePageReport.js](OnePageReport.js), [ReportGenerator.js](ReportGenerator.js), [Export_PDF.js](Export_PDF.js)
- Parent portal: [parent_portal.html](parent_portal.html), [get_student_scores.js](get_student_scores.js)
- Update checker: [AutoUpdate.js](AutoUpdate.js)

