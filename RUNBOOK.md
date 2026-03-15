# 🧰 PP5 Smart — Runbook (Deploy / Ops / Troubleshooting)

เอกสารนี้ไว้สำหรับคน/AI ที่ต้อง “ดูแลระบบ” และทำงานเดิมซ้ำ ๆ ได้เร็ว (deploy, ส่งออก, ติดตั้งโรงเรียนใหม่)

## Deployment Basics
### `/dev` vs `/exec`
- `/dev` = ใช้โค้ดล่าสุดของโปรเจกต์ (เหมาะทดสอบ)
- `/exec` = ใช้เวอร์ชันที่ deploy ไว้ (ผู้ใช้จริง) ต้อง redeploy ทุกครั้งเมื่อเปลี่ยนไฟล์

### ทำไม push แล้ว `/exec` ยังไม่เปลี่ยน
- เพราะ `clasp push` อัปเดต “โค้ดในโปรเจกต์” แต่ `/exec` ผูกกับ “deployment version”
- หลังแก้ไฟล์สำคัญ ควร deploy ใหม่เสมอ

## ค่าที่เก็บใน Script Properties (สำคัญ)
- `SPREADSHEET_ID`: Spreadsheet หลักของโรงเรียน
- `SCRIPT_ID`: กันกรณี copy project แล้วไปทับของเดิม
- `TOKEN_SECRET`: secret สำหรับ token (ถูกสร้างอัตโนมัติเมื่อยังไม่มี)
- `PDF_OUTPUT_FOLDER_ID`: โฟลเดอร์ปลายทางสำหรับ PDF (ถ้ามีการตั้ง)
- `TEMPLATE_SPREADSHEET_ID`: Spreadsheet แม่แบบ (เฉพาะโรงเรียนต้นแบบ)
- `ALLOW_EXPORT_TO_OTHER_SCHOOLS`: เปิดสิทธิ์ส่งออกเฉพาะต้นแบบ (true/1)
- `EXPORT_TEMPLATE_SPREADSHEET_ID`: ID ของไฟล์ export template ล่าสุดที่ระบบสร้าง

แหล่งอ้างอิง: [Code.js](Code.js), [Installer.js](Installer.js)

## Setup Wizard (ติดตั้งโรงเรียนใหม่)
### Flow
- เข้า `?page=setup_wizard` → submit → `setupNewSchool(formData)` ใน [Installer.js](Installer.js)
- ระบบจะสร้าง spreadsheet ใหม่ หรือ copy จาก template แล้วสร้างชีต/headers

### ถ้าเปิดเว็บแล้ววนกลับ Setup Wizard
- `isSetupComplete_()` จะตรวจ `SPREADSHEET_ID` + marker `installed_script_id` ใน `global_settings`
- ถ้า copy project แล้ว `SCRIPT_ID` ไม่ตรง ระบบจะล้าง properties แล้วบังคับ Setup Wizard ใหม่
- แหล่งอ้างอิง: [Code.js](Code.js)

## ส่งออกให้โรงเรียนอื่น (สำหรับโรงเรียนต้นแบบ)
### หลักการ
- วิธีส่งออกเป็น “ลิงก์ทำสำเนา Spreadsheet” (Container-bound script จะไปด้วย)
- รุ่นปัจจุบันเป็น “sanitized export”:
  - คัดลอกไฟล์ระบบปัจจุบัน → ล้างข้อมูลนักเรียน/ผู้ใช้/ข้อมูลส่วนบุคคล → คืนลิงก์ `/copy`
- แหล่งอ้างอิง: [Installer.js](Installer.js) `exportProjectAsZip(token)`

### จำกัดสิทธิ์ “ส่งออก” เฉพาะโรงเรียนต้นแบบ
- UI: เมนูถูกทำเป็น `master-only` แล้วเช็คความสามารถจาก server
- Server: บล็อก `exportProjectAsZip` ถ้าไม่ใช่ต้นแบบ
- การตัดสินว่าเป็น “ต้นแบบ”:
  - `ALLOW_EXPORT_TO_OTHER_SCHOOLS=true/1` หรือ
  - มี `TEMPLATE_SPREADSHEET_ID` ที่ไม่ใช่ placeholder

## Parent Portal
- หน้า: [parent_portal.html](parent_portal.html)
- API หลัก:
  - `getCurrentUser()` ใน [Users.js](Users.js)
  - `generateParentReportCard(studentId)` ใน [get_student_scores.js](get_student_scores.js) → `generateOnePageReportPdf()` ใน [OnePageReport.js](OnePageReport.js)

## อัปเดตระบบจาก GitHub (Check Update)
### ระบบเช็คอัปเดตทำอะไร
- ไปอ่านเวอร์ชันจาก GitHub แล้วเทียบกับเวอร์ชันในระบบ
- แสดงผล/คำแนะนำ (ไม่ทำ auto-update ทับไฟล์ในโปรเจกต์แบบ 100%)
- แหล่งอ้างอิง: [AutoUpdate.js](AutoUpdate.js)

### แนวทางอัปเดตที่เสถียรกว่า
- แนะนำใช้ `git pull` + `clasp push` + redeploy `/exec`

## Checklist หลังแก้โค้ด
- ทดสอบหน้า: login, spa main menu, parent portal, setup wizard (ถ้ามีแก้ส่วนติดตั้ง)
- ถ้าแก้ไฟล์ที่ผู้ใช้เห็น (HTML/JS ฝั่งหน้าเว็บ): redeploy `/exec`
- ถ้าแก้ auth: ทดสอบ `/exec` และ `/dev` (dev bypass)

