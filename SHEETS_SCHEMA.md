# 🗂️ PP5 Smart — Sheets Schema (Minimum Contract)

เอกสารนี้สรุป “สัญญาขั้นต่ำ” ของชีตที่ระบบใช้ เพื่อให้คน/AI ปรับปรุงต่อได้โดยไม่ทำให้ระบบพัง

แหล่งจริงของ headers อยู่ที่ `setupSheetHeaders_()` ใน [Installer.js](Installer.js)

## Core Sheets (ถาวร)
### `global_settings`
- ใช้เก็บ key/value ของระบบ (เช่น ชื่อโรงเรียน, ปีการศึกษา ฯลฯ)
- โครงสร้าง: `key`, `value`, `updatedAt`
- หมายเหตุ: มีการใช้ `installed_script_id` เป็น marker สำหรับกัน copied project ใน `isSetupComplete_()` [Code.js](Code.js)

### `Users`
- ใช้สำหรับ login/role
- โครงสร้างขั้นต่ำ (ตาม installer):
  - `username`, `password`, `role`, `firstName`, `lastName`, `email`, `className`, `active`, `createdAt`
- หมายเหตุ:
  - `role` ใช้ควบคุมสิทธิ์ (`admin`/`teacher`/`parent`)
  - ถ้าเปลี่ยนชื่อคอลัมน์ จะกระทบ session และการตรวจสิทธิ์หลายจุด (ดู [Users.js](Users.js))

### `Students`
- ข้อมูลนักเรียนหลักของระบบ
- โครงสร้างขั้นต่ำ (ตาม installer):
  - `student_id`, `id_card`, `title`, `firstname`, `lastname`, `grade`, `class_no`, `gender`, `created_date`, `status`
- หมายเหตุ:
  - หลายไฟล์อ่านคอลัมน์แบบยืดหยุ่น (มี fallback เป็น “นักเรียน”) แต่ควรยึด schema หลัก

### `รายวิชา`
- เทมเพลตรายวิชา (ควรส่งออกไปโรงเรียนอื่นได้)
- โครงสร้างขั้นต่ำ:
  - `ชั้น`, `รหัสวิชา`, `ชื่อวิชา`, `ชั่วโมง/ปี`, `ประเภทวิชา`, `ครูผู้สอน`, `คะแนนระหว่างปี`, `คะแนนปลายปี`
- หมายเหตุ:
  - ในการส่งออกแบบ sanitized จะล้างคอลัมน์ `ครูผู้สอน` แต่คงรายวิชาไว้

### `Holidays`
- วันหยุด
- โครงสร้างขั้นต่ำ: `date`, `description`, `type`
- มี fallback ชื่อชีตเป็น “วันหยุด” ในบางไฟล์

### `HomeroomTeachers`
- ครูประจำชั้น
- โครงสร้างขั้นต่ำ: `grade`, `classNo`, `teacherName`

## Yearly Sheets (ตามปีการศึกษา)
แนวคิด: ชีตบางตัวถูกสร้างแบบ `baseName_ปีการศึกษา` เช่น `SCORES_WAREHOUSE_2568` เพื่อรองรับหลายปี

รายชื่อ base ที่ใช้บ่อย (ดู `S_YEARLY_SHEETS` ใน [settings_unified.js](settings_unified.js) และ fallback ใน [Installer.js](Installer.js)):
- `SCORES_WAREHOUSE`
- `การประเมินอ่านคิดเขียน`
- `การประเมินคุณลักษณะ`
- `การประเมินกิจกรรมพัฒนาผู้เรียน`
- `การประเมินสมรรถนะ`
- `AttendanceLog`
- `ความเห็นครู`

## Contract Rules (อย่าฝ่าถ้าไม่ refactor ทั้งระบบ)
- ห้ามเปลี่ยนชื่อชีตหลักโดยไม่ไล่แก้ทุก `getSheetByName(...)`
- ห้ามลบคอลัมน์ที่ระบบใช้เป็น key:
  - `Users.username`, `Users.role`
  - `Students.student_id`, `Students.grade`, `Students.class_no`
  - `รายวิชา.ชั้น`, `รายวิชา.รหัสวิชา`
- ถ้าต้องเพิ่มคอลัมน์: เพิ่มด้านท้ายได้ปลอดภัยที่สุด
- ถ้าต้อง rename:
  - แก้ `setupSheetHeaders_()` ใน [Installer.js](Installer.js)
  - ไล่แก้ทุกจุดที่อ่านด้วย header index/ชื่อคอลัมน์
  - ทดสอบหน้า: login, รายชื่อนักเรียน, คะแนน, รายงาน, parent portal

