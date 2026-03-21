---
description: กฎป้องกันฟังก์ชัน PDF — ต้องทำตามเมื่อแก้โค้ดที่เกี่ยวกับ PDF
---

# ⚠️ PDF Protection Rules

ฟังก์ชัน PDF ทั้งหมดผ่านการปรับแต่ง layout เสร็จแล้ว ห้ามแก้ไขโดยไม่ทำตามขั้นตอนนี้

## กฎหลัก

1. **ห้าม rewrite ทั้งฟังก์ชัน** — แก้ไขเฉพาะจุดที่จำเป็น ใช้ edit tool แก้ทีละบรรทัด
2. **ห้ามเปลี่ยน layout config** — margin, font size, column width, padding, bold ห้ามแก้โดยไม่ได้รับอนุญาตจาก user
3. **ADD ไม่ใช่ REPLACE** — ถ้า user ขอ feature ใหม่ ให้เพิ่มโค้ดใหม่ ห้ามลบโค้ดเดิมที่ทำงานอยู่
4. **แจ้ง user ก่อน deploy** — ถ้าแก้ไฟล์ PDF ต้องบอก user ว่าแก้ไขอะไรบ้างก่อน deploy

## ไฟล์ที่ล็อค

| ไฟล์ | ฟังก์ชันหลัก | หมายเหตุ |
|------|-------------|----------|
| `OnePageReport.js` | `_opr_cell`, `_opr_generateSingle` | shared helper — PP6 ใช้ด้วย |
| `pp6.js` | `_pp6_buildDocPdf` | มี `testPp6PdfRegression()` |
| `Export_PDF.js` | `generateStudentListPDF`, `generateSubjectScorePDF`, `exportAttendancePDF` | |
| `Attendance.js` | `exportMonthlyAttendancePDF`, `exportYearlyAttendancePDF`, `createMonthlyPDFBlob`, `createChecklistPDFBlob` | |
| `Assessments.js` | `generateCharacteristicAssessmentPDF`, `generateActivityAssessmentPDF`, `generateSubjectScoreAssessmentPDF`, `exportCompetencyReport`, `exportCompetencySummaryReport` | |
| `pp5book.js` / `pp5cover.js` | PP5 เล่ม | |

## ⚠️ Shared Functions — แก้แล้วกระทบหลายรายงาน

ฟังก์ชันเหล่านี้ถูกใช้ร่วมโดยหลาย PDF ห้ามแก้โดยไม่ทดสอบรายงานทั้งหมดที่ใช้:

| ฟังก์ชัน | อยู่ใน | กระทบ |
|----------|--------|-------|
| `_opr_cell()` | `OnePageReport.js` | รายงานหน้าเดียว + ปพ.6 |
| `OPR_FONT` | `OnePageReport.js` | รายงานหน้าเดียว + ปพ.6 |
| `_fmtGrade()` | `OnePageReport.js` | รายงานหน้าเดียว + ปพ.6 |
| `getOutputFolder_()` | `Code.js` | ปพ.6 + รายงานรายวิชา |
| `savePDFToDrive_()` | `Assessments.js` | คุณลักษณะ/กิจกรรม/ปพ.5/สมรรถนะ (5 รายงาน) |
| `savePDFAndGetUrl()` | `Export_PDF.js` | รายชื่อนักเรียน |
| `getStudentInfo_()` | `ReportGenerator.js` | รายงานหน้าเดียว + ปพ.6 |
| `getHomeroomTeacher()` | `Assessments.js` | รายงานหน้าเดียว + ปพ.6 |
| `getStudentAssessments()` | `pp6.js` | ปพ.6 |
| `calculateGPAAndRank()` | `pp6.js` | ปพ.6 |
| `getTeacherComment_()` | `ReportGenerator.js` | รายงานหน้าเดียว + ปพ.6 |

## HTML Template Functions — แก้แล้ว layout PDF เปลี่ยน

| ฟังก์ชัน | อยู่ใน | สร้าง |
|----------|--------|-------|
| `createMonthlyPDFHTML()` | `Attendance.js` | PDF รายเดือน |
| `createYearlyPDFHTML()` | `Attendance.js` | PDF รายปี |
| `buildAttendancePdfHtml()` | `Export_PDF.js` | PDF เวลาเรียน |
| `createCharacteristicAssessmentHTML()` | `Assessments.js` | PDF คุณลักษณะ |
| `createActivityAssessmentHTML_()` | `Assessments.js` | PDF กิจกรรม |
| `createOfficialReportHTML()` | `Assessments.js` | PDF ปพ.5 |
| `createCompetencyDetailedHTML()` | `Assessments.js` | PDF สมรรถนะ |
| `createCompetencySummaryHTML()` | `Assessments.js` | PDF สรุปสมรรถนะ |

## ก่อนแก้โค้ด PDF ต้องทำ

1. อ่านโค้ดเดิมก่อนแก้ (ใช้ read_file)
2. แก้เฉพาะจุดที่ user ร้องขอ — ห้ามปรับ layout อื่นๆ ไปด้วย
3. ถ้าแก้ `_opr_cell`, `OPR_FONT`, `_fmtGrade` ต้องทดสอบทั้ง OnePageReport และ PP6
4. ถ้าแก้ `savePDFToDrive_` ต้องทดสอบรายงาน Assessments ทั้ง 5 ตัว
5. ถ้าแก้ `getStudentInfo_`, `getHomeroomTeacher`, `getTeacherComment_` ต้องทดสอบทั้งรายงานหน้าเดียว + ปพ.6
6. ถ้าแก้ HTML template functions ต้องทดสอบ PDF ที่เกี่ยวข้อง
7. ถ้าแก้ pp6.js ต้องรัน `testPp6PdfRegression()` หลัง deploy
8. แจ้ง user ก่อน deploy ว่าแก้อะไรบ้าง
