/**************************************************
 * 📄 Subject Report PDF - เวอร์ชันสุดท้าย แก้ไขแล้ว
 * - ใช้ getGlobalSettingsFixed() เพื่อ bypass ปัญหาใน settings.gs
 * - เพิ่มคอลัมน์รหัสนักเรียน และปรับ margin
 * - แสดงข้อมูลจริง: 
 **************************************************/

const REPORTS_FOLDER_NAME = 'Subject Reports';

/* =========================
   🔧 แก้ไขปัญหา settings.gs
   ========================= */
function getGlobalSettingsFixed() {
  console.log('🔧 ใช้ getGlobalSettingsFixed() เพื่อ bypass ปัญหา settings.gs');
  
  try {
    // เรียก _readKV_ โดยตรงเพื่อ bypass ปัญหา
    const result = _readKV_('global_settings');
    if (result.exists && Object.keys(result.map).length > 0) {
      console.log('✅ ได้ข้อมูลจาก _readKV_:', Object.keys(result.map).length, 'คีย์');
      console.log('👨‍💼 ชื่อผู้อำนวยการ:', result.map['ชื่อผู้อำนวยการ']);
      return result.map; // คืนข้อมูลดิบจาก _readKV_
    }
    
    console.warn('⚠️ ไม่พบข้อมูล จะใช้ค่าเริ่มต้น');
    return getDefaultSettingsForSubject();
    
  } catch (err) {
    console.error('❌ getGlobalSettingsFixed:', err);
    return getDefaultSettingsForSubject();
  }
}

function getDefaultSettingsForSubject() {
  const now = new Date();
  const thaiYear = (now.getMonth() + 1) >= 5 ? now.getFullYear() + 543 : now.getFullYear() + 542;
  
  return {
    'ชื่อโรงเรียน': 'โรงเรียน',
    'ที่อยู่โรงเรียน': '',
    'ปีการศึกษา': String(thaiYear),
    'ภาคเรียน': '1',
    'ชื่อผู้อำนวยการ': 'ผู้อำนวยการ',
    'ตำแหน่งผู้อำนวยการ': 'ผู้อำนวยการสถานศึกษา',
    'ชื่อหัวหน้างานวิชาการ': 'หัวหน้างานวิชาการ',
    'ชื่อครูผู้สอน': '',
    'หมายเหตุท้ายรายงาน': '',
    'logoFileId': '',
    'pdfSaveFolderId': ''
  };
}

/* =========================
   Utils
   ========================= */

function _getFolderByIdSafe_(id) {
  try {
    if (!id) return null;
    return DriveApp.getFolderById(id);
  } catch (e) {
    console.warn('⚠️ ใช้ pdfSaveFolderId ไม่ได้, จะสร้าง/ใช้โฟลเดอร์ชื่อ:', REPORTS_FOLDER_NAME);
    return null;
  }
}

function _safe(v, fallback = '') {
  return (v === null || v === undefined) ? fallback : v;
}

function _calculateLetterGradeFromScore_(total) {
  const t = Number(total) || 0;
  if (t >= 80) return '4';
  if (t >= 75) return '3.5';
  if (t >= 70) return '3';
  if (t >= 65) return '2.5';
  if (t >= 60) return '2';
  if (t >= 55) return '1.5';
  if (t >= 50) return '1';
  return '0';
}

function _formatFileName_(subjectName, grade, classNo, term) {
  const termText = term === '1' ? 'ภาคเรียนที่1' :
                   term === '2' ? 'ภาคเรียนที่2' : 'ทั้งสองภาคเรียน';
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm");
  const safeName = String(_safe(subjectName, 'รายวิชา')).trim();
  const safeGrade = String(_safe(grade, 'ป1')).replace('.', '').trim();
  const safeClass = String(_safe(classNo, '1')).trim();
  return `รายงาน_${safeName}_${safeGrade}_${safeClass}_${termText}_${timestamp}.pdf`;
}

/* =========================
   การดึงข้อมูลครูและคะแนน
   ========================= */
function getTeacherFromSubjectFixed(subjectName, grade, subjectCode) {
  try {
    const ss = _openSpreadsheet_(); // ใช้จาก settings.gs
    const sheet = ss.getSheetByName("รายวิชา");
    if (!sheet) return "ครูผู้สอน";

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowGrade = String(data[i][0] || "").trim();
      const rowCode = String(data[i][1] || "").trim();
      const rowName = String(data[i][2] || "").trim();
      const rowTeacher = String(data[i][5] || "").trim();
      if (rowGrade === grade &&
         (rowName === subjectName || (subjectCode && rowCode === subjectCode)) &&
          rowTeacher) {
        return rowTeacher;
      }
    }
    return "ครูผู้สอน";
  } catch (err) {
    console.error('❌ getTeacherFromSubjectFixed:', err);
    return "ครูผู้สอน";
  }
}

/**
 * ✅ แก้ไขปัญหา: เพิ่มการจัดการกรณีที่ไม่พบชีตที่ตรงกัน
 * - แสดงชีตที่มีพร้อมคำแนะนำ
 * - เสนอชีตใกล้เคียง
 */


/**
 * ✅ ปรับปรุงการดึงข้อมูลคะแนน - เพิ่ม Debug
 */
function _extractScoreDataFixed_(sheet, term) {
  try {
    const sheetName = sheet.getName();
    console.log(`📊 กำลังดึงข้อมูลจากชีต: "${sheetName}"`);
    
    const data = sheet.getDataRange().getValues();
    console.log(`📏 ชีตมี ${data.length} แถว, ${data[0] ? data[0].length : 0} คอลัมน์`);
    
    if (data.length < 5) {
      return [{ error: `❌ ข้อมูลในชีต "${sheetName}" ยังไม่พร้อม (มีแค่ ${data.length} แถว, ต้องการอย่างน้อย 5 แถว)` }];
    }

    // 🔍 Debug: ดูโครงสร้างข้อมูล
    console.log('🔍 Debug ข้อมูลในชีต:');
    console.log(`  แถวที่ 0 (header): [${data[0].slice(0, 10).map(v => `"${v}"`).join(', ')}...]`);
    console.log(`  แถวที่ 1: [${data[1].slice(0, 10).map(v => `"${v}"`).join(', ')}...]`);
    console.log(`  แถวที่ 2: [${data[2].slice(0, 10).map(v => `"${v}"`).join(', ')}...]`);
    console.log(`  แถวที่ 3: [${data[3].slice(0, 10).map(v => `"${v}"`).join(', ')}...]`);
    if (data.length > 4) {
      console.log(`  แถวที่ 4 (นักเรียนคนแรก): [${data[4].slice(0, 10).map(v => `"${v}"`).join(', ')}...]`);
    }
    
    const students = [];
    
    // ✅ ลองหาแถวเริ่มต้นข้อมูลนักเรียน (flexible)
    let studentStartRow = 4; // ค่าเริ่มต้น
    
    // หาแถวที่มีข้อมูลนักเรียน (แถวที่มี ที่, รหัส, ชื่อ)
    for (let r = 3; r < Math.min(data.length, 8); r++) {
      const row = data[r];
      // ถ้าแถวมีข้อมูลในคอลัมน์ที่ 0, 1, 2 (ที่, รหัส, ชื่อ)
      if (row[0] && row[1] && row[2]) {
        studentStartRow = r;
        console.log(`✅ พบข้อมูลนักเรียนเริ่มที่แถว: ${r}`);
        break;
      }
    }
    
    // ดึงข้อมูลนักเรียน
    for (let r = studentStartRow; r < data.length; r++) {
      const row = data[r];
      
      // ข้าม แถวว่าง
      if (!row[1] || String(row[1]).trim() === '') continue;
      
      console.log(`🔍 แถว ${r}: ที่="${row[0]}", รหัส="${row[1]}", ชื่อ="${row[2]}"`);

      // ✅ ตรวจสอบข้อมูลคะแนนในคอลัมน์ต่างๆ
      console.log(`    ภาค1: ระหว่าง[13]="${row[13]}", ปลาย[14]="${row[14]}", รวม[15]="${row[15]}", เกรด[16]="${row[16]}"`);
      console.log(`    ภาค2: ระหว่าง[27]="${row[27]}", ปลาย[28]="${row[28]}", รวม[29]="${row[29]}", เกรด[30]="${row[30]}"`);

      const term1 = {
        score70: Number(row[13]) || 0,
        final30: Number(row[14]) || 0,
        total:    Number(row[15]) || 0,
        grade:    String(_safe(row[16], '0'))
      };
      const term2 = {
        score70: Number(row[27]) || 0,
        final30: Number(row[28]) || 0,
        total:    Number(row[29]) || 0,
        grade:    String(_safe(row[30], '0'))
      };

      const term1Total = term1.total;
      const term2Total = term2.total;
      const avg = (term1Total > 0 && term2Total > 0)
        ? Math.round((term1Total + term2Total) / 2)
        : Math.max(term1Total, term2Total);

      const student = {
        no: row[0] || (students.length + 1),
        studentId: String(row[1]).trim(),
        name: String(_safe(row[2], '')).trim(),
        term1, term2,
        totalScore: avg,
        gpa: _calculateLetterGradeFromScore_(avg)
      };
      
      students.push(student);
      console.log(`✅ เพิ่มนักเรียน: ${student.name} (${student.studentId})`);
    }

    console.log(`✅ ดึงข้อมูลได้ทั้งหมด ${students.length} คน จากชีต "${sheetName}"`);
    
    if (students.length === 0) {
      return [{ error: `❌ ไม่พบข้อมูลนักเรียนในชีต "${sheetName}" (เริ่มหาจากแถว ${studentStartRow})` }];
    }
    
    return students;
    
  } catch (err) {
    console.error('❌ _extractScoreDataFixed_ error:', err);
    return [{ error: `ไม่สามารถดึงข้อมูลคะแนนได้: ${err.message}` }];
  }
}

function _mapFullGradeName_(shortGrade) {
  // กำหนด mapping ของชั้น
  const map = {
    "ป.1": "ประถมศึกษาปีที่ 1",
    "ป.2": "ประถมศึกษาปีที่ 2",
    "ป.3": "ประถมศึกษาปีที่ 3",
    "ป.4": "ประถมศึกษาปีที่ 4",
    "ป.5": "ประถมศึกษาปีที่ 5",
    "ป.6": "ประถมศึกษาปีที่ 6",
  };

  return map[shortGrade] || shortGrade; // ถ้าไม่ตรง ให้คืนค่าดั้งเดิม
}

/* =========================
   HTML Builder (ใช้ข้อมูลจริง)
   ========================= */
function _buildReportHTML_(settings, scoreData, info) {
  const safeSettings = {
    schoolName: settings['ชื่อโรงเรียน'] || "โรงเรียน",
    schoolAddress: settings['ที่อยู่โรงเรียน'] || "",
    academicYear: settings['ปีการศึกษา'] || "2568",
    directorName: settings['ชื่อผู้อำนวยการ'] || "ผู้อำนวยการ",
    directorPosition: settings['ตำแหน่งผู้อำนวยการ'] || "ผู้อำนวยการสถานศึกษา",
    academicHead: settings['ชื่อหัวหน้างานวิชาการ'] || "หัวหน้างานวิชาการ",
    footerNote: settings['หมายเหตุท้ายรายงาน'] || ""
  };

  const safeInfo = {
    subjectName: info.subjectName || "รายวิชา",
    subjectCode: info.subjectCode || "",
    grade: info.grade || "ป.1",
    classNo: info.classNo || "1",
    term: info.term || "both",
    teacherName: info.teacherName || settings['ชื่อครูผู้สอน'] || "ครูผู้สอน"
  };

  const termText = safeInfo.term === '1' ? 'ภาคเรียนที่ 1'
                 : safeInfo.term === '2' ? 'ภาคเรียนที่ 2'
                 : 'ภาคเรียนที่ 1 และ 2';

  let thead = '', rows = '';
  if (safeInfo.term === 'both') {
    // ===== รวมสองภาคเรียน =====
    thead = `
      <tr>
        <th rowspan="2" class="col-no">ที่</th>
        <th rowspan="2" class="col-name">ชื่อ-นามสกุล</th>
        <th colspan="4" class="term-header">ภาคเรียนที่ 1</th>
        <th colspan="4" class="term-header group-sep">ภาคเรียนที่ 2</th>
        <th rowspan="2" class="col-total group-sep">รวมสองภาค</th>
        <th rowspan="2" class="col-grade">ผลการประเมิน</th>
      </tr>
      <tr>
        <th class="sub-header">ระหว่างภาค</th>
        <th class="sub-header">ปลายภาค</th>
        <th class="sub-header">รวม</th>
        <th class="sub-header thick-r">เกรด</th>
        <th class="sub-header">ระหว่างภาค</th>
        <th class="sub-header">ปลายภาค</th>
        <th class="sub-header">รวม</th>
        <th class="sub-header thick-r">เกรด</th>
      </tr>`;

    rows = scoreData.map((st, i) => {
      const t1 = st.term1 || {}, t2 = st.term2 || {};
      const t1total = Number(t1.total) || 0, t2total = Number(t2.total) || 0;
      const avg = (t1total > 0 && t2total > 0) ? Math.round((t1total + t2total)/2) : Math.max(t1total, t2total);
      const finalGrade = _calculateLetterGradeFromScore_(avg);
      return `
        <tr>
          <td class="text-center">${i + 1}</td>
          <td class="student-name-both" title="${_safe(st.name)}">${_safe(st.name)}</td>

          <td class="score-cell">${_safe(t1.score70, '')}</td>
          <td class="score-cell">${_safe(t1.final30, '')}</td>
          <td class="score-cell total-cell">${_safe(t1.total, '')}</td>
          <td class="grade-cell thick-r">${_safe(t1.grade, '')}</td>

          <td class="score-cell">${_safe(t2.score70, '')}</td>
          <td class="score-cell">${_safe(t2.final30, '')}</td>
          <td class="score-cell total-cell">${_safe(t2.total, '')}</td>
          <td class="grade-cell thick-r">${_safe(t2.grade, '')}</td>

          <td class="final-total thick-r"><b>${avg || ''}</b></td>
          <td class="final-grade"><b>${finalGrade}</b></td>
        </tr>`;
    }).join('');
  } else {
    // ===== ภาคเรียนเดียว (term 1 / term 2) =====
    const key = safeInfo.term === '1' ? 'term1' : 'term2';
    thead = `
      <tr>
        <th>ที่</th>
        <th>รหัสนักเรียน</th>
        <th>ชื่อ-นามสกุล</th>
        <th>คะแนนระหว่างภาค</th>
        <th>คะแนนปลายภาค</th>
        <th>รวม</th>
        <th>ผลการประเมิน</th>
      </tr>`;
    rows = scoreData.map((st, i) => {
      const td = st[key] || {};
      return `
        <tr>
          <td>${i + 1}</td>
          <td class="student-id">${_safe(st.studentId)}</td>
          <td class="student-name">${_safe(st.name)}</td>
          <td>${_safe(td.score70, '')}</td>
          <td>${_safe(td.final30, '')}</td>
          <td>${_safe(td.total, '')}</td>
          <td>${_safe(td.grade, '')}</td>
        </tr>`;
    }).join('');
  }

  const currentDate = Utilities.formatDate(new Date(), "Asia/Bangkok", "d MMMM yyyy");
  const isBoth = safeInfo.term === 'both';
  const tableClass = isBoth ? "both-terms-table" : "report-table single-term-table";
  const emptyColspan = isBoth ? "12" : "7";

  return `
<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<style>
  @page { size: A4 portrait; margin: 12mm 12mm 14mm 12mm; }
  * { box-sizing: border-box; }
  body {
    font-family: 'TH Sarabun New','Sarabun',Arial,sans-serif;
    font-size: 12pt; color: #000; background: #fff; margin: 0;
    line-height: 1.25;
  }

  /* ===== Header (ตั้งค่า base ปกติ) ===== */
  .header { text-align: center; margin-top: 2mm; margin-bottom: 6mm; }
  .header .line { margin: 1px 0; }

  table { width: 100%; border-collapse: collapse; margin: 4mm auto; font-size: 10pt; table-layout: fixed; }
  thead { display: table-header-group; }
  tfoot { display: table-row-group; }
  tr { page-break-inside: avoid; }
  th, td {
    border: 0.6px solid #000; padding: 3px 2px; text-align: center; vertical-align: middle;
    word-wrap: break-word; overflow: hidden;
  }
  th { background: #f2f2f2; font-size: 9pt; font-weight: bold; }
  tbody tr:nth-child(even) td { background: #fcfcfc; }

  /* ===== เฉพาะภาคเรียนเดียว: ขยับตารางไปทางขวา ===== */
  .report-table { width: 92%; margin-left: auto; }

  /* ===== ภาคเรียนเดียว: สัดส่วนคอลัมน์ ===== */
  .single-term-table .student-id { font-family: 'Courier New', monospace; font-size: 9pt; text-align: center; }
  .single-term-table .student-name { text-align: left !important; padding-left: 6px !important; font-size: 10pt; }

  .single-term-table th:nth-child(1), .single-term-table td:nth-child(1) { width: 5%; }   /* ที่ */
  .single-term-table th:nth-child(2), .single-term-table td:nth-child(2) { width: 10%; }  /* รหัส */
  .single-term-table th:nth-child(3), .single-term-table td:nth-child(3) { width: 40%; }  /* ชื่อ-นามสกุล */
  .single-term-table th:nth-child(4), .single-term-table td:nth-child(4) { width: 11%; }  /* ระหว่างภาค */
  .single-term-table th:nth-child(5), .single-term-table td:nth-child(5) { width: 11%; }  /* ปลายภาค */
  .single-term-table th:nth-child(6), .single-term-table td:nth-child(6) { width: 11%; }  /* รวม */
  .single-term-table th:nth-child(7), .single-term-table td:nth-child(7) { width: 12%; }  /* เกรด */

  /* ===== รวมสองภาคเรียน ===== */
  .both-terms-table .col-no   { width: 6%; }
  .both-terms-table .col-name { width: 24%; text-align: center; padding-left: 6px; }
  .both-terms-table .col-total{ width: 8%; }
  .both-terms-table .col-grade{ width: 8%; }
  .both-terms-table .sub-header { width: 7%; }
  .both-terms-table .term-header { background: #eaeaea; }
  .student-name-both { text-align: left !important; padding-left: 6px !important; white-space: normal; line-height: 1.2; }
  .thick-r, .group-sep { border-right: 1.2px solid #000 !important; }
  .total-cell { background: #fafafa; font-weight: bold; }
  .final-total, .final-grade { background: #e8f6e8; font-weight: bold; }

  .sign { width: 100%; margin-top: 10mm; max-width: 100%; }
  .sign td { border: none; text-align: center; width: 50%; padding: 6mm 0 0; }
  .line-sign { width: 170px; height: 0; border-bottom: 1px solid #000; margin: 8px auto 4px; }
  .date { text-align: right; font-size: 10pt; margin-top: 5mm; margin-right: 6mm; }
  .footer-note { text-align: center; font-size: 9pt; margin-top: 3mm; font-style: italic; color: #555; }

  @media print { body { padding: 0 2mm; } }

  /* ===== FORCE: ให้หัวรายงานทุกบรรทัด = 14pt เท่ากัน (ทับทุกกฎก่อนหน้า) ===== */
/* ตั้งค่าทุกบรรทัดให้เหมือนกันทุกประการ */
.header .line {
  font-size: 12pt !important;
  line-height: 1.4 !important;
  font-weight: normal !important;  /* ทุกบรรทัดใช้ normal เท่ากัน */
  margin: 3px 0 !important;        /* เพิ่มระยะห่างให้เหมาะสม */
  font-family: 'TH Sarabun New','Sarabun',Arial,sans-serif !important;
}
</style>
</head>
<body>
  <div class="header">
    <div class="line school">${safeSettings.schoolName}</div>
    <div class="line title">
      รายงานผลการเรียนรายวิชา ${safeInfo.subjectName}${safeInfo.subjectCode ? ' (' + safeInfo.subjectCode + ')' : ''}
    </div>
    <div class="line">ระดับชั้น ${_mapFullGradeName_(safeInfo.grade)} ห้อง ${safeInfo.classNo} ${termText}</div>
    <div class="line">ปีการศึกษา ${safeSettings.academicYear}</div>
  </div>

  <table class="${tableClass}">
    <thead>${thead}</thead>
    <tbody>
      ${rows || `<tr><td colspan="${emptyColspan}" style="text-align:center">ไม่พบข้อมูลคะแนนสำหรับเงื่อนไขที่เลือก</td></tr>`}
    </tbody>
  </table>

  <table class="sign">
    <tr>
      <td>
        <div class="line-sign"></div>
        ( ${safeInfo.teacherName} )<br/>ครูผู้สอน
      </td>
      <td>
        <div class="line-sign"></div>
        ( ${safeSettings.directorName} )<br/>${safeSettings.directorPosition}
      </td>
    </tr>
  </table>

  ${safeSettings.footerNote && safeSettings.footerNote !== '-' ? `<div class="footer-note">${safeSettings.footerNote}</div>` : ''}
  <div class="date">วันที่พิมพ์: ${currentDate}</div>
</body>
</html>`;
}
/**
 * ==========================================
 * ฟังก์ชันหลักที่หน้าเว็บเรียกใช้ (แก้ไขแล้ว)
 * ==========================================
 */

/**
 * 🎯 ฟังก์ชันหลักที่หน้าเว็บเรียกใช้ - สร้าง PDF รายงานรายวิชา
 * แทนที่ generateSubjectPDFByTerm() เดิม
 */


/**
 * ✅ สร้าง PDF รายงานรายวิชา (เวอร์ชันแก้ไข - ไม่ใช้ template file)
 */
function generateSubjectScorePDFFixed(subjectName, subjectCode, grade, classNo, term = 'both') {
  try {
    console.log(`🔨 generateSubjectScorePDFFixed: ${subjectName}, ${grade}/${classNo}, term: ${term}`);
    
    if (!subjectName || !grade || !classNo) {
      throw new Error('กรุณาระบุข้อมูลที่จำเป็น: ชื่อวิชา, ระดับชั้น, ห้อง');
    }

    // ดึงการตั้งค่าระบบ
    const settings = S_getGlobalSettings() || {};
    const schoolName = settings.schoolName || settings['ชื่อโรงเรียน'] || 'โรงเรียนของเรา';
    const academicYear = settings.academicYear || settings['ปีการศึกษา'] || '2568';
    
    // ดึงโลโก้
    let logoDataUrl = '';
    try {
      logoDataUrl = getLogoAsBase64() || '';
    } catch (e) {
      console.warn('ไม่สามารถโหลดโลโก้ได้:', e.message);
    }

    // ดึงข้อมูลคะแนน
    const scoreData = getSubjectScoreTableFixed(subjectName, subjectCode, grade, classNo, term);
    
    if (!scoreData || !scoreData.students || scoreData.students.length === 0) {
      throw new Error('ไม่พบข้อมูลคะแนนสำหรับรายวิชานี้');
    }

    // สร้าง HTML สำหรับ PDF
    const htmlContent = createSubjectScorePDFHTML({
      schoolName,
      academicYear,
      logoDataUrl,
      subjectName,
      subjectCode: subjectCode || '',
      grade,
      classNo,
      term,
      scoreData,
      gradeFullName: U_getGradeFullName(grade)
    });

    // แปลง HTML เป็น PDF
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html')
                          .getAs('application/pdf');

    // สร้างชื่อไฟล์
    const safeSubj = String(subjectName).replace(/[\\/:*?"<>|]/g, '_');
    const safeGrade = String(grade).replace(/\./g, '');
    const termText = term === 'both' ? 'ทั้งปี' : `ภาค${term}`;
    const fileName = `ผลการเรียน_${safeSubj}_${safeGrade}-${classNo}_${termText}.pdf`;
    
    blob.setName(fileName);

    // บันทึกไฟล์
    const folderId = settings.pdfSaveFolderId || DriveApp.getRootFolder().getId();
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log(`✅ สร้างไฟล์ PDF: ${fileName}`);
    return file.getUrl();
    
  } catch (error) {
    console.error('❌ Error in generateSubjectScorePDFFixed:', error);
    throw new Error(`สร้าง PDF ล้มเหลว: ${error.message}`);
  }
}

/**
 * ✅ ดึงข้อมูลคะแนนจากชีตรายวิชา (แก้ไขแล้ว)
 */


/**
 * ✅ ดึงข้อมูลคะแนนแต่ละภาคเรียน
 */
function extractTermData(row, startCol, endCol) {
  const scores = [];
  
  // คะแนน 10 ช่อง
  for (let i = startCol; i < startCol + 10; i++) {
    scores.push(Number(row[i]) || 0);
  }
  
  return {
    scores: scores,                      // คะแนน 1-10
    midterm: Number(row[startCol + 10]) || 0,  // คะแนนระหว่างภาค
    final: Number(row[startCol + 11]) || 0,    // คะแนนปลายภาค
    total: Number(row[startCol + 12]) || 0,    // รวมคะแนน
    grade: String(row[startCol + 13] || '')    // เกรด
  };
}

/**
 * ✅ สร้าง HTML สำหรับ PDF รายงานรายวิชา
 */
function createSubjectScorePDFHTML(data) {
  const {
    schoolName, academicYear, logoDataUrl, subjectName, subjectCode, 
    grade, classNo, term, scoreData, gradeFullName
  } = data;
  
  // ส่วนหัว
  const logoHtml = logoDataUrl ? 
    `<div class="logo"><img src="${logoDataUrl}" alt="โลโก้" /></div>` : '';
  
  const termText = term === 'both' ? 'ทั้งปีการศึกษา' : 
                   term === '1' ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
  
  // สร้างตารางนักเรียน
  const studentRows = scoreData.students.map((student, index) => {
    let scoreColumns = '';
    
    if (term === 'both') {
      // แสดงทั้งสองภาค
      scoreColumns = `
        <td>${student.term1.total}</td>
        <td>${student.term1.grade}</td>
        <td>${student.term2.total}</td>
        <td>${student.term2.grade}</td>
        <td><strong>${student.average}</strong></td>
        <td><strong>${student.finalGrade}</strong></td>
      `;
    } else if (term === '1') {
      // แสดงเฉพาะภาค 1
      scoreColumns = `
        <td>${student.term1.midterm}</td>
        <td>${student.term1.final}</td>
        <td><strong>${student.term1.total}</strong></td>
        <td><strong>${student.term1.grade}</strong></td>
      `;
    } else {
      // แสดงเฉพาะภาค 2
      scoreColumns = `
        <td>${student.term2.midterm}</td>
        <td>${student.term2.final}</td>
        <td><strong>${student.term2.total}</strong></td>
        <td><strong>${student.term2.grade}</strong></td>
      `;
    }
    
    return `
      <tr>
        <td class="center">${index + 1}</td>
        <td class="center">${student.id}</td>
        <td class="left">${student.name}</td>
        ${scoreColumns}
      </tr>
    `;
  }).join('');
  
  // ส่วนหัวตาราง
  let tableHeader = '';
  if (term === 'both') {
    tableHeader = `
      <tr>
        <th rowspan="2">ที่</th>
        <th rowspan="2">รหัส</th>
        <th rowspan="2">ชื่อ - นามสกุล</th>
        <th colspan="2">ภาคเรียนที่ 1</th>
        <th colspan="2">ภาคเรียนที่ 2</th>
        <th rowspan="2">คะแนนเฉลี่ย</th>
        <th rowspan="2">เกรด</th>
      </tr>
      <tr>
        <th>คะแนน</th><th>เกรด</th>
        <th>คะแนน</th><th>เกรด</th>
      </tr>
    `;
  } else {
    tableHeader = `
      <tr>
        <th>ที่</th>
        <th>รหัส</th>
        <th>ชื่อ - นามสกุล</th>
        <th>ระหว่างภาค</th>
        <th>ปลายภาค</th>
        <th>รวม</th>
        <th>เกรด</th>
      </tr>
    `;
  }
  
  return `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>รายงานผลการเรียน</title>
      <style>
        @page { size: A4; margin: 1cm; }
        body {
          font-family: 'TH Sarabun New', 'Sarabun', sans-serif;
          font-size: 14pt;
          line-height: 1.4;
          margin: 0;
          padding: 0;
        }
        .header {
          text-align: center;
          margin-bottom: 20px;
          border-bottom: 2px solid #333;
          padding-bottom: 15px;
        }
        .logo img {
          max-height: 60px;
          margin-bottom: 10px;
        }
        .school-name {
          font-size: 18pt;
          font-weight: bold;
          margin-bottom: 5px;
        }
        .report-title {
          font-size: 16pt;
          font-weight: bold;
          margin-bottom: 10px;
        }
        .subject-info {
          font-size: 14pt;
          margin-bottom: 5px;
        }
        .info-row {
          display: flex;
          justify-content: space-between;
          margin-bottom: 15px;
          font-size: 12pt;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 10px;
        }
        th, td {
          border: 1px solid #333;
          padding: 8px 4px;
          text-align: center;
          vertical-align: middle;
        }
        th {
          background-color: #f0f0f0;
          font-weight: bold;
        }
        .left { text-align: left; padding-left: 8px; }
        .center { text-align: center; }
        .signature {
          margin-top: 30px;
          display: flex;
          justify-content: space-between;
        }
        .signature-box {
          text-align: center;
          width: 200px;
        }
        .signature-line {
          border-bottom: 1px solid #333;
          margin-bottom: 5px;
          height: 40px;
        }
      </style>
    </head>
    <body>
      <div class="header">
        ${logoHtml}
        <div class="school-name">${schoolName}</div>
        <div class="report-title">รายงานผลการเรียน ${termText}</div>
        <div class="subject-info">รายวิชา: ${subjectName} ${subjectCode ? `(${subjectCode})` : ''}</div>
        <div class="info-row">
          <span>ชั้น: ${gradeFullName} ห้อง ${classNo}</span>
          <span>ปีการศึกษา: ${academicYear}</span>
        </div>
      </div>
      
      <table>
        <thead>${tableHeader}</thead>
        <tbody>${studentRows}</tbody>
      </table>
      
      <div class="signature">
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>ผู้สอน</div>
        </div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>หัวหน้าสถานศึกษา</div>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * ==========================================
 * ฟังก์ชันสำหรับรายงานรายบุคคล (ปพ.6)
 * ==========================================
 */

/**
 * ✅ ดึงรายชื่อนักเรียนสำหรับรายงาน ปพ.6
 */


/**
 * ✅ สร้าง PDF รายงานรายบุคคล (ปพ.6) แบบสมบูรณ์
 */

/**
 * ✅ รวบรวมข้อมูลครบถ้วนสำหรับ ปพ.6
 */
function getPp6ReportDataComplete(studentId, term = 'both') {
  try {
    console.log(`📊 รวบรวมข้อมูล ปพ.6 สำหรับ: ${studentId}`);
    
    const ss = SS();
    
    // 1. ข้อมูลส่วนตัวนักเรียน
    const studentInfo = getStudentPersonalInfo(studentId);
    if (!studentInfo) {
      throw new Error('ไม่พบข้อมูลนักเรียน');
    }
    
    // 2. ผลการเรียนทุกรายวิชา
    const academicResults = getStudentAllSubjectResults(studentId, term);
    
    // 3. การเข้าเรียน
    const attendance = getStudentAttendanceSummary(studentId);
    
    // 4. กิจกรรม/คุณลักษณะ (ถ้ามี)
    const activities = getStudentActivities(studentId) || [];
    const characteristics = getStudentCharacteristics(studentId) || {};
    
    // 5. การตั้งค่าระบบ
    const settings = S_getGlobalSettings() || {};
    
    return {
      student: studentInfo,
      academic: academicResults,
      attendance: attendance,
      activities: activities,
      characteristics: characteristics,
      settings: settings,
      term: term,
      generatedDate: new Date()
    };
    
  } catch (error) {
    console.error('❌ Error in getPp6ReportDataComplete:', error);
    throw new Error(`ไม่สามารถรวบรวมข้อมูลได้: ${error.message}`);
  }
}

/**
 * ✅ ดึงข้อมูลส่วนตัวนักเรียน
 */
function getStudentPersonalInfo(studentId) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('Students');
    
    if (!sheet) {
      throw new Error('ไม่พบชีต Students');
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (String(row[0]).trim() === String(studentId).trim()) {
        return {
          id: row[0] || '',
          idCard: row[1] || '',
          title: row[2] || '',
          firstname: row[3] || '',
          lastname: row[4] || '',
          name: `${row[2] || ''}${row[3] || ''} ${row[4] || ''}`.trim(),
          grade: row[5] || '',
          class: row[6] || '',
          gender: row[7] || '',
          birthdate: row[8] || '',
          address: row[9] || '',
          fatherName: `${row[10] || ''} ${row[11] || ''}`.trim(),
          motherName: `${row[12] || ''} ${row[13] || ''}`.trim(),
          phone: row[14] || ''
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('❌ Error in getStudentPersonalInfo:', error);
    throw new Error('ไม่สามารถดึงข้อมูลส่วนตัวนักเรียนได้');
  }
}

/**
 * ✅ ดึงผลการเรียนทุกรายวิชา
 */
function getStudentAllSubjectResults(studentId, term) {
  try {
    const ss = SS();
    const allSheets = ss.getSheets();
    const subjects = [];
    
    // หาชีตที่เป็นรายวิชา (ไม่รวมชีตระบบ)
    const excludeSheets = ['Students', 'รายวิชา', 'global_settings', 'การตั้งค่าระบบ', 'การประเมินคุณลักษณะ', 'การประเมินอ่านคิดเขียน', 'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ', 'Users', 'users', 'Holidays', 'วันหยุด', 'HomeroomTeachers', 'SCORES_WAREHOUSE', 'AttendanceLog', 'ความเห็นครู', 'โปรไฟล์นักเรียน', 'BACKUP_WAREHOUSE_LATEST', 'Template_', 'สรุปวันมา', 'สรุปการมาเรียน', 'อ่านคิดเขียน', 'คุณลักษณะ', 'กิจกรรม', 'สมรรถนะ'];
    const scoreSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return !excludeSheets.some(ex => name === ex || name.startsWith(ex)) &&
             !name.includes('2567') && !name.includes('2568') && !name.includes('2569');
    });
    
    let totalGradePoints = 0;
    let totalCredits = 0;
    
    scoreSheets.forEach(sheet => {
      try {
        const data = sheet.getDataRange().getValues();
        if (data.length < 5) return;
        
        // หาข้อมูลนักเรียน
        for (let i = 4; i < data.length; i++) {
          if (String(data[i][1]).trim() === String(studentId).trim()) {
            const subjectName = data[1][1] || sheet.getName();
            const subjectCode = data[0][1] || '';
            
            // ดึงคะแนนตาม term ที่เลือก
            let displayScore = 0;
            let displayGrade = '';
            
            if (term === 'both') {
              displayScore = Number(data[i][31]) || 0; // คะแนนเฉลี่ย
              displayGrade = String(data[i][32] || ''); // เกรดสุดท้าย
            } else if (term === '1') {
              displayScore = Number(data[i][15]) || 0; // รวมภาค 1
              displayGrade = String(data[i][16] || ''); // เกรดภาค 1
            } else {
              displayScore = Number(data[i][29]) || 0; // รวมภาค 2
              displayGrade = String(data[i][30] || ''); // เกรดภาค 2
            }
            
            subjects.push({
              name: subjectName,
              code: subjectCode,
              score: displayScore,
              grade: displayGrade,
              credits: 1 // สมมติ 1 หน่วยกิต
            });
            
            // คำนวณ GPA
            const gradePoint = parseFloat(displayGrade) || 0;
            if (gradePoint > 0) {
              totalGradePoints += gradePoint;
              totalCredits += 1;
            }
            
            break;
          }
        }
      } catch (e) {
        console.log(`ข้ามชีต ${sheet.getName()}: ${e.message}`);
      }
    });
    
    const gpa = totalCredits > 0 ? (totalGradePoints / totalCredits).toFixed(2) : '0.00';
    
    return {
      subjects: subjects,
      totalSubjects: subjects.length,
      gpa: gpa,
      totalCredits: totalCredits
    };
    
  } catch (error) {
    console.error('Error in getStudentAllSubjectResults:', error);
    throw new Error('ไม่สามารถดึงผลการเรียนได้');
  }
}

/**
 * ดึงข้อมูลกิจกรรม/คุณลักษณะ (สำรอง - ถ้าไม่มีก็คืนค่าว่าง)
 */
function getStudentActivities(studentId) {
  try {
    const ss = SS();
    const sheet = S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const activities = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(studentId).trim()) {
        activities.push({
          name: data[i][1] || '',
          result: data[i][2] || '',
          note: data[i][3] || ''
        });
      }
    }
    
    return activities;
  } catch (error) {
    return [];
  }
}

/**
 * ดึงข้อมูลคุณลักษณะ (สำรอง)
 */
function getStudentCharacteristics(studentId) {
  try {
    const ss = SS();
    const sheet = S_getYearlySheet('การประเมินคุณลักษณะ');
    
    if (!sheet) return {};
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(studentId).trim()) {
        return {
          behavior: data[i][1] || '',
          responsibility: data[i][2] || '',
          morality: data[i][3] || '',
          overall: data[i][4] || ''
        };
      }
    }
    
    return {};
  } catch (error) {
    return {};
  }
}

/**
 * สร้าง HTML สำหรับ PDF รายงาน ปพ.6
 */
function createPp6PDFHTML(reportData, term) {
  const { student, academic, attendance, activities, characteristics, settings } = reportData;
  
  const schoolName = settings.schoolName || settings['ชื่อโรงเรียน'] || 'โรงเรียนของเรา';
  const academicYear = settings.academicYear || settings['ปีการศึกษา'] || '2568';
  
  // ส่วนหัว
  const logoHtml = settings.logoDataUrl ? 
    `<div class="logo"><img src="${settings.logoDataUrl}" alt="โลโก้" /></div>` : '';
  
  const termText = term === 'both' ? 'ทั้งปีการศึกษา' : 
                   term === '1' ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
  
  // ตารางผลการเรียน
  const subjectRows = academic.subjects.map((subject, index) => `
    <tr>
      <td class="center">${index + 1}</td>
      <td class="left">${subject.name}</td>
      <td class="center">${subject.code}</td>
      <td class="center">${subject.score}</td>
      <td class="center">${subject.grade}</td>
    </tr>
  `).join('');
  
  // ตารางกิจกรรม
  const activityRows = activities.length > 0 ? 
    activities.map((activity, index) => `
      <tr>
        <td class="center">${index + 1}</td>
        <td class="left">${activity.name}</td>
        <td class="center">${activity.result}</td>
        <td class="left">${activity.note}</td>
      </tr>
    `).join('') : 
    '<tr><td colspan="4" class="center">ไม่มีข้อมูลกิจกรรม</td></tr>';
  
  return `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>รายงานผลการเรียนรายบุคคล (ปพ.6)</title>
      <style>
        @page { size: A4; margin: 1cm; }
        body {
          font-family: 'TH Sarabun New', 'Sarabun', sans-serif;
          font-size: 12pt;
          line-height: 1.3;
          margin: 0;
          padding: 0;
        }
        .header {
          text-align: center;
          margin-bottom: 20px;
          border-bottom: 2px solid #333;
          padding-bottom: 15px;
        }
        .logo img {
          max-height: 50px;
          margin-bottom: 8px;
        }
        .school-name {
          font-size: 16pt;
          font-weight: bold;
          margin-bottom: 5px;
        }
        .report-title {
          font-size: 14pt;
          font-weight: bold;
          margin-bottom: 10px;
        }
        .student-info {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 20px;
          margin-bottom: 20px;
          background: #f9f9f9;
          padding: 15px;
          border: 1px solid #ddd;
        }
        .info-group h3 {
          margin: 0 0 10px 0;
          font-size: 12pt;
          color: #333;
          border-bottom: 1px solid #ccc;
          padding-bottom: 5px;
        }
        .info-row {
          display: flex;
          margin-bottom: 5px;
        }
        .info-label {
          font-weight: bold;
          width: 120px;
          flex-shrink: 0;
        }
        .info-value {
          flex: 1;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin-bottom: 20px;
        }
        th, td {
          border: 1px solid #333;
          padding: 6px 4px;
          text-align: center;
          vertical-align: middle;
          font-size: 11pt;
        }
        th {
          background-color: #f0f0f0;
          font-weight: bold;
        }
        .left { text-align: left; padding-left: 6px; }
        .center { text-align: center; }
        .section-title {
          font-size: 13pt;
          font-weight: bold;
          margin: 20px 0 10px 0;
          color: #333;
          border-left: 4px solid #333;
          padding-left: 10px;
        }
        .summary-box {
          background: #f0f8ff;
          padding: 15px;
          border: 1px solid #333;
          margin: 15px 0;
        }
        .summary-row {
          display: flex;
          justify-content: space-between;
          margin-bottom: 8px;
        }
        .summary-label {
          font-weight: bold;
        }
        .signature {
          margin-top: 30px;
          display: flex;
          justify-content: space-between;
        }
        .signature-box {
          text-align: center;
          width: 200px;
        }
        .signature-line {
          border-bottom: 1px solid #333;
          margin-bottom: 5px;
          height: 40px;
        }
        .page-break {
          page-break-before: always;
        }
      </style>
    </head>
    <body>
      <div class="header">
        ${logoHtml}
        <div class="school-name">${schoolName}</div>
        <div class="report-title">รายงานผลการเรียนรายบุคคล (ปพ.6)</div>
        <div>ปีการศึกษา ${academicYear} ${termText}</div>
      </div>
      
      <div class="student-info">
        <div class="info-group">
          <h3>ข้อมูลนักเรียน</h3>
          <div class="info-row">
            <span class="info-label">รหัสประจำตัว:</span>
            <span class="info-value">${student.id}</span>
          </div>
          <div class="info-row">
            <span class="info-label">ชื่อ-นามสกุล:</span>
            <span class="info-value">${student.name}</span>
          </div>
          <div class="info-row">
            <span class="info-label">ระดับชั้น:</span>
            <span class="info-value">${U_getGradeFullName(student.grade)} ห้อง ${student.class}</span>
          </div>
          <div class="info-row">
            <span class="info-label">เพศ:</span>
            <span class="info-value">${student.gender}</span>
          </div>
        </div>
        
        <div class="info-group">
          <h3>ข้อมูลผู้ปกครอง</h3>
          <div class="info-row">
            <span class="info-label">บิดา:</span>
            <span class="info-value">${student.fatherName}</span>
          </div>
          <div class="info-row">
            <span class="info-label">มารดา:</span>
            <span class="info-value">${student.motherName}</span>
          </div>
          <div class="info-row">
            <span class="info-label">เบอร์โทร:</span>
            <span class="info-value">${student.phone}</span>
          </div>
          <div class="info-row">
            <span class="info-label">ที่อยู่:</span>
            <span class="info-value">${student.address}</span>
          </div>
        </div>
      </div>
      
      <div class="section-title">ผลการเรียนรายวิชา</div>
      <table>
        <thead>
          <tr>
            <th style="width: 5%">ที่</th>
            <th style="width: 40%">รายวิชา</th>
            <th style="width: 15%">รหัสวิชา</th>
            <th style="width: 15%">คะแนน</th>
            <th style="width: 10%">เกรด</th>
          </tr>
        </thead>
        <tbody>
          ${subjectRows}
        </tbody>
      </table>
      
      <div class="summary-box">
        <div class="summary-row">
          <span class="summary-label">จำนวนรายวิชาทั้งหมด:</span>
          <span>${academic.totalSubjects} วิชา</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">เกรดเฉลี่ยสะสม (GPA):</span>
          <span>${academic.gpa}</span>
        </div>
      </div>
      
      <div class="section-title">สรุปการเข้าเรียน</div>
      <div class="summary-box">
        <div class="summary-row">
          <span class="summary-label">มาเรียน:</span>
          <span>${attendance.present} วัน</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">มาสาย:</span>
          <span>${attendance.late} วัน</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">ลา:</span>
          <span>${attendance.leave} วัน</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">ขาด:</span>
          <span>${attendance.absent} วัน</span>
        </div>
        <div class="summary-row">
          <span class="summary-label">รวมทั้งหมด:</span>
          <span>${attendance.totalDays} วัน</span>
        </div>
      </div>
      
      <div class="section-title">กิจกรรมและคุณลักษณะ</div>
      <table>
        <thead>
          <tr>
            <th style="width: 5%">ที่</th>
            <th style="width: 50%">กิจกรรม</th>
            <th style="width: 15%">ผลการประเมิน</th>
            <th style="width: 30%">หมายเหตุ</th>
          </tr>
        </thead>
        <tbody>
          ${activityRows}
        </tbody>
      </table>
      
      <div class="signature">
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>ครูที่ปรึกษา</div>
        </div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>หัวหน้าสถานศึกษา</div>
        </div>
      </div>
      
      <div style="text-align: center; margin-top: 20px; font-size: 10pt; color: #666;">
        สร้างเมื่อ: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')}
      </div>
    </body>
    </html>
  `;
}

/**
 * ==========================================
 * ฟังก์ชันเสริมและการตั้งค่า
 * ==========================================
 */

/**
 * แปลงรหัสระดับชั้นเป็นชื่อเต็ม
 */

/**
 * ==========================================
 * ฟังก์ชัน Warehouse Tools (สำรองไว้)
 * ==========================================
 */

/**
 * รีบิลด์คลังคะแนนสำหรับห้องเฉพาะ
 */

/**
 * รีบิลด์คลังคะแนนสำหรับทั้งชั้น
 */

/**
 * รีบิลด์คลังคะแนนทั้งไฟล์
 */

