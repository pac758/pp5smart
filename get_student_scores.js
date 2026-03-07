// ============================================================
// 📊 STUDENT SCORES SUMMARY - ฉบับสมบูรณ์
// แสดงเกรด + ส่งออก PDF + ใช้ student_id + คำนำหน้าชื่อ + โลโก้
// ============================================================

function scoreToGrade(score) {
  if (score >= 80) return 4.0;
  if (score >= 75) return 3.5;
  if (score >= 70) return 3.0;
  if (score >= 65) return 2.5;
  if (score >= 60) return 2.0;
  if (score >= 55) return 1.5;
  if (score >= 50) return 1.0;
  return 0;
}

function convertToGrade(value) {
  const numValue = parseFloat(value) || 0;
  if (numValue >= 0 && numValue <= 4) {
    return numValue;
  }
  if (numValue > 4) {
    return scoreToGrade(numValue);
  }
  return 0;
}

function getStudentScoresForWeb(grade, classNo) {
  try {
    const ss = SS();
    const studentsSheet = ss.getSheetByName("Students");
    if (!studentsSheet) {
      throw new Error("ไม่พบชีต 'Students'");
    }
    
    const studentData = studentsSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    
    const gradeCol = U_findColumnIndex(studentHeaders, ['grade', 'ชั้น', 'ระดับชั้น']);
    const classCol = U_findColumnIndex(studentHeaders, ['class_no', 'ห้อง']);
    const studentIdCol = U_findColumnIndex(studentHeaders, ['student_id', 'รหัสนักเรียน']);
    const studentNoCol = U_findColumnIndex(studentHeaders, ['student_no', 'เลขที่', 'ลำดับ']);
    const titleCol = U_findColumnIndex(studentHeaders, ['title', 'คำนำหน้า', 'คำนำหน้าชื่อ']);
    const firstNameCol = U_findColumnIndex(studentHeaders, ['firstname', 'first_name', 'ชื่อ']);
    const lastNameCol = U_findColumnIndex(studentHeaders, ['lastname', 'last_name', 'นามสกุล']);
    const fullNameCol = U_findColumnIndex(studentHeaders, ['full_name', 'ชื่อเต็ม', 'ชื่อ-นามสกุล']);
    
    if (studentIdCol === -1) {
      throw new Error("ไม่พบคอลัมน์ 'student_id' ในชีต Students");
    }
    
    const students = [];
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      const rowGrade = String(row[gradeCol] || '').trim();
      const rowClass = String(row[classCol] || '').trim();
      
      if (rowGrade === grade && rowClass === classNo) {
        const studentId = String(row[studentIdCol] || '').trim();
        const studentNo = studentNoCol !== -1 ? String(row[studentNoCol] || '').trim() : '';
        const title = titleCol !== -1 ? String(row[titleCol] || '').trim() : '';
        
        let fullName = '';
        if (fullNameCol !== -1 && row[fullNameCol]) {
          fullName = String(row[fullNameCol]).trim();
        } else if (firstNameCol !== -1 && lastNameCol !== -1) {
          const firstName = String(row[firstNameCol] || '').trim();
          const lastName = String(row[lastNameCol] || '').trim();
          fullName = `${firstName} ${lastName}`.trim();
        }
        
        // เพิ่มคำนำหน้าชื่อถ้ามี
        if (title && fullName) {
          fullName = `${title}${fullName}`;
        }
        
        students.push({
          studentId: studentId,
          studentNo: studentNo,
          fullName: fullName,
          grades: {}
        });
      }
    }
    
    if (students.length === 0) {
      return {
        success: false,
        message: `ไม่พบนักเรียนในชั้น ${grade} ห้อง ${classNo}`
      };
    }
    
    students.sort((a, b) => {
      if (a.studentNo && b.studentNo) {
        const numA = parseInt(a.studentNo) || 0;
        const numB = parseInt(b.studentNo) || 0;
        return numA - numB;
      }
      const numA = parseInt(a.studentId) || 0;
      const numB = parseInt(b.studentId) || 0;
      return numA - numB;
    });
    
    const scoresSheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!scoresSheet) {
      throw new Error("ไม่พบชีต SCORES_WAREHOUSE สำหรับปีปัจจุบัน");
    }
    
    const scoresData = scoresSheet.getDataRange().getValues();
    const scoresHeaders = scoresData[0];
    
    const scoreGradeCol = U_findColumnIndex(scoresHeaders, ['grade', 'ชั้น']);
    const scoreClassCol = U_findColumnIndex(scoresHeaders, ['class_no', 'ห้อง']);
    const scoreStudentIdCol = U_findColumnIndex(scoresHeaders, ['student_id', 'รหัสนักเรียน']);
    const subjectCol = U_findColumnIndex(scoresHeaders, ['subject_name', 'วิชา', 'ชื่อวิชา']);
    const scoreCol = U_findColumnIndex(scoresHeaders, ['final_grade', 'average', 'คะแนน', 'เกรด']);
    
    if (scoreStudentIdCol === -1) {
      throw new Error("ไม่พบคอลัมน์ 'student_id' ในชีต SCORES_WAREHOUSE");
    }
    
    const subjectsSet = new Set();
    
    for (let i = 1; i < scoresData.length; i++) {
      const row = scoresData[i];
      const rowGrade = String(row[scoreGradeCol] || '').trim();
      const rowClass = String(row[scoreClassCol] || '').trim();
      
      if (rowGrade === grade && rowClass === classNo) {
        const studentId = String(row[scoreStudentIdCol] || '').trim();
        const subject = String(row[subjectCol] || '').trim();
        const value = row[scoreCol];
        const gradeValue = convertToGrade(value);
        
        if (subject) {
          subjectsSet.add(subject);
          const student = students.find(s => s.studentId === studentId);
          if (student) {
            student.grades[subject] = gradeValue;
          }
        }
      }
    }
    
    var subjects = Array.from(subjectsSet);
    _sortBySubjectName(subjects, function(item) { return item; });
    
    students.forEach(student => {
      const grades = Object.values(student.grades);
      if (grades.length > 0) {
        const sum = grades.reduce((acc, val) => acc + val, 0);
        student.average = (sum / grades.length).toFixed(2);
      } else {
        student.average = '0.00';
      }
    });
    
   const settings = S_getGlobalSettings();
    
    return {
      success: true,
      grade: grade,
      classNo: classNo,
      gradeFullName: U_getGradeFullName(grade),
      students: students,
      subjects: subjects,
      settings: settings
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

function exportStudentScoresPDF(grade, classNo) {
  try {
    const data = getStudentScoresForWeb(grade, classNo);
    if (!data.success) {
      throw new Error(data.message);
    }
    
    // เพิ่มโลโก้ Base64 เข้าไปใน data
    data.logoBase64 = getLogoAsBase64();
    
    const html = generateScoresSummaryHTML(data);
    const fileName = `สรุปผลการเรียน_${grade}_${classNo}_${Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss")}`;
    const blob = Utilities.newBlob(html, 'text/html', `${fileName}.html`).getAs('application/pdf');
    blob.setName(`${fileName}.pdf`);
    
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (error) {
    throw new Error(error.message);
  }
}

function generateScoresSummaryHTML(data) {
  const { settings = {}, grade = '', classNo = '', gradeFullName = '', students = [], subjects = [], logoBase64 = null } = data;
  
  // ใช้โลโก้ Base64 ถ้ามี
  const logoUrl = logoBase64 || settings['logoUrl_lh3'] || settings['logo'] || settings['schoolLogo'] || '';
  const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียน';
  const semester = settings['ภาคเรียน'] || '';
  const academicYear = settings['ปีการศึกษา'] || '';
  
  const tableRows = students.map((student, index) => {
    const studentNo = student.studentNo || (index + 1);
    const gradesCells = subjects.map(subject => {
      const gradeValue = student.grades[subject];
      return gradeValue !== undefined ? gradeValue : '-';
    }).join('</td><td class="grade-cell">');
    
    return `
      <tr>
        <td class="student-no">${studentNo}</td>
        <td class="student-name">${student.fullName || '-'}</td>
        <td class="grade-cell">${gradesCells}</td>
        <td class="average">${student.average || '0.00'}</td>
      </tr>
    `;
  }).join('');
  
  // สร้างหัววิชาแบบแนวตั้ง
  const subjectHeaders = subjects.map(s => `<th class="vertical-text">${s}</th>`).join('');
  
  // สร้างส่วนแสดงโลโก้ (ถ้ามี)
  const logoSection = logoUrl ? `<img src="${logoUrl}" alt="โลโก้โรงเรียน" class="school-logo">` : '';
  
  return `
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <title>สรุปผลการเรียน ${grade} ห้อง ${classNo}</title>
  <style>
    @page { 
      size: A4 portrait; 
      margin: 8mm 10mm 8mm 12mm; /* ลดขอบซ้ายจาก 18mm เป็น 12mm */
    }
    
    body { 
      font-family: 'Sarabun', 'TH Sarabun New', 'Garuda', sans-serif; 
      font-size: 12pt; /* เพิ่มจาก 11pt เป็น 12pt */
      margin: 0; 
      padding: 0;
    }
    
    .header { 
      text-align: center; 
      margin-bottom: 10px;
      position: relative;
    }
    
    .school-logo {
      max-width: 100px; /* เพิ่มจาก 90px เป็น 100px */
      max-height: 100px;
      width: auto;
      height: auto;
      margin-bottom: 10px; /* เพิ่ม margin */
      display: block;
      margin-left: auto;
      margin-right: auto;
      object-fit: contain;
    }
    
    .header h2 { 
      margin: 4px 0; 
      font-size: 17pt; /* เพิ่มจาก 16pt เป็น 17pt */
      font-weight: bold;
    }
    
    .header p { 
      margin: 3px 0; 
      font-size: 15pt; /* เพิ่มจาก 14pt เป็น 15pt */
    }
    
    table { 
      width: 100%; 
      border-collapse: collapse; 
      margin-top: 8px;
      font-size: 14pt; /* เพิ่มจาก 13pt เป็น 14pt - ทั้งตารางเท่ากัน */
    }
    
    th, td { 
      border: 1px solid #000; 
      padding: 6px 4px; /* เพิ่ม padding จาก 5px 3px เป็น 6px 4px */
      text-align: center;
      vertical-align: middle;
    }
    
    th { 
      background-color: #e8e8e8; 
      font-weight: bold;
      font-size: 11pt; /* เพิ่มจาก 10pt เป็น 11pt */
    }
    
    /* คอลัมน์เลขที่ */
    .student-no {
      width: 45px; /* ลดจาก 60px เป็น 45px */
      min-width: 45px;
      font-size: 14pt;
      text-align: center;
      padding: 6px 4px; /* ลด padding ซ้าย-ขวา */
    }
    
    /* คอลัมน์ชื่อ-นามสกุล */
    .student-name { 
      text-align: left; 
      padding-left: 10px;
      padding-right: 5px;
      min-width: 240px; /* เพิ่มจาก 200px เป็น 240px */
      max-width: 270px; /* เพิ่มจาก 220px เป็น 270px */
      font-size: 14pt;
      word-wrap: break-word;
      white-space: normal;
      line-height: 1.3;
    }
    
    /* คอลัมน์เกรด */
    .grade-cell {
      width: 48px;
      min-width: 48px;
      font-size: 14pt; /* เพิ่มจาก 12pt เป็น 14pt */
      font-weight: bold;
      padding: 6px 5px;
    }
    
    /* คอลัมน์เฉลี่ย */
    .average { 
      font-weight: bold; 
      background-color: #f0f0f0;
      width: 60px;
      min-width: 60px;
      font-size: 14pt; /* เพิ่มจาก 12pt เป็น 14pt */
      padding: 6px 5px;
    }
    
    /* ชื่อวิชาแนวตั้ง */
    .vertical-text {
      writing-mode: vertical-rl;
      text-orientation: mixed;
      height: 230px; /* เพิ่มจาก 200px เป็น 230px สำหรับ "สังคมศึกษา ศาสนา และวัฒนธรรม" */
      padding: 10px 6px;
      font-size: 14pt;
      font-weight: 500;
      white-space: nowrap;
      vertical-align: bottom;
      line-height: 1.5;
    }
    
    /* หัวตารางรายวิชา */
    .subject-header {
      font-size: 14pt; /* เพิ่มจาก 12pt เป็น 14pt */
      padding: 5px;
      font-weight: bold;
    }
    
    /* แถวสลับสี */
    tbody tr:nth-child(even) { 
      background-color: #f9f9f9; 
    }
    
    /* ความสูงของแถวปรับตามเนื้อหา */
    tbody tr {
      height: auto;
      min-height: 30px;
    }
    
    /* หัวตารางแถวแรก */
    thead th {
      font-size: 14pt; /* เพิ่มจาก 13pt เป็น 14pt - เท่ากับตาราง */
      font-weight: bold;
    }
  </style>
</head>
<body>
  <div class="header">
    ${logoSection}
    <h2>สรุปผลการเรียน ชั้นประจำการเรียน ภาคเรียนที่ ${semester} ปีการศึกษา ${academicYear}</h2>
    <p>${schoolName}</p>
    <p>${gradeFullName} ห้อง ${classNo}</p>
  </div>
  
  <table>
    <thead>
      <tr>
        <th rowspan="2">ที่</th>
        <th rowspan="2">ชื่อ-นามสกุล</th>
        <th colspan="${Math.max(1, subjects.length)}" class="subject-header">รายวิชา</th>
        <th rowspan="2">เฉลี่ย</th>
      </tr>
      <tr>${subjectHeaders}</tr>
    </thead>
    <tbody>
      ${tableRows}
    </tbody>
  </table>
</body>
</html>
  `;
}

/**
 * ✅ ดึงรายการห้องเรียน - Fixed Version
 * @param {string} grade - ระดับชั้นที่เลือก เช่น "ป.1"
 * @returns {Array<string>} เลขห้องเรียน เช่น ["1", "2", "3"]
 */
function getAvailableClasses(grade) {  // ⚠️ เพิ่ม parameter
  try {
    Logger.log(`🔍 getAvailableClasses: grade="${grade}"`);

    if (!grade || !String(grade).trim()) {
      return [];
    }

    const ss = (getSpreadsheetId_())
      ? SS()
      : SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = ss.getSheetByName("Students") || ss.getSheetByName("นักเรียน");
    
    if (!studentsSheet) {
      Logger.log("❌ ไม่พบชีต Students/นักเรียน");
      return [];  // ⚠️ เปลี่ยนเป็น [] แทน Object
    }
    
    const data = studentsSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    const headers = data[0];

    const findCol = (names) => {
      if (typeof U_findColumnIndex === 'function') {
        return U_findColumnIndex(headers, names);
      }
      const normalized = names.map(n => String(n || '').toLowerCase().trim());
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i] || '').toLowerCase().trim();
        if (normalized.indexOf(h) !== -1) return i;
      }
      return -1;
    };

    const gradeCol = findCol(['grade', 'ชั้น', 'ระดับชั้น']);
    const classCol = findCol(['class_no', 'ห้อง', 'เลขห้อง']);
    
    Logger.log(`📋 gradeCol=${gradeCol}, classCol=${classCol}`);
    
    if (gradeCol === -1 || classCol === -1) {
      Logger.log("❌ ไม่พบคอลัมน์ที่ต้องการ");
      return [];
    }
    
    const classes = new Set();
    const targetGrade = grade.trim();  // ⚠️ เตรียม grade สำหรับเปรียบเทียบ
    
    for (let i = 1; i < data.length; i++) {
      const rowGrade = String(data[i][gradeCol] || '').trim();
      const rowClass = String(data[i][classCol] || '').trim();
      
      // ⚠️ เปลี่ยนจาก: if (grade && classNo)
      // ⚠️ เป็น: if (rowGrade === targetGrade && rowClass)
      if (rowGrade === targetGrade && rowClass) {
        classes.add(rowClass);  // ⚠️ เก็บเฉพาะเลขห้อง ไม่ใช่ `${grade}/${classNo}`
        
        // แสดง log ครั้งแรก
        if (classes.size === 1) {
          Logger.log(`✅ ตัวอย่าง: grade="${rowGrade}", class="${rowClass}"`);
        }
      }
    }
    
    if (classes.size === 0) {
      Logger.log(`❌ ไม่พบห้องเรียนสำหรับ grade="${targetGrade}"`);
      return [];
    }
    
    // ⚠️ เรียงลำดับตัวเลข
    const result = Array.from(classes).sort((a, b) => {
      const numA = parseInt(a) || 0;
      const numB = parseInt(b) || 0;
      return numA - numB;
    });
    
    Logger.log(`✅ ผลลัพธ์: ${JSON.stringify(result)}`);
    
    // ⚠️ Return แบบ Array ธรรมดา ไม่มี Object wrapper
    return result;
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    return [];  // ⚠️ เปลี่ยนเป็น [] แทน Object
  }
}
