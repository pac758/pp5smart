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

/**
 * ✅ ดึงคะแนนนักเรียนสำหรับ Parent Portal (ใช้ใน parent_portal.html)
 * @param {string} studentId - รหัสนักเรียน
 * @returns {{ success: boolean, scores: Array, message?: string }}
 */
function getParentStudentScores(studentId) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };

    var ss = SS();

    // ดึงข้อมูลจาก SCORES_WAREHOUSE
    var warehouseSheet = (typeof S_getYearlySheet === 'function')
      ? S_getYearlySheet('SCORES_WAREHOUSE')
      : ss.getSheetByName('SCORES_WAREHOUSE');

    if (!warehouseSheet) return { success: true, scores: [], message: 'ไม่พบชีต SCORES_WAREHOUSE' };

    var data = warehouseSheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, scores: [] };

    var headers = data[0];
    var col = {};
    ['student_id', 'subject_code', 'subject_name', 'subject_type', 'hours',
     'term1_total', 'term2_total', 'average', 'final_grade'].forEach(function(name) {
      col[name] = headers.indexOf(name);
    });

    var sid = String(studentId).trim();
    var scores = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[col['student_id']] || '').trim() !== sid) continue;

      var subjectType = String(row[col['subject_type']] || '').trim();
      var isActivity = /กิจกรรม/i.test(subjectType);
      var finalGrade = String(row[col['final_grade']] || '').trim();
      var term1 = row[col['term1_total']];
      var term2 = row[col['term2_total']];
      var avg = row[col['average']];

      var scoreObj = {
        subjectCode: String(row[col['subject_code']] || '').trim(),
        name: String(row[col['subject_name']] || '').trim(),
        isActivity: isActivity,
        term1Score: (term1 !== '' && term1 !== null && term1 !== undefined) ? term1 : '-',
        term2Score: (term2 !== '' && term2 !== null && term2 !== undefined) ? term2 : '-',
        average: (avg !== '' && avg !== null && avg !== undefined) ? avg : '-',
        grade: finalGrade || '-'
      };

      if (isActivity) {
        scoreObj.activityGrade = finalGrade || 'ยังไม่ประเมิน';
      }

      scores.push(scoreObj);
    }

    // ── ดึงผลประเมินจากชีต "การประเมินกิจกรรมพัฒนาผู้เรียน" มาใส่ในวิชากิจกรรม ──
    try {
      var actSheet = (typeof S_getYearlySheet === 'function')
        ? S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน')
        : ss.getSheetByName('การประเมินกิจกรรมพัฒนาผู้เรียน');
      if (actSheet) {
        var actData = actSheet.getDataRange().getValues();
        var actHeaders = actData[0];
        var actIdIdx = -1;
        for (var ah = 0; ah < actHeaders.length; ah++) {
          if (/รหัส/i.test(String(actHeaders[ah] || ''))) { actIdIdx = ah; break; }
        }

        if (actIdIdx >= 0) {
          // หาแถวของนักเรียน
          var studentActRow = null;
          for (var ai = 1; ai < actData.length; ai++) {
            if (String(actData[ai][actIdIdx] || '').trim() === sid) {
              studentActRow = actData[ai];
              break;
            }
          }

          if (studentActRow) {
            // สร้าง map ของผลประเมิน: { keyword → ผ่าน/ไม่ผ่าน }
            var actMap = {};
            var actSummary = '';
            for (var ac = 0; ac < actHeaders.length; ac++) {
              var hn = String(actHeaders[ac] || '').trim();
              var hv = String(studentActRow[ac] || '').trim();
              if (!hn || !hv) continue;
              if (hn === 'รวมกิจกรรม') { actSummary = hv; continue; }
              // เก็บ keyword → value (เช่น "แนะแนว" → "ผ่าน")
              actMap[hn] = hv;
            }

            // จับคู่ keyword กับวิชากิจกรรมใน scores
            var matchKeywords = [
              { keys: ['แนะแนว'], col: 'กิจกรรมแนะแนว' },
              { keys: ['ลูกเสือ', 'เนตรนารี'], col: 'ลูกเสือ_เนตรนารี' },
              { keys: ['ชุมนุม'], col: 'ชุมนุม' },
              { keys: ['สังคม', 'สาธารณ'], col: 'เพื่อสังคมและสาธารณประโยชน์' }
            ];

            for (var si = 0; si < scores.length; si++) {
              if (!scores[si].isActivity) continue;
              var sName = (scores[si].name || '').toLowerCase();
              for (var mk = 0; mk < matchKeywords.length; mk++) {
                var matched = matchKeywords[mk].keys.some(function(k) { return sName.indexOf(k) >= 0; });
                if (matched && actMap[matchKeywords[mk].col]) {
                  scores[si].grade = actMap[matchKeywords[mk].col];
                  scores[si].activityGrade = actMap[matchKeywords[mk].col];
                  break;
                }
              }
            }

            // เพิ่มสรุปรวมกิจกรรม
            if (actSummary) {
              scores.push({
                subjectCode: '',
                name: 'รวมกิจกรรมพัฒนาผู้เรียน',
                isActivity: true,
                term1Score: '-',
                term2Score: '-',
                average: '-',
                grade: actSummary,
                activityGrade: actSummary,
                isSummary: true
              });
            }
          }
        }
      }
    } catch (actErr) {
      Logger.log('⚠️ getParentStudentScores activity sheet: ' + actErr.message);
    }

    // เรียงลำดับ: วิชาปกติก่อน, กิจกรรมทีหลัง, เรียงตามรหัสวิชา
    scores.sort(function(a, b) {
      if (a.isActivity !== b.isActivity) return a.isActivity ? 1 : -1;
      if (a.isSummary !== b.isSummary) return a.isSummary ? 1 : -1;
      return (a.subjectCode || '').localeCompare(b.subjectCode || '');
    });

    return { success: true, scores: scores };
  } catch (e) {
    Logger.log('❌ getParentStudentScores Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ ดึงรายชื่อนักเรียนที่ผูกกับผู้ปกครอง (parent_portal)
 * อ่าน student_ids จาก Users sheet แล้วดึงข้อมูลจาก Students sheet
 */
function getParentStudents() {
  try {
    var session = getLoginSession();
    if (!session || !session.username) {
      return { success: false, message: 'กรุณาเข้าสู่ระบบก่อน' };
    }

    var ss = SS();

    // ดึง student_ids จาก Users sheet
    var userSheet = ss.getSheetByName('Users') || ss.getSheetByName('users');
    if (!userSheet) return { success: false, message: 'ไม่พบชีต Users' };

    var userData = userSheet.getDataRange().getValues();
    var uHeaders = userData[0];
    var uNameIdx = -1, sidIdx = -1;
    for (var h = 0; h < uHeaders.length; h++) {
      var hLow = String(uHeaders[h] || '').toLowerCase().trim();
      if (['username', 'user', 'userid'].indexOf(hLow) >= 0) uNameIdx = h;
      if (['student_ids', 'studentids', 'linked_students'].indexOf(hLow) >= 0) sidIdx = h;
    }

    var linkedIds = [];
    if (uNameIdx >= 0 && sidIdx >= 0) {
      for (var u = 1; u < userData.length; u++) {
        if (String(userData[u][uNameIdx] || '').trim().toLowerCase() === session.username.toLowerCase()) {
          var raw = String(userData[u][sidIdx] || '').trim();
          if (raw) {
            linkedIds = raw.split(/[,;\s]+/).map(function(s) { return s.trim(); }).filter(Boolean);
          }
          break;
        }
      }
    }

    if (linkedIds.length === 0) {
      return { success: true, data: [] };
    }

    // ดึงข้อมูลนักเรียนจาก Students sheet
    var studSheet = ss.getSheetByName('Students');
    if (!studSheet) return { success: false, message: 'ไม่พบชีต Students' };

    var studData = studSheet.getDataRange().getValues();
    var sHeaders = studData[0];
    var col = {};
    ['student_id', 'title', 'firstname', 'lastname', 'grade', 'class_no', 'gender'].forEach(function(name) {
      col[name] = sHeaders.indexOf(name);
    });

    var results = [];
    for (var i = 1; i < studData.length; i++) {
      var row = studData[i];
      var sid = String(row[col['student_id']] || '').trim();
      if (linkedIds.indexOf(sid) === -1) continue;
      results.push({
        studentId: sid,
        title: String(row[col['title']] || '').trim(),
        firstName: String(row[col['firstname']] || '').trim(),
        lastName: String(row[col['lastname']] || '').trim(),
        grade: String(row[col['grade']] || '').trim(),
        classNo: String(row[col['class_no']] || '').trim(),
        gender: String(row[col['gender']] || '').trim()
      });
    }

    return { success: true, data: results };
  } catch (e) {
    Logger.log('❌ getParentStudents Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ ดึงสรุปเวลาเรียนสำหรับ Parent Portal
 * @param {string} studentId
 */
function getParentStudentAttendance(studentId) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };

    // ดึงปีการศึกษาจาก settings
    var settings = (typeof S_getGlobalSettings === 'function') ? S_getGlobalSettings() : {};
    var academicYear = Number(settings['ปีการศึกษา'] || new Date().getFullYear() + 543);

    // หา grade, classNo ของนักเรียน
    var ss = SS();
    var studSheet = ss.getSheetByName('Students');
    if (!studSheet) return { success: false, message: 'ไม่พบชีต Students' };

    var studData = studSheet.getDataRange().getValues();
    var sHeaders = studData[0];
    var idIdx = sHeaders.indexOf('student_id');
    var gradeIdx = sHeaders.indexOf('grade');
    var classIdx = sHeaders.indexOf('class_no');
    var grade = '', classNo = '';

    for (var i = 1; i < studData.length; i++) {
      if (String(studData[i][idIdx] || '').trim() === String(studentId).trim()) {
        grade = String(studData[i][gradeIdx] || '').trim();
        classNo = String(studData[i][classIdx] || '').trim();
        break;
      }
    }

    if (!grade) return { success: true, attendance: { totalDays: 0, totalPresent: 0, totalLeave: 0, totalAbsent: 0, percentage: 0, monthly: [] } };

    // ใช้ getStudentAttendanceDetail ที่มีอยู่
    if (typeof getStudentAttendanceDetail === 'function') {
      var detail = getStudentAttendanceDetail(grade, classNo, studentId, academicYear);
      // แปลง format ให้ตรงกับ parent_portal.html
      var monthly = (detail.monthlyData || []).map(function(m) {
        return {
          month: m.monthName || '',
          present: m.present || 0,
          leave: m.leave || 0,
          absent: m.absent || 0,
          total: m.total || 0
        };
      });
      return {
        success: true,
        attendance: {
          totalDays: detail.summary.totalDays || 0,
          totalPresent: detail.summary.totalPresent || 0,
          totalLeave: detail.summary.totalLeave || 0,
          totalAbsent: detail.summary.totalAbsent || 0,
          percentage: parseFloat(detail.summary.percentage) || 0,
          monthly: monthly
        }
      };
    }

    return { success: true, attendance: { totalDays: 0, totalPresent: 0, totalLeave: 0, totalAbsent: 0, percentage: 0, monthly: [] } };
  } catch (e) {
    Logger.log('❌ getParentStudentAttendance Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ สร้างใบรายงานผลการเรียน PDF สำหรับ Parent Portal
 * @param {string} studentId
 */
function generateParentReportCard(studentId) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };

    // ใช้ generateOnePageReportPdf ที่มีอยู่
    if (typeof generateOnePageReportPdf === 'function') {
      var result = generateOnePageReportPdf(studentId);
      if (result && result.url) {
        return { success: true, viewUrl: result.url };
      } else if (result && typeof result === 'string' && result.startsWith('http')) {
        return { success: true, viewUrl: result };
      } else {
        return { success: false, message: 'ไม่สามารถสร้างใบรายงานได้' };
      }
    }

    return { success: false, message: 'ฟังก์ชันสร้างใบรายงานยังไม่พร้อมใช้งาน' };
  } catch (e) {
    Logger.log('❌ generateParentReportCard Error: ' + e.message);
    return { success: false, message: e.message };
  }
}
