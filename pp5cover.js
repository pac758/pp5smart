// ===============================
// ฟังก์ชันหลักสำหรับสร้าง PDF ปก ปพ.5 (เวอร์ชันสมบูรณ์)
// ===============================

// --- Local helpers (Web App safe) ---
function _pp5SS_() {
  return SS();
}

function _pp5FindCol_(headers, names) {
  if (typeof U_findColumnIndex === 'function') {
    return U_findColumnIndex(headers, names);
  }
  var normalized = names.map(function(n) { return String(n || '').toLowerCase().trim(); });
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase().trim();
    if (normalized.indexOf(h) !== -1) return i;
  }
  return -1;
}

function _pp5GradeFullName_(grade) {
  if (typeof U_getGradeFullName === 'function') {
    try { return U_getGradeFullName(grade); } catch(e) {}
  }
  var map = {
    'อ.1': 'อนุบาลปีที่ 1', 'อ.2': 'อนุบาลปีที่ 2', 'อ.3': 'อนุบาลปีที่ 3',
    'ป.1': 'ประถมศึกษาปีที่ 1', 'ป.2': 'ประถมศึกษาปีที่ 2', 'ป.3': 'ประถมศึกษาปีที่ 3',
    'ป.4': 'ประถมศึกษาปีที่ 4', 'ป.5': 'ประถมศึกษาปีที่ 5', 'ป.6': 'ประถมศึกษาปีที่ 6',
    'ม.1': 'มัธยมศึกษาปีที่ 1', 'ม.2': 'มัธยมศึกษาปีที่ 2', 'ม.3': 'มัธยมศึกษาปีที่ 3',
    'ม.4': 'มัธยมศึกษาปีที่ 4', 'ม.5': 'มัธยมศึกษาปีที่ 5', 'ม.6': 'มัธยมศึกษาปีที่ 6'
  };
  return map[grade] || grade;
}
function exportPp5CoverPDF(grade, classNo) {
  try {
    Logger.log(`🚀 เริ่มสร้าง PDF ปก ปพ.5`);
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();

    if (!grade || !classNo) {
      throw new Error(`ไม่ได้รับค่า grade หรือ classNo`);
    }

    // --- ดึงข้อมูลนักเรียน ---
    const ss = _pp5SS_();
    const studentsSheet = ss.getSheetByName("Students") || ss.getSheetByName("นักเรียน");
    if (!studentsSheet) throw new Error("ไม่พบชีต 'Students/นักเรียน'");
    
    const studentData = studentsSheet.getDataRange().getValues();
    const headers = studentData[0];
    const gradeColIndex = _pp5FindCol_(headers, ['grade', 'ชั้น', 'ระดับชั้น', 'Grade']);
    const classNoColIndex = _pp5FindCol_(headers, ['class_no', 'ห้อง', 'หมายเลขห้อง', 'Class', 'class']);

    const studentCount = studentData.filter((row, index) => 
      index > 0 && 
      String(row[gradeColIndex] || "").trim() === grade && 
      String(row[classNoColIndex] || "").trim() === classNo
    ).length;

    if (studentCount === 0) {
      throw new Error(`ไม่พบข้อมูลนักเรียนในชั้น ${grade} ห้อง ${classNo}`);
    }

    // --- รวบรวมข้อมูล ---
    const settings = S_getGlobalSettings(false);
    const gradeFullName = _pp5GradeFullName_(grade);
    const subjectSummary = getSubjectScoreSummary(grade, classNo);
    const attributes = getAttributeSummary(grade, classNo);
    const activity = getActivitySummary(grade, classNo);
    const reading = getReadingSummary(grade, classNo);
    settings.ชื่อครูประจำชั้น = settings.ชื่อครูประจำชั้น || getClassTeacherName(grade, classNo) || '';

    // --- สร้างโลโก้แบบ Base64 ---
    const logoFileId = settings.logoFileId || settings['logoFileId'];
    const logoDataUri = _getLogoDataUrl(logoFileId);
    Logger.log(`สร้าง Logo Base64 URI: ${logoDataUri ? 'สำเร็จ' : 'ไม่สำเร็จ'}`);
    
    // --- สร้าง HTML ---
    const templateData = {
      settings, grade, classNo, gradeFullName, studentCount, 
      subjectSummary, attributes, reading, activity, logoDataUri
    };
    const html = generatePp5CoverHtml(templateData);

    // --- สร้างไฟล์ PDF ---
    const fileName = `ปก ปพ.5 รวมวิชา_${grade}_${classNo}_${Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss")}`;
    const blob = Utilities.newBlob(html, 'text/html', `${fileName}.html`).getAs('application/pdf');
    blob.setName(`${fileName}.pdf`);
    
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    Logger.log(`✅ สร้าง PDF สำเร็จ: ${file.getUrl()}`);
    return file.getUrl();
    
  } catch (error) {
    Logger.log(`❌ เกิดข้อผิดพลาดในฟังก์ชัน exportPp5CoverPDF: ${error.message}`);
    throw new Error(error.message);
  }
}

// ===============================
// ฟังก์ชันสนับสนุนที่จำเป็น
// ===============================

/**
 * [ฟังก์ชันหัวใจสำคัญ] แปลง File ID ของรูปภาพเป็น Data URL (Base64)
 */

/**
 * [ฟังก์ชันใหม่] สร้างโค้ด HTML ของปก ปพ.5 โดยตรง
 */
function generatePp5CoverHtml(data) {
  const { 
    settings = {}, 
    gradeFullName = '', 
    classNo = '', 
    studentCount = 0, 
    subjectSummary = [], 
    attributes = {}, 
    reading = {}, 
    activity = {}, 
    logoDataUri = '' 
  } = data;

  var activityKeywords = ['กิจกรรม', 'แนะแนว', 'ลูกเสือ', 'เนตรนารี', 'ชุมนุม', 'เพื่อสังคม', 'ชมรม'];
  var filteredSubjects = (subjectSummary.length > 0 ? subjectSummary : []).filter(function(s) {
    var n = (s.name || '').trim();
    return !activityKeywords.some(function(kw) { return n.indexOf(kw) !== -1; });
  });
  const subjectRows = filteredSubjects.map((s, i) => `
    <tr>
      <td>${i + 1}</td>
      <td class="subject-col">${s.name || ''}</td>
      <td>${s.count["4"] || 0}</td><td>${s.count["3.5"] || 0}</td><td>${s.count["3"] || 0}</td><td>${s.count["2.5"] || 0}</td>
      <td>${s.count["2"] || 0}</td><td>${s.count["1.5"] || 0}</td><td>${s.count["1"] || 0}</td><td>${s.count["0"] || 0}</td>
    </tr>
  `).join('');

  return `
  <!DOCTYPE html>
  <html lang="th">
  <head>
    <meta charset="UTF-8"><title>ปก ปพ.5 รวมวิชา</title>
    <style>
      body { font-family: 'Sarabun', sans-serif; font-size: 14pt; margin: 40px; line-height: 1.2; }
      table { width: 100%; border-collapse: collapse; margin-top: 12px; }
      th, td { border: 1px solid #000; padding: 4px; text-align: center; }
      th { font-size: 14pt; } td { font-size: 11pt; }
      .subject-col { text-align: left; padding-left: 8px; }
      .header { text-align: center; font-weight: bold; margin-bottom: 10px; }
      .section-title { margin-top: 10px; font-weight: bold; }
      .signature { margin-top: 50px; display: flex; justify-content: space-around; font-size: 12px; }
      .signature div { text-align: center; }
      .logo { text-align: center; margin-bottom: 10px; }
      .logo img { height: 80px; max-width: 200px; object-fit: contain; }
    </style>
  </head>
  <body>
    <div class="logo">
      ${logoDataUri ? `<img src="${logoDataUri}" alt="โลโก้โรงเรียน">` : ''}
    </div>
    <div class="header">
      <h2 style="margin: 0;">แบบรายงานผลการพัฒนาคุณภาพผู้เรียน (ปพ.5)</h2>
      <div style="font-size: 16pt;">${settings['ชื่อโรงเรียน'] || 'โรงเรียน'}</div>
      <div style="font-size: 14pt;">${settings['ที่อยู่โรงเรียน'] || ''}</div>
      <div style="font-size: 14pt;">ปีการศึกษา ${settings['ปีการศึกษา'] || ''}</div>
    </div>
    <div class="section-title">
      ชั้น ${gradeFullName} ห้อง ${classNo} จำนวนนักเรียนทั้งหมด ${studentCount} คน
    </div>
    <table>
      <thead>
        <tr><th rowspan="2">ที่</th><th rowspan="2">ชื่อวิชา</th><th colspan="8">ระดับผลการเรียนรายวิชา (คน)</th></tr>
        <tr><th>4</th><th>3.5</th><th>3</th><th>2.5</th><th>2</th><th>1.5</th><th>1</th><th>0</th></tr>
      </thead>
      <tbody>${subjectRows}</tbody>
    </table>
    <div class="section-title">การประเมินผลแบบบูรณาการ</div>
    <table>
      <thead>
        <tr><th colspan="3">การอ่าน คิดวิเคราะห์และเขียน</th><th colspan="3">คุณลักษณะอันพึงประสงค์</th><th colspan="3">กิจกรรมพัฒนาผู้เรียน</th></tr>
        <tr><th rowspan="2">ระดับคุณภาพ</th><th colspan="2">จำนวนที่ได้</th><th rowspan="2">ระดับคุณภาพ</th><th colspan="2">จำนวนที่ได้</th><th rowspan="2">การผ่าน</th><th colspan="2">จำนวนที่ได้</th></tr>
        <tr><th>คน</th><th>ร้อยละ</th><th>คน</th><th>ร้อยละ</th><th>คน</th><th>ร้อยละ</th></tr>
      </thead>
      <tbody>
        <tr>
          <td class="subject-col">ดีเยี่ยม</td><td>${reading.level3 || 0}</td><td>${(studentCount > 0 ? ((reading.level3 || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ดีเยี่ยม</td><td>${attributes['ดีเยี่ยม'] || 0}</td><td>${(studentCount > 0 ? ((attributes['ดีเยี่ยม'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ผ่าน</td><td>${activity['ผ่าน'] || 0}</td><td>${(studentCount > 0 ? ((activity['ผ่าน'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
        </tr>
        <tr>
          <td class="subject-col">ดี</td><td>${reading.level2 || 0}</td><td>${(studentCount > 0 ? ((reading.level2 || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ดี</td><td>${attributes['ดี'] || 0}</td><td>${(studentCount > 0 ? ((attributes['ดี'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ไม่ผ่าน</td><td>${activity['ไม่ผ่าน'] || 0}</td><td>${(studentCount > 0 ? ((activity['ไม่ผ่าน'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
        </tr>
        <tr>
          <td class="subject-col">ผ่านเกณฑ์</td><td>${reading.level1 || 0}</td><td>${(studentCount > 0 ? ((reading.level1 || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ผ่านเกณฑ์</td><td>${attributes['ผ่าน'] || 0}</td><td>${(studentCount > 0 ? ((attributes['ผ่าน'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td></td><td></td><td></td>
        </tr>
        <tr>
          <td class="subject-col">ไม่ผ่านเกณฑ์</td><td>${reading.level0 || 0}</td><td>${(studentCount > 0 ? ((reading.level0 || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td class="subject-col">ไม่ผ่านเกณฑ์</td><td>${attributes['ไม่ผ่าน'] || 0}</td><td>${(studentCount > 0 ? ((attributes['ไม่ผ่าน'] || 0) / studentCount * 100).toFixed(1) : '0.0')}</td>
          <td></td><td></td><td></td>
        </tr>
      </tbody>
    </table>
    <div class="signature">
      <div>
        ลงชื่อ..............................................<br>
        (${settings['ชื่อครูประจำชั้น'] || '................................'})<br>
        ครูประจำชั้น
      </div>
      <div>
        <span style="margin-right:20px;">☐ เห็นควรอนุมัติ</span><span>☐ เห็นควรปรับปรุง</span><br><br>
        ลงชื่อ..............................................<br>
        (${settings['ชื่อหัวหน้างานวิชาการ'] || 'หัวหน้างานวิชาการ'})<br>
        หัวหน้างานวิชาการ
      </div>
      <div>
        <span style="margin-right:20px;">☐ อนุมัติ</span><span>☐ ไม่อนุมัติ</span><br><br>
        ลงชื่อ..............................................<br>
        (${settings['ชื่อผู้อำนวยการ'] || 'ผู้อำนวยการ'})<br>
        ${settings['ตำแหน่งผู้อำนวยการ'] || 'ผู้อำนวยการโรงเรียน'}
      </div>
    </div>
  </body></html>`;
}

// getGlobalSettings() → ลบแล้ว ใช้ S_getGlobalSettings() จาก settings_unified.gs

/**
 * ดึงชื่อครูประจำชั้น (เวอร์ชันแก้ไขให้ตรงกับชีต HomeroomTeachers)
 */
function getClassTeacherName(grade, classNo) {
  try {
    const ss = _pp5SS_();
    const teacherSheet = ss.getSheetByName("HomeroomTeachers");
    if (!teacherSheet) return '';
    
    const data = teacherSheet.getDataRange().getValues();
    const headers = data[0];
    
    // --- [แก้ไข] เพิ่มชื่อคอลัมน์ที่ถูกต้องเข้าไปในลิสต์ ---
    const gradeCol = _pp5FindCol_(headers, ['Grade', 'grade', 'ชั้น']);
    const classCol = _pp5FindCol_(headers, ['ClassNo', 'class_no', 'ห้อง']);
    const teacherCol = _pp5FindCol_(headers, ['TeacherName', 'teacher_name', 'ชื่อครู']);
    
    if (gradeCol === -1 || classCol === -1 || teacherCol === -1) {
      Logger.log("ไม่พบ Column ที่ต้องการในชีต HomeroomTeachers");
      return '';
    }
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][gradeCol]).trim() === grade && String(data[i][classCol]).trim() === classNo) {
        return String(data[i][teacherCol]).trim();
      }
    }
    return '';
  } catch (e) {
    Logger.log('Error in getClassTeacherName:', e.toString());
    return '';
  }
}

// ======== ฟังก์ชันสรุปข้อมูลจากชีตต่างๆ (แก้ไขแล้ว: เพิ่ม grade, classNo) ========
function getSummaryDataFromSheet(sheetName, resultColumnName, validResults, grade, classNo) {
  try {
    const ss = _pp5SS_();
    // ถ้าเป็นชีตรายปี ให้ใช้ S_getYearlySheet
    var sheet;
    if (typeof S_YEARLY_SHEETS !== 'undefined' && S_YEARLY_SHEETS.indexOf(sheetName) !== -1) {
      sheet = S_getYearlySheet(sheetName);
    } else {
      sheet = ss.getSheetByName(sheetName);
    }
    if (!sheet) {
      Logger.log('⚠️ ไม่พบชีต: ' + sheetName);
      const allSheets = ss.getSheets().map(s => s.getName());
      Logger.log('📋 ชีตทั้งหมดในไฟล์: ' + JSON.stringify(allSheets));
      return validResults.reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('⚠️ ชีต ' + sheetName + ' ไม่มีข้อมูล (แถว: ' + data.length + ')');
      return validResults.reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
    }

    const headers = data[0];
    const gradeCol = _pp5FindCol_(headers, ['grade', 'ชั้น']);
    const classCol = _pp5FindCol_(headers, ['class_no', 'ห้อง']);
    const resultCol = _pp5FindCol_(headers, [resultColumnName]);

    Logger.log(`📊 ${sheetName}: gradeCol=${gradeCol}, classCol=${classCol}, resultCol=${resultCol}, grade="${grade}", classNo="${classNo}"`);
    
    if (gradeCol === -1 || classCol === -1 || resultCol === -1) {
      Logger.log(`⚠️ ไม่พบคอลัมน์ที่ต้องการในชีต ${sheetName} (หา: ${resultColumnName})`);
      Logger.log(`📋 Headers: ${JSON.stringify(headers)}`);
      return validResults.reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
    }

    const summary = validResults.reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
    let matchCount = 0;

    data.forEach((row, index) => {
      if (index > 0 &&
          String(row[gradeCol]).trim() === String(grade).trim() &&
          String(row[classCol]).trim() === String(classNo).trim()) {
        matchCount++;
        const result = String(row[resultCol]).trim();
        if (summary.hasOwnProperty(result)) {
          summary[result]++;
        } else {
          Logger.log(`⚠️ ค่าที่ไม่รู้จักในชีต ${sheetName}: "${result}"`);
        }
      }
    });

    Logger.log(`✅ ${sheetName}: พบ ${matchCount} แถวที่ตรงกัน, สรุป: ${JSON.stringify(summary)}`);
    return summary;
  } catch (e) {
    Logger.log(`❌ Error reading summary from ${sheetName}: ${e.message}`);
    return validResults.reduce((acc, key) => ({ ...acc, [key]: 0 }), {});
  }
}

function getAttributeSummary(grade, classNo) {
  return getSummaryDataFromSheet("การประเมินคุณลักษณะ", 'ผลการประเมิน', ["ดีเยี่ยม", "ดี", "ผ่าน", "ไม่ผ่าน"], grade, classNo);
}

function getActivitySummary(grade, classNo) {
  return getSummaryDataFromSheet("การประเมินกิจกรรมพัฒนาผู้เรียน", 'รวมกิจกรรม', ["ผ่าน", "ไม่ผ่าน"], grade, classNo);
}

function getReadingSummary(grade, classNo) {
  var allKeys = ["ดีเยี่ยม", "ดี", "ผ่าน", "ไม่ผ่าน", "ผ่านเกณฑ์", "ไม่ผ่านเกณฑ์"];
  var summary = getSummaryDataFromSheet("การประเมินอ่านคิดเขียน", 'สรุปผลการประเมิน', allKeys, grade, classNo);
  return {
    level3: summary['ดีเยี่ยม'] || 0,
    level2: summary['ดี'] || 0,
    level1: (summary['ผ่าน'] || 0) + (summary['ผ่านเกณฑ์'] || 0),
    level0: (summary['ไม่ผ่าน'] || 0) + (summary['ไม่ผ่านเกณฑ์'] || 0),
  };
}

/**
 * สรุปคะแนนรายวิชา (เวอร์ชันแก้ไข: เรียงลำดับตามที่กำหนด)
 */
function getSubjectScoreSummary(grade, classNo) {
  try {
    const ss = _pp5SS_();
    const scoreSheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!scoreSheet) return [];
    
    const data = scoreSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const gradeCol = _pp5FindCol_(headers, ['grade', 'ชั้น']);
    const classCol = _pp5FindCol_(headers, ['class_no', 'ห้อง']);
    const subjectCol = _pp5FindCol_(headers, ['subject_name', 'วิชา']);
    const scoreCol = _pp5FindCol_(headers, ['final_grade', 'average', 'คะแนน']);
    
    if ([gradeCol, classCol, subjectCol, scoreCol].includes(-1)) return [];

    const summary = {};
    data.forEach((row, index) => {
      if (index > 0 && String(row[gradeCol]).trim() === grade && String(row[classCol]).trim() === classNo) {
        const subject = String(row[subjectCol]).trim();
        const score = parseFloat(row[scoreCol]) || 0;
        
        if (!subject) return;
        if (!summary[subject]) {
          summary[subject] = { name: subject, count: { "4": 0, "3.5": 0, "3": 0, "2.5": 0, "2": 0, "1.5": 0, "1": 0, "0": 0 } };
        }
        
        if (score >= 80) summary[subject].count["4"]++;
        else if (score >= 75) summary[subject].count["3.5"]++;
        else if (score >= 70) summary[subject].count["3"]++;
        else if (score >= 65) summary[subject].count["2.5"]++;
        else if (score >= 60) summary[subject].count["2"]++;
        else if (score >= 55) summary[subject].count["1.5"]++;
        else if (score >= 50) summary[subject].count["1"]++;
        else summary[subject].count["0"]++;
      }
    });

    const subjectArray = Object.values(summary);
    _sortBySubjectName(subjectArray, 'name');

    return subjectArray;

  } catch (e) {
    Logger.log("Error in getSubjectScoreSummary: " + e.toString());
    return [];
  }
}

// ======== ฟังก์ชันช่วยเหลือ ========

