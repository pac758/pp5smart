/**************************************************
 * 📚 Subject Management Functions
 * ฟังก์ชันสำหรับจัดการรายวิชาใน Google Sheets
 **************************************************/

const SUBJECT_SHEET_NAME = 'รายวิชา';

const SUBJECT_HEADERS_BASE = [
  'ชั้น', 'รหัสวิชา', 'ชื่อวิชา', 'ชั่วโมง/ปี',
  'ประเภทวิชา', 'ครูผู้สอน', 'คะแนนระหว่างปี',
  'คะแนนปลายปี', 'รวม'
];

const FULLSCORE_HEADERS = [
  'เต็ม_ครั้ง1', 'เต็ม_ครั้ง2', 'เต็ม_ครั้ง3', 'เต็ม_ครั้ง4',
  'เต็ม_ครั้ง5', 'เต็ม_ครั้ง6', 'เต็ม_ครั้ง7', 'เต็ม_ครั้ง8', 'เต็ม_ครั้ง9'
];

const SUBJECT_HEADERS_ALL = SUBJECT_HEADERS_BASE.concat(FULLSCORE_HEADERS);

/**
 * 🔍 ดึงข้อมูลรายวิชาทั้งหมด
 */
function getSubjects() {
  try {
    const ss = _openSpreadsheet_();
    let sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    
    // สร้างชีตใหม่หากไม่มี
    if (!sheet) {
      sheet = ss.insertSheet(SUBJECT_SHEET_NAME);
      // สร้างหัวตาราง
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setValues([SUBJECT_HEADERS_ALL]);
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      Logger.log('✅ สร้างชีต "รายวิชา" ใหม่พร้อมหัวตาราง');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }

    const headers = data[0];
    const subjects = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const subject = {};
      
      headers.forEach((header, index) => {
        subject[header] = row[index] || '';
      });
      
      subjects.push(subject);
    }

    Logger.log(`✅ ดึงข้อมูลรายวิชา: ${subjects.length} รายการ`);
    return subjects;

  } catch (error) {
    Logger.log('❌ getSubjects:', error);
    throw new Error(`ไม่สามารถดึงข้อมูลรายวิชาได้: ${error.message}`);
  }
}

/**
 * ➕ เพิ่มรายวิชาใหม่
 */
function addSubject(subjectData) {
  try {
    const ss = _openSpreadsheet_();
    let sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SUBJECT_SHEET_NAME);
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setValues([SUBJECT_HEADERS_ALL]);
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      ensureFullScoreHeaders_(sheet);
    }

    // ตรวจสอบข้อมูลที่จำเป็น
    if (!subjectData['ชั้น'] || !subjectData['ชื่อวิชา']) {
      throw new Error('กรุณากรอกชั้นและชื่อวิชา');
    }

    // ตรวจสอบความซ้ำ
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === subjectData['ชั้น'] && 
          existingData[i][2] === subjectData['ชื่อวิชา']) {
        throw new Error('รายวิชานี้มีอยู่แล้วในชั้นดังกล่าว');
      }
    }

    // เพิ่มข้อมูล
    const headers = existingData[0];
    const newRow = headers.map(header => subjectData[header] || '');
    
    sheet.appendRow(newRow);
    
    Logger.log(`✅ เพิ่มรายวิชา: ${subjectData['ชื่อวิชา']} ชั้น ${subjectData['ชั้น']}`);
    return 'เพิ่มรายวิชาสำเร็จ';

  } catch (error) {
    Logger.log('❌ addSubject:', error);
    throw new Error(`ไม่สามารถเพิ่มรายวิชาได้: ${error.message}`);
  }
}

/**
 * ✏️ แก้ไขรายวิชา
 */
function updateSubject(index, subjectData) {
  try {
    const ss = _openSpreadsheet_();
    const sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('ไม่พบชีตรายวิชา');
    }

    // ตรวจสอบข้อมูลที่จำเป็น
    if (!subjectData['ชั้น'] || !subjectData['ชื่อวิชา']) {
      throw new Error('กรุณากรอกชั้นและชื่อวิชา');
    }

    const data = sheet.getDataRange().getValues();
    const rowIndex = parseInt(index) + 2; // +2 เพราะ index เริ่มจาก 0 และแถวแรกเป็น header

    if (rowIndex > data.length) {
      throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');
    }

    // ตรวจสอบความซ้ำ (ยกเว้นแถวที่กำลังแก้ไข)
    for (let i = 1; i < data.length; i++) {
      if (i !== rowIndex - 1 && 
          data[i][0] === subjectData['ชั้น'] && 
          data[i][2] === subjectData['ชื่อวิชา']) {
        throw new Error('รายวิชานี้มีอยู่แล้วในชั้นดังกล่าว');
      }
    }

    // อัปเดตข้อมูล
    const headers = data[0];
    const updatedRow = headers.map(header => subjectData[header] || '');
    
    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
    
    Logger.log(`✅ แก้ไขรายวิชา: ${subjectData['ชื่อวิชา']} ชั้น ${subjectData['ชั้น']}`);
    return 'แก้ไขรายวิชาสำเร็จ';

  } catch (error) {
    Logger.log('❌ updateSubject:', error);
    throw new Error(`ไม่สามารถแก้ไขรายวิชาได้: ${error.message}`);
  }
}

/**
 * 🗑️ ลบรายวิชา
 */
function deleteSubject(index) {
  try {
    const ss = _openSpreadsheet_();
    const sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('ไม่พบชีตรายวิชา');
    }

    const data = sheet.getDataRange().getValues();
    const rowIndex = parseInt(index) + 2; // +2 เพราะ index เริ่มจาก 0 และแถวแรกเป็น header

    if (rowIndex > data.length) {
      throw new Error('ไม่พบข้อมูลที่ต้องการลบ');
    }

    const subjectName = data[rowIndex - 1][2]; // ชื่อวิชา
    const grade = data[rowIndex - 1][0]; // ชั้น
    
    sheet.deleteRow(rowIndex);
    
    Logger.log(`✅ ลบรายวิชา: ${subjectName} ชั้น ${grade}`);
    return 'ลบรายวิชาสำเร็จ';

  } catch (error) {
    Logger.log('❌ deleteSubject:', error);
    throw new Error(`ไม่สามารถลบรายวิชาได้: ${error.message}`);
  }
}

/**
 * 📄 ส่งออกรายวิชาทั้งหมดเป็น PDF
 */
function exportAllSubjectsAsPdf() {
  try {
    const subjects = getSubjects();
    
    if (subjects.length === 0) {
      throw new Error('ไม่มีข้อมูลรายวิชาให้ส่งออก');
    }

    // สร้าง HTML สำหรับ PDF
    const html = _buildSubjectsListHTML_(subjects);
    
    // สร้าง PDF
    const htmlOutput = HtmlService.createHtmlOutput(html);
    const pdfBlob = htmlOutput.getBlob().getAs('application/pdf');
    
    const fileName = `รายงานรายวิชาทั้งหมด_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm")}.pdf`;
    pdfBlob.setName(fileName);
    
    // บันทึกไฟล์
    const url = _saveBlobGetUrl_(pdfBlob, 'Subject Reports');
    Logger.log(`✅ ส่งออกรายวิชาทั้งหมด: ${subjects.length} รายการ`);
    return url;

  } catch (error) {
    Logger.log('❌ exportAllSubjectsAsPdf:', error);
    throw new Error(`ไม่สามารถส่งออก PDF ได้: ${error.message}`);
  }
}

/**
 * 🏗️ สร้าง HTML สำหรับรายงานรายวิชาทั้งหมด
 */
function _buildSubjectsListHTML_(subjects) {
  const settings = getGlobalSettingsFixed();
  const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียน';
  const academicYear = settings['ปีการศึกษา'] || String(S_getCurrentAcademicYear_());
  const currentDate = Utilities.formatDate(new Date(), "Asia/Bangkok", "d MMMM yyyy");
  
  // จัดกลุ่มตามชั้น
  const gradeGroups = {};
  subjects.forEach(subject => {
    const grade = subject['ชั้น'] || 'ไม่ระบุ';
    if (!gradeGroups[grade]) {
      gradeGroups[grade] = [];
    }
    gradeGroups[grade].push(subject);
  });
  
  // สร้างตารางสำหรับแต่ละชั้น
  let tablesHTML = '';
  const gradeOrder = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'];
  
  gradeOrder.forEach(grade => {
    if (gradeGroups[grade] && gradeGroups[grade].length > 0) {
      tablesHTML += `
        <div class="grade-section">
          <h3 class="grade-title">ระดับชั้น ${grade}</h3>
          <table class="subject-table">
            <thead>
              <tr>
                <th width="8%">ที่</th>
                <th width="12%">รหัสวิชา</th>
                <th width="25%">ชื่อวิชา</th>
                <th width="10%">ชั่วโมง/ปี</th>
                <th width="12%">ประเภทวิชา</th>
                <th width="20%">ครูผู้สอน</th>
                <th width="13%">หมายเหตุ</th>
              </tr>
            </thead>
            <tbody>`;
      
      gradeGroups[grade].forEach((subject, index) => {
        tablesHTML += `
          <tr>
            <td class="text-center">${index + 1}</td>
            <td class="text-center">${subject['รหัสวิชา'] || '-'}</td>
            <td class="text-left">${subject['ชื่อวิชา'] || '-'}</td>
            <td class="text-center">${subject['ชั่วโมง/ปี'] || '-'}</td>
            <td class="text-center">${subject['ประเภทวิชา'] || '-'}</td>
            <td class="text-left">${subject['ครูผู้สอน'] || '-'}</td>
            <td class="text-center">-</td>
          </tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>`;
    }
  });
  
  return `
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4 portrait; margin: 15mm; }
    body {
      font-family: 'TH Sarabun New', 'Sarabun', Arial, sans-serif;
      font-size: 14pt;
      line-height: 1.4;
      color: #000;
      margin: 0;
      padding: 0;
    }
    
    .header {
      text-align: center;
      margin-bottom: 20px;
      border-bottom: 2px solid #333;
      padding-bottom: 15px;
    }
    
    .header h1 {
      font-size: 18pt;
      font-weight: bold;
      margin: 5px 0;
    }
    
    .header h2 {
      font-size: 16pt;
      font-weight: normal;
      margin: 3px 0;
    }
    
    .grade-section {
      margin-bottom: 25px;
      page-break-inside: avoid;
    }
    
    .grade-title {
      background: #f0f0f0;
      padding: 8px 15px;
      margin: 0 0 10px 0;
      font-size: 14pt;
      font-weight: bold;
      border-left: 4px solid #2196F3;
    }
    
    .subject-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 15px;
    }
    
    .subject-table th,
    .subject-table td {
      border: 1px solid #333;
      padding: 8px 6px;
      vertical-align: middle;
    }
    
    .subject-table th {
      background: #e8e8e8;
      font-weight: bold;
      text-align: center;
      font-size: 12pt;
    }
    
    .subject-table td {
      font-size: 12pt;
    }
    
    .text-center { text-align: center; }
    .text-left { text-align: left; padding-left: 8px; }
    .text-right { text-align: right; padding-right: 8px; }
    
    .footer {
      margin-top: 30px;
      text-align: right;
      font-size: 12pt;
    }
    
    .summary {
      margin: 20px 0;
      padding: 15px;
      background: #f9f9f9;
      border: 1px solid #ddd;
      border-radius: 5px;
    }
    
    .summary h3 {
      margin-top: 0;
      font-size: 14pt;
      color: #333;
    }
    
    @media print {
      .grade-section {
        page-break-inside: avoid;
      }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>${schoolName}</h1>
    <h2>รายงานรายวิชาทั้งหมด</h2>
    <h2>ปีการศึกษา ${academicYear}</h2>
  </div>
  
  <div class="summary">
    <h3>สรุปจำนวนรายวิชา</h3>
    <p>รายวิชาทั้งหมด: <strong>${subjects.length}</strong> รายการ</p>
    ${Object.keys(gradeGroups).map(grade => 
      `<p>ระดับชั้น ${grade}: <strong>${gradeGroups[grade].length}</strong> รายวิชา</p>`
    ).join('')}
  </div>
  
  ${tablesHTML}
  
  <div class="footer">
    <p>วันที่พิมพ์: ${currentDate}</p>
  </div>
</body>
</html>`;
}

/**
 * 🔧 ตรวจสอบและเพิ่มคอลัมน์คะแนนเต็ม (เต็ม_ครั้ง1-9) ในชีตรายวิชาเก่าที่ยังไม่มี
 */
function ensureFullScoreHeaders_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = [];
  
  FULLSCORE_HEADERS.forEach(function(h) {
    if (headers.indexOf(h) === -1) missing.push(h);
  });
  
  if (missing.length > 0) {
    var startCol = headers.length + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    sheet.getRange(1, startCol, 1, missing.length).setFontWeight('bold');
    Logger.log('✅ เพิ่มคอลัมน์คะแนนเต็มในชีตรายวิชา: ' + missing.join(', '));
  }
}

/**
 * 📊 ดึงคะแนนเต็มรายตัวชี้วัด (9 ช่อง) จากชีตรายวิชา
 * @param {string} subjectName - ชื่อวิชา
 * @param {string} grade - ชั้น เช่น "ป.3"
 * @returns {Object} { fullScores: [s1,s2,...,s9], midMax, finalMax }
 */
function getSubjectFullScores(subjectName, grade) {
  try {
    const ss = _openSpreadsheet_();
    const sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    if (!sheet) return { fullScores: [0,0,0,0,0,0,0,0,0], midMax: 70, finalMax: 30 };
    
    ensureFullScoreHeaders_(sheet);
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { fullScores: [0,0,0,0,0,0,0,0,0], midMax: 70, finalMax: 30 };
    
    const headers = data[0];
    var idxGrade = headers.indexOf('ชั้น');
    var idxName = headers.indexOf('ชื่อวิชา');
    var idxMid = headers.indexOf('คะแนนระหว่างปี');
    var idxFinal = headers.indexOf('คะแนนปลายปี');
    
    // หา index ของ FULLSCORE_HEADERS
    var fsIdx = FULLSCORE_HEADERS.map(function(h) { return headers.indexOf(h); });
    
    // หาแถวที่ตรงกับ grade + subjectName
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[idxGrade]).trim() === String(grade).trim() &&
          String(row[idxName]).trim() === String(subjectName).trim()) {
        var fullScores = fsIdx.map(function(idx) { return idx >= 0 ? (Number(row[idx]) || 0) : 0; });
        return {
          fullScores: fullScores,
          midMax: Number(row[idxMid]) || 70,
          finalMax: Number(row[idxFinal]) || 30
        };
      }
    }
    
    return { fullScores: [0,0,0,0,0,0,0,0,0], midMax: 70, finalMax: 30 };
  } catch (e) {
    Logger.log('getSubjectFullScores error:', e.message);
    return { fullScores: [0,0,0,0,0,0,0,0,0], midMax: 70, finalMax: 30 };
  }
}

/**
 * � นำเข้าโครงสร้างคะแนนจาก CSV (base64) เข้าชีตรายวิชา
 * CSV format (2-row header):
 *   ชั้น, รหัสวิชา, ชื่อวิชา, ชั่วโมง/ปี, ประเภทวิชา,
 *   ค1, ค2, ค3, ค4, รวม14,
 *   ค5(สอบกลางภาค),
 *   ค6, ค7, ค8, รวม68,
 *   รวมระหว่างภาค, ปลายภาค, รวมทั้งหมด, ครูผู้สอน
 * @param {string} csvBase64 - CSV content encoded as base64
 * @returns {Object} { success, count, message }
 */
function importSubjectsFromCSV(csvBase64) {
  try {
    // decode base64 → string
    var csvText = Utilities.newBlob(Utilities.base64Decode(csvBase64)).getDataAsString('UTF-8');
    var lines = csvText.split(/\r?\n/).filter(function(l) { return l.trim().length > 0; });
    
    // skip 2 header rows
    if (lines.length < 3) throw new Error('ไฟล์ CSV ต้องมีอย่างน้อย 3 แถว (2 header + 1 data)');
    var dataLines = lines.slice(2);
    
    var ss = _openSpreadsheet_();
    var sheet = ss.getSheetByName(SUBJECT_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SUBJECT_SHEET_NAME);
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setValues([SUBJECT_HEADERS_ALL]);
      sheet.getRange(1, 1, 1, SUBJECT_HEADERS_ALL.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      ensureFullScoreHeaders_(sheet);
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var existingData = sheet.getDataRange().getValues();
    
    // สร้าง map ของรายวิชาที่มีอยู่แล้ว: key = "ชั้น|ชื่อวิชา"
    var existingMap = {};
    for (var e = 1; e < existingData.length; e++) {
      var key = String(existingData[e][0]).trim() + '|' + String(existingData[e][2]).trim();
      existingMap[key] = e + 1; // row number (1-indexed)
    }
    
    var addCount = 0;
    var updateCount = 0;
    
    dataLines.forEach(function(line) {
      // parse CSV line (simple split — ไม่มี quoted fields ที่ซับซ้อน)
      var cols = line.split(',').map(function(c) { return c.trim(); });
      if (cols.length < 18) return;
      
      // CSV columns mapping:
      // 0=ชั้น, 1=รหัสวิชา, 2=ชื่อวิชา, 3=ชั่วโมง/ปี, 4=ประเภทวิชา
      // 5=ค1, 6=ค2, 7=ค3, 8=ค4, 9=รวม14
      // 10=ค5(สอบกลางภาค)
      // 11=ค6, 12=ค7, 13=ค8, 14=รวม68
      // 15=รวมระหว่างภาค, 16=ปลายภาค, 17=รวมทั้งหมด
      // 18=ครูผู้สอน (อาจมีช่องว่างหลายตัว)
      
      var subjectData = {};
      subjectData['ชั้น'] = cols[0];
      subjectData['รหัสวิชา'] = cols[1];
      subjectData['ชื่อวิชา'] = cols[2];
      subjectData['ชั่วโมง/ปี'] = cols[3];
      subjectData['ประเภทวิชา'] = cols[4];
      subjectData['ครูผู้สอน'] = cols.slice(18).join(',').trim(); // กัน comma ในชื่อ
      subjectData['คะแนนระหว่างปี'] = cols[15];
      subjectData['คะแนนปลายปี'] = cols[16];
      subjectData['รวม'] = cols[17];
      
      // คะแนนเต็มรายตัวชี้วัด
      subjectData['เต็ม_ครั้ง1'] = cols[5];
      subjectData['เต็ม_ครั้ง2'] = cols[6];
      subjectData['เต็ม_ครั้ง3'] = cols[7];
      subjectData['เต็ม_ครั้ง4'] = cols[8];
      subjectData['เต็ม_ครั้ง5'] = cols[10]; // สอบกลางภาค
      subjectData['เต็ม_ครั้ง6'] = cols[11];
      subjectData['เต็ม_ครั้ง7'] = cols[12];
      subjectData['เต็ม_ครั้ง8'] = cols[13];
      subjectData['เต็ม_ครั้ง9'] = '0'; // ไม่มีใน CSV
      
      var newRow = headers.map(function(h) { return subjectData[h] || ''; });
      
      var key = String(cols[0]).trim() + '|' + String(cols[2]).trim();
      if (existingMap[key]) {
        // update existing row
        sheet.getRange(existingMap[key], 1, 1, newRow.length).setValues([newRow]);
        updateCount++;
      } else {
        // append new row
        sheet.appendRow(newRow);
        existingMap[key] = sheet.getLastRow();
        addCount++;
      }
    });
    
    return {
      success: true,
      count: addCount + updateCount,
      message: 'นำเข้าสำเร็จ: เพิ่มใหม่ ' + addCount + ' รายการ, อัปเดต ' + updateCount + ' รายการ'
    };
    
  } catch (e) {
    Logger.log('importSubjectsFromCSV error:', e);
    return { success: false, count: 0, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**
 * �🔧 ฟังก์ชันสำหรับสร้างโฟลเดอร์ (ถ้ายังไม่มีในไฟล์อื่น)
 */
function _getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

