/**************************************************
 * 📚 Subject Management Functions
 * ฟังก์ชันสำหรับจัดการรายวิชาใน Google Sheets
 **************************************************/

const SUBJECT_SHEET_NAME = 'รายวิชา';

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
      const headers = [
        'ชั้น', 'รหัสวิชา', 'ชื่อวิชา', 'ชั่วโมง/ปี', 
        'ประเภทวิชา', 'ครูผู้สอน', 'คะแนนระหว่างปี', 
        'คะแนนปลายปี', 'รวม'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      
      console.log('✅ สร้างชีต "รายวิชา" ใหม่พร้อมหัวตาราง');
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

    console.log(`✅ ดึงข้อมูลรายวิชา: ${subjects.length} รายการ`);
    return subjects;

  } catch (error) {
    console.error('❌ getSubjects:', error);
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
      // สร้างชีตใหม่หากไม่มี
      sheet = ss.insertSheet(SUBJECT_SHEET_NAME);
      const headers = [
        'ชั้น', 'รหัสวิชา', 'ชื่อวิชา', 'ชั่วโมง/ปี', 
        'ประเภทวิชา', 'ครูผู้สอน', 'คะแนนระหว่างปี', 
        'คะแนนปลายปี', 'รวม'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
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
    
    console.log(`✅ เพิ่มรายวิชา: ${subjectData['ชื่อวิชา']} ชั้น ${subjectData['ชั้น']}`);
    return 'เพิ่มรายวิชาสำเร็จ';

  } catch (error) {
    console.error('❌ addSubject:', error);
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
    
    console.log(`✅ แก้ไขรายวิชา: ${subjectData['ชื่อวิชา']} ชั้น ${subjectData['ชั้น']}`);
    return 'แก้ไขรายวิชาสำเร็จ';

  } catch (error) {
    console.error('❌ updateSubject:', error);
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
    
    console.log(`✅ ลบรายวิชา: ${subjectName} ชั้น ${grade}`);
    return 'ลบรายวิชาสำเร็จ';

  } catch (error) {
    console.error('❌ deleteSubject:', error);
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
    const folder = _getOrCreateFolder_('Subject Reports');
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    console.log(`✅ ส่งออกรายวิชาทั้งหมด: ${subjects.length} รายการ`);
    return file.getUrl();

  } catch (error) {
    console.error('❌ exportAllSubjectsAsPdf:', error);
    throw new Error(`ไม่สามารถส่งออก PDF ได้: ${error.message}`);
  }
}

/**
 * 🏗️ สร้าง HTML สำหรับรายงานรายวิชาทั้งหมด
 */
function _buildSubjectsListHTML_(subjects) {
  const settings = getGlobalSettingsFixed();
  const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียน';
  const academicYear = settings['ปีการศึกษา'] || '2568';
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
 * 🔧 ฟังก์ชันสำหรับสร้างโฟลเดอร์ (ถ้ายังไม่มีในไฟล์อื่น)
 */
function _getOrCreateFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

