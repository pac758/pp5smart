/**
 * ============================================================
 * 📊 III. PDF TEMPLATE MANAGEMENT
 * ============================================================
 */

const TEMPLATE_SHEET_NAMES = {
  readThinkWrite: 'Template_ReadThinkWrite',
  characteristic: 'Template_Characteristic',
  activity: 'Template_Activity',
  competencyDetailed: 'Template_CompetencyDetailed',
  competencySummary: 'Template_CompetencySummary'
};

/**
 * ✅ สร้าง Template ทั้งหมดในชีตหลัก (ต้องรันครั้งแรก)
 */
function createAllTemplatesInSameSheet() {
  try {
    const ss = SS();
    createReadThinkWriteTemplate(ss);
    createCharacteristicTemplate(ss);
    createActivityTemplate(ss);
    createCompetencyDetailedTemplate(ss);
    createCompetencySummaryTemplate(ss);
    return { success: true, message: 'สร้าง Template ทั้งหมดเสร็จสิ้น' };
  } catch (error) {
    throw new Error(`ไม่สามารถสร้าง Template ได้: ${error.message}`);
  }
}

// ============================================================
// 🏗️ CORE TEMPLATE PROCESSING ENGINE
// ============================================================

/**
 * สร้าง PDF จาก Template ในชีตเดียวกัน (Core Function)
 * รับข้อมูลจากไฟล์ Assessments.gs มาเติมใน Template
 */
function generatePDFFromSameSheetTemplate(templateType, grade, classNo, studentsData) {
  try {
    // สร้าง Google Doc แทน Sheet
    const settings = getGlobalSettings();
    const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียนของคุณ';
    const academicYear = settings['ปีการศึกษา'] || String(S_getCurrentAcademicYear_());
    const logoFileId = settings['logoFileId'];
    
    const docName = `ประเมินอ่าน-คิด-เขียน_${grade}_${classNo}_${new Date().toISOString().slice(0,10)}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    // เพิ่มโลโก้
    if (logoFileId && logoFileId.trim() !== '') {
      try {
        const logoFile = DriveApp.getFileById(logoFileId.trim());
        const logoBlob = logoFile.getBlob();
        const logoImage = body.appendImage(logoBlob);
        logoImage.setWidth(80).setHeight(80);
        
        // จัดให้อยู่ตรงกลาง
        const logoPara = logoImage.getParent().asParagraph();
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      } catch (error) {
        Logger.log('⚠️ ไม่สามารถแทรกโลโก้ได้: ' + error.message);
      }
    }
    
    // หัวเอกสาร
    const title = body.appendParagraph('รายงานการประเมินอ่าน คิด วิเคราะห์ เขียน');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    body.appendParagraph(`โรงเรียน: ${schoolName}`);
    body.appendParagraph(`ชั้น: ${U_getGradeFullName(grade)} ห้อง ${classNo}`);
    body.appendParagraph(`ปีการศึกษา: ${academicYear}`);
    body.appendParagraph(''); // บรรทัดว่าง
    
    // สร้างตาราง
    const tableData = [
      ['ที่', 'ชื่อ-นามสกุล', 'อ่าน 1', 'อ่าน 2', 'อ่าน 3', 
       'คิด 1', 'คิด 2', 'คิด 3', 'เขียน 1', 'เขียน 2', 'เขียน 3',
       'รวม', 'เฉลี่ย', 'ผลการประเมิน']
    ];
    
    studentsData.forEach((student, index) => {
      const a = student.assessment || {};
      const scores = [a.read1||'-', a.read2||'-', a.read3||'-', 
                      a.think1||'-', a.think2||'-', a.think3||'-',
                      a.write1||'-', a.write2||'-', a.write3||'-'];
      
      const validScores = scores.filter(s => s !== '-').map(Number);
      const total = validScores.length === 9 ? validScores.reduce((sum, s) => sum + s, 0) : '-';
      const avg = validScores.length === 9 ? (total / 9).toFixed(1) : '-';
      
      let result = '-';
      if (avg !== '-') {
        const avgNum = parseFloat(avg);
        result = avgNum >= 3 ? 'ดีเยี่ยม' : avgNum >= 2 ? 'ดี' : avgNum >= 1 ? 'พอใช้' : 'ปรับปรุง';
      }
      
      tableData.push([
        String(index + 1),
        student.name,
        ...scores.map(s => String(s)),
        String(total),
        String(avg),
        result
      ]);
    });
    
    const table = body.appendTable(tableData);
    
    // จัดรูปแบบตาราง
    for (let i = 0; i < table.getNumRows(); i++) {
      const row = table.getRow(i);
      for (let j = 0; j < row.getNumCells(); j++) {
        const cell = row.getCell(j);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        
        if (i === 0) {
          // หัวตาราง
          cell.setBackgroundColor('#ffc107');
          cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        } else {
          // ข้อมูล
          if (j === 0 || j >= 2) {
            cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          }
        }
      }
    }
    
    // บันทึกและแปลงเป็น PDF
    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(docName + '.pdf');
    
    // ลบ Doc เดิม
    docFile.setTrashed(true);
    
    return pdfFile.getUrl();
    
  } catch (error) {
    throw new Error(`ไม่สามารถสร้าง PDF ได้: ${error.message}`);
  }
}

/**
 * เติมข้อมูลลง Template
 */

/**
 * จัดรูปแบบข้อมูลใน Template
 */
function formatTemplateData(sheet, templateType, studentCount) {
  const startRow = 8; // เปลี่ยนจาก 7 เป็น 8
  const columnCount = getTemplateColumnCount(templateType);
  
  if (studentCount > 0) {
    const dataRange = sheet.getRange(startRow, 1, studentCount, columnCount);
    dataRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
    
    for (let i = 0; i < studentCount; i++) {
      if (i % 2 === 0) {
        sheet.getRange(startRow + i, 1, 1, columnCount).setBackground('#F8F9FA');
      }
    }
    
    sheet.getRange(startRow, 1, studentCount, 1).setHorizontalAlignment('center');
    formatSpecificTemplateType(sheet, templateType, startRow, studentCount);
  }
}

/**
 * Export Sheet เป็น PDF
 */
function exportSheetToPDF(spreadsheetId, sheetId, fileName) {
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    'format=pdf&' + 'size=A4&' + 'portrait=true&' + 'fitw=true&' + 'gridlines=false&' + 'printtitle=false&' + 'pagenumbers=false&' + 'fzr=false&' + `gid=${sheetId}`;
  const response = UrlFetchApp.fetch(exportUrl, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
  const pdfBlob = response.getBlob().setName(fileName + '.pdf');
  return DriveApp.createFile(pdfBlob).getUrl();
}

// ============================================================
// 🛠️ HELPER FUNCTIONS
// ============================================================

/**
 * สร้าง Template สำหรับ Read Think Write
 */
function createReadThinkWriteTemplate(spreadsheet) {
  const sheetName = TEMPLATE_SHEET_NAMES.readThinkWrite;
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const sheet = spreadsheet.insertSheet(sheetName);
  
  sheet.getRange('A1:N50').setFontFamily('Sarabun').setFontSize(12);
  
  // ⭐ Debug: ตรวจสอบการดึงโลโก้
  try {
    const settings = getGlobalSettings();
    Logger.log('📋 Settings ทั้งหมด: ' + JSON.stringify(settings));
    
    const logoFileId = settings['logoFileId'];
    Logger.log('🆔 Logo File ID: ' + logoFileId);
    
    if (logoFileId && logoFileId.toString().trim() !== '') {
      Logger.log('✅ กำลังดึงโลโก้...');
      
      const logoFile = DriveApp.getFileById(logoFileId.toString().trim());
      Logger.log('📁 ชื่อไฟล์: ' + logoFile.getName());
      
      const logoBlob = logoFile.getBlob();
      Logger.log('🖼️ ขนาด Blob: ' + logoBlob.getBytes().length);
      
      // แทรกรูปที่คอลัมน์ G (7) แถวที่ 1
      const logo = sheet.insertImage(logoBlob, 7, 1);
      logo.setWidth(80);
      logo.setHeight(80);
      
      sheet.setRowHeight(1, 90);
      Logger.log('✅ แทรกโลโก้สำเร็จ');
    } else {
      Logger.log('⚠️ ไม่พบ logoFileId ใน Settings');
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    Logger.log('❌ Stack: ' + error.stack);
  }
  
  // ส่วนที่เหลือเหมือนเดิม...
  sheet.getRange('A2:N2').merge()
    .setValue('รายงานการประเมินอ่าน คิด วิเคราะห์ เขียน')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#f8f9fa');
  
  setupHeaderInfo(sheet, 'A3', 'โรงเรียน:', '[จะใส่ชื่อโรงเรียนจากระบบ]');
  setupHeaderInfo(sheet, 'A4', 'ชั้น:', '[จะใส่ชั้น ห้อง จากระบบ]');
  setupHeaderInfo(sheet, 'A5', 'ปีการศึกษา:', '[จะใส่ปีการศึกษาจากระบบ]');
  
  sheet.getRange('A6:N6').setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  
  const headers = ['ที่', 'ชื่อ-นามสกุล', 'อ่าน\n1', 'อ่าน\n2', 'อ่าน\n3', 
                   'คิด\n1', 'คิด\n2', 'คิด\n3', 'เขียน\n1', 'เขียน\n2', 'เขียน\n3', 
                   'รวม', 'เฉลี่ย', 'ผลการประเมิน'];
  
  sheet.getRange(7, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(7, 1, 1, headers.length);
  
  headerRange.setBackground('#ffc107')
    .setFontWeight('bold')
    .setFontColor('#000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  
  for (let col = 3; col <= 11; col++) sheet.getRange(7, col).setTextRotation(90);
  
  setColumnWidths(sheet, [40, 180, 45, 45, 45, 45, 45, 45, 45, 45, 45, 50, 50, 100]);
  sheet.setRowHeight(7, 80);
}

/**
 * สร้าง Template สำหรับ Characteristic
 */
function createCharacteristicTemplate(spreadsheet) {
  const sheetName = TEMPLATE_SHEET_NAMES.characteristic;
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange('A1:M50').setFontFamily('Sarabun').setFontSize(12);
  sheet.getRange('A1:M1').merge().setValue('รายงานการประเมินคุณลักษณะอันพึงประสงค์').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f8f9fa');
  setupHeaderInfo(sheet, 'A2', 'โรงเรียน:', '[จะใส่ชื่อโรงเรียนจากระบบ]');
  setupHeaderInfo(sheet, 'A3', 'ชั้น:', '[จะใส่ชั้น ห้อง จากระบบ]');
  setupHeaderInfo(sheet, 'A4', 'ปีการศึกษา:', '[จะใส่ปีการศึกษาจากระบบ]');
  sheet.getRange('A5:M5').setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  const headers = ['ที่', 'ชื่อ-นามสกุล', 'รักชาติ\nศาสน์\nกษัตริย์', 'ซื่อสัตย์\nสุจริต', 'มีวินัย', 'ใฝ่เรียนรู้', 'อยู่อย่าง\nพอเพียง', 'มุ่งมั่น\nในการ\nทำงาน', 'รักความ\nเป็นไทย', 'มีจิต\nสาธารณะ', 'รวม', 'เฉลี่ย', 'ผลการประเมิน'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(6, 1, 1, headers.length);
  headerRange.setBackground('#0dcaf0').setFontWeight('bold').setFontColor('#000').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  for (let col = 3; col <= 10; col++) sheet.getRange(6, col).setTextRotation(90);
  setColumnWidths(sheet, [40, 160, 50, 50, 50, 50, 50, 50, 50, 50, 50, 50, 100]);
  sheet.setRowHeight(6, 80);
}

/**
 * สร้าง Template สำหรับ Activity
 */
function createActivityTemplate(spreadsheet) {
  const sheetName = TEMPLATE_SHEET_NAMES.activity;
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange('A1:G50').setFontFamily('Sarabun').setFontSize(12);
  sheet.getRange('A1:G1').merge().setValue('รายงานการประเมินกิจกรรมพัฒนาผู้เรียน').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f8f9fa');
  setupHeaderInfo(sheet, 'A2', 'โรงเรียน:', '[จะใส่ชื่อโรงเรียนจากระบบ]');
  setupHeaderInfo(sheet, 'A3', 'ชั้น:', '[จะใส่ชั้น ห้อง จากระบบ]');
  setupHeaderInfo(sheet, 'A4', 'ปีการศึกษา:', '[จะใส่ปีการศึกษาจากระบบ]');
  sheet.getRange('A5:G5').setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  const headers = ['ที่', 'ชื่อ-นามสกุล', 'กิจกรรม\nแนะแนว', 'ลูกเสือ/\nเนตรนารี', 'ชุมนุม', 'เพื่อสังคม\nและ\nสาธารณประโยชน์', 'รวมกิจกรรม'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(6, 1, 1, headers.length);
  headerRange.setBackground('#28a745').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  for (let col = 3; col <= 7; col++) sheet.getRange(6, col).setTextRotation(90);
  setColumnWidths(sheet, [40, 200, 65, 65, 65, 65, 65]);
  sheet.setRowHeight(6, 80);
}

/**
 * สร้าง Template สำหรับ Competency Detailed
 */
function createCompetencyDetailedTemplate(spreadsheet) {
  const sheetName = TEMPLATE_SHEET_NAMES.competencyDetailed;
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange('A1:AD50').setFontFamily('Sarabun').setFontSize(10);
  sheet.getRange('A1:AD1').merge().setValue('รายงานการประเมินสมรรถนะ (แบบละเอียด)').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f8f9fa');
  setupHeaderInfo(sheet, 'A2', 'โรงเรียน:', '[จะใส่ชื่อโรงเรียนจากระบบ]');
  setupHeaderInfo(sheet, 'A3', 'ชั้น:', '[จะใส่ชั้น ห้อง จากระบบ]');
  setupHeaderInfo(sheet, 'A4', 'ปีการศึกษา:', '[จะใส่ปีการศึกษาจากระบบ]');
  sheet.getRange('A5:AD5').setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  const headers = ['ที่', 'ชื่อ-นามสกุล', ...Array.from({ length: 25 }, (_, i) => (i + 1).toString()), 'รวม', 'เฉลี่ย', 'ผลการประเมิน'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(6, 1, 1, headers.length);
  headerRange.setBackground('#6f42c1').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  for (let col = 3; col <= 27; col++) sheet.getRange(6, col).setTextRotation(90);
  const widths = [30, 140, ...Array(25).fill(18), 40, 40, 70];
  setColumnWidths(sheet, widths);
  sheet.setRowHeight(6, 60);
}

/**
 * สร้าง Template สำหรับ Competency Summary
 */
function createCompetencySummaryTemplate(spreadsheet) {
  const sheetName = TEMPLATE_SHEET_NAMES.competencySummary;
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  const sheet = spreadsheet.insertSheet(sheetName);
  sheet.getRange('A1:J50').setFontFamily('Sarabun').setFontSize(12);
  sheet.getRange('A1:J1').merge().setValue('รายงานการประเมินสมรรถนะ (แบบสรุป)').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f8f9fa');
  setupHeaderInfo(sheet, 'A2', 'โรงเรียน:', '[จะใส่ชื่อโรงเรียนจากระบบ]');
  setupHeaderInfo(sheet, 'A3', 'ชั้น:', '[จะใส่ชั้น ห้อง จากระบบ]');
  setupHeaderInfo(sheet, 'A4', 'ปีการศึกษา:', '[จะใส่ปีการศึกษาจากระบบ]');
  sheet.getRange('A5:J5').setBorder(false, false, true, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID);
  const headers = ['ที่', 'ชื่อ-นามสกุล', 'การสื่อสาร', 'การคิด', 'แก้ปัญหา', 'ทักษะชีวิต', 'เทคโนโลยี', 'รวม', 'เฉลี่ย', 'ผลการประเมิน'];
  sheet.getRange(6, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(6, 1, 1, headers.length);
  headerRange.setBackground('#6f42c1').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  for (let col = 3; col <= 7; col++) sheet.getRange(6, col).setTextRotation(90);
  setColumnWidths(sheet, [40, 140, 60, 60, 60, 60, 60, 50, 50, 90]);
  sheet.setRowHeight(6, 80);
}

/**
 * เติมข้อมูลลง Template
 */

/**
 * เติมข้อมูลหัวรายงาน
 */
function fillHeaderData(sheet, grade, classNo) {
  try {
    const settings = getGlobalSettings();
    const fullGradeName = U_getGradeFullName(grade);
    
    // ขยับลงมาจากแถว 2,3,4 เป็น 3,4,5
    sheet.getRange('B3').setValue(settings['ชื่อโรงเรียน'] || 'โรงเรียนของคุณ')
      .setFontStyle('normal').setFontColor('#000');
    sheet.getRange('B4').setValue(`${fullGradeName} ห้อง ${classNo}`)
      .setFontStyle('normal').setFontColor('#000');
    sheet.getRange('B5').setValue(`ปีการศึกษา ${settings['ปีการศึกษา'] || String(S_getCurrentAcademicYear_())}`)
      .setFontStyle('normal').setFontColor('#000');
  } catch (error) {
    Logger.log('⚠️ ไม่สามารถเติมข้อมูลหัวรายงานได้:', error.message);
  }
}

/**
 * เติมข้อมูลการประเมิน Read Think Write
 */
function fillReadThinkWriteData(sheet, studentsData) {
  const startRow = 8; // เปลี่ยนจาก 7 เป็น 8
  studentsData.forEach((student, index) => {
    const row = startRow + index;
    const assessment = student.assessment || {};
    sheet.getRange(row, 1).setValue(index + 1);
    sheet.getRange(row, 2).setValue(student.name || '');
    
    const scores = [
      assessment.read1 || '', assessment.read2 || '', assessment.read3 || '', 
      assessment.think1 || '', assessment.think2 || '', assessment.think3 || '', 
      assessment.write1 || '', assessment.write2 || '', assessment.write3 || ''
    ];
    
    for (let i = 0; i < 9; i++) sheet.getRange(row, 3 + i).setValue(scores[i]);
    
    const validScores = scores.filter(s => s !== '' && !isNaN(s)).map(Number);
    const total = validScores.reduce((sum, s) => sum + s, 0);
    const average = validScores.length === 9 ? (total / 9).toFixed(1) : '';
    
    sheet.getRange(row, 12).setValue(validScores.length === 9 ? total : '');
    sheet.getRange(row, 13).setValue(average);
    
    let result = '';
    if (average !== '') {
      const avg = parseFloat(average);
      result = avg >= 3 ? 'ดีเยี่ยม' : avg >= 2 ? 'ดี' : avg >= 1 ? 'พอใช้' : 'ปรับปรุง';
    }
    sheet.getRange(row, 14).setValue(result);
  });
}

/**
 * เติมข้อมูลการประเมินคุณลักษณะ
 */
function fillCharacteristicData(sheet, studentsData) {
  const startRow = 7;
  studentsData.forEach((student, index) => {
    const row = startRow + index;
    const scores = student.scores || [];
    sheet.getRange(row, 1).setValue(index + 1);
    sheet.getRange(row, 2).setValue(student.name || '');
    for (let i = 0; i < 8; i++) sheet.getRange(row, 3 + i).setValue(scores[i] || '');
    const validScores = scores.filter(s => s !== '' && !isNaN(s)).map(Number);
    const total = validScores.reduce((sum, s) => sum + s, 0);
    const average = validScores.length === 8 ? (total / 8).toFixed(1) : '';
    sheet.getRange(row, 11).setValue(validScores.length === 8 ? total : '');
    sheet.getRange(row, 12).setValue(average);
    let result = '';
    if (average !== '') {
      const avg = parseFloat(average);
      result = avg >= 2.50 ? 'ดีเยี่ยม' : avg >= 2.00 ? 'ดี' : avg >= 1.00 ? 'ผ่าน' : 'ปรับปรุง';
    }
    sheet.getRange(row, 13).setValue(result);
  });
}

/**
 * เติมข้อมูลการประเมินกิจกรรม
 */
function fillActivityData(sheet, studentsData) {
  const startRow = 7;
  studentsData.forEach((student, index) => {
    const row = startRow + index;
    const activities = student.activities || {};
    sheet.getRange(row, 1).setValue(index + 1);
    sheet.getRange(row, 2).setValue(student.name || '');
    sheet.getRange(row, 3).setValue(activities.guidance || 'ผ่าน');
    sheet.getRange(row, 4).setValue(activities.scout || 'ผ่าน');
    sheet.getRange(row, 5).setValue(activities.club || 'ผ่าน');
    sheet.getRange(row, 6).setValue(activities.social || 'ผ่าน');
    sheet.getRange(row, 7).setValue(activities.overall || 'ผ่าน');
  });
}

/**
 * เติมข้อมูลการประเมินสมรรถนะแบบละเอียด
 */
function fillCompetencyDetailedData(sheet, studentsData) {
  const startRow = 7;
  studentsData.forEach((student, index) => {
    const row = startRow + index;
    const scores = student.scores || [];
    sheet.getRange(row, 1).setValue(index + 1);
    sheet.getRange(row, 2).setValue(student.name || '');
    for (let i = 0; i < 25; i++) sheet.getRange(row, 3 + i).setValue(scores[i] || '');
    const validScores = scores.filter(s => s !== '' && !isNaN(s)).map(Number);
    const total = validScores.reduce((sum, s) => sum + s, 0);
    const average = validScores.length === 25 ? Math.round(total / 25) : '';
    sheet.getRange(row, 28).setValue(validScores.length === 25 ? total : '');
    sheet.getRange(row, 29).setValue(average);
    let result = '';
    if (average !== '') {
      const avg = Number(average);
      result = avg >= 3 ? 'ดีเยี่ยม' : avg >= 2 ? 'ดี' : avg >= 1 ? 'พอใช้' : 'ปรับปรุง';
    }
    sheet.getRange(row, 30).setValue(result);
  });
}

/**
 * เติมข้อมูลการประเมินสมรรถนะแบบสรุป
 */
function fillCompetencySummaryData(sheet, studentsData) {
  const startRow = 7;
  studentsData.forEach((student, index) => {
    const row = startRow + index;
    const scores = student.scores || [];
    sheet.getRange(row, 1).setValue(index + 1);
    sheet.getRange(row, 2).setValue(student.name || '');
    for (let domain = 0; domain < 5; domain++) {
      let domainTotal = 0;
      let validCount = 0;
      for (let item = 0; item < 5; item++) {
        const scoreIndex = domain * 5 + item;
        const score = scores[scoreIndex];
        if (score !== '' && !isNaN(score)) {
          domainTotal += Number(score);
          validCount++;
        }
      }
      sheet.getRange(row, 3 + domain).setValue(validCount === 5 ? domainTotal : '');
    }
    const validScores = scores.filter(s => s !== '' && !isNaN(s)).map(Number);
    const total = validScores.reduce((sum, s) => sum + s, 0);
    const average = validScores.length === 25 ? (total / 25).toFixed(1) : '';
    sheet.getRange(row, 8).setValue(validScores.length === 25 ? total : '');
    sheet.getRange(row, 9).setValue(average);
    let result = '';
    if (average !== '') {
      const avg = parseFloat(average);
      result = avg >= 2.50 ? 'ดีเยี่ยม' : avg >= 2.00 ? 'ดี' : avg >= 1.00 ? 'พอใช้' : 'ปรับปรุง';
    }
    sheet.getRange(row, 10).setValue(result);
  });
}

/**
 * จัดรูปแบบเฉพาะตามประเภท Template
 */
function formatSpecificTemplateType(sheet, templateType, startRow, studentCount) {
  switch (templateType) {
    case 'readThinkWrite':
      formatReadThinkWriteData(sheet, startRow, studentCount);
      break;
    case 'characteristic':
      formatCharacteristicData(sheet, startRow, studentCount);
      break;
    case 'activity':
      formatActivityData(sheet, startRow, studentCount);
      break;
    case 'competencyDetailed':
    case 'competencySummary':
      formatCompetencyData(sheet, startRow, studentCount, templateType);
      break;
  }
}

/**
 * จัดรูปแบบ Read Think Write
 */
function formatReadThinkWriteData(sheet, startRow, studentCount) {
  const scoreRange = sheet.getRange(startRow, 3, studentCount, 11);
  scoreRange.setHorizontalAlignment('center');
  formatAssessmentResults(sheet, startRow, studentCount, 14);
}

/**
 * จัดรูปแบบ Characteristic
 */
function formatCharacteristicData(sheet, startRow, studentCount) {
  const scoreRange = sheet.getRange(startRow, 3, studentCount, 8);
  scoreRange.setHorizontalAlignment('center');
  formatAssessmentResults(sheet, startRow, studentCount, 13);
}

/**
 * จัดรูปแบบ Activity
 */
function formatActivityData(sheet, startRow, studentCount) {
  const activityRange = sheet.getRange(startRow, 3, studentCount, 5);
  activityRange.setHorizontalAlignment('center');
  const values = activityRange.getValues();
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cell = sheet.getRange(startRow + i, 3 + j);
      const value = String(values[i][j]).trim();
      if (value === 'ผ่าน') cell.setBackground('#d4edda').setFontColor('#155724');
      else if (value === 'ไม่ผ่าน') cell.setBackground('#f8d7da').setFontColor('#721c24');
    }
  }
}

/**
 * จัดรูปแบบ Competency
 */
function formatCompetencyData(sheet, startRow, studentCount, templateType) {
  const scoreStartCol = 3;
  const scoreColCount = templateType === 'competencyDetailed' ? 25 : 5;
  const scoreRange = sheet.getRange(startRow, scoreStartCol, studentCount, scoreColCount);
  scoreRange.setHorizontalAlignment('center');
  const resultCol = getTemplateColumnCount(templateType);
  formatAssessmentResults(sheet, startRow, studentCount, resultCol);
}

/**
 * จัดรูปแบบผลการประเมิน (สีตัวอักษร)
 */
function formatAssessmentResults(sheet, startRow, studentCount, resultColumn) {
  const resultRange = sheet.getRange(startRow, resultColumn, studentCount, 1);
  resultRange.setHorizontalAlignment('center').setFontWeight('bold');
  const values = resultRange.getValues();
  for (let i = 0; i < values.length; i++) {
    const cell = sheet.getRange(startRow + i, resultColumn);
    const value = String(values[i][0]).trim();
    switch (value) {
      case 'ดีเยี่ยม':
        cell.setFontColor('#198754');
        break;
      case 'ดี':
        cell.setFontColor('#0D6EFD');
        break;
      case 'พอใช้':
      case 'ผ่าน':
        cell.setFontColor('#FFC107');
        break;
      case 'ปรับปรุง':
        cell.setFontColor('#DC3545');
        break;
    }
  }
}

/**
 * ตั้งค่าข้อมูลหัวรายงาน
 */
function setupHeaderInfo(sheet, labelCell, label, placeholder) {
  sheet.getRange(labelCell).setValue(label).setFontWeight('bold');
  sheet.getRange(labelCell.replace(/\d/, (m) => String.fromCharCode(m.charCodeAt(0) + 1))).setValue(placeholder).setFontStyle('italic').setFontColor('#666666');
}

/**
 * ตั้งความกว้างคอลัมน์
 */
function setColumnWidths(sheet, widths) {
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
}

/**
 * ได้จำนวนคอลัมน์ตามประเภท Template
 */
function getTemplateColumnCount(templateType) {
  switch (templateType) {
    case 'readThinkWrite': return 14;
    case 'characteristic': return 13;
    case 'activity': return 7;
    case 'competencyDetailed': return 30;
    case 'competencySummary': return 10;
    default: return 10;
  }
}

/**
 * ฟังก์ชันแปลงชั้นเรียนให้เป็นรูปแบบเต็ม
 */

