/**
 * ฟังก์ชันสร้าง PDF รายชื่อนักเรียนพร้อมช่องคะแนน 1-10
 * ใช้ Google Sheets สร้างตารางแล้ว export เป็น PDF (ฟอนต์ TH Sarabun New คมชัด)
 */
function generateStudentListPDF(grade, classNo, gender) {
  var tmpSheet = null;
  try {
    Logger.log('สร้าง PDF: ชั้น=' + grade + ', ห้อง=' + classNo + ', เพศ=' + gender);

    var settings = S_getGlobalSettings();
    var schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียนตัวอย่าง';
    var academicYear = settings['ปีการศึกษา'] || '2567';
    var semester = settings['ภาคเรียน'] || '1';

    var students = getFilteredStudents(grade, classNo, gender);
    if (students.length === 0) {
      throw new Error('ไม่พบข้อมูลนักเรียนที่ตรงตามเงื่อนไข');
    }

    // --- แปลงชื่อชั้นเรียน ---
    var gradeMap = {
      'อ.1':'อนุบาลปีที่ 1','อ.2':'อนุบาลปีที่ 2','อ.3':'อนุบาลปีที่ 3',
      'ป.1':'ประถมศึกษาปีที่ 1','ป.2':'ประถมศึกษาปีที่ 2','ป.3':'ประถมศึกษาปีที่ 3',
      'ป.4':'ประถมศึกษาปีที่ 4','ป.5':'ประถมศึกษาปีที่ 5','ป.6':'ประถมศึกษาปีที่ 6'
    };
    var gradeText = (grade && gradeMap[grade]) ? gradeMap[grade] : (grade || 'ทุกระดับ');
    var classText = classNo ? 'ห้อง ' + classNo : 'ทุกห้อง';
    var genderText = gender ? ' (' + gender + ')' : '';

    // --- สร้าง temp sheet ---
    var ss = SS();
    var sheetName = '_TMP_PDF_' + new Date().getTime();
    tmpSheet = ss.insertSheet(sheetName);

    var FONT = 'TH SarabunPSK';
    // คอลัมน์: A=ที่, B=รหัส, C=บัตร ปชช., D=ชื่อ-สกุล, E=อายุ, F=วันเกิด, G-P=คะแนน 1-10
    var totalCols = 16;

    // --- Row 1: โลโก้โรงเรียน ---
    tmpSheet.setRowHeight(1, 50);
    tmpSheet.getRange(1, 1, 1, totalCols).merge(); // merge เพื่อเว้นที่ให้โลโก้
    try {
      var logoFileId = settings['logoFileId'];
      if (logoFileId) {
        var logoBlob = DriveApp.getFileById(logoFileId).getBlob();
        // คำนวณตำแหน่งกึ่งกลาง: ความกว้างรวม ~688px, โลโก้ 45px → offset ~322px
        var img = tmpSheet.insertImage(logoBlob, 1, 1, 322, 3);
        img.setWidth(45).setHeight(45);
      }
    } catch (logoErr) {
      Logger.log('Logo insert error: ' + logoErr.message);
    }

    // --- Row 2: ชื่อโรงเรียน ---
    tmpSheet.getRange(2, 1).setValue(schoolName);
    tmpSheet.getRange(2, 1, 1, totalCols).merge()
      .setFontFamily(FONT).setFontSize(18).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    tmpSheet.setRowHeight(2, 30);

    // --- Row 3: ข้อมูลชั้น/ห้อง ---
    tmpSheet.getRange(3, 1).setValue('รายชื่อนักเรียน ชั้น ' + gradeText + ' ' + classText + genderText);
    tmpSheet.getRange(3, 1, 1, totalCols).merge()
      .setFontFamily(FONT).setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    tmpSheet.setRowHeight(3, 26);

    // --- Row 4-5: หัวตาราง ---
    var mainHeaders = ['ที่', 'รหัส', 'บัตร ปชช.', 'ชื่อ - สกุล', 'อายุ(ปี)', 'วันเกิด'];
    for (var h = 0; h < mainHeaders.length; h++) {
      tmpSheet.getRange(4, h + 1).setValue(mainHeaders[h]);
      tmpSheet.getRange(4, h + 1, 2, 1).merge(); // merge row 4-5
    }
    // "บันทึกคะแนน" merged across G4:P4
    tmpSheet.getRange(4, 7).setValue('บันทึกคะแนน');
    tmpSheet.getRange(4, 7, 1, 10).merge();
    // Row 5: เลข 1-10
    for (var s = 1; s <= 10; s++) {
      tmpSheet.getRange(5, 6 + s).setValue(s);
    }

    // จัดรูปแบบหัวตาราง (Row 4-5)
    var headerRange = tmpSheet.getRange(4, 1, 2, totalCols);
    headerRange.setFontFamily(FONT).setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    tmpSheet.setRowHeight(4, 22);
    tmpSheet.setRowHeight(5, 20);

    // --- Row 6+: ข้อมูลนักเรียน ---
    var dataStartRow = 6;
    var rows = [];
    for (var i = 0; i < students.length; i++) {
      var st = students[i];
      var row = [st.no, st.id, st.idCard, st.name, st.age, st.birthday,
                 '', '', '', '', '', '', '', '', '', ''];
      rows.push(row);
    }
    if (rows.length > 0) {
      tmpSheet.getRange(dataStartRow, 1, rows.length, totalCols).setValues(rows);
    }

    // จัดรูปแบบข้อมูลนักเรียน
    var dataRange = tmpSheet.getRange(dataStartRow, 1, rows.length, totalCols);
    dataRange.setFontFamily(FONT).setFontSize(14)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

    // จัดกึ่งกลางคอลัมน์ที่ต้องการ
    var centerCols = [1, 2, 5]; // ที่, รหัส, อายุ
    for (var c = 0; c < centerCols.length; c++) {
      tmpSheet.getRange(dataStartRow, centerCols[c], rows.length, 1).setHorizontalAlignment('center');
    }
    // บัตร ปชช. - กึ่งกลาง, font เล็กลง
    tmpSheet.getRange(dataStartRow, 3, rows.length, 1)
      .setHorizontalAlignment('center').setFontSize(12);
    // ชื่อ-สกุล - ชิดซ้าย
    tmpSheet.getRange(dataStartRow, 4, rows.length, 1).setHorizontalAlignment('left');
    // วันเกิด - กึ่งกลาง, font เล็กลง
    tmpSheet.getRange(dataStartRow, 6, rows.length, 1)
      .setHorizontalAlignment('center').setFontSize(12);
    // ช่องคะแนน G-P กึ่งกลาง
    tmpSheet.getRange(dataStartRow, 7, rows.length, 10).setHorizontalAlignment('center');

    // ตั้งความสูงแถวข้อมูล
    for (var r = 0; r < rows.length; r++) {
      tmpSheet.setRowHeight(dataStartRow + r, 22);
    }

    // --- ตั้งความกว้างคอลัมน์ ---
    tmpSheet.setColumnWidth(1, 28);   // A: ที่
    tmpSheet.setColumnWidth(2, 38);   // B: รหัส
    tmpSheet.setColumnWidth(3, 82);   // C: บัตร ปชช.
    tmpSheet.setColumnWidth(4, 154);  // D: ชื่อ-สกุล
    tmpSheet.setColumnWidth(5, 48);   // E: อายุ
    tmpSheet.setColumnWidth(6, 68);   // F: วันเกิด
    for (var sc = 7; sc <= 16; sc++) {
      tmpSheet.setColumnWidth(sc, 27); // G-P: คะแนน 1-10
    }

    // --- Export PDF ---
    SpreadsheetApp.flush(); // บันทึกข้อมูลก่อน export
    var ssId = ss.getId();
    var gid = tmpSheet.getSheetId();
    var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export'
      + '?format=pdf'
      + '&gid=' + gid
      + '&portrait=true'
      + '&size=A4'
      + '&top_margin=0.20'
      + '&bottom_margin=0.20'
      + '&left_margin=0.59'
      + '&right_margin=0.20'
      + '&sheetnames=false'
      + '&printtitle=false'
      + '&gridlines=false'
      + '&fzr=false'
      + '&fitw=true'
      + '&fith=false';

    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(pdfUrl, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Export PDF ล้มเหลว: HTTP ' + response.getResponseCode());
    }

    var pdfBlob = response.getBlob().setName(
      createFileName(grade, classNo, gender, academicYear, semester)
    );

    // --- ลบ temp sheet ---
    ss.deleteSheet(tmpSheet);
    tmpSheet = null;

    // --- บันทึก PDF ---
    var result = savePDFAndGetUrl(pdfBlob, pdfBlob.getName(), settings);
    Logger.log('สร้าง PDF สำเร็จ: ' + pdfBlob.getName());
    return result;

  } catch (error) {
    // ลบ temp sheet ถ้าเกิด error
    if (tmpSheet) {
      try { SS().deleteSheet(tmpSheet); } catch (e) { /* ignore */ }
    }
    Logger.log('generateStudentListPDF error: ' + error.message);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

/**
 * ดึงข้อมูลนักเรียนตามเงื่อนไข
 */
/**
 * ดึงข้อมูลนักเรียนตามเงื่อนไข (ฉบับแก้ไข: เพิ่มการเรียงตามรหัสนักเรียน)
 */
function getFilteredStudents(grade, classNo, gender) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('Students') || ss.getSheetByName('นักเรียน') || ss.getSheets()[0];
    
    if (!sheet) {
      throw new Error('ไม่พบชีทข้อมูลนักเรียน');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // หาตำแหน่งคอลัมน์ (ปรับให้ตรงกับชีท Students)
    const gradeCol = U_findColumnIndex(headers, ['grade', 'ชั้น', 'ระดับชั้น']);
    const classCol = U_findColumnIndex(headers, ['class_no', 'ห้อง', 'ห้องเรียน']);
    const genderCol = U_findColumnIndex(headers, ['gender', 'เพศ']);
    const idCol = U_findColumnIndex(headers, ['student_id', 'รหัส', 'รหัสนักเรียน']);
    const idCardCol = U_findColumnIndex(headers, ['id_card', 'เลขบัตรประชาชน', 'บัตรปชช']);
    const birthCol = U_findColumnIndex(headers, ['birthdate', 'วันเกิด']);
    const ageCol = U_findColumnIndex(headers, ['age', 'อายุ']);
    const titleCol = U_findColumnIndex(headers, ['title', 'คำนำหน้า']);
    const firstnameCol = U_findColumnIndex(headers, ['firstname', 'ชื่อ']);
    const lastnameCol = U_findColumnIndex(headers, ['lastname', 'สกุล']);
    
    Logger.log(`พบคอลัมน์: ชั้น=${gradeCol}, ห้อง=${classCol}, เพศ=${genderCol}, รหัส=${idCol}, บัตรปชช=${idCardCol}`);
    
    const students = [];
    
    // กรองข้อมูลตามเงื่อนไข
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // ตรวจสอบเงื่อนไขการกรอง
      if (grade && gradeCol >= 0 && row[gradeCol] !== grade) continue;
      if (classNo && classCol >= 0 && row[classCol] != classNo) continue;
      if (gender && genderCol >= 0 && row[genderCol] !== gender) continue;
      
      // สร้างชื่อเต็มจาก title + firstname + lastname
      let fullName = '';
      if (titleCol >= 0) fullName += (row[titleCol] || '') + '';
      if (firstnameCol >= 0) fullName += (row[firstnameCol] || '') + ' ';
      if (lastnameCol >= 0) fullName += (row[lastnameCol] || '');
      fullName = fullName.trim();
      
      // เพิ่มข้อมูลนักเรียน
      const student = {
        no: students.length + 1,
        id: idCol >= 0 ? String(row[idCol] || '') : '',
        idCard: idCardCol >= 0 ? String(row[idCardCol] || '') : '',
        name: fullName || '',
        age: calcAge_(row, ageCol, birthCol),
        birthday: birthCol >= 0 ? formatDate(row[birthCol]) : '',
        grade: gradeCol >= 0 ? String(row[gradeCol] || '') : grade || '',
        class: classCol >= 0 ? String(row[classCol] || '') : classNo || '',
        gender: genderCol >= 0 ? String(row[genderCol] || '') : ''
      };
      
      students.push(student);
    }
    
    Logger.log(`พบนักเรียน ${students.length} คน`);
    
    // --- 👇 จุดที่เพิ่มเข้ามาเพื่อเรียงลำดับ ---
    students.sort((a, b) => {
      // ใช้ localeCompare เพื่อการเรียงที่ถูกต้องทั้งตัวเลขและตัวอักษรปนกัน
      return String(a.id).localeCompare(String(b.id), undefined, { numeric: true });
    });
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

    return students;
    
  } catch (error) {
    Logger.log('getFilteredStudents error:', error);
    throw new Error(`ไม่สามารถดึงข้อมูลนักเรียนได้: ${error.message}`);
  }
}

/**
 * คำนวณอายุจากวันเกิด (ถ้าคอลัมน์อายุว่าง)
 */
function calcAge_(row, ageCol, birthCol) {
  // ถ้ามีค่าอายุในชีตอยู่แล้ว ใช้เลย
  if (ageCol >= 0 && row[ageCol] !== '' && row[ageCol] !== null && row[ageCol] !== undefined) {
    return String(row[ageCol]);
  }
  // คำนวณจากวันเกิด
  if (birthCol >= 0 && row[birthCol]) {
    try {
      var bd = (row[birthCol] instanceof Date) ? row[birthCol] : new Date(row[birthCol]);
      if (!isNaN(bd.getTime())) {
        var today = new Date();
        var age = today.getFullYear() - bd.getFullYear();
        var m = today.getMonth() - bd.getMonth();
        if (m < 0 || (m === 0 && today.getDate() < bd.getDate())) age--;
        return age > 0 ? String(age) : '';
      }
    } catch (e) { /* ignore */ }
  }
  return '';
}

/**
 * แปลงโลโก้เป็น Base64 Data URL
 */
function getLogoAsBase64() {
  try {
    var settings = S_getGlobalSettings();
    var logoFileId = settings['logoFileId'];
    if (!logoFileId) { return ''; }
    var file = DriveApp.getFileById(logoFileId);
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + blob.getContentType() + ';base64,' + base64;
  } catch (error) {
    Logger.log('getLogoAsBase64 error: ' + error.message);
    return '';
  }
}

/**
 * หา Index ของคอลัมน์จากชื่อที่เป็นไปได้
 * @param {Array} headers - แถวหัวคอลัมน์
 * @param {Array} possibleNames - รายชื่อที่อาจใช้
 * @returns {number} Index ของคอลัมน์ หรือ -1 ถ้าไม่พบ
 */
function U_findColumnIndex(headers, possibleNames) {
  for (var i = 0; i < headers.length; i++) {
    var headerName = String(headers[i] || '').trim().toLowerCase();
    for (var j = 0; j < possibleNames.length; j++) {
      var name = String(possibleNames[j]).toLowerCase();
      if (headerName === name || headerName.indexOf(name) !== -1) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * จัดรูปแบบวันที่
 */
function formatDate(date) {
  if (!date) return '';
  
  try {
    if (typeof date === 'string') {
      return date;
    }
    
    if (date instanceof Date) {
      const thaiMonthsShort = [
        'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
        'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
      ];
      
      const day = date.getDate();
      const month = thaiMonthsShort[date.getMonth()];
      const yearBE = date.getFullYear() + 543;
      const yearShort = String(yearBE).slice(-2);
      
      return `${day} ${month} ${yearShort}`;
    }
    
    return date.toString();
  } catch (error) {
    return '';
  }
}

/**
 * สร้างชื่อไฟล์
 */
function createFileName(grade, classNo, gender, academicYear, semester) {
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
  
  let fileName = 'รายชื่อนักเรียน';
  
  if (grade) fileName += `_${grade}`;
  if (classNo) fileName += `_ห้อง${classNo}`;
  if (gender) fileName += `_${gender}`;
  
  fileName += `_${academicYear}_${timestamp}.pdf`;
  
  return fileName;
}

/**
 * บันทึก PDF และส่งกลับ URL
 */
function savePDFAndGetUrl(pdfBlob, fileName, settings) {
  try {
    let folder;
    
    if (settings['pdfSaveFolderId']) {
      try {
        folder = DriveApp.getFolderById(settings['pdfSaveFolderId']);
      } catch (error) {
        Logger.log('ไม่สามารถเข้าถึงโฟลเดอร์ที่กำหนด ใช้โฟลเดอร์ราก');
        folder = DriveApp.getRootFolder();
      }
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    const file = folder.createFile(pdfBlob.setName(fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    Logger.log(`บันทึกไฟล์: ${fileName} ใน ${folder.getName()}`);
    
    // ✅ แก้ตรงนี้ — เปลี่ยนจาก file.getUrl() เป็น URL ตรง
    return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    
  } catch (error) {
    Logger.log('savePDFAndGetUrl error:', error);
    throw new Error(`ไม่สามารถบันทึกไฟล์ PDF ได้: ${error.message}`);
  }
}

/**
 * ฟังก์ชันทดสอบและแก้ไขโลโก้
 */
function fixLogoIssue() {
  try {
    const settings = S_getGlobalSettings();
    const logoFileId = settings['logoFileId'];
    
    if (!logoFileId) {
      Logger.log('ไม่มี logoFileId ในการตั้งค่า');
      return 'ไม่มี logoFileId';
    }
    
    Logger.log('logoFileId:', logoFileId);
    
    // ตรวจสอบและแก้ไขไฟล์โลโก้
    const file = DriveApp.getFileById(logoFileId);
    Logger.log('พบไฟล์โลโก้:', file.getName());
    
    // ตั้งค่าการแชร์เป็น public
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('ตั้งค่าการแชร์เป็น public แล้ว');
    
    return 'แก้ไขการตั้งค่าโลโก้สำเร็จ';
    
  } catch (error) {
    Logger.log('fixLogoIssue error:', error);
    return `Error: ${error.message}`;
  }
}

/**
 * สร้างและส่งออก PDF รายงานผลการเรียนรายวิชา (เวอร์ชันแก้ไข)
 * - เปิดสเปรดชีตกำหนดจาก SPREADSHEET_ID เสมอ
 * - โหลดเทมเพลตโดยไม่ใส่ .html
 * - เซ็ต hasLogo / logoDataUrl ให้เทมเพลต
 * - รองรับ term: '1' | '2' | 'both'
 */
function generateSubjectScorePDF(subjectName, subjectCode, grade, classNo, term = 'both') {
  try {
    if (!subjectName || !grade || !classNo) {
      throw new Error('กรุณาระบุข้อมูลที่จำเป็น: ชื่อวิชา, ระดับชั้น, ห้อง');
    }

    const settings = S_getGlobalSettings() || {};
    let logoDataUrl = String(settings.logoDataUrl || '');
    if (!logoDataUrl && typeof getLogoAsBase64 === 'function') {
      logoDataUrl = getLogoAsBase64() || '';
    }
    const hasLogo = !!logoDataUrl;

    // ✏️ แก้ไขบรรทัดนี้ ให้เรียกใช้ฟังก์ชันเวอร์ชันที่ถูกต้อง
    const scoreData = getSubjectScoreTableFixed(subjectName, subjectCode, grade, classNo, term);

    const tpl = HtmlService.createTemplateFromFile('pdf_subject_score_template');

    // ✅ ส่งตัวแปรให้ครบ
    tpl.settings      = settings;
    tpl.schoolName    = settings.schoolName || 'โรงเรียนของเรา';
    tpl.hasLogo       = hasLogo;
    tpl.logoDataUrl   = logoDataUrl;
    tpl.data          = scoreData;
    tpl.subjectName   = subjectName;
    tpl.subjectCode   = subjectCode || '';
    tpl.grade         = grade;
    tpl.classNo       = classNo;
    tpl.term          = term || 'both';
    tpl.academicYear  = settings.academicYear || '';
    tpl.teacherName   = settings.teacherName  || '';
    tpl.remark        = settings.remark       || '';
    tpl.gradeFullName = convertGradeToFullName(grade);

    const html = tpl.evaluate().getContent();

    const safeSubj  = String(subjectName).replace(/[\\/:*?"<>|]/g, '_');
    const safeGrade = String(grade).replace(/\./g, '');
    const fileName  = `ผลการเรียน_${safeSubj}_${safeGrade}${classNo}_${tpl.term}.pdf`;

    const blob = Utilities.newBlob(html, 'text/html', 'tmp.html')
                           .getAs('application/pdf')
                           .setName(fileName);

    const folderId = settings.pdfSaveFolderId || DriveApp.getRootFolder().getId();
    const folder   = DriveApp.getFolderById(folderId);
    const file     = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return file.getUrl();
  } catch (error) {
    Logger.log('Error in generateSubjectScorePDF:', error);
    throw new Error(`Generate PDF ล้มเหลว: ${error && error.message ? error.message : error}`);
  }
}


// ============================================================
// 📄 PDF & EXPORT GENERATION
// ============================================================

/**
 * ✅ [เวอร์ชันสำหรับโปรเจกต์ใหม่] ส่งออก PDF ตารางเวลาเรียน (ฉบับแก้ไข: เพิ่มการเรียงลำดับ)
 */
function exportAttendancePDF(grade, classNo, year, month) {
  try {
    const monthNames = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
    
    // 1. ดึงข้อมูลนักเรียน
    const data = getAttendanceVerticalTable(grade, classNo, year, month);
    if (!data || !data.students || data.students.length === 0) {
      throw new Error(`ไม่พบข้อมูลนักเรียน`);
    }

    // --- 👇 เพิ่มโค้ดเรียงลำดับตามรหัสประจำตัวตรงนี้ ---
    data.students.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---
    
    // 3. สร้าง HTML และส่วนที่เหลือของฟังก์ชัน (เหมือนเดิมทั้งหมด)
    const settings = S_getGlobalSettings();
    const gradeFullName = convertGradeToFullName(grade);
    
    const htmlContent = buildAttendancePdfHtml(
      `ตารางเวลาเรียน ชั้น ${gradeFullName} ห้อง ${classNo} เดือน ${monthNames[month]} พ.ศ. ${year + 543}`,
      data, settings
    );

    const blob = Utilities.newBlob(htmlContent, 'text/html').getAs('application/pdf');
    const fileName = `เวลาเรียน_${grade}_${classNo}_${monthNames[month]}${year+543}.pdf`;
    blob.setName(fileName);
    
    const folderId = settings.pdfSaveFolderId || DriveApp.getRootFolder().getId();
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();

  } catch(e) {
    Logger.log("เกิดข้อผิดพลาดใน exportAttendancePDF: " + e.message);
    throw new Error(e.message);
  }
}


/**
 * ⚙️ [ฟังก์ชันเสริม] สร้างโค้ด HTML สำหรับ PDF
 */
function buildAttendancePdfHtml(title, data, settings) {
  const students = data.students;
  const days = data.days;

  let tableRows = '';
  students.forEach((s, i) => {
    let dayCells = '';
    days.forEach(d => {
      dayCells += `<td>${s.statusMap[d.label] || ''}</td>`;
    });
    tableRows += `<tr><td>${i + 1}</td><td class="name">${s.name}</td>${dayCells}</tr>`;
  });

  let tableHeader = `<tr><th rowspan="2">ที่</th><th rowspan="2" class="name">ชื่อ - สกุล</th>
    ${days.map(d => `<th>${d.label}</th>`).join('')}
    </tr><tr>
    ${days.map(d => `<th>${d.dow}</th>`).join('')}
    </tr>`;

  return `
    <!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
    <style>
      body { font-family: 'Sarabun', sans-serif; font-size: 10pt; }
      .header { text-align: center; margin-bottom: 15px; }
      .logo { max-height: 80px; }
      h3 { font-size: 14pt; margin: 5px 0; }
      h4 { font-size: 14pt; margin: 5px 0; font-weight: bold; }
      table { width: 100%; border-collapse: collapse; }
      th, td { border: 1px solid #333; padding: 2px; text-align: center; }
      th { background-color: #f2f2f2; }
      td.name { text-align: left; }
    </style></head><body>
      <div class="header">
        ${settings.logoDataUrl ? `<img src="${settings.logoDataUrl}" class="logo">` : ''}
        <h3>${settings.schoolName || ''}</h3>
        <h4>${title}</h4>
      </div>
      <table><thead>${tableHeader}</thead><tbody>${tableRows}</tbody></table>
    </body></html>`;
}

/**สร้าง pdf ข้อมูลนักเรียน........................................................................................... */

/**
 * สร้างไฟล์ PDF ข้อมูลผู้ปกครองของนักเรียนตามชั้นเรียนและห้องที่เลือก
 * (ฉบับสมบูรณ์ล่าสุด: ใช้ Folder ID จากการตั้งค่ากลาง, ปรับปรุง Layout)
 * @param {string} grade ระดับชั้น เช่น "ป.1"
 * @param {string} classNo หมายเลขห้อง เช่น "1"
 * @returns {string} URL ของไฟล์ PDF ที่สร้างขึ้น
 */
function generateParentInfoPDF(grade, classNo) {
  if (!grade || !classNo) throw new Error("กรุณาระบุชั้นและห้องให้ครบถ้วน");

  const settings = S_getGlobalSettings(false);
  const ss = (getSpreadsheetId_())
    ? SS()
    : SS();
  const studentSheet = ss.getSheetByName("Students") || ss.getSheetByName("นักเรียน");
  if (!studentSheet) throw new Error("ไม่พบชีต 'Students/นักเรียน'");
  
  const studentData = studentSheet.getDataRange().getValues();
  const header = studentData[0];
  // --- ส่วนของการหา index ของคอลัมน์ (เหมือนเดิม) ---
  const idCol = header.indexOf("student_id"); // *** เพิ่ม Index รหัสนักเรียน ***
  const titleCol = header.indexOf("title"),
        firstNameCol = header.indexOf("firstname"),
        lastNameCol = header.indexOf("lastname"),
        gradeCol = header.indexOf("grade"),
        classCol = header.indexOf("class_no"),
        fatherNameCol = header.indexOf("father_name"),
        fatherLNameCol = header.indexOf("father_lastname"),
        fatherOccCol = header.indexOf("father_occupation"),
        motherNameCol = header.indexOf("mother_name"),
        motherLNameCol = header.indexOf("mother_lastname"),
        motherOccCol = header.indexOf("mother_occupation"),
        addressCol = header.indexOf("address");

  const missing = [];
  if (gradeCol === -1) missing.push('grade');
  if (classCol === -1) missing.push('class_no');
  if (titleCol === -1) missing.push('title');
  if (firstNameCol === -1) missing.push('firstname');
  if (lastNameCol === -1) missing.push('lastname');
  if (addressCol === -1) missing.push('address');
  if (missing.length) {
    throw new Error(`หัวตารางชีต Students ไม่ถูกต้อง (missing: ${missing.join(', ')})`);
  }

  const filteredStudents = studentData.slice(1).filter(row => row[gradeCol] === grade && String(row[classCol]) === String(classNo));
  if (filteredStudents.length === 0) throw new Error(`ไม่พบข้อมูลนักเรียนในชั้น ${grade} ห้อง ${classNo}`);

  // --- 👇 เพิ่มโค้ดเรียงลำดับนักเรียนตามรหัส (student_id) ตรงนี้ ---
  if (idCol !== -1) { // ตรวจสอบว่าพบคอลัมน์รหัสนักเรียนหรือไม่
    filteredStudents.sort((a, b) => {
        const idA = a[idCol] || '';
        const idB = b[idCol] || '';
        return String(idA).localeCompare(String(idB), undefined, { numeric: true });
    });
  }
  // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

  const tableRows = filteredStudents.map((row, index) => {
    const studentName = `${row[titleCol] || ""} ${row[firstNameCol] || ""} ${row[lastNameCol] || ""}`.trim();
    const fatherFullName = `${row[fatherNameCol] || ""} ${row[fatherLNameCol] || ""}`.trim();
    const fatherOccupation = row[fatherOccCol] || "";
    const motherFullName = `${row[motherNameCol] || ""} ${row[motherLNameCol] || ""}`.trim();
    const motherOccupation = row[motherOccCol] || "";
    const addressString = row[addressCol] || "";
    const parsedAddr = parseAddress(addressString);
    return `<tr><td>${index + 1}</td><td class="left">${studentName}</td><td class="left">${fatherFullName}</td><td>${fatherOccupation}</td><td class="left">${motherFullName}</td><td>${motherOccupation}</td><td>${parsedAddr.house}</td><td>${parsedAddr.moo}</td><td>${parsedAddr.subdistrict}</td><td>${parsedAddr.district}</td><td>${parsedAddr.province}</td></tr>`;
  });
  
  const gradeFullName = convertGradeToFullName(grade);
  const schoolName = String(settings['ชื่อโรงเรียน'] || settings.schoolName || '').trim() || 'โรงเรียน';
  const academicYear = String(settings['ปีการศึกษา'] || settings.academicYear || '').trim();
  const title = `ข้อมูลผู้ปกครองและที่อยู่ปัจจุบัน ชั้น${gradeFullName} ห้อง ${classNo}`;
  const subTitle = academicYear ? `${schoolName} | ปีการศึกษา ${academicYear}` : schoolName;
  let html = `<html><head><meta charset="UTF-8"><style>body{font-family:'Sarabun',sans-serif}.header{text-align:center;font-weight:700;font-size:14pt;margin-bottom:4px}.sub-header{text-align:center;font-size:10pt;margin-bottom:12px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #000;padding:2px;text-align:center;vertical-align:middle;word-wrap:break-word}th{font-weight:700;background-color:#f2f2f2;font-size:10pt}td{font-size:7.5pt}td.left{text-align:left}.col-no{width:3%}.col-student-name{width:18%}.col-parent-name{width:16%}.col-job{width:5%}.col-addr-house{width:8%}.col-addr-moo{width:4%}.col-addr-sub{width:8%}.col-addr-dist{width:8%}.col-addr-prov{width:10%}</style></head><body><div class="header">${title}</div><div class="sub-header">${subTitle}</div><table><thead><tr><th rowspan="2" class="col-no">เลขที่</th><th rowspan="2" class="col-student-name">ชื่อ - สกุล นักเรียน</th><th colspan="2">บิดา</th><th colspan="2">มารดา</th><th colspan="5">ที่อยู่ปัจจุบัน</th></tr><tr><th class="col-parent-name">ชื่อ-สกุล บิดา</th><th class="col-job">อาชีพ</th><th class="col-parent-name">ชื่อ-สกุล มารดา</th><th class="col-job">อาชีพ</th><th class="col-addr-house">บ้านเลขที่</th><th class="col-addr-moo">หมู่ที่</th><th class="col-addr-sub">ตำบล</th><th class="col-addr-dist">อำเภอ</th><th class="col-addr-prov">จังหวัด</th></tr></thead><tbody>${tableRows.join("")}</tbody></table></body></html>`;
  
  const blob = Utilities.newBlob(html, 'text/html').getAs('application/pdf');
  const fileName = `ข้อมูลผู้ปกครอง_${grade.replace(/\./g, "")}_ห้อง${classNo}.pdf`;
  blob.setName(fileName);

  const folderId = settings.pdfSaveFolderId;
  if (!folderId) throw new Error("ยังไม่ได้ตั้งค่า 'pdfSaveFolderId' ในชีต global_settings");
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error("ไม่พบโฟลเดอร์ Google Drive หรือคุณไม่มีสิทธิ์เข้าถึง กรุณาตรวจสอบ Folder ID: " + folderId);
  }
  
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


/**
 * ฟังก์ชันเสริมสำหรับแยกส่วนประกอบของที่อยู่จากข้อความเดียว (ฉบับสมบูรณ์)
 * - รองรับรูปแบบที่ใช้ลำดับ (เช่น "3 2 - ก้านเหลือง นางรอง บุรีรัมย์")
 * - โดยจะดึงตัวเลขกลุ่มแรกเป็น "บ้านเลขที่" และกลุ่มที่สองเป็น "หมู่ที่"
 * - และยังคงรองรับรูปแบบที่มีคำนำหน้า (เช่น "123 ต.ในเมือง อ.เมือง จ.บุรีรัมย์")
 * @param {string} addressString ที่อยู่แบบเต็ม
 * @returns {object} อ็อบเจกต์ที่ประกอบด้วย house, moo, subdistrict, district, province
 */
function parseAddress(addressString) {
  // 1. ตรวจสอบข้อมูลเบื้องต้น
  if (!addressString || typeof addressString !== 'string') {
    return { house: '', moo: '', subdistrict: '', district: '', province: '' };
  }

  // 2. ทำความสะอาดข้อมูล: ตัดช่องว่างหน้า-หลัง และรวมช่องว่างหลายๆอันเป็นอันเดียว
  let cleanedString = addressString.trim().replace(/\s+/g, ' ');

  // 3. แบ่งข้อความออกเป็นส่วนๆ ด้วยช่องว่าง
  const parts = cleanedString.split(' ');

  // 4. ถ้าข้อมูลสั้นเกินไป ไม่สามารถแยกได้ ให้คืนค่าทั้งหมดไปที่ บ้านเลขที่
  if (parts.length < 3) {
    return { house: cleanedString, moo: '', subdistrict: '', district: '', province: '' };
  }

  // 5. ดึง จังหวัด, อำเภอ, ตำบล จากท้ายของ Array (แน่นอนที่สุด)
  const province = parts.pop().replace('จ.', '').trim();
  const district = parts.pop().replace('อ.', '').trim();
  const subdistrict = parts.pop().replace('ต.', '').trim();
  
  // 6. ส่วนที่เหลือคือข้อมูลบ้านเลขที่และหมู่บ้าน
  const remainder = parts.join(' ').trim();
  
  let house = '';
  let moo = '';

  // 7. ใช้ Regular Expression เพื่อดึงกลุ่มตัวเลขทั้งหมดในส่วนที่เหลือ
  const numbers = remainder.match(/\d+/g);
  
  if (numbers && numbers.length > 0) {
    // ถ้าพบกลุ่มตัวเลข
    house = numbers[0] || ''; // ตัวเลขกลุ่มแรกคือ บ้านเลขที่
    moo = numbers[1] || '';   // ตัวเลขกลุ่มที่สองคือ หมู่ที่
  } else {
    // ถ้าไม่พบตัวเลขเลย ให้ใช้ข้อความที่เหลือเป็นบ้านเลขที่
    house = remainder;
  }
  
  // 8. สร้างผลลัพธ์ที่จะส่งกลับ
  const result = {
    house: house,
    moo: moo,
    subdistrict: subdistrict,
    district: district,
    province: province
  };

  // สำหรับ Debug: แสดงผลลัพธ์ใน Log เพื่อตรวจสอบ
  Logger.log('Parsed Address: ' + JSON.stringify(result));
  
  return result;
}


/**
 * =================================================================
 * ส่วนสำหรับรายงานผลนักเรียนรายบุคคล (ปพ.6)
 * =================================================================
 */
/**
 * ✅ รวบรวมข้อมูลทั้งหมดสำหรับรายงาน ปพ.6 (รักษารูปแบบเดิม)
 */
/**
 * ✅ ปรับปรุงแล้ว - รองรับการเลือกภาคเรียน
 */
function getPp6ReportData(studentId, term = 'both') {
  // เรียกใช้ฟังก์ชันใหม่จาก individual-reports.gs
  return getPp6ReportDataComplete(studentId, term);
}

/**
 * ✅ ดึงผลการเรียนรายวิชา (รักษารูปแบบเดิม)
 */
function getStudentAcademicResults(studentId) {
  const ss = SS();
  const sheets = ss.getSheets();
  const subjects = [];
  let totalGradePoints = 0;
  let totalCredits = 0;
  
  // หาชีตที่เป็นรายวิชา
  const scoreSheets = sheets.filter(sheet => {
    const name = sheet.getName();
    return !['Students', 'รายวิชา', 'global_settings', 'คุณลักษณะ', 'อ่านคิดเขียน', 'กิจกรรม'].includes(name) &&
           !name.includes('2568') && !name.includes('2567') && !name.includes('2569');
  });
  
  scoreSheets.forEach(sheet => {
    try {
      const data = sheet.getDataRange().getValues();
      if (data.length < 5) return;
      
      // หาข้อมูลนักเรียน
      for (let i = 4; i < data.length; i++) {
        if (String(data[i][1]) === String(studentId)) {
          const subjectName = data[1][1] || sheet.getName();
          const term1Score = Number(data[i][13]) || 0; // คะแนนระหว่างภาค 1
          const term1Final = Number(data[i][14]) || 0; // คะแนนปลายภาค 1
          const term1Total = Number(data[i][15]) || 0; // รวมภาค 1
          
          const term2Score = Number(data[i][27]) || 0; // คะแนนระหว่างภาค 2
          const term2Final = Number(data[i][28]) || 0; // คะแนนปลายภาค 2
          const term2Total = Number(data[i][29]) || 0; // รวมภาค 2
          
          const averageScore = Number(data[i][31]) || 0; // คะแนนเฉลี่ย
          const grade = data[i][32] || '0'; // เกรด
          
          subjects.push({
            name: subjectName,
            term1: {
              score: term1Score,
              final: term1Final,
              total: term1Total
            },
            term2: {
              score: term2Score,
              final: term2Final,
              total: term2Total
            },
            credits: 1,
            score: averageScore,
            grade: grade
          });
          
          // คำนวณ GPA
          if (!isNaN(Number(grade))) {
            totalGradePoints += Number(grade);
            totalCredits += 1;
          }
          break;
        }
      }
    } catch (e) {
      Logger.log(`เกิดข้อผิดพลาดในชีต ${sheet.getName()}: ${e.message}`);
    }
  });
  
  const gpa = totalCredits > 0 ? (totalGradePoints / totalCredits).toFixed(2) : '0.00';
  
  return {
    subjects: subjects,
    gpa: gpa
  };
}

/**
 * ✅ สรุปการเข้าเรียนทั้งปี
 */
function getStudentAttendanceSummary(studentId) {
  const ss = SS();
  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const year = new Date().getFullYear() + 543;
  
  let present = 0, late = 0, leave = 0, absent = 0;
  
  months.forEach(month => {
    const sheetName = `${month}${year}`;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const studentRow = data.find(row => String(row[1]) === String(studentId));
    if (!studentRow) return;
    
    // นับสถานะจากคอลัมน์วันที่ (3-33)
    for (let col = 3; col <= 33; col++) {
      const status = studentRow[col];
      if (status === '/' || status === '✔') present++;
      else if (status === 'ส') late++;
      else if (status === 'ล') leave++;
      else if (status === 'ข') absent++;
    }
  });
  
  const totalDays = present + late + leave + absent;
  
  return {
    present: present,
    late: late,
    leave: leave,
    absent: absent,
    totalDays: totalDays
  };
}
// =================================================================
// 🆕 ฟังก์ชันสำหรับแก้ไขข้อมูลผู้ปกครอง (เพิ่มใหม่)
// =================================================================

/**
 * ✅ ค้นหานักเรียนสำหรับแก้ไข (ฉบับปรับปรุง: คำนวณเลขที่อัตโนมัติ)
 */
function searchStudentsForEdit(grade, classNo, name) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) throw new Error('ไม่พบชีต Students');

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงหัวคอลัมน์ออกมาก่อน

    const colMap = {};
    headers.forEach((header, index) => {
      colMap[header.trim()] = index;
    });

    // --- ส่วนคำนวณเลขที่ (ใหม่) ---
    // 1. จัดกลุ่มนักเรียนทั้งหมดตามห้องเรียน
    const studentsByClass = {};
    data.forEach(row => {
      const classKey = `${row[colMap['grade']]}/${row[colMap['class_no']]}`;
      if (!studentsByClass[classKey]) {
        studentsByClass[classKey] = [];
      }
      studentsByClass[classKey].push(row);
    });

    // 2. สร้าง Map สำหรับเก็บเลขที่ของนักเรียนทุกคน
    const studentNumberMap = {};
    for (const classKey in studentsByClass) {
      // 3. เรียงลำดับนักเรียนในแต่ละห้องตามรหัสประจำตัว
      const sortedStudents = studentsByClass[classKey].sort((a, b) => {
        return a[colMap['student_id']] - b[colMap['student_id']];
      });
      
      // 4. กำหนดเลขที่ (index + 1) ให้กับนักเรียนแต่ละคน
      sortedStudents.forEach((studentRow, index) => {
        const studentId = studentRow[colMap['student_id']];
        studentNumberMap[studentId] = index + 1;
      });
    }
    // --- สิ้นสุดส่วนคำนวณเลขที่ ---

    // 5. กรองนักเรียนตามเงื่อนไขการค้นหา
    const filteredStudents = data.filter(row => {
      const gradeMatch = !grade || row[colMap['grade']] == grade;
      const classMatch = !classNo || row[colMap['class_no']] == classNo;
      const fullName = `${row[colMap['firstname']]} ${row[colMap['lastname']]}`;
      const nameMatch = !name || fullName.includes(name);
      return gradeMatch && classMatch && nameMatch;
    });
    
    // 6. สร้างผลลัพธ์เพื่อส่งกลับไปหน้าเว็บ โดยเพิ่ม 'studentNo' ที่คำนวณไว้
    return filteredStudents.map(row => {
      const studentId = row[colMap['student_id']];
      return {
        studentId: studentId,
        fullName: `${row[colMap['title']] || ''} ${row[colMap['firstname']]} ${row[colMap['lastname']]}`.trim(),
        grade: row[colMap['grade']],
        classNo: row[colMap['class_no']],
        studentNo: studentNumberMap[studentId] || 'N/A' // ✨ ใช้เลขที่จาก Map ที่เราสร้างไว้
      };
    });

  } catch (error) {
    Logger.log('Error in searchStudentsForEdit: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการค้นหา: ' + error.message);
  }
}

/**
 * ✅ ดึงข้อมูลนักเรียนและผู้ปกครองสำหรับแก้ไข (ฉบับปรับปรุง: คำนวณเลขที่อัตโนมัติ)
 */
function getStudentParentData(studentId) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) throw new Error('ไม่พบชีต Students');

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงหัวคอลัมน์ออกมาก่อน

    // สร้าง map ของคอลัมน์เพื่อความสะดวก
    const colMap = {};
    headers.forEach((header, index) => {
      colMap[header.trim()] = index;
    });

    // --- ขั้นตอนที่ 1: ค้นหานักเรียนเป้าหมาย และดูว่าอยู่ชั้น/ห้องอะไร ---
    let targetRow = null;
    let targetGrade = null;
    let targetClassNo = null;

    for (const row of data) {
      if (row[colMap['student_id']] == studentId) {
        targetRow = row;
        targetGrade = row[colMap['grade']];
        targetClassNo = row[colMap['class_no']];
        break; // เมื่อเจอแล้วก็หยุดค้นหา
      }
    }

    // ถ้านักเรียนไม่มีอยู่จริง ให้แจ้งข้อผิดพลาด
    if (!targetRow) {
      throw new Error('ไม่พบรหัสนักเรียน ' + studentId);
    }

    // --- ขั้นตอนที่ 2: กรองเอานักเรียนทุกคนที่อยู่ห้องเดียวกัน ---
    const studentsInClass = data.filter(row => 
      row[colMap['grade']] == targetGrade && row[colMap['class_no']] == targetClassNo
    );

    // --- ขั้นตอนที่ 3: เรียงลำดับนักเรียนในห้องตามรหัสประจำตัว (น้อยไปมาก) ---
    studentsInClass.sort((a, b) => {
      return a[colMap['student_id']] - b[colMap['student_id']];
    });

    // --- ขั้นตอนที่ 4: หาว่านักเรียนของเราอยู่ลำดับที่เท่าไหร่ (นั่นคือเลขที่) ---
    const studentIndex = studentsInClass.findIndex(row => row[colMap['student_id']] == studentId);
    const calculatedStudentNo = studentIndex + 1; // Index เริ่มจาก 0 ดังนั้นต้อง +1

    // --- ส่งข้อมูลทั้งหมดกลับไปหน้าเว็บ ---
    return {
      studentId: targetRow[colMap['student_id']],
      fullName: `${targetRow[colMap['title']] || ''} ${targetRow[colMap['firstname']]} ${targetRow[colMap['lastname']]}`.trim(),
      grade: targetRow[colMap['grade']],
      classNo: targetRow[colMap['class_no']],
      studentNo: calculatedStudentNo, // ✨ ใช้เลขที่ที่คำนวณได้ใหม่
      
      // ข้อมูลบิดา
      fatherName: targetRow[colMap['father_name']] || '',
      fatherOccupation: targetRow[colMap['father_occupation']] || '',
      fatherPhone: targetRow[colMap['father_phone']] || '',
      fatherStatus: targetRow[colMap['father_status']] || 'มีชีวิต',
      
      // ข้อมูลมารดา
      motherName: targetRow[colMap['mother_name']] || '',
      motherOccupation: targetRow[colMap['mother_occupation']] || '',
      motherPhone: targetRow[colMap['mother_phone']] || '',
      motherStatus: targetRow[colMap['mother_status']] || 'มีชีวิต',
      
      // ข้อมูลผู้ปกครอง
      guardianName: targetRow[colMap['guardian_name']] || '',
      guardianRelation: targetRow[colMap['guardian_relation']] || '',
      guardianOccupation: targetRow[colMap['guardian_occupation']] || '',
      guardianPhone: targetRow[colMap['guardian_phone']] || '',
      
      // ที่อยู่
      address: targetRow[colMap['address']] || '',
      subdistrict: targetRow[colMap['subdistrict']] || '',
      district: targetRow[colMap['district']] || '',
      province: targetRow[colMap['province']] || '',
      postalCode: targetRow[colMap['postal_code']] || ''
    };
    
  } catch (error) {
    Logger.log('Error in getStudentParentData: ' + error.message);
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูล: ' + error.message);
  }
}

/**
 * ✅ บันทึกข้อมูลผู้ปกครองที่แก้ไข
 */
function saveParentData(formData) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) throw new Error('ไม่พบชีต Students');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // สร้าง map ของคอลัมน์
    const colMap = {};
    headers.forEach((header, index) => {
      colMap[header] = index;
    });

    // หาแถวของนักเรียน
    for (let i = 1; i < data.length; i++) {
      if (data[i][colMap['student_id']] == formData.studentId) {
        const row = i + 1; // แปลงเป็นแถว 1-based
        
        // อัปเดตข้อมูลบิดา
        if (colMap['father_name'] !== undefined) 
          sheet.getRange(row, colMap['father_name'] + 1).setValue(formData.fatherName || '');
        if (colMap['father_occupation'] !== undefined) 
          sheet.getRange(row, colMap['father_occupation'] + 1).setValue(formData.fatherOccupation || '');
        if (colMap['father_phone'] !== undefined) 
          sheet.getRange(row, colMap['father_phone'] + 1).setValue(formData.fatherPhone || '');
        if (colMap['father_status'] !== undefined) 
          sheet.getRange(row, colMap['father_status'] + 1).setValue(formData.fatherStatus || 'มีชีวิต');
        
        // อัปเดตข้อมูลมารดา
        if (colMap['mother_name'] !== undefined) 
          sheet.getRange(row, colMap['mother_name'] + 1).setValue(formData.motherName || '');
        if (colMap['mother_occupation'] !== undefined) 
          sheet.getRange(row, colMap['mother_occupation'] + 1).setValue(formData.motherOccupation || '');
        if (colMap['mother_phone'] !== undefined) 
          sheet.getRange(row, colMap['mother_phone'] + 1).setValue(formData.motherPhone || '');
        if (colMap['mother_status'] !== undefined) 
          sheet.getRange(row, colMap['mother_status'] + 1).setValue(formData.motherStatus || 'มีชีวิต');
        
        // อัปเดตข้อมูลผู้ปกครอง
        if (colMap['guardian_name'] !== undefined) 
          sheet.getRange(row, colMap['guardian_name'] + 1).setValue(formData.guardianName || '');
        if (colMap['guardian_relation'] !== undefined) 
          sheet.getRange(row, colMap['guardian_relation'] + 1).setValue(formData.guardianRelation || '');
        if (colMap['guardian_occupation'] !== undefined) 
          sheet.getRange(row, colMap['guardian_occupation'] + 1).setValue(formData.guardianOccupation || '');
        if (colMap['guardian_phone'] !== undefined) 
          sheet.getRange(row, colMap['guardian_phone'] + 1).setValue(formData.guardianPhone || '');
        
        // อัปเดตที่อยู่
        if (colMap['address'] !== undefined) 
          sheet.getRange(row, colMap['address'] + 1).setValue(formData.address || '');
        if (colMap['subdistrict'] !== undefined) 
          sheet.getRange(row, colMap['subdistrict'] + 1).setValue(formData.subdistrict || '');
        if (colMap['district'] !== undefined) 
          sheet.getRange(row, colMap['district'] + 1).setValue(formData.district || '');
        if (colMap['province'] !== undefined) 
          sheet.getRange(row, colMap['province'] + 1).setValue(formData.province || '');
        if (colMap['postal_code'] !== undefined) 
          sheet.getRange(row, colMap['postal_code'] + 1).setValue(formData.postalCode || '');
        
        // บันทึก log
        const username = Session.getActiveUser().getEmail();
        Logger.log(`[${new Date().toISOString()}] User: ${username} updated parent info for studentId: ${formData.studentId}`);
        
        return { success: true, message: 'บันทึกข้อมูลสำเร็จ' };
      }
    }
    
    throw new Error('ไม่พบข้อมูลนักเรียนที่ต้องการแก้ไข');
    
  } catch (error) {
    Logger.log('Error in saveParentData: ' + error);
    throw new Error('เกิดข้อผิดพลาดในการบันทึก: ' + error.message);
  }
}

/**
 * สร้างและส่งออก Excel รายชื่อนักเรียน
 */
function generateStudentListExcel(grade, classNo, gender) {
  try {
    Logger.log(`สร้าง Excel: ชั้น=${grade}, ห้อง=${classNo}, เพศ=${gender}`);

    const settings = S_getGlobalSettings();
    const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียนตัวอย่าง';
    const academicYear = settings['ปีการศึกษา'] || '2567';
    const semester = settings['ภาคเรียน'] || '1';

    // ดึงข้อมูลนักเรียน (ใช้ฟังก์ชันเดิมได้เลย)
    const students = getFilteredStudents(grade, classNo, gender);

    if (students.length === 0) {
      throw new Error('ไม่พบข้อมูลนักเรียนที่ตรงตามเงื่อนไข');
    }

    // แปลงชื่อชั้นเรียน
    const gradeMap = {
      'อ.1': 'อนุบาลปีที่ 1', 'อ.2': 'อนุบาลปีที่ 2', 'อ.3': 'อนุบาลปีที่ 3',
      'ป.1': 'ประถมศึกษาปีที่ 1', 'ป.2': 'ประถมศึกษาปีที่ 2', 'ป.3': 'ประถมศึกษาปีที่ 3',
      'ป.4': 'ประถมศึกษาปีที่ 4', 'ป.5': 'ประถมศึกษาปีที่ 5', 'ป.6': 'ประถมศึกษาปีที่ 6'
    };
    const gradeText = (grade && gradeMap[grade]) ? gradeMap[grade] : (grade || 'ทุกระดับ');
    const classText = classNo ? `ห้อง ${classNo}` : 'ทุกห้อง';
    const genderText = gender ? ` (${gender})` : '';
    const sheetTitle = `รายชื่อนักเรียน ชั้น ${gradeText} ${classText}${genderText}`;

    // สร้าง Spreadsheet ชั่วคราว
    const tempSS = SpreadsheetApp.create(`temp_excel_${Date.now()}`);
    const sheet = tempSS.getActiveSheet();
    sheet.setName('รายชื่อนักเรียน');

    // === หัวตาราง ===
    // แถว 1: ชื่อโรงเรียน
    sheet.getRange(1, 1).setValue(schoolName);
    sheet.getRange(1, 1, 1, 16).merge();
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

    // แถว 2: หัวเรื่อง
    sheet.getRange(2, 1).setValue(sheetTitle);
    sheet.getRange(2, 1, 1, 16).merge();
    sheet.getRange(2, 1).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');

    // แถว 3: หัวคอลัมน์แถวที่ 1
    const headers1 = ['ที่', 'รหัส', 'บัตร ปชช.', 'ชื่อ - สกุล', 'อายุ(ปี)', 'วันเกิด',
                       'บันทึกคะแนน', '', '', '', '', '', '', '', '', ''];
    sheet.getRange(3, 1, 1, 16).setValues([headers1]);

    // Merge ช่อง "บันทึกคะแนน"
    sheet.getRange(3, 7, 1, 10).merge();

    // Merge rowspan สำหรับคอลัมน์แรก 6 คอลัมน์ (แถว 3-4)
    for (let c = 1; c <= 6; c++) {
      sheet.getRange(3, c, 2, 1).merge();
    }

    // แถว 4: หัวคอลัมน์คะแนน 1-10
    const scoreHeaders = ['', '', '', '', '', '', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
    sheet.getRange(4, 1, 1, 16).setValues([scoreHeaders]);

    // จัดรูปแบบหัวตาราง
    sheet.getRange(3, 1, 2, 16)
      .setBackground('#f8f8f8')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);

    // === ข้อมูลนักเรียน ===
    const startRow = 5;
    students.forEach((student, index) => {
      const rowData = [
        student.no,
        student.id,
        student.idCard,
        student.name,
        student.age,
        student.birthday,
        '', '', '', '', '', '', '', '', '', '' // ช่องคะแนน 1-10
      ];
      const row = sheet.getRange(startRow + index, 1, 1, 16);
      row.setValues([rowData]);
      row.setBorder(true, true, true, true, true, true);
      row.setVerticalAlignment('middle');

      // จัดตำแหน่ง
      sheet.getRange(startRow + index, 1, 1, 6).setHorizontalAlignment('center');
      sheet.getRange(startRow + index, 4).setHorizontalAlignment('left'); // ชื่อชิดซ้าย
    });

    // === ปรับความกว้างคอลัมน์ ===
    sheet.setColumnWidth(1, 40);   // ที่
    sheet.setColumnWidth(2, 70);   // รหัส
    sheet.setColumnWidth(3, 110);  // บัตร ปชช.
    sheet.setColumnWidth(4, 180);  // ชื่อ-สกุล
    sheet.setColumnWidth(5, 55);   // อายุ
    sheet.setColumnWidth(6, 120);  // วันเกิด
    for (let c = 7; c <= 16; c++) {
      sheet.setColumnWidth(c, 35); // คะแนน 1-10
    }

// === ส่งออกเป็น .xlsx ===

// สร้างชื่อไฟล์
const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
let fileName = 'รายชื่อนักเรียน';
if (grade)   fileName += `_${grade}`;
if (classNo) fileName += `_ห้อง${classNo}`;
if (gender)  fileName += `_${gender}`;
fileName += `_${academicYear}_${timestamp}.xlsx`;

const tempFileId = tempSS.getId();
SpreadsheetApp.flush();

// ✅ วิธีที่ถูกต้อง — ใช้ Drive Export URL แทน .getAs()
const exportUrl = `https://docs.google.com/spreadsheets/d/${tempFileId}/export?format=xlsx`;
const response = UrlFetchApp.fetch(exportUrl, {
  headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
});

const xlsxBlob = response.getBlob();
xlsxBlob.setName(fileName);

// ✅ เพิ่มบรรทัดนี้ — บังคับ MIME type เป็น xlsx จริงๆ
xlsxBlob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

// บันทึกไฟล์
let folder;
if (settings['pdfSaveFolderId']) {
  try {
    folder = DriveApp.getFolderById(settings['pdfSaveFolderId']);
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }
} else {
  folder = DriveApp.getRootFolder();
}

const file = folder.createFile(xlsxBlob);
file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
DriveApp.getFileById(tempFileId).setTrashed(true);

Logger.log(`สร้าง Excel สำเร็จ: ${fileName}`);

// ✅ URL ดาวน์โหลดตรง
return `https://drive.google.com/uc?export=download&id=${file.getId()}`;

  } catch (error) {
    Logger.log('generateStudentListExcel error:', error);
    throw new Error(`ไม่สามารถสร้าง Excel ได้: ${error.message}`);
  }
}

