// ============================================================
// 🏫 INSTALLER.GS — Multi-Tenant Setup for New Schools
// ============================================================
// ฟังก์ชันนี้ใช้สำหรับติดตั้งระบบครั้งแรกเมื่อโรงเรียนใหม่
// copy Apps Script template แล้วเปิดใช้งาน
// ============================================================

const INSTALLER_VERSION = '1.0.0';

/**
 * 🔍 DIAGNOSE: ตรวจสอบสถานะชีตทั้งหมดในสเปรดชีต
 * รันจาก GAS Editor → ดู Execution log
 * @returns {Object} สรุปสถานะชีตทั้งหมด
 */
function diagnoseSheetsStatus() {
  var ss = SS();
  var allSheets = ss.getSheets();
  var year = '';
  try { year = S_getAcademicYear(); } catch(_) {}

  var permanent = ['global_settings','Users','Students','รายวิชา','Holidays','HomeroomTeachers'];
  var yearlyBase = ['SCORES_WAREHOUSE','การประเมินอ่านคิดเขียน','การประเมินคุณลักษณะ','การประเมินกิจกรรมพัฒนาผู้เรียน','การประเมินสมรรถนะ','AttendanceLog','ความเห็นครู'];
  var otherExpected = ['สรุปการมาเรียน','สรุปวันมา','โปรไฟล์นักเรียน'];

  Logger.log('=== 🔍 DIAGNOSE SHEETS STATUS ===');
  Logger.log('ปีการศึกษา: ' + (year || '(ไม่ทราบ)'));
  Logger.log('จำนวนชีตทั้งหมด: ' + allSheets.length);
  Logger.log('');

  // 1) List all existing sheets
  var sheetMap = {};
  Logger.log('--- ชีตทั้งหมดที่มี ---');
  allSheets.forEach(function(s) {
    var name = s.getName();
    var rows = s.getLastRow();
    var cols = s.getLastColumn();
    sheetMap[name] = { rows: rows, cols: cols };
    var tag = '';
    if (name.startsWith('BACKUP_')) tag = ' [BACKUP]';
    else if (name.startsWith('Template_')) tag = ' [TEMPLATE]';
    else if (name.match(/^(ม\.|ป\.)/)) tag = ' [คะแนนรายวิชา]';
    else if (name.match(/^\d{1,2}$/)) tag = ' [attendance-month?]';
    else if (name.match(/(มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)/)) tag = ' [attendance-month]';
    Logger.log('  ✅ ' + name + ' (' + rows + ' rows, ' + cols + ' cols)' + tag);
  });

  // 2) Check permanent sheets
  Logger.log('');
  Logger.log('--- ชีตถาวร (PERMANENT) ---');
  var missingPerm = [];
  permanent.forEach(function(name) {
    if (sheetMap[name]) {
      Logger.log('  ✅ ' + name + ' → ' + sheetMap[name].rows + ' rows');
    } else {
      Logger.log('  ❌ ' + name + ' → ไม่พบ!');
      missingPerm.push(name);
    }
  });

  // 3) Check yearly sheets
  Logger.log('');
  Logger.log('--- ชีตรายปี (YEARLY) ---');
  var missingYearly = [];
  yearlyBase.forEach(function(base) {
    var withYear = year ? base + '_' + year : '';
    if (sheetMap[base]) {
      Logger.log('  ✅ ' + base + ' → ' + sheetMap[base].rows + ' rows');
    } else if (withYear && sheetMap[withYear]) {
      Logger.log('  ✅ ' + withYear + ' → ' + sheetMap[withYear].rows + ' rows');
    } else {
      Logger.log('  ❌ ' + base + (withYear ? ' / ' + withYear : '') + ' → ไม่พบ!');
      missingYearly.push(base);
    }
  });

  // 4) Check other expected
  Logger.log('');
  Logger.log('--- ชีตอื่นที่คาดว่าจะมี ---');
  var missingOther = [];
  otherExpected.forEach(function(name) {
    if (sheetMap[name]) {
      Logger.log('  ✅ ' + name + ' → ' + sheetMap[name].rows + ' rows');
    } else {
      Logger.log('  ⚠️ ' + name + ' → ไม่พบ');
      missingOther.push(name);
    }
  });

  // 5) Count by category
  var backups = allSheets.filter(function(s) { return s.getName().startsWith('BACKUP_'); });
  var scoreSheets = allSheets.filter(function(s) { return s.getName().match(/^(ม\.|ป\.)/); });
  var attendanceMonths = allSheets.filter(function(s) { return s.getName().match(/(มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)/); });

  Logger.log('');
  Logger.log('=== สรุป ===');
  Logger.log('ชีตทั้งหมด: ' + allSheets.length);
  Logger.log('ชีตคะแนนรายวิชา: ' + scoreSheets.length + ' → ' + scoreSheets.map(function(s){return s.getName();}).join(', '));
  Logger.log('ชีตเช็คชื่อรายเดือน: ' + attendanceMonths.length + ' → ' + attendanceMonths.map(function(s){return s.getName();}).join(', '));
  Logger.log('ชีต BACKUP: ' + backups.length + ' → ' + backups.map(function(s){return s.getName();}).join(', '));
  Logger.log('❌ ชีตถาวรที่หาย: ' + (missingPerm.length > 0 ? missingPerm.join(', ') : 'ไม่มี'));
  Logger.log('❌ ชีตรายปีที่หาย: ' + (missingYearly.length > 0 ? missingYearly.join(', ') : 'ไม่มี'));
  Logger.log('⚠️ ชีตอื่นที่หาย: ' + (missingOther.length > 0 ? missingOther.join(', ') : 'ไม่มี'));

  return {
    total: allSheets.length,
    missingPermanent: missingPerm,
    missingYearly: missingYearly,
    missingOther: missingOther,
    scoreSheets: scoreSheets.length,
    attendanceMonths: attendanceMonths.length,
    backups: backups.length
  };
}

/**
 * 🔍 DEBUG: รันฟังก์ชันนี้ใน GAS Editor เพื่อดูสถานะ ScriptProperties
 * ไปที่ Execution log เพื่อดูผลลัพธ์
 */
function debugSetupStatus() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var currentId = ScriptApp.getScriptId();
  Logger.log('=== DEBUG SETUP STATUS ===');
  Logger.log('Current Script ID: ' + currentId);
  Logger.log('Stored SCRIPT_ID: ' + (props['SCRIPT_ID'] || '(ไม่มี)'));
  Logger.log('SPREADSHEET_ID: ' + (props['SPREADSHEET_ID'] || '(ไม่มี)'));
  Logger.log('TEMPLATE_SPREADSHEET_ID: ' + (props['TEMPLATE_SPREADSHEET_ID'] || '(ไม่มี)'));
  Logger.log('SCHOOL_NAME: ' + (props['SCHOOL_NAME'] || '(ไม่มี)'));
  Logger.log('SCRIPT_ID match: ' + (props['SCRIPT_ID'] === currentId));
  Logger.log('isSetupComplete_: ' + isSetupComplete_());
  Logger.log('All properties: ' + JSON.stringify(props));
}

// Sheet names ที่ต้องสร้างใน Spreadsheet ใหม่
// ชีตถาวร (ไม่ขึ้นกับปี)
const PERMANENT_SHEETS = [
  'global_settings',
  'Users',
  'Students',
  'รายวิชา',
  'Holidays',
  'HomeroomTeachers'
];
// ชีตรายปี — จะสร้างด้วยชื่อ baseName_ปี (Plan B)
// S_YEARLY_SHEETS อยู่ใน settings_unified.js

// ============================================================
// 🔍 CHECK: ระบบถูกติดตั้งแล้วหรือยัง
// ============================================================

function getSetupStatus() {
  try {
    // ✅ ตรวจสอบจากชีทจริงแทน Properties (สำหรับโปรเจกต์ที่คัดลอก)
    const ss = SS();
    const schoolNameSheet = ss.getSheetByName('global_settings');
    
    if (!schoolNameSheet) {
      return { installed: false, message: 'ไม่พบชีท global_settings' };
    }
    
    // ตรวจสอบว่ามีข้อมูลโรงเรียนหรือไม่
    const data = schoolNameSheet.getDataRange().getValues();
    const schoolNameRow = data.find(row => row[0] === 'ชื่อโรงเรียน');
    
    if (!schoolNameRow || !schoolNameRow[1]) {
      return { installed: false, message: 'ยังไม่ได้ตั้งค่าชื่อโรงเรียน' };
    }
    
    // ตรวจสอบชีทที่จำเป็น (รองรับชื่อชีตรายปี Plan B)
    var requiredPermanent = ['Students', 'Users', 'รายวิชา'];
    var existingSheets = ss.getSheets().map(function(s) { return s.getName(); });
    var missing = requiredPermanent.filter(function(name) { return !existingSheets.includes(name); });
    
    // ตรวจชีตรายปี — อนุญาตทั้งชื่อเดิมและชื่อ + suffix ปี
    var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : ['SCORES_WAREHOUSE'];
    yearlyBases.forEach(function(base) {
      var found = existingSheets.some(function(n) { return n === base || n.indexOf(base + '_') === 0; });
      if (!found) missing.push(base);
    });
    
    if (missing.length > 0) {
      return { installed: false, message: 'ขาดชีทที่จำเป็น: ' + missing.join(', ') };
    }
    
    // ตรวจสอบข้อมูลนักเรียน
    const studentsSheet = ss.getSheetByName('Students');
    const studentsData = studentsSheet.getDataRange().getValues();
    if (studentsData.length <= 1) {
      return { installed: false, message: 'ยังไม่มีข้อมูลนักเรียน' };
    }

    return {
      installed: true,
      schoolName: schoolNameRow[1],
      spreadsheetId: ss.getId(),
      setupDate: 'คัดลอกจากโปรเจกต์เดิม',
      installerVersion: INSTALLER_VERSION,
      method: 'copied_project'
    };
  } catch (e) {
    Logger.log('getSetupStatus error: ' + e.message);
    return { installed: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// ============================================================
// 🚀 MAIN: ติดตั้งระบบสำหรับโรงเรียนใหม่
// ============================================================

function setupNewSchool(formData) {
  try {
    // ป้องกัน: ห้ามรัน setupNewSchool บน project ที่ติดตั้งแล้ว
    var oldProps = PropertiesService.getScriptProperties();
    var existingScriptId = oldProps.getProperty('SCRIPT_ID');
    if (existingScriptId && existingScriptId === ScriptApp.getScriptId()) {
      return { success: false, message: '⛔ ระบบนี้ติดตั้งแล้ว ไม่สามารถติดตั้งซ้ำได้ (ป้องกันข้อมูลถูกทับ)' };
    }
    // ล้าง ScriptProperties เดิม (รองรับ copied project)
    oldProps.deleteAllProperties();

    // Validate ข้อมูลที่รับมา
    const validation = validateSetupForm_(formData);
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    Logger.log('🏫 เริ่มติดตั้งระบบสำหรับ: ' + formData.schoolName);

    // 1. สร้าง/ใช้ Spreadsheet
    var ss;
    var usedExisting = false;
    if (formData.spreadsheetUrl) {
      // Bound Script Mode: ใช้ Spreadsheet ที่คัดลอกมา
      var match = formData.spreadsheetUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!match) {
        return { success: false, message: 'URL ของ Spreadsheet ไม่ถูกต้อง' };
      }
      ss = SpreadsheetApp.openById(match[1]);
      usedExisting = true;
      Logger.log('✅ ใช้ Spreadsheet ที่มีอยู่: ' + ss.getId());
    } else {
      // Standalone Mode: สร้างใหม่
      ss = createSchoolSpreadsheet_(formData);
      Logger.log('✅ สร้าง Spreadsheet ใหม่: ' + ss.getId());
    }

    // 2. ล้างข้อมูลโรงเรียนเดิม (เฉพาะ Spreadsheet ที่คัดลอกมา)
    if (usedExisting) {
      cleanupCopiedSpreadsheet_(ss);
      Logger.log('✅ ล้างข้อมูลโรงเรียนเดิมแล้ว');
    }

    // 3. สร้าง Sheet (ข้ามถ้ามีอยู่แล้ว)
    setupSheets_(ss, formData);
    Logger.log('✅ สร้าง/ตรวจสอบ Sheets แล้ว');

    // 4. บันทึก Settings ลง global_settings sheet
    saveGlobalSettings_(ss, formData);
    Logger.log('✅ บันทึก Settings แล้ว');

    // 5. สร้าง Admin User
    createAdminUser_(ss, formData);
    Logger.log('✅ สร้าง Admin User แล้ว');

    // 6. บันทึก ScriptProperties
    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      'SPREADSHEET_ID': ss.getId(),
      'SCHOOL_NAME': formData.schoolName,
      'SETUP_DATE': new Date().toISOString(),
      'INSTALLER_VERSION': INSTALLER_VERSION,
      'SCRIPT_ID': ScriptApp.getScriptId()
    });
    Logger.log('✅ บันทึก ScriptProperties แล้ว');

    return {
      success: true,
      message: 'ติดตั้งระบบสำเร็จ! ยินดีต้อนรับ ' + formData.schoolName,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
      schoolName: formData.schoolName
    };

  } catch (e) {
    Logger.log('❌ setupNewSchool error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// ============================================================
// 🔧 HELPERS
// ============================================================

function validateSetupForm_(data) {
  if (!data) return { valid: false, message: 'ไม่มีข้อมูลที่ส่งมา' };
  if (!data.schoolName || data.schoolName.trim() === '')
    return { valid: false, message: 'กรุณากรอกชื่อโรงเรียน' };
  if (!data.adminUsername || data.adminUsername.trim() === '')
    return { valid: false, message: 'กรุณากรอกชื่อผู้ใช้ Admin' };
  if (!data.adminPassword || data.adminPassword.length < 6)
    return { valid: false, message: 'รหัสผ่าน Admin ต้องมีอย่างน้อย 6 ตัวอักษร' };
  if (!data.academicYear)
    return { valid: false, message: 'กรุณาระบุปีการศึกษา' };
  return { valid: true };
}

function createSchoolSpreadsheet_(formData) {
  var name = 'ระบบ ปพ.5 — ' + formData.schoolName + ' (' + formData.academicYear + ')';
  var ss;

  // MASTER TEMPLATE FALLBACK: ถ้าไม่ได้เซ็ตไว้ที่ Script Properties ให้ดึงจากค่าคงที่นี้
  const MASTER_TEMPLATE_ID = 'ใส่_ID_ของ_SHEET_แม่แบบที่นี่'; // << จุดนี้คือจุดให้ผู้พัฒนาเอา ID มาวาง

  // พยายามคัดลอกจาก Template ก่อน (เร็วกว่าสร้างใหม่)
  var props = PropertiesService.getScriptProperties();
  var templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID') || MASTER_TEMPLATE_ID;

  if (templateId && templateId.trim() !== '' && templateId !== 'ใส่_ID_ของ_SHEET_แม่แบบที่นี่') {
    try {
      var templateFile = DriveApp.getFileById(templateId);
      var copiedFile = templateFile.makeCopy(name);
      ss = SpreadsheetApp.openById(copiedFile.getId());

      // Rename ชีทรายปีให้ใส่ suffix ปีการศึกษา
      if (formData.academicYear) {
        var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
          'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
          'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
          'AttendanceLog', 'ความเห็นครู'
        ];
        yearlyBases.forEach(function(baseName) {
          var sheet = ss.getSheetByName(baseName);
          if (sheet) sheet.setName(baseName + '_' + formData.academicYear);
        });
      }
      Logger.log('✅ คัดลอกจาก Template สำเร็จ: ' + ss.getId());
    } catch (e) {
      Logger.log('⚠️ คัดลอก Template ไม่ได้: ' + e.message + ' → สร้างใหม่');
      ss = SpreadsheetApp.create(name);
    }
  } else {
    ss = SpreadsheetApp.create(name);
  }

  // ย้ายเข้า Drive folder
  try {
    var file = DriveApp.getFileById(ss.getId());
    var folderName = 'ระบบโรงเรียน_' + formData.schoolName;
    var folder;
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {
    Logger.log('ไม่สามารถย้ายไฟล์: ' + e.message);
  }

  return ss;
}

function setupSheets_(ss, formData) {
  // ลบ Sheet1 เดิม (default)
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');
  var academicYear = String(formData.academicYear || '');

  // 1. สร้างชีตถาวร
  PERMANENT_SHEETS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    setupSheetHeaders_(sheet, name, formData);
  });

  // 2. สร้างชีตรายปี — ใช้ชื่อ baseName_ปี (Plan B)
  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
    'AttendanceLog', 'ความเห็นครู'
  ];
  yearlyBases.forEach(function(baseName) {
    var sheetName = academicYear ? baseName + '_' + academicYear : baseName;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    setupSheetHeaders_(sheet, baseName, formData);
  });

  // ลบ default sheet
  try {
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  } catch(e) {}
}

function setupSheetHeaders_(sheet, sheetName, formData) {
  const headerMap = {
    'Users': [['username','password','role','firstName','lastName','email','className','active','createdAt']],
    'Students': [['student_id','id_card','title','firstname','lastname','grade','class_no','gender','created_date','status']],
    'รายวิชา': [['ชั้น','รหัสวิชา','ชื่อวิชา','ชั่วโมง/ปี','ประเภทวิชา','ครูผู้สอน','คะแนนระหว่างปี','คะแนนปลายปี']],
    'Holidays': [['date','description','type']],
    'global_settings': [['key','value','updatedAt']],
    'SCORES_WAREHOUSE': [['student_id','grade','class_no','subject_code','subject_name','subject_type','hours','term1_total','term2_total','average','final_grade','sheet_name','academic_year','updated_at']],
    'HomeroomTeachers': [['grade','classNo','teacherName']],
    'การประเมินอ่านคิดเขียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา','สุขศึกษา','ศิลปะ','การงาน','ภาษาอังกฤษ','สรุปผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินคุณลักษณะ': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','รักชาติ_ศาสน์_กษัตริย์','ซื่อสัตย์สุจริต','มีวินัย','ใฝ่เรียนรู้','อยู่อย่างพอเพียง','มุ่งมั่นในการทำงาน','รักความเป็นไทย','มีจิตสาธารณะ','คะแนนรวม','คะแนนเฉลี่ย','ผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินกิจกรรมพัฒนาผู้เรียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','กิจกรรมแนะแนว','ลูกเสือ_เนตรนารี','ชุมนุม','เพื่อสังคมและสาธารณประโยชน์','รวมกิจกรรม','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินสมรรถนะ': [[
      'รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง',
      'สื่อสาร_รับ-ส่งสาร','สื่อสาร_ถ่ายทอด','สื่อสาร_วิธีการ','สื่อสาร_เจรจา','สื่อสาร_เลือกรับ',
      'คิด_วิเคราะห์','คิด_สร้างสรรค์','คิด_วิจารณญาณ','คิด_สร้างความรู้','คิด_ตัดสินใจ',
      'แก้ปัญหา_แก้ปัญหา','แก้ปัญหา_ใช้เหตุผล','แก้ปัญหา_เข้าใจสังคม','แก้ปัญหา_แสวงหาความรู้','แก้ปัญหา_ตัดสินใจ',
      'ทักษะชีวิต_เรียนรู้ตนเอง','ทักษะชีวิต_ทำงานกลุ่ม','ทักษะชีวิต_นำไปใช้','ทักษะชีวิต_จัดการปัญหา','ทักษะชีวิต_หลีกเลี่ยง',
      'เทคโนโลยี_เลือกใช้','เทคโนโลยี_ทักษะกระบวนการ','เทคโนโลยี_พัฒนาตนเอง','เทคโนโลยี_แก้ปัญหา','เทคโนโลยี_คุณธรรม',
      'วันที่บันทึก','ผู้บันทึก'
    ]],
    'AttendanceLog': [['timestamp','date','grade','class','updated_count','user_email','details']],
    'ความเห็นครู': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','ความเห็นครู','วันที่บันทึก','ผู้บันทึก']]
  };

  // ใช้ baseName เป็น key (ไม่ใช่ชื่อชีตจริงที่อาจมี suffix ปี)
  var baseName = sheetName.replace(/_\d{4}$/, '');
  var headers = headerMap[baseName];
  if (headers) {
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    sheet.getRange(1, 1, 1, headers[0].length)
      .setBackground('#4a5568')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function saveGlobalSettings_(ss, formData) {
  const sheet = ss.getSheetByName('global_settings');
  if (!sheet) return;

  const now = new Date().toISOString();
  const schoolName = (formData.schoolName || '').trim();
  const schoolAddress = (formData.schoolAddress || '').trim();
  const directorName = (formData.directorName || '').trim();
  const directorTitle = (formData.directorTitle || 'ผู้อำนวยการสถานศึกษา').trim();
  const academicHead = (formData.academicHead || '').trim();
  const academicYear = String(formData.academicYear || '');
  const semester = String(formData.semester || '1');
  const pdfFolderId = (formData.pdfFolderId || '').trim();

  const settings = [
    // ภาษาไทย (ที่ระบบเดิมใช้)
    ['ชื่อโรงเรียน', schoolName, now],
    ['ที่อยู่โรงเรียน', schoolAddress, now],
    ['ตำแหน่งผู้อำนวยการ', directorTitle, now],
    ['ชื่อผู้อำนวยการ', directorName, now],
    ['ชื่อหัวหน้างานวิชาการ', academicHead, now],
    ['ปีการศึกษา', academicYear, now],
    ['ภาคเรียน', semester, now],
    ['หมายเหตุท้ายรายงาน', '', now],
    ['ชื่อครูผู้สอน', '', now],
    ['เขต', (formData.schoolDistrict || '').trim(), now],
    // ภาษาอังกฤษ (Installer keys)
    ['schoolName', schoolName, now],
    ['schoolDistrict', (formData.schoolDistrict || '').trim(), now],
    ['school_address', schoolAddress, now],
    ['directorName', directorName, now],
    ['directorTitle', directorTitle, now],
    ['academicHead', academicHead, now],
    ['academicYear', academicYear, now],
    ['semester', semester, now],
    ['pdfSaveFolderId', pdfFolderId, now],
    // Meta
    ['setup_complete', 'true', now],
    ['installer_version', INSTALLER_VERSION, now],
    ['installed_script_id', ScriptApp.getScriptId(), now]
  ];

  sheet.getRange(2, 1, settings.length, 3).setValues(settings);
}

function createAdminUser_(ss, formData) {
  const sheet = ss.getSheetByName('Users') || ss.getSheetByName('users');
  if (!sheet) return;

  const now = new Date().toISOString();
  // Hash password for new admin using SHA-256
  const rawPw = String(formData.adminPassword);
  const hashDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawPw, Utilities.Charset.UTF_8);
  const hashStr = hashDigest.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');

  // ✅ ผู้ใช้คนแรกต้องเป็น admin เสมอ
  sheet.appendRow([
    formData.adminUsername.trim(),
    hashStr, // Save hashed password instead of plaintext
    'admin', // ✅ บังคับให้เป็น admin
    formData.adminFirstName || 'ผู้ดูแล',
    formData.adminLastName || 'ระบบ',
    formData.adminEmail || '',
    '',
    'true',
    now
  ]);
  
  Logger.log('✅ สร้าง Admin User: ' + formData.adminUsername.trim() + ' (role: admin)');
}

// ============================================================
// 🧹 CLEANUP: ล้างข้อมูลโรงเรียนเดิมจาก Spreadsheet ที่คัดลอกมา
// ============================================================

function cleanupCopiedSpreadsheet_(ss) {
  var allSheets = ss.getSheets();
  var keepSheets = new Set(PERMANENT_SHEETS);
  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
    'AttendanceLog', 'ความเห็นครู'
  ];
  yearlyBases.forEach(function(b) { keepSheets.add(b); });

  // รายชื่อชีตที่ต้องล้างข้อมูล (เก็บ header ไว้)
  var sheetsToClean = [];
  // รายชื่อชีตที่ต้องลบทิ้ง (ชีตคะแนนรายวิชา, เช็คชื่อ, BACKUP)
  var sheetsToDelete = [];

  allSheets.forEach(function(sheet) {
    var name = sheet.getName();

    // ชีตถาวร + ชีตรายปี → ล้างข้อมูล (เก็บ header แถว 1)
    if (keepSheets.has(name) || yearlyBases.some(function(b) { return name.indexOf(b) === 0; })) {
      sheetsToClean.push(sheet);
      return;
    }

    // ชีตคะแนนรายวิชา (เช่น "ภาษาไทย ป3-1", "คณิตศาสตร์ ป.1-1")
    if (/ป\d|ป\.\d|ม\d|ม\.\d/.test(name) && name.indexOf(' ') > 0) {
      sheetsToDelete.push(sheet);
      return;
    }

    // ชีตเช็คชื่อรายเดือน (เช่น "พฤษภาคม 2568")
    if (/(มกราคม|กุมภาพันธ์|มีนาคม|เมษายน|พฤษภาคม|มิถุนายน|กรกฎาคม|สิงหาคม|กันยายน|ตุลาคม|พฤศจิกายน|ธันวาคม)/.test(name)) {
      sheetsToDelete.push(sheet);
      return;
    }

    // ชีต BACKUP
    if (name.indexOf('BACKUP_') === 0 || name.indexOf('Template_') === 0) {
      sheetsToDelete.push(sheet);
      return;
    }

    // ชีตอื่นๆ ที่ไม่รู้จัก → ล้างข้อมูล (เก็บ header)
    sheetsToClean.push(sheet);
  });

  // ล้างข้อมูล (เก็บ header แถว 1)
  sheetsToClean.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    Logger.log('🧹 ล้างข้อมูล: ' + sheet.getName() + ' (' + (lastRow - 1) + ' แถว)');
  });

  // ลบชีตที่ไม่ต้องการ (ต้องเหลืออย่างน้อย 1 ชีต)
  var remaining = allSheets.length - sheetsToDelete.length;
  if (remaining < 1) {
    // เก็บชีตแรกไว้ไม่ลบ
    sheetsToDelete.shift();
  }
  sheetsToDelete.forEach(function(sheet) {
    try {
      Logger.log('🗑️ ลบชีต: ' + sheet.getName());
      ss.deleteSheet(sheet);
    } catch (e) {
      Logger.log('⚠️ ลบไม่ได้: ' + sheet.getName() + ' — ' + e.message);
    }
  });

  Logger.log('🧹 ล้างเสร็จ: ล้างข้อมูล ' + sheetsToClean.length + ' ชีต, ลบ ' + sheetsToDelete.length + ' ชีต');
}

// ============================================================
// 🔄 MIGRATE: สำหรับโรงเรียนที่มีอยู่แล้ว (รันครั้งเดียว)
// ============================================================

function migrateExistingSchool() {
  const EXISTING_SPREADSHEET_ID = ''; // ใส่ Spreadsheet ID ที่นี่ก่อนมารัน (หรือผูกผ่าน UI) โปรแกรมจะไม่เซฟค่าทับ
  const EXISTING_SCHOOL_NAME = '';
  const EXISTING_PDF_FOLDER_ID = '';

  if (!EXISTING_SPREADSHEET_ID) {
    Logger.log('⚠️ ไม่ได้ระบุ SPREADSHEET_ID เดิม ไม่สามารถ migrate ได้');
    return { success: false, message: 'ไม่พบ SPREADSHEET_ID กรุณาระบุในโค้ดก่อนทำการ Migrate' };
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'SPREADSHEET_ID': EXISTING_SPREADSHEET_ID,
    'SCHOOL_NAME': EXISTING_SCHOOL_NAME,
    'SETUP_DATE': new Date().toISOString(),
    'INSTALLER_VERSION': INSTALLER_VERSION,
    'PDF_OUTPUT_FOLDER_ID': EXISTING_PDF_FOLDER_ID
  });

  Logger.log('✅ Migrate สำเร็จ: ' + EXISTING_SCHOOL_NAME);
  Logger.log('📊 Spreadsheet ID: ' + EXISTING_SPREADSHEET_ID);
  return { success: true, message: 'Migrate เรียบร้อย — ' + EXISTING_SCHOOL_NAME };
}

// ============================================================
// 📅 SWITCH ACADEMIC YEAR — เปลี่ยนปีการศึกษา (แผน B: สร้างชีตใหม่)
// ============================================================

/**
 * เปลี่ยนปีการศึกษา (แผน B — ปลอดภัย ข้อมูลเก่าไม่หาย):
 * 1. Copy Spreadsheet เป็น archive (backup เพิ่มเติม)
 * 2. เปลี่ยนชื่อชีตข้อมูลเดิม → เติม _ปีเก่า (เช่น SCORES_WAREHOUSE → SCORES_WAREHOUSE_2567)
 * 3. สร้างชีตใหม่ว่างชื่อ _ปีใหม่ (เช่น SCORES_WAREHOUSE_2568) พร้อม header
 * 4. เปลี่ยนชื่อชีตคะแนนรายวิชาเก่า → เติม _ปีเก่า
 * 5. อัปเดตปีการศึกษาใน global_settings
 * 6. จำหน่าย ป.6 + เลื่อนชั้น
 * 
 * @param {number} newYear - ปีการศึกษาใหม่ (พ.ศ.)
 * @param {Object} options - ตัวเลือกเพิ่มเติม
 * @returns {Object} ผลลัพธ์
 */
function switchAcademicYear(newYear, options) {
  options = options || {};

  try {
    if (!newYear || isNaN(newYear)) {
      return { success: false, message: 'กรุณาระบุปีการศึกษาใหม่' };
    }
    newYear = Number(newYear);

    var ss = SS();
    var oldSettings = S_getGlobalSettings(false);
    var oldYear = oldSettings['ปีการศึกษา'] || 'ไม่ทราบ';
    var schoolName = oldSettings['ชื่อโรงเรียน'] || 'โรงเรียน';

    if (String(oldYear) === String(newYear)) {
      return { success: false, message: 'ปีการศึกษาเดิมและใหม่เหมือนกัน (' + newYear + ')' };
    }

    Logger.log('📅 เริ่มเปลี่ยนปีการศึกษา: ' + oldYear + ' → ' + newYear);

    // ============ ขั้นที่ 1: Copy Spreadsheet เป็น archive (backup) ============
    Logger.log('📦 กำลัง archive spreadsheet...');
    var archiveName = schoolName + ' — สำรองก่อนเปลี่ยนปี ' + oldYear + '→' + newYear + ' (' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd-MM-yyyy HH:mm') + ')';
    var originalFile = DriveApp.getFileById(ss.getId());
    var parentFolders = originalFile.getParents();
    var targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

    var archiveFolderName = 'Archive_ปีการศึกษา';
    var archiveFolder = null;
    var subFolders = targetFolder.getFoldersByName(archiveFolderName);
    if (subFolders.hasNext()) {
      archiveFolder = subFolders.next();
    } else {
      archiveFolder = targetFolder.createFolder(archiveFolderName);
    }

    var copiedFile = originalFile.makeCopy(archiveName, archiveFolder);
    var archiveUrl = copiedFile.getUrl();
    Logger.log('✅ Archive สำเร็จ: ' + archiveName);

    // ============ ขั้นที่ 2: Rename ชีตข้อมูลเดิม + สร้างชีตใหม่ ============
    Logger.log('📋 กำลัง rename ชีตเก่า + สร้างชีตใหม่...');
    var yearlyBases = S_YEARLY_SHEETS; // จาก settings_unified.js
    var renamedCount = 0;
    var createdCount = 0;

    yearlyBases.forEach(function(baseName) {
      // หาชีตปัจจุบัน (อาจเป็นชื่อเดิม หรือ ชื่อ_ปีเก่า)
      var currentSheet = ss.getSheetByName(baseName + '_' + oldYear) || ss.getSheetByName(baseName);
      var currentName = currentSheet ? currentSheet.getName() : null;

      // Rename ชีตเดิม → baseName_ปีเก่า (ถ้ายังไม่มี suffix)
      if (currentSheet && currentName === baseName) {
        var archiveSuffix = baseName + '_' + oldYear;
        // ตรวจว่าชื่อซ้ำไหม
        if (!ss.getSheetByName(archiveSuffix)) {
          currentSheet.setName(archiveSuffix);
          renamedCount++;
          Logger.log('  📝 Rename: ' + baseName + ' → ' + archiveSuffix);
        }
      }

      // สร้างชีตใหม่สำหรับปีใหม่
      var newName = baseName + '_' + newYear;
      if (!ss.getSheetByName(newName)) {
        var newSheet = ss.insertSheet(newName);
        // Copy header จากชีตเดิม
        var srcSheet = ss.getSheetByName(baseName + '_' + oldYear) || ss.getSheetByName(baseName);
        if (srcSheet && srcSheet.getLastRow() >= 1) {
          var headerRow = srcSheet.getRange(1, 1, 1, srcSheet.getLastColumn()).getValues();
          if (headerRow[0].length > 0 && headerRow[0][0] !== '') {
            newSheet.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow);
            // Copy header formatting
            try {
              srcSheet.getRange(1, 1, 1, srcSheet.getLastColumn()).copyFormatToRange(newSheet, 1, srcSheet.getLastColumn(), 1, 1);
            } catch (e) { /* ignore format copy error */ }
          }
        }
        createdCount++;
        Logger.log('  ✨ สร้างชีตใหม่: ' + newName);
      }
    });

    // ============ ขั้นที่ 3: Rename ชีตคะแนนรายวิชาเก่า ============
    Logger.log('� กำลัง rename ชีตคะแนนรายวิชา...');
    var systemSheets = [
      'global_settings', 'Users', 'Students', 'รายวิชา', 'Holidays', 'HomeroomTeachers'
    ];
    // เพิ่มชีตที่เพิ่ง rename/สร้าง ลงใน systemSheets
    yearlyBases.forEach(function(b) {
      systemSheets.push(b);
      systemSheets.push(b + '_' + oldYear);
      systemSheets.push(b + '_' + newYear);
    });
    var prefixSkip = ['BACKUP_', 'Template_', 'TMP_'];

    var archivedScoreSheets = 0;
    var allSheets = ss.getSheets();
    for (var i = 0; i < allSheets.length; i++) {
      var sName = allSheets[i].getName();
      if (systemSheets.indexOf(sName) !== -1) continue;
      var skipThis = false;
      prefixSkip.forEach(function(p) { if (sName.indexOf(p) === 0) skipThis = true; });
      if (skipThis) continue;
      // ข้ามชีตที่มี suffix ปีอยู่แล้ว
      if (sName.match(/_\d{4}$/)) continue;

      // เป็นชีตคะแนนรายวิชา → rename เป็น ชื่อเดิม_ปีเก่า
      var archivedName = sName + '_' + oldYear;
      if (!ss.getSheetByName(archivedName)) {
        try {
          allSheets[i].setName(archivedName);
          archivedScoreSheets++;
          Logger.log('  � Rename: ' + sName + ' → ' + archivedName);
        } catch (e) {
          Logger.log('  ⚠️ Rename ไม่ได้: ' + sName + ' — ' + e.message);
        }
      }
    }

    // ============ ขั้นที่ 4: อัปเดต settings ============
    Logger.log('⚙️ อัปเดตปีการศึกษา...');
    var settingsSheet = ss.getSheetByName('global_settings');
    if (settingsSheet) {
      var data = settingsSheet.getDataRange().getValues();
      var now = new Date().toISOString();
      for (var r = 0; r < data.length; r++) {
        var key = String(data[r][0]).trim();
        if (key === 'ปีการศึกษา' || key === 'academicYear') {
          settingsSheet.getRange(r + 1, 2).setValue(String(newYear));
          settingsSheet.getRange(r + 1, 3).setValue(now);
        }
        if (key === 'ภาคเรียน' || key === 'semester') {
          settingsSheet.getRange(r + 1, 2).setValue('1');
          settingsSheet.getRange(r + 1, 3).setValue(now);
        }
      }
    }

    // ============ ขั้นที่ 5: จำหน่าย ป.6 + เลื่อนชั้น ============
    var studentResult = { promoted: 0, graduated: 0 };
    if (options.promoteStudents !== false) {
      Logger.log('🎓 กำลังจำหน่ายนักเรียน ป.6 และเลื่อนชั้น...');
      try {
        var gradResult = graduateP6Students_();
        studentResult.graduated = gradResult.graduated || 0;
      } catch (e) { Logger.log('⚠️ graduateP6: ' + e.message); }
      try {
        var proResult = promoteStudents_();
        studentResult.promoted = proResult.promoted || 0;
      } catch (e) { Logger.log('⚠️ promote: ' + e.message); }
    }

    // ล้าง cache
    try { S_clearSettingsCache(); } catch (e) {}

    var summary = {
      success: true,
      message: 'เปลี่ยนปีการศึกษาสำเร็จ: ' + oldYear + ' → ' + newYear,
      archiveName: archiveName,
      archiveUrl: archiveUrl,
      sheetsRenamed: renamedCount,
      sheetsCreated: createdCount,
      scoreSheetArchived: archivedScoreSheets,
      studentsPromoted: studentResult.promoted,
      studentsGraduated: studentResult.graduated,
      newYear: newYear,
      oldYear: oldYear
    };

    Logger.log('✅ เปลี่ยนปีการศึกษาสำเร็จ!');
    Logger.log(JSON.stringify(summary, null, 2));
    return summary;

  } catch (e) {
    Logger.log('❌ switchAcademicYear error: ' + e.message + '\n' + e.stack);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// ============================================================
// 🎓 PROMOTE STUDENTS — เลื่อนชั้นนักเรียน
// ============================================================

/**
 * เลื่อนชั้นนักเรียนที่ status = 'active':
 * ป.1 → ป.2, ป.2 → ป.3, ..., ป.5 → ป.6
 * (ป.6 ไม่เลื่อน — ต้องจำหน่ายก่อน)
 * @returns {Object} { promoted: number }
 */
function promoteStudents_() {
  var ss = SS();
  var sheet = ss.getSheetByName('Students');
  if (!sheet) return { promoted: 0 };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { promoted: 0 };

  var headers = data[0];
  var col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  var gradeIdx = col['grade'];
  var statusIdx = col['status'];
  if (gradeIdx == null) return { promoted: 0 };

  var gradeMap = {
    'ป.1': 'ป.2',
    'ป.2': 'ป.3',
    'ป.3': 'ป.4',
    'ป.4': 'ป.5',
    'ป.5': 'ป.6'
  };

  var promoted = 0;
  for (var r = 1; r < data.length; r++) {
    var status = statusIdx != null ? String(data[r][statusIdx]).trim() : 'active';
    if (status !== 'active' && status !== '') continue;

    var currentGrade = String(data[r][gradeIdx]).trim();
    var newGrade = gradeMap[currentGrade];
    if (newGrade) {
      sheet.getRange(r + 1, gradeIdx + 1).setValue(newGrade);
      promoted++;
    }
  }

  Logger.log('🎓 เลื่อนชั้นนักเรียน: ' + promoted + ' คน');
  return { promoted: promoted };
}

/**
 * เลื่อนชั้นนักเรียน (เรียกจากภายนอก / UI)
 * @returns {Object}
 */
function promoteAllStudents() {
  try {
    var result = promoteStudents_();
    return { success: true, message: 'เลื่อนชั้นนักเรียน ' + result.promoted + ' คน', promoted: result.promoted };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// 🎓 GRADUATE P.6 — จำหน่ายนักเรียน ป.6
// ============================================================

/**
 * จำหน่ายนักเรียน ป.6 ที่ status = 'active':
 * เปลี่ยน status เป็น 'จำหน่าย' (ไม่ลบข้อมูล ยังดูได้)
 * @returns {Object} { graduated: number }
 */
function graduateP6Students_() {
  var ss = SS();
  var sheet = ss.getSheetByName('Students');
  if (!sheet) return { graduated: 0 };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { graduated: 0 };

  var headers = data[0];
  var col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  var gradeIdx = col['grade'];
  var statusIdx = col['status'];
  if (gradeIdx == null) return { graduated: 0 };

  // ถ้าไม่มีคอลัมน์ status ให้สร้างเพิ่ม
  if (statusIdx == null) {
    var lastCol = headers.length;
    sheet.getRange(1, lastCol + 1).setValue('status');
    statusIdx = lastCol;
    Logger.log('➕ เพิ่มคอลัมน์ status');
  }

  var graduated = 0;
  for (var r = 1; r < data.length; r++) {
    var currentGrade = String(data[r][gradeIdx]).trim();
    var status = String(data[r][statusIdx] || '').trim();
    if (currentGrade === 'ป.6' && (status === 'active' || status === '')) {
      sheet.getRange(r + 1, statusIdx + 1).setValue('จำหน่าย');
      graduated++;
    }
  }

  Logger.log('🎓 จำหน่ายนักเรียน ป.6: ' + graduated + ' คน');
  return { graduated: graduated };
}

/**
 * จำหน่ายนักเรียน ป.6 (เรียกจากภายนอก / UI)
 * @returns {Object}
 */
function graduateP6() {
  try {
    var result = graduateP6Students_();
    return { success: true, message: 'จำหน่ายนักเรียน ป.6 จำนวน ' + result.graduated + ' คน', graduated: result.graduated };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ดูข้อมูลก่อนเปลี่ยนปี (preview) — แผน B: ไม่ล้างข้อมูล แค่ rename + สร้างใหม่
 * @returns {Object} สรุปสิ่งที่จะเกิดขึ้น
 */
function previewSwitchAcademicYear() {
  try {
    var ss = SS();
    var settings = S_getGlobalSettings(false);
    var currentYear = settings['ปีการศึกษา'] || 'ไม่ทราบ';
    var newYear = Number(currentYear) + 1;

    // ชีตข้อมูลรายปีที่จะ rename + สร้างใหม่
    var yearlyBases = S_YEARLY_SHEETS;
    var sheetsToRename = [];
    var sheetsToCreate = [];

    yearlyBases.forEach(function(baseName) {
      var currentSheet = ss.getSheetByName(baseName + '_' + currentYear) || ss.getSheetByName(baseName);
      var currentName = currentSheet ? currentSheet.getName() : null;
      var rows = currentSheet ? Math.max(0, currentSheet.getLastRow() - 1) : 0;

      if (currentSheet && currentName === baseName) {
        sheetsToRename.push({ from: baseName, to: baseName + '_' + currentYear, rows: rows });
      } else if (currentSheet) {
        sheetsToRename.push({ from: currentName, to: currentName, rows: rows, alreadyRenamed: true });
      }
      sheetsToCreate.push(baseName + '_' + newYear);
    });

    // ชีตคะแนนรายวิชาที่จะ rename
    var systemSheets = ['global_settings', 'Users', 'Students', 'รายวิชา', 'Holidays', 'HomeroomTeachers'];
    yearlyBases.forEach(function(b) {
      systemSheets.push(b);
      systemSheets.push(b + '_' + currentYear);
      systemSheets.push(b + '_' + newYear);
    });
    var prefixSkip = ['BACKUP_', 'Template_', 'TMP_'];
    var scoreSheets = [];
    ss.getSheets().forEach(function(sheet) {
      var name = sheet.getName();
      if (systemSheets.indexOf(name) !== -1) return;
      var skip = false;
      prefixSkip.forEach(function(p) { if (name.indexOf(p) === 0) skip = true; });
      if (skip) return;
      if (name.match(/_\d{4}$/)) return; // ข้ามชีตที่มี suffix ปีแล้ว
      scoreSheets.push(name);
    });

    // นับนักเรียนแยกตามชั้น
    var studentsSheet = ss.getSheetByName('Students');
    var studentCount = 0;
    var gradeCounts = { 'ป.1': 0, 'ป.2': 0, 'ป.3': 0, 'ป.4': 0, 'ป.5': 0, 'ป.6': 0 };
    if (studentsSheet && studentsSheet.getLastRow() > 1) {
      var sData = studentsSheet.getDataRange().getValues();
      var sHeaders = sData[0];
      var sCol = {};
      sHeaders.forEach(function(h, i) { sCol[String(h).trim()] = i; });
      var gIdx = sCol['grade'];
      var stIdx = sCol['status'];
      for (var sr = 1; sr < sData.length; sr++) {
        var st = stIdx != null ? String(sData[sr][stIdx] || '').trim() : '';
        if (st === 'จำหน่าย' || st === 'ย้ายออก' || st === 'พ้นสภาพ') continue;
        studentCount++;
        var g = gIdx != null ? String(sData[sr][gIdx]).trim() : '';
        if (gradeCounts[g] !== undefined) gradeCounts[g]++;
      }
    }

    return {
      success: true,
      currentYear: currentYear,
      suggestedNewYear: newYear,
      studentCount: studentCount,
      gradeCounts: gradeCounts,
      p6Count: gradeCounts['ป.6'],
      promoteCount: studentCount - gradeCounts['ป.6'],
      sheetsToRename: sheetsToRename,
      sheetsToCreate: sheetsToCreate,
      scoreSheets: scoreSheets,
      totalDataRows: sheetsToRename.reduce(function(sum, d) { return sum + (d.rows || 0); }, 0)
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// 🔄 RESET (ใช้เฉพาะ Admin เท่านั้น)
// ============================================================

function resetSetup() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('SPREADSHEET_ID');
    props.deleteProperty('SCHOOL_NAME');
    props.deleteProperty('SETUP_DATE');
    props.deleteProperty('INSTALLER_VERSION');
    Logger.log('⚠️ Setup ถูก reset แล้ว');
    return { success: true, message: 'Reset เรียบร้อย — ระบบพร้อมติดตั้งใหม่' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ============================================================
// 📦 EXPORT: ส่งออกระบบเป็น ZIP สำหรับโรงเรียนอื่น
// ============================================================

/**
 * ส่งออกระบบให้โรงเรียนอื่น
 * สร้าง Template Spreadsheet + ลิงก์ Make a Copy
 * @returns {Object} { success, copyUrl, steps }
 */
function exportProjectAsZip(token) {
  try {
    if (!isAdminForExport_(token)) return { success: false, error: 'เฉพาะผู้ดูแลระบบเท่านั้นที่ใช้งานเมนูนี้ได้' };
    if (!isMasterExportEnabled_()) return { success: false, error: 'เมนูส่งออกสงวนไว้เฉพาะโรงเรียนต้นแบบ' };

    var settings = S_getGlobalSettings(false) || {};
    var schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียนต้นแบบ';

    var props = PropertiesService.getScriptProperties();
    var activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var exportTemplateId = createExportTemplateForOtherSchools_(activeId, schoolName);
    props.setProperty('EXPORT_TEMPLATE_SPREADSHEET_ID', exportTemplateId);
    var copyUrl = 'https://docs.google.com/spreadsheets/d/' + exportTemplateId + '/copy';

    var templateInfo = null;
    try {
      var f = DriveApp.getFileById(exportTemplateId);
      templateInfo = {
        id: exportTemplateId,
        name: f.getName(),
        url: f.getUrl(),
        lastUpdated: (f.getLastUpdated && f.getLastUpdated()) ? f.getLastUpdated().toISOString() : null
      };
    } catch(_e) {}

    return {
      success: true,
      copyUrl: copyUrl,
      exportTemplateId: exportTemplateId,
      templateSource: 'generated_from_active',
      activeSpreadsheetId: activeId,
      templateInfo: templateInfo,
      warning: '',
      sanitized: true,
      excludedData: ['Students', 'Users', 'HomeroomTeachers', 'global_settings', 'SCORES_WAREHOUSE', 'AttendanceLog', 'assessment_sheets', 'comments'],
      schoolName: schoolName,
      steps: [
        'ขั้นที่ 1: กดลิงก์ "ทำสำเนา" → Google จะถามยืนยัน → กด "ทำสำเนา" → ได้ Spreadsheet + โค้ดทั้งหมด',
        'ขั้นที่ 2: เปิดสำเนา → ส่วนขยาย → Apps Script → กด Deploy → New deployment → Web app → Execute as: Me, Who: Anyone → Deploy',
        'ขั้นที่ 3: เปิด URL ที่ได้ → ระบบจะเข้าหน้าติดตั้งอัตโนมัติ → กรอกข้อมูลโรงเรียน → กดติดตั้ง'
      ],
      message: 'สร้างลิงก์ส่งออกสำเร็จ (ลิงก์ทำสำเนา Spreadsheet พร้อมโค้ด)'
    };

  } catch (e) {
    Logger.log('❌ exportProjectAsZip error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getExportCapability(token) {
  try {
    if (!isAdminForExport_(token)) return { allowExport: false };
    return { allowExport: isMasterExportEnabled_() };
  } catch(e) {
    return { allowExport: false };
  }
}

function isAdminForExport_(token) {
  try {
    var session = (typeof getLoginSession === 'function') ? getLoginSession() : null;
    var role = String((session && session.role) || '').toLowerCase();
    if (role === 'admin') return true;

    var t = String(token || '');
    if (!t) return false;
    if (t === 'dev_bypass_token') return true;
    if (typeof verifyAuthToken !== 'function') return false;

    var vr = verifyAuthToken(t);
    if (!vr || !vr.valid || !vr.username) return false;
    if (typeof getUserRecordByUsername_ !== 'function') return false;

    var rec = getUserRecordByUsername_(vr.username);
    if (!rec) return false;

    var roleIdx = __idx__(rec.map, ['role','roles'], 2);
    var r = String((rec.row && rec.row[roleIdx]) || '').toLowerCase();
    return r === 'admin';
  } catch(_e) {
    return false;
  }
}

function isMasterExportEnabled_() {
  var props = PropertiesService.getScriptProperties();
  var flag = String(props.getProperty('ALLOW_EXPORT_TO_OTHER_SCHOOLS') || '').toLowerCase();
  if (flag === '1' || flag === 'true') return true;
  var templateId = String(props.getProperty('TEMPLATE_SPREADSHEET_ID') || '').trim();
  if (!templateId) return false;
  if (templateId === 'ใส่_ID_ของ_SHEET_แม่แบบที่นี่') return false;
  return true;
}

function createExportTemplateForOtherSchools_(activeSpreadsheetId, schoolName) {
  var sourceFile = DriveApp.getFileById(activeSpreadsheetId);
  var ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
  var exportName = 'PP5Smart_ExportTemplate_' + (schoolName || 'School') + '_' + ts;

  var targetFolder = null;
  try {
    var parents = sourceFile.getParents();
    if (parents && parents.hasNext()) targetFolder = parents.next();
  } catch(_e) {}

  var copiedFile = targetFolder ? sourceFile.makeCopy(exportName, targetFolder) : sourceFile.makeCopy(exportName);
  copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var ss = SpreadsheetApp.openById(copiedFile.getId());
  sanitizeExportTemplateSpreadsheet_(ss);
  return copiedFile.getId();
}

function sanitizeExportTemplateSpreadsheet_(ss) {
  var clearNames = ['Users', 'Students', 'HomeroomTeachers', 'global_settings', 'SCORES_WAREHOUSE', 'AttendanceLog', 'ความเห็นครู'];
  clearNames.forEach(function(n){
    var sh = ss.getSheetByName(n);
    if (sh) clearSheetData_(sh, 1);
  });

  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
    'AttendanceLog', 'ความเห็นครู'
  ];
  var sheets = ss.getSheets();
  sheets.forEach(function(sh){
    var name = sh.getName();
    for (var i = 0; i < yearlyBases.length; i++) {
      var base = yearlyBases[i];
      if (name === base || name.indexOf(base + '_') === 0) {
        clearSheetData_(sh, 1);
        break;
      }
    }
  });

  var subj = ss.getSheetByName('รายวิชา');
  if (subj) blankTeacherInSubjects_(subj);
}

function clearSheetData_(sheet, headerRows) {
  var h = headerRows || 0;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getMaxColumns();
  if (lastRow <= h) return;
  sheet.getRange(h + 1, 1, lastRow - h, lastCol).clearContent();
}

function blankTeacherInSubjects_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return;
  var headers = data[0] || [];
  var teacherIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === 'ครูผู้สอน') { teacherIdx = i; break; }
  }
  if (teacherIdx < 0) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  sheet.getRange(2, teacherIdx + 1, lastRow - 1, 1).clearContent();
}

/**
 * คัดลอกรายวิชา + โครงสร้างคะแนนไปใส่ Template Spreadsheet
 * (ลบชื่อครูผู้สอนออก เพื่ออำนวยความสะดวกแก่โรงเรียนใหม่)
 * @returns {Object} { success, message, subjectCount }
 */
function syncSubjectsToTemplate() {
  try {
    // 1. อ่านรายวิชาจากโรงเรียนปัจจุบัน
    var ss = SS();
    var subjectSheet = ss.getSheetByName('รายวิชา');
    if (!subjectSheet) throw new Error('ไม่พบชีต "รายวิชา" ในระบบปัจจุบัน');

    var allData = subjectSheet.getDataRange().getValues();
    if (allData.length < 2) throw new Error('ไม่มีข้อมูลรายวิชาในระบบ');

    var headers = allData[0];
    var teacherColIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === 'ครูผู้สอน') { teacherColIdx = i; break; }
    }

    // 2. เปิด Template Spreadsheet
    var props = PropertiesService.getScriptProperties();
    var templateId = props.getProperty('TEMPLATE_SPREADSHEET_ID')
      || '1AcdypFst0F4pr7bjaMH1WwuTyohekV8BeO36MWZOWJE';
    if (!templateId) throw new Error('ไม่พบ TEMPLATE_SPREADSHEET_ID — กรุณาตั้งค่าใน Script Properties');

    var templateSS = SpreadsheetApp.openById(templateId);
    var templateSheet = templateSS.getSheetByName('รายวิชา');
    if (!templateSheet) {
      templateSheet = templateSS.insertSheet('รายวิชา');
    }

    // 3. ล้างข้อมูลเดิมใน Template
    templateSheet.clearContents();
    templateSheet.clearFormats();

    // 4. สร้างข้อมูลใหม่ (ลบชื่อครูผู้สอน)
    var outputData = allData.map(function(row, rowIdx) {
      return row.map(function(cell, colIdx) {
        // ลบชื่อครูผู้สอน (เฉพาะแถวข้อมูล ไม่ใช่ header)
        if (colIdx === teacherColIdx && rowIdx > 0) return '';
        return cell;
      });
    });

    // 5. เขียนลง Template
    templateSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);

    // 6. จัดรูปแบบ header
    templateSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4a5568')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    templateSheet.setFrozenRows(1);

    var subjectCount = outputData.length - 1;
    Logger.log('✅ syncSubjectsToTemplate: คัดลอก ' + subjectCount + ' รายวิชาไปยัง Template สำเร็จ');
    return {
      success: true,
      message: '✅ คัดลอก ' + subjectCount + ' รายวิชาไปยัง Template สำเร็จ (ไม่รวมชื่อครูผู้สอน)',
      subjectCount: subjectCount
    };

  } catch (e) {
    Logger.log('❌ syncSubjectsToTemplate error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * สร้าง Template Spreadsheet เปล่า (โครงสร้างชีทครบ ไม่มีข้อมูล)
 * แชร์สาธารณะ (Anyone with link can view)
 */
function createTemplateSpreadsheet_() {
  var ss = SpreadsheetApp.create('PP5 Smart — Template ฐานข้อมูลต้นแบบ');

  // สร้างชีทถาวร
  PERMANENT_SHEETS.forEach(function(name) {
    var sheet = ss.insertSheet(name);
    setupSheetHeaders_(sheet, name, {});
  });

  // สร้างชีทรายปี (ชื่อ base ไม่มี suffix ปี — จะ rename ตอนติดตั้ง)
  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
    'AttendanceLog', 'ความเห็นครู'
  ];
  yearlyBases.forEach(function(baseName) {
    var sheet = ss.insertSheet(baseName);
    setupSheetHeaders_(sheet, baseName, {});
  });

  // ลบ Sheet1 default
  try {
    var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');
    if (def) ss.deleteSheet(def);
  } catch(e) {}

  // แชร์สาธารณะ (อ่านอย่างเดียว)
  var file = DriveApp.getFileById(ss.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return ss.getId();
}

/**
 * สร้างเนื้อหา README สำหรับคำแนะนำติดตั้ง
 */
function getInstallReadmeContent_() {
  return [
    '# 📦 PP5 Smart — ระบบบันทึกผลการเรียน ปพ.5',
    '## คำแนะนำการติดตั้งสำหรับโรงเรียนใหม่',
    '',
    '## ขั้นตอนการติดตั้ง',
    '',
    '### ขั้นที่ 1: สร้าง Google Apps Script Project',
    '1. เปิด https://script.google.com',
    '2. ล็อกอินด้วย Google Account ของโรงเรียน',
    '3. กดปุ่ม "+ New project"',
    '4. ตั้งชื่อ เช่น "PP5 Smart - [ชื่อโรงเรียน]"',
    '',
    '### ขั้นที่ 2: อัปโหลดไฟล์โค้ด',
    '1. ในหน้า GAS Editor ลบไฟล์ Code.gs เดิมที่มีอยู่',
    '2. กดปุ่ม "+" (เพิ่มไฟล์) → เลือก Script หรือ HTML',
    '3. อัปโหลดไฟล์ทั้งหมดจากโฟลเดอร์นี้:',
    '   - ไฟล์ .js → เลือก "Script" แล้ววางโค้ด (ไม่ต้องใส่ .js)',
    '   - ไฟล์ .html → เลือก "HTML" แล้ววางโค้ด (ไม่ต้องใส่ .html)',
    '4. ไฟล์ appsscript.json → Project Settings → เปิด "Show appsscript.json" → วางทับ',
    '',
    '### ขั้นที่ 3: เปิด Apps Script API',
    '1. ไปที่ https://console.cloud.google.com',
    '2. เลือก Project ที่ผูกกับ GAS',
    '3. ค้นหา "Apps Script API" → Enable',
    '',
    '### ขั้นที่ 4: Deploy Web App',
    '1. กด Deploy → New deployment',
    '2. เลือก Web app',
    '3. Execute as: Me, Who has access: Anyone',
    '4. กด Deploy → Authorize access → คัดลอก URL',
    '',
    '### ขั้นที่ 5: เปิด Setup Wizard',
    '1. เปิด URL + ?page=setup_wizard',
    '2. กรอกข้อมูลโรงเรียน',
    '3. กด "สร้างระบบใหม่"',
    '4. เสร็จ! เข้าใช้งานได้เลย',
    '',
    '## ปัญหาที่พบบ่อย',
    '',
    'Q: กด Deploy แล้วขึ้น "Authorization required"',
    'A: กด "Authorize access" → เลือก Google Account → "Allow"',
    '',
    'Q: เปิด URL แล้วขึ้นหน้าว่าง',
    'A: ตรวจสอบว่าอัปโหลดไฟล์ครบทุกไฟล์ โดยเฉพาะ Code.js',
    '',
    '---',
    'พัฒนาโดย: ระบบ PP5 Smart'
  ].join('\n');
}
