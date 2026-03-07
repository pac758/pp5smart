// ============================================================
// 🏫 INSTALLER.GS — Multi-Tenant Setup for New Schools
// ============================================================
// ฟังก์ชันนี้ใช้สำหรับติดตั้งระบบครั้งแรกเมื่อโรงเรียนใหม่
// copy Apps Script template แล้วเปิดใช้งาน
// ============================================================

const INSTALLER_VERSION = '1.0.0';

// Sheet names ที่ต้องสร้างใน Spreadsheet ใหม่
const REQUIRED_SHEETS = [
  'global_settings',
  'Users',
  'Students',
  'รายวิชา',
  'SCORES_WAREHOUSE',
  'Holidays',
  'HomeroomTeachers',
  'การประเมินอ่านคิดเขียน',
  'การประเมินคุณลักษณะ',
  'การประเมินกิจกรรมพัฒนาผู้เรียน',
  'การประเมินสมรรถนะ',
  'AttendanceLog',
  'ความเห็นครู'
];

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
    
    // ตรวจสอบชีทที่จำเป็นอื่นๆ
    const requiredSheets = ['Students', 'Users', 'รายวิชา', 'SCORES_WAREHOUSE'];
    const existingSheets = ss.getSheets().map(s => s.getName());
    const missing = requiredSheets.filter(name => !existingSheets.includes(name));
    
    if (missing.length > 0) {
      return { installed: false, message: `ขาดชีทที่จำเป็น: ${missing.join(', ')}` };
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
    // ตรวจสอบว่าติดตั้งแล้วหรือยัง
    const status = getSetupStatus();
    if (status.installed) {
      return { success: false, message: 'ระบบถูกติดตั้งแล้ว ไม่สามารถติดตั้งซ้ำได้', alreadyInstalled: true };
    }

    // Validate ข้อมูลที่รับมา
    const validation = validateSetupForm_(formData);
    if (!validation.valid) {
      return { success: false, message: validation.message };
    }

    Logger.log('🏫 เริ่มติดตั้งระบบสำหรับ: ' + formData.schoolName);

    // 1. สร้าง Spreadsheet ใหม่
    const ss = createSchoolSpreadsheet_(formData);
    Logger.log('✅ สร้าง Spreadsheet แล้ว: ' + ss.getId());

    // 2. สร้าง Sheet ทั้งหมด
    setupSheets_(ss, formData);
    Logger.log('✅ สร้าง Sheets แล้ว');

    // 3. บันทึก Settings ลง global_settings sheet
    saveGlobalSettings_(ss, formData);
    Logger.log('✅ บันทึก Settings แล้ว');

    // 4. สร้าง Admin User
    createAdminUser_(ss, formData);
    Logger.log('✅ สร้าง Admin User แล้ว');

    // 5. บันทึก ScriptProperties
    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      'SPREADSHEET_ID': ss.getId(),
      'SCHOOL_NAME': formData.schoolName,
      'SETUP_DATE': new Date().toISOString(),
      'INSTALLER_VERSION': INSTALLER_VERSION
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
  const name = 'ระบบ ปพ.5 — ' + formData.schoolName + ' (' + formData.academicYear + ')';
  const ss = SpreadsheetApp.create(name);

  // ย้ายเข้า Drive folder ถ้ามี
  try {
    const file = DriveApp.getFileById(ss.getId());
    const folderName = 'ระบบโรงเรียน_' + formData.schoolName;
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
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
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');

  // สร้าง Sheets ทั้งหมด
  REQUIRED_SHEETS.forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    setupSheetHeaders_(sheet, name, formData);
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
    'SCORES_WAREHOUSE': [['studentId','subjectCode','subjectName','grade','class_no','semester','academicYear','score1','score2','score3','midterm','final','total','gradeResult','teacherUsername','updatedAt']],
    'HomeroomTeachers': [['grade','classNo','teacherName']],
    'การประเมินอ่านคิดเขียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา','สุขศึกษา','ศิลปะ','การงาน','ภาษาอังกฤษ','สรุปผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินคุณลักษณะ': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','รักชาติ_ศาสน์_กษัตริย์','ซื่อสัตย์สุจริต','มีวินัย','ใฝ่เรียนรู้','อยู่อย่างพอเพียง','มุ่งมั่นในการทำงาน','รักความเป็นไทย','มีจิตสาธารณะ','คะแนนรวม','คะแนนเฉลี่ย','ผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินกิจกรรมพัฒนาผู้เรียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','กิจกรรมแนะแนว','ลูกเสือ_เนตรนารี','ชุมนุม','เพื่อสังคมและสาธารณประโยชน์','รวมกิจกรรม','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินสมรรถนะ': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง']],
    'AttendanceLog': [['timestamp','student_id','grade','classNo','date','status','recorder']],
    'ความเห็นครู': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','ความเห็น','วันที่บันทึก','ผู้บันทึก']]
  };

  const headers = headerMap[sheetName];
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
    ['installer_version', INSTALLER_VERSION, now]
  ];

  sheet.getRange(2, 1, settings.length, 3).setValues(settings);
}

function createAdminUser_(ss, formData) {
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return;

  const now = new Date().toISOString();
  sheet.appendRow([
    formData.adminUsername.trim(),
    formData.adminPassword,
    'admin',
    formData.adminFirstName || 'ผู้ดูแล',
    formData.adminLastName || 'ระบบ',
    formData.adminEmail || '',
    '',
    'true',
    now
  ]);
}

// ============================================================
// � MIGRATE: สำหรับโรงเรียนที่มีอยู่แล้ว (รันครั้งเดียว)
// ============================================================

function migrateExistingSchool() {
  const EXISTING_SPREADSHEET_ID = '1fqfu-fUVKjELG-trQzOw7b8MEUA-61YQa0YxxohSkeA';
  const EXISTING_SCHOOL_NAME = 'โรงเรียนบ้านโคกยางหนองถนน';
  const EXISTING_SCHOOL_ADDRESS = 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาบุรีรัมย์ เขต 3';
  const EXISTING_DIRECTOR_TITLE = 'ผู้อำนวยการโรงเรียนบ้านโคกยางหนองถนน';
  const EXISTING_DIRECTOR_NAME = 'นายสมศักดิ์ อิ่มภักดี';
  const EXISTING_ACADEMIC_YEAR = '2568';
  const EXISTING_SEMESTER = '1';
  const EXISTING_PDF_FOLDER_ID = '18J4qMJEHyAt2aGWw3A9uN4LbuXP59vTc';

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
// 📅 SWITCH ACADEMIC YEAR — เปลี่ยนปีการศึกษา
// ============================================================

/**
 * เปลี่ยนปีการศึกษา:
 * 1. Copy Spreadsheet ปัจจุบันเป็น archive
 * 2. ล้างข้อมูลรายปี (คะแนน, ประเมิน, เวลาเรียน, ความเห็นครู)
 * 3. ลบชีตคะแนนรายวิชา (dynamic sheets)
 * 4. อัปเดตปีการศึกษาใน global_settings
 * 
 * @param {number} newYear - ปีการศึกษาใหม่ (พ.ศ.)
 * @param {Object} options - ตัวเลือกเพิ่มเติม
 * @returns {Object} ผลลัพธ์
 */
function switchAcademicYear(newYear, options) {
  options = options || {};
  var keepStudents = options.keepStudents !== false; // default: เก็บนักเรียนไว้
  var keepSubjects = options.keepSubjects !== false; // default: เก็บรายวิชาไว้

  try {
    if (!newYear || isNaN(newYear)) {
      return { success: false, message: 'กรุณาระบุปีการศึกษาใหม่' };
    }
    newYear = Number(newYear);

    var ss = SS();
    var oldSettings = S_getGlobalSettings(false);
    var oldYear = oldSettings['ปีการศึกษา'] || 'ไม่ทราบ';
    var schoolName = oldSettings['ชื่อโรงเรียน'] || 'โรงเรียน';

    Logger.log('📅 เริ่มเปลี่ยนปีการศึกษา: ' + oldYear + ' → ' + newYear);

    // ============ ขั้นที่ 1: Copy Spreadsheet เป็น archive ============
    Logger.log('📦 กำลัง archive spreadsheet...');
    var archiveName = schoolName + ' — ข้อมูลปีการศึกษา ' + oldYear + ' (สำรอง ' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd-MM-yyyy HH:mm') + ')';
    var originalFile = DriveApp.getFileById(ss.getId());
    var parentFolders = originalFile.getParents();
    var targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

    // สร้างโฟลเดอร์ archive ถ้ายังไม่มี
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
    Logger.log('📎 URL: ' + archiveUrl);

    // ============ ขั้นที่ 2: ล้างข้อมูลรายปี ============
    Logger.log('🧹 กำลังล้างข้อมูลรายปี...');
    var sheetsToClean = [
      'SCORES_WAREHOUSE',
      'การประเมินอ่านคิดเขียน',
      'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน',
      'การประเมินสมรรถนะ',
      'AttendanceLog',
      'ความเห็นครู'
    ];

    var cleanedCount = 0;
    sheetsToClean.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
        cleanedCount++;
        Logger.log('  🗑️ ล้าง: ' + name);
      }
    });

    // ============ ขั้นที่ 3: ลบชีตคะแนนรายวิชา (dynamic) ============
    Logger.log('🗑️ กำลังลบชีตคะแนนรายวิชา...');
    var systemSheets = [
      'global_settings', 'Users', 'Students', 'รายวิชา',
      'SCORES_WAREHOUSE', 'Holidays', 'HomeroomTeachers',
      'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
      'AttendanceLog', 'ความเห็นครู'
    ];
    var prefixSkip = ['BACKUP_', 'Template_', 'TMP_'];

    var deletedSheets = 0;
    var allSheets = ss.getSheets();
    // ต้องเก็บอย่างน้อย 1 sheet ไว้
    for (var i = allSheets.length - 1; i >= 0; i--) {
      var sheetName = allSheets[i].getName();
      // ข้ามชีตระบบ
      if (systemSheets.indexOf(sheetName) !== -1) continue;
      // ข้ามชีตที่ขึ้นต้นด้วย prefix พิเศษ
      var skipThis = false;
      prefixSkip.forEach(function(p) { if (sheetName.indexOf(p) === 0) skipThis = true; });
      if (skipThis) continue;
      // ต้องเหลืออย่างน้อย 1 sheet
      if (ss.getSheets().length <= 1) break;
      try {
        ss.deleteSheet(allSheets[i]);
        deletedSheets++;
        Logger.log('  🗑️ ลบชีต: ' + sheetName);
      } catch (e) {
        Logger.log('  ⚠️ ลบชีตไม่ได้: ' + sheetName + ' — ' + e.message);
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

    // ============ ขั้นที่ 5: จัดการนักเรียน (เลื่อนชั้น + จำหน่าย ป.6) ============
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
      sheetsCleared: cleanedCount,
      scoreSheetDeleted: deletedSheets,
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
 * ดูข้อมูลก่อนเปลี่ยนปี (preview) ไม่ทำอะไรจริง
 * @returns {Object} สรุปข้อมูลที่จะถูกล้าง
 */
function previewSwitchAcademicYear() {
  try {
    var ss = SS();
    var settings = S_getGlobalSettings(false);
    var currentYear = settings['ปีการศึกษา'] || 'ไม่ทราบ';

    // นับข้อมูลในชีตที่จะล้าง
    var sheetsToClean = [
      'SCORES_WAREHOUSE',
      'การประเมินอ่านคิดเขียน',
      'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน',
      'การประเมินสมรรถนะ',
      'AttendanceLog',
      'ความเห็นครู'
    ];

    var dataCounts = [];
    sheetsToClean.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      var rows = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
      dataCounts.push({ sheet: name, rows: rows });
    });

    // นับชีตคะแนนรายวิชาที่จะลบ
    var systemSheets = [
      'global_settings', 'Users', 'Students', 'รายวิชา',
      'SCORES_WAREHOUSE', 'Holidays', 'HomeroomTeachers',
      'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
      'AttendanceLog', 'ความเห็นครู'
    ];
    var prefixSkip = ['BACKUP_', 'Template_', 'TMP_'];
    var scoreSheets = [];
    ss.getSheets().forEach(function(sheet) {
      var name = sheet.getName();
      if (systemSheets.indexOf(name) !== -1) return;
      var skip = false;
      prefixSkip.forEach(function(p) { if (name.indexOf(p) === 0) skip = true; });
      if (skip) return;
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
      suggestedNewYear: Number(currentYear) + 1 || '',
      studentCount: studentCount,
      gradeCounts: gradeCounts,
      p6Count: gradeCounts['ป.6'],
      promoteCount: studentCount - gradeCounts['ป.6'],
      dataCounts: dataCounts,
      scoreSheets: scoreSheets,
      totalDataRows: dataCounts.reduce(function(sum, d) { return sum + d.rows; }, 0)
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
