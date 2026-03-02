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
// �� RESET (ใช้เฉพาะ Admin เท่านั้น)
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
