// ============================================================
// ⚙️ SETTINGS_UNIFIED.GS - Unified Settings Management
// ============================================================
// ไฟล์นี้รวม globalsettings.gs และ settings.gs เข้าด้วยกัน
// เพื่อลดความซ้ำซ้อนและจัดการการตั้งค่าแบบรวมศูนย์
// ============================================================

// ============================================================
// 📋 CONSTANTS
// ============================================================

const S_SETTINGS_SHEET_NAME = 'global_settings';
const S_SETTINGS_CACHE_KEY = 'system_settings_v1';
const S_SETTINGS_CACHE_TTL = 21600; // 6 ชั่วโมง

const S_Log = {
  info: (msg) => Logger.log(`[INFO] ${msg}`),
  success: (msg) => Logger.log(`[OK] ${msg}`),
  warning: (msg) => Logger.log(`[WARN] ${msg}`),
  error: (msg) => Logger.log(`[ERROR] ${msg}`)
};

function S_getCurrentAcademicYear_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  return month >= 5 ? year + 543 : year + 542;
}

function S_SS_() {
  if (getSpreadsheetId_()) {
    return SS();
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================
// 🔧 MAIN SETTINGS FUNCTIONS
// ============================================================

/**
 * ดึงการตั้งค่าระบบทั้งหมด (ฟังก์ชันหลัก)
 * รองรับทั้ง Cache และ Logo URL generation
 * @param {boolean} useCache - ใช้ Cache หรือไม่ (default: true)
 * @returns {Object} การตั้งค่าทั้งหมด
 */
function S_getGlobalSettings(useCache = true) {
  try {
    // ลอง Cache ก่อน
    if (useCache) {
      var cached = null;
      try {
        if (typeof U_CacheManager !== 'undefined' && U_CacheManager && U_CacheManager.get) {
          cached = U_CacheManager.get(S_SETTINGS_CACHE_KEY);
        } else {
          var _cache = CacheService.getScriptCache();
          var _raw = _cache.get(S_SETTINGS_CACHE_KEY);
          if (_raw) cached = JSON.parse(_raw);
        }
      } catch(ce) { Logger.log('Cache get warning: ' + ce.message); }
      if (cached) {
        const hasRealSchoolName = cached['ชื่อโรงเรียน'] && String(cached['ชื่อโรงเรียน']).trim() && String(cached['ชื่อโรงเรียน']).trim() !== 'โรงเรียนตัวอย่าง';
        const isDefaultCache = cached.__meta && cached.__meta.source === 'default';
        if (!hasRealSchoolName || isDefaultCache) {
          S_Log.warning('Cached settings looks invalid/default, ignoring cache and reloading from sheet');
        } else {
          S_Log.info('Settings loaded from cache');
          if (!cached.__meta) {
            cached.__meta = { source: 'cache', at: new Date().toISOString() };
          } else {
            cached.__meta.source = cached.__meta.source || 'cache';
            cached.__meta.at = new Date().toISOString();
          }
          return cached;
        }
      }
    }
    
    S_Log.info('Loading settings from sheet...');
    
    // อ่านจากชีต
    const settings = S_readSettingsFromSheet();
    
    // ประมวลผล Logo URLs
    if (settings.logoFileId) {
      S_processLogoUrls(settings);
    }
    
    // เพิ่มค่า default
    S_addDefaultSettings(settings);

    settings.__meta = Object.assign({}, settings.__meta || {}, {
      source: 'sheet',
      at: new Date().toISOString(),
      cacheEnabled: Boolean(useCache)
    });
    
    // บันทึกลง Cache
    if (useCache) {
      try {
        if (typeof U_CacheManager !== 'undefined' && U_CacheManager && U_CacheManager.set) {
          U_CacheManager.set(S_SETTINGS_CACHE_KEY, settings, S_SETTINGS_CACHE_TTL);
        } else {
          var _cache2 = CacheService.getScriptCache();
          _cache2.put(S_SETTINGS_CACHE_KEY, JSON.stringify(settings), S_SETTINGS_CACHE_TTL);
        }
      } catch(ce2) { Logger.log('Cache set warning: ' + ce2.message); }
    }
    
    return settings;
    
  } catch (e) {
    S_Log.error(`getGlobalSettings error: ${e.message}`);
    const defaults = S_getDefaultSettings();
    defaults.__meta = {
      source: 'default',
      at: new Date().toISOString(),
      error: String(e && e.message ? e.message : e)
    };
    return defaults;
  }
}

/**
 * Alias สำหรับความเข้ากันได้
 */
function S_getWebAppSettings() {
  return S_getGlobalSettings();
}

/**
 * Alias สำหรับความเข้ากันได้
 */
function S_getWebAppSettingsWithCache() {
  return S_getGlobalSettings(true);
}

// ============================================================
// 📖 READ SETTINGS FROM SHEET
// ============================================================

/**
 * อ่านการตั้งค่าจากชีต
 * @returns {Object} การตั้งค่า
 */
function S_readSettingsFromSheet() {
  const ss = S_SS_();
  let sheet = ss.getSheetByName(S_SETTINGS_SHEET_NAME);
  
  // ถ้าไม่มีชีต ให้สร้างใหม่
  if (!sheet) {
    S_Log.warning(`Sheet ${S_SETTINGS_SHEET_NAME} not found, creating...`);
    sheet = S_createSettingsSheet(ss);
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    S_Log.warning('Settings sheet is empty');
    const defaults = S_getDefaultSettings();
    defaults.__meta = {
      source: 'default',
      at: new Date().toISOString(),
      reason: 'empty_sheet',
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheetName: sheet.getName(),
      rowCount: data.length
    };
    return defaults;
  }
  
  const settings = {};
  
  // อ่านข้อมูล (ข้ามแถวแรกที่เป็น header)
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const value = data[i][1];
    
    if (key) {
      settings[key] = value !== null && value !== undefined ? String(value).trim() : '';
    }
  }

  settings.__meta = {
    source: 'sheet',
    at: new Date().toISOString(),
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName(),
    rowCount: data.length,
    keyCount: Object.keys(settings).length - 1
  };
  
  return settings;
}

/**
 * สร้างชีตการตั้งค่าใหม่
 * @param {Spreadsheet} ss - Spreadsheet object
 * @returns {Sheet} ชีตที่สร้างใหม่
 */
function S_createSettingsSheet(ss) {
  const sheet = ss.insertSheet(S_SETTINGS_SHEET_NAME);
  
  // ตั้งค่า header
  sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
  sheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
  
  // เพิ่มข้อมูลเริ่มต้น
  const defaults = S_getDefaultSettings();
  const rows = Object.entries(defaults).map(([key, value]) => [key, value]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  
  sheet.setFrozenRows(1);
  
  S_Log.success('Created settings sheet with default values');
  return sheet;
}

// ============================================================
// 🖼️ LOGO PROCESSING
// ============================================================

/**
 * ประมวลผล Logo URLs จาก logoFileId
 * @param {Object} settings - Settings object (จะถูก modify)
 */
function S_processLogoUrls(settings) {
  const fileId = String(settings.logoFileId || '').trim();
  
  if (!fileId || fileId.length < 10) {
    return;
  }
  
  try {
    // ตรวจสอบและแชร์ไฟล์
    try {
      const file = DriveApp.getFileById(fileId);
      const access = file.getSharingAccess();
      if (access !== DriveApp.Access.ANYONE_WITH_LINK && access !== DriveApp.Access.ANYONE) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        S_Log.success('Logo file shared successfully');
      }
    } catch (_) { /* DriveApp unavailable in web app */ }
    
    // สร้าง URLs หลายรูปแบบ
    settings.logo = `https://lh3.googleusercontent.com/d/${fileId}`;
    settings.logoDataUrl = settings.logo;
    settings.schoolLogo = settings.logo;
    settings.loginLogo = settings.logo;
    settings.logoUrl_lh3 = settings.logo;
    settings.logoUrl_uc = `https://drive.google.com/uc?export=view&id=${fileId}`;
    settings.logoUrl_backup1 = `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`;
    settings.logoUrl_backup2 = `https://drive.google.com/file/d/${fileId}/view`;
    
    settings.logoMethod = 'Simple URL (AUTO)';
    settings.logoLastFixed = new Date().toISOString();
    
    S_Log.success('Logo URLs generated successfully');
    
  } catch (e) {
    S_Log.warning(`Logo processing failed: ${e.message}`);
  }
}

// ============================================================
// 🔧 DEFAULT SETTINGS
// ============================================================

/**
 * เพิ่มค่า default ที่ขาดหายไป
 * @param {Object} settings - Settings object (จะถูก modify)
 */
function S_addDefaultSettings(settings) {
  const defaults = S_getDefaultSettings();
  
  Object.keys(defaults).forEach(key => {
    if (!settings[key]) {
      settings[key] = defaults[key];
    }
  });
}

/**
 * ดึงค่า default ทั้งหมด
 * @returns {Object} ค่า default
 */
function S_getDefaultSettings() {
  const currentYear = S_getCurrentAcademicYear_();
  
  return {
    'ชื่อโรงเรียน': 'โรงเรียนตัวอย่าง',
    'schoolName': 'โรงเรียนตัวอย่าง',
    'ที่อยู่โรงเรียน': '',
    'ปีการศึกษา': String(currentYear),
    'academicYear': String(currentYear),
    'ภาคเรียน': '1',
    'semester': '1',
    'ชื่อผู้อำนวยการ': 'ผู้อำนวยการ',
    'directorName': 'ผู้อำนวยการ',
    'ตำแหน่งผู้อำนวยการ': 'ผู้อำนวยการสถานศึกษา',
    'directorTitle': 'ผู้อำนวยการสถานศึกษา',
    'ชื่อหัวหน้างานวิชาการ': 'หัวหน้างานวิชาการ',
    'academicHead': 'หัวหน้างานวิชาการ',
    'logoFileId': '',
    'pdfSaveFolderId': '',
    'ชื่อครูผู้สอน': '',
    'หมายเหตุท้ายรายงาน': '',
    'footerNote': ''
  };
}

// ============================================================
// 💾 SAVE SETTINGS
// ============================================================

/**
 * บันทึกการตั้งค่า
 * @param {Object} settingsData - ข้อมูลการตั้งค่าที่จะบันทึก
 * @returns {string} ข้อความสถานะ
 */
function S_saveSystemSettings(settingsData) {
  try {
    S_Log.info('Saving system settings...');
    
    if (!settingsData || typeof settingsData !== 'object') {
      throw new Error('ข้อมูลการตั้งค่าไม่ถูกต้อง');
    }
    
    const ss = S_SS_();
    let sheet = ss.getSheetByName(S_SETTINGS_SHEET_NAME);
    
    if (!sheet) {
      sheet = S_createSettingsSheet(ss);
    }
    
    // อ่านข้อมูลเดิม
    const existingData = sheet.getDataRange().getValues();
    const existingKeys = {};
    
    for (let i = 1; i < existingData.length; i++) {
      const key = String(existingData[i][0] || '').trim();
      if (key) {
        existingKeys[key] = i + 1; // row number (1-based)
      }
    }
    
    // ประมวลผล logoFileId ก่อนบันทึก
    if (settingsData.logoFileId) {
      const fileId = String(settingsData.logoFileId).trim();
      
      if (fileId.length > 10) {
        // สร้าง URLs
        settingsData.logo = `https://lh3.googleusercontent.com/d/${fileId}`;
        settingsData.logoDataUrl = settingsData.logo;
        settingsData.schoolLogo = settingsData.logo;
        settingsData.loginLogo = settingsData.logo;
        settingsData.logoUrl_lh3 = settingsData.logo;
        settingsData.logoUrl_uc = `https://drive.google.com/uc?export=view&id=${fileId}`;
        settingsData.logoMethod = 'Simple URL (AUTO)';
        settingsData.logoLastFixed = new Date().toISOString();
        
        // แชร์ไฟล์
        try {
          const file = DriveApp.getFileById(fileId);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          S_Log.success('Logo file shared');
        } catch (e) {
          S_Log.warning(`Cannot share logo file: ${e.message}`);
        }
      }
    }
    
    // บันทึกข้อมูล
    let updateCount = 0;
    let insertCount = 0;
    
    Object.keys(settingsData).forEach(key => {
      const value = settingsData[key];
      const trimmedKey = String(key).trim();
      
      if (existingKeys[trimmedKey]) {
        // อัปเดตแถวที่มีอยู่
        const rowNum = existingKeys[trimmedKey];
        sheet.getRange(rowNum, 2).setValue(value);
        updateCount++;
      } else {
        // เพิ่มแถวใหม่
        const newRowNum = sheet.getLastRow() + 1;
        sheet.getRange(newRowNum, 1, 1, 2).setValues([[trimmedKey, value]]);
        insertCount++;
      }
    });
    
    S_Log.success(`Settings saved: ${updateCount} updated, ${insertCount} inserted`);
    
    // ล้าง Cache
    S_clearSettingsCache();
    
    return 'บันทึกการตั้งค่าสำเร็จ';
    
  } catch (error) {
    S_Log.error(`saveSystemSettings error: ${error.message}`);
    throw new Error(`ไม่สามารถบันทึกการตั้งค่าได้: ${error.message}`);
  }
}

/**
 * Alias สำหรับความเข้ากันได้
 */
function S_saveGlobalSettings(settingsData) {
  return S_saveSystemSettings(settingsData);
}

// ============================================================
// 🗑️ CACHE MANAGEMENT
// ============================================================

/**
 * ล้าง Cache การตั้งค่า
 * @returns {string} ข้อความสถานะ
 */
function S_clearSettingsCache() {
  try {
    const keys = [
      S_SETTINGS_CACHE_KEY,
      'settings_webapp_v3',
      'settings_webapp_v2',
      'settings_webapp',
      'global_settings'
    ];
    
    if (typeof U_CacheManager !== 'undefined' && U_CacheManager && U_CacheManager.removeAll) {
      U_CacheManager.removeAll(keys);
    } else {
      // Fallback: ใช้ CacheService โดยตรง
      var cache = CacheService.getScriptCache();
      cache.removeAll(keys);
    }
    S_Log.success('Settings cache cleared');
    
    return 'ล้าง Settings Cache สำเร็จ';
    
  } catch (e) {
    // ไม่ throw error เพราะการล้าง cache ไม่ควรทำให้การบันทึกล้มเหลว
    Logger.log('clearSettingsCache warning: ' + e.message);
    return 'บันทึกสำเร็จ (cache warning)';
  }
}

// ============================================================
// 🔍 SPECIFIC SETTINGS GETTERS
// ============================================================

/**
 * ดึงชื่อโรงเรียน
 * @returns {string}
 */
function S_getSchoolName() {
  const settings = S_getGlobalSettings();
  return settings['ชื่อโรงเรียน'] || settings['schoolName'] || 'โรงเรียน';
}

/**
 * ดึงปีการศึกษาปัจจุบัน
 * @returns {string}
 */
function S_getAcademicYear() {
  const settings = S_getGlobalSettings();
  return settings['ปีการศึกษา'] || settings['academicYear'] || String(S_getCurrentAcademicYear_());
}

/**
 * ดึงภาคเรียนปัจจุบัน
 * @returns {string}
 */
function S_getSemester() {
  const settings = S_getGlobalSettings();
  return settings['ภาคเรียน'] || settings['semester'] || '1';
}

// ============================================================
// 📋 YEARLY SHEET NAME HELPERS
// ============================================================

/** ชีตข้อมูลรายปี — ชื่อฐาน (ปีแรกที่ยังไม่มี suffix จะใช้ชื่อเดิม) */
var S_YEARLY_SHEETS = [
  'SCORES_WAREHOUSE',
  'การประเมินอ่านคิดเขียน',
  'การประเมินคุณลักษณะ',
  'การประเมินกิจกรรมพัฒนาผู้เรียน',
  'การประเมินสมรรถนะ',
  'AttendanceLog',
  'ความเห็นครู'
];

/** ชีตที่ใช้ร่วมกันทุกปี — จะถูก snapshot เป็น ชื่อ_ปี เมื่อเปลี่ยนปีการศึกษา */
var S_SHARED_SHEETS = ['Students', 'รายวิชา', 'HomeroomTeachers'];

/**
 * ดึง Sheet object ของชีตที่ใช้ร่วม (Students/รายวิชา/HomeroomTeachers)
 * - ถ้าระบุ year และมี snapshot (เช่น Students_2568) → คืน snapshot
 * - ถ้าไม่ระบุ year หรือ year = ปีปัจจุบัน → คืนชีต live
 * - fallback: ชื่อเดิมไม่มี suffix
 * @param {string} baseName - ชื่อฐานชีต เช่น 'Students'
 * @param {string} [year] - ปีการศึกษา (optional, ถ้าไม่ระบุ = ปีปัจจุบัน)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function S_getSharedSheet(baseName, year) {
  var ss = SS();
  var currentYear = S_getAcademicYear();
  var requestedYear = year ? String(year) : String(currentYear);

  // ถ้าขอปีปัจจุบัน → ใช้ชีต live
  if (requestedYear === String(currentYear)) {
    return ss.getSheetByName(baseName) || null;
  }

  // ถ้าขอปีเก่า → ลอง snapshot ก่อน แล้ว fallback ไปชีต live
  var snapshotName = baseName + '_' + requestedYear;
  return ss.getSheetByName(snapshotName) || ss.getSheetByName(baseName) || null;
}

/**
 * ดึงรายการปีการศึกษาที่มีข้อมูลในระบบ (จากชีต snapshot + ชีตรายปี)
 * @returns {string[]} เช่น ['2567', '2568', '2569']
 */
function S_getAvailableYears() {
  var ss = SS();
  var years = {};
  var currentYear = S_getAcademicYear();
  years[String(currentYear)] = true;

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    var match = name.match(/_(\d{4})$/);
    if (match) years[match[1]] = true;
  });

  return Object.keys(years).sort();
}

/**
 * สร้างชื่อชีตสำหรับปีที่ระบุ
 * เช่น S_sheetName('SCORES_WAREHOUSE', '2568') → 'SCORES_WAREHOUSE_2568'
 *      S_sheetName('SCORES_WAREHOUSE')         → ใช้ปีปัจจุบันจาก global_settings
 * @param {string} baseName - ชื่อฐานชีต
 * @param {string} [year] - ปีการศึกษา (ถ้าไม่ระบุจะดึงจาก settings)
 * @returns {string}
 */
function S_sheetName(baseName, year) {
  var y = year || S_getAcademicYear();
  return baseName + '_' + y;
}

/**
 * ดึง Sheet object สำหรับปีปัจจุบัน พร้อม fallback ไปชื่อเดิม (กรณียังไม่เคยเปลี่ยนปี)
 * @param {string} baseName - ชื่อฐานชีต เช่น 'SCORES_WAREHOUSE'
 * @param {string} [year] - ปีการศึกษา (optional)
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function S_getYearlySheet(baseName, year) {
  var ss = SS();
  var name = S_sheetName(baseName, year);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  // fallback: ลองชื่อเดิมไม่มี suffix (กรณีระบบยังไม่เคยเปลี่ยนปี)
  return ss.getSheetByName(baseName);
}

/**
 * ดึงชื่อชีตสำหรับปีที่ระบุ พร้อม fallback
 * @param {string} baseName
 * @param {string} [year]
 * @returns {string} ชื่อชีตที่มีอยู่จริง
 */
function S_resolveSheetName(baseName, year) {
  var ss = SS();
  var name = S_sheetName(baseName, year);
  if (ss.getSheetByName(name)) return name;
  if (ss.getSheetByName(baseName)) return baseName;
  return name; // ยังไม่มีชีต — คืนชื่อใหม่ไว้สร้างทีหลัง
}

/**
 * ดึง Logo URL
 * @returns {string}
 */
function S_getLogoUrl() {
  const settings = S_getGlobalSettings();
  return settings.logo || settings.logoDataUrl || settings.schoolLogo || '';
}

/**
 * ดึงชื่อผู้อำนวยการ
 * @returns {string}
 */
function S_getDirectorName() {
  const settings = S_getGlobalSettings();
  return settings['ชื่อผู้อำนวยการ'] || settings['directorName'] || 'ผู้อำนวยการ';
}

// ============================================================
// 🛠️ UTILITY FUNCTIONS
// ============================================================

/**
 * ตรวจสอบว่าการตั้งค่าครบถ้วนหรือไม่
 * @returns {Object} { complete: boolean, missing: Array }
 */
function S_validateSettings() {
  const settings = S_getGlobalSettings();
  const required = ['ชื่อโรงเรียน', 'ปีการศึกษา', 'ชื่อผู้อำนวยการ'];
  const missing = [];
  
  required.forEach(key => {
    if (!settings[key] || settings[key].trim() === '') {
      missing.push(key);
    }
  });
  
  return {
    complete: missing.length === 0,
    missing: missing
  };
}

/**
 * รีเซ็ตการตั้งค่าเป็นค่า default
 * @returns {string}
 */
function S_resetSettings() {
  try {
    const defaults = S_getDefaultSettings();
    S_saveSystemSettings(defaults);
    S_Log.success('Settings reset to defaults');
    return 'รีเซ็ตการตั้งค่าเรียบร้อยแล้ว';
  } catch (e) {
    S_Log.error(`resetSettings error: ${e.message}`);
    throw new Error(`ไม่สามารถรีเซ็ตการตั้งค่าได้: ${e.message}`);
  }
}

// ============================================================
// 🧪 TESTING FUNCTIONS
// ============================================================

/**
 * แก้ไขการแชร์ไฟล์โลโก้
 */
function fixLogoSharing_S() {
  try {
    const settings = S_getGlobalSettings();
    const logoFileId = settings.logoFileId;
    
    if (!logoFileId) {
      throw new Error('ไม่มี logoFileId ในการตั้งค่า');
    }
    
    try {
      const file = DriveApp.getFileById(logoFileId);
      S_Log.info(`Found file: ${file.getName()}`);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {
      // REST API fallback for sharing
      var _tk = ScriptApp.getOAuthToken();
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + logoFileId + '/permissions', {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + _tk },
        payload: JSON.stringify({ role: 'reader', type: 'anyone' }), muteHttpExceptions: true
      });
    }
    S_Log.success('Logo file sharing updated to PUBLIC');
    
    // ล้าง Cache เพื่อให้โหลดการตั้งค่าใหม่
    S_clearSettingsCache();
    
    return 'แก้ไขการแชร์ไฟล์โลโก้สำเร็จ';
    
  } catch (error) {
    S_Log.error(`fixLogoSharing error: ${error.message}`);
    throw new Error(`ไม่สามารถแก้ไขการแชร์ไฟล์ได้: ${error.message}`);
  }
}
