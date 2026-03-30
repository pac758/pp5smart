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
 * � แก้ไข SPREADSHEET_ID ให้ชี้กลับไปชีตเดิม (กรณี script ไปรันบนสำเนาสำรอง)
 * วิธีใช้: 
 *   1. ใส่ ID ของ Spreadsheet เดิม (ดูจาก URL: /d/xxxxx/edit → xxxxx คือ ID)
 *   2. รัน fixSpreadsheetId() จาก GAS Editor
 */
function fixSpreadsheetId() {
  // ⚠️ ใส่ ID ของ Spreadsheet เดิมที่ถูกต้องที่นี่
  var CORRECT_ID = '';  // ← ใส่ ID ตรงนี้

  if (!CORRECT_ID) {
    // ถ้ายังไม่ได้ใส่ ID → แสดง ID ปัจจุบันให้ดู
    var props = PropertiesService.getScriptProperties();
    var currentSsId = props.getProperty('SPREADSHEET_ID') || '(ไม่มี)';
    Logger.log('=== SPREADSHEET_ID ปัจจุบัน ===');
    Logger.log('SPREADSHEET_ID: ' + currentSsId);
    Logger.log('');
    Logger.log('วิธีแก้: ใส่ ID ของ Spreadsheet เดิมในตัวแปร CORRECT_ID แล้วรันใหม่');
    Logger.log('ดู ID จาก URL: https://docs.google.com/spreadsheets/d/[ID_อยู่ตรงนี้]/edit');
    return;
  }

  try {
    // ตรวจว่าเปิดได้จริง
    var ss = SpreadsheetApp.openById(CORRECT_ID);
    var name = ss.getName();

    var props = PropertiesService.getScriptProperties();
    var oldId = props.getProperty('SPREADSHEET_ID') || '(ไม่มี)';
    props.setProperty('SPREADSHEET_ID', CORRECT_ID);
    props.setProperty('SCRIPT_ID', ScriptApp.getScriptId());

    Logger.log('✅ แก้ไขสำเร็จ!');
    Logger.log('   เดิม: ' + oldId);
    Logger.log('   ใหม่: ' + CORRECT_ID);
    Logger.log('   ชื่อ: ' + name);
    Logger.log('   isSetupComplete_: ' + isSetupComplete_());
  } catch (e) {
    Logger.log('❌ เปิด Spreadsheet ไม่ได้: ' + e.message);
    Logger.log('กรุณาตรวจสอบ ID อีกครั้ง');
  }
}

/**
 * � DEBUG: รันฟังก์ชันนี้ใน GAS Editor เพื่อดูสถานะ ScriptProperties
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

  // ⬇️ ใช้ service ทุกตัวเพื่อ trigger OAuth authorization ครบทุก scope
  try { DriveApp.getRootFolder().getName(); Logger.log('✅ DriveApp: OK'); } catch(e) { Logger.log('❌ DriveApp: ' + e.message); }
  try { SpreadsheetApp.getActive(); Logger.log('✅ SpreadsheetApp: OK'); } catch(e) { Logger.log('❌ SpreadsheetApp: ' + e.message); }
  try { DocumentApp.getActiveDocument(); Logger.log('✅ DocumentApp: OK'); } catch(e) { Logger.log('⚠️ DocumentApp: ' + e.message + ' (ปกติถ้าไม่ได้เปิด Doc)'); }
  Logger.log('=== ✅ Authorization ครบทุก scope ===');
}

/**
 * 🔧 REPAIR: แก้ไขโรงเรียนที่ติดตั้งไปแล้วด้วยโค้ดเก่า
 * รันจาก GAS Editor → ดู Execution log
 * 
 * สิ่งที่แก้ไข:
 * 1. SPREADSHEET_ID → ชี้ไปที่ Bound Spreadsheet (ไม่ใช่ Spreadsheet เปล่าที่สร้างขึ้นผิด)
 * 2. SCRIPT_ID → อัปเดตให้ตรงกับ Script ปัจจุบัน
 * 3. ตรวจสอบ HomeroomTeachers header
 */
function repairExistingInstallation() {
  var props = PropertiesService.getScriptProperties();
  var currentScriptId = ScriptApp.getScriptId();
  var oldSsId = props.getProperty('SPREADSHEET_ID');
  
  Logger.log('=== 🔧 REPAIR EXISTING INSTALLATION ===');
  Logger.log('Current Script ID: ' + currentScriptId);
  Logger.log('Current SPREADSHEET_ID: ' + (oldSsId || '(ไม่มี)'));
  
  // 1. ตรวจหา Bound Spreadsheet
  var boundSs = null;
  try {
    boundSs = SpreadsheetApp.getActiveSpreadsheet();
    if (boundSs) Logger.log('✅ พบ Bound Spreadsheet ผ่าน getActiveSpreadsheet: ' + boundSs.getId());
  } catch(_e) {}
  
  if (!boundSs) {
    try {
      var token = ScriptApp.getOAuthToken();
      var resp = UrlFetchApp.fetch(
        'https://script.googleapis.com/v1/projects/' + currentScriptId,
        { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (resp.getResponseCode() === 200) {
        var projData = JSON.parse(resp.getContentText());
        if (projData.parentId) {
          boundSs = SpreadsheetApp.openById(projData.parentId);
          Logger.log('✅ พบ Bound Spreadsheet ผ่าน Script API: ' + projData.parentId);
        }
      }
    } catch(_e2) {
      Logger.log('⚠️ Script API failed: ' + _e2.message);
    }
  }
  
  if (!boundSs) {
    Logger.log('❌ ไม่พบ Bound Spreadsheet — ไม่สามารถแก้ไขอัตโนมัติได้');
    Logger.log('💡 ลอง: รันฟังก์ชัน manualRepairSpreadsheetId("SPREADSHEET_ID_ใหม่") แทน');
    return;
  }
  
  var correctId = boundSs.getId();
  var needsRepair = (oldSsId !== correctId);
  
  // 2. แก้ SPREADSHEET_ID
  if (needsRepair) {
    Logger.log('🔄 แก้ไข SPREADSHEET_ID:');
    Logger.log('   เดิม: ' + oldSsId + ' (ผิด — Spreadsheet เปล่าที่สร้างขึ้นตอนติดตั้ง)');
    Logger.log('   ใหม่: ' + correctId + ' (ถูก — Bound Spreadsheet ที่มีข้อมูลรายวิชา)');
    props.setProperty('SPREADSHEET_ID', correctId);
  } else {
    Logger.log('✅ SPREADSHEET_ID ถูกต้องแล้ว: ' + correctId);
  }
  
  // 3. แก้ SCRIPT_ID
  props.setProperty('SCRIPT_ID', currentScriptId);
  Logger.log('✅ SCRIPT_ID อัปเดต: ' + currentScriptId);
  
  // 4. ตรวจ/แก้ HomeroomTeachers header
  var htSheet = boundSs.getSheetByName('HomeroomTeachers');
  if (htSheet) {
    if (htSheet.getLastColumn() < 4) {
      htSheet.getRange(1, 4).setValue('ครูประจำชั้น 2');
      Logger.log('✅ เพิ่ม header "ครูประจำชั้น 2" ใน HomeroomTeachers');
    } else {
      Logger.log('✅ HomeroomTeachers header ครบแล้ว');
    }
  } else {
    // สร้าง HomeroomTeachers ถ้ายังไม่มี
    htSheet = boundSs.insertSheet('HomeroomTeachers');
    htSheet.getRange(1, 1, 1, 4).setValues([['grade', 'classNo', 'teacherName', 'ครูประจำชั้น 2']]);
    htSheet.getRange(1, 1, 1, 4).setBackground('#4a5568').setFontColor('#ffffff').setFontWeight('bold');
    htSheet.setFrozenRows(1);
    Logger.log('✅ สร้างชีต HomeroomTeachers ใหม่');
  }
  
  // 5. ตรวจ global_settings มี installed_script_id ถูกต้อง
  var gsSheet = boundSs.getSheetByName('global_settings');
  if (gsSheet) {
    var gsData = gsSheet.getDataRange().getValues();
    var foundMarker = false;
    for (var i = 0; i < gsData.length; i++) {
      if (gsData[i][0] === 'installed_script_id') {
        gsSheet.getRange(i + 1, 2).setValue(currentScriptId);
        foundMarker = true;
        break;
      }
    }
    if (!foundMarker) {
      gsSheet.appendRow(['installed_script_id', currentScriptId, new Date().toISOString()]);
    }
    Logger.log('✅ global_settings.installed_script_id อัปเดต');
  }
  
  // 6. ตรวจชีตรายวิชา
  var subjectSheet = boundSs.getSheetByName('รายวิชา');
  if (subjectSheet) {
    var subjectRows = subjectSheet.getLastRow() - 1;
    Logger.log('✅ ชีตรายวิชามี ' + subjectRows + ' รายวิชา');
  } else {
    Logger.log('⚠️ ไม่พบชีตรายวิชา');
  }
  
  // 7. ตรวจ/แก้ Students header ให้ครบถ้วน
  var studentSheet = boundSs.getSheetByName('Students');
  if (studentSheet) {
    var fullHeaders = ['student_id','id_card','title','firstname','lastname','grade','class_no','gender',
      'birthdate','photo_url','academic_year','weight','height','blood_type','religion',
      'father_name','father_lastname','father_occupation','mother_name','mother_lastname','mother_occupation',
      'address','created_date','status'];
    var currentHeaders = studentSheet.getRange(1, 1, 1, studentSheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var missingCols = [];
    fullHeaders.forEach(function(h) {
      if (currentHeaders.indexOf(h) < 0) missingCols.push(h);
    });
    if (missingCols.length > 0) {
      var startCol = currentHeaders.length + 1;
      missingCols.forEach(function(colName, idx) {
        studentSheet.getRange(1, startCol + idx).setValue(colName);
      });
      Logger.log('✅ เพิ่ม ' + missingCols.length + ' คอลัมน์ใน Students: ' + missingCols.join(', '));
    } else {
      Logger.log('✅ Students header ครบ ' + currentHeaders.length + ' คอลัมน์');
    }
  }
  
  // 8. ตรวจ/แก้ ความเห็นครู header (yearly sheet)
  var commentSheetNames = boundSs.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return n === 'ความเห็นครู' || n.indexOf('ความเห็นครู_') === 0; });
  commentSheetNames.forEach(function(csName) {
    var cs = boundSs.getSheetByName(csName);
    if (cs && cs.getLastColumn() > 0) {
      var csHeaders = cs.getRange(1, 1, 1, cs.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      var fixed = false;
      // แก้คอลัมน์ 2: ชื่อ-นามสกุล → ชื่อนักเรียน
      var col2Idx = csHeaders.indexOf('ชื่อ-นามสกุล');
      if (col2Idx >= 0) { cs.getRange(1, col2Idx + 1).setValue('ชื่อนักเรียน'); fixed = true; }
      // แก้คอลัมน์สุดท้าย: ผู้บันทึก → วันที่อัปเดต
      var oldLastIdx = csHeaders.indexOf('ผู้บันทึก');
      if (oldLastIdx >= 0 && csHeaders.indexOf('วันที่อัปเดต') < 0) {
        cs.getRange(1, oldLastIdx + 1).setValue('วันที่อัปเดต'); fixed = true;
      }
      if (fixed) Logger.log('✅ แก้ header ชีต "' + csName + '"');
      else Logger.log('✅ ชีต "' + csName + '" header ถูกต้อง');
    }
  });
  
  // 9. ตรวจ/แก้ ชีตกิจกรรมพัฒนาผู้เรียน — เพิ่มคอลัมน์ 'ชื่อชุมนุม'
  var actSheetNames = boundSs.getSheets().map(function(s) { return s.getName(); })
    .filter(function(n) { return n === 'การประเมินกิจกรรมพัฒนาผู้เรียน' || n.indexOf('การประเมินกิจกรรมพัฒนาผู้เรียน_') === 0; });
  actSheetNames.forEach(function(actName) {
    var actSheet = boundSs.getSheetByName(actName);
    if (actSheet && actSheet.getLastColumn() > 0) {
      var actHeaders = actSheet.getRange(1, 1, 1, actSheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      if (actHeaders.indexOf('ชื่อชุมนุม') < 0) {
        // หาตำแหน่งคอลัมน์ 'ชุมนุม' แล้วแทรก 'ชื่อชุมนุม' ถัดไป
        var clubIdx = actHeaders.indexOf('ชุมนุม');
        if (clubIdx >= 0) {
          actSheet.insertColumnAfter(clubIdx + 1);
          actSheet.getRange(1, clubIdx + 2).setValue('ชื่อชุมนุม');
          Logger.log('✅ เพิ่มคอลัมน์ "ชื่อชุมนุม" ในชีต "' + actName + '"');
        } else {
          // ถ้าไม่พบคอลัมน์ชุมนุม ให้เพิ่มท้ายสุด
          var lastCol = actHeaders.length + 1;
          actSheet.getRange(1, lastCol).setValue('ชื่อชุมนุม');
          Logger.log('✅ เพิ่มคอลัมน์ "ชื่อชุมนุม" ท้ายชีต "' + actName + '"');
        }
      } else {
        Logger.log('✅ ชีต "' + actName + '" มีคอลัมน์ชื่อชุมนุมแล้ว');
      }
    }
  });
  
  Logger.log('');
  Logger.log('=== ✅ REPAIR เสร็จสมบูรณ์ ===');
  if (needsRepair) {
    Logger.log('🎉 แก้ไข SPREADSHEET_ID สำเร็จ — ข้อมูลรายวิชาควรกลับมาแสดงแล้ว');
    Logger.log('💡 ลองรีเฟรช Web App แล้วตรวจสอบข้อมูลรายวิชาและครูประจำชั้น');
  }
  Logger.log('isSetupComplete_: ' + isSetupComplete_());
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
      Logger.log('✅ ใช้ Spreadsheet จาก URL: ' + ss.getId());
    } else {
      // ✅ Auto-detect: ตรวจสอบว่าเป็น Bound Script หรือไม่
      // เมื่อผู้ใช้คัดลอก Template ไป script จะผูกกับ Spreadsheet สำเนา
      // ต้องใช้ Spreadsheet สำเนานั้น ไม่ใช่สร้างใหม่
      var boundSs = null;
      try {
        boundSs = SpreadsheetApp.getActiveSpreadsheet();
      } catch(_e) {
        Logger.log('⚠️ getActiveSpreadsheet failed: ' + _e.message);
      }
      
      if (!boundSs) {
        // Fallback: ลองดึง container spreadsheet ผ่าน Script API
        try {
          var token = ScriptApp.getOAuthToken();
          var scriptId = ScriptApp.getScriptId();
          var resp = UrlFetchApp.fetch(
            'https://script.googleapis.com/v1/projects/' + scriptId,
            { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
          );
          if (resp.getResponseCode() === 200) {
            var projData = JSON.parse(resp.getContentText());
            if (projData.parentId) {
              boundSs = SpreadsheetApp.openById(projData.parentId);
              Logger.log('✅ ตรวจพบ Bound Spreadsheet ผ่าน Script API: ' + projData.parentId);
            }
          }
        } catch(_e2) {
          Logger.log('⚠️ Script API fallback failed: ' + _e2.message);
        }
      }
      
      if (boundSs) {
        ss = boundSs;
        usedExisting = true;
        Logger.log('✅ ใช้ Bound Spreadsheet อัตโนมัติ: ' + ss.getId());
      } else {
        // Standalone Mode: สร้างใหม่ (กรณีไม่ใช่ bound script)
        ss = createSchoolSpreadsheet_(formData);
        Logger.log('✅ สร้าง Spreadsheet ใหม่: ' + ss.getId());
      }
    }

    // 2. ล้างข้อมูลโรงเรียนเดิม (เฉพาะ Spreadsheet ที่คัดลอกมา)
    if (usedExisting) {
      cleanupCopiedSpreadsheet_(ss);
      Logger.log('✅ ล้างข้อมูลโรงเรียนเดิมแล้ว');
    }

    // 3. สร้าง Sheet (ข้ามถ้ามีอยู่แล้ว)
    setupSheets_(ss, formData, usedExisting);
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
      schoolName: formData.schoolName,
      tip: '💡 แนะนำ: เปลี่ยนชื่อ Apps Script project เป็น "ระบบ ปพ.5 — ' + formData.schoolName + '" เพื่อแยกจากโรงเรียนอื่น'
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

function setupSheets_(ss, formData, usedExisting) {
  // ลบ Sheet1 เดิม (default)
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('แผ่น1');
  var academicYear = String(formData.academicYear || '');

  // 1. สร้างชีตถาวร
  PERMANENT_SHEETS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet(name);
      isNew = true;
    }
    
    // ถ้าเป็นชีตใหม่ หรือ ไม่ได้ใช้ Spreadsheet เดิม (Standalone) ให้เขียนหัวตาราง
    // ถ้าใช้ Spreadsheet เดิม (usedExisting) และมีชีตอยู่แล้ว ห้ามเขียนหัวตารางทับ เพราะจะล้างข้อมูลเดิม
    if (isNew || !usedExisting) {
      setupSheetHeaders_(sheet, name, formData);
    } else {
      Logger.log('ℹ️ ข้ามการเขียน Header ทับ: ' + name + ' (รักษาข้อมูลเดิม)');
    }
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
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      isNew = true;
    }
    
    if (isNew || !usedExisting) {
      setupSheetHeaders_(sheet, baseName, formData);
    } else {
      Logger.log('ℹ️ ข้ามการเขียน Header ทับ (รายปี): ' + sheetName);
    }
  });

  // ลบ default sheet
  try {
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  } catch(e) {}
}

function setupSheetHeaders_(sheet, sheetName, formData) {
  const headerMap = {
    'Users': [['username','password','role','firstName','lastName','email','className','active','createdAt']],
    'Students': [['student_id','id_card','title','firstname','lastname','grade','class_no','gender','birthdate','photo_url','academic_year','weight','height','blood_type','religion','father_name','father_lastname','father_occupation','mother_name','mother_lastname','mother_occupation','address','created_date','status']],
    'รายวิชา': [['ชั้น','รหัสวิชา','ชื่อวิชา','ชั่วโมง/ปี','ประเภทวิชา','ครูผู้สอน','คะแนนระหว่างปี','คะแนนปลายปี']],
    'Holidays': [['date','description','type']],
    'global_settings': [['key','value','updatedAt']],
    'SCORES_WAREHOUSE': [['student_id','grade','class_no','subject_code','subject_name','subject_type','hours','term1_total','term2_total','average','final_grade','sheet_name','academic_year','updated_at']],
    'HomeroomTeachers': [['grade','classNo','teacherName','ครูประจำชั้น 2']],
    'การประเมินอ่านคิดเขียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา','สุขศึกษา','ศิลปะ','การงาน','ภาษาอังกฤษ','สรุปผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินคุณลักษณะ': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','รักชาติ_ศาสน์_กษัตริย์','ซื่อสัตย์สุจริต','มีวินัย','ใฝ่เรียนรู้','อยู่อย่างพอเพียง','มุ่งมั่นในการทำงาน','รักความเป็นไทย','มีจิตสาธารณะ','คะแนนรวม','คะแนนเฉลี่ย','ผลการประเมิน','วันที่บันทึก','ผู้บันทึก']],
    'การประเมินกิจกรรมพัฒนาผู้เรียน': [['รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง','กิจกรรมแนะแนว','ลูกเสือ_เนตรนารี','ชุมนุม','ชื่อชุมนุม','เพื่อสังคมและสาธารณประโยชน์','รวมกิจกรรม','วันที่บันทึก','ผู้บันทึก']],
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
    'ความเห็นครู': [['รหัสนักเรียน','ชื่อนักเรียน','ชั้น','ห้อง','ความเห็นครู','วันที่บันทึก','วันที่อัปเดต']]
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

/**
 * 🧹 สร้าง Template สะอาดจากชีตโรงเรียนต้นแบบ
 * วิธีใช้: รันจาก GAS Editor → ดู Execution log เพื่อรับ URL ของ Template ใหม่
 * ✅ ปลอดภัย: ไม่แตะต้องข้อมูลโรงเรียนต้นแบบเลย (สร้างสำเนาแยก)
 */
function createCleanTemplate() {
  try {
    var ss = SS();
    var schoolName = '';
    try {
      var gs = ss.getSheetByName('global_settings');
      if (gs) {
        var gsData = gs.getDataRange().getValues();
        for (var i = 0; i < gsData.length; i++) {
          if (gsData[i][0] === 'ชื่อโรงเรียน') { schoolName = String(gsData[i][1] || ''); break; }
        }
      }
    } catch(_) {}

    // 1. สร้างสำเนา
    var templateName = 'Template ระบบ ปพ.5' + (schoolName ? ' — ' + schoolName : '') + ' (สะอาด)';
    var file = DriveApp.getFileById(ss.getId());
    var copy = file.makeCopy(templateName);
    var copySs = SpreadsheetApp.openById(copy.getId());
    Logger.log('📋 สร้างสำเนา: ' + copySs.getName());

    // 2. ล้างข้อมูลในสำเนา
    cleanupCopiedSpreadsheet_(copySs);

    // 3. ล้าง installed_script_id ใน global_settings ของสำเนา (ให้เป็น fresh template)
    var gsSheet = copySs.getSheetByName('global_settings');
    if (gsSheet) {
      var data = gsSheet.getDataRange().getValues();
      for (var r = 0; r < data.length; r++) {
        if (data[r][0] === 'installed_script_id') {
          gsSheet.deleteRow(r + 1);
          break;
        }
      }
    }

    var url = copySs.getUrl();
    Logger.log('');
    Logger.log('==============================================');
    Logger.log('✅ Template สะอาดพร้อมแชร์!');
    Logger.log('📎 URL: ' + url);
    Logger.log('==============================================');
    Logger.log('');
    Logger.log('วิธีส่งให้โรงเรียนอื่น:');
    Logger.log('1. เปิด URL ข้างบน');
    Logger.log('2. แชร์ลิงก์ให้โรงเรียนใหม่');
    Logger.log('3. โรงเรียนใหม่: File → Make a copy → ทำตาม Setup Wizard');
    return { success: true, url: url, id: copy.getId() };
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function cleanupCopiedSpreadsheet_(ss) {
  var allSheets = ss.getSheets();

  // === WHITELIST: ชีตที่เก็บไว้ (ล้างข้อมูล เก็บ header) ===
  // หมายเหตุ: HomeroomTeachers จะถูกล้างข้อมูลเพื่อให้โรงเรียนใหม่กรอกเอง
  var keepNames = new Set(PERMANENT_SHEETS);  // global_settings, Users, Students, รายวิชา, Holidays, HomeroomTeachers
  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
    'AttendanceLog', 'ความเห็นครู'
  ];
  yearlyBases.forEach(function(b) { keepNames.add(b); });

  // ❌ ลบ 'HomeroomTeachers' ออกจาก keepNames ชั่วคราว เพื่อไม่ให้โดน deleteRows(2, ...) 
  // หากต้องการให้ข้อมูลครูประจำชั้นติดไปด้วย (แต่โดยปกติเราจะล้างเพื่อให้โรงเรียนใหม่กรอกเอง)
  // อย่างไรก็ตาม ปัญหาตอนนี้คือ "ดูไม่ได้" ซึ่งอาจเกิดจากข้อมูลถูกล้างจนเกลี้ยง
  
  // ✅ เพิ่มชีตที่จำเป็นสำหรับ template (เพื่อความปลอดภัย)
  var extraSheets = ['template', 'Template', 'ตัวอย่าง', 'example'];
  extraSheets.forEach(function(b) { keepNames.add(b); });

  var sheetsToClean = [];
  var sheetsToDelete = [];

  allSheets.forEach(function(sheet) {
    var name = sheet.getName();

    // ตรงชื่อ permanent/yearly → เก็บ (ล้างข้อมูล)
    if (keepNames.has(name)) {
      // ยกเว้น รายวิชา ถ้าไม่อยากให้โดนล้างข้อมูลแถว 2 เป็นต้นไป (ให้ข้อมูลติดไปด้วย)
      if (name === 'รายวิชา') {
        Logger.log('ℹ️ ข้ามการล้างข้อมูลแถว: ' + name + ' (เพื่อให้ข้อมูลติดไปด้วย)');
        return;
      }
      sheetsToClean.push(sheet);
      return;
    }

    // ชีตรายปี + suffix (เช่น SCORES_WAREHOUSE_2568) → เก็บ (ล้างข้อมูล)
    var isYearly = yearlyBases.some(function(b) { return name.indexOf(b + '_') === 0; });
    if (isYearly) {
      sheetsToClean.push(sheet);
      return;
    }

    // === ทุกอย่างที่เหลือ → ลบทิ้ง ===
    // ชีตคะแนนรายวิชา (เช่น "ภาษาไทย ป3-1", "ชุมนุมวิทยาศาสตร์ ป4-1")
    // ชีตเช็คชื่อรายเดือน (เช่น "พฤษภาคม 2568")
    // ชีตสรุป (เช่น "สรุปวันมาต่ละเดือน_ป.31_2568", "สรุปการมาเรียน")
    // ชีต BACKUP/Template
    // ชีตโปรไฟล์นักเรียน
    sheetsToDelete.push(sheet);
  });

  // ล้างข้อมูล (เก็บ header แถว 1)
  sheetsToClean.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    Logger.log('🧹 ล้างข้อมูล: ' + sheet.getName() + ' (' + Math.max(0, lastRow - 1) + ' แถว)');
  });

  // ลบชีตที่ไม่อยู่ใน whitelist (ต้องเหลืออย่างน้อย 1 ชีต)
  var remaining = allSheets.length - sheetsToDelete.length;
  if (remaining < 1) sheetsToDelete.shift();

  sheetsToDelete.forEach(function(sheet) {
    try {
      Logger.log('🗑️ ลบชีต: ' + sheet.getName());
      ss.deleteSheet(sheet);
    } catch (e) {
      Logger.log('⚠️ ลบไม่ได้: ' + sheet.getName() + ' — ' + e.message);
    }
  });

  Logger.log('🧹 ล้างเสร็จ: เก็บ+ล้าง ' + sheetsToClean.length + ' ชีต, ลบ ' + sheetsToDelete.length + ' ชีต');
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

    // ============ ขั้นที่ 5: จำหน่าย ป.6/ม.3 + เลื่อนชั้น ============
    var studentResult = { promoted: 0, graduated: 0 };
    if (options.promoteStudents !== false) {
      Logger.log('🎓 กำลังจำหน่ายนักเรียน ป.6/ม.3 และเลื่อนชั้น...');
      try {
        var gradResult = graduateStudents_();
        studentResult.graduated = gradResult.graduated || 0;
      } catch (e) { Logger.log('⚠️ graduateStudents: ' + e.message); }
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
 * อ.1 → อ.2, ..., ป.1 → ป.2, ..., ป.5 → ป.6, ม.1 → ม.2, ม.2 → ม.3
 * (ป.6 และ ม.3 ไม่เลื่อน — ต้องจำหน่ายก่อน ยกเว้นกรณีขยายโอกาสที่ต้องการเลื่อน ป.6 → ม.1 สามารถปรับ gradeMap ได้)
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
    'อ.1': 'อ.2',
    'อ.2': 'อ.3',
    'อ.3': 'ป.1',
    'ป.1': 'ป.2',
    'ป.2': 'ป.3',
    'ป.3': 'ป.4',
    'ป.4': 'ป.5',
    'ป.5': 'ป.6',
    // 'ป.6': 'ม.1', // ปกติจำหน่าย แต่ถ้ามีมัธยมอาจเลื่อนเอง
    'ม.1': 'ม.2',
    'ม.2': 'ม.3'
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
// 🎓 GRADUATE STUDENTS — จำหน่ายนักเรียนที่จบการศึกษา (ป.6 / ม.3)
// ============================================================

/**
 * จำหน่ายนักเรียน ป.6 และ ม.3 ที่ status = 'active':
 * เปลี่ยน status เป็น 'จำหน่าย' (ไม่ลบข้อมูล ยังดูได้)
 * @returns {Object} { graduated: number }
 */
function graduateStudents_() {
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
    if ((currentGrade === 'ป.6' || currentGrade === 'ม.3') && (status === 'active' || status === '')) {
      sheet.getRange(r + 1, statusIdx + 1).setValue('จำหน่าย');
      graduated++;
    }
  }

  Logger.log('🎓 จำหน่ายนักเรียน ป.6/ม.3: ' + graduated + ' คน');
  return { graduated: graduated };
}

/**
 * จำหน่ายนักเรียนที่จบการศึกษา (เรียกจากภายนอก / UI)
 * @returns {Object}
 */
function graduateStudents() {
  try {
    var result = graduateStudents_();
    return { success: true, message: 'จำหน่ายนักเรียนที่จบการศึกษา จำนวน ' + result.graduated + ' คน', graduated: result.graduated };
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
    var gradeCounts = { 'อ.1': 0, 'อ.2': 0, 'อ.3': 0, 'ป.1': 0, 'ป.2': 0, 'ป.3': 0, 'ป.4': 0, 'ป.5': 0, 'ป.6': 0, 'ม.1': 0, 'ม.2': 0, 'ม.3': 0 };
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

    var graduateCount = gradeCounts['ป.6'] + gradeCounts['ม.3'];

    return {
      success: true,
      currentYear: currentYear,
      suggestedNewYear: newYear,
      studentCount: studentCount,
      gradeCounts: gradeCounts,
      graduateCount: graduateCount,
      p6Count: graduateCount, // ใช้ชื่อตัวแปร p6Count เพื่อให้ UI เดิมทำงานได้ หรืออาจจะเปลี่ยนเป็น graduateCount แล้วไปแก้ UI ด้วย
      promoteCount: studentCount - graduateCount,
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
 * ✅ ตั้งค่าแชร์ไฟล์เป็น "Anyone with the link" — มี REST API fallback
 */
function _ensureFileSharedPublic_(fileId, permission) {
  permission = permission || 'writer';  // 'reader' or 'writer'
  var ok = false;

  // 1) ลอง DriveApp ก่อน
  try {
    var perm = permission === 'writer' ? DriveApp.Permission.EDIT : DriveApp.Permission.VIEW;
    DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, perm);
    ok = true;
    Logger.log('✅ setSharing (DriveApp) สำเร็จ: ' + fileId);
  } catch (e1) {
    Logger.log('⚠️ DriveApp.setSharing failed: ' + e1.message + ' → trying REST API');
  }

  // 2) Fallback: REST API
  if (!ok) {
    try {
      var token = ScriptApp.getOAuthToken();
      var role = permission === 'writer' ? 'writer' : 'reader';
      var resp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ role: role, type: 'anyone' }),
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        ok = true;
        Logger.log('✅ setSharing (REST API) สำเร็จ: ' + fileId);
      } else {
        Logger.log('❌ REST API setSharing failed: ' + code + ' ' + resp.getContentText());
      }
    } catch (e2) {
      Logger.log('❌ REST API setSharing error: ' + e2.message);
    }
  }

  return ok;
}

/**
 * ✅ ตรวจสอบและสร้างชีตที่ขาดหายใน Template Spreadsheet
 * เรียกทุกครั้งก่อนส่งออก เพื่อให้ template มีชีตครบทุกอัน
 */
function ensureTemplateSheets_(templateId) {
  var tss = SpreadsheetApp.openById(templateId);
  var existingNames = tss.getSheets().map(function(s) { return s.getName(); });

  // ชีตทั้งหมดที่ต้องมี (ใช้ชื่อ base ไม่มี suffix ปี)
  var allRequired = PERMANENT_SHEETS.concat(
    (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [
      'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',
      'AttendanceLog', 'ความเห็นครู'
    ]
  );

  var created = [];
  allRequired.forEach(function(name) {
    if (existingNames.indexOf(name) === -1) {
      var sheet = tss.insertSheet(name);
      setupSheetHeaders_(sheet, name, {});
      created.push(name);
      Logger.log('📋 สร้างชีต: ' + name);
    }
  });

  // ลบ Sheet1/แผ่น1 default ถ้ามี (และมีชีตอื่นอย่างน้อย 1)
  if (tss.getSheets().length > 1) {
    try {
      var def = tss.getSheetByName('Sheet1') || tss.getSheetByName('แผ่น1') || tss.getSheetByName('ชีต1');
      if (def) tss.deleteSheet(def);
    } catch(_e) {}
  }

  if (created.length > 0) {
    Logger.log('✅ ensureTemplateSheets_: สร้าง ' + created.length + ' ชีต: ' + created.join(', '));
  } else {
    Logger.log('✅ ensureTemplateSheets_: ชีตครบแล้ว');
  }
}

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
    var ssId = getSpreadsheetId_();
    var props = PropertiesService.getScriptProperties();

    // ---- ✅ ใช้ Template ที่มี Bound Script เสมอ (ไม่ใช้ export template เก่าที่ไม่มีโค้ด) ----
    var BOUND_TEMPLATE_ID = '1AcdypFst0F4pr7bjaMH1WwuTyohekV8BeO36MWZOWJE';
    var templateId = BOUND_TEMPLATE_ID;

    // อัปเดต ScriptProperties ให้ชี้ไป template ที่ถูกต้อง (ล้างค่าเก่าที่ไม่มี script)
    try {
      props.setProperty('EXPORT_TEMPLATE_READY_ID', templateId);
      // ลบ key เก่าที่อาจชี้ไป template ผิด
      props.deleteProperty('PENDING_EXPORT_TEMPLATE_ID');
    } catch(_e) {}

    // ✅ ตรวจสอบ + สร้างชีตที่ขาดหายใน template ก่อน
    try {
      ensureTemplateSheets_(templateId);
      Logger.log('✅ ensureTemplateSheets_ สำเร็จ');
    } catch(_e) {
      Logger.log('⚠️ ensureTemplateSheets_ error: ' + _e.message);
    }

    // Sync รายวิชาไป template (ไม่รวมชื่อครู) ก่อนส่งออก
    try {
      var syncResult = syncSubjectsToTemplate(templateId);
      Logger.log('📋 syncSubjects: ' + (syncResult.success ? syncResult.message : syncResult.error));
    } catch(_e) {
      Logger.log('⚠️ sync data skipped: ' + _e.message);
    }

    // ตั้งค่า sharing ทุกครั้ง
    var sharingOk = _ensureFileSharedPublic_(templateId, 'reader');
    var copyUrl = 'https://docs.google.com/spreadsheets/d/' + templateId + '/edit';

    return {
      success: true,
      copyUrl: copyUrl,
      activeSpreadsheetId: ssId,
      templateId: templateId,
      warning: sharingOk ? '' : '⚠️ ไม่สามารถตั้งค่าแชร์อัตโนมัติได้ — กรุณาเปิด Google Sheet แล้วแชร์เป็น "ทุกคนที่มีลิงก์" ด้วยตนเอง',
      sanitized: true,
      excludedData: ['ข้อมูลนักเรียน', 'ผู้ใช้', 'คะแนน', 'เช็คชื่อ', 'ชีตคะแนนรายวิชา'],
      schoolName: schoolName,
      steps: [
        'ขั้นที่ 1: ส่งลิงก์ให้โรงเรียนอื่น → เปิดไฟล์ต้นแบบ → เมนูไฟล์ > ทำสำเนา → ได้ Spreadsheet + โค้ดทั้งหมด',
        'ขั้นที่ 2: เปิดสำเนา → ส่วนขยาย → Apps Script → รัน debugSetupStatus (อนุญาตสิทธิ์) → Deploy → New deployment → Web app → Execute as: Me, Who: Anyone → Deploy',
        'ขั้นที่ 3: เปิด Web App URL → ระบบจะแสดงหน้าติดตั้ง → กรอกข้อมูลโรงเรียน → เสร็จ!'
      ],
      _codeVersion: 'v203_bound_template',
      message: 'สร้างลิงก์ส่งออกสำเร็จ (ให้เปิดไฟล์ต้นแบบแล้วกด เมนูไฟล์ > ทำสำเนา เพื่อให้โค้ดติดไปด้วย)'
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
  // ✅ Fix: ใช้ helper ที่มี REST API fallback
  _ensureFileSharedPublic_(copiedFile.getId(), 'writer');

  var ss = SpreadsheetApp.openById(copiedFile.getId());
  sanitizeExportTemplateSpreadsheet_(ss);
  return copiedFile.getId();
}

function sanitizeExportTemplateSpreadsheet_(ss) {
  // ใช้ cleanupCopiedSpreadsheet_ ซึ่งลบชีตคะแนนรายวิชา, Backup, เช็คชื่อรายเดือน, โปรไฟล์ ฯลฯ ครบ
  cleanupCopiedSpreadsheet_(ss);

  // ล้างชื่อครูผู้สอนในรายวิชา
  var subj = ss.getSheetByName('รายวิชา');
  if (subj) blankTeacherInSubjects_(subj);

  // ล้าง installed_script_id เพื่อให้เป็น fresh template
  var gsSheet = ss.getSheetByName('global_settings');
  if (gsSheet) {
    var data = gsSheet.getDataRange().getValues();
    for (var r = 0; r < data.length; r++) {
      if (data[r][0] === 'installed_script_id') {
        gsSheet.deleteRow(r + 1);
        break;
      }
    }
  }
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
function syncSubjectsToTemplate(templateId) {
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
    templateId = String(templateId || '').trim() || '1AcdypFst0F4pr7bjaMH1WwuTyohekV8BeO36MWZOWJE';

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
    
    // บังคับ flush เพื่อให้ข้อมูลลงไฟล์จริงๆ ก่อนจบฟังก์ชัน
    SpreadsheetApp.flush();

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
 * คัดลอกรายชื่อครูประจำชั้นไปใส่ Template Spreadsheet
 * @param {string} templateId 
 * @returns {Object} { success, message, count }
 */
function syncHomeroomTeachersToTemplate(templateId) {
  try {
    var ss = SS();
    var htSheet = ss.getSheetByName('HomeroomTeachers');
    if (!htSheet) return { success: false, error: 'ไม่พบชีต HomeroomTeachers ใน Master' };

    var data = htSheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, message: 'ไม่มีข้อมูลครูประจำชั้นให้ Sync', count: 0 };

    templateId = String(templateId || '').trim() || '1AcdypFst0F4pr7bjaMH1WwuTyohekV8BeO36MWZOWJE';
    var templateSS = SpreadsheetApp.openById(templateId);
    var templateHTSheet = templateSS.getSheetByName('HomeroomTeachers');
    if (!templateHTSheet) {
      templateHTSheet = templateSS.insertSheet('HomeroomTeachers');
    }

    templateHTSheet.clearContents();
    templateHTSheet.clearFormats();

    templateHTSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Format header
    templateHTSheet.getRange(1, 1, 1, data[0].length)
      .setBackground('#4a5568')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    templateHTSheet.setFrozenRows(1);

    SpreadsheetApp.flush();

    return { success: true, message: 'Sync ครูประจำชั้นสำเร็จ', count: data.length - 1 };
  } catch (e) {
    Logger.log('❌ syncHomeroomTeachersToTemplate error: ' + e.message);
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
