/**
 * ============================================================
 * 📝 ASSESSMENTS MANAGEMENT — ฉบับสมบูรณ์ (แทนที่ไฟล์เดิมได้ทั้งหมด)
 * ============================================================
 * 
 * ปรับปรุง Data Integrity:
 *   #11 — Server-side Data Validation
 *   #12 — Race Condition Protection (LockService)
 *   #13 — Atomic Row Lookup + Write
 * 
 * โครงสร้างไฟล์:
 *   Section A — Constants & Sheet Headers
 *   Section B — Validation Utilities          ← ใหม่
 *   Section C — Locking & Atomic Write        ← ใหม่
 *   Section D — Sheet Setup Helpers
 *   Section E — Get Functions (อ่านข้อมูล)
 *   Section F — Save Functions (เขียนข้อมูล)  ← ปรับปรุง
 *   Section G — PDF Generation
 *   Section H — HTML Templates
 * 
 * ============================================================
 */


// ============================================================
// SECTION A: CONSTANTS & SHEET HEADERS
// ============================================================

// ----- Read Think Write (อ่าน คิด วิเคราะห์ เขียน) -----
const RTW_SHEET = 'การประเมินอ่านคิดเขียน';
const RTW_HEADERS = [
  'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง',
  'ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์', 'สังคมศึกษา',
  'สุขศึกษา', 'ศิลปะ', 'การงาน', 'ภาษาอังกฤษ',
  'สรุปผลการประเมิน', 'วันที่บันทึก', 'ผู้บันทึก'
];

// ----- Characteristic (คุณลักษณะอันพึงประสงค์) -----
const CHARACTER_SHEET = 'การประเมินคุณลักษณะ';
const CHARACTER_HEADERS = [
  'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง',
  'รักชาติ_ศาสน์_กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้',
  'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ',
  'คะแนนรวม', 'คะแนนเฉลี่ย', 'ผลการประเมิน',
  'วันที่บันทึก', 'ผู้บันทึก'
];
const TRAIT_HEADERS = [
  'รักชาติ_ศาสน์_กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้',
  'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ'
];

// ----- Activity (กิจกรรมพัฒนาผู้เรียน) -----
const ACTIVITY_SHEET = 'การประเมินกิจกรรมพัฒนาผู้เรียน';
const ACTIVITY_HEADERS = [
  'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง',
  'กิจกรรมแนะแนว', 'ลูกเสือ_เนตรนารี', 'ชุมนุม',
  'เพื่อสังคมและสาธารณประโยชน์', 'รวมกิจกรรม',
  'วันที่บันทึก', 'ผู้บันทึก'
];

// ----- Competency (สมรรถนะสำคัญ) -----
const COMPETENCY_SHEET = 'การประเมินสมรรถนะ';
const COMPETENCY_HEADERS = [
  'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง',
  'สื่อสาร_รับ-ส่งสาร', 'สื่อสาร_ถ่ายทอด', 'สื่อสาร_วิธีการ', 'สื่อสาร_เจรจา', 'สื่อสาร_เลือกรับ',
  'คิด_วิเคราะห์', 'คิด_สร้างสรรค์', 'คิด_วิจารณญาณ', 'คิด_สร้างความรู้', 'คิด_ตัดสินใจ',
  'แก้ปัญหา_แก้ปัญหา', 'แก้ปัญหา_ใช้เหตุผล', 'แก้ปัญหา_เข้าใจสังคม', 'แก้ปัญหา_แสวงหาความรู้', 'แก้ปัญหา_ตัดสินใจ',
  'ทักษะชีวิต_เรียนรู้ตนเอง', 'ทักษะชีวิต_ทำงานกลุ่ม', 'ทักษะชีวิต_นำไปใช้', 'ทักษะชีวิต_จัดการปัญหา', 'ทักษะชีวิต_หลีกเลี่ยง',
  'เทคโนโลยี_เลือกใช้', 'เทคโนโลยี_ทักษะกระบวนการ', 'เทคโนโลยี_พัฒนาตนเอง', 'เทคโนโลยี_แก้ปัญหา', 'เทคโนโลยี_คุณธรรม',
  'วันที่บันทึก', 'ผู้บันทึก'
];

// ----- Validation constants -----
// ----- Validation constants (ใช้จาก validation.gs) -----
// VALID_GRADES → ใช้ V_VALID_GRADES จาก validation.gs
// VALID_CLASS_NOS → ใช้ V_VALID_CLASS_NOS จาก validation.gs
const VALID_ACTIVITY_RESULTS = ['ผ่าน', 'ไม่ผ่าน'];
const MAX_BATCH_SIZE = 100;


// ============================================================
// SECTION B: VALIDATION UTILITIES (แก้ปัญหา #11)
// ============================================================

/**
 * ตรวจสอบชั้นเรียน
 */
function validateGrade_(grade) {
  const sanitized = String(grade || '').trim();
  if (!sanitized) {
    return { valid: false, sanitized: '', error: 'ไม่ได้ระบุชั้นเรียน' };
  }
  if (!V_VALID_GRADES.includes(sanitized)) {
    return { valid: false, sanitized, error: 'ชั้นเรียน "' + sanitized + '" ไม่ถูกต้อง (ต้องเป็น ' + V_VALID_GRADES.join(', ') + ')' };
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบห้อง
 */
function validateClassNo_(classNo) {
  const sanitized = String(classNo || '').trim();
  if (!sanitized) {
    return { valid: false, sanitized: '', error: 'ไม่ได้ระบุห้อง' };
  }
    if (!V_VALID_CLASS_NOS.includes(sanitized)) {
    return { valid: false, sanitized, error: 'ห้อง "' + sanitized + '" ไม่ถูกต้อง (ต้องเป็น 1-10)' };
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบ grade + classNo รวม (ใช้กับ get functions)
 */
function validateGradeAndClass_(grade, classNo) {
  const g = validateGrade_(grade);
  if (!g.valid) throw new Error(g.error);
  const c = validateClassNo_(classNo);
  if (!c.valid) throw new Error(c.error);
  return { grade: g.sanitized, classNo: c.sanitized };
}

/**
 * ตรวจสอบรหัสนักเรียน
 */
function validateStudentId_(studentId) {
  const sanitized = String(studentId || '').trim();
  if (!sanitized) {
    return { valid: false, sanitized: '', error: 'ไม่ได้ระบุรหัสนักเรียน' };
  }
  if (sanitized.length > 20) {
    return { valid: false, sanitized, error: 'รหัสนักเรียนยาวเกินไป (สูงสุด 20 ตัวอักษร)' };
  }
  if (!/^[a-zA-Z0-9\u0E01-\u0E5B\-_]+$/.test(sanitized)) {
    return { valid: false, sanitized, error: 'รหัสนักเรียนมีอักขระที่ไม่อนุญาต' };
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบคะแนนคุณลักษณะ (0-3, จำนวน 8 ตัว)
 */
function validateCharacteristicScores_(scores) {
  if (!Array.isArray(scores)) {
    return { valid: false, sanitized: [], error: 'ข้อมูลคะแนนไม่ใช่ array' };
  }
  if (scores.length !== 8) {
    return { valid: false, sanitized: [], error: 'ต้องมีคะแนน 8 รายการ (ได้รับ ' + scores.length + ')' };
  }
  const sanitized = [];
  for (var i = 0; i < scores.length; i++) {
    var val = scores[i];
    if (val === '' || val === null || val === undefined) {
      sanitized.push('');
      continue;
    }
    var num = parseInt(val, 10);
    if (!Number.isFinite(num) || num < 0 || num > 3) {
      return { valid: false, sanitized: [], error: 'คะแนนรายการที่ ' + (i + 1) + ' ไม่ถูกต้อง: "' + val + '" (ต้องเป็น 0-3)' };
    }
    sanitized.push(num);
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบคะแนนรายวิชา (0-4)
 */
function validateSubjectScores_(scores) {
  if (!scores || typeof scores !== 'object') {
    return { valid: false, sanitized: {}, error: 'ข้อมูลคะแนนไม่ถูกต้อง' };
  }
  var subjects = ['thai', 'math', 'science', 'social', 'health', 'art', 'work', 'english'];
  var sanitized = {};
  for (var s = 0; s < subjects.length; s++) {
    var subj = subjects[s];
    var val = scores[subj];
    if (val === '' || val === null || val === undefined) {
      sanitized[subj] = '';
      continue;
    }
    var num = parseInt(val, 10);
    if (!Number.isFinite(num) || num < 0 || num > 4) {
      return { valid: false, sanitized: {}, error: 'คะแนนวิชา "' + subj + '" ไม่ถูกต้อง: "' + val + '" (ต้องเป็น 0-4)' };
    }
    sanitized[subj] = num;
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบผลกิจกรรม
 */
function validateActivityResults_(activities) {
  if (!activities || typeof activities !== 'object') {
    return { valid: false, sanitized: {}, error: 'ข้อมูลกิจกรรมไม่ถูกต้อง' };
  }
  var fields = ['guidance', 'scout', 'club', 'social', 'overall'];
  var sanitized = {};
  for (var f = 0; f < fields.length; f++) {
    var field = fields[f];
    var val = String(activities[field] || '').trim();
    if (val && !VALID_ACTIVITY_RESULTS.includes(val)) {
      return { valid: false, sanitized: {}, error: 'ผลกิจกรรม "' + field + '" ไม่ถูกต้อง: "' + val + '" (ต้องเป็น ผ่าน/ไม่ผ่าน)' };
    }
    sanitized[field] = val || 'ผ่าน';
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจสอบคะแนนสมรรถนะ (0-3, จำนวน 25 ตัว)
 */
function validateCompetencyScores_(scores) {
  if (!Array.isArray(scores) || scores.length !== 25) {
    return { valid: false, sanitized: [], error: 'ต้องมีคะแนน 25 รายการ' };
  }
  var sanitized = [];
  var errors = [];
  for (var i = 0; i < scores.length; i++) {
    var val = scores[i];
    if (val === '' || val === null || val === undefined) {
      sanitized.push('');
      continue;
    }
    var num = parseInt(val, 10);
    if (!Number.isFinite(num) || num < 0 || num > 3) {
      errors.push('ข้อย่อยที่ ' + (i + 1) + ': "' + val + '"');
      sanitized.push('');
    } else {
      sanitized.push(num);
    }
  }
  if (errors.length > 0) {
    return { valid: false, sanitized: [], error: 'คะแนนไม่ถูกต้อง (ต้อง 0-3): ' + errors.join(', ') };
  }
  return { valid: true, sanitized };
}

/**
 * ตรวจว่านักเรียนมีอยู่จริง
 */
function studentExists_(ss, studentId, grade, classNo) {
  var sheet = ss.getSheetByName('Students');
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === studentId &&
        String(data[i][5]).trim() === grade &&
        String(data[i][6]).trim() === classNo) {
      return true;
    }
  }
  return false;
}

/**
 * Validate basic fields ของ record (grade, classNo, studentId, ตรวจนักเรียนจริง)
 */
function validateRecordBasicFields_(rec, ss) {
  var errors = [];
  var gradeResult = validateGrade_(rec.grade);
  if (!gradeResult.valid) errors.push(gradeResult.error);
  var classResult = validateClassNo_(rec.classNo);
  if (!classResult.valid) errors.push(classResult.error);
  var idResult = validateStudentId_(rec.studentId);
  if (!idResult.valid) errors.push(idResult.error);

  if (errors.length === 0 && ss) {
    if (!studentExists_(ss, idResult.sanitized, gradeResult.sanitized, classResult.sanitized)) {
      errors.push('ไม่พบนักเรียนรหัส "' + idResult.sanitized + '" ในชั้น ' + gradeResult.sanitized + ' ห้อง ' + classResult.sanitized);
    }
  }
  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: {
      grade: gradeResult.valid ? gradeResult.sanitized : '',
      classNo: classResult.valid ? classResult.sanitized : '',
      studentId: idResult.valid ? idResult.sanitized : ''
    }
  };
}


// ============================================================
// SECTION C: LOCKING & ATOMIC WRITE (แก้ปัญหา #12, #13)
// ============================================================

/**
 * ทำงานภายใต้ Lock ป้องกัน race condition
 */
function withLock_(lockKey, callback, timeoutMs) {
  var lock = LockService.getScriptLock();
  var timeout = timeoutMs || 15000;
  try {
    var acquired = lock.tryLock(timeout);
    if (!acquired) {
      throw new Error('ไม่สามารถบันทึกได้ — มีผู้ใช้อื่นกำลังบันทึกข้อมูลอยู่ กรุณาลองใหม่ใน 5 วินาที');
    }
    Logger.log('🔒 Lock acquired: ' + lockKey);
    var result = callback();
    Logger.log('🔓 Lock released: ' + lockKey);
    return result;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/**
 * Atomic upsert: อ่าน → สร้าง index → เขียน ภายใน lock เดียว
 * ป้องกัน row shift ระหว่าง read/write
 */
function atomicUpsertRows_(sheet, keyColumnHeader, gradeHeader, classHeader, targetGrade, targetClassNo, records) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  var keyCol = headers.indexOf(keyColumnHeader);
  var gradeCol = headers.indexOf(gradeHeader);
  var classCol = headers.indexOf(classHeader);
  if (keyCol === -1) throw new Error('ไม่พบคอลัมน์ "' + keyColumnHeader + '"');

  // อ่านข้อมูลที่มีอยู่ → สร้าง index (ภายใน lock)
  var existingRowMap = {};
  if (lastRow > 1) {
    var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var i = 0; i < allData.length; i++) {
      var row = allData[i];
      var rowKey = String(row[keyCol] || '').trim();
      var rowGrade = gradeCol !== -1 ? String(row[gradeCol] || '').trim() : '';
      var rowClass = classCol !== -1 ? String(row[classCol] || '').trim() : '';
      var matchGrade = !targetGrade || rowGrade === targetGrade;
      var matchClass = !targetClassNo || rowClass === targetClassNo;
      if (rowKey && matchGrade && matchClass) {
        existingRowMap[rowKey] = i + 2;
      }
    }
  }

  // แยก update vs insert
  var updateBatch = [];
  var insertBatch = [];
  for (var r = 0; r < records.length; r++) {
    var rec = records[r];
    var existingRow = existingRowMap[rec.key];
    if (existingRow) {
      updateBatch.push({ row: existingRow, data: rec.rowData });
    } else {
      insertBatch.push(rec.rowData);
    }
  }

  // เขียน update ทีละแถว
  for (var u = 0; u < updateBatch.length; u++) {
    sheet.getRange(updateBatch[u].row, 1, 1, updateBatch[u].data.length).setValues([updateBatch[u].data]);
  }

  // เขียน insert batch (เร็วกว่า appendRow ทีละแถว)
  if (insertBatch.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, insertBatch.length, insertBatch[0].length).setValues(insertBatch);
  }

  Logger.log('✅ Upsert: updated=' + updateBatch.length + ', inserted=' + insertBatch.length);
  return { updated: updateBatch.length, inserted: insertBatch.length };
}


// ============================================================
// SECTION D: SHEET SETUP HELPERS
// ============================================================

function rtwHeaderIndex_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; });
  return idx;
}

function ensureRTWSheetAndHeaders_() {
  var ss = SS();
  var sheet = null;
  try { if (typeof S_getYearlySheet === 'function') sheet = S_getYearlySheet(RTW_SHEET); } catch(_) {}
  if (!sheet) sheet = ss.getSheetByName(RTW_SHEET);

  if (!sheet) {
    var sheetName = (typeof S_sheetName === 'function') ? S_sheetName(RTW_SHEET) : RTW_SHEET;
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, RTW_HEADERS.length).setValues([RTW_HEADERS]);
    var headerRange = sheet.getRange(1, 1, 1, RTW_HEADERS.length);
    headerRange.setBackground('#ffc107');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    return { sheet: sheet, idx: rtwHeaderIndex_(sheet) };
  }

  var lastC = sheet.getLastColumn();
  var oldHeaders = lastC ? sheet.getRange(1, 1, 1, lastC).getValues()[0].map(String) : [];

  RTW_HEADERS.forEach(function(h) {
    if (!oldHeaders.includes(h)) oldHeaders.push(h);
  });

  if (sheet.getLastColumn() < oldHeaders.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), oldHeaders.length - sheet.getLastColumn());
  }

  var needReorder = oldHeaders.length !== RTW_HEADERS.length || !RTW_HEADERS.every(function(h, i) { return oldHeaders[i] === h; });

  if (needReorder) {
    var lastR = sheet.getLastRow();
    var widthOld = oldHeaders.length;
    var dataOld = lastR > 1 ? sheet.getRange(2, 1, lastR - 1, widthOld).getValues() : [];
    var oldIdx = {};
    oldHeaders.forEach(function(h, i) { oldIdx[h] = i; });

    var dataNew = dataOld.map(function(row) {
      var out = new Array(RTW_HEADERS.length).fill('');
      RTW_HEADERS.forEach(function(h, ni) {
        var oi = oldIdx[h];
        if (oi !== undefined && oi < row.length) out[ni] = row[oi];
      });
      return out;
    });

    sheet.getRange(1, 1, 1, RTW_HEADERS.length).setValues([RTW_HEADERS]);
    if (dataNew.length) {
      sheet.getRange(2, 1, dataNew.length, RTW_HEADERS.length).setValues(dataNew);
    }
    var extra = sheet.getLastColumn() - RTW_HEADERS.length;
    if (extra > 0) sheet.deleteColumns(RTW_HEADERS.length + 1, extra);

    var headerRange2 = sheet.getRange(1, 1, 1, RTW_HEADERS.length);
    headerRange2.setBackground('#ffc107');
    headerRange2.setFontWeight('bold');
    headerRange2.setHorizontalAlignment('center');
  }

  return { sheet: sheet, idx: rtwHeaderIndex_(sheet) };
}

function charHeaderIndex_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; });
  return idx;
}

function ensureCharSheetAndHeaders_() {
  var ss = SS();
  var sheet = null;
  try { if (typeof S_getYearlySheet === 'function') sheet = S_getYearlySheet(CHARACTER_SHEET); } catch(_) {}
  if (!sheet) sheet = ss.getSheetByName(CHARACTER_SHEET);
  if (!sheet) {
    var sheetName = (typeof S_sheetName === 'function') ? S_sheetName(CHARACTER_SHEET) : CHARACTER_SHEET;
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, CHARACTER_HEADERS.length).setValues([CHARACTER_HEADERS]);
    return { sheet: sheet, idx: charHeaderIndex_(sheet) };
  }

  var lastC = Math.max(sheet.getLastColumn(), CHARACTER_HEADERS.length);
  if (sheet.getMaxColumns() < lastC) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastC - sheet.getMaxColumns());
  }

  var oldHeaders = sheet.getRange(1, 1, 1, lastC).getValues()[0].map(String);
  var extras = oldHeaders.filter(function(h) { return h && CHARACTER_HEADERS.indexOf(h) === -1; });
  var nextHeaders = CHARACTER_HEADERS.concat(extras);

  if (sheet.getMaxColumns() < nextHeaders.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), nextHeaders.length - sheet.getMaxColumns());
  }

  var needReorder = nextHeaders.length !== oldHeaders.length || !nextHeaders.every(function(h, i) { return oldHeaders[i] === h; });
  if (needReorder) {
    var lastR = sheet.getLastRow();
    var dataOld = lastR > 1 ? sheet.getRange(2, 1, lastR - 1, oldHeaders.length).getValues() : [];
    var oldIdx = {};
    oldHeaders.forEach(function(h, i) { if (h && oldIdx[h] === undefined) oldIdx[h] = i; });

    var dataNew = dataOld.map(function(row) {
      var out = new Array(nextHeaders.length).fill('');
      nextHeaders.forEach(function(h, ni) {
        var oi = oldIdx[h];
        if (oi !== undefined && oi < row.length) out[ni] = row[oi];
      });
      return out;
    });

    sheet.getRange(1, 1, 1, nextHeaders.length).setValues([nextHeaders]);
    if (dataNew.length) {
      sheet.getRange(2, 1, dataNew.length, nextHeaders.length).setValues(dataNew);
    }
  }

  return { sheet: sheet, idx: charHeaderIndex_(sheet) };
}

function ensureCompetencySheetAndHeaders_() {
  var ss = SS();
  var sheet = null;
  try { if (typeof S_getYearlySheet === 'function') sheet = S_getYearlySheet(COMPETENCY_SHEET); } catch(_) {}
  if (!sheet) sheet = ss.getSheetByName(COMPETENCY_SHEET);

  if (!sheet) {
    var sheetName = (typeof S_sheetName === 'function') ? S_sheetName(COMPETENCY_SHEET) : COMPETENCY_SHEET;
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, COMPETENCY_HEADERS.length).setValues([COMPETENCY_HEADERS]);
    return sheet;
  }

  var lastC = Math.max(sheet.getLastColumn(), COMPETENCY_HEADERS.length);
  if (sheet.getMaxColumns() < lastC) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), lastC - sheet.getMaxColumns());
  }

  var oldHeaders = sheet.getRange(1, 1, 1, lastC).getValues()[0].map(String);
  var extras = oldHeaders.filter(function(h) { return h && COMPETENCY_HEADERS.indexOf(h) === -1; });
  var nextHeaders = COMPETENCY_HEADERS.concat(extras);

  if (sheet.getMaxColumns() < nextHeaders.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), nextHeaders.length - sheet.getMaxColumns());
  }

  var needReorder = nextHeaders.length !== oldHeaders.length || !nextHeaders.every(function(h, i) { return oldHeaders[i] === h; });
  if (needReorder) {
    var lastR = sheet.getLastRow();
    var dataOld = lastR > 1 ? sheet.getRange(2, 1, lastR - 1, oldHeaders.length).getValues() : [];
    var oldIdx = {};
    oldHeaders.forEach(function(h, i) { if (h && oldIdx[h] === undefined) oldIdx[h] = i; });

    var dataNew = dataOld.map(function(row) {
      var out = new Array(nextHeaders.length).fill('');
      nextHeaders.forEach(function(h, ni) {
        var oi = oldIdx[h];
        if (oi !== undefined && oi < row.length) out[ni] = row[oi];
      });
      return out;
    });

    sheet.getRange(1, 1, 1, nextHeaders.length).setValues([nextHeaders]);
    if (dataNew.length) {
      sheet.getRange(2, 1, dataNew.length, nextHeaders.length).setValues(dataNew);
    }
  }

  return sheet;
}

/**
 * สรุปคะแนนคุณลักษณะ
 */
function summarizeScores_(scores) {
  var nums = scores.filter(function(s) { return s !== '' && Number.isFinite(s); });
  if (nums.length !== 8) return { sum: '', avg: '', result: '' };
  var sum = nums.reduce(function(a, b) { return a + b; }, 0);
  var avg = sum / 8;
  var result = 'ปรับปรุง';
  if (avg >= 2.5) result = 'ดีเยี่ยม';
  else if (avg >= 2.0) result = 'ดี';
  else if (avg >= 1.0) result = 'ผ่าน';
  return { sum: sum, avg: avg, result: result };
}

/**
 * คำนวณผลการประเมินรายวิชา
 */
function calculateSubjectScoreResult_(scores) {
  if (scores.length !== 8 || !scores.every(function(s) { return s !== '' && s !== null && s !== undefined; })) {
    return '-';
  }
  var validScores = scores.map(Number);
  var count0 = validScores.filter(function(s) { return s === 0; }).length;
  var count3 = validScores.filter(function(s) { return s === 3; }).length;
  var count2up = validScores.filter(function(s) { return s >= 2; }).length;
  var hasBelow2 = validScores.some(function(s) { return s < 2; });

  if (count0 >= 2) return 'ไม่ผ่าน';
  if (count3 >= 6 && !hasBelow2) return 'ดีเยี่ยม';
  if (count2up >= 6 && count0 === 0) return 'ดี';
  return 'ผ่าน';
}


// ============================================================
// SECTION E: GET FUNCTIONS (อ่านข้อมูล)
// ============================================================

// ----- E1. Characteristic -----
function getStudentsForCharacteristic(grade, classNo) {
  var validated = validateGradeAndClass_(grade, classNo);
  grade = validated.grade;
  classNo = validated.classNo;

  var ss = SS();
  var studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  var data = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = data
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo && String(r[0]); })
    .map(function(r) {
      return {
        studentId: String(r[0]),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim(),
        grade: r[5],
        classNo: r[6]
      };
    });
  students.sort(function(a, b) { return String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true }); });

  var result = ensureCharSheetAndHeaders_();
  var sheet = result.sheet;
  var existingMap = new Map();
  if (sheet && sheet.getLastRow() > 1) {
    var all = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    all.forEach(function(row) {
      if (String(row[2]) === grade && String(row[3]) === classNo) {
        existingMap.set(String(row[0]), row.slice(4, 12).map(function(v) { return parseInt(v) || ''; }));
      }
    });
  }
  return students.map(function(s) {
    return Object.assign({}, s, { scores: existingMap.get(s.studentId) || Array(8).fill('') });
  });
}

// ----- E2. Activity -----
function getStudentsForActivity(grade, classNo) {
  var validated = validateGradeAndClass_(grade, classNo);
  grade = validated.grade;
  classNo = validated.classNo;

  var ss = SS();
  var studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  var data = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = data
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo && String(r[0]); })
    .map(function(r) {
      return {
        id: String(r[0]).trim(),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim(),
        grade: r[5],
        classNo: r[6]
      };
    });
  students.sort(function(a, b) { return String(a.id).localeCompare(String(b.id), undefined, { numeric: true }); });

  var sheet = S_getYearlySheet(ACTIVITY_SHEET);
  var map = new Map();
  if (sheet && sheet.getLastRow() > 1) {
    var ad = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    ad.filter(function(r) { return String(r[2]) === grade && String(r[3]) === classNo; })
      .forEach(function(r) {
        map.set(String(r[0]).trim(), {
          guidance: r[4] || 'ผ่าน', scout: r[5] || 'ผ่าน', club: r[6] || 'ผ่าน',
          social: r[7] || 'ผ่าน', overall: r[8] || 'ผ่าน'
        });
      });
  }

  return students.map(function(s) {
    var ext = map.get(s.id) || { guidance: 'ผ่าน', scout: 'ผ่าน', club: 'ผ่าน', social: 'ผ่าน', overall: 'ผ่าน' };
    return { studentId: s.id, name: s.name, grade: s.grade, classNo: s.classNo, activities: ext };
  });
}

// ----- E3. Competency -----
function getStudentsForCompetency(grade, classNo) {
  var validated = validateGradeAndClass_(grade, classNo);
  grade = validated.grade;
  classNo = validated.classNo;

  var ss = SS();
  var studentsSheet = ss.getSheetByName('Students');
  var data = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = data
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo; })
    .map(function(r) {
      return {
        id: String(r[0]).trim(),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim(),
        grade: r[5],
        classNo: r[6]
      };
    });
  students.sort(function(a, b) { return String(a.id).localeCompare(String(b.id)); });

  var sheet = ensureCompetencySheetAndHeaders_();
  var map = new Map();
  if (sheet && sheet.getLastRow() > 1) {
    var d = sheet.getRange(2, 1, sheet.getLastRow() - 1, 31).getValues();
    d.forEach(function(r) {
      if (String(r[2]) === grade && String(r[3]) === classNo) {
        map.set(String(r[0]), r.slice(4, 29));
      }
    });
  }
  return students.map(function(s) {
    return {
      studentId: s.id, name: s.name, grade: s.grade, classNo: s.classNo,
      scores: map.get(s.id) || Array(25).fill('')
    };
  });
}

// ----- E4. Subject Score -----
function getStudentsForSubjectScore(grade, classNo) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    grade = validated.grade;
    classNo = validated.classNo;

    var ss = SS();
    var studentsSheet = ss.getSheetByName('Students');
    if (!studentsSheet) throw new Error('ไม่พบชีต "Students"');

    var lastRow = studentsSheet.getLastRow();
    if (lastRow < 2) return [];
    var studentsData = studentsSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    var students = studentsData
      .filter(function(row) { return String(row[5] || '').trim() === grade && String(row[6] || '').trim() === classNo && String(row[0] || '').trim(); })
      .map(function(row) {
        return {
          studentId: String(row[0]).trim(),
          name: (String(row[2] || '').trim() + String(row[3] || '').trim() + ' ' + String(row[4] || '').trim()).trim(),
          grade: String(row[5]).trim(),
          classNo: String(row[6]).trim()
        };
      });
    students.sort(function(a, b) { return String(a.studentId).localeCompare(String(b.studentId)); });

    var result = ensureRTWSheetAndHeaders_();
    var sheet = result.sheet;
    var idx = result.idx;
    var scoreMap = new Map();

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      data.forEach(function(r) {
        if (String(r[idx['ชั้น'] - 1] || '').trim() === grade && String(r[idx['ห้อง'] - 1] || '').trim() === classNo) {
          var sid = String(r[idx['รหัสนักเรียน'] - 1] || '').trim();
          scoreMap.set(sid, {
            thai: r[idx['ภาษาไทย'] - 1] != null ? r[idx['ภาษาไทย'] - 1] : '',
            math: r[idx['คณิตศาสตร์'] - 1] != null ? r[idx['คณิตศาสตร์'] - 1] : '',
            science: r[idx['วิทยาศาสตร์'] - 1] != null ? r[idx['วิทยาศาสตร์'] - 1] : '',
            social: r[idx['สังคมศึกษา'] - 1] != null ? r[idx['สังคมศึกษา'] - 1] : '',
            health: r[idx['สุขศึกษา'] - 1] != null ? r[idx['สุขศึกษา'] - 1] : '',
            art: r[idx['ศิลปะ'] - 1] != null ? r[idx['ศิลปะ'] - 1] : '',
            work: r[idx['การงาน'] - 1] != null ? r[idx['การงาน'] - 1] : '',
            english: r[idx['ภาษาอังกฤษ'] - 1] != null ? r[idx['ภาษาอังกฤษ'] - 1] : ''
          });
        }
      });
    }

    return students.map(function(s) {
      var ex = scoreMap.get(s.studentId) || { thai: '', math: '', science: '', social: '', health: '', art: '', work: '', english: '' };
      return { studentId: s.studentId, name: s.name, grade: s.grade, classNo: s.classNo, scores: ex };
    });
  } catch (error) {
    Logger.log('Error in getStudentsForSubjectScore: ' + error.message);
    throw new Error('ไม่สามารถดึงข้อมูลได้: ' + error.message);
  }
}


// ============================================================
// SECTION F: SAVE FUNCTIONS (เขียนข้อมูล — ปรับปรุงแล้ว)
// ============================================================

// ----- F1. Characteristic -----
function saveCharacteristicAssessmentBatch(payload) {
  var records = Array.isArray(payload) ? payload : [];
  if (!records.length) throw new Error('ไม่มีข้อมูลสำหรับบันทึก');
  if (records.length > MAX_BATCH_SIZE) throw new Error('ส่งข้อมูลมากเกินไป (' + records.length + ' รายการ, สูงสุด ' + MAX_BATCH_SIZE + ')');

  var ss = SS();
  var validationErrors = [];

  var validatedRecords = records.map(function(rec, index) {
    var basicResult = validateRecordBasicFields_(rec, ss);
    if (!basicResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ': ' + basicResult.errors.join(', '));
      return null;
    }
    var scoresResult = validateCharacteristicScores_(rec.scores);
    if (!scoresResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ' (' + rec.studentId + '): ' + scoresResult.error);
      return null;
    }
    return {
      studentId: basicResult.sanitized.studentId,
      name: String(rec.name || '').trim().substring(0, 100),
      grade: basicResult.sanitized.grade,
      classNo: basicResult.sanitized.classNo,
      scores: scoresResult.sanitized
    };
  }).filter(Boolean);

  if (validationErrors.length > 0) {
    throw new Error('ข้อมูลไม่ถูกต้อง:\n' + validationErrors.join('\n'));
  }
  if (validatedRecords.length === 0) throw new Error('ไม่มีข้อมูลที่ถูกต้องสำหรับบันทึก');

  var grade = validatedRecords[0].grade;
  var classNo = validatedRecords[0].classNo;

  return withLock_('characteristic_' + grade + '_' + classNo, function() {
    var result = ensureCharSheetAndHeaders_();
    var sheet = result.sheet;
    var idx = result.idx;
    var lastC = sheet.getLastColumn();
    var now = new Date();
    var who = Session.getActiveUser().getEmail() || '';
    var col = function(h) { return idx[h] - 1; };

    var upsertRecords = validatedRecords.map(function(rec) {
      var scores = rec.scores;
      var summary = summarizeScores_(scores);

      var rowArr = new Array(lastC).fill('');
      rowArr[col('รหัสนักเรียน')] = rec.studentId;
      rowArr[col('ชื่อ-นามสกุล')] = rec.name;
      rowArr[col('ชั้น')] = rec.grade;
      rowArr[col('ห้อง')] = rec.classNo;
      TRAIT_HEADERS.forEach(function(h, i) { rowArr[col(h)] = scores[i]; });
      rowArr[col('คะแนนรวม')] = summary.sum === '' ? '' : String(summary.sum);
      rowArr[col('คะแนนเฉลี่ย')] = summary.avg === '' ? '' : summary.avg.toFixed(2);
      rowArr[col('ผลการประเมิน')] = summary.result || '';
      rowArr[col('วันที่บันทึก')] = now;
      rowArr[col('ผู้บันทึก')] = who;

      return { key: rec.studentId, rowData: rowArr };
    });

    var stats = atomicUpsertRows_(sheet, 'รหัสนักเรียน', 'ชั้น', 'ห้อง', grade, classNo, upsertRecords);
    return 'บันทึกสำเร็จ: อัปเดต ' + stats.updated + ' คน, เพิ่มใหม่ ' + stats.inserted + ' คน';
  });
}

// ----- F2. Activity -----
function saveActivityAssessmentBatch(payload) {
  var records = Array.isArray(payload) ? payload : [];
  if (!records.length) throw new Error('ไม่มีข้อมูล');
  if (records.length > MAX_BATCH_SIZE) throw new Error('ส่งข้อมูลมากเกินไป (' + records.length + ')');

  var ss = SS();
  var validationErrors = [];

  var validatedRecords = records.map(function(rec, index) {
    var basicResult = validateRecordBasicFields_(rec, ss);
    if (!basicResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ': ' + basicResult.errors.join(', '));
      return null;
    }
    var actResult = validateActivityResults_(rec.activities);
    if (!actResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ' (' + rec.studentId + '): ' + actResult.error);
      return null;
    }
    return {
      studentId: basicResult.sanitized.studentId,
      name: String(rec.name || '').trim().substring(0, 100),
      grade: basicResult.sanitized.grade,
      classNo: basicResult.sanitized.classNo,
      activities: actResult.sanitized
    };
  }).filter(Boolean);

  if (validationErrors.length > 0) throw new Error('ข้อมูลไม่ถูกต้อง:\n' + validationErrors.join('\n'));

  var grade = validatedRecords[0].grade;
  var classNo = validatedRecords[0].classNo;

  return withLock_('activity_' + grade + '_' + classNo, function() {
    var sheet = S_getYearlySheet(ACTIVITY_SHEET);
    if (!sheet) {
      var sheetName = (typeof S_sheetName === 'function') ? S_sheetName(ACTIVITY_SHEET) : ACTIVITY_SHEET;
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(ACTIVITY_HEADERS);
    }
    var now = new Date();
    var who = Session.getActiveUser().getEmail() || '';

    var upsertRecords = validatedRecords.map(function(rec) {
      var act = rec.activities;
      return {
        key: rec.studentId,
        rowData: [rec.studentId, rec.name, rec.grade, rec.classNo,
                  act.guidance, act.scout, act.club, act.social, act.overall,
                  now, who]
      };
    });

    var stats = atomicUpsertRows_(sheet, 'รหัสนักเรียน', 'ชั้น', 'ห้อง', grade, classNo, upsertRecords);
    return 'บันทึกสำเร็จ: อัปเดต ' + stats.updated + ' คน, เพิ่มใหม่ ' + stats.inserted + ' คน';
  });
}

// ----- F3. Subject Score -----
function saveSubjectScoreAssessmentBatch(payload) {
  var records = (payload && Array.isArray(payload.records)) ? payload.records : [];
  if (!records.length) throw new Error('ไม่มีข้อมูลสำหรับบันทึก');
  if (records.length > MAX_BATCH_SIZE) throw new Error('ส่งข้อมูลมากเกินไป (' + records.length + ')');

  var ss = SS();
  var validationErrors = [];

  var validatedRecords = records.map(function(rec, index) {
    var basicResult = validateRecordBasicFields_(rec, ss);
    if (!basicResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ': ' + basicResult.errors.join(', '));
      return null;
    }
    var scoresResult = validateSubjectScores_(rec.scores);
    if (!scoresResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ' (' + rec.studentId + '): ' + scoresResult.error);
      return null;
    }
    return {
      studentId: basicResult.sanitized.studentId,
      name: String(rec.name || '').trim().substring(0, 100),
      grade: basicResult.sanitized.grade,
      classNo: basicResult.sanitized.classNo,
      scores: scoresResult.sanitized
    };
  }).filter(Boolean);

  if (validationErrors.length > 0) throw new Error('ข้อมูลไม่ถูกต้อง:\n' + validationErrors.join('\n'));

  var grade = validatedRecords[0].grade;
  var classNo = validatedRecords[0].classNo;

  return withLock_('subjectscore_' + grade + '_' + classNo, function() {
    var result = ensureRTWSheetAndHeaders_();
    var sheet = result.sheet;
    var idx = result.idx;
    var lastC = sheet.getLastColumn();
    var now = new Date();
    var who = Session.getActiveUser().getEmail() || '';

    var upsertRecords = validatedRecords.map(function(rec) {
      var s = rec.scores;
      var scoresArr = [s.thai, s.math, s.science, s.social, s.health, s.art, s.work, s.english]
        .map(function(v) { var n = parseInt(v, 10); return Number.isFinite(n) ? n : ''; });

      var calcResult = calculateSubjectScoreResult_(scoresArr);

      var rowArr = new Array(lastC).fill('');
      rowArr[idx['รหัสนักเรียน'] - 1] = rec.studentId;
      rowArr[idx['ชื่อ-นามสกุล'] - 1] = rec.name;
      rowArr[idx['ชั้น'] - 1] = rec.grade;
      rowArr[idx['ห้อง'] - 1] = rec.classNo;
      rowArr[idx['ภาษาไทย'] - 1] = scoresArr[0];
      rowArr[idx['คณิตศาสตร์'] - 1] = scoresArr[1];
      rowArr[idx['วิทยาศาสตร์'] - 1] = scoresArr[2];
      rowArr[idx['สังคมศึกษา'] - 1] = scoresArr[3];
      rowArr[idx['สุขศึกษา'] - 1] = scoresArr[4];
      rowArr[idx['ศิลปะ'] - 1] = scoresArr[5];
      rowArr[idx['การงาน'] - 1] = scoresArr[6];
      rowArr[idx['ภาษาอังกฤษ'] - 1] = scoresArr[7];
      rowArr[idx['สรุปผลการประเมิน'] - 1] = calcResult;
      rowArr[idx['วันที่บันทึก'] - 1] = now;
      rowArr[idx['ผู้บันทึก'] - 1] = who;

      return { key: rec.studentId, rowData: rowArr };
    });

    var stats = atomicUpsertRows_(sheet, 'รหัสนักเรียน', 'ชั้น', 'ห้อง', grade, classNo, upsertRecords);
    return 'บันทึกข้อมูล ' + validatedRecords.length + ' คน สำเร็จ (อัปเดต ' + stats.updated + ', เพิ่มใหม่ ' + stats.inserted + ')';
  });
}

// ----- F4. Competency -----
function saveCompetencyAssessment(payload) {
  var records = Array.isArray(payload) ? payload : [];
  if (!records.length) throw new Error('ไม่มีข้อมูล');
  if (records.length > MAX_BATCH_SIZE) throw new Error('ส่งข้อมูลมากเกินไป (' + records.length + ')');

  var ss = SS();
  var validationErrors = [];

  var validatedRecords = records.map(function(rec, index) {
    var basicResult = validateRecordBasicFields_(rec, ss);
    if (!basicResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ': ' + basicResult.errors.join(', '));
      return null;
    }
    var scoresResult = validateCompetencyScores_(rec.scores);
    if (!scoresResult.valid) {
      validationErrors.push('รายการที่ ' + (index + 1) + ' (' + rec.studentId + '): ' + scoresResult.error);
      return null;
    }
    return {
      studentId: basicResult.sanitized.studentId,
      name: String(rec.name || '').trim().substring(0, 100),
      grade: basicResult.sanitized.grade,
      classNo: basicResult.sanitized.classNo,
      scores: scoresResult.sanitized
    };
  }).filter(Boolean);

  if (validationErrors.length > 0) throw new Error('ข้อมูลไม่ถูกต้อง:\n' + validationErrors.join('\n'));

  var grade = validatedRecords[0].grade;
  var classNo = validatedRecords[0].classNo;

  return withLock_('competency_' + grade + '_' + classNo, function() {
    var sheet = ensureCompetencySheetAndHeaders_();
    var now = new Date();
    var who = Session.getActiveUser().getEmail() || '';

    var upsertRecords = validatedRecords.map(function(rec) {
      return {
        key: rec.studentId,
        rowData: [rec.studentId, rec.name, rec.grade, rec.classNo].concat(rec.scores).concat([now, who])
      };
    });

    var stats = atomicUpsertRows_(sheet, 'รหัสนักเรียน', 'ชั้น', 'ห้อง', grade, classNo, upsertRecords);
    return 'บันทึกสำเร็จ: อัปเดต ' + stats.updated + ' คน, เพิ่มใหม่ ' + stats.inserted + ' คน';
  });
}


// ============================================================
// SECTION G: PDF GENERATION
// ============================================================

/**
 * ดึงโลโก้โรงเรียนเป็น Base64 Data URL
 */
function getSchoolLogoBase64_() {
  try {
    var settings = S_getGlobalSettings();
    var logoFileId = settings['logoFileId'];
    if (!logoFileId || logoFileId.trim() === '') {
      Logger.log('⚠️ ไม่มี logoFileId ในการตั้งค่า');
      return null;
    }
    Logger.log('🔍 กำลังโหลดโลโก้จาก fileId: ' + logoFileId);
    var file = DriveApp.getFileById(logoFileId.trim());
    var blob = file.getBlob();
    var mimeType = blob.getContentType() || 'image/png';
    var base64Data = Utilities.base64Encode(blob.getBytes());
    Logger.log('✅ โหลดโลโก้สำเร็จ');
    return 'data:' + mimeType + ';base64,' + base64Data;
  } catch (error) {
    Logger.log('❌ getSchoolLogoBase64_ error: ' + error.message);
    return null;
  }
}

/**
 * ดึงชื่อครูประจำชั้นจากชีต HomeroomTeachers
 */
function getHomeroomTeacher(grade, classNo) {
  try {
    var ss = SS();
    var sheet = ss.getSheetByName('HomeroomTeachers');
    if (!sheet) { Logger.log('⚠️ ไม่พบชีต HomeroomTeachers'); return '...'; }
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === grade && String(data[i][1]).trim() === classNo) {
        Logger.log('✅ พบครูประจำชั้น: ' + data[i][2]);
        return String(data[i][2]).trim();
      }
    }
    Logger.log('⚠️ ไม่พบครูประจำชั้นสำหรับ ' + grade + ' ห้อง ' + classNo);
    return '...';
  } catch (error) {
    Logger.log('❌ Error in getHomeroomTeacher: ' + error.message);
    return '...';
  }
}

// ════════════════════════════════════════════════════════════════
// ⚠️  WARNING: PDF LAYOUT LOCKED — DO NOT MODIFY WITHOUT TESTING
// ════════════════════════════════════════════════════════════════
// ฟังก์ชันที่ล็อค: generateCharacteristicAssessmentPDF,
//   generateActivityAssessmentPDF, generateSubjectScoreAssessmentPDF,
//   exportCompetencyReport, exportCompetencySummaryReport
// ════════════════════════════════════════════════════════════════

/**
 * ดึงข้อมูลที่ใช้ร่วมกันสำหรับสร้าง PDF
 */
function getPDFCommonData_(grade, classNo) {
  var settings = getGlobalSettings();
  var logoBase64 = null;
  try { logoBase64 = getSchoolLogoBase64_(); } catch (e) { Logger.log('⚠️ โลโก้: ' + e.message); }

  return {
    schoolName: settings['ชื่อโรงเรียน'] || 'โรงเรียน...',
    academicYear: settings['ปีการศึกษา'] || '2568',
    directorName: settings['ชื่อผู้อำนวยการ'] || '...',
    teacherName: getHomeroomTeacher(grade, classNo),
    logoBase64: logoBase64
  };
}

/**
 * สร้าง PDF จาก HTML แล้วคืน base64 data URL (ไม่ใช้ DriveApp)
 */
function _pdfToBase64_(htmlContent) {
  // Inject embedded Sarabun font for proper Thai font rendering in PDF
  var fontCss = _getEmbeddedSarabunCss_();
  if (fontCss && htmlContent.indexOf('<style>') > -1) {
    htmlContent = htmlContent.replace('<style>', '<style>' + fontCss);
  }
  var blob = HtmlService.createHtmlOutput(htmlContent).getBlob().getAs('application/pdf');
  return 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes());
}

/**
 * ดึงฟอนต์ Sarabun จาก Google Fonts แล้ว embed เป็น base64 @font-face CSS
 * แก้ปัญหาฟอนต์ไม่แสดงใน HtmlService PDF renderer
 */
function _getEmbeddedSarabunCss_() {
  try {
    var cssUrl = 'https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap';
    var resp = UrlFetchApp.fetch(cssUrl, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36' },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return '';

    var css = resp.getContentText();

    // Extract unique font file URLs
    var urls = [], m, re = /url\((https:\/\/fonts\.gstatic\.com\/[^)]+)\)/g;
    while ((m = re.exec(css)) !== null) {
      if (urls.indexOf(m[1]) === -1) urls.push(m[1]);
    }
    if (urls.length === 0) return '';

    // Fetch all font files in parallel
    var requests = urls.map(function(u) { return { url: u, muteHttpExceptions: true }; });
    var responses = UrlFetchApp.fetchAll(requests);

    responses.forEach(function(r, i) {
      if (r.getResponseCode() === 200) {
        var b64 = Utilities.base64Encode(r.getBlob().getBytes());
        var mime = urls[i].indexOf('.woff2') > -1 ? 'font/woff2' : 'font/woff';
        css = css.split(urls[i]).join('data:' + mime + ';base64,' + b64);
      }
    });

    return css;
  } catch(e) {
    return '';
  }
}

/**
 * บันทึก PDF ไปยัง Drive แล้วคืน URL (เก็บไว้สำหรับ backward compat)
 */
function savePDFToDrive_(htmlContent, folderName, fileName) {
  var blob = HtmlService.createHtmlOutput(htmlContent).getBlob().getAs('application/pdf').setName(fileName);
  var folder;
  var folders = DriveApp.getFoldersByName(folderName);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('✅ PDF สร้างสำเร็จ: ' + file.getUrl());
  return file.getUrl();
}

// ----- Sort Helper for PDF -----
function sortStudentsByMode_(students, sortMode) {
  if (!students || !Array.isArray(students)) return students;
  return students.slice().sort(function(a, b) {
    if (sortMode === 'gender') {
      var nameA = String(a.name || '');
      var nameB = String(b.name || '');
      var maleA = nameA.indexOf('เด็กชาย') !== -1 || nameA.indexOf('นาย') === 0;
      var maleB = nameB.indexOf('เด็กชาย') !== -1 || nameB.indexOf('นาย') === 0;
      if (maleA && !maleB) return -1;
      if (!maleA && maleB) return 1;
    }
    return String(a.studentId || '').localeCompare(String(b.studentId || ''), undefined, { numeric: true });
  });
}

// ----- G1. Characteristic PDF -----
function generateCharacteristicAssessmentPDF(grade, classNo, sortMode) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    var studentsData = getStudentsForCharacteristic(validated.grade, validated.classNo);
    if (!studentsData || studentsData.length === 0) throw new Error('ไม่พบข้อมูลนักเรียน');
    studentsData = sortStudentsByMode_(studentsData, sortMode || 'id');

    var common = getPDFCommonData_(validated.grade, validated.classNo);
    var htmlContent = createCharacteristicAssessmentHTML(
      studentsData, validated.grade, validated.classNo,
      common.academicYear, common.schoolName, common.directorName, common.teacherName, common.logoBase64
    );

    return _pdfToBase64_(htmlContent);
  } catch (error) {
    Logger.log('❌ Characteristic PDF Error: ' + error.message);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

// ----- G2. Activity PDF -----
function generateActivityAssessmentPDF(grade, classNo, sortMode) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    var studentsData = getStudentsForActivity(validated.grade, validated.classNo);
    if (!studentsData || studentsData.length === 0) throw new Error('ไม่พบข้อมูลนักเรียน');
    studentsData = sortStudentsByMode_(studentsData, sortMode || 'id');

    var common = getPDFCommonData_(validated.grade, validated.classNo);
    var htmlContent = createActivityAssessmentHTML_(
      studentsData, validated.grade, validated.classNo,
      common.academicYear, common.schoolName, common.directorName, common.teacherName, common.logoBase64
    );

    return _pdfToBase64_(htmlContent);
  } catch (error) {
    Logger.log('❌ Activity PDF Error: ' + error.message);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

// ----- G3. Subject Score PDF (ปพ.5) -----
function generateSubjectScoreAssessmentPDF(grade, classNo, sortMode) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    var studentsData = getStudentsForSubjectScore(validated.grade, validated.classNo);
    if (!studentsData || studentsData.length === 0) throw new Error('ไม่พบข้อมูลนักเรียน');
    studentsData = sortStudentsByMode_(studentsData, sortMode || 'id');

    var common = getPDFCommonData_(validated.grade, validated.classNo);
    var htmlContent = createOfficialReportHTML(
      studentsData, validated.grade, validated.classNo,
      common.academicYear, common.schoolName, common.directorName, common.teacherName, common.logoBase64
    );

    return _pdfToBase64_(htmlContent);
  } catch (error) {
    Logger.log('❌ ปพ.5 PDF Error: ' + error.message);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

// ----- G4. Competency Report PDF -----
function exportCompetencyReport(grade, classNo, sortMode) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    var studentsData = getStudentsForCompetency(validated.grade, validated.classNo);
    if (!studentsData || studentsData.length === 0) throw new Error('ไม่พบข้อมูลนักเรียน');
    studentsData = sortStudentsByMode_(studentsData, sortMode || 'id');

    var common = getPDFCommonData_(validated.grade, validated.classNo);
    var htmlContent = createCompetencyDetailedHTML(
      studentsData, validated.grade, validated.classNo,
      common.academicYear, common.schoolName, common.directorName, common.teacherName, common.logoBase64
    );

    return _pdfToBase64_(htmlContent);
  } catch (error) {
    Logger.log('❌ Competency Report Error: ' + error.message);
    throw new Error('ไม่สามารถสร้างรายงานได้: ' + error.message);
  }
}

// ----- G5. Competency Summary PDF -----
function exportCompetencySummaryReport(grade, classNo, sortMode) {
  try {
    var validated = validateGradeAndClass_(grade, classNo);
    var studentsData = getStudentsForCompetency(validated.grade, validated.classNo);
    if (!studentsData || studentsData.length === 0) throw new Error('ไม่พบข้อมูลนักเรียน');
    studentsData = sortStudentsByMode_(studentsData, sortMode || 'id');

    var common = getPDFCommonData_(validated.grade, validated.classNo);
    var htmlContent = createCompetencySummaryHTML(
      studentsData, validated.grade, validated.classNo,
      common.academicYear, common.schoolName, common.directorName, common.teacherName, common.logoBase64
    );

    return _pdfToBase64_(htmlContent);
  } catch (error) {
    Logger.log('❌ Competency Summary Error: ' + error.message);
    throw new Error('ไม่สามารถสร้างรายงานได้: ' + error.message);
  }
}


// ============================================================
// SECTION H: HTML TEMPLATES
// ============================================================

// ----- H1. Characteristic HTML -----
function createCharacteristicAssessmentHTML(students, grade, classNo, year, school, director, teacherName, logoBase64) {
  var logoHtml = logoBase64 ? '<img src="' + logoBase64 + '" class="logo" alt="โลโก้โรงเรียน">' : '';

  var excellentCount = 0, goodCount = 0, passCount = 0, improveCount = 0;
  var rows = "";
  students.forEach(function(stu, i) {
    var scores = stu.scores || Array(8).fill('');
    var validScores = scores.filter(function(s) { return s !== '' && !isNaN(s); }).map(Number);
    var total = validScores.length === 8 ? validScores.reduce(function(a, b) { return a + b; }, 0) : '';
    var avg = validScores.length === 8 ? (total / 8).toFixed(2) : '';
    var result = '';
    if (avg !== '') {
      var avgNum = parseFloat(avg);
      if (avgNum >= 2.5) { result = 'ดีเยี่ยม'; excellentCount++; }
      else if (avgNum >= 2.0) { result = 'ดี'; goodCount++; }
      else if (avgNum >= 1.0) { result = 'ผ่าน'; passCount++; }
      else { result = 'ปรับปรุง'; improveCount++; }
    }
    rows += '<tr><td class="center">' + (i + 1) + '</td><td class="left" style="padding-left:5px;">' + stu.name + '</td>';
    scores.forEach(function(s) { rows += '<td class="center">' + (s !== '' ? s : '-') + '</td>'; });
    rows += '<td class="center">' + (total !== '' ? total : '-') + '</td>';
    rows += '<td class="center">' + (avg !== '' ? avg : '-') + '</td>';
    rows += '<td class="center">' + result + '</td></tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet"><style>'
    + '@page{size:A4 portrait;margin:1.5cm 1cm}body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;line-height:1.2;color:#000}.header{text-align:center;margin-bottom:8px}.logo{width:50px;height:50px;object-fit:contain;margin-bottom:5px}.title{font-size:14pt;font-weight:700}.school-name{font-size:12pt;text-align:center;margin-top:3px}.info-line{text-align:center;margin:5px 0 10px 0;font-size:11pt}table{width:100%;border-collapse:collapse;margin-bottom:10px}th,td{border:1px solid #000;padding:3px 2px;font-size:10pt;font-weight:300;vertical-align:middle}th{background-color:#fef9e7;font-weight:700;color:#000}.vertical-text{writing-mode:vertical-rl;text-orientation:mixed;height:140px;white-space:nowrap;padding:5px 2px}.center{text-align:center}.left{text-align:left}.font-bold{font-weight:bold}'
    + '.summary-box{margin-top:15px;padding:8px;border:1px solid #000;background-color:#f8f9fa}.summary-row{display:flex;justify-content:space-around;margin:5px 0}.summary-item{text-align:center;flex:1}.summary-label{font-size:10pt;margin-bottom:3px}.summary-value{font-size:12pt;font-weight:700}.excellent{color:#28a745}.good{color:#007bff}.pass{color:#ffc107}.improve{color:#dc3545}'
    + '.legend{margin-top:8px;font-size:9pt;text-align:center}.footer{margin-top:25px;width:100%;display:table;page-break-inside:avoid}.sign-left,.sign-right{display:table-cell;width:50%;text-align:center;vertical-align:top}.sign-line{margin-bottom:4px;font-size:10pt}'
    + '</style></head><body>'
    + '<div class="header">' + logoHtml + '<div class="title">รายงานการประเมินคุณลักษณะอันพึงประสงค์</div></div>'
    + '<div class="school-name">' + school + '</div>'
    + '<div class="info-line">ชั้น ' + grade + ' ห้อง ' + classNo + ' &nbsp;&nbsp; ปีการศึกษา ' + year + '</div>'
    + '<table><thead><tr><th rowspan="2" style="width:4%">ที่</th><th rowspan="2" style="width:25%">ชื่อ - นามสกุล</th>'
    + '<th class="vertical-text">รักชาติ ศาสน์ กษัตริย์</th><th class="vertical-text">ซื่อสัตย์สุจริต</th><th class="vertical-text">มีวินัย</th><th class="vertical-text">ใฝ่เรียนรู้</th><th class="vertical-text">อยู่อย่างพอเพียง</th><th class="vertical-text">มุ่งมั่นในการทำงาน</th><th class="vertical-text">รักความเป็นไทย</th><th class="vertical-text">มีจิตสาธารณะ</th>'
    + '<th rowspan="2" style="width:5%">รวม</th><th rowspan="2" style="width:5%">เฉลี่ย</th><th rowspan="2" style="width:8%">ผลการประเมิน</th></tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="legend"><strong>หมายเหตุ:</strong> 3 = ดีเยี่ยม, 2 = ดี, 1 = ผ่าน, 0 = ปรับปรุง &nbsp;|&nbsp; <strong>ผลประเมิน:</strong> ดีเยี่ยม ≥ 2.50, ดี ≥ 2.00, ผ่าน ≥ 1.00, ปรับปรุง < 1.00</div>'
    + '<div class="summary-box"><div style="text-align:center;font-weight:bold;margin-bottom:8px;">สรุปผลการประเมิน</div><div class="summary-row"><div class="summary-item"><div class="summary-label">ดีเยี่ยม</div><div class="summary-value excellent">' + excellentCount + '</div></div><div class="summary-item"><div class="summary-label">ดี</div><div class="summary-value good">' + goodCount + '</div></div><div class="summary-item"><div class="summary-label">ผ่าน</div><div class="summary-value pass">' + passCount + '</div></div><div class="summary-item"><div class="summary-label">ปรับปรุง</div><div class="summary-value improve">' + improveCount + '</div></div></div></div>'
    + '<div class="footer"><div class="sign-left"><div class="sign-line">ลงชื่อ...............................................ครูประจำชั้น</div><div class="sign-line">(' + teacherName + ')</div><div class="sign-line">วันที่........./........./..........</div></div><div class="sign-right"><div class="sign-line">ลงชื่อ...............................................ผู้อำนวยการ</div><div class="sign-line">(' + director + ')</div><div class="sign-line">วันที่........./........./..........</div></div></div>'
    + '</body></html>';
}

// ----- H2. Activity HTML -----
function createActivityAssessmentHTML_(students, grade, classNo, year, school, director, teacherName, logoBase64) {
  var logoHtml = (logoBase64 && logoBase64.trim() !== '') ? '<img src="' + logoBase64 + '" class="logo" alt="โลโก้โรงเรียน">' : '';
  var passAllCount = 0, notPassCount = 0;

  var rows = "";
  students.forEach(function(stu, i) {
    var act = stu.activities || {};
    var guidance = act.guidance || '-';
    var scout = act.scout || '-';
    var club = act.club || '-';
    var social = act.social || '-';
    var overall = act.overall || '-';
    if (overall === 'ผ่าน') passAllCount++;
    else if (overall === 'ไม่ผ่าน') notPassCount++;

    var getRC = function(val) { return val === 'ผ่าน' ? 'pass' : (val === 'ไม่ผ่าน' ? 'not-pass' : ''); };
    rows += '<tr><td class="center">' + (i + 1) + '</td><td class="left name-cell">' + stu.name + '</td>'
      + '<td class="center ' + getRC(guidance) + '">' + guidance + '</td>'
      + '<td class="center ' + getRC(scout) + '">' + scout + '</td>'
      + '<td class="center ' + getRC(club) + '">' + club + '</td>'
      + '<td class="center ' + getRC(social) + '">' + social + '</td>'
      + '<td class="center ' + getRC(overall) + '">' + overall + '</td></tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap" rel="stylesheet"><style>'
    + '@page{size:A4;margin:1cm 1.2cm}body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;line-height:1.2;color:#000;margin:0;padding:0}.header{text-align:center;margin-bottom:8px}.logo{width:50px;height:50px;object-fit:contain;margin-bottom:5px}.title{font-size:14pt;font-weight:700;margin-bottom:3px}.subtitle{font-size:12pt;font-weight:600}.info-box{text-align:center;margin-bottom:10px;font-size:12pt;font-weight:400}table{width:100%;border-collapse:collapse;margin-bottom:8px}th,td{border:1px solid #000;padding:4px 3px;font-size:10pt;font-weight:300;vertical-align:middle}th{background-color:#000;color:white;font-weight:700;height:30px}.center{text-align:center}.left{text-align:left;padding-left:5px!important}.font-bold{font-weight:600}.name-cell{font-size:9pt;white-space:nowrap}.pass{color:#000}.not-pass{color:#dc3545}'
    + '.summary-inline{margin:10px 0;padding:8px 15px;border:1.5px solid #000;border-radius:5px;background-color:#f8f9fa;text-align:center;font-size:10pt}.summary-inline .value{font-weight:700;font-size:12pt;margin:0 3px}.legend-inline{margin:8px 0;font-size:9pt;text-align:center}'
    + '.footer{margin-top:20px;width:100%}.footer table{border:none;margin:0}.footer td{border:none;width:50%;text-align:center;vertical-align:top;padding:0 10px}.sign-line{margin-bottom:5px;font-size:10pt}'
    + '</style></head><body>'
    + '<div class="header">' + logoHtml + '<div class="title">รายงานการประเมินกิจกรรมพัฒนาผู้เรียน</div><div class="subtitle">' + school + '</div></div>'
    + '<div class="info-box">ชั้นประถมศึกษาปีที่ ' + grade.replace('ป.', '') + ' ห้อง ' + classNo + ' &nbsp;&nbsp; ปีการศึกษา ' + year + '</div>'
    + '<table><thead><tr><th style="width:5%">ที่</th><th style="width:32%">ชื่อ - นามสกุล</th><th style="width:11%">แนะแนว</th><th style="width:13%">ลูกเสือ/เนตรนารี</th><th style="width:11%">ชุมนุม</th><th style="width:13%">เพื่อสังคมฯ</th><th style="width:11%">สรุปผล</th></tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="legend-inline"><strong>เกณฑ์:</strong> &nbsp;<span class="pass"><strong>ผ่าน</strong></span> = เข้าร่วมกิจกรรมครบตามเกณฑ์ &nbsp;&nbsp;|&nbsp;&nbsp;<span class="not-pass"><strong>ไม่ผ่าน</strong></span> = เข้าร่วมกิจกรรมไม่ครบตามเกณฑ์</div>'
    + '<div class="summary-inline"><strong>สรุปผลการประเมิน:</strong> &nbsp;&nbsp;ผ่าน <span class="value pass">' + passAllCount + '</span> คน &nbsp;&nbsp;|&nbsp;&nbsp;ไม่ผ่าน <span class="value not-pass">' + notPassCount + '</span> คน &nbsp;&nbsp;|&nbsp;&nbsp;รวมทั้งหมด <span class="value total">' + students.length + '</span> คน</div>'
    + '<div class="footer"><table><tr><td><div class="sign-line">ลงชื่อ...............................................ครูประจำชั้น</div><div class="sign-line">(' + (teacherName || '..........................................') + ')</div><div class="sign-line">ตำแหน่ง ครู &nbsp;&nbsp; วันที่........./........./..........</div></td>'
    + '<td><div class="sign-line">ลงชื่อ...............................................ผู้อำนวยการ</div><div class="sign-line">(' + director + ')</div><div class="sign-line">ผู้อำนวยการสถานศึกษา &nbsp;&nbsp; วันที่........./........./..........</div></td></tr></table></div>'
    + '</body></html>';
}

// ----- H3. Subject Score HTML (ปพ.5) -----
function createOfficialReportHTML(students, grade, classNo, year, school, director, teacherName, logoBase64) {
  var logoHtml = (logoBase64 && logoBase64.trim() !== '') ? '<img src="' + logoBase64 + '" class="logo" alt="โลโก้โรงเรียน">' : '';

  var rows = "";
  students.forEach(function(stu, i) {
    var s = stu.scores;
    var scoresArr = [s.thai, s.math, s.science, s.social, s.health, s.art, s.work, s.english].map(Number);
    var result = calculateSubjectScoreResult_(scoresArr);
    rows += '<tr><td class="center">' + (i + 1) + '</td><td class="center">' + stu.studentId + '</td><td class="left name-cell" style="padding-left:5px;">' + stu.name + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.thai)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.math)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.science)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.social)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.health)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.art)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.work)) || '-') + '</td>'
      + '<td class="center">' + (Math.floor(Number(s.english)) || '-') + '</td>'
      + '<td class="center">' + result + '</td></tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;700&display=swap" rel="stylesheet"><style>'
    + '@page{size:A4;margin:1.5cm 1.2cm}body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;line-height:1.2;color:#000}.header{text-align:center;margin-bottom:10px}.logo{width:50px;height:50px;object-fit:contain;margin-bottom:5px}.title{font-size:14pt;font-weight:700}.subtitle{font-size:12pt;font-weight:600;margin-top:3px}.info-box{text-align:center;margin-bottom:10px;font-size:12pt;font-weight:600}table{width:100%;border-collapse:collapse;margin-bottom:15px}th,td{border:1px solid #000;padding:3px 2px;font-size:10pt;font-weight:300;vertical-align:middle}th{background-color:#f0f0f0;font-weight:700;height:30px}.center{text-align:center}.left{text-align:left}.font-bold{font-weight:bold}.name-cell{font-size:10pt;white-space:nowrap}'
    + '.footer{margin-top:30px;width:100%;display:table;page-break-inside:avoid}.sign-left,.sign-right{display:table-cell;width:50%;text-align:center;vertical-align:top}.sign-line{margin-bottom:4px;font-size:10pt;white-space:nowrap}'
    + '</style></head><body>'
    + '<div class="header">' + logoHtml + '<div class="title">แบบประเมินการอ่าน คิดวิเคราะห์ เขียน</div><div class="subtitle">' + school + '</div></div>'
    + '<div class="info-box">ชั้นประถมศึกษาปีที่ ' + grade.replace('ป.', '') + ' ห้อง ' + classNo + ' &nbsp;&nbsp; ปีการศึกษา ' + year + '</div>'
    + '<table><thead><tr><th rowspan="2" style="width:5%">ที่</th><th rowspan="2" style="width:12%">รหัส</th><th rowspan="2" style="width:25%">ชื่อ - นามสกุล</th><th colspan="8">ผลการเรียนรายวิชา (เกรด 1-4)</th><th rowspan="2" style="width:12%">สรุปผล</th></tr><tr><th>ไทย</th><th>คณิต</th><th>วิทย์</th><th>สังคม</th><th>สุขะ</th><th>ศิลปะ</th><th>การงาน</th><th>อังกฤษ</th></tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="footer"><div class="sign-left"><div class="sign-line">ลงชื่อ..........................................................ครูประจำชั้น</div><div class="sign-line">(' + teacherName + ')</div><div class="sign-line">ตำแหน่ง ครู</div></div><div class="sign-right"><div class="sign-line">ลงชื่อ..........................................................ผู้อำนวยการ</div><div class="sign-line">(' + director + ')</div><div class="sign-line">ผู้อำนวยการสถานศึกษา</div></div></div>'
    + '</body></html>';
}

// ----- H4. Competency Detailed HTML -----
function createCompetencyDetailedHTML(students, grade, classNo, year, school, director, teacherName, logoBase64) {
  var logoHtml = logoBase64 ? '<img src="' + logoBase64 + '" class="logo" alt="โลโก้โรงเรียน">' : '';
  var rows = "";
  students.forEach(function(stu, i) {
    var scores = stu.scores || Array(25).fill('');
    var summaryData = [];
    for (var c = 0; c < 5; c++) {
      var start = c * 5;
      var partScores = scores.slice(start, start + 5).map(function(s) { return parseInt(s) || 0; });
      var sum = partScores.reduce(function(a, b) { return a + b; }, 0);
      summaryData.push({ sum: sum, avg: (sum / 5).toFixed(2) });
    }
    var allScores = scores.map(function(s) { return parseInt(s) || 0; });
    var totalSum = allScores.reduce(function(a, b) { return a + b; }, 0);
    var totalAvg = (totalSum / 25).toFixed(2);
    var result = 'ปรับปรุง';
    if (parseFloat(totalAvg) >= 2.5) result = 'ดีเยี่ยม';
    else if (parseFloat(totalAvg) >= 2.0) result = 'ดี';
    else if (parseFloat(totalAvg) >= 1.0) result = 'ผ่าน';

    rows += '<tr><td class="center">' + (i + 1) + '</td><td class="left" style="padding-left:5px;">' + stu.name + '</td>';
    summaryData.forEach(function(sd) { rows += '<td class="center">' + sd.avg + '</td>'; });
    rows += '<td class="center">' + totalSum + '</td><td class="center">' + totalAvg + '</td><td class="center">' + result + '</td></tr>';
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet"><style>'
    + '@page{size:A4 landscape;margin:1.5cm}body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;line-height:1.2;color:#000}.header{text-align:center;margin-bottom:8px}.logo{width:45px;height:45px;object-fit:contain;vertical-align:middle;margin-right:8px}.title{font-size:14pt;font-weight:700;display:inline;vertical-align:middle}.school-name{font-size:12pt;text-align:center;margin-top:3px}.info-line{text-align:center;margin:5px 0 10px 0;font-size:11pt}table{width:100%;border-collapse:collapse;margin-bottom:8px}th,td{border:1px solid #000;padding:3px 2px;font-size:10pt;font-weight:300;vertical-align:middle}th{background-color:#f0f0f0;font-weight:700}.center{text-align:center}.left{text-align:left}.font-bold{font-weight:bold}'
    + '.legend{margin-top:8px;font-size:9pt}.footer{margin-top:30px;width:100%}.footer table{border:none}.footer td{border:none;width:50%;text-align:center;vertical-align:top}.sign-line{margin-bottom:4px;font-size:10pt}'
    + '</style></head><body>'
    + '<div class="header">' + logoHtml + '<span class="title">รายงานการประเมินสมรรถนะสำคัญของผู้เรียน</span></div>'
    + '<div class="school-name">' + school + '</div>'
    + '<div class="info-line">ชั้น ' + grade + ' ห้อง ' + classNo + ' &nbsp;&nbsp; ปีการศึกษา ' + year + '</div>'
    + '<table><thead><tr><th rowspan="2" style="width:4%">ที่</th><th rowspan="2" style="width:20%">ชื่อ - นามสกุล</th><th colspan="5" style="background-color:#d4edda;">สมรรถนะ 5 ด้าน (คะแนนเฉลี่ย)</th><th rowspan="2" style="width:6%">รวม</th><th rowspan="2" style="width:6%">เฉลี่ย</th><th rowspan="2" style="width:9%">ผลประเมิน</th></tr><tr><th>การสื่อสาร</th><th>การคิด</th><th>แก้ปัญหา</th><th>ทักษะชีวิต</th><th>เทคโนโลยี</th></tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="legend"><strong>เกณฑ์การให้คะแนน:</strong> ดีเยี่ยม = 3 &nbsp; ดี = 2 &nbsp; พอใช้ = 1 &nbsp; ปรับปรุง = 0 &nbsp;&nbsp;|&nbsp;&nbsp; <strong>เกณฑ์ผลประเมิน:</strong> ดีเยี่ยม ≥ 2.50 &nbsp; ดี ≥ 2.00 &nbsp; ผ่าน ≥ 1.00 &nbsp; ปรับปรุง < 1.00</div>'
    + '<div class="footer"><table><tr><td><div class="sign-line">ลงชื่อ.......................................................ครูประจำชั้น</div><div class="sign-line">(' + (teacherName || '..............................................................') + ')</div><div class="sign-line">ตำแหน่ง ครู &nbsp;&nbsp; วันที่........./........./..........</div></td>'
    + '<td><div class="sign-line">ลงชื่อ.......................................................ผู้อำนวยการ</div><div class="sign-line">(' + director + ')</div><div class="sign-line">ผู้อำนวยการสถานศึกษา &nbsp;&nbsp; วันที่........./........./..........</div></td></tr></table></div>'
    + '</body></html>';
}

// ----- H5. Competency Summary HTML -----
function createCompetencySummaryHTML(students, grade, classNo, year, school, director, teacherName, logoBase64) {
  var competencyNames = ['การสื่อสาร', 'การคิด', 'การแก้ปัญหา', 'ทักษะชีวิต', 'เทคโนโลยี'];
  var logoHtml = logoBase64 ? '<img src="' + logoBase64 + '" class="logo" alt="โลโก้โรงเรียน">' : '';

  var stats = competencyNames.map(function(name, c) {
    var excellent = 0, good = 0, pass = 0, improve = 0;
    students.forEach(function(stu) {
      var scores = stu.scores || Array(25).fill('');
      var start = c * 5;
      var partScores = scores.slice(start, start + 5).map(function(s) { return parseInt(s) || 0; });
      var avg = partScores.reduce(function(a, b) { return a + b; }, 0) / 5;
      if (avg >= 2.5) excellent++;
      else if (avg >= 2.0) good++;
      else if (avg >= 1.0) pass++;
      else improve++;
    });
    return { name: name, excellent: excellent, good: good, pass: pass, improve: improve };
  });

  var overallExcellent = 0, overallGood = 0, overallPass = 0, overallImprove = 0;
  students.forEach(function(stu) {
    var scores = stu.scores || Array(25).fill('');
    var allScores = scores.map(function(s) { return parseInt(s) || 0; });
    var totalAvg = allScores.reduce(function(a, b) { return a + b; }, 0) / 25;
    if (totalAvg >= 2.5) overallExcellent++;
    else if (totalAvg >= 2.0) overallGood++;
    else if (totalAvg >= 1.0) overallPass++;
    else overallImprove++;
  });

  var statsRows = stats.map(function(s, i) {
    return '<tr><td class="center">' + (i + 1) + '</td><td class="left">' + s.name + '</td><td class="center">' + s.excellent + '</td><td class="center">' + s.good + '</td><td class="center">' + s.pass + '</td><td class="center">' + s.improve + '</td><td class="center font-bold">' + students.length + '</td></tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet"><style>'
    + '@page{size:A4;margin:1.5cm 1.2cm}body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;line-height:1.2;color:#000}.header{text-align:center;margin-bottom:8px}.logo{width:50px;height:50px;object-fit:contain;vertical-align:middle;margin-right:8px}.title{font-size:14pt;font-weight:700;display:inline;vertical-align:middle}.school-name{font-size:12pt;text-align:center;margin-top:3px}.info-line{text-align:center;margin:8px 0 12px 0;font-size:12pt}table{width:100%;border-collapse:collapse;margin-bottom:12px}th,td{border:1px solid #000;padding:5px 4px;font-size:11pt;font-weight:300;vertical-align:middle}th{background-color:#f0f0f0;font-weight:700}.center{text-align:center}.left{text-align:left;padding-left:8px}.font-bold{font-weight:bold}.total-row{background-color:#e8f5e9}'
    + '.summary-line{text-align:center;margin:15px 0;font-size:12pt}.summary-item{display:inline-block;margin:0 20px}.summary-label{font-weight:300}.summary-value{font-weight:700;font-size:14pt;margin-left:5px}.excellent{color:#28a745}.good{color:#007bff}.pass{color:#ff9800}.improve{color:#dc3545}'
    + '.footer{margin-top:30px;width:100%}.footer table{border:none}.footer td{border:none;width:50%;text-align:center;vertical-align:top}.sign-line{margin-bottom:5px;font-size:10pt;white-space:nowrap}'
    + '</style></head><body>'
    + '<div class="header">' + logoHtml + '<span class="title">รายงานสรุปผลการประเมินสมรรถนะสำคัญของผู้เรียน</span></div>'
    + '<div class="school-name">' + school + '</div>'
    + '<div class="info-line">ชั้น ' + grade + ' ห้อง ' + classNo + ' &nbsp;&nbsp; ปีการศึกษา ' + year + ' &nbsp;&nbsp; จำนวนนักเรียน ' + students.length + ' คน</div>'
    + '<table><thead><tr><th style="width:8%">ที่</th><th style="width:30%">สมรรถนะ</th><th style="width:12%">ดีเยี่ยม</th><th style="width:12%">ดี</th><th style="width:12%">ผ่าน</th><th style="width:12%">ปรับปรุง</th><th style="width:14%">รวม</th></tr></thead><tbody>' + statsRows
    + '<tr class="total-row"><td colspan="2" class="center font-bold">รวมทั้งหมด (เฉลี่ย 5 ด้าน)</td><td class="center font-bold">' + overallExcellent + '</td><td class="center font-bold">' + overallGood + '</td><td class="center font-bold">' + overallPass + '</td><td class="center font-bold">' + overallImprove + '</td><td class="center font-bold">' + students.length + '</td></tr></tbody></table>'
    + '<div class="summary-line"><strong>ภาพรวมผลการประเมิน:</strong>'
    + '<span class="summary-item"><span class="summary-label">ดีเยี่ยม</span><span class="summary-value excellent">' + overallExcellent + '</span></span>'
    + '<span class="summary-item"><span class="summary-label">ดี</span><span class="summary-value good">' + overallGood + '</span></span>'
    + '<span class="summary-item"><span class="summary-label">ผ่าน</span><span class="summary-value pass">' + overallPass + '</span></span>'
    + '<span class="summary-item"><span class="summary-label">ปรับปรุง</span><span class="summary-value improve">' + overallImprove + '</span></span></div>'
    + '<div class="footer"><table><tr><td><div class="sign-line">ลงชื่อ...................................................ครูประจำชั้น</div><div class="sign-line">(' + (teacherName || '.............................................................') + ')</div><div class="sign-line">ตำแหน่ง ครู &nbsp;&nbsp; วันที่........./........./..........</div></td>'
    + '<td><div class="sign-line">ลงชื่อ...................................................ผู้อำนวยการ</div><div class="sign-line">(' + director + ')</div><div class="sign-line">ผู้อำนวยการสถานศึกษา &nbsp;&nbsp; วันที่........./........./..........</div></td></tr></table></div>'
    + '</body></html>';
}
