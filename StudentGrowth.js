// ============================================================
// StudentGrowth.js - student weight/height tracking
// ============================================================

const SG_RECORDS_SHEET = 'StudentGrowthRecords';
const SG_CRITERIA_SHEET = 'GrowthCriteria';

const SG_RECORD_HEADERS = [
  'record_id', 'academic_year', 'round_id', 'month', 'round_no', 'measure_date',
  'student_id', 'student_name', 'gender', 'birthdate', 'age_years', 'grade',
  'class_no', 'weight_kg', 'height_cm', 'bmi', 'nutrition_status',
  'height_status', 'weight_height_status', 'recorder', 'note', 'created_at',
  'updated_at'
];

const SG_CRITERIA_HEADERS = [
  'type', 'status', 'min_value', 'max_value', 'label', 'sort_order', 'active', 'note'
];

function SG_nowIso_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function SG_academicYear_() {
  try {
    if (typeof S_getAcademicYear === 'function') return String(S_getAcademicYear());
    if (typeof S_getCurrentAcademicYear_ === 'function') return String(S_getCurrentAcademicYear_());
  } catch (_e) {}
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth() + 1;
  return String(m >= 5 ? y + 543 : y + 542);
}

function SG_month_() {
  var now = new Date();
  return String(now.getMonth() + 1);
}

function SG_user_() {
  try {
    var session = getLoginSession();
    if (session && session.username) return session.username;
  } catch (_e) {}
  try {
    return Session.getActiveUser().getEmail() || 'system';
  } catch (_e2) {
    return 'system';
  }
}

function SG_normalize_(value) {
  return String(value || '').toLowerCase().trim();
}

function SG_headerMap_(headers) {
  var map = {};
  headers.forEach(function(h, i) {
    map[SG_normalize_(h)] = i;
  });
  return map;
}

function SG_col_(map, names, fallback) {
  for (var i = 0; i < names.length; i++) {
    var key = SG_normalize_(names[i]);
    if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  }
  return typeof fallback === 'number' ? fallback : -1;
}

function SG_ensureSheet_(name, headers) {
  var ss = SS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  SG_ensureHeaders_(sheet, headers);
  return sheet;
}

function SG_ensureHeaders_(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) lastCol = 1;
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });
  var hasAny = current.some(function(v) { return v; });
  if (!hasAny) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var missing = headers.filter(function(h) { return current.indexOf(h) === -1; });
    if (missing.length) {
      sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    }
  }
  var width = Math.max(headers.length, sheet.getLastColumn());
  sheet.getRange(1, 1, 1, width)
    .setBackground('#4a5568')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function SG_ensureStudentsLatestColumns_() {
  var sheet = SS().getSheetByName('Students');
  if (!sheet) return null;
  SG_ensureHeaders_(sheet, [
    'student_id', 'id_card', 'title', 'firstname', 'lastname', 'grade', 'class_no',
    'gender', 'birthdate', 'photo_url', 'academic_year', 'weight', 'height',
    'blood_type', 'religion', 'father_name', 'father_lastname', 'father_occupation',
    'mother_name', 'mother_lastname', 'mother_occupation', 'address', 'created_date',
    'status', 'latest_weight_kg', 'latest_height_cm', 'latest_bmi',
    'latest_growth_date', 'latest_nutrition_status'
  ]);
  return sheet;
}

function SG_seedCriteria_() {
  var sheet = SG_ensureSheet_(SG_CRITERIA_SHEET, SG_CRITERIA_HEADERS);
  if (sheet.getLastRow() > 1) return;
  var rows = [
    ['bmi_basic', 'underweight', '', '18.49', 'น้ำหนักน้อย', 1, 'TRUE', 'ค่าเริ่มต้นแก้ไขได้'],
    ['bmi_basic', 'normal', '18.5', '22.99', 'ปกติ', 2, 'TRUE', 'ค่าเริ่มต้นแก้ไขได้'],
    ['bmi_basic', 'overweight', '23', '24.99', 'น้ำหนักเกิน', 3, 'TRUE', 'ค่าเริ่มต้นแก้ไขได้'],
    ['bmi_basic', 'obese', '25', '', 'อ้วน', 4, 'TRUE', 'ค่าเริ่มต้นแก้ไขได้']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function SG_getCriteria_() {
  SG_seedCriteria_();
  var sheet = SS().getSheetByName(SG_CRITERIA_SHEET);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0];
  var map = SG_headerMap_(headers);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var active = String(row[SG_col_(map, ['active'], 6)] || '').toLowerCase();
    if (active && active !== 'true' && active !== '1' && active !== 'yes') continue;
    rows.push({
      type: String(row[SG_col_(map, ['type'], 0)] || ''),
      status: String(row[SG_col_(map, ['status'], 1)] || ''),
      min: row[SG_col_(map, ['min_value'], 2)] === '' ? null : Number(row[SG_col_(map, ['min_value'], 2)]),
      max: row[SG_col_(map, ['max_value'], 3)] === '' ? null : Number(row[SG_col_(map, ['max_value'], 3)]),
      label: String(row[SG_col_(map, ['label'], 4)] || ''),
      sort: Number(row[SG_col_(map, ['sort_order'], 5)] || 99)
    });
  }
  rows.sort(function(a, b) { return a.sort - b.sort; });
  return rows;
}

function SG_calcBmi_(weightKg, heightCm) {
  var w = Number(weightKg);
  var h = Number(heightCm);
  if (!w || !h || w <= 0 || h <= 0) return '';
  var hm = h / 100;
  return Math.round((w / (hm * hm)) * 100) / 100;
}

function SG_classifyBmi_(bmi) {
  if (bmi === '' || bmi === null || isNaN(Number(bmi))) return '';
  var v = Number(bmi);
  var criteria = SG_getCriteria_().filter(function(c) { return c.type === 'bmi_basic'; });
  for (var i = 0; i < criteria.length; i++) {
    var c = criteria[i];
    var minOk = c.min === null || v >= c.min;
    var maxOk = c.max === null || v <= c.max;
    if (minOk && maxOk) return c.label || c.status;
  }
  return 'รอเทียบเกณฑ์';
}

function SG_ageYears_(birthdate, measureDate) {
  if (!birthdate) return '';
  var b = birthdate instanceof Date ? birthdate : new Date(birthdate);
  var m = measureDate ? new Date(measureDate) : new Date();
  if (isNaN(b.getTime()) || isNaN(m.getTime())) return '';
  var years = (m.getTime() - b.getTime()) / (365.2425 * 24 * 60 * 60 * 1000);
  return Math.max(0, Math.round(years * 10) / 10);
}

function SG_formatDate_(value) {
  if (!value) return '';
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value);
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function SG_studentName_(row, map) {
  var title = String(row[SG_col_(map, ['title', 'คำนำหน้า'], 2)] || '').trim();
  var first = String(row[SG_col_(map, ['firstname', 'ชื่อ'], 3)] || '').trim();
  var last = String(row[SG_col_(map, ['lastname', 'นามสกุล'], 4)] || '').trim();
  return (title + first + ' ' + last).trim();
}

function getStudentGrowthDefaults() {
  return {
    success: true,
    academicYear: SG_academicYear_(),
    month: SG_month_(),
    roundNo: '1'
  };
}

function setupStudentGrowthSheets() {
  SG_ensureSheet_(SG_RECORDS_SHEET, SG_RECORD_HEADERS);
  SG_seedCriteria_();
  SG_ensureStudentsLatestColumns_();
  return { success: true, message: 'เตรียมชีตน้ำหนัก-ส่วนสูงเรียบร้อยแล้ว' };
}

function getStudentGrowthStudents(grade, classNo, academicYear, month, roundNo) {
  try {
    setupStudentGrowthSheets();
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    academicYear = String(academicYear || SG_academicYear_()).trim();
    month = String(month || SG_month_()).trim();
    roundNo = String(roundNo || '1').trim();
    if (!grade || !classNo) return { success: false, message: 'กรุณาเลือกระดับชั้นและห้อง' };

    var ss = SS();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return { success: false, message: 'ไม่พบชีต Students' };
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, students: [] };
    var headers = values[0];
    var map = SG_headerMap_(headers);
    var idIdx = SG_col_(map, ['student_id', 'รหัสนักเรียน'], 0);
    var gradeIdx = SG_col_(map, ['grade', 'ชั้น'], 5);
    var classIdx = SG_col_(map, ['class_no', 'ห้อง'], 6);
    var genderIdx = SG_col_(map, ['gender', 'เพศ'], 7);
    var birthIdx = SG_col_(map, ['birthdate', 'วันเกิด'], 8);
    var latestWIdx = SG_col_(map, ['latest_weight_kg', 'weight'], -1);
    var latestHIdx = SG_col_(map, ['latest_height_cm', 'height'], -1);

    var existing = SG_getExistingRecordMap_(academicYear, month, roundNo, grade, classNo);
    var students = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[gradeIdx] || '').trim() !== grade) continue;
      if (String(row[classIdx] || '').trim() !== classNo) continue;
      var statusIdx = SG_col_(map, ['status', 'สถานะ'], -1);
      var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim().toLowerCase() : '';
      if (status && status !== 'active' && status !== 'กำลังศึกษา') continue;
      var sid = String(row[idIdx] || '').trim();
      if (!sid) continue;
      var rec = existing[sid] || {};
      students.push({
        student_id: sid,
        student_name: SG_studentName_(row, map),
        gender: String(row[genderIdx] || '').trim(),
        birthdate: SG_formatDate_(row[birthIdx]),
        grade: grade,
        class_no: classNo,
        weight_kg: rec.weight_kg || (latestWIdx >= 0 ? row[latestWIdx] : ''),
        height_cm: rec.height_cm || (latestHIdx >= 0 ? row[latestHIdx] : ''),
        bmi: rec.bmi || '',
        nutrition_status: rec.nutrition_status || '',
        note: rec.note || ''
      });
    }
    students.sort(function(a, b) {
      return String(a.student_id).localeCompare(String(b.student_id), 'th-TH', { numeric: true });
    });
    return { success: true, students: students, count: students.length };
  } catch (e) {
    Logger.log('getStudentGrowthStudents error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function SG_getExistingRecordMap_(academicYear, month, roundNo, grade, classNo) {
  var sheet = SG_ensureSheet_(SG_RECORDS_SHEET, SG_RECORD_HEADERS);
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return {};
  var headers = values[0];
  var map = SG_headerMap_(headers);
  var result = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[SG_col_(map, ['academic_year'], 1)] || '').trim() !== String(academicYear)) continue;
    if (String(row[SG_col_(map, ['month'], 3)] || '').trim() !== String(month)) continue;
    if (String(row[SG_col_(map, ['round_no'], 4)] || '').trim() !== String(roundNo)) continue;
    if (String(row[SG_col_(map, ['grade'], 11)] || '').trim() !== String(grade)) continue;
    if (String(row[SG_col_(map, ['class_no'], 12)] || '').trim() !== String(classNo)) continue;
    var sid = String(row[SG_col_(map, ['student_id'], 6)] || '').trim();
    if (!sid) continue;
    result[sid] = {
      weight_kg: row[SG_col_(map, ['weight_kg'], 13)],
      height_cm: row[SG_col_(map, ['height_cm'], 14)],
      bmi: row[SG_col_(map, ['bmi'], 15)],
      nutrition_status: row[SG_col_(map, ['nutrition_status'], 16)],
      note: row[SG_col_(map, ['note'], 20)]
    };
  }
  return result;
}

function saveStudentGrowthBatch(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    setupStudentGrowthSheets();
    payload = payload || {};
    var academicYear = String(payload.academicYear || SG_academicYear_()).trim();
    var month = String(payload.month || SG_month_()).trim();
    var roundNo = String(payload.roundNo || '1').trim();
    var grade = String(payload.grade || '').trim();
    var classNo = String(payload.classNo || '').trim();
    var measureDate = SG_formatDate_(payload.measureDate || new Date());
    var records = payload.records || [];
    if (!academicYear || !month || !roundNo || !grade || !classNo) {
      return { success: false, message: 'ข้อมูลปี/เดือน/รอบ/ชั้น/ห้องไม่ครบ' };
    }
    if (!records.length) return { success: false, message: 'ไม่มีรายการสำหรับบันทึก' };

    var sheet = SG_ensureSheet_(SG_RECORDS_SHEET, SG_RECORD_HEADERS);
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var map = SG_headerMap_(headers);
    var keyToRow = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var key = [
        row[SG_col_(map, ['academic_year'], 1)],
        row[SG_col_(map, ['round_id'], 2)],
        row[SG_col_(map, ['student_id'], 6)]
      ].join('|');
      keyToRow[key] = i + 1;
    }

    var now = SG_nowIso_();
    var recorder = SG_user_();
    var roundId = academicYear + '-' + month + '-' + roundNo;
    var updates = [];
    var appends = [];
    var latestRows = [];

    records.forEach(function(rec) {
      var sid = String(rec.student_id || '').trim();
      if (!sid) return;
      var weight = rec.weight_kg === '' || rec.weight_kg === null ? '' : Number(rec.weight_kg);
      var height = rec.height_cm === '' || rec.height_cm === null ? '' : Number(rec.height_cm);
      if ((weight !== '' && (isNaN(weight) || weight <= 0)) || (height !== '' && (isNaN(height) || height <= 0))) return;
      var bmi = SG_calcBmi_(weight, height);
      var nutrition = SG_classifyBmi_(bmi);
      var rowObj = {
        record_id: roundId + '-' + sid,
        academic_year: academicYear,
        round_id: roundId,
        month: month,
        round_no: roundNo,
        measure_date: measureDate,
        student_id: sid,
        student_name: String(rec.student_name || '').trim(),
        gender: String(rec.gender || '').trim(),
        birthdate: SG_formatDate_(rec.birthdate),
        age_years: SG_ageYears_(rec.birthdate, measureDate),
        grade: grade,
        class_no: classNo,
        weight_kg: weight,
        height_cm: height,
        bmi: bmi,
        nutrition_status: nutrition,
        height_status: '',
        weight_height_status: '',
        recorder: recorder,
        note: String(rec.note || '').trim(),
        created_at: now,
        updated_at: now
      };
      var rowArr = SG_RECORD_HEADERS.map(function(h) { return rowObj[h]; });
      var key = academicYear + '|' + roundId + '|' + sid;
      if (keyToRow[key]) {
        rowArr[SG_RECORD_HEADERS.indexOf('created_at')] = values[keyToRow[key] - 1][SG_col_(map, ['created_at'], 21)] || now;
        updates.push({ row: keyToRow[key], values: rowArr });
      } else {
        appends.push(rowArr);
      }
      latestRows.push(rowObj);
    });

    updates.forEach(function(item) {
      sheet.getRange(item.row, 1, 1, SG_RECORD_HEADERS.length).setValues([item.values]);
    });
    if (appends.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, SG_RECORD_HEADERS.length).setValues(appends);
    }
    SG_updateStudentsLatest_(latestRows);
    return {
      success: true,
      message: 'บันทึกน้ำหนัก-ส่วนสูงเรียบร้อยแล้ว',
      saved: updates.length + appends.length,
      updated: updates.length,
      inserted: appends.length
    };
  } catch (e) {
    Logger.log('saveStudentGrowthBatch error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_e) {}
  }
}

function SG_updateStudentsLatest_(records) {
  if (!records || !records.length) return;
  var sheet = SG_ensureStudentsLatestColumns_();
  if (!sheet) return;
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;
  var headers = values[0];
  var map = SG_headerMap_(headers);
  var idIdx = SG_col_(map, ['student_id'], 0);
  var colNames = ['latest_weight_kg', 'latest_height_cm', 'latest_bmi', 'latest_growth_date', 'latest_nutrition_status', 'weight', 'height'];
  var colIdx = {};
  colNames.forEach(function(n) { colIdx[n] = SG_col_(map, [n], -1); });
  var recMap = {};
  records.forEach(function(r) { recMap[String(r.student_id)] = r; });
  var changed = false;
  for (var i = 1; i < values.length; i++) {
    var sid = String(values[i][idIdx] || '').trim();
    if (!recMap[sid]) continue;
    var rec = recMap[sid];
    if (colIdx.latest_weight_kg >= 0) values[i][colIdx.latest_weight_kg] = rec.weight_kg;
    if (colIdx.latest_height_cm >= 0) values[i][colIdx.latest_height_cm] = rec.height_cm;
    if (colIdx.latest_bmi >= 0) values[i][colIdx.latest_bmi] = rec.bmi;
    if (colIdx.latest_growth_date >= 0) values[i][colIdx.latest_growth_date] = rec.measure_date;
    if (colIdx.latest_nutrition_status >= 0) values[i][colIdx.latest_nutrition_status] = rec.nutrition_status;
    if (colIdx.weight >= 0) values[i][colIdx.weight] = rec.weight_kg;
    if (colIdx.height >= 0) values[i][colIdx.height] = rec.height_cm;
    changed = true;
  }
  if (changed) {
    sheet.getRange(1, 1, values.length, headers.length).setValues(values);
  }
}

function getStudentGrowthReport(filters) {
  try {
    setupStudentGrowthSheets();
    filters = filters || {};
    var academicYear = String(filters.academicYear || SG_academicYear_()).trim();
    var month = String(filters.month || '').trim();
    var roundNo = String(filters.roundNo || '').trim();
    var grade = String(filters.grade || '').trim();
    var classNo = String(filters.classNo || '').trim();
    var sheet = SS().getSheetByName(SG_RECORDS_SHEET);
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) return { success: true, rows: [], summary: [], total: 0 };
    var headers = values[0];
    var map = SG_headerMap_(headers);
    var rows = [];
    var summaryMap = {};
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[SG_col_(map, ['academic_year'], 1)] || '').trim() !== academicYear) continue;
      if (month && String(row[SG_col_(map, ['month'], 3)] || '').trim() !== month) continue;
      if (roundNo && String(row[SG_col_(map, ['round_no'], 4)] || '').trim() !== roundNo) continue;
      if (grade && String(row[SG_col_(map, ['grade'], 11)] || '').trim() !== grade) continue;
      if (classNo && String(row[SG_col_(map, ['class_no'], 12)] || '').trim() !== classNo) continue;
      var status = String(row[SG_col_(map, ['nutrition_status'], 16)] || 'ไม่ระบุ').trim() || 'ไม่ระบุ';
      summaryMap[status] = (summaryMap[status] || 0) + 1;
      rows.push({
        measure_date: SG_formatDate_(row[SG_col_(map, ['measure_date'], 5)]),
        student_id: row[SG_col_(map, ['student_id'], 6)],
        student_name: row[SG_col_(map, ['student_name'], 7)],
        gender: row[SG_col_(map, ['gender'], 8)],
        grade: row[SG_col_(map, ['grade'], 11)],
        class_no: row[SG_col_(map, ['class_no'], 12)],
        weight_kg: row[SG_col_(map, ['weight_kg'], 13)],
        height_cm: row[SG_col_(map, ['height_cm'], 14)],
        bmi: row[SG_col_(map, ['bmi'], 15)],
        nutrition_status: status,
        note: row[SG_col_(map, ['note'], 20)]
      });
    }
    rows.sort(function(a, b) {
      return String(a.grade + '/' + a.class_no + '/' + a.student_id)
        .localeCompare(String(b.grade + '/' + b.class_no + '/' + b.student_id), 'th-TH', { numeric: true });
    });
    var total = rows.length;
    var summary = Object.keys(summaryMap).map(function(k) {
      return { status: k, count: summaryMap[k], percent: total ? Math.round(summaryMap[k] * 1000 / total) / 10 : 0 };
    });
    return { success: true, rows: rows, summary: summary, total: total };
  } catch (e) {
    Logger.log('getStudentGrowthReport error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function SG_escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function SG_thaiMonthName_(month) {
  var names = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  return names[Number(month)] || String(month || '');
}

function SG_reportMonths_(period) {
  period = String(period || 'term1');
  if (period === 'term2') return [11, 12, 1, 2, 3];
  if (period === 'year') return [5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
  return [5, 6, 7, 8, 9, 10];
}

function SG_periodLabel_(period) {
  if (String(period) === 'term2') return 'ภาคเรียนที่ 2';
  if (String(period) === 'year') return 'ตลอดปีการศึกษา';
  return 'ภาคเรียนที่ 1';
}

function SG_getStudentsForMatrix_(grade, classNo) {
  var sheet = SS().getSheetByName('Students');
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  var headers = values[0];
  var map = SG_headerMap_(headers);
  var idIdx = SG_col_(map, ['student_id', 'รหัสนักเรียน'], 0);
  var gradeIdx = SG_col_(map, ['grade', 'ชั้น'], 5);
  var classIdx = SG_col_(map, ['class_no', 'ห้อง'], 6);
  var statusIdx = SG_col_(map, ['status', 'สถานะ'], -1);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowGrade = String(row[gradeIdx] || '').trim();
    var rowClass = String(row[classIdx] || '').trim();
    if (grade && rowGrade !== grade) continue;
    if (classNo && rowClass !== classNo) continue;
    var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim().toLowerCase() : '';
    if (status && status !== 'active' && status !== 'กำลังศึกษา') continue;
    var sid = String(row[idIdx] || '').trim();
    if (!sid) continue;
    rows.push({
      student_id: sid,
      student_name: SG_studentName_(row, map),
      grade: rowGrade,
      class_no: rowClass
    });
  }
  rows.sort(function(a, b) {
    return String(a.grade + '/' + a.class_no + '/' + a.student_id)
      .localeCompare(String(b.grade + '/' + b.class_no + '/' + b.student_id), 'th-TH', { numeric: true });
  });
  return rows;
}

function SG_getGrowthMatrixData_(filters) {
  setupStudentGrowthSheets();
  filters = filters || {};
  var academicYear = String(filters.academicYear || SG_academicYear_()).trim();
  var grade = String(filters.grade || '').trim();
  var classNo = String(filters.classNo || '').trim();
  var roundNo = String(filters.roundNo || '').trim();
  var period = String(filters.period || 'term1').trim();
  var months = SG_reportMonths_(period);
  var students = SG_getStudentsForMatrix_(grade, classNo);
  var recordsByStudent = {};
  students.forEach(function(s) { recordsByStudent[s.student_id] = {}; });

  var sheet = SS().getSheetByName(SG_RECORDS_SHEET);
  var values = sheet ? sheet.getDataRange().getValues() : [];
  if (values.length > 1) {
    var headers = values[0];
    var map = SG_headerMap_(headers);
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[SG_col_(map, ['academic_year'], 1)] || '').trim() !== academicYear) continue;
      if (grade && String(row[SG_col_(map, ['grade'], 11)] || '').trim() !== grade) continue;
      if (classNo && String(row[SG_col_(map, ['class_no'], 12)] || '').trim() !== classNo) continue;
      var month = Number(row[SG_col_(map, ['month'], 3)] || 0);
      if (months.indexOf(month) === -1) continue;
      var recRound = String(row[SG_col_(map, ['round_no'], 4)] || '').trim();
      if (roundNo && recRound !== roundNo) continue;
      var sid = String(row[SG_col_(map, ['student_id'], 6)] || '').trim();
      if (!recordsByStudent[sid]) continue;
      var current = recordsByStudent[sid][month];
      var currentRound = current ? Number(current.round_no || 0) : -1;
      if (!current || Number(recRound || 0) >= currentRound) {
        recordsByStudent[sid][month] = {
          weight_kg: row[SG_col_(map, ['weight_kg'], 13)],
          height_cm: row[SG_col_(map, ['height_cm'], 14)],
          round_no: recRound,
          note: row[SG_col_(map, ['note'], 20)]
        };
      }
    }
  }

  var rows = students.map(function(s, idx) {
    return {
      no: idx + 1,
      student_id: s.student_id,
      student_name: s.student_name,
      grade: s.grade,
      class_no: s.class_no,
      months: recordsByStudent[s.student_id] || {}
    };
  });
  return {
    academicYear: academicYear,
    grade: grade,
    classNo: classNo,
    roundNo: roundNo,
    period: period,
    periodLabel: SG_periodLabel_(period),
    months: months,
    rows: rows,
    schoolName: (typeof getSchoolName_ === 'function' ? getSchoolName_() : '') || ''
  };
}

function SG_buildGrowthMatrixHtml_(data, forExcel) {
  var isPortrait = !forExcel;
  var colWidth = forExcel ? '' : ' width:30px;';
  var styles = '<style>' +
    '@page{size:A4 ' + (isPortrait ? 'portrait' : 'landscape') + ';margin:8mm;} body{font-family:"Sarabun",Arial,sans-serif;color:#111;margin:0;} ' +
    '.title{text-align:center;font-weight:700;font-size:' + (isPortrait ? '13px' : '15px') + ';margin-top:2px;} .subtitle{text-align:center;font-size:' + (isPortrait ? '10.5px' : '13px') + ';margin:5px 0 8px;} ' +
    'table{border-collapse:collapse;width:100%;table-layout:fixed;} th,td{border:1px solid #9ca3af;padding:' + (isPortrait ? '2px 2px' : '3px 4px') + ';font-size:' + (isPortrait ? '7.2px' : '10px') + ';text-align:center;height:' + (isPortrait ? '16px' : '21px') + ';line-height:1.15;} ' +
    'th.name,td.name{text-align:left;width:' + (isPortrait ? '82px' : '120px') + ';word-break:break-word;} th.no,td.no{width:' + (isPortrait ? '18px' : '28px') + ';} th.note,td.note{width:' + (isPortrait ? '34px' : '52px') + ';} ' +
    '.month-a{background:#d8b4fe;} .month-b{background:#fde68a;} .sub-w{background:#bbf7d0;} .sub-h{background:#fecaca;} ' +
    '.footer{text-align:center;font-size:' + (isPortrait ? '10px' : '13px') + ';margin-top:12px;font-weight:700;}' +
    '</style>';
  var titleGrade = data.grade ? data.grade : 'ทุกระดับชั้น';
  var titleClass = data.classNo ? '/' + data.classNo : '';
  var roundText = data.roundNo ? ' รอบที่ ' + data.roundNo : ' รอบล่าสุดของแต่ละเดือน';
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' + styles + '</head><body>';
  html += '<div class="title">แบบบันทึกน้ำหนักส่วนสูง</div>';
  html += '<div class="subtitle">ของนักเรียนชั้น' + SG_escapeHtml_(titleGrade + titleClass) +
    ' ' + SG_escapeHtml_(data.periodLabel) + roundText +
    ' ประจำปีการศึกษา ' + SG_escapeHtml_(data.academicYear) + '</div>';
  html += '<table><thead><tr><th class="no" rowspan="2">ที่</th><th class="name" rowspan="2">ชื่อ-สกุล</th>';
  data.months.forEach(function(m, i) {
    html += '<th colspan="2" class="' + (i % 2 ? 'month-b' : 'month-a') + '">' + SG_escapeHtml_(SG_thaiMonthName_(m)) + '</th>';
  });
  html += '<th class="note" rowspan="2">หมายเหตุ</th></tr><tr>';
  data.months.forEach(function() {
    html += '<th class="sub-w" style="' + colWidth + '">น้ำหนัก</th><th class="sub-h" style="' + colWidth + '">ส่วนสูง</th>';
  });
  html += '</tr></thead><tbody>';
  data.rows.forEach(function(r) {
    html += '<tr><td class="no">' + r.no + '</td><td class="name">' + SG_escapeHtml_(r.student_name) + '</td>';
    data.months.forEach(function(m) {
      var rec = r.months[m] || {};
      html += '<td>' + SG_escapeHtml_(rec.weight_kg || '') + '</td><td>' + SG_escapeHtml_(rec.height_cm || '') + '</td>';
    });
    html += '<td class="note"></td></tr>';
  });
  if (!data.rows.length) {
    html += '<tr><td colspan="' + (3 + data.months.length * 2) + '">ไม่พบนักเรียนตามเงื่อนไข</td></tr>';
  }
  html += '</tbody></table><div class="footer">น้ำหนักส่วนสูง</div></body></html>';
  return html;
}

function exportStudentGrowthMatrixPdf(filters) {
  try {
    var data = SG_getGrowthMatrixData_(filters || {});
    var html = SG_buildGrowthMatrixHtml_(data, false);
    var fileName = 'แบบบันทึกน้ำหนักส่วนสูง_' + (data.grade || 'ทุกชั้น') + (data.classNo ? '_' + data.classNo : '') + '_' + data.academicYear + '.pdf';
    var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(fileName);
    var url = _saveBlobGetUrl_(blob, 'StudentGrowthReports', getPdfFolderId_ && getPdfFolderId_());
    return { success: true, url: url, fileName: fileName };
  } catch (e) {
    Logger.log('exportStudentGrowthMatrixPdf error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function exportStudentGrowthMatrixXls(filters) {
  try {
    var data = SG_getGrowthMatrixData_(filters || {});
    var html = SG_buildGrowthMatrixHtml_(data, true);
    var fileName = 'แบบบันทึกน้ำหนักส่วนสูง_' + (data.grade || 'ทุกชั้น') + (data.classNo ? '_' + data.classNo : '') + '_' + data.academicYear + '.xls';
    var blob = Utilities.newBlob('\ufeff' + html, 'application/vnd.ms-excel', fileName);
    var url = _saveBlobGetUrl_(blob, 'StudentGrowthReports', getPdfFolderId_ && getPdfFolderId_());
    return { success: true, url: url, fileName: fileName };
  } catch (e) {
    Logger.log('exportStudentGrowthMatrixXls error: ' + e.message);
    return { success: false, message: e.message };
  }
}
