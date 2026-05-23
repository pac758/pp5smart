// ============================================================
// ACADEMIC YEAR AUDIT
// Read-only diagnostics for multi-year data hygiene.
// ============================================================

var AY_AUDIT_SAMPLE_LIMIT = 40;

function AY_auditTake_(arr, item) {
  if (arr.length < AY_AUDIT_SAMPLE_LIMIT) arr.push(item);
}

function AY_auditHeaders_(headers) {
  var map = {};
  (headers || []).forEach(function(h, i) {
    var key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

function AY_auditFindCol_(headersOrMap, names) {
  if (typeof AY_findColumn_ === 'function') return AY_findColumn_(headersOrMap, names);
  var map = Array.isArray(headersOrMap) ? AY_auditHeaders_(headersOrMap) : (headersOrMap || {});
  for (var i = 0; i < names.length; i++) {
    var key = String(names[i] || '').trim().toLowerCase();
    if (map[key] != null) return map[key];
  }
  return -1;
}

function AY_auditYear_(value) {
  if (typeof AY_normalizeYear_ === 'function') return AY_normalizeYear_(value);
  var s = String(value || '').trim();
  var m = s.match(/\d{4}/);
  return m ? m[0] : '';
}

function AY_auditIsInactive_(status) {
  if (typeof isInactiveStudentStatus_ === 'function') return isInactiveStudentStatus_(status);
  var st = String(status || '').trim().toLowerCase().replace(/\s+/g, '');
  return ['จำหน่าย', 'ย้ายออก', 'พ้นสภาพ', 'จบ', 'จบการศึกษา', 'สำเร็จการศึกษา', 'ลาออก', 'inactive', 'graduated', 'transferred', 'deleted'].indexOf(st) !== -1;
}

function AY_auditLabel_(rowNumber, sid, name, grade, classNo, year, status) {
  var parts = ['แถว ' + rowNumber];
  if (sid) parts.push(String(sid));
  if (name) parts.push(String(name));
  if (grade || classNo) parts.push('(' + String(grade || '-') + '/' + String(classNo || '-') + ')');
  if (year) parts.push('ปี ' + year);
  if (status) parts.push('สถานะ ' + status);
  return parts.join(' ');
}

function AY_auditStudents_(ss, targetYear) {
  var report = {
    exists: false,
    sheetName: 'Students',
    totalRows: 0,
    currentYearRows: 0,
    activeCurrentRows: 0,
    missingStudentIdCount: 0,
    blankYearCount: 0,
    oldYearActiveCount: 0,
    oldTerminalActiveCount: 0,
    currentInactiveCount: 0,
    missingStudentIdRows: [],
    blankYearRows: [],
    oldYearActiveRows: [],
    oldTerminalActiveRows: [],
    currentInactiveRows: [],
    missingAcademicYearColumn: false,
    activeCurrentIds: {}
  };

  var sheet = ss.getSheetByName('Students');
  if (!sheet) return report;

  report.exists = true;
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return report;

  var headers = data[0];
  var idIdx = AY_auditFindCol_(headers, ['student_id', 'รหัสนักเรียน', 'รหัส']);
  var titleIdx = AY_auditFindCol_(headers, ['title', 'คำนำหน้า']);
  var firstIdx = AY_auditFindCol_(headers, ['firstname', 'first_name', 'ชื่อ']);
  var lastIdx = AY_auditFindCol_(headers, ['lastname', 'last_name', 'นามสกุล']);
  var gradeIdx = AY_auditFindCol_(headers, ['grade', 'ชั้น', 'ระดับชั้น']);
  var classIdx = AY_auditFindCol_(headers, ['class_no', 'class', 'ห้อง']);
  var statusIdx = AY_auditFindCol_(headers, ['status', 'สถานะ']);
  var thaiYearKey = (typeof AY_KEY_THAI !== 'undefined') ? AY_KEY_THAI : 'ปีการศึกษา';
  var yearIdx = AY_auditFindCol_(headers, ['academic_year', 'academicyear', 'academic year', 'year', thaiYearKey]);
  report.missingAcademicYearColumn = yearIdx < 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    report.totalRows++;

    var sid = idIdx >= 0 ? String(row[idIdx] || '').trim() : '';
    var title = titleIdx >= 0 ? String(row[titleIdx] || '').trim() : '';
    var first = firstIdx >= 0 ? String(row[firstIdx] || '').trim() : '';
    var last = lastIdx >= 0 ? String(row[lastIdx] || '').trim() : '';
    var name = (title + first + ' ' + last).trim();
    var grade = gradeIdx >= 0 ? String(row[gradeIdx] || '').trim() : '';
    var classNo = classIdx >= 0 ? String(row[classIdx] || '').trim() : '';
    var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim() : '';
    var inactive = AY_auditIsInactive_(status);
    var rowYear = yearIdx >= 0 ? AY_auditYear_(row[yearIdx]) : '';
    var label = AY_auditLabel_(r + 1, sid, name, grade, classNo, rowYear, status);

    if (!sid) { report.missingStudentIdCount++; AY_auditTake_(report.missingStudentIdRows, label); }
    if (yearIdx >= 0 && !rowYear) { report.blankYearCount++; AY_auditTake_(report.blankYearRows, label); }

    var matchesTarget = (typeof AY_rowMatchesAcademicYear === 'function')
      ? AY_rowMatchesAcademicYear(row, headers, targetYear)
      : (!rowYear || rowYear === targetYear);

    if (matchesTarget) {
      report.currentYearRows++;
      if (inactive) {
        report.currentInactiveCount++;
        AY_auditTake_(report.currentInactiveRows, label);
      } else {
        report.activeCurrentRows++;
        if (sid) report.activeCurrentIds[sid] = true;
      }
    } else if (rowYear && rowYear !== targetYear && !inactive) {
      report.oldYearActiveCount++;
      AY_auditTake_(report.oldYearActiveRows, label);
      if (/^(ป\.?6|ม\.?3)$/i.test(grade)) {
        report.oldTerminalActiveCount++;
        AY_auditTake_(report.oldTerminalActiveRows, label);
      }
    }
  }

  return report;
}

function AY_auditWarehouse_(ss, targetYear, activeCurrentIds) {
  var report = {
    exists: false,
    sheetName: '',
    totalRows: 0,
    blankYearCount: 0,
    wrongYearCount: 0,
    notCurrentStudentCount: 0,
    duplicateKeyCount: 0,
    blankYearRows: [],
    wrongYearRows: [],
    notCurrentStudentRows: [],
    duplicateKeys: [],
    missingAcademicYearColumn: false
  };

  var sheet = null;
  try {
    sheet = (typeof S_getYearlySheet === 'function') ? S_getYearlySheet('SCORES_WAREHOUSE', targetYear) : ss.getSheetByName('SCORES_WAREHOUSE');
  } catch (e) {
    sheet = ss.getSheetByName('SCORES_WAREHOUSE');
  }
  if (!sheet) return report;

  report.exists = true;
  report.sheetName = sheet.getName();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return report;

  var headers = data[0];
  var sidIdx = AY_auditFindCol_(headers, ['student_id', 'รหัสนักเรียน', 'รหัส']);
  var gradeIdx = AY_auditFindCol_(headers, ['grade', 'ชั้น']);
  var classIdx = AY_auditFindCol_(headers, ['class_no', 'ห้อง']);
  var subjIdx = AY_auditFindCol_(headers, ['subject_name', 'ชื่อวิชา']);
  var codeIdx = AY_auditFindCol_(headers, ['subject_code', 'รหัสวิชา']);
  var thaiYearKey = (typeof AY_KEY_THAI !== 'undefined') ? AY_KEY_THAI : 'ปีการศึกษา';
  var yearIdx = AY_auditFindCol_(headers, ['academic_year', 'academicyear', 'academic year', 'year', thaiYearKey]);
  report.missingAcademicYearColumn = yearIdx < 0;

  var seen = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var sid = sidIdx >= 0 ? String(row[sidIdx] || '').trim() : '';
    if (!sid) continue;
    report.totalRows++;

    var grade = gradeIdx >= 0 ? String(row[gradeIdx] || '').trim() : '';
    var classNo = classIdx >= 0 ? String(row[classIdx] || '').trim() : '';
    var subjectName = subjIdx >= 0 ? String(row[subjIdx] || '').trim() : '';
    var subjectCode = codeIdx >= 0 ? String(row[codeIdx] || '').trim() : '';
    var rowYear = yearIdx >= 0 ? AY_auditYear_(row[yearIdx]) : '';
    var label = AY_auditLabel_(r + 1, sid, subjectName || subjectCode, grade, classNo, rowYear, '');

    if (yearIdx >= 0 && !rowYear) { report.blankYearCount++; AY_auditTake_(report.blankYearRows, label); }
    if (rowYear && rowYear !== targetYear) { report.wrongYearCount++; AY_auditTake_(report.wrongYearRows, label); }
    if (activeCurrentIds && !activeCurrentIds[sid]) { report.notCurrentStudentCount++; AY_auditTake_(report.notCurrentStudentRows, label); }

    var key = [sid, grade, classNo, subjectCode || subjectName, rowYear || targetYear].join('|');
    if (seen[key]) { report.duplicateKeyCount++; AY_auditTake_(report.duplicateKeys, label); }
    seen[key] = true;
  }

  return report;
}

function AY_isScoreSheet_(sheet) {
  try {
    var a1 = String(sheet.getRange(1, 1).getValue() || '').trim();
    var a2 = String(sheet.getRange(2, 1).getValue() || '').trim();
    if (a1 === 'รหัสวิชา' && a2 === 'ชื่อวิชา') return true;
    var row3 = sheet.getRange(3, 1, 1, Math.min(sheet.getLastColumn(), 8)).getValues()[0];
    return row3.join('|').indexOf('ลำดับ') !== -1 && row3.join('|').indexOf('เลขประจำตัว') !== -1;
  } catch (e) {
    return false;
  }
}

function AY_auditScoreSheets_(ss, targetYear) {
  var report = {
    totalScoreSheets: 0,
    unsuffixedCount: 0,
    oldYearSuffixCount: 0,
    suspiciousCurrentP6Count: 0,
    unsuffixedSheets: [],
    oldYearSuffixSheets: [],
    suspiciousCurrentP6Sheets: []
  };

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (!AY_isScoreSheet_(sheet)) return;
    report.totalScoreSheets++;

    var suffixMatch = name.match(/_(\d{4})$/);
    if (!suffixMatch) {
      report.unsuffixedCount++;
      AY_auditTake_(report.unsuffixedSheets, name);
      if (/\s+ป\.?6-\d+$/i.test(name)) { report.suspiciousCurrentP6Count++; AY_auditTake_(report.suspiciousCurrentP6Sheets, name); }
    } else if (suffixMatch[1] !== targetYear) {
      report.oldYearSuffixCount++;
      AY_auditTake_(report.oldYearSuffixSheets, name);
    }
  });

  return report;
}

function auditAcademicYearData(targetYear) {
  try {
    var ss = SS();
    var year = AY_auditYear_(targetYear) || ((typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '');
    if (!year) throw new Error('ไม่พบปีการศึกษาปัจจุบัน');

    var students = AY_auditStudents_(ss, year);
    var warehouse = AY_auditWarehouse_(ss, year, students.activeCurrentIds);
    var scoreSheets = AY_auditScoreSheets_(ss, year);

    var issueCount = 0;
    issueCount += students.missingAcademicYearColumn ? 1 : 0;
    issueCount += students.blankYearCount + students.oldYearActiveCount + students.oldTerminalActiveCount + students.missingStudentIdCount;
    issueCount += warehouse.missingAcademicYearColumn ? 1 : 0;
    issueCount += warehouse.blankYearCount + warehouse.wrongYearCount + warehouse.notCurrentStudentCount + warehouse.duplicateKeyCount;
    issueCount += scoreSheets.unsuffixedCount + scoreSheets.suspiciousCurrentP6Count;

    var recommendations = [];
    if (students.missingAcademicYearColumn || students.blankYearRows.length) recommendations.push('ควรเติม academic_year ให้ข้อมูลนักเรียนที่ยังว่าง');
    if (students.oldYearActiveRows.length) recommendations.push('ควรตรวจสถานะนักเรียนปีเก่าที่ยัง active');
    if (warehouse.blankYearRows.length || warehouse.wrongYearRows.length) recommendations.push('ควรซ่อม academic_year ใน SCORES_WAREHOUSE ก่อนทำรายงานข้ามปี');
    if (scoreSheets.unsuffixedSheets.length) recommendations.push('ควร snapshot/เติม suffix ปีให้ชีตคะแนนรายวิชาเก่าก่อนปิดปี');

    return {
      success: true,
      ok: issueCount === 0,
      targetYear: year,
      generatedAt: new Date().toISOString(),
      issueCount: issueCount,
      students: students,
      warehouse: warehouse,
      scoreSheets: scoreSheets,
      recommendations: recommendations
    };
  } catch (e) {
    Logger.log('auditAcademicYearData error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  }
}
