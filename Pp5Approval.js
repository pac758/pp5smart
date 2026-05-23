/**
 * PP5 approval workflow.
 * This layer stores review status only. It does not lock or modify score/student data.
 */

var PP5A_STATUS_SHEET = 'PP5_APPROVAL_STATUS';
var PP5A_HISTORY_SHEET = 'PP5_APPROVAL_HISTORY';
var PP5A_STATUS_HEADERS = [
  'academic_year', 'grade', 'class_no', 'status',
  'submitted_by', 'submitted_at',
  'reviewed_by', 'reviewed_at',
  'approved_by', 'approved_at',
  'returned_by', 'returned_at',
  'note', 'updated_by', 'updated_at'
];
var PP5A_HISTORY_HEADERS = [
  'timestamp', 'academic_year', 'grade', 'class_no',
  'action', 'from_status', 'to_status', 'user', 'role', 'note'
];

function getPp5ApprovalDashboard(academicYear) {
  try {
    var ss = SS();
    var year = pp5a_year_(academicYear);
    var statuses = pp5a_statusMap_(pp5a_readObjects_(pp5a_ensureStatusSheet_(ss)));
    var readiness = pp5a_readinessMap_(year);
    var classes = pp5a_classList_(ss);

    return {
      success: true,
      academicYear: year,
      classes: classes.map(function(cls) {
        var key = pp5a_key_(year, cls.grade, cls.classNo);
        var status = statuses[key] || pp5a_blankRecord_(year, cls.grade, cls.classNo);
        var ready = readiness[key] || null;
        return {
          academicYear: year,
          grade: cls.grade,
          classNo: cls.classNo,
          status: status.status || 'draft',
          statusLabel: pp5a_statusLabel_(status.status || 'draft'),
          note: status.note || '',
          updatedAt: status.updated_at || '',
          updatedBy: status.updated_by || '',
          submittedAt: status.submitted_at || '',
          approvedAt: status.approved_at || '',
          readinessScore: ready ? ready.score : null,
          readinessReady: ready ? !!ready.ready : null,
          blockers: ready ? ready.blockers : null,
          warnings: ready ? ready.warnings : null,
          summary: ready ? ready.summary : ''
        };
      }),
      totals: pp5a_totals_(classes, statuses, year)
    };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

function getPp5ApprovalRecord(grade, classNo, academicYear) {
  try {
    var ss = SS();
    var year = pp5a_year_(academicYear);
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    if (!grade || !classNo) throw new Error('กรุณาระบุชั้นและห้อง');

    var sheet = pp5a_ensureStatusSheet_(ss);
    var found = pp5a_findRow_(sheet, year, grade, classNo);
    var record = found ? pp5a_rowToObject_(sheet, found.row) : pp5a_blankRecord_(year, grade, classNo);
    var history = pp5a_getHistory_(ss, year, grade, classNo);
    return { success: true, record: record, history: history };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

function submitPp5ForReview(grade, classNo, academicYear, note) {
  return pp5a_transition_(grade, classNo, academicYear, 'submitted', 'submit', note || '');
}

function markPp5Reviewed(grade, classNo, academicYear, note) {
  return pp5a_transition_(grade, classNo, academicYear, 'reviewed', 'review', note || '');
}

function approvePp5Class(grade, classNo, academicYear, note) {
  return pp5a_transition_(grade, classNo, academicYear, 'approved', 'approve', note || '');
}

function returnPp5ForCorrection(grade, classNo, academicYear, note) {
  return pp5a_transition_(grade, classNo, academicYear, 'returned', 'return', note || '');
}

function pp5a_transition_(grade, classNo, academicYear, nextStatus, action, note) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SS();
    var year = pp5a_year_(academicYear);
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    if (!grade || !classNo) throw new Error('กรุณาระบุชั้นและห้อง');

    var user = pp5a_currentUser_();
    var sheet = pp5a_ensureStatusSheet_(ss);
    var historySheet = pp5a_ensureHistorySheet_(ss);
    var found = pp5a_findRow_(sheet, year, grade, classNo);
    var oldRecord = found ? pp5a_rowToObject_(sheet, found.row) : pp5a_blankRecord_(year, grade, classNo);
    var oldStatus = oldRecord.status || 'draft';

    if (oldStatus === 'approved' && nextStatus !== 'returned') {
      throw new Error('ห้องนี้อนุมัติแล้ว หากต้องแก้ไขให้ใช้สถานะส่งกลับแก้ไขก่อน');
    }
    if (nextStatus === 'approved' && oldStatus !== 'reviewed' && oldStatus !== 'submitted') {
      throw new Error('ควรส่งตรวจ/ตรวจแล้วก่อนอนุมัติ');
    }

    var now = new Date().toISOString();
    var record = Object.assign({}, oldRecord, {
      academic_year: year,
      grade: grade,
      class_no: classNo,
      status: nextStatus,
      note: String(note || '').trim(),
      updated_by: user.name,
      updated_at: now
    });

    if (nextStatus === 'submitted') {
      record.submitted_by = user.name;
      record.submitted_at = now;
    } else if (nextStatus === 'reviewed') {
      record.reviewed_by = user.name;
      record.reviewed_at = now;
    } else if (nextStatus === 'approved') {
      record.approved_by = user.name;
      record.approved_at = now;
    } else if (nextStatus === 'returned') {
      record.returned_by = user.name;
      record.returned_at = now;
    }

    pp5a_writeRecord_(sheet, found ? found.row : null, record);
    historySheet.appendRow([now, year, grade, classNo, action, oldStatus, nextStatus, user.name, user.role, String(note || '').trim()]);

    return {
      success: true,
      message: pp5a_statusLabel_(nextStatus) + ' เรียบร้อย',
      record: record
    };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function pp5a_ensureStatusSheet_(ss) {
  var sheet = ss.getSheetByName(PP5A_STATUS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PP5A_STATUS_SHEET);
    sheet.getRange(1, 1, 1, PP5A_STATUS_HEADERS.length).setValues([PP5A_STATUS_HEADERS]);
    sheet.getRange(1, 1, 1, PP5A_STATUS_HEADERS.length).setFontWeight('bold').setBackground('#e0f2fe');
  } else {
    pp5a_ensureHeaders_(sheet, PP5A_STATUS_HEADERS);
  }
  return sheet;
}

function pp5a_ensureHistorySheet_(ss) {
  var sheet = ss.getSheetByName(PP5A_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PP5A_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, PP5A_HISTORY_HEADERS.length).setValues([PP5A_HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, PP5A_HISTORY_HEADERS.length).setFontWeight('bold').setBackground('#fef3c7');
  } else {
    pp5a_ensureHeaders_(sheet, PP5A_HISTORY_HEADERS);
  }
  return sheet;
}

function pp5a_ensureHeaders_(sheet, headers) {
  var lastCol = Math.max(sheet.getLastColumn(), headers.length);
  var current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var changed = false;
  headers.forEach(function(h) {
    if (current.indexOf(h) === -1) {
      current.push(h);
      changed = true;
    }
  });
  if (changed) sheet.getRange(1, 1, 1, current.length).setValues([current]);
}

function pp5a_findRow_(sheet, year, grade, classNo) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var yIdx = headers.indexOf('academic_year');
  var gIdx = headers.indexOf('grade');
  var cIdx = headers.indexOf('class_no');
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][yIdx] || '').trim() === String(year)
      && String(values[r][gIdx] || '').trim() === String(grade)
      && String(values[r][cIdx] || '').trim() === String(classNo)) {
      return { row: r + 1 };
    }
  }
  return null;
}

function pp5a_writeRecord_(sheet, rowNumber, record) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var row = headers.map(function(h) { return record[h] || ''; });
  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function pp5a_rowToObject_(sheet, rowNumber) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  var obj = {};
  headers.forEach(function(h, i) { if (h) obj[h] = values[i]; });
  return obj;
}

function pp5a_readObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    headers.forEach(function(h, c) { if (h) obj[h] = values[r][c]; });
    rows.push(obj);
  }
  return rows;
}

function pp5a_statusMap_(rows) {
  var map = {};
  rows.forEach(function(row) {
    var key = pp5a_key_(row.academic_year, row.grade, row.class_no);
    map[key] = row;
  });
  return map;
}

function pp5a_readinessMap_(year) {
  var map = {};
  try {
    if (typeof getPp5ReadinessMatrix !== 'function') return map;
    var res = getPp5ReadinessMatrix(year);
    if (!res || !res.success || !Array.isArray(res.classes)) return map;
    res.classes.forEach(function(row) {
      map[pp5a_key_(year, row.grade, row.classNo)] = row;
    });
  } catch (e) {
    Logger.log('pp5a_readinessMap_: ' + e.message);
  }
  return map;
}

function pp5a_classList_(ss) {
  var sheet = AY_getStudentsSheetForRead();
  if (!sheet) return [];
  var rows = pp5a_readObjects_(sheet);
  var classes = {};
  rows.forEach(function(row) {
    var status = String(row.status || row['สถานะ'] || '').trim();
    if (['จำหน่าย', 'ย้ายออก', 'พ้นสภาพ', 'inactive', 'deleted'].indexOf(status) !== -1) return;
    var grade = String(row.grade || row['ชั้น'] || '').trim();
    var classNo = String(row.class_no || row.classNo || row.class || row['ห้อง'] || '').trim();
    if (!grade || !classNo) return;
    classes[grade + '|' + classNo] = { grade: grade, classNo: classNo };
  });
  return Object.keys(classes).sort().map(function(key) { return classes[key]; });
}

function pp5a_getHistory_(ss, year, grade, classNo) {
  var sheet = pp5a_ensureHistorySheet_(ss);
  return pp5a_readObjects_(sheet).filter(function(row) {
    return String(row.academic_year || '').trim() === String(year)
      && String(row.grade || '').trim() === String(grade)
      && String(row.class_no || '').trim() === String(classNo);
  }).slice(-20).reverse();
}

function pp5a_blankRecord_(year, grade, classNo) {
  return {
    academic_year: String(year || ''),
    grade: String(grade || ''),
    class_no: String(classNo || ''),
    status: 'draft',
    note: '',
    updated_by: '',
    updated_at: ''
  };
}

function pp5a_totals_(classes, statuses, year) {
  var totals = { all: classes.length, draft: 0, submitted: 0, reviewed: 0, approved: 0, returned: 0 };
  classes.forEach(function(cls) {
    var status = (statuses[pp5a_key_(year, cls.grade, cls.classNo)] || {}).status || 'draft';
    if (totals[status] === undefined) totals[status] = 0;
    totals[status]++;
  });
  return totals;
}

function pp5a_statusLabel_(status) {
  var labels = {
    draft: 'ยังไม่ส่ง',
    submitted: 'ครูส่งแล้ว',
    reviewed: 'วิชาการตรวจแล้ว',
    approved: 'อนุมัติแล้ว',
    returned: 'ส่งกลับแก้ไข'
  };
  return labels[status] || status || 'ยังไม่ส่ง';
}

function pp5a_year_(academicYear) {
  var year = String(academicYear || (typeof S_getAcademicYear === 'function' ? S_getAcademicYear() : '') || '').trim();
  if (!year) throw new Error('ไม่พบปีการศึกษา');
  return year;
}

function pp5a_key_(year, grade, classNo) {
  return [String(year || '').trim(), String(grade || '').trim(), String(classNo || '').trim()].join('|');
}

function pp5a_currentUser_() {
  try {
    if (typeof getCurrentUser === 'function') {
      var res = getCurrentUser();
      if (res && res.success) {
        return { name: res.displayName || res.username || 'user', role: res.role || 'teacher' };
      }
    }
  } catch (_) {}
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) return { name: email, role: 'user' };
  } catch (_) {}
  return { name: 'unknown', role: 'user' };
}
