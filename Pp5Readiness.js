/**
 * PP5 readiness checks for academic administration.
 * These functions are read-only: they inspect sheets and return missing items.
 */

var PP5R_YEARLY_BASES = [
  'SCORES_WAREHOUSE',
  'การประเมินอ่านคิดเขียน',
  'ประเมินอ่านคิดเขียน5เกณฑ์',
  'การประเมินคุณลักษณะ',
  'การประเมินกิจกรรมพัฒนาผู้เรียน',
  'การประเมินสมรรถนะ',
  'AttendanceLog',
  'ความเห็นครู'
];

var PP5R_ASSESSMENT_BASES = [
  { key: 'rtw', label: 'อ่าน คิด เขียน (รายวิชา)', sheet: 'การประเมินอ่านคิดเขียน' },
  { key: 'rtw5', label: 'อ่าน คิด เขียน (5 เกณฑ์)', sheet: 'ประเมินอ่านคิดเขียน5เกณฑ์' },
  { key: 'characteristic', label: 'คุณลักษณะอันพึงประสงค์', sheet: 'การประเมินคุณลักษณะ' },
  { key: 'activity', label: 'กิจกรรมพัฒนาผู้เรียน', sheet: 'การประเมินกิจกรรมพัฒนาผู้เรียน' },
  { key: 'competency', label: 'สมรรถนะสำคัญ', sheet: 'การประเมินสมรรถนะ' },
  { key: 'comment', label: 'ความคิดเห็นครู', sheet: 'ความเห็นครู' }
];

function getPp5ReadinessSummary(grade, classNo, academicYear) {
  try {
    var ss = SS();
    var year = String(academicYear || (typeof S_getAcademicYear === 'function' ? S_getAcademicYear() : '') || '').trim();
    if (!year) throw new Error('ไม่พบปีการศึกษา');

    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    if (!grade || !classNo) throw new Error('กรุณาระบุชั้นและห้อง');

    var result = {
      success: true,
      checkedAt: new Date().toISOString(),
      academicYear: year,
      grade: grade,
      classNo: classNo,
      ready: false,
      score: 0,
      totals: {},
      blockingIssues: [],
      warnings: [],
      sections: {}
    };

    pp5r_checkYearlySheets_(ss, year, result);
    var students = pp5r_getStudents_(ss, grade, classNo, result);
    var subjects = pp5r_getSubjects_(ss, grade, result);
    pp5r_checkHomeroomTeacher_(ss, grade, classNo, result);
    pp5r_checkScores_(ss, year, grade, classNo, students, subjects, result);
    pp5r_checkAssessments_(ss, year, grade, classNo, students, result);
    pp5r_checkAttendanceMonths_(ss, Number(year), result);

    var blockers = result.blockingIssues.length;
    var warnings = result.warnings.length;
    result.ready = blockers === 0;
    result.score = Math.max(0, Math.round(100 - blockers * 12 - warnings * 4));
    result.summary = result.ready
      ? 'พร้อมตรวจ/พิมพ์ ปพ.5 เบื้องต้น'
      : 'ยังไม่พร้อม: พบประเด็นที่ต้องแก้ ' + blockers + ' รายการ';
    return result;
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

function getPp5ReadinessMatrix(academicYear) {
  try {
    var ss = SS();
    var year = String(academicYear || (typeof S_getAcademicYear === 'function' ? S_getAcademicYear() : '') || '').trim();
    var studentsSheet = ss.getSheetByName('Students');
    if (!studentsSheet) throw new Error('ไม่พบชีต Students');

    var rows = pp5r_readObjects_(studentsSheet);
    var classes = {};
    rows.forEach(function(row) {
      if (pp5r_isInactive_(row)) return;
      var grade = pp5r_pick_(row, ['grade', 'ชั้น']);
      var classNo = pp5r_pick_(row, ['class_no', 'classNo', 'class', 'ห้อง']);
      if (!grade || !classNo) return;
      classes[grade + '|' + classNo] = { grade: grade, classNo: classNo };
    });

    return {
      success: true,
      academicYear: year,
      classes: Object.keys(classes).sort().map(function(key) {
        var c = classes[key];
        var r = getPp5ReadinessSummary(c.grade, c.classNo, year);
        return {
          grade: c.grade,
          classNo: c.classNo,
          ready: !!r.ready,
          score: r.score || 0,
          blockers: r.blockingIssues ? r.blockingIssues.length : 0,
          warnings: r.warnings ? r.warnings.length : 0,
          summary: r.summary || r.message || ''
        };
      })
    };
  } catch (e) {
    return { success: false, message: e.message || String(e) };
  }
}

function pp5r_checkYearlySheets_(ss, year, result) {
  var missing = [];
  PP5R_YEARLY_BASES.forEach(function(base) {
    var expected = base + '_' + year;
    if (!ss.getSheetByName(expected)) missing.push(expected);
  });
  result.sections.yearlySheets = {
    ok: missing.length === 0,
    missing: missing
  };
  if (missing.length) {
    result.blockingIssues.push('ยังไม่มีชีตรายปี: ' + missing.join(', '));
  }
}

function pp5r_getStudents_(ss, grade, classNo, result) {
  var sheet = ss.getSheetByName('Students');
  if (!sheet) {
    result.blockingIssues.push('ไม่พบชีต Students');
    result.sections.students = { ok: false, total: 0, missingIds: [], duplicateIds: [] };
    return [];
  }

  var rows = pp5r_readObjects_(sheet).filter(function(row) {
    return !pp5r_isInactive_(row)
      && pp5r_same_(pp5r_pick_(row, ['grade', 'ชั้น']), grade)
      && pp5r_same_(pp5r_pick_(row, ['class_no', 'classNo', 'class', 'ห้อง']), classNo);
  });

  var seen = {};
  var duplicateIds = [];
  var missingIds = [];
  rows.forEach(function(row) {
    var id = pp5r_studentId_(row);
    if (!id) missingIds.push(pp5r_studentName_(row) || '(ไม่ทราบชื่อ)');
    if (id) {
      if (seen[id]) duplicateIds.push(id);
      seen[id] = true;
    }
  });

  result.totals.students = rows.length;
  result.sections.students = {
    ok: rows.length > 0 && missingIds.length === 0 && duplicateIds.length === 0,
    total: rows.length,
    missingIds: missingIds,
    duplicateIds: duplicateIds
  };

  if (!rows.length) result.blockingIssues.push('ไม่พบนักเรียนในชั้น/ห้องนี้');
  if (missingIds.length) result.blockingIssues.push('มีนักเรียนไม่มีรหัส: ' + missingIds.join(', '));
  if (duplicateIds.length) result.blockingIssues.push('พบรหัสนักเรียนซ้ำ: ' + duplicateIds.join(', '));
  return rows;
}

function pp5r_getSubjects_(ss, grade, result) {
  var sheet = ss.getSheetByName('รายวิชา');
  if (!sheet) {
    result.warnings.push('ไม่พบชีตรายวิชา');
    result.sections.subjects = { ok: false, total: 0, missingTeachers: [] };
    return [];
  }

  var rows = pp5r_readObjects_(sheet).filter(function(row) {
    var rowGrade = pp5r_pick_(row, ['ชั้น', 'grade']);
    return !rowGrade || pp5r_same_(rowGrade, grade) || pp5r_same_(String(rowGrade).replace(/[\/\.]/g, ''), String(grade).replace(/[\/\.]/g, ''));
  });

  var missingTeachers = rows.filter(function(row) {
    return !pp5r_pick_(row, ['ครูผู้สอน', 'teacher', 'teacherName']);
  }).map(function(row) {
    return pp5r_pick_(row, ['ชื่อวิชา', 'subject_name', 'subjectName']) || pp5r_pick_(row, ['รหัสวิชา', 'subject_code', 'subjectCode']) || '(ไม่ทราบวิชา)';
  });

  result.totals.subjects = rows.length;
  result.sections.subjects = {
    ok: rows.length > 0,
    total: rows.length,
    missingTeachers: missingTeachers
  };
  if (!rows.length) result.blockingIssues.push('ยังไม่มีรายวิชาสำหรับชั้น ' + grade);
  if (missingTeachers.length) result.warnings.push('รายวิชาที่ยังไม่ระบุครูผู้สอน: ' + missingTeachers.join(', '));
  return rows;
}

function pp5r_checkHomeroomTeacher_(ss, grade, classNo, result) {
  var sheet = ss.getSheetByName('HomeroomTeachers');
  if (!sheet) {
    result.blockingIssues.push('ไม่พบชีต HomeroomTeachers');
    result.sections.homeroomTeacher = { ok: false };
    return;
  }

  var rows = pp5r_readObjects_(sheet);
  var found = rows.some(function(row) {
    return pp5r_same_(pp5r_pick_(row, ['grade', 'ชั้น']), grade)
      && pp5r_same_(pp5r_pick_(row, ['classNo', 'class_no', 'ห้อง']), classNo)
      && !!pp5r_pick_(row, ['teacherName', 'ครูประจำชั้น', 'ครูประจำชั้น 1']);
  });

  result.sections.homeroomTeacher = { ok: found };
  if (!found) result.blockingIssues.push('ยังไม่กำหนดครูประจำชั้นสำหรับ ' + grade + '/' + classNo);
}

function pp5r_checkScores_(ss, year, grade, classNo, students, subjects, result) {
  var sheet = pp5r_getExactYearSheet_(ss, 'SCORES_WAREHOUSE', year);
  if (!sheet) {
    result.blockingIssues.push('ไม่พบ SCORES_WAREHOUSE_' + year);
    result.sections.scores = { ok: false, rows: 0, missingStudents: [] };
    return;
  }

  var rows = pp5r_readObjects_(sheet).filter(function(row) {
    return pp5r_same_(pp5r_pick_(row, ['grade', 'ชั้น']), grade)
      && pp5r_same_(pp5r_pick_(row, ['class_no', 'classNo', 'class', 'ห้อง']), classNo);
  });
  var scoreIds = {};
  rows.forEach(function(row) {
    var id = pp5r_studentId_(row);
    if (id) scoreIds[id] = true;
  });
  var missingStudents = students.filter(function(stu) {
    return !scoreIds[pp5r_studentId_(stu)];
  }).map(pp5r_studentLabel_);

  result.sections.scores = {
    ok: rows.length > 0 && missingStudents.length === 0,
    rows: rows.length,
    missingStudents: missingStudents
  };
  if (!rows.length) result.blockingIssues.push('ยังไม่มีข้อมูลคะแนนในคลังคะแนนรวมของห้องนี้');
  if (missingStudents.length) result.blockingIssues.push('นักเรียนที่ยังไม่มีคะแนนในคลังคะแนนรวม: ' + missingStudents.join(', '));
}

function pp5r_checkAssessments_(ss, year, grade, classNo, students, result) {
  var details = {};
  PP5R_ASSESSMENT_BASES.forEach(function(item) {
    var sheet = pp5r_getExactYearSheet_(ss, item.sheet, year);
    if (!sheet) {
      details[item.key] = { ok: false, missingSheet: item.sheet + '_' + year, missingStudents: [] };
      result.blockingIssues.push('ไม่พบชีต ' + item.sheet + '_' + year);
      return;
    }

    var rows = pp5r_readObjects_(sheet).filter(function(row) {
      return pp5r_same_(pp5r_pick_(row, ['ชั้น', 'grade']), grade)
        && pp5r_same_(pp5r_pick_(row, ['ห้อง', 'class_no', 'classNo', 'class']), classNo);
    });
    var ids = {};
    rows.forEach(function(row) {
      var id = pp5r_studentId_(row);
      if (id) ids[id] = true;
    });
    var missing = students.filter(function(stu) {
      return !ids[pp5r_studentId_(stu)];
    }).map(pp5r_studentLabel_);

    details[item.key] = { ok: missing.length === 0, rows: rows.length, missingStudents: missing };
    if (missing.length) result.warnings.push(item.label + ' ยังไม่ครบ: ' + missing.length + ' คน');
  });
  result.sections.assessments = details;
}

function pp5r_checkAttendanceMonths_(ss, academicYearBE, result) {
  if (!academicYearBE || isNaN(academicYearBE)) {
    result.warnings.push('ไม่สามารถตรวจเดือนเช็คชื่อได้ เพราะปีการศึกษาไม่ถูกต้อง');
    return;
  }

  var baseYearCE = academicYearBE - 543;
  var months = [
    { m: 5, y: baseYearCE }, { m: 6, y: baseYearCE }, { m: 7, y: baseYearCE },
    { m: 8, y: baseYearCE }, { m: 9, y: baseYearCE }, { m: 10, y: baseYearCE },
    { m: 11, y: baseYearCE }, { m: 12, y: baseYearCE },
    { m: 1, y: baseYearCE + 1 }, { m: 2, y: baseYearCE + 1 }, { m: 3, y: baseYearCE + 1 }
  ];
  var names = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  var missing = months.map(function(info) {
    return names[info.m - 1] + String(info.y + 543);
  }).filter(function(name) {
    return !ss.getSheetByName(name);
  });

  result.sections.attendanceMonths = { ok: missing.length === 0, missing: missing };
  if (missing.length) result.warnings.push('ยังไม่มีชีตเช็คชื่อบางเดือน: ' + missing.join(', '));
}

function pp5r_getExactYearSheet_(ss, base, year) {
  return ss.getSheetByName(base + '_' + year);
}

function pp5r_readObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 1) return [];
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    headers.forEach(function(h, c) {
      if (h) obj[h] = values[r][c];
    });
    rows.push(obj);
  }
  return rows;
}

function pp5r_pick_(row, keys) {
  for (var i = 0; i < keys.length; i++) {
    var value = row[keys[i]];
    if (value !== undefined && value !== null && String(value).trim() !== '') return String(value).trim();
  }
  return '';
}

function pp5r_studentId_(row) {
  return pp5r_pick_(row, ['student_id', 'studentId', 'รหัสนักเรียน', 'รหัส']);
}

function pp5r_studentName_(row) {
  var full = pp5r_pick_(row, ['ชื่อ-นามสกุล', 'fullName', 'name']);
  if (full) return full;
  return [pp5r_pick_(row, ['firstname', 'ชื่อ']), pp5r_pick_(row, ['lastname', 'นามสกุล'])].filter(Boolean).join(' ');
}

function pp5r_studentLabel_(row) {
  var id = pp5r_studentId_(row);
  var name = pp5r_studentName_(row);
  return (id ? id + ' ' : '') + (name || '(ไม่ทราบชื่อ)');
}

function pp5r_same_(a, b) {
  return String(a || '').trim() === String(b || '').trim();
}

function pp5r_isInactive_(row) {
  var status = pp5r_pick_(row, ['status', 'สถานะ']);
  return ['จำหน่าย', 'ย้ายออก', 'พ้นสภาพ', 'inactive', 'deleted'].indexOf(status) !== -1;
}
