// ============================================================
// ONE PAGE REPORT - รายงานผลการพัฒนาคุณภาพผู้เรียน (หน้าเดียว)
// ตามรูปแบบมาตรฐานเอกสารวิชาการ
// ============================================================

var OPR_FONT = 'Sarabun';

function _fmtGrade(g) {
  if (g == null || isNaN(g)) return '';
  return g % 1 === 0 ? String(g) : g.toFixed(1);
}

function _opr_gradeFullName(grade) {
  var map = {
    'ป.1':'ประถมศึกษาปีที่ 1','ป.2':'ประถมศึกษาปีที่ 2',
    'ป.3':'ประถมศึกษาปีที่ 3','ป.4':'ประถมศึกษาปีที่ 4',
    'ป.5':'ประถมศึกษาปีที่ 5','ป.6':'ประถมศึกษาปีที่ 6',
    'ม.1':'มัธยมศึกษาปีที่ 1','ม.2':'มัธยมศึกษาปีที่ 2',
    'ม.3':'มัธยมศึกษาปีที่ 3','ม.4':'มัธยมศึกษาปีที่ 4',
    'ม.5':'มัธยมศึกษาปีที่ 5','ม.6':'มัธยมศึกษาปีที่ 6'
  };
  return map[grade] || grade;
}

// คำนวณ "เลขที่" จากลำดับในห้อง (เรียงตามชื่อ)
function _opr_getClassNumber(studentId, grade, classNo) {
  try {
    var students = getStudentsByClass(grade, String(classNo));
    if (!students || students.length === 0) return '';
    students.sort(function(a, b) {
      var na = (a.title||'') + (a.firstname||'') + ' ' + (a.lastname||'');
      var nb = (b.title||'') + (b.firstname||'') + ' ' + (b.lastname||'');
      return na.localeCompare(nb, 'th');
    });
    for (var i = 0; i < students.length; i++) {
      if (String(students[i].id).trim() === String(studentId).trim()) return String(i + 1);
    }
    return '';
  } catch (e) { return ''; }
}

function _opr_cell(cell, w, fs, bold, bg, align) {
  cell.setWidth(w);
  if (bg) cell.setBackgroundColor(bg);
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(2).setPaddingRight(2);
  var a = align || DocumentApp.HorizontalAlignment.CENTER;
  var txt = cell.getChild(0).asParagraph().getText();
  var lines = txt.split('\n');
  if (lines.length > 1) {
    cell.getChild(0).asParagraph().setText(lines[0]);
    var p0 = cell.getChild(0).asParagraph();
    p0.setAlignment(a);
    p0.setSpacingAfter(0).setSpacingBefore(0);
    var t0 = p0.editAsText();
    t0.setFontSize(fs).setFontFamily(OPR_FONT);
    if (bold) t0.setBold(true);
    for (var li = 1; li < lines.length; li++) {
      var pn = cell.appendParagraph(lines[li]);
      pn.setAlignment(a);
      pn.setSpacingAfter(0).setSpacingBefore(0);
      var tn = pn.editAsText();
      tn.setFontSize(fs).setFontFamily(OPR_FONT);
      if (bold) tn.setBold(true);
    }
  } else {
    var p = cell.getChild(0).asParagraph();
    p.setAlignment(a);
    p.setSpacingAfter(0).setSpacingBefore(0);
    var t = p.editAsText();
    t.setFontSize(fs).setFontFamily(OPR_FONT);
    if (bold) t.setBold(true);
  }
  return cell;
}

function _opr_addLine(cell, text, fs, bold) {
  var p = cell.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setSpacingAfter(0).setSpacingBefore(0);
  var t = p.editAsText();
  t.setFontSize(fs).setFontFamily(OPR_FONT);
  if (bold) t.setBold(true);
}

function _opr_merge(tableStart, row, col, rs, cs) {
  return {
    mergeTableCells: {
      tableRange: {
        tableCellLocation: {
          tableStartLocation: { index: tableStart },
          rowIndex: row, columnIndex: col
        },
        rowSpan: rs, columnSpan: cs
      }
    }
  };
}

/**
 * Batch: สร้าง PDF หลายคนพร้อมกัน อ่านข้อมูลชีทครั้งเดียว
 * @param {Array<string>} studentIds - รายการ student ID
 * @returns {Array<{id:string, url:string, error:string}>}
 */
function generateOnePageReportBatch(studentIds) {
  var cache = _opr_preloadData();
  var term = (arguments.length >= 2 && arguments[1] !== undefined && arguments[1] !== null) ? String(arguments[1]) : 'both';
  var results = [];
  for (var bi = 0; bi < studentIds.length; bi++) {
    try {
      var url = _opr_generateSingle(studentIds[bi], cache, term);
      results.push({id: studentIds[bi], url: url, error: ''});
    } catch (e) {
      results.push({id: studentIds[bi], url: '', error: e.message});
    }
  }
  return results;
}

/**
 * Pre-load ข้อมูลที่ใช้ร่วมกัน (อ่านครั้งเดียว)
 */
function _opr_preloadData() {
  var settings = getCachedSettings_();
  var sd = {
    name: settings['ชื่อโรงเรียน'] || 'โรงเรียน',
    director: settings['ชื่อผู้อำนวยการ'] || '',
    dirPos: settings['ตำแหน่งผู้อำนวยการ'] || 'ผู้อำนวยการสถานศึกษา',
    year: settings['ปีการศึกษา'] || String(new Date().getFullYear() + 543),
    logoId: settings['logoFileId'] || ''
  };
  var subjSheet = _readSheetToObjects('รายวิชา');
  var sInfo = {};
  subjSheet.forEach(function(s) {
    var code = String(s['รหัสวิชา'] || '').trim();
    if (code) {
      sInfo[code] = {
        hours: parseFloat(s['ชั่วโมง/ปี']) || 0,
        type: String(s['ประเภทวิชา'] || 'พื้นฐาน').trim(),
        order: parseInt(s['ลำดับ']) || 999,
        credit: parseFloat(s['หน่วยกิต'] || s['น้ำหนัก'] || s['ชั่วโมง/ปี']) || 0
      };
    }
  });

  // Pre-load ชีทประเมินทั้ง 3 แผ่น
  var readingData = [], characterData = [], activityData = [];
  try { readingData = _readSheetToObjects('การประเมินอ่านคิดเขียน'); } catch(e){}
  try { characterData = _readSheetToObjects('การประเมินคุณลักษณะ'); } catch(e){}
  try { activityData = _readSheetToObjects('การประเมินกิจกรรมพัฒนาผู้เรียน', false); } catch(e){}

  // Pre-load class lists (cache per grade+classNo)
  var classLists = {};

  return {
    settings: settings, sd: sd, sInfo: sInfo,
    whByClass: {},
    whIndexByClass: {},
    readingData: readingData, characterData: characterData, activityData: activityData,
    classLists: classLists,
    // cache GPA results per grade+classNo (คำนวณครั้งเดียวต่อห้อง)
    gpaCache: {}
  };
}

function _opr_getWarehouseForClass_(grade, classNo, cache) {
  var y = (cache && cache.sd && cache.sd.year) ? String(cache.sd.year) : '';
  var g = String(grade || '').trim();
  var c = String(classNo || '').trim();
  var classKey = y + '|' + g + '|' + c;

  if (cache && cache.whByClass && cache.whByClass[classKey]) return cache.whByClass[classKey];

  var cacheKey = (typeof _createCacheKey === 'function')
    ? _createCacheKey('opr_wh', y, g, c)
    : ('opr_wh_' + y + '_' + g + '_' + c);

  var cached = (typeof _getFromCache === 'function') ? _getFromCache(cacheKey, null) : null;
  if (cached && Array.isArray(cached)) {
    if (cache && cache.whByClass) cache.whByClass[classKey] = cached;
    if (cache && cache.whIndexByClass) {
      var idx = {};
      cached.forEach(function(r) {
        var sid = String(r.student_id || '').trim();
        if (!sid) return;
        if (!idx[sid]) idx[sid] = [];
        idx[sid].push(r);
      });
      cache.whIndexByClass[classKey] = idx;
    }
    return cached;
  }

  var sheet = null;
  try {
    sheet = (typeof S_getYearlySheet === 'function') ? S_getYearlySheet('SCORES_WAREHOUSE') : null;
  } catch (_e) {}
  try {
    if (!sheet) sheet = SS().getSheetByName('SCORES_WAREHOUSE');
  } catch (_e) {}
  if (!sheet) {
    if (cache && cache.whByClass) cache.whByClass[classKey] = [];
    if (cache && cache.whIndexByClass) cache.whIndexByClass[classKey] = {};
    return [];
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    if (cache && cache.whByClass) cache.whByClass[classKey] = [];
    if (cache && cache.whIndexByClass) cache.whIndexByClass[classKey] = {};
    return [];
  }

  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var col = function(name) { return headers.indexOf(name); };
  var findCol = function() {
    for (var i = 0; i < arguments.length; i++) {
      var cidx = col(arguments[i]);
      if (cidx !== -1) return cidx;
    }
    return -1;
  };

  var idCol = findCol('student_id', 'studentId');
  var gradeCol = findCol('grade');
  var classCol = findCol('class_no', 'classNo', 'class');
  var codeCol = findCol('subject_code', 'subjectCode');
  var nameCol = findCol('subject_name', 'subjectName');
  var typeCol = findCol('subject_type', 'subjectType');
  var hoursCol = findCol('hours', 'ชั่วโมง/ปี');
  var t1Col = findCol('term1_total', 'term1Total');
  var t2Col = findCol('term2_total', 'term2Total');
  var avgCol = findCol('average', 'avg');
  var fgCol = findCol('final_grade', 'grade_result', 'gradeResult');

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (gradeCol !== -1 && String(row[gradeCol] || '').trim() !== g) continue;
    if (classCol !== -1 && String(row[classCol] || '').trim() !== c) continue;

    out.push({
      student_id: idCol !== -1 ? String(row[idCol] || '').trim() : '',
      grade: gradeCol !== -1 ? String(row[gradeCol] || '').trim() : '',
      class_no: classCol !== -1 ? String(row[classCol] || '').trim() : '',
      subject_code: codeCol !== -1 ? String(row[codeCol] || '').trim() : '',
      subject_name: nameCol !== -1 ? String(row[nameCol] || '').trim() : '',
      subject_type: typeCol !== -1 ? String(row[typeCol] || '').trim() : '',
      hours: hoursCol !== -1 ? row[hoursCol] : '',
      term1_total: t1Col !== -1 ? row[t1Col] : '',
      term2_total: t2Col !== -1 ? row[t2Col] : '',
      average: avgCol !== -1 ? row[avgCol] : '',
      final_grade: fgCol !== -1 ? row[fgCol] : ''
    });
  }

  if (typeof _setToCache === 'function') _setToCache(cacheKey, out, (typeof CACHE_EXPIRATION !== 'undefined' ? CACHE_EXPIRATION.MEDIUM : 1800));
  if (cache && cache.whByClass) cache.whByClass[classKey] = out;
  if (cache && cache.whIndexByClass) {
    var idx2 = {};
    out.forEach(function(rr) {
      var sid2 = String(rr.student_id || '').trim();
      if (!sid2) return;
      if (!idx2[sid2]) idx2[sid2] = [];
      idx2[sid2].push(rr);
    });
    cache.whIndexByClass[classKey] = idx2;
  }
  return out;
}

/**
 * สร้าง PDF สำหรับนักเรียน 1 คน (ใช้ cache จาก preload)
 */
function _opr_generateSingle(studentId, cache) {
  var term = (arguments.length >= 3 && arguments[2] !== undefined && arguments[2] !== null) ? String(arguments[2]).trim() : 'both';
  if (term !== '1' && term !== '2' && term !== 'both') term = 'both';
  var student = getStudentInfo_(studentId);
  if (!student) throw new Error('ไม่พบข้อมูลนักเรียน');

  var sd = cache.sd;
  var settings = cache.settings;
  var sInfo = cache.sInfo;
  var sid = String(studentId).trim();
  var y = sd && sd.year ? String(sd.year) : '';
  var classKey = y + '|' + String(student.grade || '').trim() + '|' + String(student.classNo || '').trim();
  var whClass = _opr_getWarehouseForClass_(student.grade, student.classNo, cache);

  // เลขที่ในห้อง (cache per class)
  var classListKey = student.grade + '_' + student.classNo;
  if (!cache.classLists[classListKey]) {
    try {
      var students = getStudentsByClass(student.grade, String(student.classNo));
      students.sort(function(a, b) {
        return String(a.id||'').localeCompare(String(b.id||''), undefined, {numeric: true});
      });
      cache.classLists[classListKey] = students;
    } catch(e) { cache.classLists[classListKey] = []; }
  }
  var classNumber = '';
  var cl = cache.classLists[classListKey];
  for (var ci = 0; ci < cl.length; ci++) {
    if (String(cl[ci].id).trim() === sid) { classNumber = String(ci + 1); break; }
  }

  // คะแนนนักเรียน
  var studentScores = [];
  if (cache.whIndexByClass && cache.whIndexByClass[classKey] && cache.whIndexByClass[classKey][sid]) {
    studentScores = cache.whIndexByClass[classKey][sid];
  } else {
    studentScores = whClass.filter(function(r) { return String(r['student_id']).trim() === sid; });
  }

  var subjects = studentScores
    .filter(function(r) {
      var code = String(r['subject_code'] || '').trim();
      var tp = (sInfo[code] && sInfo[code].type) || String(r['subject_type'] || '').trim();
      return tp !== 'กิจกรรม';
    })
    .map(function(r) {
      var code = String(r['subject_code'] || '').trim();
      var inf = sInfo[code] || {};
      var t1 = parseFloat(r['term1_total']) || 0;
      var t2 = parseFloat(r['term2_total']) || 0;
      var avg = parseFloat(r['average']) || ((t1 + t2) / 2);
      var fg = parseFloat(r['final_grade']);
      return {
        code: code,
        name: String(r['subject_name'] || '').trim(),
        type: inf.type || String(r['subject_type'] || 'พื้นฐาน').trim(),
        credit: inf.credit || parseFloat(r['hours']) || 0,
        order: inf.order || 999,
        t1: t1, t1g: _scoreToGPA(t1).gpa,
        t2: t2, t2g: _scoreToGPA(t2).gpa,
        avg: avg, fg: !isNaN(fg) ? fg : _scoreToGPA(avg).gpa
      };
    });
  _sortBySubjectName(subjects, 'name');

  // ผลประเมิน (ใช้ข้อมูลที่ pre-load แล้ว)
  var assessments = _opr_getAssessments(sid, cache);

  var gpa = _opr_getGPAByTerm_(sid, student.grade, student.classNo, cache, whClass, term);

  var teacher = getHomeroomTeacher(student.grade, String(student.classNo));

  var validAvg = subjects.filter(function(s){return s.avg>0;}).map(function(s){return s.avg;});
  var avgAll = validAvg.length > 0 ? validAvg.reduce(function(a,b){return a+b;},0)/validAvg.length : 0;
  if (term === '1' || term === '2') {
    var validT = subjects
      .map(function(s) { return term === '1' ? Number(s.t1) : Number(s.t2); })
      .filter(function(n) { return Number.isFinite(n) && n > 0; });
    avgAll = validT.length > 0 ? validT.reduce(function(a,b){return a+b;},0)/validT.length : 0;
  }

    // =============================================
    // สร้าง Google Doc (A4)
    // =============================================
    var termTag = term === 'both' ? 'ทั้งปี' : ('ภาค' + term);
    var doc = DocumentApp.create('รายงานหน้าเดียว_' + termTag + '_' + student.id + '_' + Date.now());
    var body = doc.getBody();
    body.clear();
    body.setMarginTop(22);
    body.setMarginBottom(14);
    body.setMarginLeft(50);
    body.setMarginRight(28);
    body.setPageWidth(595);
    body.setPageHeight(842);

    var F = OPR_FONT;
    var CENTER = DocumentApp.HorizontalAlignment.CENTER;
    var LEFT = DocumentApp.HorizontalAlignment.LEFT;

    // =============================================
    // HEADER: โลโก้ + ชื่อโรงเรียน + หัวข้อรายงาน
    // =============================================
    if (sd.logoId) {
      try {
        var lp = body.appendParagraph('');
        lp.setAlignment(CENTER);
        var img = lp.appendInlineImage(DriveApp.getFileById(sd.logoId).getBlob());
        img.setHeight(45); img.setWidth(45);
        lp.setSpacingAfter(0).setSpacingBefore(0);
      } catch(e) { Logger.log('Logo: '+e.message); }
    }

    var addLine = function(txt, sz, bold, sa, align) {
      var p = body.appendParagraph(txt);
      p.setAlignment(align || CENTER);
      p.editAsText().setBold(bold).setFontSize(sz).setFontFamily(F);
      p.setSpacingAfter(sa !== undefined ? sa : 0).setSpacingBefore(0);
      return p;
    };

    var termText = term === 'both' ? 'ภาคเรียนที่ 1 และ 2' : ('ภาคเรียนที่ ' + term);
    addLine(sd.name, 14, true, 0);
    addLine('แบบรายงานผลการพัฒนาคุณภาพผู้เรียน', 13, true, 0);
    addLine('ชั้น' + _opr_gradeFullName(student.grade) + '/' + student.classNo + '  ' + termText + '  ปีการศึกษา ' + sd.year, 13, true, 1);

    // บรรทัดข้อมูลนักเรียน: เลขที่ + ชื่อ-นามสกุล จัดกลาง
    var infoText = 'เลขที่  ' + classNumber + '     ชื่อ - นามสกุล  ' + student.name;
    var infoPara = addLine(infoText, 13, false, 2, CENTER);

    // =============================================
    var HDR_BG = '#E8E8E8';
    var FS_H = 11;
    var FS_D = 12;

    if (term === 'both') {
      var T = body.appendTable();
      T.setBorderWidth(0.5);
      T.setBorderColor('#000000');

      var W = [18, 42, 197, 38, 22, 38, 26, 38, 26, 42, 30];

      var r1 = T.appendTableRow();
      var h1Vals = ['ที่','รหัสวิชา','ชื่อวิชา','ประเภท','น.น.','ภาคเรียนที่ 1','','ภาคเรียนที่ 2','','สรุปผล',''];
      h1Vals.forEach(function(h, i) { _opr_cell(r1.appendTableCell(h), W[i], FS_H, true, HDR_BG); });
      _opr_addLine(r1.getCell(3), 'วิชา', FS_H, true);
      _opr_addLine(r1.getCell(9), 'ปลายปี', FS_H, true);

      var r2 = T.appendTableRow();
      var h2Vals = ['','','','','','คะแนน','เกรด','คะแนน','เกรด','คะแนน','เกรด'];
      h2Vals.forEach(function(h, i) { _opr_cell(r2.appendTableCell(h), W[i], FS_H, true, HDR_BG); });
      _opr_addLine(r2.getCell(9), 'เฉลี่ย', FS_H, true);

      subjects.forEach(function(s, idx) {
        var row = T.appendTableRow();
        var v = [
          String(idx + 1),
          s.code,
          s.name,
          s.type,
          String(s.credit || ''),
          s.t1 > 0 ? String(Math.round(s.t1)) : '',
          s.t1g > 0 ? _fmtGrade(s.t1g) : '',
          s.t2 > 0 ? String(Math.round(s.t2)) : '',
          s.t2g > 0 ? _fmtGrade(s.t2g) : '',
          s.avg > 0 ? s.avg.toFixed(2) : '',
          s.fg >= 0 ? _fmtGrade(s.fg) : ''
        ];
        v.forEach(function(val, i) {
          var al = (i === 2) ? LEFT : null;
          _opr_cell(row.appendTableCell(val), W[i], FS_D, false, null, al);
        });
      });
    } else {
      var TT = body.appendTable();
      TT.setBorderWidth(0.5);
      TT.setBorderColor('#000000');

      var W7 = [18, 42, 250, 45, 30, 70, 62];
      var hr = TT.appendTableRow();
      var termHdr = 'ภาค ' + term;
      ['ที่','รหัสวิชา','ชื่อวิชา','ประเภท','น.น.','คะแนน ' + termHdr,'เกรด ' + termHdr]
        .forEach(function(h, i) { _opr_cell(hr.appendTableCell(h), W7[i], (i >= 5 ? 10 : FS_H), true, HDR_BG, (i === 2 ? LEFT : null)); });

      subjects.forEach(function(s, idx) {
        var row2 = TT.appendTableRow();
        var score = term === '1' ? Number(s.t1) : Number(s.t2);
        var gradeVal = term === '1' ? Number(s.t1g) : Number(s.t2g);
        var vv = [
          String(idx + 1),
          s.code,
          s.name,
          s.type,
          String(s.credit || ''),
          (Number.isFinite(score) && score > 0) ? String(Math.round(score)) : '',
          (Number.isFinite(gradeVal) && gradeVal > 0) ? _fmtGrade(gradeVal) : ''
        ];
        vv.forEach(function(val, i) {
          var al2 = (i === 2) ? LEFT : null;
          _opr_cell(row2.appendTableCell(val), W7[i], FS_D, false, null, al2);
        });
      });
    }

    // =============================================
    // สรุป GPA + คะแนนเฉลี่ย + ลำดับ (ใต้ตาราง มีเส้นกรอบ)
    // =============================================
    var sumTable = body.appendTable();
    sumTable.setBorderWidth(0.5);
    sumTable.setBorderColor('#000000');
    var sumRow = sumTable.appendTableRow();
    var sumText = 'ผลการเรียนเฉลี่ย     ' + gpa.gpa.toFixed(2) +
                  '     คะแนนเฉลี่ยทุกวิชา     ' + avgAll.toFixed(2) +
                  '     ได้ลำดับที่     ' + gpa.classRank;
    var sumCell = sumRow.appendTableCell(sumText);
    sumCell.setWidth(517);
    sumCell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(4).setPaddingRight(4);
    var sp = sumCell.getChild(0).asParagraph();
    sp.setAlignment(CENTER);
    sp.setSpacingAfter(0).setSpacingBefore(0);
    sp.editAsText().setBold(true).setFontSize(12).setFontFamily(F);

    // =============================================
    // สรุปผลการประเมินด้านต่างๆ (ตาราง 3 คอลัมน์)
    // =============================================
    var AT = body.appendTable();
    AT.setBorderWidth(0.5);
    AT.setBorderColor('#000000');
    var AW = [173, 172, 172];

    // Row 1: หัวข้อ (merged)
    var ar1 = AT.appendTableRow();
    ['สรุปผลการประเมินด้านต่างๆ','',''].forEach(function(t, i) {
      _opr_cell(ar1.appendTableCell(t), AW[i], 12, true, HDR_BG);
    });

    // Row 2: ชื่อด้าน
    var ar2 = AT.appendTableRow();
    ['คุณลักษณะอันพึงประสงค์ของสถานศึกษา','การอ่านคิดวิเคราะห์และเขียนสื่อความ','กิจกรรมพัฒนาผู้เรียน'].forEach(function(t, i) {
      _opr_cell(ar2.appendTableCell(t), AW[i], 11, false);
    });

    // Row 3: ผลประเมิน
    var ar3 = AT.appendTableRow();
    var charResult = (assessments.character && assessments.character.result) || '-';
    var readResult = (assessments.reading && assessments.reading.result) || '-';
    var actResult = (assessments.activities && assessments.activities.result) || '-';
    [charResult, readResult, actResult].forEach(function(t, i) {
      _opr_cell(ar3.appendTableCell(t), AW[i], 12, true);
    });

    // =============================================
    // ลายเซ็น 3 คอลัมน์ (ไม่มีเส้นกรอบ)
    // =============================================
    var ST = body.appendTable();
    ST.setBorderWidth(0);
    var tName = (teacher && teacher !== '...') ? teacher : '';
    var SW = [173, 172, 172];
    var nm1 = tName ? '(' + tName + ')' : '(.................................)';
    var nm2 = '(.................................)';
    var acHead = (settings && settings['ชื่อหัวหน้างานวิชาการ']) ? settings['ชื่อหัวหน้างานวิชาการ'] : '';
    var nmAc = acHead ? '(' + acHead + ')' : '(.................................)';

    // แยกแต่ละบรรทัดเป็น paragraph แยก เพื่อให้จัด CENTER ทีละบรรทัด
    var signData = [
      ['ลงชื่อ..................................', nm1, 'ครูประจำชั้นคนที่ 1'],
      ['ลงชื่อ..................................', nm2, 'ครูประจำชั้นคนที่ 2'],
      ['ลงชื่อ..................................', nmAc, 'หัวหน้า/รอง ฝ่ายวิชาการ']
    ];
    var sr = ST.appendTableRow();
    signData.forEach(function(lines, ci) {
      var c = sr.appendTableCell(lines[0]);
      c.setWidth(SW[ci]);
      c.setPaddingTop(6).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);
      // paragraph แรก (ลงชื่อ...)
      var p0 = c.getChild(0).asParagraph();
      p0.setAlignment(CENTER);
      p0.setSpacingAfter(0).setSpacingBefore(0);
      p0.editAsText().setFontSize(12).setFontFamily(F);
      // paragraph ที่ 2 (ชื่อในวงเล็บ)
      var p1 = c.appendParagraph(lines[1]);
      p1.setAlignment(CENTER);
      p1.setSpacingAfter(0).setSpacingBefore(0);
      p1.editAsText().setFontSize(12).setFontFamily(F);
      // paragraph ที่ 3 (ตำแหน่ง)
      var p2 = c.appendParagraph(lines[2]);
      p2.setAlignment(CENTER);
      p2.setSpacingAfter(0).setSpacingBefore(0);
      p2.editAsText().setFontSize(12).setFontFamily(F);
    });

    // =============================================
    // ส่วนอนุมัติ (ตาราง 4 คอลัมน์ มีเส้นกรอบ)
    // =============================================
    var BT = body.appendTable();
    BT.setBorderWidth(0.5);
    BT.setBorderColor('#000000');

    // Header row 1: สรุปผลการประเมิน (merge 2 cols), วันที่อนุมัติ, ลงชื่อ
    var bh = BT.appendTableRow();
    _opr_cell(bh.appendTableCell('สรุปผลการประเมิน'), 75, 12, true, HDR_BG);
    _opr_cell(bh.appendTableCell(''), 75, 12, true, HDR_BG);
    _opr_cell(bh.appendTableCell('วันที่อนุมัติผลการเรียน'), 185, 12, true, HDR_BG);
    _opr_cell(bh.appendTableCell('ลงชื่อ'), 182, 12, true, HDR_BG);

    // Header row 2: เลื่อนชั้น/จบ | ซ้ำชั้น | (empty) | (empty)
    var bh2 = BT.appendTableRow();
    _opr_cell(bh2.appendTableCell('เลื่อนชั้น/จบ'), 75, 11, true, HDR_BG);
    _opr_cell(bh2.appendTableCell('ซ้ำชั้น'), 75, 11, true, HDR_BG);
    _opr_cell(bh2.appendTableCell(''), 185, 11, false, HDR_BG);
    _opr_cell(bh2.appendTableCell(''), 182, 11, false, HDR_BG);

    // Data row
    // วันอนุมัติ = 31 มีนาคม ของปีการศึกษานั้น (ปี พ.ศ. + 1 เพราะจบปลายปี)
    var yearBE = parseInt(sd.year) || (new Date().getFullYear() + 543);
    var dateStr = '31 มีนาคม ' + (yearBE + 1);

    var bd = BT.appendTableRow();

    // Cell 1: เลื่อนชั้น/จบ → ✓
    var bc1 = bd.appendTableCell(term === 'both' ? '✓' : '');
    bc1.setWidth(75);
    bc1.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    bc1.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
    var bp1 = bc1.getChild(0).asParagraph();
    bp1.setAlignment(CENTER);
    bp1.setSpacingAfter(0).setSpacingBefore(0);
    bp1.editAsText().setFontSize(14).setFontFamily(F).setBold(true);

    // Cell 2: ซ้ำชั้น → ว่าง
    var bc1b = bd.appendTableCell('');
    bc1b.setWidth(75);
    bc1b.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    bc1b.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);

    // Cell 3: วันที่
    var bc2 = bd.appendTableCell(dateStr);
    bc2.setWidth(185);
    bc2.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    bc2.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(4).setPaddingRight(4);
    var bp2 = bc2.getChild(0).asParagraph();
    bp2.setAlignment(CENTER);
    bp2.setSpacingAfter(0).setSpacingBefore(0);
    bp2.editAsText().setFontSize(12).setFontFamily(F);

    // Cell 4: ลงชื่อ ผอ. (แยก paragraph เพื่อจัด CENTER ทีละบรรทัด)
    var dirName = sd.director || '..........................';
    var bc3 = bd.appendTableCell('ลงชื่อ..................................');
    bc3.setWidth(182);
    bc3.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(4).setPaddingRight(4);
    var bp3_0 = bc3.getChild(0).asParagraph();
    bp3_0.setAlignment(CENTER);
    bp3_0.setSpacingAfter(0).setSpacingBefore(0);
    bp3_0.editAsText().setFontSize(11).setFontFamily(F);
    var bp3_1 = bc3.appendParagraph('(' + dirName + ')');
    bp3_1.setAlignment(CENTER);
    bp3_1.setSpacingAfter(0).setSpacingBefore(0);
    bp3_1.editAsText().setFontSize(11).setFontFamily(F);
    var bp3_2 = bc3.appendParagraph(sd.dirPos);
    bp3_2.setAlignment(CENTER);
    bp3_2.setSpacingAfter(0).setSpacingBefore(0);
    bp3_2.editAsText().setFontSize(11).setFontFamily(F);
    if (sd.dirPos.indexOf(sd.name) === -1) {
      var bp3_3 = bc3.appendParagraph(sd.name);
      bp3_3.setAlignment(CENTER);
      bp3_3.setSpacingAfter(0).setSpacingBefore(0);
      bp3_3.editAsText().setFontSize(11).setFontFamily(F);
    }

    // =============================================
    // SAVE → MERGE CELLS → PDF
    // =============================================
    doc.saveAndClose();
    var docId = doc.getId();

    // Merge cells ผ่าน Docs API
    try {
      var dj = Docs.Documents.get(docId);
      var reqs = [];
      for (var ei = 0; ei < dj.body.content.length; ei++) {
        var el = dj.body.content[ei];
        if (!el.table) continue;
        var nc = el.table.tableRows[0].tableCells.length;
        var nr = el.table.tableRows.length;
        var si = el.startIndex;

        // ตารางผลการเรียน (11 cols): merge headers
        if (nc === 11 && nr > 2) {
          reqs.push(_opr_merge(si, 0, 5, 1, 2));   // ภาคเรียนที่ 1
          reqs.push(_opr_merge(si, 0, 7, 1, 2));   // ภาคเรียนที่ 2
          reqs.push(_opr_merge(si, 0, 9, 1, 2));   // สรุปผลปลายปี
          for (var ci = 0; ci < 5; ci++) {
            reqs.push(_opr_merge(si, 0, ci, 2, 1)); // merge row1+row2 for cols 0-4
          }
        }
        // ตารางประเมิน (3 cols, 3 rows): merge row 1
        if (nc === 3 && nr === 3) {
          reqs.push(_opr_merge(si, 0, 0, 1, 3));
        }
        // ตารางอนุมัติ (4 cols, 3 rows): merge headers
        if (nc === 4 && nr === 3) {
          reqs.push(_opr_merge(si, 0, 0, 1, 2));  // สรุปผลการประเมิน merge 2 cols
          reqs.push(_opr_merge(si, 0, 2, 2, 1));  // วันที่อนุมัติ merge 2 rows
          reqs.push(_opr_merge(si, 0, 3, 2, 1));  // ลงชื่อ merge 2 rows
        }
      }
      if (reqs.length > 0) {
        Docs.Documents.batchUpdate({requests: reqs}, docId);
      }
    } catch (me) { Logger.log('Merge: ' + me.message); }

    // แปลง PDF (with step tracking)
    var _pdfStep = 'getFileById';
    try {
      var docFile = DriveApp.getFileById(docId);
      _pdfStep = 'getAs_pdf';
      var pdfBlob = docFile.getAs('application/pdf');
      _pdfStep = 'getCachedFolder';
      var folder = getCachedFolder_(settings);
      _pdfStep = 'createFile';
      var pdfFile = folder.createFile(pdfBlob);
      _pdfStep = 'setName';
      pdfFile.setName('รายงานหน้าเดียว_' + termTag + '_' + student.id + '_' + student.name + '.pdf');
      _pdfStep = 'setSharing';
      try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(shareErr) { Logger.log('setSharing skipped (domain restriction): ' + shareErr.message); }
      _pdfStep = 'setTrashed';
      docFile.setTrashed(true);
    } catch (pdfErr) {
      throw new Error('[PDF step: ' + _pdfStep + '] ' + pdfErr.message);
    }

    Logger.log('สร้างรายงานหน้าเดียวสำเร็จ: ' + pdfFile.getUrl());
    return pdfFile.getUrl();
}

/**
 * สร้าง PDF สำหรับ 1 คน (เรียกจาก UI ทีละคน — backward compatible)
 */
function generateOnePageReportPdf(studentId) {
  try {
    var cache = _opr_preloadData();
    var term = (arguments.length >= 2 && arguments[1] !== undefined && arguments[1] !== null) ? String(arguments[1]).trim() : 'both';
    if (term !== '1' && term !== '2' && term !== 'both') term = 'both';
    return _opr_generateSingle(studentId, cache, term);
  } catch (e) {
    Logger.log('OnePageReport Error: ' + e.message + '\n' + e.stack);
    throw new Error('สร้างรายงานไม่สำเร็จ: ' + e.message);
  }
}

function _opr_getGPAByTerm_(sid, grade, classNo, cache, whClass, term) {
  var t = String(term || 'both').trim();
  if (t !== '1' && t !== '2' && t !== 'both') t = 'both';
  if (t === 'both') return _opr_getGPA(sid, grade, classNo, cache, whClass);

  if (!cache.gpaCacheTerm) cache.gpaCacheTerm = {};
  var classKey = String(grade || '') + '_' + String(classNo || '') + '_' + t;
  if (!cache.gpaCacheTerm[classKey]) {
    var creditMap = {};
    var typeMap = {};
    Object.keys(cache.sInfo).forEach(function(code) {
      creditMap[code] = cache.sInfo[code].hours || 1;
      typeMap[code] = cache.sInfo[code].type || 'พื้นฐาน';
    });

    var classStudents = {};
    (whClass || []).forEach(function(r) {
      var rsid = String(r['student_id'] || '').trim();
      if (!rsid) return;
      if (!classStudents[rsid]) classStudents[rsid] = [];
      classStudents[rsid].push(r);
    });

    var gpaList = [];
    Object.keys(classStudents).forEach(function(stId) {
      var totCredits = 0, totPoints = 0;
      classStudents[stId].forEach(function(r) {
        var code = String(r['subject_code'] || r['subjectCode'] || '').trim();
        var subType = typeMap[code] || String(r['subject_type'] || '').trim() || 'พื้นฐาน';
        if (subType === 'กิจกรรม') return;

        var credits = creditMap[code] || 1;
        var score = 0;
        if (t === '1') score = parseFloat(r['term1_total'] || r['term1Total'] || 0) || 0;
        else score = parseFloat(r['term2_total'] || r['term2Total'] || 0) || 0;

        var gpaVal = _scoreToGPA(score).gpa;
        totCredits += credits;
        totPoints += gpaVal * credits;
      });
      var gpa = totCredits > 0 ? totPoints / totCredits : 0;
      gpaList.push({ id: stId, gpa: parseFloat(gpa.toFixed(2)) });
    });

    gpaList.sort(function(a, b) { return b.gpa - a.gpa; });
    var ranked = {};
    gpaList.forEach(function(g, i) {
      ranked[g.id] = { gpa: g.gpa, classRank: i + 1, totalStudents: gpaList.length };
    });
    cache.gpaCacheTerm[classKey] = ranked;
  }

  var r = cache.gpaCacheTerm[classKey][sid];
  return r || { gpa: 0, classRank: '-', totalStudents: 0 };
}

/**
 * ดึงผลประเมิน 3 ด้าน จาก pre-loaded data (ไม่อ่านชีทซ้ำ)
 */
function _opr_getAssessments(sid, cache) {
  var result = {
    reading: { result: 'ไม่พบข้อมูล', score: 0 },
    character: { result: 'ไม่พบข้อมูล', score: 0 },
    activities: { result: 'ไม่ผ่าน' }
  };

  // คุณลักษณะ
  try {
    var charRec = cache.characterData.find(function(r) { return String(r['รหัสนักเรียน']).trim() === sid; });
    if (charRec) {
      result.character = { result: charRec['ผลการประเมิน'] || 'ปรับปรุง', score: parseFloat(charRec['คะแนนเฉลี่ย']) || 0 };
    }
  } catch(e){}

  // อ่านคิดเขียน
  try {
    var readRec = cache.readingData.find(function(r) { return String(r['รหัสนักเรียน']).trim() === sid; });
    if (readRec) {
      var allKeys = Object.keys(readRec);
      var resultKey = allKeys.find(function(k) { return k.toString().trim() === 'สรุปผลการประเมิน' || k.toString().trim().indexOf('ผลการประเมิน') !== -1; });
      var summaryResult = resultKey ? readRec[resultKey].toString().trim() : '';
      var avgKey = allKeys.find(function(k) { return k.toString().trim().indexOf('คะแนนเฉลี่ย') !== -1; });
      var avgScore = avgKey ? parseFloat(readRec[avgKey]) : 0;
      if (!summaryResult) {
        var scoreIdxs = [4,5,6,7,8,9,10,11];
        var nums = scoreIdxs.map(function(i) { var v = readRec[i]; var n = parseFloat(v ? v.toString().trim() : ''); return Number.isFinite(n) ? n : null; }).filter(function(x) { return x !== null; });
        var avg = nums.length > 0 ? nums.reduce(function(a,b){return a+b;},0)/nums.length : 0;
        summaryResult = avg >= 2.5 ? 'ดีเยี่ยม' : avg >= 2.0 ? 'ดี' : avg >= 1.0 ? 'ผ่าน' : 'ปรับปรุง';
        avgScore = avg;
      }
      result.reading = { result: summaryResult, score: avgScore };
    }
  } catch(e){}

  // กิจกรรม
  try {
    var actRec = cache.activityData.find(function(r) { return String(r['รหัสนักเรียน']).trim() === sid; });
    if (actRec) {
      result.activities = {
        result: actRec['รวมกิจกรรม'] || 'ผ่าน',
        แนะแนว: actRec['กิจกรรมแนะแนว'] || 'ผ่าน',
        ลูกเสือ: actRec['ลูกเสือ_เนตรนารี'] || 'ผ่าน',
        ชุมนุม: actRec['ชุมนุม'] || 'ผ่าน',
        สาธารณะ: actRec['เพื่อสังคมและสาธารณประโยชน์'] || 'ผ่าน'
      };
    }
  } catch(e){}

  return result;
}

/**
 * คำนวณ GPA + อันดับ จาก pre-loaded data (cache per class)
 * ใช้ logic เดียวกับ calculateGPAAndRank ต้นฉบับ
 */
function _opr_getGPA(sid, grade, classNo, cache) {
  var classKey = grade + '_' + classNo;
  if (!cache.gpaCache[classKey]) {
    // สร้าง creditMap จาก sInfo (ชั่วโมง/ปี = หน่วยกิต)
    var creditMap = {};
    var typeMap = {};
    Object.keys(cache.sInfo).forEach(function(code) {
      creditMap[code] = cache.sInfo[code].hours || 1;
      typeMap[code] = cache.sInfo[code].type || 'พื้นฐาน';
    });

    var classStudents = {};
    (arguments.length >= 5 ? arguments[4] : []).forEach(function(r) {
      var rsid = String(r['student_id'] || '').trim();
      if (!rsid) return;
      if (!classStudents[rsid]) classStudents[rsid] = [];
      classStudents[rsid].push(r);
    });

    // คำนวณ weighted GPA (เหมือน calcWeightedGPA ใน calculateGPAAndRank)
    var gpaList = [];
    Object.keys(classStudents).forEach(function(stId) {
      var totCredits = 0, totPoints = 0;
      classStudents[stId].forEach(function(r) {
        var code = String(r['subject_code'] || r['subjectCode'] || '').trim();
        var subType = typeMap[code] || 'พื้นฐาน';
        if (subType === 'กิจกรรม') return;

        var credits = creditMap[code] || 1;
        // ใช้ final_grade ก่อน (เหมือนต้นฉบับ)
        var gpaVal = 0;
        var finalGrade = parseFloat(r['final_grade'] || r['grade_result'] || '');
        if (!isNaN(finalGrade)) {
          gpaVal = finalGrade;
        } else {
          var avg = parseFloat(r['average'] || r['total'] || 0);
          gpaVal = _scoreToGPA(avg).gpa;
        }

        totCredits += credits;
        totPoints += gpaVal * credits;
      });
      var gpa = totCredits > 0 ? totPoints / totCredits : 0;
      gpaList.push({ id: stId, gpa: parseFloat(gpa.toFixed(2)) });
    });

    // เรียงอันดับ
    gpaList.sort(function(a, b) { return b.gpa - a.gpa; });
    var ranked = {};
    gpaList.forEach(function(g, i) {
      ranked[g.id] = { gpa: g.gpa, classRank: i + 1, totalStudents: gpaList.length };
    });
    cache.gpaCache[classKey] = ranked;
  }

  var r = cache.gpaCache[classKey][sid];
  return r || { gpa: 0, classRank: '-', totalStudents: 0 };
}
