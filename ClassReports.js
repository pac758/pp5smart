// ═══════════════════════════════════════════════════════════════
// ClassReports.js — รายงานสรุปคะแนนและการประเมินรายชั้น
// ═══════════════════════════════════════════════════════════════
// ⚠️ ไฟล์นี้สร้างใหม่ทั้งหมด ไม่แก้ไขฟังก์ชัน PDF เดิม
// ═══════════════════════════════════════════════════════════════

/**
 * 1. สรุปคะแนนวิชาพื้นฐาน — เก็บ/สอบ/รวม/เกรด ทุกวิชา แยกตามห้อง
 * @param {string} grade - ชั้น เช่น "ป.1"
 * @param {string} classNo - ห้อง เช่น "1"
 * @param {string} term - ภาคเรียน "1" หรือ "2"
 * @returns {Object}
 */
function getClassSubjectScoreSummary(grade, classNo, term) {
  try {
    var ss = SS();
    var termNum = parseInt(term) || 1;

    // 1. ดึงรายวิชาพื้นฐาน+เพิ่มเติม (ไม่รวมกิจกรรม) สำหรับชั้นนี้
    var subjectSheet = ss.getSheetByName('รายวิชา');
    if (!subjectSheet) throw new Error('ไม่พบชีต "รายวิชา"');
    var subjectData = subjectSheet.getDataRange().getValues();
    var subjectHeaders = subjectData[0];
    var idxGrade = subjectHeaders.indexOf('ชั้น');
    var idxName = subjectHeaders.indexOf('ชื่อวิชา');
    var idxCode = subjectHeaders.indexOf('รหัสวิชา');
    var idxType = subjectHeaders.indexOf('ประเภทวิชา');
    var idxOrder = subjectHeaders.indexOf('ลำดับ');
    var idxMidMax = subjectHeaders.indexOf('คะแนนระหว่างปี');
    var idxFinalMax = subjectHeaders.indexOf('คะแนนปลายปี');

    var subjects = [];
    for (var i = 1; i < subjectData.length; i++) {
      var row = subjectData[i];
      var sGrade = String(row[idxGrade] || '').trim();
      var sType = String(row[idxType] || '').trim();
      if (sGrade !== grade) continue;
      if (sType === 'กิจกรรม') continue;
      subjects.push({
        name: String(row[idxName] || '').trim(),
        code: String(row[idxCode] || '').trim(),
        type: sType,
        order: parseInt(row[idxOrder]) || 99,
        midMax: parseInt(row[idxMidMax]) || 60,
        finalMax: parseInt(row[idxFinalMax]) || 40
      });
    }
    subjects.sort(function(a, b) { return a.order - b.order; });

    if (subjects.length === 0) {
      return { success: false, message: 'ไม่พบรายวิชาสำหรับชั้น ' + grade };
    }

    // 2. ดึงรายชื่อนักเรียน
    var studentsSheet = ss.getSheetByName('Students');
    if (!studentsSheet) throw new Error('ไม่พบชีต "Students"');
    var studentData = studentsSheet.getDataRange().getValues();
    var stHeaders = studentData[0];
    var stGradeCol = U_findColumnIndex(stHeaders, ['grade', 'ชั้น', 'ระดับชั้น']);
    var stClassCol = U_findColumnIndex(stHeaders, ['class_no', 'ห้อง']);
    var stIdCol = U_findColumnIndex(stHeaders, ['student_id', 'รหัสนักเรียน']);
    var stNoCol = U_findColumnIndex(stHeaders, ['student_no', 'เลขที่', 'ลำดับ']);
    var stTitleCol = U_findColumnIndex(stHeaders, ['title', 'คำนำหน้า']);
    var stFirstCol = U_findColumnIndex(stHeaders, ['firstname', 'first_name', 'ชื่อ']);
    var stLastCol = U_findColumnIndex(stHeaders, ['lastname', 'last_name', 'นามสกุล']);
    var stFullCol = U_findColumnIndex(stHeaders, ['full_name', 'ชื่อเต็ม', 'ชื่อ-นามสกุล']);

    var students = [];
    for (var si = 1; si < studentData.length; si++) {
      var sr = studentData[si];
      if (String(sr[stGradeCol] || '').trim() !== grade) continue;
      if (String(sr[stClassCol] || '').trim() !== classNo) continue;
      var sId = String(sr[stIdCol] || '').trim();
      var sNo = stNoCol !== -1 ? String(sr[stNoCol] || '').trim() : '';
      var fullName = '';
      if (stFullCol !== -1 && sr[stFullCol]) {
        fullName = String(sr[stFullCol]).trim();
      } else if (stFirstCol !== -1 && stLastCol !== -1) {
        fullName = (String(sr[stFirstCol] || '') + ' ' + String(sr[stLastCol] || '')).trim();
      }
      var title = stTitleCol !== -1 ? String(sr[stTitleCol] || '').trim() : '';
      if (title && fullName) fullName = title + fullName;
      students.push({ studentId: sId, studentNo: sNo, fullName: fullName, scores: {} });
    }
    students.sort(function(a, b) {
      var nA = parseInt(a.studentNo) || 0;
      var nB = parseInt(b.studentNo) || 0;
      return nA - nB || String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true });
    });

    if (students.length === 0) {
      return { success: false, message: 'ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo };
    }

    // 3. สร้าง map studentId → index เพื่อจับคู่เร็ว
    var studentMap = {};
    students.forEach(function(s, idx) { studentMap[s.studentId] = idx; });

    // 4. อ่านคะแนนจากชีตรายวิชา
    var gradeNoSuffix = grade.replace(/\./g, '');
    subjects.forEach(function(subj) {
      // ลองหาชีตหลายรูปแบบ
      var possibleNames = [
        subj.name + ' ' + gradeNoSuffix + '-' + classNo,
        subj.name + ' ' + grade + '-' + classNo,
        subj.code + ' ' + gradeNoSuffix + '-' + classNo
      ];
      var scoreSheet = null;
      for (var p = 0; p < possibleNames.length; p++) {
        scoreSheet = ss.getSheetByName(possibleNames[p]);
        if (scoreSheet) break;
      }
      if (!scoreSheet) {
        // fuzzy search
        var allSheets = ss.getSheets();
        for (var a = 0; a < allSheets.length; a++) {
          var sn = allSheets[a].getName();
          if ((sn.indexOf(subj.name) !== -1 || (subj.code && sn.indexOf(subj.code) !== -1)) &&
              sn.indexOf(classNo) !== -1) {
            scoreSheet = allSheets[a];
            break;
          }
        }
      }
      if (!scoreSheet) {
        Logger.log('⚠️ ไม่พบชีตคะแนนสำหรับ: ' + subj.name);
        return;
      }

      var layout = detectSheetLayout_(scoreSheet);
      var termCols = termNum === 2 ? layout.term2 : layout.term1;
      var data = scoreSheet.getDataRange().getValues();

      // อ่านคะแนนเต็มจาก row 4 (index 3)
      var midMaxFromSheet = Number(data[3] ? data[3][termCols.midTotal] : 0) || 0;
      var finalMaxFromSheet = Number(data[3] ? data[3][termCols.s10] : 0) || 0;
      if (midMaxFromSheet > 0) subj.midMax = midMaxFromSheet;
      if (finalMaxFromSheet > 0) subj.finalMax = finalMaxFromSheet;

      for (var r = 4; r < data.length; r++) {
        var row = data[r];
        var sid = String(row[1] || '').trim();
        if (!sid) continue;
        var idx = studentMap[sid];
        if (idx === undefined) continue;

        students[idx].scores[subj.code || subj.name] = {
          mid: Number(row[termCols.midTotal]) || 0,
          final: Number(row[termCols.s10]) || 0,
          total: Number(row[termCols.total]) || 0,
          grade: String(row[termCols.grade] || '')
        };
      }
    });

    // 5. ดึงข้อมูลโรงเรียน
    var settings = S_getGlobalSettings();

    return {
      success: true,
      grade: grade,
      classNo: classNo,
      term: termNum,
      gradeFullName: U_getGradeFullName(grade),
      subjects: subjects,
      students: students,
      settings: settings
    };
  } catch (e) {
    Logger.log('❌ getClassSubjectScoreSummary error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * 2. สรุปการประเมินทางการเรียน — GPA, ลำดับ, อ่านฯ, คุณลักษณะฯ, สมรรถนะฯ
 * @param {string} grade - ชั้น
 * @param {string} classNo - ห้อง
 * @returns {Object}
 */
function getClassAssessmentSummary(grade, classNo) {
  try {
    var ss = SS();

    // 1. ดึงรายวิชาพื้นฐาน+เพิ่มเติม (ไม่รวมกิจกรรม)
    var subjectSheet = ss.getSheetByName('รายวิชา');
    if (!subjectSheet) throw new Error('ไม่พบชีต "รายวิชา"');
    var subjectData = subjectSheet.getDataRange().getValues();
    var sjHeaders = subjectData[0];
    var sjGradeIdx = sjHeaders.indexOf('ชั้น');
    var sjNameIdx = sjHeaders.indexOf('ชื่อวิชา');
    var sjCodeIdx = sjHeaders.indexOf('รหัสวิชา');
    var sjTypeIdx = sjHeaders.indexOf('ประเภทวิชา');
    var sjHoursIdx = sjHeaders.indexOf('ชั่วโมง/ปี');

    var academicSubjects = [];
    var creditMap = {};
    for (var si = 1; si < subjectData.length; si++) {
      var sr = subjectData[si];
      if (String(sr[sjGradeIdx] || '').trim() !== grade) continue;
      var sType = String(sr[sjTypeIdx] || '').trim();
      if (sType === 'กิจกรรม') continue;
      var code = String(sr[sjCodeIdx] || '').trim();
      var name = String(sr[sjNameIdx] || '').trim();
      var hours = parseFloat(sr[sjHoursIdx]) || 1;
      academicSubjects.push({ code: code, name: name, hours: hours });
      creditMap[code] = hours;
      creditMap[name] = hours;
    }

    // 2. ดึงรายชื่อนักเรียน
    var studentsSheet = ss.getSheetByName('Students');
    if (!studentsSheet) throw new Error('ไม่พบชีต "Students"');
    var studentData = studentsSheet.getDataRange().getValues();
    var stHeaders = studentData[0];
    var stGradeCol = U_findColumnIndex(stHeaders, ['grade', 'ชั้น', 'ระดับชั้น']);
    var stClassCol = U_findColumnIndex(stHeaders, ['class_no', 'ห้อง']);
    var stIdCol = U_findColumnIndex(stHeaders, ['student_id', 'รหัสนักเรียน']);
    var stNoCol = U_findColumnIndex(stHeaders, ['student_no', 'เลขที่']);
    var stTitleCol = U_findColumnIndex(stHeaders, ['title', 'คำนำหน้า']);
    var stFirstCol = U_findColumnIndex(stHeaders, ['firstname', 'first_name', 'ชื่อ']);
    var stLastCol = U_findColumnIndex(stHeaders, ['lastname', 'last_name', 'นามสกุล']);
    var stFullCol = U_findColumnIndex(stHeaders, ['full_name', 'ชื่อเต็ม', 'ชื่อ-นามสกุล']);

    var classStudents = [];
    for (var ci = 1; ci < studentData.length; ci++) {
      var cr = studentData[ci];
      if (String(cr[stGradeCol] || '').trim() !== grade) continue;
      if (String(cr[stClassCol] || '').trim() !== classNo) continue;
      var sId = String(cr[stIdCol] || '').trim();
      var sNo = stNoCol !== -1 ? String(cr[stNoCol] || '').trim() : '';
      var fullName = '';
      if (stFullCol !== -1 && cr[stFullCol]) {
        fullName = String(cr[stFullCol]).trim();
      } else if (stFirstCol !== -1 && stLastCol !== -1) {
        fullName = (String(cr[stFirstCol] || '') + ' ' + String(cr[stLastCol] || '')).trim();
      }
      var title = stTitleCol !== -1 ? String(cr[stTitleCol] || '').trim() : '';
      if (title && fullName) fullName = title + fullName;
      classStudents.push({ studentId: sId, studentNo: sNo, fullName: fullName });
    }
    classStudents.sort(function(a, b) {
      var nA = parseInt(a.studentNo) || 0;
      var nB = parseInt(b.studentNo) || 0;
      return nA - nB || String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true });
    });

    if (classStudents.length === 0) {
      return { success: false, message: 'ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo };
    }

    // 3. ดึงข้อมูลจาก SCORES_WAREHOUSE
    var warehouseSheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!warehouseSheet) throw new Error('ไม่พบชีต SCORES_WAREHOUSE');
    var whData = warehouseSheet.getDataRange().getValues();
    var whHeaders = whData[0];
    var whIdCol = U_findColumnIndex(whHeaders, ['student_id', 'studentId', 'รหัสนักเรียน']);
    var whGradeCol = U_findColumnIndex(whHeaders, ['grade', 'ชั้น']);
    var whClassCol = U_findColumnIndex(whHeaders, ['class_no', 'ห้อง']);
    var whSubjectCol = U_findColumnIndex(whHeaders, ['subject_name', 'วิชา', 'ชื่อวิชา']);
    var whCodeCol = U_findColumnIndex(whHeaders, ['subject_code', 'รหัสวิชา']);
    var whTypeCol = U_findColumnIndex(whHeaders, ['subject_type', 'ประเภทวิชา']);
    var whT1Col = U_findColumnIndex(whHeaders, ['term1_total', 'term1Total']);
    var whT2Col = U_findColumnIndex(whHeaders, ['term2_total', 'term2Total']);
    var whAvgCol = U_findColumnIndex(whHeaders, ['average', 'avg']);
    var whFGCol = U_findColumnIndex(whHeaders, ['final_grade', 'grade_result']);

    // สร้าง map: studentId → [{ subject, score, gpa }]
    var scoresByStudent = {};
    // สร้าง map สำหรับทุกนักเรียนในชั้นเดียวกัน (ทุกห้อง) เพื่อคำนวณลำดับชั้น
    var allGradeStudentIds = new Set();
    for (var ai = 1; ai < studentData.length; ai++) {
      if (String(studentData[ai][stGradeCol] || '').trim() === grade) {
        allGradeStudentIds.add(String(studentData[ai][stIdCol] || '').trim());
      }
    }

    for (var wi = 1; wi < whData.length; wi++) {
      var wr = whData[wi];
      var wGrade = String(wr[whGradeCol] || '').trim();
      if (wGrade !== grade) continue;

      var wId = String(wr[whIdCol] || '').trim();
      var wSubject = String(wr[whSubjectCol] || '').trim();
      var wCode = whCodeCol !== -1 ? String(wr[whCodeCol] || '').trim() : '';
      var wType = whTypeCol !== -1 ? String(wr[whTypeCol] || '').trim() : '';

      if (wType === 'กิจกรรม') continue;
      var actKw = ['แนะแนว', 'ลูกเสือ', 'เนตรนารี', 'ชุมนุม', 'กิจกรรมเพื่อสังคม', 'บำเพ็ญประโยชน์'];
      var isAct = actKw.some(function(k) { return wSubject.indexOf(k) !== -1; });
      if (isAct) continue;

      var t1 = whT1Col !== -1 ? (parseFloat(wr[whT1Col]) || 0) : 0;
      var t2 = whT2Col !== -1 ? (parseFloat(wr[whT2Col]) || 0) : 0;
      var avg = whAvgCol !== -1 ? (parseFloat(wr[whAvgCol]) || 0) : 0;
      var fg = whFGCol !== -1 ? parseFloat(wr[whFGCol]) : NaN;
      var totalScore = t1 + t2;
      if (totalScore === 0) totalScore = avg;

      var gpa = !isNaN(fg) ? fg : _scoreToGPA(avg).gpa;
      var credits = creditMap[wCode] || creditMap[wSubject] || 1;

      if (!scoresByStudent[wId]) scoresByStudent[wId] = [];
      scoresByStudent[wId].push({ totalScore: totalScore, gpa: gpa, credits: credits });
    }

    // 4. คำนวณ GPA ถ่วงน้ำหนัก + คะแนนรวม สำหรับทุกนักเรียนในชั้น (ทุกห้อง)
    var totalFullScore = academicSubjects.reduce(function(sum, s) { return sum + 100; }, 0);
    // Note: ถ้าจะใช้ full score จริง ต้องรวม max per subject; สมมติ 100 ต่อวิชา

    function calcStudentSummary(sid) {
      var scores = scoresByStudent[sid] || [];
      var sumTotal = 0, sumCredits = 0, sumGradePoints = 0;
      scores.forEach(function(s) {
        sumTotal += s.totalScore;
        sumCredits += s.credits;
        sumGradePoints += s.gpa * s.credits;
      });
      var gpa = sumCredits > 0 ? sumGradePoints / sumCredits : 0;
      var fullScore = scores.length * 100;
      var pct = fullScore > 0 ? (sumTotal / fullScore) * 100 : 0;
      return { sumTotal: sumTotal, fullScore: fullScore, pct: pct, gpa: gpa };
    }

    // คำนวณ GPA ทุกนักเรียนในชั้นเดียวกัน (ทุกห้อง) เพื่อจัดลำดับชั้น
    var gradeGPAs = [];
    allGradeStudentIds.forEach(function(sid) {
      var s = calcStudentSummary(sid);
      gradeGPAs.push({ id: sid, gpa: s.gpa });
    });
    gradeGPAs.sort(function(a, b) { return b.gpa - a.gpa; });

    // คำนวณ GPA นักเรียนในห้องเดียวกัน เพื่อจัดลำดับห้อง
    var classGPAs = classStudents.map(function(st) {
      var s = calcStudentSummary(st.studentId);
      return { id: st.studentId, gpa: s.gpa };
    });
    classGPAs.sort(function(a, b) { return b.gpa - a.gpa; });

    // 5. ดึงผลประเมิน 3 ด้าน
    var readingData = [], charData = [], compData = [];
    try { readingData = _readSheetToObjects('การประเมินอ่านคิดเขียน'); } catch(e) {}
    try { charData = _readSheetToObjects('การประเมินคุณลักษณะ'); } catch(e) {}
    try { compData = _readSheetToObjects('การประเมินสมรรถนะ'); } catch(e) {}

    function findAssessment(dataArr, sid) {
      return dataArr.find(function(r) { return String(r['รหัสนักเรียน']).trim() === sid; });
    }

    function getReadingResult(sid) {
      var rec = findAssessment(readingData, sid);
      if (!rec) return '-';
      var keys = Object.keys(rec);
      var rKey = keys.find(function(k) { return k.trim() === 'สรุปผลการประเมิน' || k.indexOf('ผลการประเมิน') !== -1; });
      if (rKey) {
        var v = String(rec[rKey]).trim();
        if (v) return v;
      }
      // fallback calculation
      var scoreIdxs = [4,5,6,7,8,9,10,11];
      var nums = scoreIdxs.map(function(i) { var val = rec[i]; var n = parseFloat(val ? val.toString().trim() : ''); return Number.isFinite(n) ? n : null; }).filter(function(x) { return x !== null; });
      if (nums.length === 0) return '-';
      var avg = nums.reduce(function(a,b) { return a+b; }, 0) / nums.length;
      return avg >= 2.5 ? 'ดีเยี่ยม' : avg >= 2.0 ? 'ดี' : avg >= 1.0 ? 'ผ่าน' : 'ปรับปรุง';
    }

    function getCharResult(sid) {
      var rec = findAssessment(charData, sid);
      return rec ? (rec['ผลการประเมิน'] || '-') : '-';
    }

    function getCompResult(sid) {
      var rec = findAssessment(compData, sid);
      if (!rec) return '-';
      // หาผลรวม — ลองหา column ที่มี "สรุป" หรือ "ผลรวม"
      var keys = Object.keys(rec);
      var rKey = keys.find(function(k) { return k.indexOf('สรุป') !== -1 || k.indexOf('ผลรวม') !== -1 || k.indexOf('ผลการประเมิน') !== -1; });
      if (rKey) return String(rec[rKey]).trim() || '-';
      return '-';
    }

    // 6. รวมผลลัพธ์สำหรับแต่ละนักเรียน
    var results = classStudents.map(function(st) {
      var sid = st.studentId;
      var summary = calcStudentSummary(sid);

      var classRank = classGPAs.findIndex(function(g) { return g.id === sid; }) + 1;
      var gradeRank = gradeGPAs.findIndex(function(g) { return g.id === sid; }) + 1;

      return {
        studentNo: st.studentNo,
        fullName: st.fullName,
        sumTotal: Math.round(summary.sumTotal * 100) / 100,
        fullScore: summary.fullScore,
        pct: Math.round(summary.pct * 100) / 100,
        gpa: Math.round(summary.gpa * 100) / 100,
        classRank: classRank,
        gradeRank: gradeRank,
        reading: getReadingResult(sid),
        character: getCharResult(sid),
        competency: getCompResult(sid)
      };
    });

    var settings = S_getGlobalSettings();

    return {
      success: true,
      grade: grade,
      classNo: classNo,
      gradeFullName: U_getGradeFullName(grade),
      totalClassStudents: classStudents.length,
      totalGradeStudents: gradeGPAs.length,
      students: results,
      settings: settings
    };
  } catch (e) {
    Logger.log('❌ getClassAssessmentSummary error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * สร้าง PDF สรุปคะแนนวิชาพื้นฐาน (Report 1)
 */
function exportClassSubjectScorePDF(grade, classNo, term) {
  var data = getClassSubjectScoreSummary(grade, classNo, term);
  if (!data.success) throw new Error(data.message);

  var settings = data.settings || {};
  var schoolName = settings.school_name || '';
  var year = settings.academic_year || '';
  var termNum = data.term;
  var subjects = data.subjects;
  var students = data.students;

  // ดึงครูประจำชั้น
  var homeroomTeacher = '';
  try { homeroomTeacher = getHomeroomTeacher(grade, classNo); } catch(e) {}

  // สร้าง HTML
  var subjectCount = subjects.length;
  var totalCols = 2 + (subjectCount * 4); // เลขที่ + ชื่อ + 4 cols per subject

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>'
    + '@page { size: A4 landscape; margin: 15mm 10mm; }'
    + 'body { font-family: "Sarabun", sans-serif; font-size: 11pt; margin: 0; }'
    + '.title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 4px; }'
    + '.info { text-align: center; font-size: 11pt; margin-bottom: 2px; }'
    + 'table { border-collapse: collapse; width: 100%; font-size: 9pt; }'
    + 'th, td { border: 1px solid #000; padding: 3px 4px; text-align: center; }'
    + 'th { background: #f0f0f0; font-weight: bold; }'
    + '.name-col { text-align: left; white-space: nowrap; }'
    + '.subject-header { background: #e8e8e8; }'
    + '</style></head><body>'
    + '<div class="title">สรุปคะแนน วิชาพื้นฐาน (' + termNum + ')</div>'
    + '<div class="info">ห้อง ' + data.gradeFullName + '/' + classNo
    + ' &nbsp; ปีการศึกษา ' + termNum + '/' + year
    + ' &nbsp; ระดับ ' + (grade.indexOf('ป.') === 0 ? 'ประถมศึกษา' : 'มัธยมศึกษา') + '</div>'
    + '<div class="info">ครูประจำชั้น ' + homeroomTeacher
    + ' &nbsp; โรงเรียน' + schoolName + '</div>';

  // สร้างตาราง
  html += '<table>';

  // แถวหัวข้อชื่อวิชา
  html += '<thead><tr><th rowspan="2" style="width:30px">เลขที่</th>'
    + '<th rowspan="2" style="min-width:120px">ชื่อ-สกุล</th>';

  var subjectShortNames = {
    'ภาษาไทย': 'ภาษาไทย',
    'คณิตศาสตร์': 'คณิตศาสตร์',
    'วิทยาศาสตร์และเทคโนโลยี': 'วิทย์และเทคโนฯ',
    'สังคมศึกษา ศาสนา และวัฒนธรรม': 'สังคมศึกษาฯ',
    'สุขศึกษาและพลศึกษา': 'สุขศึกษาฯ',
    'ศิลปะ': 'ศิลปะ',
    'การงานอาชีพ': 'การงานอาชีพ',
    'ภาษาอังกฤษ': 'ภาษาอังกฤษ',
    'ประวัติศาสตร์': 'ประวัติศาสตร์',
    'หน้าที่พลเมือง': 'หน้าที่พลเมือง'
  };

  subjects.forEach(function(subj) {
    var shortName = subjectShortNames[subj.name] || (subj.name.length > 15 ? subj.name.substring(0, 15) + 'ฯ' : subj.name);
    html += '<th colspan="4" class="subject-header">' + shortName + '</th>';
  });
  html += '</tr>';

  // แถวหัวข้อ เก็บ/สอบ/รวม/เกรด
  html += '<tr>';
  subjects.forEach(function(subj) {
    html += '<th>เก็บ' + termNum + '<br>(' + subj.midMax + ')</th>';
    html += '<th>สอบ' + termNum + '<br>(' + subj.finalMax + ')</th>';
    html += '<th>รวม' + termNum + '<br>(' + (subj.midMax + subj.finalMax) + ')</th>';
    html += '<th>เกรด' + termNum + '</th>';
  });
  html += '</tr></thead>';

  // ข้อมูลนักเรียน
  html += '<tbody>';
  students.forEach(function(st, idx) {
    html += '<tr><td>' + (idx + 1) + '</td>';
    html += '<td class="name-col">' + st.fullName + '</td>';
    subjects.forEach(function(subj) {
      var key = subj.code || subj.name;
      var sc = st.scores[key];
      if (sc) {
        html += '<td>' + sc.mid + '</td>';
        html += '<td>' + sc.final + '</td>';
        html += '<td><b>' + sc.total + '</b></td>';
        html += '<td><b>' + sc.grade + '</b></td>';
      } else {
        html += '<td>-</td><td>-</td><td>-</td><td>-</td>';
      }
    });
    html += '</tr>';
  });
  html += '</tbody></table></body></html>';

  // สร้าง PDF
  var blob = HtmlService.createHtmlOutput(html)
    .setTitle('สรุปคะแนนวิชาพื้นฐาน')
    .getBlob()
    .setName('สรุปคะแนน_' + grade + '_' + classNo + '_ภาค' + termNum + '.pdf')
    .getAs('application/pdf');

  var folder = getOutputFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/**
 * สร้าง PDF สรุปการประเมินทางการเรียน (Report 2)
 */
function exportClassAssessmentSummaryPDF(grade, classNo) {
  var data = getClassAssessmentSummary(grade, classNo);
  if (!data.success) throw new Error(data.message);

  var settings = data.settings || {};
  var schoolName = settings.school_name || '';
  var year = settings.academic_year || '';

  var homeroomTeacher = '';
  try { homeroomTeacher = getHomeroomTeacher(grade, classNo); } catch(e) {}

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>'
    + '@page { size: A4 landscape; margin: 15mm 10mm; }'
    + 'body { font-family: "Sarabun", sans-serif; font-size: 11pt; margin: 0; }'
    + '.title { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 4px; }'
    + '.info { text-align: center; font-size: 11pt; margin-bottom: 2px; }'
    + 'table { border-collapse: collapse; width: 100%; font-size: 10pt; }'
    + 'th, td { border: 1px solid #000; padding: 4px 6px; text-align: center; }'
    + 'th { background: #f0f0f0; font-weight: bold; }'
    + '.name-col { text-align: left; white-space: nowrap; }'
    + '.good { color: #059669; } .improve { color: #dc3545; }'
    + '</style></head><body>'
    + '<div class="title">สรุปการประเมินทางการเรียน</div>'
    + '<div class="info">ห้อง ' + data.gradeFullName + '/' + classNo
    + ' &nbsp; ปีการศึกษา ' + year
    + ' &nbsp; ระดับ ' + (data.grade.indexOf('ป.') === 0 ? 'ประถมศึกษา' : 'มัธยมศึกษา') + '</div>'
    + '<div class="info">ครูประจำชั้น ' + homeroomTeacher
    + ' &nbsp; โรงเรียน' + schoolName + '</div>';

  html += '<table><thead>';
  html += '<tr>'
    + '<th rowspan="2" style="width:35px">เลขที่</th>'
    + '<th rowspan="2" style="min-width:140px">ชื่อ-สกุล</th>'
    + '<th colspan="6">สรุปการประเมิน</th>'
    + '<th rowspan="2">อ่านฯ</th>'
    + '<th rowspan="2">คุณลักษณะฯ</th>'
    + '<th rowspan="2">สมรรถนะฯ</th>'
    + '</tr>';
  html += '<tr>'
    + '<th>รวมคะแนน</th>'
    + '<th>เต็ม</th>'
    + '<th>ร้อยละ</th>'
    + '<th>เกรดเฉลี่ย<br>GPA</th>'
    + '<th>ลำดับ<br>ห้อง</th>'
    + '<th>ลำดับ<br>ชั้น</th>'
    + '</tr></thead>';

  html += '<tbody>';
  data.students.forEach(function(st, idx) {
    html += '<tr>';
    html += '<td>' + (idx + 1) + '</td>';
    html += '<td class="name-col">' + st.fullName + '</td>';
    html += '<td>' + st.sumTotal + '</td>';
    html += '<td>' + st.fullScore + '</td>';
    html += '<td>' + st.pct.toFixed(2) + '</td>';
    html += '<td><b>' + st.gpa.toFixed(2) + '</b></td>';
    html += '<td>' + st.classRank + '</td>';
    html += '<td>' + st.gradeRank + '</td>';

    var readCls = (st.reading === 'ดีเยี่ยม' || st.reading === 'ดี') ? 'good' : (st.reading === 'ปรับปรุง' ? 'improve' : '');
    var charCls = (st.character === 'ดีเยี่ยม' || st.character === 'ดี') ? 'good' : (st.character === 'ปรับปรุง' ? 'improve' : '');
    var compCls = (st.competency === 'ดีเยี่ยม' || st.competency === 'ดี') ? 'good' : (st.competency === 'ปรับปรุง' ? 'improve' : '');

    html += '<td class="' + readCls + '">' + st.reading + '</td>';
    html += '<td class="' + charCls + '">' + st.character + '</td>';
    html += '<td class="' + compCls + '">' + st.competency + '</td>';
    html += '</tr>';
  });
  html += '</tbody></table></body></html>';

  var blob = HtmlService.createHtmlOutput(html)
    .setTitle('สรุปการประเมินทางการเรียน')
    .getBlob()
    .setName('สรุปประเมิน_' + grade + '_' + classNo + '.pdf')
    .getAs('application/pdf');

  var folder = getOutputFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}
