// ═══════════════════════════════════════════════════════════════
// ClassReports.js — รายงานสรุปคะแนนและการประเมินรายชั้น
// ═══════════════════════════════════════════════════════════════
// ⚠️ ไฟล์นี้สร้างใหม่ทั้งหมด ไม่แก้ไขฟังก์ชัน PDF เดิม
// ⚠️ PDF ใช้ base64 data URL (ไม่ใช้ DriveApp) เพื่อเลี่ยงปัญหา authorization
// ═══════════════════════════════════════════════════════════════

/* ── helpers ภายในไฟล์ ────────────────────────────────── */
function _cr_getStudents(ss, grade, classNo) {
  var sheet = ss.getSheetByName('Students');
  if (!sheet) throw new Error('ไม่พบชีต "Students"');
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var gi = U_findColumnIndex(h, ['grade', 'ชั้น', 'ระดับชั้น']);
  var ci = U_findColumnIndex(h, ['class_no', 'ห้อง']);
  var ii = U_findColumnIndex(h, ['student_id', 'รหัสนักเรียน']);
  var ni = U_findColumnIndex(h, ['student_no', 'เลขที่', 'ลำดับ']);
  var ti = U_findColumnIndex(h, ['title', 'คำนำหน้า']);
  var fi = U_findColumnIndex(h, ['firstname', 'first_name', 'ชื่อ']);
  var li = U_findColumnIndex(h, ['lastname', 'last_name', 'นามสกุล']);
  var ui = U_findColumnIndex(h, ['full_name', 'ชื่อเต็ม', 'ชื่อ-นามสกุล']);
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[gi]||'').trim() !== grade) continue;
    if (String(row[ci]||'').trim() !== classNo) continue;
    var fn = '';
    if (ui !== -1 && row[ui]) { fn = String(row[ui]).trim(); }
    else if (fi !== -1 && li !== -1) { fn = (String(row[fi]||'')+' '+String(row[li]||'')).trim(); }
    var title = ti !== -1 ? String(row[ti]||'').trim() : '';
    if (title && fn) fn = title + fn;
    out.push({ studentId: String(row[ii]||'').trim(), studentNo: ni!==-1?String(row[ni]||'').trim():'', fullName: fn });
  }
  out.sort(function(a,b) {
    return (parseInt(a.studentNo)||0) - (parseInt(b.studentNo)||0) ||
      String(a.studentId).localeCompare(String(b.studentId), undefined, {numeric:true});
  });
  return { students: out, allData: data, headers: h, gradeCol: gi, classCol: ci, idCol: ii };
}

function _cr_getSubjects(ss, grade) {
  var sheet = ss.getSheetByName('รายวิชา');
  if (!sheet) throw new Error('ไม่พบชีต "รายวิชา"');
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var gi = h.indexOf('ชั้น'), ni = h.indexOf('ชื่อวิชา'), ci = h.indexOf('รหัสวิชา');
  var ti = h.indexOf('ประเภทวิชา'), oi = h.indexOf('ลำดับ');
  var mi = h.indexOf('คะแนนระหว่างปี'), fi = h.indexOf('คะแนนปลายปี');
  var hi = h.indexOf('ชั่วโมง/ปี');
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[gi]||'').trim() !== grade) continue;
    var typ = String(row[ti]||'').trim();
    if (typ === 'กิจกรรม') continue;
    out.push({
      name: String(row[ni]||'').trim(), code: String(row[ci]||'').trim(), type: typ,
      order: parseInt(row[oi])||99,
      midMax: parseInt(row[mi])||60, finalMax: parseInt(row[fi])||40,
      hours: parseFloat(row[hi])||1
    });
  }
  out.sort(function(a,b) { return a.order - b.order; });
  return out;
}

var _CR_SHORT = {
  'ภาษาไทย':'ภาษาไทย','คณิตศาสตร์':'คณิตศาสตร์',
  'วิทยาศาสตร์และเทคโนโลยี':'วิทย์และเทคโนฯ',
  'สังคมศึกษา ศาสนา และวัฒนธรรม':'สังคมศึกษาฯ',
  'สุขศึกษาและพลศึกษา':'สุขศึกษาฯ','ศิลปะ':'ศิลปะ',
  'การงานอาชีพ':'การงานอาชีพ','ภาษาอังกฤษ':'ภาษาอังกฤษ',
  'ประวัติศาสตร์':'ประวัติศาสตร์','หน้าที่พลเมือง':'หน้าที่พลเมือง'
};
function _cr_shortName(n) { return _CR_SHORT[n] || (n.length>15 ? n.substring(0,15)+'ฯ' : n); }

/** สร้าง PDF blob จาก HTML แล้วคืน base64 data URL (ไม่ใช้ DriveApp) */
function _cr_htmlToBase64Pdf(html) {
  // Inject embedded Sarabun font as separate <style> tag
  var fontCss = _getEmbeddedSarabunCss_();
  if (fontCss && html.indexOf('<style>') > -1) {
    html = html.replace('<style>', '<style>' + fontCss + '</style><style>');
  }
  // Remove external Google Fonts <link> tags
  html = html.replace(/<link[^>]*fonts\.googleapis\.com[^>]*>/gi, '');
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf');
  return 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes());
}

/** ดึงโลโก้โรงเรียนผ่าน URL (ไม่ใช้ DriveApp) */
function _cr_getLogoBase64(settings) {
  try {
    var url = settings['logoUrl_lh3'] || settings['logo'] || settings['schoolLogo'] || '';
    if (!url) return null;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var blob = resp.getBlob();
    var mime = blob.getContentType() || 'image/png';
    return 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch(e) { return null; }
}

/* ════════════════════════════════════════════════════════════
   1. สรุปคะแนน — เก็บ/สอบ/รวม/เกรด ทุกวิชา แยกตามห้อง
      term: "1" | "2" | "0" (ทั้งปี)
   ════════════════════════════════════════════════════════════ */
function getClassSubjectScoreSummary(grade, classNo, term) {
  try {
    var ss = SS();
    var termNum = term === '0' || term === 0 ? 0 : (parseInt(term) || 1); // 0=ทั้งปี, 1=ภาค1, 2=ภาค2

    var subjects = _cr_getSubjects(ss, grade);
    if (subjects.length === 0) return { success: false, message: 'ไม่พบรายวิชาสำหรับชั้น ' + grade };

    var stInfo = _cr_getStudents(ss, grade, classNo);
    var students = stInfo.students;
    if (students.length === 0) return { success: false, message: 'ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo };

    var studentMap = {};
    students.forEach(function(s, idx) { s.scores = {}; studentMap[s.studentId] = idx; });

    // อ่านคะแนนจากชีตรายวิชา
    var gradeNoSuffix = grade.replace(/\./g, '');
    subjects.forEach(function(subj) {
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
        var allSheets = ss.getSheets();
        for (var a = 0; a < allSheets.length; a++) {
          var sn = allSheets[a].getName();
          if (sn.indexOf('BACKUP_') === 0) continue;
          if (/_\d{4}$/.test(sn)) continue;
          if ((sn.indexOf(subj.name) !== -1 || (subj.code && sn.indexOf(subj.code) !== -1)) &&
              (sn.indexOf(gradeNoSuffix) !== -1 || sn.indexOf(grade) !== -1) &&
              sn.indexOf('-' + classNo) !== -1) { scoreSheet = allSheets[a]; break; }
        }
      }
      if (!scoreSheet) { Logger.log('⚠️ ไม่พบชีตคะแนน: ' + subj.name); return; }

      var layout = detectSheetLayout_(scoreSheet);
      var data = scoreSheet.getDataRange().getValues();

      for (var r = 4; r < data.length; r++) {
        var row = data[r];
        var sid = String(row[1]||'').trim();
        if (!sid || studentMap[sid] === undefined) continue;
        var key = subj.code || subj.name;

        if (termNum === 0) {
          // ทั้งปี: ภาค1 total, ภาค2 total, เฉลี่ยปี, เกรดปี
          var yg = row[layout.yearGradeCol];
          students[studentMap[sid]].scores[key] = {
            t1: Number(row[layout.term1.total]) || 0,
            t2: Number(row[layout.term2.total]) || 0,
            yearAvg: Number(row[layout.yearAvgCol]) || 0,
            yearGrade: (yg !== '' && yg != null) ? String(yg) : ''
          };
        } else {
          var tc = termNum === 2 ? layout.term2 : layout.term1;
          var tg = row[tc.grade];
          students[studentMap[sid]].scores[key] = {
            mid: Number(row[tc.midTotal]) || 0,
            final: Number(row[tc.s10]) || 0,
            total: Number(row[tc.total]) || 0,
            grade: (tg !== '' && tg != null) ? String(tg) : ''
          };
        }
      }
    });

    var settings = S_getGlobalSettings();
    return {
      success: true, grade: grade, classNo: classNo, term: termNum,
      mode: termNum === 0 ? 'year' : 'term',
      gradeFullName: U_getGradeFullName(grade),
      subjects: subjects, students: students, settings: settings
    };
  } catch (e) {
    Logger.log('❌ getClassSubjectScoreSummary error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/* ════════════════════════════════════════════════════════════
   2. สรุปการประเมินทางการเรียน
   ════════════════════════════════════════════════════════════ */
function getClassAssessmentSummary(grade, classNo) {
  try {
    var ss = SS();
    var subjects = _cr_getSubjects(ss, grade);

    var creditMap = {};
    subjects.forEach(function(s) { creditMap[s.code]=s.hours; creditMap[s.name]=s.hours; });

    var stInfo = _cr_getStudents(ss, grade, classNo);
    var classStudents = stInfo.students;
    if (classStudents.length === 0) return { success: false, message: 'ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo };

    // ดึง allGradeStudentIds (ทุกห้อง) เพื่อจัดลำดับชั้น
    var allGradeIds = [];
    for (var ai = 1; ai < stInfo.allData.length; ai++) {
      if (String(stInfo.allData[ai][stInfo.gradeCol]||'').trim() === grade) {
        allGradeIds.push(String(stInfo.allData[ai][stInfo.idCol]||'').trim());
      }
    }

    // SCORES_WAREHOUSE
    var warehouseSheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!warehouseSheet) throw new Error('ไม่พบชีต SCORES_WAREHOUSE');
    var whData = warehouseSheet.getDataRange().getValues();
    var wh = whData[0];
    var wI = U_findColumnIndex(wh,['student_id','studentId','รหัสนักเรียน']);
    var wG = U_findColumnIndex(wh,['grade','ชั้น']);
    var wS = U_findColumnIndex(wh,['subject_name','วิชา','ชื่อวิชา']);
    var wC = U_findColumnIndex(wh,['subject_code','รหัสวิชา']);
    var wT = U_findColumnIndex(wh,['subject_type','ประเภทวิชา']);
    var w1 = U_findColumnIndex(wh,['term1_total','term1Total']);
    var w2 = U_findColumnIndex(wh,['term2_total','term2Total']);
    var wA = U_findColumnIndex(wh,['average','avg']);
    var wF = U_findColumnIndex(wh,['final_grade','grade_result']);

    var scoresByStudent = {};
    var actKw = ['แนะแนว','ลูกเสือ','เนตรนารี','ชุมนุม','กิจกรรมเพื่อสังคม','บำเพ็ญประโยชน์'];
    for (var wi = 1; wi < whData.length; wi++) {
      var wr = whData[wi];
      if (String(wr[wG]||'').trim() !== grade) continue;
      var wId = String(wr[wI]||'').trim();
      var wSubj = String(wr[wS]||'').trim();
      var wTyp = wT!==-1 ? String(wr[wT]||'').trim() : '';
      if (wTyp==='กิจกรรม') continue;
      if (actKw.some(function(k){return wSubj.indexOf(k)!==-1;})) continue;

      var t1 = w1!==-1?(parseFloat(wr[w1])||0):0;
      var t2 = w2!==-1?(parseFloat(wr[w2])||0):0;
      var avg = wA!==-1?(parseFloat(wr[wA])||0):0;
      var fg = wF!==-1?parseFloat(wr[wF]):NaN;
      // ใช้เฉลี่ยปี (avg, เต็ม 100) แทน t1+t2 เพื่อให้ร้อยละถูกต้อง
      var totalScore = avg > 0 ? avg : (t1+t2 > 0 ? (t1+t2)/2 : 0);
      var wCode = wC!==-1?String(wr[wC]||'').trim():'';
      var gpa = !isNaN(fg) ? fg : _scoreToGPA(avg).gpa;
      var credits = creditMap[wCode] || creditMap[wSubj] || 1;
      if (!scoresByStudent[wId]) scoresByStudent[wId] = [];
      scoresByStudent[wId].push({ totalScore:totalScore, gpa:gpa, credits:credits });
    }

    function calcSummary(sid) {
      var sc = scoresByStudent[sid] || [];
      var sum=0, cr=0, gp=0;
      sc.forEach(function(s){sum+=s.totalScore;cr+=s.credits;gp+=s.gpa*s.credits;});
      var gpa = cr>0?gp/cr:0;
      var full = sc.length * 100;
      return { sumTotal:sum, fullScore:full, pct:full>0?(sum/full)*100:0, gpa:gpa };
    }

    // ลำดับชั้น
    var gradeGPAs = allGradeIds.map(function(id){return{id:id,gpa:calcSummary(id).gpa};});
    gradeGPAs.sort(function(a,b){return b.gpa-a.gpa;});
    // ลำดับห้อง
    var classGPAs = classStudents.map(function(st){return{id:st.studentId,gpa:calcSummary(st.studentId).gpa};});
    classGPAs.sort(function(a,b){return b.gpa-a.gpa;});

    // ดึงผลประเมิน 3 ด้าน
    var readingData=[], charData=[], compData=[];
    try { readingData = _readSheetToObjects('การประเมินอ่านคิดเขียน'); } catch(e){ Logger.log('⚠️ readingData error: ' + e.message); }
    try { charData = _readSheetToObjects('การประเมินคุณลักษณะ'); } catch(e){ Logger.log('⚠️ charData error: ' + e.message); }
    try { compData = _readSheetToObjects('การประเมินสมรรถนะ'); } catch(e){ Logger.log('⚠️ compData error: ' + e.message); }
    Logger.log('📊 Assessment data counts — reading: ' + readingData.length + ', char: ' + charData.length + ', comp: ' + compData.length);
    if (compData.length > 0) {
      var sampleRec = compData[0];
      var sampleKeys = Object.keys(sampleRec).filter(function(k){ return k.indexOf('สื่อสาร_')===0 || k.indexOf('คิด_')===0 || k.indexOf('แก้ปัญหา_')===0 || k.indexOf('ทักษะชีวิต_')===0 || k.indexOf('เทคโนโลยี_')===0; });
      Logger.log('📊 compData sample keys: ' + JSON.stringify(Object.keys(sampleRec)));
      Logger.log('📊 compData score keys found: ' + sampleKeys.length + ' → ' + JSON.stringify(sampleKeys.slice(0,5)));
      Logger.log('📊 compData sample studentId: ' + String(sampleRec['รหัสนักเรียน']||'(none)'));
    }

    function findRec(arr, sid) {
      return arr.find(function(r){return String(r['รหัสนักเรียน']||'').trim()===sid;});
    }
    function getReading(sid) {
      var rec = findRec(readingData, sid);
      if (!rec) return '-';
      var rKey = Object.keys(rec).find(function(k){return k.indexOf('สรุปผลการประเมิน')!==-1||k.indexOf('ผลการประเมิน')!==-1;});
      if (rKey && String(rec[rKey]).trim()) return String(rec[rKey]).trim();
      return '-';
    }
    function getChar(sid) {
      var rec = findRec(charData, sid);
      if (!rec) return '-';
      return String(rec['ผลการประเมิน']||'').trim() || '-';
    }
    function getComp(sid) {
      var rec = findRec(compData, sid);
      if (!rec) return '-';
      // Try prefixed columns first (new format: สื่อสาร_*, คิด_*, etc.)
      var prefixedKeys = Object.keys(rec).filter(function(k) {
        return k.indexOf('\u0e2a\u0e37\u0e48\u0e2d\u0e2a\u0e32\u0e23_')===0 || k.indexOf('\u0e04\u0e34\u0e14_')===0 ||
               k.indexOf('\u0e41\u0e01\u0e49\u0e1b\u0e31\u0e0d\u0e2b\u0e32_')===0 || k.indexOf('\u0e17\u0e31\u0e01\u0e29\u0e30\u0e0a\u0e35\u0e27\u0e34\u0e15_')===0 ||
               k.indexOf('\u0e40\u0e17\u0e04\u0e42\u0e19\u0e42\u0e25\u0e22\u0e35_')===0;
      });
      var nums = prefixedKeys.map(function(k){return parseFloat(rec[k]);}).filter(function(n){return !isNaN(n) && n > 0;});
      // Fallback: try ALL columns except metadata (old format: unprefixed columns)
      if (nums.length === 0) {
        var skipKeys = {'\u0e23\u0e2b\u0e31\u0e2a\u0e19\u0e31\u0e01\u0e40\u0e23\u0e35\u0e22\u0e19':1,'\u0e0a\u0e37\u0e48\u0e2d-\u0e19\u0e32\u0e21\u0e2a\u0e01\u0e38\u0e25':1,'\u0e0a\u0e31\u0e49\u0e19':1,'\u0e2b\u0e49\u0e2d\u0e07':1,'\u0e27\u0e31\u0e19\u0e17\u0e35\u0e48\u0e1a\u0e31\u0e19\u0e17\u0e36\u0e01':1,'\u0e1c\u0e39\u0e49\u0e1a\u0e31\u0e19\u0e17\u0e36\u0e01':1};
        var prefixedSet = {};
        prefixedKeys.forEach(function(k){ prefixedSet[k]=1; });
        var fallbackKeys = Object.keys(rec).filter(function(k){ return k && !skipKeys[k] && !prefixedSet[k]; });
        nums = fallbackKeys.map(function(k){return parseFloat(rec[k]);}).filter(function(n){return !isNaN(n) && n > 0;});
      }
      if (nums.length === 0) return '-';
      var avg = nums.reduce(function(a,b){return a+b;},0) / nums.length;
      if (avg >= 2.5) return '\u0e14\u0e35\u0e40\u0e22\u0e35\u0e48\u0e22\u0e21';
      if (avg >= 2.0) return '\u0e14\u0e35';
      if (avg >= 1.0) return '\u0e1c\u0e48\u0e32\u0e19';
      return '\u0e1b\u0e23\u0e31\u0e1a\u0e1b\u0e23\u0e38\u0e07';
    }

    var results = classStudents.map(function(st) {
      var sid = st.studentId;
      var s = calcSummary(sid);
      var cRank = classGPAs.findIndex(function(g){return g.id===sid;})+1;
      var gRank = gradeGPAs.findIndex(function(g){return g.id===sid;})+1;
      return {
        studentNo: st.studentNo, fullName: st.fullName,
        sumTotal: Math.round(s.sumTotal*100)/100, fullScore: s.fullScore,
        pct: Math.round(s.pct*100)/100, gpa: Math.round(s.gpa*100)/100,
        classRank: cRank, gradeRank: gRank,
        reading: getReading(sid), character: getChar(sid), competency: getComp(sid)
      };
    });

    return {
      success: true, grade: grade, classNo: classNo,
      gradeFullName: U_getGradeFullName(grade),
      totalClassStudents: classStudents.length,
      totalGradeStudents: allGradeIds.length,
      students: results, settings: S_getGlobalSettings()
    };
  } catch (e) {
    Logger.log('❌ getClassAssessmentSummary error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/* ════════════════════════════════════════════════════════════
   PDF — สรุปคะแนน (3 วิชา/หน้า + โลโก้)
   ════════════════════════════════════════════════════════════ */
function exportClassSubjectScorePDF(grade, classNo, term) {
  var data = getClassSubjectScoreSummary(grade, classNo, term);
  if (!data.success) throw new Error(data.message);

  var settings = data.settings || {};
  var schoolName = settings['ชื่อโรงเรียน'] || settings.school_name || '';
  var year = settings['ปีการศึกษา'] || settings.academic_year || '';
  var termNum = data.term;
  var subjects = data.subjects;
  var students = data.students;
  var isYear = data.mode === 'year';

  var homeroomTeacher = '', homeroomTeacher2 = '';
  try { var _ht = getHomeroomTeachers(grade, classNo); homeroomTeacher = _ht.teacher1; homeroomTeacher2 = _ht.teacher2; } catch(e) {}
  var logoBase64 = _cr_getLogoBase64(settings);

  var termLabel = isYear ? 'ทั้งปี' : ('ภาคเรียนที่ ' + termNum);

  // แบ่งวิชา 3 วิชา/หน้า
  var subjectGroups = [];
  for (var i = 0; i < subjects.length; i += 3) {
    subjectGroups.push(subjects.slice(i, i + 3));
  }

  var css = '@page{size:A4 landscape;margin:12mm 10mm;}'
    + '*{-webkit-print-color-adjust:exact;print-color-adjust:exact}'
    + 'body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;margin:0;}'
    + '.page{page-break-after:always;}'
    + '.page:last-child{page-break-after:auto;}'
    + '.header{text-align:center;margin-bottom:8px;}'
    + '.header img{height:50px;margin-bottom:4px;}'
    + '.title{font-size:14pt;font-weight:700;margin:2px 0;}'
    + '.info{font-size:10pt;margin:1px 0;}'
    + 'table{border-collapse:collapse;width:98%;font-size:10pt;}'
    + 'th,td{border:1px solid #000;padding:4px 6px;text-align:center;font-weight:300;}'
    + 'th{background:#f0f0f0;font-weight:700;}'
    + '.subject-header{background:#e0e0e0;}'
    + '.name-col{text-align:left;white-space:nowrap;}';

  var pages = subjectGroups.map(function(group, pageIdx) {
    var html = '<div class="page"><div class="header">';
    if (logoBase64) html += '<img src="' + logoBase64 + '" alt="logo"><br>';
    html += '<div class="title">สรุปคะแนน (' + termLabel + ')</div>';
    html += '<div class="info">ชั้น ' + data.gradeFullName + '/' + classNo
      + ' &nbsp;&nbsp; ปีการศึกษา ' + year
      + ' &nbsp;&nbsp; ครูประจำชั้น ' + homeroomTeacher + (homeroomTeacher2 ? ', ' + homeroomTeacher2 : '')
      + ' &nbsp;&nbsp; โรงเรียน' + schoolName + '</div>';
    html += '<div class="info" style="font-size:9pt;color:#555;">หน้า ' + (pageIdx+1) + '/' + subjectGroups.length + '</div>';
    html += '</div>';

    // Table header row 1
    html += '<table><thead><tr><th rowspan="2" style="width:35px">เลขที่</th>'
      + '<th rowspan="2" style="min-width:140px">ชื่อ-สกุล</th>';
    group.forEach(function(subj) {
      html += '<th colspan="4" class="subject-header">' + _cr_shortName(subj.name) + '</th>';
    });
    html += '</tr>';

    // Table header row 2
    html += '<tr>';
    group.forEach(function(subj) {
      if (isYear) {
        html += '<th>ภาค1<br>(100)</th><th>ภาค2<br>(100)</th><th>เฉลี่ยปี<br>(100)</th><th>เกรดปี</th>';
      } else {
        html += '<th>เก็บ' + termNum + '<br>(' + subj.midMax + ')</th>';
        html += '<th>สอบ' + termNum + '<br>(' + subj.finalMax + ')</th>';
        html += '<th>รวม' + termNum + '<br>(' + (subj.midMax+subj.finalMax) + ')</th>';
        html += '<th>เกรด' + termNum + '</th>';
      }
    });
    html += '</tr></thead><tbody>';

    // Student rows
    students.forEach(function(st, idx) {
      html += '<tr><td>' + (idx+1) + '</td><td class="name-col">' + st.fullName + '</td>';
      group.forEach(function(subj) {
        var key = subj.code || subj.name;
        var sc = st.scores[key];
        if (sc) {
          if (isYear) {
            html += '<td>' + sc.t1 + '</td><td>' + sc.t2 + '</td>';
            html += '<td><b>' + sc.yearAvg + '</b></td><td><b>' + sc.yearGrade + '</b></td>';
          } else {
            html += '<td>' + sc.mid + '</td><td>' + sc.final + '</td>';
            html += '<td><b>' + sc.total + '</b></td><td><b>' + sc.grade + '</b></td>';
          }
        } else {
          html += '<td>-</td><td>-</td><td>-</td><td>-</td>';
        }
      });
      html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
  });

  var fullHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + css + '</style></head><body>'
    + pages.join('') + '</body></html>';

  return _cr_htmlToBase64Pdf(fullHtml);
}

/* ════════════════════════════════════════════════════════════
   PDF — สรุปการประเมินทางการเรียน (+ โลโก้)
   ════════════════════════════════════════════════════════════ */
function exportClassAssessmentSummaryPDF(grade, classNo) {
  var data = getClassAssessmentSummary(grade, classNo);
  if (!data.success) throw new Error(data.message);

  var settings = data.settings || {};
  var schoolName = settings['ชื่อโรงเรียน'] || settings.school_name || '';
  var year = settings['ปีการศึกษา'] || settings.academic_year || '';

  var homeroomTeacher = '', homeroomTeacher2 = '';
  try { var _ht2 = getHomeroomTeachers(grade, classNo); homeroomTeacher = _ht2.teacher1; homeroomTeacher2 = _ht2.teacher2; } catch(e) {}
  var logoBase64 = _cr_getLogoBase64(settings);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<style>'
    + '@page{size:A4 landscape;margin:12mm 10mm;}'
    + '*{-webkit-print-color-adjust:exact;print-color-adjust:exact}'
    + 'body{font-family:"Sarabun",sans-serif;font-size:11pt;font-weight:300;margin:0;}'
    + '.header{text-align:center;margin-bottom:8px;}'
    + '.header img{height:50px;margin-bottom:4px;}'
    + '.title{font-size:14pt;font-weight:700;margin:2px 0;}'
    + '.info{font-size:10pt;margin:1px 0;}'
    + 'table{border-collapse:collapse;width:98%;font-size:10pt;}'
    + 'th,td{border:1px solid #000;padding:5px 7px;text-align:center;font-weight:300;}'
    + 'th{background:#f0f0f0;font-weight:700;}'
    + '.name-col{text-align:left;white-space:nowrap;}'
    + '.good{color:#059669;font-weight:600;}'
    + '.improve{color:#dc3545;font-weight:600;}'
    + '</style></head><body>';

  html += '<div class="header">';
  if (logoBase64) html += '<img src="' + logoBase64 + '" alt="logo"><br>';
  html += '<div class="title">สรุปการประเมินทางการเรียน</div>';
  html += '<div class="info">ชั้น ' + data.gradeFullName + '/' + classNo
    + ' &nbsp;&nbsp; ปีการศึกษา ' + year
    + ' &nbsp;&nbsp; ครูประจำชั้น ' + homeroomTeacher
    + ' &nbsp;&nbsp; โรงเรียน' + schoolName + '</div>';
  html += '</div>';

  html += '<table><thead>';
  html += '<tr>'
    + '<th rowspan="2" style="width:35px">เลขที่</th>'
    + '<th rowspan="2" style="min-width:160px">ชื่อ-สกุล</th>'
    + '<th colspan="6">สรุปการประเมิน</th>'
    + '<th rowspan="2">อ่านฯ</th>'
    + '<th rowspan="2">คุณลักษณะฯ</th>'
    + '<th rowspan="2">สมรรถนะฯ</th></tr>';
  html += '<tr><th>รวมคะแนน</th><th>เต็ม</th><th>ร้อยละ</th>'
    + '<th>เกรดเฉลี่ย<br>GPA</th><th>ลำดับ<br>ห้อง</th><th>ลำดับ<br>ชั้น</th></tr></thead>';

  html += '<tbody>';
  data.students.forEach(function(st, idx) {
    html += '<tr><td>' + (idx+1) + '</td><td class="name-col">' + st.fullName + '</td>';
    html += '<td>' + st.sumTotal + '</td><td>' + st.fullScore + '</td>';
    html += '<td>' + st.pct.toFixed(2) + '</td><td><b>' + st.gpa.toFixed(2) + '</b></td>';
    html += '<td>' + st.classRank + '</td><td>' + st.gradeRank + '</td>';
    function cls(v){return(v==='ดีเยี่ยม'||v==='ดี')?'good':(v==='ปรับปรุง'?'improve':'');}
    html += '<td class="'+cls(st.reading)+'">' + st.reading + '</td>';
    html += '<td class="'+cls(st.character)+'">' + st.character + '</td>';
    html += '<td class="'+cls(st.competency)+'">' + st.competency + '</td></tr>';
  });
  html += '</tbody></table></body></html>';

  return _cr_htmlToBase64Pdf(html);
}
