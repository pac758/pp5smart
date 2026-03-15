// ============================================================
// 📊 STUDENT SCORES SUMMARY - ฉบับสมบูรณ์
// แสดงเกรด + ส่งออก PDF + ใช้ student_id + คำนำหน้าชื่อ + โลโก้
// ============================================================

function scoreToGrade(score) {
  if (score >= 80) return 4.0;
  if (score >= 75) return 3.5;
  if (score >= 70) return 3.0;
  if (score >= 65) return 2.5;
  if (score >= 60) return 2.0;
  if (score >= 55) return 1.5;
  if (score >= 50) return 1.0;
  return 0;
}

function convertToGrade(value) {
  const numValue = parseFloat(value) || 0;
  if (numValue >= 0 && numValue <= 4) {
    return numValue;
  }
  if (numValue > 4) {
    return scoreToGrade(numValue);
  }
  return 0;
}

function getStudentScoresForWeb(grade, classNo) {
  try {
    const ss = SS();
    const studentsSheet = ss.getSheetByName("Students");
    if (!studentsSheet) {
      throw new Error("ไม่พบชีต 'Students'");
    }
    
    const studentData = studentsSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    
    const gradeCol = U_findColumnIndex(studentHeaders, ['grade', 'ชั้น', 'ระดับชั้น']);
    const classCol = U_findColumnIndex(studentHeaders, ['class_no', 'ห้อง']);
    const studentIdCol = U_findColumnIndex(studentHeaders, ['student_id', 'รหัสนักเรียน']);
    const studentNoCol = U_findColumnIndex(studentHeaders, ['student_no', 'เลขที่', 'ลำดับ']);
    const titleCol = U_findColumnIndex(studentHeaders, ['title', 'คำนำหน้า', 'คำนำหน้าชื่อ']);
    const firstNameCol = U_findColumnIndex(studentHeaders, ['firstname', 'first_name', 'ชื่อ']);
    const lastNameCol = U_findColumnIndex(studentHeaders, ['lastname', 'last_name', 'นามสกุล']);
    const fullNameCol = U_findColumnIndex(studentHeaders, ['full_name', 'ชื่อเต็ม', 'ชื่อ-นามสกุล']);
    
    if (studentIdCol === -1) {
      throw new Error("ไม่พบคอลัมน์ 'student_id' ในชีต Students");
    }
    
    const students = [];
    for (let i = 1; i < studentData.length; i++) {
      const row = studentData[i];
      const rowGrade = String(row[gradeCol] || '').trim();
      const rowClass = String(row[classCol] || '').trim();
      
      if (rowGrade === grade && rowClass === classNo) {
        const studentId = String(row[studentIdCol] || '').trim();
        const studentNo = studentNoCol !== -1 ? String(row[studentNoCol] || '').trim() : '';
        const title = titleCol !== -1 ? String(row[titleCol] || '').trim() : '';
        
        let fullName = '';
        if (fullNameCol !== -1 && row[fullNameCol]) {
          fullName = String(row[fullNameCol]).trim();
        } else if (firstNameCol !== -1 && lastNameCol !== -1) {
          const firstName = String(row[firstNameCol] || '').trim();
          const lastName = String(row[lastNameCol] || '').trim();
          fullName = `${firstName} ${lastName}`.trim();
        }
        
        // เพิ่มคำนำหน้าชื่อถ้ามี
        if (title && fullName) {
          fullName = `${title}${fullName}`;
        }
        
        students.push({
          studentId: studentId,
          studentNo: studentNo,
          fullName: fullName,
          grades: {},
          rawGrades: {}
        });
      }
    }
    
    if (students.length === 0) {
      return {
        success: false,
        message: `ไม่พบนักเรียนในชั้น ${grade} ห้อง ${classNo}`
      };
    }
    
    students.sort((a, b) => {
      if (a.studentNo && b.studentNo) {
        const numA = parseInt(a.studentNo) || 0;
        const numB = parseInt(b.studentNo) || 0;
        return numA - numB;
      }
      const numA = parseInt(a.studentId) || 0;
      const numB = parseInt(b.studentId) || 0;
      return numA - numB;
    });
    
    const scoresSheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!scoresSheet) {
      throw new Error("ไม่พบชีต SCORES_WAREHOUSE สำหรับปีปัจจุบัน");
    }
    
    const scoresData = scoresSheet.getDataRange().getValues();
    const scoresHeaders = scoresData[0];
    
    const scoreGradeCol = U_findColumnIndex(scoresHeaders, ['grade', 'ชั้น']);
    const scoreClassCol = U_findColumnIndex(scoresHeaders, ['class_no', 'ห้อง']);
    const scoreStudentIdCol = U_findColumnIndex(scoresHeaders, ['student_id', 'รหัสนักเรียน']);
    const subjectCol = U_findColumnIndex(scoresHeaders, ['subject_name', 'วิชา', 'ชื่อวิชา']);
    const subjectTypeCol = U_findColumnIndex(scoresHeaders, ['subject_type', 'ประเภทวิชา', 'ประเภท']);
    const scoreCol = U_findColumnIndex(scoresHeaders, ['final_grade', 'average', 'คะแนน', 'เกรด']);
    
    if (scoreStudentIdCol === -1) {
      throw new Error("ไม่พบคอลัมน์ 'student_id' ในชีต SCORES_WAREHOUSE");
    }
    
    const subjectsSet = new Set();
    const subjectTypes = {};

    const isActivitySubject = (subjectName) => {
      const n = String(subjectName || '').trim();
      if (!n) return false;
      const t = String(subjectTypes[n] || '').trim();
      if (t === 'กิจกรรม') return true;
      return (
        n.indexOf('แนะแนว') !== -1 ||
        n.indexOf('ลูกเสือ') !== -1 ||
        n.indexOf('เนตรนารี') !== -1 ||
        n.indexOf('ชุมนุม') !== -1 ||
        n.indexOf('กิจกรรมเพื่อสังคม') !== -1 ||
        n.indexOf('บำเพ็ญประโยชน์') !== -1
      );
    };
    
    for (let i = 1; i < scoresData.length; i++) {
      const row = scoresData[i];
      const rowGrade = String(row[scoreGradeCol] || '').trim();
      const rowClass = String(row[scoreClassCol] || '').trim();
      
      if (rowGrade === grade && rowClass === classNo) {
        const studentId = String(row[scoreStudentIdCol] || '').trim();
        const subject = String(row[subjectCol] || '').trim();
        const value = row[scoreCol];
        const st = subjectTypeCol !== -1 ? String(row[subjectTypeCol] || '').trim() : '';
        const gradeValue = convertToGrade(value);
        
        if (subject) {
          subjectsSet.add(subject);
          if (st && !subjectTypes[subject]) subjectTypes[subject] = st;
          const student = students.find(s => s.studentId === studentId);
          if (student) {
            student.grades[subject] = gradeValue;
            student.rawGrades[subject] = value;
          }
        }
      }
    }
    
    // ── ดึงผลประเมินจากชีต "การประเมินกิจกรรมพัฒนาผู้เรียน" ──
    try {
      let actSheet = null;
      try { actSheet = S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน'); } catch(e1) {}
      if (!actSheet) { try { actSheet = SS().getSheetByName('การประเมินกิจกรรมพัฒนาผู้เรียน'); } catch(e2) {} }

      if (actSheet && actSheet.getLastRow() > 1) {
        const actData = actSheet.getDataRange().getValues();
        const actHeaders = actData[0];
        // หา index ของคอลัมน์
        let actIdIdx = -1, actGradeIdx = -1, actClassIdx = -1;
        for (let h = 0; h < actHeaders.length; h++) {
          const hn = String(actHeaders[h] || '').trim();
          if (/รหัส/i.test(hn)) actIdIdx = h;
          if (/^ชั้น$/i.test(hn)) actGradeIdx = h;
          if (/^ห้อง$/i.test(hn)) actClassIdx = h;
        }

        if (actIdIdx >= 0) {
          // สร้าง map: studentId → { headerName: value }
          const actMap = {};
          for (let r = 1; r < actData.length; r++) {
            const row = actData[r];
            const rGrade = actGradeIdx >= 0 ? String(row[actGradeIdx] || '').trim() : '';
            const rClass = actClassIdx >= 0 ? String(row[actClassIdx] || '').trim() : '';
            if (rGrade && rGrade !== grade) continue;
            if (rClass && rClass !== classNo) continue;

            const sid = String(row[actIdIdx] || '').trim();
            if (!sid) continue;
            const m = {};
            for (let c = 0; c < actHeaders.length; c++) {
              const hn = String(actHeaders[c] || '').trim();
              const hv = String(row[c] || '').trim();
              if (hn && hv) m[hn] = hv;
            }
            actMap[sid] = m;
          }

          // จับคู่กิจกรรมกับวิชาใน subjects
          const matchRules = [
            { keys: ['แนะแนว'], col: 'กิจกรรมแนะแนว' },
            { keys: ['ลูกเสือ', 'เนตรนารี'], col: 'ลูกเสือ_เนตรนารี' },
            { keys: ['ชุมนุม'], col: 'ชุมนุม' },
            { keys: ['สังคม', 'สาธารณ'], col: 'เพื่อสังคมและสาธารณประโยชน์' }
          ];

          students.forEach(student => {
            const am = actMap[student.studentId];
            if (!am) return;
            subjectsSet.forEach(subject => {
              if (!isActivitySubject(subject)) return;
              const sn = String(subject).toLowerCase();
              for (const rule of matchRules) {
                if (rule.keys.some(k => sn.indexOf(k) >= 0)) {
                  const val = am[rule.col] || '';
                  if (val) {
                    student.rawGrades[subject] = val; // "ผ่าน" หรือ "ไม่ผ่าน"
                  }
                  break;
                }
              }
            });
          });
        }
      }
    } catch (actErr) {
      Logger.log('⚠️ getStudentScoresForWeb activity sheet: ' + actErr.message);
    }

    var subjects = Array.from(subjectsSet);
    const activityPriority = (name) => {
      const n = String(name || '').trim();
      if (!n) return 999;
      if (n.indexOf('แนะแนว') !== -1) return 0;
      if (n.indexOf('ลูกเสือ') !== -1 || n.indexOf('เนตรนารี') !== -1) return 1;
      if (n.indexOf('ชุมนุม') !== -1) return 2;
      if (n.indexOf('กิจกรรมเพื่อสังคม') !== -1) return 3;
      if (n.indexOf('บำเพ็ญประโยชน์') !== -1) return 4;
      return 99;
    };

    const academicOrder = [
      'ภาษาไทย',
      'คณิตศาสตร์',
      'วิทยาศาสตร์และเทคโนโลยี',
      'วิทยาศาสตร์',
      'สังคมศึกษา ศาสนา และวัฒนธรรม',
      'สังคมศึกษา',
      'ประวัติศาสตร์',
      'สุขศึกษาและพลศึกษา',
      'สุขศึกษา',
      'พลศึกษา',
      'ศิลปะ',
      'การงานอาชีพ',
      'การงาน',
      'ภาษาอังกฤษ',
      'หน้าที่พลเมือง',
      'การป้องกันตนเองและประเทศชาติ',
      'การป้องกัน'
    ];

    const academicPriority = (name) => {
      const n = String(name || '').trim();
      if (!n) return 999;
      for (let i = 0; i < academicOrder.length; i++) {
        if (n.indexOf(academicOrder[i]) !== -1) return i;
      }
      return 999;
    };

    subjects.sort((a, b) => {
      const aa = String(a || '').trim();
      const bb = String(b || '').trim();
      const aAct = isActivitySubject(aa);
      const bAct = isActivitySubject(bb);
      if (aAct !== bAct) return aAct ? 1 : -1;
      if (aAct && bAct) {
        const pa = activityPriority(aa);
        const pb = activityPriority(bb);
        if (pa !== pb) return pa - pb;
        return aa.localeCompare(bb, 'th-TH', { numeric: true });
      }
      const oa = academicPriority(aa);
      const ob = academicPriority(bb);
      if (oa !== ob) return oa - ob;
      return aa.localeCompare(bb, 'th-TH', { numeric: true });
    });
    
    students.forEach(student => {
      const grades = subjects
        .filter(sub => !isActivitySubject(sub))
        .map(sub => student.grades[sub])
        .filter(v => typeof v === 'number' && !isNaN(v));
      if (grades.length > 0) {
        const sum = grades.reduce((acc, val) => acc + val, 0);
        student.average = (sum / grades.length).toFixed(2);
      } else {
        student.average = '0.00';
      }
    });
    
   const settings = S_getGlobalSettings();
    
    return {
      success: true,
      grade: grade,
      classNo: classNo,
      gradeFullName: U_getGradeFullName(grade),
      students: students,
      subjects: subjects,
      subjectTypes: subjectTypes,
      settings: settings
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

function exportStudentScoresPDF(grade, classNo) {
  try {
    const data = getStudentScoresForWeb(grade, classNo);
    if (!data.success) {
      throw new Error(data.message);
    }
    return exportScoresPDFviaSheet_(data);
  } catch (error) {
    throw new Error(error.message);
  }
}

// ============================================================
// 📄 EXPORT PDF ผ่าน Google Sheets (แก้ปัญหาเส้น/ฟอนต์จาก HtmlService)
// ============================================================
function exportScoresPDFviaSheet_(data) {
  const { settings = {}, grade = '', classNo = '', gradeFullName = '', students = [], subjects = [], subjectTypes = {} } = data;
  const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียน';
  const academicYear = settings['ปีการศึกษา'] || '';
  // ฟอนต์ที่ Google Sheets รองรับจริง (ไม่ใช่ TH Sarabun New ซึ่งเป็น desktop font)
  const FONT = 'Sarabun';

  const isActivitySubject = (n) => {
    const s = String(n || '').trim();
    if (!s) return false;
    const t = String(subjectTypes[s] || '').trim();
    if (t === 'กิจกรรม') return true;
    return /แนะแนว|ลูกเสือ|เนตรนารี|ชุมนุม|กิจกรรมเพื่อสังคม|บำเพ็ญประโยชน์/.test(s);
  };

  const fmtGrade = (v) => {
    if (v == null) return '-';
    if (typeof v === 'number' && !isNaN(v)) {
      return Math.abs(v - Math.round(v)) < 1e-9 ? Math.round(v) : parseFloat(v.toFixed(1));
    }
    return String(v);
  };

  const fmtActivity = (raw, gv) => {
    const s = String(raw || '').trim();
    if (!s) {
      if (typeof gv === 'number' && !isNaN(gv)) return gv > 0 ? 'ผ' : '-';
      return '-';
    }
    if (/มผ|ไม่ผ่าน/.test(s)) return 'มผ';
    if (/ผ|ผ่าน/.test(s)) return 'ผ';
    const n = parseFloat(s);
    return !isNaN(n) ? (n >= 50 ? 'ผ' : 'มผ') : '-';
  };

  // ใช้ชื่อวิชาเต็ม (ไม่ย่อ) เพื่อความเป็นทางการ
  // คำนวณความสูง sub-header จากชื่อยาวสุด (ปรับให้กระชับขึ้น)
  const calcSubHeaderHeight = (allSubs) => {
    let maxLen = 0;
    allSubs.forEach(s => { if (s.length > maxLen) maxLen = s.length; });
    // ลดจาก maxLen * 8 + 10 เป็น maxLen * 5 + 20 เพื่อให้แถวไม่สูงเกิน
    return Math.max(80, Math.min(150, maxLen * 5 + 20));
  };

  const acadSubs = subjects.filter(s => !isActivitySubject(s));
  const actSubs = subjects.filter(s => isActivitySubject(s));
  const totalCols = 2 + acadSubs.length + actSubs.length + 1;

  const ss = SS();
  const sheet = ss.insertSheet('_TMP_PDF_' + Date.now());

  try {
    // ── ตั้งค่าคอลัมน์ ──
    const COL_NO = 1, COL_NAME = 2, COL_ACAD = 3;
    const COL_ACT = COL_ACAD + acadSubs.length;
    const COL_AVG = totalCols;

    sheet.setColumnWidth(COL_NO, 24);
    sheet.setColumnWidth(COL_NAME, 155);
    for (let i = 0; i < acadSubs.length; i++) sheet.setColumnWidth(COL_ACAD + i, 28);
    for (let i = 0; i < actSubs.length; i++) sheet.setColumnWidth(COL_ACT + i, 26);
    sheet.setColumnWidth(COL_AVG, 36);

    // ── คำนวณความกว้างรวม (pixels) เพื่อวางโลโก้ตรงกลาง ──
    let totalWidthPx = 24 + 155 + (acadSubs.length * 28) + (actSubs.length * 26) + 36;

    // ── แถว 1: โลโก้ตรงกลาง ──
    let row = 1;
    const logoUrl = settings['logoUrl_lh3'] || settings['logo'] || settings['schoolLogo'] || '';
    if (logoUrl) {
      try {
        const blob = UrlFetchApp.fetch(logoUrl).getBlob();
        sheet.setRowHeight(1, 70);
        sheet.getRange(1, 1, 1, totalCols).merge();
        // วางโลโก้ที่คอลัมน์ 1 แล้วเลื่อน offset ไปกึ่งกลาง
        const img = sheet.insertImage(blob, 1, 1);
        const logoW = 55, logoH = 55;
        img.setWidth(logoW).setHeight(logoH);
        const offsetX = Math.max(0, Math.floor((totalWidthPx - logoW) / 2));
        img.setAnchorCellXOffset(offsetX);
        img.setAnchorCellYOffset(5);
        row = 2;
      } catch(e) {
        Logger.log('⚠️ Logo error: ' + e.message);
      }
    }

    // ── แถว 2: หัวเรื่อง (ไม่มีภาคเรียน) ──
    sheet.getRange(row, 1, 1, totalCols).merge()
      .setValue('แบบสรุปผลการเรียนรู้ตามกลุ่มสาระการเรียนรู้ ปีการศึกษา ' + academicYear)
      .setFontFamily(FONT).setFontSize(16).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 28);
    row++;

    // ── แถว 3: ชื่อโรงเรียน ──
    sheet.getRange(row, 1, 1, totalCols).merge()
      .setValue(schoolName)
      .setFontFamily(FONT).setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(row, 24);
    row++;

    // ── แถว 4: ชั้น/ห้อง ──
    sheet.getRange(row, 1, 1, totalCols).merge()
      .setValue(gradeFullName + ' ห้อง ' + classNo)
      .setFontFamily(FONT).setFontSize(12)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(row, 22);
    row++;

    // ── แถว header กลุ่ม ──
    const hdrRow = row;
    const subRow = row + 1;

    // "ที่" merge 2 rows
    sheet.getRange(hdrRow, COL_NO, 2, 1).merge()
      .setValue('ที่').setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // "ชื่อ-นามสกุล" merge 2 rows
    sheet.getRange(hdrRow, COL_NAME, 2, 1).merge()
      .setValue('ชื่อ-นามสกุล').setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    // "รายวิชา"
    if (acadSubs.length > 0) {
      sheet.getRange(hdrRow, COL_ACAD, 1, acadSubs.length).merge()
        .setValue('รายวิชา').setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    // "กิจกรรมพัฒนาผู้เรียน"
    if (actSubs.length > 0) {
      sheet.getRange(hdrRow, COL_ACT, 1, actSubs.length).merge()
        .setValue('กิจกรรมพัฒนาผู้เรียน').setFontFamily(FONT).setFontSize(10).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
    // "เฉลี่ย" merge 2 rows
    sheet.getRange(hdrRow, COL_AVG, 2, 1).merge()
      .setValue('เฉลี่ย').setFontFamily(FONT).setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    // ── แถว sub-header: ชื่อวิชาเต็มหมุน 90° ──
    acadSubs.forEach((s, i) => {
      sheet.getRange(subRow, COL_ACAD + i)
        .setValue(s).setFontFamily(FONT).setFontSize(10)
        .setHorizontalAlignment('center').setVerticalAlignment('bottom')
        .setTextRotation(90);
    });
    actSubs.forEach((s, i) => {
      sheet.getRange(subRow, COL_ACT + i)
        .setValue(s).setFontFamily(FONT).setFontSize(9)
        .setHorizontalAlignment('center').setVerticalAlignment('bottom')
        .setTextRotation(90);
    });
    const subHdrHeight = calcSubHeaderHeight([...acadSubs, ...actSubs]);
    sheet.setRowHeight(subRow, subHdrHeight);

    // ── header สีพื้น + border ──
    sheet.getRange(hdrRow, 1, 2, totalCols)
      .setFontWeight('bold').setBackground('#f0f0f0');

    // ── ข้อมูลนักเรียน ──
    const dataRow = subRow + 1;
    const rows = students.map((st, i) => {
      const r = [st.studentNo || (i + 1), st.fullName || '-'];
      acadSubs.forEach(sub => r.push(fmtGrade(st.grades[sub])));
      actSubs.forEach(sub => r.push(fmtActivity((st.rawGrades||{})[sub], st.grades[sub])));
      r.push(parseFloat(st.average) || 0);
      return r;
    });

    if (rows.length > 0) {
      sheet.getRange(dataRow, 1, rows.length, totalCols).setValues(rows);
    }

    const lastRow = dataRow + rows.length - 1;

    // ── จัดรูปแบบข้อมูลทั้งหมด ──
    const dataRange = sheet.getRange(dataRow, 1, rows.length, totalCols);
    dataRange.setFontFamily(FONT).setFontSize(12)
      .setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // ชื่อ: ชิดซ้าย + WRAP
    sheet.getRange(dataRow, COL_NAME, rows.length, 1)
      .setHorizontalAlignment('left')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // เฉลี่ย: bold + 2 ทศนิยม
    sheet.getRange(dataRow, COL_AVG, rows.length, 1)
      .setFontWeight('bold').setNumberFormat('0.00');

    // ── ความสูงแถว ──
    for (let r = dataRow; r <= lastRow; r++) sheet.setRowHeight(r, 24);

    // ── เส้นขอบ (สีเทาเข้ม + เส้นบาง) ──
    sheet.getRange(hdrRow, 1, lastRow - hdrRow + 1, totalCols)
      .setBorder(true, true, true, true, true, true, '#555555', SpreadsheetApp.BorderStyle.SOLID);

    // ── ลบแถว/คอลัมน์เกิน ──
    if (sheet.getMaxColumns() > totalCols) sheet.deleteColumns(totalCols + 1, sheet.getMaxColumns() - totalCols);
    if (sheet.getMaxRows() > lastRow) sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);

    SpreadsheetApp.flush();

    // ── Export PDF ──
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
      'format=pdf&gid=' + sheet.getSheetId() +
      '&size=A4&portrait=true&fitw=true&gridlines=false' +
      '&printtitle=false&sheetnames=false&pagenum=UNDEFINED&fzr=false' +
      '&top_margin=0.35&bottom_margin=0.35&left_margin=0.4&right_margin=0.4' +
      '&horizontal_alignment=CENTER';

    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (HTTP ' + resp.getResponseCode() + ')');

    const pdf = DriveApp.createFile(resp.getBlob().setName(
      'สรุปผลการเรียน_' + grade + '_' + classNo + '_' +
      Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.pdf'
    ));
    pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdf.getUrl();

  } finally {
    try { ss.deleteSheet(sheet); } catch(e) {}
  }
}

function generateScoresSummaryHTML(data) {
  const { settings = {}, grade = '', classNo = '', gradeFullName = '', students = [], subjects = [], subjectTypes = {}, logoBase64 = null } = data;
  
  const logoUrl = logoBase64 || settings['logoUrl_lh3'] || settings['logo'] || settings['schoolLogo'] || '';
  const schoolName = settings['ชื่อโรงเรียน'] || 'โรงเรียน';
  const semester = settings['ภาคเรียน'] || '';
  const academicYear = settings['ปีการศึกษา'] || '';
  
  const isActivitySubject = (subjectName) => {
    const n = String(subjectName || '').trim();
    if (!n) return false;
    const t = String(subjectTypes[n] || '').trim();
    if (t === 'กิจกรรม') return true;
    return (
      n.indexOf('แนะแนว') !== -1 ||
      n.indexOf('ลูกเสือ') !== -1 ||
      n.indexOf('เนตรนารี') !== -1 ||
      n.indexOf('ชุมนุม') !== -1 ||
      n.indexOf('กิจกรรมเพื่อสังคม') !== -1 ||
      n.indexOf('บำเพ็ญประโยชน์') !== -1
    );
  };

  const formatGradeCell = (v) => {
    if (v === undefined || v === null) return '-';
    if (typeof v === 'number' && !isNaN(v)) {
      if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
      return v.toFixed(1).replace(/\.0$/, '');
    }
    return String(v);
  };

  const activityResultText = (raw, numericFallback) => {
    const s = String(raw || '').trim();
    if (!s) {
      if (typeof numericFallback === 'number' && !isNaN(numericFallback)) {
        return numericFallback > 0 ? 'ผ่าน' : '-';
      }
      return '-';
    }
    if (s.indexOf('มผ') !== -1 || s.indexOf('ไม่ผ่าน') !== -1) return 'มผ';
    if (s.indexOf('ผ') !== -1 || s.indexOf('ผ่าน') !== -1) return 'ผ';
    const n = parseFloat(s);
    if (!isNaN(n)) return n >= 50 ? 'ผ' : 'มผ';
    return '-';
  };

  // แยกวิชาสามัญ กับ กิจกรรม
  const academicSubjects = subjects.filter(s => !isActivitySubject(s));
  const activitySubjects = subjects.filter(s => isActivitySubject(s));

  // Short name mapping สำหรับ header แนวตั้ง
  const shortNameMap = {
    'ภาษาไทย': 'ภาษาไทย',
    'คณิตศาสตร์': 'คณิตศาสตร์',
    'วิทยาศาสตร์และเทคโนโลยี': 'วิทย์ฯ',
    'วิทยาศาสตร์': 'วิทย์ฯ',
    'สังคมศึกษา ศาสนา และวัฒนธรรม': 'สังคมฯ',
    'สังคมศึกษา': 'สังคมฯ',
    'ประวัติศาสตร์': 'ประวัติศาสตร์',
    'สุขศึกษาและพลศึกษา': 'สุขศึกษาฯ',
    'สุขศึกษา': 'สุขศึกษา',
    'พลศึกษา': 'พลศึกษา',
    'ศิลปะ': 'ศิลปะ',
    'การงานอาชีพ': 'การงานฯ',
    'การงาน': 'การงานฯ',
    'ภาษาอังกฤษ': 'ภาษาอังกฤษ',
    'หน้าที่พลเมือง': 'หน้าที่ฯ',
    'การป้องกันตนเองและประเทศชาติ': 'การป้องกันฯ'
  };
  const actShortNameMap = {
    'กิจกรรมแนะแนว': 'แนะแนว',
    'ลูกเสือ/เนตรนารี': 'ลูกเสือฯ',
    'ชุมนุม': 'ชุมนุม',
    'กิจกรรมเพื่อสังคมและสาธารณประโยชน์': 'สาธารณฯ'
  };
  const getShortName = (fullName) => {
    for (const key in shortNameMap) {
      if (fullName.indexOf(key) !== -1) return shortNameMap[key];
    }
    for (const key in actShortNameMap) {
      if (fullName.indexOf(key) !== -1) return actShortNameMap[key];
    }
    // กิจกรรม fallback
    if (fullName.indexOf('แนะแนว') !== -1) return 'แนะแนว';
    if (fullName.indexOf('ลูกเสือ') !== -1 || fullName.indexOf('เนตรนารี') !== -1) return 'ลูกเสือฯ';
    if (fullName.indexOf('ชุมนุม') !== -1) return 'ชุมนุม';
    if (fullName.indexOf('สังคม') !== -1 || fullName.indexOf('สาธารณ') !== -1) return 'สาธารณฯ';
    return fullName.length > 8 ? fullName.substring(0, 8) + 'ฯ' : fullName;
  };

  // สร้าง table rows
  const tableRows = students.map((student, index) => {
    const studentNo = student.studentNo || (index + 1);
    
    // คอลัมน์เกรดวิชาสามัญ
    const academicCells = academicSubjects.map(subject => {
      const v = student.grades[subject];
      return `<td class="gc">${formatGradeCell(v)}</td>`;
    }).join('');
    
    // คอลัมน์กิจกรรม
    const activityCells = activitySubjects.map(subject => {
      const raw = (student.rawGrades || {})[subject];
      const gradeValue = student.grades[subject];
      const result = activityResultText(raw, gradeValue);
      const cls = result === 'ผ' ? 'act-pass' : (result === 'มผ' ? 'act-fail' : '');
      return `<td class="ac ${cls}">${result}</td>`;
    }).join('');
    
    return `<tr><td class="no">${studentNo}</td><td class="nm">${student.fullName || '-'}</td>${academicCells}${activityCells}<td class="avg">${student.average || '0.00'}</td></tr>`;
  }).join('\n      ');

  // คำนวณความกว้างคอลัมน์
  const pageMarginMm = 7;
  const pageWidthMm = 210;
  const availableMm = pageWidthMm - (pageMarginMm * 2);
  const noMm = 7;
  const avgMm = 10;
  const acSubCount = academicSubjects.length;
  const atSubCount = activitySubjects.length;
  const totalSubCount = acSubCount + atSubCount;
  
  // คำนวณความกว้างชื่อ ขึ้นอยู่กับจำนวนวิชา
  const actMm = 7;
  const nameMm = Math.max(35, availableMm - noMm - avgMm - (actMm * atSubCount) - (totalSubCount > 12 ? 5 * acSubCount : 6 * acSubCount));
  const acadMm = acSubCount > 0 ? Math.max(4.5, (availableMm - noMm - nameMm - avgMm - (actMm * atSubCount)) / acSubCount) : 6;

  const academicColHeaders = academicSubjects.map(s => `<th class="vt">${getShortName(s)}</th>`).join('');
  const activityColHeaders = activitySubjects.map(s => `<th class="vt vt-act">${getShortName(s)}</th>`).join('');

  const colGroup = `<colgroup>
      <col style="width:${noMm}mm">
      <col style="width:${nameMm.toFixed(1)}mm">
      ${academicSubjects.map(() => `<col style="width:${acadMm.toFixed(1)}mm">`).join('')}
      ${activitySubjects.map(() => `<col style="width:${actMm}mm">`).join('')}
      <col style="width:${avgMm}mm">
    </colgroup>`;

  const logoSection = logoUrl ? `<img src="${logoUrl}" class="logo">` : '';

  // คำนวณ font-size ตาม row count
  const rowCount = students.length;
  let bodyFontPt = '14px';
  let cellPadding = '2px 1px';
  let headerHeight = '120px';
  if (rowCount > 35) {
    bodyFontPt = '11px';
    cellPadding = '1px 0';
    headerHeight = '90px';
  } else if (rowCount > 25) {
    bodyFontPt = '12px';
    cellPadding = '1px 1px';
    headerHeight = '100px';
  } else if (rowCount > 15) {
    bodyFontPt = '13px';
    cellPadding = '2px 1px';
    headerHeight = '110px';
  }

  return `<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<title>สรุปผลการเรียน ${grade}/${classNo}</title>
<style>
  @page { size: A4 portrait; margin: ${pageMarginMm}mm; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Sarabun','TH Sarabun New','Garuda',sans-serif; font-size: ${bodyFontPt}; line-height: 1.2; color: #000; }
  
  .hdr { text-align: center; margin-bottom: 4px; }
  .logo { width: 50px; height: 50px; object-fit: contain; display: block; margin: 0 auto 2px; }
  .hdr h1 { font-size: 16px; font-weight: bold; margin: 2px 0; }
  .hdr p { font-size: 14px; margin: 1px 0; }
  .hdr .sub { font-size: 13px; }
  
  table { width: 100%; border-collapse: collapse; table-layout: fixed; margin-top: 3px; }
  th, td { border: 0.5px solid #000; padding: ${cellPadding}; text-align: center; vertical-align: middle; }
  thead th { background: #fff; font-weight: bold; font-size: 11px; }
  
  .vt { writing-mode: vertical-rl; text-orientation: mixed; height: ${headerHeight}; font-size: 11px; font-weight: normal; white-space: nowrap; vertical-align: bottom; padding: 3px 1px; line-height: 1.0; }
  .vt-act { font-size: 10px; }
  .subj-hdr { font-size: 12px; font-weight: bold; border-bottom: 0.5px solid #000; }
  
  .no { text-align: center; font-size: ${bodyFontPt}; }
  .nm { text-align: left; padding-left: 4px; font-size: ${bodyFontPt}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .gc { font-size: ${bodyFontPt}; font-weight: bold; }
  .ac { font-size: 10px; white-space: nowrap; }
  .act-pass { color: #000; }
  .act-fail { color: #c00; font-weight: bold; }
  .avg { font-weight: bold; font-size: ${bodyFontPt}; }
  
  tbody tr:nth-child(even) { background: #fafafa; }
  tbody tr { page-break-inside: avoid; }
</style>
</head>
<body>
  <div class="hdr">
    ${logoSection}
    <h1>สรุปผลการเรียน ชั้นประจำการเรียน ภาคเรียนที่ ${semester} ปีการศึกษา ${academicYear}</h1>
    <p>${schoolName}</p>
    <p class="sub">${gradeFullName} ห้อง ${classNo}</p>
  </div>
  <table>
    ${colGroup}
    <thead>
      <tr>
        <th rowspan="2" style="width:${noMm}mm">ที่</th>
        <th rowspan="2">ชื่อ-นามสกุล</th>
        ${acSubCount > 0 ? `<th colspan="${acSubCount}" class="subj-hdr">รายวิชา</th>` : ''}
        ${atSubCount > 0 ? `<th colspan="${atSubCount}" class="subj-hdr">กิจกรรมพัฒนาผู้เรียน</th>` : ''}
        <th rowspan="2" style="width:${avgMm}mm">เฉลี่ย</th>
      </tr>
      <tr>${academicColHeaders}${activityColHeaders}</tr>
    </thead>
    <tbody>
      ${tableRows}
    </tbody>
  </table>
</body>
</html>`;
}

/**
 * ✅ ดึงรายการห้องเรียน - Fixed Version
 * @param {string} grade - ระดับชั้นที่เลือก เช่น "ป.1"
 * @returns {Array<string>} เลขห้องเรียน เช่น ["1", "2", "3"]
 */
function getAvailableClasses(grade) {  // ⚠️ เพิ่ม parameter
  try {
    Logger.log(`🔍 getAvailableClasses: grade="${grade}"`);

    if (!grade || !String(grade).trim()) {
      return [];
    }

    const ss = (getSpreadsheetId_())
      ? SS()
      : SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = ss.getSheetByName("Students") || ss.getSheetByName("นักเรียน");
    
    if (!studentsSheet) {
      Logger.log("❌ ไม่พบชีต Students/นักเรียน");
      return [];  // ⚠️ เปลี่ยนเป็น [] แทน Object
    }
    
    const data = studentsSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    const headers = data[0];

    const findCol = (names) => {
      if (typeof U_findColumnIndex === 'function') {
        return U_findColumnIndex(headers, names);
      }
      const normalized = names.map(n => String(n || '').toLowerCase().trim());
      for (let i = 0; i < headers.length; i++) {
        const h = String(headers[i] || '').toLowerCase().trim();
        if (normalized.indexOf(h) !== -1) return i;
      }
      return -1;
    };

    const gradeCol = findCol(['grade', 'ชั้น', 'ระดับชั้น']);
    const classCol = findCol(['class_no', 'ห้อง', 'เลขห้อง']);
    
    Logger.log(`📋 gradeCol=${gradeCol}, classCol=${classCol}`);
    
    if (gradeCol === -1 || classCol === -1) {
      Logger.log("❌ ไม่พบคอลัมน์ที่ต้องการ");
      return [];
    }
    
    const classes = new Set();
    const targetGrade = grade.trim();  // ⚠️ เตรียม grade สำหรับเปรียบเทียบ
    
    for (let i = 1; i < data.length; i++) {
      const rowGrade = String(data[i][gradeCol] || '').trim();
      const rowClass = String(data[i][classCol] || '').trim();
      
      // ⚠️ เปลี่ยนจาก: if (grade && classNo)
      // ⚠️ เป็น: if (rowGrade === targetGrade && rowClass)
      if (rowGrade === targetGrade && rowClass) {
        classes.add(rowClass);  // ⚠️ เก็บเฉพาะเลขห้อง ไม่ใช่ `${grade}/${classNo}`
        
        // แสดง log ครั้งแรก
        if (classes.size === 1) {
          Logger.log(`✅ ตัวอย่าง: grade="${rowGrade}", class="${rowClass}"`);
        }
      }
    }
    
    if (classes.size === 0) {
      Logger.log(`❌ ไม่พบห้องเรียนสำหรับ grade="${targetGrade}"`);
      return [];
    }
    
    // ⚠️ เรียงลำดับตัวเลข
    const result = Array.from(classes).sort((a, b) => {
      const numA = parseInt(a) || 0;
      const numB = parseInt(b) || 0;
      return numA - numB;
    });
    
    Logger.log(`✅ ผลลัพธ์: ${JSON.stringify(result)}`);
    
    // ⚠️ Return แบบ Array ธรรมดา ไม่มี Object wrapper
    return result;
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.message}`);
    return [];  // ⚠️ เปลี่ยนเป็น [] แทน Object
  }
}

/**
 * ✅ ดึงคะแนนนักเรียนสำหรับ Parent Portal (ใช้ใน parent_portal.html)
 * @param {string} studentId - รหัสนักเรียน
 * @returns {{ success: boolean, scores: Array, message?: string }}
 */
function getParentStudentScores(studentId) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };

    var ss = SS();

    // ดึงข้อมูลจาก SCORES_WAREHOUSE
    var warehouseSheet = (typeof S_getYearlySheet === 'function')
      ? S_getYearlySheet('SCORES_WAREHOUSE')
      : ss.getSheetByName('SCORES_WAREHOUSE');

    if (!warehouseSheet) return { success: true, scores: [], message: 'ไม่พบชีต SCORES_WAREHOUSE' };

    var data = warehouseSheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, scores: [] };

    var headers = data[0];
    var col = {};
    ['student_id', 'subject_code', 'subject_name', 'subject_type', 'hours',
     'term1_total', 'term2_total', 'average', 'final_grade'].forEach(function(name) {
      col[name] = headers.indexOf(name);
    });

    var sid = String(studentId).trim();
    var scores = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[col['student_id']] || '').trim() !== sid) continue;

      var subjectType = String(row[col['subject_type']] || '').trim();
      var isActivity = /กิจกรรม/i.test(subjectType);
      var finalGrade = String(row[col['final_grade']] || '').trim();
      var term1 = row[col['term1_total']];
      var term2 = row[col['term2_total']];
      var avg = row[col['average']];

      var scoreObj = {
        subjectCode: String(row[col['subject_code']] || '').trim(),
        name: String(row[col['subject_name']] || '').trim(),
        isActivity: isActivity,
        term1Score: (term1 !== '' && term1 !== null && term1 !== undefined) ? term1 : '-',
        term2Score: (term2 !== '' && term2 !== null && term2 !== undefined) ? term2 : '-',
        average: (avg !== '' && avg !== null && avg !== undefined) ? avg : '-',
        grade: finalGrade || '-'
      };

      if (isActivity) {
        scoreObj.activityGrade = finalGrade || 'ยังไม่ประเมิน';
      }

      scores.push(scoreObj);
    }

    // ── ดึงผลประเมินจากชีต "การประเมินกิจกรรมพัฒนาผู้เรียน" มาใส่ในวิชากิจกรรม ──
    try {
      var actSheet = (typeof S_getYearlySheet === 'function')
        ? S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน')
        : ss.getSheetByName('การประเมินกิจกรรมพัฒนาผู้เรียน');
      if (actSheet) {
        var actData = actSheet.getDataRange().getValues();
        var actHeaders = actData[0];
        var actIdIdx = -1;
        for (var ah = 0; ah < actHeaders.length; ah++) {
          if (/รหัส/i.test(String(actHeaders[ah] || ''))) { actIdIdx = ah; break; }
        }

        if (actIdIdx >= 0) {
          // หาแถวของนักเรียน
          var studentActRow = null;
          for (var ai = 1; ai < actData.length; ai++) {
            if (String(actData[ai][actIdIdx] || '').trim() === sid) {
              studentActRow = actData[ai];
              break;
            }
          }

          if (studentActRow) {
            // สร้าง map ของผลประเมิน: { keyword → ผ่าน/ไม่ผ่าน }
            var actMap = {};
            var actSummary = '';
            for (var ac = 0; ac < actHeaders.length; ac++) {
              var hn = String(actHeaders[ac] || '').trim();
              var hv = String(studentActRow[ac] || '').trim();
              if (!hn || !hv) continue;
              if (hn === 'รวมกิจกรรม') { actSummary = hv; continue; }
              // เก็บ keyword → value (เช่น "แนะแนว" → "ผ่าน")
              actMap[hn] = hv;
            }

            // จับคู่ keyword กับวิชากิจกรรมใน scores
            var matchKeywords = [
              { keys: ['แนะแนว'], col: 'กิจกรรมแนะแนว' },
              { keys: ['ลูกเสือ', 'เนตรนารี'], col: 'ลูกเสือ_เนตรนารี' },
              { keys: ['ชุมนุม'], col: 'ชุมนุม' },
              { keys: ['สังคม', 'สาธารณ'], col: 'เพื่อสังคมและสาธารณประโยชน์' }
            ];

            for (var si = 0; si < scores.length; si++) {
              if (!scores[si].isActivity) continue;
              var sName = (scores[si].name || '').toLowerCase();
              for (var mk = 0; mk < matchKeywords.length; mk++) {
                var matched = matchKeywords[mk].keys.some(function(k) { return sName.indexOf(k) >= 0; });
                if (matched && actMap[matchKeywords[mk].col]) {
                  scores[si].grade = actMap[matchKeywords[mk].col];
                  scores[si].activityGrade = actMap[matchKeywords[mk].col];
                  break;
                }
              }
            }

            // เพิ่มสรุปรวมกิจกรรม
            if (actSummary) {
              scores.push({
                subjectCode: '',
                name: 'รวมกิจกรรมพัฒนาผู้เรียน',
                isActivity: true,
                term1Score: '-',
                term2Score: '-',
                average: '-',
                grade: actSummary,
                activityGrade: actSummary,
                isSummary: true
              });
            }
          }
        }
      }
    } catch (actErr) {
      Logger.log('⚠️ getParentStudentScores activity sheet: ' + actErr.message);
    }

    // เรียงลำดับ: วิชาปกติก่อน, กิจกรรมทีหลัง, เรียงตามรหัสวิชา
    scores.sort(function(a, b) {
      if (a.isActivity !== b.isActivity) return a.isActivity ? 1 : -1;
      if (a.isSummary !== b.isSummary) return a.isSummary ? 1 : -1;
      return (a.subjectCode || '').localeCompare(b.subjectCode || '');
    });

    return { success: true, scores: scores };
  } catch (e) {
    Logger.log('❌ getParentStudentScores Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ ดึงรายชื่อนักเรียนที่ผูกกับผู้ปกครอง (parent_portal)
 * อ่าน student_ids จาก Users sheet แล้วดึงข้อมูลจาก Students sheet
 */
function getParentStudents() {
  try {
    var session = getLoginSession();
    if (!session || !session.username) {
      return { success: false, message: 'กรุณาเข้าสู่ระบบก่อน' };
    }

    var ss = SS();

    // ดึง student_ids จาก Users sheet
    var userSheet = ss.getSheetByName('Users') || ss.getSheetByName('users');
    if (!userSheet) return { success: false, message: 'ไม่พบชีต Users' };

    var userData = userSheet.getDataRange().getValues();
    var uHeaders = userData[0];
    var uNameIdx = -1, sidIdx = -1;
    for (var h = 0; h < uHeaders.length; h++) {
      var hLow = String(uHeaders[h] || '').toLowerCase().trim();
      if (['username', 'user', 'userid'].indexOf(hLow) >= 0) uNameIdx = h;
      if (['student_ids', 'studentids', 'linked_students'].indexOf(hLow) >= 0) sidIdx = h;
    }

    var linkedIds = [];
    if (uNameIdx >= 0 && sidIdx >= 0) {
      for (var u = 1; u < userData.length; u++) {
        if (String(userData[u][uNameIdx] || '').trim().toLowerCase() === session.username.toLowerCase()) {
          var raw = String(userData[u][sidIdx] || '').trim();
          if (raw) {
            linkedIds = raw.split(/[,;\s]+/).map(function(s) { return s.trim(); }).filter(Boolean);
          }
          break;
        }
      }
    }

    if (linkedIds.length === 0) {
      return { success: true, data: [] };
    }

    // ดึงข้อมูลนักเรียนจาก Students sheet
    var studSheet = ss.getSheetByName('Students');
    if (!studSheet) return { success: false, message: 'ไม่พบชีต Students' };

    var studData = studSheet.getDataRange().getValues();
    var sHeaders = studData[0];
    var col = {};
    ['student_id', 'title', 'firstname', 'lastname', 'grade', 'class_no', 'gender'].forEach(function(name) {
      col[name] = sHeaders.indexOf(name);
    });

    var results = [];
    for (var i = 1; i < studData.length; i++) {
      var row = studData[i];
      var sid = String(row[col['student_id']] || '').trim();
      if (linkedIds.indexOf(sid) === -1) continue;
      results.push({
        studentId: sid,
        title: String(row[col['title']] || '').trim(),
        firstName: String(row[col['firstname']] || '').trim(),
        lastName: String(row[col['lastname']] || '').trim(),
        grade: String(row[col['grade']] || '').trim(),
        classNo: String(row[col['class_no']] || '').trim(),
        gender: String(row[col['gender']] || '').trim()
      });
    }

    return { success: true, data: results };
  } catch (e) {
    Logger.log('❌ getParentStudents Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ ดึงสรุปเวลาเรียนสำหรับ Parent Portal
 * @param {string} studentId
 */
function getParentStudentAttendance(studentId) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };

    // ดึงปีการศึกษาจาก settings
    var settings = (typeof S_getGlobalSettings === 'function') ? S_getGlobalSettings() : {};
    var academicYear = Number(settings['ปีการศึกษา'] || new Date().getFullYear() + 543);

    // หา grade, classNo ของนักเรียน
    var ss = SS();
    var studSheet = ss.getSheetByName('Students');
    if (!studSheet) return { success: false, message: 'ไม่พบชีต Students' };

    var studData = studSheet.getDataRange().getValues();
    var sHeaders = studData[0];
    var idIdx = sHeaders.indexOf('student_id');
    var gradeIdx = sHeaders.indexOf('grade');
    var classIdx = sHeaders.indexOf('class_no');
    var grade = '', classNo = '';

    for (var i = 1; i < studData.length; i++) {
      if (String(studData[i][idIdx] || '').trim() === String(studentId).trim()) {
        grade = String(studData[i][gradeIdx] || '').trim();
        classNo = String(studData[i][classIdx] || '').trim();
        break;
      }
    }

    if (!grade) return { success: true, attendance: { totalDays: 0, totalPresent: 0, totalLeave: 0, totalAbsent: 0, percentage: 0, monthly: [] } };

    // ใช้ getStudentAttendanceDetail ที่มีอยู่
    if (typeof getStudentAttendanceDetail === 'function') {
      var detail = getStudentAttendanceDetail(grade, classNo, studentId, academicYear);
      // แปลง format ให้ตรงกับ parent_portal.html
      var monthly = (detail.monthlyData || []).map(function(m) {
        return {
          month: m.monthName || '',
          present: m.present || 0,
          leave: m.leave || 0,
          absent: m.absent || 0,
          total: m.total || 0
        };
      });
      return {
        success: true,
        attendance: {
          totalDays: detail.summary.totalDays || 0,
          totalPresent: detail.summary.totalPresent || 0,
          totalLeave: detail.summary.totalLeave || 0,
          totalAbsent: detail.summary.totalAbsent || 0,
          percentage: parseFloat(detail.summary.percentage) || 0,
          monthly: monthly
        }
      };
    }

    return { success: true, attendance: { totalDays: 0, totalPresent: 0, totalLeave: 0, totalAbsent: 0, percentage: 0, monthly: [] } };
  } catch (e) {
    Logger.log('❌ getParentStudentAttendance Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ✅ สร้างใบรายงานผลการเรียน PDF สำหรับ Parent Portal
 * @param {string} studentId
 */
function generateParentReportCard(studentId, options) {
  try {
    if (!studentId) return { success: false, message: 'ไม่ได้ระบุรหัสนักเรียน' };
    options = options || {};
    var term = options.term || 'both';

    // ใช้ generateOnePageReportPdf ที่มีอยู่
    if (typeof generateOnePageReportPdf === 'function') {
      var result = generateOnePageReportPdf(studentId, term);
      if (result && result.url) {
        return { success: true, viewUrl: result.url };
      } else if (result && typeof result === 'string' && result.startsWith('http')) {
        return { success: true, viewUrl: result };
      } else {
        return { success: false, message: 'ไม่สามารถสร้างใบรายงานได้' };
      }
    }

    return { success: false, message: 'ฟังก์ชันสร้างใบรายงานยังไม่พร้อมใช้งาน' };
  } catch (e) {
    Logger.log('❌ generateParentReportCard Error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function generateStudentReportPdfUnified(studentId, options) {
  options = options || {};
  var term = options.term || 'both';
  try {
    var r = generateParentReportCard(studentId, { term: term });
    if (r && r.success && r.viewUrl) {
      return { mode: 'url', url: r.viewUrl };
    }
  } catch (e) {}

  var raw = null;
  if (typeof generatePp6PDFCompleteNoDrive === 'function') {
    raw = generatePp6PDFCompleteNoDrive(studentId, term);
  } else if (typeof generatePp6PDFComplete === 'function') {
    try {
      var res = generatePp6PDFComplete(studentId, term);
      var pdfUrl = typeof res === 'string' ? res : (res && (res.previewUrl || res.downloadUrl) ? (res.previewUrl || res.downloadUrl) : '');
      if (pdfUrl) return { mode: 'url', url: pdfUrl };
    } catch (e2) {}
  }

  if (raw && raw.base64) {
    return { mode: 'base64', fileName: raw.fileName, mimeType: raw.mimeType, base64: raw.base64 };
  }
  throw new Error('ไม่สามารถสร้างรายงานได้');
}
