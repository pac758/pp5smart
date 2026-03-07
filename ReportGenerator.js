// ============================================================
// 🖨️ REPORT GENERATOR (ปพ.6) - FINAL VERSION
// รวมระบบโหลดเร็ว + ดีไซน์ใหม่ + สรุปผลการประเมิน
// ============================================================

// --- Global Cache ---
let CACHED_LOGO = null;
let CACHED_SETTINGS = null;
let SETTINGS_CACHE_TIME = 0;
const CACHE_DURATION = 5 * 60 * 1000; // 5 นาที
let CACHED_FOLDER = null;
let CACHED_STUDENTS_DATA = null;
let CACHED_STUDENTS_TIME = 0;
let CACHED_WAREHOUSE_DATA = null;
let CACHED_WAREHOUSE_TIME = 0;
let CACHED_SUBJECT_DATA = null;
let CACHED_SUBJECT_TIME = 0;
let CACHED_COMMENT_DATA = null;
let CACHED_COMMENT_TIME = 0;

// ============================================================
// ⚡ ฟังก์ชันหลัก: โหลดข้อมูลนักเรียน (Optimized)
// ============================================================
function getAllStudentDataOptimized() {
  const ss = SS();
  const sheet = ss.getSheetByName('Students');
  
  if (!sheet) throw new Error('ไม่พบ sheet Students');
  
  const data = sheet.getDataRange().getValues();
  const result = {};
  
  for (let i = 1; i < data.length; i++) {
    const grade = String(data[i][5] || '').trim();
    const classNo = String(data[i][6] || '').trim();
    
    if (!grade || !classNo) continue;
    
    if (!result[grade]) result[grade] = {};
    if (!result[grade][classNo]) result[grade][classNo] = [];
    
    result[grade][classNo].push({
      id: String(data[i][0] || '').trim(),
      fullName: `${data[i][2] || ''}${data[i][3] || ''} ${data[i][4] || ''}`.trim()
    });
  }
  
  Object.keys(result).forEach(grade => {
    Object.keys(result[grade]).forEach(classNo => {
      result[grade][classNo].sort((a, b) => a.id.localeCompare(b.id));
    });
  });
  
  return result;
}

// ============================================================
// 🖨️ สร้าง PDF ผ่าน Google Docs (เวอร์ชันสมบูรณ์)
// ============================================================
// ============================================================
// 1. ฟังก์ชันช่วย: ดึงความเห็นครู (เพิ่มส่วนนี้เข้าไปก่อนฟังก์ชันหลัก)
// ============================================================
function getTeacherComment_(studentId) {
  try {
    const data = getCachedSheetData_('ความเห็นครู');
    if (!data || data.length < 2) return '-';
    
    // ค้นหาคอลัมน์จาก header (ไม่ hardcode index)
    const headers = data[0];
    let idCol = -1, commentCol = -1;
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c]).trim();
      if (h === 'รหัสนักเรียน' || h === 'student_id') idCol = c;
      if (h === 'ความเห็นครู' || h === 'comment') commentCol = c;
    }
    // fallback: ถ้าไม่เจอ header ให้ใช้ col 0 / col 4
    if (idCol === -1) idCol = 0;
    if (commentCol === -1) commentCol = 4;
    
    const sid = String(studentId).trim();
    for (let i = 1; i < data.length; i++) {
      const sheetId = String(data[i][idCol]).trim();
      if (sheetId === sid) {
        const val = data[i][commentCol];
        return (val != null && String(val).trim() !== '') ? String(val).trim() : '-';
      }
    }
    return '-';
  } catch (e) {
    Logger.log('Error fetching teacher comment: ' + e.message);
    return '-';
  }
}

// ============================================================
// 2. ฟังก์ชันหลัก: สร้าง PDF ผ่าน Google Docs (Fix: Layout Spacing)
// ============================================================
function generatePp6Pdf(studentId, showRank) {
  if (showRank === undefined || showRank === null) showRank = true;
  try {
    // 1. เตรียมข้อมูล (ใช้ Cache เพื่อความเร็ว)
    const student = getStudentInfo_(studentId);
    if (!student) throw new Error('ไม่พบข้อมูลนักเรียน');
    
    const settings = getCachedSettings_();
    const schoolData = {
      name: settings['ชื่อโรงเรียน'] || 'โรงเรียน',
      office: settings['ที่อยู่โรงเรียน'] || 'สำนักงานเขตพื้นที่การศึกษา',
      district: settings['เขต'] || '',
      director: settings['ชื่อผู้อำนวยการ'] || '',
      directorPosition: settings['ตำแหน่งผู้อำนวยการ'] || 'ผู้อำนวยการสถานศึกษา',
      year: settings['ปีการศึกษา'] || String(new Date().getFullYear() + 543),
      logoFileId: settings['logoFileId'] || ''
    };

    const assessments = getStudentAssessments(studentId);
    const scores = getStudentScores_(studentId, schoolData.year, assessments);
    const attendance = getStudentAttendance_(studentId, schoolData.year);
    const gpaInfo = calculateGPAAndRank(studentId, student.grade, student.classNo);
    const teacherComment = getTeacherComment_(studentId);
    
    // 2. สร้าง Google Doc ชั่วคราว
    const docName = `ปพ.6_${student.id}_${student.name}_${new Date().getTime()}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    body.clear();
    
    // ตั้งค่าหน้ากระดาษ (A4) — compact 1 page
    body.setMarginTop(24);
    body.setMarginBottom(20);
    body.setMarginLeft(85);
    body.setMarginRight(28);
    body.setPageWidth(595);
    body.setPageHeight(842);
    
    // === ส่วนหัว (Header) ===
    if (schoolData.logoFileId) {
      try {
        const logoFile = DriveApp.getFileById(schoolData.logoFileId);
        const logoBlob = logoFile.getBlob();
        const logoPara = body.appendParagraph('');
        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        const image = logoPara.appendInlineImage(logoBlob);
        image.setHeight(42);
        image.setWidth(42);
        logoPara.setSpacingAfter(0);
      } catch (logoErr) {
        Logger.log('⚠️ ไม่สามารถโหลดโลโก้: ' + logoErr.message);
      }
    }
    
    const styleHeader = (para) => {
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        para.editAsText().setBold(true).setFontSize(12).setFontFamily('Sarabun');
        para.setSpacingAfter(0).setSpacingBefore(0);
    };

    styleHeader(body.appendParagraph('แบบรายงานผลพัฒนาคุณภาพผู้เรียนรายบุคคล (ปพ.6)'));
    styleHeader(body.appendParagraph(`ปีการศึกษา ${schoolData.year}`));
    styleHeader(body.appendParagraph(schoolData.name));
    
    let line4Text = schoolData.office;
    if (schoolData.district) line4Text += ` เขต ${schoolData.district}`;
    const line4 = body.appendParagraph(line4Text);
    styleHeader(line4);
    line4.setSpacingAfter(3);
    
    // === ข้อมูลนักเรียน ===
    const studentInfo = body.appendParagraph(
      `รหัสประจำตัวนักเรียน : ${student.id}     ชื่อ-นามสกุล : ${student.name}     ชั้น ${student.grade} ห้อง ${student.classNo}`
    );
    studentInfo.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    studentInfo.editAsText().setFontSize(10).setFontFamily('Sarabun');
    studentInfo.setSpacingAfter(1);
    
    // === ตารางเกรด (Main Table) ===
    const basicSubjects = scores.filter(s => s.type === 'พื้นฐาน' || !s.type);
    const addSubjects = scores.filter(s => s.type === 'เพิ่มเติม' || s.type === 'กิจกรรม');
    const allSubjects = [...basicSubjects, ...addSubjects];
    
    const table = body.appendTable();
    table.setBorderWidth(1);
    table.setBorderColor('#000000');
    
    const headers = ['ลำดับ', 'รหัสวิชา', 'รายวิชา', 'ประเภท', 'จำนวนชั่วโมง', 'ระดับผลการเรียน', 'หมายเหตุ'];
    const widths = [25, 50, 150, 50, 55, 55, 50];
    
    const headerRow = table.appendTableRow();
    headers.forEach((header, i) => {
      const cell = headerRow.appendTableCell(header);
      cell.setWidth(widths[i]);
      cell.setBackgroundColor('#f0f0f0');
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);
      const para = cell.getChild(0).asParagraph();
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      para.setSpacingAfter(0).setSpacingBefore(0);
      para.editAsText().setBold(true).setFontSize(10).setFontFamily('Sarabun');
    });
    
    allSubjects.forEach((s, idx) => {
      const row = table.appendTableRow();
      const values = [
        String(idx + 1), s.code || '', s.name || '-', s.type || 'พื้นฐาน',
        String(s.hours != null && s.hours !== '' ? s.hours : '-'), 
        String(s.grade != null && s.grade !== '' ? s.grade : '-'), ''
      ];
      values.forEach((val, i) => {
        const cell = row.appendTableCell(val);
        cell.setWidth(widths[i]);
        cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);
        const para = cell.getChild(0).asParagraph();
        const align = i === 2 ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER;
        para.setAlignment(align);
        para.setSpacingAfter(0).setSpacingBefore(0);
        const text = para.editAsText();
        text.setFontSize(10).setFontFamily('Sarabun').setBold(false);
      });
    });
    
    // ⭐ ลดช่องว่างหลังตารางเกรด (เปลี่ยนจาก Paragraph ว่าง เป็นการกำหนด SpaceBefore ที่หัวข้อถัดไปแทน)
    // body.appendParagraph('').setSpacingAfter(8); <--- ลบทิ้ง
    
    // ===================================================
    // ⭐ สรุปผลการประเมิน: outer table (ไม่มีเส้น) > [ซ้าย: nested table มีเส้น | ขวา: ลายเซ็น]
    // ===================================================
    const outerTable = body.appendTable();
    outerTable.setBorderWidth(0);
    outerTable.setBorderColor('#FFFFFF');

    const outerRow = outerTable.appendTableRow();
    const leftCell = outerRow.appendTableCell('');
    leftCell.setWidth(290);
    leftCell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(2).setPaddingRight(4);
    const rightCell = outerRow.appendTableCell('');
    rightCell.setWidth(213);
    rightCell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(8).setPaddingRight(0);

    // --- ลบ paragraph ว่างเริ่มต้นของ cell ---
    if (leftCell.getNumChildren() > 0) leftCell.getChild(0).removeFromParent();
    if (rightCell.getNumChildren() > 0) rightCell.getChild(0).removeFromParent();

    // ========== ซ้าย: Nested table มีเส้นขอบ ==========
    const smFont = 'Sarabun';
    const smSize = 10;
    const innerTable = leftCell.appendTable();
    innerTable.setBorderWidth(1);
    innerTable.setBorderColor('#000000');

    const smStyle = (cell, bold, width, align, bgColor) => {
      const p = cell.getChild(0).asParagraph();
      p.setSpacingAfter(0).setSpacingBefore(0);
      p.editAsText().setFontSize(smSize).setFontFamily(smFont).setBold(bold || false);
      if (align) p.setAlignment(align);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(4);
      if (width) cell.setWidth(width);
      if (bgColor) cell.setBackgroundColor(bgColor);
    };

    const mergeRows = [];
    let innerRowIdx = 0;
    const addInnerRow = (label, value, headerBg) => {
      if (headerBg) mergeRows.push(innerRowIdx);
      innerRowIdx++;
      const r = innerTable.appendTableRow();
      const c1 = r.appendTableCell(label);
      smStyle(c1, !!headerBg, 210, headerBg ? DocumentApp.HorizontalAlignment.CENTER : null, headerBg || null);
      const c2 = r.appendTableCell(String(value));
      smStyle(c2, false, 70, DocumentApp.HorizontalAlignment.CENTER, headerBg || null);
    };

    addInnerRow('สรุปผลการประเมิน', '', '#D6EAF8');
    addInnerRow('จำนวนหน่วยกิต/น้ำหนักรายวิชาพื้นฐาน', gpaInfo.basicCredits.toFixed(2));
    addInnerRow('จำนวนหน่วยกิต/น้ำหนักรายวิชาเพิ่มเติม', gpaInfo.additionalCredits.toFixed(2));
    addInnerRow('รวมหน่วยกิต/น้ำหนักที่ได้', gpaInfo.totalCredits.toFixed(2));
    addInnerRow('ระดับผลการเรียนเฉลี่ย', gpaInfo.gpa.toFixed(2));
    if (showRank) {
      addInnerRow('อันดับที่ใน ' + gpaInfo.totalStudents + ' คน', gpaInfo.classRank);
    }
    addInnerRow('สรุปเวลาเรียนมาเรียน ' + attendance.totalPresent + '/' + attendance.totalDays + ' วัน', 'ร้อยละ ' + attendance.percentage.toFixed(2));
    addInnerRow('การประเมินคุณลักษณะ', '', '#D6EAF8');
    addInnerRow('คุณลักษณะอันพึงประสงค์ของสถานศึกษา', assessments.character?.result || '-');
    addInnerRow('การอ่าน คิดวิเคราะห์และเขียน', assessments.reading?.result || '-');
    addInnerRow('กิจกรรมพัฒนาผู้เรียน', assessments.activities?.result || '-');
    // ความเห็นครูประจำชั้น — merge เป็นแถวเดียว
    const commentText = 'ความเห็นครูประจำชั้น: ' + (teacherComment || '-');
    mergeRows.push(innerRowIdx);
    innerRowIdx++;
    const cmRow = innerTable.appendTableRow();
    const cmC1 = cmRow.appendTableCell(commentText);
    smStyle(cmC1, false, 210, null, null);
    const cmC2 = cmRow.appendTableCell('');
    smStyle(cmC2, false, 70, null, null);

    // ========== ขวา: ลายเซ็น (paragraph ไม่มีเส้น) ==========
    const addSign = (text, isRole, topSpace) => {
      const p = rightCell.appendParagraph(text);
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.editAsText().setFontSize(10).setFontFamily('Sarabun');
      p.setSpacingBefore(topSpace != null ? topSpace : (isRole ? 0 : 10)).setSpacingAfter(0);
    };

    // ดึงชื่อครูประจำชั้น
    const homeroomName = getHomeroomTeacher(student.grade, String(student.classNo));

    addSign('ลงชื่อ.........................................', false, 30);
    addSign('(' + (homeroomName && homeroomName !== '...' ? homeroomName : '                                    ') + ')', false, 0);
    addSign('ครูที่ปรึกษา/ครูประจำชั้น', true);
    addSign('ลงชื่อ.........................................', false);
    addSign('(' + (schoolData.director || '                                    ') + ')', false, 0);
    addSign(schoolData.directorPosition, true);
    addSign('ลงชื่อ.........................................', false);
    addSign('(                                          )', false, 0);
    addSign('ผู้ปกครอง', true);

    // 4. บันทึก → merge cells → แปลงเป็น PDF
    doc.saveAndClose();
    const docId = doc.getId();

    // ⭐ Merge header cells ด้วย Docs Advanced Service
    try {
      if (mergeRows.length > 0) {
        const docJson = Docs.Documents.get(docId);
        let innerTableIdx = null;
        for (const elem of docJson.body.content) {
          if (elem.table && !innerTableIdx) {
            const rows = elem.table.tableRows;
            if (rows && rows[0] && rows[0].tableCells) {
              for (const cellElem of (rows[0].tableCells[0].content || [])) {
                if (cellElem.table) {
                  innerTableIdx = cellElem.startIndex;
                  break;
                }
              }
            }
          }
        }
        if (innerTableIdx != null) {
          const requests = mergeRows.map(rowIdx => ({
            mergeTableCells: {
              tableRange: {
                tableCellLocation: {
                  tableStartLocation: { index: innerTableIdx },
                  rowIndex: rowIdx,
                  columnIndex: 0
                },
                rowSpan: 1,
                columnSpan: 2
              }
            }
          }));
          Docs.Documents.batchUpdate({ requests: requests }, docId);
        }
      }
    } catch (mergeErr) {
      Logger.log('⚠️ Merge cells failed: ' + mergeErr.message);
    }

    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs('application/pdf');
    
    // ใช้ cached folder
    const folder = getCachedFolder_(settings);
    
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(`ปพ.6_${student.id}_${student.name}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    docFile.setTrashed(true);
    
    Logger.log('✅ สร้าง PDF สำเร็จ: ' + pdfFile.getUrl());
    return pdfFile.getUrl();
    
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log(e.stack);
    throw new Error('สร้าง PDF ไม่สำเร็จ: ' + e.message);
  }
}
// ============================================================
// 🔧 Helper Functions
// ============================================================

function getCachedSettings_() {
  const now = Date.now();
  if (CACHED_SETTINGS && (now - SETTINGS_CACHE_TIME) < CACHE_DURATION) {
    return CACHED_SETTINGS;
  }
  CACHED_SETTINGS = getGlobalSettings();
  SETTINGS_CACHE_TIME = now;
  
  if (!CACHED_LOGO && CACHED_SETTINGS['logoFileId']) {
    try {
      CACHED_LOGO = getLogoDataUrl(CACHED_SETTINGS['logoFileId']);
      CACHED_SETTINGS['_cachedLogoUrl'] = CACHED_LOGO;
    } catch (e) {
      Logger.log('⚠️ ไม่สามารถโหลดโลโก้: ' + e.message);
    }
  }
  return CACHED_SETTINGS;
}

function getCachedFolder_(settings) {
  if (CACHED_FOLDER) return CACHED_FOLDER;
  
  const folderName = "StudentReports_PP6";
  if (settings['pdfSaveFolderId']) {
    try { CACHED_FOLDER = DriveApp.getFolderById(settings['pdfSaveFolderId']); return CACHED_FOLDER; } 
    catch (e) { /* fallback */ }
  }
  const folders = DriveApp.getFoldersByName(folderName);
  CACHED_FOLDER = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  return CACHED_FOLDER;
}

function getCachedSheetData_(sheetName) {
  const now = Date.now();
  switch (sheetName) {
    case 'Students':
      if (CACHED_STUDENTS_DATA && (now - CACHED_STUDENTS_TIME) < CACHE_DURATION) return CACHED_STUDENTS_DATA;
      CACHED_STUDENTS_DATA = SS().getSheetByName('Students').getDataRange().getValues();
      CACHED_STUDENTS_TIME = now;
      return CACHED_STUDENTS_DATA;
    case 'SCORES_WAREHOUSE':
      if (CACHED_WAREHOUSE_DATA && (now - CACHED_WAREHOUSE_TIME) < CACHE_DURATION) return CACHED_WAREHOUSE_DATA;
      var whSheet = S_getYearlySheet('SCORES_WAREHOUSE');
      CACHED_WAREHOUSE_DATA = whSheet ? whSheet.getDataRange().getValues() : [];
      CACHED_WAREHOUSE_TIME = now;
      return CACHED_WAREHOUSE_DATA;
    case 'ความเห็นครู':
      if (CACHED_COMMENT_DATA && (now - CACHED_COMMENT_TIME) < CACHE_DURATION) return CACHED_COMMENT_DATA;
      const sh = S_getYearlySheet('ความเห็นครู');
      CACHED_COMMENT_DATA = sh ? sh.getDataRange().getValues() : null;
      CACHED_COMMENT_TIME = now;
      return CACHED_COMMENT_DATA;
    default:
      return SS().getSheetByName(sheetName).getDataRange().getValues();
  }
}

function getStudentInfo_(id) {
  const data = getCachedSheetData_('Students');
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(id)) {
      return {
        id: data[i][0],
        name: `${data[i][2] || ''}${data[i][3]} ${data[i][4]}`.trim(),
        grade: data[i][5],
        classNo: data[i][6],
        no: data[i][1] || '-'
      };
    }
  }
  return null;
}

function getStudentScores_(id, year, assessmentsCache) {
  const data = getCachedSheetData_('SCORES_WAREHOUSE');
  if (!data || data.length === 0) return [];
  const headers = data[0];
  
  // Header-based column lookup (flexible)
  const col = (name) => headers.indexOf(name);
  const findCol = (...names) => { for (const n of names) { const c = col(n); if (c !== -1) return c; } return -1; };
  
  const idCol = findCol('student_id', 'studentId');
  const codeCol = findCol('subject_code', 'subjectCode');
  const nameCol = findCol('subject_name', 'subjectName');
  const gradeCol = findCol('grade');
  const finalGradeCol = findCol('final_grade', 'grade_result', 'gradeResult');
  const averageCol = findCol('average', 'avg');
  
  // สร้าง map ชั่วโมง + ประเภทวิชา จากชีต "รายวิชา" (cached)
  const hoursMap = {};
  const typeMap = {};
  try {
    const now = Date.now();
    if (!CACHED_SUBJECT_DATA || (now - CACHED_SUBJECT_TIME) >= CACHE_DURATION) {
      const subjectSheet = SS().getSheetByName('รายวิชา');
      CACHED_SUBJECT_DATA = subjectSheet ? subjectSheet.getDataRange().getValues() : null;
      CACHED_SUBJECT_TIME = now;
    }
    if (CACHED_SUBJECT_DATA) {
      const subjectData = CACHED_SUBJECT_DATA;
      const sHeaders = subjectData[0];
      const sCodeIdx = sHeaders.indexOf('รหัสวิชา');
      const sHoursIdx = sHeaders.indexOf('ชั่วโมง/ปี');
      const sTypeIdx = sHeaders.indexOf('ประเภทวิชา');
      for (let i = 1; i < subjectData.length; i++) {
        const code = String(subjectData[i][sCodeIdx] || '').trim();
        if (code) {
          hoursMap[code] = parseInt(subjectData[i][sHoursIdx]) || 0;
          typeMap[code] = subjectData[i][sTypeIdx] || 'พื้นฐาน';
        }
      }
    }
  } catch (e) { Logger.log('⚠️ อ่านชีตรายวิชาไม่ได้: ' + e.message); }

  const scores =[];
  let studentGrade = '';
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol] || '').trim() === String(id).trim()) {
      if (!studentGrade) studentGrade = data[i][gradeCol];
      
      const subjectCode = String(data[i][codeCol] || '').trim();
      
      // ดึง grade: ใช้ final_grade ก่อน, ถ้าไม่มีคำนวณจาก average
      let gradeValue = '';
      if (finalGradeCol !== -1 && data[i][finalGradeCol] !== '' && data[i][finalGradeCol] != null) {
        gradeValue = parseFloat(data[i][finalGradeCol]);
        if (!isNaN(gradeValue)) {
          gradeValue = gradeValue.toFixed(1);
        } else {
          gradeValue = String(data[i][finalGradeCol]);
        }
      } else if (averageCol !== -1) {
        const avg = parseFloat(data[i][averageCol]);
        if (!isNaN(avg)) {
          gradeValue = _scoreToGPA(avg).gpa.toFixed(1);
        }
      }
      
      scores.push({
        code: subjectCode,
        name: data[i][nameCol] || '',
        type: typeMap[subjectCode] || 'พื้นฐาน',
        hours: hoursMap[subjectCode] || 0,
        grade: gradeValue
      });
    }
  }
  
  // เพิ่มกิจกรรมพิเศษ
  try {
    const assessments = assessmentsCache || getStudentAssessments(id);
    
    if (CACHED_SUBJECT_DATA && studentGrade) {
      const subjectData = CACHED_SUBJECT_DATA;
      
      const headers = subjectData[0];
      const gradeColIdx = headers.indexOf('ชั้น');
      const nameColIdx = headers.indexOf('ชื่อวิชา');
      const typeColIdx = headers.indexOf('ประเภทวิชา');
      const hoursColIdx = headers.indexOf('ชั่วโมง/ปี');
      
      const existingActivityNames = scores
        .filter(s => s.type === 'กิจกรรม')
        .map(s => s.name.toLowerCase());
      
      for (let i = 1; i < subjectData.length; i++) {
        const row = subjectData[i];
        
        if (row[gradeColIdx] === studentGrade && row[typeColIdx] === 'กิจกรรม') {
          const subjectName = row[nameColIdx];
          const subjectNameLower = subjectName.toLowerCase();
          
          if (!existingActivityNames.some(n => n === subjectNameLower)) {
            
            let activityGrade = 'มผ';
            
            if (subjectNameLower.includes('แนะแนว')) {
              activityGrade = assessments.activities.แนะแนว || 'มผ';
            } else if (subjectNameLower.includes('ลูกเสือ') || subjectNameLower.includes('เนตรนารี')) {
              activityGrade = assessments.activities.ลูกเสือ || 'มผ';
            } else if (subjectNameLower.includes('ชุมนุม')) {
              activityGrade = assessments.activities.ชุมนุม || 'มผ';
            } else if (subjectNameLower.includes('สังคม') || subjectNameLower.includes('สาธารณ')) {
              activityGrade = assessments.activities.สาธารณะ || 'มผ';
            }
            
            scores.push({
              code: '',
              name: subjectName,
              type: 'กิจกรรม',
              hours: parseInt(row[hoursColIdx]) || 0,
              grade: activityGrade
            });
            
            Logger.log(`✅ เพิ่มกิจกรรม: ${subjectName} = ${activityGrade}`);
          }
        }
      }
    }
    
  } catch (e) {
    Logger.log(`⚠️ ไม่สามารถดึงข้อมูลกิจกรรม: ${e.message}`);
  }
  
  return scores;
}

function getStudentAttendance_(id, year) {
  try {
    const academicYear = parseInt(year) || (new Date().getFullYear() + 543);
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const baseYearCE = academicYear - 543;
    const months = [
      { month: 5, yearCE: baseYearCE }, { month: 6, yearCE: baseYearCE },
      { month: 7, yearCE: baseYearCE }, { month: 8, yearCE: baseYearCE },
      { month: 9, yearCE: baseYearCE }, { month: 10, yearCE: baseYearCE },
      { month: 11, yearCE: baseYearCE }, { month: 12, yearCE: baseYearCE },
      { month: 1, yearCE: baseYearCE + 1 }, { month: 2, yearCE: baseYearCE + 1 },
      { month: 3, yearCE: baseYearCE + 1 }, { month: 4, yearCE: baseYearCE + 1 }
    ];
    
    const ss = SS();
    const sid = String(id).trim();
    let totalPresent = 0, totalLeave = 0, totalAbsent = 0;
    
    for (const m of months) {
      const sheetName = monthNames[m.month - 1] + (m.yearCE + 543);
      
      // ใช้ cache จาก _attendanceSheetCache (ประกาศใน Attendance.js)
      let data;
      if (typeof _attendanceSheetCache !== 'undefined' && _attendanceSheetCache[sheetName] !== undefined) {
        data = _attendanceSheetCache[sheetName];
      } else {
        const sheet = ss.getSheetByName(sheetName);
        data = sheet ? sheet.getDataRange().getValues() : null;
        if (typeof _attendanceSheetCache !== 'undefined') _attendanceSheetCache[sheetName] = data;
      }
      
      if (!data || data.length < 2) continue;
      
      // หานักเรียน
      let studentRow = null;
      for (let i = 1; i < data.length; i++) {
        const rowId = String(data[i][1] || '').trim();
        const match = rowId.match(/^(\d+)/);
        if (match && match[1] === sid) { studentRow = data[i]; break; }
      }
      if (!studentRow) continue;
      
      // นับสถานะ
      const headers = data[0];
      for (let c = 3; c < headers.length; c++) {
        if (!String(headers[c]).trim().match(/^\d{1,2}/)) continue;
        const status = String(studentRow[c] || '').trim();
        if (status === '/' || status === 'ม' || status === '1' || status === '✓') totalPresent++;
        else if (status === 'ล' || status.toLowerCase() === 'l') totalLeave++;
        else if (status === 'ข' || status === '0') totalAbsent++;
      }
    }
    
    const totalDays = totalPresent + totalLeave + totalAbsent;
    const percentage = totalDays > 0 ? (totalPresent / totalDays) * 100 : 0;
    return { percentage: percentage, totalDays: totalDays, totalPresent: totalPresent };
  } catch (e) {
    Logger.log('⚠️ getStudentAttendance_ fallback: ' + e.message);
    return { percentage: 0, totalDays: 0, totalPresent: 0 };
  }
}

function clearReportCache() {
  CACHED_LOGO = null;
  CACHED_SETTINGS = null;
  SETTINGS_CACHE_TIME = 0;
  CACHED_FOLDER = null;
  CACHED_STUDENTS_DATA = null;
  CACHED_STUDENTS_TIME = 0;
  CACHED_WAREHOUSE_DATA = null;
  CACHED_WAREHOUSE_TIME = 0;
  CACHED_SUBJECT_DATA = null;
  CACHED_SUBJECT_TIME = 0;
  CACHED_COMMENT_DATA = null;
  CACHED_COMMENT_TIME = 0;
  Logger.log('✅ Cache cleared');
}

