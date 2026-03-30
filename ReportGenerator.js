// ============================================================

// 🖨️ REPORT GENERATOR (ปพ.6) - FINAL VERSION

// รวมระบบโหลดเร็ว + ดีไซน์ใหม่ + สรุปผลการประเมิน

// ============================================================



// --- Helper: แสดงผลการเรียน (กิจกรรม → ผ/มผ, วิชาปกติ → ตัวเลข) ---

const ACTIVITY_KEYWORDS_ = ['กิจกรรม','แนะแนว','ลูกเสือ','เนตรนารี','ชุมนุม','เพื่อสังคม','สาธารณประโยชน์'];

function isActivitySubject_(s) {

  return String(s.type || '').trim().includes('กิจกรรม') ||

    ACTIVITY_KEYWORDS_.some(kw => String(s.name || '').includes(kw));

}

function formatGradeForDisplay_(s) {

  const g = s.grade;

  if (isActivitySubject_(s)) {

    if (!g) return 'มผ';

    const t = String(g).trim();

    if (t === 'ผ่าน' || t === 'ผ') return 'ผ';

    if (t === 'ไม่ผ่าน' || t === 'มผ') return 'มผ';

    if (!isNaN(parseFloat(t))) return 'ผ';

    return t;

  }

  if (g == null || g === '') return '-';

  const num = parseFloat(g);

  if (!isNaN(num)) return num % 1 === 0 ? String(Math.round(num)) : String(num);

  return String(g);

}



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

function getAllStudentDataOptimized(year) {

  const ss = SS();

  const sheet = (year && typeof S_getSharedSheet === 'function') ? S_getSharedSheet('Students', year) : ss.getSheetByName('Students');

  

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

      result[grade][classNo].sort((a, b) => a.id.localeCompare(b.id, undefined, {numeric: true}));

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

function _ensureDrivePermissionForPdf_() {
  try {
    DriveApp.getRootFolder().getName();
    return true;
  } catch (e) {
    // Fallback: ลองผ่าน REST API
    try {
      var token = ScriptApp.getOAuthToken();
      var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?pageSize=1&fields=files(id)', {
        headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) return true;
    } catch (_) {}
    throw new Error(
      'ระบบยังไม่ได้รับอนุญาตให้เข้าถึง Google Drive สำหรับการสร้าง PDF\n' +
      'วิธีแก้:\n' +
      '1) เปิดโปรเจกต์ Apps Script ด้วยบัญชีเจ้าของ (Deploying account)\n' +
      '2) Run ฟังก์ชัน testDrivePermission() 1 ครั้งเพื่อ authorize\n' +
      '3) Deploy > Manage deployments > Edit > New version > Deploy\n' +
      'หมายเหตุ: ถ้า Deploy ตั้งค่า Execute as = User accessing ให้เปลี่ยนเป็น Me'
    );
  }
}

function generatePp6Pdf(studentId, showRank) {

  if (showRank === undefined || showRank === null) showRank = true;

  try {
    _ensureDrivePermissionForPdf_();

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
    
    // ตั้งค่าหน้ากระดาษ (A4)
    body.setMarginTop(14);
    body.setMarginBottom(10);
    body.setMarginLeft(50);
    body.setMarginRight(20);
    body.setPageWidth(595);
    body.setPageHeight(842);

    

    // === ส่วนหัว (Header) ===

    if (schoolData.logoFileId) {

      try {

        const logoBlob = _getFileBlobCompat_(schoolData.logoFileId);

        const logoPara = body.appendParagraph('');

        logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

        const image = logoPara.appendInlineImage(logoBlob);

        image.setHeight(34);

        image.setWidth(34);

        logoPara.setSpacingAfter(0);

      } catch (logoErr) {

        Logger.log('⚠️ ไม่สามารถโหลดโลโก้: ' + logoErr.message);

      }

    }

    

    const styleHeader = (para) => {
        para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        para.editAsText().setBold(true).setFontSize(16).setFontFamily('Sarabun');
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
    studentInfo.editAsText().setFontSize(14).setFontFamily('Sarabun');
    studentInfo.setSpacingAfter(0);
    studentInfo.setSpacingBefore(0);

    

    // === ตารางเกรด (Main Table) ===

    const basicSubjects = scores.filter(s => s.type === 'พื้นฐาน' || !s.type);

    const addSubjects = scores.filter(s => s.type === 'เพิ่มเติม' || s.type === 'กิจกรรม');

    const allSubjects = [...basicSubjects, ...addSubjects];

    

    const table = body.appendTable();

    table.setBorderWidth(1);

    table.setBorderColor('#000000');
    table.setAttributes({
      [DocumentApp.Attribute.MARGIN_TOP]: 0,
      [DocumentApp.Attribute.MARGIN_BOTTOM]: 0
    });

    

    const headers = ['ลำดับ', 'รหัสวิชา', 'รายวิชา', 'ประเภท', 'จำนวนชั่วโมง', 'ระดับผลการเรียน', 'หมายเหตุ'];

    const widths = [26, 55, 175, 52, 50, 58, 52];

    

    const headerRow = table.appendTableRow();

    headers.forEach((header, i) => {

      const cell = headerRow.appendTableCell(header);

      cell.setWidth(widths[i]);

      cell.setBackgroundColor('#f0f0f0');

      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);

      const para = cell.getChild(0).asParagraph();

      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      para.setSpacingAfter(0).setSpacingBefore(0);

      para.editAsText().setBold(true).setFontSize(12).setFontFamily('Sarabun');
    });

    allSubjects.forEach((s, idx) => {
      const row = table.appendTableRow();
      const values = [
        String(idx + 1), s.code || '', s.name || '-', s.type || 'พื้นฐาน',
        String(s.hours != null && s.hours !== '' ? s.hours : '-'), 
        formatGradeForDisplay_(s), ''
      ];
      values.forEach((val, i) => {
        const cell = row.appendTableCell(val);
        cell.setWidth(widths[i]);
        cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(2).setPaddingRight(2);
        const para = cell.getChild(0).asParagraph();
        const align = i === 2 ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER;
        para.setAlignment(align);
        para.setSpacingAfter(0).setSpacingBefore(0);
        const text = para.editAsText();
        text.setFontSize(12).setFontFamily('Sarabun').setBold(false);
      });
    });

    // 3. ส่วนสรุปผลและลายเซ็น (ดึงขึ้นชิดตารางด้านบนให้มากที่สุด)
    // ลบช่องว่างที่เกิดขึ้นโดยอัตโนมัติระหว่างตาราง
    const lastIdx = body.getNumChildren() - 1;
    const spacerPara = body.getChild(lastIdx).asParagraph();
    spacerPara.setSpacingAfter(0).setSpacingBefore(0).editAsText().setFontSize(1);
    
    const outerTable = body.appendTable();
    outerTable.setBorderWidth(0);
    outerTable.setBorderColor('#FFFFFF');
    outerTable.setAttributes({
      [DocumentApp.Attribute.MARGIN_TOP]: 0,
      [DocumentApp.Attribute.MARGIN_BOTTOM]: 0
    });

    const outerRow = outerTable.appendTableRow();
    const leftCell = outerRow.appendTableCell('');
    leftCell.setWidth(310);
    leftCell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(2).setPaddingRight(4);
    const rightCell = outerRow.appendTableCell('');
    rightCell.setWidth(200);
    rightCell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(8).setPaddingRight(0);

    // --- ลบ paragraph ว่างเริ่มต้นของ cell ---
    if (leftCell.getNumChildren() > 0) leftCell.getChild(0).removeFromParent();
    if (rightCell.getNumChildren() > 0) rightCell.getChild(0).removeFromParent();

    const smFont = 'Sarabun';
    const smSize = 11.5; // ปรับลดขนาดตัวอักษรลงเล็กน้อย (11.5pt) เพื่อช่วยดึงเนื้อหา
    const innerTable = leftCell.appendTable();
    innerTable.setBorderWidth(1);
    innerTable.setBorderColor('#000000');

    const smStyle = (cell, bold, width, align, bgColor) => {
      const p = cell.getChild(0).asParagraph();
      p.setSpacingAfter(0).setSpacingBefore(0);
      p.editAsText().setFontSize(smSize).setFontFamily(smFont).setBold(bold || false);
      if (align) p.setAlignment(align);
      cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(4).setPaddingRight(4); // ลด padding เพื่อประหยัดพื้นที่แนวตั้ง
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
    addInnerRow('สรุปเวลาเรียน ' + attendance.totalPresent + '/' + attendance.totalDays + ' วัน', 'ร้อยละ ' + attendance.percentage.toFixed(2));
    addInnerRow('การประเมินคุณลักษณะ', '', '#D6EAF8');
    addInnerRow('คุณลักษณะอันพึงประสงค์', assessments.character?.result || '-');
    addInnerRow('การอ่าน คิดวิเคราะห์และเขียน', assessments.reading?.result || '-');
    addInnerRow('กิจกรรมพัฒนาผู้เรียน', assessments.activities?.result || '-');

    // ความเห็นครูประจำชั้น — ปรับให้ประหยัดพื้นที่แนวตั้งที่สุด
    const commentText = 'ความเห็นครู: ' + (teacherComment || '-');
    mergeRows.push(innerRowIdx);
    innerRowIdx++;
    const cmRow = innerTable.appendTableRow();
    const cmC1 = cmRow.appendTableCell(commentText);
    smStyle(cmC1, false, 210, DocumentApp.HorizontalAlignment.LEFT, null);
    const cmC2 = cmRow.appendTableCell('');
    smStyle(cmC2, false, 70, DocumentApp.HorizontalAlignment.LEFT, null);
    cmC1.setPaddingTop(2).setPaddingBottom(2); 

    // ========== ขวา: ลายเซ็น (paragraph ไม่มีเส้น) ==========
    const addSign = (text, isRole, topSpace) => {
      const p = rightCell.appendParagraph(text);
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.editAsText().setFontSize(11.5).setFontFamily('Sarabun'); 
      p.setSpacingBefore(topSpace != null ? topSpace : (isRole ? 0 : 4)).setSpacingAfter(0); // ลดระยะห่างลายเซ็น
    };

    // ดึงชื่อครูประจำชั้นทั้ง 2 คน
    const homeroomTeachers = getHomeroomTeachers(student.grade, String(student.classNo));
    const homeroomName = homeroomTeachers.teacher1;
    const homeroomName2 = homeroomTeachers.teacher2;

    addSign('ลงชื่อ.........................................', false, 8); // ลดระยะห่างด้านบนลายเซ็น
    addSign('(' + (homeroomName && homeroomName !== '...' ? homeroomName : '                                    ') + ')', false, 0);
    addSign('ครูประจำชั้นคนที่ 1', true);
    addSign('ลงชื่อ.........................................', false, 10);
    addSign('(' + (homeroomName2 || '                                    ') + ')', false, 0);
    addSign('ครูประจำชั้นคนที่ 2', true);
    addSign('ลงชื่อ.........................................', false, 10);
    addSign('(' + (schoolData.director || '                                    ') + ')', false, 0);
    addSign(schoolData.directorPosition, true);



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



    // แปลง PDF (ใช้ pattern เดียวกับ OnePageReport เพื่อความเสถียร)
    let _pdfStep = 'getFileById';
    try {
      _pdfStep = 'getAs_pdf';
      const pdfBlob = _getFileBlobCompat_(docId).getAs('application/pdf');
      pdfBlob.setName(`ปพ.6_${student.id}_${student.name}.pdf`);

      _pdfStep = 'saveBlobGetUrl';
      const pdfUrl = _saveBlobGetUrl_(pdfBlob, null, settings.pdfSaveFolderId || null);

      _pdfStep = 'setTrashed';
      try { DriveApp.getFileById(docId).setTrashed(true); } catch(_) {
        try {
          var _tk = ScriptApp.getOAuthToken();
          UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId, {
            method: 'patch', contentType: 'application/json',
            headers: { Authorization: 'Bearer ' + _tk },
            payload: JSON.stringify({ trashed: true }), muteHttpExceptions: true
          });
        } catch(_2) {}
      }

      Logger.log('✅ สร้าง PDF สำเร็จ: ' + pdfUrl);
      return pdfUrl;
      
    } catch (pdfErr) {
      Logger.log(`❌ PDF Conversion Error at [${_pdfStep}]: ` + pdfErr.message);
      throw new Error(`[PDF step: ${_pdfStep}] ` + pdfErr.message);
    }

    

  } catch (e) {

    Logger.log('❌ Error: ' + e.message);

    Logger.log(e.stack);

    if (String(e && e.message ? e.message : e).indexOf('Drive') !== -1 || String(e && e.message ? e.message : e).indexOf('DriveApp') !== -1) {
      throw new Error(String(e.message || e));
    }
    throw new Error('สร้าง PDF ไม่สำเร็จ: ' + (e && e.message ? e.message : e));

  }

}

// ============================================================

// 🧪 ทดสอบสิทธิ์ DriveApp (Run ใน Editor เพื่อ authorize /dev)

// ============================================================

function testDrivePermission() {

  const files = DriveApp.getFiles();

  Logger.log('✅ DriveApp authorized! First file: ' + (files.hasNext() ? files.next().getName() : 'none'));

  const doc = DocumentApp.create('test_permission_' + Date.now());

  Logger.log('✅ DocumentApp authorized! Doc ID: ' + doc.getId());

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  Logger.log('✅ ทุกสิทธิ์พร้อมใช้งาน — /dev จะทำงานได้แล้ว');

  return 'OK';

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

  try {
    const folders = DriveApp.getFoldersByName(folderName);
    CACHED_FOLDER = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  } catch (e) {
    Logger.log('getCachedFolder_ DriveApp fallback: ' + e.message);
    CACHED_FOLDER = null;
  }

  return CACHED_FOLDER;

}



function getCachedSheetData_(sheetName, year) {

  const now = Date.now();
  var currentYear = (typeof S_getAcademicYear === 'function') ? String(S_getAcademicYear()) : '';
  var isOldYear = year && String(year) !== currentYear;

  switch (sheetName) {

    case 'Students':
      if (isOldYear) {
        var snapSheet = (typeof S_getSharedSheet === 'function') ? S_getSharedSheet('Students', year) : SS().getSheetByName('Students');
        return snapSheet ? snapSheet.getDataRange().getValues() : [];
      }

      if (CACHED_STUDENTS_DATA && (now - CACHED_STUDENTS_TIME) < CACHE_DURATION) return CACHED_STUDENTS_DATA;

      CACHED_STUDENTS_DATA = SS().getSheetByName('Students').getDataRange().getValues();

      CACHED_STUDENTS_TIME = now;

      return CACHED_STUDENTS_DATA;

    case 'SCORES_WAREHOUSE':

      if (isOldYear) {
        var whOld = S_getYearlySheet('SCORES_WAREHOUSE', year);
        return whOld ? whOld.getDataRange().getValues() : [];
      }

      if (CACHED_WAREHOUSE_DATA && (now - CACHED_WAREHOUSE_TIME) < CACHE_DURATION) return CACHED_WAREHOUSE_DATA;

      var whSheet = S_getYearlySheet('SCORES_WAREHOUSE');

      CACHED_WAREHOUSE_DATA = whSheet ? whSheet.getDataRange().getValues() : [];

      CACHED_WAREHOUSE_TIME = now;

      return CACHED_WAREHOUSE_DATA;

    case 'ความเห็นครู':

      if (isOldYear) {
        var cmOld = S_getYearlySheet('ความเห็นครู', year);
        return cmOld ? cmOld.getDataRange().getValues() : null;
      }

      if (CACHED_COMMENT_DATA && (now - CACHED_COMMENT_TIME) < CACHE_DURATION) return CACHED_COMMENT_DATA;

      const sh = S_getYearlySheet('ความเห็นครู');

      CACHED_COMMENT_DATA = sh ? sh.getDataRange().getValues() : null;

      CACHED_COMMENT_TIME = now;

      return CACHED_COMMENT_DATA;

    case 'รายวิชา':
      if (isOldYear) {
        var subjOld = (typeof S_getSharedSheet === 'function') ? S_getSharedSheet('รายวิชา', year) : SS().getSheetByName('รายวิชา');
        return subjOld ? subjOld.getDataRange().getValues() : null;
      }
      if (!CACHED_SUBJECT_DATA || (now - CACHED_SUBJECT_TIME) >= CACHE_DURATION) {
        var subjSheet = SS().getSheetByName('รายวิชา');
        CACHED_SUBJECT_DATA = subjSheet ? subjSheet.getDataRange().getValues() : null;
        CACHED_SUBJECT_TIME = now;
      }
      return CACHED_SUBJECT_DATA;

    default:

      return SS().getSheetByName(sheetName).getDataRange().getValues();

  }

}



function getStudentInfo_(id, year) {

  const data = getCachedSheetData_('Students', year);

  

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

  const data = getCachedSheetData_('SCORES_WAREHOUSE', year);

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

  

  // สร้าง map ชั่วโมง + ประเภทวิชา จากชีต "รายวิชา" (cached / year-aware)

  const hoursMap = {};

  const typeMap = {};

  try {

    const subjectDataForYear = getCachedSheetData_('รายวิชา', year);

    if (subjectDataForYear) {

      const subjectData = subjectDataForYear;

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

          gradeValue = gradeValue % 1 === 0 ? String(gradeValue) : gradeValue.toFixed(1);

        } else {

          gradeValue = String(data[i][finalGradeCol]);

        }

      } else if (averageCol !== -1) {

        const avg = parseFloat(data[i][averageCol]);

        if (!isNaN(avg)) {

          var g = _scoreToGPA(avg).gpa;

          gradeValue = g % 1 === 0 ? String(g) : g.toFixed(1);

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

  

  // ⭐ แก้ไข: อ่านผลประเมินกิจกรรมตรงจากชีต (ไม่ผ่าน cache) แล้วแทนที่เกรดตัวเลข

  try {

    // 1. อ่านชีตการประเมินกิจกรรมพัฒนาผู้เรียนโดยตรง

    let actSheet = null;

    try {

      actSheet = S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน', year);

    } catch (e1) {

      actSheet = SS().getSheetByName('การประเมินกิจกรรมพัฒนาผู้เรียน');

    }

    

    let activityMap = {}; // { แนะแนว: 'ผ่าน', ลูกเสือ: 'ผ่าน', ... }

    if (actSheet) {

      const actData = actSheet.getDataRange().getValues();

      const actHeaders = actData[0];

      Logger.log('📋 [Activity Sheet] headers: ' + actHeaders.join(', '));

      

      // หาคอลัมน์จาก header

      const actIdCol = actHeaders.findIndex(h => String(h).trim() === 'รหัสนักเรียน');

      const sid = String(id).trim();

      

      for (let r = 1; r < actData.length; r++) {

        if (String(actData[r][actIdCol] || '').trim() === sid) {

          // เจอแถวนักเรียน → สร้าง map จาก header ทั้งหมด

          actHeaders.forEach((h, ci) => {

            const hTrim = String(h).trim();

            if (hTrim && ci > 3) { // ข้ามคอลัมน์ รหัส, ชื่อ, ชั้น, ห้อง

              activityMap[hTrim] = String(actData[r][ci] || '').trim();

            }

          });

          Logger.log('✅ [Activity Sheet] found student ' + sid + ': ' + JSON.stringify(activityMap));

          break;

        }

      }

    } else {

      Logger.log('⚠️ ไม่พบชีต การประเมินกิจกรรมพัฒนาผู้เรียน');

    }

    

    // 2. แก้เกรดวิชากิจกรรมทุกตัวใน scores

    Logger.log('📋 [DEBUG] scores count=' + scores.length + ', activityMap keys=' + Object.keys(activityMap).join(','));

    scores.forEach(s => {

      try {

        if (isActivitySubject_(s)) {

          s.type = 'กิจกรรม';

          const n = String(s.name || '');

          let actGrade = 'ผ่าน';

          

          // จับคู่ชื่อวิชากับคอลัมน์ในชีตประเมิน

          if (n.includes('แนะแนว')) actGrade = activityMap['กิจกรรมแนะแนว'] || activityMap['แนะแนว'] || 'ผ่าน';

          else if (n.includes('ลูกเสือ') || n.includes('เนตรนารี')) actGrade = activityMap['ลูกเสือ_เนตรนารี'] || activityMap['ลูกเสือ/เนตรนารี'] || 'ผ่าน';

          else if (n.includes('ชุมนุม')) {
            actGrade = activityMap['ชุมนุม'] || 'ผ่าน';
            // ถ้ามีชื่อชุมนุมรายคน → ใช้แทนชื่อจากรายวิชา
            if (activityMap['ชื่อชุมนุม']) s.name = activityMap['ชื่อชุมนุม'];
          }

          else if (n.includes('สังคม') || n.includes('สาธารณ')) actGrade = activityMap['เพื่อสังคมและสาธารณประโยชน์'] || activityMap['เพื่อสังคม'] || 'ผ่าน';

          else actGrade = activityMap['รวมกิจกรรม'] || 'ผ่าน';

          

          // ถ้าได้ค่าเป็นตัวเลข (เช่น 0, 1) → แปลงเป็น ผ่าน/ไม่ผ่าน

          if (actGrade === '' || actGrade === '0' || actGrade === '0.0') actGrade = 'ไม่ผ่าน';

          else if (!isNaN(parseFloat(actGrade))) actGrade = 'ผ่าน';

          

          Logger.log('✅ Fix activity: "' + n + '" old=' + s.grade + ' → ' + actGrade);

          s.grade = actGrade;

        }

      } catch (rowErr) {

        Logger.log('⚠️ Activity fix row error: ' + rowErr.message);

      }

    });

    

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



