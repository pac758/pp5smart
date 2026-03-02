// ============================================================
// 📖 PP5 FULL BOOK — ปพ.5 รวมเล่ม (PDF)
// สร้าง HTML รวมทุกส่วนแล้วแปลงเป็น PDF เดียว
// ============================================================

/**
 * ฟังก์ชันหลัก: สร้าง PDF ปพ.5 รวมเล่ม
 * @param {string} grade ชั้น เช่น "ป.3"
 * @param {string} classNo ห้อง เช่น "1"
 * @returns {string} URL ของไฟล์ PDF ที่สร้าง
 */
function exportPp5FullBook(grade, classNo) {
  try {
    Logger.log('📖 เริ่มสร้าง ปพ.5 รวมเล่ม: ' + grade + '/' + classNo);
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    if (!grade || !classNo) throw new Error('กรุณาระบุชั้นและห้อง');

    // --- ดึงข้อมูลพื้นฐาน ---
    var settings = S_getGlobalSettings(false);
    var academicYear = settings['ปีการศึกษา'] || '';
    var schoolName = settings['ชื่อโรงเรียน'] || '';
    var logoDataUri = '';
    try { logoDataUri = _getLogoDataUrl(settings.logoFileId || ''); } catch (e) {}
    var teacherName = settings['ชื่อครูประจำชั้น'] || getClassTeacherName(grade, classNo) || '';
    var directorName = settings['ชื่อผู้อำนวยการ'] || '';
    var directorTitle = settings['ตำแหน่งผู้อำนวยการ'] || 'ผู้อำนวยการโรงเรียน';
    var academicHead = settings['ชื่อหัวหน้างานวิชาการ'] || '';

    var gradeFullName = _pp5GradeFullName_(grade);

    // --- ดึงรายชื่อนักเรียน + ผู้ปกครอง ---
    var students = getFilteredStudentsInline(grade, classNo);
    if (!students || students.length === 0) throw new Error('ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo);

    // --- ดึงคะแนนรายวิชาทั้งหมด ---
    var scoreSheets = getExistingScoreSheets();
    var cleanGrade = grade.replace(/[\/\.]/g, '');
    var classScoreSheets = scoreSheets.filter(function(s) {
      var sheetGrade = String(s.grade).trim().replace(/[\/\.]/g, '');
      return (sheetGrade === cleanGrade || sheetGrade === grade) && String(s.classNo).trim() === classNo;
    });
    Logger.log('📊 พบ ' + classScoreSheets.length + ' วิชาสำหรับ ' + grade + '/' + classNo + ' (cleanGrade=' + cleanGrade + ')');

    // --- ดึงเวลาเรียน ---
    var attendance1 = [], attendance2 = [];
    try {
      attendance1 = getSemesterAttendanceSummary(grade, classNo, 1, Number(academicYear));
    } catch (e) { Logger.log('⚠️ ดึงเวลาเรียนภาค 1 ไม่ได้: ' + e.message); }
    try {
      attendance2 = getSemesterAttendanceSummary(grade, classNo, 2, Number(academicYear));
    } catch (e) { Logger.log('⚠️ ดึงเวลาเรียนภาค 2 ไม่ได้: ' + e.message); }

    // --- ดึงข้อมูลการประเมิน ---
    var charData = [], actData = [], rtwSummary = {}, attrSummary = {}, actSummary = {};
    try { charData = getStudentsForCharacteristic(grade, classNo); } catch (e) { Logger.log('⚠️ ดึงคุณลักษณะไม่ได้: ' + e.message); }
    try { actData = getStudentsForActivity(grade, classNo); } catch (e) {}
    try { rtwSummary = getReadingSummary(grade, classNo); } catch (e) {}
    try { attrSummary = getAttributeSummary(grade, classNo); } catch (e) {}
    try { actSummary = getActivitySummary(grade, classNo); } catch (e) {}

    // --- ดึงข้อมูล อ่าน คิดวิเคราะห์ เขียน ---
    var subjectScoreStudents = [];
    try { subjectScoreStudents = getStudentsForSubjectScore(grade, classNo); } catch (e) { Logger.log('⚠️ ดึง RTW/SubjectScore ไม่ได้: ' + e.message); }

    // --- ดึงข้อมูลสำหรับ PDF ---
    var pdfCommon = { schoolName: schoolName, academicYear: academicYear, directorName: directorName, teacherName: teacherName, logoBase64: logoDataUri }

    // --- ดึงข้อมูลสรุปวิชา ---
    var subjectSummary = [];
    try { subjectSummary = getSubjectScoreSummary(grade, classNo); } catch (e) {}

    // --- รวม HTML ทั้งหมด ---
    var meta = {
      grade: grade, classNo: classNo, gradeFullName: gradeFullName,
      schoolName: schoolName, academicYear: academicYear,
      logoDataUri: logoDataUri, teacherName: teacherName,
      directorName: directorName, directorTitle: directorTitle,
      academicHead: academicHead, settings: settings
    };

    var htmlParts = [];

    // ส่วนที่ 1: ปก ปพ.5
    htmlParts.push(pp5book_coverHtml_(meta, students.length, subjectSummary, rtwSummary, attrSummary, actSummary));

    // ส่วนที่ 2: รายชื่อนักเรียน
    htmlParts.push(pp5book_studentListHtml_(meta, students));

    // ส่วนที่ 3: ข้อมูลผู้ปกครอง
    htmlParts.push(pp5book_parentInfoHtml_(meta, students));

    // ====== ส่วนที่ 4: บันทึกเวลาเรียน (ก่อนผลการเรียนรายวิชา) ======

    // 4A: สรุปเวลาเรียน (รวมภาค)
    htmlParts.push(pp5book_attendanceHtml_(meta, students, attendance1, attendance2));

    // 4B: เวลาเรียนรายเดือน (สรุป มา/ลา/ขาด) + รายวัน (ตาราง ✓/ข/ล)
    try {
      var yearCE = Number(academicYear) - 543;
      var acadMonths = [
        { m: 5, y: yearCE, name: 'พฤษภาคม', m0: 4 },
        { m: 6, y: yearCE, name: 'มิถุนายน', m0: 5 },
        { m: 7, y: yearCE, name: 'กรกฎาคม', m0: 6 },
        { m: 8, y: yearCE, name: 'สิงหาคม', m0: 7 },
        { m: 9, y: yearCE, name: 'กันยายน', m0: 8 },
        { m: 10, y: yearCE, name: 'ตุลาคม', m0: 9 },
        { m: 11, y: yearCE, name: 'พฤศจิกายน', m0: 10 },
        { m: 12, y: yearCE, name: 'ธันวาคม', m0: 11 },
        { m: 1, y: yearCE + 1, name: 'มกราคม', m0: 0 },
        { m: 2, y: yearCE + 1, name: 'กุมภาพันธ์', m0: 1 },
        { m: 3, y: yearCE + 1, name: 'มีนาคม', m0: 2 }
      ];

      // สรุปรายเดือน
      var monthlyDataAll = [];
      acadMonths.forEach(function(mo) {
        try {
          var mData = getMonthlyAttendanceSummary(grade, classNo, mo.y, mo.m);
          if (mData && mData.length > 0) {
            monthlyDataAll.push({ month: mo.m, year: mo.y, name: mo.name, yearBE: mo.y + 543, data: mData });
          }
        } catch (e) { Logger.log('⚠️ ข้ามเดือน ' + mo.name + ': ' + e.message); }
      });
      if (monthlyDataAll.length > 0) {
        htmlParts.push(pp5book_monthlyAttendanceHtml_(meta, students, monthlyDataAll));
      }

      // 4C: ตารางรายวัน (เหมือน PDF จากหน้าเช็คชื่อ — ✓/ข/ล แต่ละวัน)
      acadMonths.forEach(function(mo) {
        try {
          var tableData = getAttendanceVerticalTableFiltered(grade, classNo, mo.y, mo.m0);
          if (tableData && tableData.students && tableData.students.length > 0 && tableData.days && tableData.days.length > 0) {
            var semester = (mo.m0 >= 4 && mo.m0 <= 9) ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
            htmlParts.push(pp5book_dailyAttendanceHtml_(meta, tableData, mo.name, mo.y + 543, semester));
          }
        } catch (e) { Logger.log('⚠️ ข้ามตารางรายวัน ' + mo.name + ': ' + e.message); }
      });
    } catch (e) { Logger.log('⚠️ สร้างเวลาเรียนไม่ได้: ' + e.message); }

    // ====== ส่วนที่ 5: คะแนนรวม 2 ภาค แต่ละวิชา (เฉพาะ 11 วิชาหลัก ไม่รวมกิจกรรม) ======
    _sortBySubjectName(classScoreSheets, 'subjectName');

    var activityKeywords = ['กิจกรรม', 'แนะแนว', 'ลูกเสือ', 'เนตรนารี', 'ชุมนุม', 'เพื่อสังคม', 'ชมรม'];
    classScoreSheets.forEach(function(sheetInfo) {
      try {
        var subName = (sheetInfo.subjectName || '').trim();
        var isActivity = activityKeywords.some(function(kw) { return subName.indexOf(kw) !== -1; });
        if (isActivity) {
          Logger.log('⏭️ ข้ามวิชากิจกรรม: ' + subName);
          return;
        }
        var scoreData = getScoreSheetData(sheetInfo.sheetName);
        htmlParts.push(pp5book_yearScoreHtml_(meta, sheetInfo, scoreData));
      } catch (e) {
        Logger.log('⚠️ ข้ามวิชา ' + sheetInfo.subjectName + ': ' + e.message);
      }
    });

    // ส่วนที่ 6: คุณลักษณะอันพึงประสงค์ (ใช้ HTML เดิมจาก Assessments.js)
    try {
      if (charData && charData.length > 0) {
        var charHtml = createCharacteristicAssessmentHTML(charData, grade, classNo, academicYear, schoolName, directorName, teacherName, logoDataUri);
        htmlParts.push(pp5book_extractBody_(charHtml));
      }
    } catch (e) { Logger.log('⚠️ สร้าง HTML คุณลักษณะไม่ได้: ' + e.message); }

    // ส่วนที่ 7: กิจกรรมพัฒนาผู้เรียน (ใช้ HTML เดิมจาก Assessments.js)
    try {
      if (actData && actData.length > 0) {
        var actHtml = createActivityAssessmentHTML_(actData, grade, classNo, academicYear, schoolName, directorName, teacherName, logoDataUri);
        htmlParts.push(pp5book_extractBody_(actHtml));
      }
    } catch (e) { Logger.log('⚠️ สร้าง HTML กิจกรรมไม่ได้: ' + e.message); }

    // ส่วนที่ 8: อ่าน คิดวิเคราะห์ เขียน (ใช้ HTML เดิมจาก Assessments.js)
    try {
      if (subjectScoreStudents && subjectScoreStudents.length > 0) {
        var rtwHtml = createOfficialReportHTML(subjectScoreStudents, grade, classNo, academicYear, schoolName, directorName, teacherName, logoDataUri);
        htmlParts.push(pp5book_extractBody_(rtwHtml));
      }
    } catch (e) { Logger.log('⚠️ สร้าง HTML อ่านคิดเขียนไม่ได้: ' + e.message); }

    // --- รวมเป็น HTML เดียว ---
    var fullHtml = pp5book_wrapHtml_(htmlParts);

    // --- สร้าง PDF ---
    var fileName = 'ปพ.5 รวมเล่ม_' + grade + '_' + classNo + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    var blob = Utilities.newBlob(fullHtml, 'text/html', fileName + '.html').getAs('application/pdf');
    blob.setName(fileName + '.pdf');

    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('✅ สร้าง ปพ.5 รวมเล่มสำเร็จ: ' + file.getUrl());
    return file.getUrl();

  } catch (error) {
    Logger.log('❌ exportPp5FullBook error: ' + error.message);
    throw new Error(error.message);
  }
}


// ============================================================
// 🔧 ดึง <style> + <body> จาก full HTML document
// ============================================================
function pp5book_extractBody_(fullHtml) {
  if (!fullHtml) return '';
  // ดึง <style> content
  var styleMatch = fullHtml.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
  var styleContent = styleMatch ? styleMatch[1] : '';

  // ดึง <body> content
  var bodyMatch = fullHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  var bodyContent = bodyMatch ? bodyMatch[1] : '';

  if (!bodyContent) return '';

  // สร้าง scoped section ด้วย unique class เพื่อไม่ให้ style ชนกัน
  var scopeId = 'pp5sec_' + Math.random().toString(36).substr(2, 6);
  var scopedStyle = styleContent.replace(/(^|\})\s*([^@{}][^{]*)\{/g, function(match, prefix, selector) {
    // แยก comma-separated selectors แล้ว prefix แต่ละตัว
    var parts = selector.split(',').map(function(s) { return '.' + scopeId + ' ' + s.trim(); });
    return prefix + ' ' + parts.join(', ') + ' {';
  });

  return '<div class="page-break"></div>'
    + '<style>' + scopedStyle + '</style>'
    + '<div class="' + scopeId + '">' + bodyContent + '</div>';
}


// ============================================================
// 🎨 CSS พื้นฐาน
// ============================================================
function pp5book_baseCss_() {
  return '\
    @page { size: A4 portrait; margin: 15mm 12mm 15mm 15mm; }\
    * { box-sizing: border-box; }\
    body { font-family: "Sarabun", sans-serif; font-size: 14pt; line-height: 1.3; color: #000; margin: 0; padding: 0; font-weight: 400; }\
    .page-break { page-break-before: always; }\
    table { width: 100%; border-collapse: collapse; }\
    th, td { border: 1px solid #000; padding: 3px 5px; vertical-align: middle; }\
    th { background: #f2f2f2; font-weight: 700; text-align: center; }\
    td { font-size: 12pt; font-weight: 400 !important; }\
    .center { text-align: center; }\
    .left { text-align: left; padding-left: 5px; }\
    .nowrap { white-space: nowrap; }\
    .section-title { font-weight: 700; font-size: 14pt; margin: 8px 0 4px; }\
    .page-header { text-align: center; margin-bottom: 8px; }\
    .page-header img { height: 70px; }\
    .page-header h2 { margin: 4px 0; font-size: 16pt; }\
    .page-header h3 { margin: 2px 0; font-size: 14pt; font-weight: 400; }\
    .signature-box { margin-top: 30px; display: flex; justify-content: space-around; }\
    .signature-box div { text-align: center; font-size: 12pt; }\
  ';
}


// ============================================================
// 📄 Wrapper HTML
// ============================================================
function pp5book_wrapHtml_(parts) {
  return '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>ปพ.5 รวมเล่ม</title>'
    + '<link href="https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=Prompt:wght@300;400;600&display=swap" rel="stylesheet">'
    + '<style>' + pp5book_baseCss_() + '</style>'
    + '</head><body>'
    + parts.join('')
    + '</body></html>';
}


// ============================================================
// 📄 ส่วนที่ 1: ปก ปพ.5
// ============================================================
function pp5book_coverHtml_(meta, studentCount, subjectSummary, reading, attributes, activity) {
  var logoHtml = meta.logoDataUri ? '<img src="' + meta.logoDataUri + '" style="height:60px;">' : '';

  // --- เรียงลำดับวิชาตามที่กำหนด ---
  var coverOrder = [
    'ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์และเทคโนโลยี',
    'สังคมศึกษา ศาสนา และวัฒนธรรม', 'ประวัติศาสตร์',
    'สุขศึกษาและพลศึกษา', 'ศิลปะ', 'การงานอาชีพ', 'ภาษาอังกฤษ',
    'หน้าที่พลเมือง', 'การป้องกันการทุจริต'
  ];
  var sorted = (subjectSummary || []).slice().sort(function(a, b) {
    var ia = -1, ib = -1;
    for (var k = 0; k < coverOrder.length; k++) {
      if ((a.name || '').indexOf(coverOrder[k]) !== -1 || coverOrder[k].indexOf(a.name || '') !== -1) ia = k;
      if ((b.name || '').indexOf(coverOrder[k]) !== -1 || coverOrder[k].indexOf(b.name || '') !== -1) ib = k;
    }
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return (a.name || '').localeCompare(b.name || '', 'th');
  });

  var subjectRows = sorted.map(function(s, i) {
    return '<tr><td class="center">' + (i + 1) + '</td><td class="left">' + (s.name || '') + '</td>'
      + '<td class="center">' + (s.count["4"] || 0) + '</td><td class="center">' + (s.count["3.5"] || 0) + '</td>'
      + '<td class="center">' + (s.count["3"] || 0) + '</td><td class="center">' + (s.count["2.5"] || 0) + '</td>'
      + '<td class="center">' + (s.count["2"] || 0) + '</td><td class="center">' + (s.count["1.5"] || 0) + '</td>'
      + '<td class="center">' + (s.count["1"] || 0) + '</td><td class="center">' + (s.count["0"] || 0) + '</td></tr>';
  }).join('');

  var pct = function(n) { return studentCount > 0 ? ((n || 0) / studentCount * 100).toFixed(1) : '0.0'; };

  return '<style>'
    + '.cover-page { font-size: 13pt; }'
    + '.cover-page .page-header { margin-bottom: 6px; }'
    + '.cover-page .page-header img { height: 70px; }'
    + '.cover-page .page-header h2 { margin: 3px 0; font-size: 16pt; }'
    + '.cover-page .page-header h3 { margin: 2px 0; font-size: 13pt; font-weight: 400; }'
    + '.cover-page table { font-size: 13pt; }'
    + '.cover-page th, .cover-page td { padding: 3px 4px; font-size: 13pt; }'
    + '.cover-page .section-title { font-size: 14pt; margin: 8px 0 4px; }'
    + '.cover-page .signature-box { margin-top: 30px; font-size: 12pt; }'
    + '.cover-page .signature-box div { line-height: 1.6; }'
    + '</style>'
    + '<div class="cover-page">'
    + '<div class="page-header">'
    + logoHtml
    + '<h2>แบบบันทึกผลการเรียนประจำรายวิชา (ปพ.5)</h2>'
    + '<h3>' + (meta.schoolName || '') + '</h3>'
    + '<h3>' + (meta.settings['ที่อยู่โรงเรียน'] || '') + '</h3>'
    + '<h3>ปีการศึกษา ' + meta.academicYear + '</h3>'
    + '</div>'
    + '<div class="section-title">ชั้น' + meta.gradeFullName + ' ห้อง ' + meta.classNo + ' จำนวนนักเรียน ' + studentCount + ' คน</div>'

    // ตารางสรุปเกรดรายวิชา
    + '<table><thead>'
    + '<tr><th rowspan="2" style="width:22px;">ที่</th><th rowspan="2">ชื่อวิชา</th><th colspan="8">ระดับผลการเรียนรายวิชา (คน)</th></tr>'
    + '<tr><th style="width:30px;">4</th><th style="width:30px;">3.5</th><th style="width:30px;">3</th><th style="width:30px;">2.5</th><th style="width:30px;">2</th><th style="width:30px;">1.5</th><th style="width:30px;">1</th><th style="width:30px;">0</th></tr>'
    + '</thead><tbody>' + subjectRows + '</tbody></table>'

    // ตารางประเมินบูรณาการ
    + '<div class="section-title" style="margin-top:10px;">การประเมินผลแบบบูรณาการ</div>'
    + '<table><thead>'
    + '<tr><th colspan="3">การอ่าน คิดวิเคราะห์และเขียน</th><th colspan="3">คุณลักษณะอันพึงประสงค์</th><th colspan="3">กิจกรรมพัฒนาผู้เรียน</th></tr>'
    + '<tr><th>ระดับคุณภาพ</th><th>คน</th><th>ร้อยละ</th><th>ระดับคุณภาพ</th><th>คน</th><th>ร้อยละ</th><th>การผ่าน</th><th>คน</th><th>ร้อยละ</th></tr>'
    + '</thead><tbody>'
    + '<tr><td class="left">ดีเยี่ยม</td><td class="center">' + (reading.level3||0) + '</td><td class="center">' + pct(reading.level3) + '</td>'
    + '<td class="left">ดีเยี่ยม</td><td class="center">' + (attributes['ดีเยี่ยม']||0) + '</td><td class="center">' + pct(attributes['ดีเยี่ยม']) + '</td>'
    + '<td class="left">ผ่าน</td><td class="center">' + (activity['ผ่าน']||0) + '</td><td class="center">' + pct(activity['ผ่าน']) + '</td></tr>'
    + '<tr><td class="left">ดี</td><td class="center">' + (reading.level2||0) + '</td><td class="center">' + pct(reading.level2) + '</td>'
    + '<td class="left">ดี</td><td class="center">' + (attributes['ดี']||0) + '</td><td class="center">' + pct(attributes['ดี']) + '</td>'
    + '<td class="left">ไม่ผ่าน</td><td class="center">' + (activity['ไม่ผ่าน']||0) + '</td><td class="center">' + pct(activity['ไม่ผ่าน']) + '</td></tr>'
    + '<tr><td class="left">ผ่านเกณฑ์</td><td class="center">' + (reading.level1||0) + '</td><td class="center">' + pct(reading.level1) + '</td>'
    + '<td class="left">ผ่านเกณฑ์</td><td class="center">' + (attributes['ผ่าน']||0) + '</td><td class="center">' + pct(attributes['ผ่าน']) + '</td>'
    + '<td></td><td></td><td></td></tr>'
    + '<tr><td class="left">ไม่ผ่านเกณฑ์</td><td class="center">' + (reading.level0||0) + '</td><td class="center">' + pct(reading.level0) + '</td>'
    + '<td class="left">ไม่ผ่านเกณฑ์</td><td class="center">' + (attributes['ไม่ผ่าน']||0) + '</td><td class="center">' + pct(attributes['ไม่ผ่าน']) + '</td>'
    + '<td></td><td></td><td></td></tr>'
    + '</tbody></table>'

    // ลงชื่อ
    + '<div class="signature-box">'
    + '<div>ลงชื่อ......................................<br>(' + (meta.teacherName || '................................') + ')<br>ครูประจำชั้น</div>'
    + '<div>☐ เห็นควรอนุมัติ ☐ เห็นควรปรับปรุง<br>ลงชื่อ......................................<br>(' + (meta.academicHead || 'หัวหน้างานวิชาการ') + ')<br>หัวหน้างานวิชาการ</div>'
    + '<div>☐ อนุมัติ ☐ ไม่อนุมัติ<br>ลงชื่อ......................................<br>(' + (meta.directorName || 'ผู้อำนวยการ') + ')<br>' + meta.directorTitle + '</div>'
    + '</div>'
    + '</div>';
}


// ============================================================
// 📄 ส่วนที่ 2: รายชื่อนักเรียน
// ============================================================
function pp5book_studentListHtml_(meta, students) {
  var rows = students.map(function(s, i) {
    var bd = s.birthdate || '';
    var parts = bd.split('-');
    var dayStr = '', monthStr = '', yearStr = '';
    if (parts.length === 3) {
      dayStr = parts[2]; monthStr = parts[1]; yearStr = String(Number(parts[0]) + 543);
    }
    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="center" style="font-size:10pt;">' + (s.id || '') + '</td>'
      + '<td class="center" style="font-size:10pt;">' + (s.idCard || '') + '</td>'
      + '<td class="left" style="font-size:11pt;">' + (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '') + '</td>'
      + '<td class="center">' + dayStr + '</td>'
      + '<td class="center">' + monthStr + '</td>'
      + '<td class="center">' + yearStr + '</td>'
      + '<td class="center">' + (s.age || '') + '</td>'
      + '<td class="center">' + (s.weight || '') + '</td>'
      + '<td class="center">' + (s.height || '') + '</td>'
      + '</tr>';
  }).join('');

  return '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">ข้อมูลนักเรียน ชั้น' + meta.gradeFullName + ' ห้อง ' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<table>'
    + '<thead><tr>'
    + '<th rowspan="2" style="width:30px;">ที่</th>'
    + '<th rowspan="2" style="width:75px;">เลขประจำตัว</th>'
    + '<th rowspan="2" style="width:100px;">เลขประจำตัวประชาชน</th>'
    + '<th rowspan="2" style="width:180px;">ชื่อ - สกุล</th>'
    + '<th colspan="3">วัน เดือน ปี เกิด</th>'
    + '<th rowspan="2" style="width:35px;">อายุ</th>'
    + '<th rowspan="2" style="width:40px;">น้ำหนัก</th>'
    + '<th rowspan="2" style="width:40px;">ส่วนสูง</th>'
    + '</tr><tr>'
    + '<th style="width:30px;">วัน</th><th style="width:30px;">เดือน</th><th style="width:35px;">ปี</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody></table>';
}


// ============================================================
// 📄 ส่วนที่ 3: ข้อมูลผู้ปกครอง
// ============================================================
function pp5book_parentInfoHtml_(meta, students) {
  var rows = students.map(function(s, i) {
    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:11pt;">' + (s.father_name || '-') + '</td>'
      + '<td class="center" style="font-size:11pt;">' + (s.father_occupation || '-') + '</td>'
      + '<td class="left" style="font-size:11pt;">' + (s.mother_name || '-') + '</td>'
      + '<td class="center" style="font-size:11pt;">' + (s.mother_occupation || '-') + '</td>'
      + '<td class="left" style="font-size:10pt;">' + (s.address || '-') + '</td>'
      + '</tr>';
  }).join('');

  return '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">ข้อมูลผู้ปกครอง ชั้น' + meta.gradeFullName + ' ห้อง ' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<table>'
    + '<thead><tr>'
    + '<th rowspan="2" style="width:30px;">ที่</th>'
    + '<th colspan="2">ข้อมูลบิดา</th>'
    + '<th colspan="2">ข้อมูลมารดา</th>'
    + '<th rowspan="2">ที่อยู่ปัจจุบัน</th>'
    + '</tr><tr>'
    + '<th style="width:130px;">ชื่อ-สกุล</th><th style="width:80px;">อาชีพ</th>'
    + '<th style="width:130px;">ชื่อ-สกุล</th><th style="width:80px;">อาชีพ</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody></table>';
}


// ============================================================
// 📄 ส่วนที่ 4: คะแนนรวม 2 ภาค (รูปแบบเหมือน PDF จากหน้าบันทึกคะแนน)
// ============================================================
function pp5book_yearScoreHtml_(meta, sheetInfo, scoreData) {
  var t1 = scoreData.term1 || {};
  var t2 = scoreData.term2 || {};
  if ((!t1.rows || t1.rows.length === 0) && (!t2.rows || t2.rows.length === 0)) return '';

  var rows1 = t1.rows || [];
  var rows2 = t2.rows || [];
  var yearSummary = scoreData.yearSummary || [];

  // สร้าง map จาก yearSummary เพื่อ lookup ด้วย id
  var yearMap = {};
  yearSummary.forEach(function(ys) { yearMap[String(ys.id)] = ys; });

  var logoHtml = meta.logoDataUri ? '<img src="' + meta.logoDataUri + '" style="width:60px;height:60px;object-fit:contain;">' : '';

  var dataRows = rows1.map(function(r1, i) {
    var r2 = rows2.find(function(r) { return r.id === r1.id; }) || rows2[i] || {};
    var total1 = Number(r1.total) || 0;
    var total2 = Number(r2.total) || 0;
    var grade1 = r1.grade || '0';
    var grade2 = r2.grade || '0';

    // ใช้ค่าจากชีตโดยตรง (คอลัมน์ AH-AI)
    var ys = yearMap[String(r1.id)] || {};
    var yearAvg = ys.yearAvg || 0;
    var yearGrade = ys.yearGrade || '';

    // fallback: ถ้าชีตยังไม่มีค่า ให้คำนวณเอง
    if (!yearGrade && (total1 > 0 || total2 > 0)) {
      yearAvg = Math.round((total1 + total2) / 2);
      yearGrade = calculateFinalGrade(yearAvg);
    }

    var bgColor = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    return '<tr style="background:' + bgColor + ';">'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="center" style="font-size:11pt;">' + (r1.id || '') + '</td>'
      + '<td class="left" style="font-size:11pt; padding-left:5px; white-space:nowrap;">' + (r1.name || '') + '</td>'
      + '<td class="center">' + total1 + '</td>'
      + '<td class="center" style="font-weight:600;">' + grade1 + '</td>'
      + '<td class="center">' + total2 + '</td>'
      + '<td class="center" style="font-weight:600;">' + grade2 + '</td>'
      + '<td class="center" style="font-weight:700; background:#eaeff7;">' + yearAvg + '</td>'
      + '<td class="center" style="font-weight:700; background:#e8f5e9; font-size:13pt;">' + yearGrade + '</td>'
      + '</tr>';
  }).join('');

  return '<div class="page-break"></div>'
    + '<div style="text-align:center; margin-bottom:5px;">' + logoHtml + '</div>'
    + '<div style="text-align:center; font-size:14pt; font-weight:700;">ผลการประเมิน รายวิชา: ' + (sheetInfo.subjectName || '') + ' (' + (sheetInfo.subjectCode || '') + ')</div>'
    + '<div style="text-align:center; font-size:12pt; margin:3px 0 8px 0;">' + meta.grade + '/' + meta.classNo + ' | รวมทั้งปีการศึกษา | ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<table style="margin-top:4px; font-size:12pt;">'
    + '<thead>'
    + '<tr style="background:#eaeff7;">'
    + '<th rowspan="2" style="width:35px;">เลขที่</th>'
    + '<th rowspan="2" style="width:70px;">รหัส</th>'
    + '<th rowspan="2" style="width:200px;">ชื่อ-นามสกุล</th>'
    + '<th colspan="2" style="background:#e3f2fd;">ภาคเรียนที่ 1</th>'
    + '<th colspan="2" style="background:#fff3e0;">ภาคเรียนที่ 2</th>'
    + '<th colspan="2" style="background:#e8f5e9;">รวมทั้งปี</th>'
    + '</tr>'
    + '<tr style="background:#eaeff7;">'
    + '<th style="width:55px; background:#e3f2fd;">คะแนน</th><th style="width:50px; background:#e3f2fd;">ผลการเรียน</th>'
    + '<th style="width:55px; background:#fff3e0;">คะแนน</th><th style="width:50px; background:#fff3e0;">ผลการเรียน</th>'
    + '<th style="width:60px; background:#e8f5e9;">คะแนนเฉลี่ย 2 ภาค</th><th style="width:65px; background:#e8f5e9;">ระดับผลการเรียน</th>'
    + '</tr>'
    + '</thead>'
    + '<tbody>' + dataRows + '</tbody></table>'
    + '<div style="margin-top:25px; display:flex; justify-content:space-around; font-size:11pt;">'
    + '<div style="text-align:center;"><div>ลงชื่อ...............................................ครูผู้สอน</div><div style="margin-top:3px;">(.............................................)</div><div>ตำแหน่ง ครู</div></div>'
    + '<div style="text-align:center;"><div>ลงชื่อ...............................................ผู้อำนวยการ</div><div style="margin-top:3px;">(' + (meta.directorName || '..............................................') + ')</div><div>' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา') + '</div></div>'
    + '</div>';
}


// ============================================================
// 📄 ส่วนที่ 5: สรุปเวลาเรียน
// ============================================================
function pp5book_attendanceHtml_(meta, students, att1, att2) {
  var att1Map = {};
  (att1 || []).forEach(function(a) { att1Map[a.studentId] = a; });
  var att2Map = {};
  (att2 || []).forEach(function(a) { att2Map[a.studentId] = a; });

  var rows = students.map(function(s, i) {
    var a1 = att1Map[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
    var a2 = att2Map[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
    var totalPresent = a1.present + a2.present;
    var totalLeave = a1.leave + a2.leave;
    var totalAbsent = a1.absent + a2.absent;
    var totalAll = a1.total + a2.total;
    var pct = totalAll > 0 ? ((totalPresent / totalAll) * 100).toFixed(1) : '0.0';

    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:11pt;">' + (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '') + '</td>'
      + '<td class="center">' + a1.present + '</td><td class="center">' + a1.leave + '</td><td class="center">' + a1.absent + '</td><td class="center">' + a1.total + '</td>'
      + '<td class="center">' + a2.present + '</td><td class="center">' + a2.leave + '</td><td class="center">' + a2.absent + '</td><td class="center">' + a2.total + '</td>'
      + '<td class="center" style="font-weight:600;">' + totalPresent + '</td>'
      + '<td class="center">' + totalLeave + '</td>'
      + '<td class="center">' + totalAbsent + '</td>'
      + '<td class="center" style="font-weight:600;">' + totalAll + '</td>'
      + '<td class="center" style="font-weight:700;">' + pct + '</td>'
      + '</tr>';
  }).join('');

  return '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">สรุปเวลาเรียน ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<table style="font-size:11pt;">'
    + '<thead><tr>'
    + '<th rowspan="2" style="width:25px;">ที่</th>'
    + '<th rowspan="2" style="width:170px;">ชื่อ-นามสกุล</th>'
    + '<th colspan="4">ภาคเรียนที่ 1 (วัน)</th>'
    + '<th colspan="4">ภาคเรียนที่ 2 (วัน)</th>'
    + '<th colspan="5">รวมทั้งปี</th>'
    + '</tr><tr>'
    + '<th style="width:30px;">มา</th><th style="width:30px;">ลา</th><th style="width:30px;">ขาด</th><th style="width:30px;">รวม</th>'
    + '<th style="width:30px;">มา</th><th style="width:30px;">ลา</th><th style="width:30px;">ขาด</th><th style="width:30px;">รวม</th>'
    + '<th style="width:30px;">มา</th><th style="width:30px;">ลา</th><th style="width:30px;">ขาด</th><th style="width:30px;">รวม</th><th style="width:40px;">%มา</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody></table>';
}


// ============================================================
// 📄 ส่วนที่ 5.1: เวลาเรียนรายเดือน (เช็คชื่อรายเดือน)
// ============================================================
function pp5book_monthlyAttendanceHtml_(meta, students, monthlyDataAll) {
  if (!monthlyDataAll || monthlyDataAll.length === 0) return '';

  var logoHtml = meta.logoDataUri ? '<img src="' + meta.logoDataUri + '" style="width:50px;height:50px;object-fit:contain;">' : '';

  // สร้าง map นักเรียน id → ข้อมูล
  var studentList = students.map(function(s) {
    return { id: s.id, name: (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '') };
  });

  var pages = [];

  monthlyDataAll.forEach(function(monthInfo) {
    // สร้าง map ข้อมูลเดือนนี้: studentId → { present, leave, absent }
    var dataMap = {};
    (monthInfo.data || []).forEach(function(d) {
      dataMap[d.studentId] = d;
    });

    var rows = studentList.map(function(s, i) {
      var d = dataMap[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
      var bgColor = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
      return '<tr style="background:' + bgColor + ';">'
        + '<td class="center">' + (i + 1) + '</td>'
        + '<td class="left" style="font-size:11pt; white-space:nowrap;">' + s.name + '</td>'
        + '<td class="center" style="color:#2e7d32; font-weight:600;">' + (d.present || 0) + '</td>'
        + '<td class="center" style="color:#e65100; font-weight:600;">' + (d.leave || 0) + '</td>'
        + '<td class="center" style="color:#c62828; font-weight:600;">' + (d.absent || 0) + '</td>'
        + '<td class="center" style="font-weight:700;">' + (d.total || 0) + '</td>'
        + '</tr>';
    }).join('');

    // คำนวณรวมทั้งห้อง
    var sumPresent = 0, sumLeave = 0, sumAbsent = 0;
    studentList.forEach(function(s) {
      var d = dataMap[s.id] || {};
      sumPresent += d.present || 0;
      sumLeave += d.leave || 0;
      sumAbsent += d.absent || 0;
    });
    var sumTotal = sumPresent + sumLeave + sumAbsent;

    var footerRow = '<tr style="background:#e3f2fd; font-weight:700;">'
      + '<td class="center" colspan="2">รวมทั้งห้อง</td>'
      + '<td class="center" style="color:#2e7d32;">' + sumPresent + '</td>'
      + '<td class="center" style="color:#e65100;">' + sumLeave + '</td>'
      + '<td class="center" style="color:#c62828;">' + sumAbsent + '</td>'
      + '<td class="center">' + sumTotal + '</td>'
      + '</tr>';

    pages.push('<div class="page-break"></div>'
      + '<div style="text-align:center; margin-bottom:3px;">' + logoHtml + '</div>'
      + '<div style="text-align:center; font-size:14pt; font-weight:700;">บันทึกเวลาเรียน (เช็คชื่อรายเดือน)</div>'
      + '<div style="text-align:center; font-size:12pt; margin:2px 0 6px 0;">'
      + meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo
      + ' | เดือน' + monthInfo.name + ' ' + monthInfo.yearBE
      + ' | ปีการศึกษา ' + meta.academicYear + '</div>'
      + '<table style="font-size:12pt;">'
      + '<thead>'
      + '<tr style="background:#eaeff7;">'
      + '<th style="width:30px;">ที่</th>'
      + '<th style="width:200px;">ชื่อ-นามสกุล</th>'
      + '<th style="width:55px; background:#e8f5e9;">มา</th>'
      + '<th style="width:55px; background:#fff3e0;">ลา</th>'
      + '<th style="width:55px; background:#ffebee;">ขาด</th>'
      + '<th style="width:55px;">รวม</th>'
      + '</tr>'
      + '</thead>'
      + '<tbody>' + rows + footerRow + '</tbody></table>'
      + '<div style="margin-top:20px; display:flex; justify-content:space-around; font-size:11pt;">'
      + '<div style="text-align:center;"><div>ลงชื่อ...............................................ครูประจำชั้น</div><div style="margin-top:3px;">(' + (meta.teacherName || '..............................................') + ')</div></div>'
      + '<div style="text-align:center;"><div>ลงชื่อ...............................................ผู้อำนวยการ</div><div style="margin-top:3px;">(' + (meta.directorName || '..............................................') + ')</div><div>' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา') + '</div></div>'
      + '</div>'
    );
  });

  return pages.join('');
}


// ============================================================
// 📄 ส่วนที่ 4C: ตารางเช็คชื่อรายวัน (เหมือน PDF จากหน้าเช็คชื่อ)
// ============================================================
function pp5book_dailyAttendanceHtml_(meta, tableData, monthName, yearBE, semester) {
  var students = tableData.students || [];
  var days = tableData.days || [];
  var actualSchoolDays = tableData.actualSchoolDays || days.length;
  var totalDaysInMonth = tableData.totalDaysInMonth || days.length;

  if (students.length === 0 || days.length === 0) return '';

  var logoHtml = meta.logoDataUri ? '<img src="' + meta.logoDataUri + '" style="width:45px;height:45px;object-fit:contain;">' : '';

  // คำนวณ font size ตามจำนวนวัน
  var dateFontSize = days.length > 22 ? '7pt' : (days.length > 18 ? '8pt' : '9pt');
  var cellWidth = days.length > 22 ? '16px' : (days.length > 18 ? '18px' : '20px');

  // header วันที่ (แถว 1: เลขวัน)
  var dayHeaders = days.map(function(d) {
    return '<th style="width:' + cellWidth + '; font-size:' + dateFontSize + '; padding:1px;">' + d.label + '</th>';
  }).join('');

  // header วันในสัปดาห์ (แถว 2: จ, อ, พ...)
  var dowHeaders = days.map(function(d) {
    return '<th style="width:' + cellWidth + '; font-size:' + dateFontSize + '; padding:1px; font-weight:400;">' + d.dow + '</th>';
  }).join('');

  // แถวข้อมูลนักเรียน
  var dataRows = students.map(function(stu, i) {
    var bgColor = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
    var cells = days.map(function(d) {
      var status = stu.statusMap[d.label] || '';
      var display = '';
      var color = '';
      if (status === '/' || status === 'ม' || status === '1') {
        display = '✓'; color = '#2e7d32';
      } else if (status === 'ล' || String(status).toLowerCase() === 'l') {
        display = 'ล'; color = '#e65100';
      } else if (status === 'ข' || status === '0') {
        display = 'ข'; color = '#c62828';
      }
      var fw = (display === 'ล' || display === 'ข') ? 'font-weight:600;' : 'font-weight:normal;';
      return '<td class="center" style="font-size:' + dateFontSize + '; padding:1px;' + (color ? ' color:' + color + '; ' + fw : '') + '">' + display + '</td>';
    }).join('');

    return '<tr style="background:' + bgColor + ';">'
      + '<td class="center" style="font-size:9pt; padding:1px 2px;">' + (i + 1) + '</td>'
      + '<td class="center" style="font-size:8pt; padding:1px 2px;">' + (stu.id || '') + '</td>'
      + '<td class="left" style="font-size:8pt; padding:1px 3px; white-space:nowrap;">' + (stu.name || '') + '</td>'
      + cells + '</tr>';
  }).join('');

  // หมายเหตุ
  var noteHtml = '<div style="margin-top:8px; font-size:9pt;">'
    + 'หมายเหตุ: ✓ = มาเรียน &nbsp; ล = ลา &nbsp; ข = ขาด &nbsp;| นักเรียน: ' + students.length + ' คน &nbsp;| วันเรียน: ' + actualSchoolDays + ' วัน'
    + '<div style="text-align:right; font-size:8pt; color:#888; margin-top:2px;">พิมพ์: ' + new Date().toLocaleDateString('th-TH') + ' ' + new Date().toLocaleTimeString('th-TH', {hour:'2-digit', minute:'2-digit'}) + ' น.</div>'
    + '</div>';

  return '<div class="page-break"></div>'
    + '<style>.daily-att-page { font-family:"Prompt",sans-serif; font-weight:300; } .daily-att-page table { font-size: 9pt; border-collapse:collapse; width:100%; } .daily-att-page th, .daily-att-page td { padding: 1px 2px; font-family:"Prompt",sans-serif; } .daily-att-page tbody td { font-weight:300; font-family:"Prompt",sans-serif; }</style>'
    + '<div class="daily-att-page">'
    + '<div style="text-align:center; margin-bottom:2px;">' + logoHtml + '</div>'
    + '<div style="text-align:center; font-size:13pt; font-weight:700;">' + (meta.schoolName || '') + '</div>'
    + '<div style="text-align:center; font-size:14pt; font-weight:700;">บันทึกเวลาเรียน ชั้น ' + meta.grade + '/' + meta.classNo + ' เดือน ' + monthName + ' ' + yearBE + '</div>'
    + '<div style="text-align:center; font-size:11pt; margin:1px 0 3px 0;">' + semester + ' | ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<div style="text-align:center; font-size:10pt; margin-bottom:4px;">แสดงเฉพาะวันที่มีการเรียน: ' + actualSchoolDays + ' วัน (จาก ' + totalDaysInMonth + ' วันในเดือน)</div>'
    + '<table>'
    + '<thead>'
    + '<tr style="background:#f3f3f3;">'
    + '<th rowspan="2" style="width:22px; font-size:9pt;">ที่</th>'
    + '<th rowspan="2" style="width:45px; font-size:9pt;">รหัส</th>'
    + '<th rowspan="2" style="width:140px; font-size:9pt;">ชื่อ - สกุล</th>'
    + dayHeaders + '</tr>'
    + '<tr style="background:#f3f3f3;">' + dowHeaders + '</tr>'
    + '</thead>'
    + '<tbody>' + dataRows + '</tbody></table>'
    + noteHtml
    + '</div>';
}


// ============================================================
// 📄 ส่วนที่ 6: สรุปการประเมิน (คุณลักษณะ + กิจกรรม)
// ============================================================
function pp5book_assessmentHtml_(meta, charData, actData) {
  // --- คุณลักษณะอันพึงประสงค์ ---
  var traitNames = ['รักชาติ ศาสน์ กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้', 'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ'];

  var charRows = (charData || []).map(function(s, i) {
    var scores = s.scores || [];
    var cells = '';
    for (var k = 0; k < 8; k++) {
      cells += '<td class="center">' + (scores[k] !== '' && scores[k] !== undefined ? scores[k] : '') + '</td>';
    }
    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:10pt;">' + (s.name || '') + '</td>'
      + cells
      + '<td class="center" style="font-weight:600;">' + (s.sum || '') + '</td>'
      + '<td class="center" style="font-weight:600;">' + (s.avg || '') + '</td>'
      + '<td class="center" style="font-weight:700;">' + (s.result || '') + '</td>'
      + '</tr>';
  }).join('');

  var charTable = '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">ผลการประเมินคุณลักษณะอันพึงประสงค์ ชั้น' + meta.gradeFullName + '/' + meta.classNo + '</div>'
    + '<table style="font-size:10pt;">'
    + '<thead><tr>'
    + '<th rowspan="2" style="width:25px;">ที่</th>'
    + '<th rowspan="2" style="width:140px;">ชื่อ-นามสกุล</th>';
  for (var t = 0; t < traitNames.length; t++) {
    charTable += '<th style="width:45px; font-size:9pt;">' + traitNames[t] + '</th>';
  }
  charTable += '<th rowspan="2" style="width:35px;">รวม</th>'
    + '<th rowspan="2" style="width:35px;">เฉลี่ย</th>'
    + '<th rowspan="2" style="width:50px;">ผลประเมิน</th>'
    + '</tr></thead>'
    + '<tbody>' + charRows + '</tbody></table>';

  // --- กิจกรรมพัฒนาผู้เรียน ---
  var actRows = (actData || []).map(function(s, i) {
    var a = s.activities || {};
    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:10pt;">' + (s.name || '') + '</td>'
      + '<td class="center">' + (a.guidance || 'ผ่าน') + '</td>'
      + '<td class="center">' + (a.scout || 'ผ่าน') + '</td>'
      + '<td class="center">' + (a.club || 'ผ่าน') + '</td>'
      + '<td class="center">' + (a.social || 'ผ่าน') + '</td>'
      + '<td class="center" style="font-weight:700;">' + (a.overall || 'ผ่าน') + '</td>'
      + '</tr>';
  }).join('');

  var actTable = '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">ผลการประเมินกิจกรรมพัฒนาผู้เรียน ชั้น' + meta.gradeFullName + '/' + meta.classNo + '</div>'
    + '<table>'
    + '<thead><tr>'
    + '<th style="width:30px;">ที่</th>'
    + '<th style="width:170px;">ชื่อ-นามสกุล</th>'
    + '<th>แนะแนว</th><th>ลูกเสือ/เนตรนารี</th><th>ชุมนุม</th>'
    + '<th>จิตอาสา</th><th>สรุปผล</th>'
    + '</tr></thead>'
    + '<tbody>' + actRows + '</tbody></table>';

  return charTable + actTable;
}


// ============================================================
// 📊 ดึงข้อมูลคุณลักษณะอันพึงประสงค์จากชีตโดยตรง
// (ใช้ค่า คะแนนรวม เฉลี่ย ผลการประเมิน ที่บันทึกจริง)
// ============================================================
function pp5book_getCharacteristicData_(grade, classNo) {
  grade = String(grade || '').trim();
  classNo = String(classNo || '').trim();
  var ss = SS();

  // --- ดึงรายชื่อนักเรียน ---
  var studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  var stuData = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = stuData
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo && String(r[0]); })
    .map(function(r) {
      return {
        studentId: String(r[0]).trim(),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim()
      };
    });
  students.sort(function(a, b) { return String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true }); });

  // --- ดึงข้อมูลจากชีต การประเมินคุณลักษณะ ---
  var charSheet = ss.getSheetByName('การประเมินคุณลักษณะ');
  var charMap = {};
  if (charSheet && charSheet.getLastRow() > 1) {
    var headers = charSheet.getRange(1, 1, 1, charSheet.getLastColumn()).getValues()[0].map(String);
    var hIdx = {};
    headers.forEach(function(h, i) { hIdx[h.trim()] = i; });

    var allData = charSheet.getRange(2, 1, charSheet.getLastRow() - 1, charSheet.getLastColumn()).getValues();
    var traitCols = [
      'รักชาติ_ศาสน์_กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้',
      'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ'
    ];

    allData.forEach(function(row) {
      var rGrade = String(row[hIdx['ชั้น']] || '').trim();
      var rClass = String(row[hIdx['ห้อง']] || '').trim();
      if (rGrade !== grade || rClass !== classNo) return;

      var sid = String(row[hIdx['รหัสนักเรียน']] || '').trim();
      var scores = traitCols.map(function(col) {
        var v = hIdx[col] !== undefined ? row[hIdx[col]] : '';
        return (v !== '' && v !== null && v !== undefined) ? Number(v) : '';
      });

      var sum = hIdx['คะแนนรวม'] !== undefined ? row[hIdx['คะแนนรวม']] : '';
      var avg = hIdx['คะแนนเฉลี่ย'] !== undefined ? row[hIdx['คะแนนเฉลี่ย']] : '';
      var result = hIdx['ผลการประเมิน'] !== undefined ? String(row[hIdx['ผลการประเมิน']] || '').trim() : '';

      charMap[sid] = { scores: scores, sum: sum, avg: avg, result: result };
    });
  }

  // --- รวมข้อมูล ---
  return students.map(function(s) {
    var saved = charMap[s.studentId] || {};
    return {
      studentId: s.studentId,
      name: s.name,
      scores: saved.scores || Array(8).fill(''),
      sum: saved.sum || '',
      avg: saved.avg || '',
      result: saved.result || ''
    };
  });
}


// ============================================================
// 📊 ดึงข้อมูล อ่าน คิดวิเคราะห์ เขียน จากชีตโดยตรง
// ============================================================
function pp5book_getRTWData_(grade, classNo) {
  grade = String(grade || '').trim();
  classNo = String(classNo || '').trim();
  var ss = SS();

  // --- ดึงรายชื่อนักเรียน ---
  var studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  var stuData = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = stuData
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo && String(r[0]); })
    .map(function(r) {
      return {
        studentId: String(r[0]).trim(),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim()
      };
    });
  students.sort(function(a, b) { return String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true }); });

  // --- ดึงจากชีต การประเมินอ่านคิดเขียน ---
  var rtwSheet = ss.getSheetByName('การประเมินอ่านคิดเขียน');
  var rtwMap = {};
  var subjectCols = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์', 'สังคมศึกษา', 'สุขศึกษา', 'ศิลปะ', 'การงาน', 'ภาษาอังกฤษ'];

  if (rtwSheet && rtwSheet.getLastRow() > 1) {
    var headers = rtwSheet.getRange(1, 1, 1, rtwSheet.getLastColumn()).getValues()[0].map(String);
    var hIdx = {};
    headers.forEach(function(h, i) { hIdx[h.trim()] = i; });

    var allData = rtwSheet.getRange(2, 1, rtwSheet.getLastRow() - 1, rtwSheet.getLastColumn()).getValues();
    allData.forEach(function(row) {
      var rGrade = String(row[hIdx['ชั้น']] || '').trim();
      var rClass = String(row[hIdx['ห้อง']] || '').trim();
      if (rGrade !== grade || rClass !== classNo) return;

      var sid = String(row[hIdx['รหัสนักเรียน']] || '').trim();
      var scores = subjectCols.map(function(col) {
        return hIdx[col] !== undefined ? row[hIdx[col]] : '';
      });
      var summary = hIdx['สรุปผลการประเมิน'] !== undefined ? String(row[hIdx['สรุปผลการประเมิน']] || '').trim() : '';

      rtwMap[sid] = { scores: scores, summary: summary };
    });
  }

  return students.map(function(s) {
    var saved = rtwMap[s.studentId] || {};
    return {
      studentId: s.studentId,
      name: s.name,
      scores: saved.scores || Array(8).fill(''),
      summary: saved.summary || ''
    };
  });
}


// ============================================================
// 📄 ส่วนที่ 7: อ่าน คิดวิเคราะห์ เขียน (HTML)
// ============================================================
function pp5book_rtwHtml_(meta, rtwData) {
  if (!rtwData || rtwData.length === 0) return '';

  var subjectNames = ['ไทย', 'คณิต', 'วิทย์', 'สังคม', 'สุขศึกษา', 'ศิลปะ', 'การงาน', 'อังกฤษ'];

  var rows = rtwData.map(function(s, i) {
    var cells = '';
    for (var k = 0; k < 8; k++) {
      cells += '<td class="center">' + (s.scores[k] !== '' && s.scores[k] !== undefined ? s.scores[k] : '') + '</td>';
    }
    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:10pt;">' + (s.name || '') + '</td>'
      + cells
      + '<td class="center" style="font-weight:700;">' + (s.summary || '') + '</td>'
      + '</tr>';
  }).join('');

  var headerCells = '';
  for (var j = 0; j < subjectNames.length; j++) {
    headerCells += '<th style="width:45px;">' + subjectNames[j] + '</th>';
  }

  return '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">ผลการประเมินการอ่าน คิดวิเคราะห์ และเขียน</div>'
    + '<div style="text-align:center; font-size:12pt;">ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<div style="text-align:center; font-size:10pt; margin:4px 0 6px;">'
    + '0 = ต้องปรับปรุง &nbsp; 1 = ผ่านเกณฑ์ขั้นต่ำ &nbsp; 2 = ดี &nbsp; 3 = ดีเยี่ยม'
    + ' &nbsp;|&nbsp; <b>ปรับปรุง</b> 0.00-0.99 &nbsp; <b>ผ่าน</b> 1.00-1.99 &nbsp; <b>ดี</b> 2.00-2.49 &nbsp; <b>ดีเยี่ยม</b> 2.50-3.00'
    + '</div>'
    + '<table style="font-size:10pt;">'
    + '<thead><tr>'
    + '<th style="width:25px;">ที่</th>'
    + '<th style="width:140px;">ชื่อ-นามสกุล</th>'
    + headerCells
    + '<th style="width:55px;">สรุป</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody></table>';
}


// ============================================================
// 📊 ดึงข้อมูลคะแนนรายวิชา (Subject Score) จากชีต อ่านคิดเขียน
// (ตรงกับหน้า "บันทึกผลการประเมิน (ปพ.5)" ที่แสดงคะแนน 8 วิชา)
// ============================================================
function pp5book_getSubjectScoreData_(grade, classNo) {
  grade = String(grade || '').trim();
  classNo = String(classNo || '').trim();
  var ss = SS();

  // --- ดึงรายชื่อนักเรียน ---
  var studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  var stuData = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 7).getValues();
  var students = stuData
    .filter(function(r) { return String(r[5]) === grade && String(r[6]) === classNo && String(r[0]); })
    .map(function(r) {
      return {
        studentId: String(r[0]).trim(),
        name: ((r[2] || '') + (r[3] || '') + ' ' + (r[4] || '')).trim()
      };
    });
  students.sort(function(a, b) { return String(a.studentId).localeCompare(String(b.studentId), undefined, { numeric: true }); });

  // --- ดึงจากชีต การประเมินอ่านคิดเขียน (Subject Score ใช้ชีตเดียวกัน) ---
  // Subject Score Assessment จะดึงจาก getStudentsForSubjectScore ที่ใช้ชีต RTW
  // แต่เราต้องดึงโดยตรงเพื่อได้ข้อมูลครบ
  var rtwSheet = ss.getSheetByName('การประเมินอ่านคิดเขียน');
  var scoreMap = {};
  var subjectCols = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์', 'สังคมศึกษา', 'สุขศึกษา', 'ศิลปะ', 'การงาน', 'ภาษาอังกฤษ'];

  if (rtwSheet && rtwSheet.getLastRow() > 1) {
    var headers = rtwSheet.getRange(1, 1, 1, rtwSheet.getLastColumn()).getValues()[0].map(String);
    var hIdx = {};
    headers.forEach(function(h, i) { hIdx[h.trim()] = i; });

    var allData = rtwSheet.getRange(2, 1, rtwSheet.getLastRow() - 1, rtwSheet.getLastColumn()).getValues();
    allData.forEach(function(row) {
      var rGrade = String(row[hIdx['ชั้น']] || '').trim();
      var rClass = String(row[hIdx['ห้อง']] || '').trim();
      if (rGrade !== grade || rClass !== classNo) return;

      var sid = String(row[hIdx['รหัสนักเรียน']] || '').trim();
      var scores = {};
      subjectCols.forEach(function(col) {
        scores[col] = hIdx[col] !== undefined ? row[hIdx[col]] : '';
      });
      var summary = hIdx['สรุปผลการประเมิน'] !== undefined ? String(row[hIdx['สรุปผลการประเมิน']] || '').trim() : '';
      scoreMap[sid] = { scores: scores, summary: summary };
    });
  }

  return students.map(function(s) {
    var saved = scoreMap[s.studentId] || {};
    return {
      studentId: s.studentId,
      name: s.name,
      scores: saved.scores || {},
      summary: saved.summary || ''
    };
  });
}


// ============================================================
// 📄 ส่วนที่ 8: คะแนนรายวิชา (Subject Score Assessment) HTML
// (ตรงกับหน้า "บันทึกผลการประเมิน (ปพ.5)" — คะแนน 8 วิชา + สรุป)
// ============================================================
function pp5book_subjectScoreAssessmentHtml_(meta, data) {
  if (!data || data.length === 0) return '';

  var subjectNames = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์', 'สังคมศึกษา', 'สุขศึกษา', 'ศิลปะ', 'การงาน', 'ภาษาอังกฤษ'];
  var shortNames = ['ไทย', 'คณิต', 'วิทย์', 'สังคม', 'สุขศึกษา', 'ศิลปะ', 'การงาน', 'อังกฤษ'];

  var rows = data.map(function(s, i) {
    var scores = s.scores || {};
    var cells = '';
    subjectNames.forEach(function(subj) {
      var v = scores[subj];
      cells += '<td class="center">' + (v !== '' && v !== undefined && v !== null ? v : '') + '</td>';
    });

    return '<tr>'
      + '<td class="center">' + (i + 1) + '</td>'
      + '<td class="left" style="font-size:10pt;">' + (s.name || '') + '</td>'
      + cells
      + '<td class="center" style="font-weight:700;">' + (s.summary || '') + '</td>'
      + '</tr>';
  }).join('');

  var headerCells = '';
  shortNames.forEach(function(n) {
    headerCells += '<th style="width:45px;">' + n + '</th>';
  });

  return '<div class="page-break"></div>'
    + '<div class="section-title" style="text-align:center;">บันทึกผลการประเมิน (ปพ.5) อ่าน คิดวิเคราะห์ และเขียน</div>'
    + '<div style="text-align:center; font-size:12pt;">ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear + '</div>'
    + '<div style="text-align:center; font-size:10pt; margin:4px 0 6px;">'
    + '0 = ต้องปรับปรุง &nbsp; 1 = ผ่านเกณฑ์ขั้นต่ำ &nbsp; 2 = ดี &nbsp; 3 = ดีเยี่ยม'
    + '</div>'
    + '<table style="font-size:10pt;">'
    + '<thead><tr>'
    + '<th style="width:25px;">ที่</th>'
    + '<th style="width:140px;">ชื่อ-นามสกุล</th>'
    + headerCells
    + '<th style="width:55px;">สรุป</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody></table>';
}
