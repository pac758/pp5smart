// ============================================================
// 🔧 DEBUG: ตรวจสอบข้อมูลเช็คชื่อในชีต vs สิ่งที่ getSavedAttendance ส่งกลับ
// ============================================================
function debugAttendanceData() {
  var grade = 'ป.1', classNo = '1';
  var thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  var ss = SS();
  Logger.log('=== DEBUG v4: เปรียบเทียบภาคเรียน 1 vs 2 ===');

  // === 1. หา ป.1/1 IDs จาก Students ===
  var studentsSheet = ss.getSheetByName('Students');
  var sData = studentsSheet.getDataRange().getValues();
  var classMatchIds = new Set();
  for (var i = 1; i < sData.length; i++) {
    var sid = String(sData[i][0] || '').trim();
    var g = String(sData[i][5] || '').trim();
    var c = String(sData[i][6] || '').trim();
    if (g === grade && c === classNo && sid) classMatchIds.add(sid);
  }
  Logger.log('ป.1/1 IDs from Students: ' + Array.from(classMatchIds).join(', '));

  // === 2. ตรวจหลายเดือน: ภาคเรียน1 + ภาคเรียน2 ===
  var monthsToCheck = [
    { y: 2025, m: 5, label: 'มิ.ย.68 (ภาค1)' },   // มิถุนายน2568
    { y: 2025, m: 7, label: 'ส.ค.68 (ภาค1)' },     // สิงหาคม2568
    { y: 2025, m: 10, label: 'พ.ย.68 (ภาค2)' },    // พฤศจิกายน2568
    { y: 2025, m: 11, label: 'ธ.ค.68 (ภาค2)' },    // ธันวาคม2568
    { y: 2026, m: 0, label: 'ม.ค.69 (ภาค2)' },     // มกราคม2569
    { y: 2026, m: 2, label: 'มี.ค.69 (ภาค2)' }     // มีนาคม2569
  ];

  monthsToCheck.forEach(function(mc) {
    var sheetName = thaiMonths[mc.m] + (mc.y + 543);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('❌ ' + mc.label + ' (' + sheetName + '): ไม่พบชีต');
      return;
    }
    var data = sheet.getDataRange().getValues();
    var totalRows = data.length - 1;
    var matchCount = 0;
    var matchedIds = [];
    var allSheetIds = [];
    var classSummary = {};

    for (var i = 1; i < data.length; i++) {
      var rawId = String(data[i][1] || '').trim().match(/^(\d+)/);
      var sheetId = rawId ? rawId[1] : '';
      if (!sheetId) continue;
      allSheetIds.push(sheetId);
      if (classMatchIds.has(sheetId)) {
        matchCount++;
        matchedIds.push(sheetId);
      }
      // นับจำนวนนร.แต่ละชั้นจาก Students master
      var sInfo = null;
      for (var j = 1; j < sData.length; j++) {
        if (String(sData[j][0] || '').trim() === sheetId) {
          sInfo = String(sData[j][5] || '').trim() + '/' + String(sData[j][6] || '').trim();
          break;
        }
      }
      var key = sInfo || 'ไม่พบใน Students';
      classSummary[key] = (classSummary[key] || 0) + 1;
    }

    Logger.log('✅ ' + mc.label + ' (' + sheetName + '): ' + totalRows + ' นร., ป.1/1 match=' + matchCount + ' IDs=' + matchedIds.join(','));
    // แสดงจำนวนนร.แต่ละชั้นที่อยู่ในชีตนี้
    var summaryParts = [];
    Object.keys(classSummary).sort().forEach(function(k) {
      summaryParts.push(k + ':' + classSummary[k]);
    });
    Logger.log('   ชั้นในชีต: ' + summaryParts.join(', '));
  });

  Logger.log('=== END DEBUG v4 ===');
}

// ============================================================
// 🔧 CORE FUNCTIONS - อัปเดตให้รองรับระบบปีการศึกษา
// ============================================================
// ⚠️ หมายเหตุ: SPREADSHEET_ID และ monthNames ได้ประกาศไว้ใน code.gs แล้ว

/**
 * 🆕 ฟังก์ชันหาปีการศึกษาจากปีค.ศ.และเดือนที่ส่งมา
 * @param {number} yearCE - ปี ค.ศ.
 * @param {number} month - เดือน (1-12)
 * @returns {number} ปีการศึกษา (พ.ศ.)
 */
function getAcademicYearFromDate(yearCE, month) {
  const yearBE = yearCE + 543;
  
  // ถ้าเดือนมกราคม-เมษายน = ปีการศึกษาปีก่อน
  // ถ้าเดือนพฤษภาคม-ธันวาคม = ปีการศึกษาปีนั้น
  if (month >= 1 && month <= 4) {
    return yearBE - 1; // ปีการศึกษาปีก่อน
  } else {
    return yearBE; // ปีการศึกษาปีนั้น
  }
}

/**
 * 🆕 ตรวจสอบว่าเดือน/ปีอยู่ในปีการศึกษาที่กำหนดหรือไม่
 * @param {number} yearCE - ปี ค.ศ.
 * @param {number} month - เดือน (1-12)  
 * @param {number} targetAcademicYear - ปีการศึกษาเป้าหมาย (พ.ศ.)
 * @returns {boolean}
 */
function isInAcademicYear(yearCE, month, targetAcademicYear) {
  const academicYear = getAcademicYearFromDate(yearCE, month);
  return academicYear === targetAcademicYear;
}

/**
 * 🆕 แปลงปีการศึกษาเป็นช่วงเดือน/ปีที่ครอบคลุม
 * @param {number} academicYearBE - ปีการศึกษา (พ.ศ.)
 * @returns {Array} รายการ {month, yearCE} ในปีการศึกษา
 */
function getMonthsInAcademicYear(academicYearBE) {
  const baseYearCE = academicYearBE - 543;
  
  return [
    // ภาคเรียนที่ 1: พฤษภาคม-ตุลาคม (ปีฐาน)
    { month: 5, yearCE: baseYearCE },
    { month: 6, yearCE: baseYearCE },
    { month: 7, yearCE: baseYearCE },
    { month: 8, yearCE: baseYearCE },
    { month: 9, yearCE: baseYearCE },
    { month: 10, yearCE: baseYearCE },
    // ภาคเรียนที่ 2: พฤศจิกายน-ธันวาคม (ปีฐาน)
    { month: 11, yearCE: baseYearCE },
    { month: 12, yearCE: baseYearCE },
    // ภาคเรียนที่ 2 (ต่อ): มกราคม-มีนาคม (ปีถัดไป)
    { month: 1, yearCE: baseYearCE + 1 },
    { month: 2, yearCE: baseYearCE + 1 },
    { month: 3, yearCE: baseYearCE + 1 }
  ];
}

function setupNewAcademicYear() {
  try {
    var yearRaw = '';
    try {
      if (typeof S_getAcademicYear === 'function') {
        yearRaw = S_getAcademicYear();
      }
    } catch (_e) {}

    if (!yearRaw) {
      try {
        if (typeof S_getGlobalSettings === 'function') {
          var s = S_getGlobalSettings(true) || {};
          yearRaw = s['ปีการศึกษา'] || s['academicYear'] || '';
        }
      } catch (_e2) {}
    }

    var academicYearBE = parseInt(String(yearRaw || '').replace(/[^0-9]/g, ''), 10);
    if (!academicYearBE || isNaN(academicYearBE)) {
      throw new Error('ไม่พบปีการศึกษาใน global_settings');
    }

    var ss = SS();
    var months = getMonthsInAcademicYear(academicYearBE);
    var monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    var thaiWeekdaysShort = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];

    var created = [];
    var skipped = [];

    months.forEach(function(info) {
      var sheetName = monthNames[info.month - 1] + String(info.yearCE + 543);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        skipped.push(sheetName);
        return;
      }

      sheet = ss.insertSheet(sheetName);
      var daysInMonth = new Date(info.yearCE, info.month, 0).getDate();

      var headers = ["เลขที่", "รหัส", "ชื่อ-สกุล"];
      for (var d = 1; d <= daysInMonth; d++) {
        var dt = new Date(info.yearCE, info.month - 1, d);
        headers.push(String(d) + ' ' + thaiWeekdaysShort[dt.getDay()]);
      }
      headers.push("มา", "ลา", "ขาด");

      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      created.push(sheetName);
    });

    return [
      'ปีการศึกษา: ' + academicYearBE,
      'สร้างใหม่: ' + created.length + ' ชีต',
      created.length ? ('- ' + created.join(', ')) : '',
      'มีอยู่แล้ว: ' + skipped.length + ' ชีต'
    ].filter(function(x){ return x; }).join('\n');
  } catch (e) {
    throw new Error('setupNewAcademicYear: ' + (e.message || String(e)));
  }
}

/**
 * ✅ ดึงรายชื่อนักเรียนสำหรับเช็คชื่อ (ใช้โค้ดเดิมที่ทำงานได้)
 */
function getStudentsForAttendance(grade, classNo) {
  const sheet = SS().getSheetByName("Students");
  if (!sheet) throw new Error('ไม่พบชีต "Students" กรุณาตรวจสอบว่ามีชีตนักเรียนในระบบ');
  const data = sheet.getDataRange().getValues();
  const result = [];

  Logger.log(`🔍 getStudentsForAttendance: ${grade} ห้อง ${classNo}`);

  // หา index ของ status column
  const headers = data[0];
  const statusCol = headers.indexOf ? headers.indexOf('status') : -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // กรองนักเรียนที่จำหน่าย/ย้ายออก/พ้นสภาพ
    if (statusCol !== -1) {
      const st = String(row[statusCol] || '').trim();
      if (st === 'จำหน่าย' || st === 'ย้ายออก' || st === 'พ้นสภาพ') continue;
    }

    if (String(row[5]).trim() === grade && String(row[6]).trim() === classNo) {
      const student = {
        id: String(row[0]).trim(),          // student_id
        title: row[2] || '',          // title
        firstname: row[3] || '',          // firstname
        lastname: row[4] || '',          // lastname
        grade: String(row[5]).trim(),          // grade
        classNo: String(row[6]).trim()          // class_no
      };
      
      result.push(student);
    }
  }

  // --- 👇 เพิ่มโค้ดเรียงลำดับนักเรียนตามรหัส (id) ตรงนี้ ---
  result.sort((a, b) => {
    const idA = a.id || '';
    const idB = b.id || '';
    return String(idA).localeCompare(String(idB), undefined, {numeric: true});
});
  // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

  Logger.log(`✅ พบนักเรียน ${result.length} คน (เรียงลำดับแล้ว)`);
  return result;
}


/**
 * ดึงข้อมูลการเช็คชื่อที่บันทึกไว้แล้ว - แก้ไขให้อ่านข้อมูลถูกต้อง
 * @param {string} grade - ระดับชั้น
 * @param {string} classNo - ห้อง
 * @param {number} year - ปี ค.ศ.
 * @param {number} month - เดือน (0-indexed: 0=ม.ค., 4=พ.ค., 11=ธ.ค.) ⚠️ ต่างจาก getMonthlyAttendanceSummary ที่ใช้ 1-indexed
 * @returns {Object} ข้อมูลการเช็คชื่อที่บันทึกไว้
 */
function getSavedAttendance(grade, classNo, year, month) {
  try {
    Logger.log(`getSavedAttendance called with: grade=${grade}, classNo=${classNo}, year=${year}, month=${month}`);
    
    // ตรวจสอบพารามิเตอร์
    if (!grade || !classNo || year === undefined || month === undefined) {
      Logger.log('Invalid parameters:', {grade, classNo, year, month});
      return { records: {} };
    }
    
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const ss = SS();
    const sheetName = `${thaiMonths[month]}${year + 543}`;
    
    Logger.log(`Looking for sheet: ${sheetName}`);
    
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`ไม่พบชีต: ${sheetName}`);
      return { records: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log(`ไม่มีข้อมูลในชีต: ${sheetName}`);
      return { records: {} };
    }
    
    const headers = data[0];
    Logger.log(`Headers found: ${headers.length} columns`);
    Logger.log(`First few headers: ${headers.slice(0, 10)}`);

    // หาวันที่เป็นคอลัมน์วันที่ (มักจะเริ่มจากคอลัมน์ที่ 3 ขึ้นไป)
    const dateCols = [];
    for (let i = 3; i < headers.length; i++) {
      const h = String(headers[i]).trim();
      // ดึงตัวเลขวันที่จาก header เช่น "1 อา", "2 จ", "10 พ"
      const dayMatch = h.match(/^(\d{1,2})/);
      if (dayMatch) {
        const day = parseInt(dayMatch[1]);
        dateCols.push({ col: i, day: day });
      }
    }
    
    Logger.log(`Found ${dateCols.length} date columns:`, dateCols.slice(0, 5));
    
    // ✅ อ่านชีต Students ทั้งหมด (รวมนร.จำหน่าย/ย้ายออก) เพื่อสร้าง ID → class map
    const studentsSheet = ss.getSheetByName("Students");
    const allStudentIdToClass = {}; // map: studentId → {grade, classNo}
    if (studentsSheet) {
      const sData = studentsSheet.getDataRange().getValues();
      for (let i = 1; i < sData.length; i++) {
        const row = sData[i];
        const sid = String(row[0] || '').trim();
        const g = String(row[5] || '').trim();
        const c = String(row[6] || '').trim();
        if (sid && g && c) allStudentIdToClass[sid] = { grade: g, classNo: c };
      }
    }
    Logger.log(`All students map (incl. inactive): ${Object.keys(allStudentIdToClass).length} entries`);

    // ✅ หา IDs ที่อยู่ในห้องที่เลือก (จาก master ทั้งหมด)
    const classMatchIds = new Set();
    Object.keys(allStudentIdToClass).forEach(sid => {
      const info = allStudentIdToClass[sid];
      if (info.grade === grade && info.classNo === classNo) {
        classMatchIds.add(sid);
      }
    });
    Logger.log(`IDs in ${grade}/${classNo} from full master: ${classMatchIds.size} (sample: ${Array.from(classMatchIds).slice(0,5)})`);

    // ✅ อ่าน records เฉพาะนักเรียนในห้องที่เลือก
    const records = {};
    const sheetStudents = []; // นักเรียนจากชีตเช็คชื่อที่ตรงห้อง
    let matchedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const match = String(row[1] || '').trim().match(/^(\d+)/);
      const sheetId = match ? match[1] : '';
      if (!sheetId) continue;

      // ตรวจสอบว่า sheetId อยู่ในห้องที่เลือกหรือไม่
      if (!classMatchIds.has(sheetId)) continue;
      matchedCount++;
      sheetStudents.push({ id: sheetId, name: String(row[2] || '').trim() });

      dateCols.forEach(dateCol => {
        const cellValue = String(row[dateCol.col] || '').trim();
        const key = `${sheetId}_${dateCol.day}`;
        
        let normalizedStatus = '';
        if (cellValue === 'ข' || cellValue === '0') {
          normalizedStatus = 'absent';
        } else if (cellValue === 'ล' || cellValue.toLowerCase() === 'l') {
          normalizedStatus = 'leave';
        } else if (cellValue === '/' || cellValue === 'ม') {
          normalizedStatus = 'present';
        }
        
        if (normalizedStatus) {
          records[key] = normalizedStatus;
        }
      });
    }
    
    Logger.log(`Matched ${matchedCount} sheet students to ${grade}/${classNo}, records: ${Object.keys(records).length}`);

    // ดึงข้อมูลวันหยุดแนบไปด้วย
    let holidays = {};
    try {
      const holidayArr = loadHolidaysFromSheet();
      if (Array.isArray(holidayArr)) {
        holidayArr.forEach(function(row) {
          if (row && row[0] && row[1]) {
            holidays[String(row[0]).trim()] = String(row[1]).trim();
          }
        });
      }
      Logger.log(`📅 Holidays attached: ${Object.keys(holidays).length} items`);
    } catch (he) {
      Logger.log('⚠️ Could not load holidays: ' + he.message);
    }
    
    return { records: records, holidays: holidays, sheetStudents: sheetStudents };
    
  } catch (e) {
    Logger.log("Error in getSavedAttendance: " + e.message);
    Logger.log("Stack trace:", e.stack);
    return { records: {} };
  }
}



// ============================================================
// 🔧 อัปเดตฟังก์ชันสรุปรายเดือน
// ============================================================

/**
 * ✅ [แก้ไขใหม่] ดึงข้อมูลสรุปการมาเรียนรายเดือน (เร็วและแม่นยำขึ้น)
 * @param {string} grade ระดับชั้น
 * @param {string} classNo ห้อง
 * @param {number} yearCE ปี ค.ศ.
 * @param {number} month เดือน (1-12)
 * @returns {Array} ข้อมูลสรุปของนักเรียนในเดือนนั้น
 */
function getMonthlyAttendanceSummary(grade, classNo, yearCE, month) {
  try {
    const ss = SS();
    
    // 1. ดึงรายชื่อนักเรียนที่ต้องการทั้งหมด "ครั้งเดียว"
    const studentsInClass = getStudentsForAttendance(grade, classNo);
    if (!studentsInClass || studentsInClass.length === 0) {
      Logger.log(`ไม่พบนักเรียนในชั้น ${grade}/${classNo}`);
      return [];
    }
    
    // 2. สร้าง Map เพื่อเก็บข้อมูลและค้นหาได้รวดเร็ว
    const studentMap = new Map();
    studentsInClass.forEach(stu => {
      studentMap.set(stu.id, {
        studentId: stu.id,
        name: `${stu.title || ''}${stu.firstname} ${stu.lastname}`.trim(),
        present: 0,
        leave: 0,
        absent: 0,
        total: 0
      });
    });

    // 3. เปิดชีตของเดือนที่ต้องการ
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const sheetName = `${monthNames[month - 1]}${yearCE + 543}`;
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`ไม่พบชีต: ${sheetName}, คืนค่าว่าง`);
      return []; // ถ้าไม่มีชีต ก็ไม่มีข้อมูล
    }

    // 4. อ่านข้อมูลจากชีต "ครั้งเดียว" ทั้งหมด
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // ไม่มีข้อมูลนักเรียน

    const headers = data[0];
    const idColIdx = headers.findIndex(h => String(h).includes("รหัส"));
    const presentColIdx = headers.findIndex(h => String(h).startsWith("มา"));
    const leaveColIdx = headers.findIndex(h => String(h).startsWith("ลา"));
    const absentColIdx = headers.findIndex(h => String(h).startsWith("ขาด"));
    
    // ตรวจสอบว่ามีคอลัมน์สรุปหรือไม่
    if ([idColIdx, presentColIdx, leaveColIdx, absentColIdx].includes(-1)) {
        Logger.log(`ชีต ${sheetName} ไม่มีคอลัมน์สรุป (มา/ลา/ขาด)`);
        return []; // ถ้าไม่มีคอลัมน์สรุป จะไม่สามารถรวมยอดได้
    }

    // 5. วนลูปประมวลผลข้อมูลในหน่วยความจำ (เร็วมาก)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentIdMatch = String(row[idColIdx] || '').match(/^(\d+)/);
      const studentId = studentIdMatch ? studentIdMatch[1] : null;

      // ถ้าเป็นนักเรียนในห้องที่ต้องการ ให้บวกค่า
      if (studentId && studentMap.has(studentId)) {
        const studentData = studentMap.get(studentId);
        studentData.present += Number(row[presentColIdx]) || 0;
        studentData.leave += Number(row[leaveColIdx]) || 0;
        studentData.absent += Number(row[absentColIdx]) || 0;
      }
    }

    // 6. แปลง Map เป็น Array และคำนวณผลรวม
    const result = Array.from(studentMap.values()).map(stu => {
        stu.total = stu.present + stu.leave + stu.absent;
        return stu;
    }).filter(stu => stu.total > 0); // กรองเฉพาะคนที่มีข้อมูล

    // --- 👇 เพิ่มโค้ดเรียงลำดับตรงนี้ ก่อน return ---
    result.sort((a, b) => String(a.studentId).localeCompare(String(b.studentId), undefined, {numeric: true}));
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

    Logger.log(`สรุปข้อมูลเดือน ${month}/${yearCE} สำเร็จ, พบนักเรียน ${result.length} คน`);
    return result;

  } catch (e) {
    Logger.log(`Error in getMonthlyAttendanceSummary: ${e.message}`);
    throw e;
  }
}


/**
 * 🔧 ฟังก์ชันสำรอง: ดึงข้อมูลจากคอลัมน์สรุป (มา/ลา/ขาด)
 */
function getMonthlyFromSummaryColumns(data, headers, grade, classNo) {
  const idxId = headers.findIndex(h => String(h).includes("รหัส"));
  const idxName = headers.findIndex(h => String(h).includes("ชื่อ"));
  const idxPresent = headers.findIndex(h => String(h).startsWith("มา"));
  const idxLeave = headers.findIndex(h => String(h).startsWith("ลา"));
  const idxAbsent = headers.findIndex(h => String(h).startsWith("ขาด"));

  if ([idxId, idxPresent, idxLeave, idxAbsent].includes(-1)) {
    Logger.log("ไม่พบคอลัมน์สรุปที่จำเป็น");
    return [];
  }

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentId = String(row[idxId]).trim();
    
    if (!isStudentInClass(studentId, grade, classNo)) continue;

    const name = idxName >= 0 ? row[idxName] : '-';
    const present = Number(row[idxPresent]) || 0;
    const leave = Number(row[idxLeave]) || 0;
    const absent = Number(row[idxAbsent]) || 0;
    const total = present + leave + absent;

    if (total > 0) {
      result.push({
        studentId,
        name,
        present,
        leave,
        absent,
        total
      });
    }
  }
  return result;
}

/**
 * 🔧 ตรวจสอบว่านักเรียนอยู่ในชั้น/ห้องที่ต้องการหรือไม่
 */
function isStudentInClass(studentId, targetGrade, targetClassNo) {
  try {
    const ss = SS();
    const studentsSheet = ss.getSheetByName('Students');
    
    if (!studentsSheet) return false;
    
    const data = studentsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf("student_id");
    const gradeCol = headers.indexOf("grade");
    const classCol = headers.indexOf("class_no");
    
    if ([idCol, gradeCol, classCol].includes(-1)) return false;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[idCol]).trim() === String(studentId).trim()) {
        const grade = String(row[gradeCol]).trim();
        const classNo = String(row[classCol]).trim();
        return grade === targetGrade && classNo === targetClassNo;
      }
    }
    
    return false;
  } catch (error) {
    Logger.log('Error in isStudentInClass:', error);
    return false;
  }
}


/**
 * บันทึก log การเช็คชื่อ
 * @param {string} date - วันที่
 * @param {string} grade - ระดับชั้น
 * @param {string} classNo - ห้อง
 * @param {Object} records - ข้อมูลการเช็คชื่อ
 * @param {number} updatedCount - จำนวนที่อัพเดท
 */
function saveAttendanceLog(date, grade, classNo, records, updatedCount) {
  try {
    const ss = SS();
    let logSheet = S_getYearlySheet('AttendanceLog');
    
    if (!logSheet) {
      logSheet = ss.insertSheet(S_sheetName('AttendanceLog'));
      const headers = ["timestamp", "date", "grade", "class", "updated_count", "user_email", "details"];
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // จัดรูปแบบหัวตาราง
      const headerRange = logSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#34495e");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
    }
    
    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || "anonymous";
    const details = JSON.stringify(records);
    
    const logData = [timestamp, date, grade, classNo, updatedCount, userEmail, details];
    logSheet.appendRow(logData);
    
  } catch (error) {
    Logger.log("Error in saveAttendanceLog: " + error.message);
  }
}

/**
 * ตารางเช็คชื่อ รายเดือน (ใช้กับชีทที่มีอยู่แล้ว)
 * (ฉบับสมบูรณ์: มีการเรียงลำดับนักเรียนตามรหัส)
 */
// ฟังก์ชันดึงข้อมูลวันหยุด
function getHolidaysMap() {
  try {
    const ss = SS();
    const holidaySheet = ss.getSheetByName('Holidays') || ss.getSheetByName('วันหยุด');
    if (!holidaySheet) {
      Logger.log('❌ ไม่พบชีต "Holidays" หรือ "วันหยุด"');
      return {};
    }
    
    const data = holidaySheet.getDataRange().getValues();
    const holidays = {};
    const dateCol = 0;
    const typeCol = 1;
    
    for (let i = 1; i < data.length; i++) {
      const rawDate = data[i][dateCol];
      if (!rawDate) continue;
      
      let year, month, day;
      
      if (rawDate instanceof Date) {
        // Date object จาก Google Sheets
        year = rawDate.getFullYear();
        month = rawDate.getMonth() + 1;
        day = rawDate.getDate();
      } else {
        // Text date เช่น "2025-08-11" หรือ "2568-08-11"
        const dateStr = String(rawDate).trim();
        const match = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (match) {
          year = parseInt(match[1]);
          month = parseInt(match[2]);
          day = parseInt(match[3]);
          // ถ้าเป็น พ.ศ. (ปี > 2500) แปลงเป็น ค.ศ.
          if (year > 2500) year -= 543;
        } else {
          continue;
        }
      }
      
      const type = String(data[i][typeCol] || '').trim();
      const ceDateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      holidays[ceDateStr] = type || 'วันหยุด';
    }
    
    Logger.log(`📊 Holidays loaded: ${Object.keys(holidays).length}, keys: ${JSON.stringify(Object.keys(holidays))}`);
    return holidays;
  } catch (e) {
    Logger.log('Error loading holidays: ' + e.message);
    return {};
  }
}

function getAttendanceVerticalTable(grade, classNo, year, month) {
  try {
    const ss = SS();
    const sheetName = `${monthNames[month]}${year + 543}`;
    const monthSheet = ss.getSheetByName(sheetName);
    if (!monthSheet) return { students: [], days: [] };

    // ดึงข้อมูลวันหยุด
    const holidaysMap = getHolidaysMap();

    const mData = monthSheet.getDataRange().getValues();
    const mHead = mData[0];
    const idCol = 1;     // B
    const nameCol = 2;   // C

    const dateCols = mHead.map((h, i) => {
      const match = String(h).match(/^(\d{1,2})/);
      return match ? { col: i, label: match[1] } : null;
    }).filter(Boolean);

    Logger.log(`📋 holidaysMap keys: ${JSON.stringify(Object.keys(holidaysMap))}`);
    Logger.log(`📋 year=${year}, month=${month} (0-indexed)`);
    
    const days = dateCols.map(dc => {
      const d = parseInt(dc.label);
      const date = new Date(year, month, d);
      const dow = ["อา.", "จ.", "อ.", "พ.", "พฤ.", "ศ.", "ส."][date.getDay()];
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const isHoliday = holidaysMap[dateStr] ? true : false;
      const holidayType = holidaysMap[dateStr] || '';
      if (d >= 10 && d <= 13) Logger.log(`🔍 day ${d}: dateStr=${dateStr}, isHoliday=${isHoliday}`);
      return { label: `${d}`, dow, date: dateStr, isHoliday, holidayType };
    });

    const studentsSheet = ss.getSheetByName("Students");
    const studentsData = studentsSheet.getDataRange().getValues();
    const stuHeaders = studentsData[0];
    const stuIdCol = stuHeaders.indexOf("student_id");
    const stuGradeCol = stuHeaders.indexOf("grade");
    const stuClassCol = stuHeaders.indexOf("class_no");

    const idGradeMap = {};
    studentsData.slice(1).forEach(row => {
      const id = String(row[stuIdCol]).trim();
      idGradeMap[id] = {
        grade: String(row[stuGradeCol]).trim(),
        classNo: String(row[stuClassCol]).trim()
      };
    });

    const students = mData.slice(1)
      .map(row => {
        const idCell = String(row[idCol]).trim();
        const name = String(row[nameCol]).trim();
        
        const idMatch = idCell.match(/^(\d+)/);
        const id = idMatch ? idMatch[1] : '';
        
        const info = idGradeMap[id];
        if (!info) return null;
        if (info.grade !== grade || info.classNo !== classNo) return null;

        const statusMap = {};
        dateCols.forEach(dc => {
          statusMap[dc.label] = row[dc.col] || '';
        });
        return { id, name, statusMap };
      })
      .filter(stu => stu && stu.id && stu.name);

    // --- 👇 เรียงลำดับนักเรียนตามรหัส (id) ---
    students.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
    // --- 👆 สิ้นสุดการเรียงลำดับ ---

    return { students, days };

  } catch (e) {
    Logger.log("Error in getAttendanceVerticalTable: " + e.message);
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลตาราง: " + e.message);
  }
}

/**
 * ฟังก์ชันบันทึกข้อมูลเช็คชื่อทั้งเดือน (payload: {grade, classNo, year, month, records})
 * records: { "studentId_day": "present"/"leave"/"absent" }
 */
// ✅ [เวอร์ชันแก้ไข] วางโค้ดนี้แทนที่ฟังก์ชัน saveMonthlyAttendance เดิมทั้งหมด
function saveMonthlyAttendance(payload) {
  if (!payload || !payload.grade || !payload.classNo || payload.year === undefined || payload.month === undefined) {
    throw new Error("ข้อมูล Payload ไม่ครบถ้วน (grade, classNo, year, month)");
  }

  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      throw new Error('ระบบกำลังบันทึกข้อมูลอยู่ กรุณารอสักครู่แล้วลองใหม่');
    }

    const ss = SS();
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthName = `${monthNames[payload.month]}${payload.year + 543}`;
    let sheet = ss.getSheetByName(monthName);

    if (!sheet) {
      sheet = ss.insertSheet(monthName);
      const daysInMonth = new Date(payload.year, payload.month + 1, 0).getDate();
      const hdr = ["เลขที่", "รหัส", "ชื่อ-สกุล"];
      const thaiWeekdaysShort = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
      for (let i = 1; i <= daysInMonth; i++) {
        const d = new Date(payload.year, payload.month, i);
        hdr.push(`${i} ${thaiWeekdaysShort[d.getDay()]}`);
      }
      hdr.push("มา", "ลา", "ขาด");
      sheet.appendRow(hdr);
      sheet.getRange(1, 1, 1, hdr.length).setFontWeight("bold");
    }

    // อ่าน header + ข้อมูลทั้งหมด 1 ครั้ง
    const sheetData = sheet.getDataRange().getValues();
    const headerRow = sheetData[0];
    const numCols = headerRow.length;

    // Map วัน → colIndex (0-based) + หาขอบเขตคอลัมน์วันที่
    const dayToCol = {};
    let firstDayCol = -1, lastDayCol = -1;
    for (let i = 2; i < headerRow.length; i++) {
      const h = String(headerRow[i]).match(/^(\d+)/);
      if (h) {
        dayToCol[parseInt(h[1])] = i; // 0-based index
        if (firstDayCol === -1) firstDayCol = i;
        lastDayCol = i;
      }
    }

    // หาคอลัมน์สรุป มา/ลา/ขาด
    let idxPresent = -1, idxLeave = -1, idxAbsent = -1;
    for (let i = 0; i < headerRow.length; i++) {
      const ht = String(headerRow[i] || '').trim();
      if ((ht === 'มา' || ht.startsWith('มา')) && idxPresent === -1) idxPresent = i;
      else if ((ht === 'ลา' || ht.startsWith('ลา')) && idxLeave === -1) idxLeave = i;
      else if ((ht === 'ขาด' || ht.startsWith('ขาด')) && idxAbsent === -1) idxAbsent = i;
    }

    // Map studentId → index ใน sheetData (0-based ใน sheetData)
    const studentRowMap = new Map();
    sheetData.slice(1).forEach((row, i) => {
      const match = String(row[1] || '').trim().match(/^(\d+)/);
      if (match) studentRowMap.set(match[1], i + 1); // index ใน sheetData (row 0 = header)
    });

    const studentsFromMaster = getStudentsForAttendance(payload.grade, payload.classNo);
    const statusMap = { 'present': '/', 'absent': 'ข', 'leave': 'ล' };
    const records = payload.records || {};
    let updateCount = 0;

    // แถวใหม่ที่ต้อง append (นักเรียนที่ยังไม่มีในชีต)
    const newRows = [];

    studentsFromMaster.forEach(student => {
      const studentId = String(student.id).trim();
      let dataIdx = studentRowMap.get(studentId); // index ใน sheetData

      if (dataIdx == null) {
        // นักเรียนใหม่ — สร้างแถวใหม่
        const newRow = new Array(numCols).fill('');
        newRow[1] = studentId;
        newRow[2] = `${student.title || ''}${student.firstname} ${student.lastname}`.trim();
        sheetData.push(newRow);
        dataIdx = sheetData.length - 1; // index ของแถวที่เพิ่งเพิ่ม
        newRows.push(newRow);
        studentRowMap.set(studentId, dataIdx);
      }

      // ใช้ in-memory sheetData แทน getValue/setValue ทีละ cell
      Object.keys(records).forEach(key => {
        const [rid, day] = key.split('_');
        if (rid !== studentId) return;
        const status = records[key];
        const colIndex = dayToCol[parseInt(day)]; // 0-based
        if (colIndex === undefined || !status) return;
        const symbol = statusMap[status];
        if (!symbol) return;
        if (sheetData[dataIdx][colIndex] !== symbol) {
          sheetData[dataIdx][colIndex] = symbol;
          updateCount++;
        }
      });
    });

    // ✅ เขียนสูตร COUNTIF กลับทุกแถว (ป้องกันสูตรหาย)
    if (firstDayCol > -1 && lastDayCol > -1 && idxPresent > -1 && idxLeave > -1 && idxAbsent > -1) {
      const startCol = columnNumberToLetter(firstDayCol + 1);
      const endCol = columnNumberToLetter(lastDayCol + 1);
      for (let i = 1; i < sheetData.length; i++) {
        const rowNum = i + 1;
        const rangeForRow = `${startCol}${rowNum}:${endCol}${rowNum}`;
        sheetData[i][idxPresent] = `=COUNTIF(${rangeForRow},"/")`;
        sheetData[i][idxLeave] = `=COUNTIF(${rangeForRow},"ล")`;
        sheetData[i][idxAbsent] = `=COUNTIF(${rangeForRow},"ข")`;
      }
    }

    // === Batch write ===
    const existingRows = sheetData.length - newRows.length;
    if (existingRows > 1) {
      sheet.getRange(2, 1, existingRows - 1, numCols)
           .setValues(sheetData.slice(1, existingRows));
    }
    if (newRows.length > 0) {
      if (sheetData.length > sheet.getLastRow()) {
        const rowsToAdd = sheetData.length - sheet.getLastRow();
        sheet.insertRows(sheet.getLastRow() + 1, rowsToAdd);
      }
      const startRow = existingRows + 1;
      sheet.getRange(startRow, 1, newRows.length, numCols).setValues(newRows);
    }

    SpreadsheetApp.flush();

    Logger.log(`saveMonthlyAttendance เสร็จ: อัปเดต ${updateCount} ช่อง, เพิ่มใหม่ ${newRows.length} แถว`);
    return { message: `บันทึกข้อมูลสำเร็จ (อัปเดต ${updateCount} ช่อง)` };

  } catch (e) {
    Logger.log('saveMonthlyAttendance error: ' + e.message + '\n' + e.stack);
    throw new Error('บันทึกข้อมูลเช็คชื่อไม่สำเร็จ: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}


// ✅ เพิ่มฟังก์ชันแปลงเลขคอลัมน์เป็นตัวอักษร
function columnNumberToLetter(columnNumber) {
  let dividend = columnNumber;
  let columnName = '';
  let modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

// ✅ ฟังก์ชันหลักที่แก้ไขตามโครงสร้างจริง
function saveMonthlyAttendanceOptimized(payload) {
  try {
    Logger.log("🚀 [Optimized] เริ่มกระบวนการบันทึก");
    const { grade, classNo, year, month, records } = payload;
    
    const ss = SS();
    const MONTH_NAMES_TH = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
const monthName = `${MONTH_NAMES_TH[month]}${year + 543}`;
    let sheet = ss.getSheetByName(monthName);

    // ✅ สร้างชีตใหม่ถ้ายังไม่มี (แทนที่จะ throw error)
    if (!sheet) {
      Logger.log(`📄 ไม่พบชีต ${monthName} — กำลังสร้างใหม่...`);
      sheet = ss.insertSheet(monthName);
      const daysInMonthForHeader = new Date(year, month + 1, 0).getDate();
      const thaiWeekdaysShort = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
      const hdr = ["เลขที่", "รหัส", "ชื่อ-สกุล"];
      for (let d = 1; d <= daysInMonthForHeader; d++) {
        const dt = new Date(year, month, d);
        hdr.push(`${d} ${thaiWeekdaysShort[dt.getDay()]}`);
      }
      hdr.push("มา", "ลา", "ขาด");
      sheet.appendRow(hdr);
      sheet.getRange(1, 1, 1, hdr.length).setFontWeight("bold");
      Logger.log(`✅ สร้างชีต ${monthName} สำเร็จ (${hdr.length} คอลัมน์)`);
    }

    const studentsFromMaster = getStudentsForAttendance(grade, classNo);
    if (studentsFromMaster.length === 0) {
      return { message: "ไม่พบรายชื่อนักเรียนสำหรับชั้นเรียนนี้" };
    }

    const dataRange = sheet.getDataRange();
    let sheetData = dataRange.getValues();
    const headers = sheetData[0];
    
    Logger.log(`📋 Headers (first 20): ${JSON.stringify(headers.slice(0, 20))}`);
    Logger.log(`📋 Headers (last 10): ${JSON.stringify(headers.slice(-10))}`);
    
    const studentRowMap = new Map();
    sheetData.slice(1).forEach((row, i) => {
      const studentId = String(row[1] || '').match(/^(\d+)/);
      if (studentId) studentRowMap.set(studentId[1], i + 1);
    });

    let newStudentsAdded = false;
    // ตรวจสอบและเพิ่มนักเรียนที่ยังไม่มีในชีต
    studentsFromMaster.forEach((student, index) => {
      const studentId = String(student.id).trim();
      if (!studentRowMap.has(studentId)) {
        newStudentsAdded = true;
        const newRow = new Array(headers.length).fill('');
        newRow[0] = studentRowMap.size + 1;
        newRow[1] = studentId;
        newRow[2] = `${student.title || ''}${student.firstname} ${student.lastname}`.trim();
        sheetData.push(newRow);
        studentRowMap.set(studentId, sheetData.length - 1);
        Logger.log(`➕ เพิ่มนักเรียนใหม่: ${newRow[2]}`);
      }
    });

    // ✅ คำนวณจำนวนวันจริงในเดือนนี้
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    Logger.log(`📅 เดือนนี้มี ${daysInMonth} วัน (เดือน ${month + 1}/${year})`);
    
    // ✅ หาตำแหน่งคอลัมน์วันที่ (ตามจำนวนวันจริง)
    const dayToCol = {};
    let firstDayCol = -1, lastDayCol = -1;
    
    // หาคอลัมน์วันที่ตามจำนวนวันจริงในเดือน
    for (let day = 1; day <= daysInMonth; day++) {
      const colIndex = 3 + (day - 1); // คอลัมน์ D (index 3) เป็นวันที่ 1
      dayToCol[day] = colIndex;
      if (firstDayCol === -1) firstDayCol = colIndex;
      lastDayCol = colIndex;
      Logger.log(`📅 วันที่ ${day} = คอลัมน์ ${colIndex} (${columnNumberToLetter(colIndex + 1)})`);
    }

    // ✅ หาคอลัมน์สถิติ (คำนวณตำแหน่งจากจำนวนวันจริง)
    const expectedPresentCol = 3 + daysInMonth;     // คอลัมน์หลังวันสุดท้าย
    const expectedLeaveCol = expectedPresentCol + 1;
    const expectedAbsentCol = expectedLeaveCol + 1;
    
    let idxPresent = -1, idxLeave = -1, idxAbsent = -1;
    
    Logger.log(`🎯 คาดว่าคอลัมน์สถิติอยู่ที่: มา=${expectedPresentCol}(${columnNumberToLetter(expectedPresentCol + 1)}) ลา=${expectedLeaveCol}(${columnNumberToLetter(expectedLeaveCol + 1)}) ขาด=${expectedAbsentCol}(${columnNumberToLetter(expectedAbsentCol + 1)})`);
    
    // ตรวจสอบคอลัมน์ที่คาดไว้ก่อน
    if (expectedPresentCol < headers.length) {
      const presentHeader = String(headers[expectedPresentCol] || '').trim();
      if (presentHeader === 'มา' || presentHeader.includes('มา') || presentHeader.match(/^\d+$/)) {
        idxPresent = expectedPresentCol;
        Logger.log(`✅ พบคอลัมน์ "มา" ที่ตำแหน่งที่คาดไว้ ${expectedPresentCol} (${columnNumberToLetter(expectedPresentCol + 1)}): "${presentHeader}"`);
      }
    }
    
    if (expectedLeaveCol < headers.length) {
      const leaveHeader = String(headers[expectedLeaveCol] || '').trim();
      if (leaveHeader === 'ลา' || leaveHeader.includes('ลา') || leaveHeader.match(/^\d+$/)) {
        idxLeave = expectedLeaveCol;
        Logger.log(`✅ พบคอลัมน์ "ลา" ที่ตำแหน่งที่คาดไว้ ${expectedLeaveCol} (${columnNumberToLetter(expectedLeaveCol + 1)}): "${leaveHeader}"`);
      }
    }
    
    if (expectedAbsentCol < headers.length) {
      const absentHeader = String(headers[expectedAbsentCol] || '').trim();
      if (absentHeader === 'ขาด' || absentHeader.includes('ขาด') || absentHeader.match(/^\d+$/)) {
        idxAbsent = expectedAbsentCol;
        Logger.log(`✅ พบคอลัมน์ "ขาด" ที่ตำแหน่งที่คาดไว้ ${expectedAbsentCol} (${columnNumberToLetter(expectedAbsentCol + 1)}): "${absentHeader}"`);
      }
    }
    
    // ถ้าไม่พบที่ตำแหน่งที่คาดไว้ ให้ค้นหาจากท้ายชีต
    if (idxPresent === -1 || idxLeave === -1 || idxAbsent === -1) {
      Logger.log(`🔍 ไม่พบคอลัมน์สถิติที่ตำแหน่งที่คาดไว้ กำลังค้นหาจากท้ายชีต...`);
      
      for (let i = Math.max(0, expectedPresentCol - 3); i < headers.length; i++) {
        const headerText = String(headers[i] || '').trim();
        Logger.log(`🔍 ตรวจสอบคอลัมน์ ${i} (${columnNumberToLetter(i + 1)}): "${headerText}"`);
        
        if ((headerText === 'มา' || headerText.includes('มา')) && idxPresent === -1) {
          idxPresent = i;
          Logger.log(`✅ พบคอลัมน์ "มา" ที่ตำแหน่ง ${i} (${columnNumberToLetter(i + 1)})`);
        } else if ((headerText === 'ลา' || headerText.includes('ลา')) && idxLeave === -1) {
          idxLeave = i;
          Logger.log(`✅ พบคอลัมน์ "ลา" ที่ตำแหน่ง ${i} (${columnNumberToLetter(i + 1)})`);
        } else if ((headerText === 'ขาด' || headerText.includes('ขาด')) && idxAbsent === -1) {
          idxAbsent = i;
          Logger.log(`✅ พบคอลัมน์ "ขาด" ที่ตำแหน่ง ${i} (${columnNumberToLetter(i + 1)})`);
        }
      }
    }

    Logger.log(`📊 สรุปตำแหน่งคอลัมน์:`);
    Logger.log(`   วันที่: ${firstDayCol}-${lastDayCol} (${Object.keys(dayToCol).length} วัน)`);
    Logger.log(`   สถิติ: มา=${idxPresent} ลา=${idxLeave} ขาด=${idxAbsent}`);

    // ตรวจสอบว่าพบคอลัมน์ที่จำเป็นหรือไม่
    if (firstDayCol === -1) {
      Logger.log("❌ ไม่พบคอลัมน์วันที่");
      return { message: "ไม่พบคอลัมน์วันที่ในชีต" };
    }
    if (idxPresent === -1 || idxLeave === -1 || idxAbsent === -1) {
      Logger.log("❌ ไม่พบคอลัมน์สถิติครบ");
      return { message: "ไม่พบคอลัมน์สถิติ (มา/ลา/ขาด) ในชีต" };
    }

    // ✅ อัปเดตข้อมูลการเช็คชื่อ
    let updateCount = 0;
    Object.keys(records).forEach(key => {
      const [studentId, day] = key.split('_');
      const rowIndex = studentRowMap.get(studentId);
      const colIndex = dayToCol[parseInt(day)];
      
      if (rowIndex !== undefined && colIndex !== undefined) {
        const statusMap = { 'present': '/', 'absent': 'ข', 'leave': 'ล', 'none': '' };
        const newSymbol = statusMap[records[key]];
        const oldValue = sheetData[rowIndex][colIndex];
        
        if (oldValue !== newSymbol) {
          sheetData[rowIndex][colIndex] = newSymbol;
          updateCount++;
          Logger.log(`🔄 อัปเดต: นักเรียน ${studentId} วันที่ ${day} จาก "${oldValue}" เป็น "${newSymbol}"`);
        }
      } else {
        if (rowIndex === undefined) {
          Logger.log(`⚠️ ไม่พบนักเรียน ${studentId} ในชีต`);
        }
        if (colIndex === undefined) {
          Logger.log(`⚠️ ไม่พบคอลัมน์วันที่ ${day}`);
        }
      }
    });

    // ✅ อัปเดตสูตรสถิติทุกแถว
    let formulasUpdated = 0;
    if (idxPresent > -1 && idxLeave > -1 && idxAbsent > -1 && firstDayCol > -1 && lastDayCol > -1) {
      
      const startCol = columnNumberToLetter(firstDayCol + 1);
      const endCol = columnNumberToLetter(lastDayCol + 1);
      
      Logger.log(`🔄 เริ่มอัปเดตสูตรสถิติ: ${startCol}:${endCol}`);
      
      for (let i = 1; i < sheetData.length; i++) {
        const rowNum = i + 1;
        const rangeForRow = `${startCol}${rowNum}:${endCol}${rowNum}`;
        
        // ✅ ใส่สูตรใหม่ทุกแถว
        sheetData[i][idxPresent] = `=COUNTIF(${rangeForRow},"/")`;
        sheetData[i][idxLeave] = `=COUNTIF(${rangeForRow},"ล")`;
        sheetData[i][idxAbsent] = `=COUNTIF(${rangeForRow},"ข")`;
        
        formulasUpdated++;
        
        if (i <= 3) { // แสดง log เฉพาะ 3 แถวแรก
          Logger.log(`📝 แถว ${rowNum}: สูตร = COUNTIF(${rangeForRow},...)`);
        }
      }
      
      Logger.log(`✅ อัปเดตสูตรเสร็จ: ${formulasUpdated} แถว`);
    }

    // ✅ บันทึกข้อมูลกลับลงชีต
    if (updateCount > 0 || newStudentsAdded || formulasUpdated > 0) {
      try {
        Logger.log(`💾 กำลังบันทึกข้อมูล: ${sheetData.length} แถว x ${headers.length} คอลัมน์`);
        
        // ขยายขนาดชีตถ้าจำเป็น
        if (newStudentsAdded && sheetData.length > sheet.getLastRow()) {
          const rowsToAdd = sheetData.length - sheet.getLastRow();
          sheet.insertRows(sheet.getLastRow() + 1, rowsToAdd);
          Logger.log(`📏 เพิ่มแถวใหม่: ${rowsToAdd} แถว`);
        }
        
        // บันทึกข้อมูลทั้งหมด
        sheet.getRange(1, 1, sheetData.length, headers.length).setValues(sheetData);
        Logger.log(`💾 บันทึกข้อมูลแล้ว`);
        
        // ✅ บังคับให้ Google Sheets คำนวณสูตรใหม่
        SpreadsheetApp.flush();
        Logger.log(`🔄 บังคับคำนวณสูตรแล้ว`);
        
        // รอให้คำนวณเสร็จ
        Utilities.sleep(2000);
        Logger.log(`⏰ รอการคำนวณเสร็จแล้ว`);
        
        const resultMessage = `บันทึกข้อมูลและอัปเดตสถิติสำเร็จ (อัปเดต ${updateCount} ช่อง, สูตร ${formulasUpdated} แถว${newStudentsAdded ? ', เพิ่มนักเรียนใหม่' : ''})`;
        
        Logger.log(`✅ ${resultMessage}`);
        
        return { 
          message: resultMessage,
          updated: updateCount,
          newStudents: newStudentsAdded,
          formulasUpdated: formulasUpdated,
          success: true
        };
        
      } catch (saveError) {
        Logger.log("❌ เกิดข้อผิดพลาดในการบันทึก:", saveError);
        throw new Error(`เกิดข้อผิดพลาดในการบันทึก: ${saveError.message}`);
      }
    } else {
      Logger.log(`ℹ️ ไม่มีข้อมูลที่ต้องอัปเดต`);
      return { message: `ไม่มีข้อมูลที่ต้องอัปเดต` };
    }

  } catch (e) {
    Logger.log("❌ [Optimized] Error:", e);
    Logger.log("❌ Stack trace:", e.stack);
    throw new Error(`เกิดข้อผิดพลาดในการบันทึก: ${e.message}`);
  }
}

// ✅ ฟังก์ชันตรวจสอบและแก้ไขสูตรที่เสีย (เรียกใช้แยกได้)
function fixAttendanceFormulas(monthName) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) throw new Error(`ไม่พบชีต ${monthName}`);
    
    const dataRange = sheet.getDataRange();
    const sheetData = dataRange.getValues();
    const headers = sheetData[0];
    
    // หาตำแหน่งคอลัมน์
    let firstDayCol = -1, lastDayCol = -1;
    headers.forEach((h, i) => {
      const match = String(h).match(/^(\d+)$/);
      if (match) {
        if (firstDayCol === -1) firstDayCol = i;
        lastDayCol = i;
      }
    });
    
    const idxPresent = headers.indexOf("มา");
    const idxLeave = headers.indexOf("ลา");
    const idxAbsent = headers.indexOf("ขาด");
    
    if (firstDayCol === -1 || idxPresent === -1 || idxLeave === -1 || idxAbsent === -1) {
      throw new Error("ไม่พบคอลัมน์ที่จำเป็น");
    }
    
    const startCol = columnNumberToLetter(firstDayCol + 1);
    const endCol = columnNumberToLetter(lastDayCol + 1);
    
    // แก้ไขสูตรทุกแถว
    for (let i = 1; i < sheetData.length; i++) {
      const rowNum = i + 1;
      const rangeForRow = `${startCol}${rowNum}:${endCol}${rowNum}`;
      
      sheetData[i][idxPresent] = `=COUNTIF(${rangeForRow}, "/")`;
      sheetData[i][idxLeave] = `=COUNTIF(${rangeForRow}, "ล")`;
      sheetData[i][idxAbsent] = `=COUNTIF(${rangeForRow}, "ข")`;
    }
    
    // บันทึกและบังคับคำนวณ
    dataRange.setValues(sheetData);
    SpreadsheetApp.flush();
    
    return { message: `แก้ไขสูตรในชีต ${monthName} สำเร็จ` };
    
  } catch (e) {
    Logger.log("❌ Error fixing formulas:", e);
    throw new Error(`เกิดข้อผิดพลาดในการแก้ไขสูตร: ${e.message}`);
  }
}
/**
 * ✅ ฟังก์ชันบันทึกการเช็คชื่อ - รองรับทั้งรายวันและรายเดือน
 * บันทึกลงชีตรายเดือนเดียวกัน
 * @param {Object} payload - ข้อมูลการเช็คชื่อ
 * @returns {Object} ผลลัพธ์การบันทึก
 */
function saveAttendance(payload) {
  try {
    Logger.log("💾 saveAttendance called with mode:", payload.mode || 'monthly');
    Logger.log("📝 Payload:", JSON.stringify(payload, null, 2));
    
    // ตรวจสอบข้อมูลพื้นฐาน
    if (!payload || !payload.grade || !payload.classNo || !payload.records) {
      throw new Error("ข้อมูลไม่ครบถ้วน: ต้องมี grade, classNo, records");
    }
    
    let finalPayload;
    
    if (payload.mode === 'daily' && payload.date) {
      // 📅 โหมดรายวัน: แปลงจาก {studentId: status} เป็น {studentId_day: status}
      Logger.log("📅 โหมดบันทึกรายวัน");
      finalPayload = convertDailyToMonthly(payload);
    } else {
      // 📊 โหมดรายเดือน: ใช้ payload ตรงๆ
      Logger.log("📊 โหมดบันทึกรายเดือน");
      finalPayload = payload;
    }
    
    // บันทึกลงชีตรายเดือน (ใช้ฟังก์ชันที่มีอยู่แล้ว)
    Logger.log("🔄 เรียก saveMonthlyAttendance...");
    const result = saveMonthlyAttendance(finalPayload);
    
    // บันทึก log การทำงาน
    const logDate = payload.date || `${finalPayload.year}-${String(finalPayload.month + 1).padStart(2, '0')}-01`;
    const recordCount = Object.keys(payload.records).length;
    saveAttendanceLog(logDate, payload.grade, payload.classNo, payload.records, recordCount);
    
    // ส่งผลลัพธ์กลับ
    return {
      success: true,
      message: result.message || "✅ บันทึกข้อมูลสำเร็จ",
      mode: payload.mode || 'monthly',
      count: recordCount,
      details: result
    };
    
  } catch (error) {
    Logger.log("❌ Error in saveAttendance:", error.message);
    Logger.log("Error stack:", error.stack);
    
    return {
      success: false,
      error: error.message,
      message: `❌ การบันทึกล้มเหลว: ${error.message}`
    };
  }
}

/**
 * 🔄 แปลงข้อมูลรายวันเป็นรูปแบบรายเดือน
 * @param {Object} dailyPayload - ข้อมูลรายวัน
 * @returns {Object} ข้อมูลรูปแบบรายเดือน
 */
function convertDailyToMonthly(dailyPayload) {
  const { date, grade, classNo, records } = dailyPayload;
  
  Logger.log(`🔄 แปลงข้อมูลรายวัน: ${date}`);
  
  // แปลงวันที่
  const d = new Date(date);
  const year = d.getFullYear();
  const month = d.getMonth(); // 0-11
  const day = d.getDate();
  
  Logger.log(`📅 Date parsing: ${year}-${month + 1}-${day}`);
  
  // แปลง records จาก {studentId: status} เป็น {studentId_day: status}
  const monthlyRecords = {};
  Object.keys(records).forEach(studentId => {
    const key = `${studentId}_${day}`;
    monthlyRecords[key] = records[studentId];
    Logger.log(`🔑 Convert: ${studentId} -> ${key} = ${records[studentId]}`);
  });
  
  const monthlyPayload = {
    grade: grade,
    classNo: classNo,
    year: year,
    month: month,
    records: monthlyRecords
  };
  
  Logger.log("✅ แปลงเสร็จ:", JSON.stringify(monthlyPayload, null, 2));
  return monthlyPayload;
}


/**
 * ✅ สรุปการมาเรียนรายภาคเรียน (รองรับปีการศึกษา)
 * @param {string} grade ระดับชั้น
 * @param {string} classNo ห้อง
 * @param {number} semester ภาคเรียนที่ (1 หรือ 2)
 * @param {number} academicYear ปีการศึกษา (พ.ศ.)
 * @returns {Array} ข้อมูลสรุปภาคเรียน
 */
function getSemesterAttendanceSummary(grade, classNo, semester, academicYear) {
  try {
    const ss = SS();
    const yearCE = academicYear - 543; // แปลงเป็น ค.ศ.
    
    let months = [];
    if (semester == 1) {
      // ภาคเรียนที่ 1: พ.ค.-ต.ค.
      months = [
        {m: 5, y: yearCE},
        {m: 6, y: yearCE},
        {m: 7, y: yearCE},
        {m: 8, y: yearCE},
        {m: 9, y: yearCE},
        {m: 10, y: yearCE}
      ];
    } else if (semester == 2) {
      // ภาคเรียนที่ 2: พ.ย.-มี.ค.
      months = [
        {m: 11, y: yearCE},
        {m: 12, y: yearCE},
        {m: 1, y: yearCE + 1},
        {m: 2, y: yearCE + 1},
        {m: 3, y: yearCE + 1}
      ];
    }
    
    // โหลดรายชื่อนักเรียน
    const students = getStudentsForAttendance(grade, classNo);
    const studentMap = {};
    students.forEach(stu => {
      studentMap[stu.id] = {
        studentId: stu.id,
        name: `${stu.title||''}${stu.firstname} ${stu.lastname}`,
        present: 0,
        leave: 0,
        absent: 0,
        total: 0
      };
    });
    
    // รวมข้อมูลจากแต่ละเดือน
    months.forEach(obj => {
      const monthlyData = getMonthlyAttendanceSummary(grade, classNo, obj.y, obj.m);
      
      monthlyData.forEach(stuData => {
        if (studentMap[stuData.studentId]) {
          studentMap[stuData.studentId].present += stuData.present || 0;
          studentMap[stuData.studentId].leave += stuData.leave || 0;
          studentMap[stuData.studentId].absent += stuData.absent || 0;
        }
      });
    });
    
    // แปลงเป็น array
    const result = Object.values(studentMap).map(stu => {
      stu.total = stu.present + stu.leave + stu.absent;
      return stu;
    }).filter(stu => stu.total > 0);

    // --- 👇 เพิ่มโค้ดเรียงลำดับตรงนี้ ก่อน return ---
    result.sort((a, b) => String(a.studentId).localeCompare(String(b.studentId), undefined, {numeric: true}));
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---
    
    Logger.log(`✅ สรุปภาคเรียนที่ ${semester} ปีการศึกษา ${academicYear}: ${result.length} คน`);
    return result;
    
  } catch (error) {
    Logger.log('Error in getSemesterAttendanceSummary:', error);
    throw error;
  }
}

/**
 * ✅ ส่งออก PDF รายงานภาคเรียน
 */
function exportSemesterAttendancePDF(grade, classNo, semester, academicYear) {
  try {
    const summary = getSemesterAttendanceSummary(grade, classNo, semester, academicYear);
    const semesterName = semester == 1 ? 'ภาคเรียนที่ 1 (พฤษภาคม - ตุลาคม)' : 'ภาคเรียนที่ 2 (พฤศจิกายน - มีนาคม)';
    
    let schoolName = "";
    try {
      const settingsSheet = SS().getSheetByName("global_settings");
      if (settingsSheet) {
        const settings = settingsSheet.getDataRange().getValues();
        settings.forEach(row => {
          if (row[0] === "schoolName" || row[0] === "ชื่อโรงเรียน") schoolName = row[1];
        });
      }
    } catch(e) {}
    
    let html = `
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { font-family: 'Sarabun', Arial, sans-serif; margin: 30px; }
            h2 { color: #1a237e; text-align: center; margin-bottom: 4px;}
            .info { text-align: center; color: #444; margin-bottom: 10px;}
            table { width: 100%; border-collapse: collapse; margin-top: 20px;}
            th, td { border: 1px solid #888; padding: 6px 8px; text-align: center; }
            th { background: #e3f2fd; }
            .text-start { text-align: left; }
          </style>
        </head>
        <body>
          <h2>สรุปการมาเรียนรายภาคเรียน</h2>
          <div class="info">
            ${schoolName ? "โรงเรียน" + schoolName + "<br>" : ""}
            ชั้น ${grade} ห้อง ${classNo}<br>
            ${semesterName}<br>
            ปีการศึกษา ${academicYear}
          </div>
          <table>
            <thead>
              <tr>
                <th>เลขที่</th>
                <th>รหัส</th>
                <th class="text-start">ชื่อ-นามสกุล</th>
                <th>มา</th>
                <th>ลา</th>
                <th>ขาด</th>
                <th>รวมวันเรียน</th>
                <th>ร้อยละการมาเรียน</th>
              </tr>
            </thead>
            <tbody>
    `;
    
    summary.forEach((student, idx) => {
      const present = student.present || 0;
      const leave = student.leave || 0;
      const absent = student.absent || 0;
      const total = present + leave + absent;
      const percent = total > 0 ? ((present / total) * 100).toFixed(1) : "0.0";
      
      html += `<tr>
        <td>${idx + 1}</td>
        <td>${student.studentId}</td>
        <td class="text-start">${student.name}</td>
        <td>${present}</td>
        <td>${leave}</td>
        <td>${absent}</td>
        <td>${total}</td>
        <td>${percent}%</td>
      </tr>`;
    });
    
    html += `
            </tbody>
          </table>
        </body>
      </html>
    `;
    
    const blob = HtmlService.createHtmlOutput(html).getBlob()
      .setName(`สรุปภาคเรียนที่${semester}_${grade}_ห้อง${classNo}_ปี${academicYear}.pdf`)
      .getAs('application/pdf');
    
    return _saveBlobGetUrl_(blob);
  } catch (error) {
    throw new Error(`ไม่สามารถส่งออก PDF: ${error.message}`);
  }
}

// ============================================================
// 🔧 แก้ไขปัญหาการคำนวณเปอร์เซนต์ในการส่งออก PDF
// ============================================================



/**
 * ส่งออก PDF ทุกเดือนในปีเดียวกัน (alternative approach)
 * ใช้เมื่อต้องการส่งออกครั้งเดียวจาก backend
 */
function exportAllMonthsPDF(grade, classNo, year) {
  try {
    Logger.log(`🔍 exportAllMonthsPDF: ${grade}/${classNo}/${year}`);
    
    if (!grade || !classNo || !year) {
      throw new Error("พารามิเตอร์ไม่ครบถ้วน");
    }
    
    const results = [];
    const errors = [];
    
    // ส่งออกทีละเดือน
    for (let month = 1; month <= 12; month++) {
      try {
        const pdfUrl = exportMonthlyAttendancePDF(grade, classNo, year, month);
        results.push({
          month: month,
          monthName: getThaiMonthName(month),
          url: pdfUrl,
          success: true
        });
        
        // หน่วงเวลาเล็กน้อยเพื่อไม่ให้ระบบหนัก
        Utilities.sleep(500);
        
      } catch (error) {
        errors.push({
          month: month,
          monthName: getThaiMonthName(month),
          error: error.message,
          success: false
        });
        Logger.log(`Error exporting month ${month}:`, error.message);
      }
    }
    
    // สร้างรายงานสรุป
    const summary = {
      totalMonths: 12,
      successCount: results.length,
      errorCount: errors.length,
      results: results,
      errors: errors,
      grade: grade,
      classNo: classNo,
      year: year
    };
    
    Logger.log(`✅ Bulk export completed: ${results.length}/${12} successful`);
    return summary;
    
  } catch (error) {
    Logger.log("❌ Error in exportAllMonthsPDF:", error.message);
    throw new Error(`ไม่สามารถส่งออกทุกเดือนได้: ${error.message}`);
  }
}


/**
 * ดึงรายการเดือนที่มีข้อมูล
 */
function getAvailableMonths(grade, classNo, year) {
  const availableMonths = [];
  
  for (let month = 1; month <= 12; month++) {
    if (hasDataForMonth(grade, classNo, year, month)) {
      availableMonths.push({
        month: month,
        monthName: getThaiMonthName(month)
      });
    }
  }
  
  return availableMonths;
}

/**
 * ✅ ฟังก์ชันส่งออก PDF รายคน (แก้ไขการคำนวณเปอร์เซนต์)
 */
function exportStudentAttendancePDF(grade, classNo, studentId, academicYear) {
  try {
    Logger.log(`🔍 exportStudentAttendancePDF: ${grade}/${classNo}/${studentId}/${academicYear}`);
    
    // ดึงข้อมูลรายละเอียดนักเรียน
    const data = getStudentAttendanceDetail(grade, classNo, studentId, academicYear);
    
    if (!data || !data.student) {
      throw new Error(`ไม่พบข้อมูลนักเรียนรหัส ${studentId}`);
    }
    
    let html = `
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { 
              font-family: 'Sarabun', 'THSarabunNew', Arial, sans-serif; 
              margin: 20px; 
              font-size: 14px;
            }
            h2 { 
              color: #1a237e; 
              text-align: center; 
              margin-bottom: 10px;
            }
            .student-info {
              background: #f8f9fa;
              padding: 15px;
              border-radius: 8px;
              margin: 20px 0;
            }
            .summary-stats {
              display: flex;
              justify-content: space-around;
              margin: 20px 0;
              padding: 15px;
              background: #e3f2fd;
              border-radius: 8px;
            }
            .stat-item { text-align: center; }
            .stat-number { 
              font-size: 20px; 
              font-weight: bold; 
              margin-bottom: 5px; 
            }
            table { 
              width: 100%; 
              border-collapse: collapse; 
              margin-top: 20px;
              font-size: 14px;
            }
            th, td { 
              border: 1px solid #888; 
              padding: 6px 4px; 
              text-align: center; 
            }
            th { background: #e3f2fd; }
          </style>
        </head>
        <body>
          <h2>👤 รายงานการเข้าเรียนรายบุคคล</h2>
          
          <div class="student-info">
            <h4>ข้อมูลนักเรียน</h4>
            <p><strong>รหัสนักเรียน:</strong> ${data.student.id}</p>
            <p><strong>ชื่อ-นามสกุล:</strong> ${data.student.fullname}</p>
            <p><strong>ชั้น/ห้อง:</strong> ${grade} ห้อง ${classNo}</p>
            <p><strong>ปีการศึกษา:</strong> ${academicYear}</p>
          </div>
          
          <div class="summary-stats">
            <div class="stat-item">
              <div class="stat-number" style="color: #28a745;">${data.summary.totalPresent}</div>
              <div>วันมาเรียน</div>
            </div>
            <div class="stat-item">
              <div class="stat-number" style="color: #ffc107;">${data.summary.totalLeave}</div>
              <div>วันลา</div>
            </div>
            <div class="stat-item">
              <div class="stat-number" style="color: #dc3545;">${data.summary.totalAbsent}</div>
              <div>วันขาด</div>
            </div>
            <div class="stat-item">
              <div class="stat-number" style="color: #007bff;">${data.summary.percentage}%</div>
              <div>ร้อยละมาเรียน</div>
            </div>
          </div>
          
          <h4>📊 สรุปรายเดือน</h4>
          <table>
            <thead>
              <tr>
                <th>เดือน</th>
                <th>ปี</th>
                <th>มา</th>
                <th>ลา</th>
                <th>ขาด</th>
                <th>รวม</th>
                <th>%มาเรียน</th>
              </tr>
            </thead>
            <tbody>
    `;
    
    // แสดงข้อมูลรายเดือนที่มีการเรียน
    data.monthlyData.forEach(month => {
      if (month.total > 0) {
        // ✅ คำนวณเปอร์เซนต์ใหม่ให้แน่ใจ
        const monthPercentage = month.total > 0 ? ((month.present / month.total) * 100).toFixed(1) : "0.0";
        
        html += `
          <tr>
            <td>${month.monthName}</td>
            <td>${month.year}</td>
            <td>${month.present}</td>
            <td>${month.leave}</td>
            <td>${month.absent}</td>
            <td>${month.total}</td>
            <td><strong>${monthPercentage}%</strong></td>
          </tr>
        `;
      }
    });
    
    html += `
            </tbody>
          </table>
          
          <div style="margin-top: 30px; text-align: center;">
            <h4>ผลการประเมิน: <span style="color: #007bff;">${data.summary.gradeLevel}</span></h4>
            <p>จากการเข้าเรียนทั้งหมด ${data.summary.totalDays} วัน</p>
            <p style="font-size: 12px; color: #666;">
              สร้างเมื่อ: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss น.')}
            </p>
          </div>
        </body>
      </html>
    `;
    
    // สร้าง PDF
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(`รายงานรายคน_${studentId}_${data.student.firstname}_ปี${academicYear}.pdf`)
      .getAs('application/pdf');
    
    // บันทึกลง Drive
    const url = _saveBlobGetUrl_(blob);
    Logger.log(`✅ Student PDF created: ${url}`);
    return url;
    
  } catch (error) {
    Logger.log("❌ Error in exportStudentAttendancePDF:", error.message);
    throw new Error(`ไม่สามารถสร้าง PDF รายคนได้: ${error.message}`);
  }
}


/**
 * ✅ ฟังก์ชันบันทึกข้อมูลสรุปลงชีต (แก้ไขแล้ว - ตรงกับโครงสร้างจริง)
 */
function saveAttendanceSummaryToSheet(summary, academicYear) {
  Logger.log("🔧 saveAttendanceSummaryToSheet called:", {
    summaryLength: summary.length,
    academicYear: academicYear,
    firstRecord: summary[0]
  });

  try {
    var sheetName = 'สรุปการมาเรียนรวม_' + academicYear;
    var ss = SS();
    var sheet = ss.getSheetByName(sheetName);
    
    // ✅ สร้างชีตใหม่ด้วย Headers ที่ถูกต้อง
    if (!sheet) {
      Logger.log("🆕 Creating new sheet:", sheetName);
      sheet = ss.insertSheet(sheetName);
      
      // ✅ แก้ไข: ใช้ Headers ตรงกับโครงสร้างจริง
      const headers = [
        'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง', 
        'มา(ภ.1)', 'ลา(ภ.1)', 'ขาด(ภ.1)', 
        'มา(ภ.2)', 'ลา(ภ.2)', 'ขาด(ภ.2)', 
        'รวมมา', 'รวมลา', 'รวมขาด', 'รวมวันทั้งหมด', '% การมาเรียน', 'วันที่อัปเดต'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // จัดรูปแบบ header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      headerRange.setWrap(true);
      sheet.autoResizeColumns(1, headers.length);
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    Logger.log("📋 Current headers:", headers);
    
    // ✅ แก้ไข: หาตำแหน่งคอลัมน์ที่ถูกต้อง
    var columnIndex = {
      id: headers.indexOf('รหัสนักเรียน'),
      name: headers.indexOf('ชื่อ-นามสกุล'),
      grade: headers.indexOf('ชั้น'),
      classNo: headers.indexOf('ห้อง'),
      // ✅ บันทึกลงคอลัมน์ "รวม" ที่ถูกต้อง
      totalPresent: headers.indexOf('รวมมา'),
      totalLeave: headers.indexOf('รวมลา'),
      totalAbsent: headers.indexOf('รวมขาด'),
      totalDays: headers.indexOf('รวมวันทั้งหมด'),
      percentage: headers.indexOf('% การมาเรียน'),
      updateDate: headers.indexOf('วันที่อัปเดต')
    };
    
    Logger.log("📍 Column indices:", columnIndex);
    
    // ตรวจสอบว่าพบคอลัมน์หลักหรือไม่
    const requiredColumns = ['id', 'name', 'grade', 'classNo', 'totalPresent', 'totalLeave', 'totalAbsent', 'totalDays'];
    for (const col of requiredColumns) {
      if (columnIndex[col] === -1) {
        throw new Error(`ไม่พบคอลัมน์: ${col} (${headers})`);
      }
    }

    let processedCount = 0;

    summary.forEach(function(row) {
      Logger.log(`🔍 Processing: ${row.studentId} - ${row.name}`);
      
      var found = false;
      // หาแถวที่มีข้อมูลนักเรียนคนนี้อยู่แล้ว
      for (var i = 1; i < data.length; i++) {
        if (
          String(data[i][columnIndex.id]).trim() === String(row.studentId).trim() &&
          String(data[i][columnIndex.grade]).trim() === String(row.grade).trim() &&
          String(data[i][columnIndex.classNo]).trim() === String(row.classNo).trim()
        ) {
          // ✅ แก้ไข: อัพเดทคอลัมน์ที่ถูกต้อง
          Logger.log(`⬆️ Updating existing row ${i+1}`);
          
          const rowNum = i + 1; // Google Sheets เริ่มนับที่ 1
          const total = row.total || (row.present + row.leave + row.absent);
          const percentage = total > 0 ? ((row.present / total) * 100).toFixed(1) + "%" : "0.0%";
          const updateTime = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
          
          // บันทึกลงคอลัมน์ "รวม" ที่ถูกต้อง
          sheet.getRange(rowNum, columnIndex.totalPresent + 1).setValue(row.present);
          sheet.getRange(rowNum, columnIndex.totalLeave + 1).setValue(row.leave);
          sheet.getRange(rowNum, columnIndex.totalAbsent + 1).setValue(row.absent);
          sheet.getRange(rowNum, columnIndex.totalDays + 1).setValue(total);
          
          if (columnIndex.percentage !== -1) {
            sheet.getRange(rowNum, columnIndex.percentage + 1).setValue(percentage);
          }
          if (columnIndex.updateDate !== -1) {
            sheet.getRange(rowNum, columnIndex.updateDate + 1).setValue(updateTime);
          }
          
          found = true;
          break;
        }
      }
      
      if (!found) {
        // ✅ แก้ไข: เพิ่มแถวใหม่ด้วยลำดับที่ถูกต้อง
        Logger.log(`➕ Adding new row`);
        
        const total = row.total || (row.present + row.leave + row.absent);
        const percentage = total > 0 ? ((row.present / total) * 100).toFixed(1) + "%" : "0.0%";
        const updateTime = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
        
        // สร้างแถวใหม่ตามลำดับ headers ที่ถูกต้อง
        const newRow = [
          row.studentId,        // รหัสนักเรียน
          row.name,             // ชื่อ-นามสกุล
          row.grade,            // ชั้น
          row.classNo,          // ห้อง
          '',                   // มา(ภ.1) - ว่างไว้
          '',                   // ลา(ภ.1) - ว่างไว้
          '',                   // ขาด(ภ.1) - ว่างไว้
          '',                   // มา(ภ.2) - ว่างไว้
          '',                   // ลา(ภ.2) - ว่างไว้
          '',                   // ขาด(ภ.2) - ว่างไว้
          row.present,          // รวมมา ✅
          row.leave,            // รวมลา ✅
          row.absent,           // รวมขาด ✅
          total,                // รวมวันทั้งหมด ✅
          percentage,           // % การมาเรียน ✅
          updateTime            // วันที่อัปเดต ✅
        ];
        
        sheet.appendRow(newRow);
      }
      
      processedCount++;
    });
    
    Logger.log(`✅ Processed ${processedCount} records successfully`);
    return true;
    
  } catch (error) {
    Logger.log("❌ Error in saveAttendanceSummaryToSheet:", error.message);
    Logger.log("Stack trace:", error.stack);
    throw error;
  }
}


/**
 * ✅ บันทึกสรุปภาคเรียนลงชีต
 */
function saveSemesterSummaryToSheet(summary, academicYear, semester) {
  try {
    const sheetName = `สรุปภาคเรียนที่${semester}_${academicYear}`;
    const ss = SS();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = ['รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ชั้น', 'ห้อง', 'มา', 'ลา', 'ขาด', 'รวม', '% การมาเรียน', 'วันที่อัปเดต'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#4285f4").setFontColor("white").setFontWeight("bold");
    }
    
    const data = sheet.getDataRange().getValues();
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');
    
    summary.forEach(row => {
      const total = row.total || (row.present + row.leave + row.absent);
      const percentage = total > 0 ? ((row.present / total) * 100).toFixed(1) + "%" : "0.0%";
      
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(row.studentId).trim() &&
            String(data[i][2]).trim() === String(row.grade).trim() &&
            String(data[i][3]).trim() === String(row.classNo).trim()) {
          // อัพเดทข้อมูล
          sheet.getRange(i + 1, 5).setValue(row.present);
          sheet.getRange(i + 1, 6).setValue(row.leave);
          sheet.getRange(i + 1, 7).setValue(row.absent);
          sheet.getRange(i + 1, 8).setValue(total);
          sheet.getRange(i + 1, 9).setValue(percentage);
          sheet.getRange(i + 1, 10).setValue(nowStr);
          found = true;
          break;
        }
      }
      
      if (!found) {
        sheet.appendRow([
          row.studentId, row.name, row.grade, row.classNo,
          row.present, row.leave, row.absent, total, percentage, nowStr
        ]);
      }
    });
    
    return true;
  } catch (error) {
    throw error;
  }
}

// ============================================================
// 🔧 อัปเดตฟังก์ชันรายงานรายคน
// ============================================================

/**
 * ✅ [แก้ไขอีกครั้ง] ดึงข้อมูลรายบุคคล (ใช้โครงสร้างที่หน้าบ้านต้องการ)
 * @param {string} grade ระดับชั้น
 * @param {string} classNo ห้อง
 * @param {string} studentId รหัสนักเรียน
 * @param {number} academicYear ปีการศึกษา (พ.ศ.)
 * @returns {object} ข้อมูลสรุปของนักเรียน
 */
function getStudentAttendanceDetail(grade, classNo, studentId, academicYear) {
  try {
    // 1. ดึงข้อมูลพื้นฐานของนักเรียน
    const studentInfo = getStudentBasicInfo(studentId);
    if (!studentInfo) throw new Error(`ไม่พบนักเรียนรหัส ${studentId}`);

    // 2. วนลูปตามเดือนใน "ปีการศึกษา" เพื่อเก็บสถิติ
    const monthsInYear = getMonthsInAcademicYear(academicYear);
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    
    let totalPresent = 0, totalLeave = 0, totalAbsent = 0;
    const monthlyData = [];

    monthsInYear.forEach(({ month, yearCE }) => {
        const stats = getStudentMonthlyStats(studentId, yearCE, month);
        totalPresent += stats.present;
        totalLeave += stats.leave;
        totalAbsent += stats.absent;
        
        monthlyData.push({
            monthName: thaiMonths[month - 1],
            year: yearCE + 543,
            present: stats.present,
            leave: stats.leave,
            absent: stats.absent,
            total: stats.total,
            percentage: stats.total > 0 ? ((stats.present / stats.total) * 100).toFixed(1) : "0.0"
        });
    });

    // 3. คำนวณผลรวมและประเมินผล
    const totalDays = totalPresent + totalLeave + totalAbsent;
    const percentage = totalDays > 0 ? ((totalPresent / totalDays) * 100).toFixed(1) : "0.0";
    
    let gradeLevel, gradeClass;
    const percentNum = parseFloat(percentage);
    if (percentNum >= 80) { gradeLevel = "ดีมาก"; gradeClass = "success"; }
    else if (percentNum >= 70) { gradeLevel = "ดี"; gradeClass = "info"; }
    else if (percentNum >= 60) { gradeLevel = "พอใช้"; gradeClass = "warning"; }
    else { gradeLevel = "ต้องปรับปรุง"; gradeClass = "danger"; }
    
    // 4. จัดโครงสร้างข้อมูลให้ตรงกับที่ Frontend ต้องการ
    return {
      student: studentInfo,
      academicYear: academicYear,
      summary: {
        totalPresent: totalPresent,
        totalLeave: totalLeave,
        totalAbsent: totalAbsent,
        totalDays: totalDays,
        percentage: percentage,
        gradeLevel: gradeLevel,
        gradeClass: gradeClass
      },
      monthlyData: monthlyData.filter(m => m.total > 0)
    };

  } catch (e) {
    Logger.log(`Error in getStudentAttendanceDetail: ${e.message}`);
    throw new Error("ไม่สามารถดึงข้อมูลนักเรียนได้: " + e.message);
  }
}


/**
 * 🔧 ดึงสถิติรายเดือนของนักเรียนคนเดียว
 */
// Cache สำหรับ attendance sheet data (ลด API calls จาก 12 เหลือไม่ซ้ำ)
var _attendanceSheetCache = {};

function getStudentMonthlyStats(studentId, yearCE, month) {
  const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const sheetName = `${monthNames[month-1]}${yearCE + 543}`;
  
  // ใช้ cache ถ้ามี
  if (_attendanceSheetCache[sheetName] !== undefined) {
    var data = _attendanceSheetCache[sheetName];
    if (!data) return { present: 0, leave: 0, absent: 0, total: 0 };
  } else {
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      _attendanceSheetCache[sheetName] = null;
      return { present: 0, leave: 0, absent: 0, total: 0 };
    }
    
    var data = sheet.getDataRange().getValues();
    _attendanceSheetCache[sheetName] = data;
  }
  const headers = data[0];
  
  // หาแถวของนักเรียนคนนี้
  let studentRow = null;
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][1]).trim(); // คอลัมน์ B = รหัสนักเรียน
    const match = rowId.match(/^(\d+)/);
    if (match && match[1] === studentId) {
      studentRow = data[i];
      break;
    }
  }
  
  if (!studentRow) {
    return { present: 0, leave: 0, absent: 0, total: 0 };
  }
  
  // หาคอลัมน์วันที่
  const dateCols = [];
  for (let i = 3; i < headers.length; i++) {
    const h = String(headers[i]).trim();
    const dayMatch = h.match(/^(\d{1,2})/);
    if (dayMatch) {
      const day = parseInt(dayMatch[1]);
      dateCols.push({ col: i, day: day });
    }
  }
  
  // นับสถานะแต่ละวัน
  let present = 0, leave = 0, absent = 0;
  
  dateCols.forEach(dateCol => {
    const status = String(studentRow[dateCol.col] || '').trim();
    if (status === '/' || status === 'ม' || status === '1' || status === '✓') {
      present++;
    } else if (status === 'ล' || status.toLowerCase() === 'l') {
      leave++;
    } else if (status === 'ข' || status === '0') {
      absent++;
    }
  });
  
  return {
    present: present,
    leave: leave,
    absent: absent,
    total: present + leave + absent
  };
}

// ============================================================
// 🔧 ฟังก์ชันช่วยเหลืออื่นๆ
// ============================================================

/**
 * 🔧 ดึงข้อมูลพื้นฐานของนักเรียน
 */
function getStudentBasicInfo(studentId) {
  try {
    const ss = SS();
    const studentsSheet = ss.getSheetByName('Students');
    
    if (!studentsSheet) {
      throw new Error("ไม่พบชีต Students");
    }
    
    const data = studentsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const row = data.find(r => String(r[headers.indexOf("student_id")]) === String(studentId));
    if (!row) {
      return null;
    }
    
    return {
      id: studentId,
      title: row[headers.indexOf("title")] || '',
      firstname: row[headers.indexOf("firstname")] || '',
      lastname: row[headers.indexOf("lastname")] || '',
      fullname: `${row[headers.indexOf("title")] || ''}${row[headers.indexOf("firstname")] || ''} ${row[headers.indexOf("lastname")] || ''}`.trim(),
      grade: row[headers.indexOf("grade")] || '',
      classNo: row[headers.indexOf("class_no")] || ''
    };
    
  } catch (error) {
    Logger.log('Error in getStudentBasicInfo:', error);
    return null;
  }
}

/**
 * 🔧 ดึงรายชื่อนักเรียนสำหรับ dropdown
 */
function getStudentsList(grade, classNo) {
  try {
    const students = getStudentsForAttendance(grade, classNo);
    return students.map(student => ({
      id: student.id,
      name: `${student.title || ''}${student.firstname} ${student.lastname}`.trim()
    }));
  } catch (error) {
    Logger.log('Error in getStudentsList:', error);
    return [];
  }
}

/**
 * ✅ [แก้ไขใหม่] ดึงข้อมูลสรุปการมาเรียน "ทั้งปีการศึกษา" (เร็วและแม่นยำขึ้น)
 * @param {string} grade ระดับชั้น
 * @param {string} classNo ห้อง
 * @param {number} academicYear ปีการศึกษา (พ.ศ.)
 * @returns {Array} ข้อมูลสรุปของนักเรียนทั้งปี
 */
function getYearlyAttendanceSummary(grade, classNo, academicYear) {
  try {
    const ss = SS();
    
    // 1. ดึงรายชื่อนักเรียนที่ต้องการทั้งหมด "ครั้งเดียว"
    const studentsInClass = getStudentsForAttendance(grade, classNo);
    if (!studentsInClass || studentsInClass.length === 0) {
      Logger.log(`ไม่พบนักเรียนในชั้น ${grade}/${classNo} สำหรับปีการศึกษา ${academicYear}`);
      return [];
    }

    // 2. สร้าง Map เพื่อเก็บผลรวมของนักเรียนทุกคน
    const studentStats = new Map();
    studentsInClass.forEach(stu => {
      studentStats.set(stu.id, {
        studentId: stu.id,
        name: `${stu.title || ''}${stu.firstname} ${stu.lastname}`.trim(),
        present: 0,
        leave: 0,
        absent: 0
      });
    });
    
    // 3. วนลูปตามเดือนใน "ปีการศึกษา" (พ.ค. - มี.ค.)
    const monthsInYear = getMonthsInAcademicYear(academicYear);
    
    monthsInYear.forEach(({ month, yearCE }) => {
      // ใช้ฟังก์ชันรายเดือนที่แก้ไขใหม่แล้ว ซึ่งจะทำงานได้เร็วขึ้นมาก
      const monthlySummary = getMonthlyAttendanceSummary(grade, classNo, yearCE, month);
      
      // นำข้อมูลรายเดือนมารวมใน Map ใหญ่
      monthlySummary.forEach(stuData => {
        if (studentStats.has(stuData.studentId)) {
          const stats = studentStats.get(stuData.studentId);
          stats.present += stuData.present;
          stats.leave += stuData.leave;
          stats.absent += stuData.absent;
        }
      });
    });

    // 4. แปลง Map เป็น Array และคำนวณผลรวมสุดท้าย
    const result = Array.from(studentStats.values()).map(stu => {
      const total = stu.present + stu.leave + stu.absent;
      return { ...stu, total };
    }).filter(stu => stu.total > 0); // กรองเฉพาะคนที่มีข้อมูล

    // --- 👇 เพิ่มโค้ดเรียงลำดับตรงนี้ ก่อน return ---
    result.sort((a, b) => String(a.studentId).localeCompare(String(b.studentId), undefined, {numeric: true}));
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

    Logger.log(`สรุปข้อมูลปีการศึกษา ${academicYear} สำเร็จ, พบนักเรียน ${result.length} คน`);
    return result;
    
  } catch (e) {
    Logger.log(`Error in getYearlyAttendanceSummary: ${e.message}`);
    throw e;
  }
}



/**
 * วิธีที่ 1: ลองดึงจากชีต "สรุปการมาเรียนรวม_ปี"
 */
function tryGetFromSummarySheet(grade, classNo, academicYear) {
  try {
    const ss = SS();
    const sheetName = `สรุปการมาเรียนรวม_${academicYear}`;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`ไม่พบชีต: ${sheetName}`);
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log(`ชีต ${sheetName} ไม่มีข้อมูล`);
      return null;
    }
    
    const headers = data[0];
    Logger.log(`Headers ใน ${sheetName}:`, headers);
    
    // หาตำแหน่งคอลัมน์
    const columnIndex = {
      id: U_findColumnIndex(headers, ["รหัสนักเรียน", "รหัส", "student_id"]),
      name: U_findColumnIndex(headers, ["ชื่อ-นามสกุล", "ชื่อ", "name"]),
      grade: U_findColumnIndex(headers, ["ชั้น", "grade"]),
      classNo: U_findColumnIndex(headers, ["ห้อง", "class"]),
      present: U_findColumnIndex(headers, ["รวมมา", "มา", "present"]),
      leave: U_findColumnIndex(headers, ["รวมลา", "ลา", "leave"]),
      absent: U_findColumnIndex(headers, ["รวมขาด", "ขาด", "absent"]),
      total: U_findColumnIndex(headers, ["รวมวันทั้งหมด", "รวม", "total"]),
      percent: U_findColumnIndex(headers, ["% การมาเรียน", "%", "percent"])
    };
    
    // ตรวจสอบว่าพบคอลัมน์หลักหรือไม่
    const requiredColumns = ['id', 'name', 'grade', 'classNo', 'present', 'leave', 'absent'];
    for (const col of requiredColumns) {
      if (columnIndex[col] === -1) {
        Logger.log(`ไม่พบคอลัมน์: ${col}`);
        return null;
      }
    }
    
    // กรองข้อมูลตามชั้น/ห้อง
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = String(row[columnIndex.grade] || '').trim();
      const rowClass = String(row[columnIndex.classNo] || '').trim();
      
      if (rowGrade === grade && rowClass === classNo) {
        const present = parseInt(row[columnIndex.present]) || 0;
        const leave = parseInt(row[columnIndex.leave]) || 0;
        const absent = parseInt(row[columnIndex.absent]) || 0;
        const total = parseInt(row[columnIndex.total]) || (present + leave + absent);
        const percent = total > 0 ? ((present / total) * 100).toFixed(1) + "%" : "0.0%";
        
        result.push({
          studentId: String(row[columnIndex.id] || '').trim(),
          name: String(row[columnIndex.name] || '').trim(),
          present: present,
          leave: leave,
          absent: absent,
          total: total,
          percent: percent
        });
      }
    }
    
    return result;
    
  } catch (error) {
    Logger.log("Error in tryGetFromSummarySheet:", error.message);
    return null;
  }
}

/**
 * วิธีที่ 2: รวมข้อมูลจากชีตรายเดือน
 */
function calculateFromMonthlySheets(grade, classNo, academicYear) {
  try {
    const students = getStudentsForAttendance(grade, classNo);
    if (!students || students.length === 0) {
      Logger.log("ไม่พบรายชื่อนักเรียน");
      return [];
    }
    
    const result = [];
    
    students.forEach(student => {
      let totalPresent = 0, totalLeave = 0, totalAbsent = 0, totalDays = 0;
      
      // รวมข้อมูลจาก 12 เดือน
      for (let month = 1; month <= 12; month++) {
        try {
          const monthlyData = getMonthlyAttendanceSummary(grade, classNo, academicYear - 543, month);
          const studentData = monthlyData.find(s => String(s.studentId).trim() === String(student.id).trim());
          
          if (studentData) {
            totalPresent += studentData.present || 0;
            totalLeave += studentData.leave || 0;
            totalAbsent += studentData.absent || 0;
            totalDays += studentData.total || 0;
          }
        } catch (monthError) {
          Logger.log(`ไม่มีข้อมูลเดือน ${month}:`, monthError.message);
          // ไม่ต้อง throw error เพราะอาจมีบางเดือนที่ไม่มีข้อมูล
        }
      }
      
      const percentage = totalDays > 0 ? ((totalPresent / totalDays) * 100).toFixed(1) + "%" : "0.0%";
      
      result.push({
        studentId: student.id,
        name: `${student.title || ''}${student.firstname} ${student.lastname}`.trim(),
        present: totalPresent,
        leave: totalLeave,
        absent: totalAbsent,
        total: totalDays,
        percent: percentage
      });
    });
    
    return result;
    
  } catch (error) {
    Logger.log("Error in calculateFromMonthlySheets:", error.message);
    throw error;
  }
}

/**
 * ฟังก์ชันช่วยหาตำแหน่งคอลัมน์
 */


/**
 * ✅ ฟังก์ชันตรวจสอบ Sheets ที่มีอยู่
 */
function listAvailableSheets() {
  try {
    const ss = SS();
    const sheets = ss.getSheets().map(sheet => ({
      name: sheet.getName(),
      rows: sheet.getLastRow(),
      cols: sheet.getLastColumn()
    }));
    Logger.log("📊 Available sheets:", sheets);
    return sheets;
  } catch (error) {
    Logger.log("Error listing sheets:", error.message);
    return [];
  }
}

/**
 * ✅ ฟังก์ชันสร้างข้อมูลตัวอย่าง (รันครั้งเดียว)
 */
function createSampleYearlyData() {
  try {
    const ss = SS();
    const sheetName = "สรุปการมาเรียนรวม_2568";
    
    // ลบ sheet เก่าถ้ามี
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // สร้าง sheet ใหม่
    const sheet = ss.insertSheet(sheetName);
    
    // หัวตาราง
    const headers = [
      "รหัสนักเรียน", "ชื่อ-นามสกุล", "ชั้น", "ห้อง", 
      "รวมมา", "รวมลา", "รวมขาด", "รวมวันทั้งหมด", "% การมาเรียน"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ข้อมูลตัวอย่าง
    const sampleData = [
      ["001", "เด็กชายสมชาย ใจดี", "ป.1", "1", 180, 5, 3, 188, "95.7%"],
      ["002", "เด็กหญิงสมใส สุขใส", "ป.1", "1", 175, 8, 5, 188, "93.1%"],
      ["003", "เด็กชายสมศักดิ์ มั่นคง", "ป.1", "1", 185, 2, 1, 188, "98.4%"],
      ["004", "เด็กหญิงสมหวัง รุ่งเรือง", "ป.1", "2", 178, 6, 4, 188, "94.7%"],
      ["005", "เด็กชายสมปอง เก่งเก็บ", "ป.1", "2", 182, 4, 2, 188, "96.8%"]
    ];
    
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
    
    // จัดรูปแบบ
    sheet.getRange(1, 1, 1, headers.length).setBackground("#4285f4").setFontColor("white").setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✅ สร้าง ${sheetName} พร้อมข้อมูลตัวอย่างสำเร็จ`);
    return `สร้าง ${sheetName} สำเร็จ`;
    
  } catch (error) {
    Logger.log("Error creating sample data:", error.message);
    return `สร้างข้อมูลไม่สำเร็จ: ${error.message}`;
  }
}
/***** ====== CONFIG / HELPERS ====== *****/
const MONTH_NAMES_TH = [
  'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
];

/** แปลงค่าเป็นจำนวนเต็มแบบปลอดภัย (กันวันที่/สตริงปน) */
function toIntSafe(v) {
  if (v instanceof Date) return 0;
  const s = String(v ?? '').trim();
  if (/\d{4}-\d{2}-\d{2}/.test(s) || /:/.test(s)) return 0; // กันสตริงวัน-เวลา
  const m = s.match(/-?\d+/);
  const n = m ? parseInt(m[0], 10) : 0;
  return Number.isFinite(n) && n > 0 ? n : 0;
}

/** header ชีตสรุปปี */
function headersYearly() {
  return [
    'รหัสนักเรียน','ชื่อ-นามสกุล','ชั้น','ห้อง',
    'มา(ภ.1)','ลา(ภ.1)','ขาด(ภ.1)',
    'มา(ภ.2)','ลา(ภ.2)','ขาด(ภ.2)',
    'รวมมา','รวมลา','รวมขาด','รวมวันทั้งหมด','% การมาเรียน','วันที่อัปเดต'
  ];
}

/** ค้น/คำนวณปีการศึกษา (พ.ศ.) กันค่าหาย/undefined */
function resolveAcademicYearBE_(arg) {
  // 1) ถ้าส่งมาเป็นตัวเลข พ.ศ. ใช้เลย
  const n = Number(arg);
  if (Number.isFinite(n) && n >= 2500 && n <= 2700) return n;

  // 2) ลองอ่านจาก global_settings!A:B แถว "ปีการศึกษา"
  try {
    const ss = SS();
    const gs = ss.getSheetByName('global_settings');
    if (gs) {
      const data = gs.getDataRange().getValues();
      const row = data.find(r => String(r[0]).trim() === 'ปีการศึกษา');
      const y = row ? Number(row[1]) : NaN;
      if (Number.isFinite(y) && y >= 2500 && y <= 2700) return y;
    }
  } catch (_) {}

  // 3) คำนวณจากวันที่ปัจจุบัน (ปีการศึกษาเริ่ม พ.ค.)
  const now = new Date();
  const gYear = now.getFullYear();
  const gMonth = now.getMonth() + 1;
  const academicYearG = (gMonth >= 5 ? gYear : gYear - 1); // ค.ศ.
  return academicYearG + 543; // พ.ศ.
}

/** สร้าง/รีเซ็ตชีตสรุปปีให้ตรงสเปก */
function ensureYearlySheet_(ss, academicYearBE_arg) {
  const academicYearBE = resolveAcademicYearBE_(academicYearBE_arg);
  const name = `สรุปการมาเรียนรวม_${academicYearBE}`;
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name); else sh.clearContents();

  const H = headersYearly();
  sh.getRange(1,1,1,H.length).setValues([H]);
  sh.getRange(1,1,1,H.length)
    .setBackground('#4285f4').setFontColor('#fff').setFontWeight('bold').setWrap(true);
  return sh;
}

/***** ====== CORE: รวมผลจากชีตรายเดือน พ.ค.–เม.ย. ====== *****/
/**
 * สร้าง/เขียนชีต "สรุปการมาเรียนรวม_<ปีการศึกษา>" จากชีตรายเดือน
 * - ใช้คอลัมน์ท้ายตาราง "มา/ลา/ขาด" ของแต่ละเดือนเป็นแหล่งความจริง
 * - ภาคเรียน 1: พ.ค.–ต.ค. ของปีฐาน (ค.ศ.=ปีการศึกษา-543)
 * - ภาคเรียน 2: พ.ย.–ธ.ค. ปีฐาน + ม.ค.–เม.ย. ปีถัดไป
 * @param {number} academicYearBE_arg - ปีการศึกษา (พ.ศ.) เช่น 2568 (เว้นว่างได้)
 */
function rebuildYearlySummarySheet(academicYearBE_arg) {
  const academicYearBE = resolveAcademicYearBE_(academicYearBE_arg);
  const baseYearG = academicYearBE - 543; // ค.ศ. ปีฐานของภาคเรียน
  const ss = SS();

  // โหลด Students → map id -> {name, grade, classNo}
  const studentsSheet = ss.getSheetByName('Students');
  if (!studentsSheet) throw new Error('ไม่พบชีต Students');
  const sVals = studentsSheet.getDataRange().getValues();
  const stuMap = {}; // id -> {name, grade, classNo}
  for (let i = 1; i < sVals.length; i++) {
    const r = sVals[i];
    const id = String(r[0]).trim();
    if (!id) continue;
    stuMap[id] = {
      name: `${String(r[2]||'')}${String(r[3]||'')} ${String(r[4]||'')}`.trim(),
      grade: String(r[5]||'').trim(),
      classNo: String(r[6]||'').trim()
    };
  }

  // ตัวสะสมต่อคน: s1 (ภาค 1), s2 (ภาค 2)
  const acc = {};
  Object.keys(stuMap).forEach(id => {
    const s = stuMap[id];
    acc[id] = { name: s.name, grade: s.grade, classNo: s.classNo, s1:{p:0,l:0,a:0}, s2:{p:0,l:0,a:0} };
  });

  // อ่านชีตรายเดือน → ดึงคอลัมน์ มา/ลา/ขาด ของแต่ละคน
  function pullMonth_(month1to12, gregYear) {
    const sheetName = `${MONTH_NAMES_TH[month1to12-1]}${gregYear+543}`;
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const headers = values[0].map(v => String(v).trim());

    const idxId      = headers.findIndex(h => h.includes('รหัส'));
    const idxPresent = headers.findIndex(h => h.startsWith('มา'));
    const idxLeave   = headers.findIndex(h => h.startsWith('ลา'));
    const idxAbsent  = headers.findIndex(h => h.startsWith('ขาด'));
    if (idxId === -1 || idxPresent === -1 || idxLeave === -1 || idxAbsent === -1) return;

    // เดือนนี้อยู่ภาคไหน (อิงปีการศึกษา academicYearBE)
    const isS1 = (gregYear === baseYearG && month1to12 >= 5 && month1to12 <= 10); // พ.ค.–ต.ค.
    const isS2 = ((gregYear === baseYearG && (month1to12 === 11 || month1to12 === 12)) ||
                  (gregYear === baseYearG+1 && month1to12 >= 1 && month1to12 <= 4));
    const semKey = isS1 ? 's1' : (isS2 ? 's2' : null);
    if (!semKey) return; // นอกช่วงปีการศึกษานี้

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      // รหัสในชีตรายเดือนบางครั้งมีตัวหนังสือพ่วง → เอาเฉพาะตัวเลขนำหน้า
      const rawId = String(row[idxId] || '').trim();
      const m = rawId.match(/^(\d+)/);
      const id = m ? m[1] : '';
      if (!id || !acc[id]) continue; // ไม่พบใน Students → ข้าม

      acc[id][semKey].p += toIntSafe(row[idxPresent]);
      acc[id][semKey].l += toIntSafe(row[idxLeave]);
      acc[id][semKey].a += toIntSafe(row[idxAbsent]);
    }
  }

  // ภาคเรียน 1: พ.ค.–ต.ค. (ปีฐาน)
  for (let m = 5; m <= 10; m++) pullMonth_(m, baseYearG);
  // ภาคเรียน 2: พ.ย.–ธ.ค. (ปีฐาน)
  for (let m = 11; m <= 12; m++) pullMonth_(m, baseYearG);
  // ภาคเรียน 2: ม.ค.–เม.ย. (ปีถัดไป)
  for (let m = 1; m <= 4; m++) pullMonth_(m, baseYearG + 1);

  // เขียนชีตสรุปปี
  const yearly = ensureYearlySheet_(ss, academicYearBE);
  const H = headersYearly();
  const rows = [];
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss');

  // เรียงข้อมูล: ชั้น → ห้อง → รหัส
  const ids = Object.keys(acc).sort((a,b) => {
    const A = acc[a], B = acc[b];
    return (A.grade||'').localeCompare(B.grade||'', 'th') ||
           (A.classNo||'').localeCompare(B.classNo||'', 'th') ||
           String(a).localeCompare(String(b), 'th');
  });

  ids.forEach(id => {
    const s = acc[id];
    const s1p=s.s1.p, s1l=s.s1.l, s1a=s.s1.a;
    const s2p=s.s2.p, s2l=s.s2.l, s2a=s.s2.a;
    const sumP = s1p + s2p;
    const sumL = s1l + s2l;
    const sumA = s1a + s2a;
    const sumT = sumP + sumL + sumA;
    const pct  = sumT > 0 ? (sumP / sumT) : 0;

    rows.push([
      id, s.name, s.grade, s.classNo,
      s1p, s1l, s1a,
      s2p, s2l, s2a,
      sumP, sumL, sumA, sumT,
      pct, nowStr
    ]);
  });

  if (rows.length) yearly.getRange(2,1,rows.length,H.length).setValues(rows);

  // จัดรูปแบบตัวเลข/เปอร์เซ็นต์
  const numCols = ['มา(ภ.1)','ลา(ภ.1)','ขาด(ภ.1)','มา(ภ.2)','ลา(ภ.2)','ขาด(ภ.2)','รวมมา','รวมลา','รวมขาด','รวมวันทั้งหมด'];
  numCols.forEach(h => {
    const c = H.indexOf(h)+1; if (c>0) yearly.getRange(2,c,Math.max(rows.length,0),1).setNumberFormat('0');
  });
  const pc = H.indexOf('% การมาเรียน')+1;
  if (pc>0) yearly.getRange(2,pc,Math.max(rows.length,0),1).setNumberFormat('0.0%');
  const uc = H.indexOf('วันที่อัปเดต')+1;
  if (uc>0) yearly.getRange(2,uc,Math.max(rows.length,0),1).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  yearly.autoResizeColumns(1, H.length);
  return `✅ สร้างชีต "สรุปการมาเรียนรวม_${academicYearBE}" สำเร็จ (นักเรียน: ${rows.length})`;
}


// ============================================================
// 🆕 ฟังก์ชัน Backend รองรับระบบปีการศึกษา (พ.ค.-มี.ค.)
// ============================================================

/**
 * 🆕 ส่งออก PDF ทั้งปีการศึกษา (พ.ค.-มี.ค.)
 */
function exportAcademicYearPDF(grade, classNo, academicYearBE) {
  try {
    Logger.log(`🔍 exportAcademicYearPDF: ${grade}/${classNo}/${academicYearBE}`);
    
    if (!grade || !classNo || !academicYearBE) {
      throw new Error("พารามิเตอร์ไม่ครบถ้วน");
    }
    
    const results = [];
    const errors = [];
    const baseYearAD = academicYearBE - 543;
    
    // สร้างรายการเดือนตามปีการศึกษา
    const academicMonths = [
      // ภาคเรียนที่ 1: พ.ค.-ต.ค. (ปีฐาน)
      { month: 5, year: baseYearAD, semester: 1 },
      { month: 6, year: baseYearAD, semester: 1 },
      { month: 7, year: baseYearAD, semester: 1 },
      { month: 8, year: baseYearAD, semester: 1 },
      { month: 9, year: baseYearAD, semester: 1 },
      { month: 10, year: baseYearAD, semester: 1 },
      // ภาคเรียนที่ 2: พ.ย.-ธ.ค. (ปีฐาน)
      { month: 11, year: baseYearAD, semester: 2 },
      { month: 12, year: baseYearAD, semester: 2 },
      // ภาคเรียนที่ 2 (ต่อ): ม.ค.-มี.ค. (ปีถัดไป)
      { month: 1, year: baseYearAD + 1, semester: 2 },
      { month: 2, year: baseYearAD + 1, semester: 2 },
      { month: 3, year: baseYearAD + 1, semester: 2 }
    ];
    
    // ส่งออกทีละเดือน
    academicMonths.forEach((monthInfo, index) => {
      try {
        Logger.log(`Processing month ${monthInfo.month}/${monthInfo.year} (${index + 1}/${academicMonths.length})`);
        
        const pdfUrl = exportMonthlyAttendancePDF(grade, classNo, monthInfo.year, monthInfo.month);
        
        results.push({
          month: monthInfo.month,
          year: monthInfo.year,
          yearBE: monthInfo.year + 543,
          semester: monthInfo.semester,
          monthName: getThaiMonthName(monthInfo.month),
          displayName: `${getThaiMonthName(monthInfo.month)} ${monthInfo.year + 543}`,
          url: pdfUrl,
          success: true
        });
        
        // หน่วงเวลาเล็กน้อยเพื่อไม่ให้ระบบหนัก
        Utilities.sleep(500);
        
      } catch (error) {
        errors.push({
          month: monthInfo.month,
          year: monthInfo.year,
          yearBE: monthInfo.year + 543,
          semester: monthInfo.semester,
          monthName: getThaiMonthName(monthInfo.month),
          displayName: `${getThaiMonthName(monthInfo.month)} ${monthInfo.year + 543}`,
          error: error.message,
          success: false
        });
        Logger.log(`Error exporting ${monthInfo.month}/${monthInfo.year}:`, error.message);
      }
    });
    
    // สร้างรายงานสรุป
    const summary = {
      academicYear: academicYearBE,
      grade: grade,
      classNo: classNo,
      totalMonths: academicMonths.length,
      successCount: results.length,
      errorCount: errors.length,
      semester1Success: results.filter(r => r.semester === 1).length,
      semester2Success: results.filter(r => r.semester === 2).length,
      results: results,
      errors: errors,
      exportDate: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
    
    Logger.log(`✅ Academic year export completed: ${results.length}/${academicMonths.length} successful`);
    return summary;
    
  } catch (error) {
    Logger.log("❌ Error in exportAcademicYearPDF:", error.message);
    throw new Error(`ไม่สามารถส่งออกทั้งปีการศึกษาได้: ${error.message}`);
  }
}

// ============================================================
// 🆕 ฟังก์ชัน Backend รองรับระบบปีการศึกษา (พ.ค.-มี.ค.)
// ============================================================

/**
 * 🆕 ส่งออก PDF ทั้งปีการศึกษาเป็น ZIP (พ.ค.-มี.ค.) - เวอร์ชันที่สร้าง ZIP โดยตรง
 */
function exportAcademicYearZIP(grade, classNo, academicYearBE) {
  try {
    Logger.log(`🔍 exportAcademicYearZIP: ${grade}/${classNo}/${academicYearBE}`);
    
    if (!grade || !classNo || !academicYearBE) {
      throw new Error("พารามิเตอร์ไม่ครบถ้วน");
    }
    
    const pdfBlobs = [];
    const results = [];
    const errors = [];
    const baseYearAD = academicYearBE - 543;
    
    // สร้างรายการเดือนตามปีการศึกษา
    const academicMonths = [
      // ภาคเรียนที่ 1: พ.ค.-ต.ค. (ปีฐาน)
      { month: 5, year: baseYearAD, semester: 1 },
      { month: 6, year: baseYearAD, semester: 1 },
      { month: 7, year: baseYearAD, semester: 1 },
      { month: 8, year: baseYearAD, semester: 1 },
      { month: 9, year: baseYearAD, semester: 1 },
      { month: 10, year: baseYearAD, semester: 1 },
      // ภาคเรียนที่ 2: พ.ย.-ธ.ค. (ปีฐาน)
      { month: 11, year: baseYearAD, semester: 2 },
      { month: 12, year: baseYearAD, semester: 2 },
      // ภาคเรียนที่ 2 (ต่อ): ม.ค.-มี.ค. (ปีถัดไป)
      { month: 1, year: baseYearAD + 1, semester: 2 },
      { month: 2, year: baseYearAD + 1, semester: 2 },
      { month: 3, year: baseYearAD + 1, semester: 2 }
    ];
    
    // สร้าง PDF แต่ละเดือนและเก็บ blob
    academicMonths.forEach((monthInfo, index) => {
      try {
        Logger.log(`Processing month ${monthInfo.month}/${monthInfo.year} (${index + 1}/${academicMonths.length})`);
        
        // สร้าง PDF blob โดยตรงแทนการสร้างไฟล์ใน Drive
        const pdfBlob = createMonthlyPDFBlob(grade, classNo, monthInfo.year, monthInfo.month);
        
        if (pdfBlob) {
          // สร้างชื่อไฟล์ที่เป็นระเบียบ
          const paddedIndex = String(index + 1).padStart(2, '0');
          const monthName = getThaiMonthName(monthInfo.month);
          const yearBE = monthInfo.year + 543;
          const fileName = `${paddedIndex}_${monthName}${yearBE}_ภาคเรียนที่${monthInfo.semester}.pdf`;
          
          pdfBlob.setName(fileName);
          pdfBlobs.push(pdfBlob);
          
          results.push({
            month: monthInfo.month,
            year: monthInfo.year,
            semester: monthInfo.semester,
            fileName: fileName,
            success: true
          });
        }
        
        // หน่วงเวลาเล็กน้อย
        Utilities.sleep(300);
        
      } catch (error) {
        errors.push({
          month: monthInfo.month,
          year: monthInfo.year,
          semester: monthInfo.semester,
          error: error.message,
          success: false
        });
        Logger.log(`Error creating PDF for ${monthInfo.month}/${monthInfo.year}:`, error.message);
      }
    });
    
    if (pdfBlobs.length === 0) {
      throw new Error("ไม่สามารถสร้าง PDF ใดๆ ได้");
    }
    
    // สร้าง ZIP
    Logger.log(`Creating ZIP with ${pdfBlobs.length} PDF files...`);
    const zipBlob = Utilities.zip(pdfBlobs, `รายงานการเข้าเรียน_${grade}_ห้อง${classNo}_ปีการศึกษา${academicYearBE}.zip`);
    
    // บันทึกลง Google Drive
    const zipUrl = _saveBlobGetUrl_(zipBlob);
    
    // สร้างรายงานสรุป
    const summary = {
      academicYear: academicYearBE,
      grade: grade,
      classNo: classNo,
      zipUrl: zipUrl,
      zipFileName: zipBlob.getName(),
      totalMonths: academicMonths.length,
      successCount: results.length,
      errorCount: errors.length,
      semester1Success: results.filter(r => r.semester === 1).length,
      semester2Success: results.filter(r => r.semester === 2).length,
      results: results,
      errors: errors,
      exportDate: Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss')
    };
    
    Logger.log(`✅ Academic year ZIP export completed: ${results.length}/${academicMonths.length} successful`);
    return zipFile.getUrl(); // ส่งคืนเฉพาะ URL เพื่อให้ frontend ใช้งานง่าย
    
  } catch (error) {
    Logger.log("❌ Error in exportAcademicYearZIP:", error.message);
    throw new Error(`ไม่สามารถส่งออกทั้งปีการศึกษาเป็น ZIP ได้: ${error.message}`);
  }
}



/**
 * 🆕 ดึงรายการเดือนที่มีข้อมูลในปีการศึกษา
 */
function getAvailableMonthsInAcademicYear(grade, classNo, academicYearBE) {
  const availableMonths = [];
  const baseYearAD = academicYearBE - 543;
  
  // เดือนในปีการศึกษา
  const academicMonths = [
    { month: 5, year: baseYearAD },
    { month: 6, year: baseYearAD },
    { month: 7, year: baseYearAD },
    { month: 8, year: baseYearAD },
    { month: 9, year: baseYearAD },
    { month: 10, year: baseYearAD },
    { month: 11, year: baseYearAD },
    { month: 12, year: baseYearAD },
    { month: 1, year: baseYearAD + 1 },
    { month: 2, year: baseYearAD + 1 },
    { month: 3, year: baseYearAD + 1 }
  ];
  
  academicMonths.forEach(monthInfo => {
    if (hasDataForMonth(grade, classNo, monthInfo.year, monthInfo.month)) {
      availableMonths.push({
        month: monthInfo.month,
        year: monthInfo.year,
        yearBE: monthInfo.year + 543,
        monthName: getThaiMonthName(monthInfo.month),
        displayName: `${getThaiMonthName(monthInfo.month)} ${monthInfo.year + 543}`
      });
    }
  });
  
  return availableMonths;
}

/**
 * 🆕 แปลงเลขเดือนเป็นชื่อเดือนไทย
 */
function getThaiMonthName(month) {
  const thaiMonths = [
    'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
  ];
  return thaiMonths[month - 1] || 'ไม่ทราบ';
}

/**
 * 🆕 ตรวจสอบว่ามีข้อมูลในเดือนนั้นหรือไม่
 */
function hasDataForMonth(grade, classNo, year, month) {
  try {
    const data = getMonthlyAttendanceSummary(grade, classNo, year, month);
    return data && data.length > 0;
  } catch (error) {
    return false;
  }
}
// ============================================================
// 📋 อัปเดตฟังก์ชัน PDF Export ให้รองรับระบบปีการศึกษา
// ============================================================




// ════════════════════════════════════════════════════════════════
// ⚠️  WARNING: PDF LAYOUT LOCKED — DO NOT MODIFY WITHOUT TESTING
// ════════════════════════════════════════════════════════════════
// ฟังก์ชันที่ล็อค: exportMonthlyAttendancePDF, exportYearlyAttendancePDF,
//   createEmptyMonthPDF, createMonthlyPDFBlob, createChecklistPDFBlob
// ════════════════════════════════════════════════════════════════

/**
 * ✅ ส่งออก PDF รายเดือน - รองรับระบบปีการศึกษา
 */
function exportMonthlyAttendancePDF(grade, classNo, yearCE, month) {
  try {
    Logger.log(`🔍 exportMonthlyAttendancePDF: ${grade}/${classNo}/${yearCE}/${month}`);
    
    // ดึงข้อมูลด้วยฟังก์ชันที่แก้ไขแล้ว
    const data = getMonthlyAttendanceSummary(grade, classNo, yearCE, month);
    const academicYear = getAcademicYearFromDate(yearCE, month);
    
    if (!data || data.length === 0) {
      Logger.log(`⚠️ ไม่พบข้อมูลสำหรับ ${grade}/${classNo} เดือน ${month}/${yearCE}`);
      return createEmptyMonthPDF(grade, classNo, yearCE, month);
    }
    
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthName = thaiMonths[month - 1];
    const yearBE = yearCE + 543;
    
    // กำหนดภาคเรียนตามปีการศึกษา
    let semester = "";
    if (month >= 5 && month <= 10) {
      semester = "ภาคเรียนที่ 1";
    } else if (month >= 11 && month <= 12) {
      semester = "ภาคเรียนที่ 2";
    } else if (month >= 1 && month <= 3) {
      semester = "ภาคเรียนที่ 2";
    } else {
      semester = "ช่วงพิเศษ";
    }
    
    // สร้าง HTML และ PDF (ใช้โค้ดเดิมที่มีอยู่)
    const html = createMonthlyPDFHTML(data, grade, classNo, monthName, yearBE, academicYear, semester);
    
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(`รายงานรายเดือน_${grade}_ห้อง${classNo}_${monthName}${yearBE}_${semester}.pdf`)
      .getAs('application/pdf');
    
    const file2 = _saveBlobGetUrl_(blob);
    Logger.log(`✅ Monthly PDF created: ${file2}`);
    return file2;
    
  } catch (error) {
    Logger.log(`❌ Error in exportMonthlyAttendancePDF (${month}/${yearCE}):`, error.message);
    throw new Error(`ไม่สามารถสร้าง PDF เดือน ${month}/${yearCE} ได้: ${error.message}`);
  }
}

/**
 * ✅ ส่งออก PDF รายปี - รองรับระบบปีการศึกษา
 */
function exportYearlyAttendancePDF(grade, classNo, academicYear) {
  try {
    Logger.log(`🔍 exportYearlyAttendancePDF: ${grade}/${classNo}/${academicYear}`);
    
    // ใช้ฟังก์ชันที่แก้ไขแล้ว
    const data = getYearlyAttendanceSummary(grade, classNo, academicYear);
    
    if (!data || data.length === 0) {
      throw new Error(`ไม่พบข้อมูลสำหรับชั้น ${grade} ห้อง ${classNo} ปีการศึกษา ${academicYear}`);
    }
    
    // สร้าง HTML และ PDF (ใช้โค้ดเดิมที่มีอยู่)
    const html = createYearlyPDFHTML(data, grade, classNo, academicYear);
    
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(`รายงานสรุปการมาเรียน_${grade}_ห้อง${classNo}_ปี${academicYear}.pdf`)
      .getAs('application/pdf');
    
    const yearlyUrl = _saveBlobGetUrl_(blob);
    Logger.log(`✅ Yearly PDF created successfully: ${yearlyUrl}`);
    return yearlyUrl;
    
  } catch (error) {
    Logger.log("❌ Error in exportYearlyAttendancePDF:", error.message);
    throw new Error(`ไม่สามารถสร้าง PDF ได้: ${error.message}`);
  }
}

/**
 * ✅ สร้าง PDF เปล่าสำหรับเดือนที่ไม่มีข้อมูล (รูปแบบเป็นทางการ)
 */
function createEmptyMonthPDF(grade, classNo, year, month) {
  try {
    const thaiMonths = [
      'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
      'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
    ];
    
    const monthName = thaiMonths[month - 1];
    const yearBE = year + 543;
    
    // กำหนดภาคเรียน
    let semester = "";
    let academicYear = "";
    if (month >= 5 && month <= 10) {
      semester = "ภาคเรียนที่ 1";
      academicYear = yearBE;
    } else if (month >= 11 && month <= 12) {
      semester = "ภาคเรียนที่ 2";
      academicYear = yearBE;
    } else if (month >= 1 && month <= 3) {
      semester = "ภาคเรียนที่ 2";
      academicYear = yearBE - 1;
    }
    
    const html = `
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { 
              font-family: 'Sarabun', 'THSarabunNew', Arial, sans-serif; 
              margin: 30px; 
              line-height: 1.4;
            }
            .header {
              text-align: center;
              margin-bottom: 25px;
              border-bottom: 2px solid #1a237e;
              padding-bottom: 15px;
            }
            .header h1 { 
              color: #1a237e; 
              margin: 0 0 8px 0;
              font-size: 20px;
              font-weight: bold;
            }
            .header .info { 
              color: #444; 
              font-size: 16px;
              margin: 5px 0;
            }
            .academic-info {
              background: #f8f9fa;
              padding: 8px 15px;
              border-radius: 4px;
              text-align: center;
              margin: 15px 0;
              font-size: 13px;
              color: #1565c0;
              border-left: 4px solid #1565c0;
            }
            .no-data {
              margin: 50px 0;
              padding: 40px;
              background: #f8f9fa;
              border: 1px solid #dee2e6;
              border-radius: 8px;
              text-align: center;
            }
            .no-data h3 {
              color: #6c757d;
              margin-bottom: 15px;
            }
            .no-data p {
              color: #868e96;
              margin: 10px 0;
            }
            .footer {
              margin-top: 40px;
              text-align: center;
              font-size: 12px;
              color: #666;
              border-top: 1px solid #ddd;
              padding-top: 15px;
            }
          </style>
        </head>
        <body>
          <div class="header">
            <h1>รายงานสรุปการมาเรียนรายเดือน</h1>
            <div class="info">ชั้น ${grade} ห้อง ${classNo} | ${monthName} ${yearBE}</div>
          </div>
          
          <div class="academic-info">
            <strong>${semester}</strong> | ปีการศึกษา ${academicYear}
          </div>
          
          <div class="no-data">
            <h3>ไม่พบข้อมูลการเข้าเรียน</h3>
            <p>ยังไม่มีการบันทึกข้อมูลการเข้าเรียนในเดือนนี้</p>
            <p>หรือยังไม่มีนักเรียนในชั้น/ห้องที่เลือก</p>
          </div>
          
          <div class="footer">
            <p>${semester} ปีการศึกษา ${academicYear}</p>
          </div>
        </body>
      </html>
    `;
    
    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(`รายงานรายเดือน_${grade}_ห้อง${classNo}_${monthName}${yearBE}_ไม่มีข้อมูล.pdf`)
      .getAs('application/pdf');
    
    return _saveBlobGetUrl_(blob);
    
  } catch (error) {
    Logger.log("Error creating empty PDF:", error.message);
    throw error;
  }
}

/**
 * ✅ สร้าง PDF Blob โดยตรงสำหรับรายเดือน (ไม่บันทึกลง Drive) - สำหรับ ZIP
 */
function createMonthlyPDFBlob(grade, classNo, year, month) {
  try {
    // ดึงข้อมูลจากฟังก์ชันที่มีอยู่แล้ว
    const data = getMonthlyAttendanceSummary(grade, classNo, year, month);
    
    const thaiMonths = [
      'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
      'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
    ];
    
    const monthName = thaiMonths[month - 1];
    const yearBE = year + 543;
    
    // กำหนดภาคเรียน
    let semester = "";
    let academicYear = "";
    if (month >= 5 && month <= 10) {
      semester = "ภาคเรียนที่ 1";
      academicYear = yearBE;
    } else if (month >= 11 && month <= 12) {
      semester = "ภาคเรียนที่ 2";
      academicYear = yearBE;
    } else if (month >= 1 && month <= 3) {
      semester = "ภาคเรียนที่ 2";
      academicYear = yearBE - 1;
    }
    
    let html = "";
    
    if (!data || data.length === 0) {
      // สร้าง HTML สำหรับไฟล์เปล่า
      html = `
        <html>
          <head>
            <meta charset="utf-8">
            <style>
              body { 
                font-family: 'Sarabun', 'THSarabunNew', Arial, sans-serif; 
                margin: 30px; 
                line-height: 1.4;
              }
              .header {
                text-align: center;
                margin-bottom: 25px;
                border-bottom: 2px solid #1a237e;
                padding-bottom: 15px;
              }
              .header h1 { 
                color: #1a237e; 
                margin: 0 0 8px 0;
                font-size: 20px;
                font-weight: bold;
              }
              .header .info { 
                color: #444; 
                font-size: 16px;
                margin: 5px 0;
              }
              .academic-info {
                background: #f8f9fa;
                padding: 8px 15px;
                border-radius: 4px;
                text-align: center;
                margin: 15px 0;
                font-size: 13px;
                color: #1565c0;
                border-left: 4px solid #1565c0;
              }
              .no-data {
                margin: 50px 0;
                padding: 40px;
                background: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                text-align: center;
              }
              .no-data h3 {
                color: #6c757d;
                margin-bottom: 15px;
              }
              .no-data p {
                color: #868e96;
                margin: 10px 0;
              }
              .footer {
                margin-top: 40px;
                text-align: center;
                font-size: 12px;
                color: #666;
                border-top: 1px solid #ddd;
                padding-top: 15px;
              }
            </style>
          </head>
          <body>
            <div class="header">
              <h1>รายงานสรุปการมาเรียนรายเดือน</h1>
              <div class="info">ชั้น ${grade} ห้อง ${classNo} | ${monthName} ${yearBE}</div>
            </div>
            
            <div class="academic-info">
              <strong>${semester}</strong> | ปีการศึกษา ${academicYear}
            </div>
            
            <div class="no-data">
              <h3>ไม่พบข้อมูลการเข้าเรียน</h3>
              <p>ยังไม่มีการบันทึกข้อมูลการเข้าเรียนในเดือนนี้</p>
              <p>หรือยังไม่มีนักเรียนในชั้น/ห้องที่เลือก</p>
            </div>
            
            <div class="footer">
              <p>${semester} ปีการศึกษา ${academicYear}</p>
            </div>
          </body>
        </html>
      `;
    } else {
      // คำนวณสถิติรวม
      let totalPresent = 0, totalLeave = 0, totalAbsent = 0, totalDays = 0;
      data.forEach(student => {
        totalPresent += parseInt(student.present) || 0;
        totalLeave += parseInt(student.leave) || 0;
        totalAbsent += parseInt(student.absent) || 0;
        totalDays += parseInt(student.total) || 0;
      });
      
      const overallPercentage = totalDays > 0 ? ((totalPresent / totalDays) * 100).toFixed(1) : "0.0";
      
      // สร้าง HTML สำหรับไฟล์ที่มีข้อมูล
      html = `
        <html>
          <head>
            <meta charset="utf-8">
            <style>
              body { 
                font-family: 'Sarabun', 'THSarabunNew', Arial, sans-serif; 
                margin: 30px; 
                font-size: 14px;
                line-height: 1.4;
              }
              .header {
                text-align: center;
                margin-bottom: 25px;
                border-bottom: 2px solid #1a237e;
                padding-bottom: 15px;
              }
              .header h1 { 
                color: #1a237e; 
                margin: 0 0 8px 0;
                font-size: 20px;
                font-weight: bold;
              }
              .header .info { 
                color: #444; 
                font-size: 16px;
                margin: 5px 0;
              }
              .academic-info {
                background: #f8f9fa;
                padding: 8px 15px;
                border-radius: 4px;
                text-align: center;
                margin: 15px 0;
                font-size: 13px;
                color: #1565c0;
                border-left: 4px solid #1565c0;
              }
              .summary-box {
                background: #fafafa;
                padding: 15px;
                border-radius: 6px;
                margin: 20px 0;
                border: 1px solid #e0e0e0;
              }
              .summary-title {
                font-weight: bold;
                text-align: center;
                margin-bottom: 15px;
                color: #333;
                font-size: 15px;
              }
              .summary-stats {
                display: flex;
                justify-content: space-around;
                margin: 10px 0;
              }
              .stat-item {
                text-align: center;
                flex: 1;
              }
              .stat-number {
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 3px;
              }
              .stat-label {
                font-size: 12px;
                color: #666;
              }
              table { 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 20px;
                font-size: 14px;
              }
              th, td { 
                border: 1px solid #888; 
                padding: 8px 6px; 
                text-align: center; 
              }
              th { 
                background: #f5f5f5; 
                font-weight: bold;
                color: #333;
              }
              .text-start { text-align: left; }
              .footer {
                margin-top: 25px;
                padding-top: 15px;
                border-top: 1px solid #ddd;
                text-align: center;
                font-size: 12px;
                color: #666;
              }
              .status-pass { color: #2e7d32; font-weight: bold; }
              .status-fail { color: #c62828; font-weight: bold; }
            </style>
          </head>
          <body>
            <div class="header">
              <h1>รายงานสรุปการมาเรียนรายเดือน</h1>
              <div class="info">ชั้น ${grade} ห้อง ${classNo} | ${monthName} ${yearBE}</div>
            </div>
            
            <div class="academic-info">
              <strong>${semester}</strong> | ปีการศึกษา ${academicYear}
            </div>
            
            <div class="summary-box">
              <div class="summary-title">สถิติการเข้าเรียน</div>
              <div class="summary-stats">
                <div class="stat-item">
                  <div class="stat-number" style="color: #2e7d32;">${totalPresent}</div>
                  <div class="stat-label">วันมาเรียน</div>
                </div>
                <div class="stat-item">
                  <div class="stat-number" style="color: #f57c00;">${totalLeave}</div>
                  <div class="stat-label">วันลา</div>
                </div>
                <div class="stat-item">
                  <div class="stat-number" style="color: #c62828;">${totalAbsent}</div>
                  <div class="stat-label">วันขาด</div>
                </div>
                <div class="stat-item">
                  <div class="stat-number" style="color: #1565c0;">${overallPercentage}%</div>
                  <div class="stat-label">เฉลี่ยการมาเรียน</div>
                </div>
              </div>
            </div>
            
            <table>
              <thead>
                <tr>
                  <th style="width: 8%;">เลขที่</th>
                  <th style="width: 12%;">รหัสนักเรียน</th>
                  <th class="text-start" style="width: 35%;">ชื่อ-นามสกุล</th>
                  <th style="width: 10%;">มา</th>
                  <th style="width: 10%;">ลา</th>
                  <th style="width: 10%;">ขาด</th>
                  <th style="width: 10%;">รวมวันเรียน</th>
                  <th style="width: 15%;">ร้อยละมาเรียน</th>
                </tr>
              </thead>
              <tbody>
      `;
      
      data.forEach((student, index) => {
        const present = parseInt(student.present) || 0;
        const leave = parseInt(student.leave) || 0;
        const absent = parseInt(student.absent) || 0;
        const total = parseInt(student.total) || (present + leave + absent);
        
        const percentage = total > 0 ? ((present / total) * 100).toFixed(1) : "0.0";
        const percentNum = parseFloat(percentage);
        const statusClass = percentNum >= 80 ? "status-pass" : "status-fail";
        
        html += `
          <tr>
            <td>${index + 1}</td>
            <td>${student.studentId}</td>
            <td class="text-start">${student.name || '-'}</td>
            <td>${present}</td>
            <td>${leave}</td>
            <td>${absent}</td>
            <td>${total}</td>
            <td class="${statusClass}"><strong>${percentage}%</strong></td>
          </tr>
        `;
      });
      
      html += `
              </tbody>
            </table>
                        
          </body>
        </html>
      `;
    }
    
    // สร้าง PDF blob
    const pdfBlob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf');
    
    return pdfBlob;
    
  } catch (error) {
    Logger.log(`Error creating PDF blob for ${month}/${year}:`, error.message);
    return null;
  }
}

function createZipFromResults(grade, classNo, academicYear, successResults) {
  try {
    Logger.log(`🔍 createZipFromResults: ${grade}/${classNo}/${academicYear} - ${successResults.length} files`);
    
    if (!successResults || successResults.length === 0) {
      throw new Error("ไม่มีไฟล์ที่จะรวมเป็น ZIP");
    }
    
    const blobs = [];
    const fileNames = [];
    
    // ดาวน์โหลดและเตรียม blobs สำหรับแต่ละ PDF
    successResults.forEach((result, index) => {
      try {
        // ดึง PDF จาก URL
        const response = UrlFetchApp.fetch(result.url);
        const blob = response.getBlob();
        
        // สร้างชื่อไฟล์ที่เป็นระเบียบ
        const paddedIndex = String(index + 1).padStart(2, '0');
        const semester = result.semester;
        const fileName = `${paddedIndex}_${result.displayName}_ภาคเรียนที่${semester}.pdf`;
        
        blob.setName(fileName);
        blobs.push(blob);
        fileNames.push(fileName);
        
        Logger.log(`✅ Added to ZIP: ${fileName}`);
        
      } catch (error) {
        Logger.log(`❌ Error processing ${result.displayName}:`, error.message);
        // ข้ามไฟล์ที่มีปัญหาและดำเนินการต่อ
      }
    });
    
    if (blobs.length === 0) {
      throw new Error("ไม่สามารถเตรียมไฟล์สำหรับ ZIP ได้");
    }
    
    // สร้าง ZIP
    const zipBlob = Utilities.zip(blobs, `รายงานการเข้าเรียน_${grade}_ห้อง${classNo}_ปีการศึกษา${academicYear}.zip`);
    
    // บันทึกลง Google Drive
    const zipUrl3 = _saveBlobGetUrl_(zipBlob);
    
    Logger.log(`✅ ZIP created successfully with ${blobs.length} files: ${zipUrl3}`);
    
    return {
      url: zipUrl3,
      fileName: zipBlob.getName(),
      fileCount: blobs.length,
      totalFiles: successResults.length,
      fileNames: fileNames
    };
    
  } catch (error) {
    Logger.log("❌ Error in createZipFromResults:", error.message);
    throw new Error(`ไม่สามารถสร้างไฟล์ ZIP ได้: ${error.message}`);
  }
}


/**
 * Adapter สำหรับฝั่งหน้าเว็บ
 * รับ payload จาก AttendanceMonthly.exportMonthlyPDF() แล้วเรียกฟังก์ชันหลักที่มีอยู่
 * payload: { grade, classNo, year, monthIndex1, ... }
 * คืนค่า: URL ไฟล์ PDF (string) หรือ object { url, fileId, message }
 */
function generateMonthlyAttendancePDF(payload) {
  if (!payload) throw new Error('payload is required');
  const grade = String(payload.grade || '').trim();
  const classNo = String(payload.classNo || '').trim();
  const year = Number(payload.year);          // ปี ค.ศ.
  const month = Number(payload.monthIndex1);  // 1–12

  if (!grade || !classNo || !year || !month) {
    throw new Error('ข้อมูลไม่ครบ: ต้องมี grade, classNo, year (ค.ศ.), month (1–12)');
  }

  // เรียกฟังก์ชันหลักของคุณที่มีอยู่แล้ว
  const url = exportMonthlyAttendancePDF(grade, classNo, year, month);

  // ส่งกลับเป็น string หรือ object ก็ได้ (frontend รองรับทั้งคู่)
  return { url, message: `สร้าง PDF สำเร็จ: ${grade}/${classNo} ${month}/${year}` };
}

// ✅ [เวอร์ชันแก้ไขล่าสุด] ลบ Ui().alert() ออกแล้ว
function addSummaryFormulasToAllSheets() {
  const ss = SS();
  const allSheets = ss.getSheets();
  const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  let updatedSheetsCount = 0;

  // ใช้ console.log แทน Ui().alert()
  Logger.log("🚀 เริ่มกระบวนการเพิ่มสูตรสรุปผลอัตโนมัติ (เวอร์ชันไดนามิก)...");

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const isMonthlySheet = monthNames.some(month => sheetName.startsWith(month));

    if (isMonthlySheet) {
      try {
        Logger.log(`- - - กำลังประมวลผลชีต: ${sheetName} - - -`);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          Logger.log(`  -> ข้าม: ไม่มีข้อมูลนักเรียน`);
          return;
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const idxPresent = headers.indexOf("มา");
        const idxLeave = headers.indexOf("ลา");
        const idxAbsent = headers.indexOf("ขาด");

        if (idxPresent === -1 || idxLeave === -1 || idxAbsent === -1) {
          Logger.log(`  -> ⚠️ ข้าม: ไม่พบคอลัมน์ มา/ลา/ขาด`);
          return;
        }

        const startDataCol = headers.indexOf("ชื่อ-สกุล") + 1;
        const endDataCol = idxPresent;
        
        if (startDataCol >= endDataCol) {
          Logger.log(`  -> ⚠️ ข้าม: ไม่สามารถระบุช่วงของวันได้`);
          return;
        }

        const startColLetter = sheet.getRange(1, startDataCol + 1).getA1Notation().replace("1", "");
        const endColLetter = sheet.getRange(1, endDataCol).getA1Notation().replace("1", "");
        
        const rangeForRow2 = `${startColLetter}2:${endColLetter}2`;
        const presentFormula = `=COUNTIF(${rangeForRow2}, "/")`;
        const leaveFormula = `=COUNTIF(${rangeForRow2}, "ล")`;
        const absentFormula = `=COUNTIF(${rangeForRow2}, "ข")`;
        
        Logger.log(`  -> สูตรสำหรับชีตนี้: ${presentFormula}`);

        sheet.getRange(2, idxPresent + 1).setFormula(presentFormula);
        sheet.getRange(2, idxLeave + 1).setFormula(leaveFormula);
        sheet.getRange(2, idxAbsent + 1).setFormula(absentFormula);

        const sourceRange = sheet.getRange(2, idxPresent + 1, 1, 3);
        const destinationRange = sheet.getRange(2, idxPresent + 1, lastRow - 1, 3);
        sourceRange.autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
        
        Logger.log(`  -> ✅ เพิ่มสูตรในชีต ${sheetName} สำเร็จ`);
        updatedSheetsCount++;
      } catch (e) {
        Logger.log(`  -> ❌ เกิดข้อผิดพลาดในชีต ${sheetName}: ${e.message}`);
      }
    }
  });

  // ใช้ console.log แทน Ui().alert()
  Logger.log(`🎉 ทำรายการเสร็จสิ้น! เพิ่ม/อัปเดตสูตรสรุปผลในชีตรายเดือนทั้งหมด ${updatedSheetsCount} ชีต`);
}

/**
 * ✅ [โค้ดอัปเกรดล่าสุด] แก้ไขสูตร COUNTIF ในทุกชีตรายเดือน
 * เวอร์ชันนี้จะค้นหา "คอลัมน์วันที่สุดท้าย" จริงๆ ทำให้แม่นยำกว่าเดิม
 * และจะสร้างสูตรที่นับเฉพาะ "/", "ล", "ข" เท่านั้น
 */
function fixAllSheetFormulas() {
  const ss = SS();
  const allSheets = ss.getSheets();
  const monthNames = [
    'มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'
  ];
  let updatedSheetsCount = 0;
  
  Logger.log("🚀 [v2] เริ่มกระบวนการแก้ไขสูตรสรุปผลในทุกชีต...");

  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // ตรวจสอบว่าเป็นชีตรายเดือนหรือไม่
    const isMonthlySheet = monthNames.some(month => sheetName.startsWith(month));

    if (isMonthlySheet) {
      try {
        Logger.log(`- - - ⚙️ กำลังประมวลผลชีต: "${sheetName}" - - -`);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          Logger.log(`  -> ⏭️ ข้าม: ไม่มีข้อมูลนักเรียน`);
          return;
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const nameColIdx = headers.findIndex(h => String(h).includes("ชื่อ-สกุล"));
        const presentColIdx = headers.findIndex(h => String(h).startsWith("มา"));
        const leaveColIdx = headers.findIndex(h => String(h).startsWith("ลา"));
        const absentColIdx = headers.findIndex(h => String(h).startsWith("ขาด"));
        
        if (nameColIdx === -1 || presentColIdx === -1) {
          Logger.log(`  -> ⚠️ ข้าม: ไม่พบคอลัมน์ "ชื่อ-สกุล" หรือ "มา"`);
          return;
        }

        // ค้นหาคอลัมน์วันที่สุดท้ายจริงๆ โดยไล่จากขวามาซ้าย
        let lastDayColIdx = -1;
        for (let i = presentColIdx - 1; i > nameColIdx; i--) {
          const headerText = String(headers[i]).trim();
          if (/^\d/.test(headerText)) {
            lastDayColIdx = i;
            break; 
          }
        }
        
        if (lastDayColIdx === -1) {
          Logger.log(`  -> ⚠️ ข้าม: ไม่พบคอลัมน์วันที่ที่ถูกต้อง`);
          return;
        }

        const startDataCol = nameColIdx + 1;
        
        const startColLetter = sheet.getRange(1, startDataCol + 1).getA1Notation().replace(/\d+/, '');
        const endColLetter = sheet.getRange(1, lastDayColIdx + 1).getA1Notation().replace(/\d+/, '');
        
        Logger.log(`  -> 🎯 ช่วงข้อมูลวันที่ที่พบ: ${startColLetter} ถึง ${endColLetter}`);

        const formulas = [];
        for (let row = 2; row <= lastRow; row++) {
          const range = `${startColLetter}${row}:${endColLetter}${row}`;
          // --- สร้างสูตรที่นับเฉพาะเครื่องหมาย ---
          const presentFormula = `=COUNTIF(${range}, "/")`; 
          const leaveFormula = `=COUNTIF(${range}, "ล")`;
          const absentFormula = `=COUNTIF(${range}, "ข")`;
          formulas.push([presentFormula, leaveFormula, absentFormula]);
        }
        
        // เขียนสูตรทั้งหมดลงชีตในครั้งเดียว
        sheet.getRange(2, presentColIdx + 1, formulas.length, 3).setFormulas(formulas);
        
        Logger.log(`  -> ✅ แก้ไขสูตรในชีต "${sheetName}" จำนวน ${formulas.length} แถว สำเร็จ!`);
        updatedSheetsCount++;

      } catch (e) {
        Logger.log(`  -> ❌ เกิดข้อผิดพลาดในชีต "${sheetName}": ${e.message}`);
      }
    }
  });
  
  const summaryMessage = `🎉🎉🎉 กระบวนการเสร็จสิ้น! แก้ไขสูตรในชีตทั้งหมด ${updatedSheetsCount} ชีตเรียบร้อยแล้ว`;
  Logger.log(summaryMessage);
  
  if (typeof SpreadsheetApp.getUi === 'function') {
    SpreadsheetApp.getUi().alert('แก้ไขสูตรสำเร็จ!', summaryMessage, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================
// 📋 โค้ดแก้ไข PDF - ตัวหนังสือดำสนิท พื้นหลังขาวทั้งหมด
// ============================================================
// วิธีใช้: 
// 1. ค้นหาฟังก์ชัน createMonthlyPDFHTML, createYearlyPDFHTML, createRealChecklistHTML ในโค้ดเดิม
// 2. ลบฟังก์ชันเหล่านั้นทิ้ง
// 3. วางโค้ดนี้แทน
// ============================================================

/**
 * ✅ สร้าง HTML สำหรับ PDF รายเดือน (ตัวหนังสือดำ พื้นขาว)
 */
function createMonthlyPDFHTML(data, grade, classNo, monthName, yearBE, academicYear, semester) {
  // คำนวณสถิติรวม
  let totalPresentInstances = 0;
  let totalLeaveInstances = 0;
  let totalAbsentInstances = 0;
  let schoolDaysInMonth = 0;
  let totalPossibleAttendance = 0;

  if (data && data.length > 0) {
    schoolDaysInMonth = parseInt(data[0].total) || 0;
    data.forEach(student => {
      totalPresentInstances += parseInt(student.present) || 0;
      totalLeaveInstances += parseInt(student.leave) || 0;
      totalAbsentInstances += parseInt(student.absent) || 0;
    });
    totalPossibleAttendance = schoolDaysInMonth * data.length;
  }
  
  const overallPercentage = totalPossibleAttendance > 0 
    ? ((totalPresentInstances / totalPossibleAttendance) * 100).toFixed(1) 
    : "0.0";
  
  let html = `
    <!DOCTYPE html>
    <html lang="th">
      <head>
        <meta charset="utf-8">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
          
          @page { 
            size: A4 portrait; 
            margin: 15mm 10mm 15mm 15mm;
          }
          
          * {
            box-sizing: border-box;
            font-family: 'Sarabun', 'Tahoma', 'Arial', sans-serif !important;
            color: #000 !important;
          }
          
          body { 
            font-family: 'Sarabun', 'Tahoma', 'Arial', sans-serif !important;
            font-size: 14px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
            margin: 0; 
            padding: 10px;
            line-height: 1.5;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          
          .pdf-header {
            text-align: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #000;
            background: #fff !important;
          }
          
          .pdf-header h1 {
            font-size: 20px;
            font-weight: 700;
            color: #000 !important;
            margin: 0 0 8px 0;
          }
          
          .pdf-header h2 {
            font-size: 16px;
            font-weight: 600;
            color: #000 !important;
            margin: 0 0 5px 0;
          }
          
          .pdf-header .info {
            font-size: 14px;
            font-weight: 400;
            color: #000 !important;
            margin: 3px 0;
          }
          
          .academic-info {
            background: #fff !important;
            padding: 8px 15px;
            border: 1px solid #000;
            text-align: center;
            margin: 10px 0 15px 0;
            font-size: 13px;
            font-weight: 500;
            color: #000 !important;
          }
          
          table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 15px;
            background: #fff !important;
          }
          
          th, td { 
            border: 1px solid #000 !important;
            padding: 8px 6px; 
            text-align: center; 
            vertical-align: middle;
            font-size: 13px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
          }
          
          th { 
            background: #fff !important;
            font-weight: 700;
            color: #000 !important;
          }
          
          .col-no { 
            width: 40px; 
            text-align: center;
          }
          
          .col-id { 
            width: 80px; 
            text-align: center;
            font-size: 12px;
          }
          
          .col-name { 
            text-align: left; 
            padding-left: 10px !important;
            font-size: 13px;
            font-weight: 400;
            min-width: 150px;
          }
          
          .col-name-header {
            text-align: center;
            font-weight: 700;
          }
          
          tbody tr {
            background: #fff !important;
          }
          
          .summary-box {
            background: #fff !important;
            padding: 15px;
            margin: 20px 0;
            border: 1px solid #000;
          }
          
          .summary-title {
            font-weight: 700;
            font-size: 15px;
            text-align: center;
            margin-bottom: 15px;
            color: #000 !important;
          }
          
          .summary-stats {
            display: flex;
            justify-content: space-around;
            flex-wrap: wrap;
            gap: 10px;
          }
          
          .stat-item {
            text-align: center;
            flex: 1;
            min-width: 80px;
          }
          
          .stat-number {
            font-size: 18px;
            font-weight: 700;
            margin-bottom: 3px;
            color: #000 !important;
          }
          
          .stat-label {
            font-size: 12px;
            font-weight: 400;
            color: #000 !important;
          }
          
          .pdf-footer {
            margin-top: 20px;
            padding-top: 10px;
            border-top: 1px solid #000;
            text-align: center;
            font-size: 11px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
          }
        </style>
      </head>
      <body>
        <div class="pdf-header">
          <h1>รายงานสรุปการมาเรียนรายเดือน</h1>
          <h2>ชั้น ${grade} ห้อง ${classNo}</h2>
          <div class="info">${monthName} ${yearBE}</div>
        </div>
        
        <div class="academic-info">
          <strong>${semester}</strong> | ปีการศึกษา ${academicYear}
        </div>
        
        <table>
          <thead>
            <tr>
              <th class="col-no">เลขที่</th>
              <th class="col-id">รหัสนักเรียน</th>
              <th class="col-name-header" style="width: 35%;">ชื่อ-นามสกุล</th>
              <th style="width: 8%;">มา</th>
              <th style="width: 8%;">ลา</th>
              <th style="width: 8%;">ขาด</th>
              <th style="width: 10%;">รวมวัน</th>
              <th style="width: 12%;">ร้อยละ</th>
            </tr>
          </thead>
          <tbody>
  `;
  
  data.forEach((student, index) => {
    const present = parseInt(student.present) || 0;
    const leave = parseInt(student.leave) || 0;
    const absent = parseInt(student.absent) || 0;
    const total = parseInt(student.total) || (present + leave + absent);
    
    const percentage = total > 0 ? ((present / total) * 100).toFixed(1) : "0.0";
    
    const displayName = (student.name || '-').trim();
    
    html += `
      <tr>
        <td class="col-no">${index + 1}</td>
        <td class="col-id">${student.studentId || '-'}</td>
        <td class="col-name">${displayName}</td>
        <td>${present}</td>
        <td>${leave}</td>
        <td>${absent}</td>
        <td>${total}</td>
        <td><strong>${percentage}%</strong></td>
      </tr>
    `;
  });
  
  html += `
          </tbody>
        </table>
        
        <div class="summary-box">
          <div class="summary-title">สถิติการเข้าเรียน (ภาพรวมทั้งห้อง)</div>
          <div class="summary-stats">
            <div class="stat-item">
              <div class="stat-number">${schoolDaysInMonth}</div>
              <div class="stat-label">จำนวนวันเรียน</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${totalLeaveInstances}</div>
              <div class="stat-label">รวมวันลา (ครั้ง)</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${totalAbsentInstances}</div>
              <div class="stat-label">รวมวันขาด (ครั้ง)</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${overallPercentage}%</div>
              <div class="stat-label">เฉลี่ยการมาเรียน</div>
            </div>
          </div>
        </div>

        <div class="pdf-footer">
          <p>พิมพ์เมื่อ: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss น.')}</p>
        </div>
      </body>
    </html>
  `;
  
  return html;
}


/**
 * ✅ สร้าง HTML สำหรับ PDF รายปี (ตัวหนังสือดำ พื้นขาว)
 */
function createYearlyPDFHTML(data, grade, classNo, academicYear) {
  // คำนวณสถิติรวม
  let totalPresent = 0, totalLeave = 0, totalAbsent = 0, totalDays = 0;
  data.forEach(student => {
    totalPresent += parseInt(student.present) || 0;
    totalLeave += parseInt(student.leave) || 0;
    totalAbsent += parseInt(student.absent) || 0;
  });
  totalDays = totalPresent + totalLeave + totalAbsent;
  const overallPercentage = totalDays > 0 ? ((totalPresent / totalDays) * 100).toFixed(1) : "0.0";

  let html = `
    <!DOCTYPE html>
    <html lang="th">
      <head>
        <meta charset="utf-8">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');
          
          @page { 
            size: A4 portrait; 
            margin: 15mm 10mm 15mm 15mm;
          }
          
          * {
            box-sizing: border-box;
            font-family: 'Sarabun', 'Tahoma', 'Arial', sans-serif !important;
            color: #000 !important;
          }
          
          body { 
            font-family: 'Sarabun', 'Tahoma', 'Arial', sans-serif !important;
            font-size: 14px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
            margin: 0; 
            padding: 10px;
            line-height: 1.5;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          
          .pdf-header {
            text-align: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #000;
            background: #fff !important;
          }
          
          .pdf-header h1 {
            font-size: 20px;
            font-weight: 700;
            color: #000 !important;
            margin: 0 0 8px 0;
          }
          
          .pdf-header h2 {
            font-size: 16px;
            font-weight: 600;
            color: #000 !important;
            margin: 0 0 5px 0;
          }
          
          .pdf-header .info {
            font-size: 14px;
            font-weight: 400;
            color: #000 !important;
            margin: 3px 0;
          }
          
          table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 15px;
            background: #fff !important;
          }
          
          th, td { 
            border: 1px solid #000 !important;
            padding: 8px 6px; 
            text-align: center; 
            vertical-align: middle;
            font-size: 13px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
          }
          
          th { 
            background: #fff !important;
            font-weight: 700;
            color: #000 !important;
          }
          
          .col-no { 
            width: 40px; 
            text-align: center;
          }
          
          .col-id { 
            width: 80px; 
            text-align: center;
            font-size: 12px;
          }
          
          .col-name { 
            text-align: left; 
            padding-left: 10px !important;
            font-size: 13px;
            font-weight: 400;
            min-width: 150px;
          }
          
          .col-name-header {
            text-align: center;
            font-weight: 700;
          }
          
          tbody tr {
            background: #fff !important;
          }
          
          .summary-box {
            background: #fff !important;
            padding: 15px;
            margin: 20px 0;
            border: 1px solid #000;
          }
          
          .summary-title {
            font-weight: 700;
            font-size: 15px;
            text-align: center;
            margin-bottom: 15px;
            color: #000 !important;
          }
          
          .summary-stats {
            display: flex;
            justify-content: space-around;
            flex-wrap: wrap;
            gap: 10px;
          }
          
          .stat-item {
            text-align: center;
            flex: 1;
            min-width: 80px;
          }
          
          .stat-number {
            font-size: 18px;
            font-weight: 700;
            margin-bottom: 3px;
            color: #000 !important;
          }
          
          .stat-label {
            font-size: 12px;
            font-weight: 400;
            color: #000 !important;
          }
          
          .pdf-footer {
            margin-top: 20px;
            padding-top: 10px;
            border-top: 1px solid #000;
            text-align: center;
            font-size: 11px;
            font-weight: 400;
            color: #000 !important;
            background: #fff !important;
          }
        </style>
      </head>
      <body>
        <div class="pdf-header">
          <h1>รายงานสรุปการมาเรียนประจำปี</h1>
          <h2>ชั้น ${grade} ห้อง ${classNo}</h2>
          <div class="info">ปีการศึกษา ${academicYear}</div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th class="col-no">ลำดับ</th>
              <th class="col-id">รหัส</th>
              <th class="col-name-header" style="width: 35%;">ชื่อ-นามสกุล</th>
              <th style="width: 8%;">มา</th>
              <th style="width: 8%;">ลา</th>
              <th style="width: 8%;">ขาด</th>
              <th style="width: 10%;">รวมวัน</th>
              <th style="width: 12%;">ร้อยละ</th>
            </tr>
          </thead>
          <tbody>
  `;
  
  data.forEach((student, index) => {
    const present = parseInt(student.present) || 0;
    const leave = parseInt(student.leave) || 0;
    const absent = parseInt(student.absent) || 0;
    const total = parseInt(student.total) || (present + leave + absent);
    
    const percentage = total > 0 ? ((present / total) * 100).toFixed(1) : "0.0";
    
    const displayName = (student.name || '-').trim();
    
    html += `
      <tr>
        <td class="col-no">${index + 1}</td>
        <td class="col-id"><strong>${student.studentId || '-'}</strong></td>
        <td class="col-name">${displayName}</td>
        <td>${present}</td>
        <td>${leave}</td>
        <td>${absent}</td>
        <td><strong>${total}</strong></td>
        <td><strong>${percentage}%</strong></td>
      </tr>
    `;
  });
  
  html += `
          </tbody>
        </table>
        
        <div class="summary-box">
          <div class="summary-title">สถิติการเข้าเรียน (ภาพรวมทั้งปีการศึกษา)</div>
          <div class="summary-stats">
            <div class="stat-item">
              <div class="stat-number">${totalPresent}</div>
              <div class="stat-label">วันมาเรียนรวม</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${totalLeave}</div>
              <div class="stat-label">วันลารวม</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${totalAbsent}</div>
              <div class="stat-label">วันขาดรวม</div>
            </div>
            <div class="stat-item">
              <div class="stat-number">${overallPercentage}%</div>
              <div class="stat-label">เฉลี่ยการมาเรียน</div>
            </div>
          </div>
        </div>
        
        <div class="pdf-footer">
          <p>พิมพ์เมื่อ: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss น.')}</p>
        </div>
      </body>
    </html>
  `;
  
  return html;
}



/**
 * ✅ สร้าง HTML ตารางเช็คชื่อ - ขยับขวามาก + ชื่อนักเรียนใหญ่ขึ้น
 */
function createRealChecklistHTML(tableData, grade, classNo, monthName, yearBE, academicYear, semester) {
  const { students, days } = tableData;
  const actualSchoolDays = tableData.actualSchoolDays || days.length;
  const totalDaysInMonth = tableData.totalDaysInMonth || days.length;

  // ดึงการตั้งค่าระบบและโลโก้
  let logoBase64 = '';
  let schoolName = '';
  
  try {
    const settings = getGlobalSettings();
    schoolName = settings['ชื่อโรงเรียน'] || settings['schoolName'] || '';
    
    const logoFileId = settings['logoFileId'];
    if (logoFileId) {
      try {
        const blob = _getFileBlobCompat_(logoFileId);
        const mime = blob.getContentType() || 'image/png';
        const b64 = Utilities.base64Encode(blob.getBytes());
        logoBase64 = `data:${mime};base64,${b64}`;
      } catch (logoError) {
        Logger.log('ไม่สามารถโหลดโลโก้: ' + logoError.message);
      }
    }
  } catch (e) {
    Logger.log('ไม่สามารถดึงการตั้งค่าได้: ' + e.message);
  }

  // คำนวณขนาดคอลัมน์ตามจำนวนวัน
  let dateColWidth, dateColFontSize, dayOfWeekFontSize;
  
  if (actualSchoolDays <= 15) {
    dateColWidth = '28px';
    dateColFontSize = '11px';
    dayOfWeekFontSize = '9px';
  } else if (actualSchoolDays <= 20) {
    dateColWidth = '24px';
    dateColFontSize = '10px';
    dayOfWeekFontSize = '8px';
  } else if (actualSchoolDays <= 25) {
    dateColWidth = '20px';
    dateColFontSize = '9px';
    dayOfWeekFontSize = '7px';
  } else {
    dateColWidth = '18px';
    dateColFontSize = '8px';
    dayOfWeekFontSize = '6px';
  }

  // สร้างหัวตารางวันที่
  let dateHeaders = '';
  let dayHeaders = '';
  days.forEach(day => {
    dateHeaders += `<th class="date-col">${day.label}</th>`;
    dayHeaders += `<th class="day-col">${day.dow}</th>`;
  });

  let logoSection = '';
  if (logoBase64) {
    logoSection = `
      <div class="logo-container">
        <img src="${logoBase64}" alt="โลโก้โรงเรียน" class="school-logo">
      </div>
    `;
  }

  let html = `
    <!DOCTYPE html>
    <html lang="th">
      <head>
        <meta charset="utf-8">
        <style>
          /* ✅ เพิ่ม left margin มากขึ้นเป็น 40mm */
          @page { 
            size: A4 portrait; 
            margin: 10mm 8mm 10mm 40mm;  /* top right bottom left - เพิ่มซ้ายเป็น 40mm */
          }
          
          * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
          }
          
          body { 
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            margin: 0; 
            padding: 8px;
            font-size: 14px; 
            font-weight: normal;
            line-height: 1.4;
            color: #000;
            background: #fff;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          
          .logo-container {
            text-align: center;
            margin-bottom: 8px;
          }
          
          .school-logo {
            max-height: 50px;
            max-width: 50px;
            object-fit: contain;
          }
          
          .header { 
            text-align: center; 
            margin-bottom: 12px; 
            border-bottom: 2px solid #000; 
            padding-bottom: 8px;
          }
          
          .header .school-name {
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 18px;
            font-weight: bold;
            color: #000;
            margin-bottom: 4px;
          }
          
          .header h2 { 
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            margin: 4px 0; 
            font-size: 20px; 
            font-weight: bold; 
            line-height: 1.3; 
            color: #000;
          }
          
          .header .info { 
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 15px; 
            font-weight: normal;
            color: #000;
            margin: 4px 0; 
          }
          
          .header .school-days-info {
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 13px;
            color: #333;
            margin-top: 6px;
            padding: 5px 12px;
            background: #f0f0f0;
            border: 1px solid #ccc;
            display: inline-block;
          }
          
          table { 
            width: 100%; 
            border-collapse: collapse; 
            table-layout: fixed;
            background: #fff;
            margin-top: 8px;
          }
          
          th, td { 
            border: 1px solid #000;
            padding: 4px 3px; 
            text-align: center; 
            vertical-align: middle;
            color: #000;
            background: #fff;
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
          }
          
          th { 
            font-weight: bold;
            background: #fff;
          }
          
          /* ✅ เพิ่มขนาดคอลัมน์เลขที่ */
          .col-no { 
            width: 28px; 
            font-size: 13px;
          }
          
          /* ✅ เพิ่มขนาดคอลัมน์รหัส */
          .col-id { 
            width: 48px; 
            font-size: 13px;
          }
          
          .col-name-header {
            width: 140px;  /* ✅ เพิ่มจาก 100px */
            text-align: center;
            font-size: 14px;  /* ✅ เพิ่มจาก 12px */
            font-weight: bold;
          }
          
          /* ✅ เพิ่มขนาดชื่อนักเรียน */
          .col-name { 
            width: 140px;  /* ✅ เพิ่มจาก 100px */
            text-align: left; 
            padding-left: 6px; 
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 14px;  /* ✅ เพิ่มจาก 12px */
            font-weight: normal;
            line-height: 1.3;  /* ✅ เพิ่มระยะห่างบรรทัด */
          }
          
          .date-col, .day-col { 
            width: ${dateColWidth}; 
            min-width: ${dateColWidth};
            max-width: ${dateColWidth};
            font-size: ${dateColFontSize}; 
            padding: 3px 1px;
          }
          
          .day-col {
            font-size: ${dayOfWeekFontSize};
            font-weight: normal;
          }
          
          .attendance-cell { 
            width: ${dateColWidth}; 
            min-width: ${dateColWidth};
            max-width: ${dateColWidth};
            height: 26px;
            font-size: 13px;
            font-weight: normal;
            padding: 0;
          }
          
          /* ✅ เพิ่มความสูงแถว */
          .student-row td {
            height: 28px;  /* ✅ เพิ่มจาก 24px */
          }
          
          .student-name {
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 14px;  /* ✅ เพิ่มจาก 12px */
            font-weight: normal;
            line-height: 1.3;  /* ✅ เพิ่มระยะห่าง */
            color: #000;
            text-align: left;
            padding-left: 6px;
          }
          
          .student-name span {
            display: block;
            font-family: inherit;
          }
          
          .footer { 
            margin-top: 10px; 
            font-family: 'TH Sarabun New', 'Tahoma', 'Arial', sans-serif;
            font-size: 12px;  /* ✅ เพิ่มจาก 11px */
            font-weight: normal;
            border-top: 1px solid #000; 
            padding-top: 6px; 
            color: #000;
          }
          
          .footer p {
            margin: 3px 0;
          }
          
          .legend {
            display: inline-block;
            margin-right: 15px;
          }
        </style>
      </head>
      <body>
        ${logoSection}
        
        <div class="header">
          ${schoolName ? `<div class="school-name">${schoolName}</div>` : ''}
          <h2>บันทึกเวลาเรียน ชั้น ${grade}/${classNo} เดือน ${monthName} ${yearBE}</h2>
          <div class="info">${semester} | ปีการศึกษา ${academicYear}</div>
          <div class="school-days-info">
            แสดงเฉพาะวันที่มีการเรียน: <strong>${actualSchoolDays}</strong> วัน 
            ${totalDaysInMonth !== actualSchoolDays ? `(จาก ${totalDaysInMonth} วันในเดือน)` : ''}
          </div>
        </div>
        
        <table>
          <thead>
            <tr>
              <th rowspan="2" class="col-no">ที่</th>
              <th rowspan="2" class="col-id">รหัส</th>
              <th rowspan="2" class="col-name-header">ชื่อ - สกุล</th>
              ${dateHeaders}
            </tr>
            <tr>
              ${dayHeaders}
            </tr>
          </thead>
          <tbody>
  `;

  students.forEach((student, index) => {
    // แยกชื่อและนามสกุล
    let firstName = '', lastName = '';
    if (student.name) {
      const nameParts = student.name.trim().split(' ');
      if (nameParts.length <= 2) {
        firstName = nameParts[0] || '';
        lastName = nameParts[1] || '';
      } else {
        lastName = nameParts.pop();
        firstName = nameParts.join(' ');
      }
    }
    
    const studentId = student.id || student.studentId || '-';
    
    html += `
      <tr class="student-row">
        <td class="col-no">${index + 1}</td>
        <td class="col-id">${studentId}</td>
        <td class="col-name student-name">
          <span>${firstName}</span>
          <span>${lastName}</span>
        </td>
    `;
    
    days.forEach(day => {
      const status = student.statusMap[day.label] || '';
      let displayStatus = '';
      
      if (status === '/' || status === 'ม' || status === '1') { 
        displayStatus = '<span style="font-family:Arial,Helvetica,sans-serif;font-size:9px;font-weight:normal;line-height:1;">✓</span>';
      }
      else if (status === 'ล' || status.toLowerCase() === 'l') { 
        displayStatus = 'ล'; 
      }
      else if (status === 'ข' || status === '0') { 
        displayStatus = 'ข'; 
      }
      
      html += `<td class="attendance-cell">${displayStatus}</td>`;
    });
    
    html += `</tr>`;
  });

  html += `
          </tbody>
        </table>
        
        <div class="footer">
          <p>
            <span class="legend"><strong>หมายเหตุ:</strong> ✓ = มาเรียน</span>
            <span class="legend">ล = ลา</span>
            <span class="legend">ข = ขาด</span>
            <span class="legend">| <strong>นักเรียน:</strong> ${students.length} คน</span>
            <span class="legend">| <strong>วันเรียน:</strong> ${actualSchoolDays} วัน</span>
          </p>
          <p><strong>พิมพ์:</strong> ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm น.')}</p>
        </div>
      </body>
    </html>
  `;
  
  return html;
}
// ============================================================
// 🔧 สร้าง PDF ผ่าน Google Sheets - รองรับผสานเซลล์
// ============================================================

/**
 * ✅ ดึงข้อมูลตารางเช็คชื่อ - กรองเฉพาะวันที่มีการเรียน
 */
function getAttendanceVerticalTableFiltered(grade, classNo, year, month) {
  try {
    const ss = SS();
    const sheetName = `${monthNames[month]}${year + 543}`;
    const monthSheet = ss.getSheetByName(sheetName);
    if (!monthSheet) return { students: [], days: [], schoolDaysOnly: true };

    const mData = monthSheet.getDataRange().getValues();
    const mHead = mData[0];
    const idCol = 1;
    const nameCol = 2;

    const allDateCols = mHead.map((h, i) => {
      const match = String(h).match(/^(\d{1,2})/);
      return match ? { col: i, label: match[1] } : null;
    }).filter(Boolean);

    const studentsSheet = ss.getSheetByName("Students");
    const studentsData = studentsSheet.getDataRange().getValues();
    const stuHeaders = studentsData[0];
    const stuIdCol = stuHeaders.indexOf("student_id");
    const stuGradeCol = stuHeaders.indexOf("grade");
    const stuClassCol = stuHeaders.indexOf("class_no");

    const idGradeMap = {};
    studentsData.slice(1).forEach(row => {
      const id = String(row[stuIdCol]).trim();
      idGradeMap[id] = {
        grade: String(row[stuGradeCol]).trim(),
        classNo: String(row[stuClassCol]).trim()
      };
    });

    const schoolDayLabels = new Set();
    
    for (let i = 1; i < mData.length; i++) {
      const row = mData[i];
      const idCell = String(row[idCol]).trim();
      const idMatch = idCell.match(/^(\d+)/);
      const id = idMatch ? idMatch[1] : '';
      
      const info = idGradeMap[id];
      if (!info || info.grade !== grade || info.classNo !== classNo) continue;
      
      allDateCols.forEach(dc => {
        const cellValue = String(row[dc.col] || '').trim();
        if (cellValue === '/' || cellValue === 'ม' || cellValue === '1' || 
            cellValue === 'ล' || cellValue.toLowerCase() === 'l' ||
            cellValue === 'ข' || cellValue === '0' || cellValue === '✓') {
          schoolDayLabels.add(dc.label);
        }
      });
    }

    const filteredDateCols = allDateCols.filter(dc => schoolDayLabels.has(dc.label));

    const days = filteredDateCols.map(dc => {
      const d = parseInt(dc.label);
      const date = new Date(year, month, d);
      const dow = ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"][date.getDay()];
      return { label: `${d}`, dow, col: dc.col };
    });

    const students = mData.slice(1)
      .map(row => {
        const idCell = String(row[idCol]).trim();
        const name = String(row[nameCol]).trim();
        
        const idMatch = idCell.match(/^(\d+)/);
        const id = idMatch ? idMatch[1] : '';
        
        const info = idGradeMap[id];
        if (!info) return null;
        if (info.grade !== grade || info.classNo !== classNo) return null;

        const statusMap = {};
        filteredDateCols.forEach(dc => {
          statusMap[dc.label] = row[dc.col] || '';
        });
        return { id, name, statusMap };
      })
      .filter(stu => stu && stu.id && stu.name);

    students.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));

    return { 
      students, 
      days,
      schoolDaysOnly: true,
      totalDaysInMonth: allDateCols.length,
      actualSchoolDays: filteredDateCols.length
    };

  } catch (e) {
    Logger.log("Error: " + e.message);
    throw new Error("เกิดข้อผิดพลาด: " + e.message);
  }
}

// ============================================================
// 🔧 แก้ไขเครื่องหมายถูกให้เบาลง (ใช้ ✓ แทน ✔)
// ============================================================
// คำแนะนำ: ค้นหาและแทนที่ฟังก์ชันเหล่านี้ในไฟล์ .gs ของคุณ
// ============================================================

/**
 * ✅ สร้าง PDF ตารางเช็คชื่อผ่าน Google Sheets - แก้ไขเครื่องหมายถูกให้เบาลง
 */
function exportAttendanceChecklistPDF(grade, classNo, year, month) {
  let tempSpreadsheet = null;
  
  try {
    Logger.log(`🔍 exportAttendanceChecklistPDF: ${grade}/${classNo}/${year}/${month}`);
    
    const tableData = getAttendanceVerticalTableFiltered(grade, classNo, year, month);
    
    if (!tableData || !tableData.students || tableData.students.length === 0) {
      throw new Error(`ไม่พบข้อมูลสำหรับชั้น ${grade} ห้อง ${classNo}`);
    }
    
    const { students, days } = tableData;
    const actualSchoolDays = tableData.actualSchoolDays || days.length;
    const totalDaysInMonth = tableData.totalDaysInMonth || days.length;
    
    Logger.log(`📊 พบ ${students.length} คน, ${days.length} วัน`);
    
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                        'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthName = thaiMonths[month];
    const yearBE = year + 543;
    
    const academicYear = getAcademicYearFromDate(year, month + 1);
    let semester = (month >= 4 && month <= 9) ? "ภาคเรียนที่ 1" : "ภาคเรียนที่ 2";
    
    let schoolName = '';
    let logoFileId = '';
    try {
      const settings = getGlobalSettings();
      schoolName = settings['ชื่อโรงเรียน'] || settings['schoolName'] || '';
      logoFileId = settings['logoFileId'] || '';
    } catch (e) {}
    
    // ============================================================
    // สร้าง Google Sheets ชั่วคราว
    // ============================================================
    const ssName = `ตารางเช็คชื่อ_${grade}_${classNo}_${monthName}${yearBE}_temp`;
    tempSpreadsheet = SpreadsheetApp.create(ssName);
    const sheet = tempSpreadsheet.getActiveSheet();
    
    const numDays = days.length;
    const numCols = 3 + numDays; // ที่, รหัส, ชื่อ-สกุล, วันที่...
    
    // ============================================================
    // แถวที่ 1-5: หัวเรื่อง
    // ============================================================
    let currentRow = 1;
    
    // แถว 1: โลโก้
    if (logoFileId) {
      try {
        sheet.getRange(currentRow, 1, 1, numCols).merge();
        sheet.setRowHeight(currentRow, 60);
        const logoCell = sheet.getRange(currentRow, 1);
        const logoBlob = _getFileBlobCompat_(logoFileId);
        const logo = sheet.insertImage(logoBlob, 1, currentRow);
        
        const col1Width = 35;
        const col2Width = 60;
        const col3Width = 180;
        const dateColWidth = numDays > 22 ? 20 : (numDays > 18 ? 22 : 24);
        const totalWidth = col1Width + col2Width + col3Width + (numDays * dateColWidth);
        
        logo.setWidth(50);
        logo.setHeight(50);
        logo.setAnchorCell(logoCell);
        logo.setAnchorCellXOffset(Math.floor((totalWidth - 50) / 2));
        logo.setAnchorCellYOffset(5);
        
        currentRow++;
      } catch (e) {
        Logger.log('ไม่สามารถโหลดโลโก้: ' + e.message);
      }
    }
    
    // แถว: ชื่อโรงเรียน
    if (schoolName) {
      sheet.getRange(currentRow, 1, 1, numCols).merge();
      sheet.getRange(currentRow, 1).setValue(schoolName)
        .setFontFamily('Sarabun')
        .setFontSize(16)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.setRowHeight(currentRow, 24);
      currentRow++;
    }
    
    // แถว: หัวเรื่องหลัก
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`บันทึกเวลาเรียน ชั้น ${grade}/${classNo} เดือน ${monthName} ${yearBE}`)
      .setFontFamily('Sarabun')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;
    
    // แถว: ภาคเรียน
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`${semester} | ปีการศึกษา ${academicYear}`)
      .setFontFamily('Sarabun')
      .setFontSize(13)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 22);
    currentRow++;
    
    // แถว: จำนวนวันเรียน
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`แสดงเฉพาะวันที่มีการเรียน: ${actualSchoolDays} วัน (จาก ${totalDaysInMonth} วันในเดือน)`)
      .setFontFamily('Sarabun')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;
    
    // แถวว่าง
    sheet.setRowHeight(currentRow, 12);
    currentRow++;
    
    const headerStartRow = currentRow;
    
    // ============================================================
    // แถวหัวตาราง
    // ============================================================
    
    // แถวที่ 1 ของตาราง: วันที่
    sheet.getRange(currentRow, 1).setValue('ที่');
    sheet.getRange(currentRow, 2).setValue('รหัส');
    sheet.getRange(currentRow, 3).setValue('ชื่อ - สกุล');
    for (let i = 0; i < numDays; i++) {
      sheet.getRange(currentRow, 4 + i).setValue(days[i].label);
    }
    currentRow++;
    
    // แถวที่ 2 ของตาราง: วันในสัปดาห์
    sheet.getRange(currentRow, 1).setValue('');
    sheet.getRange(currentRow, 2).setValue('');
    sheet.getRange(currentRow, 3).setValue('');
    for (let i = 0; i < numDays; i++) {
      sheet.getRange(currentRow, 4 + i).setValue(days[i].dow);
    }
    currentRow++;
    
    // ✅ ผสานเซลล์ "ที่" (2 แถว)
    sheet.getRange(headerStartRow, 1, 2, 1).merge();
    
    // ✅ ผสานเซลล์ "รหัส" (2 แถว)
    sheet.getRange(headerStartRow, 2, 2, 1).merge();
    
    // ✅ ผสานเซลล์ "ชื่อ - สกุล" (2 แถว)
    sheet.getRange(headerStartRow, 3, 2, 1).merge();
    
    // จัดรูปแบบหัวตาราง
    const headerRange = sheet.getRange(headerStartRow, 1, 2, numCols);
    headerRange.setFontFamily('Sarabun')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#f3f3f3')
      .setBorder(true, true, true, true, true, true);
    
    const dataStartRow = currentRow;
    
    // ============================================================
    // แถวข้อมูลนักเรียน - 🔧 แก้ไขเครื่องหมายถูกให้เบาลง
    // ============================================================
    students.forEach((student, index) => {
      sheet.getRange(currentRow, 1).setValue(index + 1);
      sheet.getRange(currentRow, 2).setValue(student.id);
      sheet.getRange(currentRow, 3).setValue(student.name);
      
      for (let i = 0; i < numDays; i++) {
        const status = student.statusMap[days[i].label] || '';
        let displayStatus = '';
        
        // 🔧 แก้ไข: ใช้ ✓ (U+2713 - CHECK MARK) แทน ✔ (U+2714 - HEAVY CHECK MARK)
        if (status === '/' || status === 'ม' || status === '1') {
          displayStatus = '✓';  // ✅ เครื่องหมายถูกแบบเบา (U+2713)
        }
        else if (status === 'ล' || status.toLowerCase() === 'l') {
          displayStatus = 'ล';
        }
        else if (status === 'ข' || status === '0') {
          displayStatus = 'ข';
        }
        
        sheet.getRange(currentRow, 4 + i).setValue(displayStatus);
      }
      currentRow++;
    });
    
    // จัดรูปแบบข้อมูลนักเรียน
    const dataRange = sheet.getRange(dataStartRow, 1, students.length, numCols);
    dataRange.setFontFamily('Sarabun')
      .setFontSize(13)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    
    // คอลัมน์ "ที่" - จัดกลาง
    sheet.getRange(dataStartRow, 1, students.length, 1)
      .setHorizontalAlignment('center');
    
    // คอลัมน์ "รหัส" - จัดกลาง
    sheet.getRange(dataStartRow, 2, students.length, 1)
      .setHorizontalAlignment('center');
    
    // คอลัมน์ "ชื่อ-สกุล" - ชิดซ้ายเล็กน้อย + เพิ่มขนาดฟอนต์
    sheet.getRange(dataStartRow, 3, students.length, 1)
      .setHorizontalAlignment('left')
      .setFontSize(14);
    
    // 🔧 คอลัมน์วันที่ - จัดกลาง + font-weight normal + ขนาดเล็กลง
    sheet.getRange(dataStartRow, 4, students.length, numDays)
      .setHorizontalAlignment('center')
      .setFontFamily('Arial')
      .setFontWeight('normal')
      .setFontSize(9);
    
    // ============================================================
    // ส่วนท้าย
    // ============================================================
    currentRow++;
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`หมายเหตุ: ✓ = มาเรียน   ล = ลา   ข = ขาด   |   นักเรียน: ${students.length} คน   |   วันเรียน: ${actualSchoolDays} วัน`)
      .setFontFamily('Sarabun')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;
    
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`พิมพ์: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm น.')}`)
      .setFontFamily('Sarabun')
      .setFontSize(9)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(currentRow, 18);
    
    // ============================================================
    // กำหนดความกว้างคอลัมน์
    // ============================================================
    sheet.setColumnWidth(1, 35);   // ที่
    sheet.setColumnWidth(2, 60);   // รหัส
    sheet.setColumnWidth(3, 180);  // ชื่อ-สกุล
    
    const dateColWidth = numDays > 22 ? 20 : (numDays > 18 ? 22 : 24);
    for (let i = 0; i < numDays; i++) {
      sheet.setColumnWidth(4 + i, dateColWidth);
    }
    
    // กำหนดความสูงแถว
    sheet.setRowHeight(headerStartRow, 22);
    sheet.setRowHeight(headerStartRow + 1, 20);
    for (let i = 0; i < students.length; i++) {
      sheet.setRowHeight(dataStartRow + i, 26);
    }
    
    // ============================================================
    // Export เป็น PDF
    // ============================================================
    SpreadsheetApp.flush();
    
    const ssId = tempSpreadsheet.getId();
    const sheetId = sheet.getSheetId();
    
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
      `format=pdf&` +
      `size=A4&` +
      `portrait=true&` +
      `fitw=true&` +
      `gridlines=false&` +
      `printtitle=false&` +
      `sheetnames=false&` +
      `pagenum=false&` +
      `fzr=false&` +
      `left_margin=0.8&` +
      `right_margin=0.3&` +
      `top_margin=0.4&` +
      `bottom_margin=0.3&` +
      `gid=${sheetId}`;
    
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const pdfBlob = response.getBlob().setName(`ตารางเช็คชื่อ_${grade}_ห้อง${classNo}_${monthName}${yearBE}.pdf`);
    
    // บันทึก PDF
    const pdfUrl = _saveBlobGetUrl_(pdfBlob);
    
    // ลบ Spreadsheet ชั่วคราว
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch(_) {
      try { var _tk=ScriptApp.getOAuthToken(); UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+ssId,{method:'patch',contentType:'application/json',headers:{Authorization:'Bearer '+_tk},payload:JSON.stringify({trashed:true}),muteHttpExceptions:true}); } catch(_e) {}
    }
    
    Logger.log(`✅ PDF สร้างเสร็จ: ${pdfUrl}`);
    return pdfUrl;
    
  } catch (error) {
    Logger.log(`❌ Error:`, error.message);
    if (tempSpreadsheet) {
      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch (_e2) {
        try { var _tk2=ScriptApp.getOAuthToken(); UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+tempSpreadsheet.getId(),{method:'patch',contentType:'application/json',headers:{Authorization:'Bearer '+_tk2},payload:JSON.stringify({trashed:true}),muteHttpExceptions:true}); } catch(_) {}
      }
    }
    throw new Error(`ไม่สามารถสร้าง PDF ได้: ${error.message}`);
  }
}


/**
 * ✅ สร้าง PDF Blob สำหรับ ZIP - แก้ไขเครื่องหมายถูกให้เบาลง
 */
function createChecklistPDFBlob(grade, classNo, year, month) {
  let tempSpreadsheet = null;
  
  try {
    const monthIndex = month - 1;
    const tableData = getAttendanceVerticalTableFiltered(grade, classNo, year, monthIndex);
    
    if (!tableData || !tableData.students || tableData.students.length === 0) {
      return null;
    }
    
    const { students, days } = tableData;
    const actualSchoolDays = tableData.actualSchoolDays || days.length;
    const totalDaysInMonth = tableData.totalDaysInMonth || days.length;
    
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                        'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthName = thaiMonths[monthIndex];
    const yearBE = year + 543;
    
    const academicYear = getAcademicYearFromDate(year, month);
    let semester = (month >= 5 && month <= 10) ? "ภาคเรียนที่ 1" : "ภาคเรียนที่ 2";
    
    let schoolName = '';
    let logoFileId = '';
    try {
      const settings = getGlobalSettings();
      schoolName = settings['ชื่อโรงเรียน'] || settings['schoolName'] || '';
      logoFileId = settings['logoFileId'] || '';
    } catch (e) {}
    
    // สร้าง Spreadsheet ชั่วคราว
    const ssName = `temp_checklist_${Date.now()}`;
    tempSpreadsheet = SpreadsheetApp.create(ssName);
    const sheet = tempSpreadsheet.getActiveSheet();
    
    const numDays = days.length;
    const numCols = 3 + numDays;
    
    let currentRow = 1;
    
    // โลโก้
    if (logoFileId) {
      try {
        sheet.getRange(currentRow, 1, 1, numCols).merge();
        sheet.setRowHeight(currentRow, 60);
        const logoCell = sheet.getRange(currentRow, 1);
        const logoBlob = _getFileBlobCompat_(logoFileId);
        const logo = sheet.insertImage(logoBlob, 1, currentRow);
        
        const col1Width = 35;
        const col2Width = 60;
        const col3Width = 180;
        const dateColWidth = numDays > 22 ? 20 : (numDays > 18 ? 22 : 24);
        const totalWidth = col1Width + col2Width + col3Width + (numDays * dateColWidth);
        
        logo.setWidth(50);
        logo.setHeight(50);
        logo.setAnchorCell(logoCell);
        logo.setAnchorCellXOffset(Math.floor((totalWidth - 50) / 2));
        logo.setAnchorCellYOffset(5);
        
        currentRow++;
      } catch (e) {}
    }
    
    if (schoolName) {
      sheet.getRange(currentRow, 1, 1, numCols).merge();
      sheet.getRange(currentRow, 1).setValue(schoolName)
        .setFontFamily('Sarabun')
        .setFontSize(16)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
      sheet.setRowHeight(currentRow, 24);
      currentRow++;
    }
    
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`บันทึกเวลาเรียน ชั้น ${grade}/${classNo} เดือน ${monthName} ${yearBE}`)
      .setFontFamily('Sarabun')
      .setFontSize(18)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 28);
    currentRow++;
    
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`${semester} | ปีการศึกษา ${academicYear}`)
      .setFontFamily('Sarabun')
      .setFontSize(13)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 22);
    currentRow++;
    
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`แสดงเฉพาะวันที่มีการเรียน: ${actualSchoolDays} วัน (จาก ${totalDaysInMonth} วันในเดือน)`)
      .setFontFamily('Sarabun')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;
    
    sheet.setRowHeight(currentRow, 12);
    currentRow++;
    
    const headerStartRow = currentRow;
    
    // หัวตาราง
    sheet.getRange(currentRow, 1).setValue('ที่');
    sheet.getRange(currentRow, 2).setValue('รหัส');
    sheet.getRange(currentRow, 3).setValue('ชื่อ - สกุล');
    for (let i = 0; i < numDays; i++) {
      sheet.getRange(currentRow, 4 + i).setValue(days[i].label);
    }
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue('');
    sheet.getRange(currentRow, 2).setValue('');
    sheet.getRange(currentRow, 3).setValue('');
    for (let i = 0; i < numDays; i++) {
      sheet.getRange(currentRow, 4 + i).setValue(days[i].dow);
    }
    currentRow++;
    
    // ผสานเซลล์
    sheet.getRange(headerStartRow, 1, 2, 1).merge();
    sheet.getRange(headerStartRow, 2, 2, 1).merge();
    sheet.getRange(headerStartRow, 3, 2, 1).merge();
    
    const headerRange = sheet.getRange(headerStartRow, 1, 2, numCols);
    headerRange.setFontFamily('Sarabun')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#f3f3f3')
      .setBorder(true, true, true, true, true, true);
    
    const dataStartRow = currentRow;
    
    // ข้อมูลนักเรียน - 🔧 แก้ไขเครื่องหมายถูกให้เบาลง
    students.forEach((student, index) => {
      sheet.getRange(currentRow, 1).setValue(index + 1);
      sheet.getRange(currentRow, 2).setValue(student.id);
      sheet.getRange(currentRow, 3).setValue(student.name);
      
      for (let i = 0; i < numDays; i++) {
        const status = student.statusMap[days[i].label] || '';
        let displayStatus = '';
        
        // 🔧 แก้ไข: ใช้ ✓ (U+2713 - CHECK MARK) แทน ✔ (U+2714 - HEAVY CHECK MARK)
        if (status === '/' || status === 'ม' || status === '1') {
          displayStatus = '✓';  // ✅ เครื่องหมายถูกแบบเบา (U+2713)
        }
        else if (status === 'ล' || status.toLowerCase() === 'l') {
          displayStatus = 'ล';
        }
        else if (status === 'ข' || status === '0') {
          displayStatus = 'ข';
        }
        
        sheet.getRange(currentRow, 4 + i).setValue(displayStatus);
      }
      currentRow++;
    });
    
    const dataRange = sheet.getRange(dataStartRow, 1, students.length, numCols);
    dataRange.setFontFamily('Sarabun')
      .setFontSize(13)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);
    
    sheet.getRange(dataStartRow, 1, students.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataStartRow, 2, students.length, 1).setHorizontalAlignment('center');
    sheet.getRange(dataStartRow, 3, students.length, 1)
      .setHorizontalAlignment('left')
      .setFontSize(14);
    
    // 🔧 ปรับ font weight ให้เบาลง + ลดขนาด
    sheet.getRange(dataStartRow, 4, students.length, numDays)
      .setHorizontalAlignment('center')
      .setFontFamily('Arial')
      .setFontWeight('normal')
      .setFontSize(9);
    
    currentRow++;
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`หมายเหตุ: ✓ = มาเรียน   ล = ลา   ข = ขาด   |   นักเรียน: ${students.length} คน   |   วันเรียน: ${actualSchoolDays} วัน`)
      .setFontFamily('Sarabun')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(currentRow, 20);
    currentRow++;
    
    sheet.getRange(currentRow, 1, 1, numCols).merge();
    sheet.getRange(currentRow, 1).setValue(`พิมพ์: ${Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm น.')}`)
      .setFontFamily('Sarabun')
      .setFontSize(9)
      .setHorizontalAlignment('center');
    sheet.setRowHeight(currentRow, 18);
    
    // กำหนดความกว้างคอลัมน์
    sheet.setColumnWidth(1, 35);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(3, 180);
    const dateColWidth = numDays > 22 ? 20 : (numDays > 18 ? 22 : 24);
    for (let i = 0; i < numDays; i++) {
      sheet.setColumnWidth(4 + i, dateColWidth);
    }
    
    sheet.setRowHeight(headerStartRow, 22);
    sheet.setRowHeight(headerStartRow + 1, 20);
    for (let i = 0; i < students.length; i++) {
      sheet.setRowHeight(dataStartRow + i, 26);
    }
    
    SpreadsheetApp.flush();
    
    const ssId = tempSpreadsheet.getId();
    const sheetId = sheet.getSheetId();
    
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
      `format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&` +
      `printtitle=false&sheetnames=false&pagenum=false&fzr=false&` +
      `left_margin=0.8&right_margin=0.3&top_margin=0.4&bottom_margin=0.3&gid=${sheetId}`;
    
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    const pdfBlob = response.getBlob();
    
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch(_) {
      try { UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+ssId,{method:'patch',contentType:'application/json',headers:{Authorization:'Bearer '+token},payload:JSON.stringify({trashed:true}),muteHttpExceptions:true}); } catch(_e) {}
    }
    
    return pdfBlob;
    
  } catch (error) {
    if (tempSpreadsheet) {
      try { DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); } catch (_e2) {
        try { var _tk3=ScriptApp.getOAuthToken(); UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/'+tempSpreadsheet.getId(),{method:'patch',contentType:'application/json',headers:{Authorization:'Bearer '+_tk3},payload:JSON.stringify({trashed:true}),muteHttpExceptions:true}); } catch(_) {}
      }
    }
    Logger.log(`❌ Error:`, error.message);
    return null;
  }
}

// ============================================================
// 📋 ฟังก์ชันเสริม: ตรวจสอบวันที่มีการเรียนในเดือน
// ============================================================

/**
 * ✅ ดึงรายการวันที่มีการเรียนจริงในเดือน (สำหรับ debug)
 */
function getSchoolDaysInMonth(grade, classNo, year, month) {
  try {
    const ss = SS();
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                        'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const sheetName = `${monthNames[month]}${year + 543}`;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { error: `ไม่พบชีต ${sheetName}`, schoolDays: [] };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // หาคอลัมน์วันที่
    const dateCols = headers.map((h, i) => {
      const match = String(h).match(/^(\d{1,2})/);
      return match ? { col: i, day: parseInt(match[1]) } : null;
    }).filter(Boolean);
    
    // หาวันที่มีข้อมูล
    const schoolDays = new Set();
    
    for (let i = 1; i < data.length; i++) {
      dateCols.forEach(dc => {
        const value = String(data[i][dc.col] || '').trim();
        if (value === '/' || value === 'ม' || value === '1' || 
            value === 'ล' || value.toLowerCase() === 'l' ||
            value === 'ข' || value === '0') {
          schoolDays.add(dc.day);
        }
      });
    }
    
    const sortedDays = Array.from(schoolDays).sort((a, b) => a - b);
    
    return {
      sheetName: sheetName,
      totalDaysInMonth: dateCols.length,
      schoolDaysCount: sortedDays.length,
      schoolDays: sortedDays,
      nonSchoolDays: dateCols.map(dc => dc.day).filter(d => !schoolDays.has(d)).sort((a, b) => a - b)
    };
    
  } catch (error) {
    return { error: error.message, schoolDays: [] };
  }
}


/**
 * 🔧 ฟังก์ชันหลัก: ส่งออกตารางเช็คชื่อทั้งปีการศึกษาเป็น ZIP (ข้อมูลจริง)
 * @param {string} grade - ระดับชั้น
 * @param {string} classNo - ห้อง
 * @param {number} academicYearBE - ปีการศึกษา (พ.ศ.)
 * @returns {string} URL ไฟล์ ZIP
 */
function exportAcademicYearChecklistZIP(grade, classNo, academicYearBE) {
  try {
    Logger.log(`🔍 exportAcademicYearChecklistZIP: ${grade}/${classNo}/${academicYearBE}`);
    
    if (!grade || !classNo || !academicYearBE) {
      throw new Error("พารามิเตอร์ไม่ครบถ้วน");
    }
    
    const pdfBlobs = [];
    const results = [];
    const errors = [];
    const baseYearAD = academicYearBE - 543;
    
    // สร้างรายการเดือนตามปีการศึกษา (พ.ค. - มี.ค.)
    const academicMonths = [
      // ภาคเรียนที่ 1: พ.ค.-ต.ค. (ปีฐาน)
      { month: 5, year: baseYearAD, semester: 1 },
      { month: 6, year: baseYearAD, semester: 1 },
      { month: 7, year: baseYearAD, semester: 1 },
      { month: 8, year: baseYearAD, semester: 1 },
      { month: 9, year: baseYearAD, semester: 1 },
      { month: 10, year: baseYearAD, semester: 1 },
      // ภาคเรียนที่ 2: พ.ย.-ธ.ค. (ปีฐาน)
      { month: 11, year: baseYearAD, semester: 2 },
      { month: 12, year: baseYearAD, semester: 2 },
      // ภาคเรียนที่ 2 (ต่อ): ม.ค.-มี.ค. (ปีถัดไป)
      { month: 1, year: baseYearAD + 1, semester: 2 },
      { month: 2, year: baseYearAD + 1, semester: 2 },
      { month: 3, year: baseYearAD + 1, semester: 2 }
    ];
    
    // สร้าง PDF blob แต่ละเดือน (ข้อมูลจริง)
    academicMonths.forEach((monthInfo, index) => {
      try {
        Logger.log(`Processing real checklist ${monthInfo.month}/${monthInfo.year} (${index + 1}/${academicMonths.length})`);
        
        const pdfBlob = createChecklistPDFBlob(grade, classNo, monthInfo.year, monthInfo.month);
        
        if (pdfBlob) {
          // สร้างชื่อไฟล์ที่เป็นระเบียบ
          const paddedIndex = String(index + 1).padStart(2, '0');
          const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                              'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
          const monthName = thaiMonths[monthInfo.month - 1];
          const yearBE = monthInfo.year + 543;
          const fileName = `${paddedIndex}_ตารางเช็คชื่อ_${monthName}${yearBE}_ภาคเรียนที่${monthInfo.semester}.pdf`;
          
          pdfBlob.setName(fileName);
          pdfBlobs.push(pdfBlob);
          
          results.push({
            month: monthInfo.month,
            year: monthInfo.year,
            semester: monthInfo.semester,
            fileName: fileName,
            success: true
          });
        }
        
        // หน่วงเวลาเล็กน้อย
        Utilities.sleep(300);
        
      } catch (error) {
        errors.push({
          month: monthInfo.month,
          year: monthInfo.year,
          semester: monthInfo.semester,
          error: error.message,
          success: false
        });
        Logger.log(`Error creating real checklist for ${monthInfo.month}/${monthInfo.year}:`, error.message);
      }
    });
    
    if (pdfBlobs.length === 0) {
      throw new Error("ไม่สามารถสร้างตารางเช็คชื่อใดๆ ได้");
    }
    
    // สร้าง ZIP
    Logger.log(`Creating real checklist ZIP with ${pdfBlobs.length} PDF files...`);
    const zipBlob = Utilities.zip(pdfBlobs, `ตารางเช็คชื่อ_${grade}_ห้อง${classNo}_ปีการศึกษา${academicYearBE}.zip`);
    
    // บันทึกลง Google Drive
    const zipUrl2 = _saveBlobGetUrl_(zipBlob);
    
    Logger.log(`✅ Real checklist ZIP export completed: ${results.length}/${academicMonths.length} successful`);
    return zipUrl2;
    
  } catch (error) {
    Logger.log("❌ Error in exportAcademicYearChecklistZIP:", error.message);
    throw new Error(`ไม่สามารถส่งออกตารางเช็คชื่อทั้งปีเป็น ZIP ได้: ${error.message}`);
  }
}

/**
 * ✅ ฟังก์ชันใหม่: ดึงสรุปวันมาของนักเรียนแต่ละเดือนในปีการศึกษา
 * @param {string} grade ระดับชั้น
 * @param {string} classNo ห้อง
 * @param {number} academicYear ปีการศึกษา (พ.ศ.)
 * @returns {Array} ข้อมูลสรุปแต่ละเดือน
 */
function getMonthlyAttendanceStats(grade, classNo, academicYear) {
  try {
    Logger.log(`🔍 getMonthlyAttendanceStats: ${grade}/${classNo}/${academicYear}`);
    
    const monthsInYear = getMonthsInAcademicYear(academicYear);
    const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                        'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    
    const monthlyStats = [];
    
    monthsInYear.forEach(({ month, yearCE }) => {
      try {
        // ดึงข้อมูลรายเดือนจากฟังก์ชันที่มีอยู่แล้ว
        const monthlyData = getMonthlyAttendanceSummary(grade, classNo, yearCE, month);
        
        // คำนวณสถิติของเดือนนี้
        let totalStudents = 0;
        let totalPresent = 0;
        let totalLeave = 0;
        let totalAbsent = 0;
        let totalDays = 0;
        let studentsWithData = 0;
        
        monthlyData.forEach(student => {
          totalStudents++;
          const present = parseInt(student.present) || 0;
          const leave = parseInt(student.leave) || 0;
          const absent = parseInt(student.absent) || 0;
          const total = present + leave + absent;
          
          if (total > 0) {
            studentsWithData++;
            totalPresent += present;
            totalLeave += leave;
            totalAbsent += absent;
            totalDays += total;
          }
        });
        
        // คำนวณเปอร์เซ็นต์
        const avgPresent = totalDays > 0 ? ((totalPresent / totalDays) * 100).toFixed(1) : "0.0";
        const avgLeave = totalDays > 0 ? ((totalLeave / totalDays) * 100).toFixed(1) : "0.0";
        const avgAbsent = totalDays > 0 ? ((totalAbsent / totalDays) * 100).toFixed(1) : "0.0";
        
        // กำหนดภาคเรียน
        let semester = 1;
        if (month >= 11 || month <= 3) {
          semester = 2;
        }
        
        // กำหนดสีตามเปอร์เซ็นต์การมาเรียน
        let statusColor = '#4caf50'; // เขียว - ดีมาก
        const avgPresentNum = parseFloat(avgPresent);
        if (avgPresentNum < 60) {
          statusColor = '#f44336'; // แดง - ต้องปรับปรุง
        } else if (avgPresentNum < 70) {
          statusColor = '#ff9800'; // ส้ม - พอใช้
        } else if (avgPresentNum < 80) {
          statusColor = '#2196f3'; // น้ำเงิน - ดี
        }
        
        monthlyStats.push({
          month: month,
          yearCE: yearCE,
          yearBE: yearCE + 543,
          monthName: thaiMonths[month - 1],
          displayName: `${thaiMonths[month - 1]} ${yearCE + 543}`,
          semester: semester,
          semesterName: `ภาคเรียนที่ ${semester}`,
          
          // สถิติพื้นฐาน
          totalStudents: totalStudents,
          studentsWithData: studentsWithData,
          totalPresent: totalPresent,
          totalLeave: totalLeave,
          totalAbsent: totalAbsent,
          totalDays: totalDays,
          
          // เปอร์เซ็นต์
          avgPresent: parseFloat(avgPresent),
          avgLeave: parseFloat(avgLeave),
          avgAbsent: parseFloat(avgAbsent),
          avgPresentStr: avgPresent + '%',
          avgLeaveStr: avgLeave + '%',
          avgAbsentStr: avgAbsent + '%',
          
          // สถานะและสี
          statusColor: statusColor,
          statusText: avgPresentNum >= 80 ? 'ดีมาก' : 
                     avgPresentNum >= 70 ? 'ดี' : 
                     avgPresentNum >= 60 ? 'พอใช้' : 'ต้องปรับปรุง',
          
          // ข้อมูลสำหรับกราฟ
          chartData: {
            present: totalPresent,
            leave: totalLeave,
            absent: totalAbsent
          },
          
          // วันที่มีข้อมูล (จำนวนวันที่มีการเรียน)
          schoolDays: totalDays > 0 ? Math.round(totalDays / studentsWithData) : 0
        });
        
        Logger.log(`✅ สรุปเดือน ${thaiMonths[month - 1]} ${yearCE + 543}: ${studentsWithData}/${totalStudents} คน, มาเรียน ${avgPresent}%`);
        
      } catch (error) {
        Logger.log(`❌ Error processing month ${month}/${yearCE}:`, error.message);
        // เพิ่มเดือนว่างถ้าไม่มีข้อมูล
        monthlyStats.push({
          month: month,
          yearCE: yearCE,
          yearBE: yearCE + 543,
          monthName: thaiMonths[month - 1],
          displayName: `${thaiMonths[month - 1]} ${yearCE + 543}`,
          semester: month >= 11 || month <= 3 ? 2 : 1,
          semesterName: `ภาคเรียนที่ ${month >= 11 || month <= 3 ? 2 : 1}`,
          
          totalStudents: 0,
          studentsWithData: 0,
          totalPresent: 0,
          totalLeave: 0,
          totalAbsent: 0,
          totalDays: 0,
          
          avgPresent: 0,
          avgLeave: 0,
          avgAbsent: 0,
          avgPresentStr: "0.0%",
          avgLeaveStr: "0.0%",
          avgAbsentStr: "0.0%",
          
          statusColor: '#bdbdbd',
          statusText: 'ไม่มีข้อมูล',
          
          chartData: { present: 0, leave: 0, absent: 0 },
          schoolDays: 0
        });
      }
    });
    
    // เรียงลำดับตามเดือนในปีการศึกษา
    monthlyStats.sort((a, b) => {
      // ภาคเรียนที่ 1 ก่อน (พ.ค.-ต.ค.)
      if (a.semester !== b.semester) {
        return a.semester - b.semester;
      }
      // เรียงตามเดือนภายในภาคเรียน
      if (a.semester === 1) {
        return a.month - b.month;
      } else {
        // ภาคเรียนที่ 2: พ.ย., ธ.ค., ม.ค., ก.พ., มี.ค.
        const monthOrder = [11, 12, 1, 2, 3];
        return monthOrder.indexOf(a.month) - monthOrder.indexOf(b.month);
      }
    });
    
    Logger.log(`✅ getMonthlyAttendanceStats completed: ${monthlyStats.length} months processed`);
    return monthlyStats;
    
  } catch (error) {
    Logger.log(`❌ Error in getMonthlyAttendanceStats:`, error.message);
    throw new Error(`ไม่สามารถดึงสถิติรายเดือนได้: ${error.message}`);
  }
}

/**
 * ดึงข้อมูลสรุปวันมาแต่ละเดือนในรูปแบบ Matrix
 * (ใช้โครงสร้างเดียวกันกับ getMonthlyAttendanceSummary)
 */
function getMonthlyAttendanceMatrix(grade, classNo, academicYear) {
  try {
    const ss = SS();
    
    // 1. ดึงรายชื่อนักเรียนที่ต้องการทั้งหมด "ครั้งเดียว"
    const studentsInClass = getStudentsForAttendance(grade, classNo);
    if (!studentsInClass || studentsInClass.length === 0) {
      throw new Error(`ไม่พบนักเรียนในชั้น ${grade}/${classNo}`);
    }
    
    // 2. สร้างรายการเดือนในปีการศึกษา
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const monthNamesShort = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
    
    const months = [];
    const academicYearInt = parseInt(academicYear);
    
    // ภาคเรียนที่ 1: พฤษภาคม - ธันวาคม (เดือน 5-12)
    for (let month = 5; month <= 12; month++) {
      months.push({
        key: `${academicYearInt - 543}-${month.toString().padStart(2, '0')}`,
        shortName: monthNamesShort[month - 1],
        year: academicYearInt,
        fullName: `${monthNames[month - 1]} ${academicYearInt}`,
        sheetName: `${monthNames[month - 1]}${academicYearInt}`,
        monthIndex: month,
        yearCE: academicYearInt - 543
      });
    }
    
    // ภาคเรียนที่ 2: มกราคม - มีนาคม (เดือน 1-3)
    for (let month = 1; month <= 3; month++) {
      months.push({
        key: `${academicYearInt - 543 + 1}-${month.toString().padStart(2, '0')}`,
        shortName: monthNamesShort[month - 1], 
        year: academicYearInt + 1,
        fullName: `${monthNames[month - 1]} ${academicYearInt + 1}`,
        sheetName: `${monthNames[month - 1]}${academicYearInt + 1}`,
        monthIndex: month,
        yearCE: academicYearInt - 543 + 1
      });
    }
    
    // 3. สร้าง Map เพื่อเก็บข้อมูลและค้นหาได้รวดเร็ว
    const studentMap = new Map();
    studentsInClass.forEach(stu => {
      studentMap.set(stu.id, {
        id: stu.id,
        name: `${stu.title || ''}${stu.firstname} ${stu.lastname}`.trim(),
        monthlyData: {}, // เก็บข้อมูลแต่ละเดือน
        totalPresent: 0,
        totalDays: 0
      });
    });
    
    // 4. วนลูปแต่ละเดือนเพื่อดึงข้อมูล
    months.forEach(month => {
      try {
        const sheet = ss.getSheetByName(month.sheetName);
        if (!sheet) {
          Logger.log(`ไม่พบชีต: ${month.sheetName}, ข้ามไป`);
          // ตั้งค่าเป็น 0 สำหรับเดือนที่ไม่มีชีต
          studentMap.forEach(student => {
            student.monthlyData[month.key] = 0;
          });
          return;
        }
        
        // อ่านข้อมูลจากชีต "ครั้งเดียว" ทั้งหมด
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) {
          // ไม่มีข้อมูลนักเรียน ตั้งค่าเป็น 0
          studentMap.forEach(student => {
            student.monthlyData[month.key] = 0;
          });
          return;
        }
        
        const headers = data[0];
        const idColIdx = headers.findIndex(h => String(h).includes("รหัส"));
        const presentColIdx = headers.findIndex(h => String(h).startsWith("มา"));
        const leaveColIdx = headers.findIndex(h => String(h).startsWith("ลา"));
        const absentColIdx = headers.findIndex(h => String(h).startsWith("ขาด"));
        
         // ตรวจสอบว่ามีคอลัมน์สรุปหรือไม่
        if ([idColIdx, presentColIdx, leaveColIdx, absentColIdx].includes(-1)) {
          Logger.log(`ชีต ${month.sheetName} ไม่มีคอลัมน์สรุป (มา/ลา/ขาด), ข้ามไป`);
          studentMap.forEach(student => {
            student.monthlyData[month.key] = 0;
          });
          return;
        }
        
        // หาคอลัมน์วันที่
        const dateCols = headers.map((h, i) => {
          const m = String(h).match(/^(\d{1,2})/);
          return m ? { col: i, day: parseInt(m[1]) } : null;
        }).filter(Boolean);
        
        // 5. วนลูปประมวลผลข้อมูล + นับวันเรียนจริง (เฉพาะนักเรียนในห้องที่เลือก)
        const schoolDaysSet = new Set();
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const studentIdMatch = String(row[idColIdx] || '').match(/^(\d+)/);
          const studentId = studentIdMatch ? studentIdMatch[1] : null;
          
          // ถ้าเป็นนักเรียนในห้องที่ต้องการ
          if (studentId && studentMap.has(studentId)) {
            // นับวันที่มีเช็คชื่อจริงของห้องนี้ (ทุกวันที่มีการเช็คชื่อ รวมเสาร์-อาทิตย์-วันหยุด)
            dateCols.forEach(dc => {
              const v = String(row[dc.col] || '').trim();
              if (v === '/' || v === 'ม' || v === 'ล' || v === 'ข' || v === '0' || v === '1' || v.toLowerCase() === 'l') {
                schoolDaysSet.add(dc.day);
              }
            });
            const presentCount = Number(row[presentColIdx]) || 0;
            const leaveCount = Number(row[leaveColIdx]) || 0;
            const absentCount = Number(row[absentColIdx]) || 0;
            const totalCount = presentCount + leaveCount + absentCount;
            
            // เก็บข้อมูลเฉพาะวันมาในเดือนนี้
            const student = studentMap.get(studentId);
            student.monthlyData[month.key] = presentCount;
            student.totalPresent += presentCount;
            student.totalDays += totalCount;
          }
         }
        month.schoolDays = schoolDaysSet.size;
        
        // ตั้งค่า 0 สำหรับนักเรียนที่ไม่มีข้อมูลในเดือนนี้
        studentMap.forEach(student => {
          if (!(month.key in student.monthlyData)) {
            student.monthlyData[month.key] = 0;
          }
        });
        
        Logger.log(`ดึงข้อมูลจากชีต ${month.sheetName} สำเร็จ (วันเรียน: ${month.schoolDays} วัน)`);
        

      } catch (error) {
        Logger.log(`ข้อผิดพลาดในการอ่านชีต "${month.sheetName}": ${error.message}`);
        // ตั้งค่าเป็น 0 สำหรับเดือนที่มีปัญหา
        studentMap.forEach(student => {
          student.monthlyData[month.key] = 0;
        });
      }
    });
    
    // 6. แปลง Map เป็น Array และคำนวณร้อยละ
    const processedStudents = Array.from(studentMap.values()).map(student => {
      const percentage = student.totalDays > 0 ? 
        ((student.totalPresent / student.totalDays) * 100).toFixed(1) : '0.0';
      
      return {
        id: student.id,
        name: student.name,
        monthlyData: student.monthlyData,
        totalPresent: student.totalPresent,
        totalDays: student.totalDays,
        percentage: parseFloat(percentage)
      };
    });
    
    // --- 👇 เพิ่มโค้ดเรียงลำดับตรงนี้ ก่อน return (ใช้ property "id") ---
    processedStudents.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, {numeric: true}));
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

    Logger.log(`สรุปข้อมูลสำเร็จ ปีการศึกษา ${academicYear}, พบนักเรียน ${processedStudents.length} คน`);
    
    return {
      students: processedStudents,
      months: months,
      academicYear: academicYear,
      grade: grade,
      classNo: classNo
    };
    
  } catch (error) {
    throw new Error(`ไม่สามารถดึงข้อมูลสรุปวันมาแต่ละเดือน: ${error.message}`);
  }
}


/**
 * ส่งออก PDF สรุปวันมาแต่ละเดือน
 */

/**
 * ส่งออก PDF สรุปวันมาแต่ละเดือน (ใช้ HtmlService - ไม่ต้องขออนุญาต)
 */

/**
 * แปลงชื่อชั้นเรียน ป.1–ป.6 → ประถมศึกษาปีที่ 1–6
 */
function getGradeFullName(grade) {
  const map = {
    "ป.1": "ประถมศึกษาปีที่ 1",
    "ป.2": "ประถมศึกษาปีที่ 2",
    "ป.3": "ประถมศึกษาปีที่ 3",
    "ป.4": "ประถมศึกษาปีที่ 4",
    "ป.5": "ประถมศึกษาปีที่ 5",
    "ป.6": "ประถมศึกษาปีที่ 6",
  };
  return map[grade] || grade;
}

/**
 * สร้าง HTML สำหรับ PDF สรุปวันมาแต่ละเดือน (A4 แนวตั้ง)
 * - ภาคเรียนที่ 1: พ.ค.–ต.ค.
 * - ภาคเรียนที่ 2: พ.ย.–มี.ค.
 * - เส้นตารางดำทั้งหมด / ข้อความทุกช่องสีดำ
 */
function createMonthlyAttendanceMatrixHTML(data, grade, classNo, academicYear) {
  const gradeName = getGradeFullName(grade);
  // ===== คำนวณสถิติโดยรวม =====
  const totalStudents = data.students.length;
  const avgAttendance = data.students.reduce((sum, s) => sum + s.percentage, 0) / Math.max(totalStudents, 1);
  const excellentCount = data.students.filter(s => s.percentage >= 80).length;
  const goodCount      = data.students.filter(s => s.percentage >= 70 && s.percentage < 80).length;
  const fairCount      = data.students.filter(s => s.percentage >= 60 && s.percentage < 70).length;
  const poorCount      = data.students.filter(s => s.percentage < 60).length;

  // ===== แบ่งเดือนตามภาคเรียน =====
  const semester1 = data.months
    .filter(m => m.monthIndex >= 5 && m.monthIndex <= 10)
    .sort((a, b) => a.monthIndex - b.monthIndex);

  const semester2Order = [11, 12, 1, 2, 3];
  const semester2 = data.months
    .filter(m => (m.monthIndex >= 11 || m.monthIndex <= 3))
    .sort((a, b) => semester2Order.indexOf(a.monthIndex) - semester2Order.indexOf(b.monthIndex));

  // ===== ส่วนหัวคอลัมน์เดือน =====
  const makeHeaderRows = (months) => {
    let row1 = '', row2 = '';
    months.forEach(month => {
      row1 += `<th class="month-header">${month.shortName}</th>`;
      row2 += `<th class="year-header">${month.year}</th>`;
    });
    return { row1, row2 };
  };
  const s1Hdr = makeHeaderRows(semester1);
  const s2Hdr = makeHeaderRows(semester2);

  // ===== แถวข้อมูลนักเรียน =====
  let studentRows = '';
  data.students.forEach((student, index) => {
    let s1Cells = '';
    let s1Total = 0;
    semester1.forEach(m => {
      const v = Number(student.monthlyData[m.key] || 0);
      s1Cells += `<td>${v > 0 ? v : '-'}</td>`;
      s1Total += v;
    });

    let s2Cells = '';
    let s2Total = 0;
    semester2.forEach(m => {
      const v = Number(student.monthlyData[m.key] || 0);
      s2Cells += `<td>${v > 0 ? v : '-'}</td>`;
      s2Total += v;
    });

    const rowClass = index % 2 === 0 ? 'even-row' : 'odd-row';
    studentRows += `
      <tr class="${rowClass}">
        <td class="center">${index + 1}</td>
        <td class="left student-name">${student.name}</td>
        ${s1Cells}
        <td class="center bold">${s1Total}</td>
        ${s2Cells}
        <td class="center bold">${s2Total}</td>
        <td class="center bold">${student.totalPresent}</td>
        <td class="center bold">${student.percentage}%</td>
      </tr>
    `;
  });

  // ===== HTML =====
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4 portrait; margin: 1.5cm 1cm; }

    body {
      font-family: 'Sarabun','Tahoma',sans-serif;
      font-size: 14px; color: #000; margin: 0; padding: 0; line-height: 1.4;
    }

    .header { text-align: center; margin-bottom: 25px; border-bottom: 3px solid #000; padding-bottom: 20px; }
    .header h1 { color: #000; font-size: 24px; font-weight: bold; margin: 0 0 10px 0; }
    .header h2 { color: #000; font-size: 18px; font-weight: normal; margin: 0; }

    .table-container { width: 100%; margin-bottom: 30px; overflow-x: auto; }

    table { width: 100%; border-collapse: collapse; border: 1px solid #000 !important; font-size: 12px; }
    th, td { border: 0.6px solid #000 !important; vertical-align: middle; }

    td { color: #000 !important; padding: 6px 3px; text-align: center; font-size: 12px; }

    /* หัวตาราง: ขาวพื้นหลัง ข้อความดำ */
    th {
      background: #fff !important;
      color: #000 !important;
      font-weight: bold;
      text-align: center;
      padding: 10px 3px;
      font-size: 12px;
    }
    .month-header, .year-header, .semester-header {
      background: #fff !important;
      color: #000 !important;
    }
    .month-header { font-size: 11px; padding: 8px 2px; min-width: 32px; max-width: 32px; }
    .year-header  { font-size: 10px; padding: 5px 2px; min-width: 32px; max-width: 32px; }
    .semester-header { font-size: 11px; padding: 8px 4px; }

    .left { text-align: left; padding-left: 10px; }
    .student-name { min-width: 180px; max-width: 200px; }
    .center { text-align: center; }
    .bold { font-weight: bold; }

    .even-row { background: #f8f9fa; }
    .odd-row  { background: #fff; }

    .summary-section {
      page-break-before: always; margin-top: 30px; padding: 25px;
      background: #f6f6f6; border: 2px solid #000; border-radius: 10px;
    }
    .summary-title {
      color: #000; font-size: 24px; font-weight: bold; margin-bottom: 25px; text-align: center;
      border-bottom: 2px solid #000; padding-bottom: 15px;
    }
    .summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; margin-bottom: 30px; }
    .summary-item { padding: 20px; background: #fff; border: 1px solid #000; border-radius: 10px; text-align: center; }
    .summary-label { font-weight: bold; color: #000; font-size: 16px; margin-bottom: 10px; }
    .summary-value { font-size: 24px; font-weight: bold; color: #000; }

    .level-stats { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-top: 20px; }
    .level-item  { padding: 18px; background: #fff; border: 1px solid #000; border-radius: 8px; text-align: center; }
    .level-label { font-size: 14px; color: #000; }
    .level-value { font-size: 20px; font-weight: bold; color: #000; }
  </style>
</head>
<body>
  <div class="header">
    <h1>สรุปเวลาเรียน ปีการศึกษา ${academicYear} </h1>
    <h2>ชั้น ${gradeName} ห้อง ${classNo} </h2>
  </div>

  <div class="table-container">
    <table>
      <thead>
        <tr>
          <th rowspan="3" style="width:30px;">เลขที่</th>
          <th rowspan="3" style="width:35%;">ชื่อ-นามสกุล</th>
          <th colspan="${semester1.length}" class="semester-header">ภาคเรียนที่ 1</th>
          <th rowspan="3">รวม<br>ภาค 1</th>
          <th colspan="${semester2.length}" class="semester-header">ภาคเรียนที่ 2</th>
          <th rowspan="3">รวม<br>ภาค 2</th>
          <th rowspan="3">รวม<br>ทั้งปี</th>
          <th rowspan="3">ร้อยละ</th>
        </tr>
        <tr>
          ${s1Hdr.row1}
          ${s2Hdr.row1}
        </tr>
        <tr>
          ${s1Hdr.row2}
          ${s2Hdr.row2}
        </tr>
      </thead>
      <tbody>
        ${studentRows}
      </tbody>
    </table>
  </div>

  <!-- สถิติสรุปในหน้า 2 -->
  <div class="summary-section">
    <div class="summary-title">📈 สถิติสรุป</div>
    <div class="summary-grid">
      <div class="summary-item">
        <div class="summary-label">👥 จำนวนนักเรียนทั้งหมด</div>
        <div class="summary-value">${totalStudents} คน</div>
      </div>
      <div class="summary-item">
        <div class="summary-label">📊 เฉลี่ยร้อยละการมาเรียน</div>
        <div class="summary-value">${avgAttendance.toFixed(1)}%</div>
      </div>
    </div>

    <div class="level-stats">
      <div class="level-item">
        <div class="level-label">🟢 ระดับดีมาก (≥80%)</div>
        <div class="level-value">${excellentCount} คน</div>
      </div>
      <div class="level-item">
        <div class="level-label">🔵 ระดับดี (70–79%)</div>
        <div class="level-value">${goodCount} คน</div>
      </div>
      <div class="level-item">
        <div class="level-label">🟡 พอใช้ (60–69%)</div>
        <div class="level-value">${fairCount} คน</div>
      </div>
      <div class="level-item">
        <div class="level-label">🔴 ต้องปรับปรุง (&lt;60%)</div>
        <div class="level-value">${poorCount} คน</div>
      </div>
    </div>
  </div>
</body>
</html>
  `;
}

/**
 * สร้าง PDF สรุปวันมาแต่ละเดือน (ใช้ HtmlService - เวอร์ชันเก่า - ลบได้)
 */
function generateMonthlyAttendanceMatrixPDF(data) {
  // ฟังก์ชันนี้ไม่ใช้แล้ว - ใช้ exportMonthlyAttendanceMatrixPDF แทน
  throw new Error('ฟังก์ชันนี้ถูกแทนที่ด้วย exportMonthlyAttendanceMatrixPDF');
}

/**
 * บันทึกข้อมูลสรุปวันมาแต่ละเดือนลงชีต
 */
function saveMonthlyAttendanceMatrixToSheet(data, grade, classNo, academicYear) {
  try {
    const ss = SS();
    const sheetName = `สรุปวันมาแต่ละเดือน_${grade}${classNo}_${academicYear}`;
    
    // ลบชีตเก่าถ้ามี
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // สร้างชีตใหม่
    const sheet = ss.insertSheet(sheetName);
    
    // ข้อมูลหัวตาราง
    const headers1 = ['ที่', 'รหัส', 'ชื่อ-นามสกุล'];
    data.months.forEach(month => {
      headers1.push(month.shortName);
    });
    headers1.push('รวมวันมา', 'ร้อยละ');
    
    const headers2 = ['', '', ''];
    data.months.forEach(month => {
      headers2.push(month.year.toString());
    });
    headers2.push('', '');
    
    // เขียนหัวตาราง
    sheet.getRange(1, 1, 1, headers1.length).setValues([headers1]);
    sheet.getRange(2, 1, 1, headers2.length).setValues([headers2]);
    
    // ข้อมูลนักเรียน
    const studentData = [];
    data.students.forEach((student, index) => {
      const row = [
        index + 1,
        student.id,
        student.name
      ];
      
      // วันมาแต่ละเดือน
      data.months.forEach(month => {
        const attendanceCount = student.monthlyData[month.key] || 0;
        row.push(attendanceCount > 0 ? attendanceCount : '');
      });
      
      // รวมและร้อยละ
      row.push(student.totalPresent);
      row.push(student.percentage);
      
      studentData.push(row);
    });
    
    // เขียนข้อมูลนักเรียน
    if (studentData.length > 0) {
      sheet.getRange(3, 1, studentData.length, headers1.length).setValues(studentData);
    }
    
    // จัดรูปแบบ
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // หัวตาราง
    sheet.getRange(1, 1, 2, lastCol).setBackground('#4472c4').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(1, 1, 2, lastCol).setHorizontalAlignment('center');
    
    // ข้อบอร์เดอร์
    sheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true);
    
    // จัดรูปแบบคอลัมน์
    sheet.getRange(3, 1, lastRow - 2, 1).setHorizontalAlignment('center'); // ลำดับ
    sheet.getRange(3, 2, lastRow - 2, 1).setHorizontalAlignment('center'); // รหัส
    sheet.getRange(3, 3, lastRow - 2, 1).setHorizontalAlignment('left');   // ชื่อ
    sheet.getRange(3, 4, lastRow - 2, data.months.length).setHorizontalAlignment('center'); // เดือน
    sheet.getRange(3, lastCol - 1, lastRow - 2, 2).setHorizontalAlignment('center'); // รวม/ร้อยละ
    
    // ปรับขนาดคอลัมน์
    sheet.autoResizeColumns(1, lastCol);
    
    return 'บันทึกข้อมูลสำเร็จ';
    
  } catch (error) {
    throw new Error(`ไม่สามารถบันทึกข้อมูลลงชีต: ${error.message}`);
  }
}
/**
 * แก้ไขเมื่อบันทึกผิดวัน หรือรีเซ็ตค่าวันในชีตมาเรียนรายวัน
 */
function deleteAttendanceData(payload) {
  try {
    const { grade, classNo, year, month, day } = payload;
    
    // ใช้ monthNames ที่ประกาศไว้แล้ว
    const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const sheetName = `${monthNames[month]}${year + 543}`;
    
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: `ไม่พบชีต ${sheetName}` };
    }
    
    // หาคอลัมน์ของวันที่
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dayCol = headers.findIndex(h => String(h).match(new RegExp(`^${day}\\s`)));
    
    if (dayCol === -1) {
      return { success: false, message: `ไม่พบคอลัมน์วันที่ ${day}` };
    }
    
    // ดึงรายชื่อนักเรียนในชั้น/ห้องนี้
    const studentsInClass = getStudentsForAttendance(grade, classNo);
    const studentIds = new Set(studentsInClass.map(s => String(s.id)));
    
    // ล้างข้อมูลเฉพาะนักเรียนในชั้น/ห้องนี้
    const data = sheet.getDataRange().getValues();
    let clearedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rowStudentId = String(data[i][1] || '').match(/^(\d+)/);
      if (rowStudentId && studentIds.has(rowStudentId[1])) {
        sheet.getRange(i + 1, dayCol + 1).clearContent();
        clearedCount++;
      }
    }
    
    return { 
      success: true, 
      message: `ลบข้อมูลวันที่ ${day}/${month + 1}/${year} สำเร็จ (${clearedCount} คน)` 
    };
    
  } catch (error) {
    Logger.log('Error in deleteAttendanceData:', error);
    return { 
      success: false, 
      message: `เกิดข้อผิดพลาด: ${error.message}` 
    };
  }
}
/**
 * ดึงเนื้อหาหน้าพร้อมพารามิเตอร์
 * @param {string} page ชื่อหน้า
 * @param {Object} params พารามิเตอร์เพิ่มเติม (เช่น month: "2025-10")
 * @returns {string} HTML เนื้อหาของหน้า
 */
function getPageContentWithParams(page, params) {
  try {
    const MONTH_NAMES_TH = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    Logger.log(`🔍 getPageContentWithParams: page=${page}, params=${JSON.stringify(params)}`);

    // ดึงเดือนและปีจาก params.month (รูปแบบ YYYY-MM)
    let year, month;
    if (params && params.month && params.month.match(/^\d{4}-\d{2}$/)) {
      const [y, m] = params.month.split('-').map(Number);
      year = y;
      month = m - 1; // แปลงเป็น 0-11 สำหรับใช้ใน MONTH_NAMES_TH
    } else {
      // Fallback เดือนมีนาคม
      year = 2025;
      month = 2; // มีนาคม
    }

    // ดึงปีการศึกษาจาก global_settings
    const ss = SS();
    const yearSetting = ss.getSheetByName('global_settings');
    let academicYear = U_getCurrentAcademicYear();
    
    // สร้างชื่อชีต
    const sheetName = `${MONTH_NAMES_TH[month]}${year + 543}`;
    
    // สร้าง HTML ตามหน้า
    if (page === 'semester_attendance_summary') {
      const grade = params.grade || 'ป.1';
      const classNo = params.classNo || '1';
      const data = getMonthlyAttendanceSummary(grade, classNo, year, month + 1);

      let html = `
        <div class="container animate-fade-in">
          <h2 class="mb-4">สรุปการมาเรียน - เดือน ${MONTH_NAMES_TH[month]}${year + 543}</h2>
          <div class="card mb-4">
            <div class="card-body">
              <h5>ชั้น ${grade} ห้อง ${classNo} ปีการศึกษา ${academicYear}</h5>
              <table class="table table-bordered table-hover">
                <thead class="table-primary">
                  <tr>
                    <th>เลขที่</th>
                    <th>รหัสนักเรียน</th>
                    <th>ชื่อ-นามสกุล</th>
                    <th>มา</th>
                    <th>ลา</th>
                    <th>ขาด</th>
                    <th>รวมวันเรียน</th>
                    <th>ร้อยละการมาเรียน</th>
                  </tr>
                </thead>
                <tbody>
      `;

      if (data.length === 0) {
        html += `
          <tr>
            <td colspan="8" class="text-center text-muted">ไม่มีข้อมูลสำหรับเดือน ${MONTH_NAMES_TH[month]}${year + 543}</td>
          </tr>
        `;
      } else {
        data.forEach((student, index) => {
          const total = student.total || (student.present + student.leave + student.absent);
          const percentage = total > 0 ? ((student.present / total) * 100).toFixed(1) : "0.0";
          html += `
            <tr>
              <td>${index + 1}</td>
              <td>${student.studentId}</td>
              <td>${student.name}</td>
              <td>${student.present}</td>
              <td>${student.leave}</td>
              <td>${student.absent}</td>
              <td>${total}</td>
              <td>${percentage}%</td>
            </tr>
          `;
        });
      }

      html += `
                </tbody>
              </table>
            </div>
          </div>
        </div>
      `;
      return html;
    } else if (page === 'attendance_monthly' || page === 'attendance_table') {
      const grade = params.grade || 'ป.1';
      const classNo = params.classNo || '1';
      const tableData = getAttendanceVerticalTable(grade, classNo, year, month);

      let html = `
        <div class="container animate-fade-in">
          <h2 class="mb-4">ตารางเช็คชื่อ - เดือน ${MONTH_NAMES_TH[month]}${year + 543}</h2>
          <div class="card mb-4">
            <div class="card-body">
              <h5>ชั้น ${grade} ห้อง ${classNo} ปีการศึกษา ${academicYear}</h5>
              <table class="table table-bordered table-sm">
                <thead class="table-primary">
                  <tr>
                    <th>ที่</th>
                    <th>รหัส</th>
                    <th>ชื่อ-สกุล</th>
      `;

      if (tableData.days && tableData.days.length > 0) {
        tableData.days.forEach(day => {
          html += `<th>${day.label} ${day.dow}</th>`;
        });
      }

      html += `
                </thead>
                <tbody>
      `;

      if (!tableData.students || tableData.students.length === 0) {
        html += `
          <tr>
            <td colspan="${tableData.days ? tableData.days.length + 3 : 3}" class="text-center text-muted">
              ไม่มีข้อมูลสำหรับเดือน ${MONTH_NAMES_TH[month]}${year + 543}
            </td>
          </tr>
        `;
      } else {
        tableData.students.forEach((student, index) => {
          html += `
            <tr>
              <td>${index + 1}</td>
              <td>${student.id}</td>
              <td>${student.name}</td>
          `;
          tableData.days.forEach(day => {
            const status = student.statusMap[day.label] || '';
            html += `<td>${status}</td>`;
          });
          html += `</tr>`;
        });
      }

      html += `
                </tbody>
              </table>
            </div>
          </div>
        </div>
      `;
      return html;
    } else {
      return getPageContent(page);
    }
  } catch (e) {
    Logger.log(`Error in getPageContentWithParams: ${e.message}`);
    return `
      <div class="alert alert-danger">
        <h4>เกิดข้อผิดพลาด</h4>
        <p>ไม่สามารถโหลดหน้า ${page} ได้ กรุณาลองใหม่</p>
        <small>ข้อผิดพลาด: ${e.message}</small>
      </div>
    `;
  }
}

/**
 * exportMonthlyAttendanceMatrixPDF
 */


/**
 * exportMonthlyAttendanceMatrixPDF
 */


/**
 * exportMonthlyAttendanceMatrixPDF
 */


/**
 * exportMonthlyAttendanceMatrixPDF
 */


/**
 * exportMonthlyAttendanceMatrixPDF
 */
function exportMonthlyAttendanceMatrixPDF(grade, classNo, academicYear) {
  try {
    const data = getMonthlyAttendanceMatrix(grade, classNo, academicYear);
    if (!data || !data.students || data.students.length === 0) {
      throw new Error('��辺����������Ѻ���͡');
    }
    const html = createMonthlyAttendanceMatrixHTML(data, grade, classNo, academicYear);
    const blob = HtmlService.createHtmlOutput(html).getBlob().setName('Attendance.pdf').getAs('application/pdf');
    return _saveBlobGetUrl_(blob);
  } catch (error) {
    throw new Error('Error creating PDF');
  }
}

