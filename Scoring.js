// ============================================================
// 📚 SCORING.GS - ระบบตารางคะแนนรายวิชา (แก้ไขสมบูรณ์)
// ============================================================

/**
 * ✅ ดึงรายการระดับชั้นที่มีอยู่ (รวมจากทุกแหล่ง)
 * ฟังก์ชั่นนี้จะพยายามอ่านจากทั้ง "Students" และ "รายวิชา"
 * รองรับการเรียกใช้งานแบบเดิมทั้งหมด
 * * @returns {Array<string>} อาร์เรย์ของระดับชั้นที่เรียงลำดับแล้ว
 */
function getAvailableGrades() {
  const gradesSet = new Set();
   
  try {
    // พยายามเชื่อมต่อ Spreadsheet
    let ss;
    try {
      // ลองใช้ SPREADSHEET_ID ก่อน (ถ้ามี)
      if (getSpreadsheetId_()) {
        ss = SS();
      }
    } catch (e) {
      // ถ้าไม่ได้ ใช้ Active Spreadsheet แทน
      ss = SS();
    }
    
    if (!ss) {
      ss = SS();
    }
    
    // ============================================
    // 🎯 วิธีที่ 1: อ่านจากชีต "Students"
    // ============================================
    try {
      const studentsSheet = ss.getSheetByName("Students");
      if (studentsSheet) {
        const data = studentsSheet.getDataRange().getValues();
        
        if (data.length > 1) {
          const headers = data[0];
          const gradeIndex = headers.indexOf("grade");
          
          if (gradeIndex !== -1) {
            for (let i = 1; i < data.length; i++) {
              const grade = String(data[i][gradeIndex] || '').trim();
              if (grade) {
                gradesSet.add(grade);
              }
            }
            Logger.log(`✅ อ่านจากชีต Students: ${gradesSet.size} ระดับชั้น`);
          }
        }
      }
    } catch (e) {
      Logger.log(`⚠️ ไม่สามารถอ่านจากชีต Students: ${e.message}`);
    }
    
    // ============================================
    // 🎯 วิธีที่ 2: อ่านจากชีต "รายวิชา"
    // ============================================
    try {
      const coursesSheet = ss.getSheetByName("รายวิชา");
      if (coursesSheet) {
        const data = coursesSheet.getDataRange().getValues();
        
        if (data.length > 1) {
          // อ่านจากคอลัมน์ A (index 0)
          for (let i = 1; i < data.length; i++) {
            const grade = String(data[i][0] || '').trim();
            if (grade) {
              gradesSet.add(grade);
            }
          }
          Logger.log(`✅ อ่านจากชีต รายวิชา: ${gradesSet.size} ระดับชั้น`);
        }
      }
    } catch (e) {
      Logger.log(`⚠️ ไม่สามารถอ่านจากชีต รายวิชา: ${e.message}`);
    }
    
    // ============================================
    // 🎯 เรียงลำดับและส่งคืนผลลัพธ์
    // ============================================
    if (gradesSet.size === 0) {
      Logger.log("❌ ไม่พบข้อมูลระดับชั้นในชีตใดๆ");
      return [];
    }
    
    // เรียงลำดับแบบภาษาไทย รองรับทั้ง ป.1-6, ม.1-6, ป1-6, ม1-6
    const result = Array.from(gradesSet).sort((a, b) => {
      return a.localeCompare(b, 'th-TH', { 
        numeric: true,
        sensitivity: 'base'
      });
    });
    
    Logger.log(`✅ ผลลัพธ์สุดท้าย: ${JSON.stringify(result)}`);
    return result;
    
  } catch (e) {
    Logger.log(`❌ Error in getAvailableGrades: ${e.message}`);
    Logger.log(`📚 Stack: ${e.stack}`);
    return [];
  }
}

/**
 * ✅ ดึงรายการห้องเรียนที่มีนักเรียนอยู่ สำหรับระดับชั้นที่เลือก
 * @param {string} grade - ระดับชั้น เช่น "ป.3"
 * @returns {Array<string>} อาร์เรย์ของหมายเลขห้อง เรียงลำดับแล้ว
 */
function getAvailableClassNos(grade) {
  try {
    grade = String(grade || '').trim();
    if (!grade) return [];
    var ss = SS();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var gradeIdx = headers.indexOf('grade');
    var classIdx = headers.indexOf('class_no');
    if (gradeIdx === -1 || classIdx === -1) return [];
    var classSet = new Set();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][gradeIdx] || '').trim() === grade) {
        var c = String(data[i][classIdx] || '').trim();
        if (c) classSet.add(c);
      }
    }
    return Array.from(classSet).sort(function(a, b) {
      return a.localeCompare(b, 'th-TH', { numeric: true });
    });
  } catch (e) {
    Logger.log('❌ Error in getAvailableClassNos: ' + e.message);
    return [];
  }
}

/**
 * ✅ ดึงรายวิชาทั้งหมดสำหรับระดับชั้นที่เลือก
 * @param {string} grade - ระดับชั้นที่ต้องการค้นหา
 * @returns {Array<Object>} อาร์เรย์ของอ็อบเจกต์รายวิชา
 */
function getSubjectsForGrade(grade) {
  try {
    Logger.log(`=== เริ่มต้น getSubjectsForGrade("${grade}") ===`);
    
    if (!getSpreadsheetId_()) {
      Logger.log("❌ ไม่พบ SPREADSHEET_ID");
      throw new Error("SPREADSHEET_ID ไม่ได้ถูกกำหนด");
    }
    
    const ss = SS();
    const sheet = ss.getSheetByName("รายวิชา");
    
    if (!sheet) {
      Logger.log("❌ ไม่พบชีต 'รายวิชา'");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("❌ ไม่มีข้อมูลในชีต 'รายวิชา'");
      return [];
    }

    const subjects = [];

    // วนลูปจากแถวที่ 2 (index 1) เพื่อข้าม header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = row[0]?.toString().trim();       // คอลัมน์ A = ชั้น
      const subjectCode = row[1]?.toString().trim();    // คอลัมน์ B = รหัสวิชา
      const subjectName = row[2]?.toString().trim();    // คอลัมน์ C = ชื่อวิชา
      const timeInfo = row[3]?.toString().trim() || ""; // คอลัมน์ D = ข้อมูลเวลา
      
      Logger.log(`📝 แถวที่ ${i + 1}: ${rowGrade} | ${subjectCode} | ${subjectName} | ${timeInfo}`);
      
      if (rowGrade === grade.trim() && subjectCode && subjectName) {
        // ✅ [FIXED] แยกตัวเลขจากข้อความในวงเล็บ เช่น "(40)" → 40
        let hour = 0;
        // ใช้ Regex จับตัวเลขในวงเล็บ \((\d+)\) หรือจับตัวเลขเฉยๆ (\d+) ถ้าไม่มีวงเล็บ
        const hourMatch = timeInfo.match(/\((\d+)\)/) || timeInfo.match(/(\d+)/); 
        
        if (hourMatch) {
          hour = parseInt(hourMatch[1], 10); // ใส่ base 10 เพื่อความปลอดภัย
        }
        
        const subject = {
          code: subjectCode,
          name: subjectName,
          hour: hour,
          timeInfo: timeInfo
        };
        
        subjects.push(subject);
        Logger.log(`✅ เพิ่มรายวิชา: ${JSON.stringify(subject)}`);
      }
    }
    
    Logger.log(`✅ พบรายวิชาสำหรับ ${grade}: ${subjects.length} วิชา`);
    return subjects;
    
  } catch(e) {
    Logger.log("❌ Error in getSubjectsForGrade: " + e.message);
    Logger.log("📚 Stack Trace: " + e.stack);
    return [];
  }
}

/**
 * ✅ สร้างตารางคะแนนในชีตใหม่ (เวอร์ชันปรับปรุง)
 * @param {string} subjectCode - รหัสวิชา
 * @param {string} subjectName - ชื่อวิชา
 * @param {number|string} hour - จำนวนชั่วโมง
 * @param {string} grade - ระดับชั้น
 * @param {string} classNo - หมายเลขห้อง
 * @returns {string} ข้อความสถานะ
 */
function createScoreSheet(subjectCode, subjectName, hour, grade, classNo) {
  try {
    // ❗ [ปรับปรุง] ตรวจสอบว่าได้รับข้อมูลห้องเรียนหรือไม่
    if (!classNo) {
      throw new Error("กรุณาเลือกระดับชั้นและห้องเรียนให้ครบถ้วน");
    }

    // ❗ [ปรับปรุง] สร้างชื่อชีตที่ไม่ซ้ำกันโดยการเพิ่มระดับชั้นและห้อง
    const cleanGrade = grade.replace(/[\/\.]/g, ''); // เอาอักขระพิเศษออกจากชื่อชั้น เช่น ป.1 -> ป1
    const sheetName = `${subjectName} ${cleanGrade}-${classNo}`;
    Logger.log(`📋 ชื่อชีตที่จะสร้าง: "${sheetName}"`);
    
    const ss = SS();

    // ตรวจสอบว่ามีชีตชื่อนี้อยู่แล้วหรือไม่
    const existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      throw new Error(`ชีตชื่อ "${sheetName}" มีอยู่แล้ว`);
    }

    // ❗ [ปรับปรุง] ดึงข้อมูลนักเรียนตาม grade และ classNo ที่ผู้ใช้เลือก
    Logger.log(`👨‍🎓 กำลังดึงข้อมูลนักเรียน ${grade}/${classNo}`);
    const students = getStudentsForAttendance(grade, classNo);
    
    // ❗ [ปรับปรุง] ตรวจสอบว่ามีข้อมูลนักเรียนก่อนสร้างชีต
    if (students.length === 0) {
      throw new Error(`ไม่พบข้อมูลนักเรียนสำหรับชั้น ${grade}/${classNo} ไม่สามารถสร้างตารางได้`);
    }

    Logger.log(`👨‍🎓 พบนักเรียน: ${students.length} คน`);
    
    // สร้างชีตใหม่
    Logger.log(`🛠️ กำลังสร้างชีต "${sheetName}"`);
    const sheet = ss.insertSheet(sheetName);

    // ส่วนหัวของตาราง
    // โครงสร้างตาม SchoolMIS (16 คอลัมน์ต่อภาค)
    const header3 = [
      "ลำดับ", "เลขประจำตัว", "ชื่อ - สกุล",
      // Term 1 (16 cols: D-S, index 3-18)
      "ครั้งที่1","ครั้งที่2","ครั้งที่3","ครั้งที่4","รวม",
      "ครั้งที่5","แก้ตัวกลางภาค",
      "ครั้งที่6","ครั้งที่7","ครั้งที่8","ครั้งที่9","รวม",
      "รวมระหว่างภาค","ครั้งที่10","รวมทั้งหมด","เกรด",
      // Term 2 (16 cols: T-AI, index 19-34)
      "ครั้งที่1","ครั้งที่2","ครั้งที่3","ครั้งที่4","รวม",
      "ครั้งที่5","แก้ตัวกลางภาค",
      "ครั้งที่6","ครั้งที่7","ครั้งที่8","ครั้งที่9","รวม",
      "รวมระหว่างภาค","ครั้งที่10","รวมทั้งหมด","เกรด",
      // Summary (2 cols: AJ-AK, index 35-36)
      "คะแนนเฉลี่ย 2 ภาค", "ระดับผลการเรียน"
    ];

    // ใส่ข้อมูลพื้นฐาน
    sheet.getRange(1, 1).setValue("รหัสวิชา").setFontWeight("bold").setBackground("#E8F4FD");
    sheet.getRange(1, 2).setValue(subjectCode).setBackground("#E8F4FD");
    sheet.getRange(2, 1).setValue("ชื่อวิชา").setFontWeight("bold").setBackground("#E8F4FD");
    sheet.getRange(2, 2).setValue(subjectName).setBackground("#E8F4FD");
    sheet.getRange(2, 3).setValue(`${hour} ชั่วโมง/ปี`).setBackground("#E8F4FD");
    
    // ใส่ส่วนหัวตาราง
    sheet.getRange(3, 1, 1, header3.length).setValues([header3]).setFontWeight("bold").setBackground("#D4EDDA");
    
    // แถวคะแนนเต็ม
    sheet.getRange(4, 3).setValue("คะแนนเต็ม").setFontWeight("bold").setBackground("#FFF3CD");

    // ใส่ข้อมูลนักเรียน
    const studentRows = students.map((s, i) => [i + 1, s.id, (s.name || ((s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '')).trim())]);
    sheet.getRange(5, 1, studentRows.length, 3).setValues(studentRows);
    
    // ใส่สีพื้นหลังให้แถวนักเรียนแต่ละคน
    for (let i = 0; i < studentRows.length; i++) {
      const rowNum = 5 + i;
      const bgColor = i % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
      sheet.getRange(rowNum, 1, 1, header3.length).setBackground(bgColor);
    }

    // จัดรูปแบบ
    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(3);
    
    // ปรับขนาดคอลัมน์
    sheet.setColumnWidth(1, 50);   // ลำดับ
    sheet.setColumnWidth(2, 100);  // เลขประจำตัว
    sheet.setColumnWidth(3, 200);  // ชื่อ-สกุล
    
    for (let col = 4; col <= header3.length; col++) {
      sheet.setColumnWidth(col, 80);
    }

    const resultMessage = `✅ สร้างชีต "${sheetName}" สำเร็จแล้ว พร้อมข้อมูลนักเรียน ${students.length} คน`;
    Logger.log(resultMessage);
    return resultMessage;

  } catch(e) {
    Logger.log("❌ Error in createScoreSheet: " + e.message);
    // ส่งข้อความ Error ที่ชัดเจนกลับไปให้หน้าเว็บ
    throw new Error("เกิดข้อผิดพลาด: " + e.message);
  }
}

/**
 * 🆕 ดึงรายการตารางคะแนนที่มีอยู่แล้ว
 * @returns {Array<Object>} รายการตารางคะแนน
 */
function getExistingScoreSheets() {
  try {
    Logger.log("=== ดึงรายการตารางคะแนน ===");
    
    const ss = SS();
    const allSheets = ss.getSheets();
    const scoreSheets = [];
    
    // ✅ [FIXED] ดึงเวลาแก้ไขล่าสุดของไฟล์ Spreadsheet (เพราะ Sheet ไม่มีเมธอด getLastEditTime)
    let fileLastUpdated = new Date();
    try {
      fileLastUpdated = DriveApp.getFileById(getSpreadsheetId_()).getLastUpdated();
    } catch (err) {
      Logger.log("⚠️ ไม่สามารถดึงเวลาแก้ไขจาก DriveApp ได้: " + err.message);
    }

    // ชีตที่ไม่ใช่ตารางคะแนน
    const excludeSheets = ['Students', 'รายวิชา', 'Users', 'users', 'Settings', 'Attendance', 'global_settings', 'การตั้งค่าระบบ', 'Holidays', 'วันหยุด', 'SCORES_WAREHOUSE', 'AttendanceLog', 'HomeroomTeachers', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ', 'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ', 'ความเห็นครู', 'โปรไฟล์นักเรียน', 'BACKUP_WAREHOUSE_LATEST', 'Template_', 'สรุปวันมา', 'สรุปการมาเรียน', 'อ่านคิดเขียน', 'คุณลักษณะ', 'กิจกรรม', 'สมรรถนะ'];
    
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // ข้ามชีตที่ไม่ใช่ตารางคะแนน
      if (excludeSheets.some(ex => sheetName === ex || sheetName.startsWith(ex))) {
        return;
      }
      
      // ตรวจสอบว่าเป็นตารางคะแนนหรือไม่โดยดูโครงสร้าง
      try {
        const cell_A1 = sheet.getRange(1, 1).getValue();
        const cell_A2 = sheet.getRange(2, 1).getValue();
        
        if (cell_A1 === "รหัสวิชา" && cell_A2 === "ชื่อวิชา") {
          const subjectCode = sheet.getRange(1, 2).getValue();
          const subjectName = sheet.getRange(2, 2).getValue();
          const timeInfo = sheet.getRange(2, 3).getValue();
          
          // แยกชั้นและห้องจากชื่อชีต
          const nameMatch = sheetName.match(/(.+)\s+([^-]+)-(\d+)$/);
          let grade = "", classNo = "";
          
          if (nameMatch) {
            grade = nameMatch[2];
            classNo = nameMatch[3];
          }
          
          scoreSheets.push({
            sheetName: sheetName,
            subjectCode: subjectCode,
            subjectName: subjectName,
            timeInfo: timeInfo,
            grade: grade,
            classNo: classNo,
            // ✅ [FIXED] ใช้เวลาของไฟล์รวม หรือใช้เวลาปัจจุบัน
            lastModified: fileLastUpdated
          });
        }
      } catch (e) {
        // ข้ามชีตที่มีปัญหา
        Logger.log(`ข้ามชีต ${sheetName}: ${e.message}`);
      }
    });
    
    // เรียงตามวันที่แก้ไขล่าสุด
    scoreSheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
    
    Logger.log(`✅ พบตารางคะแนน: ${scoreSheets.length} ตาราง`);
    return scoreSheets;
    
  } catch (e) {
    Logger.log("❌ Error in getExistingScoreSheets: " + e.message);
    return [];
  }
}

/**
 * 🆕 ดึงข้อมูลคะแนนจากตารางที่เลือก
 * @param {string} sheetName - ชื่อชีต
 * @returns {Object} ข้อมูลคะแนนทั้งหมด
 */
function getScoreData(sheetName) {
  try {
    Logger.log(`=== ดึงข้อมูลคะแนนจาก "${sheetName}" ===`);
    
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`ไม่พบชีต "${sheetName}"`);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // ดึงข้อมูลพื้นฐาน
    const subjectCode = data[0][1] || "";
    const subjectName = data[1][1] || "";
    const timeInfo = data[1][2] || "";
    
    // ดึงหัวตาราง (แถวที่ 3)
    const headers = data[2] || [];
    
    // ดึงข้อมูลนักเรียนและคะแนน (เริ่มจากแถวที่ 5)
    const studentData = [];
    for (let i = 4; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1] && row[2]) { // มี ลำดับ, รหัส, ชื่อ
        const studentRow = {
          order: row[0],
          studentId: row[1],
          studentName: row[2],
          scores: row.slice(3) // คะแนนทั้งหมด
        };
        studentData.push(studentRow);
      }
    }
    
    const result = {
      sheetName: sheetName,
      subjectCode: subjectCode,
      subjectName: subjectName,
      timeInfo: timeInfo,
      headers: headers,
      students: studentData,
      totalStudents: studentData.length
    };
    
    Logger.log(`✅ ดึงข้อมูลสำเร็จ: ${studentData.length} นักเรียน`);
    return result;
    
  } catch (e) {
    Logger.log("❌ Error in getScoreData: " + e.message);
    throw new Error("ไม่สามารถดึงข้อมูลคะแนนได้: " + e.message);
  }
}

/**
 * 🆕 บันทึกคะแนน
 * @param {string} sheetName - ชื่อชีต
 * @param {Array} scoreData - ข้อมูลคะแนนที่จะบันทึก
 * @returns {string} ข้อความสถานะ
 */
function saveScores(sheetName, scoreData) {
  try {
    Logger.log(`=== บันทึกคะแนนลง "${sheetName}" ===`);
    
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`ไม่พบชีต "${sheetName}"`);
    }
    
    // บันทึกข้อมูลคะแนน
    scoreData.forEach((studentData, index) => {
      const rowIndex = 5 + index; // เริ่มจากแถวที่ 5
      const startCol = 4; // เริ่มจากคอลัมน์ D (คะแนน)
      
      // บันทึกคะแนนแต่ละคอลัมน์
      studentData.scores.forEach((score, colIndex) => {
        if (score !== null && score !== undefined && score !== "") {
          sheet.getRange(rowIndex, startCol + colIndex).setValue(score);
        }
      });
    });
    
    const message = `✅ บันทึกคะแนนเรียบร้อยแล้ว (${scoreData.length} นักเรียน)`;
    Logger.log(message);
    return message;
    
  } catch (e) {
    Logger.log("❌ Error in saveScores: " + e.message);
    throw new Error("ไม่สามารถบันทึกคะแนนได้: " + e.message);
  }
}


// === ฟังก์ชันเดิมที่ยังต้องใช้ ===

/**
 * เพิ่มรายวิชาใหม่ลงในชีต "รายวิชา"
 */

/**
 * อัปเดตข้อมูลรายวิชา
 */

/**
 * ลบรายวิชา
 */

/**
 * แปลงคะแนนเป็นเกรด (4, 3.5, 3, ...)
 */
function getGrade(score) {
  if (score >= 80) return 4;
  if (score >= 75) return 3.5;
  if (score >= 70) return 3;
  if (score >= 65) return 2.5;
  if (score >= 60) return 2;
  if (score >= 55) return 1.5;
  if (score >= 50) return 1;
  return 0;
}

/**
 * ✅ ตรวจสอบว่าชีตมีอยู่แล้วหรือไม่
 * @param {string} sheetName - ชื่อชีตที่ต้องการตรวจสอบ
 * @returns {boolean} - true หากมีชีตอยู่แล้ว, false หากไม่มี
 */
function checkSheetExists(sheetName) {
  try {
    if (!getSpreadsheetId_()) {
      Logger.log("❌ ไม่พบ SPREADSHEET_ID");
      return false;
    }
    
    const ss = SS();
    return ss.getSheetByName(sheetName) !== null;
  } catch (e) {
    Logger.log("❌ Error in checkSheetExists: " + e.message);
    return false;
  }
}

/**
 * ✅ ฟังก์ชันใหม่ - ดึงระดับชั้น (ชื่อใหม่เพื่อไม่ให้ซ้ำ)
 * แก้ปัญหาการส่งข้อมูลจาก GAS ไป HTML
 */
function loadAvailableGrades() {
  try {
    console.log("=== loadAvailableGrades เริ่มทำงาน ===");
    
    if (!getSpreadsheetId_()) {
      console.log("❌ SPREADSHEET_ID ไม่ได้ถูกกำหนด");
      return [];
    }
    
    const ss = SS();
    const sheet = ss.getSheetByName("รายวิชา");
    
    if (!sheet) {
      console.log("❌ ไม่พบชีท 'รายวิชา'");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log("❌ ไม่มีข้อมูลในชีท");
      return [];
    }

    const grades = [];
    for (let i = 1; i < data.length; i++) {
      const grade = data[i][0]; 
      if (grade) {
        const gradeStr = String(grade).trim();
        if (gradeStr !== "" && !grades.includes(gradeStr)) {
          grades.push(gradeStr);
        }
      }
    }
    
    const sortedGrades = grades.sort();
    console.log("✅ loadAvailableGrades - ระดับชั้นที่พบ:", sortedGrades);
    
    // ส่งข้อมูลแบบ JSON string เพื่อป้องกันปัญหาการ serialize
    return sortedGrades;
    
  } catch (error) {
    console.log("❌ Error in loadAvailableGrades:", error.message);
    return [];
  }
}

/**
 * ✅ ฟังก์ชันใหม่ - ดึงรายวิชา (ชื่อใหม่เพื่อไม่ให้ซ้ำ)
 * แก้ปัญหาการส่งข้อมูลจาก GAS ไป HTML
 */
function loadSubjectsForGrade(grade) {
  try {
    console.log(`=== loadSubjectsForGrade("${grade}") เริ่มทำงาน ===`);
    
    if (!grade || typeof grade !== 'string') {
      console.log("❌ Parameter grade ไม่ถูกต้อง:", grade);
      return [];
    }
    
    if (!getSpreadsheetId_()) {
      console.log("❌ SPREADSHEET_ID ไม่ได้ถูกกำหนด");
      return [];
    }
    
    const ss = SS();
    const sheet = ss.getSheetByName("รายวิชา");
    
    if (!sheet) {
      console.log("❌ ไม่พบชีท 'รายวิชา'");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log("❌ ไม่มีข้อมูลในชีท");
      return [];
    }

    const subjects = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = row[0] ? String(row[0]).trim() : "";      
      const subjectCode = row[1] ? String(row[1]).trim() : "";   
      const subjectName = row[2] ? String(row[2]).trim() : "";   
      const hour = row[3] || 0;                                 
      
      if (rowGrade === grade && subjectCode !== "" && subjectName !== "") {
        subjects.push({
          code: subjectCode,
          name: subjectName,
          hour: hour,
          timeInfo: `${hour} ชั่วโมง/ปี`
        });
      }
    }
    
    console.log(`✅ loadSubjectsForGrade - พบ ${subjects.length} วิชาสำหรับ ${grade}`);
    return subjects;
    
  } catch (error) {
    console.log("❌ Error in loadSubjectsForGrade:", error.message);
    return [];
  }
}

/**
 * ดึง midMax/finMax จากชื่อชีต (format: "ชื่อวิชา_ชั้น-ห้อง" เช่น "คณิตศาสตร์_ป.1-1")
 * ใช้ภายใน saveScoreSheetData เพื่อแปลงสัดส่วนคะแนน
 */
function getAssessmentConfigFromSheet_(sheetName) {
  try {
    // แยกชื่อวิชาและระดับชั้นจาก sheetName
    // format: "ชื่อวิชา_ชั้น-ห้อง" เช่น "คณิตศาสตร์_ป.1-1"
    const parts = sheetName.split('_');
    if (parts.length < 2) return { midMax: 70, finalMax: 30 };
    
    const subjectName = parts[0];
    const gradeClassPart = parts.slice(1).join('_'); // ป.1-1
    const gradeParts = gradeClassPart.split('-');
    const grade = gradeParts[0]; // ป.1

    // เรียก getSubjectAssessmentConfig ด้วย subjectName, code ว่าง, grade
    const cfg = getSubjectAssessmentConfig(subjectName, '', grade);
    return { midMax: Number(cfg.midMax) || 70, finalMax: Number(cfg.finalMax) || 30 };
  } catch (e) {
    console.warn('getAssessmentConfigFromSheet_ error:', e.message);
    return { midMax: 70, finalMax: 30 };
  }
}

/**
 * บันทึกคะแนนเฉพาะภาคเรียน (term) - แก้ไขแล้ว
 * พร้อมคำนวณคะแนนเฉลี่ย 2 ภาคและเกรด (อัปเดตทุกครั้ง)
 */
function saveScoreSheetData(sheetName, term, studentScores, fullScores, fullFinal) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`ไม่พบชีต: ${sheetName}`);

    const startRow = 5;
    
    // ✅ โครงสร้างตาม SchoolMIS (16 คอลัมน์ต่อภาค)
    // Per term layout (0-indexed from termStart):
    //   0-3: ครั้งที่ 1-4 (ก่อนกลางภาค)
    //   4:   รวม(ก่อนกลางภาค)
    //   5:   ครั้งที่ 5 (กลางภาค)
    //   6:   แก้ตัวกลางภาค
    //   7-10: ครั้งที่ 6-9 (หลังกลางภาค)
    //   11:  รวม(หลังกลางภาค)
    //   12:  รวมระหว่างภาค
    //   13:  ครั้งที่ 10 (ปลายภาค)
    //   14:  รวมทั้งหมด
    //   15:  เกรด
    const termCols = {
      term1: { 
        start: 3,           // D
        s1: 3, s2: 4, s3: 5, s4: 6, sum14: 7,
        s5: 8, makeup: 9,
        s6: 10, s7: 11, s8: 12, s9: 13, sum69: 14,
        midTotal: 15,
        s10: 16,
        total: 17,
        grade: 18
      },
      term2: { 
        start: 19,          // T
        s1: 19, s2: 20, s3: 21, s4: 22, sum14: 23,
        s5: 24, makeup: 25,
        s6: 26, s7: 27, s8: 28, s9: 29, sum69: 30,
        midTotal: 31,
        s10: 32,
        total: 33,
        grade: 34
      }
    };
    
    const cols = termCols[term];
    if (!cols) throw new Error("term ต้องเป็น 'term1' หรือ 'term2'");

    // fullScores จาก frontend: [s1,s2,s3,s4, s5, s6,s7,s8,s9] = 9 ช่อง (ตัวชี้วัด)
    // fullFinal จาก frontend: คะแนนปลายภาค (ครั้งที่ 10)
    // บันทึกคะแนนเต็มในแถวที่ 4
    var scoreSlots = [cols.s1, cols.s2, cols.s3, cols.s4, cols.s5, cols.s6, cols.s7, cols.s8, cols.s9];
    for (let i = 0; i < 9; i++) {
      sheet.getRange(4, scoreSlots[i] + 1).setValue(fullScores[i] || "");
    }
    // ดึง midMax/finMax จากชีตรายวิชา เพื่อแปลงสัดส่วน
    // sheetName format: "ชื่อวิชา_ชั้น-ห้อง" → ต้องหา subject จาก sheetName
    const cfg = getAssessmentConfigFromSheet_(sheetName);
    const midMax = cfg.midMax;
    const finMax = cfg.finalMax;
    const fullScoresSum = fullScores.reduce((a, b) => a + (Number(b) || 0), 0);

    // บันทึกคะแนนเต็ม midMax/finMax ในแถวที่ 4
    sheet.getRange(4, cols.midTotal + 1).setValue(midMax);
    sheet.getRange(4, cols.s10 + 1).setValue(finMax);
    sheet.getRange(4, cols.total + 1).setValue(midMax + finMax);

    // บันทึกคะแนนนักเรียน
    studentScores.forEach((student, idx) => {
      const row = startRow + idx;
      var scores = student.scores || [];

      // บันทึกคะแนน 9 ช่อง (ครั้งที่ 1-9 ตัวชี้วัด) — คะแนนดิบ
      for (let i = 0; i < 9; i++) {
        let val = scores[i];
        if (val === null || val === undefined || val === "") val = "";
        else {
          val = Number(val) || 0;
          var maxVal = Number(fullScores[i]) || 0;
          if (maxVal > 0 && val > maxVal) val = maxVal;
        }
        sheet.getRange(row, scoreSlots[i] + 1).setValue(val);
      }

      // บันทึกปลายภาค (ครั้งที่ 10) จาก student.final — คะแนนดิบ (เต็ม finMax)
      var s10val = Number(student.final) || 0;
      if (s10val > finMax) s10val = finMax;
      sheet.getRange(row, cols.s10 + 1).setValue(s10val);

      // คำนวณรวมก่อนกลางภาค (1-4) — คะแนนดิบ
      var sum14 = (Number(scores[0])||0) + (Number(scores[1])||0) + (Number(scores[2])||0) + (Number(scores[3])||0);
      sheet.getRange(row, cols.sum14 + 1).setValue(sum14);

      // คำนวณรวมหลังกลางภาค (6-9) — คะแนนดิบ
      var sum69 = (Number(scores[5])||0) + (Number(scores[6])||0) + (Number(scores[7])||0) + (Number(scores[8])||0);
      sheet.getRange(row, cols.sum69 + 1).setValue(sum69);

      // แก้ตัวกลางภาค
      sheet.getRange(row, cols.makeup + 1).setValue(student.makeup || "");

      // แปลงสัดส่วน: รวมระหว่างภาค = (คะแนนดิบรวมตัวชี้วัด / คะแนนเต็มรวม) * midMax
      var itemsSum = (Number(scores[0])||0) + (Number(scores[1])||0) + (Number(scores[2])||0) + (Number(scores[3])||0)
        + (Number(scores[4])||0) + (Number(scores[5])||0) + (Number(scores[6])||0) + (Number(scores[7])||0) + (Number(scores[8])||0);
      var midTotal = 0;
      if (fullScoresSum > 0 && itemsSum > 0) {
        midTotal = Math.round((itemsSum / fullScoresSum) * midMax);
      }
      sheet.getRange(row, cols.midTotal + 1).setValue(midTotal);

      // รวมทั้งหมด = รวมระหว่างภาค(แปลงแล้ว) + ปลายภาค (สเกล 100)
      var total = midTotal + s10val;
      sheet.getRange(row, cols.total + 1).setValue(total);

      // เกรด — คำนวณจาก total ที่เป็นสเกล 100 แล้ว
      sheet.getRange(row, cols.grade + 1).setValue(student.grade || calculateFinalGrade(total));

      // ✅ คำนวณคะแนนเฉลี่ย 2 ภาคและเกรดรวม
      const totalTerm1 = Number(sheet.getRange(row, termCols.term1.total + 1).getValue()) || 0;
      const totalTerm2 = Number(sheet.getRange(row, termCols.term2.total + 1).getValue()) || 0;
      
      let average = 0;
      let finalGradeVal = "";
      
      if (totalTerm1 > 0 && totalTerm2 > 0) {
        average = Math.round((totalTerm1 + totalTerm2) / 2);
        finalGradeVal = calculateFinalGrade(average);
      } else if (totalTerm1 > 0 || totalTerm2 > 0) {
        average = totalTerm1 > 0 ? totalTerm1 : totalTerm2;
        finalGradeVal = calculateFinalGrade(average);
      }

      // บันทึกคะแนนเฉลี่ยที่ AJ (column 36) และเกรดที่ AK (column 37)
      sheet.getRange(row, 36).setValue(average);
      sheet.getRange(row, 37).setValue(finalGradeVal);
    });

    return "บันทึกคะแนนเรียบร้อย พร้อมคำนวณเกรดเฉลี่ย";
  } catch (e) {
    Logger.log("❌ Error in saveScoreSheetData: " + e.message);
    throw new Error("ไม่สามารถบันทึกคะแนนได้: " + e.message);
  }
}

/**
 * ✅ ฟังก์ชันคำนวณเกรดปรับปรุง
 */
function calculateFinalGrade(score) {
  if (score >= 80) return "4";
  if (score >= 75) return "3.5";
  if (score >= 70) return "3";
  if (score >= 65) return "2.5";
  if (score >= 60) return "2";
  if (score >= 55) return "1.5";
  if (score >= 50) return "1";
  return "0";
}

/**
 * ✅ ดึงข้อมูลคะแนนจากชีตตาม term1 และ term2 - แก้ไขแล้ว
 */
function getScoreSheetData(sheetName) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`ไม่พบชีต: ${sheetName}`);

    const data = sheet.getDataRange().getValues();
    const startRow = 5;

    // ✅ โครงสร้างตาม SchoolMIS (16 คอลัมน์ต่อภาค)
    const termCols = {
      term1: { 
        s1: 3, s2: 4, s3: 5, s4: 6, sum14: 7,
        s5: 8, makeup: 9,
        s6: 10, s7: 11, s8: 12, s9: 13, sum69: 14,
        midTotal: 15,
        s10: 16,
        total: 17,
        grade: 18
      },
      term2: { 
        s1: 19, s2: 20, s3: 21, s4: 22, sum14: 23,
        s5: 24, makeup: 25,
        s6: 26, s7: 27, s8: 28, s9: 29, sum69: 30,
        midTotal: 31,
        s10: 32,
        total: 33,
        grade: 34
      }
    };

    function extractTermData(cols) {
      // คะแนนเต็ม 9 ช่อง (ครั้งที่ 1-9 ตัวชี้วัด) จากแถวที่ 4 (index 3)
      var scoreSlots = [cols.s1, cols.s2, cols.s3, cols.s4, cols.s5, cols.s6, cols.s7, cols.s8, cols.s9];
      const fullScores = scoreSlots.map(function(idx) { return Number(data[3][idx]) || 0; });
      const fullFinal = Number(data[3][cols.s10]) || 0; // คะแนนเต็มปลายภาค (ครั้งที่ 10)
      
      const rows = [];
      for (let i = startRow - 1; i < data.length; i++) {
        const row = data[i];
        if (!row[1] || !row[2]) continue;
        
        // ดึงคะแนน 9 ช่อง (ครั้งที่ 1-9 ตัวชี้วัด)
        const scores = scoreSlots.map(function(idx) { return Number(row[idx]) || 0; });
        const mid = Number(row[cols.midTotal]) || 0;
        const final_ = Number(row[cols.s10]) || 0; // ปลายภาค (ครั้งที่ 10) แยก
        const total = Number(row[cols.total]) || 0;
        const grade = String(row[cols.grade] || "");
        const makeup = Number(row[cols.makeup]) || 0;
        
        rows.push({ 
          id: row[1], 
          name: row[2], 
          scores, 
          mid, 
          final: final_, 
          total, 
          grade,
          makeup
        });
      }
      
      return { fullScores, fullFinal, rows };
    }

    // ดึงข้อมูลสรุปรวม 2 ภาค (คอลัมน์ AJ-AK, index 35-36)
    var yearSummary = [];
    for (var yi = startRow - 1; yi < data.length; yi++) {
      var yRow = data[yi];
      if (!yRow[1] || !yRow[2]) continue;
      yearSummary.push({
        id: yRow[1],
        name: yRow[2],
        yearAvg: yRow[35] !== undefined && yRow[35] !== '' ? Number(yRow[35]) || 0 : 0,
        yearGrade: String(yRow[36] || '')
      });
    }

    return {
      term1: extractTermData(termCols.term1),
      term2: extractTermData(termCols.term2),
      yearSummary: yearSummary
    };
  } catch (e) {
    Logger.log("❌ Error in getScoreSheetData: " + e.message);
    throw new Error("ไม่สามารถดึงข้อมูลคะแนนได้: " + e.message);
  }
}
/**
 * 🎯 ฟังก์ชันหลักที่หน้าเว็บเรียกใช้ - ดึงรายวิชาตามระดับชั้น
 * แทนที่ getSubjectListNoCache() เดิม
 */
function getSubjectListNoCache(grade, timestamp) {
  try {
    console.log(`🔍 [NEW] ดึงรายวิชาสำหรับ ${grade} | timestamp: ${timestamp}`);
    
    const ss = SS();
    const subjectSheet = ss.getSheetByName("รายวิชา");
    
    if (!subjectSheet) {
      console.error('❌ ไม่พบชีต "รายวิชา"');
      return { subjects: [], error: 'ไม่พบข้อมูลรายวิชา' };
    }
    
    const data = subjectSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log('⚠️ ไม่มีข้อมูลรายวิชา');
      return { subjects: [], error: 'ไม่มีข้อมูลรายวิชา' };
    }
    
    // หา header columns
    const headers = data[0];
    const gradeCol = U_findColumnIndex(headers, ['ชั้น', 'grade', 'ระดับชั้น']);
    const codeCol = U_findColumnIndex(headers, ['รหัสวิชา', 'code', 'subject_code']);
    const nameCol = U_findColumnIndex(headers, ['ชื่อวิชา', 'name', 'subject_name']);
    const hourCol = U_findColumnIndex(headers, ['ชั่วโมง', 'hour', 'hours']);
    const teacherCol = U_findColumnIndex(headers, ['ครูผู้สอน', 'teacher', 'instructor']);
    
    console.log(`📊 Column mapping: grade=${gradeCol}, code=${codeCol}, name=${nameCol}, hour=${hourCol}, teacher=${teacherCol}`);
    
    const subjects = [];
    
    // วนลูปหารายวิชาที่ตรงกับชั้น
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = String(row[gradeCol] || '').trim();
      
      if (rowGrade === grade) {
        const subject = {
          code: String(row[codeCol] || '').trim(),
          name: String(row[nameCol] || '').trim(),
          hours: row[hourCol] || 0,
          teacher: String(row[teacherCol] || '').trim(),
          grade: rowGrade,
          loadTimestamp: timestamp
        };
        
        if (subject.name && subject.code) {
          subjects.push(subject);
          console.log(`✅ เพิ่มวิชา: ${subject.name} (${subject.code})`);
        }
      }
    }
    
    console.log(`🎯 พบรายวิชาสำหรับ ${grade}: ${subjects.length} วิชา`);
    
    return {
      subjects: subjects,
      totalCount: subjects.length,
      grade: grade,
      loadTime: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ getSubjectListNoCache error:', error);
    return { 
      subjects: [], 
      error: `ไม่สามารถดึงข้อมูลได้: ${error.message}` 
    };
  }
}
/**
 * 📊 ดึงข้อมูลคะแนนจากชีตรายวิชา
 */
function getSubjectScoreData(subjectName, subjectCode, grade, classNo) {
  try {
    console.log(`📊 ดึงข้อมูลคะแนน: ${subjectName} ${grade}/${classNo}`);
    
    const ss = SS();
    
    // หาชีตคะแนน - ลองหลายรูปแบบ
    const possibleSheetNames = [
      `${subjectName} ${grade.replace(/\./g, '')}-${classNo}`,
      `${subjectName} ${grade}-${classNo}`,
      `${subjectName}_${grade}_${classNo}`,
      `${subjectCode} ${grade}-${classNo}`,
      subjectName
    ];
    
    let scoreSheet = null;
    let foundSheetName = '';
    
    for (const sheetName of possibleSheetNames) {
      scoreSheet = ss.getSheetByName(sheetName);
      if (scoreSheet) {
        foundSheetName = sheetName;
        console.log(`✅ พบชีตคะแนน: ${sheetName}`);
        break;
      }
    }
    
    if (!scoreSheet) {
      // หาชีตที่มีชื่อคล้ายๆ
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const sheetName = sheet.getName();
        if (sheetName.includes(subjectName) || 
            (subjectCode && sheetName.includes(subjectCode))) {
          scoreSheet = sheet;
          foundSheetName = sheetName;
          console.log(`✅ พบชีตคะแนนคล้าย: ${sheetName}`);
          break;
        }
      }
    }
    
    if (!scoreSheet) {
      throw new Error(`ไม่พบชีตคะแนนสำหรับ ${subjectName}`);
    }
    
    // ดึงข้อมูลจากชีต
    const data = scoreSheet.getDataRange().getValues();
    
    if (data.length < 5) {
      throw new Error('ข้อมูลในชีตคะแนนไม่ครบถ้วน');
    }
    
    // ดึงข้อมูลพื้นฐาน
    const sheetSubjectCode = data[0][1] || subjectCode;
    const sheetSubjectName = data[1][1] || subjectName;
    
    // ดึงข้อมูลนักเรียนและคะแนน
    const students = [];
    
    for (let i = 4; i < data.length; i++) {
      const row = data[i];
      
      // ข้ามแถวที่ไม่มีข้อมูลครบ
      if (!row[0] || !row[1] || !row[2]) continue;
      
      const student = {
        no: row[0],
        studentId: String(row[1]).trim(),
        name: String(row[2]).trim(),
        term1: extractTermScores(row, 3, 16),   // คอลัมน์ D-Q
        term2: extractTermScores(row, 17, 30),  // คอลัมน์ R-AE
        yearAverage: Number(row[31]) || 0,      // คะแนนเฉลี่ยปี
        finalGrade: String(row[32] || '0')      // เกรดสุดท้าย
      };
      
      students.push(student);
    }
    
    // --- 👇 เพิ่มโค้ดเรียงลำดับนักเรียนตามรหัส ---
    students.sort((a, b) => {
      const idA = a.studentId || '';
      const idB = b.studentId || '';
      return String(idA).localeCompare(String(idB), undefined, {numeric: true});
    });
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---
    
    console.log(`📝 พบข้อมูลนักเรียน: ${students.length} คน`);
    
    return {
      sheetName: foundSheetName,
      subjectCode: sheetSubjectCode,
      subjectName: sheetSubjectName,
      grade,
      classNo,
      students,
      totalStudents: students.length,
      lastUpdated: scoreSheet.getLastEditTime()
    };
    
  } catch (error) {
    console.error('❌ getSubjectScoreData error:', error);
    throw new Error(`ไม่สามารถดึงข้อมูลคะแนนได้: ${error.message}`);
  }
}
/**
 * 📊 แยกคะแนนแต่ละภาคเรียนจากแถวข้อมูล
 */
function extractTermScores(row, startCol, endCol) {
  try {
    const scores = [];
    
    // คะแนนย่อย 10 ช่อง
    for (let i = 0; i < 10; i++) {
      scores.push(Number(row[startCol + i]) || 0);
    }
    
    return {
      scores: scores,                                   // คะแนนย่อย 1-10
      midtermScore: Number(row[startCol + 10]) || 0,  // คะแนนระหว่างภาค  
      finalScore: Number(row[startCol + 11]) || 0,    // คะแนนปลายภาค
      totalScore: Number(row[startCol + 12]) || 0,    // รวมคะแนน
      grade: String(row[startCol + 13] || '0')        // เกรด
    };
  } catch (error) {
    console.error('❌ extractTermScores error:', error);
    return {
      scores: Array(10).fill(0),
      midtermScore: 0,
      finalScore: 0, 
      totalScore: 0,
      grade: '0'
    };
  }
}

/**
 * ⚙️ ดึงการตั้งค่าระบบ
 */
function getSystemSettings() {
  try {
    const ss = SS();
    const settingsSheet = ss.getSheetByName('global_settings');
    
    const defaultSettings = {
      schoolName: 'โรงเรียน',
      academicYear: '2568',
      directorName: 'ผู้อำนวยการ',
      directorTitle: 'ผู้อำนวยการสถานศึกษา',
      academicHead: 'หัวหน้างานวิชาการ',
      logoDataUrl: '',
      footerNote: ''
    };
    
    if (!settingsSheet) {
      console.log('⚠️ ไม่พบชีตการตั้งค่า ใช้ค่าเริ่มต้น');
      return defaultSettings;
    }
    
    const data = settingsSheet.getDataRange().getValues();
    const settings = { ...defaultSettings };
    
    // อ่านการตั้งค่าจากชีต
    for (let i = 0; i < data.length; i++) {
      const [key, value] = data[i];
      if (key && value) {
        // แปลงชื่อไทยเป็น key อังกฤษ
        const keyMap = {
          'ชื่อโรงเรียน': 'schoolName',
          'ปีการศึกษา': 'academicYear', 
          'ชื่อผู้อำนวยการ': 'directorName',
          'ตำแหน่งผู้อำนวยการ': 'directorTitle',
          'ชื่อหัวหน้างานวิชาการ': 'academicHead',
          'หมายเหตุท้ายรายงาน': 'footerNote'
        };
        
        const mappedKey = keyMap[key] || key;
        settings[mappedKey] = value;
      }
    }
    
    return settings;
    
  } catch (error) {
    console.error('❌ getSystemSettings error:', error);
    return {
      schoolName: 'โรงเรียน',
      academicYear: '2568', 
      directorName: 'ผู้อำนวยการ',
      directorTitle: 'ผู้อำนวยการสถานศึกษา',
      academicHead: 'หัวหน้างานวิชาการ',
      logoDataUrl: '',
      footerNote: ''
    };
  }
}

/**
 * 🎨 สร้าง HTML สำหรับ PDF รายงานรายวิชา
 */
function createSubjectReportHTML(data) {
  const {
    subjectName, subjectCode, grade, classNo, term, 
    scoreData, settings, reportDate
  } = data;
  
  const termText = term === 'both' ? 'ทั้งปีการศึกษา' :
                   term === '1' ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
  
  const gradeFullName = U_getGradeFullName(grade);
  
  // สร้างแถวข้อมูลนักเรียน
  const studentRows = scoreData.students.map((student, index) => {
    if (term === 'both') {
      // แสดงทั้งสองภาค
      return `
        <tr>
          <td class="center">${index + 1}</td>
          <td class="student-name">${student.name}</td>
          <td class="center">${student.term1.totalScore}</td>
          <td class="center">${student.term1.grade}</td>
          <td class="center">${student.term2.totalScore}</td>
          <td class="center">${student.term2.grade}</td>
          <td class="center highlight">${student.yearAverage}</td>
          <td class="center highlight">${student.finalGrade}</td>
        </tr>`;
    } else {
      // แสดงภาคเรียนเดียว
      const termData = term === '1' ? student.term1 : student.term2;
      return `
        <tr>
          <td class="center">${index + 1}</td>
          <td class="center">${student.studentId}</td>
          <td class="student-name">${student.name}</td>
          <td class="center">${termData.midtermScore}</td>
          <td class="center">${termData.finalScore}</td>
          <td class="center highlight">${termData.totalScore}</td>
          <td class="center highlight">${termData.grade}</td>
        </tr>`;
    }
  }).join('');
  
  // สร้างส่วนหัวตาราง
  const tableHeader = term === 'both' ? `
    <tr>
      <th rowspan="2">ที่</th>
      <th rowspan="2">ชื่อ - นามสกุล</th>
      <th colspan="2">ภาคเรียนที่ 1</th>
      <th colspan="2">ภาคเรียนที่ 2</th>
      <th rowspan="2">คะแนนเฉลี่ย</th>
      <th rowspan="2">เกรด</th>
    </tr>
    <tr>
      <th>คะแนน</th><th>เกรด</th>
      <th>คะแนน</th><th>เกรด</th>
    </tr>` : `
    <tr>
      <th>ที่</th>
      <th>รหัสนักเรียน</th>
      <th>ชื่อ - นามสกุล</th>
      <th>ระหว่างภาค</th>
      <th>ปลายภาค</th>
      <th>รวม</th>
      <th>เกรด</th>
    </tr>`;
  
  return `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <title>รายงานผลการเรียนรายวิชา</title>
      <style>
        @page { size: A4; margin: 1cm; }
        body {
          font-family: 'TH Sarabun New', 'Sarabun', sans-serif;
          font-size: 14pt;
          line-height: 1.4;
          margin: 0;
          padding: 0;
        }
        .header {
          text-align: center;
          margin-bottom: 20px;
          border-bottom: 2px solid #333;
          padding-bottom: 15px;
        }
        .school-name {
          font-size: 18pt;
          font-weight: bold;
          margin-bottom: 5px;
        }
        .report-title {
          font-size: 16pt;
          font-weight: bold;
          margin-bottom: 10px;
        }
        .subject-info {
          font-size: 14pt;
          margin-bottom: 5px;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 15px 0;
        }
        th, td {
          border: 1px solid #333;
          padding: 8px 4px;
          text-align: center;
          vertical-align: middle;
        }
        th {
          background-color: #f0f0f0;
          font-weight: bold;
          font-size: 12pt;
        }
        td {
          font-size: 11pt;
        }
        .student-name {
          text-align: left;
          padding-left: 8px;
          min-width: 150px;
        }
        .center { text-align: center; }
        .highlight {
          background-color: #fffacd;
          font-weight: bold;
        }
        .signature {
          margin-top: 40px;
          display: flex;
          justify-content: space-between;
        }
        .signature-box {
          text-align: center;
          width: 200px;
        }
        .signature-line {
          border-bottom: 1px solid #333;
          margin-bottom: 5px;
          height: 40px;
        }
        .footer {
          margin-top: 20px;
          text-align: right;
          font-size: 10pt;
          color: #666;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="school-name">${settings.schoolName}</div>
        <div class="report-title">รายงานผลการเรียนรายวิชา</div>
        <div class="subject-info">
          รายวิชา: ${subjectName} ${subjectCode ? `(${subjectCode})` : ''}
        </div>
        <div class="subject-info">
          ระดับชั้น: ${gradeFullName} ห้อง ${classNo} | ${termText}
        </div>
        <div class="subject-info">
          ปีการศึกษา: ${settings.academicYear}
        </div>
      </div>
      
      <table>
        <thead>${tableHeader}</thead>
        <tbody>${studentRows}</tbody>
      </table>
      
      <div class="signature">
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>ครูผู้สอน</div>
        </div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>${settings.directorName}<br/>${settings.directorTitle}</div>
        </div>
      </div>
      
      ${settings.footerNote ? `<div style="text-align: center; margin-top: 20px; font-style: italic;">${settings.footerNote}</div>` : ''}
      
      <div class="footer">
        วันที่พิมพ์: ${Utilities.formatDate(reportDate, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')}
      </div>
    </body>
    </html>
  `;
}

/**
 * 📄 แปลง HTML เป็น PDF และบันทึกไฟล์
 */
function convertAndSavePDF(htmlContent, fileInfo) {
  try {
    // แปลง HTML เป็น PDF
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html')
                      .getAs('application/pdf');
    
    // สร้างชื่อไฟล์
    const termText = fileInfo.term === 'both' ? 'ทั้งปี' : `ภาค${fileInfo.term}`;
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmm');
    const fileName = `รายงาน_${fileInfo.subjectName}_${fileInfo.grade.replace('.', '')}_${fileInfo.classNo}_${termText}_${timestamp}.pdf`;
    
    blob.setName(fileName);
    
    // บันทึกไฟล์
    const settings = getSystemSettings();
    let folder;
    
    try {
      if (settings.pdfSaveFolderId) {
        folder = DriveApp.getFolderById(settings.pdfSaveFolderId);
      } else {
        folder = DriveApp.getRootFolder();
      }
    } catch (e) {
      console.warn('ไม่สามารถใช้โฟลเดอร์ที่กำหนดได้ ใช้โฟลเดอร์ราก');
      folder = DriveApp.getRootFolder();
    }
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    console.log(`✅ บันทึกไฟล์ PDF: ${fileName}`);
    return file.getUrl();
    
  } catch (error) {
    console.error('❌ convertAndSavePDF error:', error);
    throw new Error(`ไม่สามารถสร้างไฟล์ PDF ได้: ${error.message}`);
  }
}

/**
 * 🔧 ฟังก์ชันช่วย - หาตำแหน่งคอลัมน์
 */

/**
 * ✅ คืนค่าเพดานคะแนนกลางภาค (midMax) / ปลายภาค (finalMax)
 * โดยอ่านจากชีต "รายวิชา" ตรงตามระดับชั้น รหัสวิชา และชื่อวิชา
 */
function getSubjectAssessmentConfig(subjectName, subjectCode, grade) {
  try {
    console.log(`🔍 getSubjectAssessmentConfig: ${subjectName} (${subjectCode}) | ${grade}`);
    
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName("รายวิชา"); // ✅ อ่านจากชีต "รายวิชา"
    
    if (!sheet) {
      console.warn('⚠️ ไม่พบชีต "รายวิชา" - ใช้ค่าเริ่มต้น 70/30');
      return { midMax: 70, finalMax: 30 };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.warn('⚠️ ชีต "รายวิชา" ไม่มีข้อมูล - ใช้ค่าเริ่มต้น 70/30');
      return { midMax: 70, finalMax: 30 };
    }

    // หา index หัวคอลัมน์
    const header = data[0];
    const idxGrade = header.indexOf('ชั้น');
    const idxCode  = header.indexOf('รหัสวิชา');
    const idxName  = header.indexOf('ชื่อวิชา');
    const idxMid   = header.indexOf('คะแนนระหว่างปี');
    const idxFinal = header.indexOf('คะแนนปลายปี');

    // ตรวจสอบว่าพบคอลัมน์ที่จำเป็นหรือไม่
    if (idxMid < 0 || idxFinal < 0) {
      console.warn('⚠️ ไม่พบคอลัมน์ "คะแนนระหว่างปี" หรือ "คะแนนปลายปี" - ใช้ค่าเริ่มต้น 70/30');
      return { midMax: 70, finalMax: 30 };
    }

    console.log(`📋 Column indices: grade=${idxGrade}, code=${idxCode}, name=${idxName}, mid=${idxMid}, final=${idxFinal}`);

    // หาแถวของวิชาที่ตรงกับเงื่อนไข
    let targetRow = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = String(row[idxGrade] || '').trim();
      const rowCode = String(row[idxCode] || '').trim();
      const rowName = String(row[idxName] || '').trim();
      
      // ตรวจสอบว่าตรงกับเงื่อนไข: ระดับชั้นตรง + (รหัสวิชาตรง หรือ ชื่อวิชาตรง)
      const gradeMatch = rowGrade === String(grade).trim();
      const codeMatch = rowCode === String(subjectCode).trim();
      const nameMatch = rowName === String(subjectName).trim();
      
      if (gradeMatch && (codeMatch || nameMatch)) {
        targetRow = row;
        console.log(`✅ พบแถวที่ตรงกัน: ${rowName} (${rowCode})`);
        break;
      }
    }

    if (!targetRow) {
      console.warn(`⚠️ ไม่พบรายวิชา ${subjectName} (${subjectCode}) ในระดับชั้น ${grade} - ใช้ค่าเริ่มต้น 70/30`);
      return { midMax: 70, finalMax: 30 };
    }

    // อ่านค่าคะแนนเต็ม
    const midMax = Number(targetRow[idxMid]) || 70;
    const finalMax = Number(targetRow[idxFinal]) || 30;

    console.log(`✅ คะแนนเต็ม: กลางภาค=${midMax}, ปลายภาค=${finalMax}`);
    
    return { midMax, finalMax };

  } catch (error) {
    console.error('❌ getSubjectAssessmentConfig error:', error);
    return { midMax: 70, finalMax: 30 };
  }
}





/**
 * 🔧 ฟังก์ชันช่วย - แปลงรหัสชั้นเป็นชื่อเต็ม
 */

/**
 * 🔧 ฟังก์ชันช่วย - ดึงข้อมูลครูผู้สอนจากรายวิชา
 */
function getSubjectTeacher(subjectName, subjectCode, grade) {
  try {
    const ss = SS();
    const subjectSheet = ss.getSheetByName("รายวิชา");
    
    if (!subjectSheet) return "ครูผู้สอน";
    
    const data = subjectSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = String(row[0] || '').trim();
      const rowCode = String(row[1] || '').trim();
      const rowName = String(row[2] || '').trim();
      const teacher = String(row[5] || '').trim(); // สมมติว่าครูอยู่คอลัมน์ที่ 6
      
      if (rowGrade === grade && 
          (rowName === subjectName || rowCode === subjectCode) && 
          teacher) {
        return teacher;
      }
    }
    
    return "ครูผู้สอน";
    
  } catch (error) {
    console.error('❌ getSubjectTeacher error:', error);
    return "ครูผู้สอน";
  }
}

/**
 * 🔧 ฟังก์ชันช่วย - คำนวณสถิติคะแนน
 */
function calculateScoreStatistics(students, term) {
  try {
    if (!students || students.length === 0) {
      return { average: 0, highest: 0, lowest: 0, passCount: 0, failCount: 0 };
    }
    
    const scores = students.map(student => {
      if (term === 'both') {
        return student.yearAverage || 0;
      } else if (term === '1') {
        return student.term1.totalScore || 0;
      } else {
        return student.term2.totalScore || 0;
      }
    }).filter(score => score > 0);
    
    if (scores.length === 0) {
      return { average: 0, highest: 0, lowest: 0, passCount: 0, failCount: 0 };
    }
    
    const total = scores.reduce((sum, score) => sum + score, 0);
    const average = Math.round(total / scores.length * 100) / 100;
    const highest = Math.max(...scores);
    const lowest = Math.min(...scores);
    const passCount = scores.filter(score => score >= 50).length;
    const failCount = scores.filter(score => score < 50).length;
    
    return {
      average,
      highest,
      lowest,
      passCount,
      failCount,
      totalStudents: scores.length
    };
    
  } catch (error) {
    console.error('❌ calculateScoreStatistics error:', error);
    return { average: 0, highest: 0, lowest: 0, passCount: 0, failCount: 0 };
  }
}

/**
 * 🎯 ปรับปรุง createSubjectReportHTML ให้มีสถิติและข้อมูลครู
 */
function createSubjectReportHTMLWithStats(data) {
  const {
    subjectName, subjectCode, grade, classNo, term, 
    scoreData, settings, reportDate
  } = data;
  
  const termText = term === 'both' ? 'ทั้งปีการศึกษา' :
                   term === '1' ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
  
  const gradeFullName = U_getGradeFullName(grade);
  const teacher = getSubjectTeacher(subjectName, subjectCode, grade);
  const stats = calculateScoreStatistics(scoreData.students, term);
  
  // สร้างแถวข้อมูลนักเรียน
  const studentRows = scoreData.students.map((student, index) => {
    if (term === 'both') {
      // แสดงทั้งสองภาค
      return `
        <tr>
          <td class="center">${index + 1}</td>
          <td class="student-name">${student.name}</td>
          <td class="center">${student.term1.totalScore}</td>
          <td class="center">${student.term1.grade}</td>
          <td class="center">${student.term2.totalScore}</td>
          <td class="center">${student.term2.grade}</td>
          <td class="center highlight">${student.yearAverage}</td>
          <td class="center highlight">${student.finalGrade}</td>
        </tr>`;
    } else {
      // แสดงภาคเรียนเดียว
      const termData = term === '1' ? student.term1 : student.term2;
      return `
        <tr>
          <td class="center">${index + 1}</td>
          <td class="center">${student.studentId}</td>
          <td class="student-name">${student.name}</td>
          <td class="center">${termData.midtermScore}</td>
          <td class="center">${termData.finalScore}</td>
          <td class="center highlight">${termData.totalScore}</td>
          <td class="center highlight">${termData.grade}</td>
        </tr>`;
    }
  }).join('');
  
  // สร้างส่วนหัวตาราง
  const tableHeader = term === 'both' ? `
    <tr>
      <th rowspan="2">ที่</th>
      <th rowspan="2">ชื่อ - นามสกุล</th>
      <th colspan="2">ภาคเรียนที่ 1</th>
      <th colspan="2">ภาคเรียนที่ 2</th>
      <th rowspan="2">คะแนนเฉลี่ย</th>
      <th rowspan="2">เกรด</th>
    </tr>
    <tr>
      <th>คะแนน</th><th>เกรด</th>
      <th>คะแนน</th><th>เกรด</th>
    </tr>` : `
    <tr>
      <th>ที่</th>
      <th>รหัสนักเรียน</th>
      <th>ชื่อ - นามสกุล</th>
      <th>ระหว่างภาค</th>
      <th>ปลายภาค</th>
      <th>รวม</th>
      <th>เกรด</th>
    </tr>`;
  
  // สร้างส่วนสถิติ
  const statisticsSection = `
    <div class="statistics">
      <h3>สรุปผลการเรียน</h3>
      <div class="stats-grid">
        <div class="stat-item">
          <span class="stat-label">จำนวนนักเรียนทั้งหมด:</span>
          <span class="stat-value">${stats.totalStudents} คน</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">คะแนนเฉลี่ย:</span>
          <span class="stat-value">${stats.average}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">คะแนนสูงสุด:</span>
          <span class="stat-value">${stats.highest}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">คะแนนต่ำสุด:</span>
          <span class="stat-value">${stats.lowest}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">ผ่าน (≥50):</span>
          <span class="stat-value pass">${stats.passCount} คน</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">ไม่ผ่าน (<50):</span>
          <span class="stat-value fail">${stats.failCount} คน</span>
        </div>
      </div>
    </div>
  `;
  
  return `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <title>รายงานผลการเรียนรายวิชา</title>
      <style>
        @page { size: A4; margin: 1cm; }
        body {
          font-family: 'TH Sarabun New', 'Sarabun', sans-serif;
          font-size: 14pt;
          line-height: 1.4;
          margin: 0;
          padding: 0;
        }
        .header {
          text-align: center;
          margin-bottom: 20px;
          border-bottom: 2px solid #333;
          padding-bottom: 15px;
        }
        .school-name {
          font-size: 18pt;
          font-weight: bold;
          margin-bottom: 5px;
        }
        .report-title {
          font-size: 16pt;
          font-weight: bold;
          margin-bottom: 10px;
        }
        .subject-info {
          font-size: 14pt;
          margin-bottom: 5px;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin: 15px 0;
        }
        th, td {
          border: 1px solid #333;
          padding: 8px 4px;
          text-align: center;
          vertical-align: middle;
        }
        th {
          background-color: #f0f0f0;
          font-weight: bold;
          font-size: 12pt;
        }
        td {
          font-size: 11pt;
        }
        .student-name {
          text-align: left;
          padding-left: 8px;
          min-width: 150px;
        }
        .center { text-align: center; }
        .highlight {
          background-color: #fffacd;
          font-weight: bold;
        }
        .statistics {
          background-color: #f8f9fa;
          border: 1px solid #dee2e6;
          border-radius: 5px;
          padding: 15px;
          margin: 20px 0;
        }
        .statistics h3 {
          margin: 0 0 10px 0;
          font-size: 14pt;
          text-align: center;
        }
        .stats-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 8px;
        }
        .stat-item {
          display: flex;
          justify-content: space-between;
          padding: 5px 0;
        }
        .stat-label {
          font-weight: bold;
        }
        .stat-value {
          color: #333;
        }
        .stat-value.pass {
          color: #28a745;
          font-weight: bold;
        }
        .stat-value.fail {
          color: #dc3545;
          font-weight: bold;
        }
        .signature {
          margin-top: 40px;
          display: flex;
          justify-content: space-between;
        }
        .signature-box {
          text-align: center;
          width: 200px;
        }
        .signature-line {
          border-bottom: 1px solid #333;
          margin-bottom: 5px;
          height: 40px;
        }
        .footer {
          margin-top: 20px;
          text-align: right;
          font-size: 10pt;
          color: #666;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="school-name">${settings.schoolName}</div>
        <div class="report-title">รายงานผลการเรียนรายวิชา</div>
        <div class="subject-info">
          รายวิชา: ${subjectName} ${subjectCode ? `(${subjectCode})` : ''}
        </div>
        <div class="subject-info">
          ระดับชั้น: ${gradeFullName} ห้อง ${classNo} | ${termText}
        </div>
        <div class="subject-info">
          ปีการศึกษา: ${settings.academicYear} | ครูผู้สอน: ${teacher}
        </div>
      </div>
      
      <table>
        <thead>${tableHeader}</thead>
        <tbody>${studentRows}</tbody>
      </table>
      
      ${statisticsSection}
      
      <div class="signature">
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>${teacher}<br/>ครูผู้สอน</div>
        </div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div>${settings.directorName}<br/>${settings.directorTitle}</div>
        </div>
      </div>
      
      ${settings.footerNote ? `<div style="text-align: center; margin-top: 20px; font-style: italic;">${settings.footerNote}</div>` : ''}
      
      <div class="footer">
        วันที่พิมพ์: ${Utilities.formatDate(reportDate, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm')}
      </div>
    </body>
    </html>
  `;
}

/**
 * 🆕 อัปเดตข้อมูลรายชื่อนักเรียนในชีตที่มีอยู่แล้ว
 * (เพิ่มนักเรียนใหม่ ไม่ลบคะแนนเดิม)
 * @param {Sheet} existingSheet - ชีตที่มีอยู่แล้ว
 * @param {Array<Object>} students - รายชื่อนักเรียนใหม่
 * @returns {string} ข้อความสถานะ
 */
function updateExistingScoreSheet(existingSheet, students) {
  try {
    console.log('อัปเดตชีตที่มีอยู่แล้ว:', existingSheet.getName());
    
    // ดึงข้อมูลนักเรียนเดิมจากชีต
    const existingData = existingSheet.getDataRange().getValues();
    const existingStudents = new Set();
    
    // เก็บรายชื่อนักเรียนเดิม (เริ่มจากแถวที่ 5)
    for (let i = 4; i < existingData.length; i++) {
      const studentId = existingData[i][1]; // Column B = เลขประจำตัว
      if (studentId) {
        existingStudents.add(String(studentId).trim());
      }
    }
    
    // หานักเรียนใหม่ที่ยังไม่มีในชีต
    const newStudents = students.filter(student => {
      const studentId = String(student.studentId || student.id || '').trim();
      return studentId && !existingStudents.has(studentId);
    });
    
    if (newStudents.length === 0) {
      return `ชีต "${existingSheet.getName()}" มีรายชื่อนักเรียนครบแล้ว ไม่มีการเปลี่ยนแปลง`;
    }
    
    // เพิ่มนักเรียนใหม่ที่ด้านล่าง
    const lastRow = existingSheet.getLastRow();
    const startRow = lastRow + 1;
    
    const newStudentRows = newStudents.map((student, index) => [
      existingStudents.size + index + 1, // ลำดับต่อจากเดิม
      student.studentId || student.id || '', // เลขประจำตัว
      student.name || `${student.firstname || ''} ${student.lastname || ''}`.trim() // ชื่อ-นามสกุล
    ]);
    
    // ใส่ข้อมูลนักเรียนใหม่
    existingSheet.getRange(startRow, 1, newStudentRows.length, 3).setValues(newStudentRows);
    
    // ใส่สีพื้นหลัง
    for (let i = 0; i < newStudentRows.length; i++) {
      const rowNum = startRow + i;
      const bgColor = (existingStudents.size + i) % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
      const totalCols = existingSheet.getLastColumn();
      existingSheet.getRange(rowNum, 1, 1, totalCols).setBackground(bgColor);
    }
    
    const message = `อัปเดตชีต "${existingSheet.getName()}" เรียบร้อย: เพิ่มนักเรียนใหม่ ${newStudents.length} คน`;
    console.log(message);
    return message;
    
  } catch (error) {
    console.log('Error in updateExistingScoreSheet:', error.message);
    throw new Error('ไม่สามารถอัปเดตชีตได้: ' + error.message);
  }
}

/**
 * 🆕 ดึงรายชื่อนักเรียนตามระดับชั้นและห้องเรียน
 * ฟังก์ชันนี้ถูกเรียกจากหน้าเว็บ
 * @param {string} grade - ระดับชั้น เช่น "ป.1"
 * @param {string} classNo - หมายเลขห้อง เช่น "1" 
 * @returns {Array<Object>} รายชื่อนักเรียน
 */
/**
 * 🆕 สร้างตารางคะแนนพร้อมรายชื่อนักเรียน (แก้ไขแล้ว - ป้องกันชื่อซ้ำ)
 * ฟังก์ชันนี้ถูกเรียกจากหน้าเว็บแทน createScoreSheet()
 * @param {string} subjectCode - รหัสวิชา
 * @param {string} subjectName - ชื่อวิชา  
 * @param {number|string} hour - จำนวนชั่วโมง
 * @param {string} grade - ระดับชั้น
 * @param {string} classNo - หมายเลขห้อง
 * @param {Array<Object>} students - รายชื่อนักเรียน
 * @returns {string} ข้อความสถานะ
 */
function createScoreSheetWithStudents(subjectCode, subjectName, hour, grade, classNo, students) {
  try {
    // --- (ส่วนตรวจสอบเดิม คงไว้ตามโค้ดคุณ) ---

    if (!getSpreadsheetId_()) {
      throw new Error("SPREADSHEET_ID ไม่ได้ถูกกำหนด");
    }
    if (!subjectCode || !subjectName || !grade || !classNo) {
      throw new Error("กรุณากรอกข้อมูลให้ครบถ้วน");
    }
    if (!students || !Array.isArray(students) || students.length === 0) {
      throw new Error("ไม่พบรายชื่อนักเรียนในห้องนี้");
    }

    const cleanGrade = grade.replace(/[\/\.]/g, '');
    const baseSheetName = `${subjectName} ${cleanGrade}-${classNo}`;

    const ss = SS();
    const existingSheet = ss.getSheetByName(baseSheetName);

    // ⛔ เปลี่ยนพฤติกรรม: ถ้ามีชีตอยู่แล้ว → อัปเดต ไม่สร้าง _2
    if (existingSheet) {
      // เรียงนักเรียนก่อน (เหมือนเดิม)
      const sortedStudents = [...students].sort((a, b) => {
        const idA = (a.studentId || a.id || '').toString();
        const idB = (b.studentId || b.id || '').toString();
        return idA.localeCompare(idB, 'th', { numeric: true });
      });

      // ใช้ฟังก์ชันอัปเดตที่คุณมีอยู่แล้ว (จะเพิ่มเฉพาะคนที่ยังไม่มี)
      const msg = updateExistingScoreSheet(existingSheet, sortedStudents);
      return `✅ พบชีตเดิม "${baseSheetName}" และได้อัปเดตรายชื่อนักเรียนแล้ว: ${msg}`;
    }

    // ▼▼▼ ส่วน “สร้างชีตใหม่” ทำงานเฉพาะกรณีไม่พบชีตเดิมเท่านั้น ▼▼▼
    const sheet = ss.insertSheet(baseSheetName);

    // โครงสร้างตาม SchoolMIS (16 คอลัมน์ต่อภาค)
    const header3 = [
      "ลำดับ", "เลขประจำตัว", "ชื่อ - สกุล",
      // Term 1 (16 cols: D-S, index 3-18)
      "ครั้งที่1","ครั้งที่2","ครั้งที่3","ครั้งที่4","รวม",
      "ครั้งที่5","แก้ตัวกลางภาค",
      "ครั้งที่6","ครั้งที่7","ครั้งที่8","ครั้งที่9","รวม",
      "รวมระหว่างภาค","ครั้งที่10","รวมทั้งหมด","เกรด",
      // Term 2 (16 cols: T-AI, index 19-34)
      "ครั้งที่1","ครั้งที่2","ครั้งที่3","ครั้งที่4","รวม",
      "ครั้งที่5","แก้ตัวกลางภาค",
      "ครั้งที่6","ครั้งที่7","ครั้งที่8","ครั้งที่9","รวม",
      "รวมระหว่างภาค","ครั้งที่10","รวมทั้งหมด","เกรด",
      // Summary (2 cols: AJ-AK, index 35-36)
      "คะแนนเฉลี่ย 2 ภาค", "ระดับผลการเรียน"
    ];

    sheet.getRange(1, 1).setValue("รหัสวิชา").setFontWeight("bold").setBackground("#E8F4FD");
    sheet.getRange(1, 2).setValue(subjectCode).setBackground("#E8F4FD");
    sheet.getRange(2, 1).setValue("ชื่อวิชา").setFontWeight("bold").setBackground("#E8F4FD");
    sheet.getRange(2, 2).setValue(subjectName).setBackground("#E8F4FD");
    sheet.getRange(2, 3).setValue(`${hour} ชั่วโมง/ปี`).setBackground("#E8F4FD");

    sheet.getRange(3, 1, 1, header3.length)
      .setValues([header3]).setFontWeight("bold").setBackground("#D4EDDA");

    sheet.getRange(4, 3).setValue("คะแนนเต็ม").setFontWeight("bold").setBackground("#FFF3CD");

    const sortedStudents = [...students].sort((a, b) => {
      const idA = (a.studentId || a.id || '').toString();
      const idB = (b.studentId || b.id || '').toString();
      return idA.localeCompare(idB, 'th', { numeric: true });
    });

    const studentRows = sortedStudents.map((student, index) => {
      let fullName = '';
      if (student.name && student.name.trim()) fullName = student.name.trim();
      else if (student.title && student.firstname && student.lastname) {
        fullName = `${student.title}${student.firstname} ${student.lastname}`.trim();
      } else if (student.firstname && student.lastname) {
        fullName = `${student.firstname} ${student.lastname}`.trim();
      } else if (student.firstname) {
        fullName = student.firstname.trim();
      } else {
        fullName = `ไม่มีชื่อ (${student.id || student.studentId || index + 1})`;
      }
      return [index + 1, student.studentId || student.id || '', fullName];
    });

    if (studentRows.length) {
      sheet.getRange(5, 1, studentRows.length, 3).setValues(studentRows);
    }

    for (let i = 0; i < studentRows.length; i++) {
      const rowNum = 5 + i;
      const bgColor = i % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
      sheet.getRange(rowNum, 1, 1, header3.length).setBackground(bgColor);
    }

    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(3);
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 200);
    for (let col = 4; col <= header3.length; col++) sheet.setColumnWidth(col, 80);

    return `✅ สร้างชีต "${baseSheetName}" สำเร็จแล้ว พร้อมข้อมูลนักเรียน ${students.length} คน (เรียงตามรหัสนักเรียน)`;

  } catch (error) {
    console.log("❌ Error in createScoreSheetWithStudents: " + error.message);
    console.log("Stack trace:", error.stack);
    throw new Error("เกิดข้อผิดพลาดในการสร้าง/อัปเดตตาราง: " + error.message);
  }
}




/**
 * Helper: คำนวณขนาดโลโก้ที่เหมาะสม (กว้าง x สูง) ในหน่วยพิกเซล
 */
function computeLogoSize_(totalWidth, settings) {
  // --- ✅ [แก้ไข] ปรับขนาดโลโก้ของคุณที่นี่ ---
  const LOGO_HEIGHT = 120; // ✨ เดิม 90

  const LOGO_ASPECT_RATIO = 1.0; 

  return {
    width: LOGO_HEIGHT * LOGO_ASPECT_RATIO,
    height: LOGO_HEIGHT
  };
}

/**
 * วางหัวรายงาน - ✨ [ปรับปรุง] ปรับขนาดตัวอักษรหัวกระดาษให้สมดุล
 */
function placeHeader_(sheet, settings, meta, colCount) {
  const FONT = 'Sarabun';
  const FS_SCHOOL = 18;
  const FS_HEAD = 14;
  const LOGO_ROW_HEIGHT = 120;

  sheet.setRowHeight(1, LOGO_ROW_HEIGHT); 
  sheet.setRowHeight(2, 30);
  sheet.setRowHeight(3, 28);

  // Merge ทั้งแถวก่อนเพื่อให้โลโก้อยู่ตรงกลาง
  const row1 = sheet.getRange(1, 1, 1, colCount).merge();
  row1.setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ใส่โลโก้ตรงกลาง
  try {
    if (settings && settings.logoFileId) {
      const file = DriveApp.getFileById(settings.logoFileId);
      const blob = file.getBlob();
      
      // คำนวณความกว้างรวมของตาราง
      let totalWidth = 0;
      for (let c = 1; c <= colCount; c++) {
        totalWidth += sheet.getColumnWidth(c);
      }
      
      // ขนาดโลโก้
      const logoWidth = 110;
      const logoHeight = 110;
      
      // คำนวณตำแหน่งกึ่งกลาง
      const centerX = (totalWidth - logoWidth) / 2;
      
      const img = sheet.insertImage(blob, 1, 1);
      img.setWidth(logoWidth)
         .setHeight(logoHeight)
         .setAnchorCell(sheet.getRange(1, 1))
         .setAnchorCellXOffset(centerX)
         .setAnchorCellYOffset(5); // เว้นจากขอบบน 5px
    }
  } catch (e) {
    Logger.log('placeHeader_ logo error: ' + e);
  }

  // หัวข้อแถวที่ 2
  sheet.getRange(2, 1, 1, colCount).merge()
    .setValue(`ผลการประเมิน รายวิชา: ${meta.subjectName} (${meta.subjectCode})`)
    .setFontFamily(FONT).setFontSize(FS_SCHOOL).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
    
  // หัวข้อแถวที่ 3
  const year = (settings && settings['ปีการศึกษา']) ? settings['ปีการศึกษา'] : '';
  const extra = year ? ` | ปีการศึกษา ${year}` : '';
  
  sheet.getRange(3, 1, 1, colCount).merge()
    .setValue(`${meta.grade}/${meta.classNo} | ${meta.termName}${extra}`)
    .setFontFamily(FONT).setFontSize(FS_HEAD)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
    
  return 5;
}
/***********************
 * 2) ส่งออกภาคเรียนเดียว (ไม่มีคอลัมน์ 1–10)
 ***********************/


/**
 * Helper: เรียก URL Export ของ Google Sheets → ได้ PDF Blob
 * @param {string} spreadsheetId
 * @param {number} sheetId
 * @param {object} opt {filename, landscape, printGridlines, scale, top,right,bottom,left}
 */
/**
 * สร้าง PDF Blob จากชีตด้วย endpoint export (แก้ 400)
 * - ใช้พารามิเตอร์เท่าที่จำเป็น/เชื่อถือได้
 * - margin หน่วย "นิ้ว" (ตัวเลขทศนิยม), ไม่ใส่ mm
 * - ตรวจ HTTP/Content-Type และโยน error พร้อมข้อความย่อ
 */
/**
 * ใช้แทน exportSheetToPdfBlob_ เดิม: เพิ่ม flush, retry และ fallback
 */
function exportSheetToPdfBlob_(spreadsheetId, sheetId, opt) {
  SpreadsheetApp.flush(); // คอมมิตการแก้ไขทั้งหมดให้เสร็จ
  Utilities.sleep(300);   // กัน race condition เบาๆ

  opt = opt || {};
  var urlBase = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?';

  var params = {
    format: 'pdf',
    size: 'A4',
    portrait: (opt.landscape ? 'false' : 'true'),
    fitw: 'true',
    sheetnames: 'false',
    printtitle: 'false',
    pagenum: 'RIGHT',
    gridlines: (opt.printGridlines ? 'true' : 'false'),
    fzr: 'true',
    gid: String(sheetId),
    top_margin:   String(typeof opt.top    === 'number' ? opt.top    : 0.5), // นิ้ว
    bottom_margin:String(typeof opt.bottom === 'number' ? opt.bottom : 0.5),
    left_margin:  String(typeof opt.left   === 'number' ? opt.left   : 0.5),
    right_margin: String(typeof opt.right  === 'number' ? opt.right  : 0.5),
    scale:        String(opt.scale || 2)
  };

  var query = Object.keys(params).map(function(k){ return k + '=' + encodeURIComponent(params[k]); }).join('&');
  var token = ScriptApp.getOAuthToken();

  // ===== 1) RETRY กับ export URL เดิม =====
  var attempts = 3, lastErr;
  for (var i = 0; i < attempts; i++) {
    try {
      var res = UrlFetchApp.fetch(urlBase + query, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      var headers = res.getAllHeaders();
      var ctype = (headers['Content-Type'] || headers['content-type'] || '') + '';

      if (code === 200 && ctype.indexOf('application/pdf') !== -1) {
        return res.getBlob()
                  .setName(((opt && opt.filename) || 'scores') + '.pdf')
                  .setContentType('application/pdf');
      }

      // ถ้า 5xx ให้ลองใหม่ (backoff)
      if (String(code).startsWith('5')) {
        lastErr = new Error('HTTP ' + code + ' — ' + res.getContentText().slice(0, 300));
        Utilities.sleep(500 * Math.pow(2, i)); // 0.5s, 1s, 2s
        continue;
      }

      // อื่นๆ (เช่น 4xx หรือ content-type เพี้ยน) โยนเลย
      throw new Error('Export failed with HTTP ' + code + ' — ' + res.getContentText().slice(0, 300));
    } catch (e) {
      lastErr = e;
      // ถ้าเจอ network error ที่ไม่ใช่ response ให้ backoff เช่นกัน
      Utilities.sleep(500 * Math.pow(2, i));
    }
  }

  // ===== 2) FALLBACK: copy ชีตไป Spreadsheet ใหม่ แล้ว export ทั้งไฟล์ =====
  try {
    var blob = fallbackExportByCopy_(spreadsheetId, sheetId, opt);
    if (blob) return blob;
    throw lastErr || new Error('Unknown error in fallback export.');
  } catch (e2) {
    // ส่ง error รายละเอียดกลับไปให้ frontend
    throw new Error('Fallback export failed — ' + (e2 && e2.message ? e2.message : e2));
  }
}

/**
 * คัดลอกชีตไปยังสเปรดชีตว่างใบใหม่ → export ทั้งไฟล์
 * หมายเหตุ: วิธีนี้จะไม่ได้ตั้งค่ามาร์จิน/เลขหน้าได้ละเอียดเท่าวิธีหลัก
 */
function fallbackExportByCopy_(spreadsheetId, sheetId, opt) {
  var ssSrc = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ssSrc.getSheets().filter(function(s){ return s.getSheetId() === sheetId; })[0];
  if (!sheet) throw new Error('Source sheet not found for gid=' + sheetId);

  SpreadsheetApp.flush();
  Utilities.sleep(200);

  // สร้างไฟล์ใหม่ แล้วคัดลอกชีตเป้าหมายไปลงเป็นชีตเดียว
  var tmp = SpreadsheetApp.create('TMP_EXPORT_' + Date.now());
  var tmpId = tmp.getId();
  var copied = sheet.copyTo(tmp).setName(sheet.getName());
  // ลบชีตแรกที่ติดมากับไฟล์ใหม่
  tmp.deleteSheet(tmp.getSheets()[0]);

  SpreadsheetApp.flush();
  Utilities.sleep(300);

  // พยายาม export ทั้งไฟล์ผ่าน Drive API (Advanced Service) ก่อน
  // (เปิดใช้งาน Advanced Drive Service ใน Apps Script: Services > Drive API)
  try {
    var blob = Drive.Files.export(tmpId, 'application/pdf');
    if (blob && blob.getBytes().length > 0) {
      return blob.setName(((opt && opt.filename) || 'scores') + '.pdf')
                  .setContentType('application/pdf');
    }
  } catch (e) {
    // ถ้า Advanced Drive ใช้ไม่ได้ จะลองวิธี URL fetch แบบทั้งไฟล์
  }

  // ถ้า Drive.Files.export ใช้ไม่ได้ ใช้ export URL แบบทั้งไฟล์ (ไม่ระบุ gid)
  var url = 'https://docs.google.com/spreadsheets/d/' + tmpId + '/export?format=pdf&size=A4&portrait='
            + (opt.landscape ? 'false' : 'true') + '&fitw=true&sheetnames=false&printtitle=false&pagenum=RIGHT&gridlines='
            + (opt.printGridlines ? 'true' : 'false') + '&fzr=true&scale=' + (opt.scale || 2);

  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var headers = res.getAllHeaders();
  var ctype = (headers['Content-Type'] || headers['content-type'] || '') + '';
  if (code !== 200 || ctype.indexOf('application/pdf') === -1) {
    throw new Error('Fallback HTTP ' + code + ' — ' + res.getContentText().slice(0, 300));
  }

  return res.getBlob()
            .setName(((opt && opt.filename) || 'scores') + '.pdf')
            .setContentType('application/pdf');

} /* end fallback */

/**
 * 🆕 Wrapper สำหรับ Frontend - ส่งออก PDF ภาคเรียนเดียว
 */
function exportScoresPdf_SingleTerm(subjectName, subjectCode, grade, classNo, termId, termData) {
  try {
    Logger.log('=== exportScoresPdf_SingleTerm ===');
    Logger.log('Subject:', subjectName, subjectCode);
    Logger.log('Grade/Class:', grade, classNo);
    Logger.log('Term:', termId);
    Logger.log('Rows:', termData?.rows?.length || 0);
    
    if (!termData || !termData.rows || termData.rows.length === 0) {
      throw new Error('ไม่พบข้อมูลคะแนน');
    }
    
    var termName = termId === 'term1' ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
    
    var meta = {
      subjectName: subjectName,
      subjectCode: subjectCode || '',
      grade: grade,
      classNo: classNo,
      termName: termName
    };
    
    var result = exportScoresPdf_SingleTerm_(meta, termData);
    
    Logger.log('SingleTerm PDF generated successfully');
    return result.base64;
    
  } catch (error) {
    Logger.log('❌ exportScoresPdf_SingleTerm error:', error);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

/**
 * 🆕 Wrapper สำหรับ Frontend - ส่งออก PDF รวม 2 ภาคเรียน
 */
function exportScoresPdf_Year(subjectName, subjectCode, grade, classNo, term1Data, term2Data) {
  try {
    Logger.log('=== exportScoresPdf_Year ===');
    Logger.log('Subject:', subjectName, subjectCode);
    Logger.log('Grade/Class:', grade, classNo);
    Logger.log('Term1 rows:', term1Data?.rows?.length || 0);
    Logger.log('Term2 rows:', term2Data?.rows?.length || 0);
    
    if (!term1Data || !term1Data.rows || term1Data.rows.length === 0) {
      throw new Error('ไม่พบข้อมูลภาคเรียนที่ 1');
    }
    if (!term2Data || !term2Data.rows || term2Data.rows.length === 0) {
      throw new Error('ไม่พบข้อมูลภาคเรียนที่ 2');
    }
    
    const mergedRows = term1Data.rows.map((t1Row, index) => {
      const t2Row = term2Data.rows[index] || {};
      
      const term1Total = Number(t1Row.total || 0);
      const term2Total = Number(t2Row.total || 0);
      const yearTotal = term1Total + term2Total;
      const yearPercent = yearTotal > 0 ? Math.round(yearTotal / 2) : 0;
      const yearGrade = calculateFinalGrade(yearPercent);

      return {
        no: index + 1,
        id: t1Row.id || '',
        name: t1Row.name || '',  // ✅ เพิ่มบรรทัดนี้
        term1: { total: term1Total, grade: t1Row.grade || '0' },
        term2: { total: term2Total, grade: t2Row.grade || '0' },
        year: { total: yearTotal, percent: yearPercent, grade: yearGrade }
      };
    });
    
    const meta = {
      subjectName: subjectName,
      subjectCode: subjectCode || '', // ⭐ ใช้ค่าที่ส่งมา
      grade: grade,
      classNo: classNo,
      termName: 'รวมทั้งปีการศึกษา'
    };
    
    const result = exportScoresPdf_Year_(meta, mergedRows);
    
    Logger.log('Year PDF generated successfully');
    return result.base64;
    
  } catch (error) {
    Logger.log('❌ exportScoresPdf_Year error:', error);
    throw new Error('ไม่สามารถสร้าง PDF ได้: ' + error.message);
  }
}

/**
 * เปลี่ยนชื่อฟังก์ชันเดิมเป็น internal functions (เพิ่ม _ ท้ายชื่อ)
 */
function exportScoresPdf_SingleTerm_(meta, data) {
  // โค้ดเดิมของ exportScoresPdf_SingleTerm
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet('TMP_PDF_' + Date.now());
  try {
    const settings = S_getGlobalSettings();
    const FONT = 'Sarabun';
    const FS_TABLE = 16;

    const fullMid   = Number(data.fullScores?.reduce((a,b)=>a+(b||0),0) || 0);
    const fullFinal = Number(data.fullFinal || 0);
    const fullTotal = fullMid + fullFinal;

    sheet.insertColumnBefore(1);
    sheet.setColumnWidth(1, 30);

    sheet.setColumnWidths(2, 1, 60);
    sheet.setColumnWidths(3, 1, 110);
    sheet.setColumnWidths(4, 1, 300);
    sheet.setColumnWidths(5, 3, 100);

    const HEADER = [
      'เลขที่','รหัส','ชื่อ-นามสกุล',
      `กลางภาค ${fullMid}`, `ปลายภาค ${fullFinal}`, `รวม ${fullTotal}`, 'เกรด'
    ];
    const startRow = placeHeader_(sheet, settings, meta, HEADER.length + 1);

    sheet.getRange(startRow, 2, 1, HEADER.length).setValues([HEADER])
      .setFontFamily(FONT).setFontSize(FS_TABLE).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground('#eaeff7');

    const values = (data.rows || []).map((r,i)=>[
      '', 
      r.no || (i+1), r.id || '', r.name || '',
      Number(r.mid||0), Number(r.final||0), Number(r.total||0), (r.grade||'0')
    ]);
    
    if (values.length) {
      sheet.getRange(startRow + 1, 1, values.length, HEADER.length + 1)
        .setValues(values)
        .setFontFamily(FONT).setFontSize(FS_TABLE);
    }

    const lastRow = sheet.getLastRow();
    sheet.getRange(startRow, 2, Math.max(1 + values.length, 1), HEADER.length)
      .setBorder(true,true,true,true,true,true,'#9aa4b2',SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow, 3, lastRow - startRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow, 5, lastRow - startRow + 1, 3).setHorizontalAlignment('center');
    sheet.getRange(startRow, 8, lastRow - startRow + 1, 1).setHorizontalAlignment('center');

    sheet.getRange(1,1,lastRow,HEADER.length + 1).setFontFamily(FONT);
    sheet.setFrozenRows(startRow);

    SpreadsheetApp.flush();
    Utilities.sleep(200);

    const blob = exportSheetToPdfBlob_(ss.getId(), sheet.getSheetId(), {
      filename: `${meta.subjectName}_${meta.grade}-${meta.classNo}_${meta.termName}`.replaceAll(' ','_'),
      landscape: false, printGridlines: false, scale: 2,
      top: 0.5, right: 0.5, bottom: 0.5, left: 1.0
    });
    
    return { base64: Utilities.base64Encode(blob.getBytes()) };
  } finally {
    ss.deleteSheet(sheet);
  }
}

function exportScoresPdf_Year_(meta, mergedRows) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.insertSheet('TMP_PDF_YEAR_' + Date.now());
  try {
    const settings = S_getGlobalSettings();
    const FONT = 'Sarabun';
    const FS_HEADER = 22;  // ⭐ เพิ่มเป็น 22 (จาก 18)
    const FS_DATA = 20;    // ⭐ เพิ่มเป็น 20 (จาก 18)

    const HEADER = [
      'เลขที่', 'รหัส', 'ชื่อ-นามสกุล',
      'คะแนนภาคเรียนที่ 1', 'ผลการเรียน',
      'คะแนนภาคเรียนที่ 2', 'ผลการเรียน',
      'คะแนนรวม', 'คะแนนเฉลี่ย', 'ผลการเรียนตลอดปี'
    ];

    sheet.insertColumnBefore(1);
    sheet.setColumnWidth(1, 30);
    
    sheet.setColumnWidths(2, 1, 50);
    sheet.setColumnWidths(3, 1, 80);   // เพิ่มเล็กน้อย
    sheet.setColumnWidths(4, 1, 280);  // เพิ่มชื่อ-นามสกุล
    sheet.setColumnWidths(5, 7, 55);   // เพิ่มคอลัมน์คะแนนเล็กน้อย

    const startRow = placeHeader_(sheet, settings, meta, HEADER.length + 1);
    
    sheet.setRowHeight(startRow, 200);  // ⭐ เพิ่มเป็น 200 (จาก 125)

    const headerRange = sheet.getRange(startRow, 2, 1, HEADER.length);
    headerRange.setValues([HEADER])
      .setFontFamily(FONT)
      .setFontSize(FS_HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground('#eaeff7')
      .setWrap(true);

    sheet.getRange(startRow, 5, 1, 7).setTextRotation(90);

    const values = (mergedRows || []).map((r,i)=>[
      '', 
      r.no || (i+1), r.id || '', r.name || '',
      Number(r.term1?.total||0), r.term1?.grade||'0',
      Number(r.term2?.total||0), r.term2?.grade||'0',
      Number(r.year?.total||0), Number(r.year?.percent||0), r.year?.grade||'0'
    ]);
    
    if (values.length) {
      const dataRange = sheet.getRange(startRow + 1, 1, values.length, HEADER.length + 1);
      dataRange.setValues(values)
        .setFontFamily(FONT)
        .setFontSize(FS_DATA)
        .setVerticalAlignment('middle');
      
      for (let i = 0; i < values.length; i++) {
        sheet.setRowHeight(startRow + 1 + i, 34);  // ⭐ เพิ่มเป็น 32 (จาก 24)
      }
    }
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(startRow, 2, Math.max(1 + values.length, 1), HEADER.length)
      .setBorder(true,true,true,true,true,true,'#9aa4b2',SpreadsheetApp.BorderStyle.SOLID);
      
    sheet.getRange(startRow, 2, lastRow - startRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow, 3, lastRow - startRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(startRow, 5, lastRow - startRow + 1, 7).setHorizontalAlignment('center');
    sheet.setFrozenRows(startRow);

    SpreadsheetApp.flush();
    Utilities.sleep(300);

    const blob = exportSheetToPdfBlob_(ss.getId(), sheet.getSheetId(), {
      filename: `${meta.subjectName}_${meta.grade}-${meta.classNo}_${meta.termName}`.replaceAll(' ','_'),
      landscape: false, 
      printGridlines: false, 
      scale: 2,  // ⭐ สำคัญมาก: เปลี่ยนเป็น 2 (จาก 4)
      top: 0.3,       
      right: 0.4,
      bottom: 0.4,
      left: 1.0
    });
    
    return { base64: Utilities.base64Encode(blob.getBytes()) };
  } finally {
    ss.deleteSheet(sheet);
  }
}
/**
 * ✅ ตรวจสอบว่าชีตหลายตัวมีอยู่หรือไม่
 * @param {Array<string>} sheetNames - ชื่อชีตที่ต้องการตรวจสอบ
 * @returns {Array<string>} - รายการชีตที่มีอยู่
 */
function checkMultipleSheets(sheetNames) {
  try {
    Logger.log(`=== checkMultipleSheets: ${sheetNames.length} sheets ===`);
    
    if (!sheetNames || sheetNames.length === 0) {
      return [];
    }
    
    if (!getSpreadsheetId_()) {
      Logger.log("❌ ไม่พบ SPREADSHEET_ID");
      return [];
    }
    
    const ss = SS();
    const existingSheets = [];
    
    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        existingSheets.push(sheetName);
        Logger.log(`✅ Found: ${sheetName}`);
      } else {
        Logger.log(`❌ Not found: ${sheetName}`);
      }
    });
    
    Logger.log(`Found ${existingSheets.length}/${sheetNames.length} existing sheets`);
    return existingSheets;
    
  } catch (error) {
    Logger.log('❌ Error in checkMultipleSheets: ' + error.message);
    return [];
  }
}

/**
 * =====================================================
 * 💾 BACKUP: สำรองชีตคะแนนก่อนปรับโครงสร้าง
 * =====================================================
 */

/**
 * สำรองข้อมูลชีตคะแนนทั้งหมดไปเป็นชีต BACKUP_xxx
 * @returns {Object} ผลการสำรอง
 */
function backupAllScoreSheets() {
  try {
    const ss = SS();
    const allSheets = ss.getSheets();
    const excludeSheets = ['Students', 'รายวิชา', 'Users', 'users', 'Settings', 'Attendance', 'global_settings', 'การตั้งค่าระบบ', 'Holidays', 'วันหยุด', 'SCORES_WAREHOUSE', 'AttendanceLog', 'HomeroomTeachers', 'Template_', 'BACKUP_'];
    
    let backed = 0;
    let skipped = 0;
    
    allSheets.forEach(function(sheet) {
      var name = sheet.getName();
      
      // ข้ามชีตที่ไม่ใช่ตารางคะแนน
      if (excludeSheets.some(function(ex) { return name === ex || name.startsWith(ex); })) {
        return;
      }
      
      // ตรวจสอบว่าเป็นตารางคะแนน
      try {
        if (sheet.getRange(1, 1).getValue() !== 'รหัสวิชา') {
          skipped++;
          return;
        }
      } catch (_) {
        skipped++;
        return;
      }
      
      // สร้างชีต backup
      var backupName = 'BACKUP_' + name;
      var existing = ss.getSheetByName(backupName);
      if (existing) {
        ss.deleteSheet(existing);
      }
      
      var backup = sheet.copyTo(ss);
      backup.setName(backupName);
      backup.setTabColor('#FF9800');
      backed++;
      Logger.log('✅ สำรอง: ' + name + ' → ' + backupName);
    });
    
    return {
      success: true,
      backed: backed,
      skipped: skipped,
      message: '✅ สำรองข้อมูลสำเร็จ ' + backed + ' ชีต'
    };
    
  } catch (error) {
    Logger.log('❌ Error in backupAllScoreSheets: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ลบชีตรายวิชา (ตารางคะแนน) เก่าทั้งหมด
 * ยกเว้น: ชีตระบบ, BACKUP_, Template_
 * ⚠️ ต้องรัน backupAllScoreSheets() ก่อนเสมอ!
 * @returns {Object} ผลการลบ
 */
function deleteOldScoreSheets() {
  try {
    var ss = SS();
    var allSheets = ss.getSheets();
    var excludeSheets = ['Students', 'รายวิชา', 'Users', 'users', 'Settings', 'Attendance', 'global_settings', 'การตั้งค่าระบบ', 'Holidays', 'วันหยุด', 'SCORES_WAREHOUSE', 'AttendanceLog', 'HomeroomTeachers', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ', 'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ', 'ความเห็นครู', 'โปรไฟล์นักเรียน', 'BACKUP_WAREHOUSE_LATEST', 'สรุปวันมา', 'สรุปการมาเรียน', 'อ่านคิดเขียน', 'คุณลักษณะ', 'กิจกรรม', 'สมรรถนะ'];
    var prefixExclude = ['BACKUP_', 'Template_'];
    
    var deleted = 0;
    var skipped = 0;
    var deletedNames = [];
    
    allSheets.forEach(function(sheet) {
      var name = sheet.getName();
      
      // ข้ามชีตระบบ
      if (excludeSheets.some(function(ex) { return name === ex; })) {
        skipped++;
        return;
      }
      
      // ข้ามชีต BACKUP_ และ Template_
      if (prefixExclude.some(function(prefix) { return name.startsWith(prefix); })) {
        skipped++;
        return;
      }
      
      // ตรวจว่าเป็นชีตคะแนน (A1 = "รหัสวิชา")
      try {
        if (sheet.getRange(1, 1).getValue() !== 'รหัสวิชา') {
          skipped++;
          return;
        }
      } catch (_) {
        skipped++;
        return;
      }
      
      // ลบชีตคะแนนเก่า
      Logger.log('🗑️ ลบ: ' + name);
      deletedNames.push(name);
      ss.deleteSheet(sheet);
      deleted++;
    });
    
    Logger.log('✅ ลบชีตรายวิชาเก่า ' + deleted + ' ชีต, ข้าม ' + skipped + ' ชีต');
    return {
      success: true,
      deleted: deleted,
      skipped: skipped,
      deletedNames: deletedNames,
      message: '✅ ลบชีตรายวิชาเก่าสำเร็จ ' + deleted + ' ชีต'
    };
    
  } catch (error) {
    Logger.log('❌ Error in deleteOldScoreSheets: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * =====================================================
 * 🔄 MIGRATE: ย้ายคะแนนจาก BACKUP → ชีตใหม่ (SchoolMIS)
 * =====================================================
 * 
 * โครงสร้างเดิม (ต่อ 1 ภาค): 10 ช่องคะแนนย่อย + คะแนนระหว่างภาค + คะแนนปลายภาค + รวม + เกรด = 14 คอลัมน์
 *   Term1 เดิม: col 3-12 (D-M) = ช่อง1-10, col 13 = กลางภาค, col 14 = ปลายภาค, col 15 = รวม, col 16 = เกรด
 *   Term2 เดิม: col 17-26 (R-AA) = ช่อง11-20, col 27 = กลางภาค2, col 28 = ปลายภาค2, col 29 = รวม2, col 30 = เกรด2
 *
 * โครงสร้างใหม่ (ต่อ 1 ภาค): 16 คอลัมน์
 *   Term1 ใหม่: s1-s4(3-6), sum14(7), s5(8), makeup(9), s6-s9(10-13), sum69(14), midTotal(15), s10(16), total(17), grade(18)
 *   Term2 ใหม่: s1-s4(19-22), sum14(23), s5(24), makeup(25), s6-s9(26-29), sum69(30), midTotal(31), s10(32), total(33), grade(34)
 *
 * MAPPING (ผู้ใช้เลือก):
 *   เดิม ช่อง 1-4 → ครั้งที่ 1-4
 *   เดิม ช่อง 5   → ครั้งที่ 5 (กลางภาค)
 *   เดิม ช่อง 6-8 → ครั้งที่ 6-8
 *   เดิม "ปลายภาค" → ครั้งที่ 10 (ปลายภาค)
 *   เดิม ช่อง 9-10 → ไม่ใช้
 */
function migrateScoresFromBackup() {
  try {
    var ss = SS();
    var allSheets = ss.getSheets();
    var migrated = 0;
    var errors = [];
    
    allSheets.forEach(function(backupSheet) {
      var bName = backupSheet.getName();
      if (!bName.startsWith('BACKUP_')) return;
      
      // หาชีตปลายทาง (ชื่อเดียวกันไม่มี BACKUP_)
      var targetName = bName.replace('BACKUP_', '');
      var targetSheet = ss.getSheetByName(targetName);
      
      if (!targetSheet) {
        Logger.log('⚠️ ไม่พบชีตปลายทาง: ' + targetName + ' — ข้าม');
        return;
      }
      
      // ตรวจว่า BACKUP เป็นชีตคะแนน
      try {
        if (backupSheet.getRange(1, 1).getValue() !== 'รหัสวิชา') return;
      } catch (_) { return; }
      
      Logger.log('🔄 Migrate: ' + bName + ' → ' + targetName);
      
      var bData = backupSheet.getDataRange().getValues();
      
      // โครงสร้างเดิม (0-indexed)
      var OLD = {
        term1: { scores: 3, final: 14 },  // col D(3)-M(12)=ช่อง1-10, col O(14)=ปลายภาค
        term2: { scores: 17, final: 28 }   // col R(17)-AA(26)=ช่อง11-20, col AC(28)=ปลายภาค2
      };
      
      // โครงสร้างใหม่ (0-indexed, 16 คอลัมน์ต่อภาค)
      var NEW = {
        term1: {
          s1: 3, s2: 4, s3: 5, s4: 6, sum14: 7,
          s5: 8, makeup: 9,
          s6: 10, s7: 11, s8: 12, s9: 13, sum69: 14,
          midTotal: 15, s10: 16, total: 17, grade: 18
        },
        term2: {
          s1: 19, s2: 20, s3: 21, s4: 22, sum14: 23,
          s5: 24, makeup: 25,
          s6: 26, s7: 27, s8: 28, s9: 29, sum69: 30,
          midTotal: 31, s10: 32, total: 33, grade: 34
        }
      };
      
      // ย้ายคะแนนเต็ม (แถว 4, index 3)
      if (bData.length > 3) {
        migrateTermFullScores_(targetSheet, bData[3], OLD.term1, NEW.term1);
        migrateTermFullScores_(targetSheet, bData[3], OLD.term2, NEW.term2);
      }
      
      // ย้ายคะแนนนักเรียน (แถว 5+, index 4+)
      // สร้าง map ของนักเรียนในชีตปลายทาง (by studentId)
      var tData = targetSheet.getDataRange().getValues();
      var targetMap = {};
      for (var t = 4; t < tData.length; t++) {
        var tid = String(tData[t][1]).trim();
        if (tid) targetMap[tid] = t + 1; // 1-indexed row
      }
      
      var studentCount = 0;
      for (var i = 4; i < bData.length; i++) {
        var bRow = bData[i];
        var studentId = String(bRow[1]).trim();
        if (!studentId) continue;
        
        var targetRow = targetMap[studentId];
        if (!targetRow) {
          Logger.log('  ⚠️ ไม่พบ ' + studentId + ' ในชีตใหม่ — ข้าม');
          continue;
        }
        
        // Migrate ทั้ง 2 เทอม
        migrateTermRow_(targetSheet, targetRow, bRow, OLD.term1, NEW.term1);
        migrateTermRow_(targetSheet, targetRow, bRow, OLD.term2, NEW.term2);
        studentCount++;
      }
      
      // คำนวณเฉลี่ย 2 ภาค
      for (var sid in targetMap) {
        var row = targetMap[sid];
        var t1Total = Number(targetSheet.getRange(row, NEW.term1.total + 1).getValue()) || 0;
        var t2Total = Number(targetSheet.getRange(row, NEW.term2.total + 1).getValue()) || 0;
        var avg = 0, fg = '';
        if (t1Total > 0 && t2Total > 0) {
          avg = Math.round((t1Total + t2Total) / 2);
          fg = calculateFinalGrade(avg);
        } else if (t1Total > 0 || t2Total > 0) {
          avg = t1Total > 0 ? t1Total : t2Total;
          fg = calculateFinalGrade(avg);
        }
        targetSheet.getRange(row, 36).setValue(avg);
        targetSheet.getRange(row, 37).setValue(fg);
      }
      
      Logger.log('  ✅ Migrate ' + studentCount + ' คน สำเร็จ');
      migrated++;
    });
    
    var msg = '✅ Migrate คะแนนสำเร็จ ' + migrated + ' ชีต';
    Logger.log(msg);
    return { success: true, migrated: migrated, message: msg };
    
  } catch (error) {
    Logger.log('❌ Error in migrateScoresFromBackup: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ย้ายคะแนนเต็มของ 1 เทอม (แถว 4)
 */
function migrateTermFullScores_(targetSheet, bFullRow, oldCols, newCols) {
  // เดิม: ช่อง 1-4 → ครั้งที่ 1-4
  for (var k = 0; k < 4; k++) {
    var val = Number(bFullRow[oldCols.scores + k]) || 0;
    if (val > 0) targetSheet.getRange(4, newCols.s1 + k + 1).setValue(val);
  }
  // เดิม: ช่อง 5 → ครั้งที่ 5
  var v5 = Number(bFullRow[oldCols.scores + 4]) || 0;
  if (v5 > 0) targetSheet.getRange(4, newCols.s5 + 1).setValue(v5);
  // เดิม: ช่อง 6-8 → ครั้งที่ 6-8
  for (var k = 0; k < 3; k++) {
    var val = Number(bFullRow[oldCols.scores + 5 + k]) || 0;
    if (val > 0) targetSheet.getRange(4, newCols.s6 + k + 1).setValue(val);
  }
  // เดิม: ปลายภาค → ครั้งที่ 10
  var vFinal = Number(bFullRow[oldCols.final]) || 0;
  if (vFinal > 0) targetSheet.getRange(4, newCols.s10 + 1).setValue(vFinal);
}

/**
 * ย้ายคะแนนนักเรียน 1 แถว ของ 1 เทอม
 */
function migrateTermRow_(targetSheet, targetRow, bRow, oldCols, newCols) {
  // เดิม: ช่อง 1-4 → ครั้งที่ 1-4
  var s = [0,0,0,0, 0, 0,0,0,0, 0]; // scores index 0-9 (10 ช่อง)
  for (var k = 0; k < 4; k++) {
    s[k] = Number(bRow[oldCols.scores + k]) || 0;
    targetSheet.getRange(targetRow, newCols.s1 + k + 1).setValue(s[k]);
  }
  // เดิม: ช่อง 5 → ครั้งที่ 5 (กลางภาค)
  s[4] = Number(bRow[oldCols.scores + 4]) || 0;
  targetSheet.getRange(targetRow, newCols.s5 + 1).setValue(s[4]);
  // เดิม: ช่อง 6-8 → ครั้งที่ 6-8 (ครั้งที่ 9 ไม่มีในข้อมูลเดิม)
  for (var k = 0; k < 3; k++) {
    s[5 + k] = Number(bRow[oldCols.scores + 5 + k]) || 0;
    targetSheet.getRange(targetRow, newCols.s6 + k + 1).setValue(s[5 + k]);
  }
  // เดิม: ปลายภาค → ครั้งที่ 10
  s[9] = Number(bRow[oldCols.final]) || 0;
  targetSheet.getRange(targetRow, newCols.s10 + 1).setValue(s[9]);
  
  // คำนวณรวม
  var sum14 = s[0] + s[1] + s[2] + s[3];
  var sum69 = s[5] + s[6] + s[7] + s[8]; // s[8]=ครั้ง9 (0 จากข้อมูลเดิม)
  var midTotal = sum14 + s[4] + sum69;
  var total = midTotal + s[9];
  
  targetSheet.getRange(targetRow, newCols.sum14 + 1).setValue(sum14);
  targetSheet.getRange(targetRow, newCols.sum69 + 1).setValue(sum69);
  targetSheet.getRange(targetRow, newCols.midTotal + 1).setValue(midTotal);
  targetSheet.getRange(targetRow, newCols.total + 1).setValue(total);
  targetSheet.getRange(targetRow, newCols.grade + 1).setValue(calculateFinalGrade(total));
}

