// ============================================================
// 👩‍🏫 TEACHER MANAGEMENT (Simple Version)
// ============================================================

/**
 * [ฟังก์ชันใหม่แบบง่าย] บันทึกข้อมูลครูประจำชั้น 1 รายการ
 * @param {Object} teacherInfo - อ็อบเจกต์ที่มีข้อมูล grade, classNo, teacherName
 * @returns {string} ข้อความสถานะการบันทึก
 */
function saveSingleTeacher(teacherInfo) {
  try {
    const grade = teacherInfo.grade;
    const classNo = teacherInfo.classNo;
    const teacherName = teacherInfo.teacherName;
    const teacher2Name = teacherInfo.teacher2Name || '';

    // ตรวจสอบข้อมูลเบื้องต้น
    if (!grade || !classNo || !teacherName) {
      throw new Error("กรุณากรอกข้อมูลให้ครบ (ระดับชั้น, ห้อง, ชื่อครูคนที่ 1)");
    }

    const ss = SS();
    const sheet = ss.getSheetByName("HomeroomTeachers");
    if (!sheet) {
      throw new Error("ไม่พบชีต 'HomeroomTeachers'");
    }

    // ตรวจสอบ/เพิ่ม header คอลัมน์ที่ 4 ถ้ายังไม่มี
    if (sheet.getLastColumn() < 4) {
      sheet.getRange(1, 4).setValue('ครูประจำชั้น 2');
    }

    // --- ส่วนตรรกะการบันทึก ---
    // 1. อ่านข้อมูลทั้งหมดที่มีอยู่
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงหัวตารางออก
    
    let recordUpdated = false;
    // 2. ตรวจสอบว่ามีข้อมูลของชั้น/ห้องนี้อยู่แล้วหรือไม่
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == grade && data[i][1] == classNo) {
        // ถ้ามี ให้อัปเดตชื่อครูทั้ง 2 คน
        sheet.getRange(i + 2, 3).setValue(teacherName);
        sheet.getRange(i + 2, 4).setValue(teacher2Name);
        recordUpdated = true;
        break;
      }
    }

    // 3. ถ้ายังไม่มีข้อมูลของชั้น/ห้องนี้ ให้เพิ่มแถวใหม่
    if (!recordUpdated) {
      sheet.appendRow([grade, classNo, teacherName, teacher2Name]);
    }
    
    var msg = `✅ บันทึกข้อมูลสำเร็จ: ชั้น ${grade} ห้อง ${classNo} ครู ${teacherName}`;
    if (teacher2Name) msg += `, ${teacher2Name}`;
    return msg;

  } catch (e) {
    Logger.log("Error in saveSingleTeacher: " + e.message);
    throw new Error("เกิดข้อผิดพลาดในการบันทึก: " + e.message);
  }
}
/**
 * [ฟังก์ชันใหม่] ดึงข้อมูลครูประจำชั้นทั้งหมดมาแสดงในตารางแก้ไข
 * โดยจะรวมข้อมูลจากชีต Students และ HomeroomTeachers เข้าด้วยกัน
 */
function getTeacherListForEditor() {
  const ss = SS();
  
  // 1. ดึงข้อมูลครูที่บันทึกไว้ใน HomeroomTeachers
  const teacherSheet = ss.getSheetByName("HomeroomTeachers");
  const teacherData = teacherSheet ? teacherSheet.getDataRange().getValues().slice(1) : [];
  const teacherMap = new Map();
  teacherData.forEach(function(row) {
    var g = String(row[0] || '').trim();
    var c = String(row[1] || '').trim();
    if (!g || !c) return;
    const key = g + '-' + c;
    teacherMap.set(key, {teacher1: String(row[2] || ''), teacher2: String(row[3] || '')});
  });

  // 2. ดึงโครงสร้างชั้นเรียนจาก Students (ถ้ามี)
  const classMap = new Map();
  var studentSheet = ss.getSheetByName("Students");
  if (studentSheet && studentSheet.getLastRow() > 1) {
    var allStudentData = studentSheet.getDataRange().getValues();
    var sHeaders = allStudentData[0];
    var gradeIdx = sHeaders.indexOf('grade');
    var classIdx = sHeaders.indexOf('class_no');
    if (gradeIdx < 0) { for (var h = 0; h < sHeaders.length; h++) { if (/grade|ชั้น/i.test(String(sHeaders[h]))) { gradeIdx = h; break; } } }
    if (classIdx < 0) { for (var h = 0; h < sHeaders.length; h++) { if (/class_no|class|ห้อง/i.test(String(sHeaders[h]))) { classIdx = h; break; } } }
    if (gradeIdx < 0) gradeIdx = 5;
    if (classIdx < 0) classIdx = 6;
    allStudentData.slice(1).forEach(function(row) {
      var grade = String(row[gradeIdx] || "").trim();
      var classNo = String(row[classIdx] || "").trim();
      if (grade && classNo) {
        if (!classMap.has(grade)) classMap.set(grade, new Set());
        classMap.get(grade).add(classNo);
      }
    });
  }

  // 3. รวมชั้นเรียนจาก HomeroomTeachers เข้ามาด้วย (กรณี Students ว่าง หรือมีชั้นเรียนที่เพิ่มเองใน HomeroomTeachers)
  teacherMap.forEach(function(_val, key) {
    var parts = key.split('-');
    var grade = parts[0];
    var classNo = parts.slice(1).join('-');
    if (!classMap.has(grade)) classMap.set(grade, new Set());
    classMap.get(grade).add(classNo);
  });

  // 4. ประกอบร่างข้อมูลทั้งหมด
  const result = [];
  const sortedGrades = Array.from(classMap.keys()).sort();

  sortedGrades.forEach(function(grade) {
    const sortedClasses = Array.from(classMap.get(grade)).sort();
    sortedClasses.forEach(function(classNo) {
      const key = grade + '-' + classNo;
      var t = teacherMap.get(key) || {teacher1: '', teacher2: ''};
      result.push({
        grade: grade,
        classNo: classNo,
        teacherName: t.teacher1 || '',
        teacher2Name: t.teacher2 || ''
      });
    });
  });

  return result;
}