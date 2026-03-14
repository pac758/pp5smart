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

    // ตรวจสอบข้อมูลเบื้องต้น
    if (!grade || !classNo || !teacherName) {
      throw new Error("กรุณากรอกข้อมูลให้ครบทั้ง 3 ช่อง");
    }

    const ss = SS();
    const sheet = ss.getSheetByName("HomeroomTeachers");
    if (!sheet) {
      throw new Error("ไม่พบชีต 'HomeroomTeachers'");
    }

    // --- ส่วนตรรกะการบันทึก ---
    // 1. อ่านข้อมูลทั้งหมดที่มีอยู่
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // ดึงหัวตารางออก
    
    let recordUpdated = false;
    // 2. ตรวจสอบว่ามีข้อมูลของชั้น/ห้องนี้อยู่แล้วหรือไม่
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == grade && data[i][1] == classNo) {
        // ถ้ามี ให้อัปเดตชื่อครู
        sheet.getRange(i + 2, 3).setValue(teacherName); // +2 เพราะเราดึง header ออกและแถวเริ่มที่ 1
        recordUpdated = true;
        break;
      }
    }

    // 3. ถ้ายังไม่มีข้อมูลของชั้น/ห้องนี้ ให้เพิ่มแถวใหม่
    if (!recordUpdated) {
      sheet.appendRow([grade, classNo, teacherName]);
    }
    
    return `✅ บันทึกข้อมูลสำเร็จ: ชั้น ${grade} ห้อง ${classNo} คือครู ${teacherName}`;

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
  
  // 1. ดึงข้อมูลครูที่บันทึกไว้
  const teacherSheet = ss.getSheetByName("HomeroomTeachers");
  const teacherData = teacherSheet ? teacherSheet.getDataRange().getValues().slice(1) : [];
  const teacherMap = new Map();
  teacherData.forEach(row => {
    const key = `${row[0]}-${row[1]}`; // "Grade-ClassNo"
    teacherMap.set(key, row[2]);
  });

  // 2. ดึงโครงสร้างชั้นเรียนและห้องทั้งหมด
  const studentSheet = ss.getSheetByName("Students");
  if (!studentSheet) throw new Error("ไม่พบชีต Students");
  const allStudentData = studentSheet.getDataRange().getValues();
  const sHeaders = allStudentData[0];
  var gradeIdx = sHeaders.indexOf('grade');
  var classIdx = sHeaders.indexOf('class_no');
  if (gradeIdx < 0) { for (var h = 0; h < sHeaders.length; h++) { if (/grade|ชั้น/i.test(String(sHeaders[h]))) { gradeIdx = h; break; } } }
  if (classIdx < 0) { for (var h = 0; h < sHeaders.length; h++) { if (/class_no|class|ห้อง/i.test(String(sHeaders[h]))) { classIdx = h; break; } } }
  if (gradeIdx < 0) gradeIdx = 5;
  if (classIdx < 0) classIdx = 6;
  const studentData = allStudentData.slice(1);
  const classMap = new Map();
  studentData.forEach(row => {
    const grade = String(row[gradeIdx] || "").trim();
    const classNo = String(row[classIdx] || "").trim();
    if (grade && classNo) {
      if (!classMap.has(grade)) {
        classMap.set(grade, new Set());
      }
      classMap.get(grade).add(classNo);
    }
  });

  // 3. ประกอบร่างข้อมูลทั้งหมด
  const result = [];
  const sortedGrades = Array.from(classMap.keys()).sort();

  sortedGrades.forEach(grade => {
    const sortedClasses = Array.from(classMap.get(grade)).sort();
    sortedClasses.forEach(classNo => {
      const key = `${grade}-${classNo}`;
      result.push({
        grade: grade,
        classNo: classNo,
        teacherName: teacherMap.get(key) || "" // ถ้าไม่มีให้เป็นค่าว่าง
      });
    });
  });

  return result;
}