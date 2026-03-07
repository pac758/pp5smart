// ============================================================
// 🧑‍🎓 STUDENTS.GS - จัดการข้อมุลนักเรียน (แก้ไข Syntax Error)
// ============================================================

/**
 * ✅ ดึงรายการห้องเรียนที่มีนักเรียนในระดับชั้นที่เลือก (Fixed Version)
 * @param {string} grade - ระดับชั้น เช่น "ป.1", "ป.2"
 * @returns {Array<string>} อาร์เรย์ของเลขห้อง
 */
function getAvailableClasses_DEBUG(grade) {
  try {
    Logger.log("=== 🔍 getAvailableClasses DEBUG ===");
    Logger.log(`📥 Input grade: "${grade}"`);
    
    // 1️⃣ เชื่อมต่อ Spreadsheet
    const ss = typeof SPREADSHEET_ID !== 'undefined' 
      ? SS()
      : SS();
    
    Logger.log(`✅ Spreadsheet: ${ss.getName()}`);
    
    // 2️⃣ หาชีต Students
    const sheet = ss.getSheetByName("Students");
    if (!sheet) {
      Logger.log("❌ ไม่พบชีต 'Students'");
      const allSheets = ss.getSheets().map(s => s.getName());
      Logger.log(`📋 ชีตที่มี: ${JSON.stringify(allSheets)}`);
      return [];
    }
    
    Logger.log(`✅ พบชีต Students`);
    
    // 3️⃣ ดึงข้อมูล
    const data = sheet.getDataRange().getValues();
    Logger.log(`📊 จำนวนแถว: ${data.length}`);
    
    if (data.length <= 1) {
      Logger.log("❌ ไม่มีข้อมูล (มีแค่ header)");
      return [];
    }
    
    // 4️⃣ หา Header Index (รองรับหลายชื่อ)
    const headers = data[0];
    Logger.log(`📋 Headers: ${JSON.stringify(headers)}`);
    
    const gradeIndex = headers.findIndex(h => 
      ['grade', 'ชั้น', 'ระดับชั้น', 'level'].includes(String(h).toLowerCase())
    );
    
    const classIndex = headers.findIndex(h => 
      ['class', 'class_no', 'classroom', 'ห้อง', 'ห้องเรียน'].includes(String(h).toLowerCase())
    );
    
    Logger.log(`🔢 gradeIndex: ${gradeIndex}, classIndex: ${classIndex}`);
    
    if (gradeIndex === -1) {
      Logger.log("❌ ไม่พบคอลัมน์ระดับชั้น");
      return [];
    }
    
    if (classIndex === -1) {
      Logger.log("❌ ไม่พบคอลัมน์ห้องเรียน");
      return [];
    }
    
    // 5️⃣ ดึงข้อมูลห้องเรียน
    const classes = new Set();
    const gradeNormalized = grade.trim().toLowerCase();
    
    Logger.log(`🔍 กำลังหา grade: "${gradeNormalized}"`);
    
    for (let i = 1; i < data.length; i++) {
      const rowGrade = String(data[i][gradeIndex] || '').trim().toLowerCase();
      const rowClass = String(data[i][classIndex] || '').trim();
      
      // ตัวอย่างแถวแรก (เพื่อ debug)
      if (i === 1) {
        Logger.log(`📝 ตัวอย่างแถวที่ 1: grade="${rowGrade}", class="${rowClass}"`);
      }
      
      // เช็คแบบหลวม (รองรับทั้ง "ป.3" และ "ประถมศึกษาปีที่ 3")
      const isMatch = 
        rowGrade === gradeNormalized ||
        rowGrade.includes(gradeNormalized) ||
        gradeNormalized.includes(rowGrade);
      
      if (isMatch && rowClass) {
        classes.add(rowClass);
        Logger.log(`✅ เจอ! grade="${rowGrade}" → class="${rowClass}"`);
      }
    }
    
    // 6️⃣ เรียงลำดับและส่งกลับ
    if (classes.size === 0) {
      Logger.log("❌ ไม่พบห้องเรียนที่ตรงกับเงื่อนไข");
      return [];
    }
    
    const result = Array.from(classes).sort((a, b) => {
      const numA = parseInt(a) || 0;
      const numB = parseInt(b) || 0;
      return numA - numB;
    });
    
    Logger.log(`✅ ผลลัพธ์: ${JSON.stringify(result)}`);
    return result;
    
  } catch (e) {
    Logger.log("❌ ERROR ===");
    Logger.log(`Message: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return [];
  }
}

function saveStudentData(data) {
  try {
    data = data || {};
    const studentData = {
      student_id: String(data.studentCode || '').trim(),
      id_card: String(data.nationalId || '').trim(),
      title: String(data.prefix || '').trim(),
      firstname: String(data.firstName || '').trim(),
      lastname: String(data.lastName || '').trim(),
      grade: String(data.grade || '').trim(),
      class_no: String(data.classroom || '').trim(),
      gender: String(data.gender || '').trim(),
      birthdate: String(data.birthdate || '').trim(),
      weight: String(data.weight || '').trim(),
      height: String(data.height || '').trim(),
      photo_url: String(data.photoUrl || '').trim(),
      address: String(data.address || '').trim(),
      father_name: String(data.fatherName || '').trim(),
      father_lastname: String(data.fatherLastname || '').trim(),
      mother_name: String(data.motherName || '').trim(),
      mother_lastname: String(data.motherLastname || '').trim(),
      academic_year: String(data.academicYear || '').trim()
    };

    if (!studentData.student_id || !studentData.firstname || !studentData.lastname) {
      throw new Error('กรุณากรอกข้อมูลที่จำเป็น: รหัสนักเรียน, ชื่อ, นามสกุล');
    }

    const result = addStudent(studentData);
    if (result && result.success) {
      try {
        // เติมข้อมูลเพิ่มเติมลงคอลัมน์ที่มีอยู่ (ถ้าชีตมี header รองรับ)
        const ss = SS();
        const sheet = ss.getSheetByName('Students');
        if (sheet) {
          const values = sheet.getDataRange().getValues();
          const headers = values[0] || [];
          const col = {};
          headers.forEach((h, i) => {
            const key = String(h || '').trim();
            if (key) col[key] = i;
          });

          // หาแถวที่เพิ่งเพิ่ม (จาก student_id)
          const idIdx = col['student_id'];
          if (idIdx != null) {
            let rowIndex = -1;
            for (let r = values.length - 1; r >= 1; r--) {
              if (String(values[r][idIdx] || '').trim() === studentData.student_id) {
                rowIndex = r + 1;
                break;
              }
            }
            if (rowIndex > 0) {
              const updates = {
                birthdate: studentData.birthdate,
                weight: studentData.weight,
                height: studentData.height,
                photo_url: studentData.photo_url,
                address: studentData.address,
                father_name: studentData.father_name,
                father_lastname: studentData.father_lastname,
                mother_name: studentData.mother_name,
                mother_lastname: studentData.mother_lastname,
                academic_year: studentData.academic_year
              };
              Object.keys(updates).forEach((k) => {
                if (col[k] != null) {
                  sheet.getRange(rowIndex, col[k] + 1).setValue(updates[k]);
                }
              });
            }
          }
        }
      } catch (_e) {
        // ignore optional enrich
      }
      return result.message || 'เพิ่มนักเรียนเรียบร้อยแล้ว';
    }
    return (result && result.message) || 'ไม่สามารถเพิ่มนักเรียนได้';
  } catch (e) {
    throw new Error(e.message || String(e));
  }
}

function updateStudentInline(studentData) {
  try {
    studentData = studentData || {};
    const id = String(studentData.id || '').trim();
    if (!id) throw new Error('ไม่พบรหัสนักเรียน');

    const ss = SS();
    const sheet = ss.getSheetByName('Students');
    if (!sheet) throw new Error('ไม่พบชีต Students');

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) throw new Error('ไม่มีข้อมูลในชีต Students');

    const headers = values[0] || [];
    const col = {};
    headers.forEach((h, i) => {
      const key = String(h || '').trim();
      if (key) col[key] = i;
    });

    const idIdx = col['student_id'];
    if (idIdx == null) throw new Error('ไม่พบคอลัมน์ student_id');

    let rowIndex = -1;
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIdx] || '').trim() === id) {
        rowIndex = r + 1;
        break;
      }
    }
    if (rowIndex < 0) throw new Error('ไม่พบข้อมูลนักเรียน');

    const updates = {
      firstname: String(studentData.firstname || '').trim(),
      lastname: String(studentData.lastname || '').trim(),
      grade: String(studentData.grade || '').trim(),
      class_no: String(studentData.classNo || '').trim(),
      gender: String(studentData.gender || '').trim(),
      birthdate: String(studentData.birthdate || '').trim(),
      weight: String(studentData.weight || '').trim(),
      height: String(studentData.height || '').trim()
    };

    Object.keys(updates).forEach((k) => {
      if (col[k] != null && updates[k] !== '') {
        sheet.getRange(rowIndex, col[k] + 1).setValue(updates[k]);
      } else if (col[k] != null && (k === 'weight' || k === 'height')) {
        // allow clearing numeric
        sheet.getRange(rowIndex, col[k] + 1).setValue(updates[k]);
      }
    });

    return 'บันทึกข้อมูลเรียบร้อยแล้ว';
  } catch (e) {
    throw new Error(e.message || String(e));
  }
}

/**
 * ✅ ดึงรายชื่อนักเรียนในชั้นและห้องที่เลือก
 */
function getStudentsByClass(grade, classNo) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const headers = data[0];
    const result = [];
    
    const colIndexes = {
      id: headers.indexOf("student_id"),
      title: headers.indexOf("title"),
      firstname: headers.indexOf("firstname"),
      lastname: headers.indexOf("lastname"),
      grade: headers.indexOf("grade"),
      classNo: headers.indexOf("class_no"),
      gender: headers.indexOf("gender")
    };
    
    // ตรวจสอบว่าพบคอลัมน์ที่จำเป็นหรือไม่
    if (colIndexes.id === -1) {
      Logger.log("Error: ไม่พบคอลัมน์ 'student_id' ในชีต Students");
      return [];
    }
    
    const statusIdx = headers.indexOf("status");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowGrade = String(row[colIndexes.grade] || '').trim();
      const rowClass = String(row[colIndexes.classNo] || '').trim();

      // กรองนักเรียนที่จำหน่าย/ย้ายออก/พ้นสภาพ
      if (statusIdx !== -1) {
        const st = String(row[statusIdx] || '').trim();
        if (st === 'จำหน่าย' || st === 'ย้ายออก' || st === 'พ้นสภาพ') continue;
      }

      if (rowGrade === grade && rowClass === classNo) {
        result.push({
          id: String(row[colIndexes.id] || "").trim(),
          title: String(row[colIndexes.title] || "").trim(),
          firstname: String(row[colIndexes.firstname] || "").trim(),
          lastname: String(row[colIndexes.lastname] || "").trim(),
          grade: String(row[colIndexes.grade] || "").trim(),
          classNo: String(row[colIndexes.classNo] || "").trim(),
          gender: String(row[colIndexes.gender] || "").trim()
        });
      }
    }
    
    // --- 👇 เพิ่มโค้ดเรียงลำดับตรงนี้ ---
    result.sort((a, b) => {
      const idA = a.id || '';
      const idB = b.id || '';
      return String(idA).localeCompare(String(idB), undefined, {numeric: true});
    });
    // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---
    
    return result;
  } catch (error) {
    Logger.log(`Error in getStudentsByClass: ${error.message}`);
    return [];
  }
}

/**
 * ✅ ค้นหานักเรียนตามเงื่อนไขที่กำหนด
 */
function searchStudentBy(searchType, searchValue) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    const headers = data[0];
    const studentData = getStudentById(data, headers, searchType, searchValue);
    return studentData;

  } catch (error) {
    Logger.log(`Error in searchStudentBy: ${error.message}`);
    return null;
  }
}

/**
 * ✅ ดึงข้อมูลนักเรียน 1 คนตาม ID (ใช้ภายใน)
 */
function getStudentById(data, headers, searchType, searchValue) {
    const colIndexes = {};
    headers.forEach((h, i) => colIndexes[h] = i);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let found = false;
      
      switch (searchType) {
        case 'student_id':
          found = String(row[colIndexes.student_id] || '').trim() === searchValue.trim();
          break;
        case 'name':
          const fullName = `${row[colIndexes.title] || ''}${row[colIndexes.firstname] || ''} ${row[colIndexes.lastname] || ''}`.toLowerCase();
          found = fullName.includes(searchValue.toLowerCase());
          break;
        case 'id_card':
          found = String(row[colIndexes.id_card] || '').trim() === searchValue.trim();
          break;
        default: // Allow direct ID search
          found = String(row[colIndexes.student_id] || '').trim() === String(searchType).trim();
      }
      
      if (found) {
        return {
          rowIndex: i + 1,
          id: String(row[colIndexes.student_id] || "").trim(),
          name: `${row[colIndexes.title] || ''}${row[colIndexes.firstname] || ''} ${row[colIndexes.lastname] || ''}`.trim(),
          grade: String(row[colIndexes.grade] || "").trim(),
          classNo: String(row[colIndexes.class_no] || "").trim(),
          title: String(row[colIndexes.title] || "").trim(),
          firstname: String(row[colIndexes.firstname] || "").trim(),
          lastname: String(row[colIndexes.lastname] || "").trim(),
          gender: String(row[colIndexes.gender] || "").trim(),
          idCard: String(row[colIndexes.id_card] || "").trim()
        };
      }
    }
    return null;
}

/**
 * ✅ ลบข้อมูลนักเรียน
 */
function deleteStudent(deleteData) {
  try {
    // ตรวจสอบรหัสผ่านผู้ดูแลระบบ
    const ADMIN_PASSWORD = 'admin123'; 
    if (deleteData.adminPassword !== ADMIN_PASSWORD) {
      return { success: false, message: 'รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง' };
    }
    
    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    if (!sheet) return { success: false, message: 'ไม่พบชีต Students' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const student = getStudentById(data, headers, 'student_id', deleteData.studentId);

    if (!student) {
      return { success: false, message: 'ไม่พบข้อมูลนักเรียนที่ต้องการลบ' };
    }
    
    // บันทึก log การลบ (สามารถเพิ่มได้ถ้าต้องการ)
    Logger.log(`Deleting student: ${student.name} (ID: ${student.id})`);
    
    sheet.deleteRow(student.rowIndex);
    
    return { success: true, message: 'ดำเนินการลบข้อมูลเรียบร้อยแล้ว' };
    
  } catch (error) {
    Logger.log(`Error in deleteStudent: ${error.message}`);
    return { success: false, message: 'เกิดข้อผิดพลาดในการลบข้อมูล: ' + error.message };
  }
}

/**
 * ✅ เพิ่มนักเรียนใหม่
 */
function addStudent(studentData) {
  try {
    const ss = SS();
    let sheet = ss.getSheetByName("Students");
    
    // สร้างชีต Students ถ้ายังไม่มี
    if (!sheet) {
      sheet = createStudentsSheet(ss);
    }
    
    // ตรวจสอบข้อมูลที่จำเป็น
    if (!studentData.student_id || !studentData.firstname || !studentData.lastname) {
      return { success: false, message: 'กรุณากรอกข้อมูลที่จำเป็น: รหัส, ชื่อ, นามสกุล' };
    }
    
    // ตรวจสอบรหัสซ้ำ
    const existingData = sheet.getDataRange().getValues();
    const headers = existingData[0];
    const idIndex = headers.indexOf("student_id");
    
    if (idIndex !== -1) {
      for (let i = 1; i < existingData.length; i++) {
        if (String(existingData[i][idIndex]).trim() === String(studentData.student_id).trim()) {
          return { success: false, message: 'รหัสนักเรียนนี้มีอยู่แล้ว' };
        }
      }
    }
    
    // เพิ่มข้อมูล
    const newRow = [
      studentData.student_id || '',
      studentData.id_card || '',
      studentData.title || '',
      studentData.firstname || '',
      studentData.lastname || '',
      studentData.grade || '',
      studentData.class_no || '',
      studentData.gender || '',
      new Date(), // created_date
      'active'    // status
    ];
    
    sheet.appendRow(newRow);
    
    Logger.log(`Added student: ${studentData.firstname} ${studentData.lastname} (ID: ${studentData.student_id})`);
    return { success: true, message: 'เพิ่มนักเรียนเรียบร้อยแล้ว' };
    
  } catch (error) {
    Logger.log('Error in addStudent: ' + error.message);
    return { success: false, message: 'เกิดข้อผิดพลาดในการเพิ่มนักเรียน: ' + error.message };
  }
}

/**
 * ✅ สร้างชีต Students พร้อมโครงสร้างเริ่มต้น
 */
function createStudentsSheet(ss) {
  try {
    const sheet = ss.insertSheet("Students");
    
    // สร้าง header
    const headers = [
      'student_id', 'id_card', 'title', 'firstname', 'lastname', 
      'grade', 'class_no', 'gender', 'created_date', 'status'
    ];
    
    // สร้าง header ภาษาไทย (แถวที่ 2 สำหรับอ้างอิง)
    const thaiHeaders = [
      'รหัสนักเรียน', 'เลขบัตรประชาชน', 'คำนำหน้า', 'ชื่อ', 'นามสกุล',
      'ชั้น', 'ห้อง', 'เพศ', 'วันที่สร้าง', 'สถานะ'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // จัดรูปแบบ header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // กำหนดความกว้างคอลัมน์
    sheet.setColumnWidth(1, 120); // student_id
    sheet.setColumnWidth(2, 150); // id_card
    sheet.setColumnWidth(3, 80);  // title
    sheet.setColumnWidth(4, 150); // firstname
    sheet.setColumnWidth(5, 150); // lastname
    sheet.setColumnWidth(6, 80);  // grade
    sheet.setColumnWidth(7, 60);  // class_no
    sheet.setColumnWidth(8, 80);  // gender
    sheet.setColumnWidth(9, 120); // created_date
    sheet.setColumnWidth(10, 100); // status
    
    // เพิ่มข้อมูลตัวอย่าง
    const sampleData = [
      ['2024001', '1234567890123', 'เด็กชาย', 'สมชาย', 'ใจดี', 'ป.1', '1', 'ชาย', new Date(), 'active'],
      ['2024002', '1234567890124', 'เด็กหญิง', 'สมหญิง', 'ใจงาม', 'ป.1', '1', 'หญิง', new Date(), 'active'],
      ['2024003', '1234567890125', 'เด็กชาย', 'สมศักดิ์', 'ใจสู้', 'ป.1', '2', 'ชาย', new Date(), 'active']
    ];
    
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
    
    Logger.log('Created Students sheet with sample data');
    return sheet;
    
  } catch (error) {
    Logger.log('Error creating Students sheet: ' + error.message);
    throw error;
  }
}

/**
 * ✅ ฟังก์ชันเสริมสำหรับแปลงวันที่สำหรับ input[type="date"]
 */

/**
 * ✅ ฟังก์ชันหลักสำหรับนำเข้าไฟล์
 */
function importStudentFile(fileType, csvContent, fileName) {
  if (fileType !== 'csv') {
    throw new Error('อนุญาตให้นำเข้าเฉพาะไฟล์ CSV เท่านั้น');
  }
  return importCsvStudents(csvContent);
}

/**
 * ✅ ประมวลผลและนำเข้าข้อมูลจาก CSV
 */
function importCsvStudents(csvContent) {
  try {
    const rows = Utilities.parseCsv(csvContent);
    if (!rows || rows.length <= 2) {
      throw new Error('ไม่พบข้อมูลนักเรียนในแถวที่ 3 เป็นต้นไป');
    }

    const dataToImport = rows.slice(2); // ข้อมูลเริ่มจากแถวที่ 3
    
    const ss = SS();
    let sheet = ss.getSheetByName("Students");
    if (!sheet) {
      sheet = createStudentsSheet(ss);
    }

    const existingData = sheet.getDataRange().getValues();
    const studentMap = new Map();
    
    // สร้าง Map จากข้อมูลเดิมเพื่อตรวจสอบข้อมูลซ้ำ
    for (let i = 1; i < existingData.length; i++) {
      const studentId = String(existingData[i][0]).trim(); // student_id
      if (studentId) {
        studentMap.set(studentId, i + 1); // Key: studentId, Value: rowIndex
      }
    }

    let updatedCount = 0;
    let insertedCount = 0;

    dataToImport.forEach(row => {
      const studentId = String(row[5] || '').trim(); // รหัสนักเรียนจากคอลัมน์ F
      if (!studentId) return; // ข้ามแถวที่ไม่มีรหัสนักเรียน

      const newRowData = [
        studentId,                                    // student_id
        String(row[2] || '').trim(),                 // id_card (เลข ปชช.)
        String(row[7] || '').trim(),                 // title (คำนำหน้า)
        String(row[8] || '').trim(),                 // firstname (ชื่อ)
        String(row[9] || '').trim(),                 // lastname (นามสกุล)
        String(row[3] || '').trim(),                 // grade (ระดับชั้น)
        String(row[4] || '').trim(),                 // class_no (ห้อง)
        (String(row[6] || '').trim() === 'ช' ? 'ชาย' : 'หญิง'), // gender (เพศ)
        new Date(),                                  // created_date
        'active'                                     // status
      ];

      const rowIndex = studentMap.get(studentId);
      if (rowIndex) {
        // อัปเดตข้อมูลเดิม
        sheet.getRange(rowIndex, 1, 1, newRowData.length).setValues([newRowData]);
        updatedCount++;
      } else {
        // เพิ่มข้อมูลใหม่
        sheet.appendRow(newRowData);
        insertedCount++;
      }
    });

    return `✅ นำเข้าเสร็จสิ้น: เพิ่มใหม่ ${insertedCount} รายการ, อัปเดต ${updatedCount} รายการ`;
  } catch (e) {
    Logger.log("Error in importCsvStudents: " + e.message);
    throw new Error("เกิดข้อผิดพลาดในการอ่านไฟล์ CSV: " + e.message);
  }
}

/**
 * ✅ ดึงสถิติจำนวนนักเรียน
 */
function getStudentStats() {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    
    if (!sheet) {
      return { total: 0, byGrade: {}, byClass: {} };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { total: 0, byGrade: {}, byClass: {} };
    
    const headers = data[0];
    const gradeIndex = headers.indexOf("grade");
    const classIndex = headers.indexOf("class_no");
    
    const students = data.slice(1).filter(row => row[0]); // กรองแถวที่มี student_id
    
    const stats = {
      total: students.length,
      byGrade: {},
      byClass: {}
    };
    
    students.forEach(row => {
      const grade = String(row[gradeIndex] || '').trim();
      const classNo = String(row[classIndex] || '').trim();
      
      // นับตามชั้น
      if (grade) {
        stats.byGrade[grade] = (stats.byGrade[grade] || 0) + 1;
      }
      
      // นับตามห้อง
      const gradeClass = `${grade}/${classNo}`;
      if (grade && classNo) {
        stats.byClass[gradeClass] = (stats.byClass[gradeClass] || 0) + 1;
      }
    });
    
    return stats;
    
  } catch (error) {
    Logger.log('Error in getStudentStats: ' + error.message);
    return { total: 0, byGrade: {}, byClass: {} };
  }
}


// ในไฟล์ Code.gs (นำไปวางทับของเดิม)

function getFilteredStudentsInline(grade, classNo) {
  const sheet = SS().getSheetByName("Students");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const result = [];

  var U_formatToDateInput = (typeof globalThis !== 'undefined' && typeof globalThis.U_formatToDateInput === 'function')
    ? globalThis.U_formatToDateInput
    : function(value) {
        if (!value) return '';
        if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
          return Utilities.formatDate(value, 'Asia/Bangkok', 'yyyy-MM-dd');
        }
        const s = String(value).trim();
        const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (m) return `${m[1]}-${String(m[2]).padStart(2, '0')}-${String(m[3]).padStart(2, '0')}`;
        return s;
      };

  // สร้าง mapping ของ header ไปยัง index เพื่อให้อนาคตแก้ไขง่าย
  const colMap = {};
  headers.forEach((h, i) => colMap[h.trim()] = i);

  // กำหนด index ของคอลัมน์ตามที่คุณระบุ (O=14, P=15, Q=16, R=17, S=18, T=19)
  const fatherFirstNameIndex = 14; // O
  const fatherLastNameIndex  = 15; // P
  const motherFirstNameIndex = 16; // Q
  const motherLastNameIndex  = 17; // R
  const addressIndex         = 18; // S
  const occupationIndex      = 19; // T (ใช้สำหรับทั้งบิดาและมารดา)

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowGrade = String(row[colMap['grade']] || '').trim();
    const rowClass = String(row[colMap['class_no']] || '').trim();

    // กรองนักเรียนที่จำหน่าย/ย้ายออก/พ้นสภาพ
    var rowStatus = colMap['status'] != null ? String(row[colMap['status']] || '').trim() : '';
    if (rowStatus === 'จำหน่าย' || rowStatus === 'ย้ายออก' || rowStatus === 'พ้นสภาพ') continue;

    const matchGrade = !grade || grade === rowGrade;
    const matchClass = !classNo || classNo === rowClass;

    if (matchGrade && matchClass) {
      const father_name = `${String(row[fatherFirstNameIndex] || '').trim()} ${String(row[fatherLastNameIndex] || '').trim()}`.trim();
      const mother_name = `${String(row[motherFirstNameIndex] || '').trim()} ${String(row[motherLastNameIndex] || '').trim()}`.trim();
      result.push({
        id: String(row[colMap['student_id']] || '').trim(),
        idCard: String(row[colMap['id_card']] || '').trim(),
        title: String(row[colMap['title']] || '').trim(),
        firstname: String(row[colMap['firstname']] || '').trim(),
        lastname: String(row[colMap['lastname']] || '').trim(),
        grade: rowGrade,
        classNo: rowClass,
        gender: String(row[colMap['gender']] || '').trim(),
        birthdate: U_formatToDateInput(row[colMap['birthdate']]),
        father_name: father_name,
        father_occupation: String(row[occupationIndex] || '').trim(),
        mother_name: mother_name,
        mother_occupation: String(row[occupationIndex] || '').trim(),
        address: String(row[addressIndex] || '').trim()
      });
    }
  }

  // --- 👇 เพิ่มโค้ดเรียงลำดับนักเรียนตามรหัส (id) ---
  result.sort((a, b) => {
      const idA = a.id || '';
      const idB = b.id || '';
      return String(idA).localeCompare(String(idB), undefined, { numeric: true });
  });
  // --- 👆 สิ้นสุดโค้ดที่เพิ่ม ---

  return result;
}

// ============================================================
// 📋 STUDENT LIST PDF EXPORT
// ============================================================

// exportStudentListPDF() → ย้ายไปด้านล่าง (ป้องกันประกาศซ้ำ)

/**
 * 🆕 [อัปเดตล่าสุด] สร้าง HTML สำหรับ PDF ทะเบียนประวัติ (แบบ 2 หน้า)
 */
function createStudentRegistryHTML_Simple(students, grade) {
  // ฟังก์ชันแปลงวันที่ (Helper function)
  function formatThaiDateForRegistry(dateStr) {
    if (!dateStr) return { d: '', m: '', y: '' };
    try {
      const date = new Date(dateStr);
      const d = date.getDate();
      const m = date.getMonth() + 1;
      const y = date.getFullYear() + 543;
      return { d, m, y };
    } catch {
      return { d: '', m: '', y: '' };
    }
  }

  // สร้างแถวข้อมูลนักเรียนสำหรับตารางทั้งสอง
  let studentRows1 = ''; // แถวสำหรับตารางหน้า 1
  let studentRows2 = ''; // แถวสำหรับตารางหน้า 2
  const maxRows = 35; // จำนวนแถวสูงสุดต่อหน้า

  for (let i = 0; i < maxRows; i++) {
    const s = students[i]; // ดึงข้อมูลนักเรียน, ถ้าไม่มีจะได้ undefined
    const dob = s ? formatThaiDateForRegistry(s.birthdate) : { d: '', m: '', y: '' };
    
    // สร้างแถวสำหรับตารางที่ 1
    studentRows1 += `
      <tr>
        <td class="center">${i + 1}</td>
        <td>${s ? s.id : ''}</td>
        <td>${s ? s.idCard : ''}</td>
        <td class="left">${s ? `${s.firstname} ${s.lastname}` : ''}</td>
        <td class="center">${dob.d}</td>
        <td class="center">${dob.m}</td>
        <td class="center">${dob.y}</td>
        <td class="center">${s ? s.age || '' : ''}</td>
        <td class="center">${s ? s.blood_type || '' : ''}</td>
      </tr>
    `;
    
    // สร้างแถวสำหรับตารางที่ 2
    studentRows2 += `
       <tr>
        <td class="center">${i + 1}</td>
        <td class="left">${s ? s.father_name : ''}</td>
        <td class="left">${s ? s.father_occupation : ''}</td>
        <td class="left">${s ? s.mother_name : ''}</td>
        <td class="left">${s ? s.mother_occupation : ''}</td>
        <td class="left address">${s ? s.address : ''}</td>
      </tr>
    `;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
        
        /* ✨ [แก้ไข] ตั้งค่าหน้ากระดาษเป็นแนวตั้ง ✨ */
        @page { 
          size: A4 portrait; 
          margin: 1.5cm; 
        }

        body { font-family: 'Sarabun', sans-serif; font-size: 11px; }
        
        /* ✨ [ใหม่] สไตล์สำหรับบังคับขึ้นหน้าใหม่ ✨ */
        .page-break {
          page-break-before: always;
        }

        .table-container { width: 100%; }
        h2 { font-size: 16px; font-weight: 700; text-align: center; margin-bottom: 15px; }
        table { width: 100%; border-collapse: collapse; border: 2px solid black; }
        th, td { border: 1px solid black; padding: 4px; vertical-align: middle; }
        th { font-weight: 700; text-align: center; background-color: #f2f2f2; }
        td { height: 18px; }
        .center { text-align: center; }
        .left { text-align: left; padding-left: 5px; }
        .address { font-size: 10px; }
        .vertical-text { writing-mode: vertical-rl; text-orientation: mixed; white-space: nowrap; }
      </style>
    </head>
    <body>
      <div class="table-container">
        <h2>ข้อมูลนักเรียน</h2>
        <table>
          <thead>
            <tr>
              <th rowspan="2" class="vertical-text">เลขที่</th>
              <th>เลข</th><th>เลขประจำตัว</th><th rowspan="2">ชื่อ - สกุล</th>
              <th colspan="3">วัน เดือน ปี เกิด</th><th rowspan="2">อายุ</th><th rowspan="2">หมู่โลหิต</th>
            </tr>
            <tr>
              <th>ประจำตัวนักเรียน</th><th>ประชาชน</th>
              <th>วัน</th><th>เดือน</th><th>พ.ศ.</th>
            </tr>
          </thead>
          <tbody>${studentRows1}</tbody>
        </table>
      </div>

      <div class="table-container page-break">
        <h2>ข้อมูลนักเรียน ชั้นประถมศึกษาปีที่ ${grade.replace('ป.','')}</h2>
        <table>
          <thead>
            <tr>
              <th rowspan="2">เลขที่</th>
              <th colspan="2">ข้อมูลบิดา</th><th colspan="2">ข้อมูลมารดา</th>
              <th rowspan="2">ที่อยู่ปัจจุบัน</th>
            </tr>
            <tr>
              <th>ชื่อ-สกุล</th><th>อาชีพ</th><th>ชื่อ-สกุล</th><th>อาชีพ</th>
            </tr>
          </thead>
          <tbody>${studentRows2}</tbody>
        </table>
      </div>
    </body>
    </html>
  `;
}
// ============================================================
// 📋 STUDENT LIST PDF EXPORT (สำหรับตารางปกติ)
// ============================================================

/**
 * ✅ ฟังก์ชันหลักสำหรับส่งออกรายชื่อนักเรียนเป็น PDF (ตารางปกติ)
 * @param {string} grade - ระดับชั้นที่ต้องการ (ถ้าว่างคือทั้งหมด)
 * @param {string} classNo - ห้องที่ต้องการ (ถ้าว่างคือทั้งหมด)
 * @returns {string} URL ของไฟล์ PDF ที่สร้างเสร็จแล้ว
 */
function exportStudentListPDF(grade, classNo) {
  try {
    const students = getFilteredStudentsInline(grade, classNo);
    if (!students || students.length === 0) {
      throw new Error("ไม่พบข้อมูลนักเรียนตามเงื่อนไขที่เลือก");
    }

    // เรียกใช้ฟังก์ชันสร้าง HTML ที่ขาดไป
    const html = createStudentListHTMLForPDF(students, grade, classNo);

    const blob = HtmlService.createHtmlOutput(html)
      .getBlob()
      .setName(`รายชื่อนักเรียน_${grade || 'ทุกชั้น'}_ห้อง${classNo || 'ทุกห้อง'}.pdf`)
      .getAs('application/pdf');

    const folderName = "AttendancePDFs";
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log(`✅ Student List PDF created: ${file.getUrl()}`);
    return file.getUrl();

  } catch (error) {
    Logger.log(`❌ Error in exportStudentListPDF: ${error.message}`);
    throw new Error(`ไม่สามารถสร้าง PDF ได้: ${error.message}`);
  }
}

/**
 * ✅ ฟังก์ชันสร้าง HTML สำหรับ PDF (ตารางปกติ)
 */
function createStudentListHTMLForPDF(students, grade, classNo) {
  function U_formatThaiDate(dateStr) {
    if (!dateStr) return "-";
    const months = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
    try {
      const [y, m, d] = dateStr.split("-");
      const thaiYear = parseInt(y) + 543;
      return `${parseInt(d)} ${months[parseInt(m)]} ${thaiYear}`;
    } catch {
      return dateStr;
    }
  }

  let studentRows = '';
  students.forEach(s => {
    studentRows += `
      <tr>
        <td>${s.id}</td>
        <td>${s.idCard}</td>
        <td class="text-start">${s.firstname} ${s.lastname}</td>
        <td>${s.grade}</td>
        <td>${s.classNo}</td>
        <td>${s.gender}</td>
        <td>${U_formatThaiDate(s.birthdate)}</td>
        <td>${s.age || "-"}</td>
        <td>${s.weight || "-"}</td>
        <td>${s.height || "-"}</td>
      </tr>
    `;
  });
  
  const filterText = `ชั้น ${grade || 'ทั้งหมด'} | ห้อง ${classNo || 'ทั้งหมด'}`;

  return `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
        body { font-family: 'Sarabun', sans-serif; font-size: 12px; }
        h1 { font-size: 20px; text-align: center; font-weight: 700; margin-bottom: 5px; }
        .info { text-align: center; font-size: 14px; color: #555; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ccc; padding: 6px; text-align: center; vertical-align: middle; }
        th { background-color: #f2f2f2; font-weight: 700; }
        .text-start { text-align: left; }
      </style>
    </head>
    <body>
      <h1>📋 รายชื่อนักเรียน</h1>
      <div class="info">${filterText}</div>
      <table>
        <thead>
          <tr>
            <th>รหัส</th>
            <th>เลข ปชช.</th>
            <th class="text-start">ชื่อ - สกุล</th>
            <th>ชั้น</th>
            <th>ห้อง</th>
            <th>เพศ</th>
            <th>วันเกิด</th>
            <th>อายุ</th>
            <th>น้ำหนัก</th>
            <th>ส่วนสูง</th>
          </tr>
        </thead>
        <tbody>
          ${studentRows}
        </tbody>
      </table>
    </body>
    </html>
  `;
}
// ============================================================
// 📋 STUDENT REGISTRY PDF EXPORT (ทะเบียนประวัติ 2 หน้า)
// ============================================================

/**
 * HELPER [ใหม่]: แปลงไฟล์รูปภาพใน Drive เป็น Base64 Data URL
 * @param {string} fileId - ID ของไฟล์รูปภาพ
 * @returns {string} Data URL ของรูปภาพสำหรับใช้ใน <img>
 */
function getLogoAsBase64_(fileId) {
  try {
    if (!fileId) return '';
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    return `data:${contentType};base64,${base64Data}`;
  } catch (e) {
    Logger.log(`ไม่สามารถแปลงโลโก้เป็น Base64 ได้: ${e.message}`);
    return ''; // คืนค่าว่างถ้าเกิดข้อผิดพลาด
  }
}

/**
 * ✅ [ฉบับแก้ไขสมบูรณ์] ส่งออก PDF รูปแบบทะเบียนประวัติ (ปพ.5)
 * - ✨ ใหม่: ปรับขนาดตัวอักษรเป็น 13px และปรับปรุง CSS ให้สมบูรณ์
 */
function exportStudentRegistryPDF(grade, classNo) {
  try {
    const settings = S_getGlobalSettings(false);
    const schoolNameResolved = [
      settings['ชื่อโรงเรียน'],
      settings.schoolName,
      settings['school_name'],
      settings['ชื่อสถานศึกษา']
    ]
      .map(v => String(v || '').trim())
      .find(v => v && v !== 'โรงเรียนตัวอย่าง') || String(settings['ชื่อโรงเรียน'] || settings.schoolName || '').trim();

    if (schoolNameResolved) {
      settings['ชื่อโรงเรียน'] = schoolNameResolved;
      settings.schoolName = schoolNameResolved;
    }

    const logoBase64 = getLogoAsBase64_(settings['logoFileId']);
    const students = getFilteredStudentsInline(grade, classNo);
    if (!students || students.length === 0) {
      throw new Error("ไม่พบข้อมูลนักเรียนตามเงื่อนไขที่เลือก");
    }

    const STUDENTS_PER_PAGE = 20;
    let finalHtmlBody = '';

    for (let i = 0; i < students.length; i += STUDENTS_PER_PAGE) {
      const studentChunk = students.slice(i, i + STUDENTS_PER_PAGE);
      const startingIndex = i;
      finalHtmlBody += createRegistryPageHTML(studentChunk, grade, startingIndex, settings, logoBase64);
      if (i + STUDENTS_PER_PAGE < students.length) {
        finalHtmlBody += '<div style="page-break-after: always;"></div>';
      }
    }

    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
          @page { size: A4 portrait; margin: 1.2cm 1cm; }
          /* ✨ [แก้ไข] เพิ่มขนาดตัวอักษรพื้นฐานเป็น 13px ✨ */
          body { font-family: 'Sarabun', sans-serif; font-size: 13px; margin: 0; padding: 0; }
          .page-container { page-break-inside: avoid; }
          .page-break-internal { page-break-before: always; }
          .header-table { width: 100%; border: none; margin-bottom: 10px; }
          .header-table td { border: none; vertical-align: middle; padding: 0; }
          .logo-cell { width: 80px; text-align: left; }
          .logo { max-width: 70px; max-height: 70px; }
          .school-info-cell { text-align: center; }
          .school-info-cell h2 { font-size: 18px; font-weight: 700; margin: 0; }
          .school-info-cell h3 { font-size: 16px; margin: 2px 0; }
          .school-info-cell p { font-size: 14px; margin: 0; }
          table { width: 100%; border-collapse: collapse; border: 2px solid black; }
          th, td { border: 1px solid black; padding: 5px; vertical-align: middle; height: 20px; line-height: 1.3; }
          th { font-weight: 700; text-align: center; background-color: #f2f2f2; }
          .center { text-align: center; }
          .left { text-align: left; padding-left: 5px; }
          .address { font-size: 12px; word-wrap: break-word; }
          .vertical-text { writing-mode: vertical-rl; text-orientation: mixed; white-space: nowrap; }
        </style>
      </head>
      <body>${finalHtmlBody}</body>
      </html>`;

    const blob = Utilities.newBlob(fullHtml, MimeType.HTML)
      .setName(`ทะเบียนประวัตินักเรียน_${grade || 'ทุกชั้น'}_ห้อง${classNo || 'ทุกห้อง'}.pdf`)
      .getAs('application/pdf');

    const folderName = "AttendancePDFs";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();

  } catch (error) {
    Logger.log(`❌ Error in exportStudentRegistryPDF: ${error.message}`);
    throw new Error(`ไม่สามารถสร้าง PDF ได้: ${error.message}`);
  }
}

/**
 * HELPER: สร้าง HTML สำหรับทะเบียนประวัติ 1 ชุด (2 หน้า)
 * - ✨ [แก้ไข] กำหนดความกว้างของคอลัมน์ (width) ใน <th> ให้สมบูรณ์
 */
function createRegistryPageHTML(studentChunk, grade, startingIndex, settings, logoBase64) {
  const schoolName = settings.schoolName || settings['ชื่อโรงเรียน'] || 'โรงเรียนไม่ระบุ';
  const academicYear = settings['ปีการศึกษา'] || (new Date().getFullYear() + 543);
  const gradeDisplay = grade ? grade.replace('ป.', '') : '';

  const headerHtmlPage1 = `...`; // (ส่วน Header นี้เหมือนเดิมครับ)
  const headerHtmlPage2 = `...`; // (ส่วน Header นี้เหมือนเดิมครับ)
  
  // ย่อส่วน Header ไว้เพื่อความกระชับ
  const fullHeaderHtmlPage1 = `
    <table class="header-table">
      <tr>
        <td class="logo-cell">${logoBase64 ? `<img src="${logoBase64}" class="logo">` : ''}</td>
        <td class="school-info-cell">
          <h2>ทะเบียนข้อมูลนักเรียน (ปพ.5)</h2>
          <h3>ชั้นประถมศึกษาปีที่ ${gradeDisplay} | ปีการศึกษา ${academicYear}</h3>
          <p style="font-size: 14px; margin-top: 4px;">${schoolName}</p>
        </td>
      </tr>
    </table>`;
  const fullHeaderHtmlPage2 = `
    <table class="header-table">
      <tr>
        <td class="logo-cell">${logoBase64 ? `<img src="${logoBase64}" class="logo">` : ''}</td>
        <td class="school-info-cell">
          <h2>ข้อมูลผู้ปกครองนักเรียน</h2>
          <h3>ชั้นประถมศึกษาปีที่ ${gradeDisplay} | ปีการศึกษา ${academicYear}</h3>
          <p style="font-size: 14px; margin-top: 4px;">${schoolName}</p>
        </td>
      </tr>
    </table>`;

  let studentRows1 = '', studentRows2 = '';
  studentChunk.forEach((s, index) => {
    // ... ส่วนสร้างข้อมูลนักเรียนตรงนี้เหมือนเดิมทุกประการ ...
    const studentNumber = startingIndex + index + 1;
    const fullName = `${s.title || ''}${s.firstname} ${s.lastname}`.trim();
    const dob = s.birthdate ? new Date(s.birthdate) : null;
    studentRows1 += `<tr><td class="center">${studentNumber}</td><td class="center">${s.id || ''}</td><td class="center">${s.idCard || ''}</td><td class="left">${fullName}</td><td class="center">${dob ? dob.getDate() : ''}</td><td class="center">${dob ? dob.getMonth() + 1 : ''}</td><td class="center">${dob ? dob.getFullYear() + 543 : ''}</td><td class="center">${s.age || ''}</td><td class="center">${s.blood_type || ''}</td></tr>`;
    studentRows2 += `<tr><td class="center">${studentNumber}</td><td class="left">${s.father_name || ''}</td><td class="center">${s.father_occupation || ''}</td><td class="left">${s.mother_name || ''}</td><td class="center">${s.mother_occupation || ''}</td><td class="left address">${s.address || ''}</td></tr>`;
  });

  return `
    <div class="page-container">
      ${fullHeaderHtmlPage1}
      <table>
        <thead>
          <tr>
            <th rowspan="2" class="vertical-text" style="width: 4%;">เลขที่</th>
            <th rowspan="2" style="width: 12%;">เลขประจำตัวนักเรียน</th>
            <th rowspan="2" style="width: 13%;">เลขประจำตัวประชาชน</th>
            <th rowspan="2" style="width: 30%;">ชื่อ - สกุล</th>
            <th colspan="3">วัน เดือน ปี เกิด</th>
            <th rowspan="2" style="width: 6%;">อายุ</th>
            <th rowspan="2" style="width: 10%;">หมู่โลหิต</th>
          </tr>
          <tr>
            <th style="width: 7%;">วัน</th>
            <th style="width: 7%;">เดือน</th>
            <th style="width: 7%;">พ.ศ.</th>
          </tr>
        </thead>
        <tbody>${studentRows1}</tbody>
      </table>
    </div>
    <div class="page-container page-break-internal">
      ${fullHeaderHtmlPage2}
      <table>
        <thead>
          <tr>
            <th rowspan="2" style="width: 5%;">เลขที่</th>
            <th colspan="2">ข้อมูลบิดา</th>
            <th colspan="2">ข้อมูลมารดา</th>
            <th rowspan="2" style="width: 35%;">ที่อยู่ปัจจุบัน</th>
          </tr>
          <tr>
            <th style="width: 15%;">ชื่อ-สกุล</th>
            <th style="width: 10%;">อาชีพ</th>
            <th style="width: 15%;">ชื่อ-สกุล</th>
            <th style="width: 10%;">อาชีพ</th>
          </tr>
        </thead>
        <tbody>${studentRows2}</tbody>
      </table>
    </div>`;
}

/**
 * 🆕 สร้าง HTML สำหรับ PDF ทะเบียนประวัติ (หน้าละ 30 คน พอดีกับหน้า A4 + มีคำนำหน้า)
 */
function createStudentRegistryHTML(studentChunk, grade, pageNumber) {
  // ฟังก์ชันแปลงวันที่ (Helper function)
  function formatThaiDateForRegistry(dateStr) {
    if (!dateStr) return { d: '', m: '', y: '' };
    try {
      const date = new Date(dateStr);
      return { d: date.getDate(), m: date.getMonth() + 1, y: date.getFullYear() + 543 };
    } catch {
      return { d: '', m: '', y: '' };
    }
  }

  let studentRows1 = ''; // แถวสำหรับตารางหน้า 1
  let studentRows2 = ''; // แถวสำหรับตารางหน้า 2
  const maxRows = 30; // หน้าละ 30 คน
  const startingIndex = (pageNumber - 1) * maxRows;

  for (let i = 0; i < maxRows; i++) {
    const s = studentChunk[i];
    const dob = s ? formatThaiDateForRegistry(s.birthdate) : { d: '', m: '', y: '' };
    const currentNumber = startingIndex + i + 1;

    // ✨ เพิ่มคำนำหน้าในชื่อ-สกุล
    const fullName = s ? `${s.title || ''}${s.firstname} ${s.lastname}` : '';

    // สร้างแถวสำหรับตารางที่ 1
    studentRows1 += `
      <tr>
        <td class="center">${s ? currentNumber : ''}</td>
        <td class="center">${s ? s.id : ''}</td>
        <td class="center">${s ? s.idCard : ''}</td>
        <td class="left">${fullName}</td>
        <td class="center">${dob.d}</td>
        <td class="center">${dob.m}</td>
        <td class="center">${dob.y}</td>
        <td class="center">${s ? s.age || '' : ''}</td>
        <td class="center">${s ? s.blood_type || '' : ''}</td>
      </tr>
    `;
    
    // สร้างแถวสำหรับตารางที่ 2
    studentRows2 += `
       <tr>
        <td class="center">${s ? currentNumber : ''}</td>
        <td class="left">${s ? s.father_name : ''}</td>
        <td class="center">${s ? s.father_occupation : ''}</td>
        <td class="left">${s ? s.mother_name : ''}</td>
        <td class="center">${s ? s.mother_occupation : ''}</td>
        <td class="left address">${s ? s.address : ''}</td>
      </tr>
    `;
  }

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
        
        @page { 
          size: A4 portrait; 
          margin: 1.2cm 1cm;
        }

        body { 
          font-family: 'Sarabun', sans-serif; 
          font-size: 14px;
          margin: 0;
          padding: 0;
        }
        
        .page-break { 
          page-break-after: always;
        }
        
        .table-container { 
          width: 100%;
          page-break-inside: avoid;
        }
        
        h2 { 
          font-size: 16px; 
          font-weight: 700; 
          text-align: center; 
          margin: 8px 0 10px 0;
        }

        table { 
          width: 100%; 
          border-collapse: collapse; 
          border: 2px solid black;
        }
        
        th, td { 
          border: 1px solid black; 
          padding: 5px 4px;
          vertical-align: middle;
          font-size: 14px;
          line-height: 1.2;
        }
        
        td { 
          height: 20px;
        }

        th { 
          font-weight: 700; 
          text-align: center; 
          background-color: #f2f2f2;
          padding: 6px 4px;
        }
        
        .center { 
          text-align: center; 
        }
        
        .left { 
          text-align: left; 
          padding-left: 6px; 
        }
        
        .address { 
          font-size: 13px;
          word-wrap: break-word;
          word-break: break-word;
        }
        
        .vertical-text { 
          writing-mode: vertical-rl; 
          text-orientation: mixed; 
          white-space: nowrap;
          padding: 6px 3px;
        }
      </style>
    </head>
    <body>
      <div class="table-container page-break">
        <h2>ข้อมูลนักเรียน ชั้นประถมศึกษาปีที่ ${grade ? grade.replace('ป.',''): ''}</h2>
        <table>
          <thead>
            <tr>
              <th rowspan="2" class="vertical-text">เลขที่</th>
              <th rowspan="2">เลขประจำตัวนักเรียน</th>
              <th rowspan="2">เลขประจำตัวประชาชน</th>
              <th rowspan="2">ชื่อ - สกุล</th>
              <th colspan="3">วัน เดือน ปี เกิด</th>
              <th rowspan="2">อายุ</th>
              <th rowspan="2">หมู่โลหิต</th>
            </tr>
            <tr>
              <th>วัน</th>
              <th>เดือน</th>
              <th>พ.ศ.</th>
            </tr>
          </thead>
          <tbody>${studentRows1}</tbody>
        </table>
      </div>
      <div class="table-container">
        <h2>ข้อมูลนักเรียน ชั้นประถมศึกษาปีที่ ${grade ? grade.replace('ป.',''): ''}</h2>
        <table>
          <thead>
            <tr>
              <th rowspan="2">เลขที่</th>
              <th colspan="2">ข้อมูลบิดา</th>
              <th colspan="2">ข้อมูลมารดา</th>
              <th rowspan="2">ที่อยู่ปัจจุบัน</th>
            </tr>
            <tr>
              <th>ชื่อ-สกุล</th>
              <th>อาชีพ</th>
              <th>ชื่อ-สกุล</th>
              <th>อาชีพ</th>
            </tr>
          </thead>
          <tbody>${studentRows2}</tbody>
        </table>
      </div>
    </body>
    </html>
  `;
}
// ============================================================
// 🏫 TEACHER COMMENT SYSTEM - BACKEND
// ============================================================



/**
 * 🆕 [ใหม่] ดึงรายชื่อนักเรียนตามชั้นและห้อง (สำหรับ dropdown)
 */
function getStudentsByGradeClass(grade, classNo) {
  try {
    const students = getFilteredStudentsInline(grade, classNo); // ใช้ฟังก์ชันเดิมที่มีอยู่
    
    // --- 👇 แก้ไข: เพิ่ม {numeric: true} ---
    if (students && Array.isArray(students)) {
      students.sort((a, b) => {
        const idA = a.id || '';
        const idB = b.id || '';
        return String(idA).localeCompare(String(idB), undefined, {numeric: true});
      });
    }
    // --- 👆 สิ้นสุดโค้ดที่แก้ไข ---

    return students.map(s => ({
      id: s.id,
      fullName: `${s.title}${s.firstname} ${s.lastname}`.trim()
    }));
  } catch (e) {
    Logger.log(`Error in getStudentsByGradeClass: ${e.message}`);
    return [];
  }
}


/**
 * 🆕 [ใหม่] ดึงข้อมูลนักเรียนพร้อมความคิดเห็นเดิม
 */
function getStudentFullData(studentId) {
  try {
    const students = getFilteredStudentsInline('', ''); // ดึงนักเรียนทั้งหมด
    const student = students.find(s => String(s.id).trim() === String(studentId).trim());

    if (!student) {
      return { success: false, error: 'ไม่พบนักเรียน' };
    }
    
    // อ่านความเห็นจริงจากชีต "ความเห็นครู"
    let realComment = '';
    try {
      const ss = SS();
      const commentSheet = ss.getSheetByName('ความเห็นครู');
      if (commentSheet) {
        const data = commentSheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          let idCol = -1, commentCol = -1;
          for (let c = 0; c < headers.length; c++) {
            const h = String(headers[c]).trim();
            if (h === 'รหัสนักเรียน' || h === 'student_id') idCol = c;
            if (h === 'ความเห็นครู' || h === 'comment') commentCol = c;
          }
          if (idCol === -1) idCol = 0;
          if (commentCol === -1) commentCol = 4;
          
          const sid = String(studentId).trim();
          for (let i = 1; i < data.length; i++) {
            if (String(data[i][idCol]).trim() === sid) {
              realComment = data[i][commentCol] ? String(data[i][commentCol]).trim() : '';
              break;
            }
          }
        }
      }
    } catch (commentErr) {
      Logger.log('⚠️ Error reading comment sheet: ' + commentErr.message);
    }

    return {
      success: true,
      student: {
        id: student.id,
        fullName: `${student.title}${student.firstname} ${student.lastname}`.trim(),
        grade: student.grade,
        classNo: student.classNo
      },
      comment: realComment
    };

  } catch (e) {
    Logger.log(`Error in getStudentFullData: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * 🆕 [แก้ไขล่าสุด] บันทึกความคิดเห็นลงในชีต "ความเห็นครู" (โครงสร้างที่ถูกต้อง)
 * @param {string} studentId - รหัสนักเรียน
 * @param {string} comment - ข้อความความคิดเห็น
 * @returns {object} ผลลัพธ์การบันทึก
 */
function saveCommentWithValidation(studentId, comment) {
  try {
    // --- 1. ตรวจสอบข้อมูลเบื้องต้น ---
    if (!studentId) {
      return { success: false, error: 'ไม่พบรหัสนักเรียน' };
    }
    if (!comment || comment.trim().length < 10) {
      return { success: false, error: 'ความคิดเห็นต้องมีอย่างน้อย 10 ตัวอักษร' };
    }

    // --- 2. [ใหม่] ดึงข้อมูลนักเรียนทั้งหมดก่อน เพื่อให้มีข้อมูลครบถ้วน ---
    const studentInfo = getStudentFullData(studentId);
    if (!studentInfo.success) {
      return { success: false, error: 'ไม่พบข้อมูลนักเรียนรหัส ' + studentId };
    }
    const student = studentInfo.student;

    const ss = SS();
    const sheetName = "ความเห็นครู"; // ใช้ชีต "ความเห็นครู"
    let sheet = ss.getSheetByName(sheetName);
    const timestamp = new Date();

// --- 3. สร้างชีต "ความเห็นครู" หากยังไม่มี (พร้อม Header ที่ถูกต้อง) ---
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      // ‼️ เปลี่ยนเป็นภาษาไทยให้ตรงกับภาพ
      const headers = ["รหัสนักเรียน", "ชื่อนักเรียน", "ชั้น", "ห้อง", "ความเห็นครู", "วันที่บันทึก", "วันที่อัปเดต"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      // ...
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const commentColIndex = headers.indexOf('ความเห็นครู');
const updatedColIndex = headers.indexOf('วันที่อัปเดต');

    // --- 4. [เพิ่มการตรวจสอบ] ตรวจสอบว่าหาหัวตารางเจอหรือไม่ ---
    if (commentColIndex === -1 || updatedColIndex === -1) {
      Logger.log(`Error: ไม่พบหัวตาราง 'comment' (พบที่: ${commentColIndex}) หรือ 'last_updated' (พบที่: ${updatedColIndex}) ในชีต ${sheetName}`);
      return { 
        success: false, 
        error: `โครงสร้างชีต '${sheetName}' ผิดพลาด (ไม่พบหัวตาราง 'comment' หรือ 'last_updated')`
      };
    }

    // --- 5. ค้นหาแถวของนักเรียน ---
    let studentRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(studentId).trim()) {
        studentRowIndex = i + 1; // ‼️ (i คือ index ของ array แต่แถวในชีตคือ i + 1)
        break;
      }
    }

    // --- 6. อัปเดตข้อมูลหรือเพิ่มแถวใหม่ (ด้วยข้อมูลที่ครบถ้วน) ---
    if (studentRowIndex !== -1) {
      // ถ้าเจอแถวเดิม -> อัปเดตเฉพาะช่อง comment และ last_updated
      // ‼️ (commentColIndex คือ 0-based index แต่คอลัมน์ในชีตคือ index + 1)
      sheet.getRange(studentRowIndex, commentColIndex + 1).setValue(comment);
      sheet.getRange(studentRowIndex, updatedColIndex + 1).setValue(timestamp);
      Logger.log(`Updated comment for student ID: ${studentId}`);
    } else {
      // ถ้าไม่เจอ -> เพิ่มแถวใหม่พร้อมข้อมูลทั้งหมด
      const newRow = [
        student.id,
        student.fullName,
        student.grade,
        student.classNo,
        comment,
        timestamp, // created_date
        timestamp  // last_updated
      ];
      sheet.appendRow(newRow);
      Logger.log(`Added new comment for student ID: ${studentId}`);
    }

    // --- 7. ส่งผลลัพธ์กลับ ---
    return { 
      success: true, 
      message: 'บันทึกความคิดเห็นเรียบร้อยแล้ว',
      student: student // ส่งข้อมูลนักเรียนกลับไปอัปเดตหน้าเว็บ
    };

  } catch (e) {
    Logger.log(`Error in saveCommentWithValidation: ${e.message}`);
    return { success: false, error: e.message };
  }
}

