// ============================================================
// ✅ VALIDATION.GS - Validation Module
// ============================================================
// ไฟล์นี้รวม Validation Functions ทั้งหมดที่ใช้ร่วมกันในระบบ
// เพื่อลดการซ้ำซ้อนของโค้ดและรักษามาตรฐานการตรวจสอบข้อมูล
// ============================================================

// ============================================================
// 📋 CONSTANTS - ค่าคงที่สำหรับ Validation
// ============================================================

const V_VALID_GRADES = ['อ.1', 'อ.2', 'อ.3', 'ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6', 'ม.1', 'ม.2', 'ม.3'];
const V_VALID_CLASS_NOS = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
const V_VALID_GENDERS = ['ชาย', 'หญิง', 'ช', 'ญ'];
const V_VALID_SEMESTERS = ['1', '2'];

// ============================================================
// 🎯 BASIC VALIDATION FUNCTIONS
// ============================================================

/**
 * ตรวจสอบว่าค่าไม่เป็น null, undefined หรือ empty string
 * @param {any} value - ค่าที่ต้องการตรวจสอบ
 * @param {string} fieldName - ชื่อฟิลด์ (สำหรับ error message)
 * @returns {Object} { valid: boolean, error: string }
 */
function validateRequired(value, fieldName = 'ข้อมูล') {
  if (value === null || value === undefined || value === '' || 
      (typeof value === 'string' && value.trim() === '')) {
    return {
      valid: false,
      error: `${fieldName}จำเป็นต้องระบุ`
    };
  }
  return { valid: true };
}

/**
 * ตรวจสอบว่าค่าเป็นตัวเลข
 * @param {any} value - ค่าที่ต้องการตรวจสอบ
 * @param {string} fieldName - ชื่อฟิลด์
 * @returns {Object} { valid: boolean, error: string, value: number }
 */
function validateNumber(value, fieldName = 'ตัวเลข') {
  const num = Number(value);
  if (isNaN(num)) {
    return {
      valid: false,
      error: `${fieldName}ต้องเป็นตัวเลข`
    };
  }
  return { valid: true, value: num };
}

/**
 * ตรวจสอบว่าค่าอยู่ในช่วงที่กำหนด
 * @param {number} value - ค่าที่ต้องการตรวจสอบ
 * @param {number} min - ค่าต่ำสุด
 * @param {number} max - ค่าสูงสุด
 * @param {string} fieldName - ชื่อฟิลด์
 * @returns {Object} { valid: boolean, error: string }
 */
function validateRange(value, min, max, fieldName = 'ค่า') {
  const numCheck = validateNumber(value, fieldName);
  if (!numCheck.valid) return numCheck;
  
  const num = numCheck.value;
  if (num < min || num > max) {
    return {
      valid: false,
      error: `${fieldName}ต้องอยู่ระหว่าง ${min} ถึง ${max}`
    };
  }
  return { valid: true, value: num };
}

/**
 * ตรวจสอบว่าค่าอยู่ใน Array ที่กำหนด
 * @param {any} value - ค่าที่ต้องการตรวจสอบ
 * @param {Array} validValues - Array ของค่าที่ถูกต้อง
 * @param {string} fieldName - ชื่อฟิลด์
 * @returns {Object} { valid: boolean, error: string }
 */
function validateInArray(value, validValues, fieldName = 'ค่า') {
  const strValue = String(value).trim();
  if (!validValues.includes(strValue)) {
    return {
      valid: false,
      error: `${fieldName}ไม่ถูกต้อง (ต้องเป็น: ${validValues.join(', ')})`
    };
  }
  return { valid: true, value: strValue };
}

// ============================================================
// 🎓 GRADE & CLASS VALIDATION
// ============================================================

/**
 * ตรวจสอบระดับชั้น
 * @param {string} grade - ระดับชั้น
 * @returns {Object} { valid: boolean, error: string, value: string }
 */
function validateGrade(grade) {
  const requiredCheck = validateRequired(grade, 'ระดับชั้น');
  if (!requiredCheck.valid) return requiredCheck;
  
  return validateInArray(grade, V_VALID_GRADES, 'ระดับชั้น');
}

/**
 * ตรวจสอบหมายเลขห้อง
 * @param {string|number} classNo - หมายเลขห้อง
 * @returns {Object} { valid: boolean, error: string, value: string }
 */
function validateClassNo(classNo) {
  const requiredCheck = validateRequired(classNo, 'หมายเลขห้อง');
  if (!requiredCheck.valid) return requiredCheck;
  
  return validateInArray(String(classNo), V_VALID_CLASS_NOS, 'หมายเลขห้อง');
}

/**
 * ตรวจสอบระดับชั้นและห้องพร้อมกัน
 * @param {string} grade - ระดับชั้น
 * @param {string|number} classNo - หมายเลขห้อง
 * @returns {Object} { valid: boolean, error: string, grade: string, classNo: string }
 */
function validateGradeAndClass(grade, classNo) {
  const gradeCheck = validateGrade(grade);
  if (!gradeCheck.valid) return gradeCheck;
  
  const classCheck = validateClassNo(classNo);
  if (!classCheck.valid) return classCheck;
  
  return {
    valid: true,
    grade: gradeCheck.value,
    classNo: classCheck.value
  };
}

// ============================================================
// 👤 STUDENT VALIDATION
// ============================================================

/**
 * ตรวจสอบรหัสนักเรียน
 * @param {string|number} studentId - รหัสนักเรียน
 * @returns {Object} { valid: boolean, error: string, value: string }
 */
function validateStudentId(studentId) {
  const requiredCheck = validateRequired(studentId, 'รหัสนักเรียน');
  if (!requiredCheck.valid) return requiredCheck;
  
  const strId = String(studentId).trim();
  
  // ตรวจสอบว่าเป็นตัวเลข
  if (!/^\d+$/.test(strId)) {
    return {
      valid: false,
      error: 'รหัสนักเรียนต้องเป็นตัวเลขเท่านั้น'
    };
  }
  
  // ตรวจสอบความยาว (ปรับตามระบบของคุณ)
  if (strId.length < 3 || strId.length > 10) {
    return {
      valid: false,
      error: 'รหัสนักเรียนต้องมีความยาว 3-10 หลัก'
    };
  }
  
  return { valid: true, value: strId };
}

/**
 * ตรวจสอบว่ารหัสนักเรียนมีอยู่ในระบบ
 * @param {string|number} studentId - รหัสนักเรียน
 * @returns {Object} { valid: boolean, error: string, student: Object }
 */
function validateStudentExists(studentId) {
  const idCheck = validateStudentId(studentId);
  if (!idCheck.valid) return idCheck;
  
  try {
    const student = U_getStudentById(idCheck.value);
    if (!student) {
      return {
        valid: false,
        error: `ไม่พบรหัสนักเรียน ${idCheck.value} ในระบบ`
      };
    }
    
    return {
      valid: true,
      student: student
    };
  } catch (e) {
    return {
      valid: false,
      error: `เกิดข้อผิดพลาดในการตรวจสอบรหัสนักเรียน: ${e.message}`
    };
  }
}

/**
 * ตรวจสอบเพศ
 * @param {string} gender - เพศ
 * @returns {Object} { valid: boolean, error: string, value: string }
 */
function validateGender(gender) {
  if (!gender) return { valid: true, value: '' }; // เพศไม่บังคับ
  
  return validateInArray(gender, V_VALID_GENDERS, 'เพศ');
}

// ============================================================
// 📊 SCORE VALIDATION
// ============================================================

/**
 * ตรวจสอบคะแนน
 * @param {number} score - คะแนน
 * @param {number} maxScore - คะแนนเต็ม
 * @param {string} fieldName - ชื่อฟิลด์
 * @returns {Object} { valid: boolean, error: string, value: number }
 */
function validateScore(score, maxScore = 100, fieldName = 'คะแนน') {
  // อนุญาตให้เป็นค่าว่างได้
  if (score === null || score === undefined || score === '') {
    return { valid: true, value: 0 };
  }
  
  return validateRange(score, 0, maxScore, fieldName);
}

/**
 * ตรวจสอบเกรด
 * @param {number} grade - เกรด
 * @returns {Object} { valid: boolean, error: string, value: number }
 */
function validateGradeScore(grade) {
  const validGrades = [0, 1, 1.5, 2, 2.5, 3, 3.5, 4];
  
  if (grade === null || grade === undefined || grade === '') {
    return { valid: true, value: 0 };
  }
  
  const numCheck = validateNumber(grade, 'เกรด');
  if (!numCheck.valid) return numCheck;
  
  if (!validGrades.includes(numCheck.value)) {
    return {
      valid: false,
      error: `เกรดต้องเป็นค่าใดค่าหนึ่งต่อไปนี้: ${validGrades.join(', ')}`
    };
  }
  
  return { valid: true, value: numCheck.value };
}

// ============================================================
// 📅 DATE & SEMESTER VALIDATION
// ============================================================

/**
 * ตรวจสอบภาคเรียน
 * @param {string|number} semester - ภาคเรียน
 * @returns {Object} { valid: boolean, error: string, value: string }
 */
function validateSemester(semester) {
  const requiredCheck = validateRequired(semester, 'ภาคเรียน');
  if (!requiredCheck.valid) return requiredCheck;
  
  return validateInArray(String(semester), V_VALID_SEMESTERS, 'ภาคเรียน');
}

/**
 * ตรวจสอบปีการศึกษา
 * @param {number} year - ปีการศึกษา (พ.ศ.)
 * @returns {Object} { valid: boolean, error: string, value: number }
 */
function validateAcademicYear(year) {
  const requiredCheck = validateRequired(year, 'ปีการศึกษา');
  if (!requiredCheck.valid) return requiredCheck;
  
  const numCheck = validateNumber(year, 'ปีการศึกษา');
  if (!numCheck.valid) return numCheck;
  
  const currentYear = U_getCurrentAcademicYear();
  
  // ตรวจสอบว่าปีการศึกษาอยู่ในช่วงที่สมเหตุสมผล (ย้อนหลัง 5 ปี ถึง ล่วงหน้า 2 ปี)
  if (numCheck.value < currentYear - 5 || numCheck.value > currentYear + 2) {
    return {
      valid: false,
      error: `ปีการศึกษาต้องอยู่ระหว่าง ${currentYear - 5} ถึง ${currentYear + 2}`
    };
  }
  
  return { valid: true, value: numCheck.value };
}

/**
 * ตรวจสอบวันที่
 * @param {string|Date} date - วันที่
 * @returns {Object} { valid: boolean, error: string, value: Date }
 */
function validateDate(date) {
  const requiredCheck = validateRequired(date, 'วันที่');
  if (!requiredCheck.valid) return requiredCheck;
  
  let dateObj;
  
  if (date instanceof Date) {
    dateObj = date;
  } else {
    dateObj = new Date(date);
  }
  
  if (isNaN(dateObj.getTime())) {
    return {
      valid: false,
      error: 'รูปแบบวันที่ไม่ถูกต้อง'
    };
  }
  
  return { valid: true, value: dateObj };
}

// ============================================================
// 📝 ASSESSMENT VALIDATION
// ============================================================

/**
 * ตรวจสอบข้อมูลการประเมินคุณลักษณะ
 * @param {Object} data - ข้อมูลการประเมิน
 * @returns {Object} { valid: boolean, errors: Array, data: Object }
 */
function validateCharacteristicAssessment(data) {
  const errors = [];
  const validated = {};
  
  // ตรวจสอบ studentId
  const studentCheck = validateStudentId(data.studentId);
  if (!studentCheck.valid) {
    errors.push(studentCheck.error);
  } else {
    validated.studentId = studentCheck.value;
  }
  
  // ตรวจสอบ grade และ classNo
  const gradeClassCheck = validateGradeAndClass(data.grade, data.classNo);
  if (!gradeClassCheck.valid) {
    errors.push(gradeClassCheck.error);
  } else {
    validated.grade = gradeClassCheck.grade;
    validated.classNo = gradeClassCheck.classNo;
  }
  
  // ตรวจสอบ semester
  const semesterCheck = validateSemester(data.semester);
  if (!semesterCheck.valid) {
    errors.push(semesterCheck.error);
  } else {
    validated.semester = semesterCheck.value;
  }
  
  // ตรวจสอบ academicYear
  const yearCheck = validateAcademicYear(data.academicYear);
  if (!yearCheck.valid) {
    errors.push(yearCheck.error);
  } else {
    validated.academicYear = yearCheck.value;
  }
  
  // ตรวจสอบคะแนนแต่ละด้าน (ถ้ามี)
  if (data.scores) {
    validated.scores = {};
    Object.keys(data.scores).forEach(key => {
      const scoreCheck = validateScore(data.scores[key], 4, key);
      if (!scoreCheck.valid) {
        errors.push(scoreCheck.error);
      } else {
        validated.scores[key] = scoreCheck.value;
      }
    });
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    data: validated
  };
}

/**
 * ตรวจสอบข้อมูลการประเมินกิจกรรม
 * @param {Object} data - ข้อมูลการประเมิน
 * @returns {Object} { valid: boolean, errors: Array, data: Object }
 */
function validateActivityAssessment(data) {
  const errors = [];
  const validated = {};
  
  // ตรวจสอบพื้นฐานเหมือน characteristic
  const baseCheck = validateCharacteristicAssessment(data);
  if (!baseCheck.valid) {
    return baseCheck;
  }
  
  validated.studentId = baseCheck.data.studentId;
  validated.grade = baseCheck.data.grade;
  validated.classNo = baseCheck.data.classNo;
  validated.semester = baseCheck.data.semester;
  validated.academicYear = baseCheck.data.academicYear;
  
  // ตรวจสอบชั่วโมงกิจกรรม
  if (data.hours !== undefined) {
    const hoursCheck = validateRange(data.hours, 0, 200, 'ชั่วโมงกิจกรรม');
    if (!hoursCheck.valid) {
      errors.push(hoursCheck.error);
    } else {
      validated.hours = hoursCheck.value;
    }
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    data: validated
  };
}

// ============================================================
// 🔐 BATCH VALIDATION
// ============================================================

/**
 * ตรวจสอบข้อมูลแบบ Batch
 * @param {Array} dataArray - Array ของข้อมูลที่ต้องการตรวจสอบ
 * @param {Function} validatorFn - ฟังก์ชัน Validator
 * @returns {Object} { valid: boolean, results: Array, errors: Array }
 */
function validateBatch(dataArray, validatorFn) {
  if (!Array.isArray(dataArray)) {
    return {
      valid: false,
      errors: ['ข้อมูลต้องเป็น Array'],
      results: []
    };
  }
  
  const results = [];
  const allErrors = [];
  
  dataArray.forEach((item, index) => {
    const result = validatorFn(item);
    results.push(result);
    
    if (!result.valid) {
      allErrors.push({
        index: index,
        errors: result.errors || [result.error]
      });
    }
  });
  
  return {
    valid: allErrors.length === 0,
    results: results,
    errors: allErrors
  };
}

// ============================================================
// 🛡️ SANITIZATION FUNCTIONS
// ============================================================

/**
 * ทำความสะอาดข้อความ (ป้องกัน XSS)
 * @param {string} text - ข้อความ
 * @returns {string} ข้อความที่สะอาดแล้ว
 */
function sanitizeText(text) {
  if (!text) return '';
  
  return String(text)
    .trim()
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;');
}

/**
 * ทำความสะอาดตัวเลข
 * @param {any} value - ค่าที่ต้องการทำความสะอาด
 * @returns {number} ตัวเลขที่สะอาดแล้ว
 */
function sanitizeNumber(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * ทำความสะอาด Object
 * @param {Object} obj - Object ที่ต้องการทำความสะอาด
 * @returns {Object} Object ที่สะอาดแล้ว
 */
function sanitizeObject(obj) {
  if (!obj || typeof obj !== 'object') return {};
  
  const cleaned = {};
  
  Object.keys(obj).forEach(key => {
    const value = obj[key];
    
    if (typeof value === 'string') {
      cleaned[key] = sanitizeText(value);
    } else if (typeof value === 'number') {
      cleaned[key] = sanitizeNumber(value);
    } else if (Array.isArray(value)) {
      cleaned[key] = value.map(item => 
        typeof item === 'object' ? sanitizeObject(item) : sanitizeText(String(item))
      );
    } else if (typeof value === 'object' && value !== null) {
      cleaned[key] = sanitizeObject(value);
    } else {
      cleaned[key] = value;
    }
  });
  
  return cleaned;
}

// ============================================================
// 🎯 HELPER FUNCTIONS
// ============================================================

/**
 * สร้าง Error Response แบบมาตรฐาน
 * @param {string|Array} errors - Error message(s)
 * @returns {Object} Error response
 */
function createValidationError(errors) {
  const errorArray = Array.isArray(errors) ? errors : [errors];
  
  return {
    valid: false,
    success: false,
    errors: errorArray,
    message: errorArray.join(', ')
  };
}

/**
 * สร้าง Success Response แบบมาตรฐาน
 * @param {Object} data - ข้อมูลที่ validated แล้ว
 * @returns {Object} Success response
 */
function createValidationSuccess(data) {
  return {
    valid: true,
    success: true,
    data: data
  };
}

// ============================================================
// 🧪 TESTING FUNCTION
// ============================================================

