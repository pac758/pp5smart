// ============================================================
// 🛠️ UTILS.GS - Utility Functions Module
// ============================================================
// ไฟล์นี้รวม Helper Functions ทั้งหมดที่ใช้ร่วมกันในระบบ
// เพื่อลดการซ้ำซ้อนของโค้ดและง่ายต่อการบำรุงรักษา
// ============================================================

/** * เปิด Spreadsheet หลักของระบบ * @returns {Spreadsheet} Spreadsheet object */function U_SS() {  if (typeof SPREADSHEET_ID === 'undefined' || !SPREADSHEET_ID) {    throw new Error('SPREADSHEET_ID ไม่ได้ถูกกำหนด');  }  return SS();}

/** * ดึงชีตตามชื่อ พร้อมตรวจสอบว่ามีอยู่จริง * @param {string} sheetName - ชื่อชีต * @returns {Sheet} Sheet object */function U_getSheet(sheetName) {  const ss = U_SS();  const sheet = ss.getSheetByName(sheetName);  if (!sheet) {    throw new Error(`ไม่พบชีต: ${sheetName}`);  }  return sheet;}

/** * ตรวจสอบว่าชีตมีอยู่หรือไม่ * @param {string} sheetName - ชื่อชีต * @returns {boolean} */function U_sheetExists(sheetName) {  try {    const ss = U_SS();    return ss.getSheetByName(sheetName) !== null;  } catch (e) {    return false;  }}

/** * อ่านข้อมูลจากชีตเป็น Array of Objects * @param {string} sheetName - ชื่อชีต * @returns {Array<Object>} ข้อมูลในรูปแบบ Object */function U_getSheetData(sheetName) {  const sheet = U_getSheet(sheetName);  const data = sheet.getDataRange().getValues();    if (data.length === 0) {    return [];  }    const headers = data[0];  const rows = data.slice(1);    return rows.map(row => {    const obj = {};    headers.forEach((header, index) => {      obj[header] = row[index];    });    return obj;  });}

/** * ดึงรายชื่อนักเรียนตามชั้นและห้อง (ฟังก์ชันหลักที่ใช้ทั่วทั้งระบบ) * @param {string} grade - ระดับชั้น เช่น "ป.1" * @param {string} classNo - หมายเลขห้อง เช่น "1" * @returns {Array<Object>} รายชื่อนักเรียน */function U_getStudentsByClass(grade, classNo) {  const sheet = U_getSheet('Students');  const data = sheet.getDataRange().getValues();    if (data.length <= 1) {    return [];  }    const headers = data[0];  const rows = data.slice(1);     const colMap = {};  headers.forEach((header, index) => {    colMap[header.trim()] = index;  });    const students = [];    rows.forEach(row => {    const rowGrade = String(row[colMap['grade']] || '').trim();    const rowClass = String(row[colMap['class_no']] || '').trim();        if (rowGrade === grade.trim() && rowClass === classNo.toString().trim()) {      const student = {        id: row[colMap['student_id']] || '',        studentId: row[colMap['student_id']] || '',        idCard: row[colMap['id_card']] || '',        title: String(row[colMap['title']] || '').trim(),        firstname: String(row[colMap['firstname']] || '').trim(),        lastname: String(row[colMap['lastname']] || '').trim(),        grade: rowGrade,        classNo: rowClass,        class_no: rowClass,        gender: row[colMap['gender']] || ''      };             student.name = `${student.title}${student.firstname} ${student.lastname}`.trim();            students.push(student);    }  });     students.sort((a, b) => {    const idA = String(a.id || '');    const idB = String(b.id || '');    return idA.localeCompare(idB, 'th', { numeric: true });  });    return students;}

/** * Alias สำหรับความเข้ากันได้กับโค้ดเดิม */function U_getStudentsForAttendance(grade, classNo) {  return U_getStudentsByClass(grade, classNo);}

/** * ดึงข้อมูลนักเรียนทั้งหมด * @returns {Array<Object>} รายชื่อนักเรียนทั้งหมด */function U_getAllStudents() {  return U_getSheetData('Students');}

/** * ดึงข้อมูลนักเรียนจากรหัส * @param {string} studentId - รหัสนักเรียน * @returns {Object|null} ข้อมูลนักเรียน */function U_getStudentById(studentId) {  const students = U_getAllStudents();  return students.find(s => String(s.student_id) === String(studentId)) || null;}

/** * แปลงปี ค.ศ. เป็น พ.ศ. * @param {number} year - ปี ค.ศ. * @returns {number} ปี พ.ศ. */function U_toBuddhistYear(year) {  return parseInt(year) + 543;}

/** * แปลงปี พ.ศ. เป็น ค.ศ. * @param {number} year - ปี พ.ศ. * @returns {number} ปี ค.ศ. */function U_toChristianYear(year) {  return parseInt(year) - 543;}

/** * คำนวณปีการศึกษาปัจจุบัน (พ.ศ.) * @returns {number} ปีการศึกษา พ.ศ. */function U_getCurrentAcademicYear() {  const now = new Date();  const year = now.getFullYear();  const month = now.getMonth() + 1;      return month >= 5 ? year + 543 : year + 542;}

/** * คำนวณปีการศึกษาจากวันที่ * @param {Date} date - วันที่ * @returns {number} ปีการศึกษา พ.ศ. */function U_getAcademicYearFromDate(date) {  const year = date.getFullYear();  const month = date.getMonth() + 1;  return month >= 5 ? year + 543 : year + 542;}

/** * ชื่อเดือนภาษาไทย */const U_MONTH_NAMES_TH = [  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน',  'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม',  'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];/** * ชื่อเดือนภาษาไทยแบบย่อ */const U_MONTH_NAMES_TH_SHORT = [  'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',  'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];/** * แปลงวันที่เป็นข้อความภาษาไทย * @param {Date} date - วันที่ * @param {boolean} short - ใช้ชื่อเดือนแบบย่อ * @returns {string} วันที่ในรูปแบบไทย */function U_formatThaiDate(date, short = false) {  const day = date.getDate();  const month = short ? U_MONTH_NAMES_TH_SHORT[date.getMonth()] : U_MONTH_NAMES_TH[date.getMonth()];  const year = U_toBuddhistYear(date.getFullYear());  return `${day} ${month} ${year}`;}

/** * แปลงเดือนปัจจุบันเป็นรูปแบบภาษาไทยพร้อมปี พ.ศ. * @returns {string} เดือนปัจจุบัน เช่น "ตุลาคม2568" */function U_getCurrentThaiMonth() {  const now = new Date();  const month = now.getMonth();  const year = U_toBuddhistYear(now.getFullYear());  return U_MONTH_NAMES_TH[month] + year;}

/** * Cache Manager แบบรวมศูนย์ */const U_CacheManager = {  /**   * ดึงข้อมูลจาก Cache   * @param {string} key - Cache key   * @returns {any|null} ข้อมูลที่แคช หรือ null ถ้าไม่มี   */  get(key) {    try {      const cache = CacheService.getScriptCache();      const cached = cache.get(key);      if (cached) {        Logger.log(`✅ Cache HIT: ${key}`);        return JSON.parse(cached);      }      Logger.log(`❌ Cache MISS: ${key}`);      return null;    } catch (e) {      Logger.log(`⚠️ Cache get error: ${e.message}`);      return null;    }  },    /**   * บันทึกข้อมูลลง Cache   * @param {string} key - Cache key   * @param {any} data - ข้อมูลที่จะแคช   * @param {number} ttl - เวลาหมดอายุ (วินาที) default 300 = 5 นาที   */  set(key, data, ttl = 300) {    try {      const cache = CacheService.getScriptCache();      cache.put(key, JSON.stringify(data), ttl);      Logger.log(`💾 Cache SET: ${key} (TTL: ${ttl}s)`);    } catch (e) {      Logger.log(`⚠️ Cache set error: ${e.message}`);    }  },    /**   * ลบข้อมูลจาก Cache   * @param {string} key - Cache key   */  remove(key) {    try {      const cache = CacheService.getScriptCache();      cache.remove(key);      Logger.log(`🗑️ Cache REMOVED: ${key}`);    } catch (e) {      Logger.log(`⚠️ Cache remove error: ${e.message}`);    }  },    /**   * ลบข้อมูลหลายตัวจาก Cache   * @param {Array<string>} keys - Array ของ cache keys   */  removeAll(keys) {    try {      const cache = CacheService.getScriptCache();      cache.removeAll(keys);      Logger.log(`🗑️ Cache REMOVED ALL: ${keys.join(', ')}`);    } catch (e) {      Logger.log(`⚠️ Cache removeAll error: ${e.message}`);    }  }};   /** * ใช้ Lock เพื่อป้องกัน Race Condition * @param {string} lockName - ชื่อ Lock * @param {Function} callback - ฟังก์ชันที่จะทำงานภายใต้ Lock * @param {number} timeout - เวลารอ Lock (มิลลิวินาที) default 30000 * @returns {any} ผลลัพธ์จาก callback */function U_withLock(lockName, callback, timeout = 30000) {  const lock = LockService.getScriptLock();    try {    lock.waitLock(timeout);    Logger.log(`🔒 Lock acquired: ${lockName}`);        const result = callback();        return result;  } catch (e) {    Logger.log(`❌ Lock error (${lockName}): ${e.message}`);    throw new Error(`ไม่สามารถล็อกทรัพยากรได้: ${e.message}`);  } finally {    lock.releaseLock();    Logger.log(`🔓 Lock released: ${lockName}`);  }}

/** * ตรวจสอบว่าเป็นค่าว่างหรือไม่ * @param {any} value - ค่าที่ต้องการตรวจสอบ * @returns {boolean} */function U_isEmpty(value) {  return value === null || value === undefined || value === '' ||          (typeof value === 'string' && value.trim() === '');}

/** * แปลงค่าเป็นตัวเลข * @param {any} value - ค่าที่ต้องการแปลง * @param {number} defaultValue - ค่า default ถ้าแปลงไม่ได้ * @returns {number} */function U_toNumber(value, defaultValue = 0) {  const num = Number(value);  return isNaN(num) ? defaultValue : num;}

/** * แปลงค่าเป็น String และ trim * @param {any} value - ค่าที่ต้องการแปลง * @returns {string} */function U_toString(value) {  return String(value || '').trim();}

/** * ตัดทศนิยม * @param {number} value - ตัวเลข * @param {number} decimals - จำนวนทศนิยม * @returns {number} */function U_round(value, decimals = 2) {  const multiplier = Math.pow(10, decimals);  return Math.round(value * multiplier) / multiplier;}

/** * แปลงชื่อชั้นเรียนแบบย่อเป็นชื่อเต็ม * @param {string} grade - ชั้นเรียนแบบย่อ เช่น "ป.1" * @returns {string} ชื่อเต็ม เช่น "ประถมศึกษาปีที่ 1" */function U_getGradeFullName(grade) {  const gradeMap = {    'อ.1': 'อนุบาลปีที่ 1',    'อ.2': 'อนุบาลปีที่ 2',    'อ.3': 'อนุบาลปีที่ 3',    'ป.1': 'ประถมศึกษาปีที่ 1',    'ป.2': 'ประถมศึกษาปีที่ 2',    'ป.3': 'ประถมศึกษาปีที่ 3',    'ป.4': 'ประถมศึกษาปีที่ 4',    'ป.5': 'ประถมศึกษาปีที่ 5',    'ป.6': 'ประถมศึกษาปีที่ 6',    'ม.1': 'มัธยมศึกษาปีที่ 1',    'ม.2': 'มัธยมศึกษาปีที่ 2',    'ม.3': 'มัธยมศึกษาปีที่ 3'  };  return gradeMap[grade] || grade;}

/** * คำนวณเกรดจากคะแนน * @param {number} score - คะแนน * @returns {number} เกรด */function U_calculateGrade(score) {  if (score >= 80) return 4;  if (score >= 75) return 3.5;  if (score >= 70) return 3;  if (score >= 65) return 2.5;  if (score >= 60) return 2;  if (score >= 55) return 1.5;  if (score >= 50) return 1;  return 0;}

/** * แปลงเกรดเป็นข้อความ * @param {number} grade - เกรด * @returns {string} */function U_gradeToText(grade) {  const gradeMap = {    4: 'ดีเยี่ยม',    3.5: 'ดีมาก',    3: 'ดี',    2.5: 'ค่อนข้างดี',    2: 'พอใช้',    1.5: 'ผ่าน',    1: 'ผ่านเกณฑ์ขั้นต่ำ',    0: 'ไม่ผ่าน'  };  return gradeMap[grade] || 'ไม่ระบุ';}

/** * จัดการ Error แบบรวมศูนย์ * @param {Error} error - Error object * @param {string} context - บริบทของ Error * @returns {Object} Error response */function U_handleError(error, context = '') {  const errorMessage = error.message || 'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ';  const fullMessage = context ? `${context}: ${errorMessage}` : errorMessage;    Logger.log(`❌ ERROR ${context ? `[${context}]` : ''}: ${errorMessage}`);  if (error.stack) {    Logger.log(`📚 Stack: ${error.stack}`);  }    return {    success: false,    error: true,    message: fullMessage,    originalError: errorMessage  };}

/** * Wrapper สำหรับ try-catch ที่ส่ง Error response กลับ * @param {Function} fn - ฟังก์ชันที่จะ execute * @param {string} context - บริบท * @returns {any} ผลลัพธ์หรือ Error response */function U_tryCatch(fn, context = '') {  try {    return fn();  } catch (error) {    return U_handleError(error, context);  }}

/** * ลบค่าซ้ำออกจาก Array * @param {Array} array - Array ต้นฉบับ * @returns {Array} Array ที่ไม่มีค่าซ้ำ */function U_unique(array) {  return [...new Set(array)];}

/** * จัดกลุ่มข้อมูลตาม key * @param {Array} array - Array ของ Objects * @param {string} key - Key ที่จะใช้จัดกลุ่ม * @returns {Object} Object ที่จัดกลุ่มแล้ว */function U_groupBy(array, key) {  return array.reduce((result, item) => {    const groupKey = item[key];    if (!result[groupKey]) {      result[groupKey] = [];    }    result[groupKey].push(item);    return result;  }, {});}

/** * เรียงลำดับ Array ของ Objects * @param {Array} array - Array ที่จะเรียง * @param {string} key - Key ที่จะใช้เรียง * @param {boolean} ascending - เรียงจากน้อยไปมาก (default: true) * @returns {Array} Array ที่เรียงแล้ว */function U_sortBy(array, key, ascending = true) {  return array.sort((a, b) => {    const aVal = a[key];    const bVal = b[key];        if (typeof aVal === 'string') {      return ascending         ? aVal.localeCompare(bVal, 'th', { numeric: true })        : bVal.localeCompare(aVal, 'th', { numeric: true });    }        return ascending ? aVal - bVal : bVal - aVal;  });}

/** Log แบบมีสี (ใช้ใน Logger) */
const U_Log = {  info: (msg) => Logger.log(`ℹ️ ${msg}`),  success: (msg) => Logger.log(`✅ ${msg}`),  warning: (msg) => Logger.log(`⚠️ ${msg}`),  error: (msg) => Logger.log(`❌ ${msg}`),  debug: (msg) => Logger.log(`🐛 ${msg}`)};

/** ฟังก์ชันทดสอบว่า utils.gs ทำงานได้ */
function testUtils() {  U_Log.info('Testing utils.gs...');     const currentYear = U_getCurrentAcademicYear();  U_Log.success(`Current academic year: ${currentYear}`);     U_CacheManager.set('test_key', { value: 'test' }, 60);  const cached = U_CacheManager.get('test_key');  U_Log.success(`Cache test: ${cached ? 'PASS' : 'FAIL'}`);     try {    const students = U_getStudentsByClass('ป.1', '1');    U_Log.success(`Found ${students.length} students`);  } catch (e) {    U_Log.warning(`Students test skipped: ${e.message}`);  }    U_Log.success('Utils.gs is working!');  return 'All tests passed';}

/** * ���� Index �ͧ�������ҡ���ͷ������� * @param {Array} headers - ��������ͧ��Ǥ������ * @param {Array} possibleNames - ��������ͧ���ͷ���Ҩ����� * @returns {number} Index �ͧ������� ���� -1 �����辺 */function U_findColumnIndex(headers, possibleNames) {  for (let i = 0; i < headers.length; i++) {    const headerName = String(headers[i] || '').trim().toLowerCase();    for (const name of possibleNames) {      if (headerName === String(name).toLowerCase() || headerName.includes(String(name).toLowerCase())) {        return i;      }    }  }  return -1;}

/** * Format string �ѹ������������ٻẺ YYYY-MM-DD * @param {Date|string} date - �ѹ��� * @returns {string} �ѹ�����ٻẺ YYYY-MM-DD ����ʵ�ԧ��ҧ */function U_formatToDateInput(date) {  if (!date) return '';  try {    let d = (date instanceof Date) ? date : new Date(date);    if (isNaN(d.getTime())) return '';    const year = d.getFullYear();    const month = (d.getMonth() + 1).toString().padStart(2, '0');    const day = d.getDate().toString().padStart(2, '0');    return `${year}-${month}-${day}`;  } catch (error) {    return '';  }}



// ============================================================
// 📚 STANDARD SUBJECT ORDER
// ============================================================
var STANDARD_SUBJECT_ORDER = [
  'ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา',
  'ประวัติศาสตร์','สุขศึกษา','ศิลปะ','การงาน',
  'ภาษาอังกฤษ','หน้าที่พลเมือง','การป้องกัน'
];

/**
 * เรียงลำดับ array ตามชื่อวิชามาตรฐาน
 * @param {Array} arr - array ที่ต้องการเรียง
 * @param {string|function} nameKey - ชื่อ property หรือ function ที่ดึงชื่อวิชาจาก element
 * @returns {Array} array ที่เรียงแล้ว (in-place)
 */
function _sortBySubjectName(arr, nameKey) {
  var getName = typeof nameKey === 'function' ? nameKey : function(item) {
    return String(item[nameKey] || '').trim();
  };
  arr.sort(function(a, b) {
    var nameA = getName(a);
    var nameB = getName(b);
    var idxA = -1, idxB = -1;
    for (var si = 0; si < STANDARD_SUBJECT_ORDER.length; si++) {
      if (nameA.indexOf(STANDARD_SUBJECT_ORDER[si]) !== -1) idxA = si;
      if (nameB.indexOf(STANDARD_SUBJECT_ORDER[si]) !== -1) idxB = si;
    }
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return nameA.localeCompare(nameB, 'th');
  });
  return arr;
}