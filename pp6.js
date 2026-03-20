/**
 * ========================================
 * Backend Functions สำหรับระบบรายงานผลการเรียน
 * ========================================
 */


// Cache Settings
const CACHE_EXPIRATION = {
  SHORT: 300,    // 5 นาที - สำหรับข้อมูลที่เปลี่ยนบ่อย
  MEDIUM: 1800,  // 30 นาที - สำหรับข้อมูลที่เปลี่ยนปานกลาง
  LONG: 7200,    // 2 ชั่วโมง - สำหรับข้อมูลที่เปลี่ยนไม่บ่อย
  STATIC: 21600  // 6 ชั่วโมง - สำหรับข้อมูลคงที่
};

// เกณฑ์การแปลงคะแนนเป็น GPA (มาตรฐานไทย)
const GRADE_SCALE = [
  { min: 80, max: 100, gpa: 4.0, letter: 'ดีเยี่ยม' },
  { min: 75, max: 79,  gpa: 3.5, letter: 'ดี' },
  { min: 70, max: 74,  gpa: 3.0, letter: 'ดี' },
  { min: 65, max: 69,  gpa: 2.5, letter: 'พอใช้' },
  { min: 60, max: 64,  gpa: 2.0, letter: 'พอใช้' },
  { min: 55, max: 59,  gpa: 1.5, letter: 'ปรับปรุง' },
  { min: 50, max: 54,  gpa: 1.0, letter: 'ปรับปรุง' },
  { min: 0,  max: 49,  gpa: 0.0, letter: 'ไม่ผ่าน' }
];

// ======== Cache Management Functions ========

/**
 * สร้าง cache key ที่เป็นเอกลักษณ์
 */
function _createCacheKey(prefix, ...args) {
  return `${prefix}_${args.filter(arg => arg !== null && arg !== undefined).join('_')}`;
}

/**
 * ดึงข้อมูลจาก cache
 */
function _getFromCache(key, defaultValue = null) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(key);
    
    if (cachedData) {
      Logger.log(`✅ Cache HIT: ${key}`);
      return JSON.parse(cachedData);
    }
    
    Logger.log(`❌ Cache MISS: ${key}`);
    return defaultValue;
  } catch (e) {
    Logger.log(`⚠️ Cache ERROR: ${key} - ${e.message}`);
    return defaultValue;
  }
}

/**
 * บันทึกข้อมูลลง cache
 */
function _setToCache(key, data, expirationInSeconds = CACHE_EXPIRATION.MEDIUM) {
  try {
    const cache = CacheService.getScriptCache();
    const dataString = JSON.stringify(data);
    
    // ตรวจสอบขนาดข้อมูล (Google Apps Script Cache จำกัดที่ 100KB per item)
    if (dataString.length > 90000) { // เหลือ buffer 10KB
      Logger.log(`⚠️ Cache data too large: ${key} (${dataString.length} chars)`);
      return false;
    }
    
    cache.put(key, dataString, expirationInSeconds);
    Logger.log(`💾 Cache SET: ${key} (expires in ${expirationInSeconds}s)`);
    return true;
  } catch (e) {
    Logger.log(`❌ Cache SET ERROR: ${key} - ${e.message}`);
    return false;
  }
}

/**
 * ล้าง cache ที่เกี่ยวข้องกับข้อมูลที่อัพเดท
 */
function _clearRelatedCache(pattern) {
  try {
    const cache = CacheService.getScriptCache();
    // GAS ไม่มี pattern matching สำหรับ cache keys
    // ต้องใช้ manual cache invalidation
    Logger.log(`🧹 Cache invalidation requested for pattern: ${pattern}`);
    
    // Clear common cache keys ที่เกี่ยวข้อง
    const commonKeys = [
      'subjects_all',
      'warehouse_all',
      'settings_app',
      'assessments_all'
    ];
    
    commonKeys.forEach(key => {
      cache.remove(key);
      Logger.log(`🗑️ Removed cache: ${key}`);
    });
    
  } catch (e) {
    Logger.log(`❌ Cache clear error: ${e.message}`);
  }
}

/**
 * ล้าง cache ทั้งหมด (ใช้ในกรณีฉุกเฉิน) - แก้ไขแล้ว
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    
    // รวบรวม Cache Keys หลักๆ ที่ระบบสร้างขึ้นทั้งหมด
    const keysToRemove = [
      'settings_webapp',
      'sheet_SCORES_WAREHOUSE',
      'sheet_รายวิชา',
      'sheet_การประเมินอ่านคิดเขียน',
      'sheet_การประเมินคุณลักษณะ',
      'sheet_การประเมินกิจกรรมพัฒนาผู้เรียน'
    ];

    // เพิ่ม Keys ของรายวิชาทุกระดับชั้น
    const grades = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'];
    grades.forEach(grade => {
      keysToRemove.push(_createCacheKey('subjects', grade));
    });

    // สั่งลบ Keys ทั้งหมดที่รวบรวมมา
    if (keysToRemove.length > 0) {
      cache.removeAll(keysToRemove);
    }
    
    Logger.log(`🧹 Cleared ${keysToRemove.length} known cache keys.`);
    return `ล้าง Cache หลัก ${keysToRemove.length} รายการเรียบร้อยแล้ว`;

  } catch (e) {
    Logger.log(`❌ Clear all cache error: ${e.message}`);
    throw new Error('ไม่สามารถล้าง Cache ได้: ' + e.message);
  }
}

// ======== Enhanced Utility Functions ========

/**
 * เปิด Spreadsheet หลัก
 */
function _openSpreadsheet() {
  return SS();
}

/**
 * แปลงคะแนนเป็น GPA
 */
function _scoreToGPA(score) {
  const numScore = parseFloat(score);
  if (isNaN(numScore)) return { gpa: 0, letter: 'ไม่ได้คะแนน' };
  
  for (const scale of GRADE_SCALE) {
    if (numScore >= scale.min && numScore <= scale.max) {
      return { gpa: scale.gpa, letter: scale.letter };
    }
  }
  return { gpa: 0, letter: 'ไม่ผ่าน' };
}

/**
 * อ่านข้อมูลจากชีตเป็น Array of Objects (พร้อม Cache)
 */
function _readSheetToObjects(sheetName, useCache = true) {
  const cacheKey = _createCacheKey('sheet', sheetName);
  
  if (useCache) {
    const cachedData = _getFromCache(cacheKey);
    if (cachedData) return cachedData;
  }
  
  try {
    const ss = _openSpreadsheet();
    // ถ้าเป็นชีตรายปี ให้ใช้ S_getYearlySheet (resolve ชื่อ + fallback อัตโนมัติ)
    var sheet;
    if (typeof S_YEARLY_SHEETS !== 'undefined' && S_YEARLY_SHEETS.indexOf(sheetName) !== -1) {
      sheet = S_getYearlySheet(sheetName);
    } else {
      sheet = ss.getSheetByName(sheetName);
    }
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const headers = values[0];
    const data = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = {};
      headers.forEach((header, index) => {
        row[header] = values[i][index] || '';
      });
      data.push(row);
    }
    
    // แคชข้อมูลตามประเภท
    let cacheExpiration = CACHE_EXPIRATION.MEDIUM;
    if (['รายวิชา', 'global_settings'].includes(sheetName)) {
      cacheExpiration = CACHE_EXPIRATION.STATIC; // ข้อมูลคงที่
    } else if (sheetName.includes('การประเมิน')) {
      cacheExpiration = CACHE_EXPIRATION.LONG; // การประเมิน
    }
    
    _setToCache(cacheKey, data, cacheExpiration);
    return data;
  } catch (e) {
    Logger.log(`Error reading sheet ${sheetName}: ${e.message}`);
    return [];
  }
}



// ======== Core Data Functions ========

/**
 * 1. ดึงรายวิชาตามระดับชั้น (พร้อม Cache)
 */
function getSubjectList(grade) {
  const cacheKey = _createCacheKey('subjects', grade);
  const cachedSubjects = _getFromCache(cacheKey);
  
  if (cachedSubjects) {
    return cachedSubjects;
  }

  try {
    const subjects = _readSheetToObjects('รายวิชา', false); // ไม่ใช้ cache ซ้อน
    const gradeSubjects = subjects
      .filter(subject => subject['ชั้น'] === grade)
      .map(subject => ({
        code: subject['รหัสวิชา'],
        name: subject['ชื่อวิชา'],
        hours: parseInt(subject['ชั่วโมง/ปี']) || 0,
        type: subject['ประเภทวิชา'],
        teacher: subject['ครูผู้สอน'],
        midScore: parseInt(subject['คะแนนระหว่างปี']) || 70,
        finalScore: parseInt(subject['คะแนนปลายปี']) || 30,
        total: parseInt(subject['รวม']) || 100
      }));

    // แคชเป็นเวลานาน เพราะหลักสูตรไม่ค่อยเปลี่ยน
    _setToCache(cacheKey, gradeSubjects, CACHE_EXPIRATION.STATIC);
    return gradeSubjects;
  } catch (error) {
    Logger.log('Error in getSubjectList:', error.message);
    throw new Error('ไม่สามารถดึงข้อมูลรายวิชาได้: ' + error.message);
  }
}



// ======== PDF Cache Functions ========

/**
 * สร้าง PDF hash เพื่อใช้เป็น cache key
 */
function _createPDFHash(type, ...params) {
  const hashString = `${type}_${params.join('_')}`;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, hashString)
    .map(byte => (byte + 256).toString(16).slice(-2))
    .join('');
}

/**
 * ตรวจสอบว่า PDF ถูกสร้างไว้แล้วหรือไม่
 */
function _checkPDFCache(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedPDFData = cache.get(cacheKey);
    
    if (cachedPDFData) {
      const pdfInfo = JSON.parse(cachedPDFData);
      
      // ตรวจสอบว่าไฟล์ยังอยู่ใน Drive หรือไม่
      try {
        const file = DriveApp.getFileById(pdfInfo.fileId);
        if (file) {
          Logger.log(`📄 PDF Cache HIT: ${cacheKey}`);
          return {
            exists: true,
            url: file.getUrl(),
            fileName: pdfInfo.fileName,
            createdAt: pdfInfo.createdAt
          };
        }
      } catch (e) {
        // ไฟล์ถูกลบแล้ว ให้ล้าง cache
        cache.remove(cacheKey);
        Logger.log(`🗑️ PDF Cache expired (file deleted): ${cacheKey}`);
      }
    }
    
    return { exists: false };
  } catch (e) {
    Logger.log(`⚠️ PDF Cache check error: ${e.message}`);
    return { exists: false };
  }
}

/**
 * บันทึก PDF ลง cache
 */
function _savePDFCache(cacheKey, fileId, fileName) {
  try {
    const cache = CacheService.getScriptCache();
    const pdfInfo = {
      fileId: fileId,
      fileName: fileName,
      createdAt: new Date().toISOString()
    };
    
    // แคช PDF ไว้ 1 ชั่วโมง
    cache.put(cacheKey, JSON.stringify(pdfInfo), 3600);
    Logger.log(`💾 PDF Cache saved: ${cacheKey}`);
    
    return true;
  } catch (e) {
    Logger.log(`❌ PDF Cache save error: ${e.message}`);
    return false;
  }
}

function getStudentAssessments(studentId) {
  const cacheKey = _createCacheKey('assessments', String(studentId).trim());
  let cachedAssessments = _getFromCache(cacheKey);
  
  // ⭐ Force MISS สำหรับ debug (comment หลังทดสอบ)
  // CacheService.getScriptCache().remove(cacheKey);
  // cachedAssessments = null;
  
  if (cachedAssessments) {
    Logger.log(`Cache HIT for ${studentId}: ${cachedAssessments.reading.result}`);
    return cachedAssessments;
  }

  try {
    const result = {
      reading: { result: 'ไม่พบข้อมูล', score: 0 },
      character: { result: 'ไม่พบข้อมูล', score: 0 },
      activities: { result: 'ไม่ผ่าน' }
    };

    // 1. การประเมินอ่าน คิด เขียน (ปรับใหม่: nums ก่อน if + no error fallback in catch)
    try {
      const readingData = _readSheetToObjects('การประเมินอ่านคิดเขียน');
      Logger.log(`Total records in RTW sheet: ${readingData.length}`);
      
      const readingRecord = readingData.find(row => String(row['รหัสนักเรียน']).trim() === String(studentId).trim());
      
      if (readingRecord) {
        Logger.log(`✅ Record พบสำหรับ ${studentId}: ${JSON.stringify({รหัส: readingRecord['รหัสนักเรียน'], ชื่อ: readingRecord['ชื่อ-นามสกุล']})}`);
        
        const allKeys = Object.keys(readingRecord);
        Logger.log(`All keys in record: [${allKeys.slice(0,15).join(', ')}${allKeys.length >15 ? '...' : ''}]`);
        
        let summaryResult = '';
        const resultKey = allKeys.find(key => {
          const trimmedKey = key.toString().trim();
          return trimmedKey === 'สรุปผลการประเมิน' || trimmedKey.includes('ผลการประเมิน');  // Exact + includes
        });
        if (resultKey) {
          summaryResult = readingRecord[resultKey].toString().trim();
          Logger.log(`Matched header "${resultKey.trim()}" = "${summaryResult}"`);
        } else {
          Logger.log('❌ No match for "สรุปผลการประเมิน" – force calculation');
        }
        
        // ⭐ แก้ไข: คำนวณ nums เสมอ (ก่อน if) สำหรับ avgScore
        const scoreIndices = [4,5,6,7,8,9,10,11,12];  // E-M: ภาษาไทย to ภาษาอังกฤษ? Wait – จาก keys: index 4=ภาษาไทย to 11=ภาษาอังกฤษ, 12=สรุป
        // จาก log keys: index 0=รหัส,1=ชื่อ,2=ชั้น,3=ห้อง,4=ภาษาไทย,5=คณิต,6=วิทย์,7=สังคม,8=สุข,9=ศิลปะ,10=การงาน,11=อังกฤษ,12=สรุป
        // แต่ RTW ควร 9 คะแนน: อ่าน1-3,คิด1-3,เขียน1-3 – keys ผิด? Screenshot แสดง 8 วิชา + สรุป (subject score?)
        const scores = scoreIndices.slice(0,9).map(idx => {  // ถ้า 9 คะแนน = index 4-12
          const val = readingRecord[idx];
          const n = parseFloat(val ? val.toString().trim() : '');
          return Number.isFinite(n) ? n : '';
        });
        
        const nums = scores.filter(x => x !== '' && Number.isFinite(x));
        let calcAvg = nums.length === 9 ? nums.reduce((a, b) => a + b, 0) / 9 : 0;
        
        // ถ้ามีสรุป → ใช้
        if (summaryResult && summaryResult !== '') {
          Logger.log(`ใช้สรุปจากคอลัมน์ M: "${summaryResult}" (calc avg for score: ${calcAvg.toFixed(2)})`);
        } else {
          // Force calc result
          let calcResult = 'ปรับปรุง';
          if (nums.length === 9) {
            if (calcAvg >= 2.5) calcResult = 'ดีเยี่ยม';
            else if (calcAvg >= 2.0) calcResult = 'ดี';
            else if (calcAvg >= 1.0) calcResult = 'ผ่าน';
          }
          summaryResult = calcResult;
          Logger.log(`Force calc: nums.length=${nums.length}, avg=${calcAvg.toFixed(2)} → ${summaryResult}`);
        }
        
        // หา avgKey
        const avgKey = allKeys.find(key => key.toString().trim().includes('คะแนนเฉลี่ย'));
        const avgScore = avgKey ? parseFloat(readingRecord[avgKey]) : calcAvg;
        
        result.reading = {
          result: summaryResult,
          score: avgScore
        };
        Logger.log(`Final reading for ${studentId}: "${summaryResult}" (score: ${avgScore.toFixed(2)})`);
      } else {
        Logger.log(`❌ ไม่พบ record สำหรับ ${studentId} – sample รหัส: ${readingData.slice(0,3).map(r => String(r['รหัสนักเรียน']).trim()).join(', ')}`);
      }
    } catch (e) {
      Logger.log('Cannot read reading assessment: ' + e.message);
      result.reading.result = 'ไม่พบข้อมูล';  // ⭐ เปลี่ยน fallback ไม่มี (error)
    }

    // 2. การประเมินคุณลักษณะ (trim match)
    try {
      const characterData = _readSheetToObjects('การประเมินคุณลักษณะ');
      const characterRecord = characterData.find(row => String(row['รหัสนักเรียน']).trim() === String(studentId).trim());
      if (characterRecord) {
        result.character = {
          result: characterRecord['ผลการประเมิน'] || 'ปรับปรุง',
          score: parseFloat(characterRecord['คะแนนเฉลี่ย']) || 0
        };
      }
    } catch (e) {
      Logger.log('Cannot read character assessment: ' + e.message);
    }

    // 3. การประเมินกิจกรรม (trim match)
    try {
      const activityData = _readSheetToObjects('การประเมินกิจกรรมพัฒนาผู้เรียน', false);
      const activityRecord = activityData.find(row => String(row['รหัสนักเรียน']).trim() === String(studentId).trim());
      if (activityRecord) {
        result.activities = {
          result: activityRecord['รวมกิจกรรม'] || 'ผ่าน',
          แนะแนว: activityRecord['กิจกรรมแนะแนว'] || 'ผ่าน',
          ลูกเสือ: activityRecord['ลูกเสือ_เนตรนารี'] || 'ผ่าน',
          ชุมนุม: activityRecord['ชุมนุม'] || 'ผ่าน',
          สาธารณะ: activityRecord['เพื่อสังคมและสาธารณประโยชน์'] || 'ผ่าน'
        };
      }
    } catch (e) {
      Logger.log('Cannot read activity assessment: ' + e.message);
    }

    _setToCache(cacheKey, result, CACHE_EXPIRATION.LONG);
    Logger.log(`Cache NEW: ${studentId} reading: ${result.reading.result}`);
    return result;
  } catch (error) {
    Logger.log('Error in getStudentAssessments:', error.message);
    throw new Error('ไม่สามารถดึงข้อมูลการประเมินได้: ' + error.message);
  }
}

/**
 * 4. คำนวณ GPA และอันดับของนักเรียน
 */
function calculateGPAAndRank(studentId, grade, classNo) {
  try {
    const warehouse = _readSheetToObjects('SCORES_WAREHOUSE');
    const subjectSheet = _readSheetToObjects('รายวิชา');

    // สร้าง map หน่วยกิต + ประเภทวิชา จากชีตรายวิชา
    const creditMap = {};  // subject_code -> hours (หน่วยกิต)
    const typeMap = {};    // subject_code -> ประเภทวิชา (พื้นฐาน/เพิ่มเติม)
    subjectSheet.forEach(s => {
      const code = String(s['รหัสวิชา'] || '').trim();
      if (code) {
        creditMap[code] = parseFloat(s['ชั่วโมง/ปี']) || 1;
        typeMap[code] = String(s['ประเภทวิชา'] || 'พื้นฐาน').trim();
      }
    });

    // ฟังก์ชันคำนวณ weighted GPA สำหรับนักเรียน 1 คน
    const calcWeightedGPA = (scores) => {
      let totCredits = 0, totGradePoints = 0, basic = 0, additional = 0;
      scores.forEach(subject => {
        const code = String(subject['subject_code'] || subject['subjectCode'] || '').trim();
        const subType = typeMap[code] || 'พื้นฐาน';
        // ข้ามกิจกรรม (ไม่มีหน่วยกิต)
        if (subType === 'กิจกรรม') return;

        const credits = creditMap[code] || 1;
        // ใช้ final_grade ก่อน, ถ้าไม่มีคำนวณจาก average หรือ total
        let gpa = 0;
        const finalGrade = parseFloat(subject['final_grade'] || subject['grade_result'] || '');
        if (!isNaN(finalGrade)) {
          gpa = finalGrade;
        } else {
          const avg = parseFloat(subject['average'] || subject['total'] || 0);
          gpa = _scoreToGPA(avg).gpa;
        }

        totCredits += credits;
        totGradePoints += gpa * credits;
        if (subType === 'เพิ่มเติม') {
          additional += credits;
        } else {
          basic += credits;
        }
      });
      return { totCredits, totGradePoints, basic, additional, gpa: totCredits > 0 ? totGradePoints / totCredits : 0 };
    };

    // ดึงคะแนนของนักเรียนคนนี้
    const studentScores = warehouse.filter(row => {
      const sid = String(row['studentId'] || row['student_id'] || '').trim();
      return sid === String(studentId).trim() && 
        String(row['grade']).trim() === String(grade).trim() && 
        String(row['class_no']).trim() === String(classNo).trim();
    });

    const myGPA = calcWeightedGPA(studentScores);

    // คำนวณอันดับในห้อง (ใช้ weighted GPA เหมือนกัน)
    const classRows = warehouse.filter(row => 
      String(row['grade']).trim() === String(grade).trim() && 
      String(row['class_no']).trim() === String(classNo).trim()
    );
    const studentGroups = {};
    classRows.forEach(row => {
      const sid = String(row['studentId'] || row['student_id'] || '').trim();
      if (!studentGroups[sid]) studentGroups[sid] = [];
      studentGroups[sid].push(row);
    });

    const studentGPAs = Object.entries(studentGroups)
      .map(([sid, scores]) => ({ id: sid, gpa: calcWeightedGPA(scores).gpa }))
      .sort((a, b) => b.gpa - a.gpa);

    const studentRank = studentGPAs.findIndex(s => s.id === String(studentId).trim()) + 1;

    return {
      basicCredits: parseFloat(myGPA.basic.toFixed(2)),
      additionalCredits: parseFloat(myGPA.additional.toFixed(2)),
      totalCredits: parseFloat(myGPA.totCredits.toFixed(2)),
      gpa: parseFloat(myGPA.gpa.toFixed(2)),
      classRank: studentRank,
      totalStudents: studentGPAs.length
    };

  } catch (error) {
    Logger.log('Error in calculateGPAAndRank:', error.message);
    throw new Error('ไม่สามารถคำนวณ GPA และอันดับได้: ' + error.message);
  }
}

/**
 * เรียงลำดับวิชาสำหรับ ปพ.6: พื้นฐาน → เพิ่มเติม → กิจกรรม ตามลำดับมาตรฐาน
 */
function _pp6SortSubjects(subjects) {
  var SUBJECT_ORDER = [
    'ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา',
    'ประวัติศาสตร์','สุขศึกษา','ศิลปะ','การงาน',
    'ภาษาอังกฤษ','หน้าที่พลเมือง','การป้องกัน'
  ];
  var ACTIVITY_ORDER = ['แนะแนว','ลูกเสือ','เนตรนารี','ชุมนุม','สาธารณ','กิจกรรม'];

  function getTypeRank(s) {
    var t = String(s.type || '').trim();
    if (t.includes('กิจกรรม')) return 2;
    if (t.includes('เพิ่มเติม')) return 1;
    return 0;
  }

  function getSubjectIdx(name, orderList) {
    var n = String(name || '').trim();
    for (var i = 0; i < orderList.length; i++) {
      if (n.indexOf(orderList[i]) !== -1) return i;
    }
    return 999;
  }

  subjects.sort(function(a, b) {
    var typeA = getTypeRank(a), typeB = getTypeRank(b);
    if (typeA !== typeB) return typeA - typeB;
    if (typeA === 2) {
      return getSubjectIdx(a.name, ACTIVITY_ORDER) - getSubjectIdx(b.name, ACTIVITY_ORDER);
    }
    return getSubjectIdx(a.name, SUBJECT_ORDER) - getSubjectIdx(b.name, SUBJECT_ORDER);
  });
  return subjects;
}

/**
 * 5. ดึงคะแนนทุกวิชาของนักเรียน (ฉบับ Final - ยึดชีต "รายวิชา" เป็นหลัก)
 */
function getStudentAllSubjectScores(studentId, term = 'both') {
  try {
    // 1. ดึงข้อมูลหลักทั้งหมด
    const studentData = _readSheetToObjects('SCORES_WAREHOUSE').find(row => row['student_id'] == studentId);
    if (!studentData) throw new Error(`ไม่พบข้อมูลนักเรียนรหัส ${studentId}`);
    
    const grade = studentData['grade'];
    const allSubjectsFromSheet = _readSheetToObjects('รายวิชา');
    const studentScoresFromWarehouse = _readSheetToObjects('SCORES_WAREHOUSE').filter(row => row['student_id'] == studentId);
    const assessments = getStudentAssessments(studentId);

    // 2. สร้าง scoreMap
    const scoreMap = new Map();
    studentScoresFromWarehouse.forEach(scoreRow => {
      const average = parseFloat(scoreRow['average']) || 0;
      let displayScore = average;
      if (term === '1') displayScore = parseFloat(scoreRow['term1_total']) || null;
      else if (term === '2') displayScore = parseFloat(scoreRow['term2_total']) || null;
      
      const finalGrade = _scoreToGPA(displayScore);
      const key = scoreRow['subject_code'] || scoreRow['subject_name'];
      
      scoreMap.set(key, {
        score: displayScore,
        grade: finalGrade.gpa > 0 ? finalGrade.gpa.toFixed(1) : (displayScore !== null ? '0.0' : '')
      });
    });
    
    // ⭐ ลบ assessmentMap ออก แล้วใช้ฟังก์ชันนี้แทน
    const getActivityResult = (subjectName) => {
      if (!subjectName) return 'มผ';
      const name = subjectName.toLowerCase();
      
      if (name.includes('แนะแนว')) {
        return assessments.activities.แนะแนว || 'มผ';
      }
      if (name.includes('ลูกเสือ')) {
        return assessments.activities.ลูกเสือ || 'มผ';
      }
      if (name.includes('ชุมนุม')) {
        return assessments.activities.ชุมนุม || 'มผ';
      }
      if (name.includes('สังคม') || name.includes('สาธารณ')) {
        return assessments.activities.สาธารณะ || 'มผ';
      }

      // fallback: ใช้ผลรวมกิจกรรม (รวมกิจกรรม) หากไม่ตรง keyword เฉพาะ
      return assessments.activities.result || 'มผ';
    };

    // 3. สร้างรายการวิชาจาก Master Sheet
    const subjectsForStudentGrade = allSubjectsFromSheet.filter(subject => subject['ชั้น'] === grade);
    
    let finalSubjectList = subjectsForStudentGrade.map(subjectInfo => {
      const subjectName = subjectInfo['ชื่อวิชา'];
      const subjectCode = subjectInfo['รหัสวิชา'];
      const subjectType = subjectInfo['ประเภทวิชา'];
      
      let subjectResult = {
        code: subjectCode,
        name: subjectName,
        type: subjectType,
        hours: parseInt(subjectInfo['ชั่วโมง/ปี']) || 0,
        order: parseInt(subjectInfo['ลำดับ']) || 99,
        score: null,
        grade: ''
      };

      // ตรวจสอบว่าเป็นวิชากิจกรรมหรือไม่ (ตรวจทั้งประเภทวิชาและชื่อวิชา เพื่อรองรับข้อมูลเก่า)
      const activityKeywords = ['แนะแนว', 'ลูกเสือ', 'เนตรนารี', 'ชุมนุม', 'เพื่อสังคม', 'สาธารณประโยชน์', 'กิจกรรมพัฒนา'];
      const typeTrimmed = String(subjectType || '').trim();
      const isActivitySubject = typeTrimmed.includes('กิจกรรม') ||
        activityKeywords.some(kw => String(subjectName || '').includes(kw));
      Logger.log('[PP6 Subject] name="' + subjectName + '" code="' + subjectCode + '" type="' + typeTrimmed + '" isActivity=' + isActivitySubject);
      if (isActivitySubject) {
        subjectResult.isActivity = true;
        subjectResult.grade = getActivityResult(subjectName);
        subjectResult.score = null; // กิจกรรมไม่มีคะแนน
        Logger.log('[PP6 Activity] name="' + subjectName + '" grade="' + subjectResult.grade + '"');
      } else {
        // หาคะแนนจากรหัสวิชาก่อน
        let savedScore = scoreMap.get(subjectCode);
        if (!savedScore) {
          savedScore = scoreMap.get(subjectName);
        }
        
        if (savedScore) {
          subjectResult.score = savedScore.score;
          subjectResult.grade = savedScore.grade;
        }
      }
      return subjectResult;
    });

    // 4. เรียงลำดับ
    finalSubjectList.sort((a, b) => a.order - b.order);
    return finalSubjectList.map((item, index) => ({ ...item, order: index + 1 }));

  } catch (error) {
    Logger.log('Error in getStudentAllSubjectScores:', error.message);
    throw new Error('ไม่สามารถดึงคะแนนทุกวิชาได้: ' + error.message);
  }
}

// ======== PDF Generation Functions ========

/**
 * 6. สร้าง PDF รายงานรายวิชา (พร้อม Cache)
 */
function generateSubjectPDFByTerm(options) {
  try {
    const { subjectName, subjectCode, grade, classNo, term } = options;
    
    // สร้าง cache key สำหรับ PDF นี้
    const pdfHash = _createPDFHash('subject', subjectCode, grade, classNo, term);
    const cacheKey = `pdf_subject_${pdfHash}`;
    
// ⚠️ ปิด PDF Cache ชั่วคราวเพื่อทดสอบ
// const cachedPDF = _checkPDFCache(cacheKey);
// if (cachedPDF.exists) {
//   Logger.log(`📄 Returning cached PP6 PDF for student ${studentId}`);
//   return cachedPDF.url;
// }
    
    Logger.log(`🔨 Generating new PDF: รายงานรายวิชา ${subjectName}`);
    
    // 1. ดึงข้อมูลคะแนนของวิชานี้
    const warehouse = _readSheetToObjects('SCORES_WAREHOUSE');
    const subjectData = warehouse.filter(row =>
      row['subject_code'] === subjectCode &&
      row['grade'] === grade &&
      row['class_no'] == classNo
    );

    if (subjectData.length === 0) {
      throw new Error(`ไม่พบข้อมูลคะแนนสำหรับวิชา ${subjectName} ชั้น ${grade} ห้อง ${classNo}`);
    }

    // 2. ดึงการตั้งค่าระบบ
    const settings = getWebAppSettingsWithCache();
    
    // 3. สร้าง HTML Template
    const html = _createSubjectReportHTML({
      schoolName: settings.schoolName,
      logoDataUrl: settings.logoDataUrl,
      subjectName,
      subjectCode,
      grade,
      classNo,
      term,
      students: subjectData
    });

    // --- ส่วนที่แก้ไข ---
    // 4. แปลงเป็น PDF และบันทึก (แก้ไขแล้ว)
    const fileName = `รายงานรายวิชา_${subjectName}_${grade}-${classNo}_${term === 'both' ? 'ทั้งปี' : 'ภาค' + term}_${new Date().toISOString().split('T')[0]}.pdf`;
    const pdfBlob = HtmlService.createHtmlOutput(html)
                             .getAs('application/pdf') // บังคับให้เป็นไฟล์ PDF ชัดเจน
                             .setName(fileName);
    
    // ใช้ getOutputFolder_() จาก code.gs
    const folder = getOutputFolder_();
    const file = folder.createFile(pdfBlob);
    // --- สิ้นสุดส่วนที่แก้ไข ---
    
    // 5. บันทึกลง PDF Cache
    _savePDFCache(cacheKey, file.getId(), fileName);
    
    Logger.log(`✅ PDF generated successfully: ${fileName}`);
    return file.getUrl();

  } catch (error) {
    Logger.log('Error in generateSubjectPDFByTerm:', error.message);
    throw new Error('ไม่สามารถสร้างรายงานรายวิชาได้: ' + error.message);
  }
}

/**
 * ================================================================
 * == โค้ดส่วนดึงข้อมูลและสร้าง PDF ปพ.6 (ฉบับสมบูรณ์สุดท้าย) ==
 * ================================================================
 */

/**
 * 1. ดึงการตั้งค่าเว็บแอปจากชีต global_settings (อ่านแบบ Key-Value)
 */

/**
 * 2. Helper: แปลง File ID ของรูปภาพเป็น Data URL (Base64)
 */
function _getLogoDataUrl(fileId) {
  if (!fileId) return null;
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const contentType = blob.getContentType();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    return `data:${contentType};base64,${base64Data}`;
  } catch (e) {
    Logger.log('ไม่สามารถโหลดไฟล์โลโก้ได้: ' + e.message);
    return null;
  }
}

/**
 * 3. ดึงรายชื่อนักเรียนสำหรับ ปพ.6 (จากชีต Students)
 */
function getStudentListForPp6(grade, classNo) {
  const cacheKey = _createCacheKey('students_master', grade, classNo);
  const cachedStudents = _getFromCache(cacheKey);
  if (cachedStudents) return cachedStudents;

  try {
    // โหลดทะเบียนนักเรียนจากชีต Students
    const allStudentsData = _readSheetToObjects('Students'); 
    const studentNameMap = new Map();
    allStudentsData.forEach(student => {
      const fullName = `${student.title || ''}${student.firstname || ''} ${student.lastname || ''}`;
      studentNameMap.set(student.student_id.toString().trim(), fullName.trim());
    });

    // ค้นหารหัสนักเรียนในชั้นเรียนจาก SCORES_WAREHOUSE
    const warehouse = _readSheetToObjects('SCORES_WAREHOUSE');
    const studentIdsInClass = new Set();
    warehouse
      .filter(row => row['grade'] === grade && row['class_no'] == classNo)
      .forEach(row => {
        if (row['student_id']) studentIdsInClass.add(row['student_id'].toString().trim());
      });

    // สร้างผลลัพธ์โดยการนำรหัสไปหาชื่อในทะเบียน
    const result = Array.from(studentIdsInClass).map(id => ({
      id: id,
      name: studentNameMap.get(id) || `นักเรียนรหัส ${id} (ไม่มีชื่อในทะเบียน)`
    }));

    // --- 👇 แก้ไข/เพิ่มโค้ดเรียงลำดับให้ใช้ localeCompare มาตรฐาน ---
    result.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, {numeric: true}));
    // --- 👆 สิ้นสุดโค้ดที่แก้ไข ---

    _setToCache(cacheKey, result, CACHE_EXPIRATION.LONG);
    return result;
  } catch (error) {
    Logger.log('Error in getStudentListForPp6:', error.message);
    throw new Error('ไม่สามารถดึงข้อมูลรายชื่อนักเรียนได้: ' + error.message);
  }
}


/**
 * ฟังก์ชันดึงชื่อครูประจำชั้นจากชีต HomeroomTeachers
 * @param {string} grade - ระดับชั้น เช่น 'ป.1', 'ป.2'
 * @param {string} classNo - ห้อง เช่น '1', '2'
 * @returns {string} ชื่อครูประจำชั้น
 */

/**
 * สร้าง PDF จาก Google Docs (ฟอนต์ Sarabun เหมือนรายงานหน้าเดียว)
 */
function _pp6_buildDocPdf(data) {
  var F = OPR_FONT; // 'Sarabun'
  var CENTER = DocumentApp.HorizontalAlignment.CENTER;
  var LEFT = DocumentApp.HorizontalAlignment.LEFT;

  var doc = DocumentApp.create('ปพ6_' + data.studentId + '_' + Date.now());
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(22);
  body.setMarginBottom(14);
  body.setMarginLeft(50);
  body.setMarginRight(28);
  body.setPageWidth(595);
  body.setPageHeight(842);

  var addLine = function(txt, sz, bold, sa, align) {
    var p = body.appendParagraph(txt);
    p.setAlignment(align || CENTER);
    p.editAsText().setBold(bold).setFontSize(sz).setFontFamily(F);
    p.setSpacingAfter(sa !== undefined ? sa : 0).setSpacingBefore(0);
    return p;
  };

  // === HEADER: โลโก้ + ชื่อโรงเรียน ===
  if (data.logoFileId) {
    try {
      var lp = body.appendParagraph('');
      lp.setAlignment(CENTER);
      var img = lp.appendInlineImage(DriveApp.getFileById(data.logoFileId).getBlob());
      img.setHeight(45); img.setWidth(45);
      lp.setSpacingAfter(0).setSpacingBefore(0);
    } catch(e) { Logger.log('PP6 Logo: ' + e.message); }
  }

  addLine('แบบรายงานผลพัฒนาคุณภาพผู้เรียนรายบุคคล', 15, true, 0);
  addLine('ปีการศึกษา ' + data.academicYear, 13, false, 0);
  addLine(data.schoolName, 14, true, 0);
  if (data.schoolAddress) addLine(data.schoolAddress, 13, false, 1);

  // === ข้อมูลนักเรียน ===
  var infoText = 'ชื่อ - นามสกุล: ' + data.studentName + '     รหัสประจำตัว: ' + data.studentId + '     ชั้น: ' + data.grade + '/' + data.classNo;
  addLine(infoText, 13, false, 2, LEFT);

  // === ตารางรายวิชา (7 คอลัมน์) ===
  var HDR_BG = '#E8E8E8';
  var FS_H = 11;
  var FS_D = 12;
  var T = body.appendTable();
  T.setBorderWidth(0.5);
  T.setBorderColor('#000000');

  var W = [25, 52, 170, 55, 55, 70, 90];
  var headers = ['ลำดับ', 'รหัสวิชา', 'ชื่อวิชา', 'ประเภท', 'จำนวน\nชั่วโมง', 'คะแนน\nที่ได้', 'ระดับผล\nการเรียน'];
  var hr = T.appendTableRow();
  headers.forEach(function(h, i) {
    _opr_cell(hr.appendTableCell(h), W[i], FS_H, true, HDR_BG);
  });

  var fmtAct = function(txt) {
    if (!txt) return 'มผ';
    var s = String(txt).trim();
    if (s === 'ผ่าน' || s === 'ผ') return 'ผ่าน';
    if (s === 'ไม่ผ่าน' || s === 'มผ') return 'มผ';
    if (!isNaN(parseFloat(s))) return 'ผ่าน';
    return s;
  };

  data.subjects.forEach(function(sub, idx) {
    var isAct = sub.isActivity || String(sub.type || '').trim().indexOf('กิจกรรม') !== -1;
    var row = T.appendTableRow();
    var vals = [
      String(idx + 1),
      sub.code || '',
      sub.name || '',
      sub.type || '',
      String(sub.hours || ''),
      sub.score ? String(sub.score) : '',
      isAct ? fmtAct(sub.grade) : (sub.grade || '')
    ];
    vals.forEach(function(val, i) {
      _opr_cell(row.appendTableCell(val), W[i], FS_D, (i >= 5), null, (i === 2 ? LEFT : null));
    });
  });

  // === สรุปผลการประเมิน (หัว 1 คอลัมน์ + ข้อมูล 4 คอลัมน์) ===
  var shT = body.appendTable();
  shT.setBorderWidth(0.5);
  shT.setBorderColor('#000000');
  var shR = shT.appendTableRow();
  _opr_cell(shR.appendTableCell('สรุปผลการประเมิน'), 517, 13, true, HDR_BG);

  var SW = [170, 60, 220, 67];
  var sT = body.appendTable();
  sT.setBorderWidth(0.5);
  sT.setBorderColor('#000000');

  var charRes = (data.assessments.character && data.assessments.character.result) || '-';
  var readRes = (data.assessments.reading && data.assessments.reading.result) || '-';
  var actRes = fmtAct(data.assessments.activities && data.assessments.activities.result);

  var sumRows = [
    ['จำนวนหน่วยกิต/น้ำหนักวิชาพื้นฐาน:', String(data.gpaInfo.basicCredits || '0'), 'ผลการประเมินคุณลักษณะอันพึงประสงค์:', charRes],
    ['จำนวนหน่วยกิต/น้ำหนักวิชาเพิ่มเติม:', String(data.gpaInfo.additionalCredits || '0'), 'ผลการประเมินการอ่าน คิดวิเคราะห์ และเขียน:', readRes],
    ['รวมจำนวนหน่วยกิต/น้ำหนัก:', String(data.gpaInfo.totalCredits || '0'), 'ผลการประเมินกิจกรรมพัฒนาผู้เรียน:', actRes],
    ['ระดับผลการเรียนเฉลี่ย (GPA):', String(data.gpaInfo.gpa || '0.00'),
      data.showRank ? ('ได้ลำดับที่ จากนักเรียน ' + (data.gpaInfo.totalStudents || '-') + ' คน') : '',
      data.showRank ? String(data.gpaInfo.classRank || '-') : '']
  ];
  sumRows.forEach(function(rd) {
    var r = sT.appendTableRow();
    rd.forEach(function(val, i) {
      var align = (i === 0 || i === 2) ? LEFT : null; // label=LEFT, value=CENTER
      _opr_cell(r.appendTableCell(val), SW[i], 11, (i === 1 || i === 3), null, align);
    });
  });

  // === ความคิดเห็นของครูประจำชั้น ===
  var cmtH = body.appendParagraph('ความคิดเห็นของครูประจำชั้น / ครูที่ปรึกษา:');
  cmtH.editAsText().setBold(true).setFontSize(12).setFontFamily(F);
  cmtH.setSpacingAfter(1).setSpacingBefore(6);

  var cmtText = data.teacherComment && data.teacherComment !== '-' ? data.teacherComment : '..........................................................................................................................................................................................';
  var cmtP = body.appendParagraph(cmtText);
  cmtP.editAsText().setBold(false).setFontSize(12).setFontFamily(F);
  cmtP.setSpacingAfter(4).setSpacingBefore(0);

  // === ลายเซ็น 3 คอลัมน์ (ไม่มีเส้นกรอบ) ===
  var sigT = body.appendTable();
  sigT.setBorderWidth(0);
  var SgW = [173, 172, 172];

  var tName = (data.homeroomTeacher && data.homeroomTeacher !== '...') ? data.homeroomTeacher : '';
  var nm1 = tName ? '(' + tName + ')' : '(.................................)';
  var nm2 = data.principalName ? '(' + data.principalName + ')' : '(.................................)';

  var signData = [
    ['ลงชื่อ..................................', nm1, 'ครูที่ปรึกษา'],
    ['ลงชื่อ..................................', nm2, 'ผู้บริหารสถานศึกษา'],
    ['ลงชื่อ..................................', '(.................................)', 'ผู้ปกครอง']
  ];
  var sr = sigT.appendTableRow();
  signData.forEach(function(lines, ci) {
    var c = sr.appendTableCell(lines[0]);
    c.setWidth(SgW[ci]);
    c.setPaddingTop(6).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);
    var p0 = c.getChild(0).asParagraph();
    p0.setAlignment(CENTER);
    p0.setSpacingAfter(0).setSpacingBefore(0);
    p0.editAsText().setFontSize(12).setFontFamily(F);
    var p1 = c.appendParagraph(lines[1]);
    p1.setAlignment(CENTER);
    p1.setSpacingAfter(0).setSpacingBefore(0);
    p1.editAsText().setFontSize(12).setFontFamily(F);
    var p2 = c.appendParagraph(lines[2]);
    p2.setAlignment(CENTER);
    p2.setSpacingAfter(0).setSpacingBefore(0);
    p2.editAsText().setFontSize(12).setFontFamily(F);
  });

  // === Export PDF ===
  doc.saveAndClose();
  var docId = doc.getId();
  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName(data.fileName);
  docFile.setTrashed(true);

  return pdfBlob;
}

/**
 * ฟังก์ชันสร้าง PDF รายงานรายบุคคล (ปพ.6) - ใช้ Google Docs (ฟอนต์ Sarabun)
 */
function generatePp6PDFComplete(studentId, term = 'both', showRank = true) {
  try {
    try { DriveApp.getRootFolder().getName(); } catch (e) {
      throw new Error(
        'ระบบยังไม่ได้รับอนุญาตให้เข้าถึง Google Drive สำหรับการสร้าง PDF\n' +
        'วิธีแก้:\n' +
        '1) เปิดโปรเจกต์ Apps Script ด้วยบัญชีเจ้าของ (Deploying account)\n' +
        '2) Run ฟังก์ชัน testDrivePermission() 1 ครั้งเพื่อ authorize\n' +
        '3) Deploy > Manage deployments > Edit > New version > Deploy\n' +
        'หมายเหตุ: ถ้า Deploy ตั้งค่า Execute as = User accessing ให้เปลี่ยนเป็น Me'
      );
    }

    const pdfHash = _createPDFHash('pp6_final_v3', studentId, term);
    const cacheKey = `pdf_pp6_${pdfHash}`;
    
    // ปิด Cache ชั่วคราวเพื่อทดสอบ
    Logger.log(`🔨 Generating new PP6 PDF for student ${studentId}, term: ${term}`);
    
    const settings = getWebAppSettings();
    if (!settings['ชื่อโรงเรียน']) {
      throw new Error('ไม่พบการตั้งค่าโรงเรียน');
    }

    const allStudentsData = _readSheetToObjects('Students');
    const studentMasterData = allStudentsData.find(row => String(row.student_id).trim() === String(studentId).trim());
    
    if (!studentMasterData) {
      throw new Error(`ไม่พบข้อมูลนักเรียนรหัส ${studentId} ในชีต Students`);
    }

    const studentFullName = `${studentMasterData.title || ''}${studentMasterData.firstname || ''} ${studentMasterData.lastname || ''}`.trim();

    const warehouse = _readSheetToObjects('SCORES_WAREHOUSE');
    const studentScoreData = warehouse.find(row => String(row['student_id']).trim() === String(studentId).trim());
    if (!studentScoreData) {
      throw new Error(`ไม่พบข้อมูลคะแนนของนักเรียนรหัส ${studentId} ใน SCORES_WAREHOUSE`);
    }

    const grade = studentScoreData['grade'];
    const classNo = studentScoreData['class_no'];
    if (!grade || !classNo) {
      throw new Error(`ข้อมูล grade (${grade}) หรือ class_no (${classNo}) ไม่ครบถ้วน`);
    }

    const subjects = getStudentAllSubjectScores(studentId, term);
    if (!subjects || subjects.length === 0) {
      throw new Error(`ไม่พบข้อมูลวิชาสำหรับนักเรียน ${studentId}`);
    }
    _pp6SortSubjects(subjects);

    const gpaInfo = calculateGPAAndRank(studentId, grade, classNo);
    const assessments = getStudentAssessments(studentId);

    // ✅ ใช้ getStudentInfo_ เหมือนรายงานหน้าเดียว เพื่อให้ grade/classNo format ตรงกับ HomeroomTeachers
    const studentInfo = typeof getStudentInfo_ === 'function' ? getStudentInfo_(studentId) : null;
    const teacherGrade = studentInfo ? studentInfo.grade : grade;
    const teacherClassNo = studentInfo ? String(studentInfo.classNo) : String(classNo);
    const homeroomTeacher = getHomeroomTeacher(teacherGrade, teacherClassNo);
    Logger.log('🔍 PP6 teacher lookup: grade=' + teacherGrade + ', classNo=' + teacherClassNo + ', result=' + homeroomTeacher);
    const teacherComment = typeof getTeacherComment_ === 'function' ? getTeacherComment_(studentId) : '';

    const fileName = `ปพ6_${studentFullName || studentId}_${studentId}_${term}.pdf`;

    const pdfBlob = _pp6_buildDocPdf({
      schoolName: settings['ชื่อโรงเรียน'],
      schoolAddress: settings['ที่อยู่โรงเรียน'],
      logoFileId: settings['logoFileId'],
      studentId: studentId,
      studentName: studentFullName,
      grade: grade,
      classNo: classNo,
      term: term,
      subjects: subjects,
      gpaInfo: gpaInfo,
      assessments: assessments,
      principalName: settings['ชื่อผู้อำนวยการ'],
      homeroomTeacher: homeroomTeacher,
      teacherComment: teacherComment,
      academicYear: settings['ปีการศึกษา'] || new Date().getFullYear() + 543,
      fileName: fileName,
      showRank: showRank
    });
    
    const folder = getOutputFolder_();
    const file = folder.createFile(pdfBlob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(shareErr) { Logger.log('setSharing skipped: ' + shareErr.message); }
    
    _savePDFCache(cacheKey, file.getId(), fileName);
    Logger.log(`✅ PP6 PDF generated: ${fileName}`);

    const id = file.getId();
    const previewUrl  = `https://drive.google.com/file/d/${id}/preview`;
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${id}`;

    return { previewUrl, downloadUrl, fileId: id, name: fileName };
    
  } catch (error) {
    Logger.log(`❌ Error in generatePp6PDFComplete for student ${studentId}: ${error.message}`);
    throw new Error(`ไม่สามารถสร้างรายงาน ปพ.6: ${error.message}`);
  }
}

function generatePp6PDFCompleteNoDrive(studentId, term = 'both', showRank = true) {
  try {
    const settings = getWebAppSettings();
    if (!settings['ชื่อโรงเรียน']) {
      throw new Error('ไม่พบการตั้งค่าโรงเรียน');
    }

    const allStudentsData = _readSheetToObjects('Students');
    const studentMasterData = allStudentsData.find(row => String(row.student_id).trim() === String(studentId).trim());
    if (!studentMasterData) {
      throw new Error(`ไม่พบข้อมูลนักเรียนรหัส ${studentId} ในชีต Students`);
    }

    const studentFullName = `${studentMasterData.title || ''}${studentMasterData.firstname || ''} ${studentMasterData.lastname || ''}`.trim();

    const warehouse = _readSheetToObjects('SCORES_WAREHOUSE');
    const studentScoreData = warehouse.find(row => String(row['student_id']).trim() === String(studentId).trim());
    if (!studentScoreData) {
      throw new Error(`ไม่พบข้อมูลคะแนนของนักเรียนรหัส ${studentId} ใน SCORES_WAREHOUSE`);
    }

    const grade = studentScoreData['grade'];
    const classNo = studentScoreData['class_no'];

    const subjects = getStudentAllSubjectScores(studentId, term);
    _pp6SortSubjects(subjects);
    const gpaInfo = calculateGPAAndRank(studentId, grade, classNo);
    const assessments = getStudentAssessments(studentId);

    // ✅ ใช้ getStudentInfo_ เหมือนรายงานหน้าเดียว เพื่อให้ grade/classNo format ตรงกับ HomeroomTeachers
    const studentInfo = typeof getStudentInfo_ === 'function' ? getStudentInfo_(studentId) : null;
    const teacherGrade = studentInfo ? studentInfo.grade : grade;
    const teacherClassNo = studentInfo ? String(studentInfo.classNo) : String(classNo);
    const homeroomTeacher = getHomeroomTeacher(teacherGrade, teacherClassNo);
    Logger.log('🔍 PP6 NoDrive teacher lookup: grade=' + teacherGrade + ', classNo=' + teacherClassNo + ', result=' + homeroomTeacher);
    const teacherComment = typeof getTeacherComment_ === 'function' ? getTeacherComment_(studentId) : '';

    const fileName = `ปพ6_${studentFullName || studentId}_${studentId}_${term}.pdf`;

    const pdfBlob = _pp6_buildDocPdf({
      schoolName: settings['ชื่อโรงเรียน'],
      schoolAddress: settings['ที่อยู่โรงเรียน'],
      logoFileId: settings['logoFileId'],
      studentId: studentId,
      studentName: studentFullName,
      grade: grade,
      classNo: classNo,
      term: term,
      subjects: subjects,
      gpaInfo: gpaInfo,
      assessments: assessments,
      principalName: settings['ชื่อผู้อำนวยการ'],
      homeroomTeacher: homeroomTeacher,
      teacherComment: teacherComment,
      academicYear: settings['ปีการศึกษา'] || new Date().getFullYear() + 543,
      fileName: fileName,
      showRank: showRank
    });

    const base64 = Utilities.base64Encode(pdfBlob.getBytes());
    return { fileName: fileName, mimeType: 'application/pdf', base64: base64 };
  } catch (error) {
    Logger.log(`❌ Error in generatePp6PDFCompleteNoDrive for student ${studentId}: ${error.message}`);
    throw new Error(`ไม่สามารถสร้างรายงาน ปพ.6: ${error.message}`);
  }
}

/**
 * เทมเพลต HTML ที่ปรับปรุงให้แสดงชื่อครูประจำชั้น และชื่อวิชาชิดซ้าย
 */
function _createPp6ReportHTML(data) {
  const { 
    schoolName, schoolAddress, logoDataUrl, studentId, studentName, grade, classNo, 
    term, subjects, gpaInfo, assessments, principalName, homeroomTeacher, teacherComment, academicYear 
  } = data;
  
  const termText = term === 'both' ? 'ภาคเรียนที่ 1 และ 2' : `ภาคเรียนที่ ${term}`;
  
const formatActivityResult = (resultText) => {
  if (!resultText) return 'มผ';
  const text = String(resultText).trim();
  if (text === 'ผ่าน' || text === 'ผ') return 'ผ';
  if (text === 'ไม่ผ่าน' || text === 'มผ') return 'มผ';
  // ถ้าเป็นตัวเลข (0.0, 3.5 ฯลฯ) → กิจกรรมไม่มีระดับคะแนนเป็นตัวเลข → แสดง ผ แทน
  if (!isNaN(parseFloat(text))) return 'ผ';
  return text;
};
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4; margin: 10mm 12mm; }
    body { font-family: 'Sarabun', 'TH SarabunPSK', sans-serif; font-size: 12px; line-height: 1.2; color: #000; margin: 0; padding: 0; }
    .header { text-align: center; margin-bottom: 8px; }
    .logo { width: 45px; height: 45px; margin-bottom: 5px; }
    .report-title { font-size: 15px; font-weight: bold; margin-bottom: 2px;}
    .school-info { font-size: 13px; margin-bottom: 1px;}
    .student-info { margin: 8px 0; padding: 4px 0; font-size: 13px; }
    .main-table { width: 100%; border-collapse: collapse; margin-bottom: 6px; font-size: 12px; table-layout: fixed; }
    .main-table th, .main-table td { border: 1px solid #000; padding: 5px 6px; text-align: center; word-wrap: break-word; }
    .main-table th { background: #f0f0f0; font-weight: bold; }
    .main-table td:nth-child(3) { text-align: left; padding-left: 6px; }
    .summary-section { margin: 8px 0; }
    .summary-title { font-weight: bold; font-size: 13px; margin-bottom: 4px; text-align: center; background: #e0e0e0; padding: 3px; border: 1px solid #ccc; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 12px; }
    .summary-table td { padding: 2px 4px; vertical-align: top; }
    .summary-table td.value { font-weight: bold; }
    .summary-table td:nth-child(1) { width: 28%; }
    .summary-table td:nth-child(2) { width: 15%; }
    .summary-table td:nth-child(3) { width: 39%; }
    .summary-table td:nth-child(4) { width: 18%; }
    .signatures { margin-top: 10px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; text-align: center; }
    .signature-box { padding: 4px; }
    .signature-line { border-top: 1px solid #000; margin-top: 20px; padding-top: 3px; }
    .footer { margin-top: 8px; text-align: right; font-size: 10px; }
  </style>
</head>
<body>
  <div class="header">
    ${logoDataUrl ? `<img src="${logoDataUrl}" class="logo" alt="โลโก้โรงเรียน">` : ''}
    <div class="report-title">แบบรายงานผลพัฒนาคุณภาพผู้เรียนรายบุคคล</div>
    <div class="school-info">ปีการศึกษา ${academicYear} </div>
    <div class="school-info"><strong>${schoolName}</strong></div>
    <div class="school-info">${schoolAddress || ''}</div>
  </div>

  <div class="student-info">
    <strong>ชื่อ - นามสกุล:</strong> ${studentName} &nbsp;&nbsp; <strong>รหัสประจำตัว:</strong> ${studentId} &nbsp;&nbsp; <strong>ชั้น:</strong> ${grade}/${classNo}
  </div>

  <table class="main-table">
    <thead>
      <tr>
        <th style="width: 6%">ลำดับ</th><th style="width: 10%">รหัสวิชา</th><th style="width: 30%">ชื่อวิชา</th>
        <th style="width: 12%">ประเภท</th><th style="width: 12%">จำนวนชั่วโมง</th><th style="width: 14%">คะแนนที่ได้</th><th style="width: 16%">ระดับผลการเรียน</th>
      </tr>
    </thead>
    <tbody>
      ${subjects.map((subject, idx) => `
        <tr>
          <td>${idx + 1}</td>
          <td>${subject.code}</td>
          <td>${subject.name}</td>
          <td>${subject.type}</td>
          <td>${subject.hours}</td>
          <td><strong>${subject.score ? subject.score : ''}</strong></td>
          <td><strong>${(subject.isActivity || String(subject.type || '').trim().includes('กิจกรรม')) ? formatActivityResult(subject.grade) : (subject.grade || '')}</strong></td>
        </tr>
      `).join('')}
    </tbody>
  </table>

  <div class="summary-section">
    <div class="summary-title">สรุปผลการประเมิน</div>
    <table class="summary-table">
        <tr>
            <td class="label">จำนวนหน่วยกิต/น้ำหนักวิชาพื้นฐาน:</td>
            <td class="value">${gpaInfo.basicCredits || '0.00'}</td>
            <td class="label" style="padding-left: 8px;">ผลการประเมินคุณลักษณะอันพึงประสงค์:</td>
            <td class="value">${assessments.character?.result || 'ไม่พบข้อมูล'}</td>
        </tr>
        <tr>
            <td class="label">จำนวนหน่วยกิต/น้ำหนักวิชาเพิ่มเติม:</td>
            <td class="value">${gpaInfo.additionalCredits || '0.00'}</td>
            <td class="label" style="padding-left: 8px;">ผลการประเมินการอ่าน คิดวิเคราะห์ และเขียน:</td>
            <td class="value">${assessments.reading?.result || 'ไม่พบข้อมูล'}</td>
        </tr>
        <tr>
            <td class="label">รวมจำนวนหน่วยกิต/น้ำหนัก:</td>
            <td class="value">${gpaInfo.totalCredits || '0.00'}</td>
            <td class="label" style="padding-left: 8px;">ผลการประเมินกิจกรรมพัฒนาผู้เรียน:</td>
            <td class="value">${formatActivityResult(assessments.activities?.result)}</td>
        </tr>
                <tr>
            <td class="label">ระดับผลการเรียนเฉลี่ย (GPA):</td>
            <td class="value">${gpaInfo.gpa || '0.00'}</td>
            <td class="label" style="padding-left: 8px;"></td>
            <td class="value"></td>
        </tr>
    </table>
  </div>

  <div class="comment-section" style="margin: 8px 0; border: 1px solid #ccc; padding: 6px 8px;">
    <div style="font-weight: bold; font-size: 12px; margin-bottom: 3px;">ความคิดเห็นของครูประจำชั้น / ครูที่ปรึกษา:</div>
    <div style="min-height: 30px; font-size: 12px; line-height: 1.4; padding: 2px;">${teacherComment && teacherComment !== '-' ? teacherComment : '......................................................................................................................................................................................................................................................'}</div>
  </div>

  <div class="signatures">
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ${homeroomTeacher && homeroomTeacher !== '...' ? homeroomTeacher : '.................................'} )</div>
       <div>ครูที่ปรึกษา</div>
     </div>
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ${principalName || '.................................'} )</div>
       <div>ผู้บริหารสถานศึกษา</div>
     </div>
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ................................. )</div>
       <div>ผู้ปกครอง</div>
     </div>
  </div>

  <div class="footer">
    สร้างเมื่อ: ${new Date().toLocaleDateString('th-TH')} ${new Date().toLocaleTimeString('th-TH')}
  </div>
</body>
</html>`;
}
// ======== Cache Management API ========

/**
 * API สำหรับจัดการ Cache (สำหรับ Web App)
 */
function manageCacheAPI(action, params = {}) {
  try {
    switch (action) {
      case 'warmup':
        return { success: true, result: warmupCache() };
        
      case 'clear':
        const sheetName = params.sheetName || null;
        return { success: true, result: invalidateCacheAfterDataUpdate(sheetName) };
        
      case 'clearAll':
        return { success: true, result: clearAllCache() };
        
      case 'stats':
        return { success: true, result: getCacheStatistics() };
        
      case 'performance':
        const { functionName, args = [] } = params;
        return { success: true, result: measureFunctionPerformance(functionName, ...args) };
        
      default:
        throw new Error(`ไม่รองรับคำสั่ง cache: ${action}`);
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 🧹 ลบชีต backup warehouse ทั้งหมด
 */
function deleteAllWarehouseBackups() {
  try {
    const ss = _openSpreadsheet();
    const sheets = ss.getSheets();
    let deletedCount = 0;
    const deletedNames = [];
    
    Logger.log('🔍 กำลังค้นหาชีต backup warehouse...');
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      
      // ลบชีตที่ชื่อขึ้นต้นด้วย "BACKUP_WAREHOUSE"
      if (name.startsWith('BACKUP_WAREHOUSE')) {
        try {
          ss.deleteSheet(sheet);
          deletedNames.push(name);
          deletedCount++;
          Logger.log(`🗑️ ลบแล้ว: ${name}`);
        } catch (e) {
          Logger.log(`❌ ไม่สามารถลบ: ${name} - ${e.message}`);
        }
      }
    });
    
    Logger.log('='.repeat(50));
    Logger.log(`✅ ลบชีต backup เสร็จสิ้น`);
    Logger.log(`📊 จำนวนที่ลบ: ${deletedCount} ชีต`);
    
    if (deletedCount > 0) {
      Logger.log('รายการที่ลบ:');
      deletedNames.forEach((name, i) => Logger.log(`  ${i+1}. ${name}`));
      return `ลบชีต backup warehouse สำเร็จ ${deletedCount} ชีต:\n${deletedNames.join('\n')}`;
    } else {
      return 'ไม่พบชีต backup warehouse ที่ต้องลบ';
    }
    
  } catch (error) {
    Logger.log(`❌ Error in deleteAllWarehouseBackups: ${error.message}`);
    throw new Error('ไม่สามารถลบชีต backup ได้: ' + error.message);
  }
}
