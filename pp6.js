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
    const sheet = ss.getSheetByName(sheetName);
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
    console.error(`Error reading sheet ${sheetName}: ${e.message}`);
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
    console.error('Error in getSubjectList:', error.message);
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
      console.warn('Cannot read reading assessment: ' + e.message);
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
      console.warn('Cannot read character assessment: ' + e.message);
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
      console.warn('Cannot read activity assessment: ' + e.message);
    }

    _setToCache(cacheKey, result, CACHE_EXPIRATION.LONG);
    Logger.log(`Cache NEW: ${studentId} reading: ${result.reading.result}`);
    return result;
  } catch (error) {
    console.error('Error in getStudentAssessments:', error.message);
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
    console.error('Error in calculateGPAAndRank:', error.message);
    throw new Error('ไม่สามารถคำนวณ GPA และอันดับได้: ' + error.message);
  }
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
      
      return 'มผ';
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

      // ตรวจสอบว่าเป็นวิชากิจกรรมหรือไม่
      if (subjectType === 'กิจกรรม') {
        subjectResult.code = '';
        subjectResult.grade = getActivityResult(subjectName); // ⭐ ใช้ฟังก์ชันใหม่
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
    console.error('Error in getStudentAllSubjectScores:', error.message);
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
    console.error('Error in generateSubjectPDFByTerm:', error.message);
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
    console.error('ไม่สามารถโหลดไฟล์โลโก้ได้: ' + e.message);
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
    console.error('Error in getStudentListForPp6:', error.message);
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
 * ฟังก์ชันสร้าง PDF รายงานรายบุคคล (ปพ.6) - เวอร์ชันปรับปรุงให้รองรับครูประจำชั้น
 */
function generatePp6PDFComplete(studentId, term = 'both') {
  try {
    const pdfHash = _createPDFHash('pp6_final_v3', studentId, term);
    const cacheKey = `pdf_pp6_${pdfHash}`;
    
    // ปิด Cache ชั่วคราวเพื่อทดสอบ
    Logger.log(`🔨 Generating new PP6 PDF for student ${studentId}, term: ${term}`);
    
    const settings = getWebAppSettings();
    if (!settings['ชื่อโรงเรียน']) {
      throw new Error('ไม่พบการตั้งค่าโรงเรียน');
    }

    const logoDataUrl = _getLogoDataUrl(settings['logoFileId']);
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

    const gpaInfo = calculateGPAAndRank(studentId, grade, classNo);
    const assessments = getStudentAssessments(studentId);
    const homeroomTeacher = getHomeroomTeacher(grade, classNo);

    const html = _createPp6ReportHTML({
      schoolName: settings['ชื่อโรงเรียน'],
      schoolAddress: settings['ที่อยู่โรงเรียน'],
      logoDataUrl: logoDataUrl,
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
      academicYear: settings['ปีการศึกษา'] || new Date().getFullYear() + 543
    });

    const fileName = `ปพ6_${studentFullName || studentId}_${studentId}_${term}.pdf`;
    const pdfBlob = HtmlService.createHtmlOutput(html)
                               .getAs('application/pdf')
                               .setName(fileName);
    
    const folder = getOutputFolder_();
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
_savePDFCache(cacheKey, file.getId(), fileName);
Logger.log(`✅ PP6 PDF generated: ${fileName}`);

const id = file.getId();
const previewUrl  = `https://drive.google.com/file/d/${id}/preview`;          // ใช้กับ <iframe>
const downloadUrl = `https://drive.google.com/uc?export=download&id=${id}`;   // ใช้กับปุ่มดาวน์โหลด

// คืนเป็น "อ็อบเจกต์" เพื่อให้ฝั่งหน้าเว็บเลือกใช้ได้ถูกบริบท
return { previewUrl, downloadUrl, fileId: id, name: fileName };

    
  } catch (error) {
    Logger.log(`❌ Error in generatePp6PDFComplete for student ${studentId}: ${error.message}`);
    throw new Error(`ไม่สามารถสร้างรายงาน ปพ.6: ${error.message}`);
  }
}

// ======== Advanced Cache Functions ========

/**
 * Cache Warming - โหลดข้อมูลล่วงหน้า
 */
function warmupCache() {
  const startTime = new Date().getTime();
  let warmedCount = 0;
  
  try {
    Logger.log('🔥 Starting cache warmup...');
    
    // 1. Warm up การตั้งค่าระบบ
    getWebAppSettingsWithCache();
    warmedCount++;
    
    // 2. Warm up รายวิชาทุกระดับชั้น
    const grades = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'];
    grades.forEach(grade => {
      getSubjectList(grade);
      warmedCount++;
    });
    
    // 3. Warm up ข้อมูล SCORES_WAREHOUSE
    _readSheetToObjects('SCORES_WAREHOUSE');
    warmedCount++;
    
    // 4. Warm up ข้อมูลการประเมิน
    _readSheetToObjects('การประเมินอ่านคิดเขียน');
    _readSheetToObjects('การประเมินคุณลักษณะ');
    _readSheetToObjects('การประเมินกิจกรรมพัฒนาผู้เรียน');
    warmedCount += 3;
    
    const endTime = new Date().getTime();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`🔥 Cache warmup completed: ${warmedCount} items in ${duration}s`);
    return `Cache Warmup เสร็จสิ้น: โหลดข้อมูล ${warmedCount} รายการใน ${duration} วินาที`;
    
  } catch (error) {
    Logger.log('❌ Cache warmup error: ' + error.message);
    throw new Error('ไม่สามารถ Warmup Cache ได้: ' + error.message);
  }
}

/**
 * ดูสถิติ Cache
 */
function getCacheStatistics() {
  try {
    const cache = CacheService.getScriptCache();
    const stats = {
      timestamp: new Date().toLocaleString('th-TH'),
      cacheType: 'Script Cache',
      limits: {
        maxSize: '10 MB',
        maxItemSize: '100 KB', 
        maxItems: 'ไม่จำกัด',
        defaultExpiration: '10 นาที'
      },
      recommendations: [
        'ใช้ Cache Warmup เพื่อโหลดข้อมูลล่วงหน้า',
        'ล้าง Cache เมื่อข้อมูลในชีตมีการเปลี่ยนแปลง',
        'PDF Cache จะช่วยลดเวลาสร้างรายงานที่เหมือนเดิม',
        'การตั้งค่าระบบและรายวิชาถูกแคชไว้นานเพราะไม่ค่อยเปลี่ยน'
      ]
    };
    
    return stats;
  } catch (error) {
    Logger.log('Error in getCacheStatistics: ' + error.message);
    throw new Error('ไม่สามารถดูสถิติ Cache ได้: ' + error.message);
  }
}

/**
 * ล้าง Cache เมื่อข้อมูลใน Sheets มีการเปลี่ยนแปลง
 */
function invalidateCacheAfterDataUpdate(sheetName = null) {
  try {
    const cache = CacheService.getScriptCache();
    let clearedCount = 0;
    
    if (sheetName) {
      // ล้างแคชเฉพาะที่เกี่ยวข้องกับชีตนี้
      const relatedKeys = [];
      
      if (sheetName === 'รายวิชา') {
        ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'].forEach(grade => {
          relatedKeys.push(_createCacheKey('subjects', grade));
        });
      } else if (sheetName === 'SCORES_WAREHOUSE') {
        relatedKeys.push(_createCacheKey('sheet', 'SCORES_WAREHOUSE'));
        // ล้าง PDF cache ทั้งหมด เพราะคะแนนเปลี่ยน
        Logger.log('🗑️ Clearing PDF cache due to score updates');
      } else if (sheetName.includes('การประเมิน')) {
        relatedKeys.push(_createCacheKey('sheet', sheetName));
      }
      
      relatedKeys.forEach(key => {
        cache.remove(key);
        clearedCount++;
      });
      
    } else {
      // ล้างทั้งหมด
      cache.removeAll();
      clearedCount = 'ทั้งหมด';
    }
    
    const message = `ล้าง Cache เรียบร้อย: ${clearedCount} รายการ สำหรับ ${sheetName || 'ทุกข้อมูล'}`;
    Logger.log(`🧹 ${message}`);
    return message;
    
  } catch (error) {
    Logger.log('Error in invalidateCacheAfterDataUpdate: ' + error.message);
    throw new Error('ไม่สามารถล้าง Cache ได้: ' + error.message);
  }
}

/**
 * ฟังก์ชันเสริมสำหรับ Performance Monitoring
 */
function measureFunctionPerformance(functionName, ...args) {
  const startTime = new Date().getTime();
  
  try {
    let result;
    switch (functionName) {
      case 'getSubjectList':
        result = getSubjectList(...args);
        break;
      case 'getStudentListForPp6':
        result = getStudentListForPp6(...args);
        break;
      case 'getStudentAssessments':
        result = getStudentAssessments(...args);
        break;
      case 'generatePp6PDFComplete':
        result = generatePp6PDFComplete(...args);
        break;
      default:
        throw new Error(`ไม่รองรับการวัด performance ของฟังก์ชัน ${functionName}`);
    }
    
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    Logger.log(`⏱️ ${functionName}(${args.join(', ')}) took ${duration}ms`);
    
    return {
      result: result,
      performance: {
        duration: duration,
        timestamp: new Date().toISOString()
      }
    };
    
  } catch (error) {
    const endTime = new Date().getTime();
    const duration = endTime - startTime;
    
    Logger.log(`❌ ${functionName}(${args.join(', ')}) failed after ${duration}ms: ${error.message}`);
    throw error;
  }
}

// ======== Warehouse Management Functions ========

/**
 * รีบิลด์คลังคะแนนสำหรับห้องเฉพาะ (เวอร์ชันสมบูรณ์)
 */
/**
 * รีบิลด์คลังคะแนนสำหรับห้องเฉพาะ (แก้ไขชื่อชีตให้ถูกต้อง)
 */
function rebuildScoresWarehouseForClass(grade, classNo, academicYear) {
  try {
    Logger.log(`🔨 Starting rebuild for ${grade} ห้อง ${classNo}`);
    
    const ss = _openSpreadsheet();
    const warehouseSheet = ss.getSheetByName('SCORES_WAREHOUSE');
    
    if (!warehouseSheet) {
      throw new Error('ไม่พบชีต SCORES_WAREHOUSE');
    }
    
    // สำรอง: คัดลอกชีต SCORES_WAREHOUSE ก่อนแก้ไข
    // ✅ เก็บ backup เพียง 1 ชีตล่าสุด
const backupName = 'BACKUP_WAREHOUSE_LATEST';

// ลบ backup เก่าก่อน (ถ้ามี)
const oldBackup = ss.getSheetByName(backupName);
if (oldBackup) {
  ss.deleteSheet(oldBackup);
  Logger.log(`🗑️ ลบ backup เก่า: ${backupName}`);
}

// สร้าง backup ใหม่
warehouseSheet.copyTo(ss).setName(backupName);
Logger.log(`💾 สร้าง backup ใหม่: ${backupName} (${new Date().toLocaleString('th-TH')})`);
    
    // 1. อ่านข้อมูลเดิมทั้งหมด
    const allData = warehouseSheet.getDataRange().getValues();
    const headers = allData[0];
    
    // 2. กรองเอาข้อมูลชั้นอื่นๆ ออก (เก็บไว้)
    const otherClassData = [];
    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      const gradeColIndex = headers.indexOf('grade');
      const classColIndex = headers.indexOf('class_no');
      
      if (row[gradeColIndex] !== grade || row[classColIndex] != classNo) {
        otherClassData.push(row);
      }
    }
    
    Logger.log(`📋 Kept ${otherClassData.length} rows from other classes`);
    
    // 3. ดึงรายวิชาของชั้นนี้จากชีต "รายวิชา"
    const subjects = getSubjectList(grade);
    Logger.log(`📚 Found ${subjects.length} subjects for ${grade}`);
    
    const newRows = [];
    let successCount = 0;
    let errorCount = 0;
    const errorDetails = [];
    
    // 4. วนลูปแต่ละวิชา
    subjects.forEach(subject => {
      try {
        // ⭐ แก้ไข: ลบจุดออกจากชื่อชั้น (ป.3 → ป3)
        const sheetName = `${subject.name} ${grade.replace('.', '')}-${classNo}`;
        const scoreSheet = ss.getSheetByName(sheetName);
        
        if (!scoreSheet) {
          const error = `ไม่พบชีต: ${sheetName}`;
          Logger.log(`⚠️ ${error}`);
          errorDetails.push(error);
          errorCount++;
          return;
        }
        
        Logger.log(`✅ Processing: ${sheetName}`);
        
        // อ่านข้อมูลจากชีตคะแนน
        const scoreData = scoreSheet.getDataRange().getValues();
        
        if (scoreData.length < 5) {
          const error = `ชีต ${sheetName} มีข้อมูลไม่ครบ (${scoreData.length} แถว)`;
          Logger.log(`⚠️ ${error}`);
          errorDetails.push(error);
          errorCount++;
          return;
        }
        
        let studentCount = 0;
        
        // แถวที่ 5 (index 4) เป็นต้นไป คือข้อมูลนักเรียน
        for (let i = 4; i < scoreData.length; i++) {
          const row = scoreData[i];
          
          const studentId = String(row[1] || '').trim(); // Col 1 = รหัสนักเรียน
          if (!studentId || isNaN(studentId)) {
            Logger.log(`⚠️ Skip row ${i+1}: รหัสนักเรียนไม่ถูกต้อง (${studentId})`);
            continue;
          }
          
          const term1Total = parseFloat(row[15]) || 0; // Col 15 = รวมภาค 1
          const term2Total = parseFloat(row[29]) || 0; // Col 29 = รวมภาค 2
          const average = parseFloat(row[31]) || ((term1Total + term2Total) / 2);
          
          // คำนวณ final_grade
          const gradeInfo = _scoreToGPA(average);
          
          // เพิ่มข้อมูลลง warehouse
          newRows.push([
            studentId,                    // student_id
            grade,                        // grade
            classNo,                      // class_no
            subject.code,                 // subject_code
            subject.name,                 // subject_name
            subject.type,                 // subject_type
            subject.hours,                // hours
            term1Total,                   // term1_total
            term2Total,                   // term2_total
            average,                      // average
            gradeInfo.gpa,                // final_grade
            sheetName,                    // sheet_name
            academicYear || 2568,         // academic_year
            new Date()                    // updated_at
          ]);
          
          studentCount++;
        }
        
        Logger.log(`   📝 บันทึก ${studentCount} นักเรียน`);
        successCount++;
        
      } catch (e) {
        const error = `Error ${subject.name}: ${e.message}`;
        Logger.log(`❌ ${error}`);
        errorDetails.push(error);
        errorCount++;
      }
    });
    
    // 5. ตรวจสอบว่ามีข้อมูลใหม่หรือไม่
    if (newRows.length === 0) {
      throw new Error(
        `ไม่สามารถดึงคะแนนได้เลย!\n\n` +
        `ปัญหาที่พบ:\n${errorDetails.join('\n')}\n\n` +
        `กรุณาตรวจสอบว่า:\n` +
        `1. ชีตคะแนนมีชื่อถูกต้อง (เช่น "ศิลปะ ป3-1")\n` +
        `2. ชีตคะแนนมีข้อมูลนักเรียนตั้งแต่แถวที่ 5\n` +
        `3. คอลัมน์ที่ 1 เป็นรหัสนักเรียน`
      );
    }
    
    // 6. เขียนข้อมูลกลับลงชีต
    const finalData = [headers, ...otherClassData, ...newRows];
    
    warehouseSheet.clear();
    if (finalData.length > 0) {
      warehouseSheet.getRange(1, 1, finalData.length, headers.length)
                    .setValues(finalData);
    }
    
    // 7. ล้าง cache
    invalidateCacheAfterDataUpdate('SCORES_WAREHOUSE');
    
    const message = `✅ รีบิลด์สำเร็จ: ${grade} ห้อง ${classNo}\n\n` +
                   `📊 วิชาที่ประมวลผลสำเร็จ: ${successCount}/${subjects.length}\n` +
                   `📝 นักเรียนที่บันทึก: ${newRows.length} รายการ\n` +
                   `ปีการศึกษา: ${academicYear || 2568}\n\n` +
                   (errorCount > 0 ? `⚠️ ข้อผิดพลาด ${errorCount} รายการ:\n${errorDetails.join('\n')}` : '');
    
    Logger.log(message);
    return message;
    
  } catch (error) {
    Logger.log(`❌ Error in rebuildScoresWarehouseForClass: ${error.message}`);
    throw error;
  }
}

/**
 * 9. รีบิลด์คลังคะแนนสำหรับทั้งชั้น (พร้อม Cache Invalidation)
 */
function rebuildScoresWarehouseForGrade(grade, academicYear) {
  try {
    rebuildScoresWarehouseForClass(grade, '1', academicYear);
    rebuildScoresWarehouseForClass(grade, '2', academicYear);
    
    // ล้างแคชที่เกี่ยวข้อง
    invalidateCacheAfterDataUpdate('SCORES_WAREHOUSE');
    
    const message = `รีบิลด์คลังคะแนนทั้งชั้นสำเร็จสำหรับ ${grade} ปีการศึกษา ${academicYear}`;
    console.log(message);
    return message;
  } catch (error) {
    console.error('Error in rebuildScoresWarehouseForGrade:', error.message);
    throw new Error('ไม่สามารถรีบิลด์คลังคะแนนทั้งชั้นได้: ' + error.message);
  }
}

/**
 * 10. รีบิลด์คลังคะแนนทั้งหมด (พร้อม Cache Invalidation)
 */
function rebuildScoresWarehouseAll(academicYear) {
  try {
    const grades = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6'];
    
    grades.forEach(grade => {
      rebuildScoresWarehouseForGrade(grade, academicYear);
    });
    
    // ล้างแคชทั้งหมด
    clearAllCache();
    
    const message = `รีบิลด์คลังคะแนนทั้งหมดสำเร็จสำหรับปีการศึกษา ${academicYear}`;
    console.log(message);
    return message;
  } catch (error) {
    console.error('Error in rebuildScoresWarehouseAll:', error.message);
    throw new Error('ไม่สามารถรีบิลด์คลังคะแนนทั้งหมดได้: ' + error.message);
  }
}

// ======== HTML Template Functions ========

/**
 * สร้าง HTML Template สำหรับรายงานรายวิชา
 */
function _createSubjectReportHTML(data) {
  const { schoolName, logoDataUrl, subjectName, subjectCode, grade, classNo, term, students } = data;
  
  // คำนวณสถิติคะแนน
  const scores = students.map(s => parseFloat(s.average) || 0).filter(score => score > 0);
  const avgScore = scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length) : 0;
  const maxScore = scores.length ? Math.max(...scores) : 0;
  const minScore = scores.length ? Math.min(...scores) : 0;
  
  const termText = term === 'both' ? 'ทั้งสองภาคเรียน' : `ภาคเรียนที่ ${term}`;
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page { 
      size: A4; 
      margin: 15mm; 
    }
    body { 
      font-family: 'Sarabun', 'TH SarabunPSK', sans-serif; 
      font-size: 14px; 
      line-height: 1.4; 
      color: #333; 
    }
    .header { 
      text-align: center; 
      margin-bottom: 20px; 
      border-bottom: 2px solid #4a90e2;
      padding-bottom: 15px;
    }
    .logo { 
      width: 60px; 
      height: 60px; 
      margin-bottom: 10px; 
    }
    .school-name { 
      font-size: 18px; 
      font-weight: bold; 
      margin-bottom: 5px; 
      color: #2c5282;
    }
    .report-title { 
      font-size: 16px; 
      font-weight: bold; 
      margin-bottom: 3px;
    }
    .report-info { 
      font-size: 14px; 
      color: #4a5568;
    }
    .info-section { 
      display: flex; 
      justify-content: space-between; 
      margin-bottom: 15px; 
      background: #f7fafc;
      padding: 10px;
      border-radius: 5px;
    }
    table { 
      width: 100%; 
      border-collapse: collapse; 
      margin-bottom: 15px; 
    }
    th, td { 
      border: 1px solid #ccc; 
      padding: 8px; 
      text-align: center; 
    }
    th { 
      background: #4a90e2; 
      color: white; 
      font-weight: bold; 
    }
    tr:nth-child(even) { 
      background: #f8f9fa; 
    }
    tr:hover { 
      background: #e2f4ff; 
    }
    .stats-section { 
      margin-top: 20px; 
      background: #f0f8ff;
      padding: 15px;
      border-radius: 8px;
    }
    .stats-title { 
      font-weight: bold; 
      margin-bottom: 10px;
      color: #2c5282;
    }
    .stats-grid { 
      display: grid; 
      grid-template-columns: repeat(3, 1fr); 
      gap: 15px; 
    }
    .stat-item { 
      text-align: center; 
      padding: 10px;
      background: white;
      border-radius: 5px;
      border: 1px solid #e2e8f0;
    }
    .stat-value { 
      font-size: 18px; 
      font-weight: bold; 
      color: #2b6cb0; 
    }
    .stat-label { 
      font-size: 12px; 
      color: #4a5568; 
    }
    .footer { 
      margin-top: 30px; 
      text-align: center; 
      font-size: 12px; 
      color: #666;
    }
    .grade-a { color: #22c55e; font-weight: bold; }
    .grade-b { color: #3b82f6; font-weight: bold; }
    .grade-c { color: #f59e0b; font-weight: bold; }
    .grade-f { color: #ef4444; font-weight: bold; }
  </style>
</head>
<body>
  <div class="header">
    ${logoDataUrl ? `<img src="${logoDataUrl}" class="logo" alt="โลโก้โรงเรียน">` : ''}
    <div class="school-name">${schoolName || 'โรงเรียน'}</div>
    <div class="report-title">รายงานผลการเรียนรายวิชา</div>
    <div class="report-info">${termText} ปีการศึกษา ${new Date().getFullYear() + 543}</div>
  </div>

  <div class="info-section">
    <div><strong>วิชา:</strong> ${subjectName} (${subjectCode})</div>
    <div><strong>ระดับชั้น:</strong> ${grade} ห้อง ${classNo}</div>
    <div><strong>จำนวนนักเรียน:</strong> ${students.length} คน</div>
  </div>

  <table>
    <thead>
      <tr>
        <th style="width: 8%">ลำดับ</th>
        <th style="width: 15%">รหัสนักเรียน</th>
        <th style="width: 35%">ชื่อ - นามสกุล</th>
        <th style="width: 12%">คะแนนเฉลี่ย</th>
        <th style="width: 15%">ระดับผลการเรียน</th>
        <th style="width: 15%">ผลการเรียน</th>
      </tr>
    </thead>
    <tbody>
      ${students.map((student, index) => {
        const score = parseFloat(student.average) || 0;
        const gradeInfo = _scoreToGPA(score);
        const gradeClass = gradeInfo.gpa >= 3.5 ? 'grade-a' : gradeInfo.gpa >= 2.5 ? 'grade-b' : gradeInfo.gpa >= 1.5 ? 'grade-c' : 'grade-f';
        
        return `
        <tr>
          <td>${index + 1}</td>
          <td>${student.student_id}</td>
          <td style="text-align: left; padding-left: 10px;">${student.student_name || 'นักเรียนรหัส ' + student.student_id}</td>
          <td><strong>${score.toFixed(1)}</strong></td>
          <td class="${gradeClass}">${gradeInfo.gpa}</td>
          <td class="${gradeClass}">${gradeInfo.letter}</td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>

  <div class="stats-section">
    <div class="stats-title">สถิติผลการเรียน</div>
    <div class="stats-grid">
      <div class="stat-item">
        <div class="stat-value">${avgScore.toFixed(1)}</div>
        <div class="stat-label">คะแนนเฉลี่ย</div>
      </div>
      <div class="stat-item">
        <div class="stat-value">${maxScore.toFixed(1)}</div>
        <div class="stat-label">คะแนนสูงสุด</div>
      </div>
      <div class="stat-item">
        <div class="stat-value">${minScore.toFixed(1)}</div>
        <div class="stat-label">คะแนนต่ำสุด</div>
      </div>
    </div>
  </div>

  <div class="footer">
    สร้างเมื่อ: ${new Date().toLocaleString('th-TH')}
    <br>ระบบรายงานผลการเรียน - ${schoolName || 'โรงเรียน'}
  </div>
</body>
</html>`;
}

/**
 * เทมเพลต HTML ที่ปรับปรุงให้แสดงชื่อครูประจำชั้น และชื่อวิชาชิดซ้าย
 */
function _createPp6ReportHTML(data) {
  const { 
    schoolName, schoolAddress, logoDataUrl, studentId, studentName, grade, classNo, 
    term, subjects, gpaInfo, assessments, principalName, homeroomTeacher, academicYear 
  } = data;
  
  const termText = term === 'both' ? 'ภาคเรียนที่ 1 และ 2' : `ภาคเรียนที่ ${term}`;
  
const formatActivityResult = (resultText) => {
  if (!resultText) return 'มผ'; // ถ้าไม่มีค่า
  if (resultText === 'ผ่าน' || resultText === 'ผ') return 'ผ';
  if (resultText === 'ไม่ผ่าน' || resultText === 'มผ') return 'มผ';
  return resultText; // แสดงค่าตามที่ได้มา
};
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4; margin: 12mm; }
    body { font-family: 'Sarabun', 'TH SarabunPSK', sans-serif; font-size: 13px; line-height: 1.3; color: #000; }
    .header { text-align: center; margin-bottom: 15px; }
    .logo { width: 50px; height: 50px; margin-bottom: 8px; }
    .report-title { font-size: 16px; font-weight: bold; margin-bottom: 3px;}
    .school-info { font-size: 14px; margin-bottom: 2px;}
    .student-info { margin: 15px 0; background: #f9f9f9; padding: 8px; border: 1px solid #ddd; }
    .student-info div { margin-bottom: 3px; }
    .main-table { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 12px; }
    .main-table th, .main-table td { border: 1px solid #000; padding: 4px 6px; text-align: center; }
    .main-table th { background: #f0f0f0; font-weight: bold; }
    
    /* ปรับให้คอลัมน์ชื่อวิชา (คอลัมน์ที่ 3) ชิดซ้าย ยกเว้นหัวข้อ */
    .main-table td:nth-child(3) { text-align: left; padding-left: 8px; }
    
    .summary-section { margin: 15px 0; }
    .summary-title { font-weight: bold; font-size: 14px; margin-bottom: 8px; text-align: center; background: #e0e0e0; padding: 5px; border: 1px solid #ccc; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .summary-table td { padding: 4px; vertical-align: top; }
    .summary-table td.value { font-weight: bold; }

    /* ความกว้างคอลัมน์ที่ปรับแก้แล้ว */
    .summary-table td:nth-child(1) { width: 28%; }
    .summary-table td:nth-child(2) { width: 15%; }
    .summary-table td:nth-child(3) { width: 39%; }
    .summary-table td:nth-child(4) { width: 18%; }

    .signatures { margin-top: 20px; display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; text-align: center; }
    .signature-box { padding: 8px; }
    .signature-line { border-top: 1px solid #000; margin-top: 30px; padding-top: 5px; }
    .footer { margin-top: 15px; text-align: right; font-size: 11px; }
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
    <div><strong>รหัสประจำตัวนักเรียน:</strong> ${studentId} <strong>ชั้น:</strong> ${grade} ห้อง ${classNo}</div>
    <div><strong>ชื่อ - นามสกุล:</strong> ${studentName}</div>
  </div>

  <table class="main-table">
    <thead>
      <tr>
        <th style="width: 8%">ลำดับ</th><th style="width: 12%">รหัสวิชา</th><th style="width: 35%">ชื่อวิชา</th>
        <th style="width: 10%">ประเภท</th><th style="width: 10%">จำนวนชั่วโมง</th><th style="width: 12%">คะแนนที่ได้</th><th style="width: 13%">ระดับผลการเรียน</th>
      </tr>
    </thead>
    <tbody>
      ${subjects.map(subject => `
        <tr>
          <td>${subject.order}</td>
          <td>${subject.code}</td>
          <td>${subject.name}</td>
          <td>${subject.type}</td>
          <td>${subject.hours}</td>
          <td><strong>${subject.score ? subject.score : ''}</strong></td>
          <td><strong>${subject.type === 'กิจกรรม' ? formatActivityResult(subject.grade) : (subject.grade || '')}</strong></td>
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

  <div class="signatures">
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ${homeroomTeacher || 'ครูประจำชั้น'} )</div>
       <div>ครูที่ปรึกษา</div>
     </div>
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ${principalName || 'ผู้อำนวยการโรงเรียน'} )</div>
       <div>ผู้บริหารสถานศึกษา</div>
     </div>
     <div class="signature-box">
       <div>ลงชื่อ</div>
       <div class="signature-line">( ............................................................... )</div>
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
