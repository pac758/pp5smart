// ============================================================
// 🔄 SCHOOLMIS_SYNC.JS - ระบบ Sync 2 ทาง กับ SchoolMIS
// ============================================================

/**
 * =====================================================
 * 📋 CONFIG: การ Map คอลัมน์ระหว่างสองระบบ
 * =====================================================
 */
const SCHOOLMIS_CONFIG = {
  // คอลัมน์ใน SchoolMIS CSV (0-indexed)
  csv: {
    NO: 0,              // ที่
    STUDENT_CODE: 1,    // รหัสนักเรียน
    ID_CARD: 2,         // เลข ปชช.
    FIRSTNAME: 3,       // ชื่อ
    LASTNAME: 4,        // นามสกุล
    SCORE_1: 5,         // ครั้งที่ 1
    SCORE_2: 6,         // ครั้งที่ 2
    SCORE_3: 7,         // ครั้งที่ 3
    SCORE_4: 8,         // ครั้งที่ 4
    SUM_1_4: 9,         // รวม (1-4)
    SCORE_5: 10,        // ครั้งที่ 5
    MAKEUP_MID: 11,     // แก้ตัวกลางภาค
    SCORE_6: 12,        // ครั้งที่ 6
    SCORE_7: 13,        // ครั้งที่ 7
    SCORE_8: 14,        // ครั้งที่ 8
    SCORE_9: 15,        // ครั้งที่ 9
    SUM_5_9: 16,        // รวม (5-9)
    MIDTERM_TOTAL: 17,  // รวมระหว่างภาค
    FINAL_SCORE: 18,    // ครั้งที่10ปลายภาค
    TOTAL: 19,          // ทั้งหมด
    GRADE: 20           // เกรด
  },
  
  // Header ของ SchoolMIS
  csvHeaders: [
    'ที่', 'รหัสนักเรียน', 'เลข ปชช.', 'ชื่อ', 'นามสกุล',
    'ครั้งที่ 1', 'ครั้งที่ 2', 'ครั้งที่ 3', 'ครั้งที่ 4', 'รวม',
    'ครั้งที่5', 'แก้ตัวกลางภาค', 'ครั้งที่6', 'ครั้งที่7', 'ครั้งที่8', 'ครั้งที่9', 'รวม',
    'รวมระหว่างภาค', 'ครั้งที่10ปลายภาค', 'ทั้งหมด', 'เกรด'
  ]
};

/**
 * =====================================================
 * 📥 IMPORT: นำเข้าข้อมูลจาก SchoolMIS CSV
 * =====================================================
 */

/**
 * นำเข้าข้อมูลคะแนนจากไฟล์ CSV ของ SchoolMIS
 * @param {string} csvContent - เนื้อหาไฟล์ CSV
 * @param {string} sheetName - ชื่อชีตคะแนนในระบบของเรา
 * @param {string} term - ภาคเรียน ('term1' หรือ 'term2')
 * @returns {Object} ผลการนำเข้า
 */
function importFromSchoolMIS(csvContent, sheetName, term) {
  try {
    Logger.log('=== เริ่มนำเข้าจาก SchoolMIS ===');
    Logger.log('ชีตเป้าหมาย: ' + sheetName + ', ภาคเรียน: ' + term);
    
    // 1. Parse CSV
    const csvData = parseSchoolMISCsv(csvContent);
    if (!csvData || csvData.length === 0) {
      throw new Error('ไม่สามารถอ่านข้อมูล CSV ได้');
    }
    
    Logger.log('พบข้อมูลนักเรียน: ' + csvData.length + ' คน');
    
    // 2. เปิดชีตเป้าหมาย
    const ss = SS();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('ไม่พบชีต: ' + sheetName);
    }
    
    // 3. ดึงข้อมูลนักเรียนในชีตเพื่อ Match
    const sheetData = sheet.getDataRange().getValues();
    const studentMap = buildStudentMapSync_(sheetData);
    
    // 4. กำหนดคอลัมน์ตาม term
    const cols = getTermColumnsSync_(term);
    
    // 5. นำเข้าข้อมูล
    var importCount = 0;
    var skipCount = 0;
    var notFoundStudents = [];
    
    csvData.forEach(function(csvRow, index) {
      // หานักเรียนในชีต (ใช้รหัสนักเรียนหรือเลขบัตรประชาชน)
      var studentCode = String(csvRow.studentCode || '').trim();
      var idCard = String(csvRow.idCard || '').trim();
      var fullName = (csvRow.firstname + ' ' + csvRow.lastname).trim();
      
      var rowIndex = findStudentRowSync_(studentMap, studentCode, idCard, fullName);
      
      if (rowIndex === -1) {
        notFoundStudents.push({ code: studentCode, name: fullName });
        skipCount++;
        return;
      }
      
      // บันทึกคะแนน
      var row = rowIndex + 1; // +1 เพราะ getRange เริ่มที่ 1
      
      // scores[0-8] = ครั้ง1-4, ครั้ง5, ครั้ง6-9 → scoreSlots[0-8]
      // scores[9] = ครั้ง10ปลายภาค → scoreSlots[9] = s10
      var slots = cols.scoreSlots;
      var scoresCount = Math.min(csvRow.scores.length, slots.length);
      
      for (var i = 0; i < scoresCount; i++) {
        var score = csvRow.scores[i];
        if (score !== null && score !== undefined && score !== '') {
          sheet.getRange(row, slots[i] + 1).setValue(score);
        }
      }

      // คำนวณรวมก่อนกลางภาค (ครั้ง1-4)
      var sum14 = 0;
      for (var k = 0; k < 4; k++) sum14 += Number(csvRow.scores[k]) || 0;
      sheet.getRange(row, cols.sum14 + 1).setValue(sum14);

      // คำนวณรวมหลังกลางภาค (ครั้ง6-9) — scores[5-8]
      var sum69 = 0;
      for (var k = 5; k < 9; k++) sum69 += Number(csvRow.scores[k]) || 0;
      sheet.getRange(row, cols.sum69 + 1).setValue(sum69);
      
      // รวมระหว่างภาค
      if (csvRow.midtermTotal !== null && csvRow.midtermTotal !== undefined) {
        sheet.getRange(row, cols.midTotal + 1).setValue(csvRow.midtermTotal);
      }
      
      // ปลายภาค (ครั้ง10) อยู่ใน scores[9] → s10 แล้ว (ผ่าน scoreSlots loop)
      
      // รวมทั้งหมด
      if (csvRow.total !== null && csvRow.total !== undefined) {
        sheet.getRange(row, cols.total + 1).setValue(csvRow.total);
      }
      
      // เกรด
      if (csvRow.grade && csvRow.grade !== '') {
        sheet.getRange(row, cols.grade + 1).setValue(csvRow.grade);
      }
      
      importCount++;
      Logger.log('✅ นำเข้า: ' + fullName + ' (แถว ' + row + ')');
    });
    
    var result = {
      success: true,
      imported: importCount,
      skipped: skipCount,
      notFound: notFoundStudents,
      message: '✅ นำเข้าสำเร็จ ' + importCount + ' คน, ข้าม ' + skipCount + ' คน'
    };
    
    Logger.log(result.message);
    return result;
    
  } catch (error) {
    Logger.log('❌ Error in importFromSchoolMIS: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Parse ไฟล์ CSV จาก SchoolMIS
 * @param {string} csvContent - เนื้อหา CSV
 * @returns {Array<Object>} ข้อมูลนักเรียน
 */
function parseSchoolMISCsv(csvContent) {
  try {
    // ลบ BOM ถ้ามี
    if (csvContent.charCodeAt(0) === 0xFEFF) {
      csvContent = csvContent.substring(1);
    }
    
    var lines = csvContent.split(/\r?\n/);
    var students = [];
    
    // ข้าม header (บรรทัดแรก)
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      
      // Parse CSV line (รองรับ comma ใน quotes)
      var cols = parseCSVLineSync_(line);
      
      if (cols.length < 21) {
        Logger.log('⚠️ แถว ' + (i + 1) + ' มีคอลัมน์ไม่ครบ: ' + cols.length);
        continue;
      }
      
      var cfg = SCHOOLMIS_CONFIG.csv;
      
      var student = {
        no: parseNumberSync_(cols[cfg.NO]),
        studentCode: (cols[cfg.STUDENT_CODE] || '').trim(),
        idCard: (cols[cfg.ID_CARD] || '').trim(),
        firstname: (cols[cfg.FIRSTNAME] || '').trim(),
        lastname: (cols[cfg.LASTNAME] || '').trim(),
        // SchoolMIS มี 10 ครั้ง (1-9 + ปลายภาค) ตรง 1:1 กับ pp5smart
        // scores[0-3] = ครั้ง1-4 (ก่อนกลางภาค)
        // scores[4]   = ครั้ง5 (กลางภาค)
        // scores[5-8] = ครั้ง6-9 (หลังกลางภาค)
        // scores[9]   = ครั้ง10ปลายภาค
        scores: [
          parseNumberSync_(cols[cfg.SCORE_1]),
          parseNumberSync_(cols[cfg.SCORE_2]),
          parseNumberSync_(cols[cfg.SCORE_3]),
          parseNumberSync_(cols[cfg.SCORE_4]),
          parseNumberSync_(cols[cfg.SCORE_5]),
          parseNumberSync_(cols[cfg.SCORE_6]),
          parseNumberSync_(cols[cfg.SCORE_7]),
          parseNumberSync_(cols[cfg.SCORE_8]),
          parseNumberSync_(cols[cfg.SCORE_9]),
          parseNumberSync_(cols[cfg.FINAL_SCORE])  // ครั้งที่10ปลายภาค
        ],
        makeupMid: parseNumberSync_(cols[cfg.MAKEUP_MID]),
        midtermTotal: parseNumberSync_(cols[cfg.MIDTERM_TOTAL]),
        finalScore: parseNumberSync_(cols[cfg.FINAL_SCORE]),
        total: parseNumberSync_(cols[cfg.TOTAL]),
        grade: (cols[cfg.GRADE] || '').trim()
      };
      
      students.push(student);
    }
    
    return students;
    
  } catch (error) {
    Logger.log('❌ Error parsing CSV: ' + error.message);
    return [];
  }
}

/**
 * =====================================================
 * 📤 EXPORT: ส่งออกข้อมูลไปยัง SchoolMIS Format
 * =====================================================
 */

/**
 * ส่งออกข้อมูลคะแนนเป็น CSV สำหรับ SchoolMIS
 * @param {string} sheetName - ชื่อชีตคะแนน
 * @param {string} term - ภาคเรียน ('term1' หรือ 'term2')
 * @returns {Object} ผลการส่งออก พร้อม CSV content
 */
function exportToSchoolMIS(sheetName, term) {
  try {
    Logger.log('=== เริ่มส่งออกไป SchoolMIS ===');
    Logger.log('ชีตต้นทาง: ' + sheetName + ', ภาคเรียน: ' + term);
    
    // 1. เปิดชีต
    var ss = SS();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('ไม่พบชีต: ' + sheetName);
    }
    
    // 2. ดึงข้อมูล
    var data = sheet.getDataRange().getValues();
    var cols = getTermColumnsSync_(term);
    
    // 3. ดึงข้อมูลนักเรียนจากชีต Students เพื่อเอา idCard
    var studentsSheet = ss.getSheetByName('Students');
    var studentsData = studentsSheet ? studentsSheet.getDataRange().getValues() : [];
    var studentIdCardMap = buildStudentIdCardMapSync_(studentsData);
    
    // 4. สร้าง CSV content
    var csvLines = [];
    
    // Header
    csvLines.push(SCHOOLMIS_CONFIG.csvHeaders.join(','));
    
    // Data rows (เริ่มจากแถวที่ 5, index 4)
    for (var i = 4; i < data.length; i++) {
      var row = data[i];
      
      // ข้ามแถวว่าง
      if (!row[1]) continue;
      
      var studentId = String(row[1]).trim(); // เลขประจำตัว
      var fullName = String(row[2]).trim();  // ชื่อ-นามสกุล
      
      // แยกชื่อ-นามสกุล
      var nameParts = splitThaiNameSync_(fullName);
      
      // หา idCard จาก Students sheet
      var idCard = studentIdCardMap[studentId] || '';
      
      // ดึงคะแนนจาก pp5smart scoreSlots (10 ช่อง ตรง 1:1 กับ SchoolMIS)
      var slots = cols.scoreSlots;
      var scores = [];
      for (var j = 0; j < slots.length; j++) {
        scores.push(row[slots[j]] || '');
      }
      
      // รวม (1-4) และ (6-9) อ่านจากชีตโดยตรง
      var sum1_4 = row[cols.sum14] || '';
      var sum5_9 = row[cols.sum69] || '';
      
      var midtermTotal = row[cols.midTotal] || '';
      var finalScore = row[cols.s10] || '';  // pp5smart s10 = ปลายภาค = SchoolMIS ครั้ง10ปลายภาค
      var total = row[cols.total] || '';
      var grade = row[cols.grade] || '';
      var makeup = row[cols.makeup] || '';
      
      // สร้างแถว CSV ตามรูปแบบ SchoolMIS (21 คอลัมน์) — ตรง 1:1
      var csvRow = [
        i - 3,              // ที่ (ลำดับ)
        studentId,          // รหัสนักเรียน
        idCard,             // เลข ปชช.
        nameParts.firstname,// ชื่อ
        nameParts.lastname, // นามสกุล
        scores[0],          // ครั้งที่ 1
        scores[1],          // ครั้งที่ 2
        scores[2],          // ครั้งที่ 3
        scores[3],          // ครั้งที่ 4
        sum1_4,             // รวม (1-4)
        scores[4],          // ครั้งที่ 5 (กลางภาค)
        makeup,             // แก้ตัวกลางภาค
        scores[5],          // ครั้งที่ 6
        scores[6],          // ครั้งที่ 7
        scores[7],          // ครั้งที่ 8
        scores[8],          // ครั้งที่ 9
        sum5_9,             // รวม (5-9)
        midtermTotal,       // รวมระหว่างภาค
        finalScore,         // ครั้งที่10ปลายภาค
        total,              // ทั้งหมด
        grade               // เกรด
      ];
      
      // Escape values for CSV
      var escapedRow = csvRow.map(function(val) { return escapeCSVValueSync_(val); });
      csvLines.push(escapedRow.join(','));
    }
    
    var csvContent = csvLines.join('\n');
    
    // 5. บันทึกเป็นไฟล์ใน Drive
    var fileName = sheetName + '_' + term + '_SchoolMIS_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.csv';
    
    // เพิ่ม BOM สำหรับ UTF-8
    var bom = '\uFEFF';
    var csvWithBom = bom + csvContent;
    var blobWithBom = Utilities.newBlob(csvWithBom, 'text/csv', fileName);
    
    var file = DriveApp.createFile(blobWithBom);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var result = {
      success: true,
      fileName: fileName,
      fileUrl: file.getUrl(),
      downloadUrl: file.getDownloadUrl(),
      csvContent: csvContent,
      studentCount: csvLines.length - 1,
      message: '✅ ส่งออกสำเร็จ ' + (csvLines.length - 1) + ' คน'
    };
    
    Logger.log(result.message);
    return result;
    
  } catch (error) {
    Logger.log('❌ Error in exportToSchoolMIS: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ส่งออกคะแนนเป็น Base64 CSV สำหรับดาวน์โหลดจากหน้าเว็บ
 * @param {string} sheetName - ชื่อชีต
 * @param {string} term - ภาคเรียน
 * @returns {Object} CSV เป็น base64
 */
function exportToSchoolMIS_Base64(sheetName, term) {
  try {
    var result = exportToSchoolMIS(sheetName, term);
    
    if (!result.success) {
      return result;
    }
    
    // เพิ่ม BOM และแปลงเป็น Base64
    var bom = '\uFEFF';
    var csvWithBom = bom + result.csvContent;
    var base64 = Utilities.base64Encode(csvWithBom, Utilities.Charset.UTF_8);
    
    return {
      success: true,
      base64: base64,
      fileName: result.fileName,
      studentCount: result.studentCount,
      message: result.message
    };
    
  } catch (error) {
    Logger.log('❌ Error in exportToSchoolMIS_Base64: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * =====================================================
 * � EXPORT AVERAGE: เฉลี่ยคะแนน 2 ภาคเรียน → SchoolMIS Format
 * =====================================================
 */

/**
 * ส่งออกคะแนนเฉลี่ย 2 ภาคเรียน เป็น CSV สำหรับ SchoolMIS
 * คำนวณ: (คะแนนภาค1 + คะแนนภาค2) / 2 สำหรับทุกช่อง
 * @param {string} sheetName - ชื่อชีตคะแนน
 * @returns {Object} ผลการส่งออก พร้อม CSV content
 */
function exportToSchoolMIS_Average(sheetName) {
  try {
    Logger.log('=== เริ่มส่งออกเฉลี่ย 2 ภาค ไป SchoolMIS ===');
    Logger.log('ชีตต้นทาง: ' + sheetName);
    
    // 1. เปิดชีต
    var ss = SS();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('ไม่พบชีต: ' + sheetName);
    }
    
    // 2. ดึงข้อมูล
    var data = sheet.getDataRange().getValues();
    var cols1 = getTermColumnsSync_('term1');
    var cols2 = getTermColumnsSync_('term2');
    
    // 3. ดึงข้อมูลนักเรียนจากชีต Students เพื่อเอา idCard
    var studentsSheet = ss.getSheetByName('Students');
    var studentsData = studentsSheet ? studentsSheet.getDataRange().getValues() : [];
    var studentIdCardMap = buildStudentIdCardMapSync_(studentsData);
    
    // 4. สร้าง CSV content
    var csvLines = [];
    
    // Header
    csvLines.push(SCHOOLMIS_CONFIG.csvHeaders.join(','));
    
    // Data rows (เริ่มจากแถวที่ 5, index 4)
    for (var i = 4; i < data.length; i++) {
      var row = data[i];
      
      // ข้ามแถวว่าง
      if (!row[1]) continue;
      
      var studentId = String(row[1]).trim();
      var fullName = String(row[2]).trim();
      var nameParts = splitThaiNameSync_(fullName);
      var idCard = studentIdCardMap[studentId] || '';
      
      // ดึงคะแนนทั้ง 2 ภาค แล้วเฉลี่ย (ใช้ scoreSlots 9 ช่อง)
      var slots1 = cols1.scoreSlots;
      var slots2 = cols2.scoreSlots;
      var avgScores = [];
      for (var j = 0; j < slots1.length; j++) {
        var s1 = Number(row[slots1[j]]) || 0;
        var s2 = Number(row[slots2[j]]) || 0;
        var hasT1 = (row[slots1[j]] !== '' && row[slots1[j]] !== null && row[slots1[j]] !== undefined);
        var hasT2 = (row[slots2[j]] !== '' && row[slots2[j]] !== null && row[slots2[j]] !== undefined);
        if (hasT1 && hasT2) {
          avgScores.push(Math.round(((s1 + s2) / 2) * 100) / 100);
        } else if (hasT1) {
          avgScores.push(s1);
        } else if (hasT2) {
          avgScores.push(s2);
        } else {
          avgScores.push('');
        }
      }
      
      // เฉลี่ยรวมระหว่างภาค
      var mid1 = Number(row[cols1.midTotal]) || 0;
      var mid2 = Number(row[cols2.midTotal]) || 0;
      var hasMid1 = (row[cols1.midTotal] !== '' && row[cols1.midTotal] !== null && row[cols1.midTotal] !== undefined);
      var hasMid2 = (row[cols2.midTotal] !== '' && row[cols2.midTotal] !== null && row[cols2.midTotal] !== undefined);
      var avgMid = '';
      if (hasMid1 && hasMid2) avgMid = Math.round(((mid1 + mid2) / 2) * 100) / 100;
      else if (hasMid1) avgMid = mid1;
      else if (hasMid2) avgMid = mid2;
      
      // เฉลี่ยครั้งที่ 10 (ปลายภาค)
      var fin1 = Number(row[cols1.s10]) || 0;
      var fin2 = Number(row[cols2.s10]) || 0;
      var hasFin1 = (row[cols1.s10] !== '' && row[cols1.s10] !== null && row[cols1.s10] !== undefined);
      var hasFin2 = (row[cols2.s10] !== '' && row[cols2.s10] !== null && row[cols2.s10] !== undefined);
      var avgFinal = '';
      if (hasFin1 && hasFin2) avgFinal = Math.round(((fin1 + fin2) / 2) * 100) / 100;
      else if (hasFin1) avgFinal = fin1;
      else if (hasFin2) avgFinal = fin2;
      
      // เฉลี่ยรวมทั้งหมด
      var tot1 = Number(row[cols1.total]) || 0;
      var tot2 = Number(row[cols2.total]) || 0;
      var hasTot1 = (row[cols1.total] !== '' && row[cols1.total] !== null && row[cols1.total] !== undefined);
      var hasTot2 = (row[cols2.total] !== '' && row[cols2.total] !== null && row[cols2.total] !== undefined);
      var avgTotal = '';
      if (hasTot1 && hasTot2) avgTotal = Math.round(((tot1 + tot2) / 2) * 100) / 100;
      else if (hasTot1) avgTotal = tot1;
      else if (hasTot2) avgTotal = tot2;
      
      // เฉลี่ยเกรด
      var g1 = Number(row[cols1.grade]) || 0;
      var g2 = Number(row[cols2.grade]) || 0;
      var hasG1 = (row[cols1.grade] !== '' && row[cols1.grade] !== null && row[cols1.grade] !== undefined);
      var hasG2 = (row[cols2.grade] !== '' && row[cols2.grade] !== null && row[cols2.grade] !== undefined);
      var avgGrade = '';
      if (hasG1 && hasG2) avgGrade = Math.round(((g1 + g2) / 2) * 100) / 100;
      else if (hasG1) avgGrade = g1;
      else if (hasG2) avgGrade = g2;
      
      // คำนวณรวม (1-4) และ (6-9) จากค่าเฉลี่ย
      var sum1_4 = sumScoresSync_(avgScores.slice(0, 4));
      var sum5_9 = sumScoresSync_(avgScores.slice(5, 9));
      
      // สร้างแถว CSV ตามรูปแบบ SchoolMIS (21 คอลัมน์) — ตรง 1:1
      var csvRow = [
        i - 3,                // ที่ (ลำดับ)
        studentId,            // รหัสนักเรียน
        idCard,               // เลข ปชช.
        nameParts.firstname,  // ชื่อ
        nameParts.lastname,   // นามสกุล
        avgScores[0] || '',   // ครั้งที่ 1
        avgScores[1] || '',   // ครั้งที่ 2
        avgScores[2] || '',   // ครั้งที่ 3
        avgScores[3] || '',   // ครั้งที่ 4
        sum1_4,               // รวม (1-4)
        avgScores[4] || '',   // ครั้งที่ 5 (กลางภาค)
        '',                   // แก้ตัวกลางภาค
        avgScores[5] || '',   // ครั้งที่ 6
        avgScores[6] || '',   // ครั้งที่ 7
        avgScores[7] || '',   // ครั้งที่ 8
        avgScores[8] || '',   // ครั้งที่ 9
        sum5_9,               // รวม (5-9)
        avgMid,               // รวมระหว่างภาค
        avgFinal,             // ครั้งที่10ปลายภาค
        avgTotal,             // ทั้งหมด
        avgGrade              // เกรด
      ];
      
      var escapedRow = csvRow.map(function(val) { return escapeCSVValueSync_(val); });
      csvLines.push(escapedRow.join(','));
    }
    
    var csvContent = csvLines.join('\n');
    
    // 5. บันทึกเป็นไฟล์ใน Drive
    var fileName = sheetName + '_เฉลี่ย2ภาค_SchoolMIS_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.csv';
    var bom = '\uFEFF';
    var csvWithBom = bom + csvContent;
    var blobWithBom = Utilities.newBlob(csvWithBom, 'text/csv', fileName);
    
    var file = DriveApp.createFile(blobWithBom);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var result = {
      success: true,
      fileName: fileName,
      fileUrl: file.getUrl(),
      downloadUrl: file.getDownloadUrl(),
      csvContent: csvContent,
      studentCount: csvLines.length - 1,
      message: '✅ ส่งออกเฉลี่ย 2 ภาคสำเร็จ ' + (csvLines.length - 1) + ' คน'
    };
    
    Logger.log(result.message);
    return result;
    
  } catch (error) {
    Logger.log('❌ Error in exportToSchoolMIS_Average: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ส่งออกเฉลี่ย 2 ภาค เป็น Base64 CSV สำหรับดาวน์โหลดจากหน้าเว็บ
 * @param {string} sheetName - ชื่อชีต
 * @returns {Object} CSV เป็น base64
 */
function exportToSchoolMIS_Average_Base64(sheetName) {
  try {
    var result = exportToSchoolMIS_Average(sheetName);
    
    if (!result.success) {
      return result;
    }
    
    var bom = '\uFEFF';
    var csvWithBom = bom + result.csvContent;
    var base64 = Utilities.base64Encode(csvWithBom, Utilities.Charset.UTF_8);
    
    return {
      success: true,
      base64: base64,
      fileName: result.fileName,
      studentCount: result.studentCount,
      message: result.message
    };
    
  } catch (error) {
    Logger.log('❌ Error in exportToSchoolMIS_Average_Base64: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * =====================================================
 * � SYNC: เปรียบเทียบและ Sync ข้อมูล
 * =====================================================
 */

/**
 * เปรียบเทียบข้อมูลระหว่าง SchoolMIS CSV กับระบบของเรา
 * @param {string} csvContent - เนื้อหา CSV จาก SchoolMIS
 * @param {string} sheetName - ชื่อชีต
 * @param {string} term - ภาคเรียน
 * @returns {Object} ผลการเปรียบเทียบ
 */
function compareWithSchoolMIS(csvContent, sheetName, term) {
  try {
    Logger.log('=== เปรียบเทียบข้อมูล ===');
    
    // Parse CSV
    var csvData = parseSchoolMISCsv(csvContent);
    
    // ดึงข้อมูลจากระบบ
    var ss = SS();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('ไม่พบชีต: ' + sheetName);
    }
    
    var sheetData = sheet.getDataRange().getValues();
    var studentMap = buildStudentMapSync_(sheetData);
    var cols = getTermColumnsSync_(term);
    
    var comparison = {
      matched: [],
      differences: [],
      notInSystem: [],
      notInSchoolMIS: [],
      summary: {}
    };
    
    // เปรียบเทียบแต่ละนักเรียนจาก CSV
    csvData.forEach(function(csvStudent) {
      var studentCode = String(csvStudent.studentCode).trim();
      var idCard = String(csvStudent.idCard).trim();
      var fullName = (csvStudent.firstname + ' ' + csvStudent.lastname).trim();
      
      var rowIndex = findStudentRowSync_(studentMap, studentCode, idCard, fullName);
      
      if (rowIndex === -1) {
        comparison.notInSystem.push({
          code: studentCode,
          name: fullName,
          source: 'SchoolMIS'
        });
        return;
      }
      
      // เปรียบเทียบคะแนน
      var row = sheetData[rowIndex];
      var systemScores = {
        midterm: Number(row[cols.midTotal]) || 0,
        final: Number(row[cols.s10]) || 0,
        total: Number(row[cols.total]) || 0,
        grade: String(row[cols.grade] || '')
      };
      
      var schoolmisScores = {
        midterm: csvStudent.midtermTotal || 0,
        final: csvStudent.finalScore || 0,
        total: csvStudent.total || 0,
        grade: csvStudent.grade || ''
      };
      
      var hasDiff = (
        systemScores.midterm !== schoolmisScores.midterm ||
        systemScores.final !== schoolmisScores.final ||
        systemScores.total !== schoolmisScores.total ||
        systemScores.grade !== schoolmisScores.grade
      );
      
      if (hasDiff) {
        comparison.differences.push({
          code: studentCode,
          name: fullName,
          system: systemScores,
          schoolmis: schoolmisScores
        });
      } else {
        comparison.matched.push({
          code: studentCode,
          name: fullName
        });
      }
    });
    
    // สรุป
    comparison.summary = {
      totalInCSV: csvData.length,
      matched: comparison.matched.length,
      different: comparison.differences.length,
      notInSystem: comparison.notInSystem.length
    };
    
    Logger.log('สรุป: ตรงกัน ' + comparison.matched.length + ', ต่างกัน ' + comparison.differences.length + ', ไม่พบในระบบ ' + comparison.notInSystem.length);
    
    return comparison;
    
  } catch (error) {
    Logger.log('❌ Error in compareWithSchoolMIS: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * =====================================================
 * 🛠️ HELPER FUNCTIONS (ใช้ suffix Sync_ เพื่อไม่ชนกับไฟล์อื่น)
 * =====================================================
 */

/**
 * กำหนดคอลัมน์ตาม term
 * 
 * โครงสร้างชีตคะแนน:
 * - A (0): ลำดับ
 * - B (1): เลขประจำตัว
 * - C (2): ชื่อ-สกุล
 * 
 * ภาคเรียน 1:
 * - D-M (3-12): คะแนนย่อย 1-10 (10 ช่อง)
 * - N (13): คะแนนระหว่างภาค
 * - O (14): คะแนนปลายภาค
 * - P (15): รวมคะแนน
 * - Q (16): เกรด
 * 
 * ภาคเรียน 2:
 * - R-AA (17-26): คะแนนย่อย 11-20 (10 ช่อง)
 * - AB (27): คะแนนระหว่างภาค
 * - AC (28): คะแนนปลายภาค
 * - AD (29): รวมคะแนน
 * - AE (30): เกรด
 */
function getTermColumnsSync_(term) {
  // โครงสร้างตาม SchoolMIS (16 คอลัมน์ต่อภาค)
  // Per term: s1-s4, sum14, s5, makeup, s6-s9, sum69, midTotal, s10, total, grade
  if (term === 'term1') {
    return {
      s1: 3, s2: 4, s3: 5, s4: 6, sum14: 7,
      s5: 8, makeup: 9,
      s6: 10, s7: 11, s8: 12, s9: 13, sum69: 14,
      midTotal: 15,
      s10: 16,
      total: 17,
      grade: 18,
      // scoreSlots: 10 ช่อง ตรง 1:1 กับ SchoolMIS ครั้ง1-9 + ครั้ง10ปลายภาค
      scoreSlots: [3, 4, 5, 6, 8, 10, 11, 12, 13, 16],
      scoresCount: 10,
      mid: 15,
      final: 16
    };
  } else {
    return {
      s1: 19, s2: 20, s3: 21, s4: 22, sum14: 23,
      s5: 24, makeup: 25,
      s6: 26, s7: 27, s8: 28, s9: 29, sum69: 30,
      midTotal: 31,
      s10: 32,
      total: 33,
      grade: 34,
      // scoreSlots: 10 ช่อง ตรง 1:1 กับ SchoolMIS ครั้ง1-9 + ครั้ง10ปลายภาค
      scoreSlots: [19, 20, 21, 22, 24, 26, 27, 28, 29, 32],
      scoresCount: 10,
      mid: 31,
      final: 32
    };
  }
}

/**
 * สร้าง Map ของนักเรียนจากข้อมูลชีต
 */
function buildStudentMapSync_(sheetData) {
  var map = {
    byCode: {},
    byName: {},
    byIdCard: {}
  };
  
  // เริ่มจากแถวที่ 5 (index 4)
  for (var i = 4; i < sheetData.length; i++) {
    var row = sheetData[i];
    var code = String(row[1] || '').trim();  // เลขประจำตัว
    var name = String(row[2] || '').trim();  // ชื่อ-สกุล
    
    if (code) {
      map.byCode[code] = i;
    }
    if (name) {
      map.byName[name.toLowerCase()] = i;
    }
  }
  
  return map;
}

/**
 * สร้าง Map ของ idCard จากชีต Students
 */
function buildStudentIdCardMapSync_(studentsData) {
  var map = {};
  
  if (studentsData.length <= 1) return map;
  
  var headers = studentsData[0];
  var idCol = headers.indexOf('student_id');
  var cardCol = headers.indexOf('id_card');
  
  if (idCol === -1 || cardCol === -1) return map;
  
  for (var i = 1; i < studentsData.length; i++) {
    var studentId = String(studentsData[i][idCol] || '').trim();
    var idCard = String(studentsData[i][cardCol] || '').trim();
    if (studentId && idCard) {
      map[studentId] = idCard;
    }
  }
  
  return map;
}

/**
 * หาแถวของนักเรียนในชีต
 */
function findStudentRowSync_(studentMap, code, idCard, name) {
  // 1. ลองหาจากรหัส
  if (code && studentMap.byCode[code] !== undefined) {
    return studentMap.byCode[code];
  }
  
  // 2. ลองหาจากชื่อ (ตรงทั้งหมด)
  if (name && studentMap.byName[name.toLowerCase()] !== undefined) {
    return studentMap.byName[name.toLowerCase()];
  }
  
  // 3. ลองหาจากชื่อ (partial match)
  if (name) {
    var nameLower = name.toLowerCase();
    var mapEntries = Object.keys(studentMap.byName);
    for (var k = 0; k < mapEntries.length; k++) {
      var mapName = mapEntries[k];
      if (mapName.indexOf(nameLower) !== -1 || nameLower.indexOf(mapName) !== -1) {
        return studentMap.byName[mapName];
      }
    }
  }
  
  return -1;
}

/**
 * Parse CSV line รองรับ quotes
 */
function parseCSVLineSync_(line) {
  var result = [];
  var current = '';
  var inQuotes = false;
  
  for (var i = 0; i < line.length; i++) {
    var char = line[i];
    
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  result.push(current.trim());
  return result;
}

/**
 * Parse ตัวเลข
 */
function parseNumberSync_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  var num = Number(value);
  return isNaN(num) ? null : num;
}

/**
 * รวมคะแนน
 */
function sumScoresSync_(scores) {
  var sum = 0;
  for (var i = 0; i < scores.length; i++) {
    sum += Number(scores[i]) || 0;
  }
  return sum;
}

/**
 * แยกชื่อ-นามสกุลภาษาไทย
 */
function splitThaiNameSync_(fullName) {
  // ลบคำนำหน้า
  var prefixes = ['เด็กชาย', 'เด็กหญิง', 'ด.ช.', 'ด.ญ.', 'นาย', 'นางสาว', 'น.ส.', 'นาง'];
  var name = fullName;
  
  for (var p = 0; p < prefixes.length; p++) {
    if (name.indexOf(prefixes[p]) === 0) {
      name = name.substring(prefixes[p].length).trim();
      break;
    }
  }
  
  // แยกชื่อ-นามสกุล
  var parts = name.split(/\s+/);
  
  if (parts.length >= 2) {
    return {
      firstname: parts[0],
      lastname: parts.slice(1).join(' ')
    };
  } else {
    return {
      firstname: name,
      lastname: ''
    };
  }
}

/**
 * Escape ค่าสำหรับ CSV
 */
function escapeCSVValueSync_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  
  var str = String(value);
  
  // ถ้ามี comma, quotes, หรือ newline ให้ครอบด้วย quotes
  if (str.indexOf(',') !== -1 || str.indexOf('"') !== -1 || str.indexOf('\n') !== -1) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  
  return str;
}

/**
 * =====================================================
 * 📱 UI FUNCTIONS (สำหรับเรียกจาก HTML)
 * =====================================================
 */

/**
 * ดึงรายการชีตคะแนนสำหรับ dropdown แยกตามชั้น/ห้อง
 * ชื่อชีต pattern: "{วิชา} {ชั้น}-{ห้อง}" เช่น "ภาษาไทย ป3-1"
 * @returns {Object} { grades: [...], sheets: [...] }
 */
function getScoreSheetsForSync() {
  try {
    var ss = SS();
    var sheets = ss.getSheets();
    var scoreSheets = [];
    var gradeSet = {};
    
    var excludeSheets = ['Students', 'รายวิชา', 'Users', 'Settings', 'Attendance', 'global_settings', 'วันหยุด'];
    
    sheets.forEach(function(sheet) {
      var name = sheet.getName();
      
      if (excludeSheets.indexOf(name) !== -1 || name.indexOf('TMP_') === 0) {
        return;
      }
      
      // ตรวจสอบว่าเป็นชีตคะแนนหรือไม่
      try {
        var cell_A1 = sheet.getRange(1, 1).getValue();
        if (cell_A1 === 'รหัสวิชา') {
          var subjectCode = sheet.getRange(1, 2).getValue();
          var subjectName = sheet.getRange(2, 2).getValue();
          
          // Parse ชั้น/ห้องจากชื่อชีต: "{วิชา} {ชั้น}-{ห้อง}"
          var gradeClass = '';
          var match = name.match(/\s+(ป\d+)-(\d+)$/);
          if (match) {
            gradeClass = match[1] + '/' + match[2]; // เช่น "ป3/1"
          } else {
            // fallback: ลองหา pattern อื่น
            var match2 = name.match(/\s+(.+)-(\d+)$/);
            if (match2) {
              gradeClass = match2[1] + '/' + match2[2];
            }
          }
          
          if (gradeClass) {
            gradeSet[gradeClass] = true;
          }
          
          scoreSheets.push({
            sheetName: name,
            subjectCode: subjectCode,
            subjectName: subjectName,
            gradeClass: gradeClass
          });
        }
      } catch (e) {
        // ข้ามชีตที่มีปัญหา
      }
    });
    
    // เรียงลำดับชั้น
    var grades = Object.keys(gradeSet).sort(function(a, b) {
      var numA = parseInt(a.replace(/[^\d]/g, '')) || 0;
      var numB = parseInt(b.replace(/[^\d]/g, '')) || 0;
      return numA - numB;
    });
    
    return { grades: grades, sheets: scoreSheets };
    
  } catch (error) {
    Logger.log('❌ Error in getScoreSheetsForSync: ' + error.message);
    return { grades: [], sheets: [] };
  }
}

/**
 * Import จาก CSV ที่อัพโหลดจาก Frontend — ใส่ทั้ง 2 ภาคเท่ากัน
 */
function importSchoolMISFromUpload(base64Content, sheetName, term) {
  try {
    // Decode base64
    var csvContent = Utilities.newBlob(Utilities.base64Decode(base64Content)).getDataAsString('UTF-8');
    
    // นำเข้าข้อมูลทั้ง 2 ภาค
    var result1 = importFromSchoolMIS(csvContent, sheetName, 'term1');
    var result2 = importFromSchoolMIS(csvContent, sheetName, 'term2');
    
    if (!result1.success && !result2.success) {
      return { success: false, error: 'นำเข้าทั้ง 2 ภาคไม่สำเร็จ: ' + (result1.error || result2.error) };
    }
    
    var imported = Math.max(result1.imported || 0, result2.imported || 0);
    var skipped = Math.max(result1.skipped || 0, result2.skipped || 0);
    var notFound = result1.notFound || result2.notFound || [];
    
    return {
      success: true,
      imported: imported,
      skipped: skipped,
      notFound: notFound,
      message: '✅ นำเข้าทั้ง 2 ภาคเรียนสำเร็จ ' + imported + ' คน, ข้าม ' + skipped + ' คน'
    };
    
  } catch (error) {
    Logger.log('❌ Error in importSchoolMISFromUpload: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}
