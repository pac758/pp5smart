/**
 * ดึงสถานะการจบการศึกษาสำหรับ UI
 * @returns {Object} สถานะปัจจุบัน
 */
function getGraduationStatus() {
  try {
    var settings = S_getGlobalSettings();
    var currentYear = settings['ปีการศึกษา'] || 'ไม่ทราบ';
    
    // ดึงข้อมูลนักเรียน ป.6 ทั้งหมด
    var ss = SS();
    var sheet = AY_getStudentsSheetForRead();
    if (!sheet) {
      return {
        currentYear: currentYear,
        totalP6: 0,
        activeP6: 0,
        graduatedP6: 0
      };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {
        currentYear: currentYear,
        totalP6: 0,
        activeP6: 0,
        graduatedP6: 0
      };
    }
    
    var headers = data[0];
    var col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });
    
    var gradeIdx = col['grade'];
    var statusIdx = col['status'];
    
    var totalP6 = 0, activeP6 = 0, graduatedP6 = 0;
    
    for (var r = 1; r < data.length; r++) {
      var grade = String(data[r][gradeIdx] || '').trim();
      var status = String(data[r][statusIdx] || '').trim();
      
      if (grade === 'ป.6') {
        totalP6++;
        if (status === 'จำหน่าย') {
          graduatedP6++;
        } else {
          activeP6++;
        }
      }
    }
    
    return {
      currentYear: currentYear,
      totalP6: totalP6,
      activeP6: activeP6,
      graduatedP6: graduatedP6
    };
    
  } catch (e) {
    Logger.log('getGraduationStatus error: ' + e.message);
    return {
      currentYear: 'ผิดพลาด',
      totalP6: 0,
      activeP6: 0,
      graduatedP6: 0
    };
  }
}

/**
 * ดูตัวอย่างการเปลี่ยนแปลงก่อนจำหน่าย/เปลี่ยนปี
 * @returns {Object} ตัวอย่างสิ่งที่จะเกิดขึ้น
 */
function previewGraduation() {
  try {
    var settings = S_getGlobalSettings();
    var currentYear = settings['ปีการศึกษา'] || ((typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : String(S_getCurrentAcademicYear_()));
    var newYear = Number(currentYear) + 1;
    
    // ดึงข้อมูลนักเรียน
    var ss = SS();
    var sheet = AY_getStudentsSheetForRead();
    if (!sheet) {
      return { success: false, message: 'ไม่พบชีต Students' };
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, message: 'ไม่มีข้อมูลนักเรียน' };
    }
    
    var headers = data[0];
    var col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });
    
    var gradeIdx = col['grade'];
    var statusIdx = col['status'];
    var studentIdIdx = col['student_id'];
    var firstnameIdx = col['firstname'];
    var lastnameIdx = col['lastname'];
    
    var willGraduate = [];
    var willPromote = 0;
    
    for (var r = 1; r < data.length; r++) {
      var grade = String(data[r][gradeIdx] || '').trim();
      var status = String(data[r][statusIdx] || '').trim();
      var studentId = String(data[r][studentIdIdx] || '').trim();
      var firstname = String(data[r][firstnameIdx] || '').trim();
      var lastname = String(data[r][lastnameIdx] || '').trim();
      var name = firstname + ' ' + lastname;
      
      // นักเรียนที่จะจำหน่าย
      if ((grade === 'ป.6' || grade === 'ม.3') && (status === 'active' || status === '')) {
        willGraduate.push({
          studentId: studentId,
          name: name,
          grade: grade
        });
      }
      
      // นักเรียนที่จะเลื่อนชั้น (ทุกคนที่ไม่ใช่ ป.6 และ ม.3)
      if (grade !== 'ป.6' && grade !== 'ม.3' && (status === 'active' || status === '')) {
        willPromote++;
      }
    }
    
    // ชีทที่จะสร้างใหม่
    var yearlyBases = typeof S_YEARLY_SHEETS !== 'undefined' ? S_YEARLY_SHEETS : [
      'SCORES_WAREHOUSE',
      'การประเมินอ่านคิดเขียน',
      'การประเมินคุณลักษณะ',
      'การประเมินกิจกรรมพัฒนาผู้เรียน',
      'การประเมินสมรรถนะ',
      'AttendanceLog',
      'ความเห็นครู'
    ];
    
    var willCreateSheets = yearlyBases.map(function(base) {
      return base + '_' + newYear;
    });
    
    return {
      success: true,
      willGraduate: willGraduate.slice(0, 10), // แสดงแค่ 10 คนแรก
      willPromote: willPromote,
      willCreateSheets: willCreateSheets,
      totalGraduate: willGraduate.length
    };
    
  } catch (e) {
    Logger.log('previewGraduation error: ' + e.message);
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}
