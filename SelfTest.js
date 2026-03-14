/**
 * Self-test สำหรับตรวจสอบฟังก์ชันหลักที่แก้ไขล่าสุด
 * วิธีใช้: เปิด Apps Script Editor → เลือก runSelfTest → กด Run
 */
function runSelfTest() {
  const results = [];

  function pass(name) { results.push('✅ ' + name); }
  function fail(name, err) { results.push('❌ ' + name + ': ' + err); }

  // --- 1. Students sheet exists ---
  try {
    const sheet = SS().getSheetByName('Students');
    if (!sheet) throw new Error('ไม่พบ sheet Students');
    pass('Students sheet exists');

    // --- 2. Headers มี birthdate, weight, height ---
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const required = ['student_id', 'id_card', 'firstname', 'lastname', 'grade', 'class_no', 'gender', 'birthdate', 'weight', 'height'];
    required.forEach(h => {
      if (headers.indexOf(h) === -1) fail('header: ' + h, 'ไม่พบ');
      else pass('header: ' + h);
    });

    // --- 3. getFilteredStudentsInline คืนค่าถูกต้อง ---
    const students = getFilteredStudentsInline('', '');
    if (!students || students.length === 0) {
      fail('getFilteredStudentsInline', 'คืนค่าว่าง');
    } else {
      const s = students[0];
      const fields = ['id', 'firstname', 'lastname', 'grade', 'classNo', 'gender', 'birthdate', 'age', 'weight', 'height', 'father_name', 'mother_name', 'address'];
      fields.forEach(f => {
        if (!(f in s)) fail('student field: ' + f, 'ไม่มีใน object');
        else pass('student field: ' + f);
      });
      pass('getFilteredStudentsInline returns ' + students.length + ' students');
    }
  } catch (e) {
    fail('Students sheet check', e.message);
  }

  // --- 4. parseThaiBirthdate (ทดสอบภายใน importCsvStudents) ---
  try {
    const testCsv = [
      'col0,col1,1234567890123,ป.1,1,TEST001,ช,เด็กชาย,ทดสอบ,ระบบ,24/03/2565,3,25,110,-,พุทธ,ไทย,ไทย,1,1,-,ต.ทดสอบ,อ.ทดสอบ,จ.ทดสอบ,ผู้ปก,ครอง,รับจ้าง,บิดา,พิทักษ์,นาม,รับจ้าง,มารดา,นาม,รับจ้าง',
      'header1,header2,header3,header4,header5,header6,header7,header8,header9,header10,header11,header12,header13,header14,header15,header16,header17,header18,header19,header20,header21,header22,header23,header24,header25,header26,header27,header28,header29,header30,header31,header32,header33,header34',
      '31030002,โรงเรียน,1319800689643,ป.6,1,763,ช,เด็กชาย,ทดสอบ,ระบบ,13/09/2556,12,32,147,-,พุทธ,ไทย,ไทย,1,1,-,ต.ทดสอบ,อ.ทดสอบ,จ.ทดสอบ,มารดา,นามสกุล,รับจ้าง,มารดา,บิดา,นามสกุล,รับจ้าง,มารดา,นามสกุล,รับจ้าง,เด็กยากจน,-'
    ].join('\n');

    const result = importCsvStudents(testCsv);
    if (result && result.includes('✅')) pass('importCsvStudents: ทำงานโดยไม่ error');
    else fail('importCsvStudents', 'ผลลัพธ์ไม่ถูกต้อง: ' + result);
  } catch (e) {
    fail('importCsvStudents', e.message);
  }

  // --- 5. exportStudentRegistryPDF ไม่ crash (ไม่ generate จริง แค่เช็ค function exists) ---
  try {
    if (typeof exportStudentRegistryPDF !== 'function') throw new Error('ไม่พบฟังก์ชัน');
    pass('exportStudentRegistryPDF: function exists');
  } catch (e) {
    fail('exportStudentRegistryPDF', e.message);
  }

  // --- สรุปผล ---
  const passed = results.filter(r => r.startsWith('✅')).length;
  const failed = results.filter(r => r.startsWith('❌')).length;
  Logger.log('=== SELF TEST RESULTS ===');
  results.forEach(r => Logger.log(r));
  Logger.log('=========================');
  Logger.log(`สรุป: ผ่าน ${passed} / ไม่ผ่าน ${failed}`);

  return results.join('\n') + `\n\nสรุป: ผ่าน ${passed} / ไม่ผ่าน ${failed}`;
}
