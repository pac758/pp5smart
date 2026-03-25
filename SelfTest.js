/**
 * Self-test สำหรับตรวจสอบฟังก์ชันหลักที่แก้ไขล่าสุด
 * วิธีใช้: เปิด Apps Script Editor → เลือก runSelfTest → กด Run
 */
/**
 * ค้นหานักเรียนรหัส 763 (เด็กชายจิรายุ การชารี) ในทุกชีต
 * วิธีใช้: Apps Script Editor → เลือก findStudent763 → Run → ดูผลใน Log
 */
/**
 * ลบชีต BACKUP_ ทั้งหมดออกจาก Spreadsheet
 * วิธีใช้: Apps Script Editor → เลือก deleteAllBackupSheets → Run → ดูผลใน Log
 */
/**
 * ตรวจสอบการนับวันเรียนจริงของแต่ละเดือน (ปีการศึกษา 2568)
 * วิธีใช้: Apps Script Editor → เลือก debugSchoolDays → Run → ดูผลใน Log
 */
function debugSchoolDays() {
  var ss = SS();
  var monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  var thaiDays = ['อา','จ','อ','พ','พฤ','ศ','ส'];

  // ปีการศึกษา 2568: พ.ค.2568 - มี.ค.2569
  var sheetsToCheck = [
    {name: 'พฤษภาคม2568', label: 'พ.ค.68'},
    {name: 'มิถุนายน2568', label: 'มิ.ย.68'},
    {name: 'กรกฎาคม2568', label: 'ก.ค.68'},
    {name: 'สิงหาคม2568', label: 'ส.ค.68'},
    {name: 'กันยายน2568', label: 'ก.ย.68'},
    {name: 'ตุลาคม2568', label: 'ต.ค.68'},
    {name: 'พฤศจิกายน2568', label: 'พ.ย.68'},
    {name: 'ธันวาคม2568', label: 'ธ.ค.68'},
    {name: 'มกราคม2569', label: 'ม.ค.69'},
    {name: 'กุมภาพันธ์2569', label: 'ก.พ.69'},
    {name: 'มีนาคม2569', label: 'มี.ค.69'}
  ];

  sheetsToCheck.forEach(function(info) {
    var sheet = ss.getSheetByName(info.name);
    if (!sheet) {
      Logger.log('❌ ' + info.label + ': ไม่พบชีต ' + info.name);
      return;
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // หา date columns
    var dateCols = [];
    for (var i = 0; i < headers.length; i++) {
      var m = String(headers[i]).match(/^(\d{1,2})/);
      if (m) dateCols.push({col: i, day: parseInt(m[1]), header: String(headers[i])});
    }

    // นับวันที่มีข้อมูล
    var schoolDaysSet = {};
    for (var r = 1; r < data.length; r++) {
      dateCols.forEach(function(dc) {
        var v = String(data[r][dc.col] || '').trim();
        if (v === '/' || v === 'ม' || v === 'ล' || v === 'ข' || v === '0' || v === '1' || v.toLowerCase() === 'l') {
          schoolDaysSet[dc.day] = (schoolDaysSet[dc.day] || 0) + 1;
        }
      });
    }

    var schoolDays = Object.keys(schoolDaysSet).map(Number).sort(function(a,b){return a-b;});
    var totalStudentRows = data.length - 1;

    Logger.log('=== ' + info.label + ' (' + info.name + ') ===');
    Logger.log('  คอลัมน์วันที่ทั้งหมด: ' + dateCols.length + ' คอลัมน์');
    Logger.log('  แถวนักเรียน: ' + totalStudentRows + ' แถว');
    Logger.log('  วันเรียนจริง (มีเช็คชื่อ): ' + schoolDays.length + ' วัน');
    Logger.log('  วันที่: ' + schoolDays.join(', '));

    // แสดงวันที่ไม่มีข้อมูล
    var allDays = dateCols.map(function(dc){return dc.day;});
    var nonSchoolDays = allDays.filter(function(d){return schoolDays.indexOf(d) === -1;});
    if (nonSchoolDays.length > 0) {
      Logger.log('  วันที่ไม่มีเช็คชื่อ: ' + nonSchoolDays.join(', '));
    }

    // ตรวจหาวันเสาร์-อาทิตย์ที่มีข้อมูล (อาจเป็นข้อมูลผิด)
    var yearCE = info.name.match(/(\d{4})/);
    if (yearCE) {
      var y = parseInt(yearCE[1]) - 543;
      var mIdx = monthNames.indexOf(info.name.replace(/\d+/, ''));
      if (mIdx >= 0) {
        var weekendWithData = [];
        schoolDays.forEach(function(day) {
          var d = new Date(y, mIdx, day);
          var dow = d.getDay();
          if (dow === 0 || dow === 6) {
            weekendWithData.push(day + '(' + thaiDays[dow] + ') มี' + schoolDaysSet[day] + 'คน');
          }
        });
        if (weekendWithData.length > 0) {
          Logger.log('  ⚠️ วันหยุด(ส./อา.)ที่มีข้อมูล: ' + weekendWithData.join(', '));
        }
      }
    }
  });
}

function deleteAllBackupSheets() {
  var ss = SS();
  var sheets = ss.getSheets();
  var deleted = 0;
  var names = [];

  // นับก่อน
  var backupSheets = sheets.filter(function(s) { return s.getName().startsWith('BACKUP_'); });
  Logger.log('พบชีต BACKUP_ จำนวน ' + backupSheets.length + ' ชีต');

  backupSheets.forEach(function(sheet) {
    var name = sheet.getName();
    try {
      ss.deleteSheet(sheet);
      deleted++;
      names.push(name);
      Logger.log('🗑️ ลบ: ' + name);
    } catch(e) {
      Logger.log('❌ ลบไม่ได้: ' + name + ' — ' + e.message);
    }
  });

  var msg = '✅ ลบชีต BACKUP_ สำเร็จ ' + deleted + ' ชีต';
  Logger.log(msg);
  Logger.log('รายชื่อที่ลบ: ' + names.join(', '));
  return msg;
}

function findStudent763() {
  var ss = SS();
  var results = [];
  var searchId = '763';
  var searchName = 'จิรายุ';

  // 1. ค้นหาในชีต Students
  var studSheet = ss.getSheetByName('Students');
  if (studSheet) {
    var data = studSheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf('student_id');
    var fnCol = headers.indexOf('firstname');
    var lnCol = headers.indexOf('lastname');
    var grCol = headers.indexOf('grade');
    var clCol = headers.indexOf('class_no');
    var found = false;
    for (var i = 1; i < data.length; i++) {
      var sid = String(data[i][idCol] || '').trim();
      var fn = String(data[i][fnCol] || '');
      if (sid === searchId || fn.indexOf(searchName) !== -1) {
        results.push('✅ Students แถว ' + (i+1) + ': id=' + sid + ' ชื่อ=' + fn + ' ' + data[i][lnCol] + ' ชั้น=' + data[i][grCol] + '/' + data[i][clCol]);
        found = true;
      }
    }
    if (!found) results.push('❌ ไม่พบรหัส ' + searchId + ' หรือชื่อ ' + searchName + ' ในชีต Students (' + (data.length-1) + ' แถว)');
  }

  // 2. ค้นหาในชีตคะแนนทั้งหมด (ป.6)
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var name = sh.getName();
    if (name.indexOf('ป6') === -1 && name.indexOf('ป.6') === -1) continue;
    if (name.indexOf('BACKUP') !== -1) continue;
    try {
      var d = sh.getDataRange().getValues();
      for (var r = 0; r < d.length; r++) {
        for (var c = 0; c < Math.min(d[r].length, 5); c++) {
          var val = String(d[r][c] || '');
          if (val === searchId || val.indexOf(searchName) !== -1) {
            results.push('📋 ชีต "' + name + '" แถว ' + (r+1) + ' col ' + (c+1) + ': ' + val);
          }
        }
      }
    } catch(e) {}
  }

  // 3. ค้นหาใน SCORES_WAREHOUSE
  var wh = ss.getSheetByName('SCORES_WAREHOUSE');
  if (wh) {
    var wd = wh.getDataRange().getValues();
    for (var r = 0; r < wd.length; r++) {
      var row = wd[r];
      for (var c = 0; c < Math.min(row.length, 5); c++) {
        var v = String(row[c] || '');
        if (v === searchId || v.indexOf(searchName) !== -1) {
          results.push('🏪 SCORES_WAREHOUSE แถว ' + (r+1) + ' col ' + (c+1) + ': ' + JSON.stringify(row.slice(0,6)));
          break;
        }
      }
    }
  }

  Logger.log('=== ค้นหานักเรียน 763 จิรายุ ===');
  if (results.length === 0) {
    Logger.log('❌ ไม่พบนักเรียนรหัส 763 ในระบบเลย — อาจถูกลบหรือไม่เคยนำเข้า');
  } else {
    results.forEach(function(r) { Logger.log(r); });
  }
  Logger.log('================================');
  return results.join('\n') || 'ไม่พบข้อมูล';
}

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
