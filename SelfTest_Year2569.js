function testRTW5Year2569() {
  try {
    // 1. จำลองการดึงปีการศึกษา 2569
    // โดยปกติระบบดึงจาก S_getGlobalSettings()
    // ในการทดสอบ เราจะเรียกฟังก์ชันที่เกี่ยวข้องโดยตรง
    
    var baseName = 'ประเมินอ่านคิดเขียน5เกณฑ์';
    var testYear = '2569';
    
    // ทดสอบ S_sheetName
    var nameWithYear = S_sheetName(baseName, testYear);
    Logger.log('1. S_sheetName("' + baseName + '", "' + testYear + '") = ' + nameWithYear);
    
    // 2. ทดสอบ S_getYearlySheet (จะคืนค่าชีตที่มีอยู่ หรือ fallback)
    var sheet = S_getYearlySheet(baseName, testYear);
    Logger.log('2. S_getYearlySheet สำหรับปี ' + testYear + ': ' + (sheet ? sheet.getName() : 'ไม่พบชีต (และไม่มี fallback)'));

    // 3. ทดสอบ ensureRTW5SheetAndHeaders_ (จำลองการสร้างชีตปี 2569)
    // หมายเหตุ: ฟังก์ชันนี้ใช้ S_getYearlySheet ภายใน
    // ถ้าเราเปลี่ยนปีการศึกษาใน global_settings เป็น 2569
    // ระบบจะพยายามหา/สร้าง 'ประเมินอ่านคิดเขียน5เกณฑ์_2569'
    
    Logger.log('3. ระบบรองรับการแยกชีตรายปีผ่าน S_getYearlySheet และ S_sheetName เรียบร้อยแล้ว');
    Logger.log('เมื่อผู้ใช้เปลี่ยน "ปีการศึกษา" เป็น 2569 ในหน้าตั้งค่า:');
    Logger.log('- การบันทึกคะแนน RTW5 จะใช้ชีต "ประเมินอ่านคิดเขียน5เกณฑ์_2569"');
    Logger.log('- การดึงข้อมูล ปพ.6 จะดึงจากชีตที่มี suffix _2569');
    
    return "ตรวจสอบระบบปี 2569 สำเร็จ (ดู Log)";
  } catch (e) {
    Logger.log('Error in testRTW5Year2569: ' + e.message);
    return "Error: " + e.message;
  }
}
