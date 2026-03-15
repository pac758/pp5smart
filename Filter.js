/**
 * ✅ [ฟังก์ชันใหม่แบบง่าย] ดึงรายการระดับชั้นทั้งหมดที่มีในระบบ
 * ฟังก์ชันนี้ถูกออกแบบมาให้ทำงานอย่างเรียบง่ายและแน่นอนที่สุด
 */
function getGradeList() {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    
    if (!sheet) {
      Logger.log("❌ ไม่พบชีต 'Students'");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("⚠️ ชีต 'Students' ไม่มีข้อมูลนักเรียน");
      return [];
    }

    // ระบุว่าข้อมูลระดับชั้นอยู่ที่คอลัมน์ F (index ที่ 5)
    const gradeColIndex = 5; 

    const grades = data.slice(1)
      .map(row => row[gradeColIndex])
      .filter(grade => grade) 
      .map(grade => String(grade).trim());

    const uniqueGrades = [...new Set(grades)].sort();
    
    Logger.log(`✅ พบระดับชั้น: ${uniqueGrades.join(', ')}`);
    return uniqueGrades;
    
  } catch (error) {
    Logger.log(`❌ เกิดข้อผิดพลาดใน getGradeList: ${error.message}`);
    return [];
  }
}