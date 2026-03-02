// ============================================================
// 📅 HOLIDAY SERVICE — Server-side functions for holiday management
// ย้ายจาก Holidays.html (ซึ่งเป็น .html ทำให้ google.script.run เรียกไม่ได้)
// ============================================================

/**
 * 🟢 ฟังก์ชันที่หน้าเว็บเรียกใช้เพื่อขอดึงวันหยุด
 */
function getSpecialHolidays(year) {
  var rawData = loadHolidaysFromSheet();
  var holidayMap = {};
  
  rawData.forEach(function(row) {
    var dateStr = row[0];
    var name = row[1];
    holidayMap[dateStr] = name;
  });
  
  Logger.log("ส่งข้อมูลวันหยุดไปหน้าเว็บ: " + JSON.stringify(holidayMap));
  return holidayMap;
}

/**
 * 🧪 TEST: รันจาก Apps Script editor เพื่อดูข้อมูลจริงในชีต
 */
function debugHolidaySheet() {
  const ss = SS();
  const allSheets = ss.getSheets().map(s => s.getName());
  Logger.log('📋 ชีตทั้งหมด: ' + JSON.stringify(allSheets));
  
  const sheet = ss.getSheetByName("วันหยุด");
  if (!sheet) {
    Logger.log('❌ ไม่พบชีต "วันหยุด"');
    return;
  }
  
  const values = sheet.getDataRange().getValues();
  Logger.log('📅 Rows: ' + values.length + ', Cols: ' + (values[0] ? values[0].length : 0));
  Logger.log('📅 Header: ' + JSON.stringify(values[0]));
  
  for (let i = 0; i < Math.min(values.length, 10); i++) {
    const row = values[i];
    const dateVal = row[0];
    const typeOf = dateVal instanceof Date ? 'Date' : typeof dateVal;
    Logger.log(`Row ${i}: [${typeOf}] "${dateVal}" | "${row[1]}" | "${row[2] || ''}"`);
  }
  
  // ทดสอบ loadHolidaysFromSheet
  const result = loadHolidaysFromSheet();
  Logger.log('🎯 loadHolidaysFromSheet result: ' + JSON.stringify(result));
}

/**
 * 📥 โหลดข้อมูลวันหยุดจาก Sheet
 */
function loadHolidaysFromSheet() {
  try {
    const ss = SS();
    Logger.log('📅 loadHolidaysFromSheet: SS() OK, id=' + ss.getId());
    
    const sheet = ss.getSheetByName("วันหยุด");
    
    if (!sheet) {
      Logger.log('❌ ไม่พบชีต "วันหยุด"');
      const allSheets = ss.getSheets().map(s => s.getName());
      Logger.log('📋 ชีตทั้งหมด: ' + JSON.stringify(allSheets));
      return [];
    }

    const values = sheet.getDataRange().getValues();
    Logger.log('📅 Holidays sheet rows: ' + values.length + ', header: ' + JSON.stringify(values[0]));

    const holidays = values.slice(1).map(row => {
      const dateValue = row[0];
      const detail = row[1];

      if (dateValue && detail) {
        let dateStr = '';
        if (dateValue instanceof Date) {
          dateStr = Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          dateStr = String(dateValue).trim();
        }
        return [dateStr, String(detail).trim()];
      }
      return null;
    }).filter(item => item !== null);

    Logger.log('✅ loadHolidaysFromSheet: loaded ' + holidays.length + ' holidays');
    if (holidays.length > 0) {
      Logger.log('📌 First 3: ' + JSON.stringify(holidays.slice(0, 3)));
    }
    return holidays;

  } catch (e) {
    Logger.log("❌ Error in loadHolidaysFromSheet: " + e.message);
    return [];
  }
}

/**
 * 💾 บันทึกข้อมูลวันหยุดลง Sheet
 */
function saveHolidaySheet(holidayData) {
  try {
    const ss = SS();
    let sheet = ss.getSheetByName("วันหยุด");

    if (!sheet) {
      sheet = ss.insertSheet("วันหยุด");
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
    }

    sheet.getRange("A1:B1")
      .setValues([["วันที่", "รายละเอียด"]])
      .setFontWeight("bold")
      .setBackground("#e0e0e0");
    
    const validData = (holidayData || [])
      .filter(row => row && row[0] && row[1]) 
      .map(row => {
        let d = row[0];
        if (d instanceof Date) {
          d = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return [String(d).trim(), String(row[1] || "").trim()];
      });

    if (validData.length > 0) {
      validData.sort((a, b) => a[0].localeCompare(b[0]));
      
      sheet.getRange(2, 1, validData.length, 2).setValues(validData);
      sheet.getRange(2, 1, validData.length, 1).setNumberFormat("yyyy-MM-dd");
      sheet.autoResizeColumn(1);
      sheet.autoResizeColumn(2);
    }

    return { success: true, message: "บันทึกข้อมูลวันหยุดเรียบร้อย (" + validData.length + " รายการ)" };
  } catch (e) {
    Logger.log("Error saving holidays: " + e.message);
    throw new Error("บันทึกไม่สำเร็จ: " + e.message);
  }
}
