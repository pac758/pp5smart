// ============================================================
// 📖 PP5 FULL BOOK — Google Sheets → PDF Version
// สร้างแต่ละส่วนใน temp Google Sheets แล้ว export PDF → merge ด้วย pdf-lib
// ============================================================

// ============================================================
// � PROGRESS: ระบบแจ้งสถานะสร้าง PDF แบบ real-time
// ============================================================
function pp5fb_setProgress_(pct, step, detail) {
  try {
    var data = JSON.stringify({ pct: pct, step: step, detail: detail || '', ts: Date.now() });
    CacheService.getUserCache().put('pp5fb_progress', data, 300);
  } catch (e) {}
}

function pp5fb_getProgress() {
  try {
    var raw = CacheService.getUserCache().get('pp5fb_progress');
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return { pct: 0, step: 0, detail: '' };
}

// ============================================================
// � UTILITY: สร้าง temp spreadsheet, export PDF, merge PDFs
// ============================================================

/**
 * Export ทั้ง Spreadsheet (ทุก sheet) เป็น PDF blob รวดเดียว — เร็วมาก (1 request)
 */
function pp5s_exportAllAsPdf_(ss, opts) {
  opts = opts || {};
  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  var ssId = ss.getId();
  var portrait = opts.portrait !== false;
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?'
    + 'format=pdf&'
    + 'size=A4&'
    + 'portrait=' + portrait + '&'
    + 'fitw=' + (opts.fitw !== false ? 'true' : 'false') + '&'
    + 'gridlines=false&'
    + 'printtitle=false&'
    + 'sheetnames=false&'
    + 'pagenum=false&'
    + 'fzr=false&'
    + 'left_margin=' + (opts.left || 1.2) + '&'
    + 'right_margin=' + (opts.right || 0.3) + '&'
    + 'top_margin=' + (opts.top || 0.4) + '&'
    + 'bottom_margin=' + (opts.bottom || 0.3);
  // ไม่ระบุ gid = export ทุก sheet
  var token = ScriptApp.getOAuthToken();
  return UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();
}

/**
 * Export ทีละ sheet เป็น PDF blob (แต่ละ sheet fit A4 ของตัวเอง)
 */
function pp5s_exportSheetAsPdf_(ss, sheet, opts) {
  opts = opts || {};
  SpreadsheetApp.flush();
  var ssId = ss.getId();
  var gid = sheet.getSheetId();
  var portrait = opts.portrait !== false;
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?'
    + 'format=pdf&'
    + 'size=A4&'
    + 'portrait=' + portrait + '&'
    + 'fitw=true&'
    + 'gridlines=false&'
    + 'printtitle=false&'
    + 'sheetnames=false&'
    + 'pagenum=false&'
    + 'fzr=false&'
    + 'left_margin=' + (opts.left || 1.2) + '&'
    + 'right_margin=' + (opts.right || 0.3) + '&'
    + 'top_margin=' + (opts.top || 0.4) + '&'
    + 'bottom_margin=' + (opts.bottom || 0.3) + '&'
    + 'gid=' + gid;
  var token = ScriptApp.getOAuthToken();
  return UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();
}

/**
 * Export ทุก sheet เป็น PDF blobs แยกกัน (พร้อม rate-limit handling)
 */
function pp5s_exportAllSheetsAsPdfs_(ss, opts) {
  opts = opts || {};
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  var sheets = ss.getSheets();
  var blobs = [];
  for (var i = 0; i < sheets.length; i++) {
    var blob = null;
    var maxRetries = 3;
    for (var attempt = 0; attempt < maxRetries; attempt++) {
      try {
        blob = pp5s_exportSheetAsPdf_(ss, sheets[i], opts);
        break;
      } catch (e) {
        if (e.message && e.message.indexOf('429') !== -1 && attempt < maxRetries - 1) {
          var wait = (attempt + 1) * 3000; // 3s, 6s, 9s
          Logger.log('  ⏳ 429 retry ' + (attempt+1) + ' for ' + sheets[i].getName() + ' (wait ' + (wait/1000) + 's)');
          Utilities.sleep(wait);
        } else {
          Logger.log('  ⚠️ skip sheet ' + sheets[i].getName() + ': ' + e.message);
          break;
        }
      }
    }
    if (blob) {
      blobs.push(blob);
      Logger.log('  PDF sheet ' + (i+1) + '/' + sheets.length + ': ' + sheets[i].getName());
    }
    // delay ระหว่าง requests เพื่อหลีกเลี่ยง 429
    if (i < sheets.length - 1) Utilities.sleep(1500);
  }
  return blobs;
}

/**
 * Merge หลาย PDF blobs ด้วย pdf-lib (async — ใช้ได้ใน GAS V8)
 */
async function pp5s_mergePdfs_(blobs, fileName) {
  if (!blobs || blobs.length === 0) throw new Error('ไม่มี PDF blobs');
  if (blobs.length === 1) return blobs[0].setName(fileName);

  // Load pdf-lib from CDN
  var cdnjs = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';
  eval(UrlFetchApp.fetch(cdnjs).getContentText().replace(/setTimeout\(.*?,.*?(\d*?)\)/g, 'Utilities.sleep($1);return t();'));

  var pdfDoc = await PDFLib.PDFDocument.create();
  for (var i = 0; i < blobs.length; i++) {
    var data = new Uint8Array(blobs[i].getBytes());
    var srcDoc = await PDFLib.PDFDocument.load(data);
    var pages = await pdfDoc.copyPages(srcDoc, srcDoc.getPageIndices());
    pages.forEach(function(page) { pdfDoc.addPage(page); });
  }
  var bytes = await pdfDoc.save();
  return Utilities.newBlob([].slice.call(new Int8Array(bytes)), MimeType.PDF, fileName);
}

/**
 * ตั้งค่า font Sarabun + style พื้นฐานให้ range
 */
function pp5s_baseStyle_(range, fontSize) {
  range.setFontFamily('Sarabun')
    .setFontSize(fontSize || 12)
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setFontColor('#000000');
}

/**
 * ตั้งค่า header style
 */
function pp5s_headerStyle_(range, fontSize) {
  range.setFontFamily('Sarabun')
    .setFontSize(fontSize || 11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#f2f2f2')
    .setBorder(true, true, true, true, true, true);
}

/**
 * ใส่โลโก้ที่ cell
 */
function pp5s_insertLogo_(sheet, row, col, logoDataUri, logoSize, totalWidth) {
  if (!logoDataUri) return;
  try {
    var match = logoDataUri.match(/^data:image\/(png|jpeg|jpg|gif);base64,(.+)$/i);
    if (!match) return;
    var mimeType = 'image/' + match[1];
    var b64 = match[2];
    var blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, 'logo.' + match[1]);
    var sz = logoSize || 55;
    // ถ้าระบุ totalWidth ให้ center โลโก้แนวนอน โดย anchor ที่ col=1
    var anchorCol = totalWidth ? 1 : col;
    var logo = sheet.insertImage(blob, anchorCol, row);
    logo.setWidth(sz).setHeight(sz);
    if (totalWidth && totalWidth > sz) {
      logo.setAnchorCellXOffset(Math.round((totalWidth - sz) / 2));
    }
    logo.setAnchorCellYOffset(4);
  } catch (e) {
    Logger.log('⚠️ ใส่โลโก้ไม่ได้: ' + e.message);
  }
}

// ============================================================
// 📄 ส่วนที่ 1: ปก ปพ.5 (ใช้ Google Doc ควบคุม layout เต็มที่)
// ============================================================

/**
 * สร้างหน้าปก ปพ.5 ด้วย Google Doc แล้ว export เป็น PDF blob
 * @returns {Blob} PDF blob ของหน้าปก
 */
function pp5s_coverDoc_(meta, studentCount, subjectSummary, reading, attributes, activity) {
  var doc = DocumentApp.create('_pp5cover_' + Date.now());
  var body = doc.getBody();
  var A = DocumentApp.Attribute;

  // ตั้งค่าหน้า A4 portrait
  body.setPageWidth(595.28).setPageHeight(841.89)
    .setMarginTop(36).setMarginBottom(28).setMarginLeft(50).setMarginRight(28);

  // helper: style paragraph text
  function stylePara(p, size, bold) {
    var style = {};
    style[A.FONT_FAMILY] = 'Sarabun';
    style[A.FONT_SIZE] = size;
    style[A.BOLD] = bold;
    p.setAttributes(style);
  }
  // helper: style table cell
  function styleCell(cell, size, bold) {
    var style = {};
    style[A.FONT_FAMILY] = 'Sarabun';
    style[A.FONT_SIZE] = size;
    style[A.BOLD] = bold;
    cell.setAttributes(style);
  }

  // ใช้ default paragraph (index 0) สำหรับ element แรก
  var firstPara = body.getChild(0).asParagraph();
  var usedFirstPara = false;

  // --- โลโก้ ---
  if (meta.logoDataUri) {
    try {
      var match = meta.logoDataUri.match(/^data:image\/(png|jpeg|jpg|gif);base64,(.+)$/i);
      if (match) {
        var imgBlob = Utilities.newBlob(Utilities.base64Decode(match[2]), 'image/' + match[1], 'logo');
        firstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        firstPara.appendInlineImage(imgBlob).setWidth(50).setHeight(50);
        usedFirstPara = true;
      }
    } catch (e) { Logger.log('⚠️ โลโก้: ' + e.message); }
  }

  // --- หัวเรื่อง ---
  var h1;
  if (!usedFirstPara) {
    firstPara.setText('แบบรายงานผลการพัฒนาคุณภาพผู้เรียน (ปพ.5)');
    firstPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    stylePara(firstPara, 18, true);
    firstPara.setLineSpacing(1.0).setSpacingBefore(2).setSpacingAfter(1);
    h1 = firstPara;
  } else {
    h1 = body.appendParagraph('แบบรายงานผลการพัฒนาคุณภาพผู้เรียน (ปพ.5)');
    h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    stylePara(h1, 18, true);
    h1.setLineSpacing(1.0).setSpacingBefore(2).setSpacingAfter(1);
  }

  var h2 = body.appendParagraph(meta.schoolName || '');
  h2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  stylePara(h2, 16, true);
  h2.setLineSpacing(1.0).setSpacingBefore(1).setSpacingAfter(0);

  if (meta.settings && meta.settings['ที่อยู่โรงเรียน']) {
    var h3 = body.appendParagraph(meta.settings['ที่อยู่โรงเรียน']);
    h3.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    stylePara(h3, 11, false);
    h3.setLineSpacing(1.0).setSpacingBefore(0).setSpacingAfter(0);
  }

  var h4 = body.appendParagraph('ปีการศึกษา ' + meta.academicYear);
  h4.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  stylePara(h4, 16, true);
  h4.setLineSpacing(1.0).setSpacingBefore(3).setSpacingAfter(1);

  var h5 = body.appendParagraph('ชั้น' + meta.gradeFullName + '  ห้อง ' + meta.classNo + '    จำนวนนักเรียน ' + studentCount + ' คน');
  h5.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  stylePara(h5, 14, true);
  h5.setLineSpacing(1.0).setSpacingBefore(1).setSpacingAfter(4);

  // --- ตารางสรุปเกรดรายวิชา ---
  var coverOrder = ['ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา','ประวัติศาสตร์','สุขศึกษา','ศิลปะ','การงาน','ภาษาอังกฤษ','หน้าที่พลเมือง','การป้องกัน'];
  var sorted = (subjectSummary || []).slice().sort(function(a, b) {
    var ia = -1, ib = -1;
    for (var k = 0; k < coverOrder.length; k++) {
      if ((a.name || '').indexOf(coverOrder[k]) !== -1) ia = k;
      if ((b.name || '').indexOf(coverOrder[k]) !== -1) ib = k;
    }
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return (a.name || '').localeCompare(b.name || '', 'th');
  });

  var numSubjects = sorted.length;
  var gradeTable = body.appendTable();

  // แถวที่ 1: ที่, ชื่อวิชา, "ระดับผลการเรียนรายวิชา (คน)"
  var hdrRow1 = gradeTable.appendTableRow();
  ['ที่', 'ชื่อวิชา', 'ระดับผลการเรียนรายวิชา (คน)', '', '', '', '', '', '', ''].forEach(function(h) {
    var cell = hdrRow1.appendTableCell(h);
    styleCell(cell, 12, true);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.setBackgroundColor('#f0f0f0');
    cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(1).setPaddingRight(1);
  });

  // แถวที่ 2: 4 3.5 3 2.5 2 1.5 1 0
  var hdrRow2 = gradeTable.appendTableRow();
  ['', '', '4', '3.5', '3', '2.5', '2', '1.5', '1', '0'].forEach(function(h) {
    var cell = hdrRow2.appendTableCell(h);
    styleCell(cell, 12, true);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.setBackgroundColor('#f0f0f0');
    cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(1).setPaddingRight(1);
  });

  sorted.forEach(function(s, i) {
    var row = gradeTable.appendTableRow();
    var vals = [String(i + 1), s.name || '',
      String(s.count['4'] || 0), String(s.count['3.5'] || 0), String(s.count['3'] || 0), String(s.count['2.5'] || 0),
      String(s.count['2'] || 0), String(s.count['1.5'] || 0), String(s.count['1'] || 0), String(s.count['0'] || 0)];
    vals.forEach(function(v, ci) {
      var cell = row.appendTableCell(v);
      styleCell(cell, 12, false);
      cell.getChild(0).asParagraph().setAlignment(ci === 1 ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER);
      cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(1).setPaddingRight(1);
    });
  });

  gradeTable.setColumnWidth(0, 25);
  gradeTable.setColumnWidth(1, 140);
  for (var gc = 2; gc <= 9; gc++) gradeTable.setColumnWidth(gc, 45);

  // ลบ default row 0 (appendTable สร้าง row ว่าง 1 แถว)
  try {
    if (gradeTable.getNumRows() > (numSubjects + 2) && gradeTable.getRow(0).getNumCells() <= 1) {
      gradeTable.removeRow(0);
    }
  } catch (e) {}

  body.appendParagraph('').setLineSpacing(1.0).setSpacingBefore(2).setSpacingAfter(0);

  // --- ตารางประเมินบูรณาการ ---
  var evalTitle = body.appendParagraph('การประเมินผลแบบบูรณาการ');
  evalTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  stylePara(evalTitle, 13, true);
  evalTitle.setLineSpacing(1.0).setSpacingBefore(2).setSpacingAfter(2);

  var pct = function(n) { return studentCount > 0 ? ((n || 0) / studentCount * 100).toFixed(1) : '0.0'; };

  var evalTable = body.appendTable();

  // header row 1: 3 กลุ่ม ในแต่ละกลุ่ม 3 คอลัมน์
  var eh1 = evalTable.appendTableRow();
  ['การอ่านฯ', '', '', 'คุณลักษณะฯ', '', '', 'กิจกรรมฯ', '', ''].forEach(function(h) {
    var cell = eh1.appendTableCell(h);
    styleCell(cell, 10, true);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.setBackgroundColor('#f0f0f0');
    cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(1).setPaddingRight(1);
  });

  // header row 2
  var eh2 = evalTable.appendTableRow();
  ['ระดับคุณภาพ', 'คน', 'ร้อยละ', 'ระดับคุณภาพ', 'คน', 'ร้อยละ', 'การผ่าน', 'คน', 'ร้อยละ'].forEach(function(h) {
    var cell = eh2.appendTableCell(h);
    styleCell(cell, 10, true);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.setBackgroundColor('#f0f0f0');
    cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(1).setPaddingRight(1);
  });

  // data rows
  [
    ['ดีเยี่ยม', String(reading.level3||0), pct(reading.level3), 'ดีเยี่ยม', String(attributes['ดีเยี่ยม']||0), pct(attributes['ดีเยี่ยม']), 'ผ่าน', String(activity['ผ่าน']||0), pct(activity['ผ่าน'])],
    ['ดี', String(reading.level2||0), pct(reading.level2), 'ดี', String(attributes['ดี']||0), pct(attributes['ดี']), 'ไม่ผ่าน', String(activity['ไม่ผ่าน']||0), pct(activity['ไม่ผ่าน'])],
    ['ผ่านเกณฑ์', String(reading.level1||0), pct(reading.level1), 'ผ่านเกณฑ์', String(attributes['ผ่าน']||0), pct(attributes['ผ่าน']), '', '', ''],
    ['ไม่ผ่านเกณฑ์', String(reading.level0||0), pct(reading.level0), 'ไม่ผ่านเกณฑ์', String(attributes['ไม่ผ่าน']||0), pct(attributes['ไม่ผ่าน']), '', '', '']
  ].forEach(function(rowData) {
    var row = evalTable.appendTableRow();
    rowData.forEach(function(v, ci) {
      var cell = row.appendTableCell(v);
      styleCell(cell, 10, false);
      var align = (ci === 0 || ci === 3 || ci === 6) ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER;
      cell.getChild(0).asParagraph().setAlignment(align);
      cell.setPaddingTop(1).setPaddingBottom(1).setPaddingLeft(2).setPaddingRight(2);
    });
  });

  // 90+40+50+90+40+50+80+40+45=525 (เท่ากับ gradeTable)
  evalTable.setColumnWidth(0, 90); evalTable.setColumnWidth(1, 40); evalTable.setColumnWidth(2, 50);
  evalTable.setColumnWidth(3, 90); evalTable.setColumnWidth(4, 40); evalTable.setColumnWidth(5, 50);
  evalTable.setColumnWidth(6, 80); evalTable.setColumnWidth(7, 40); evalTable.setColumnWidth(8, 45);

  // ลบ default row 0
  try {
    if (evalTable.getNumRows() > 6 && evalTable.getRow(0).getNumCells() <= 1) {
      evalTable.removeRow(0);
    }
  } catch (e) {}

  body.appendParagraph('').setLineSpacing(1.0).setSpacingBefore(4).setSpacingAfter(0);

  // --- ลงชื่อ ---
  var signTable = body.appendTable();
  var signR1 = signTable.appendTableRow();
  [['', ''], ['☐ เห็นควรอนุมัติ  ☐ เห็นควรปรับปรุง', ''], ['☐ อนุมัติ  ☐ ไม่อนุมัติ', '']].forEach(function(pair) {
    var cell = signR1.appendTableCell(pair[0]);
    styleCell(cell, 12, false);
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.setPaddingTop(1).setPaddingBottom(1);
  });
  var signR2 = signTable.appendTableRow();
  var signLines = [
    ['ลงชื่อ......................................', '(' + (meta.teacherName || '...............................................') + ')', 'ครูประจำชั้น'],
    ['ลงชื่อ......................................', '(' + (meta.academicHead || '...............................................') + ')', 'หัวหน้างานวิชาการ'],
    ['ลงชื่อ......................................', '(' + (meta.directorName || '...............................................') + ')', (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา')]
  ];
  signLines.forEach(function(lines) {
    var cell = signR2.appendTableCell('');
    cell.clear();
    lines.forEach(function(line, li) {
      var para = (li === 0) ? cell.getChild(0).asParagraph() : cell.appendParagraph('');
      para.setText(line);
      para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      para.editAsText().setFontFamily('Sarabun').setFontSize(12).setBold(false);
      para.setLineSpacing(1.15).setSpacingBefore(0).setSpacingAfter(0);
    });
    cell.setPaddingTop(3).setPaddingBottom(3);
  });

  signTable.setColumnWidth(0, 175); signTable.setColumnWidth(1, 175); signTable.setColumnWidth(2, 175);
  signTable.setBorderWidth(0);

  // ลบ default row 0
  try {
    if (signTable.getNumRows() > 2 && signTable.getRow(0).getNumCells() <= 1) {
      signTable.removeRow(0);
    }
  } catch (e) {}

  // --- Export Doc เป็น PDF ---
  doc.saveAndClose();

  // Merge cells หัวตาราง gradeTable ด้วย Docs API
  var docId = doc.getId();
  try {
    var docData = Docs.Documents.get(docId);
    var tables = [];
    docData.body.content.forEach(function(el) {
      if (el.table) tables.push(el);
    });
    // gradeTable = ตารางแรกที่มี 10+ columns
    var mergeReqs = [];
    for (var ti = 0; ti < tables.length; ti++) {
      var tbl = tables[ti].table;
      if (tbl.columns >= 10 && tbl.rows >= 3) {
        var tsi = tables[ti].startIndex;
        // แถว 1 (row 0): merge col 2-9 เป็น "ระดับผลการเรียนรายวิชา (คน)"
        mergeReqs.push({ mergeTableCells: { tableRange: { tableCellLocation: { tableStartLocation: { index: tsi }, rowIndex: 0, columnIndex: 2 }, rowSpan: 1, columnSpan: 8 } } });
        // แถว 1-2 (row 0-1): merge col 0 เป็น "ที่"
        mergeReqs.push({ mergeTableCells: { tableRange: { tableCellLocation: { tableStartLocation: { index: tsi }, rowIndex: 0, columnIndex: 0 }, rowSpan: 2, columnSpan: 1 } } });
        // แถว 1-2 (row 0-1): merge col 1 เป็น "ชื่อวิชา"
        mergeReqs.push({ mergeTableCells: { tableRange: { tableCellLocation: { tableStartLocation: { index: tsi }, rowIndex: 0, columnIndex: 1 }, rowSpan: 2, columnSpan: 1 } } });
        break;
      }
    }
    if (mergeReqs.length > 0) {
      Docs.Documents.batchUpdate({ requests: mergeReqs }, docId);
    }
  } catch (me) { Logger.log('Cover merge: ' + me.message); }

  var docFile = DriveApp.getFileById(docId);
  var pdfBlob = docFile.getAs('application/pdf');
  try { docFile.setTrashed(true); } catch (e) {}
  Logger.log('✅ ปก (Google Doc → PDF)');
  return pdfBlob;
}


// ============================================================
// 📄 ส่วนที่ 2: รายชื่อนักเรียน
// ============================================================
function pp5s_studentList_(ss, meta, students) {
  var sheet = ss.insertSheet('รายชื่อ');
  var r = 1;

  // header
  sheet.getRange(r, 1, 1, 10).merge();
  sheet.getRange(r, 1).setValue('ข้อมูลนักเรียน ชั้น' + meta.gradeFullName + ' ห้อง ' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 32);
  r++;

  // table header row 1
  var h1 = ['ที่', 'เลขประจำตัว', 'เลขประจำตัวประชาชน', 'ชื่อ - สกุล', 'วัน', 'เดือน', 'ปี', 'อายุ', 'น้ำหนัก', 'ส่วนสูง'];
  sheet.getRange(r, 1, 1, 10).setValues([h1]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 10), 12);
  // merge header "วัน เดือน ปี เกิด" — ข้ามเพราะ simple header ก็พอ
  sheet.setRowHeight(r, 26);
  r++;

  // data rows
  students.forEach(function(s, i) {
    var bd = s.birthdate || '';
    var parts = bd.split('-');
    var day = '', month = '', year = '';
    if (parts.length === 3) {
      day = parts[2]; month = parts[1]; year = String(Number(parts[0]) + 543);
    }
    var vals = [i + 1, s.id || '', s.idCard || '',
      (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || ''),
      day, month, year, s.age || '', s.weight || '', s.height || ''];
    sheet.getRange(r, 1, 1, 10).setValues([vals]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, 10), 12);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 2).setHorizontalAlignment('center');
    sheet.getRange(r, 3).setHorizontalAlignment('center');
    sheet.getRange(r, 5, 1, 6).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, 10).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // ~700px: 30+70+120+180+35+35+35+40+35+35=615
  sheet.setColumnWidth(1, 30);   // ที่
  sheet.setColumnWidth(2, 70);   // เลขประจำตัว
  sheet.setColumnWidth(3, 120);  // เลขประจำตัวประชาชน
  sheet.setColumnWidth(4, 180);  // ชื่อ-สกุล
  for (var c = 5; c <= 7; c++) sheet.setColumnWidth(c, 35);  // วัน เดือน ปี
  sheet.setColumnWidth(8, 40);   // อายุ
  sheet.setColumnWidth(9, 35);   // น้ำหนัก
  sheet.setColumnWidth(10, 35);  // ส่วนสูง

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 3: ข้อมูลผู้ปกครอง
// ============================================================
function pp5s_parentInfo_(ss, meta, students) {
  var sheet = ss.insertSheet('ผู้ปกครอง');
  var r = 1;

  sheet.getRange(r, 1, 1, 6).merge();
  sheet.getRange(r, 1).setValue('ข้อมูลผู้ปกครอง ชั้น' + meta.gradeFullName + ' ห้อง ' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 32);
  r++;

  var h1 = ['ที่', 'ชื่อ-สกุลบิดา', 'อาชีพ', 'ชื่อ-สกุลมารดา', 'อาชีพ', 'ที่อยู่ปัจจุบัน'];
  sheet.getRange(r, 1, 1, 6).setValues([h1]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 6), 12);
  sheet.setRowHeight(r, 26);
  r++;

  students.forEach(function(s, i) {
    var vals = [i + 1, s.father_name || '-', s.father_occupation || '-',
      s.mother_name || '-', s.mother_occupation || '-', s.address || '-'];
    sheet.getRange(r, 1, 1, 6).setValues([vals]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, 6), 12);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 3).setHorizontalAlignment('center');
    sheet.getRange(r, 5).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, 6).setBorder(true, true, true, true, true, true).setWrap(true);
    sheet.setRowHeight(r, 30);
    r++;
  });

  // 28+140+70+140+70+210=658
  sheet.setColumnWidth(1, 28);   // ที่
  sheet.setColumnWidth(2, 140);  // ชื่อบิดา
  sheet.setColumnWidth(3, 70);   // อาชีพ
  sheet.setColumnWidth(4, 140);  // ชื่อมารดา
  sheet.setColumnWidth(5, 70);   // อาชีพ
  sheet.setColumnWidth(6, 210);  // ที่อยู่ปัจจุบัน

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 4A: สรุปเวลาเรียน (รวมภาค)
// ============================================================
function pp5s_attendanceSummary_(ss, meta, students, att1, att2) {
  var sheet = ss.insertSheet('สรุปเวลาเรียน');
  var r = 1;

  sheet.getRange(r, 1, 1, 15).merge();
  sheet.getRange(r, 1).setValue('สรุปเวลาเรียน ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 32);
  r++;

  // header row 1
  sheet.getRange(r, 1).setValue('ที่');
  sheet.getRange(r, 2).setValue('ชื่อ-นามสกุล');
  sheet.getRange(r, 3, 1, 4).merge().setValue('ภาคเรียนที่ 1 (วัน)');
  sheet.getRange(r, 7, 1, 4).merge().setValue('ภาคเรียนที่ 2 (วัน)');
  sheet.getRange(r, 11, 1, 5).merge().setValue('รวมทั้งปี');
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 15), 12);
  sheet.setRowHeight(r, 24);
  r++;

  // header row 2
  var h2 = ['', '', 'มา', 'ลา', 'ขาด', 'รวม', 'มา', 'ลา', 'ขาด', 'รวม', 'มา', 'ลา', 'ขาด', 'รวม', '%มา'];
  sheet.getRange(r, 1, 1, 15).setValues([h2]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 15), 12);
  sheet.setRowHeight(r, 22);
  r++;

  var att1Map = {};
  (att1 || []).forEach(function(a) { att1Map[a.studentId] = a; });
  var att2Map = {};
  (att2 || []).forEach(function(a) { att2Map[a.studentId] = a; });

  students.forEach(function(s, i) {
    var a1 = att1Map[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
    var a2 = att2Map[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
    var tp = a1.present + a2.present;
    var tl = a1.leave + a2.leave;
    var ta = a1.absent + a2.absent;
    var tt = a1.total + a2.total;
    var pctVal = tt > 0 ? ((tp / tt) * 100).toFixed(1) : '0.0';

    var vals = [i + 1, (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || ''),
      a1.present, a1.leave, a1.absent, a1.total,
      a2.present, a2.leave, a2.absent, a2.total,
      tp, tl, ta, tt, pctVal];
    sheet.getRange(r, 1, 1, 15).setValues([vals]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, 15), 12);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 3, 1, 13).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, 15).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // 30+170+13×33=631
  sheet.setColumnWidth(1, 30);
  sheet.setColumnWidth(2, 170);
  for (var c = 3; c <= 15; c++) sheet.setColumnWidth(c, 33);

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 4B: เวลาเรียนรายเดือน (สรุป มา/ลา/ขาด)
// ============================================================
function pp5s_monthlyAttendance_(ss, meta, students, monthlyDataAll) {
  if (!monthlyDataAll || monthlyDataAll.length === 0) return [];
  var sheets = [];

  var studentList = students.map(function(s) {
    return { id: s.id, name: (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '') };
  });

  monthlyDataAll.forEach(function(monthInfo) {
    var sheetName = 'เดือน' + monthInfo.name;
    var sheet = ss.insertSheet(sheetName);
    var r = 1;

    // โลโก้
    sheet.setRowHeight(r, 60);
    sheet.getRange(r, 1, 1, 6).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
    pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 580);
    r++;

    // header
    sheet.getRange(r, 1, 1, 6).merge();
    sheet.getRange(r, 1).setValue('บันทึกเวลาเรียน (เช็คชื่อรายเดือน)')
      .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setRowHeight(r, 28);
    r++;

    sheet.getRange(r, 1, 1, 6).merge();
    sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | เดือน' + monthInfo.name + ' ' + monthInfo.yearBE + ' | ปีการศึกษา ' + meta.academicYear)
      .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
    sheet.setRowHeight(r, 24);
    r++;

    // table header
    var hdr = ['ที่', 'ชื่อ-นามสกุล', 'มา', 'ลา', 'ขาด', 'รวม'];
    sheet.getRange(r, 1, 1, 6).setValues([hdr]);
    pp5s_headerStyle_(sheet.getRange(r, 1, 1, 6), 13);
    sheet.setRowHeight(r, 24);
    r++;

    var dataMap = {};
    (monthInfo.data || []).forEach(function(d) { dataMap[d.studentId] = d; });

    var sumP = 0, sumL = 0, sumA = 0;
    studentList.forEach(function(s, i) {
      var d = dataMap[s.id] || { present: 0, leave: 0, absent: 0, total: 0 };
      sumP += d.present || 0; sumL += d.leave || 0; sumA += d.absent || 0;
      var vals = [i + 1, s.name, d.present || 0, d.leave || 0, d.absent || 0, d.total || 0];
      sheet.getRange(r, 1, 1, 6).setValues([vals]);
      pp5s_baseStyle_(sheet.getRange(r, 1, 1, 6), 13);
      sheet.getRange(r, 1).setHorizontalAlignment('center');
      sheet.getRange(r, 3, 1, 4).setHorizontalAlignment('center');
      sheet.getRange(r, 1, 1, 6).setBorder(true, true, true, true, true, true);
      if (i % 2 === 0) sheet.getRange(r, 1, 1, 6).setBackground('#f8f9fa');
      sheet.setRowHeight(r, 24);
      r++;
    });

    // footer
    sheet.getRange(r, 1, 1, 2).merge().setValue('รวมทั้งห้อง');
    sheet.getRange(r, 3).setValue(sumP);
    sheet.getRange(r, 4).setValue(sumL);
    sheet.getRange(r, 5).setValue(sumA);
    sheet.getRange(r, 6).setValue(sumP + sumL + sumA);
    sheet.getRange(r, 1, 1, 6).setFontFamily('Sarabun').setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center').setBorder(true, true, true, true, true, true).setBackground('#e3f2fd');
    sheet.setRowHeight(r, 24);

    // 30+200+4×88=382 → total 580px
    sheet.setColumnWidth(1, 30);
    sheet.setColumnWidth(2, 200);
    for (var c = 3; c <= 6; c++) sheet.setColumnWidth(c, 88);

    sheets.push(sheet);
  });

  return sheets;
}


// ============================================================
// 📄 ส่วนที่ 4C: ตารางเช็คชื่อรายวัน (✓/ข/ล แต่ละวัน)
// ============================================================
function pp5s_dailyAttendance_(ss, meta, tableData, monthName, yearBE, semester) {
  var students = tableData.students || [];
  var days = tableData.days || [];
  var actualSchoolDays = tableData.actualSchoolDays || days.length;
  var totalDaysInMonth = tableData.totalDaysInMonth || days.length;

  if (students.length === 0 || days.length === 0) return null;

  var sheetName = 'รายวัน' + monthName;
  var sheet = ss.insertSheet(sheetName);
  var numDays = days.length;
  var numCols = 3 + numDays; // ที่, รหัส, ชื่อ-สกุล, วันที่...
  var r = 1;

  // โลโก้
  sheet.setRowHeight(r, 55);
  sheet.getRange(r, 1, 1, numCols).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 45, 620);
  r++;

  // header
  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName || '')
    .setFontFamily('Sarabun').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 26);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('บันทึกเวลาเรียน ชั้น ' + meta.grade + '/' + meta.classNo + ' เดือน ' + monthName + ' ' + yearBE)
    .setFontFamily('Sarabun').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 26);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue(semester + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 22);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('แสดงเฉพาะวันที่มีการเรียน: ' + actualSchoolDays + ' วัน (จาก ' + totalDaysInMonth + ' วันในเดือน)')
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 20);
  r++;

  // table header: row 1 (เลขวัน)
  var headerRow1 = ['ที่', 'รหัส', 'ชื่อ - สกุล'];
  days.forEach(function(d) { headerRow1.push(d.label); });
  sheet.getRange(r, 1, 1, numCols).setValues([headerRow1]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 10);
  sheet.setRowHeight(r, 20);
  r++;

  // table header: row 2 (วันในสัปดาห์)
  var headerRow2 = ['', '', ''];
  days.forEach(function(d) { headerRow2.push(d.dow); });
  sheet.getRange(r, 1, 1, numCols).setValues([headerRow2]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 9);
  sheet.getRange(r, 1, 1, numCols).setFontWeight('normal');
  sheet.setRowHeight(r, 18);
  r++;

  // data rows
  var dataStartRow = r;
  students.forEach(function(stu, i) {
    var rowData = [i + 1, stu.id || '', stu.name || ''];
    days.forEach(function(d) {
      var status = stu.statusMap[d.label] || '';
      var display = '';
      if (status === '/' || status === 'ม' || status === '1') display = '✓';
      else if (status === 'ล' || String(status).toLowerCase() === 'l') display = 'ล';
      else if (status === 'ข' || status === '0') display = 'ข';
      rowData.push(display);
    });
    sheet.getRange(r, 1, 1, numCols).setValues([rowData]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, numCols), 11);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 2).setHorizontalAlignment('center');
    sheet.getRange(r, 4, 1, numDays).setHorizontalAlignment('center').setFontWeight('normal').setFontSize(11);
    sheet.getRange(r, 1, 1, numCols).setBorder(true, true, true, true, true, true);

    // สีตาม status
    for (var dc = 0; dc < numDays; dc++) {
      var cellVal = rowData[3 + dc];
      if (cellVal === '✓') sheet.getRange(r, 4 + dc).setFontColor('#000000');
      else if (cellVal === 'ล') sheet.getRange(r, 4 + dc).setFontColor('#e65100').setFontWeight('bold');
      else if (cellVal === 'ข') sheet.getRange(r, 4 + dc).setFontColor('#c62828').setFontWeight('bold');
    }

    if (i % 2 === 0) sheet.getRange(r, 1, 1, numCols).setBackground('#f8f9fa');
    sheet.setRowHeight(r, 25);
    r++;
  });

  // หมายเหตุ
  r++;
  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('หมายเหตุ: ✓ = มาเรียน   ล = ลา   ข = ขาด   |   นักเรียน: ' + students.length + ' คน   |   วันเรียน: ' + actualSchoolDays + ' วัน')
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 20);

  // column widths — total ~620px: 26+45+fixedName + days×dateColWidth
  var fixedWidth = 26 + 45; // ที่ + รหัส = 71
  var availForNameAndDays = 620 - fixedWidth; // = 549
  var dateColWidth = Math.max(16, Math.min(22, Math.floor((availForNameAndDays - 140) / Math.max(numDays, 1))));
  var nameWidth = availForNameAndDays - (numDays * dateColWidth);
  if (nameWidth < 120) nameWidth = 120;
  sheet.setColumnWidth(1, 26);
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, nameWidth);
  for (var dc = 0; dc < numDays; dc++) {
    sheet.setColumnWidth(4 + dc, dateColWidth);
  }

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 5: คะแนนรวม 2 ภาค แต่ละวิชา
// ============================================================
function pp5s_yearScore_(ss, meta, sheetInfo, scoreData) {
  var t1 = scoreData.term1 || {};
  var t2 = scoreData.term2 || {};
  if ((!t1.rows || t1.rows.length === 0) && (!t2.rows || t2.rows.length === 0)) return null;

  var rows1 = t1.rows || [];
  var rows2 = t2.rows || [];
  var yearSummary = scoreData.yearSummary || [];
  var yearMap = {};
  yearSummary.forEach(function(ys) { yearMap[String(ys.id)] = ys; });

  var subName = (sheetInfo.subjectName || '').replace(/[\/\\?*[\]]/g, '').substring(0, 20);
  // ป้องกันชื่อ sheet ซ้ำ (เช่น วิชาเดียวกัน 2 ภาคเรียน)
  var uniqueName = subName;
  var suffix = 2;
  while (ss.getSheetByName(uniqueName)) {
    uniqueName = subName.substring(0, 17) + '(' + suffix + ')';
    suffix++;
  }
  var sheet = ss.insertSheet(uniqueName);
  var r = 1;

  // โลโก้
  sheet.setRowHeight(r, 65);
  sheet.getRange(r, 1, 1, 9).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 55, 610);
  r++;

  // header
  sheet.getRange(r, 1, 1, 9).merge();
  sheet.getRange(r, 1).setValue('ผลการประเมิน รายวิชา: ' + (sheetInfo.subjectName || '') + ' (' + (sheetInfo.subjectCode || '') + ')')
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 30);
  r++;

  sheet.getRange(r, 1, 1, 9).merge();
  sheet.getRange(r, 1).setValue(meta.grade + '/' + meta.classNo + ' | รวมทั้งปีการศึกษา | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 24);
  r++;

  // table header
  sheet.getRange(r, 1).setValue('เลขที่');
  sheet.getRange(r, 2).setValue('รหัส');
  sheet.getRange(r, 3).setValue('ชื่อ-นามสกุล');
  sheet.getRange(r, 4, 1, 2).merge().setValue('ภาคเรียนที่ 1');
  sheet.getRange(r, 6, 1, 2).merge().setValue('ภาคเรียนที่ 2');
  sheet.getRange(r, 8, 1, 2).merge().setValue('รวมทั้งปี');
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 9), 11);
  sheet.getRange(r, 4, 1, 2).setBackground('#e3f2fd');
  sheet.getRange(r, 6, 1, 2).setBackground('#fff3e0');
  sheet.getRange(r, 8, 1, 2).setBackground('#e8f5e9');
  sheet.setRowHeight(r, 24);
  r++;

  var h2 = ['', '', '', 'คะแนน', 'ผลการเรียน', 'คะแนน', 'ผลการเรียน', 'เฉลี่ย 2 ภาค', 'ระดับผลการเรียน'];
  sheet.getRange(r, 1, 1, 9).setValues([h2]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, 9), 11);
  sheet.getRange(r, 4, 1, 2).setBackground('#e3f2fd');
  sheet.getRange(r, 6, 1, 2).setBackground('#fff3e0');
  sheet.getRange(r, 8, 1, 2).setBackground('#e8f5e9');
  sheet.setRowHeight(r, 22);
  r++;

  // data
  rows1.forEach(function(r1, i) {
    var r2 = rows2.find(function(rr) { return rr.id === r1.id; }) || rows2[i] || {};
    var total1 = Number(r1.total) || 0;
    var total2 = Number(r2.total) || 0;
    var grade1 = r1.grade || '0';
    var grade2 = r2.grade || '0';
    var ys = yearMap[String(r1.id)] || {};
    var yearAvg = ys.yearAvg || 0;
    var yearGrade = ys.yearGrade || '';
    if (!yearGrade && (total1 > 0 || total2 > 0)) {
      yearAvg = Math.round((total1 + total2) / 2);
      yearGrade = calculateFinalGrade(yearAvg);
    }

    var vals = [i + 1, r1.id || '', r1.name || '', total1, grade1, total2, grade2, yearAvg, yearGrade];
    sheet.getRange(r, 1, 1, 9).setValues([vals]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, 9), 12);
    sheet.getRange(r, 1, 1, 2).setHorizontalAlignment('center');
    sheet.getRange(r, 4, 1, 6).setHorizontalAlignment('center');
    sheet.getRange(r, 5).setFontWeight('bold');
    sheet.getRange(r, 7).setFontWeight('bold');
    sheet.getRange(r, 8, 1, 2).setFontWeight('bold').setBackground('#f0f7f0');
    sheet.getRange(r, 1, 1, 9).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  r += 2;
  // ลงชื่อ
  var subjectTeacher = sheetInfo.teacherName || '';
  var teacherLine = subjectTeacher ? '(' + subjectTeacher + ')' : '(..............................................)';
  sheet.getRange(r, 1, 1, 4).merge().setValue('ลงชื่อ...............................................ครูผู้สอน\n' + teacherLine + '\nตำแหน่ง ครู')
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, 5, 1, 5).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '..............................................') + ')\n' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา'))
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // 30+50+170+55+60+55+60+55+75=610
  sheet.setColumnWidth(1, 30);  // เลขที่
  sheet.setColumnWidth(2, 50);  // รหัส
  sheet.setColumnWidth(3, 170); // ชื่อ
  sheet.setColumnWidth(4, 55);  // คะแนน ภ 1
  sheet.setColumnWidth(5, 60);  // ผลการเรียน ภ 1
  sheet.setColumnWidth(6, 55);  // คะแนน ภ 2
  sheet.setColumnWidth(7, 60);  // ผลการเรียน ภ 2
  sheet.setColumnWidth(8, 55);  // เฉลี่ย
  sheet.setColumnWidth(9, 75);  // ระดับผลการเรียน

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 6: คุณลักษณะอันพึงประสงค์
// ============================================================
function pp5s_characteristic_(ss, meta, charData) {
  if (!charData || charData.length === 0) return null;
  var sheet = ss.insertSheet('คุณลักษณะ');
  var r = 1;
  var traitNames = ['รักชาติ ศาสน์ กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้',
    'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ'];

  // โลโก้
  sheet.setRowHeight(r, 60);
  sheet.getRange(r, 1, 1, 12).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 580);
  r++;

  sheet.getRange(r, 1, 1, 12).merge();
  sheet.getRange(r, 1).setValue('รายงานการประเมินคุณลักษณะอันพึงประสงค์')
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 30);
  r++;

  sheet.getRange(r, 1, 1, 12).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 24);
  r++;

  // header
  var hdr = ['ที่', 'ชื่อ-นามสกุล'];
  traitNames.forEach(function(t) { hdr.push(t); });
  hdr.push('รวม'); hdr.push('ผลประเมิน');
  sheet.getRange(r, 1, 1, hdr.length).setValues([hdr]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, hdr.length), 9);
  // หมุนหัวข้อคุณลักษณะ 8 ข้อเป็นแนวตั้ง (col 3-10)
  sheet.getRange(r, 3, 1, 8).setTextRotation(90).setVerticalAlignment('bottom');
  sheet.getRange(r, 11, 1, 2).setTextRotation(90).setVerticalAlignment('bottom');
  sheet.setRowHeight(r, 110);
  r++;

  // data
  charData.forEach(function(s, i) {
    var scores = s.scores || [];
    var rowData = [i + 1, s.name || ''];
    var validScores = scores.filter(function(sc) { return sc !== '' && !isNaN(sc); }).map(Number);
    var sum = 0;
    for (var k = 0; k < 8; k++) {
      var v = scores[k];
      if (v !== '' && !isNaN(v)) sum += Number(v);
      rowData.push(v !== '' && v !== undefined ? v : '-');
    }
    var avg = validScores.length === 8 ? sum / 8 : 0;
    var result = validScores.length === 8 ? (avg >= 2.5 ? 'ดีเยี่ยม' : avg >= 2.0 ? 'ดี' : avg >= 1.0 ? 'ผ่าน' : 'ปรับปรุง') : '-';
    rowData.push(avg.toFixed(2));
    rowData.push(result);
    sheet.getRange(r, 1, 1, rowData.length).setValues([rowData]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, rowData.length), 11);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 3, 1, 10).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, rowData.length).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // ลงชื่อ
  r += 2;
  sheet.getRange(r, 1, 1, 6).merge().setValue('ลงชื่อ...............................................ครูประจำชั้น\n(' + (meta.teacherName || '...............................................') + ')')
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, 7, 1, 6).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '...............................................') + ')\n' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา'))
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // 28+140+8×40+40+60=580
  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 140);
  for (var c = 3; c <= 10; c++) sheet.setColumnWidth(c, 40);
  sheet.setColumnWidth(11, 40);
  sheet.setColumnWidth(12, 60);

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 7: กิจกรรมพัฒนาผู้เรียน
// ============================================================
function pp5s_activity_(ss, meta, actData) {
  if (!actData || actData.length === 0) return null;
  var sheet = ss.insertSheet('กิจกรรม');
  var r = 1;

  // โลโก้
  sheet.setRowHeight(r, 60);
  sheet.getRange(r, 1, 1, 8).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 604);
  r++;

  sheet.getRange(r, 1, 1, 8).merge();
  sheet.getRange(r, 1).setValue('รายงานการประเมินกิจกรรมพัฒนาผู้เรียน')
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 30);
  r++;

  sheet.getRange(r, 1, 1, 8).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 24);
  r++;

  var actNames = ['แนะแนว', 'ลูกเสือ/เนตรนารี', 'ชุมนุม', 'เพื่อสังคมและสาธารณประโยชน์'];
  var hdr = ['ที่', 'ชื่อ-นามสกุล'];
  actNames.forEach(function(a) { hdr.push(a); });
  hdr.push('สรุป'); hdr.push('ผลประเมิน');
  sheet.getRange(r, 1, 1, hdr.length).setValues([hdr]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, hdr.length), 11);
  sheet.setRowHeight(r, 26);
  r++;

  actData.forEach(function(s, i) {
    var act = s.activities || {};
    var g = act.guidance || 'ผ่าน';
    var sc = act.scout || 'ผ่าน';
    var cl = act.club || 'ผ่าน';
    var so = act.social || 'ผ่าน';
    var overall = act.overall || '';
    var scores = [g, sc, cl, so];
    var rowData = [i + 1, s.name || ''];
    var allPass = true;
    scores.forEach(function(v) {
      if (v === 'มผ' || v === 'ไม่ผ่าน' || v === '0') allPass = false;
      rowData.push(v);
    });
    rowData.push(overall || (allPass ? 'ผ่าน' : 'ไม่ผ่าน'));
    rowData.push(allPass ? 'ผ' : 'มผ');
    sheet.getRange(r, 1, 1, rowData.length).setValues([rowData]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, rowData.length), 12);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 3, 1, 6).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, rowData.length).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // ลงชื่อ
  r += 2;
  sheet.getRange(r, 1, 1, 4).merge().setValue('ลงชื่อ...............................................ครูประจำชั้น\n(' + (meta.teacherName || '...............................................') + ')')
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, 5, 1, 4).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '...............................................') + ')\n' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา'))
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // 28+180+4×75+48+48=604
  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 180);
  for (var c = 3; c <= 6; c++) sheet.setColumnWidth(c, 75);
  sheet.setColumnWidth(7, 48);
  sheet.setColumnWidth(8, 48);

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 8: อ่าน คิดวิเคราะห์ เขียน
// ============================================================
function pp5s_readThinkWrite_(ss, meta, subjectScoreStudents) {
  if (!subjectScoreStudents || subjectScoreStudents.length === 0) return null;
  var sheet = ss.insertSheet('อ่านคิดเขียน');
  var r = 1;
  var numCols = 12; // ที่, รหัส, ชื่อ, 8 วิชา, สรุปผล

  // โลโก้
  sheet.setRowHeight(r, 60);
  sheet.getRange(r, 1, 1, numCols).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 607);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('แบบประเมินการอ่าน คิดวิเคราะห์ เขียน')
    .setFontFamily('Sarabun').setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 30);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 24);
  r++;

  // header row 1
  sheet.getRange(r, 1).setValue('ที่');
  sheet.getRange(r, 2).setValue('รหัส');
  sheet.getRange(r, 3).setValue('ชื่อ - นามสกุล');
  sheet.getRange(r, 4, 1, 8).merge().setValue('ผลการเรียนรายวิชา (เกรด 1-4)');
  sheet.getRange(r, 12).setValue('สรุปผล');
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 11);
  sheet.setRowHeight(r, 24);
  r++;

  // header row 2
  var h2 = ['', '', '', 'ไทย', 'คณิต', 'วิทย์', 'สังคม', 'สุขะ', 'ศิลปะ', 'การงาน', 'อังกฤษ', ''];
  sheet.getRange(r, 1, 1, numCols).setValues([h2]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 11);
  sheet.setRowHeight(r, 22);
  r++;

  subjectScoreStudents.forEach(function(s, i) {
    var sc = s.scores || {};
    var scoresArr = [sc.thai, sc.math, sc.science, sc.social, sc.health, sc.art, sc.work, sc.english].map(function(v) { return Number(v) || 0; });
    var result = '';
    try { result = calculateSubjectScoreResult_(scoresArr); } catch (e) {
      // fallback: คำนวณเอง
      var allZero = scoresArr.every(function(v) { return v === 0; });
      if (!allZero) {
        var sum = scoresArr.reduce(function(a, b) { return a + b; }, 0);
        var avg = sum / 8;
        result = avg >= 3.5 ? 'ดีเยี่ยม' : avg >= 2.5 ? 'ดี' : avg >= 1.5 ? 'ผ่าน' : 'ปรับปรุง';
      }
    }

    var vals = [i + 1, s.studentId || '', s.name || '',
      Math.floor(Number(sc.thai)) || '-', Math.floor(Number(sc.math)) || '-',
      Math.floor(Number(sc.science)) || '-', Math.floor(Number(sc.social)) || '-',
      Math.floor(Number(sc.health)) || '-', Math.floor(Number(sc.art)) || '-',
      Math.floor(Number(sc.work)) || '-', Math.floor(Number(sc.english)) || '-',
      result];
    sheet.getRange(r, 1, 1, numCols).setValues([vals]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, numCols), 11);
    sheet.getRange(r, 1, 1, 2).setHorizontalAlignment('center');
    sheet.getRange(r, 4, 1, 9).setHorizontalAlignment('center');
    sheet.getRange(r, 12).setFontWeight('bold');
    sheet.getRange(r, 1, 1, numCols).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 24);
    r++;
  });

  // ลงชื่อ
  r += 2;
  sheet.getRange(r, 1, 1, 6).merge().setValue('ลงชื่อ...............................................ครูประจำชั้น\n(' + (meta.teacherName || '...............................................') + ')\nตำแหน่ง ครู')
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, 7, 1, 6).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '...............................................') + ')\nผู้อำนวยการสถานศึกษา')
    .setFontFamily('Sarabun').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // 28+50+170+8×38+55=607
  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 170);
  for (var c = 4; c <= 11; c++) sheet.setColumnWidth(c, 38);
  sheet.setColumnWidth(12, 55);

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 9: สรุปผลการเรียนรายวิชา (เกรดทุกวิชา + GPA + อันดับ)
// ============================================================
function pp5s_gradesSummary_(ss, meta, students, warehouseData, subjectList) {
  if (!students || students.length === 0) return null;
  if (!warehouseData || warehouseData.length === 0) return null;

  var sheet = ss.insertSheet('สรุปผลการเรียน');
  var r = 1;
  var grade = meta.grade;
  var classNo = meta.classNo;

  // --- สร้าง creditMap / typeMap จากชีตรายวิชา ---
  var creditMap = {}, typeMap = {};
  (subjectList || []).forEach(function(s) {
    var code = String(s['รหัสวิชา'] || '').trim();
    if (code) {
      creditMap[code] = parseFloat(s['ชั่วโมง/ปี']) || 1;
      typeMap[code] = String(s['ประเภทวิชา'] || 'พื้นฐาน').trim();
    }
  });

  // --- กรองเฉพาะชั้น/ห้องนี้ ---
  var classRows = warehouseData.filter(function(row) {
    return String(row['grade'] || '').trim() === grade && String(row['class_no'] || '').trim() === classNo;
  });

  // --- หาวิชาทั้งหมด (ไม่รวมกิจกรรม) เรียงตาม coverOrder ---
  var subjectSet = {};
  classRows.forEach(function(row) {
    var code = String(row['subject_code'] || row['subjectCode'] || '').trim();
    var name = String(row['subject_name'] || '').trim();
    var subType = typeMap[code] || 'พื้นฐาน';
    if (subType === 'กิจกรรม') return;
    if (name && !subjectSet[name]) {
      subjectSet[name] = code;
    }
  });

  var coverOrder = ['ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','สังคมศึกษา','ประวัติศาสตร์','สุขศึกษา','ศิลปะ','การงาน','ภาษาอังกฤษ','หน้าที่พลเมือง'];
  var subjectNames = Object.keys(subjectSet).sort(function(a, b) {
    var ia = -1, ib = -1;
    for (var k = 0; k < coverOrder.length; k++) {
      if (a.indexOf(coverOrder[k]) !== -1) ia = k;
      if (b.indexOf(coverOrder[k]) !== -1) ib = k;
    }
    if (ia === -1) ia = 999;
    if (ib === -1) ib = 999;
    return ia - ib;
  });

  var numSubjects = subjectNames.length;
  var numCols = 3 + numSubjects + 2; // ที่, รหัส, ชื่อ, ...วิชา..., GPA, อันดับ

  // --- คำนวณเกรดรายนักเรียน ---
  var studentGradeMap = {}; // studentId -> { subjectName: grade, gpa: x, rank: y }
  var studentGroups = {};
  classRows.forEach(function(row) {
    var sid = String(row['studentId'] || row['student_id'] || '').trim();
    if (!studentGroups[sid]) studentGroups[sid] = [];
    studentGroups[sid].push(row);
  });

  // คำนวณ GPA ของแต่ละคน
  var gpaList = [];
  Object.keys(studentGroups).forEach(function(sid) {
    var scores = studentGroups[sid];
    var totCredits = 0, totGradePoints = 0;
    var gradeBySubject = {};

    scores.forEach(function(subject) {
      var code = String(subject['subject_code'] || subject['subjectCode'] || '').trim();
      var name = String(subject['subject_name'] || '').trim();
      var subType = typeMap[code] || 'พื้นฐาน';
      if (subType === 'กิจกรรม') return;

      var credits = creditMap[code] || 1;
      var gpa = 0;
      var finalGrade = parseFloat(subject['final_grade'] || subject['grade_result'] || '');
      if (!isNaN(finalGrade)) {
        gpa = finalGrade;
      } else {
        var avg = parseFloat(subject['average'] || subject['total'] || 0);
        if (avg >= 80) gpa = 4; else if (avg >= 75) gpa = 3.5; else if (avg >= 70) gpa = 3;
        else if (avg >= 65) gpa = 2.5; else if (avg >= 60) gpa = 2; else if (avg >= 55) gpa = 1.5;
        else if (avg >= 50) gpa = 1; else gpa = 0;
      }

      if (name) gradeBySubject[name] = gpa;
      totCredits += credits;
      totGradePoints += gpa * credits;
    });

    var gpaVal = totCredits > 0 ? totGradePoints / totCredits : 0;
    studentGradeMap[sid] = { grades: gradeBySubject, gpa: gpaVal };
    gpaList.push({ id: sid, gpa: gpaVal });
  });

  // คำนวณอันดับ
  gpaList.sort(function(a, b) { return b.gpa - a.gpa; });
  gpaList.forEach(function(item, idx) {
    if (studentGradeMap[item.id]) {
      studentGradeMap[item.id].rank = idx + 1;
    }
  });

  // --- โลโก้ ---
  sheet.setRowHeight(r, 60);
  sheet.getRange(r, 1, 1, numCols).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 580);
  r++;

  // --- หัวเรื่อง ---
  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('แบบสรุปผลการเรียนรู้ตามกลุ่มสาระการเรียนรู้')
    .setFontFamily('Sarabun').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 28);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 22);
  r++;

  // --- Header: ที่ | รหัส | ชื่อ-สกุล | วิชา... | GPA | อันดับ ---
  var hdr = ['ที่', 'รหัส', 'ชื่อ-นามสกุล'];
  subjectNames.forEach(function(name) {
    // ชื่อย่อ: ตัดให้สั้นถ้ายาวเกิน
    var short = name.length > 8 ? name.substring(0, 8) : name;
    hdr.push(short);
  });
  hdr.push('GPA');
  hdr.push('อันดับ');
  sheet.getRange(r, 1, 1, numCols).setValues([hdr]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 9);
  // หมุนชื่อวิชาเป็นแนวตั้ง
  if (numSubjects > 0) {
    sheet.getRange(r, 4, 1, numSubjects).setTextRotation(90).setVerticalAlignment('bottom');
  }
  sheet.getRange(r, 4 + numSubjects, 1, 2).setTextRotation(90).setVerticalAlignment('bottom');
  sheet.setRowHeight(r, 110);
  r++;

  // --- ข้อมูลนักเรียน ---
  students.forEach(function(s, i) {
    var sid = String(s.id || s.studentId || '').trim();
    var fullName = (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '');
    if (!fullName.trim() && s.name) fullName = s.name;
    var info = studentGradeMap[sid] || { grades: {}, gpa: 0, rank: '-' };

    var rowData = [i + 1, sid, fullName.trim()];
    subjectNames.forEach(function(subName) {
      var g = info.grades[subName];
      rowData.push(g !== undefined ? g : '-');
    });
    rowData.push(info.gpa ? info.gpa.toFixed(2) : '-');
    rowData.push(info.rank || '-');

    sheet.getRange(r, 1, 1, numCols).setValues([rowData]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, numCols), 10);
    sheet.getRange(r, 1, 1, 2).setHorizontalAlignment('center');
    sheet.getRange(r, 4, 1, numSubjects + 2).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, numCols).setBorder(true, true, true, true, true, true);
    sheet.setRowHeight(r, 22);
    r++;
  });

  // --- ลงชื่อ ---
  r += 2;
  var halfCols = Math.max(Math.floor(numCols / 2), 3);
  sheet.getRange(r, 1, 1, halfCols).merge().setValue('ลงชื่อ...............................................ครูประจำชั้น\n(' + (meta.teacherName || '...............................................') + ')\nตำแหน่ง ครู')
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, halfCols + 1, 1, numCols - halfCols).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '...............................................') + ')\n' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา'))
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // --- ปรับขนาดคอลัมน์ ---
  sheet.setColumnWidth(1, 24);   // ที่
  sheet.setColumnWidth(2, 50);   // รหัส
  sheet.setColumnWidth(3, 130);  // ชื่อ
  var subColWidth = numSubjects > 0 ? Math.max(30, Math.min(40, Math.floor(340 / numSubjects))) : 40;
  for (var c = 4; c <= 3 + numSubjects; c++) sheet.setColumnWidth(c, subColWidth);
  sheet.setColumnWidth(4 + numSubjects, 36);  // GPA
  sheet.setColumnWidth(5 + numSubjects, 36);  // อันดับ

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 10: สรุปผลการประเมินรวม (อ่านคิดเขียน + คุณลักษณะ + กิจกรรม)
// ============================================================
function pp5s_assessmentSummary_(ss, meta, students, rtwSummary, attrSummary, actSummary) {
  if (!students || students.length === 0) return null;

  var sheet = ss.insertSheet('สรุปผลประเมิน');
  var r = 1;
  var numCols = 6; // ที่, ชื่อ, อ่านคิดเขียน, คุณลักษณะ, กิจกรรม, สรุป

  // --- โลโก้ ---
  sheet.setRowHeight(r, 60);
  sheet.getRange(r, 1, 1, numCols).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 50, 580);
  r++;

  // --- หัวเรื่อง ---
  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('แบบสรุปผลการประเมินคุณลักษณะอันพึงประสงค์ การอ่านคิดเขียน และกิจกรรมพัฒนาผู้เรียน')
    .setFontFamily('Sarabun').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sheet.setRowHeight(r, 36);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue(meta.schoolName + ' | ชั้น' + meta.gradeFullName + '/' + meta.classNo + ' | ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setFontWeight('normal');
  sheet.setRowHeight(r, 22);
  r++;

  // --- Header ---
  var hdr = ['ที่', 'ชื่อ-นามสกุล', 'อ่าน คิด เขียน', 'คุณลักษณะอันพึงประสงค์', 'กิจกรรมพัฒนาผู้เรียน', 'สรุปผลประเมิน'];
  sheet.getRange(r, 1, 1, numCols).setValues([hdr]);
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 10);
  sheet.setRowHeight(r, 30);
  r++;

  // --- สร้าง lookup maps ---
  var rtwMap = {};
  if (rtwSummary && rtwSummary.students) {
    rtwSummary.students.forEach(function(s) {
      rtwMap[String(s.studentId || s.id || '').trim()] = s.result || s.summary || '-';
    });
  }
  var attrMap = {};
  if (attrSummary && attrSummary.students) {
    attrSummary.students.forEach(function(s) {
      attrMap[String(s.studentId || s.id || '').trim()] = s.result || s.summary || '-';
    });
  }
  var actMap = {};
  if (actSummary && actSummary.students) {
    actSummary.students.forEach(function(s) {
      actMap[String(s.studentId || s.id || '').trim()] = s.result || s.summary || '-';
    });
  }

  // --- ข้อมูลนักเรียน ---
  students.forEach(function(s, i) {
    var sid = String(s.id || s.studentId || '').trim();
    var fullName = (s.title || '') + (s.firstname || '') + ' ' + (s.lastname || '');
    if (!fullName.trim() && s.name) fullName = s.name;

    var rtw = rtwMap[sid] || '-';
    var attr = attrMap[sid] || '-';
    var act = actMap[sid] || '-';

    // สรุปรวม: "ผ่าน" ถ้าทุกหมวดผ่าน
    var isPass = true;
    [rtw, attr, act].forEach(function(v) {
      var vl = String(v).trim().toLowerCase();
      if (vl === '-' || vl === '' || vl === 'มผ' || vl === 'ไม่ผ่าน' || vl === 'ปรับปรุง') isPass = false;
    });
    var overall = isPass ? 'ผ่าน' : 'ไม่ผ่าน';

    var rowData = [i + 1, fullName.trim(), rtw, attr, act, overall];
    sheet.getRange(r, 1, 1, numCols).setValues([rowData]);
    pp5s_baseStyle_(sheet.getRange(r, 1, 1, numCols), 11);
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    sheet.getRange(r, 3, 1, 4).setHorizontalAlignment('center');
    sheet.getRange(r, 1, 1, numCols).setBorder(true, true, true, true, true, true);
    // เน้นสีแดงถ้าไม่ผ่าน
    if (!isPass) {
      sheet.getRange(r, 6).setFontColor('#dc2626').setFontWeight('bold');
    }
    sheet.setRowHeight(r, 24);
    r++;
  });

  // --- ลงชื่อ ---
  r += 2;
  sheet.getRange(r, 1, 1, 3).merge().setValue('ลงชื่อ...............................................ครูประจำชั้น\n(' + (meta.teacherName || '...............................................') + ')\nตำแหน่ง ครู')
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.getRange(r, 4, 1, 3).merge().setValue('ลงชื่อ...............................................ผู้อำนวยการ\n(' + (meta.directorName || '...............................................') + ')\n' + (meta.directorTitle || 'ผู้อำนวยการสถานศึกษา'))
    .setFontFamily('Sarabun').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('top').setFontWeight('normal').setWrap(true);
  sheet.setRowHeight(r, 70);

  // --- ปรับขนาดคอลัมน์ ---
  sheet.setColumnWidth(1, 28);   // ที่
  sheet.setColumnWidth(2, 170);  // ชื่อ
  sheet.setColumnWidth(3, 95);   // อ่านคิดเขียน
  sheet.setColumnWidth(4, 95);   // คุณลักษณะ
  sheet.setColumnWidth(5, 95);   // กิจกรรม
  sheet.setColumnWidth(6, 80);   // สรุป

  return sheet;
}


// ============================================================
// 📄 ส่วนที่ 4D: สรุปเวลาเรียน ปีการศึกษา (รายเดือน ภาค1+2 รวมปี ร้อยละ)
// ============================================================
function pp5s_yearlyAttendanceSummary_(ss, meta, students, cachedMonthlyData) {
  var sheet = ss.insertSheet('สรุปเวลาเรียนปี');
  var r = 1;
  var academicYear = Number(meta.academicYear);
  var yearCE = academicYear - 543;
  cachedMonthlyData = cachedMonthlyData || {};

  // เดือนภาค 1 (พ.ค.-ต.ค.) + ภาค 2 (พ.ย.-มี.ค.)
  var sem1Months = [
    { m: 5, y: yearCE, label: 'พ.ค.' },
    { m: 6, y: yearCE, label: 'มิ.ย.' },
    { m: 7, y: yearCE, label: 'ก.ค.' },
    { m: 8, y: yearCE, label: 'ส.ค.' },
    { m: 9, y: yearCE, label: 'ก.ย.' },
    { m: 10, y: yearCE, label: 'ต.ค.' }
  ];
  var sem2Months = [
    { m: 11, y: yearCE, label: 'พ.ย.' },
    { m: 12, y: yearCE, label: 'ธ.ค.' },
    { m: 1, y: yearCE + 1, label: 'ม.ค.' },
    { m: 2, y: yearCE + 1, label: 'ก.พ.' },
    { m: 3, y: yearCE + 1, label: 'มี.ค.' }
  ];
  var allMonths = sem1Months.concat(sem2Months);

  // ใช้ cachedMonthlyData ที่ดึงไว้แล้ว (ไม่ต้องเรียก API ซ้ำ)
  var monthlyDataMap = {};
  students.forEach(function(s) {
    var sid = String(s.student_id || s.studentId || s.id || '').trim();
    monthlyDataMap[sid] = { months: {}, sem1: 0, sem2: 0, total: 0 };
  });

  allMonths.forEach(function(mo, moIdx) {
    var monthKey = mo.m + '_' + mo.y;
    var mData = cachedMonthlyData[monthKey];
    if (mData && mData.length > 0) {
      mData.forEach(function(stuData) {
        var sid = String(stuData.studentId || '').trim();
        if (monthlyDataMap[sid]) {
          var present = stuData.present || 0;
          monthlyDataMap[sid].months[monthKey] = present;
          if (moIdx < 6) {
            monthlyDataMap[sid].sem1 += present;
          } else {
            monthlyDataMap[sid].sem2 += present;
          }
          monthlyDataMap[sid].total += present;
          if (!monthlyDataMap[sid].totalDays) monthlyDataMap[sid].totalDays = 0;
          monthlyDataMap[sid].totalDays += (present + (stuData.leave || 0) + (stuData.absent || 0));
        }
      });
    }
  });

  // คอลัมน์: เลขที่, ชื่อ-นามสกุล, พ.ค.(ปี), มิ.ย., ก.ค., ส.ค., ก.ย., ต.ค., รวมภาค1, พ.ย., ธ.ค., ม.ค., ก.พ., มี.ค., รวมภาค2, รวมทั้งปี, ร้อยละ
  var numCols = 17;

  // โลโก้ (totalWidth = 28+160+35*6+35+35*5+35+35+45 = 723)
  sheet.setRowHeight(r, 55);
  sheet.getRange(r, 1, 1, numCols).merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
  pp5s_insertLogo_(sheet, r, 1, meta.logoDataUri, 45, 723);
  r++;

  // หัวเรื่อง
  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('สรุปเวลาเรียน ปีการศึกษา ' + meta.academicYear)
    .setFontFamily('Sarabun').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 28);
  r++;

  sheet.getRange(r, 1, 1, numCols).merge();
  sheet.getRange(r, 1).setValue('ชั้น ' + meta.gradeFullName + ' ห้อง ' + meta.classNo)
    .setFontFamily('Sarabun').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(r, 26);
  r++;

  // Header Row 1: กลุ่มภาคเรียน
  var hdr1 = sheet.getRange(r, 1, 1, numCols);
  // เลขที่
  sheet.getRange(r, 1).setValue('เลข\nที่');
  // ชื่อ-นามสกุล
  sheet.getRange(r, 2).setValue('ชื่อ-นามสกุล');
  // ภาคเรียนที่ 1
  sheet.getRange(r, 3, 1, 6).merge().setValue('ภาคเรียนที่ 1');
  // รวมภาค 1
  sheet.getRange(r, 9).setValue('รวม\nภาค\n1');
  // ภาคเรียนที่ 2
  sheet.getRange(r, 10, 1, 5).merge().setValue('ภาคเรียนที่ 2');
  // รวมภาค 2
  sheet.getRange(r, 15).setValue('รวม\nภาค\n2');
  // รวมทั้งปี
  sheet.getRange(r, 16).setValue('รวม\nทั้ง\nปี');
  // ร้อยละ
  sheet.getRange(r, 17).setValue('ร้อยละ');
  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 10);
  sheet.getRange(r, 1, 1, numCols).setWrap(true);
  sheet.setRowHeight(r, 36);
  r++;

  // Header Row 2: ชื่อเดือน + ปี
  var monthLabels = [];
  sem1Months.forEach(function(mo) { monthLabels.push(mo.label + '\n' + (mo.y + 543)); });
  // ข้ามรวมภาค1
  sem2Months.forEach(function(mo) { monthLabels.push(mo.label + '\n' + (mo.y + 543)); });

  sheet.getRange(r, 1).setValue('');
  sheet.getRange(r, 2).setValue('');
  for (var mi = 0; mi < 6; mi++) {
    sheet.getRange(r, 3 + mi).setValue(monthLabels[mi]);
  }
  sheet.getRange(r, 9).setValue('');
  for (var mi2 = 0; mi2 < 5; mi2++) {
    sheet.getRange(r, 10 + mi2).setValue(monthLabels[6 + mi2]);
  }
  sheet.getRange(r, 15).setValue('');
  sheet.getRange(r, 16).setValue('');
  sheet.getRange(r, 17).setValue('');

  // merge เลขที่ + ชื่อ กับ row ก่อนหน้า
  sheet.getRange(r - 1, 1, 2, 1).merge();
  sheet.getRange(r - 1, 2, 2, 1).merge();
  sheet.getRange(r - 1, 9, 2, 1).merge();
  sheet.getRange(r - 1, 15, 2, 1).merge();
  sheet.getRange(r - 1, 16, 2, 1).merge();
  sheet.getRange(r - 1, 17, 2, 1).merge();

  pp5s_headerStyle_(sheet.getRange(r, 1, 1, numCols), 9);
  sheet.getRange(r, 1, 1, numCols).setWrap(true);
  sheet.setRowHeight(r, 30);
  r++;

  // Data rows
  students.forEach(function(stu, idx) {
    var sid = String(stu.student_id || stu.studentId || stu.id || '').trim();
    var name = String(stu.title || '') + String(stu.firstname || '') + ' ' + String(stu.lastname || '');
    name = name.trim();
    var data = monthlyDataMap[sid] || { months: {}, sem1: 0, sem2: 0, total: 0, totalDays: 0 };

    var vals = [idx + 1, name];
    // ภาค 1 (6 เดือน)
    sem1Months.forEach(function(mo) {
      var key = mo.m + '_' + mo.y;
      vals.push(data.months[key] || 0);
    });
    vals.push(data.sem1); // รวมภาค1

    // ภาค 2 (5 เดือน)
    sem2Months.forEach(function(mo) {
      var key = mo.m + '_' + mo.y;
      vals.push(data.months[key] || 0);
    });
    vals.push(data.sem2); // รวมภาค2
    vals.push(data.total); // รวมทั้งปี

    var totalDays = data.totalDays || 0;
    var pct = totalDays > 0 ? ((data.total / totalDays) * 100).toFixed(1) + '%' : '0.0%';
    vals.push(pct);

    var row = sheet.getRange(r, 1, 1, numCols);
    row.setValues([vals]);
    pp5s_baseStyle_(row, 11);
    sheet.getRange(r, 2).setHorizontalAlignment('left');
    sheet.getRange(r, 1).setHorizontalAlignment('center');
    for (var ci = 3; ci <= numCols; ci++) sheet.getRange(r, ci).setHorizontalAlignment('center');
    sheet.setRowHeight(r, 22);
    r++;
  });

  // ตั้งค่าความกว้างคอลัมน์
  sheet.setColumnWidth(1, 28);
  sheet.setColumnWidth(2, 160);
  for (var c = 3; c <= 8; c++) sheet.setColumnWidth(c, 35);
  sheet.setColumnWidth(9, 35);
  for (var c2 = 10; c2 <= 14; c2++) sheet.setColumnWidth(c2, 35);
  sheet.setColumnWidth(15, 35);
  sheet.setColumnWidth(16, 35);
  sheet.setColumnWidth(17, 45);

  // border ทั้งหมด
  var dataRange = sheet.getRange(4, 1, r - 4, numCols);
  dataRange.setBorder(true, true, true, true, true, true);

  return sheet;
}


// ============================================================
// � MAIN: สร้าง ปพ.5 รวมเล่ม (Google Sheets → PDF)
// ============================================================
async function exportPp5FullBookSheets(grade, classNo, parts) {
  try {
    // default: สร้างทุกส่วน ถ้าไม่ส่ง parts มา
    // parts อาจเป็น JSON string จาก frontend
    if (typeof parts === 'string') {
      try { parts = JSON.parse(parts); } catch (e) { parts = null; }
    }
    if (!parts || typeof parts !== 'object') {
      parts = { part1: true, part2: true, part3: true };
    }
    // normalize: แปลง truthy/falsy เป็น boolean จริงๆ
    parts = {
      part1: parts.part1 === true || parts.part1 === 'true',
      part2: parts.part2 === true || parts.part2 === 'true',
      part3: parts.part3 === true || parts.part3 === 'true'
    };
    Logger.log('📖 เริ่มสร้าง ปพ.5 (Sheets→PDF): ' + grade + '/' + classNo + ' parts=' + JSON.stringify(parts));
    pp5fb_setProgress_(0, 1, 'เริ่มต้น...');
    grade = String(grade || '').trim();
    classNo = String(classNo || '').trim();
    if (!grade || !classNo) throw new Error('กรุณาระบุชั้นและห้อง');

    // --- ดึงข้อมูลพื้นฐาน (เหมือนเดิม) ---
    pp5fb_setProgress_(3, 1, 'ดึงข้อมูลนักเรียน...');
    var settings = S_getGlobalSettings(false);
    var academicYear = settings['ปีการศึกษา'] || '';
    var schoolName = settings['ชื่อโรงเรียน'] || '';
    var logoDataUri = '';
    try { logoDataUri = _getLogoDataUrl(settings.logoFileId || ''); } catch (e) {}
    var teacherName = settings['ชื่อครูประจำชั้น'] || getClassTeacherName(grade, classNo) || '';
    var directorName = settings['ชื่อผู้อำนวยการ'] || '';
    var directorTitle = settings['ตำแหน่งผู้อำนวยการ'] || 'ผู้อำนวยการโรงเรียน';
    var academicHead = settings['ชื่อหัวหน้างานวิชาการ'] || '';
    var gradeFullName = _pp5GradeFullName_(grade);

    var students = getFilteredStudentsInline(grade, classNo);
    if (!students || students.length === 0) throw new Error('ไม่พบนักเรียนในชั้น ' + grade + ' ห้อง ' + classNo);
    pp5fb_setProgress_(5, 1, 'พบนักเรียน ' + students.length + ' คน');

    // ดึงข้อมูลเฉพาะส่วนที่เลือก
    var scoreSheets = [], classScoreSheets = [];
    var subjectTeacherMap = {};
    if (parts.part1 || parts.part2) {
      pp5fb_setProgress_(8, 2, 'ดึงคะแนนทุกวิชา...');
      scoreSheets = getExistingScoreSheets();
      var cleanGrade = grade.replace(/[\/\.]/g, '');
      classScoreSheets = scoreSheets.filter(function(s) {
        var sheetGrade = String(s.grade).trim().replace(/[\/\.]/g, '');
        return (sheetGrade === cleanGrade || sheetGrade === grade) && String(s.classNo).trim() === classNo;
      });

      // ดึงชื่อครูผู้สอนจากชีต "รายวิชา" (คอลัมน์ F)
      try {
        var subjectSheet = SS().getSheetByName('รายวิชา');
        if (subjectSheet) {
          var subData = subjectSheet.getDataRange().getValues();
          for (var si = 1; si < subData.length; si++) {
            var subCode = String(subData[si][1] || '').trim(); // B: รหัสวิชา
            var subTeacher = String(subData[si][5] || '').trim(); // F: ครูผู้สอน
            if (subCode && subTeacher) {
              subjectTeacherMap[subCode] = subTeacher;
            }
          }
        }
      } catch (e) { Logger.log('⚠️ ดึงครูผู้สอนจากรายวิชา: ' + e.message); }

      // แนบชื่อครูเข้า sheetInfo
      classScoreSheets.forEach(function(si) {
        si.teacherName = subjectTeacherMap[String(si.subjectCode || '').trim()] || '';
      });
    }

    var attendance1 = [], attendance2 = [];
    if (parts.part1) {
      pp5fb_setProgress_(12, 3, 'ดึงเวลาเรียน...');
      try { attendance1 = getSemesterAttendanceSummary(grade, classNo, 1, Number(academicYear)); } catch (e) {}
      try { attendance2 = getSemesterAttendanceSummary(grade, classNo, 2, Number(academicYear)); } catch (e) {}
    }

    var charData = [], actData = [], rtwSummary = {}, attrSummary = {}, actSummary = {};
    var subjectScoreStudents = [];
    var subjectSummary = [];
    var warehouseData = [], subjectListData = [];
    if (parts.part1 || parts.part3) {
      pp5fb_setProgress_(16, 4, 'ดึงข้อมูลการประเมิน...');
      try { charData = getStudentsForCharacteristic(grade, classNo); } catch (e) {}
      try { actData = getStudentsForActivity(grade, classNo); } catch (e) {}
      try { rtwSummary = getReadingSummary(grade, classNo); } catch (e) {}
      try { attrSummary = getAttributeSummary(grade, classNo); } catch (e) {}
      try { actSummary = getActivitySummary(grade, classNo); } catch (e) {}
      try { subjectScoreStudents = getStudentsForSubjectScore(grade, classNo); } catch (e) {}
      try { subjectSummary = getSubjectScoreSummary(grade, classNo); } catch (e) {}
      try { warehouseData = _readSheetToObjects('SCORES_WAREHOUSE'); } catch (e) { Logger.log('⚠️ SCORES_WAREHOUSE: ' + e.message); }
      try { subjectListData = _readSheetToObjects('รายวิชา'); } catch (e) { Logger.log('⚠️ รายวิชา: ' + e.message); }
    }

    var meta = {
      grade: grade, classNo: classNo, gradeFullName: gradeFullName,
      schoolName: schoolName, academicYear: academicYear,
      logoDataUri: logoDataUri, teacherName: teacherName,
      directorName: directorName, directorTitle: directorTitle,
      academicHead: academicHead, settings: settings
    };

    pp5fb_setProgress_(20, 5, 'ดึงข้อมูลครบแล้ว เริ่มสร้าง PDF...');

    // --- สร้าง temp Spreadsheet ---
    var tempSS = SpreadsheetApp.create('_pp5temp_' + grade + '_' + classNo + '_' + Date.now());
    Logger.log('📄 สร้าง temp spreadsheet: ' + tempSS.getId());

    var startTime = new Date();
    Logger.log('⏱ เริ่ม: ' + startTime.toISOString());

    // ======= ส่วนที่ 1: ปก + ข้อมูลนักเรียน + เวลาเรียน =======
    var coverPdfBlob = null;
    var defaultSheetDeleted = false;
    if (parts.part1) {
      pp5fb_setProgress_(22, 5, 'สร้างหน้าปก...');
      try {
        coverPdfBlob = pp5s_coverDoc_(meta, students.length, subjectSummary, rtwSummary, attrSummary, actSummary);
        Logger.log('📄 coverPdfBlob: ' + (coverPdfBlob ? coverPdfBlob.getBytes().length + ' bytes' : 'NULL'));
      } catch (e) { Logger.log('❌ ปก ERROR: ' + e.message + '\n' + e.stack); }

      pp5fb_setProgress_(26, 5, 'สร้างตารางรายชื่อนักเรียน...');
      try {
        pp5s_studentList_(tempSS, meta, students);
        Logger.log('✅ รายชื่อ');
      } catch (e) { Logger.log('⚠️ รายชื่อ: ' + e.message); }

      // --- ลบ default sheet ทันทีหลังมี sheet แรก ---
      try {
        var sheets0 = tempSS.getSheets();
        for (var di = sheets0.length - 1; di >= 0; di--) {
          if (sheets0[di].getName() !== 'รายชื่อ' && tempSS.getSheets().length > 1) {
            Logger.log('🗑️ ลบ default sheet: ' + sheets0[di].getName());
            tempSS.deleteSheet(sheets0[di]);
          }
        }
        defaultSheetDeleted = true;
      } catch (e) { Logger.log('⚠️ ลบ default: ' + e.message); }

      pp5fb_setProgress_(30, 5, 'สร้างตารางผู้ปกครอง...');
      try {
        pp5s_parentInfo_(tempSS, meta, students);
        Logger.log('✅ ผู้ปกครอง');
      } catch (e) { Logger.log('⚠️ ผู้ปกครอง: ' + e.message); }
    } else {
      Logger.log('⏭ ข้ามส่วนที่ 1 (ปก+เวลาเรียน)');
    }

    // ส่วนที่ 4A: สรุปเวลาเรียน
    if (parts.part1) {
      pp5fb_setProgress_(34, 5, 'สร้างสรุปเวลาเรียน...');
      try {
        pp5s_attendanceSummary_(tempSS, meta, students, attendance1, attendance2);
        Logger.log('✅ สรุปเวลาเรียน');
      } catch (e) { Logger.log('⚠️ สรุปเวลาเรียน: ' + e.message); }

      // ส่วนที่ 4B+4C: เช็คชื่อรายวัน + สรุปรายเดือน
      pp5fb_setProgress_(38, 5, 'สร้างเวลาเรียนรายเดือน...');
    try {
      var yearCE = Number(academicYear) - 543;
      var acadMonths = [
        { m: 5, y: yearCE, name: 'พฤษภาคม', m0: 4 },
        { m: 6, y: yearCE, name: 'มิถุนายน', m0: 5 },
        { m: 7, y: yearCE, name: 'กรกฎาคม', m0: 6 },
        { m: 8, y: yearCE, name: 'สิงหาคม', m0: 7 },
        { m: 9, y: yearCE, name: 'กันยายน', m0: 8 },
        { m: 10, y: yearCE, name: 'ตุลาคม', m0: 9 },
        { m: 11, y: yearCE, name: 'พฤศจิกายน', m0: 10 },
        { m: 12, y: yearCE, name: 'ธันวาคม', m0: 11 },
        { m: 1, y: yearCE + 1, name: 'มกราคม', m0: 0 },
        { m: 2, y: yearCE + 1, name: 'กุมภาพันธ์', m0: 1 },
        { m: 3, y: yearCE + 1, name: 'มีนาคม', m0: 2 }
      ];

      // เก็บ monthly data ไว้ใช้ซ้ำกับสรุปเวลาเรียนปี
      var cachedMonthlyData = {};

      // สร้างสลับ: เช็คชื่อรายวัน (daily) → สรุปรายเดือน (monthly) ในแต่ละเดือน
      acadMonths.forEach(function(mo, moIdx) {
        pp5fb_setProgress_(38 + Math.round(moIdx / acadMonths.length * 17), 5, 'เดือน' + mo.name + '...');

        // 1) เช็คชื่อรายวัน (daily) ก่อน
        try {
          var tableData = getAttendanceVerticalTableFiltered(grade, classNo, mo.y, mo.m0);
          if (tableData && tableData.students && tableData.students.length > 0 && tableData.days && tableData.days.length > 0) {
            var semester = (mo.m0 >= 4 && mo.m0 <= 9) ? 'ภาคเรียนที่ 1' : 'ภาคเรียนที่ 2';
            var dailySheet = pp5s_dailyAttendance_(tempSS, meta, tableData, mo.name, mo.y + 543, semester);
            if (dailySheet) Logger.log('✅ รายวัน ' + mo.name);
          }
        } catch (e) { Logger.log('⚠️ ข้ามรายวัน ' + mo.name + ': ' + e.message); }

        // 2) สรุปรายเดือน (monthly) ตามหลัง
        try {
          var mData = getMonthlyAttendanceSummary(grade, classNo, mo.y, mo.m);
          if (mData && mData.length > 0) {
            cachedMonthlyData[mo.m + '_' + mo.y] = mData;
            var monthInfo = { month: mo.m, year: mo.y, name: mo.name, yearBE: mo.y + 543, data: mData };
            pp5s_monthlyAttendance_(tempSS, meta, students, [monthInfo]);
            Logger.log('✅ สรุปเดือน ' + mo.name);
          }
        } catch (e) { Logger.log('⚠️ ข้ามสรุปเดือน ' + mo.name + ': ' + e.message); }
      });
    } catch (e) { Logger.log('⚠️ เวลาเรียน: ' + e.message); }

      // ส่วนที่ 4D: สรุปเวลาเรียนปีการศึกษา (รายเดือน ภาค1+2 รวมปี ร้อยละ)
      pp5fb_setProgress_(54, 5, 'สร้างสรุปเวลาเรียนรายปี...');
      try {
        pp5s_yearlyAttendanceSummary_(tempSS, meta, students, cachedMonthlyData);
        Logger.log('✅ สรุปเวลาเรียนรายปี');
      } catch (e) { Logger.log('⚠️ สรุปเวลาเรียนรายปี: ' + e.message); }
    } // end parts.part1 เวลาเรียน

    // ลบ default sheet (ถ้ายังไม่ได้ลบ)
    if (!defaultSheetDeleted) {
      try {
        var sheets0b = tempSS.getSheets();
        if (sheets0b.length > 1) {
          for (var di2 = sheets0b.length - 1; di2 >= 0; di2--) {
            if (sheets0b[di2].getName() === 'Sheet1' && tempSS.getSheets().length > 1) {
              tempSS.deleteSheet(sheets0b[di2]);
              break;
            }
          }
        }
      } catch (e) {}
    }

    // ======= ส่วนที่ 2: คะแนนรายวิชา =======
    // ส่วนที่ 5: คะแนนรายวิชา
    if (parts.part2) {
    pp5fb_setProgress_(55, 5, 'สร้างผลการเรียนรายวิชา...');
    _sortBySubjectName(classScoreSheets, 'subjectName');

    var activityKeywords = ['กิจกรรม', 'แนะแนว', 'ลูกเสือ', 'เนตรนารี', 'ชุมนุม', 'เพื่อสังคม', 'ชมรม'];
    var processedSubjects = {};
    var subIdx = 0;
    var totalSubs = classScoreSheets.length;
    classScoreSheets.forEach(function(sheetInfo) {
      try {
        var subName = (sheetInfo.subjectName || '').trim();
        var isActivity = activityKeywords.some(function(kw) { return subName.indexOf(kw) !== -1; });
        if (isActivity) return;
        if (processedSubjects[subName]) { Logger.log('⏭ ข้าม (ซ้ำ): ' + subName); return; }
        processedSubjects[subName] = true;
        subIdx++;
        pp5fb_setProgress_(55 + Math.round(subIdx / Math.max(totalSubs, 1) * 20), 5, 'สร้างผลวิชา: ' + subName);
        var scoreData = getScoreSheetData(sheetInfo.sheetName);
        var scoreSheet = pp5s_yearScore_(tempSS, meta, sheetInfo, scoreData);
        if (scoreSheet) Logger.log('✅ วิชา: ' + subName);
      } catch (e) { Logger.log('⚠️ วิชา ' + sheetInfo.subjectName + ': ' + e.message); }
    });
    } else {
      Logger.log('⏭ ข้ามส่วนที่ 2 (คะแนนรายวิชา)');
    }

    // ======= ส่วนที่ 3: ผลการประเมิน =======
    // ส่วนที่ 6: คุณลักษณะ
    if (parts.part3) {
      pp5fb_setProgress_(78, 5, 'สร้างตารางคุณลักษณะ...');
      try {
        pp5s_characteristic_(tempSS, meta, charData);
        Logger.log('✅ คุณลักษณะ');
      } catch (e) { Logger.log('⚠️ คุณลักษณะ: ' + e.message); }

      // ส่วนที่ 7: กิจกรรม
      pp5fb_setProgress_(82, 5, 'สร้างตารางกิจกรรม...');
      try {
        pp5s_activity_(tempSS, meta, actData);
        Logger.log('✅ กิจกรรม');
      } catch (e) { Logger.log('⚠️ กิจกรรม: ' + e.message); }

      // ส่วนที่ 8: อ่าน คิดวิเคราะห์ เขียน
      pp5fb_setProgress_(85, 5, 'สร้างตารางอ่านคิดเขียน...');
      try {
        pp5s_readThinkWrite_(tempSS, meta, subjectScoreStudents);
        Logger.log('✅ อ่านคิดเขียน');
      } catch (e) { Logger.log('⚠️ อ่านคิดเขียน: ' + e.message); }

      // ส่วนที่ 9: สรุปผลการเรียนรวม (เกรดทุกวิชา + GPA + อันดับ)
      pp5fb_setProgress_(87, 5, 'สร้างตารางสรุปผลการเรียน...');
      try {
        pp5s_gradesSummary_(tempSS, meta, students, warehouseData, subjectListData);
        Logger.log('✅ สรุปผลการเรียน');
      } catch (e) { Logger.log('⚠️ สรุปผลการเรียน: ' + e.message); }

      // ส่วนที่ 10: สรุปผลการประเมินรวม
      pp5fb_setProgress_(89, 5, 'สร้างตารางสรุปผลประเมิน...');
      try {
        pp5s_assessmentSummary_(tempSS, meta, students, rtwSummary, attrSummary, actSummary);
        Logger.log('✅ สรุปผลประเมิน');
      } catch (e) { Logger.log('⚠️ สรุปผลประเมิน: ' + e.message); }
    } else {
      Logger.log('⏭ ข้ามส่วนที่ 3 (ผลการประเมิน)');
    }

    // --- Log sheets ทั้งหมดก่อน export ---
    var allSheets = tempSS.getSheets();
    Logger.log('📄 sheets ทั้งหมด (' + allSheets.length + '): ' + allSheets.map(function(s){return s.getName();}).join(', '));
    var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log('⏱ สร้าง sheets เสร็จ: ' + elapsed + 's');

    // --- Export ทั้ง Spreadsheet เป็น PDF 1 ไฟล์ (เร็วที่สุด) ---
    pp5fb_setProgress_(90, 5, 'กำลัง export PDF...');
    Logger.log('📄 เริ่ม export ทั้ง Spreadsheet (เนื้อหา) เป็น 1 PDF blob...');
    var contentPdfBlob = null;
    try {
      // ใช้ opts fitw=true หรือ false ตามต้องการ (ส่วนใหญ่ fitw=true ช่วยไม่ให้ตารางตกขอบ)
      contentPdfBlob = pp5s_exportAllAsPdf_(tempSS, { top: 0.5, fitw: true });
      Logger.log('✅ Export เนื้อหาสำเร็จ');
    } catch (e) {
      Logger.log('❌ Export เนื้อหา ERROR: ' + e.message);
      throw e;
    }

    // รวม: ปก (จาก Doc) + เนื้อหาทั้งหมด (จาก Spreadsheet)
    var allPdfBlobs = [];
    if (coverPdfBlob) allPdfBlobs.push(coverPdfBlob);
    if (contentPdfBlob) allPdfBlobs.push(contentPdfBlob);
    pp5fb_setProgress_(95, 5, 'กำลังรวมไฟล์ PDF...');
    Logger.log('📄 รวม ' + allPdfBlobs.length + ' PDF blobs (ปก + เนื้อหาทั้งหมด), เริ่ม merge...');

    var fileName = 'ปพ.5 รวมเล่ม_' + grade + '_' + classNo + '_' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '.pdf';
    var pdfBlob = await pp5s_mergePdfs_(allPdfBlobs, fileName);

    var file = DriveApp.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ลบ temp spreadsheet
    try { DriveApp.getFileById(tempSS.getId()).setTrashed(true); } catch (e) {}

    pp5fb_setProgress_(100, 5, 'เสร็จสมบูรณ์!');
    Logger.log('✅ สร้าง ปพ.5 รวมเล่มสำเร็จ (Sheets→PDF): ' + file.getUrl());
    return file.getUrl();

  } catch (error) {
    Logger.log('❌ exportPp5FullBookSheets error: ' + error.message);
    throw new Error(error.message);
  }
}
