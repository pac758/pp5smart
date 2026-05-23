// ============================================================
// ACADEMIC YEAR ARCHIVE
// Stage 1 separation: create year-suffixed snapshots only.
// This does not rename live sheets, delete data, promote students,
// or change global_settings.
// ============================================================

var AY_ARCHIVE_MAX_SHEET_NAME = 100;

function AY_archiveTargetYear_(targetYear) {
  var year = (typeof AY_normalizeYear_ === 'function')
    ? AY_normalizeYear_(targetYear)
    : String(targetYear || '').replace(/\D/g, '').slice(0, 4);
  if (year) return year;
  if (typeof AY_getCurrentAcademicYear === 'function') return AY_getCurrentAcademicYear(false);
  if (typeof S_getAcademicYear === 'function') return String(S_getAcademicYear());
  return '';
}

function AY_archiveSafeSheetName_(baseName, year) {
  var suffix = '_' + year;
  var raw = String(baseName || '').trim();
  var maxBase = AY_ARCHIVE_MAX_SHEET_NAME - suffix.length;
  if (raw.length > maxBase) raw = raw.slice(0, maxBase);
  return raw + suffix;
}

function AY_archiveIsSystemSheet_(name, targetYear) {
  if (!name) return true;
  if (/_\d{4}$/.test(name)) return true;
  var exact = [
    'global_settings', 'Users', 'users', 'Settings', 'Holidays', 'วันหยุด',
    'Attendance', 'BACKUP_WAREHOUSE_LATEST'
  ];
  if (exact.indexOf(name) !== -1) return true;

  var bases = [];
  if (typeof S_SHARED_SHEETS !== 'undefined') bases = bases.concat(S_SHARED_SHEETS);
  if (typeof S_YEARLY_SHEETS !== 'undefined') bases = bases.concat(S_YEARLY_SHEETS);
  for (var i = 0; i < bases.length; i++) {
    if (name === bases[i]) return true;
    if (name === AY_archiveSafeSheetName_(bases[i], targetYear)) return true;
  }

  var prefixes = ['BACKUP_', 'Template_', 'TMP_', 'Copy of '];
  return prefixes.some(function(p) { return name.indexOf(p) === 0; });
}

function AY_archiveBuildPlan_(targetYear) {
  var ss = SS();
  var year = AY_archiveTargetYear_(targetYear);
  if (!year) throw new Error('ไม่พบปีการศึกษาสำหรับสร้าง snapshot');

  var plan = {
    targetYear: year,
    sharedSheets: [],
    yearlySheets: [],
    scoreSheets: [],
    skipped: [],
    totals: {
      toCreate: 0,
      alreadyExists: 0,
      missingSource: 0
    }
  };

  function addItem(group, sourceName, targetName, type) {
    var source = ss.getSheetByName(sourceName);
    var target = ss.getSheetByName(targetName);
    var item = {
      type: type,
      source: sourceName,
      target: targetName,
      exists: !!source,
      alreadyExists: !!target,
      rows: source ? Math.max(0, source.getLastRow() - 1) : 0,
      columns: source ? source.getLastColumn() : 0
    };
    group.push(item);
    if (!source) plan.totals.missingSource++;
    else if (target) plan.totals.alreadyExists++;
    else plan.totals.toCreate++;
  }

  var sharedBases = (typeof S_SHARED_SHEETS !== 'undefined')
    ? S_SHARED_SHEETS
    : ['Students', 'รายวิชา', 'HomeroomTeachers'];
  sharedBases.forEach(function(baseName) {
    addItem(plan.sharedSheets, baseName, AY_archiveSafeSheetName_(baseName, year), 'shared');
  });

  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined')
    ? S_YEARLY_SHEETS
    : ['SCORES_WAREHOUSE', 'AttendanceLog'];
  yearlyBases.forEach(function(baseName) {
    addItem(plan.yearlySheets, baseName, AY_archiveSafeSheetName_(baseName, year), 'yearly');
  });

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (AY_archiveIsSystemSheet_(name, year)) return;
    if (typeof AY_isScoreSheet_ === 'function' && !AY_isScoreSheet_(sheet)) return;
    addItem(plan.scoreSheets, name, AY_archiveSafeSheetName_(name, year), 'score');
  });

  return plan;
}

function previewAcademicYearSnapshots(targetYear) {
  try {
    var plan = AY_archiveBuildPlan_(targetYear);
    return {
      success: true,
      readOnly: true,
      targetYear: plan.targetYear,
      plan: plan
    };
  } catch (e) {
    Logger.log('previewAcademicYearSnapshots error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function createAcademicYearSnapshots(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { success: false, message: 'ระบบกำลังทำงานอยู่ กรุณาลองใหม่อีกครั้ง' };
  }

  try {
    var ss = SS();
    var plan = AY_archiveBuildPlan_(targetYear);
    var created = [];
    var skipped = [];
    var errors = [];

    function copyItem(item) {
      if (!item.exists) {
        skipped.push(item.target + ' - ไม่พบชีตต้นทาง');
        return;
      }
      if (item.alreadyExists) {
        skipped.push(item.target + ' - มีอยู่แล้ว');
        return;
      }
      try {
        var source = ss.getSheetByName(item.source);
        var copied = source.copyTo(ss);
        copied.setName(item.target);
        created.push(item.target);
      } catch (e) {
        errors.push(item.source + ' → ' + item.target + ': ' + e.message);
      }
    }

    plan.sharedSheets.forEach(copyItem);
    plan.yearlySheets.forEach(copyItem);
    plan.scoreSheets.forEach(copyItem);

    return {
      success: errors.length === 0,
      targetYear: plan.targetYear,
      createdCount: created.length,
      skippedCount: skipped.length,
      errorCount: errors.length,
      created: created.slice(0, 80),
      skipped: skipped.slice(0, 80),
      errors: errors.slice(0, 40),
      message: errors.length
        ? 'สร้าง snapshot บางส่วนสำเร็จ แต่มีข้อผิดพลาด ' + errors.length + ' รายการ'
        : 'สร้าง snapshot ปีการศึกษา ' + plan.targetYear + ' สำเร็จ ' + created.length + ' ชีต'
    };
  } catch (e) {
    Logger.log('createAcademicYearSnapshots error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}



function AY_prepareCanCopyRow_(row, headers, targetYear) {
  if (typeof AY_rowMatchesAcademicYear !== 'function') return false;
  return AY_rowMatchesAcademicYear(row, headers, targetYear, {
    allowMissingColumn: false,
    allowBlank: false
  });
}

function AY_prepareIsActiveStudentRow_(row, headers) {
  var statusIdx = (typeof AY_findColumn_ === 'function')
    ? AY_findColumn_(headers, ['status', 'สถานะ'])
    : -1;
  var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim() : '';
  if (typeof isInactiveStudentStatus_ === 'function') return !isInactiveStudentStatus_(status);
  var st = status.toLowerCase().replace(/\s+/g, '');
  return ['จำหน่าย', 'ย้ายออก', 'พ้นสภาพ', 'จบ', 'จบการศึกษา', 'สำเร็จการศึกษา', 'ลาออก', 'inactive', 'graduated', 'transferred', 'deleted'].indexOf(st) === -1;
}



function AY_prepareSourceSheet_(ss, baseName, targetYear) {
  return ss.getSheetByName(baseName) || ss.getSheetByName(AY_archiveSafeSheetName_(baseName, targetYear));
}











function previewCurrentYearWorkingSheets(targetYear) {
  try {
    var ss = SS();
    var year = AY_archiveTargetYear_(targetYear);
    if (!year) throw new Error('ไม่พบปีการศึกษาสำหรับเตรียมชีตปีปัจจุบัน');

    var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined')
      ? S_YEARLY_SHEETS
      : ['SCORES_WAREHOUSE', 'AttendanceLog'];

    var sheets = [AY_preparePlanForSheet_(ss, 'Students', year)].concat(yearlyBases.map(function(baseName) {
      return AY_preparePlanForSheet_(ss, baseName, year);
    }));

    var totals = {
      toCreate: 0,
      alreadyExists: 0,
      explicitTargetYearRows: 0,
      blankYearRows: 0,
      missingAcademicYearColumn: 0
    };
    sheets.forEach(function(item) {
      if (item.targetExists) totals.alreadyExists++;
      else totals.toCreate++;
      totals.explicitTargetYearRows += item.explicitTargetYearRows || 0;
      totals.blankYearRows += item.blankYearRows || 0;
      if (item.missingAcademicYearColumn) totals.missingAcademicYearColumn++;
    });

    return {
      success: true,
      readOnly: true,
      targetYear: year,
      sheets: sheets,
      totals: totals,
      note: 'จะสร้างเฉพาะชีตรายปี _' + year + ' และคัดลอกเฉพาะแถวที่มี academic_year ตรงปีนี้เท่านั้น'
    };
  } catch (e) {
    Logger.log('previewCurrentYearWorkingSheets error: ' + e.message);
    return { success: false, message: e.message };
  }
}



// Final current-year preparation overrides.
// Keep this block at the end so it wins over earlier legacy definitions.
function AY_prepareFindColumnFinal_(headers, names) {
  var map = {};
  (headers || []).forEach(function(h, i) {
    var key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  for (var n = 0; n < names.length; n++) {
    var name = String(names[n] || '').trim().toLowerCase();
    if (map[name] != null) return map[name];
  }
  return -1;
}

function AY_prepareIsInactiveStatusFinal_(status) {
  var st = String(status || '').trim().toLowerCase().replace(/\s+/g, '');
  if (!st) return false;
  var inactive = [
    '\u0e08\u0e33\u0e2b\u0e19\u0e48\u0e32\u0e22',
    '\u0e22\u0e49\u0e32\u0e22\u0e2d\u0e2d\u0e01',
    '\u0e1e\u0e49\u0e19\u0e2a\u0e20\u0e32\u0e1e',
    '\u0e08\u0e1a',
    '\u0e08\u0e1a\u0e01\u0e32\u0e23\u0e28\u0e36\u0e01\u0e29\u0e32',
    '\u0e2a\u0e33\u0e40\u0e23\u0e47\u0e08\u0e01\u0e32\u0e23\u0e28\u0e36\u0e01\u0e29\u0e32',
    '\u0e25\u0e32\u0e2d\u0e2d\u0e01',
    '\u0e1e\u0e31\u0e01\u0e01\u0e32\u0e23\u0e40\u0e23\u0e35\u0e22\u0e19',
    'inactive',
    'graduated',
    'transferred',
    'deleted'
  ];
  if (inactive.indexOf(st) !== -1) return true;
  return st.indexOf('\u0e08\u0e33\u0e2b\u0e19\u0e48\u0e32\u0e22') !== -1 ||
    st.indexOf('\u0e22\u0e49\u0e32\u0e22\u0e2d\u0e2d\u0e01') !== -1 ||
    st.indexOf('\u0e08\u0e1a') !== -1 ||
    st.indexOf('inactive') !== -1 ||
    st.indexOf('graduated') !== -1 ||
    st.indexOf('transferred') !== -1;
}

function AY_prepareRowYearFinal_(row, headers) {
  var yearIdx = AY_prepareFindColumnFinal_(headers, [
    'academic_year',
    'academicyear',
    'academic year',
    '\u0e1b\u0e35\u0e01\u0e32\u0e23\u0e28\u0e36\u0e01\u0e29\u0e32'
  ]);
  return {
    index: yearIdx,
    value: yearIdx >= 0 ? AY_archiveTargetYear_(row[yearIdx]) : ''
  };
}

function AY_prepareShouldCopyRow_(baseName, row, headers, targetYear) {
  var target = String(targetYear || '');
  var yearInfo = AY_prepareRowYearFinal_(row, headers);
  var rowYear = yearInfo.value;

  if (baseName !== 'Students') {
    return !!rowYear && rowYear === target;
  }

  var statusIdx = AY_prepareFindColumnFinal_(headers, ['status', '\u0e2a\u0e16\u0e32\u0e19\u0e30']);
  var status = statusIdx >= 0 ? row[statusIdx] : '';
  if (AY_prepareIsInactiveStatusFinal_(status)) return false;

  // Legacy Students rows may have blank academic_year. Treat active blank rows
  // as the current working year only when preparing Students_<year>.
  return !rowYear || rowYear === target;
}

function AY_prepareNormalizeRowForTarget_(baseName, row, headers, targetYear) {
  var out = row.slice();
  if (baseName === 'Students') {
    var yearInfo = AY_prepareRowYearFinal_(out, headers);
    if (yearInfo.index >= 0) out[yearInfo.index] = String(targetYear);
  }
  return out;
}

function AY_preparePlanForSheet_(ss, baseName, targetYear) {
  var targetName = AY_archiveSafeSheetName_(baseName, targetYear);
  var target = ss.getSheetByName(targetName);
  var source = AY_prepareSourceSheet_(ss, baseName, targetYear);
  var item = {
    baseName: baseName,
    source: source ? source.getName() : '',
    target: targetName,
    sourceExists: !!source,
    targetExists: !!target,
    sourceRows: source ? Math.max(0, source.getLastRow() - 1) : 0,
    targetRows: target ? Math.max(0, target.getLastRow() - 1) : 0,
    explicitTargetYearRows: 0,
    blankYearRows: 0,
    missingAcademicYearColumn: false
  };

  if (!source || source.getLastRow() <= 1) return item;
  if (target) return item;

  var lastRow = source.getLastRow();
  var lastCol = source.getLastColumn();
  var headers = lastCol > 0 ? source.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var yearIdx = AY_prepareFindColumnFinal_(headers, [
    'academic_year',
    'academicyear',
    'academic year',
    '\u0e1b\u0e35\u0e01\u0e32\u0e23\u0e28\u0e36\u0e01\u0e29\u0e32'
  ]);
  item.missingAcademicYearColumn = yearIdx < 0;

  if (baseName === 'Students') {
    var data = source.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var rowYear = yearIdx >= 0 ? AY_archiveTargetYear_(row[yearIdx]) : '';
      if (!rowYear) item.blankYearRows++;
      if (AY_prepareShouldCopyRow_(baseName, row, headers, targetYear)) item.explicitTargetYearRows++;
    }
    return item;
  }

  if (yearIdx < 0) return item;
  var yearValues = source.getRange(2, yearIdx + 1, lastRow - 1, 1).getValues();
  for (var y = 0; y < yearValues.length; y++) {
    var yv = AY_archiveTargetYear_(yearValues[y][0]);
    if (!yv) item.blankYearRows++;
    if (yv === String(targetYear)) item.explicitTargetYearRows++;
  }
  return item;
}

// Final user-facing implementation placed after all legacy definitions.
function createCurrentYearWorkingSheets(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'ระบบยังทำงานคำสั่งก่อนหน้าอยู่ กรุณารอสักครู่แล้วกดใหม่' };
  }

  try {
    var ss = SS();
    var preview = previewCurrentYearWorkingSheets(targetYear);
    if (!preview.success) return preview;

    var created = [];
    var skipped = [];
    var copiedRows = 0;
    var errors = [];

    (preview.sheets || []).forEach(function(item) {
      try {
        if (!item.sourceExists) {
          skipped.push(item.target + ' - ไม่พบชีตต้นทาง');
          return;
        }
        if (ss.getSheetByName(item.target)) {
          skipped.push(item.target + ' - มีอยู่แล้ว');
          return;
        }

        var source = ss.getSheetByName(item.source);
        if (!source) {
          skipped.push(item.target + ' - ไม่พบชีตต้นทาง');
          return;
        }

        var lastRow = source.getLastRow();
        var lastCol = source.getLastColumn();
        if (lastRow < 1 || lastCol < 1) {
          skipped.push(item.target + ' - ชีตต้นทางว่าง');
          return;
        }

        var headers = source.getRange(1, 1, 1, lastCol).getValues()[0];
        var rowsToCopy = [];
        if (item.explicitTargetYearRows > 0 && lastRow > 1) {
          var data = source.getRange(2, 1, lastRow - 1, lastCol).getValues();
          for (var r = 0; r < data.length; r++) {
            if (AY_prepareShouldCopyRow_(item.baseName, data[r], headers, preview.targetYear)) {
              rowsToCopy.push(AY_prepareNormalizeRowForTarget_(item.baseName, data[r], headers, preview.targetYear));
            }
          }
        }

        var target = ss.insertSheet(item.target);
        target.getRange(1, 1, 1, lastCol).setValues([headers]);
        target.setFrozenRows(1);
        if (rowsToCopy.length > 0) {
          target.getRange(2, 1, rowsToCopy.length, lastCol).setValues(rowsToCopy);
        }
        try {
          target.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#e8f0fe');
        } catch (_) {}

        created.push(item.target);
        copiedRows += rowsToCopy.length;
      } catch (err) {
        errors.push(item.target + ': ' + err.message);
      }
    });

    var message = errors.length
      ? 'เตรียมชีตปีปัจจุบันบางส่วนสำเร็จ แต่มีข้อผิดพลาด ' + errors.length + ' รายการ'
      : (created.length === 0
        ? 'ชีตปีปัจจุบัน ' + preview.targetYear + ' มีอยู่แล้ว ไม่ต้องสร้างเพิ่ม'
        : 'เตรียมชีตปีปัจจุบัน ' + preview.targetYear + ' สำเร็จ ' + created.length + ' ชีต');

    return {
      success: errors.length === 0,
      targetYear: preview.targetYear,
      createdCount: created.length,
      skippedCount: skipped.length,
      copiedRows: copiedRows,
      errorCount: errors.length,
      created: created.slice(0, 80),
      skipped: skipped.slice(0, 80),
      errors: errors.slice(0, 40),
      message: message
    };
  } catch (e) {
    Logger.log('createCurrentYearWorkingSheets final user-facing error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}



// ============================================================
// EXTERNAL ACADEMIC YEAR ARCHIVE
// Moves old year-suffixed sheets to a separate spreadsheet so the
// working file stays small. Destructive cleanup is blocked until the
// external archive is verified.
// ============================================================

var AY_EXTERNAL_ARCHIVE_REGISTRY = 'AcademicYearArchiveRegistry';
var AY_EXTERNAL_ARCHIVE_INFO = '_ARCHIVE_INFO';

function AY_externalArchiveRegistrySheet_() {
  var ss = SS();
  var sheet = ss.getSheetByName(AY_EXTERNAL_ARCHIVE_REGISTRY);
  if (!sheet) {
    sheet = ss.insertSheet(AY_EXTERNAL_ARCHIVE_REGISTRY);
    sheet.getRange(1, 1, 1, 11).setValues([[
      'academic_year',
      'spreadsheet_id',
      'url',
      'name',
      'created_at',
      'archived_sheet_count',
      'cleaned_at',
      'notes',
      'status',
      'verified_at',
      'last_restored_at'
    ]]);
    try { sheet.hideSheet(); } catch (_) {}
  } else {
    // อัปเกรดคอลัมน์เก่าหากจำเป็น (ต้องมีอย่างน้อย 11 คอลัมน์)
    var maxCols = sheet.getMaxColumns();
    if (maxCols < 11) {
      try { sheet.insertColumnsAfter(maxCols, 11 - maxCols); } catch (_) {}
    }
    // อัปเดต headers เสมอเพื่อให้แน่ใจว่าคอลัมน์ตรงกัน
    sheet.getRange(1, 1, 1, 11).setValues([[
      'academic_year',
      'spreadsheet_id',
      'url',
      'name',
      'created_at',
      'archived_sheet_count',
      'cleaned_at',
      'notes',
      'status',
      'verified_at',
      'last_restored_at'
    ]]);
  }
  return sheet;
}

function AY_externalArchiveFindRecord_(year) {
  var sheet = AY_externalArchiveRegistrySheet_();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(year)) {
      return {
        row: i + 1,
        year: String(values[i][0] || ''),
        spreadsheetId: String(values[i][1] || ''),
        url: String(values[i][2] || ''),
        name: String(values[i][3] || ''),
        createdAt: values[i][4] || '',
        archivedSheetCount: values[i][5] || 0,
        cleanedAt: values[i][6] || '',
        notes: String(values[i][7] || ''),
        status: String(values[i][8] || ''),
        verifiedAt: values[i][9] || '',
        lastRestoredAt: values[i][10] || ''
      };
    }
  }
  return null;
}

function AY_externalArchiveSaveRecord_(year, archive, sheetCount, notes) {
  var sheet = AY_externalArchiveRegistrySheet_();
  var record = AY_externalArchiveFindRecord_(year);
  var row = record ? record.row : sheet.getLastRow() + 1;
  var existingCleanedAt = record ? record.cleanedAt : '';
  var existingStatus = record ? record.status : 'archived'; // ค่าเริ่มต้น
  var existingVerifiedAt = record ? record.verifiedAt : '';
  var existingLastRestoredAt = record ? record.lastRestoredAt : '';
  
  sheet.getRange(row, 1, 1, 11).setValues([[
    String(year),
    archive.getId(),
    archive.getUrl(),
    archive.getName(),
    record ? record.createdAt || new Date() : new Date(),
    sheetCount || 0,
    existingCleanedAt || '',
    notes || (record ? record.notes || '' : ''),
    existingStatus || 'archived',
    existingVerifiedAt || '',
    existingLastRestoredAt || ''
  ]]);
  try { sheet.hideSheet(); } catch (_) {}
}

function AY_registrySetStatus_(year, status, extraFields) {
  var sheet = AY_externalArchiveRegistrySheet_();
  var record = AY_externalArchiveFindRecord_(year);
  if (!record) return;
  
  sheet.getRange(record.row, 9).setValue(status);
  
  if (extraFields) {
    if (extraFields.verifiedAt !== undefined) {
      sheet.getRange(record.row, 10).setValue(extraFields.verifiedAt);
    }
    if (extraFields.lastRestoredAt !== undefined) {
      sheet.getRange(record.row, 11).setValue(extraFields.lastRestoredAt);
    }
    if (extraFields.notes !== undefined) {
      sheet.getRange(record.row, 8).setValue(extraFields.notes);
    }
    if (extraFields.cleanedAt !== undefined) {
      sheet.getRange(record.row, 7).setValue(extraFields.cleanedAt);
    }
  }
}

function AY_registryGetStatus_(year) {
  var record = AY_externalArchiveFindRecord_(year);
  return record ? record.status || 'archived' : '';
}

function AY_getRegistryAll_() {
  var sheet = AY_externalArchiveRegistrySheet_();
  var values = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var y = String(values[i][0] || '').trim();
    if (!y) continue;
    list.push({
      year: y,
      spreadsheetId: String(values[i][1] || ''),
      url: String(values[i][2] || ''),
      name: String(values[i][3] || ''),
      createdAt: values[i][4] || '',
      archivedSheetCount: values[i][5] || 0,
      cleanedAt: values[i][6] || '',
      notes: String(values[i][7] || ''),
      status: String(values[i][8] || 'archived'),
      verifiedAt: values[i][9] || '',
      lastRestoredAt: values[i][10] || ''
    });
  }
  return list;
}

function AY_externalArchiveMarkCleaned_(year, deletedCount) {
  var sheet = AY_externalArchiveRegistrySheet_();
  var record = AY_externalArchiveFindRecord_(year);
  if (!record) return;
  sheet.getRange(record.row, 7, 1, 2).setValues([[new Date(), 'cleaned local sheets: ' + deletedCount]]);
  
  // อัปเดตสถานะเป็น verified และบันทึก verifiedAt
  AY_registrySetStatus_(year, 'verified', { verifiedAt: new Date() });
}

function AY_externalArchiveName_(year) {
  var sourceName = '';
  try { sourceName = SS().getName(); } catch (_) {}
  return 'PP5Smart Archive ' + year + (sourceName ? ' - ' + sourceName : '');
}

function AY_externalArchiveOpenOrCreate_(year) {
  var record = AY_externalArchiveFindRecord_(year);
  if (record && record.spreadsheetId) {
    try {
      return SpreadsheetApp.openById(record.spreadsheetId);
    } catch (e) {
      Logger.log('AY_externalArchiveOpenOrCreate_ old file unavailable: ' + e.message);
    }
  }

  var archive = SpreadsheetApp.create(AY_externalArchiveName_(year));
  try {
    var sourceFile = DriveApp.getFileById(SS().getId());
    var archiveFile = DriveApp.getFileById(archive.getId());
    var parents = sourceFile.getParents();
    if (parents.hasNext()) {
      var folder = parents.next();
      folder.addFile(archiveFile);
      try { DriveApp.getRootFolder().removeFile(archiveFile); } catch (_) {}
    }
  } catch (e) {
    Logger.log('AY_externalArchiveOpenOrCreate_ folder placement skipped: ' + e.message);
  }
  AY_externalArchiveSaveRecord_(year, archive, 0, 'created');
  return archive;
}

function AY_externalArchiveIsCandidate_(sheetName, year) {
  if (!sheetName) return false;
  if (sheetName === AY_EXTERNAL_ARCHIVE_REGISTRY) return false;
  return new RegExp('_' + String(year) + '$').test(sheetName);
}

function AY_externalArchiveCandidates_(year) {
  var ss = SS();
  return ss.getSheets()
    .filter(function(sheet) { return AY_externalArchiveIsCandidate_(sheet.getName(), year); })
    .map(function(sheet) {
      return {
        name: sheet.getName(),
        rows: sheet.getLastRow(),
        dataRows: Math.max(0, sheet.getLastRow() - 1),
        columns: sheet.getLastColumn()
      };
    })
    .sort(function(a, b) { return a.name.localeCompare(b.name, 'th'); });
}

function AY_externalArchiveGetExisting_(record) {
  if (!record || !record.spreadsheetId) return { available: false, sheets: {} };
  try {
    var archive = SpreadsheetApp.openById(record.spreadsheetId);
    var sheets = {};
    archive.getSheets().forEach(function(sheet) {
      sheets[sheet.getName()] = {
        rows: sheet.getLastRow(),
        columns: sheet.getLastColumn(),
        gridRows: sheet.getMaxRows(),
        gridColumns: sheet.getMaxColumns()
      };
    });
    return {
      available: true,
      spreadsheetId: archive.getId(),
      url: archive.getUrl(),
      name: archive.getName(),
      sheets: sheets
    };
  } catch (e) {
    return { available: false, error: e.message, sheets: {} };
  }
}

function AY_externalArchiveVerification_(year) {
  var candidates = AY_externalArchiveCandidates_(year);
  var record = AY_externalArchiveFindRecord_(year);
  var existing = AY_externalArchiveGetExisting_(record);
  var missing = [];
  var mismatched = [];
  var verified = 0;

  candidates.forEach(function(item) {
    var archived = existing.sheets[item.name];
    if (!archived) {
      missing.push(item.name);
      return;
    }
    var archiveRows = Math.max(archived.rows || 0, archived.gridRows || 0);
    var archiveColumns = Math.max(archived.columns || 0, archived.gridColumns || 0);
    if (archiveRows < item.rows || archiveColumns < item.columns) {
      mismatched.push(item.name + ' (ต้นทาง ' + item.rows + 'x' + item.columns + ', archive ' + (archived.rows || 0) + 'x' + (archived.columns || 0) + ', grid ' + archiveRows + 'x' + archiveColumns + ')');
      return;
    }
    verified++;
  });

  return {
    record: record,
    archiveAvailable: existing.available,
    archive: existing,
    candidates: candidates,
    verifiedCount: verified,
    missing: missing,
    mismatched: mismatched,
    readyToCleanup: candidates.length > 0 && existing.available && missing.length === 0 && mismatched.length === 0
  };
}

function AY_externalArchiveWriteInfo_(archive, year, sourceId, sheetCount) {
  var info = archive.getSheetByName(AY_EXTERNAL_ARCHIVE_INFO);
  if (!info) info = archive.insertSheet(AY_EXTERNAL_ARCHIVE_INFO);
  info.clear();
  info.getRange(1, 1, 7, 2).setValues([
    ['archive_year', String(year)],
    ['source_spreadsheet_id', sourceId],
    ['created_at', new Date()],
    ['sheet_count', sheetCount],
    ['mode', 'data-only read-only archive'],
    ['warning', 'This archive is for old academic year reference. Edit the current year in the main PP5Smart file.'],
    ['generated_by', 'PP5Smart']
  ]);
  try { info.hideSheet(); } catch (_) {}
}

function AY_errorMessage_(err) {
  if (!err) return 'ไม่ทราบสาเหตุ';
  return String(err.message || err.stack || err);
}

function AY_externalArchiveReplaceSheet_(source, targetSpreadsheet) {
  var name = source.getName();
  var target = targetSpreadsheet.getSheetByName(name);
  var existed = !!target;
  var tempName = '__AY_TMP_' + String(new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
  var copied = null;

  try {
    copied = source.copyTo(targetSpreadsheet);
    copied.setName(tempName);
    SpreadsheetApp.flush();

    if (target && targetSpreadsheet.getSheets().length > 1) {
      targetSpreadsheet.deleteSheet(target);
      SpreadsheetApp.flush();
    } else if (target) {
      target.clear();
      target.setName(tempName + '_old');
      SpreadsheetApp.flush();
    }

    copied.setName(name);
    target = copied;
  } catch (copyErr) {
    try {
      if (copied && targetSpreadsheet.getSheets().length > 1) targetSpreadsheet.deleteSheet(copied);
    } catch (_) {}

    target = targetSpreadsheet.getSheetByName(name);
    if (!target) target = targetSpreadsheet.insertSheet(name);
    AY_externalArchiveWriteSheetValues_(source, target);
    SpreadsheetApp.flush();

    var sourceRows = source.getLastRow();
    var sourceCols = source.getLastColumn();
    var targetRows = Math.max(target.getLastRow(), target.getMaxRows());
    var targetCols = Math.max(target.getLastColumn(), target.getMaxColumns());
    if (targetRows < sourceRows || targetCols < sourceCols) {
      throw new Error('copyTo ล้มเหลว: ' + AY_errorMessage_(copyErr) + ' และ fallback value-only ยังไม่ครบ (ต้นทาง ' + sourceRows + 'x' + sourceCols + ', archive ' + targetRows + 'x' + targetCols + ')');
    }
  }

  try {
    target.setFrozenRows(source.getFrozenRows());
    target.setFrozenColumns(source.getFrozenColumns());
  } catch (_) {}
  return { copied: !existed, updated: existed, reason: existed ? 'อัปเดตแล้ว' : '' };
}



function AY_externalArchiveRepairProblems_(year, verification, forceAll) {
  var ss = SS();
  var archive = AY_externalArchiveOpenOrCreate_(year);
  var existing = AY_externalArchiveGetExisting_({ spreadsheetId: archive.getId() });
  var refreshed = [];
  var skipped = [];
  var errors = [];

  (verification && verification.candidates ? verification.candidates : AY_externalArchiveCandidates_(year)).forEach(function(item) {
    var archived = existing.sheets[item.name];
    var archiveRows = archived ? Math.max(archived.rows || 0, archived.gridRows || 0) : 0;
    var archiveColumns = archived ? Math.max(archived.columns || 0, archived.gridColumns || 0) : 0;
    var needsRefresh = forceAll || !archived || archiveRows < item.rows || archiveColumns < item.columns;
    if (!needsRefresh) {
      skipped.push(item.name + ' - ตรงกับไฟล์หลักแล้ว');
      return;
    }

    try {
      var source = ss.getSheetByName(item.name);
      if (!source) {
        errors.push(item.name + ': ไม่พบชีตต้นทาง');
        return;
      }
      AY_externalArchiveReplaceSheet_(source, archive);
      refreshed.push(item.name);
    } catch (err) {
      errors.push(item.name + ': ' + AY_errorMessage_(err));
    }
  });

  SpreadsheetApp.flush();
  Utilities.sleep(1000);
  AY_externalArchiveWriteInfo_(archive, year, ss.getId(), archive.getSheets().length);
  AY_externalArchiveSaveRecord_(year, archive, archive.getSheets().length, 'auto-repaired');

  return {
    archive: archive,
    refreshed: refreshed,
    skipped: skipped,
    errors: errors
  };
}

function previewAcademicYearExternalArchive(targetYear) {
  try {
    var year = AY_archiveTargetYear_(targetYear);
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    var current = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    var verification = AY_externalArchiveVerification_(year);
    return {
      success: true,
      readOnly: true,
      targetYear: year,
      currentYear: current,
      isCurrentYear: current && String(current) === String(year),
      registry: verification.record,
      archiveAvailable: verification.archiveAvailable,
      archive: verification.archive,
      candidates: verification.candidates,
      candidateCount: verification.candidates.length,
      verifiedCount: verification.verifiedCount,
      missing: verification.missing,
      mismatched: verification.mismatched,
      readyToCleanup: verification.readyToCleanup,
      message: verification.candidates.length
        ? 'พบชีตปี ' + year + ' ในไฟล์หลัก ' + verification.candidates.length + ' ชีต'
        : 'ยังไม่พบชีต _' + year + ' ในไฟล์หลัก ให้สร้าง Snapshot ปีเก่าก่อน'
    };
  } catch (e) {
    Logger.log('previewAcademicYearExternalArchive error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function createAcademicYearExternalArchive(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'ระบบยังทำงานคำสั่งก่อนหน้าอยู่ กรุณารอสักครู่แล้วกดใหม่' };
  }

  try {
    var year = AY_archiveTargetYear_(targetYear);
    var current = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    var registryStatus = AY_registryGetStatus_(year);
    if (current && String(year) === String(current)) {
      if (registryStatus !== 'editing') {
        return { success: false, message: 'ไม่ควรส่งออกปีปัจจุบัน ' + year + ' ไปคลังเก่า ให้ใช้กับปีเก่าหรือปีที่อยู่ในโหมดแก้ไขย้อนหลังเท่านั้น' };
      }
    }

    var candidates = AY_externalArchiveCandidates_(year);
    if (!candidates.length) {
      return { success: false, message: 'ยังไม่พบชีต _' + year + ' ในไฟล์หลัก ให้สร้าง Snapshot ปีเก่าก่อน' };
    }

    var verification = AY_externalArchiveVerification_(year);
    var repair = AY_externalArchiveRepairProblems_(year, verification, true);
    verification = AY_externalArchiveVerification_(year);
    return {
      success: repair.errors.length === 0 && verification.readyToCleanup,
      targetYear: year,
      archiveId: repair.archive.getId(),
      archiveUrl: repair.archive.getUrl(),
      archiveName: repair.archive.getName(),
      createdCount: repair.refreshed.length,
      skippedCount: repair.skipped.length,
      errorCount: repair.errors.length,
      verifiedCount: verification.verifiedCount,
      readyToCleanup: verification.readyToCleanup,
      missing: verification.missing,
      mismatched: verification.mismatched,
      created: repair.refreshed.slice(0, 80),
      skipped: repair.skipped.slice(0, 80),
      errors: repair.errors.slice(0, 40),
      message: verification.readyToCleanup
        ? 'ส่งออกปี ' + year + ' ไปไฟล์คลังสำเร็จ ตรวจครบแล้ว พร้อมล้างชีตปีเก่าจากไฟล์หลัก'
        : 'ส่งออกไฟล์คลังแล้ว แต่ยังตรวจไม่ครบ กรุณาดูรายการที่ขาด'
    };
  } catch (e) {
    Logger.log('createAcademicYearExternalArchive error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function cleanupArchivedYearSheetsFromMain(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'ระบบยังทำงานคำสั่งก่อนหน้าอยู่ กรุณารอสักครู่แล้วกดใหม่' };
  }

  try {
    var year = AY_archiveTargetYear_(targetYear);
    var current = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    if (current && String(year) === String(current)) {
      return { success: false, message: 'ไม่อนุญาตให้ล้างชีตปีปัจจุบัน ' + year };
    }

    var verification = AY_externalArchiveVerification_(year);
    if (!verification.readyToCleanup) {
      var repair = AY_externalArchiveRepairProblems_(year, verification, false);
      if (repair.errors.length) {
        return {
          success: false,
          message: 'ซ่อมไฟล์คลังก่อนล้างไม่สำเร็จ กรุณาดูข้อผิดพลาด',
          missing: verification.missing,
          mismatched: verification.mismatched,
          errors: repair.errors
        };
      }
      verification = AY_externalArchiveVerification_(year);
    }

    if (!verification.readyToCleanup) {
      return {
        success: false,
        message: 'ยังล้างไม่ได้ เพราะไฟล์คลังยังไม่ครบหรือยังตรวจไม่ผ่าน แม้ระบบซ่อมให้อัตโนมัติแล้ว',
        missing: verification.missing,
        mismatched: verification.mismatched
      };
    }

    var ss = SS();
    var deleted = [];
    verification.candidates.forEach(function(item) {
      var sheet = ss.getSheetByName(item.name);
      if (sheet && ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
        deleted.push(item.name);
      }
    });

    AY_externalArchiveMarkCleaned_(year, deleted.length);
    return {
      success: true,
      targetYear: year,
      deletedCount: deleted.length,
      deleted: deleted.slice(0, 120),
      archiveUrl: verification.archive.url || '',
      message: 'ล้างชีตปีเก่า ' + year + ' จากไฟล์หลักสำเร็จ ' + deleted.length + ' ชีต'
    };
  } catch (e) {
    Logger.log('cleanupArchivedYearSheetsFromMain error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// ============================================================
// OLD YEAR EDIT WORKFLOW
// Restore old year sheets from archive, switch the active academic year
// temporarily, then export back to archive and clean up again.
// ============================================================

function AY_externalArchiveWriteSheetValues_(source, target) {
  var lastRow = source.getLastRow();
  var lastCol = source.getLastColumn();
  target.clear();
  var maxRows = target.getMaxRows();
  var maxCols = target.getMaxColumns();
  if (lastRow > maxRows) target.insertRowsAfter(maxRows, lastRow - maxRows);
  if (lastCol > maxCols) target.insertColumnsAfter(maxCols, lastCol - maxCols);
  if (lastRow > 0 && lastCol > 0) {
    var chunkSize = 500;
    for (var start = 1; start <= lastRow; start += chunkSize) {
      var numRows = Math.min(chunkSize, lastRow - start + 1);
      var values = source.getRange(start, 1, numRows, lastCol).getValues();
      target.getRange(start, 1, numRows, lastCol).setValues(values);
    }
    try {
      target.setFrozenRows(source.getFrozenRows());
      target.setFrozenColumns(source.getFrozenColumns());
      target.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#e8f0fe');
    } catch (_) {}
  }
}

function AY_externalArchiveCopyValues_(source, targetSpreadsheet) {
  return AY_externalArchiveReplaceSheet_(source, targetSpreadsheet);
}

function restoreAcademicYearFromExternalArchive(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'ระบบยังทำงานคำสั่งก่อนหน้าอยู่ กรุณารอสักครู่แล้วกดใหม่' };
  }

  try {
    var year = AY_archiveTargetYear_(targetYear);
    var current = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    if (current && String(year) === String(current)) {
      return { success: false, message: 'ปี ' + year + ' เป็นปีปัจจุบันอยู่แล้ว ไม่ต้องกู้คืนจากคลัง' };
    }

    var record = AY_externalArchiveFindRecord_(year);
    if (!record || !record.spreadsheetId) {
      return { success: false, message: 'ยังไม่พบไฟล์คลังของปี ' + year + ' กรุณาตรวจไฟล์คลังก่อน' };
    }

    var archive = SpreadsheetApp.openById(record.spreadsheetId);
    var ss = SS();
    var restored = [];
    var skipped = [];
    var errors = [];

    archive.getSheets().forEach(function(source) {
      var name = source.getName();
      if (name === AY_EXTERNAL_ARCHIVE_INFO) return;
      if (!AY_externalArchiveIsCandidate_(name, year)) return;

      try {
        if (ss.getSheetByName(name)) {
          skipped.push(name + ' - มีอยู่แล้วในไฟล์หลัก');
          return;
        }
        var target = ss.insertSheet(name);
        AY_externalArchiveWriteSheetValues_(source, target);
        restored.push(name);
      } catch (err) {
        errors.push(name + ': ' + err.message);
      }
    });

    if (errors.length === 0 && restored.length > 0) {
      // อัปเดตสถานะในคลังเป็น 'editing'
      AY_registrySetStatus_(year, 'editing', { lastRestoredAt: new Date() });
    }

    return {
      success: errors.length === 0,
      targetYear: year,
      archiveUrl: archive.getUrl(),
      restoredCount: restored.length,
      skippedCount: skipped.length,
      errorCount: errors.length,
      restored: restored.slice(0, 120),
      skipped: skipped.slice(0, 120),
      errors: errors.slice(0, 40),
      message: errors.length
        ? 'กู้คืนปี ' + year + ' บางส่วนสำเร็จ แต่มีข้อผิดพลาด ' + errors.length + ' รายการ'
        : 'กู้คืนชีตปี ' + year + ' จากไฟล์คลังกลับเข้าไฟล์หลักสำเร็จ ' + restored.length + ' ชีต'
    };
  } catch (e) {
    Logger.log('restoreAcademicYearFromExternalArchive error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function switchAcademicYearOnly(targetYear) {
  try {
    var year = AY_archiveTargetYear_(targetYear);
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    var before = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    var settings = (typeof S_getGlobalSettings === 'function') ? S_getGlobalSettings(false) : {};
    settings['ปีการศึกษา'] = String(year);
    settings.academicYear = String(year);
    settings.academic_year = String(year);
    settings.currentAcademicYear = String(year);

    if (typeof S_saveGlobalSettings === 'function') {
      S_saveGlobalSettings(settings);
    } else if (typeof saveGlobalSettings === 'function') {
      saveGlobalSettings(settings);
    } else {
      throw new Error('ไม่พบฟังก์ชันบันทึกปีการศึกษา');
    }
    try { if (typeof S_clearSettingsCache === 'function') S_clearSettingsCache(); } catch (_) {}

    return {
      success: true,
      beforeYear: before,
      currentYear: String(year),
      message: 'สลับปีการศึกษาเป็น ' + year + ' แล้ว'
    };
  } catch (e) {
    Logger.log('switchAcademicYearOnly error: ' + e.message + '\n' + (e.stack || ''));
    return { success: false, message: e.message };
  }
}

// 📦 ฟังก์ชันยกเลิกการแก้ไขปีเก่า
function AY_cancelEditingOldYear(targetYear) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: 'ระบบยังทำงานคำสั่งก่อนหน้าอยู่ กรุณารอสักครู่แล้วกดใหม่' };
  }
  try {
    var year = AY_archiveTargetYear_(targetYear);
    if (!year) throw new Error('ไม่พบปีการศึกษา');
    var registryStatus = AY_registryGetStatus_(year);
    if (registryStatus !== 'editing') {
      throw new Error('ปี ' + year + ' ไม่ได้อยู่ในโหมดแก้ไขย้อนหลัง จึงไม่ต้องยกเลิกการแก้ไข');
    }

    var computedYear = (typeof AY_computeAcademicYearFromDate === 'function')
      ? AY_computeAcademicYearFromDate(new Date())
      : '';
    if (!computedYear) throw new Error('ไม่สามารถคำนวณปีการศึกษาปัจจุบันได้');
    if (String(year) === String(computedYear)) {
      throw new Error('ปี ' + year + ' เป็นปีปัจจุบันตามปฏิทิน จึงไม่ใช่ปีเก่าย้อนหลังที่ต้องยกเลิก');
    }

    var switchResult = switchAcademicYearOnly(computedYear);
    if (!switchResult || !switchResult.success) {
      throw new Error('สลับกลับปีปัจจุบัน ' + computedYear + ' ไม่สำเร็จ: ' + ((switchResult && switchResult.message) || 'ไม่ทราบสาเหตุ'));
    }
    
    var ss = SS();
    var candidates = AY_externalArchiveCandidates_(year);
    var deleted = [];
    candidates.forEach(function(item) {
      var sheet = ss.getSheetByName(item.name);
      if (sheet && ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
        deleted.push(item.name);
      }
    });
    
    AY_registrySetStatus_(year, 'verified', { notes: 'ยกเลิกแก้ไข คืนสถานะสมบูรณ์' });
    return {
      success: true,
      targetYear: year,
      restoredToYear: computedYear,
      deletedCount: deleted.length,
      deleted: deleted,
      message: 'ยกเลิกการดึงข้อมูลปี ' + year + ' มาแก้ไข สลับกลับปีปัจจุบัน ' + computedYear + ' และนำชีตชั่วคราวออกจากไฟล์หลักแล้ว ' + deleted.length + ' ชีต'
    };
  } catch (e) {
    Logger.log('AY_cancelEditingOldYear error: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// 📊 ดึงคลังข้อมูลสำหรับ UI
function getYearDatabaseRegistry() {
  try {
    var currentYear = (typeof AY_getCurrentAcademicYear === 'function') ? AY_getCurrentAcademicYear(false) : '';
    
    // คำนวณปีจริงตามปฏิทิน (เช่น พ.ค. เป็นต้นไป คือปีการศึกษาใหม่)
    var computedYear = AY_computeAcademicYearFromDate(new Date());
    var isEditingOldYear = currentYear && String(currentYear) !== String(computedYear);
    
    var years = AY_getRegistryAll_();
    
    // นำเข้าข้อมูลปีปัจจุบัน (active) เข้าไปในรายการเพื่อแสดงผลใน UI ด้วย
    var hasCurrentYear = false;
    for (var i = 0; i < years.length; i++) {
      if (String(years[i].year) === String(currentYear)) {
        years[i].status = 'active'; // บังคับแสดงผลเป็น active หากกำลังใช้อยู่
        hasCurrentYear = true;
      }
    }
    
    // แปลงรูปแบบวันที่ใน Registry ให้สวยงามเหมาะสมสำหรับ UI
    var formattedYears = years.map(function(y) {
      return {
        year: y.year,
        spreadsheetId: y.spreadsheetId,
        url: y.url,
        name: y.name,
        createdAt: y.createdAt ? Utilities.formatDate(new Date(y.createdAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        archivedSheetCount: y.archivedSheetCount,
        cleanedAt: y.cleanedAt ? Utilities.formatDate(new Date(y.cleanedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        notes: y.notes,
        status: String(currentYear) === String(y.year) ? 'active' : (y.status || 'archived'),
        verifiedAt: y.verifiedAt ? Utilities.formatDate(new Date(y.verifiedAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        lastRestoredAt: y.lastRestoredAt ? Utilities.formatDate(new Date(y.lastRestoredAt), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : ''
      };
    });
    
    // ถ้ายังไม่มีปีปัจจุบันใน registry เลย ให้ใส่รายการจำลองสำหรับแสดง active
    if (!hasCurrentYear && currentYear) {
      formattedYears.unshift({
        year: currentYear,
        spreadsheetId: SS().getId(),
        url: SS().getUrl(),
        name: SS().getName(),
        createdAt: '',
        archivedSheetCount: 0,
        cleanedAt: '',
        notes: 'ปีการศึกษาที่ใช้งานในปัจจุบัน',
        status: 'active',
        verifiedAt: '',
        lastRestoredAt: ''
      });
    }
    
    // เรียงลำดับปีจากมากไปน้อย
    formattedYears.sort(function(a, b) {
      return Number(b.year) - Number(a.year);
    });
    
    return {
      success: true,
      currentYear: currentYear,
      computedYear: computedYear,
      isEditingOldYear: isEditingOldYear,
      years: formattedYears
    };
  } catch (e) {
    Logger.log('getYearDatabaseRegistry error: ' + e.message);
    return { success: false, message: e.message };
  }
}
