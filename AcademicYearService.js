// ============================================================
// ACADEMIC YEAR SERVICE
// Safe helpers for current-year filtering and future archives.
// These helpers are intentionally soft: if an old sheet has no
// academic_year column, existing workflows continue to work.
// ============================================================

var AY_KEY_THAI = '\u0e1b\u0e35\u0e01\u0e32\u0e23\u0e28\u0e36\u0e01\u0e29\u0e32'; // ปีการศึกษา

function AY_normalizeYear_(value) {
  var s = String(value || '').trim();
  var m = s.match(/\d{4}/);
  return m ? m[0] : '';
}

function AY_computeAcademicYearFromDate(date) {
  var d = date || new Date();
  var year = d.getFullYear() + 543;
  var month = d.getMonth() + 1;
  return String(month >= 5 ? year : year - 1);
}

function AY_getCurrentAcademicYear(useCache) {
  try {
    var settings = {};
    if (typeof S_getGlobalSettings === 'function') {
      settings = S_getGlobalSettings(useCache !== false) || {};
    } else if (typeof getGlobalSettings === 'function') {
      settings = getGlobalSettings(useCache !== false) || {};
    }

    var year = AY_normalizeYear_(
      settings[AY_KEY_THAI] ||
      settings.academicYear ||
      settings.academic_year ||
      settings.currentAcademicYear
    );
    if (year) return year;
  } catch (e) {
    Logger.log('AY_getCurrentAcademicYear fallback: ' + e.message);
  }

  if (typeof S_getCurrentAcademicYear_ === 'function') {
    return String(S_getCurrentAcademicYear_());
  }
  if (typeof U_getCurrentAcademicYear === 'function') {
    return String(U_getCurrentAcademicYear());
  }
  return AY_computeAcademicYearFromDate(new Date());
}

function AY_headerMap_(headers) {
  var map = {};
  (headers || []).forEach(function(h, i) {
    var key = String(h || '').trim().toLowerCase();
    if (key) map[key] = i;
  });
  return map;
}

function AY_findColumn_(headersOrMap, names) {
  var map = Array.isArray(headersOrMap) ? AY_headerMap_(headersOrMap) : (headersOrMap || {});
  for (var i = 0; i < names.length; i++) {
    var key = String(names[i] || '').trim().toLowerCase();
    if (map[key] != null) return map[key];
  }
  return -1;
}

function AY_getRowAcademicYear(row, headersOrMap) {
  var idx = AY_findColumn_(headersOrMap, [
    'academic_year',
    'academicyear',
    'academic year',
    'year',
    AY_KEY_THAI
  ]);
  if (idx < 0) return '';
  return AY_normalizeYear_(row[idx]);
}

function AY_rowMatchesAcademicYear(row, headersOrMap, targetYear, options) {
  options = options || {};
  var idx = AY_findColumn_(headersOrMap, [
    'academic_year',
    'academicyear',
    'academic year',
    'year',
    AY_KEY_THAI
  ]);

  // Old sheets without year columns are allowed until migration is complete.
  if (idx < 0) return options.allowMissingColumn !== false;

  var rowYear = AY_normalizeYear_(row[idx]);
  if (!rowYear) return options.allowBlank !== false;

  var target = AY_normalizeYear_(targetYear || AY_getCurrentAcademicYear(false));
  return !target || rowYear === target;
}

function AY_getAcademicYearInfo() {
  var currentYear = AY_getCurrentAcademicYear(false);
  return {
    currentYear: currentYear,
    computedYear: AY_computeAcademicYearFromDate(new Date()),
    availableYears: (typeof S_getAvailableYears === 'function') ? S_getAvailableYears() : [currentYear]
  };
}

function AY_getStudentsSheetForRead(targetYear) {
  var ss = SS();
  var year = AY_normalizeYear_(targetYear || AY_getCurrentAcademicYear(false));
  if (year) {
    var yearly = ss.getSheetByName('Students_' + year);
    if (yearly && yearly.getLastRow() > 1) return yearly;
  }
  return ss.getSheetByName('Students');
}
