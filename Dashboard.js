// ============================================================
// 🚀 DASHBOARD & CACHING WITH ATTENDANCE DATA
// ============================================================

const DashboardCacheManager = {
  cache: new Map(),
  get(key) {
    const item = this.cache.get(key);
    if (item && item.expiry > Date.now()) return item.data;
    if (item) this.cache.delete(key);
    return null;
  },
  set(key, data, ttl = 120000) {
    this.cache.set(key, { data: data, expiry: Date.now() + ttl, created: new Date().toISOString() });
  },
  clearAll() { this.cache.clear(); }
};

const AcademicCacheManager = {
  cache: new Map(),
  get(key) {
    const item = this.cache.get(key);
    if (item && item.expiry > Date.now()) return item.data;
    if (item) this.cache.delete(key);
    return null;
  },
  set(key, data, ttl = 300000) {
    this.cache.set(key, { data: data, expiry: Date.now() + ttl, created: new Date().toISOString() });
  },
  clearAll() { this.cache.clear(); }
};

// === Inline implementations (ไม่พึ่ง utils.js ที่ minified) ===
function getCurrentThaiMonth() {
  const thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const now = new Date();
  return thaiMonths[now.getMonth()] + String(now.getFullYear() + 543);
}
function getCurrentAcademicYear() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  return month >= 5 ? year + 543 : year + 542;
}

// ============================================================
// 📊 DASHBOARD SUMMARY
// ============================================================

function getCachedDashboardSummary() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_summary_v5';
    const cachedData = cache.get(cacheKey);
    if (cachedData != null) return JSON.parse(cachedData);

    const startTime = Date.now();
    const summary = getDashboardSummary();
    summary.processingTime = Date.now() - startTime;
    summary.lastUpdated = new Date().toISOString();
    try { cache.put(cacheKey, JSON.stringify(summary), 300); } catch (ce) { Logger.log('Cache put failed: ' + ce.message); }
    return summary;
  } catch (e) {
    Logger.log(`❌ getCachedDashboardSummary: ${e.message}`);
    return { error: true, message: `ไม่สามารถโหลดข้อมูลสรุปได้: ${e.message}` };
  }
}

function getDashboardSummary() {
  const ss = SS();
  const sheet = ss.getSheetByName("Students");
  if (!sheet) throw new Error("ไม่พบชีต Students");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const genderCol = headers.indexOf("gender");
  const gradeCol = headers.indexOf("grade");
  const studentIdCol = headers.indexOf("student_id");

  if (genderCol === -1 || gradeCol === -1 || studentIdCol === -1) {
    throw new Error("ไม่พบคอลัมน์ที่จำเป็น (gender, grade, student_id)");
  }

  let male = 0, female = 0;
  const classMap = {};
  const studentGradeMap = {};

  rows.forEach(row => {
    const gender = String(row[genderCol] || "").trim();
    const grade = String(row[gradeCol] || "").trim();
    const studentId = String(row[studentIdCol] || '').trim();
    if (!grade || !studentId) return;

    studentGradeMap[studentId] = grade;
    if (!classMap[grade]) classMap[grade] = { male: 0, female: 0 };

    if (gender === "ชาย" || gender === "ช") { male++; classMap[grade].male++; }
    else if (gender === "หญิง" || gender === "ญ") { female++; classMap[grade].female++; }
  });

  const attendanceData = getAttendanceData(ss, studentGradeMap);

  return { total: male + female, male, female, classes: classMap, attendance: attendanceData };
}

function getFastStudentStats() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'fast_student_stats_v1';
    const hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);

    const ss = SS();
    const sheet = ss.getSheetByName("Students");
    if (!sheet) return { total: 0, male: 0, female: 0 };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let genderCol = headers.indexOf("gender");
    if (genderCol === -1) genderCol = headers.indexOf("เพศ");

    let male = 0, female = 0;
    for (let i = 1; i < data.length; i++) {
      const gender = String(data[i][genderCol] || "").trim();
      if (["ชาย", "ช", "ด.ช.", "นาย"].includes(gender)) male++;
      else if (["หญิง", "ญ", "ด.ญ.", "นาง", "นางสาว"].includes(gender)) female++;
    }
    const result = { total: male + female, male, female, lastUpdated: new Date().toISOString() };
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch (e) {
    return { total: 0, male: 0, female: 0, error: e.message };
  }
}

// ============================================================
// 📅 ATTENDANCE DATA
// ============================================================

function getAttendanceData(ss, studentGradeMap) {
  // อ่านปีการศึกษาจาก settings อัตโนมัติ (ไม่ hardcode)
  const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  let academicYear;
  try {
    academicYear = getCurrentAcademicYear();
  } catch (e) {
    Logger.log('⚠️ ไม่สามารถอ่านปีการศึกษาได้: ' + e.message);
    return { monthly: {}, yearlyByGrade: {}, availableMonths: [], currentMonth: '', _debug: 'yearError:' + e.message };
  }
  const monthsInYear = getMonthsInAcademicYear(academicYear);
  const months = monthsInYear.map(m => monthNames[m.month - 1] + String(m.yearCE + 543));

  // [DEBUG] เก็บชื่อชีตที่มีจริงใน Spreadsheet
  const allSheetNames = ss.getSheets().map(s => s.getName());
  const _debug = {
    academicYear: academicYear,
    expectedSheets: months,
    allSheets: allSheetNames.filter(n => /[ก-๙]/.test(n)),
    studentGradeMapSize: Object.keys(studentGradeMap).length
  };

  const currentMonth = getCurrentThaiMonth();
  const monthlyData = {};
  const gradeStats = {};

  const _monthDebug = [];
  months.forEach(monthName => {
    try {
      const monthSheet = ss.getSheetByName(monthName);
      if (!monthSheet) { _monthDebug.push(monthName + ':NOT_FOUND'); return; }

      const allData = monthSheet.getDataRange().getValues();
      if (allData.length < 2) { _monthDebug.push(monthName + ':NO_DATA(rows=' + allData.length + ')'); return; }

      const headers = allData[0];
      const trimHeaders = headers.map(h => String(h || '').trim());
      const rows = allData.slice(1);
      const studentIdCol = trimHeaders.indexOf('รหัส');
      const presentCol = trimHeaders.indexOf('มา');
      const leaveCol = trimHeaders.indexOf('ลา');
      const absentCol = trimHeaders.indexOf('ขาด');
      if (studentIdCol === -1 || presentCol === -1) { _monthDebug.push(monthName + ':HEADER_MISS(id=' + studentIdCol + ',มา=' + presentCol + ',hdr0-5=' + JSON.stringify(trimHeaders.slice(0,6)) + ',hdr-4=' + JSON.stringify(trimHeaders.slice(-4)) + ')'); return; }

      let schoolDaysInMonth = 0;
      const firstDateCol = 3;
      const lastDateCol = presentCol - 1;
      for (let col = firstDateCol; col <= lastDateCol; col++) {
        for (let row = 1; row < allData.length; row++) {
          if (allData[row][col] && allData[row][col].toString().trim() !== '') {
            schoolDaysInMonth++;
            break;
          }
        }
      }

      const gradeData = {};
      rows.forEach(row => {
        const studentId = String(row[studentIdCol] || '').trim();
        const grade = studentGradeMap[studentId];
        if (!studentId || !grade) return;

        const present = parseInt(row[presentCol]) || 0;
        const leave = parseInt(row[leaveCol]) || 0;
        const absent = parseInt(row[absentCol]) || 0;

        if (!gradeData[grade]) {
          gradeData[grade] = {
            totalPresent: 0, totalLeave: 0, totalAbsent: 0, totalDays: 0, studentCount: 0,
            schoolDaysInMonth: schoolDaysInMonth,
            studentsWithLeave: new Set(), studentsWithAbsent: new Set()
          };
        }

        gradeData[grade].totalPresent += present;
        gradeData[grade].totalLeave += leave;
        gradeData[grade].totalAbsent += absent;
        gradeData[grade].totalDays += present + leave + absent;
        gradeData[grade].studentCount++;
        if (leave > 0) gradeData[grade].studentsWithLeave.add(studentId);
        if (absent > 0) gradeData[grade].studentsWithAbsent.add(studentId);

        if (!gradeStats[grade]) gradeStats[grade] = { totalPresent: 0, totalLeave: 0, totalAbsent: 0, totalDays: 0 };
        gradeStats[grade].totalPresent += present;
        gradeStats[grade].totalLeave += leave;
        gradeStats[grade].totalAbsent += absent;
        gradeStats[grade].totalDays += present + leave + absent;
      });

      Object.keys(gradeData).forEach(grade => {
        const d = gradeData[grade];
        d.attendanceRate = d.totalDays > 0 ? (d.totalPresent / d.totalDays * 100).toFixed(2) : 0;
        d.uniqueAbsentCount = d.studentsWithAbsent.size;
        d.uniqueLeaveCount = d.studentsWithLeave.size;
        delete d.studentsWithAbsent;
        delete d.studentsWithLeave;
      });

      monthlyData[monthName] = gradeData;
      const matchCount = Object.values(gradeData).reduce((s, g) => s + g.studentCount, 0);
      _monthDebug.push(monthName + ':OK(rows=' + rows.length + ',matched=' + matchCount + ',grades=' + Object.keys(gradeData).join('/') + ')');
    } catch (error) {
      _monthDebug.push(monthName + ':ERROR(' + error.message + ')');
      Logger.log(`❌ Error reading ${monthName}: ${error.message}`);
    }
  });

  Object.keys(gradeStats).forEach(grade => {
    const s = gradeStats[grade];
    s.attendanceRate = s.totalDays > 0 ? (s.totalPresent / s.totalDays * 100).toFixed(2) : 0;
  });

  _debug.monthResults = _monthDebug;
  return { monthly: monthlyData, yearlyByGrade: gradeStats, availableMonths: Object.keys(monthlyData), currentMonth, _debug: _debug };
}

function getAttendanceByGrade(grade) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `attendance_grade_${grade}_v2`;
  const hit = cache.get(cacheKey);
  if (hit != null) return JSON.parse(hit);

  const summary = getDashboardSummary();
  if (!summary.attendance) throw new Error('ไม่พบข้อมูลการเช็คชื่อ');

  const result = { grade, monthly: {}, yearly: { totalPresent: 0, totalLeave: 0, totalAbsent: 0, totalDays: 0, attendanceRate: 0 } };
  const monthly = summary.attendance.monthly || {};

  Object.keys(monthly).forEach(monthName => {
    const g = monthly[monthName][grade];
    if (!g) return;
    const rec = {
      totalPresent: g.totalPresent || 0, totalLeave: g.totalLeave || 0,
      totalAbsent: g.totalAbsent || 0, totalDays: g.totalDays || 0,
      studentCount: g.studentCount || 0,
      attendanceRate: g.totalDays > 0 ? Number((g.totalPresent / g.totalDays * 100).toFixed(2)) : 0
    };
    result.monthly[monthName] = rec;
    result.yearly.totalPresent += rec.totalPresent;
    result.yearly.totalLeave += rec.totalLeave;
    result.yearly.totalAbsent += rec.totalAbsent;
    result.yearly.totalDays += rec.totalDays;
  });

  result.yearly.attendanceRate = result.yearly.totalDays > 0 ? Number((result.yearly.totalPresent / result.yearly.totalDays * 100).toFixed(2)) : 0;
  cache.put(cacheKey, JSON.stringify(result), 300);
  return result;
}

function getAttendanceByMonth(monthName = null) {
  if (!monthName) monthName = getCurrentThaiMonth();
  const cacheKey = `attendance_month_${monthName}`;
  let data = DashboardCacheManager.get(cacheKey);
  if (!data) {
    const summary = getDashboardSummary();
    if (!summary.attendance || !summary.attendance.monthly) throw new Error("ไม่พบข้อมูลการเช็คชื่อ");
    data = { month: monthName, grades: summary.attendance.monthly[monthName] || {}, totalStats: calculateMonthTotalStats(summary.attendance.monthly[monthName]) };
    DashboardCacheManager.set(cacheKey, data, 120000);
  }
  return data;
}

function getCurrentMonthAttendance() { return getAttendanceByMonth(); }

function calculateMonthTotalStats(monthData) {
  if (!monthData) return null;
  let totalPresent = 0, totalLeave = 0, totalAbsent = 0, totalDays = 0;
  Object.values(monthData).forEach(g => { totalPresent += g.totalPresent; totalLeave += g.totalLeave; totalAbsent += g.totalAbsent; totalDays += g.totalDays; });
  return { totalPresent, totalLeave, totalAbsent, totalDays, attendanceRate: totalDays > 0 ? (totalPresent / totalDays * 100).toFixed(2) : 0 };
}

function compareAttendanceByGrades(monthName = null) {
  const summary = getDashboardSummary();
  if (!summary.attendance) throw new Error("ไม่พบข้อมูลการเช็คชื่อ");
  let compareData = [];

  if (monthName && summary.attendance.monthly[monthName]) {
    const md = summary.attendance.monthly[monthName];
    Object.keys(md).forEach(grade => {
      compareData.push({ grade, attendanceRate: parseFloat(md[grade].attendanceRate), totalPresent: md[grade].totalPresent, totalDays: md[grade].totalDays, studentCount: md[grade].studentCount, period: monthName });
    });
  } else {
    const yd = summary.attendance.yearlyByGrade;
    Object.keys(yd).forEach(grade => {
      compareData.push({ grade, attendanceRate: parseFloat(yd[grade].attendanceRate), totalPresent: yd[grade].totalPresent, totalDays: yd[grade].totalDays, period: 'รายปี' });
    });
  }
  compareData.sort((a, b) => b.attendanceRate - a.attendanceRate);
  return compareData;
}

function getLowestAttendanceGrades(limit = 3) { return compareAttendanceByGrades().slice(-limit).reverse(); }
function getHighestAttendanceGrades(limit = 3) { return compareAttendanceByGrades().slice(0, limit); }
// 📊 ATTENDANCE CHARTS
// ============================================================

function getAttendanceTrendChartData(grade) {
  const data = getAttendanceByGrade(grade);
  if (!data || !data.monthly) throw new Error(`ไม่พบข้อมูลสำหรับห้อง ${grade}`);

  // สร้าง monthOrder แบบ dynamic จากปีการศึกษา
  const mn = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  let ay;
  try { ay = getCurrentAcademicYear(); } catch (e) { ay = 2568; }
  const miy = getMonthsInAcademicYear(ay);
  const monthOrder = miy.map(m => mn[m.month - 1] + String(m.yearCE + 543));
  const labels = [], attendanceRates = [], presentCounts = [], absentCounts = [];

  monthOrder.forEach(month => {
    if (data.monthly[month]) {
      const yearStr = String(month.match(/\d+/));
      labels.push(month.replace(yearStr, '').trim());
      attendanceRates.push(parseFloat(data.monthly[month].attendanceRate));
      presentCounts.push(data.monthly[month].totalPresent);
      absentCounts.push(data.monthly[month].totalAbsent);
    }
  });

  return {
    labels,
    datasets: [
      { label: 'อัตราการมาเรียน (%)', data: attendanceRates, borderColor: 'rgb(75, 192, 192)', backgroundColor: 'rgba(75, 192, 192, 0.2)', yAxisID: 'y-percentage' },
      { label: 'จำนวนวันที่มา', data: presentCounts, borderColor: 'rgb(54, 162, 235)', backgroundColor: 'rgba(54, 162, 235, 0.2)', yAxisID: 'y-count' },
      { label: 'จำนวนวันที่ขาด', data: absentCounts, borderColor: 'rgb(255, 99, 132)', backgroundColor: 'rgba(255, 99, 132, 0.2)', yAxisID: 'y-count' }
    ]
  };
}

function getAttendanceGradeComparisonChartData(monthName = null) {
  if (monthName === null) monthName = getCurrentThaiMonth();
  const comparison = compareAttendanceByGrades(monthName === 'yearly' ? null : monthName);
  const labels = comparison.map(i => i.grade);
  const rates = comparison.map(i => i.attendanceRate);

  return {
    labels,
    datasets: [{
      label: 'อัตราการมาเรียน (%)', data: rates,
      backgroundColor: rates.map(r => r >= 95 ? 'rgba(75, 192, 192, 0.6)' : r >= 90 ? 'rgba(255, 206, 86, 0.6)' : 'rgba(255, 99, 132, 0.6)'),
      borderColor: rates.map(r => r >= 95 ? 'rgb(75, 192, 192)' : r >= 90 ? 'rgb(255, 206, 86)' : 'rgb(255, 99, 132)'),
      borderWidth: 2
    }]
  };
}

// ============================================================
// 📋 ATTENDANCE REPORTS
// ============================================================

function generateMonthlyAttendanceReport(monthName) {
  const data = getAttendanceByMonth(monthName);
  if (!data || !data.grades || Object.keys(data.grades).length === 0) throw new Error(`ไม่พบข้อมูลสำหรับเดือน ${monthName}`);

  const report = { title: `รายงานการเข้าเรียน - ${monthName}`, generatedAt: new Date().toISOString(), summary: data.totalStats, gradeDetails: [] };
  const sortedGrades = Object.keys(data.grades).sort((a, b) => { const p = g => { const m = g.match(/^ป\.(\d+)$/); return m ? parseInt(m[1]) : 999; }; return p(a) - p(b); });

  sortedGrades.forEach(grade => {
    const gd = data.grades[grade];
    report.gradeDetails.push({ grade, stats: gd, avgAttendancePerStudent: gd.studentCount > 0 ? (gd.totalPresent / gd.studentCount).toFixed(1) : 0 });
  });
  return report;
}

function generateYearlyAttendanceReport(grade) {
  const data = getAttendanceByGrade(grade);
  if (!data || !data.monthly) throw new Error(`ไม่พบข้อมูลสำหรับห้อง ${grade}`);

  const report = { title: `รายงานการเข้าเรียนรายปี - ${grade}`, generatedAt: new Date().toISOString(), yearlyStats: data.yearly, monthlyBreakdown: [] };
  const mn = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  let ay; try { ay = getCurrentAcademicYear(); } catch (e) { ay = 2568; }
  const monthOrder = getMonthsInAcademicYear(ay).map(m => mn[m.month - 1] + String(m.yearCE + 543));
  monthOrder.forEach(month => { if (data.monthly[month]) report.monthlyBreakdown.push({ month, stats: data.monthly[month] }); });
  return report;
}

// ============================================================
// 📚 ACADEMIC PERFORMANCE
// ============================================================

function normalizeGradeName(g) { return g ? String(g).trim().replace(/\s+/g,'').replace(/\/\d+$/,'') : ''; }
function findHeaderIndex(headers, candidates) { const lower = headers.map(h => String(h||'').trim().toLowerCase()); for (const c of candidates) { const idx = lower.indexOf(String(c).toLowerCase()); if (idx !== -1) return idx; } return -1; }
function pick(row, idx) { return idx < 0 ? '' : row[idx]; }
function gradeRank(g) { const m = String(g||'').trim().match(/^ป\.(\d+)/); return m ? parseInt(m[1], 10) : 999; }
function getUniqueStudentCount(scores) { return new Set(scores.map(s => s.student_id)).size; }
function getUniqueSubjectCount(scores) { const u = new Set(scores.map(s => s.subject_code || '')); return u.has('') ? u.size - 1 : u.size; }

function getScoresData(academicYear = null) {
  try {
    const ss = SS();
    const sheet = S_getYearlySheet('SCORES_WAREHOUSE');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0]; const rows = data.slice(1);
    const idx = {
      student_id: findHeaderIndex(headers, ['student_id','studentid','รหัสนักเรียน','รหัส']),
      grade: findHeaderIndex(headers, ['grade','ชั้น','ชั้นเรียน','ระดับชั้น']),
      subject_code: findHeaderIndex(headers, ['subject_code','subjectcode','รหัสวิชา']),
      subject_name: findHeaderIndex(headers, ['subject_name','subject','ชื่อวิชา']),
      subject_type: findHeaderIndex(headers, ['subject_type','ประเภทวิชา','กลุ่มสาระ']),
      average: findHeaderIndex(headers, ['average','avg','คะแนนเฉลี่ย','เฉลี่ย']),
      final_grade: findHeaderIndex(headers, ['final_grade','grade_final','เกรด','ผลการประเมิน','final']),
      term1_total: findHeaderIndex(headers, ['term1_total','เทอม1','คะแนนรวมเทอม1','คะแนนระหว่างภาค']),
      term2_total: findHeaderIndex(headers, ['term2_total','เทอม2','คะแนนรวมเทอม2','คะแนนปลายภาค']),
      academic_year: findHeaderIndex(headers, ['academic_year','ปีการศึกษา'])
    };

    let scores = rows.map(r => ({
      student_id: String(pick(r, idx.student_id)).trim(),
      grade: normalizeGradeName(pick(r, idx.grade)),
      subject_code: String(pick(r, idx.subject_code)).trim(),
      subject_name: String(pick(r, idx.subject_name)).trim(),
      subject_type: String(pick(r, idx.subject_type)).trim(),
      average: Number(pick(r, idx.average)) || 0,
      final_grade: Number(pick(r, idx.final_grade)) || 0,
      term1_total: Number(pick(r, idx.term1_total)) || 0,
      term2_total: Number(pick(r, idx.term2_total)) || 0,
      academic_year: String(pick(r, idx.academic_year)).trim()
    })).filter(o => o.student_id && o.grade);

    if (academicYear) scores = scores.filter(s => String(s.academic_year) === String(academicYear));
    return scores;
  } catch (e) { Logger.log("❌ getScoresData Error: " + e.message); return []; }
}

function calculateAcademicSummary(academicYear = null) {
  const scores = getScoresData(academicYear);
  if (scores.length === 0) return { hasData: false, message: "ยังไม่มีข้อมูลผลการเรียน", overall: null, byGrade: {} };

  const byGrade = {};
  const gradesList = [...new Set(scores.map(s => s.grade))].filter(Boolean);

  gradesList.forEach(grade => {
    const gs = scores.filter(s => s.grade === grade);
    const vs = gs.filter(s => s.average > 0);
    const vg = gs.filter(s => s.final_grade > 0);
    const avgScore = vs.length ? vs.reduce((sum, s) => sum + s.average, 0) / vs.length : 0;
    const avgGrade = vg.length ? vg.reduce((sum, s) => sum + s.final_grade, 0) / vg.length : 0;
    const passCount = vg.filter(s => s.final_grade >= 1).length;

    byGrade[grade] = {
      hasData: gs.length > 0, avgScore: Number(avgScore.toFixed(2)), avgGrade: Number(avgGrade.toFixed(2)),
      passRate: Number((vg.length ? passCount / vg.length * 100 : 0).toFixed(1)),
      studentCount: getUniqueStudentCount(gs), subjectCount: getUniqueSubjectCount(gs), recordCount: gs.length
    };
  });

  const vsAll = scores.filter(s => s.average > 0);
  const vgAll = scores.filter(s => s.final_grade > 0);
  const overall = {
    avgScore: Number((vsAll.length ? vsAll.reduce((a, b) => a + b.average, 0) / vsAll.length : 0).toFixed(2)),
    avgGrade: Number((vgAll.length ? vgAll.reduce((a, b) => a + b.final_grade, 0) / vgAll.length : 0).toFixed(2)),
    passRate: Number((vgAll.length ? vgAll.filter(s => s.final_grade >= 1).length / vgAll.length * 100 : 0).toFixed(1)),
    totalRecords: vgAll.length, totalStudents: getUniqueStudentCount(scores)
  };

  return { hasData: true, overall, byGrade };
}

function getAcademicDashboardSummary(academicYear = null, force = false) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `academic_summary_${academicYear || 'all'}_v2`;
    if (force) cache.remove(cacheKey);

    const hit = cache.get(cacheKey);
    if (hit != null) return JSON.parse(hit);

    const start = Date.now();
    const summary = calculateAcademicSummary(academicYear);
    summary.processingTime = Date.now() - start;
    summary.lastUpdated = new Date().toISOString();
    cache.put(cacheKey, JSON.stringify(summary), 900);
    return summary;
  } catch (e) {
    return { error: true, message: 'ไม่สามารถโหลดข้อมูลผลการเรียนได้: ' + e.message };
  }
}

function refreshAcademicSummary(academicYear = null) { return getAcademicDashboardSummary(academicYear, true); }

// ============================================================
// 📊 ACADEMIC CHARTS
// ============================================================

function getGradeComparisonChartData(academicYear = null) {
  const summary = calculateAcademicSummary(academicYear);
  if (!summary.hasData) return { labels: [], datasets: [], hasData: false };

  const gradesList = Object.keys(summary.byGrade || {}).filter(g => summary.byGrade[g]?.hasData).sort((a, b) => gradeRank(a) - gradeRank(b));
  const labels = [], avgScores = [], avgGrades = [], studentCounts = [], bgColors = [], bdColors = [];
  const palette = ['rgba(54,162,235,0.8)','rgba(255,99,132,0.8)','rgba(75,192,192,0.8)','rgba(255,206,86,0.8)','rgba(153,102,255,0.8)','rgba(255,159,64,0.8)'];

  gradesList.forEach((grade, i) => {
    const gd = summary.byGrade[grade];
    labels.push(grade); avgScores.push(gd.avgScore); avgGrades.push(gd.avgGrade); studentCounts.push(gd.studentCount);
    bgColors.push(palette[i % palette.length]); bdColors.push(palette[i % palette.length].replace('0.8', '1'));
  });

  return { labels, datasets: [{ label: 'คะแนนเฉลี่ย', data: avgScores, backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 0 }], metadata: { grades: avgGrades, studentCounts }, hasData: labels.length > 0 };
}

function getOverallGradeDistributionChartData(academicYear = null) {
  const scores = getScoresData(academicYear).filter(s => !isNaN(Number(s.final_grade)));
  if (scores.length === 0) return { hasData: false };

  const gradeTypes = ['4','3.5','3','2.5','2','1.5','1','0'];
  const counts = gradeTypes.map(gt => scores.filter(s => Number(s.final_grade).toFixed(1) === Number(gt).toFixed(1)).length);
  const total = counts.reduce((a, b) => a + b, 0);

  return {
    hasData: true,
    labels: gradeTypes.map(g => `เกรด ${g}`),
    datasets: [{ label: 'สัดส่วนรวมทุกชั้น', data: counts.map(c => total ? Math.round((c / total) * 1000) / 10 : 0), backgroundColor: gradeTypes.map(getGradeColor), borderWidth: 1 }],
    totalRecords: total, rawCounts: counts
  };
}

function getGradeDistributionChartData(academicYear = null) {
  const summary = calculateAcademicSummary(academicYear);
  if (!summary.hasData) return { labels: [], datasets: [], hasData: false };

  const allScores = getScoresData(academicYear);
  const gradesList = Object.keys(summary.byGrade || {}).filter(g => summary.byGrade[g]?.hasData).sort((a, b) => gradeRank(a) - gradeRank(b));
  const labels = [];
  const gradeTypes = ['4','3.5','3','2.5','2','1.5','1','0'];
  const datasets = gradeTypes.map(gt => ({ label: `เกรด ${gt}`, data: [], backgroundColor: getGradeColor(gt), borderWidth: 0 }));

  gradesList.forEach(grade => {
    labels.push(grade);
    const gs = allScores.filter(s => s.grade === grade && s.final_grade > 0);
    gradeTypes.forEach((gType, i) => {
      const cnt = gs.filter(s => Number(s.final_grade).toFixed(1) === Number(gType).toFixed(1)).length;
      datasets[i].data.push(Number((gs.length ? cnt / gs.length * 100 : 0).toFixed(1)));
    });
  });

  return { labels, datasets, hasData: labels.length > 0 };
}

function getGradeColor(grade) {
  const g = parseFloat(grade);
  if (g >= 4) return 'rgba(40, 167, 69, 0.8)';
  if (g >= 3.5) return 'rgba(40, 167, 69, 0.6)';
  if (g >= 3) return 'rgba(23, 162, 184, 0.8)';
  if (g >= 2.5) return 'rgba(255, 193, 7, 0.8)';
  if (g >= 2) return 'rgba(255, 152, 0, 0.8)';
  if (g >= 1.5) return 'rgba(255, 111, 0, 0.8)';
  if (g >= 1) return 'rgba(255, 99, 132, 0.8)';
  return 'rgba(180, 60, 90, 0.8)';
}

// ============================================================
// 🗑️ CACHE & UTILITIES
// ============================================================

/**
 * รวม API calls กราฟ 2 ตัวเป็น call เดียว — ลดการอ่าน SCORES_WAREHOUSE จาก 3 รอบเหลือ 1 รอบ
 */
function getAllDashboardChartData(academicYear) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'all_chart_data_' + (academicYear || 'all') + '_v1';
    const hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);

    const scores = getScoresData(academicYear);
    if (scores.length === 0) return { gradeComparison: { labels: [], datasets: [], hasData: false }, gradeDistribution: { hasData: false } };

    // --- Grade Comparison ---
    const summary = calculateAcademicSummary(academicYear);
    const gradesList = Object.keys(summary.byGrade || {}).filter(g => summary.byGrade[g]?.hasData).sort((a, b) => gradeRank(a) - gradeRank(b));
    const labels = [], avgScores = [], avgGrades = [], studentCounts = [], bgColors = [], bdColors = [];
    const palette = ['rgba(54,162,235,0.8)','rgba(255,99,132,0.8)','rgba(75,192,192,0.8)','rgba(255,206,86,0.8)','rgba(153,102,255,0.8)','rgba(255,159,64,0.8)'];
    gradesList.forEach((grade, i) => {
      const gd = summary.byGrade[grade];
      labels.push(grade); avgScores.push(gd.avgScore); avgGrades.push(gd.avgGrade); studentCounts.push(gd.studentCount);
      bgColors.push(palette[i % palette.length]); bdColors.push(palette[i % palette.length].replace('0.8', '1'));
    });
    const gradeComparison = { labels, datasets: [{ label: 'คะแนนเฉลี่ย', data: avgScores, backgroundColor: bgColors, borderColor: bdColors, borderWidth: 1, borderRadius: 0 }], metadata: { grades: avgGrades, studentCounts }, hasData: labels.length > 0 };

    // --- Grade Distribution ---
    const validScores = scores.filter(s => !isNaN(Number(s.final_grade)));
    const gradeTypes = ['4','3.5','3','2.5','2','1.5','1','0'];
    const counts = gradeTypes.map(gt => validScores.filter(s => Number(s.final_grade).toFixed(1) === Number(gt).toFixed(1)).length);
    const total = counts.reduce((a, b) => a + b, 0);
    const gradeDistribution = {
      hasData: total > 0,
      labels: gradeTypes.map(g => 'เกรด ' + g),
      datasets: [{ label: 'สัดส่วนรวมทุกชั้น', data: counts.map(c => total ? Math.round((c / total) * 1000) / 10 : 0), backgroundColor: gradeTypes.map(getGradeColor), borderWidth: 1 }],
      totalRecords: total, rawCounts: counts
    };

    const result = { gradeComparison, gradeDistribution, academicSummary: summary };
    cache.put(cacheKey, JSON.stringify(result), 900);
    return result;
  } catch(e) {
    Logger.log('getAllDashboardChartData error: ' + e.message);
    return { gradeComparison: { labels: [], datasets: [], hasData: false }, gradeDistribution: { hasData: false }, error: e.message };
  }
}

function clearDashboardCache() { DashboardCacheManager.clearAll(); return "✅ ล้าง Dashboard Cache เรียบร้อย"; }
function clearAcademicCache() { AcademicCacheManager.clearAll(); return "✅ ล้าง Academic Cache เรียบร้อย"; }

function clearServerCache() {
  try {
    CacheService.getScriptCache().removeAll(['dashboard_summary_v5', 'dashboard_summary_v3', 'academic_summary_all_v2', 'fast_student_stats_v1', 'all_chart_data_all_v1', 'dashboard_progress_v1', 'dashboard_alerts_v1', 'grade_recording_chart_v1', 'grade_recording_chart_v2']);
    return "✅ ล้าง Server Cache เรียบร้อย";
  } catch (e) { return { error: true, message: e.message }; }
}

function getDashboardWithData() {
  try {
    var html = HtmlService.createTemplateFromFile('spa_dashboard').evaluate().getContent();
    var summary = getCachedDashboardSummary();
    var settings = getWebAppSettingsWithCache();
    var currentMonth = getCurrentThaiMonth();
    summary.currentMonth = currentMonth;
    return { html, summary, settings, currentMonth };
  } catch (e) {
    throw new Error('Failed to load dashboard data: ' + e.message);
  }
}

// ============================================================
// 📋 DASHBOARD PROGRESS & ALERTS
// ============================================================

/**
 * ดึงสถานะการกรอกข้อมูลสำหรับ Dashboard
 */
function getDashboardProgress() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_progress_v1';
    const hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);

    const ss = SS();
    const items = [];

    // 1. ตรวจสอบคะแนนรายวิชา (SCORES_WAREHOUSE)
    try {
      const scoresSheet = S_getYearlySheet('SCORES_WAREHOUSE');
      if (scoresSheet && scoresSheet.getLastRow() > 1) {
        const scoresData = scoresSheet.getDataRange().getValues();
        const headers = scoresData[0];
        const gradeIdx = findHeaderIndex(headers, ['grade','ชั้น','ระดับชั้น']);
        const subjectIdx = findHeaderIndex(headers, ['subject_code','subjectcode','รหัสวิชา']);
        
        const uniqueGrades = new Set();
        const uniqueSubjects = new Set();
        for (let i = 1; i < scoresData.length; i++) {
          if (gradeIdx >= 0) uniqueGrades.add(String(scoresData[i][gradeIdx]).trim());
          if (subjectIdx >= 0) uniqueSubjects.add(String(scoresData[i][subjectIdx]).trim());
        }
        uniqueGrades.delete('');
        uniqueSubjects.delete('');

        items.push({
          label: 'คะแนนรายวิชา',
          completed: uniqueSubjects.size,
          total: Math.max(uniqueSubjects.size, 8),
          unit: 'วิชา',
          icon: 'bi-pencil-square',
          iconBg: '#6366f1'
        });

        items.push({
          label: 'ชั้นเรียนที่มีข้อมูล',
          completed: uniqueGrades.size,
          total: Math.max(uniqueGrades.size, 6),
          unit: 'ชั้น',
          icon: 'bi-mortarboard',
          iconBg: '#8b5cf6'
        });
      } else {
        items.push({ label: 'คะแนนรายวิชา', completed: 0, total: 8, unit: 'วิชา', icon: 'bi-pencil-square', iconBg: '#6366f1' });
      }
    } catch(e) { Logger.log('Progress scores error: ' + e.message); }

    // 2. ตรวจสอบการเช็คชื่อ (เดือนปัจจุบัน)
    try {
      const currentMonth = getCurrentThaiMonth();
      const monthSheet = ss.getSheetByName(currentMonth);
      if (monthSheet && monthSheet.getLastRow() > 1) {
        const mData = monthSheet.getDataRange().getValues();
        const mHeaders = mData[0];
        const presentIdx = mHeaders.indexOf('มา');
        let checkedCount = 0;
        for (let i = 1; i < mData.length; i++) {
          if (presentIdx >= 0 && Number(mData[i][presentIdx]) > 0) checkedCount++;
        }
        const totalStudents = mData.length - 1;
        items.push({
          label: 'เช็คชื่อเดือนนี้ (' + currentMonth.replace(/\d+/g,'') + ')',
          completed: checkedCount,
          total: totalStudents,
          unit: 'คน',
          icon: 'bi-calendar-check',
          iconBg: '#10b981'
        });
      }
    } catch(e) { Logger.log('Progress attendance error: ' + e.message); }

    // 3. ตรวจสอบการประเมินอ่าน คิด เขียน
    try {
      const rtwSheet = S_getYearlySheet('การประเมินอ่านคิดเขียน');
      if (rtwSheet && rtwSheet.getLastRow() > 1) {
        const rtwData = rtwSheet.getDataRange().getValues();
        const studentsSheet = ss.getSheetByName('Students');
        const totalStudents = studentsSheet ? Math.max(studentsSheet.getLastRow() - 1, 0) : 0;
        const rtwStudents = new Set();
        for (let i = 1; i < rtwData.length; i++) {
          const sid = String(rtwData[i][0] || '').trim();
          if (sid) rtwStudents.add(sid);
        }
        items.push({
          label: 'ประเมินอ่าน คิด เขียน',
          completed: rtwStudents.size,
          total: totalStudents || rtwStudents.size,
          unit: 'คน',
          icon: 'bi-journal-text',
          iconBg: '#ec4899'
        });
      } else {
        items.push({ label: 'ประเมินอ่าน คิด เขียน', completed: 0, total: 1, unit: '', icon: 'bi-journal-text', iconBg: '#ec4899' });
      }
    } catch(e) { Logger.log('Progress RTW error: ' + e.message); }

    // 4. ตรวจสอบการประเมินคุณลักษณะ
    try {
      const charSheet = S_getYearlySheet('การประเมินคุณลักษณะ');
      if (charSheet && charSheet.getLastRow() > 1) {
        items.push({
          label: 'ประเมินคุณลักษณะ',
          completed: charSheet.getLastRow() - 1,
          total: charSheet.getLastRow() - 1,
          unit: 'รายการ',
          icon: 'bi-check-circle',
          iconBg: '#f59e0b'
        });
      } else {
        items.push({ label: 'ประเมินคุณลักษณะ', completed: 0, total: 1, unit: '', icon: 'bi-check-circle', iconBg: '#f59e0b' });
      }
    } catch(e) { Logger.log('Progress char error: ' + e.message); }

    const result = { items: items };
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch(e) {
    Logger.log('getDashboardProgress error: ' + e.message);
    return { error: true, message: e.message, items: [] };
  }
}

/**
 * ดึงแจ้งเตือนสำหรับ Dashboard
 */
function getDashboardAlerts() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_alerts_v1';
    const hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);

    const ss = SS();
    const alerts = [];

    // 1. ตรวจสอบนักเรียนที่ได้เกรด 0 + สถานะการบันทึกเกรดแต่ละชั้น/วิชา
    try {
      const scoresSheet = S_getYearlySheet('SCORES_WAREHOUSE');
      if (scoresSheet && scoresSheet.getLastRow() > 1) {
        const sData = scoresSheet.getDataRange().getValues();
        const sHeaders = sData[0];
        const finalGradeIdx = findHeaderIndex(sHeaders, ['final_grade','grade_final','เกรด','ผลการประเมิน','final']);
        const classIdx = findHeaderIndex(sHeaders, ['grade','ชั้น','ระดับชั้น']);
        const subjectIdx = findHeaderIndex(sHeaders, ['subject_name','subject','ชื่อวิชา']);
        
        // วิชากิจกรรม — ไม่นับ (เหมือน getGradeRecordingChartData)
        var actKw = ['แนะแนว','ลูกเสือ','เนตรนารี','ชุมนุม','กิจกรรมเพื่อสังคม','กิจกรรมพัฒนาผู้เรียน','ยุวกาชาด','ลูกเสือ-เนตรนารี'];
        function _isAct(n) { var t = String(n||'').trim(); return actKw.some(function(k){return t.indexOf(k)>=0;}); }

        // นับนักเรียนต่อชั้น
        var _studPerGrade = {};
        try {
          var _stSheet = ss.getSheetByName('Students');
          if (_stSheet && _stSheet.getLastRow() > 1) {
            var _stD = _stSheet.getDataRange().getValues();
            var _gC = _stD[0].indexOf('grade'); if (_gC<0) _gC = _stD[0].indexOf('ชั้น');
            if (_gC >= 0) { for (var si=1;si<_stD.length;si++) { var _sg=String(_stD[si][_gC]||'').trim(); _studPerGrade[_sg]=(_studPerGrade[_sg]||0)+1; } }
          }
        } catch(e2) {}

        let zeroGradeCount = 0;
        const zeroSubjects = new Set();
        // gradeSubjCount[grade][subject] = จำนวนนักเรียน
        var gradeSubjCount = {};
        
        for (let i = 1; i < sData.length; i++) {
          var fg = finalGradeIdx >= 0 ? Number(sData[i][finalGradeIdx]) : -1;
          var _subj = subjectIdx >= 0 ? String(sData[i][subjectIdx] || '').trim() : '';
          var _cls = classIdx >= 0 ? String(sData[i][classIdx] || '').trim() : '';

          if (fg === 0 && !_isAct(_subj)) {
            zeroGradeCount++;
            if (_subj) zeroSubjects.add(_subj);
          }
          
          // สะสมข้อมูลชั้น/วิชาที่บันทึกเกรดแล้ว (ไม่นับกิจกรรม)
          if (classIdx >= 0 && subjectIdx >= 0 && fg > 0 && _cls && _subj && !_isAct(_subj)) {
            if (!gradeSubjCount[_cls]) gradeSubjCount[_cls] = {};
            gradeSubjCount[_cls][_subj] = (gradeSubjCount[_cls][_subj] || 0) + 1;
          }
        }
        
        if (zeroGradeCount > 0) {
          alerts.push({
            type: 'danger',
            icon: 'bi-exclamation-triangle-fill',
            message: 'พบนักเรียนได้เกรด 0 จำนวน <strong>' + zeroGradeCount + '</strong> รายการ ใน ' + zeroSubjects.size + ' วิชา'
          });
        }
        
        // กรองเฉพาะวิชาที่มีนักเรียน >= 50% ของชั้น
        var gradeSubjectMap = {};
        Object.keys(gradeSubjCount).forEach(function(g) {
          var total = _studPerGrade[g] || 1;
          var minReq = Math.max(Math.ceil(total * 0.98), 2);
          gradeSubjectMap[g] = [];
          Object.keys(gradeSubjCount[g]).forEach(function(s) {
            if (gradeSubjCount[g][s] >= minReq) gradeSubjectMap[g].push(s);
          });
          if (gradeSubjectMap[g].length === 0) delete gradeSubjectMap[g];
        });

        // (สถานะการบันทึกเกรดแสดงเป็นแผนภูมิแยกด้วย — ดู getGradeRecordingChartData)
        var sortedGrades = Object.keys(gradeSubjectMap).sort(function(a, b) {
          var pa = a.match(/(\d+)/); var pb = b.match(/(\d+)/);
          return (pa ? parseInt(pa[1]) : 999) - (pb ? parseInt(pb[1]) : 999);
        });
        
        if (sortedGrades.length > 0) {
          var totalSubjects = new Set();
          sortedGrades.forEach(function(g) { gradeSubjectMap[g].forEach(function(s) { totalSubjects.add(s); }); });
          
          var msg = '<strong>สถานะการบันทึกเกรด</strong> (' + totalSubjects.size + ' วิชา, ' + sortedGrades.length + ' ชั้น)';
          msg += '<div style="margin-top:8px;font-size:0.9em">';
          sortedGrades.forEach(function(grade) {
            var subjects = gradeSubjectMap[grade].sort();
            msg += '<div style="margin-bottom:4px;padding:4px 8px;background:rgba(16,185,129,0.08);border-radius:6px">'
              + '<i class="bi bi-check-circle-fill text-success me-1"></i>'
              + '<strong>' + grade + '</strong> — '
              + subjects.length + ' วิชา: '
              + '<span style="color:#059669">' + subjects.join(', ') + '</span>'
              + '</div>';
          });
          msg += '</div>';
          
          alerts.push({
            type: 'success',
            icon: 'bi-journal-check',
            message: msg
          });
        }
      }
    } catch(e) { Logger.log('Alert grade0 error: ' + e.message); }

    // 2. ตรวจสอบ settings ที่ยังไม่ได้ตั้งค่า
    try {
      var settings = getGlobalSettings(true);
      var schoolName = String(settings['ชื่อโรงเรียน'] || '').trim();
      if (!schoolName || schoolName === 'โรงเรียนตัวอย่าง' || schoolName === 'โรงเรียน...') {
        alerts.push({
          type: 'warning',
          icon: 'bi-gear-fill',
          message: 'ยังไม่ได้ตั้งค่า<strong>ชื่อโรงเรียน</strong> — ไปที่ ระบบ > ตั้งค่าระบบ'
        });
      }
      var director = String(settings['ชื่อผู้อำนวยการ'] || '').trim();
      if (!director || director === '...' || director === 'ผู้อำนวยการ') {
        alerts.push({
          type: 'warning',
          icon: 'bi-person-fill-gear',
          message: 'ยังไม่ได้ตั้งค่า<strong>ชื่อผู้อำนวยการ</strong>'
        });
      }
    } catch(e) { Logger.log('Alert settings error: ' + e.message); }

    // 3. ตรวจสอบการเช็คชื่อเดือนนี้
    try {
      const currentMonth = getCurrentThaiMonth();
      const monthSheet = ss.getSheetByName(currentMonth);
      if (!monthSheet) {
        alerts.push({
          type: 'info',
          icon: 'bi-calendar-x',
          message: 'ยังไม่มีข้อมูลเช็คชื่อเดือน <strong>' + currentMonth.replace(/\d+/g,'') + '</strong>'
        });
      }
    } catch(e) {}

    const result = { alerts: alerts };
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch(e) {
    Logger.log('getDashboardAlerts error: ' + e.message);
    return { error: true, alerts: [] };
  }
}

// ============================================================
// GRADE RECORDING CHART DATA
// ============================================================

function getGradeRecordingChartData() {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'grade_recording_chart_v2';
    var hit = cache.get(cacheKey);
    if (hit) return JSON.parse(hit);

    var ss = SS();
    var allGrades = ['ป.1','ป.2','ป.3','ป.4','ป.5','ป.6'];

    // วิชากิจกรรม — ไม่นับ
    var activityKeywords = ['แนะแนว','ลูกเสือ','เนตรนารี','ชุมนุม','กิจกรรมเพื่อสังคม','กิจกรรมพัฒนาผู้เรียน',
      'ยุวกาชาด','ชุมนุมวิทยาศาสตร์','ชุมนุมคณิตศาสตร์','ลูกเสือ-เนตรนารี'];
    function isActivitySubject(name) {
      var n = String(name || '').trim();
      return activityKeywords.some(function(kw) { return n.indexOf(kw) >= 0; });
    }

    // 0. นับจำนวนนักเรียนต่อชั้น (จาก Students sheet)
    var studentsPerGrade = {};
    allGrades.forEach(function(g) { studentsPerGrade[g] = 0; });
    try {
      var studSheet = ss.getSheetByName('Students');
      if (studSheet && studSheet.getLastRow() > 1) {
        var stData = studSheet.getDataRange().getValues();
        var stHeaders = stData[0];
        var gCol = stHeaders.indexOf('grade');
        if (gCol < 0) gCol = stHeaders.indexOf('ชั้น');
        if (gCol >= 0) {
          for (var s = 1; s < stData.length; s++) {
            var g = String(stData[s][gCol] || '').trim();
            if (studentsPerGrade.hasOwnProperty(g)) studentsPerGrade[g]++;
          }
        }
      }
    } catch(e) { Logger.log('Chart students count: ' + e.message); }

    // 1. อ่านวิชาที่คาดหวังจากชีตรายวิชา (ไม่รวมกิจกรรม)
    var expectedByGrade = {};
    allGrades.forEach(function(g) { expectedByGrade[g] = []; });
    try {
      var subjSheet = ss.getSheetByName('รายวิชา');
      if (subjSheet) {
        var sd = subjSheet.getDataRange().getValues();
        var sh = sd[0];
        var nameCol = sh.findIndex(function(v) { return /ชื่อวิชา/i.test(v); });
        var gradeCol = sh.findIndex(function(v) { return /^ชั้น$|ระดับชั้น/i.test(String(v || '').trim()); });
        if (nameCol >= 0) {
          for (var i = 1; i < sd.length; i++) {
            var sn = String(sd[i][nameCol] || '').trim();
            var sg = gradeCol >= 0 ? String(sd[i][gradeCol] || '').trim() : '';
            if (sn && sg && expectedByGrade[sg] && !isActivitySubject(sn)) {
              if (expectedByGrade[sg].indexOf(sn) < 0) expectedByGrade[sg].push(sn);
            }
          }
        }
      }
    } catch(e) { Logger.log('Chart subject lookup: ' + e.message); }

    // 2. อ่าน SCORES_WAREHOUSE — นับจำนวนนักเรียนต่อชั้น/วิชา (เฉพาะเกรด > 0)
    var recordedByGrade = {};
    allGrades.forEach(function(g) { recordedByGrade[g] = []; });
    var allRecordedSubjects = [];

    try {
      var scoresSheet = S_getYearlySheet('SCORES_WAREHOUSE');
      if (scoresSheet && scoresSheet.getLastRow() > 1) {
        var sData = scoresSheet.getDataRange().getValues();
        var sHeaders = sData[0];
        var fgIdx = findHeaderIndex(sHeaders, ['final_grade','grade_final','เกรด','ผลการประเมิน','final']);
        var clsIdx = findHeaderIndex(sHeaders, ['grade','ชั้น','ระดับชั้น']);
        var subjIdx = findHeaderIndex(sHeaders, ['subject_name','subject','ชื่อวิชา']);

        // gradeSubjCount[grade][subject] = จำนวนนักเรียนที่มีเกรด > 0
        var gradeSubjCount = {};
        for (var j = 1; j < sData.length; j++) {
          var fg = fgIdx >= 0 ? Number(sData[j][fgIdx]) : -1;
          if (fg <= 0) continue;
          var cls = clsIdx >= 0 ? String(sData[j][clsIdx] || '').trim() : '';
          var subj = subjIdx >= 0 ? String(sData[j][subjIdx] || '').trim() : '';
          if (!cls || !subj) continue;
          if (isActivitySubject(subj)) continue;
          if (!gradeSubjCount[cls]) gradeSubjCount[cls] = {};
          gradeSubjCount[cls][subj] = (gradeSubjCount[cls][subj] || 0) + 1;
        }

        // ต้องมีนักเรียนอย่างน้อย 50% ของชั้นจึงจะนับว่า "บันทึกแล้ว"
        allGrades.forEach(function(g) {
          if (!gradeSubjCount[g]) return;
          var totalStudents = studentsPerGrade[g] || 1;
          var minRequired = Math.max(Math.ceil(totalStudents * 0.98), 2);
          var recorded = [];
          Object.keys(gradeSubjCount[g]).forEach(function(subj) {
            if (gradeSubjCount[g][subj] >= minRequired) {
              recorded.push(subj);
            }
          });
          recordedByGrade[g] = recorded.sort();
        });

        // รวมวิชาทั้งหมดที่บันทึกครบ (ไม่รวมกิจกรรม)
        var allSet = {};
        allGrades.forEach(function(g) {
          recordedByGrade[g].forEach(function(s) { allSet[s] = true; });
        });
        allRecordedSubjects = Object.keys(allSet).sort();
      }
    } catch(e) { Logger.log('Chart warehouse read: ' + e.message); }

    // 3. สร้างข้อมูลแผนภูมิ
    var grades = [];
    allGrades.forEach(function(g) {
      var expected = expectedByGrade[g].length > 0 ? expectedByGrade[g] : allRecordedSubjects;
      // กรองกิจกรรมออกจาก expected ด้วย
      expected = expected.filter(function(s) { return !isActivitySubject(s); });
      var recorded = recordedByGrade[g];
      var total = Math.max(expected.length, recorded.length, 1);
      var pct = total > 0 ? Math.round(recorded.length / total * 100) : 0;
      grades.push({
        name: g,
        recordedCount: recorded.length,
        total: total,
        pct: pct
      });
    });

    var totalSubjects = 0;
    var allExpected = {};
    allGrades.forEach(function(g) {
      expectedByGrade[g].forEach(function(s) { allExpected[s] = true; });
    });
    totalSubjects = Object.keys(allExpected).length || allRecordedSubjects.length;

    var result = { grades: grades, totalSubjects: totalSubjects, hasData: allRecordedSubjects.length > 0 || totalSubjects > 0 };
    cache.put(cacheKey, JSON.stringify(result), 300);
    return result;
  } catch(e) {
    Logger.log('getGradeRecordingChartData error: ' + e.message);
    return { grades: [], totalSubjects: 0, hasData: false };
  }
}

// ============================================================
// DEBUG: Attendance Dashboard
// ============================================================
function debugAttendanceDashboard() {
  const result = [];
  const ss = SS();

  // 1. ปีการศึกษา
  let academicYear;
  try {
    academicYear = getCurrentAcademicYear();
    result.push('✅ ปีการศึกษา: ' + academicYear);
  } catch (e) {
    result.push('❌ อ่านปีการศึกษาไม่ได้: ' + e.message);
    return result.join('\n');
  }

  // 2. ชื่อชีตที่คาดหวัง
  const monthNames = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const monthsInYear = getMonthsInAcademicYear(academicYear);
  const expectedSheets = monthsInYear.map(m => monthNames[m.month - 1] + String(m.yearCE + 543));
  result.push('📋 ชีตที่คาดหวัง: ' + expectedSheets.join(', '));

  // 3. ตรวจว่าชีตมีอยู่จริงไหม
  const allSheets = ss.getSheets().map(s => s.getName());
  result.push('📄 ชีตทั้งหมดใน Spreadsheet (' + allSheets.length + '): ' + allSheets.join(', '));

  let foundSheets = [];
  expectedSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      foundSheets.push(name);
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idxId = headers.indexOf('รหัส');
      const idxPresent = headers.indexOf('มา');
      const idxLeave = headers.indexOf('ลา');
      const idxAbsent = headers.indexOf('ขาด');
      result.push('  ✅ ' + name + ' (แถว: ' + data.length + ', headers: รหัส=' + idxId + ' มา=' + idxPresent + ' ลา=' + idxLeave + ' ขาด=' + idxAbsent + ')');

      // ตรวจสอบข้อมูลนักเรียน 3 คนแรก
      if (data.length > 1 && idxId >= 0 && idxPresent >= 0) {
        for (let i = 1; i <= Math.min(3, data.length - 1); i++) {
          const sid = data[i][idxId];
          const present = data[i][idxPresent];
          const leave = idxLeave >= 0 ? data[i][idxLeave] : '-';
          const absent = idxAbsent >= 0 ? data[i][idxAbsent] : '-';
          result.push('    นร.' + i + ': รหัส=' + sid + ' (type=' + typeof sid + ') มา=' + present + ' ลา=' + leave + ' ขาด=' + absent);
        }
      }
    } else {
      result.push('  ❌ ไม่พบ: ' + name);
    }
  });

  // 4. ตรวจสอบ Students sheet
  const studSheet = ss.getSheetByName('Students');
  if (studSheet) {
    const studData = studSheet.getDataRange().getValues();
    const studHeaders = studData[0];
    const sidCol = studHeaders.indexOf('student_id');
    const gradeCol = studHeaders.indexOf('grade');
    result.push('👨‍🎓 Students sheet: แถว=' + studData.length + ' student_id col=' + sidCol + ' grade col=' + gradeCol);
    if (sidCol >= 0 && gradeCol >= 0 && studData.length > 1) {
      for (let i = 1; i <= Math.min(3, studData.length - 1); i++) {
        result.push('    นร.' + i + ': student_id=' + studData[i][sidCol] + ' (type=' + typeof studData[i][sidCol] + ') grade=' + studData[i][gradeCol]);
      }
    }
  } else {
    result.push('❌ ไม่พบชีต Students');
  }

  // 5. สรุป
  result.push('---');
  result.push('พบชีตเช็คชื่อ: ' + foundSheets.length + '/' + expectedSheets.length);

  const output = result.join('\n');
  Logger.log(output);
  try {
    SpreadsheetApp.getUi().alert('🔍 Debug Attendance Dashboard', output, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
  return output;
}