/**
 * ============================================================
 * 📤 EXPORT CSV (FIXED UTF-8 BOM VERSION)
 * ============================================================
 */

function handleExportClassCSV(grade, classNo, sortMode) {
  try {
    const settings = S_getGlobalSettings();
    const academicYear = settings['ปีการศึกษา'] || String(S_getCurrentAcademicYear_());
    
    const result = exportClassScoresToCSV(grade, classNo, academicYear, sortMode);
    
    if (result.success) {
      // ✅ เพิ่ม UTF-8 BOM กลับมา
      const csvWithBOM = '\uFEFF' + result.csvContent;
      
      // แปลงเป็น Base64
      const base64Content = Utilities.base64Encode(csvWithBOM, Utilities.Charset.UTF_8);
      
      return {
        success: true,
        base64Content: base64Content,
        filename: result.filename,
        message: 'สร้าง CSV สำเร็จ'
      };
    } else {
      return {
        success: false,
        error: result.error
      };
    }
    
  } catch (error) {
    Logger.log('handleExportClassCSV Error: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

function exportClassScoresToCSV(grade, classNo, academicYear, sortMode) {
  try {
    const ss = SS();
    
    const students = getStudentsInClass(grade, classNo, sortMode);
    if (!students || students.length === 0) {
      throw new Error('ไม่พบข้อมูลนักเรียนในชั้นนี้');
    }
    
    const gradesData = getGradesFromWarehouseAllYears(grade, classNo, academicYear);
    const activitiesData = getActivitiesDataAllYears(grade, classNo, academicYear);
    const csvContent = buildCSVFromWarehouse(students, gradesData, activitiesData, grade, classNo);
    
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss');
    const safeGrade = grade.replace(/\./g, '');
    const filename = `คะแนน_ชั้น${safeGrade}_ห้อง${classNo}_${academicYear}_${timestamp}.csv`;
    
    return {
      success: true,
      csvContent: csvContent,
      filename: filename
    };
    
  } catch (error) {
    Logger.log('Export CSV Error: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

function getStudentsInClass(grade, classNo, sortMode) {
  const ss = SS();
  const studentsSheet = ss.getSheetByName('Students');
  
  if (!studentsSheet) {
    throw new Error('ไม่พบชีต Students');
  }
  
  const data = studentsSheet.getDataRange().getValues();
  const headers = data[0];
  
  const idCol = headers.indexOf('student_id');
  const titleCol = headers.indexOf('title');
  const firstnameCol = headers.indexOf('firstname');
  const lastnameCol = headers.indexOf('lastname');
  const gradeCol = headers.indexOf('grade');
  const classCol = headers.indexOf('class_no');
  
  const students = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (String(row[gradeCol]) === grade && String(row[classCol]) === classNo) {
      students.push({
        id: String(row[idCol]),
        name: `${row[titleCol] || ''}${row[firstnameCol] || ''} ${row[lastnameCol] || ''}`.trim(),
        grade: String(row[gradeCol]),
        classNo: String(row[classCol])
      });
    }
  }
  
  students.sort(function(a, b) {
    if (sortMode === 'gender') {
      var maleA = a.name.indexOf('เด็กชาย') !== -1 || a.name.indexOf('นาย') === 0;
      var maleB = b.name.indexOf('เด็กชาย') !== -1 || b.name.indexOf('นาย') === 0;
      if (maleA && !maleB) return -1;
      if (!maleA && maleB) return 1;
    }
    return String(a.id).localeCompare(String(b.id), undefined, { numeric: true });
  });
  
  return students;
}

function getGradesFromWarehouseAllYears(grade, classNo, academicYear) {
  const ss = SS();
  const gradesMap = {};
  
  const currentGradeNumber = parseInt(grade.replace('ป.', ''));
  const yearsToFetch = [];
  
  for (let i = 1; i <= currentGradeNumber; i++) {
    const yearOffset = currentGradeNumber - i;
    const targetYear = String(parseInt(academicYear) - yearOffset);
    yearsToFetch.push({
      year: targetYear,
      grade: 'ป.' + i
    });
  }
  
  Logger.log('📚 ดึงข้อมูล ' + yearsToFetch.length + ' ปี: ' + JSON.stringify(yearsToFetch));
  
  // ดึงข้อมูลจากชีตแต่ละปี (แผน B: ชีตแยกตามปี)
  yearsToFetch.forEach(function(yf) {
    var sheet = S_getYearlySheet('SCORES_WAREHOUSE', yf.year);
    if (!sheet) {
      Logger.log('⚠️ ไม่พบชีต SCORES_WAREHOUSE สำหรับปี ' + yf.year);
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    var headers = data[0];
    
    var studentIdCol = headers.indexOf('student_id');
    var gradeCol = headers.indexOf('grade');
    var classCol = headers.indexOf('class_no');
    var subjectCodeCol = headers.indexOf('subject_code');
    var finalGradeCol = headers.indexOf('final_grade');
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowGrade = String(row[gradeCol]);
      var rowClass = String(row[classCol]);
      
      if (rowGrade === yf.grade && rowClass === classNo) {
        var studentId = String(row[studentIdCol]);
        var subjectCode = String(row[subjectCodeCol]);
        var finalGrade = row[finalGradeCol] || '';
        
        if (!gradesMap[studentId]) {
          gradesMap[studentId] = {};
        }
        gradesMap[studentId][subjectCode] = finalGrade;
      }
    }
  });
  
  return gradesMap;
}

function getActivitiesDataAllYears(grade, classNo, academicYear) {
  var activitiesMap = {};
  var currentGradeNumber = parseInt(grade.replace('ป.', ''));

  // ดึงข้อมูลจากชีตกิจกรรมแต่ละปี (แผน B: ชีตแยกตามปี)
  for (var g = 1; g <= currentGradeNumber; g++) {
    var yearOffset = currentGradeNumber - g;
    var targetYear = String(parseInt(academicYear) - yearOffset);

    var sheet = S_getYearlySheet('การประเมินกิจกรรมพัฒนาผู้เรียน', targetYear);
    if (!sheet) {
      Logger.log('⚠️ ไม่พบชีตกิจกรรมสำหรับปี ' + targetYear);
      continue;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    var headers = data[0];

    var studentIdCol = headers.indexOf('รหัสนักเรียน');
    var gradeCol = headers.indexOf('ชั้น');
    var classCol = headers.indexOf('ห้อง');
    var activity1Col = headers.indexOf('กิจกรรมแนะแนว');
    var activity2Col = headers.indexOf('ลูกเสือ_เนตรนารี');
    var activity3Col = headers.indexOf('ชุมนุม');
    var activity4Col = headers.indexOf('เพื่อสังคมและสาธารณประโยชน์');

    var targetGrade = 'ป.' + g;
    var prefix = 'year' + g + '_';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentGrade = String(row[gradeCol]);
      var studentClass = String(row[classCol]);

      if (studentGrade === targetGrade && studentClass === classNo) {
        var studentId = String(row[studentIdCol]);
        if (!activitiesMap[studentId]) activitiesMap[studentId] = {};
        activitiesMap[studentId][prefix + 'activity1'] = row[activity1Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][prefix + 'activity2'] = row[activity2Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][prefix + 'activity3'] = row[activity3Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][prefix + 'activity4'] = row[activity4Col] === 'ผ่าน' ? 'ผ' : '';
      }
    }
  }

  return activitiesMap;
}

function buildCSVFromWarehouse(students, gradesData, activitiesData, grade, classNo) {
  const currentGradeNumber = parseInt(grade.replace('ป.', ''));
  
  const allSubjects = [];
  for (let gradeNum = 1; gradeNum <= currentGradeNumber; gradeNum++) {
    const subjects = [
      { code: `ท1${gradeNum}101`, name: `ภาษาไทย ${gradeNum}` },
      { code: `ค1${gradeNum}101`, name: `คณิตศาสตร์ ${gradeNum}` },
      { code: `ว1${gradeNum}101`, name: `วิทยาศาสตร์และเทคโนโลยี ${gradeNum}` },
      { code: `ส1${gradeNum}101`, name: `สังคมศึกษา ศาสนาและวัฒนธรรม ${gradeNum}` },
      { code: `ส1${gradeNum}102`, name: `ประวัติศาสตร์ ${gradeNum}` },
      { code: `พ1${gradeNum}101`, name: `สุขศึกษาและพลศึกษา ${gradeNum}` },
      { code: `ศ1${gradeNum}101`, name: `ศิลปะ ${gradeNum}` },
      { code: `ง1${gradeNum}101`, name: gradeNum <= 2 ? `การงานอาชีพ ${gradeNum}` : `การงานอาชีพและเทคโนโลยี ${gradeNum}` },
      { code: `อ1${gradeNum}101`, name: `ภาษาอังกฤษ ${gradeNum}` },
      { code: `ส1${gradeNum}${230 + gradeNum}`, name: `หน้าที่พลเมือง ${gradeNum}` },
      { code: `ส1${gradeNum}201`, name: `การป้องกันการทุจริต ${gradeNum}` }
    ];
    allSubjects.push(...subjects);
  }
  
  const headers = ['#', 'รหัสนักเรียน', 'ชื่อ-สกุล'];
  
  allSubjects.forEach(subject => {
    headers.push(`${subject.code} ${subject.name}`);
  });
  
  for (let gradeNum = 1; gradeNum <= currentGradeNumber; gradeNum++) {
    headers.push(`ก1${gradeNum}901 แนะแนว ${gradeNum}`);
    headers.push(`ก1${gradeNum}902 ลูกเสือ-เนตรนารี ${gradeNum}`);
    headers.push(`ก1${gradeNum}903 ${gradeNum === 1 ? 'ชุมนุมนันทนาการ' : 'ชุมนุมกีฬา ' + gradeNum}`);
    headers.push(`ก1${gradeNum}904 กิจกรรมเพื่อสังคมและสาธารณประโยชน์`);
  }
  
  const lines = [headers.join(',')];
  
  students.forEach((student, index) => {
    const studentGrades = gradesData[student.id] || {};
    const studentActivities = activitiesData[student.id] || {};
    
    const row = [
      index + 1,
      student.id,
      `"${student.name}"`
    ];
    
    allSubjects.forEach(subject => {
      row.push(studentGrades[subject.code] || '');
    });
    
    for (let gradeNum = 1; gradeNum <= currentGradeNumber; gradeNum++) {
      const prefix = `year${gradeNum}_`;
      row.push(studentActivities[`${prefix}activity1`] || '');
      row.push(studentActivities[`${prefix}activity2`] || '');
      row.push(studentActivities[`${prefix}activity3`] || '');
      row.push(studentActivities[`${prefix}activity4`] || '');
    }
    
    lines.push(row.join(','));
  });
  
  // ✅ ไม่ใส่ BOM ที่นี่ (จะใส่ใน handleExportClassCSV แทน)
  return lines.join('\n');
}