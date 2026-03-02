/**
 * ============================================================
 * 📤 EXPORT CSV (FIXED UTF-8 BOM VERSION)
 * ============================================================
 */

function handleExportClassCSV(grade, classNo) {
  try {
    const settings = S_getGlobalSettings();
    const academicYear = settings['ปีการศึกษา'] || '2568';
    
    const result = exportClassScoresToCSV(grade, classNo, academicYear);
    
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

function exportClassScoresToCSV(grade, classNo, academicYear) {
  try {
    const ss = SS();
    
    const students = getStudentsInClass(grade, classNo);
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

function getStudentsInClass(grade, classNo) {
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
  
  students.sort((a, b) => String(a.id).localeCompare(String(b.id), undefined, { numeric: true }));
  
  return students;
}

function getGradesFromWarehouseAllYears(grade, classNo, academicYear) {
  const ss = SS();
  const warehouseSheet = ss.getSheetByName('SCORES_WAREHOUSE');
  
  if (!warehouseSheet) {
    Logger.log('⚠️ ไม่พบชีต SCORES_WAREHOUSE');
    return {};
  }
  
  const data = warehouseSheet.getDataRange().getValues();
  const headers = data[0];
  
  const studentIdCol = headers.indexOf('student_id');
  const gradeCol = headers.indexOf('grade');
  const classCol = headers.indexOf('class_no');
  const subjectCodeCol = headers.indexOf('subject_code');
  const finalGradeCol = headers.indexOf('final_grade');
  const yearCol = headers.indexOf('academic_year');
  
  const gradesMap = {};
  
  const currentGradeNumber = parseInt(grade.replace('ป.', ''));
  const yearsToFetch = [];
  
  for (let i = 1; i <= currentGradeNumber; i++) {
    const yearOffset = currentGradeNumber - i;
    const targetYear = String(parseInt(academicYear) - yearOffset);
    yearsToFetch.push({
      year: targetYear,
      grade: `ป.${i}`
    });
  }
  
  Logger.log(`📚 ดึงข้อมูล ${yearsToFetch.length} ปี: ${JSON.stringify(yearsToFetch)}`);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const rowYear = String(row[yearCol]);
    const rowGrade = String(row[gradeCol]);
    const rowClass = String(row[classCol]);
    
    const isRelevant = yearsToFetch.some(y => 
      y.year === rowYear && y.grade === rowGrade && rowClass === classNo
    );
    
    if (isRelevant) {
      const studentId = String(row[studentIdCol]);
      const subjectCode = String(row[subjectCodeCol]);
      const finalGrade = row[finalGradeCol] || '';
      
      if (!gradesMap[studentId]) {
        gradesMap[studentId] = {};
      }
      
      gradesMap[studentId][subjectCode] = finalGrade;
    }
  }
  
  return gradesMap;
}

function getActivitiesDataAllYears(grade, classNo, academicYear) {
  const ss = SS();
  const activitySheet = ss.getSheetByName('การประเมินกิจกรรมพัฒนาผู้เรียน');
  
  if (!activitySheet) {
    Logger.log('⚠️ ไม่พบชีตกิจกรรม');
    return {};
  }
  
  const data = activitySheet.getDataRange().getValues();
  const headers = data[0];
  
  const studentIdCol = headers.indexOf('รหัสนักเรียน');
  const gradeCol = headers.indexOf('ชั้น');
  const classCol = headers.indexOf('ห้อง');
  
  const activity1Col = headers.indexOf('กิจกรรมแนะแนว');
  const activity2Col = headers.indexOf('ลูกเสือ_เนตรนารี');
  const activity3Col = headers.indexOf('ชุมนุม');
  const activity4Col = headers.indexOf('เพื่อสังคมและสาธารณประโยชน์');
  
  const activitiesMap = {};
  
  const currentGradeNumber = parseInt(grade.replace('ป.', ''));
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentId = String(row[studentIdCol]);
    const studentGrade = String(row[gradeCol]);
    const studentClass = String(row[classCol]);
    
    if (studentClass === classNo) {
      const gradeNum = parseInt(studentGrade.replace('ป.', ''));
      
      if (gradeNum <= currentGradeNumber) {
        if (!activitiesMap[studentId]) {
          activitiesMap[studentId] = {};
        }
        
        const prefix = `year${gradeNum}_`;
        activitiesMap[studentId][`${prefix}activity1`] = row[activity1Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][`${prefix}activity2`] = row[activity2Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][`${prefix}activity3`] = row[activity3Col] === 'ผ่าน' ? 'ผ' : '';
        activitiesMap[studentId][`${prefix}activity4`] = row[activity4Col] === 'ผ่าน' ? 'ผ' : '';
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