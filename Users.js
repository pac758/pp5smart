// ================================================
// UsersModule.gs — Hotfix (no global re‑declare)
// ใช้เมื่อ Code.gs มีประกาศ SPREADSHEET_ID / USER_SHEET_NAME / SETTINGS_SHEET_NAME และ SS() อยู่แล้ว
// 👉 วางไฟล์นี้แทน UsersModule เวอร์ชันก่อนหน้า (หรือแก้เฉพาะส่วนที่คอมเมนต์บอก)
// ================================================

/**************************************
 * ❌ ลบ/อย่าใส่ซ้ำในไฟล์นี้
 * - SPREADSHEET_ID / USER_SHEET_NAME / SETTINGS_SHEET_NAME (มีแล้วใน Code.gs)
 * - ฟังก์ชัน SS() (มีแล้วใน Code.gs)
 **************************************/

/**************************************
 * � Password Hashing (SHA-256)
 **************************************/
function hashPassword_(password) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return raw.map(function(b) { return ('0' + ((b < 0 ? b + 256 : b)).toString(16)).slice(-2); }).join('');
}

function isHashed_(value) {
  return /^[a-f0-9]{64}$/.test(String(value || '').trim());
}

/**************************************
 * �📦 Helpers (ใช้ของ Code.gs ที่มีอยู่)
 **************************************/
function __normalizeHeaderKey__(s){
  return String(s||'').trim().toLowerCase().replace(/[^a-z0-9]+/g,'');
}

function getHeadersMap_(sheet){
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1,1,1,lastCol).getValues()[0] : [];
  var map = {};
  headers.forEach(function(h, i){
    var raw = String(h||'').trim();
    if (!raw) return;
    var norm = __normalizeHeaderKey__(raw);
    map[raw] = i;
    map[norm] = i;
  });
  return { headers: headers, map: map };
}

function __idx__(map, candidates, fallbackIdx){
  for (var i=0;i<candidates.length;i++) if (candidates[i] in map) return map[candidates[i]];
  return fallbackIdx;
}


/**************************************
 * 👤 User Helpers
 **************************************/
function getUserRecordByUsername_(username){
  var sh = SS().getSheetByName(USER_SHEET_NAME);
  if (!sh) throw new Error('ไม่พบชีต "'+USER_SHEET_NAME+'"');
  var info = getHeadersMap_(sh);
  var map = info.map; var data = sh.getDataRange().getValues();
  var unameIdx = __idx__(map, ['username','user','userid','ชื่อผู้ใช้'], 0);
  for (var i=1;i<data.length;i++){
    var row = data[i];
    var u = String(row[unameIdx]||'').trim();
    if (u && u.toLowerCase() === String(username).toLowerCase()){
      return { rowIndex: i+1, row: row, headers: info.headers, map: map, sheet: sh };
    }
  }
  return null;
}

/**************************************
 * 🔐 Login & Session
 **************************************/
function setLoginSession(username){ PropertiesService.getUserProperties().setProperty('username', String(username)); }
function getLoginSession() {
  try {
    var up = PropertiesService.getUserProperties();
    var username = up.getProperty('username');
    if (!username) return null;
    var rec = getUserRecordByUsername_(username);
    if (!rec) return { username: username, role: 'teacher', displayName: username, firstName: '', lastName: '' };
    var map = rec.map, row = rec.row;
    var roleIdx = __idx__(map, ['role','roles'], 2);
    var fnIdx = __idx__(map, ['firstname','first name','ชื่อ','ชื่อจริง'], 3);
    var lnIdx = __idx__(map, ['lastname','last name','นามสกุล'], 4);
    var fn = String(row[fnIdx]||'').trim();
    var ln = String(row[lnIdx]||'').trim();
    return { username: username, role: String(row[roleIdx]||'teacher'), displayName: (fn+' '+ln).trim()||username, firstName: fn, lastName: ln };
  } catch(e) { Logger.log('getLoginSession error: '+e.message); return null; }
}
function getCurrentUser(){
  const up = PropertiesService.getUserProperties();
  const username = up.getProperty('username');
  let email = ''; try { email = Session.getActiveUser().getEmail(); } catch(e) {}
  return { success: !!username, username: username || null, email: email || null };
}

function verifyLogin(username, password){
  try{
    if (!username || !password)
      return { success:false, message:'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };

    var rec = getUserRecordByUsername_(username);
    if (!rec)
      return { success:false, message:'ไม่พบชื่อผู้ใช้นี้ในระบบ' };

    var map = rec.map, row = rec.row;
    var pwIdx   = __idx__(map, ['password','pass','pwd'], 1);
    var roleIdx = __idx__(map, ['role','roles'], 2);
    var fnIdx   = __idx__(map, ['firstname','first name','ชื่อ','ชื่อจริง'], 3);
    var lnIdx   = __idx__(map, ['lastname','last name','นามสกุล'], 4);

    // เทียบรหัสผ่าน: รองรับทั้ง hash (SHA-256) และ plain text เดิม
    var pwCell = String(row[pwIdx] == null ? '' : row[pwIdx]).trim();
    var pwIn   = String(password).trim();
    var matched = false;
    if (isHashed_(pwCell)) {
      matched = (pwCell === hashPassword_(pwIn));
    } else {
      matched = (pwCell === pwIn);
      // auto-migrate: อัปเกรดเป็น hash อัตโนมัติเมื่อ login สำเร็จ
      if (matched && pwIn) {
        try { rec.sheet.getRange(rec.rowIndex, pwIdx + 1).setValue(hashPassword_(pwIn)); } catch(e) {}
      }
    }
    if (!matched)
      return { success:false, message:'รหัสผ่านไม่ถูกต้อง' };

    var fn = String(row[fnIdx]||'').trim();
    var ln = String(row[lnIdx]||'').trim();

    setLoginSession(username);

    var authToken = generateAuthToken(String(username));

    return {
      success: true,
      username: String(username),
      role: String(row[roleIdx]||''),
      firstName: fn,
      lastName: ln,
      fullName: (fn+' '+ln).trim() || String(username),
      authToken: authToken
    };
  } catch(e){
    return { success:false, message:'เกิดข้อผิดพลาด: '+e.message };
  }
}


/**************************************
 * 🔑 Change / Reset Password
 **************************************/
function changePassword(username, oldPassword, newPassword){
  try{
    if (!username || !oldPassword || !newPassword) return { success:false, message:'กรุณากรอกข้อมูลให้ครบถ้วน' };
    if (String(newPassword).length < 6) return { success:false, message:'รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร' };
    var check = verifyLogin(username, oldPassword); if (!check.success) return { success:false, message:'รหัสผ่านเก่าไม่ถูกต้อง' };
    var rec = getUserRecordByUsername_(username); if (!rec) return { success:false, message:'ไม่พบผู้ใช้ในระบบ' };
    var pwIdx = __idx__(rec.map, ['password','pass','pwd'], 1);
    rec.sheet.getRange(rec.rowIndex, pwIdx+1).setValue(hashPassword_(newPassword));
    return { success:true, message:'เปลี่ยนรหัสผ่านสำเร็จ' };
  } catch(e){ return { success:false, message:'เกิดข้อผิดพลาด: '+e.message }; }
}

function resetUserPasswordAdmin(payload){
  try{
    var username = payload && payload.username; var newPassword = payload && payload.newPassword;
    if (!username || !newPassword) return { success:false, message:'กรุณากรอกข้อมูลให้ครบถ้วน' };
    if (String(newPassword).length < 6) return { success:false, message:'รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร' };
    var rec = getUserRecordByUsername_(username); if (!rec) return { success:false, message:'ไม่พบชื่อผู้ใช้นี้ในระบบ' };
    var pwIdx = __idx__(rec.map, ['password','pass','pwd'], 1);
    rec.sheet.getRange(rec.rowIndex, pwIdx+1).setValue(hashPassword_(newPassword));
    return { success:true, message:'รีเซ็ตรหัสผ่านสำหรับ '+username+' สำเร็จ' };
  } catch(e){ return { success:false, message:'เกิดข้อผิดพลาด: '+e.message }; }
}

function deleteUserAdmin(payload){
  try{
    var username = payload && payload.username; if (!username) return { success:false, message:'กรุณาระบุชื่อผู้ใช้' };
    var rec = getUserRecordByUsername_(username); if (!rec) return { success:false, message:'ไม่พบชื่อผู้ใช้นี้ในระบบ' };
    rec.sheet.deleteRow(rec.rowIndex);
    return { success:true, message:'ลบผู้ใช้ '+username+' สำเร็จ' };
  } catch(e){ return { success:false, message:'เกิดข้อผิดพลาด: '+e.message }; }
}

/**************************************
 * ✏️ Update User (สำหรับ Admin)
 **************************************/
function updateUserAdmin(payload) {
  try {
    var originalUsername = payload && payload.originalUsername;
    var newUsername = payload && payload.username;
    var fullName = payload && payload.fullName;
    var role = payload && payload.role;

    // Validate inputs
    if (!originalUsername) {
      return { success: false, message: 'ไม่พบข้อมูลผู้ใช้เดิม' };
    }
    if (!newUsername || String(newUsername).trim().length < 3) {
      return { success: false, message: 'ชื่อผู้ใช้ต้องมีความยาวอย่างน้อย 3 ตัวอักษร' };
    }

    // Find existing user
    var rec = getUserRecordByUsername_(originalUsername);
    if (!rec) {
      return { success: false, message: 'ไม่พบผู้ใช้ "' + originalUsername + '" ในระบบ' };
    }

    // Check if new username already exists (if changed)
    if (newUsername.toLowerCase() !== originalUsername.toLowerCase()) {
      var existingUser = getUserRecordByUsername_(newUsername);
      if (existingUser) {
        return { success: false, message: 'ชื่อผู้ใช้ "' + newUsername + '" มีอยู่ในระบบแล้ว' };
      }
    }

    var map = rec.map;
    var sheet = rec.sheet;
    var rowIndex = rec.rowIndex;

    // Get column indices
    var unameIdx = __idx__(map, ['username', 'user', 'userid', 'ชื่อผู้ใช้'], 0);
    var roleIdx = __idx__(map, ['role', 'roles'], 2);
    var fnIdx = __idx__(map, ['firstname', 'first name', 'ชื่อ', 'ชื่อจริง'], 3);
    var lnIdx = __idx__(map, ['lastname', 'last name', 'นามสกุล'], 4);

    // แยก fullName เป็น firstName และ lastName
    var nameParts = String(fullName || '').trim().split(/\s+/);
    var firstName = nameParts[0] || '';
    var lastName = nameParts.slice(1).join(' ') || '';

    // Update values (column index + 1 for getRange)
    sheet.getRange(rowIndex, unameIdx + 1).setValue(newUsername.trim());
    sheet.getRange(rowIndex, roleIdx + 1).setValue(role || 'teacher');
    sheet.getRange(rowIndex, fnIdx + 1).setValue(firstName);
    sheet.getRange(rowIndex, lnIdx + 1).setValue(lastName);

    return { 
      success: true, 
      message: 'แก้ไขข้อมูลผู้ใช้ "' + newUsername + '" สำเร็จ' 
    };

  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

// Alias function
function updateUser(payload) { return updateUserAdmin(payload); }
/**************************************
 * ➕ Add New User (สำหรับ Admin)
 **************************************/
function addNewUser(payload) {
  try {
    var username = payload && payload.username;
    var password = payload && payload.password;
    var firstName = payload && payload.firstName;
    var lastName = payload && payload.lastName;
    var role = payload && payload.role;

    // Validate inputs
    if (!username || !password || !firstName || !lastName) {
      return { success: false, message: 'กรุณากรอกข้อมูลให้ครบถ้วน' };
    }

    // Check if username already exists
    var existing = getUserRecordByUsername_(username);
    if (existing) {
      return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่ในระบบแล้ว' };
    }

    // Get the users sheet
    var sh = SS().getSheetByName(USER_SHEET_NAME);
    if (!sh) throw new Error('ไม่พบชีต "' + USER_SHEET_NAME + '"');

    // Get headers to find correct columns
    var info = getHeadersMap_(sh);
    var map = info.map;
    var headers = info.headers;

    // Find column indices
    var unameIdx = __idx__(map, ['username', 'user', 'userid', 'ชื่อผู้ใช้'], 0);
    var pwIdx = __idx__(map, ['password', 'pass', 'pwd'], 1);
    var roleIdx = __idx__(map, ['role', 'roles'], 2);
    var fnIdx = __idx__(map, ['firstname', 'first name', 'ชื่อ', 'ชื่อจริง'], 3);
    var lnIdx = __idx__(map, ['lastname', 'last name', 'นามสกุล'], 4);

    // Create new row array with correct length
    var newRow = [];
    for (var i = 0; i < headers.length; i++) {
      newRow.push('');
    }

    // Set values in correct positions
    newRow[unameIdx] = username;
    newRow[pwIdx] = password;
    newRow[roleIdx] = role || 'teacher';
    newRow[fnIdx] = firstName;
    newRow[lnIdx] = lastName;

    // Append the new row
    sh.appendRow(newRow);

    return { 
      success: true, 
      message: 'เพิ่มผู้ใช้ "' + username + '" สำเร็จแล้ว' 
    };

  } catch (e) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.message };
  }
}

/**************************************
 * 🔁 Forgot Password (Email)
 **************************************/
function requestPasswordReset(params){
  try{
    var username = params && params.username; if (!username) return { success:false, message:'กรุณากรอกชื่อผู้ใช้' };
    var rec = getUserRecordByUsername_(username); if (!rec) return { success:false, message:'ไม่พบชื่อผู้ใช้นี้ในระบบ' };
    var map = rec.map, row = rec.row;
    var emailIdx = __idx__(map, ['email','อีเมล','อีเมล์'], 5);
    var pwIdx    = __idx__(map, ['password','pass','pwd'], 1);
    var fnIdx    = __idx__(map, ['firstname','first name','ชื่อ','ชื่อจริง'], 3);
    var lnIdx    = __idx__(map, ['lastname','last name','นามสกุล'], 4);
    var email = String(row[emailIdx]||'').trim(); if (!email || !isValidEmail(email)) return { success:false, message:'ไม่พบอีเมลในระบบ กรุณาติดต่อผู้ดูแลระบบ' };
    var temp = generateTempPassword(); rec.sheet.getRange(rec.rowIndex, pwIdx+1).setValue(temp);
    var fullName = (String(row[fnIdx]||'').trim() + ' ' + String(row[lnIdx]||'').trim()).trim() || username;
    var sent = sendPasswordResetEmail(email, fullName, username, temp);
    if (!sent) return { success:false, message:'ไม่สามารถส่งอีเมลได้ กรุณาลองใหม่อีกครั้งหรือติดต่อผู้ดูแลระบบ' };
    return { success:true, message:'✅ ส่งรหัสผ่านชั่วคราวไปที่ '+maskEmail(email)+' แล้ว' };
  } catch(e){ return { success:false, message:'เกิดข้อผิดพลาดในระบบ: '+e.message }; }
}

function generateTempPassword(){
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var pw = ''; for (var i=0;i<8;i++) pw += chars.charAt(Math.floor(Math.random()*chars.length));
  return pw;
}

function sendPasswordResetEmail(email, fullName, username, tempPassword){
  try{
    var settings = {}; try { settings = getGlobalSettings(); } catch(e){}
    var schoolName = settings['ชื่อโรงเรียน'] || settings['schoolName'] || 'โรงเรียน';
    var subject = '🔐 รหัสผ่านชั่วคราว - ระบบ ปพ.5 ออนไลน์ '+schoolName;
    var plainBody = 'รีเซ็ตรหัสผ่าน - ระบบ ปพ.5 ออนไลน์\n\n'+
      'เรียน คุณ'+fullName+'\n\n'+
      'คุณได้ทำการขอรีเซ็ตรหัสผ่านสำหรับบัญชี: '+username+'\n\n'+
      'รหัสผ่านชั่วคราวของคุณคือ: '+tempPassword+'\n\n'+
      '⚠️ โปรดเปลี่ยนรหัสผ่านทันทีหลังเข้าสู่ระบบ';
    var htmlBody = '<p>เรียน คุณ'+fullName+'</p>'+
      '<p>คุณได้ทำการขอรีเซ็ตรหัสผ่านสำหรับบัญชี: <b>'+username+'</b></p>'+
      '<p><b>รหัสผ่านชั่วคราว: '+tempPassword+'</b></p>'+
      '<p>⚠️ โปรดเปลี่ยนรหัสผ่านทันทีหลังเข้าสู่ระบบ</p>'+
      '<p>'+schoolName+'</p>';
    MailApp.sendEmail({ to: email, subject: subject, body: plainBody, htmlBody: htmlBody, name: 'ระบบ ปพ.5 '+schoolName });
    return true;
  } catch(e){ Logger.log('sendPasswordResetEmail error: '+e); return false; }
}

function isValidEmail(email){ var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; return re.test(String(email||'')); }
function maskEmail(email){ if (!email || email.indexOf('@')<0) return email; var p=email.split('@'), l=p[0], d=p[1]; return l.length<=2? l[0]+'***@'+d : l[0]+'***'+l[l.length-1]+'@'+d; }

/**************************************
 * 🧹 Admin Utilities
 **************************************/
function getAllUsers(){
  try{
    var sh = SS().getSheetByName(USER_SHEET_NAME);
    if (!sh) throw new Error('ไม่พบชีต "'+USER_SHEET_NAME+'"');
    var info = getHeadersMap_(sh); var map = info.map; var data = sh.getDataRange().getValues();
    var users = [];
    for (var i=1;i<data.length;i++){
      var r = data[i];
      var username = String(r[__idx__(map,['username','user','userid','ชื่อผู้ใช้'],0)]||'').trim(); if (!username) continue;
      var fn = String(r[__idx__(map,['firstname','first name','ชื่อ','ชื่อจริง'],3)]||'').trim();
      var ln = String(r[__idx__(map,['lastname','last name','นามสกุล'],4)]||'').trim();
      var role = String(r[__idx__(map,['role','roles'],2)]||'user').trim();
      users.push({ username:username, role:role, firstName:fn, lastName:ln, fullName:(fn+' '+ln).trim()||username });
    }
    return { success:true, users:users, count:users.length };
  } catch(e){ return { success:false, error:e.message, users:[] }; }
}
function resetUserPassword(payload) { return resetUserPasswordAdmin(payload); }
function deleteUser(payload) { return deleteUserAdmin(payload); }

