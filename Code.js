// ============================================================

// ⚙️ CORE CONFIGURATION

// ============================================================



const USER_SHEET_NAME = "users";

const APP_VERSION = 'v2.3.0_20250115_TOKEN_AUTH';



// ============================================================

// 🏫 MULTI-TENANT: ดึง ID จาก ScriptProperties (ไม่ hardcode)

// ============================================================

function getSpreadsheetId_() {

  const props = PropertiesService.getScriptProperties();

  return props.getProperty('SPREADSHEET_ID') || null;

}



function getPdfFolderId_() {

  const props = PropertiesService.getScriptProperties();

  return props.getProperty('PDF_OUTPUT_FOLDER_ID') || null;

}



function isSetupComplete_() {

  var props = PropertiesService.getScriptProperties();

  var ssId = props.getProperty('SPREADSHEET_ID');

  if (!ssId) return false;



  var currentScriptId = ScriptApp.getScriptId();

  var storedScriptId = props.getProperty('SCRIPT_ID');



  // SCRIPT_ID ตรงกัน → project เดิม → OK

  if (storedScriptId === currentScriptId) return true;



  // SCRIPT_ID มีแต่ไม่ตรง → copied project → ล้าง → Setup Wizard

  if (storedScriptId && storedScriptId !== currentScriptId) {

    props.deleteAllProperties();

    return false;

  }



  // SCRIPT_ID ไม่มี → ตรวจจาก Spreadsheet

  try {

    var ss = SpreadsheetApp.openById(ssId);

    var sheet = ss.getSheetByName('global_settings');

    if (!sheet) { props.deleteAllProperties(); return false; }



    var data = sheet.getDataRange().getValues();

    var markerRow = null;

    for (var i = 0; i < data.length; i++) {

      if (data[i][0] === 'installed_script_id') { markerRow = data[i]; break; }

    }



    if (markerRow && markerRow[1]) {

      if (String(markerRow[1]) === currentScriptId) {

        props.setProperty('SCRIPT_ID', currentScriptId);

        return true;

      } else {

        props.deleteAllProperties();

        return false;

      }

    } else {

      props.setProperty('SCRIPT_ID', currentScriptId);

      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues([['installed_script_id', currentScriptId]]);

      return true;

    }

  } catch (e) {

    Logger.log('isSetupComplete_ error: ' + e.message);

    props.deleteAllProperties();

    return false;

  }

}



const TOKEN_EXPIRY_HOURS = 24;



const monthNames = [

  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',

  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'

];



// ============================================================

// 🛠️ UTILITY FUNCTIONS (ใช้ภายใน Code.gs)

// ============================================================



function SS() {

  const id = getSpreadsheetId_();

  if (!id) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน Setup Wizard ก่อน');

  return SpreadsheetApp.openById(id);

}

function getSchoolName_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('school_name');
    if (cached) return cached;
    var ss = SS();
    var sheet = ss.getSheetByName('global_settings');
    if (!sheet) return '';
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'ชื่อโรงเรียน') {
        var name = String(data[i][1] || '').trim();
        if (name) cache.put('school_name', name, 600);
        return name;
      }
    }
    return '';
  } catch (e) { return ''; }
}



function escapeHtml(text) {

  if (!text) return '';

  return String(text)

    .replace(/&/g, '&amp;')

    .replace(/</g, '&lt;')

    .replace(/>/g, '&gt;')

    .replace(/"/g, '&quot;')

    .replace(/'/g, '&#039;');

}



function findColumn(headers, possibleNames) {

  for (const name of possibleNames) {

    const index = headers.findIndex(h => 

      String(h || '').toLowerCase().trim() === name.toLowerCase()

    );

    if (index !== -1) return index;

  }

  return -1;

}



function createEmptySummary() {

  return {

    summary: {

      total: 0, male: 0, female: 0,

      lastUpdated: new Date().toISOString()

    }

  };

}



function getDefaultSettings() {

  return S_getDefaultSettings();

}



function checkVersion() {

  return {

    version: APP_VERSION,

    timestamp: new Date().toISOString(),

    spreadsheetId: getSpreadsheetId_() || '(ยังไม่ได้ตั้งค่า)',

    setupComplete: isSetupComplete_(),

    status: 'OK'

  };

}





function getOutputFolder_() {

  const FOLDER_NAME = 'รายงานผลการเรียน_PDF';

  try {

    const folders = DriveApp.getFoldersByName(FOLDER_NAME);

    if (folders.hasNext()) return folders.next();

    return DriveApp.createFolder(FOLDER_NAME);

  } catch (e) {

    Logger.log('Error getting output folder: ' + e.message);

    return DriveApp.getRootFolder();

  }

}



// ============================================================

// 🔐 SECURE TOKEN SECRET MANAGEMENT

// ============================================================



function getTokenSecret_() {

  const props = PropertiesService.getScriptProperties();

  let secret = props.getProperty('TOKEN_SECRET');

  if (!secret) {

    secret = Utilities.getUuid() + '-' + Utilities.getUuid();

    props.setProperty('TOKEN_SECRET', secret);

    Logger.log('🔐 New TOKEN_SECRET generated and stored');

  }

  return secret;

}



function resetTokenSecret() {

  const props = PropertiesService.getScriptProperties();

  const newSecret = Utilities.getUuid() + '-' + Utilities.getUuid();

  props.setProperty('TOKEN_SECRET', newSecret);

  Logger.log('🔐 TOKEN_SECRET has been reset.');

  return { success: true, message: 'Token secret reset. All users must login again.' };

}



// ============================================================

// 🔐 TOKEN-BASED AUTHENTICATION

// ============================================================



function generateAuthToken(username) {

  try {

    const secret = getTokenSecret_();

    const timestamp = Date.now();

    const expiry = timestamp + (TOKEN_EXPIRY_HOURS * 60 * 60 * 1000);

    const data = username + '|' + expiry + '|' + timestamp;

    

    const signature = Utilities.base64Encode(

      Utilities.computeHmacSha256Signature(data, secret)

    ).replace(/[+/=]/g, c => ({ '+': '-', '/': '_', '=': '' }[c]));

    

    const encodedData = Utilities.base64Encode(data).replace(/[+/=]/g, c => ({ '+': '-', '/': '_', '=': '' }[c]));

    const token = encodedData + '.' + signature.substring(0, 32);

    

    Logger.log('✅ Token generated for: ' + username);

    return token;

  } catch (e) {

    Logger.log('❌ generateAuthToken error: ' + e.message);

    return null;

  }

}



function verifyAuthToken(token) {

  try {

    if (!token || typeof token !== 'string') {

      return { valid: false, reason: 'NO_TOKEN' };

    }

    

    const parts = token.split('.');

    if (parts.length !== 2) {

      return { valid: false, reason: 'INVALID_FORMAT' };

    }

    

    const secret = getTokenSecret_();

    const encodedData = parts[0].replace(/-/g, '+').replace(/_/g, '/');

    const data = Utilities.newBlob(Utilities.base64Decode(encodedData)).getDataAsString();

    const providedSig = parts[1];

    

    const expectedSig = Utilities.base64Encode(

      Utilities.computeHmacSha256Signature(data, secret)

    ).replace(/[+/=]/g, c => ({ '+': '-', '/': '_', '=': '' }[c])).substring(0, 32);

    

    if (providedSig !== expectedSig) {

      return { valid: false, reason: 'INVALID_SIGNATURE' };

    }

    

    const dataParts = data.split('|');

    if (dataParts.length < 2) {

      return { valid: false, reason: 'INVALID_DATA' };

    }

    

    const username = dataParts[0];

    const expiry = parseInt(dataParts[1]);

    

    if (Date.now() > expiry) {

      return { valid: false, reason: 'EXPIRED', username: username };

    }

    

    return { valid: true, username: username, expiry: expiry };

  } catch (e) {

    Logger.log('❌ verifyAuthToken error: ' + e.message);

    return { valid: false, reason: 'ERROR', message: e.message };

  }

}



function validateToken(token) {

  const result = verifyAuthToken(token);

  return { valid: result.valid, username: result.username || null, reason: result.reason || null };

}



function refreshToken(oldToken) {

  const result = verifyAuthToken(oldToken);

  if (result.valid) {

    return { success: true, token: generateAuthToken(result.username), username: result.username };

  }

  return { success: false, reason: result.reason, message: 'Token ไม่ถูกต้องหรือหมดอายุ กรุณาเข้าสู่ระบบใหม่' };

}



// ============================================================

// 🛡️ PROTECTED API WRAPPER

// ============================================================



function requireAuth_(token, callback) {

  const authResult = verifyAuthToken(token);

  if (!authResult.valid) {

    return { success: false, error: 'UNAUTHORIZED', reason: authResult.reason, message: 'กรุณาเข้าสู่ระบบใหม่' };

  }

  return callback(authResult.username);

}



function getProtectedDashboardData(token) {

  return requireAuth_(token, (username) => {

    return { success: true, username: username, data: getCachedDashboardSummary() };

  });

}



function getProtectedStudentData(token, studentId) {

  return requireAuth_(token, (username) => {

    return { success: true, username: username, studentId: studentId, data: {} };

  });

}



function saveProtectedData(token, dataType, data) {

  return requireAuth_(token, (username) => {

    Logger.log('📝 User ' + username + ' saving ' + dataType);

    return { success: true, username: username, message: 'บันทึกข้อมูลสำเร็จ' };

  });

}



// ============================================================

// 🌐 WEB APP ENTRY POINTS

// ============================================================



function doGet(e) {

  try {

    const params = e && e.parameter ? e.parameter : {};

    const page = (params.page || '').toLowerCase();



    // 📦 SETUP WIZARD: ?page=setup_wizard → แสดงหน้าติดตั้ง (เฉพาะ project ที่ยังไม่ติดตั้ง/copied)

    if (page === 'setup_wizard') {

      var wizProps = PropertiesService.getScriptProperties();

      var wizScriptId = wizProps.getProperty('SCRIPT_ID');

      if (wizScriptId && wizScriptId === ScriptApp.getScriptId()) {

        // project เดิมที่ติดตั้งแล้ว → ไม่อนุญาตให้ติดตั้งซ้ำ

        return serveLoginPage();

      }

      return serveSetupWizard();

    }



    // 🏫 MULTI-TENANT: ถ้ายังไม่ได้ติดตั้ง → Setup Wizard
    if (!isSetupComplete_()) {
      return serveSetupWizard();
    }

    // 🛠️ DEV BYPASS: ?dev=1 หรือ test deployment (/dev) → bypass login
    if (params.dev === '1' || isTestDeployment_()) {
      return serveDevBypass();
    }

    if (page === 'dashboard' || page === 'app') return serveMainApp();

    if (page === 'parent_portal') return serveParentPortal_();

    if (page === 'forgot_password') return serveForgotPassword();



    // หน้าอื่นๆ → login

    return serveLoginPage();

  } catch (error) {

    Logger.log('❌ doGet Error: ' + error.message);

    return createErrorPage(error);

  }

}



function serveSetupWizard() {

  try {

    const output = HtmlService.createHtmlOutputFromFile('setup_wizard')

      .setTitle('ติดตั้งระบบ ปพ.5 — Setup Wizard')

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    addNoCacheMeta_(output);

    return output;

  } catch (error) {

    Logger.log('❌ serveSetupWizard Error: ' + error.message);

    return createErrorPage(error);

  }

}



function isTestDeployment_() {

  try {

    var url = ScriptApp.getService().getUrl();

    var isDev = url && url.indexOf('/dev') !== -1;

    Logger.log('🔍 isTestDeployment_: url=' + url + ', isDev=' + isDev);

    return isDev;

  } catch (e) {

    Logger.log('⚠️ isTestDeployment_ error: ' + e.message);

    return false;

  }

}



function getDeploymentUrl() {

  try { return ScriptApp.getService().getUrl(); }

  catch (e) { return ''; }

}



function doPost(e) {

  const out = ContentService.createTextOutput().setMimeType(ContentService.MimeType.JSON);

  try {

    if (e && e.parameter && e.parameter.action === 'login') {

      const result = verifyLogin(e.parameter.username, e.parameter.password);

      if (result && result.success) {

        const token = generateAuthToken(result.username);

        return out.setContent(JSON.stringify({

          success: true, message: 'เข้าสู่ระบบสำเร็จ',

          username: result.username, role: result.role,

          firstName: result.firstName, lastName: result.lastName,

          fullName: result.fullName, className: result.className || "",

          appVersion: APP_VERSION, authToken: token

        }));

      } else {

        return out.setContent(JSON.stringify({

          success: false, message: (result && result.message) || 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'

        }));

      }

    }

    return out.setContent(JSON.stringify({ success: false, message: 'ไม่รองรับ action นี้' }));

  } catch (err) {

    return out.setContent(JSON.stringify({ success: false, message: String(err) }));

  }

}



// ============================================================

// 📄 PAGE SERVING FUNCTIONS

// ============================================================



function serveLoginPage() {

  try {

    const output = HtmlService.createHtmlOutputFromFile("login")

      .setTitle(getSchoolName_() ? 'เข้าสู่ระบบ - ' + getSchoolName_() : 'เข้าสู่ระบบ - ระบบ ปพ.5')

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    addNoCacheMeta_(output);

    return output;

  } catch (error) {

    Logger.log('❌ serveLoginPage Error: ' + error.message);

    throw error;

  }

}



function serveParentPortal_() {
  try {
    var output = HtmlService.createHtmlOutputFromFile('parent_portal')
      .setTitle('\u0e23\u0e30\u0e1a\u0e1a\u0e1c\u0e39\u0e49\u0e1b\u0e01\u0e04\u0e23\u0e2d\u0e07 - PP5 Smart')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    addNoCacheMeta_(output);
    return output;
  } catch (e) {
    Logger.log('\u274c serveParentPortal_ Error: ' + e.message);
    return serveLoginPage();
  }
}

function serveDevBypass() {

  try {

    // ✅ ตั้ง server-side session ด้วย เพื่อให้ getCurrentUser() ทำงานได้
    setLoginSession('dev_admin');

    var output = HtmlService.createHtmlOutputFromFile("spa_main_menu")

      .setTitle("ระบบจัดการนักเรียน ปพ.5 [DEV]")

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    // Inject dev session script ต่อท้าย HTML

    output.append('<script>localStorage.setItem("loggedIn","true");localStorage.setItem("username","dev_admin");localStorage.setItem("role","admin");localStorage.setItem("displayName","Dev Admin");localStorage.setItem("firstName","Developer");localStorage.setItem("lastName","Admin");localStorage.setItem("className","ทดสอบระบบ");localStorage.setItem("authToken","dev_bypass_token");console.log("🛠️ DEV BYPASS: session set from server");</script>');

    addNoCacheMeta_(output);

    Logger.log('✅ serveDevBypass: served successfully');

    return output;

  } catch (error) {

    Logger.log('❌ serveDevBypass Error: ' + error.message);

    return serveMainApp();

  }

}



function serveMainApp() {

  try {

    var sn = getSchoolName_();

    const output = HtmlService.createHtmlOutputFromFile("spa_main_menu")

      .setTitle(sn ? 'ระบบ ปพ.5 — ' + sn : 'ระบบจัดการนักเรียน ปพ.5')

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    addNoCacheMeta_(output);

    return output;

  } catch (error) {

    Logger.log('❌ serveMainApp Error: ' + error.message);

    throw error;

  }

}



function escapeJs_(str) {

  return String(str || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/'/g, "\\'").replace(/</g, '\\x3c').replace(/>/g, '\\x3e');

}



function serveMainAppWithSession_(session) {

  try {

    var output = HtmlService.createHtmlOutputFromFile("spa_main_menu");

    var sessionScript = '<script>' +

      'localStorage.setItem("loggedIn","true");' +

      'localStorage.setItem("username","' + escapeJs_(session.username) + '");' +

      'localStorage.setItem("role","' + escapeJs_(session.role || 'teacher') + '");' +

      'localStorage.setItem("displayName","' + escapeJs_(session.displayName || session.username) + '");' +

      'localStorage.setItem("firstName","' + escapeJs_(session.firstName) + '");' +

      'localStorage.setItem("lastName","' + escapeJs_(session.lastName) + '");' +

      'Logger.log("✅ Session injected from server for: ' + escapeJs_(session.username) + '");' +

      '</script>';

    output.append(sessionScript);

    var sn2 = getSchoolName_();

    output.setTitle(sn2 ? 'ระบบ ปพ.5 — ' + sn2 : 'ระบบจัดการนักเรียน ปพ.5')

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    addNoCacheMeta_(output);

    return output;

  } catch (error) {

    Logger.log('❌ serveMainAppWithSession_ Error: ' + error.message);

    return serveMainApp();

  }

}



function serveForgotPassword() {

  try {

    const output = HtmlService.createHtmlOutputFromFile('forgot_password')

      .setTitle('ลืมรหัสผ่าน - ระบบ ปพ.5')

      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    addNoCacheMeta_(output);

    return output;

  } catch (error) {

    Logger.log('❌ serveForgotPassword Error: ' + error.message);

    throw error;

  }

}



function createErrorPage(error) {

  const html = `

<!DOCTYPE html>

<html lang="th">

<head>

  <meta charset="UTF-8">

  <meta name="viewport" content="width=device-width, initial-scale=1.0">

  <title>System Error</title>

  <style>

    body { font-family: -apple-system, BlinkMacSystemFont, 'Sarabun', sans-serif; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; }

    .error { background: white; padding: 40px; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); max-width: 500px; text-align: center; }

    h1 { color: #e74c3c; margin: 0 0 15px; }

    pre { background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: left; overflow-x: auto; font-size: 12px; }

    button { margin-top: 20px; padding: 12px 24px; background: #667eea; color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 16px; }

  </style>

</head>

<body>

  <div class="error">

    <h1>⚠️ System Error</h1>

    <p>เกิดข้อผิดพลาดในการโหลดหน้าเว็บ</p>

    <pre>${escapeHtml(error.message)}</pre>

    <button onclick="window.location.reload()">🔄 Reload</button>

  </div>

</body>

</html>`;

  return HtmlService.createHtmlOutput(html)

    .setTitle('System Error')

    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}



function addNoCacheMeta_(output) {

  try {

    output.addMetaTag('viewport', 'width=device-width, initial-scale=1.0')

          .addMetaTag('x-app-version', APP_VERSION)

          .addMetaTag('Cache-Control', 'no-cache, no-store, must-revalidate')

          .addMetaTag('Pragma', 'no-cache')

          .addMetaTag('Expires', '0');

  } catch (e) {}

  return output;

}



// ============================================================

// 🗂️ PAGE CONTENT LOADER

// ============================================================



function getPageContent(pageName, token) {

  const protectedPages = ['dashboard', 'students', 'scores', 'reports', 'settings'];

  

  if (protectedPages.includes(pageName)) {

    const authResult = verifyAuthToken(token);

    if (!authResult.valid) {

      return `

        <div class="alert alert-danger p-4 text-center">

          <h4><i class="bi bi-shield-exclamation"></i> Session หมดอายุ</h4>

          <p>กรุณาเข้าสู่ระบบใหม่</p>

          <button class="btn btn-primary" onclick="logout()">

            <i class="bi bi-box-arrow-right"></i> เข้าสู่ระบบใหม่

          </button>

        </div>`;

    }

  }

  

  try {

    const content = HtmlService.createHtmlOutputFromFile(pageName).getContent();

    return `<!-- Version: ${APP_VERSION} | Page: ${pageName} | Time: ${new Date().toISOString()} -->\n` + content;

  } catch (e) {

    return `

      <div class="alert alert-warning p-3">

        <h4><i class="bi bi-exclamation-triangle"></i> ไม่พบหน้า: ${escapeHtml(pageName)}</h4>

        <p>หน้านี้ยังไม่ได้สร้างหรือถูกลบออกจากระบบ</p>

        <button class="btn btn-primary" onclick="loadPage('dashboard')">

          <i class="bi bi-house"></i> กลับหน้าหลัก

        </button>

      </div>`;

  }

}



function include(filename) {

  return HtmlService.createHtmlOutputFromFile(filename).getContent();

}



function injectIncludes_(html) {

  var includes = '';

  try { includes += HtmlService.createHtmlOutputFromFile('grade_class_helper').getContent(); } catch(e) { Logger.log('⚠️ include grade_class_helper failed: ' + e.message); }

  return html.replace('<!-- INCLUDES_PLACEHOLDER -->', includes);

}



// ============================================================

// 📊 DASHBOARD FUNCTIONS

// ============================================================





function getGlobalSettings() {

  const args = Array.prototype.slice.call(arguments);

  const useCache = args.length ? Boolean(args[0]) : true;

  return S_getGlobalSettings(useCache);

}



// ============================================================

// � DIAGNOSTIC: ตรวจสอบชีตที่ขาดหายไป

// ============================================================

function checkMissingSheets() {

  const permanentSheets = [

    'global_settings', 'Users', 'Students', 'รายวิชา',

    'Holidays', 'HomeroomTeachers'

  ];

  // ชีตรายปี — ใช้ S_resolveSheetName เพื่อหาชื่อจริง

  var yearlyBases = (typeof S_YEARLY_SHEETS !== 'undefined') ? S_YEARLY_SHEETS : [

    'SCORES_WAREHOUSE', 'การประเมินอ่านคิดเขียน', 'การประเมินคุณลักษณะ',

    'การประเมินกิจกรรมพัฒนาผู้เรียน', 'การประเมินสมรรถนะ',

    'AttendanceLog', 'ความเห็นครู'

  ];

  var requiredSheets = permanentSheets.slice();

  yearlyBases.forEach(function(base) {

    requiredSheets.push((typeof S_resolveSheetName === 'function') ? S_resolveSheetName(base) : base);

  });

  

  const ss = SS();

  const existingSheets = ss.getSheets().map(s => s.getName());

  

  const missing = requiredSheets.filter(name => !existingSheets.includes(name));

  const extra = existingSheets.filter(name => !requiredSheets.includes(name));

  

  const report = {

    spreadsheetName: ss.getName(),

    spreadsheetId: ss.getId(),

    totalSheets: existingSheets.length,

    existingSheets: existingSheets,

    requiredSheets: requiredSheets,

    missingSheets: missing,

    extraSheets: extra,

    status: missing.length === 0 ? '✅ ครบทุกชีต' : `❌ ขาด ${missing.length} ชีต`

  };

  

  Logger.log(JSON.stringify(report, null, 2));

  return report;

}



// ============================================================

// 📚 ACADEMIC FUNCTIONS

// ============================================================

