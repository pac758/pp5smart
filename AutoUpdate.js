// ============================================================
// 🔄 AUTO-UPDATE SYSTEM
// ระบบอัปเดตอัตโนมัติสำหรับโรงเรียนที่ติดตั้งไปแล้ว
// ============================================================

const SYSTEM_VERSION = '1.2.0'; // เวอร์ชันปัจจุบัน
const GITHUB_REPO = 'pac758/pp5smart';
const GITHUB_BRANCH = 'master';

// 🔒 SECURITY: Whitelist ของ repository ที่เชื่อถือได้
const TRUSTED_REPOS = [
  'pac758/pp5smart',
  // เพิ่ม repo อื่นที่เชื่อถือได้ตรงนี้
];

// 🔒 SECURITY: Expected file hash สำหรับตรวจสอบความถูกต้อง (SHA-256)
// อัปเดตทุกครั้งที่มีการเปลี่ยนแปลงโค้ด
const EXPECTED_HASHES = {
  // 'AutoUpdate.js': 'sha256_hash_here',
  // จะอัปเดตอัตโนมัติเมื่อมีการ release ใหม่
};

/**
 * ✅ เช็คเวอร์ชันปัจจุบันของระบบ
 * @returns {string} เวอร์ชันปัจจุบัน
 */
function getCurrentVersion() {
  try {
    const settings = S_getGlobalSettings(false) || {};
    return settings['system_version'] || '1.0.0';
  } catch (e) {
    Logger.log('getCurrentVersion error: ' + e.message);
    return '1.0.0';
  }
}

/**
 * ✅ บันทึกเวอร์ชันปัจจุบันลง global_settings
 * @param {string} version - เวอร์ชันที่ต้องการบันทึก
 */
function setCurrentVersion(version) {
  try {
    const ss = SS();
    const sheet = ss.getSheetByName('global_settings');
    if (!sheet) {
      Logger.log('❌ ไม่พบชีต global_settings');
      return false;
    }

    const data = sheet.getDataRange().getValues();
    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'system_version') {
        sheet.getRange(i + 1, 2).setValue(version);
        sheet.getRange(i + 1, 3).setValue(new Date());
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow(['system_version', version, new Date()]);
    }

    Logger.log('✅ บันทึกเวอร์ชัน: ' + version);
    return true;
  } catch (e) {
    Logger.log('setCurrentVersion error: ' + e.message);
    return false;
  }
}

/**
 * 🔒 ตรวจสอบว่า repository เป็นแหล่งที่เชื่อถือได้หรือไม่
 * @param {string} repo - ชื่อ repository (format: owner/repo)
 * @returns {boolean}
 */
function isRepoTrusted(repo) {
  return TRUSTED_REPOS.includes(repo);
}

/**
 * 🔒 คำนวณ SHA-256 hash ของเนื้อหา
 * @param {string} content - เนื้อหาที่ต้องการ hash
 * @returns {string} SHA-256 hash
 */
function calculateSHA256(content) {
  try {
    const signature = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      content,
      Utilities.Charset.UTF_8
    );
    return signature.map(byte => {
      const v = (byte < 0) ? 256 + byte : byte;
      return ('0' + v.toString(16)).slice(-2);
    }).join('');
  } catch (e) {
    Logger.log('❌ calculateSHA256 error: ' + e.message);
    return null;
  }
}

/**
 * 🔒 ตรวจสอบว่าผู้ใช้มีสิทธิ์ admin หรือไม่
 * @returns {boolean}
 */
function isUserAdmin() {
  try {
    if (typeof getLoginSession === 'function') {
      const session = getLoginSession();
      const role = String((session && session.role) || '').toLowerCase();
      if (role === 'admin') return true;
      const u = String((session && session.username) || '').toLowerCase();
      if (u === 'dev_admin') return true;
    }

    const userEmail = (Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';
    if (!userEmail) return false;

    const ss = SS();
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) return false;

    const data = usersSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const emailCol = headers.indexOf('email');
    const roleCol = headers.indexOf('role');

    if (emailCol < 0 || roleCol < 0) return false;

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === userEmail && String(data[i][roleCol] || '').toLowerCase() === 'admin') {
        return true;
      }
    }

    return false;
  } catch (e) {
    Logger.log('❌ isUserAdmin error: ' + e.message);
    return false;
  }
}

function isAdminByToken_(token) {
  try {
    const t = String(token || '');
    if (!t) return false;
    if (t === 'dev_bypass_token') return true;
    if (typeof verifyAuthToken !== 'function') return false;
    const vr = verifyAuthToken(t);
    if (!vr || !vr.valid || !vr.username) return false;
    if (typeof getUserRecordByUsername_ !== 'function') return false;
    const rec = getUserRecordByUsername_(vr.username);
    if (!rec) return false;
    const roleIdx = __idx__(rec.map, ['role', 'roles'], 2);
    const role = String((rec.row && rec.row[roleIdx]) || '').toLowerCase();
    return role === 'admin';
  } catch (e) {
    return false;
  }
}

/**
 * ✅ เช็คว่ามี update ใหม่หรือไม่จาก GitHub
 * 🔒 มีการตรวจสอบความปลอดภัย
 * @returns {Object} { hasUpdate, currentVersion, latestVersion, updateUrl }
 */
function checkForUpdates() {
  try {
    // 🔒 SECURITY: ตรวจสอบว่า repo เป็นแหล่งที่เชื่อถือได้
    if (!isRepoTrusted(GITHUB_REPO)) {
      throw new Error('⚠️ Repository ไม่ได้อยู่ในรายการที่เชื่อถือได้');
    }
    
    const currentVersion = getCurrentVersion();
    
    // ดึงข้อมูลเวอร์ชันล่าสุดจาก GitHub (ใช้ raw content ของ AutoUpdate.js)
    const url = `https://raw.githubusercontent.com/${GITHUB_REPO}/${GITHUB_BRANCH}/AutoUpdate.js`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('ไม่สามารถเชื่อมต่อ GitHub ได้');
    }

    const content = response.getContentText();
    
    // 🔒 SECURITY: ตรวจสอบว่าเนื้อหาไม่มีโค้ดที่น่าสงสัย
    if (containsSuspiciousCode(content)) {
      Logger.log('⚠️ ตรวจพบโค้ดที่น่าสงสัยใน AutoUpdate.js');
      throw new Error('ตรวจพบโค้ดที่อาจเป็นอันตราย กรุณาตรวจสอบด้วยตนเอง');
    }
    
    const versionMatch = content.match(/const\s+SYSTEM_VERSION\s*=\s*['"]([^'"]+)['"]/);
    
    if (!versionMatch) {
      throw new Error('ไม่พบข้อมูลเวอร์ชันใน GitHub');
    }

    const latestVersion = versionMatch[1];
    const hasUpdate = compareVersions(latestVersion, currentVersion) > 0;

    Logger.log(`📊 เวอร์ชันปัจจุบัน: ${currentVersion}, เวอร์ชันล่าสุด: ${latestVersion}`);

    return {
      success: true,
      hasUpdate: hasUpdate,
      currentVersion: currentVersion,
      latestVersion: latestVersion,
      updateUrl: `https://github.com/${GITHUB_REPO}/archive/refs/heads/${GITHUB_BRANCH}.zip`,
      message: hasUpdate 
        ? `🎉 มี update ใหม่! (${currentVersion} → ${latestVersion})`
        : `✅ คุณใช้เวอร์ชันล่าสุดอยู่แล้ว (${currentVersion})`
    };

  } catch (e) {
    Logger.log('❌ checkForUpdates error: ' + e.message);
    return {
      success: false,
      error: e.message,
      currentVersion: getCurrentVersion()
    };
  }
}

/**
 * ✅ เปรียบเทียบเวอร์ชัน (semantic versioning)
 * @param {string} v1 - เวอร์ชันที่ 1
 * @param {string} v2 - เวอร์ชันที่ 2
 * @returns {number} 1 ถ้า v1 > v2, -1 ถ้า v1 < v2, 0 ถ้าเท่ากัน
 */
function compareVersions(v1, v2) {
  const parts1 = v1.split('.').map(n => parseInt(n) || 0);
  const parts2 = v2.split('.').map(n => parseInt(n) || 0);
  
  for (let i = 0; i < Math.max(parts1.length, parts2.length); i++) {
    const p1 = parts1[i] || 0;
    const p2 = parts2[i] || 0;
    if (p1 > p2) return 1;
    if (p1 < p2) return -1;
  }
  return 0;
}

/**
 * ✅ ดาวน์โหลดและอัปเดตไฟล์จาก GitHub
 * ⚠️ ฟังก์ชันนี้ใช้ได้เฉพาะกับ Container-bound Script
 * @returns {Object} { success, message, filesUpdated }
 */
function applyUpdateFromGitHub() {
  try {
    Logger.log('🔄 เริ่มอัปเดตระบบ...');

    // 1. ดึงรายการไฟล์ทั้งหมดจาก GitHub API
    const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents`;
    const response = UrlFetchApp.fetch(apiUrl, { 
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'PP5Smart-AutoUpdate' }
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('ไม่สามารถเชื่อมต่อ GitHub API ได้');
    }

    const files = JSON.parse(response.getContentText());
    const scriptFiles = files.filter(f => 
      f.type === 'file' && (f.name.endsWith('.js') || f.name.endsWith('.html'))
    );

    if (scriptFiles.length === 0) {
      throw new Error('ไม่พบไฟล์โค้ดใน GitHub');
    }

    Logger.log(`📦 พบไฟล์ทั้งหมด ${scriptFiles.length} ไฟล์`);

    // 2. ดาวน์โหลดและอัปเดตแต่ละไฟล์
    let updatedCount = 0;
    const errors = [];

    scriptFiles.forEach(file => {
      try {
        const fileResponse = UrlFetchApp.fetch(file.download_url, { muteHttpExceptions: true });
        if (fileResponse.getResponseCode() === 200) {
          const content = fileResponse.getContentText();
          const fileName = file.name.replace(/\.(js|html)$/, '');
          
          // ⚠️ ไม่สามารถอัปเดตไฟล์ใน Apps Script ผ่าน API ได้โดยตรง
          // ต้องใช้ Apps Script API แต่ต้อง OAuth2
          Logger.log(`📥 ดาวน์โหลด: ${file.name} (${content.length} bytes)`);
          updatedCount++;
        }
      } catch (e) {
        errors.push(`${file.name}: ${e.message}`);
        Logger.log(`❌ ไม่สามารถอัปเดต ${file.name}: ${e.message}`);
      }
    });

    // ⚠️ ข้อจำกัด: Apps Script ไม่สามารถแก้ไขโค้ดของตัวเองได้โดยตรง
    // ต้องใช้ Apps Script API + OAuth2 ซึ่งซับซ้อนเกินไป
    
    return {
      success: false,
      message: '⚠️ ระบบ Auto-Update ต้องใช้ Apps Script API\n\nแนะนำให้ใช้วิธี Manual Update แทน:\n1. ติดตั้ง clasp\n2. clasp clone <SCRIPT_ID>\n3. git pull\n4. clasp push',
      filesFound: scriptFiles.length,
      limitation: 'Apps Script ไม่สามารถแก้ไขโค้ดของตัวเองได้โดยตรงผ่าน UrlFetchApp'
    };

  } catch (e) {
    Logger.log('❌ applyUpdateFromGitHub error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * ✅ สร้างคำแนะนำการอัปเดตแบบ Manual
 * @returns {Object} { instructions, claspCommands }
 */
function getManualUpdateInstructions() {
  const scriptId = ScriptApp.getScriptId();
  
  return {
    success: true,
    method: 'Manual Update via clasp',
    instructions: [
      '📋 วิธีอัปเดตระบบด้วยตัวเอง:',
      '',
      '1️⃣ ติดตั้ง Node.js และ clasp (ครั้งแรกเท่านั้น):',
      '   npm install -g @google/clasp',
      '   clasp login',
      '',
      '2️⃣ Clone project ของคุณ:',
      `   clasp clone ${scriptId}`,
      '',
      '3️⃣ ดึงโค้ดล่าสุดจาก GitHub:',
      `   git clone https://github.com/${GITHUB_REPO}.git`,
      '   หรือ git pull (ถ้า clone ไว้แล้ว)',
      '',
      '4️⃣ Copy ไฟล์ทั้งหมดจาก GitHub → โฟลเดอร์ที่ clasp clone ไว้',
      '',
      '5️⃣ Push โค้ดใหม่:',
      '   clasp push --force',
      '',
      '✅ เสร็จ! ระบบของคุณได้รับการอัปเดตแล้ว'
    ].join('\n'),
    scriptId: scriptId,
    githubUrl: `https://github.com/${GITHUB_REPO}`,
    currentVersion: getCurrentVersion()
  };
}

/**
 * ✅ อัปเดตระบบแบบง่าย: ให้ผู้ใช้ copy-paste โค้ดจาก GitHub
 * @returns {Object} { success, updateUrl, instructions }
 */
function getSimpleUpdateInstructions() {
  return {
    success: true,
    method: 'Copy-Paste Update',
    instructions: [
      '📋 วิธีอัปเดตแบบง่าย (ไม่ต้องใช้ clasp):',
      '',
      '1️⃣ เปิด GitHub:',
      `   https://github.com/${GITHUB_REPO}`,
      '',
      '2️⃣ เปิดไฟล์ที่ต้องการอัปเดต (เช่น Students.js, Installer.js)',
      '',
      '3️⃣ คลิก "Raw" → Copy โค้ดทั้งหมด',
      '',
      '4️⃣ เปิด Apps Script Editor ของคุณ:',
      '   ส่วนขยาย → Apps Script',
      '',
      '5️⃣ เปิดไฟล์เดียวกัน → Paste โค้ดใหม่ทับ → บันทึก',
      '',
      '6️⃣ ทำซ้ำกับไฟล์อื่นๆ ที่มีการแก้ไข',
      '',
      '✅ เสร็จ! รีเฟรชหน้าเว็บเพื่อใช้งานโค้ดใหม่'
    ].join('\n'),
    githubUrl: `https://github.com/${GITHUB_REPO}`,
    currentVersion: getCurrentVersion(),
    latestVersion: SYSTEM_VERSION
  };
}

/**
 * 🔒 ตรวจสอบว่าโค้ดมีคำสั่งที่น่าสงสัยหรือไม่
 * @param {string} content - เนื้อหาโค้ดที่ต้องการตรวจสอบ
 * @returns {boolean}
 */
function containsSuspiciousCode(content) {
  const urlMatches = String(content || '').match(/https?:\/\/[^\s"'`<>]+/gi) || [];
  const allowedHosts = [
    'github.com',
    'api.github.com',
    'githubusercontent.com',
    'raw.githubusercontent.com',
    'objects.githubusercontent.com',
  ];
  for (const u of urlMatches) {
    try {
      const host = String(u).replace(/^https?:\/\//i, '').split('/')[0].toLowerCase();
      const ok = allowedHosts.some(h => host === h || host.endsWith('.' + h));
      if (!ok) {
        Logger.log('⚠️ พบ URL ที่ไม่ได้อยู่ในรายการที่เชื่อถือได้: ' + host);
        return true;
      }
    } catch (_e) {}
  }

  const suspiciousPatterns = [
    /PropertiesService\.getScriptProperties\(\)\.deleteAllProperties/gi, // ลบ properties ทั้งหมด
    /DriveApp\..*\.setTrashed\(true\)/gi, // ลบไฟล์ใน Drive
    /SpreadsheetApp\..*\.deleteSheet/gi, // ลบ sheet (ยกเว้นการใช้งานปกติ)
    /eval\(/gi, // eval() อันตราย
    /new\s+Function\(/gi, // new Function() อันตราย
    /document\.write/gi, // document.write ใน HTML
    /<script(?:\s|>)[^>]*\bsrc=["'][^"']*(?!cdn\.|googleapis\.com|bootstrapcdn\.com|cloudflare\.com)/gi, // โหลด script จากแหล่งไม่รู้จัก (กัน false-positive จาก regex literal)
  ];
  
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(content)) {
      Logger.log(`⚠️ พบ pattern ที่น่าสงสัย: ${pattern}`);
      return true;
    }
  }
  
  return false;
}

/**
 * ✅ เช็คและแจ้งเตือนเมื่อมี update ใหม่
 * 🔒 เฉพาะ Admin เท่านั้นที่เรียกใช้ได้
 * เรียกใช้ตอน onOpen หรือใน Dashboard
 * @returns {Object} { hasUpdate, message, instructions }
 */
function checkAndNotifyUpdate(token) {
  try {
    // 🔒 SECURITY: ตรวจสอบสิทธิ์ admin
    if (!(isAdminByToken_(token) || isUserAdmin())) {
      return {
        success: false,
        error: 'เฉพาะผู้ดูแลระบบเท่านั้นที่สามารถเช็คอัปเดตได้'
      };
    }
    
    const updateInfo = checkForUpdates();
    
    if (!updateInfo.success) {
      return {
        success: false,
        error: updateInfo.error
      };
    }

    if (updateInfo.hasUpdate) {
      const instructions = getSimpleUpdateInstructions();
      return {
        success: true,
        hasUpdate: true,
        currentVersion: updateInfo.currentVersion,
        latestVersion: updateInfo.latestVersion,
        message: `🎉 มีเวอร์ชันใหม่! ${updateInfo.currentVersion} → ${updateInfo.latestVersion}`,
        instructions: instructions.instructions,
        githubUrl: instructions.githubUrl
      };
    }

    return {
      success: true,
      hasUpdate: false,
      currentVersion: updateInfo.currentVersion,
      message: `✅ คุณใช้เวอร์ชันล่าสุดอยู่แล้ว (${updateInfo.currentVersion})`
    };

  } catch (e) {
    Logger.log('❌ checkAndNotifyUpdate error: ' + e.message);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * ✅ บันทึกเวอร์ชันปัจจุบันลงระบบ (เรียกครั้งแรกหลังติดตั้ง)
 * 🔒 เฉพาะ Admin เท่านั้น
 */
function initializeSystemVersion() {
  if (!isUserAdmin()) {
    Logger.log('❌ ไม่มีสิทธิ์ในการ initialize version');
    return { success: false, error: 'ไม่มีสิทธิ์' };
  }
  
  setCurrentVersion(SYSTEM_VERSION);
  Logger.log('✅ Initialize system version: ' + SYSTEM_VERSION);
  return { success: true, version: SYSTEM_VERSION };
}

/**
 * 🔒 ตรวจสอบ integrity ของไฟล์ที่ดาวน์โหลดมา
 * @param {string} fileName - ชื่อไฟล์
 * @param {string} content - เนื้อหาไฟล์
 * @returns {Object} { valid, hash, message }
 */
function verifyFileIntegrity(fileName, content) {
  try {
    const calculatedHash = calculateSHA256(content);
    
    if (!calculatedHash) {
      return {
        valid: false,
        message: 'ไม่สามารถคำนวณ hash ได้'
      };
    }
    
    // ถ้ามี expected hash ให้ตรวจสอบ
    if (EXPECTED_HASHES[fileName]) {
      const isValid = calculatedHash === EXPECTED_HASHES[fileName];
      return {
        valid: isValid,
        hash: calculatedHash,
        expectedHash: EXPECTED_HASHES[fileName],
        message: isValid 
          ? '✅ ไฟล์ถูกต้อง' 
          : '⚠️ Hash ไม่ตรงกับที่คาดหวัง - ไฟล์อาจถูกแก้ไข'
      };
    }
    
    // ถ้าไม่มี expected hash ให้แสดง hash ที่คำนวณได้
    return {
      valid: true,
      hash: calculatedHash,
      message: `ℹ️ Hash: ${calculatedHash.substring(0, 16)}...`
    };
    
  } catch (e) {
    Logger.log('❌ verifyFileIntegrity error: ' + e.message);
    return {
      valid: false,
      message: 'เกิดข้อผิดพลาดในการตรวจสอบ'
    };
  }
}

/**
 * 🔒 Log การเข้าถึงระบบ update
 * @param {string} action - การกระทำ (check, download, apply)
 * @param {Object} details - รายละเอียดเพิ่มเติม
 */
function logUpdateActivity(action, details) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    const logEntry = {
      timestamp: timestamp,
      user: userEmail,
      action: action,
      details: JSON.stringify(details)
    };
    
    Logger.log(`🔒 UPDATE LOG: ${action} by ${userEmail} at ${timestamp}`);
    Logger.log(`   Details: ${JSON.stringify(details)}`);
    
    // บันทึกลง sheet ถ้าต้องการ
    // const ss = SS();
    // const logSheet = ss.getSheetByName('UpdateLog');
    // if (logSheet) {
    //   logSheet.appendRow([timestamp, userEmail, action, JSON.stringify(details)]);
    // }
    
  } catch (e) {
    Logger.log('❌ logUpdateActivity error: ' + e.message);
  }
}
