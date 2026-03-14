/**
 * แยก fileId จากลิงก์ Google Drive หรือรับเป็น id ตรง ๆ
 * รองรับ: /file/d/<id>/view, uc?id=<id>, open?id=<id>, share?id=<id> ฯลฯ
 */
function parseDriveFileId(raw) {
  try {
    if (!raw) return { ok:false, msg:'ไม่มีข้อมูลลิงก์/ไอดี' };
    const s = String(raw).trim();

    // 1) จากพาธ /file/d/<id>/
    let m = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
    if (m && m[1]) return { ok:true, fileId:m[1] };

    // 2) จาก query id=
    m = s.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
    if (m && m[1]) return { ok:true, fileId:m[1] };

    // 3) เป็นไอดีล้วน
    m = s.match(/^([a-zA-Z0-9_-]{25,})$/);
    if (m && m[1]) return { ok:true, fileId:m[1] };

    // 4) เผื่อรูปแบบอื่น ๆ – ดึง token ที่ยาวพอ
    m = s.match(/([a-zA-Z0-9_-]{25,})/);
    if (m && m[1]) return { ok:true, fileId:m[1] };

    return { ok:false, msg:'หา fileId ไม่พบจากลิงก์/ข้อความนี้' };
  } catch (e) {
    return { ok:false, msg:'parseDriveFileId error: ' + e.message };
  }
}

/**
 * ตั้งค่า logoFileId แล้วคืน settings ปัจจุบัน
 */
function setLogoFileId(fileId) {
  if (!fileId) throw new Error('ไม่มี fileId');
  // บันทึกค่า
  saveSystemSettings({ logoFileId: fileId });
  try { clearSettingsCache(); } catch (_e) {}
  // ตรวจสอบการแชร์เพื่อแจ้งเตือนใน log (ไม่โยน error)
  try {
    const f = DriveApp.getFileById(fileId);
    const acc = f.getSharingAccess();
    if (acc !== DriveApp.Access.ANYONE_WITH_LINK && acc !== DriveApp.Access.ANYONE) {
      Logger.log('⚠️ โลโก้ยังไม่ได้แชร์ Public');
    }
  } catch (e) {
    Logger.log('ตรวจสอบแชร์โลกล้มเหลว: ' + e.message);
  }
  return getGlobalSettings(false);
}

/**
 * อัปโหลดโลโก้จาก dataURL -> Drive
 * - สร้าง/ใช้โฟลเดอร์ชื่อ "SystemAssets" (หรือจะเปลี่ยนชื่อได้)
 * - ตั้งค่าแชร์ Anyone with the link (VIEW)
 * - บันทึก logoFileId ให้อัตโนมัติ
 */
function uploadLogoFromDataUrl(dataUrl, filename) {
  try {
    if (!dataUrl || typeof dataUrl !== 'string' || dataUrl.indexOf('data:') !== 0) {
      return { ok:false, msg:'dataUrl ไม่ถูกต้อง' };
    }
    filename = filename || 'school-logo.png';

    // แยก MIME + base64
    const match = dataUrl.match(/^data:([^;]+);base64,(.*)$/);
    if (!match) return { ok:false, msg:'รูปแบบ dataUrl ไม่ถูกต้อง' };
    const mime = match[1];
    const base64 = match[2];

    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mime, filename);

    // หาโฟลเดอร์เป้าหมาย
    const folder = _ensureSystemAssetsFolder_();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ตั้งเป็นโลโก้
    saveSystemSettings({ logoFileId: file.getId() });
    try { clearSettingsCache(); } catch (_e) {}

    return { ok:true, fileId:file.getId(), name:file.getName() };
  } catch (e) {
    return { ok:false, msg:'uploadLogoFromDataUrl error: ' + e.message };
  }
}

/**
 * สร้าง/คืนโฟลเดอร์เก็บไฟล์ระบบ
 * - พยายามใช้ pdfSaveFolderId ถ้าตั้งค่าไว้
 * - ถ้าไม่มี ให้สร้างโฟลเดอร์ชื่อ "SystemAssets"
 */
function _ensureSystemAssetsFolder_() {
  try {
    const settings = getGlobalSettings ? getGlobalSettings() : {};
    const pdfFolderId = settings && settings['pdfSaveFolderId'];

    if (pdfFolderId) {
      try {
        return DriveApp.getFolderById(pdfFolderId);
      } catch (e) {
        Logger.log('pdfSaveFolderId ใช้ไม่ได้: ' + e.message);
      }
    }

    // หาโฟลเดอร์ SystemAssets ในราก
    const it = DriveApp.getFoldersByName('SystemAssets');
    if (it.hasNext()) return it.next();

    // สร้างใหม่
    const folder = DriveApp.createFolder('SystemAssets');
    return folder;
  } catch (e) {
    throw new Error('_ensureSystemAssetsFolder_ error: ' + e.message);
  }
}
/**
 * ✨ ฟังก์ชันทำให้โลโก้เป็น Public
 * เรียกใช้โดยปุ่ม "ทำให้ไฟล์โลโก้เป็น Public" ในหน้าเว็บ
 */
function fixLogoSharing() {
  try {
    const settings = S_getGlobalSettings(false);
    const fileId = settings['logoFileId'];
    
    if (!fileId || fileId.trim() === '') {
      return '⚠️ ไม่พบไอดีโลโก้ในระบบ กรุณาตั้งค่าโลโก้ก่อน';
    }
    
    // ดึงไฟล์จาก Drive
    const file = DriveApp.getFileById(fileId.trim());
    
    // ตรวจสอบสถานะปัจจุบัน
    const currentAccess = file.getSharingAccess();
    Logger.log('🔐 Sharing Access ปัจจุบัน: ' + currentAccess);
    
    // ตั้งค่าให้เป็น Public (Anyone with the link can VIEW)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    try { clearSettingsCache(); } catch (_e) {}
    
    Logger.log('✅ ตั้งค่าแชร์โลโก้เป็น Public สำเร็จ: ' + fileId);
    Logger.log('🔗 URL: https://drive.google.com/uc?export=view&id=' + fileId);
    
    return '✅ ตั้งค่าแชร์โลโก้เป็น Public สำเร็จแล้ว!\n' +
           'ชื่อไฟล์: ' + file.getName() + '\n' +
           'ขนาดไฟล์: ' + Math.round(file.getSize() / 1024) + ' KB';
    
  } catch (e) {
    Logger.log('❌ fixLogoSharing error: ' + e.message);
    
    // ตรวจสอบประเภท error
    if (e.message.includes('not found')) {
      throw new Error('ไม่พบไฟล์โลโก้ในระบบ (File ID อาจผิดหรือไฟล์ถูกลบแล้ว)');
    } else if (e.message.includes('Permission')) {
      throw new Error('ไม่มีสิทธิ์เข้าถึงไฟล์นี้ กรุณาตรวจสอบว่าคุณเป็นเจ้าของไฟล์');
    }
    
    throw new Error('ไม่สามารถตั้งค่าแชร์ได้: ' + e.message);
  }
}

/**
 * ✨ ฟังก์ชันทดสอบโลโก้
 * เรียกใช้โดยปุ่ม "ทดสอบโลโก้ (Console)" ในหน้าเว็บ
 * ผลลัพธ์จะแสดงใน Logger (View → Execution log)
 */
function testLogoUrl() {
  try {
    const settings = S_getGlobalSettings(false);
    const fileId = settings['logoFileId'] || '';
    
    Logger.log('');
    Logger.log('========================================');
    Logger.log('=== 🔍 ทดสอบโลโก้โรงเรียน ===');
    Logger.log('========================================');
    Logger.log('');
    
    // ตรวจสอบว่ามี File ID หรือไม่
    if (!fileId || fileId.trim() === '') {
      Logger.log('❌ ไม่มีไอดีโลโก้ในระบบ');
      Logger.log('💡 กรุณาอัปโหลดโลโก้หรือตั้งค่าไอดีก่อน');
      return;
    }
    
    Logger.log('📋 File ID: ' + fileId);
    Logger.log('');
    
    // ทดสอบดึงไฟล์จาก Drive
    try {
      const file = DriveApp.getFileById(fileId.trim());
      
      Logger.log('✅ พบไฟล์ในระบบ');
      Logger.log('📄 ชื่อไฟล์: ' + file.getName());
      Logger.log('📊 ขนาดไฟล์: ' + Math.round(file.getSize() / 1024) + ' KB');
      Logger.log('📁 MIME Type: ' + file.getMimeType());
      Logger.log('📅 วันที่สร้าง: ' + file.getDateCreated());
      Logger.log('📅 วันที่แก้ไข: ' + file.getLastUpdated());
      Logger.log('');
      
      // ตรวจสอบการแชร์
      const access = file.getSharingAccess();
      const permission = file.getSharingPermission();
      
      Logger.log('🔐 การตั้งค่าการแชร์:');
      Logger.log('   - Sharing Access: ' + access);
      Logger.log('   - Sharing Permission: ' + permission);
      Logger.log('');
      
      // ตรวจสอบว่าเป็น Public หรือไม่
      const isPublic = (access === DriveApp.Access.ANYONE_WITH_LINK || 
                        access === DriveApp.Access.ANYONE);
      
      if (isPublic) {
        Logger.log('✅ สถานะ: ไฟล์เป็น Public แล้ว (ใช้งานได้)');
      } else {
        Logger.log('⚠️ สถานะ: ไฟล์ยังไม่เป็น Public (จะแสดงไม่ได้)');
        Logger.log('💡 คำแนะนำ: กรุณากดปุ่ม "ทำให้ไฟล์โลโก้เป็น Public"');
      }
      Logger.log('');
      
      // แสดง URL ต่าง ๆ
      Logger.log('🔗 URL สำหรับเข้าถึงไฟล์:');
      Logger.log('   1. Direct View: https://drive.google.com/uc?export=view&id=' + fileId);
      Logger.log('   2. Alternative: https://drive.google.com/uc?id=' + fileId);
      Logger.log('   3. Drive UI: https://drive.google.com/file/d/' + fileId + '/view');
      Logger.log('');
      
      // ทดสอบเจ้าของไฟล์
      try {
        const owner = file.getOwner();
        Logger.log('👤 เจ้าของไฟล์: ' + owner.getEmail());
      } catch (e) {
        Logger.log('⚠️ ไม่สามารถตรวจสอบเจ้าของไฟล์ได้');
      }
      
    } catch (e) {
      Logger.log('❌ ไม่สามารถเข้าถึงไฟล์ได้');
      Logger.log('📋 Error: ' + e.message);
      Logger.log('');
      Logger.log('💡 สาเหตุที่เป็นไปได้:');
      Logger.log('   1. File ID ไม่ถูกต้อง');
      Logger.log('   2. ไฟล์ถูกลบแล้ว');
      Logger.log('   3. ไม่มีสิทธิ์เข้าถึงไฟล์');
    }
    
    Logger.log('');
    Logger.log('========================================');
    Logger.log('=== จบการทดสอบ ===');
    Logger.log('========================================');
    
  } catch (e) {
    Logger.log('❌ testLogoUrl error: ' + e.message);
  }
}

