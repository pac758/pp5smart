/**************************************************
 * ⚙️ globalsettings.gs - Wrapper (เรียกใช้ settings_unified.gs)
 * ไฟล์นี้เก็บ Wrapper functions เพื่อความเข้ากันได้กับโค้ดเดิม
 * ฟังก์ชันหลักอยู่ใน settings_unified.gs แล้ว
 **************************************************/
 
const GLOBAL_SETTINGS_SHEET_NAME = 'global_settings';
const LEGACY_SETTINGS_SHEET_NAME = 'การตั้งค่าระบบ';
 
// === Wrapper Functions (เรียกไฟล์ใหม่) ===
 
function saveSystemSettings(settingsData) {
  return S_saveSystemSettings(settingsData);
}
 
function getWebAppSettingsWithCache() {
  return S_getWebAppSettingsWithCache();
}
 
function getWebAppSettings() {
  return S_getWebAppSettings();
}
 
function clearSettingsCache() {
  return S_clearSettingsCache();
}
 
function saveGlobalSettings(settingsData) {
  return S_saveGlobalSettings(settingsData);
}
 
 
 
function _getDefaultGlobalSettings_() {
  return S_getDefaultSettings();
}
 
function _openSpreadsheet_() {
  if (getSpreadsheetId_()) {
    return SS();
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
 
function _defaultAcademicYear_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  return month >= 5 ? year + 543 : year + 542;
}
 
function _createDefaultGlobalSettingsSheet_(spreadsheet) {
  return S_createSettingsSheet(spreadsheet);
}
 
// === ฟังก์ชันเฉพาะทาง (เก็บไว้) ===
 
function getLogoDataUrl(fileId) {
  if (!fileId) throw new Error('missing fileId');
  const blob = DriveApp.getFileById(fileId).getBlob();
  const mime = blob.getContentType() || 'image/png';
  const b64  = Utilities.base64Encode(blob.getBytes());
  return `data:${mime};base64,${b64}`;
}
 
function getPublicLogoUrl_(fileId) {
  if (!fileId) return '';
  try {
    var file = DriveApp.getFileById(fileId);
    var access = file.getSharingAccess();
    if (access !== DriveApp.Access.ANYONE_WITH_LINK && access !== DriveApp.Access.ANYONE) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  } catch (e) {
    Logger.log('❌ getPublicLogoUrl_ error: ' + e.message);
    return '';
  }
}
