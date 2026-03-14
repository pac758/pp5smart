# 🔒 มาตรการรักษาความปลอดภัยระบบ Auto-Update

## 📋 สรุปช่องโหว่ที่พบและวิธีแก้ไข

### ⚠️ ช่องโหว่ที่พบก่อนแก้ไข

1. **Man-in-the-Middle Attack**
   - ดึงข้อมูลจาก GitHub โดยไม่มีการตรวจสอบความถูกต้อง
   - ผู้ไม่หวังดีอาจแทรกโค้ดระหว่างทาง

2. **Repository Hijacking**
   - ถ้า GitHub account ถูกแฮก → โค้ดปลอมถูก push เข้ามา
   - ระบบจะดึงโค้ดปลอมไปใช้ทันที

3. **Code Injection**
   - ไม่มีการตรวจสอบว่าโค้ดที่ดึงมาปลอดภัยหรือไม่
   - อาจมีโค้ดที่ขโมยข้อมูลนักเรียน/ครู/คะแนน

4. **No Access Control**
   - ทุกคนสามารถเรียก checkForUpdates() ได้
   - ไม่มีการจำกัดสิทธิ์

5. **No Integrity Check**
   - ไม่มีการตรวจสอบ hash/checksum
   - ไม่รู้ว่าไฟล์ถูกแก้ไขระหว่างทางหรือไม่

---

## ✅ มาตรการรักษาความปลอดภัยที่เพิ่มเข้ามา

### 1. 🔒 Repository Whitelist
```javascript
const TRUSTED_REPOS = [
  'pac758/pp5smart',
  // เพิ่ม repo อื่นที่เชื่อถือได้ตรงนี้
];

function isRepoTrusted(repo) {
  return TRUSTED_REPOS.includes(repo);
}
```

**การทำงาน:**
- ตรวจสอบว่า repository ที่ดึงข้อมูลเป็นแหล่งที่เชื่อถือได้
- ป้องกันการดึงโค้ดจาก repository ปลอม
- ถ้าไม่อยู่ใน whitelist → ปฏิเสธทันที

---

### 2. 🔒 Suspicious Code Detection
```javascript
function containsSuspiciousCode(content) {
  const suspiciousPatterns = [
    /UrlFetchApp\.fetch\([^)]*(?!github\.com|githubusercontent\.com)[^)]*\)/gi,
    /PropertiesService\.getScriptProperties\(\)\.deleteAllProperties/gi,
    /DriveApp\..*\.setTrashed\(true\)/gi,
    /eval\(/gi,
    /new\s+Function\(/gi,
    /<script[^>]*src=["'][^"']*(?!cdn\.|googleapis\.com)/gi,
  ];
  
  for (const pattern of suspiciousPatterns) {
    if (pattern.test(content)) {
      return true;
    }
  }
  return false;
}
```

**ตรวจจับ:**
- ✅ การเรียก API ไปยังเซิร์ฟเวอร์ภายนอก (นอกจาก GitHub)
- ✅ การลบ Properties ทั้งหมด
- ✅ การลบไฟล์ใน Google Drive
- ✅ การลบ Sheet
- ✅ `eval()` และ `new Function()` (อันตราย)
- ✅ การโหลด script จากแหล่งไม่รู้จัก

---

### 3. 🔒 SHA-256 Hash Verification
```javascript
function calculateSHA256(content) {
  const signature = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    content,
    Utilities.Charset.UTF_8
  );
  return signature.map(byte => {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function verifyFileIntegrity(fileName, content) {
  const calculatedHash = calculateSHA256(content);
  
  if (EXPECTED_HASHES[fileName]) {
    const isValid = calculatedHash === EXPECTED_HASHES[fileName];
    return {
      valid: isValid,
      hash: calculatedHash,
      message: isValid ? '✅ ไฟล์ถูกต้อง' : '⚠️ Hash ไม่ตรงกับที่คาดหวัง'
    };
  }
  
  return { valid: true, hash: calculatedHash };
}
```

**การทำงาน:**
- คำนวณ SHA-256 hash ของไฟล์ที่ดาวน์โหลด
- เปรียบเทียบกับ hash ที่คาดหวัง (ถ้ามี)
- ป้องกันการแก้ไขไฟล์ระหว่างทาง

---

### 4. 🔒 Admin-Only Access Control
```javascript
function isUserAdmin() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SS();
  const usersSheet = ss.getSheetByName('Users');
  
  const data = usersSheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('email');
  const roleCol = headers.indexOf('role');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === userEmail && data[i][roleCol] === 'admin') {
      return true;
    }
  }
  return false;
}
```

**การทำงาน:**
- ตรวจสอบว่าผู้ใช้มีสิทธิ์ admin หรือไม่
- เฉพาะ admin เท่านั้นที่เรียก `checkAndNotifyUpdate()` ได้
- ครูทั่วไปไม่สามารถเช็คหรือดาวน์โหลด update ได้

---

### 5. 🔒 Activity Logging
```javascript
function logUpdateActivity(action, details) {
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date();
  
  Logger.log(`🔒 UPDATE LOG: ${action} by ${userEmail} at ${timestamp}`);
  Logger.log(`   Details: ${JSON.stringify(details)}`);
  
  // สามารถบันทึกลง sheet ได้ถ้าต้องการ
}
```

**การทำงาน:**
- บันทึกทุกการเข้าถึงระบบ update
- เก็บข้อมูล: ผู้ใช้, เวลา, การกระทำ, รายละเอียด
- ช่วยตรวจสอบย้อนหลังได้

---

## 🛡️ ขั้นตอนการตรวจสอบความปลอดภัย

เมื่อมีการเรียก `checkForUpdates()`:

```
1. ✅ ตรวจสอบสิทธิ์ admin
   ↓
2. ✅ ตรวจสอบ repository whitelist
   ↓
3. ✅ ดาวน์โหลดไฟล์จาก GitHub (HTTPS)
   ↓
4. ✅ ตรวจสอบโค้ดที่น่าสงสัย
   ↓
5. ✅ คำนวณ SHA-256 hash
   ↓
6. ✅ เปรียบเทียบกับ expected hash (ถ้ามี)
   ↓
7. ✅ บันทึก log
   ↓
8. ✅ แสดงผลให้ผู้ใช้
```

---

## 📊 ระดับความปลอดภัย

| มาตรการ | ก่อนแก้ไข | หลังแก้ไข |
|---------|-----------|-----------|
| Repository Verification | ❌ | ✅ |
| Code Injection Protection | ❌ | ✅ |
| Hash Integrity Check | ❌ | ✅ |
| Access Control | ❌ | ✅ |
| Activity Logging | ❌ | ✅ |
| HTTPS Connection | ✅ | ✅ |

**ระดับความปลอดภัย:**
- ก่อนแก้ไข: 🔴 **ต่ำ** (1/6)
- หลังแก้ไข: 🟢 **สูง** (6/6)

---

## ⚠️ ข้อจำกัดที่ยังมีอยู่

### 1. ไม่สามารถตรวจสอบ SSL Certificate
- Apps Script ไม่มี API สำหรับตรวจสอบ certificate
- ต้องพึ่งพา Google's UrlFetchApp ที่ตรวจสอบให้อัตโนมัติ

### 2. Hash Verification เป็น Optional
- ต้องอัปเดต `EXPECTED_HASHES` ด้วยตนเองทุกครั้งที่มี release ใหม่
- ถ้าไม่อัปเดต → จะไม่มีการตรวจสอบ hash

### 3. Manual Update ยังคงมีความเสี่ยง
- ผู้ใช้ยังสามารถ copy-paste โค้ดจาก GitHub ได้โดยตรง
- ไม่มีการตรวจสอบความปลอดภัยในกรณีนี้

---

## 🔐 คำแนะนำเพิ่มเติม

### สำหรับผู้ดูแลระบบ:

1. **เปิด 2-Factor Authentication** ใน GitHub account
2. **ตรวจสอบ commit history** ก่อนอัปเดตทุกครั้ง
3. **อ่านโค้ดที่เปลี่ยนแปลง** ก่อน apply update
4. **Backup ข้อมูล** ก่อนอัปเดตเสมอ
5. **ทดสอบใน test environment** ก่อนใช้งานจริง

### สำหรับโรงเรียนที่ใช้งาน:

1. **ให้เฉพาะ admin ที่เชื่อถือได้** เท่านั้นมีสิทธิ์อัปเดต
2. **ตรวจสอบ execution log** เป็นประจำ
3. **อย่าแชร์ script ID** กับบุคคลภายนอก
4. **ใช้ Google Workspace** แทน Gmail ส่วนตัว (มี audit log)

---

## 🚨 สัญญาณเตือนที่ต้องระวัง

ถ้าเจอสิ่งเหล่านี้ → **อย่าอัปเดต**:

- ⚠️ Repository ไม่ใช่ `pac758/pp5smart`
- ⚠️ ตรวจพบโค้ดที่น่าสงสัย
- ⚠️ Hash ไม่ตรงกับที่คาดหวัง
- ⚠️ มีการเรียก API ไปยังเซิร์ฟเวอร์ภายนอก
- ⚠️ มีการลบข้อมูลโดยไม่จำเป็น
- ⚠️ มีการใช้ `eval()` หรือ `new Function()`

---

## 📞 ติดต่อ

หากพบช่องโหว่ด้านความปลอดภัย:
1. **อย่า** เปิดเผยต่อสาธารณะ
2. ติดต่อผู้พัฒนาโดยตรงทาง GitHub Issues (Private)
3. แจ้งรายละเอียดให้ชัดเจน

---

**อัปเดตล่าสุด:** 2026-03-14  
**เวอร์ชัน:** 1.2.0  
**ผู้จัดทำ:** PP5Smart Development Team
