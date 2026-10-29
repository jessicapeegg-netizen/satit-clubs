/**
 * ====================================================================
 * ระบบเลือกชุมนุมออนไลน์ (Satit VRU Clubs - High Concurrency Engine)
 * ออกแบบเพื่อรองรับนักเรียน 1,500 คนเข้าใช้งานพร้อมกัน
 * Google Sheets Database ID: 1K2JpQD5NsfKfXt27OvxgQ3-vVw0-dFSBpjqcplNqiDw
 * ====================================================================
 */

const SPREADSHEET_ID = '1K2JpQD5NsfKfXt27OvxgQ3-vVw0-dFSBpjqcplNqiDw';

function getSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID && SPREADSHEET_ID !== 'ใส่_SPREADSHEET_ID_ตรงนี้') {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      Logger.log('Error opening spreadsheet by ID: ' + e);
    }
  }
  return ss;
}

/**
 * เพิ่มเมนูด้านบนของ Google Sheets อัตโนมัติเมื่อเปิดไฟล์ชีต
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🚀 ระบบเลือกชุมนุม')
      .addItem('✨ สร้าง/รีเซ็ตแผ่นงานฐานข้อมูล (setupDatabase)', 'setupDatabase')
      .addItem('⚡ ล้างแคชระบบ (Clear Cache)', 'clearSystemCache')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen notice: ' + e);
  }
}

function clearSystemCache() {
  clearCachedData('db_full');
  try {
    SpreadsheetApp.getUi().alert('⚡ ล้างแคชระบบเรียบร้อยแล้ว ข้อมูลจะดึงสดจากชีตในการโหลดครั้งถัดไป');
  } catch (e) {}
}

/**
 * ฟังก์ชันสร้างแผ่นงาน (Sheets) และข้อมูลตั้งต้นอัตโนมัติ
 */
function setupDatabase() {
  const ss = getSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Google Spreadsheet กรุณาเปิดสคริปต์นี้จากเมนู "ส่วนขยาย (Extensions) > Apps Script" ภายใน Google Sheets');
  }

  // 1. แผ่นงาน Settings
  let sSettings = ss.getSheetByName('Settings');
  if (!sSettings) sSettings = ss.insertSheet('Settings');
  if (sSettings.getLastRow() === 0) {
    sSettings.appendRow(['Key', 'Value']);
    sSettings.appendRow(['is_open', 'true']);
    sSettings.appendRow(['term', 'ภาคเรียนที่ 2/2569']);
    sSettings.appendRow(['logo', '🧸']);
    formatHeader(sSettings);
  }

  // 2. แผ่นงาน Students
  let sStudents = ss.getSheetByName('Students');
  if (!sStudents) sStudents = ss.insertSheet('Students');
  if (sStudents.getLastRow() === 0) {
    sStudents.appendRow(['id', 'prefix', 'name', 'surname', 'class', 'cancel_count']);
    sStudents.appendRow(['10001', 'เด็กชาย', 'ปัญญา', 'ใจดี', 'ม.1/1', 0]);
    sStudents.appendRow(['10002', 'เด็กหญิง', 'สมศรี', 'มีสุข', 'ม.1/1', 0]);
    sStudents.appendRow(['10003', 'นาย', 'กิตติ', 'รักเรียน', 'ม.4/1', 0]);
    formatHeader(sStudents);
  }

  // 3. แผ่นงาน Clubs
  let sClubs = ss.getSheetByName('Clubs');
  if (!sClubs) sClubs = ss.insertSheet('Clubs');
  if (sClubs.getLastRow() === 0) {
    sClubs.appendRow(['id', 'name', 'teacher', 'capacity', 'image']);
    sClubs.appendRow([1, 'ชุมนุมหุ่นยนต์และ AI', 'ครูสมศักดิ์ นวัตกรรม', 30, 'https://images.unsplash.com/photo-1485827404703-89b55fcc595e?w=600']);
    sClubs.appendRow([2, 'ชุมนุมดนตรีสากล', 'ครูดนัย เสียงทอง, ครูกิตติ รักเรียน', 60, 'https://images.unsplash.com/photo-1511671782779-c97d3d27a1d4?w=600']);
    sClubs.appendRow([3, 'ชุมนุมบอร์ดเกมเพื่อการเรียนรู้', 'ครูวิภา พัฒนาคิด', 30, 'https://images.unsplash.com/photo-1610890716171-6b1bb98ffd09?w=600']);
    sClubs.appendRow([4, 'ชุมนุมกีฬาบาสเกตบอล', 'ครูธนกร สุขภาพดี, ครูสมศักดิ์ นวัตกรรม', 60, 'https://images.unsplash.com/photo-1546519638-68e109498ffc?w=600']);
    formatHeader(sClubs);
  }

  // 4. แผ่นงาน Enrollments
  let sEnrollments = ss.getSheetByName('Enrollments');
  if (!sEnrollments) sEnrollments = ss.insertSheet('Enrollments');
  if (sEnrollments.getLastRow() === 0) {
    sEnrollments.appendRow(['student_id', 'club_id', 'timestamp']);
    formatHeader(sEnrollments);
  }

  // 5. แผ่นงาน Admins
  let sAdmins = ss.getSheetByName('Admins');
  if (!sAdmins) sAdmins = ss.insertSheet('Admins');
  if (sAdmins.getLastRow() === 0) {
    sAdmins.appendRow(['username', 'password', 'name']);
    sAdmins.appendRow(['admin', 'admin1234', 'ผู้ดูแลระบบโรงเรียน']);
    formatHeader(sAdmins);
  }

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }

  SpreadsheetApp.flush();
  clearCachedData('db_full');

  Logger.log('✅ ตั้งค่าฐานข้อมูลสำเร็จครบทุก Sheet เรียบร้อยแล้ว!');
  try {
    SpreadsheetApp.getUi().alert('✅ ตั้งค่าฐานข้อมูลสำเร็จ!\nสร้างแผ่นงาน Settings, Students, Clubs, Enrollments, Admins ครบถ้วนแล้ว');
  } catch (e) {}
}

function formatHeader(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  header.setBackground('#ffd54f');
  header.setFontColor('#37474f');
  header.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * ระบบ Chunks Caching สำหรับรองรับข้อมูลขนาดใหญ่เกิน 100KB ใน CacheService
 */
function setCachedData(key, obj, expirationSeconds) {
  try {
    const cache = CacheService.getScriptCache();
    const json = JSON.stringify(obj);
    const CHUNK_SIZE = 85000; // 85KB ปลอดภัยไม่เกินลิมิต 100KB ต่อคีย์
    const numChunks = Math.ceil(json.length / CHUNK_SIZE);
    
    const entries = {};
    entries[key + '_count'] = String(numChunks);
    for (let i = 0; i < numChunks; i++) {
      entries[key + '_' + i] = json.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    cache.putAll(entries, expirationSeconds || 120);
  } catch (err) {
    Logger.log('Cache set warning: ' + err);
  }
}

function getCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get(key + '_count');
    if (!countStr) return null;
    const numChunks = parseInt(countStr, 10);
    const keys = [];
    for (let i = 0; i < numChunks; i++) keys.push(key + '_' + i);
    const chunks = cache.getAll(keys);
    let json = '';
    for (let i = 0; i < numChunks; i++) {
      if (!chunks[key + '_' + i]) return null;
      json += chunks[key + '_' + i];
    }
    return JSON.parse(json);
  } catch (err) {
    Logger.log('Cache get warning: ' + err);
    return null;
  }
}

function clearCachedData(key) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get(key + '_count');
    if (countStr) {
      const numChunks = parseInt(countStr, 10);
      const keys = [key + '_count'];
      for (let i = 0; i < numChunks; i++) keys.push(key + '_' + i);
      cache.removeAll(keys);
    }
  } catch (err) {}
}

/**
 * GET Request
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Satit VRU Clubs High-Concurrency API is active.',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST Request: แยก Read กับ Write ชัดเจนเพื่อรองรับ 1,500 คนพร้อมกัน
 */
function doPost(e) {
  let postData = {};
  try {
    postData = e.postData ? JSON.parse(e.postData.contents) : {};
  } catch (err) {
    return createJsonResponse({ status: 'error', message: 'รูปแบบข้อมูลไม่ถูกต้อง' });
  }

  const action = postData.action;

  // ⚡ 1. คำขออ่านข้อมูล (getData): ไม่ใช้ Lock! ใช้ CacheService ตอบกลับใน 5-15ms!
  // วิธีนี้ทำให้เด็ก 1,500 คนกดเข้าเว็บพร้อมกันได้โดยไม่มีทางล่มหรือ Timeout
  if (action === 'getData') {
    let data = getCachedData('db_full');
    if (!data) {
      const ss = getSpreadsheet();
      data = getFullDatabase(ss);
      setCachedData('db_full', data, 60); // แคชไว้ 60 วินาที
    }
    return createJsonResponse({ status: 'success', data: data });
  }

  // ⚡ 2. คำขอเขียนข้อมูล (enroll, cancel, adminSave): ใช้ Micro-Lock สั้นที่สุด
  const lock = LockService.getScriptLock();
  // รอคิวไม่เกิน 15 วินาที
  const success = lock.tryLock(15000);
  if (!success) {
    return createJsonResponse({
      status: 'error',
      busy: true, // แจ้งฝั่งหน้าเว็บให้ทำการ Auto-retry อัตโนมัติ
      message: 'มีผู้ใช้งานเข้าคิวพร้อมกันจำนวนมาก ระบบกำลังจัดคิวให้อัตโนมัติ กรุณารอสักครู่...'
    });
  }

  try {
    const ss = getSpreadsheet();

    // ----------------------------------------------------
    // Action: enroll (สมัครชุมนุม) - Micro-transaction
    // ----------------------------------------------------
    if (action === 'enroll') {
      const studentId = String(postData.student_id).trim();
      const clubId = Number(postData.club_id);

      // ตรวจสอบสถานะเปิดรับสมัคร
      const sSettings = ss.getSheetByName('Settings');
      let isOpen = false;
      if (sSettings && sSettings.getLastRow() > 1) {
        const sVals = sSettings.getRange(2, 1, sSettings.getLastRow() - 1, 2).getValues();
        for (let r of sVals) {
          if (String(r[0]).trim() === 'is_open') {
            isOpen = (String(r[1]).toLowerCase() === 'true' || r[1] === true);
            break;
          }
        }
      }
      if (!isOpen) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ขออภัย ขณะนี้ระบบปิดรับสมัครแล้ว' });
      }

      // ตรวจสอบประวัติการสมัครใน Enrollments
      const sEnroll = ss.getSheetByName('Enrollments');
      const lastEnrollRow = sEnroll.getLastRow();
      let existingEnroll = false;
      let clubEnrolledCount = 0;

      if (lastEnrollRow > 1) {
        const enrollData = sEnroll.getRange(2, 1, lastEnrollRow - 1, 2).getValues();
        for (let row of enrollData) {
          if (String(row[0]).trim() === studentId) {
            existingEnroll = true;
            break;
          }
          if (Number(row[1]) === clubId) {
            clubEnrolledCount++;
          }
        }
      }

      if (existingEnroll) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'คุณได้เลือกชุมนุมไปแล้ว หากต้องการเปลี่ยนกรุณายกเลิกชุมนุมเดิมก่อน' });
      }

      // ตรวจสอบจำนวนรับของชุมนุมจาก Clubs
      const sClubs = ss.getSheetByName('Clubs');
      const lastClubRow = sClubs.getLastRow();
      let clubCapacity = 30;
      let clubFound = false;

      if (lastClubRow > 1) {
        const clubData = sClubs.getRange(2, 1, lastClubRow - 1, 4).getValues();
        for (let row of clubData) {
          if (Number(row[0]) === clubId) {
            clubCapacity = Number(row[3]) || 30;
            clubFound = true;
            break;
          }
        }
      }

      if (!clubFound) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ไม่พบชุมนุมที่เลือก' });
      }

      if (clubEnrolledCount >= clubCapacity) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ขออภัย ชุมนุมนี้มีผู้สมัครเต็มจำนวนแล้ว' });
      }

      // บันทึกลง Sheet Enrollments ทันที
      sEnroll.appendRow([studentId, clubId, new Date().toLocaleString('th-TH')]);
      SpreadsheetApp.flush();

      // ⚡ ปล่อย Lock ทันทีที่บันทึกเสร็จ เพื่อให้คิวถัดไปทำงานได้เร็วที่สุด (< 150ms ต่อคน!)
      lock.releaseLock();

      // อัปเดต Cache สดใหม่
      clearCachedData('db_full');
      const updatedDb = getFullDatabase(ss);
      setCachedData('db_full', updatedDb, 60);

      return createJsonResponse({ status: 'success', data: updatedDb, message: 'สมัครชุมนุมสำเร็จเรียบร้อยแล้ว' });
    }

    // ----------------------------------------------------
    // Action: cancel (ยกเลิกชุมนุม)
    // ----------------------------------------------------
    if (action === 'cancel') {
      const studentId = String(postData.student_id).trim();

      const sSettings = ss.getSheetByName('Settings');
      let isOpen = false;
      if (sSettings && sSettings.getLastRow() > 1) {
        const sVals = sSettings.getRange(2, 1, sSettings.getLastRow() - 1, 2).getValues();
        for (let r of sVals) {
          if (String(r[0]).trim() === 'is_open') {
            isOpen = (String(r[1]).toLowerCase() === 'true' || r[1] === true);
            break;
          }
        }
      }
      if (!isOpen) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ระบบปิดรับสมัครแล้ว ไม่สามารถยกเลิกได้' });
      }

      const sStudent = ss.getSheetByName('Students');
      const studentData = sStudent.getDataRange().getValues();
      let studentRowIndex = -1;
      let currentCancelCount = 0;

      for (let i = 1; i < studentData.length; i++) {
        if (String(studentData[i][0]).trim() === studentId) {
          studentRowIndex = i + 1;
          currentCancelCount = Number(studentData[i][5]) || 0;
          break;
        }
      }

      if (studentRowIndex === -1) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ไม่พบข้อมูลนักเรียน' });
      }

      if (currentCancelCount >= 1) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'คุณได้ใช้สิทธิ์ยกเลิกชุมนุมครบกำหนดแล้ว (จำกัด 1 ครั้ง)' });
      }

      const sEnroll = ss.getSheetByName('Enrollments');
      const enrollData = sEnroll.getDataRange().getValues();
      let foundEnroll = false;
      for (let i = enrollData.length - 1; i >= 1; i--) {
        if (String(enrollData[i][0]).trim() === studentId) {
          sEnroll.deleteRow(i + 1);
          foundEnroll = true;
          break;
        }
      }

      if (!foundEnroll) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ไม่พบประวัติการเลือกชุมนุมของคุณ' });
      }

      sStudent.getRange(studentRowIndex, 6).setValue(currentCancelCount + 1);
      SpreadsheetApp.flush();

      lock.releaseLock();

      clearCachedData('db_full');
      const updatedDb = getFullDatabase(ss);
      setCachedData('db_full', updatedDb, 60);

      return createJsonResponse({ status: 'success', data: updatedDb, message: 'ยกเลิกชุมนุมเรียบร้อยแล้ว' });
    }

    // ----------------------------------------------------
    // Action: adminSave
    // ----------------------------------------------------
    if (action === 'adminSave') {
      const newDb = postData.db;
      if (!newDb) {
        lock.releaseLock();
        return createJsonResponse({ status: 'error', message: 'ไม่พบข้อมูลที่ส่งมาบันทึก' });
      }

      saveFullDatabase(ss, newDb);
      SpreadsheetApp.flush();
      lock.releaseLock();

      clearCachedData('db_full');
      const updatedDb = getFullDatabase(ss);
      setCachedData('db_full', updatedDb, 60);

      return createJsonResponse({ status: 'success', data: updatedDb, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' });
    }

    lock.releaseLock();
    return createJsonResponse({ status: 'error', message: 'ไม่พบ Action ที่ระบุ' });

  } catch (err) {
    try { lock.releaseLock(); } catch(e) {}
    return createJsonResponse({ status: 'error', message: 'เกิดข้อผิดพลาด: ' + err.toString() });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ดึงข้อมูลทั้งหมดจาก Google Sheets มาประกอบเป็น Object Database
 */
function getFullDatabase(ss) {
  // 1. Settings
  const settings = { is_open: false, term: 'ภาคเรียนที่ 2/2569', logo: '🧸' };
  const sSettings = ss.getSheetByName('Settings');
  if (sSettings && sSettings.getLastRow() > 1) {
    const sRows = sSettings.getRange(2, 1, sSettings.getLastRow() - 1, 2).getValues();
    sRows.forEach(r => {
      const k = String(r[0]).trim();
      const v = r[1];
      if (k === 'is_open') settings.is_open = (String(v).toLowerCase() === 'true' || v === true);
      else if (k === 'term') settings.term = String(v);
      else if (k === 'logo') settings.logo = String(v);
    });
  }

  // 2. Students (รองรับ 1,500 คนได้อย่างรวดเร็ว)
  const students = [];
  const sStudents = ss.getSheetByName('Students');
  if (sStudents && sStudents.getLastRow() > 1) {
    const stRows = sStudents.getRange(2, 1, sStudents.getLastRow() - 1, 6).getValues();
    stRows.forEach(r => {
      if (r[0]) {
        students.push({
          id: String(r[0]).trim(),
          prefix: String(r[1] || ''),
          name: String(r[2] || ''),
          surname: String(r[3] || ''),
          class: String(r[4] || ''),
          cancel_count: Number(r[5]) || 0
        });
      }
    });
  }

  // 3. Enrollments
  const enrollments = [];
  const sEnroll = ss.getSheetByName('Enrollments');
  if (sEnroll && sEnroll.getLastRow() > 1) {
    const enRows = sEnroll.getRange(2, 1, sEnroll.getLastRow() - 1, 3).getValues();
    enRows.forEach(r => {
      if (r[0] && r[1]) {
        enrollments.push({
          student_id: String(r[0]).trim(),
          club_id: Number(r[1]),
          timestamp: String(r[2] || '')
        });
      }
    });
  }

  // 4. Clubs
  const clubs = [];
  const sClubs = ss.getSheetByName('Clubs');
  if (sClubs && sClubs.getLastRow() > 1) {
    const clRows = sClubs.getRange(2, 1, sClubs.getLastRow() - 1, 5).getValues();
    clRows.forEach(r => {
      if (r[0]) {
        const cId = Number(r[0]);
        const enrolled = enrollments.filter(e => e.club_id === cId).length;
        clubs.push({
          id: cId,
          name: String(r[1] || ''),
          teacher: String(r[2] || ''),
          capacity: Number(r[3]) || 30,
          current_enrolled: enrolled,
          image: String(r[4] || 'https://images.unsplash.com/photo-1541339907198-e08756dedf3f?w=400')
        });
      }
    });
  }

  // 5. Admins
  const admins = [];
  const sAdmins = ss.getSheetByName('Admins');
  if (sAdmins && sAdmins.getLastRow() > 1) {
    const adRows = sAdmins.getRange(2, 1, sAdmins.getLastRow() - 1, 3).getValues();
    adRows.forEach(r => {
      if (r[0]) {
        admins.push({
          username: String(r[0]).trim(),
          password: String(r[1] || '').trim(),
          name: String(r[2] || 'ผู้ดูแลระบบ')
        });
      }
    });
  } else {
    admins.push({ username: 'admin', password: 'password', name: 'ผู้ดูแลระบบ' });
  }

  return { settings, students, clubs, enrollments, admins };
}

/**
 * บันทึกข้อมูลทั้งหมดลงใน Sheets เมื่อ Admin มีการเปลี่ยนแปลง
 */
function saveFullDatabase(ss, newDb) {
  // 1. Update Settings
  if (newDb.settings) {
    let sSettings = ss.getSheetByName('Settings');
    if (!sSettings) {
      sSettings = ss.insertSheet('Settings');
      formatHeader(sSettings);
    }
    sSettings.clearContents();
    sSettings.appendRow(['Key', 'Value']);
    sSettings.appendRow(['is_open', String(newDb.settings.is_open)]);
    sSettings.appendRow(['term', String(newDb.settings.term || '')]);
    sSettings.appendRow(['logo', String(newDb.settings.logo || '🧸')]);
  }

  // 2. Update Students
  if (Array.isArray(newDb.students)) {
    let sStudents = ss.getSheetByName('Students');
    if (!sStudents) {
      sStudents = ss.insertSheet('Students');
      formatHeader(sStudents);
    }
    sStudents.clearContents();
    sStudents.appendRow(['id', 'prefix', 'name', 'surname', 'class', 'cancel_count']);
    const stRows = newDb.students.map(s => [
      String(s.id),
      String(s.prefix || ''),
      String(s.name || ''),
      String(s.surname || ''),
      String(s.class || ''),
      Number(s.cancel_count) || 0
    ]);
    if (stRows.length > 0) {
      sStudents.getRange(2, 1, stRows.length, 6).setValues(stRows);
    }
  }

  // 3. Update Clubs
  if (Array.isArray(newDb.clubs)) {
    let sClubs = ss.getSheetByName('Clubs');
    if (!sClubs) {
      sClubs = ss.insertSheet('Clubs');
      formatHeader(sClubs);
    }
    sClubs.clearContents();
    sClubs.appendRow(['id', 'name', 'teacher', 'capacity', 'image']);
    const clRows = newDb.clubs.map(c => [
      Number(c.id),
      String(c.name || ''),
      String(c.teacher || ''),
      Number(c.capacity) || 30,
      String(c.image || '')
    ]);
    if (clRows.length > 0) {
      sClubs.getRange(2, 1, clRows.length, 5).setValues(clRows);
    }
  }

  // 4. Update Enrollments
  if (Array.isArray(newDb.enrollments)) {
    let sEnroll = ss.getSheetByName('Enrollments');
    if (!sEnroll) {
      sEnroll = ss.insertSheet('Enrollments');
      formatHeader(sEnroll);
    }
    sEnroll.clearContents();
    sEnroll.appendRow(['student_id', 'club_id', 'timestamp']);
    const enRows = newDb.enrollments.map(e => [
      String(e.student_id),
      Number(e.club_id),
      String(e.timestamp || new Date().toLocaleString('th-TH'))
    ]);
    if (enRows.length > 0) {
      sEnroll.getRange(2, 1, enRows.length, 3).setValues(enRows);
    }
  }

  // 5. Update Admins
  if (Array.isArray(newDb.admins) && newDb.admins.length > 0) {
    let sAdmins = ss.getSheetByName('Admins');
    if (!sAdmins) {
      sAdmins = ss.insertSheet('Admins');
      formatHeader(sAdmins);
    }
    sAdmins.clearContents();
    sAdmins.appendRow(['username', 'password', 'name']);
    const adRows = newDb.admins.map(a => [
      String(a.username),
      String(a.password),
      String(a.name || 'ผู้ดูแลระบบ')
    ]);
    sAdmins.getRange(2, 1, adRows.length, 3).setValues(adRows);
  }
}
