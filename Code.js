// ============================================================
// Code.gs — iPad Management System v2
// ============================================================

const CONFIG = {
  SHEETS: {
    IPAD: 'iPad_Data', USERS: 'Users', LOG: 'Log', STUDENTS: 'Students',
    ADMIN: 'Admin', DATABASE: 'Database', ISSUES: 'Issues', ACCOUNTS: 'Accounts',
    TEACHERS: 'Teachers', APP_REQUESTS: 'AppRequests'
  },
  STATUS: {
    AVAILABLE: 'ยังไม่ยืม', BORROWED: 'ยืม', RETURNED: 'คืน', CLAIMED: 'เคลม',
    PENDING_ISSUE: 'อยู่ระหว่างการส่งเรื่อง',
    UNDER_INSPECTION: 'อยู่ระหว่างการตรวจสอบ',
    DEFECTIVE: 'ใช้งานไม่ได้',
    UNDER_CLAIM: 'อยู่ระหว่างการเคลม'
  },
  ROLES: { USER: 'user', ADMIN: 'admin', SUPER: 'super admin' }
};

// iPad_Data columns: 1:ID 2:Serial 3:AssetCode 4:Prefix 5:FirstName 6:LastName
//                    7:PersonCode 8:Position 9:Grade 10:Room 11:Status
//                    12:BorrowDate 13:ReturnDate 14:Notes
// Students columns:  1:ลำดับ 2:รหัส 3:คำนำหน้า 4:ชื่อ 5:นามสกุล 6:ชั้น 7:ห้อง 8:รูปนักเรียน
// Teachers columns:  1:ลำดับ 2:รหัส 3:คำนำหน้า 4:ชื่อ 5:นามสกุล 6:วิชา 7:ห้องประจำ
// Issues columns:    1:ID 2:รหัสนักเรียน 3:ชื่อ 4:Serial 5:ประเภท 6:รายละเอียด
//                    7:ไฟล์แนบ 8:สถานะเดิม 9:สถานะปัญหา 10:วันที่แจ้ง
//                    11:หมายเหตุAdmin 12:อัพเดทล่าสุด 13:กำหนดส่ง
// Accounts columns:  1:Type 2:Code 3:Username 4:Password 5:Email 6:สร้างเมื่อ 7:Loginล่าสุด

function doGet(e) {
  if (e && e.parameter) {
    // Service worker — enables Android Chrome standalone mode
    if (e.parameter.sw === '1') {
      var sw = 'self.addEventListener("install",function(){self.skipWaiting();});' +
               'self.addEventListener("activate",function(ev){ev.waitUntil(clients.claim());});' +
               'self.addEventListener("fetch",function(){});';
      return ContentService.createTextOutput(sw).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    // SVG icon served as real HTTPS URL (no auth required)
    if (e.parameter.icon === '1') {
      var svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">' +
        '<rect width="512" height="512" rx="110" fill="#4F7EF7"/>' +
        '<rect x="120" y="72" width="272" height="368" rx="30" fill="none" stroke="#fff" stroke-width="28"/>' +
        '<circle cx="256" cy="410" r="22" fill="#fff"/>' +
        '<rect x="184" y="104" width="144" height="14" rx="7" fill="#fff" opacity=".5"/>' +
        '</svg>';
      return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.XML);
    }
    // Web App Manifest
    if (e.parameter.manifest === '1') {
      var baseUrl = ScriptApp.getService().getUrl();
      var iconUrl = baseUrl + '?icon=1';
      var manifest = {
        name: 'ระบบบริหารจัดการไอแพด',
        short_name: 'iPad JV',
        description: 'โรงเรียนจักราชวิทยา — ระบบบริหารจัดการไอแพด',
        start_url: baseUrl,
        scope: baseUrl,
        display: 'standalone',
        orientation: 'any',
        background_color: '#F0F4FF',
        theme_color: '#4F7EF7',
        icons: [
          { src: iconUrl, sizes: '192x192', type: 'image/svg+xml', purpose: 'any maskable' },
          { src: iconUrl, sizes: '512x512', type: 'image/svg+xml', purpose: 'any maskable' }
        ]
      };
      return ContentService.createTextOutput(JSON.stringify(manifest))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  var tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.initRmt = (e && e.parameter && e.parameter._rmt) ? String(e.parameter._rmt) : '';
  // Device ID for iOS auto-login (no cross-origin communication needed)
  tmpl.initDeviceId = (e && e.parameter && e.parameter._devid) ? String(e.parameter._devid).trim().substring(0, 64) : '';
  tmpl.initSection = (e && e.parameter && e.parameter.section) ? String(e.parameter.section).trim().substring(0, 40) : '';
  return tmpl.evaluate()
    .setTitle('ระบบบริหารจัดการไอแพด')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }
function ss()       { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(name)   { return ss().getSheetByName(name); }

// ── Cache ─────────────────────────────────────────────────────────────────────

function getCache_()        { return CacheService.getScriptCache(); }
function invalidateCache_() {
  try { getCache_().removeAll(['dashboard_stats', 'all_ipads']); } catch(e) {}
}

// ── Auth ──────────────────────────────────────────────────────────────────────

function getCurrentUser(tok) {
  // Check token-based session first (for password/OTP admins on "Anyone" deployment)
  if (tok) {
    try {
      const cached = CacheService.getScriptCache().get('_sess_' + tok);
      if (cached) return JSON.parse(cached);
    } catch(e) {}
  }
  const email = String(Session.getActiveUser().getEmail() || '').trim();
  if (!email) return { email: '', name: 'Guest', code: '', role: CONFIG.ROLES.USER, position: '' };
  const data = sh(CONFIG.SHEETS.USERS).getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase())
      return { email: data[i][0], name: data[i][1], code: data[i][2], role: data[i][3], position: data[i][4] };
  }
  sh(CONFIG.SHEETS.USERS).appendRow([email, email.split('@')[0], '', CONFIG.ROLES.USER, '']);
  log_(email, 'ลงทะเบียนใหม่', 'Auto-registered');
  return { email, name: email.split('@')[0], code: '', role: CONFIG.ROLES.USER, position: '' };
}

function createSession_(user) {
  const tok = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('_sess_' + tok, JSON.stringify(user), 28800);
  return tok;
}

function createRememberToken_(user) {
  const tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const expires = Date.now() + (30 * 24 * 60 * 60 * 1000);
  PropertiesService.getScriptProperties().setProperty('_remember_' + tok, JSON.stringify({
    user: user,
    expires: expires
  }));
  return tok;
}

function consumeRememberToken(token) {
  if (!token) return err('No remembered login');
  try {
    const key = '_remember_' + String(token).trim();
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return err('Remembered login expired');
    const data = JSON.parse(raw);
    if (!data || !data.user || !data.expires || Date.now() > Number(data.expires)) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      return err('Remembered login expired');
    }
    if (data.user.role === CONFIG.ROLES.USER) return err('Invalid remembered login');
    const fresh = createSession_(data.user);
    return ok('Remembered login restored', Object.assign({}, data.user, { token: fresh, rememberToken: token }));
  } catch(e) {
    return err('Remembered login failed: ' + e.message);
  }
}

function forgetRememberToken(token) {
  if (token) {
    try { PropertiesService.getScriptProperties().deleteProperty('_remember_' + String(token).trim()); } catch(e) {}
  }
  return ok('Remembered login removed');
}

function createRememberForCurrentSession(tok) {
  const user = getCurrentUser(tok);
  if (!user || user.role === CONFIG.ROLES.USER) return err('Admin session required');
  return ok('Remembered login saved', { rememberToken: createRememberToken_(user) });
}

// ── Device-based auto-login (iOS PWA — no postMessage required) ───────────────

function registerDeviceRemember(deviceId, tok) {
  if (!deviceId || String(deviceId).length < 8) return err('Invalid device ID');
  const user = getCurrentUser(tok);
  if (!user || user.role === CONFIG.ROLES.USER) return err('Admin session required');
  const expires = Date.now() + (30 * 24 * 60 * 60 * 1000);
  const key = '_dev_' + String(deviceId).trim().substring(0, 64);
  try {
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify({ user: user, expires: expires }));
  } catch(e) { return err('Register failed: ' + e.message); }
  return ok('Device registered for auto-login');
}

function autoLoginByDevice(deviceId) {
  if (!deviceId) return err('No device ID');
  try {
    const key = '_dev_' + String(deviceId).trim().substring(0, 64);
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return err('No device remember');
    const data = JSON.parse(raw);
    if (!data || !data.user || !data.expires || Date.now() > Number(data.expires)) {
      PropertiesService.getScriptProperties().deleteProperty(key);
      return err('Device remember expired');
    }
    if (data.user.role === CONFIG.ROLES.USER) return err('Invalid device remember');
    const fresh = createSession_(data.user);
    return ok('Auto-login successful', Object.assign({}, data.user, { token: fresh }));
  } catch(e) {
    return err('Device auto-login failed: ' + e.message);
  }
}

function forgetDeviceRemember(deviceId) {
  if (deviceId) {
    try { PropertiesService.getScriptProperties().deleteProperty('_dev_' + String(deviceId).trim().substring(0, 64)); } catch(e) {}
  }
  return ok('Device remember cleared');
}

// ── iPad CRUD ─────────────────────────────────────────────────────────────────

function getAllIPads() {
  const cache  = getCache_();
  const cached = cache.get('all_ipads');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  // Build photo map from Students sheet: personCode → photoUrl
  const photoMap = {};
  const stuSheet = sh(CONFIG.SHEETS.STUDENTS);
  if (stuSheet && stuSheet.getLastRow() >= 2) {
    stuSheet.getRange(2, 2, stuSheet.getLastRow() - 1, 7).getValues().forEach(r => {
      const code = String(r[0] || '').trim();
      const url  = String(r[6] || '').trim();
      if (code && url) photoMap[code] = url;
    });
  }

  const data = sh(CONFIG.SHEETS.IPAD).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[0] === '' && r[1] === '') continue;
    const personCode = String(r[6] || '').trim();
    rows.push({
      row: i + 1, id: r[0], serial: String(r[1]), assetCode: r[2],
      prefix: r[3], firstName: r[4], lastName: r[5], personCode,
      position: r[7], grade: r[8], room: r[9], status: r[10],
      borrowDate: r[11] ? fmt(r[11]) : '',
      returnDate: r[12] ? fmt(r[12]) : '',
      notes: r[13] || '',
      photoUrl: photoMap[personCode] || ''
    });
  }
  try { cache.put('all_ipads', JSON.stringify(rows), 30); } catch(e) {}
  return rows;
}

function addIPad(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์เพิ่มข้อมูล');
  if (!/^[A-Z0-9]{10}$/.test(String(d.serial).toUpperCase())) return err('Serial ต้องเป็นตัวอักษร/ตัวเลข 10 ตัว');
  const sheet = sh(CONFIG.SHEETS.IPAD);
  const newId = d.id || nextIPadId_(sheet);
  sheet.appendRow([newId, d.serial, d.assetCode, d.prefix, d.firstName, d.lastName,
    d.personCode, d.position, d.grade || '', d.room || '', CONFIG.STATUS.AVAILABLE, '', '', d.notes || '']);
  invalidateCache_();
  log_(u.email, 'เพิ่มไอแพด', `Serial: ${d.serial}`);
  return ok('เพิ่มข้อมูลสำเร็จ', { id: newId });
}

function nextIPadId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const nums = ids.map(v => Number(v)).filter(n => Number.isFinite(n));
  return nums.length ? Math.max(...nums) + 1 : lastRow;
}

function updateIPad(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์แก้ไข');
  if (!/^[A-Z0-9]{10}$/.test(String(d.serial).toUpperCase())) return err('Serial ต้องเป็นตัวอักษร/ตัวเลข 10 ตัว');
  const sheet = sh(CONFIG.SHEETS.IPAD);
  sheet.getRange(row, 2, 1, 9).setValues([[d.serial, d.assetCode, d.prefix, d.firstName, d.lastName,
    d.personCode, d.position, d.grade || '', d.room || '']]);
  sheet.getRange(row, 14).setValue(d.notes || '');
  invalidateCache_();
  log_(u.email, 'แก้ไขไอแพด', `Serial: ${d.serial}`);
  return ok('แก้ไขสำเร็จ');
}

function updateIPadStatus(row, newStatus, notes) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์เปลี่ยนสถานะ');
  const sheet = sh(CONFIG.SHEETS.IPAD);
  const now   = new Date();
  const S     = CONFIG.STATUS;
  sheet.getRange(row, 11).setValue(newStatus);
  if (newStatus === S.BORROWED) {
    sheet.getRange(row, 12).setValue(now); sheet.getRange(row, 13).setValue('');
  } else if (newStatus === S.RETURNED || newStatus === S.CLAIMED) {
    if (!sheet.getRange(row, 12).getValue()) sheet.getRange(row, 12).setValue(now);
    sheet.getRange(row, 13).setValue(now);
  } else if (newStatus === S.AVAILABLE) {
    sheet.getRange(row, 12).setValue(''); sheet.getRange(row, 13).setValue('');
  }
  if (notes) {
    const prev = sheet.getRange(row, 14).getValue();
    sheet.getRange(row, 14).setValue(prev ? prev + '\n' + notes : notes);
  }
  const serial = sheet.getRange(row, 2).getValue();
  invalidateCache_();
  log_(u.email, `เปลี่ยนสถานะ → ${newStatus}`, `Serial: ${serial}`);
  return ok('อัพเดทสถานะสำเร็จ');
}

function deleteIPad(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const serial = sh(CONFIG.SHEETS.IPAD).getRange(row, 2).getValue();
  sh(CONFIG.SHEETS.IPAD).deleteRow(row);
  invalidateCache_();
  log_(u.email, 'ลบไอแพด', `Serial: ${serial}`);
  return ok('ลบสำเร็จ');
}

// ── Search ────────────────────────────────────────────────────────────────────

function searchIPads(f) {
  let result = getAllIPads();
  if (f.query) {
    const q = f.query.toLowerCase();
    result = result.filter(r =>
      String(r.id).includes(q) || r.serial.includes(q) ||
      r.assetCode.toLowerCase().includes(q) ||
      r.firstName.toLowerCase().includes(q) || r.lastName.toLowerCase().includes(q) ||
      r.personCode.toLowerCase().includes(q)
    );
  }
  if (f.status   && f.status   !== 'all') result = result.filter(r => r.status   === f.status);
  if (f.position && f.position !== 'all') result = result.filter(r => r.position === f.position);
  if (f.grade    && f.grade    !== 'all') result = result.filter(r => String(r.grade) === f.grade);
  if (f.room     && f.room     !== 'all') result = result.filter(r => String(r.room)  === f.room);
  return result;
}

// ── Dashboard Stats ───────────────────────────────────────────────────────────

function getDashboardStats() {
  const cache  = getCache_();
  const cached = cache.get('dashboard_stats');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  const ipads = getAllIPads();
  const S = CONFIG.STATUS;
  const borrowed = ipads.filter(r => r.status === S.BORROWED).length;

  // True total from Database master sheet
  const dbSheet_ = findDatabaseSheet_();
  let totalFromDB = 0;
  if (dbSheet_ && dbSheet_.getLastRow() >= 2) {
    const dbVals = dbSheet_.getRange(1, 1, dbSheet_.getLastRow(), 1).getValues().slice(1);
    totalFromDB = dbVals.filter(r => String(r[0]).trim()).length;
  }
  const totalReal = totalFromDB || ipads.length;

  const inProgressStatuses = [S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM];
  const unavailableStatuses = inProgressStatuses.concat([S.BORROWED, S.CLAIMED]);
  const inProgressCount = ipads.filter(r => inProgressStatuses.includes(r.status)).length;
  const stats = {
    total:            totalReal,
    available:        Math.max(0, totalReal - ipads.filter(r => unavailableStatuses.includes(r.status)).length),
    borrowed:         borrowed,
    returned:         ipads.filter(r => r.status === S.RETURNED).length,
    claimed:          ipads.filter(r => r.status === S.CLAIMED).length,
    inProgress:       inProgressCount,
    pendingIssue:     ipads.filter(r => r.status === S.PENDING_ISSUE).length,
    underInspection:  ipads.filter(r => r.status === S.UNDER_INSPECTION).length,
    defective:        ipads.filter(r => r.status === S.DEFECTIVE).length,
    underClaim:       ipads.filter(r => r.status === S.UNDER_CLAIM).length,
    teachers:         ipads.filter(r => r.position === 'ครู').length,
    students:         ipads.filter(r => r.position === 'นักเรียน').length,
    gradeMap: {}
  };
  ipads.filter(r => r.position === 'นักเรียน' && r.grade).forEach(r => {
    const g = String(r.grade), rm = String(r.room);
    if (!stats.gradeMap[g])     stats.gradeMap[g]     = {};
    if (!stats.gradeMap[g][rm]) stats.gradeMap[g][rm] = { total:0, available:0, borrowed:0, returned:0, claimed:0, inProgress:0 };
    stats.gradeMap[g][rm].total++;
    if      (r.status === S.AVAILABLE)                   stats.gradeMap[g][rm].available++;
    else if (r.status === S.BORROWED)                    stats.gradeMap[g][rm].borrowed++;
    else if (r.status === S.RETURNED)                    stats.gradeMap[g][rm].returned++;
    else if (r.status === S.CLAIMED)                     stats.gradeMap[g][rm].claimed++;
    else if (inProgressStatuses.includes(r.status))      stats.gradeMap[g][rm].inProgress++;
  });
  try { cache.put('dashboard_stats', JSON.stringify(stats), 60); } catch(e) {}
  return stats;
}

// ── User Management ───────────────────────────────────────────────────────────

function getAllUsers() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const data = sh(CONFIG.SHEETS.USERS).getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map((r, i) => ({
    row: i + 2, email: r[0], name: r[1], code: r[2], role: r[3], position: r[4]
  }));
}

function addUser(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const data = sh(CONFIG.SHEETS.USERS).getDataRange().getValues();
  if (data.slice(1).some(r => r[0] === d.email)) return err('Email นี้มีในระบบแล้ว');
  sh(CONFIG.SHEETS.USERS).appendRow([d.email, d.name, d.code, d.role, d.position]);
  log_(u.email, 'เพิ่มผู้ใช้', d.email);
  return ok('เพิ่มผู้ใช้สำเร็จ');
}

function updateUser(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  sh(CONFIG.SHEETS.USERS).getRange(row, 1, 1, 5).setValues([[d.email, d.name, d.code, d.role, d.position]]);
  log_(u.email, 'แก้ไขผู้ใช้', d.email);
  return ok('แก้ไขสำเร็จ');
}

function deleteUser(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const email = sh(CONFIG.SHEETS.USERS).getRange(row, 1).getValue();
  if (email === u.email) return err('ไม่สามารถลบบัญชีตัวเองได้');
  sh(CONFIG.SHEETS.USERS).deleteRow(row);
  log_(u.email, 'ลบผู้ใช้', email);
  return ok('ลบสำเร็จ');
}

// ── Link Admin Google Email ───────────────────────────────────────────────────

function linkAdminEmail(email) {
  const tok = arguments[arguments.length - 1];
  const u   = getCurrentUser(tok);
  if (!u || u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');

  const requestedEmail = String(email || '').trim();
  if (!requestedEmail) return err('กรุณากรอก email');
  if (!requestedEmail.toLowerCase().endsWith('@jv.ac.th'))
    return err('กรุณาใช้อีเมลโรงเรียน (@jv.ac.th) เท่านั้น');

  const usersSheet = sh(CONFIG.SHEETS.USERS);
  const data       = usersSheet.getDataRange().getValues();
  const reqLow     = requestedEmail.toLowerCase();

  // Block if another non-local entry already has this email
  const duplicate = data.slice(1).find(r =>
    String(r[0]).trim().toLowerCase() === reqLow &&
    !String(r[0]).trim().toLowerCase().endsWith('@admin.local')
  );
  if (duplicate) return err('อีเมลนี้มีในระบบแล้ว กรุณาติดต่อผู้ดูแล');

  // Find this admin's row (by current email, which may be @admin.local or already real)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === u.email.toLowerCase()) {
      usersSheet.getRange(i + 1, 1).setValue(requestedEmail);
      // Update session cache so further calls use the new email
      const updated = Object.assign({}, u, { email: requestedEmail });
      try { CacheService.getScriptCache().put('_sess_' + tok, JSON.stringify(updated), 28800); } catch(e) {}
      log_(requestedEmail, 'เชื่อมอีเมลโรงเรียน', u.email + ' → ' + requestedEmail);
      return ok('เชื่อมอีเมลสำเร็จ ครั้งต่อไปสามารถ login ด้วย OTP ได้เลย', { email: requestedEmail });
    }
  }
  return err('ไม่พบบัญชีของท่านในระบบ');
}

// ── Admin Account Management ──────────────────────────────────────────────────

function getAdminAccounts() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  ensureDefaultAdminAccount_();
  const sheet = sh(CONFIG.SHEETS.ADMIN);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[0]).map((r, i) => ({
    row: i + 2, username: String(r[0]), role: String(r[2] || CONFIG.ROLES.SUPER), email: String(r[3] || '')
  }));
}

function addAdminAccount(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  if (!d.username || !d.password) return err('กรุณากรอก Username และ Password');
  ensureDefaultAdminAccount_();
  const sheet = sh(CONFIG.SHEETS.ADMIN);
  const data = sheet.getDataRange().getValues();
  if (data.slice(1).some(r => String(r[0]).trim().toLowerCase() === String(d.username).trim().toLowerCase()))
    return err('Username นี้มีในระบบแล้ว');
  sheet.appendRow([d.username.trim(), d.password, d.role || CONFIG.ROLES.ADMIN, d.email || '']);
  log_(u.email, 'เพิ่มบัญชีแอดมิน', d.username);
  return ok('เพิ่มบัญชีแอดมินสำเร็จ');
}

function updateAdminAccount(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const sheet = sh(CONFIG.SHEETS.ADMIN);
  const existing = String(sheet.getRange(row, 1).getValue()).trim();
  const newPass = d.password || String(sheet.getRange(row, 2).getValue());
  const newUser = d.username || existing;
  const existingEmail = String(sheet.getRange(row, 4).getValue());
  sheet.getRange(row, 1, 1, 4).setValues([[newUser, newPass, d.role || CONFIG.ROLES.ADMIN, d.email !== undefined ? d.email : existingEmail]]);
  log_(u.email, 'แก้ไขบัญชีแอดมิน', newUser);
  return ok('แก้ไขสำเร็จ');
}

function deleteAdminAccount(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const sheet = sh(CONFIG.SHEETS.ADMIN);
  const username = String(sheet.getRange(row, 1).getValue()).trim();
  if (username === 'admin') return err('ไม่สามารถลบบัญชี admin หลักได้');
  sheet.deleteRow(row);
  log_(u.email, 'ลบบัญชีแอดมิน', username);
  return ok('ลบสำเร็จ');
}

// ── Activity Log ──────────────────────────────────────────────────────────────

function getActivityLog() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return [];
  const data = sh(CONFIG.SHEETS.LOG).getDataRange().getValues();
  const result = [];
  for (let i = data.length - 1; i >= 1 && result.length < 300; i--) {
    if (data[i][0]) result.push({
      date: Utilities.formatDate(new Date(data[i][0]), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'),
      email: data[i][1], action: data[i][2], details: data[i][3]
    });
  }
  return result;
}

// ── Export ────────────────────────────────────────────────────────────────────

function exportCSV() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const headers = ['ID','Serial','รหัสครุภัณฑ์','คำนำหน้า','ชื่อ','นามสกุล',
                   'รหัส','ตำแหน่ง','ระดับชั้น','ห้อง','สถานะ','วันที่ยืม','วันที่คืน','หมายเหตุ'];
  const rows = [headers.join(',')];
  getAllIPads().forEach(r => {
    const fields = [r.id, r.serial, r.assetCode, r.prefix, r.firstName, r.lastName,
      r.personCode, r.position, r.grade, r.room, r.status, r.borrowDate, r.returnDate, r.notes || ''];
    rows.push(fields.map(c => {
      const s = String(c == null ? '' : c).replace(/"/g, '""');
      return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s}"` : s;
    }).join(','));
  });
  return ok('', { csv: rows.join('\n') });
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function fmt(d)        { return Utilities.formatDate(new Date(d), 'Asia/Bangkok', 'dd/MM/yyyy'); }
function fmtTs(d)      { return Utilities.formatDate(new Date(d), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm'); }
function ok(msg, ext)  { return Object.assign({ success: true,  message: msg }, ext || {}); }
function err(msg)      { return { success: false, message: msg }; }
function log_(email, action, details) {
  try { sh(CONFIG.SHEETS.LOG).appendRow([new Date(), email, action, details]); } catch(e) {}
}

function getAdminEmails_() {
  try {
    const data = sh(CONFIG.SHEETS.USERS).getDataRange().getValues();
    return data.slice(1)
      .filter(r => r[3] === CONFIG.ROLES.ADMIN || r[3] === CONFIG.ROLES.SUPER)
      .map(r => r[0]).filter(e => e);
  } catch(e) { return []; }
}

function notifyAdmins_(subject, body) {
  const emails = getAdminEmails_();
  if (!emails.length) return;
  try { MailApp.sendEmail(emails.join(','), subject, body); } catch(e) {
    log_('system', 'แจ้งเตือน admin ล้มเหลว', e.message);
  }
}

// ── Admin Auth ────────────────────────────────────────────────────────────────

function testPushNotification() {
  return debugPushNotification();
}

function debugPushNotification() {
  const props = PropertiesService.getScriptProperties();
  const appId  = String(props.getProperty('ONESIGNAL_APP_ID')       || '').trim();
  const apiKey = String(props.getProperty('ONESIGNAL_REST_API_KEY') || '').trim();
  const pwaUrl = String(props.getProperty('IPAD_JV_PWA_URL')        || '').trim();

  const report = {
    ONESIGNAL_APP_ID:       appId  ? appId.slice(0,8)+'...' : '❌ ไม่ได้ตั้งค่า',
    ONESIGNAL_REST_API_KEY: apiKey ? apiKey.slice(0,8)+'...' : '❌ ไม่ได้ตั้งค่า',
    IPAD_JV_PWA_URL:        pwaUrl || '(ใช้ค่า default)',
  };

  if (!appId || !apiKey) {
    report.result = '❌ ไม่มี API key — หยุดที่นี่';
    Logger.log(JSON.stringify(report, null, 2));
    return report;
  }

  // ── ดู subscribers ──
  try {
    const subRes = UrlFetchApp.fetch(
      'https://onesignal.com/api/v1/apps/' + appId,
      { headers: { Authorization: 'Basic ' + apiKey }, muteHttpExceptions: true }
    );
    const subJson = JSON.parse(subRes.getContentText());
    report.app_name          = subJson.name || '?';
    report.subscribers_total = subJson.players || 0;
  } catch(e) {
    report.subscribers_error = e.message;
  }

  // ── ส่ง test notification ──
  const payload = {
    app_id: appId,
    target_channel: 'push',
    filters: [{ field: 'tag', key: 'role', relation: '=', value: 'admin' }],
    headings: { en: '🔔 ทดสอบ', th: '🔔 ทดสอบ' },
    contents: { en: 'ทดสอบการแจ้งเตือนจาก iPad JV', th: 'ทดสอบการแจ้งเตือนจาก iPad JV' },
    url: pwaUrl || 'https://wisanu15.github.io/IPAD_JV/',
    data: {}
  };
  try {
    const res = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    report.http_status  = res.getResponseCode();
    report.recipients   = json.recipients;
    report.notification_id = json.id || '?';
    report.api_errors   = json.errors || null;
    report.result = (json.recipients > 0) ? '✅ ส่งสำเร็จ' : '⚠️ ส่งแล้วแต่ไม่มี subscriber ตรงเงื่อนไข (role=admin)';
  } catch(e) {
    report.result = '❌ ส่งล้มเหลว: ' + e.message;
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function notifyAdminPush_(title, message, section, data) {
  try {
    const props = PropertiesService.getScriptProperties();
    const appId = String(props.getProperty('ONESIGNAL_APP_ID') || '').trim();
    const apiKey = String(props.getProperty('ONESIGNAL_REST_API_KEY') || '').trim();
    if (!appId || !apiKey) return;

    const pwaBase = String(props.getProperty('IPAD_JV_PWA_URL') || 'https://wisanu15.github.io/IPAD_JV/');
    const sep = pwaBase.includes('?') ? '&' : '?';
    const url = section
      ? pwaBase + sep + 'embed=1&section=' + encodeURIComponent(section)
      : pwaBase + sep + 'embed=1';

    const payload = {
      app_id: appId,
      target_channel: 'push',
      filters: [{ field: 'tag', key: 'role', relation: '=', value: 'admin' }],
      headings: { en: title || 'iPad JV', th: title || 'iPad JV' },
      contents: { en: message || '', th: message || '' },
      url: url,
      data: data || {}
    };

    const res = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      log_('system', 'Push notification failed', code + ': ' + res.getContentText());
    }
  } catch(e) {
    log_('system', 'Push notification error', e.message);
  }
}

function adminLogin(user, pass, remember) {
  ensureDefaultAdminAccount_();
  const sheet = sh(CONFIG.SHEETS.ADMIN);
  if (!sheet) return err('ไม่พบ Sheet Admin');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(user).trim() &&
        String(data[i][1]).trim() === String(pass).trim()) {
      // Use Google email if available, otherwise fall back to username-based identifier
      const googleEmail = String(Session.getActiveUser().getEmail() || '').trim();
      const email = googleEmail || (String(user).trim().toLowerCase() + '@admin.local');
      const role = String(data[i][2] || CONFIG.ROLES.SUPER).trim();
      const current = upsertAdminUserByEmail_(email, user, role);
      const tok = createSession_(current);
      const rememberTok = remember ? createRememberToken_(current) : '';
      log_(email, 'Admin Password Login', user);
      return ok('เข้าสู่ระบบสำเร็จ', Object.assign({}, current, { token: tok, rememberToken: rememberTok || '' }));
    }
  }
  return err('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
}

function ensureDefaultAdminAccount_() {
  let sheet = sh(CONFIG.SHEETS.ADMIN);
  if (!sheet) {
    sheet = ss().insertSheet(CONFIG.SHEETS.ADMIN);
    sheet.appendRow(['user', 'รหัส', 'role']);
    try { sheet.getRange('1:1').setFontWeight('bold').setBackground('#6a1b9a').setFontColor('white'); sheet.setFrozenRows(1); } catch(e) {}
    sheet.appendRow(['admin', '112233', CONFIG.ROLES.SUPER]);
    return;
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'admin') {
      if (!String(data[i][1]).trim() || String(data[i][1]).trim() === '1234') {
        sheet.getRange(i + 1, 2).setValue('112233');
      }
      if (!String(data[i][2]).trim()) sheet.getRange(i + 1, 3).setValue(CONFIG.ROLES.SUPER);
      return;
    }
  }
  sheet.appendRow(['admin', '112233', CONFIG.ROLES.SUPER]);
}

function upsertAdminUserByEmail_(email, name, role) {
  const usersSheet = sh(CONFIG.SHEETS.USERS);
  const data = usersSheet.getDataRange().getValues();
  const normalized = String(email).trim().toLowerCase();
  const isLocalEmail = normalized.endsWith('@admin.local');
  const displayName = name === 'admin' ? 'Admin' : name;
  // Generate a stable code per admin username
  const adminCode = name === 'admin' ? 'ADMIN001' : ('ADM_' + String(name).toUpperCase().replace(/[^A-Z0-9]/g, '').substr(0, 8));
  const user = {
    email,
    name: displayName,
    code: adminCode,
    role: role || CONFIG.ROLES.SUPER,
    position: 'ผู้ดูแลระบบ'
  };
  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][0]).trim().toLowerCase();
    const rowCode  = String(data[i][2]).trim();
    // Match by exact email, OR by code when using local email (prevents duplicates)
    const matched = rowEmail === normalized || (isLocalEmail && rowCode === adminCode);
    if (!matched) continue;
    // If we now have a real Google email and the stored one was local, upgrade it
    const storedEmail = (!isLocalEmail && rowEmail.endsWith('@admin.local')) ? email : data[i][0];
    usersSheet.getRange(i + 1, 1, 1, 5).setValues([[
      storedEmail,
      data[i][1] || user.name,
      data[i][2] || user.code,
      user.role,
      data[i][4] || user.position
    ]]);
    return {
      email: storedEmail,
      name: data[i][1] || user.name,
      code: data[i][2] || user.code,
      role: user.role,
      position: data[i][4] || user.position
    };
  }
  usersSheet.appendRow([user.email, user.name, user.code, user.role, user.position]);
  return user;
}

// ── Registration ──────────────────────────────────────────────────────────────

function getGradesRooms() {
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const map = {};
  data.slice(1).forEach(r => {
    if (r[5] && r[6]) map[`${r[5]}|${r[6]}`] = { grade: String(r[5]), room: String(r[6]) };
  });
  return Object.values(map).sort((a, b) =>
    a.grade !== b.grade ? (a.grade < b.grade ? -1 : 1) : Number(a.room) - Number(b.room)
  );
}

function getStudentsForRegistration(grade, room) {
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[1] && !r[3]) continue;
    if (String(r[5]) !== String(grade) || String(r[6]) !== String(room)) continue;
    students.push({ row: i + 1, seq: r[0], code: String(r[1]), prefix: r[2], firstName: r[3], lastName: r[4], grade: r[5], room: r[6] });
  }
  const assignedCodes = {};
  getAllIPads().filter(r => r.personCode).forEach(r => { assignedCodes[String(r.personCode)] = r.serial; });
  return students.map(s => ({ ...s, hasIPad: !!assignedCodes[String(s.code)], assignedSerial: assignedCodes[String(s.code)] || '' }));
}

function findDatabaseSheet_() {
  let s = sh(CONFIG.SHEETS.DATABASE);
  if (s) return s;
  const all = ss().getSheets();
  for (let i = 0; i < all.length; i++) {
    if (String(all[i].getName()).trim().toLowerCase().indexOf('database') >= 0) return all[i];
  }
  return null;
}

function getAvailableSerials(cachedIpads) {
  const ipads = cachedIpads || getAllIPads();
  const S = CONFIG.STATUS;
  const unavailable = new Set([S.BORROWED, S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM, S.CLAIMED]);
  const assigned = {}, assetBySerial = {};
  ipads.forEach(r => {
    if (unavailable.has(r.status)) assigned[r.serial] = true;
    if (r.assetCode) assetBySerial[r.serial] = r.assetCode;
  });
  const dbSheet = findDatabaseSheet_();
  const masterSerials = [];
  if (dbSheet && dbSheet.getLastRow() >= 2 && dbSheet.getLastColumn() >= 1) {
    const data = dbSheet.getRange(1, 1, dbSheet.getLastRow(), dbSheet.getLastColumn()).getValues();
    const dataRows = data.slice(1); // skip header row
    dataRows.forEach(row => {
      const s = String(row[0] == null ? '' : row[0]).trim();
      if (s) masterSerials.push(s);
    });
    if (!masterSerials.length) {
      dataRows.forEach(row => row.forEach(cell => {
        const s = String(cell == null ? '' : cell).trim();
        if (/^[A-Za-z0-9]{6,}$/.test(s)) masterSerials.push(s);
      }));
    }
  }
  if (!masterSerials.length) {
    return ipads.filter(r => !unavailable.has(r.status))
                .map(r => ({ serial: r.serial, assetCode: r.assetCode || '' }));
  }
  const seen = {}, out = [];
  masterSerials.forEach(s => {
    if (seen[s]) return; seen[s] = true;
    if (assigned[s]) return;
    out.push({ serial: s, assetCode: assetBySerial[s] || '' });
  });
  return out;
}

function debugAvailableSerials() {
  const dbSheet = findDatabaseSheet_();
  const info = {
    เจอชีตDatabase: !!dbSheet, ชื่อชีตที่เจอ: dbSheet ? dbSheet.getName() : '(ไม่เจอ)',
    จำนวนแถว: dbSheet ? dbSheet.getLastRow() : 0, จำนวนคอลัมน์: dbSheet ? dbSheet.getLastColumn() : 0,
    ชื่อชีตทั้งหมด: ss().getSheets().map(s => s.getName())
  };
  const result = getAvailableSerials();
  info.ซีเรียลว่างที่ได้ = result.length;
  info.ตัวอย่าง5ตัวแรก = result.slice(0, 5).map(r => r.serial);
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function registerIPadBatch(registrations, grade, room) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์ลงทะเบียนไอแพด');
  return registerIPadBatch_(registrations, grade, room, u.email || 'visitor');
}

function registerIPadBatchMultiple(classes) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์ลงทะเบียนไอแพด');
  const userEmail = u.email || 'visitor';
  let totalCount = 0;
  const results = [], allErrors = [];
  for (const cls of classes) {
    if (!cls.items || !cls.items.length) continue;
    const r = registerIPadBatch_(cls.items, cls.grade, cls.room, userEmail);
    results.push(`${cls.grade}/${cls.room}: ${r.message}`);
    if (!r.success) { allErrors.push(`${cls.grade}/${cls.room}: ${r.message}`); }
    else { const m = r.message.match(/(\d+)/); if (m) totalCount += Number(m[1]); }
  }
  if (allErrors.length && !totalCount) return err(allErrors.join('\n'));
  return ok(`ลงทะเบียนสำเร็จรวม ${totalCount} เครื่อง จาก ${classes.length} ห้อง` +
    (allErrors.length ? `\n⚠️ บางห้องมีข้อผิดพลาด` : ''), { results });
}

function registerIPadBatch_(registrations, grade, room, userEmail) {
  const sheet = sh(CONFIG.SHEETS.IPAD);
  const data  = sheet.getDataRange().getValues();
  const serialRowMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][1]) serialRowMap[String(data[i][1])] = i + 1;
  }
  let count = 0;
  const errors = [], manualEntries = [];
  registrations.forEach(reg => {
    if (!reg.serial) return;
    const serial = String(reg.serial).trim().toUpperCase();
    if (!/^[A-Z0-9]{10}$/.test(serial)) { errors.push(`Serial "${serial}" ต้องเป็น 10 ตัว`); return; }
    if (reg.isManualStudent || reg.isManualSerial) {
      manualEntries.push({ serial, firstName: reg.firstName, lastName: reg.lastName,
        isManualStudent: !!reg.isManualStudent, isManualSerial: !!reg.isManualSerial });
    }
    const existingRow = serialRowMap[serial];
    if (existingRow) {
      const existingCode = data[existingRow - 1][6];
      if (existingCode && String(existingCode) !== String(reg.studentCode)) {
        errors.push(`Serial ${serial} ถูกมอบหมายแล้ว`); return;
      }
      sheet.getRange(existingRow, 4, 1, 7).setValues([[reg.prefix, reg.firstName, reg.lastName, reg.studentCode || '', reg.position || 'นักเรียน', grade, room]]);
      if (reg.assetCode && !data[existingRow - 1][2]) sheet.getRange(existingRow, 3).setValue(reg.assetCode);
      if (reg.notes && !data[existingRow - 1][13]) sheet.getRange(existingRow, 14).setValue(reg.notes);
      sheet.getRange(existingRow, 11).setValue(CONFIG.STATUS.BORROWED);
      if (!data[existingRow - 1][11]) sheet.getRange(existingRow, 12).setValue(new Date());
    } else {
      const lastRow = sheet.getLastRow();
      const newId   = lastRow > 1 ? Number(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
      sheet.appendRow([newId, serial, reg.assetCode || '', reg.prefix, reg.firstName, reg.lastName,
        reg.studentCode || '', reg.position || 'นักเรียน', grade, room, CONFIG.STATUS.BORROWED, new Date(), '', reg.notes || '']);
    }
    count++;
  });
  if (count > 0) { invalidateCache_(); log_(userEmail, 'ลงทะเบียนไอแพด', `${grade}/${room} จำนวน ${count} เครื่อง`); }
  if (manualEntries.length) {
    const lines = manualEntries.map(e => {
      const parts = [];
      if (e.isManualStudent) parts.push('ชื่อ: ' + e.firstName + ' ' + e.lastName + ' (กรอกเอง)');
      if (e.isManualSerial)  parts.push('Serial: ' + e.serial + ' (ไม่อยู่ในรายการ)');
      return parts.join(' | ');
    });
    const ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm');
    notifyAdmins_(`[ไอแพด] แจ้งเตือน: มีการลงทะเบียนด้วยข้อมูลที่กรอกเอง`,
      `มีการลงทะเบียนไอแพดชั้น ${grade}/${room} ที่ใช้ข้อมูลนอกรายการ เมื่อ ${ts}\n\n` +
      lines.join('\n') + `\n\nกรุณาตรวจสอบที่ระบบไอแพดโรงเรียนจักราชวิทยา`);
    notifyAdminPush_(
      '⚠️ ลงทะเบียนข้อมูลนอกรายการ',
      `ชั้น ${grade}/${room} | ${lines[0]}${lines.length > 1 ? ` +อีก ${lines.length - 1} รายการ` : ''}`,
      '',
      { grade, room, manualCount: manualEntries.length }
    );
    log_(userEmail, 'แจ้งเตือน admin', `Manual entry: ${lines.join('; ')}`);
  }
  if (errors.length && !count) return err(errors.join('\n'));
  return ok(`ลงทะเบียนสำเร็จ ${count} เครื่อง` + (errors.length ? `\n⚠️ ${errors.join(', ')}` : ''));
}

function returnIpadByStudent(serial, studentCode) {
  serial = String(serial || '').trim().toUpperCase();
  if (!serial) return err('กรุณาระบุ Serial');
  const ipadSheet = sh(CONFIG.SHEETS.IPAD);
  const data = ipadSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toUpperCase() !== serial) continue;
    const status = String(data[i][10]).trim();
    if (status !== CONFIG.STATUS.BORROWED) return err('ไม่สามารถคืนได้ — สถานะปัจจุบัน: ' + status);
    if (studentCode) {
      const rowCode = String(data[i][6]).trim();
      if (rowCode && rowCode !== String(studentCode).trim()) return err('ข้อมูลนักเรียนไม่ตรงกับไอแพดเครื่องนี้');
    }
    ipadSheet.getRange(i + 1, 11).setValue(CONFIG.STATUS.RETURNED);
    ipadSheet.getRange(i + 1, 13).setValue(new Date());
    invalidateCache_();
    const name = [data[i][3], data[i][4], data[i][5]].filter(Boolean).join(' ');
    log_('student', 'คืนไอแพด', `Serial ${serial} — ${name}`);
    return ok('คืนไอแพดสำเร็จ', { serial, name });
  }
  return err('ไม่พบ Serial ' + serial + ' ในระบบ');
}

// ── Public Status ─────────────────────────────────────────────────────────────

function getPublicStatus() {
  const ipads = getAllIPads();
  const S = CONFIG.STATUS;

  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) {
    const stats = { total:0, borrowed:0, available:0, returned:0, claimed:0, availableFromDB:0 };
    return { stats, grades: {}, teacherList: [] };
  }
  const rows = sheet.getDataRange().getValues().slice(1);

  // helpers
  function normStr(s)  { return String(s).normalize('NFC').replace(/\s+/g,' ').trim(); }
  function normName(s) { return normStr(s).toLowerCase(); }

  // 1. Primary: personCode exact match
  const ipadByCode = {};
  ipads.forEach(r => {
    const c = normStr(r.personCode);
    if (c) ipadByCode[c] = r;
  });

  // 2. Fallback: firstName|lastName exact (NFC-normalized)
  const ipadByName = {};
  ipads.forEach(r => {
    const pos = normStr(r.position);
    if (!pos.includes('นักเรียน') && !pos.includes('student')) return;
    const key = normStr(r.firstName) + '|' + normStr(r.lastName);
    if (key !== '|' && !ipadByName[key]) ipadByName[key] = r;
  });

  // 3. Second fallback: firstName|lastName case-insensitive
  const ipadByNameCI = {};
  ipads.forEach(r => {
    const pos = normStr(r.position);
    if (!pos.includes('นักเรียน') && !pos.includes('student')) return;
    const key = normName(r.firstName) + '|' + normName(r.lastName);
    if (key !== '|' && !ipadByNameCI[key]) ipadByNameCI[key] = r;
  });

  // 4. Any-position fallback by name (catch iPads registered without position)
  const ipadByNameAny = {};
  ipads.forEach(r => {
    const key = normStr(r.firstName) + '|' + normStr(r.lastName);
    if (key !== '|' && !ipadByNameAny[key]) ipadByNameAny[key] = r;
  });

  // Normalize grade: strip leading "ม." so "ม.4" → "4", keeps plain "4" as "4"
  function normGrade(g) { return String(g).replace(/^ม\.\s*/, '').trim(); }

  const grades = {};
  let stuTotal = 0, stuBorrowed = 0, stuReturned = 0;

  rows.forEach(r => {
    if (!r[3] && !r[1]) return;
    const grade = normGrade(r[5]), room = String(r[6]).trim();
    if (!grade || !room || grade === 'undefined') return;
    if (!grades[grade])       grades[grade]       = {};
    if (!grades[grade][room]) grades[grade][room]  = { total:0, borrowed:0, available:0, returned:0, claimed:0, noIpad:0, students:[] };
    const code    = normStr(r[1]);
    const nameKey = normStr(r[3]) + '|' + normStr(r[4]);
    const nameCI  = normName(r[3]) + '|' + normName(r[4]);
    const ipad    = ipadByCode[code] || ipadByName[nameKey] || ipadByNameCI[nameCI] || ipadByNameAny[nameKey];
    const status = ipad ? ipad.status : 'ไม่มีเครื่อง';
    const rm = grades[grade][room];
    rm.total++;
    stuTotal++;
    if (!ipad)                      rm.noIpad++;
    else if (status === S.BORROWED) { rm.borrowed++; stuBorrowed++; }
    else if (status === S.AVAILABLE)  rm.available++;
    else if (status === S.RETURNED) { rm.returned++; stuReturned++; }
    else if (status === S.CLAIMED)    rm.claimed++;
    else                              rm.problem = (rm.problem||0) + 1;
    rm.students.push({ code, prefix: String(r[2]), firstName: String(r[3]), lastName: String(r[4]),
      serial: ipad ? ipad.serial : '', assetCode: ipad ? ipad.assetCode : '',
      ipadId: ipad ? ipad.id : '',
      status, borrowDate: ipad ? ipad.borrowDate : '', photoUrl: String(r[7] || '') });
  });

  // True total from Database master sheet
  const dbSheet_ = findDatabaseSheet_();
  let totalFromDB = 0;
  if (dbSheet_ && dbSheet_.getLastRow() >= 2) {
    const dbVals = dbSheet_.getRange(1, 1, dbSheet_.getLastRow(), 1).getValues().slice(1);
    totalFromDB = dbVals.filter(r => String(r[0]).trim()).length;
  }
  const availCount = getAvailableSerials(ipads).length;
  const inProgressStatuses = [S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM];
  const unavailableStatuses = inProgressStatuses.concat([S.BORROWED, S.CLAIMED]);
  const borrowedFromDB = ipads.filter(r => r.status === S.BORROWED).length;
  const inProgressCount = ipads.filter(r => inProgressStatuses.includes(r.status)).length;
  const totalReal = totalFromDB || ipads.length;
  const stats = {
    total:          stuTotal,
    borrowed:       stuBorrowed,
    borrowedFromDB: borrowedFromDB,
    returned:       stuReturned,
    available:      ipads.filter(r => r.position === 'นักเรียน' && r.status === S.AVAILABLE).length,
    claimed:        ipads.filter(r => r.position === 'นักเรียน' && r.status === S.CLAIMED).length,
    availableFromDB: availCount,
    adminAvailable: Math.max(0, totalReal - ipads.filter(r => unavailableStatuses.includes(r.status)).length),
    problem:        ipads.filter(r => [S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM, S.CLAIMED].includes(r.status)).length,
    totalIPads:     ipads.length,
    totalFromDB:    totalFromDB || (availCount + ipads.filter(r => r.personCode).length),
    stuIpadBorrowed: stuBorrowed,
    stuIpadTotal:    stuTotal
  };

  // Build borrow lookup from iPad_Data (teachers)
  const tBorrowByCode = {}, tBorrowByName = {};
  ipads.filter(r => r.position === 'ครู').forEach(r => {
    const entry = { serial: r.serial, assetCode: r.assetCode || '', status: r.status, borrowDate: r.borrowDate || '' };
    if (r.personCode) tBorrowByCode[String(r.personCode)] = entry;
    tBorrowByName[(String(r.firstName) + '|' + String(r.lastName)).toLowerCase()] = entry;
  });
  // All teachers from Teachers sheet
  const tSheet = sh(CONFIG.SHEETS.TEACHERS);
  let teacherList;
  if (tSheet && tSheet.getLastRow() >= 2) {
    const tData = tSheet.getDataRange().getValues().slice(1);
    teacherList = tData.filter(r => r[3] || r[1]).map(r => {
      const code = String(r[1] || '').trim();
      const nameKey = (String(r[3]) + '|' + String(r[4])).toLowerCase();
      const borrow = (code && tBorrowByCode[code]) || tBorrowByName[nameKey];
      return {
        prefix: String(r[2] || ''), firstName: String(r[3] || ''), lastName: String(r[4] || ''),
        photoUrl: String(r[7] || ''),
        personCode: code, subject: String(r[5] || ''), room: String(r[6] || ''),
        serial: borrow ? borrow.serial : '', assetCode: borrow ? borrow.assetCode : '',
        status: borrow ? borrow.status : 'ยังไม่ยืม', borrowDate: borrow ? borrow.borrowDate : ''
      };
    });
  } else {
    teacherList = ipads.filter(r => r.position === 'ครู').map(r => ({
      prefix: r.prefix, firstName: r.firstName, lastName: r.lastName,
      personCode: r.personCode, subject: '', room: '',
      serial: r.serial, assetCode: r.assetCode, status: r.status, borrowDate: r.borrowDate
    }));
  }
  // เรียงฝ่ายบริหารขึ้นก่อน
  const mgmtP = (r) => {
    if (r.subject === 'ผู้อำนวยการโรงเรียน' || r.firstName === 'ศราวุธ') return 0;
    if (r.subject === 'รองผู้อำนวยการโรงเรียน' || r.subject === 'ฝ่ายบริหาร') return 1;
    return 2;
  };
  teacherList.sort((a, b) => mgmtP(a) - mgmtP(b));
  return { stats, grades, teacherList };
}

function searchSerial(serial) {
  const result = getAllIPads().find(r => r.serial === String(serial).trim());
  return result ? { found: true, data: result } : { found: false };
}

function saveBorrowDoc(base64Data, filename, serial, pages) {
  try {
    const folderName = 'เอกสารยืมไอแพด';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', filename);
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const u = Session.getActiveUser().getEmail() || 'visitor';
    log_(u, 'อัพโหลดเอกสารยืม', `Serial: ${serial} | Pages: ${pages} | ${file.getUrl()}`);
    return ok('อัพโหลดสำเร็จ', { url: file.getUrl() });
  } catch(e) { return err('อัพโหลดไม่สำเร็จ: ' + e.message); }
}

// ── Students ──────────────────────────────────────────────────────────────────

function getStudents(filters) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  const data = sheet.getDataRange().getValues();
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[1] && !r[3]) continue;
    rows.push({ row: i + 1, seq: r[0], code: String(r[1]), prefix: r[2], firstName: r[3], lastName: r[4], grade: r[5], room: r[6], photoUrl: String(r[7] || '') });
  }
  if (filters) {
    if (filters.grade && filters.grade !== 'all') rows = rows.filter(r => String(r.grade) === filters.grade);
    if (filters.room  && filters.room  !== 'all') rows = rows.filter(r => String(r.room)  === filters.room);
    if (filters.query) {
      const q = filters.query.toLowerCase();
      rows = rows.filter(r => r.code.includes(q) || r.firstName.toLowerCase().includes(q) || r.lastName.toLowerCase().includes(q));
    }
  }
  return rows;
}

function addStudent(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  sheet.appendRow(['', d.code, d.prefix, d.firstName, d.lastName, d.grade, d.room]);
  renumberStudents_();
  invalidateCache_();
  log_(u.email, 'เพิ่มนักเรียน', `${d.prefix}${d.firstName} ${d.lastName}`);
  return ok('เพิ่มนักเรียนสำเร็จ');
}

function updateStudent(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  sh(CONFIG.SHEETS.STUDENTS).getRange(row, 2, 1, 6).setValues([[d.code, d.prefix, d.firstName, d.lastName, d.grade, d.room]]);
  invalidateCache_();
  log_(u.email, 'แก้ไขนักเรียน', `${d.firstName} ${d.lastName}`);
  return ok('แก้ไขสำเร็จ');
}

function checkStudentClass(grade, room) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return { exists: false, count: 0 };
  const data = sheet.getDataRange().getValues();
  const count = data.slice(1).filter(r => String(r[5]) === String(grade) && String(r[6]) === String(room)).length;
  return { exists: count > 0, count };
}

function importStudents(studentsData, grade, room, overwrite) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์นำเข้าข้อมูล');
  if (!studentsData || !studentsData.length) return err('ไม่มีข้อมูลนักเรียน');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  if (overwrite) {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][5]) === String(grade) && String(data[i][6]) === String(room)) sheet.deleteRow(i + 1);
    }
  }
  const rows = studentsData.map(s => ['', s.code, s.prefix, s.firstName, s.lastName, grade, room]);
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  renumberStudents_();
  log_(u.email, 'Import นักเรียน', `${grade}/${room} จำนวน ${studentsData.length} คน`);
  return ok(`นำเข้าสำเร็จ ${studentsData.length} คน`);
}

function importMultipleClasses(classesData, overwrite) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์นำเข้าข้อมูล');
  if (!classesData || !classesData.length) return err('ไม่มีข้อมูล');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  let totalImported = 0;
  const results = [];
  classesData.forEach(cls => {
    if (!cls.students || !cls.students.length) return;
    if (overwrite) {
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][5]) === String(cls.grade) && String(data[i][6]) === String(cls.room)) sheet.deleteRow(i + 1);
      }
    }
    const rows = cls.students.map(s => ['', s.code, s.prefix, s.firstName, s.lastName, cls.grade, cls.room]);
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    totalImported += cls.students.length;
    results.push(`${cls.grade}/${cls.room} (${cls.students.length} คน)`);
  });
  renumberStudents_();
  log_(u.email, 'Import นักเรียนหลายห้อง', results.join(', '));
  return ok(`นำเข้าสำเร็จ ${totalImported} คน จาก ${results.length} ห้อง`, { results });
}

function deleteStudentClass(grade, room) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  const data  = sheet.getDataRange().getValues();
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][5]) === String(grade) && String(data[i][6]) === String(room)) { sheet.deleteRow(i + 1); deleted++; }
  }
  renumberStudents_();
  log_(u.email, 'ลบข้อมูลนักเรียน', `${grade}/${room} จำนวน ${deleted} คน`);
  return ok(`ลบสำเร็จ ${deleted} คน`);
}

function deleteStudentRow(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const name = sh(CONFIG.SHEETS.STUDENTS).getRange(row, 4).getValue();
  sh(CONFIG.SHEETS.STUDENTS).deleteRow(row);
  renumberStudents_();
  log_(u.email, 'ลบนักเรียน', name);
  return ok('ลบสำเร็จ');
}

// ── Sync iPad_Data with Students sheet ────────────────────────────────────────

function syncIpadWithStudents() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');

  const stuSheet  = sh(CONFIG.SHEETS.STUDENTS);
  const ipadSheet = sh(CONFIG.SHEETS.IPAD);
  if (!stuSheet || !ipadSheet) return err('ไม่พบ Sheet');

  const stuData  = stuSheet.getDataRange().getValues().slice(1);
  const ipadData = ipadSheet.getDataRange().getValues();

  // Build student lookup: normalized firstName|lastName → {code, grade, room}
  function norm(s) { return String(s).normalize('NFC').replace(/\s+/g,' ').trim().toLowerCase(); }
  const stuByName = {};
  const stuByCode = {};
  stuData.forEach(r => {
    const code  = String(r[1]).trim();
    const grade = String(r[5]).trim();
    const room  = String(r[6]).trim();
    const key   = norm(r[3]) + '|' + norm(r[4]);
    const obj   = { code, grade, room, firstName: String(r[3]).trim(), lastName: String(r[4]).trim() };
    if (key !== '|') stuByName[key] = obj;
    if (code)        stuByCode[code] = obj;
  });

  let updated = 0, notFound = 0, skipped = 0;
  const updates = []; // {row, code, grade, room, firstName, lastName, position}

  for (let i = 1; i < ipadData.length; i++) {
    const row       = ipadData[i];
    const pCode     = String(row[6]).trim();
    const firstName = String(row[4]).trim();
    const lastName  = String(row[5]).trim();
    const nameKey   = norm(firstName) + '|' + norm(lastName);

    // Try match: personCode first, then name
    const stu = stuByCode[pCode] || stuByName[nameKey];

    if (!stu) {
      // If row has no name and no code → truly blank row, skip silently
      if (!firstName && !lastName && !pCode) continue;
      notFound++;
      continue;
    }

    const currentCode  = pCode;
    const currentGrade = String(row[8]).trim();
    const currentRoom  = String(row[9]).trim();
    const currentPos   = String(row[7]).trim();

    const needUpdate = currentCode !== stu.code
      || currentGrade !== stu.grade
      || currentRoom  !== stu.room
      || !currentPos.includes('นักเรียน');

    if (!needUpdate) { skipped++; continue; }

    updates.push({ row: i + 1, code: stu.code, grade: stu.grade, room: stu.room,
                   firstName: stu.firstName, lastName: stu.lastName });
    updated++;
  }

  // Apply all updates in batch (one setValues per row to keep it simple)
  updates.forEach(upd => {
    ipadSheet.getRange(upd.row, 7).setValue(upd.code);        // personCode
    ipadSheet.getRange(upd.row, 8).setValue('นักเรียน');      // position
    ipadSheet.getRange(upd.row, 9).setValue(upd.grade);       // grade
    ipadSheet.getRange(upd.row, 10).setValue(upd.room);       // room
  });

  if (updated > 0) invalidateCache_();
  log_(u.email, 'Sync iPad↔Students', `อัพเดท ${updated} | ข้ามแล้ว ${skipped} | ไม่พบ ${notFound}`);
  return ok(`Sync: อัพเดท ${updated} เครื่อง | ข้ามแล้ว ${skipped} | จับคู่ไม่ได้ ${notFound}`, { updated, skipped, notFound });
}

// ── Grade Promotion & Bulk Delete ─────────────────────────────────────────────

function getGradeStats() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return ok('', { grades: {} });
  const data = sheet.getDataRange().getValues().slice(1);
  const grades = {};
  data.forEach(r => {
    const g = String(r[5]).trim();
    if (!g) return;
    grades[g] = (grades[g] || 0) + 1;
  });
  return ok('', { grades });
}

function promoteGrades(opts) {
  // opts: { deleteGraduates: bool, unassignIpads: bool }
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');

  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  const data = sheet.getDataRange().getValues();

  function gradeNum(g) { return parseInt(String(g).replace(/^ม\.\s*/, ''), 10); }
  function nextGrade(g) {
    const n = gradeNum(g);
    if (isNaN(n)) return null;
    const hasPrefix = String(g).includes('ม.');
    return hasPrefix ? 'ม.' + (n + 1) : String(n + 1);
  }

  const promoted = {};
  const grad6Rows = []; // rows with ม.6 students (1-based, collected before editing)

  // Collect ม.6 rows first (process bottom-up later)
  for (let i = data.length - 1; i >= 1; i--) {
    const g = String(data[i][5]).trim();
    if (gradeNum(g) === 6) grad6Rows.push(i + 1); // 1-based
  }

  // Promote ม.5→ม.6, ม.4→ม.5, etc. (process bottom-up to keep row indices stable)
  for (let i = data.length - 1; i >= 1; i--) {
    const g = String(data[i][5]).trim();
    const n = gradeNum(g);
    if (isNaN(n) || n === 6) continue; // skip ม.6 — handle separately
    const ng = nextGrade(g);
    sheet.getRange(i + 1, 6).setValue(ng);
    promoted[g] = (promoted[g] || 0) + 1;
  }

  let deletedCount = 0;
  let unassignedCount = 0;

  // Optionally unassign iPads for ม.6 students (now promoted to ม.7 — but if deleteGraduates, they'll be gone)
  // Find by name since personCode may be missing
  if (opts.unassignIpads && grad6Rows.length) {
    const ipadSheet = sh(CONFIG.SHEETS.IPAD);
    if (ipadSheet) {
      const ipadData = ipadSheet.getDataRange().getValues();
      // collect names of ม.6 students
      const grad6Names = new Set();
      grad6Rows.forEach(r => {
        const row = data[r - 1];
        const key = String(row[3]).trim() + '|' + String(row[4]).trim();
        const code = String(row[1]).trim();
        if (key !== '|') grad6Names.add(key);
        if (code) grad6Names.add('code:' + code);
      });
      for (let j = ipadData.length - 1; j >= 1; j--) {
        const firstName = String(ipadData[j][4]).trim();
        const lastName  = String(ipadData[j][5]).trim();
        const pCode     = String(ipadData[j][6]).trim();
        const nameKey   = firstName + '|' + lastName;
        if (!grad6Names.has(nameKey) && !grad6Names.has('code:' + pCode)) continue;
        // Clear assignment fields, reset status
        ipadSheet.getRange(j + 1, 4, 1, 8).setValues([[
          '', '', '', '', '', '', '', CONFIG.STATUS.AVAILABLE
        ]]);
        ipadSheet.getRange(j + 1, 12, 1, 2).setValues([['', '']]);
        unassignedCount++;
      }
    }
  }

  // Delete ม.6 (bottom-up, use original row numbers)
  if (opts.deleteGraduates) {
    grad6Rows.forEach(r => sheet.deleteRow(r));
    deletedCount = grad6Rows.length;
  }

  renumberStudents_();
  invalidateCache_();
  log_(u.email, 'เลื่อนชั้น', `เลื่อน: ${JSON.stringify(promoted)} | ลบจบ: ${deletedCount} | คืนไอแพด: ${unassignedCount}`);
  return ok('เลื่อนชั้นสำเร็จ', { promoted, deletedCount, unassignedCount });
}

function deleteStudentsByGrade(grade, unassignIpads) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');

  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  const data = sheet.getDataRange().getValues();

  const targetRows = [];
  const targetNames = new Set();
  const targetCodes = new Set();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][5]).trim() !== String(grade).trim()) continue;
    targetRows.push(i + 1);
    const k = String(data[i][3]).trim() + '|' + String(data[i][4]).trim();
    if (k !== '|') targetNames.add(k);
    const c = String(data[i][1]).trim();
    if (c) targetCodes.add(c);
  }

  if (targetRows.length === 0) return err(`ไม่พบนักเรียนชั้น ${grade}`);

  let unassignedCount = 0;
  if (unassignIpads) {
    const ipadSheet = sh(CONFIG.SHEETS.IPAD);
    if (ipadSheet) {
      const ipadData = ipadSheet.getDataRange().getValues();
      for (let j = ipadData.length - 1; j >= 1; j--) {
        const nameKey = String(ipadData[j][4]).trim() + '|' + String(ipadData[j][5]).trim();
        const pCode   = String(ipadData[j][6]).trim();
        if (!targetNames.has(nameKey) && !targetCodes.has(pCode)) continue;
        ipadSheet.getRange(j + 1, 4, 1, 8).setValues([[
          '', '', '', '', '', '', '', CONFIG.STATUS.AVAILABLE
        ]]);
        ipadSheet.getRange(j + 1, 12, 1, 2).setValues([['', '']]);
        unassignedCount++;
      }
    }
  }

  targetRows.forEach(r => sheet.deleteRow(r));
  renumberStudents_();
  invalidateCache_();
  log_(u.email, 'ลบทั้งสายชั้น', `${grade} จำนวน ${targetRows.length} คน | คืนไอแพด: ${unassignedCount}`);
  return ok(`ลบนักเรียนชั้น ${grade} สำเร็จ ${targetRows.length} คน`, { deleted: targetRows.length, unassigned: unassignedCount });
}

function renumberStudents_() {
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  const last  = sheet.getLastRow();
  if (last < 2) return;
  const nums = Array.from({length: last - 1}, (_, i) => [i + 1]);
  sheet.getRange(2, 1, last - 1, 1).setValues(nums);
}

// ── Teachers ──────────────────────────────────────────────────────────────────

function getTeachers(filters) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[3] && !r[1]) continue;
    rows.push({ row: i + 1, seq: r[0], code: String(r[1]), prefix: r[2], firstName: r[3], lastName: r[4], subject: r[5] || '', room: r[6] || '', photoUrl: String(r[7] || '') });
  }
  if (filters && filters.query) {
    const q = filters.query.toLowerCase();
    rows = rows.filter(r => r.code.includes(q) || r.firstName.toLowerCase().includes(q) || r.lastName.toLowerCase().includes(q) || (r.subject || '').toLowerCase().includes(q));
  }
  // เรียงฝ่ายบริหารขึ้นก่อน (รองรับทั้ง subject ใหม่และ "ฝ่ายบริหาร" เดิม)
  const mgmtPriority = (r) => {
    if (r.subject === 'ผู้อำนวยการโรงเรียน' || r.firstName === 'ศราวุธ') return 0;
    if (r.subject === 'รองผู้อำนวยการโรงเรียน' || r.subject === 'ฝ่ายบริหาร') return 1;
    return 2;
  };
  rows.sort((a, b) => mgmtPriority(a) - mgmtPriority(b));
  return rows;
}

function updateManagementTitles() {
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return { ok: false, error: 'ไม่พบ Sheet Teachers' };
  const data = sheet.getDataRange().getValues();
  const mgmt = {
    'ศราวุธ': 'ผู้อำนวยการโรงเรียน',
    'บุญไพร': 'รองผู้อำนวยการโรงเรียน',
    'กรวัลล์': 'รองผู้อำนวยการโรงเรียน',
    'พิมพ์ประภา': 'รองผู้อำนวยการโรงเรียน',
    'พัชรวลี': 'รองผู้อำนวยการโรงเรียน',
  };
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const firstName = String(data[i][3] || '').trim();
    if (mgmt[firstName]) {
      sheet.getRange(i + 1, 6).setValue(mgmt[firstName]);
      updated++;
    }
  }
  return { ok: true, updated };
}

function addTeacher(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบ Sheet Teachers');
  sheet.appendRow(['', d.code || '', d.prefix || 'นาย', d.firstName, d.lastName, d.subject || '', d.room || '', '']);
  renumberTeachers_();
  log_(u.email, 'เพิ่มครู', `${d.prefix}${d.firstName} ${d.lastName}`);
  return ok('เพิ่มครูสำเร็จ');
}

function updateTeacher(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  sh(CONFIG.SHEETS.TEACHERS).getRange(row, 2, 1, 6).setValues([[d.code || '', d.prefix || 'นาย', d.firstName, d.lastName, d.subject || '', d.room || '']]);
  log_(u.email, 'แก้ไขครู', `${d.firstName} ${d.lastName}`);
  return ok('แก้ไขสำเร็จ');
}

function autoMatchTeacherPhotos() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');

  // ข้อมูลจาก jv.ac.th: "ชื่อ|นามสกุล" → ชื่อไฟล์รูปบนเว็บ rr
  const WEB_BASE = 'https://jv.ac.th/assets/images/personal/';
  const webMap = {
    // ผู้บริหาร
    'ศราวุธ|ศรีหาบุญทัน':      '17361547195085pic.jpg',
    'บุญไพร|หมู่ทองหลาง':      '17069350223551pic.jpg',
    'กรวัลล์|สุวรรณทา':        '17124810061905pic.jpg',
    'พิมพ์ประภา|วรรณา':        '17420287240679pic.jpg',
    'พัชรวลี|ทวรรณกุล':        '17371870944353pic.jpg',
    // ภาษาไทย
    'น้องนุช|บุญจันทร์':       '16719571949456pic.jpg',
    'ปราณี|พิศุทธิสุวรรณ':     '16714307952138pic.jpg',
    'นิภาวรรณ|ภูมิทน':         '16719571840887pic.jpg',
    'ทัศนี|ฆารเจริญ':           '16719656375994pic.jpg',
    'ดวงพร|วรสีหะ':             '16719640098671pic.jpg',
    'วราภรณ์|ศรีสันต์':         '16718618830687pic.jpg',
    'พิจิตรา|หัตถกิจ':          '16719572056732pic.jpg',
    'ทิพย์สุคนธ์|ไตรพรม':      '16725487953140pic.jpg',
    'บุณยวีร์|การชงัด':         '16719656488422pic.jpg',
    'จุฑารัตน์|ศิริวัฒน์ธนรักษ์': '17616264955252pic.jpg',
    // คณิตศาสตร์
    'วาสนา|กุนอก':              '16718738883961pic.jpg',
    'ธณัฐตา|ราษฎร์เจริญ':       '16719571279807pic.jpg',
    'กุสุมา|กุนหนองแดง':        '16719473053563pic.jpg',
    'อรัญญา|อะทอยรัมย์':        '16719571387177pic.jpg',
    'อัญญาลักษณ์|ทินกระโทก':   '16719655789708pic.jpg',
    'เชษฐ์|รักกลาง':            '16731579013448pic.jpg',
    'วาสนา|พิระชัย':            '17059793387861pic.jpeg',
    'สถิต|ถูระพี':              '16719642890413pic.jpg',
    'สรรญ์|มากทรัพย์':          '16725495613521pic.jpg',
    'ประภัสสร|เพชรสุ่ม':        '16719473205691pic.jpg',
    'ยุวรินทร์|สามารถ':         '16719473340568pic.jpg',
    'กรกนก|ร้อยอำแพง':          '16719571666659pic.jpg',
    'ภัทรชริญา|วรทองหลาง':      '17318984966636pic.jpg',
    'ศิริพักษ์|พานิช':           '17318985533767pic.jpg',
    'ฉัตรสุดา|ยายพิมพ์':        '16731579166776pic.jpg',
    'กมลรัตน์|สิงห์กุล':        '16718619873105pic.jpg',
    'พาขวัญ|สุมหิรัญ':          '16719655977923pic.jpg',
    'เอกลักษณ์|วรรธนาจิรานนท์': '16718619743553pic.jpg',
    // วิทยาศาสตร์
    'ปาริชาติ|เนตรทองหลาง':     '16719656719043pic.jpg',
    'สมศรี|ตวยกระโทก':          '16719604011647pic.jpg',
    'ศรีนวล|เช่นพิมาย':         '16719604143461pic.jpg',
    'ชลิดา|เขียวปาน':           '16719656801037pic.jpg',
    'จิรนันท์|พรหมลิ':          '16719472741528pic.jpg',
    'นุชจรี|ชวนขุนทด':          '16719604231348pic.jpg',
    'สุกัญญา|สุนทร':            '16714412363835pic.jpg',
    'วินัย|หนุนกระโทก':         '16719468752153pic.jpg',
    'สุดากาญจน์|รัตนสุข':       '17167091047720pic.jpg',
    'เทอดศักดิ์|โพธิ์ขาว':      '17721644026208pic.jpg',
    'อัจฉรา|เผ่าจินดา':         '16714293250361pic.jpg',
    'ธนาธิป|พันธุ์โหมด':        '16719604588851pic.jpg',
    'นพรุจ|แก่นกระโทก':         '16719604724295pic.jpg',
    'นัทธมน|เฝ้ากระโทก':        '16719605109071pic.jpg',
    'ณัฐนนท์|สายพิมพ์พงษ์':     '17318332295648pic.jpg',
    'ณัฐธิดา|จินากูล':           '17228278381963pic.jpg',
    'ไพศาล|พู่เจริญ':            '17318334124021pic.jpg',
    'รัตนาภรณ์|สุทธิสิน':       '17392604712458pic.jpg',
    // ภาษาต่างประเทศ
    'ชณัญกาญจน์|สนธิพันธ์':     '16714509674324pic.jpg',
    'ณัฐเสฐ|โคตรบรรเทา':        '16716042302839pic.jpg',
    'นิภาพร|คุ้มกลาง':          '16719572447961pic.jpg',
    'จันนิกา|ถองกระโทก':        '16719658466530pic.jpg',
    'ฑิพาพรรณ|นุชมี':           '16714489014225pic.jpg',
    'นัชนันท์|ปลื้มญาติ':        '16714506543893pic.jpg',
    'วิกานดา|คงคางาม':           '16718622359829pic.jpg',
    'ศิริลักษณ์|เขียวสาคู':      '16714413088147pic.jpg',
    'อัศวิณีย์|ศาลารักษ์':       '16714495464956pic.jpg',
    'อัคเรศ|ศรัณย์ธรรมกุล':     '16719660790448pic.jpg',
    'นุชสรา|บ่าพิมาย':           '16719659070319pic.jpg',
    'ชนวีร์|ขันติวงษ์':          '16719658677025pic.jpg',
    'กิรดา|คำตา':               '16731580056874pic.jpg',
    'ธัญญารัตน์|ชมไชยรัตน์':    '16931107026540pic.jpg',
    'ญาสุมินทร์|บ่อพิมาย':       '17318337300630pic.jpg',
    'ธัญชนก|จันทะดวง':          '17318336422517pic.jpg',
    // ศิลปะ
    'ประยุทธ|ช่างเกวียน':        '16718621202309pic.jpg',
    'ณัฐดนัย|เงาเกาะ':           '16719607296332pic.jpg',
    'ปางคณา|ชำนิประโคน':        '16719607088832pic.jpg',
    'ชโยดม|ประภาสโนบล':         '16719661995674pic.jpg',
    'อำนาจ|จันที':               '16719662068867pic.jpg',
    // การงานอาชีพ
    'อุทัย|วงณรา':               '16719573070655pic.jpg',
    'ประสิทธิ์|ภูมิทน':          '16719573153362pic.jpg',
    'มนัส|วรสีหะ':               '16719573330710pic.jpg',
    'สุคนธ์ทิพย์|เล็งกลาง':     '16731581314344pic.jpg',
    'รุจิรา|เทพสาร':             '16719661840724pic.jpg',
    'ปุญญิสา|ปุณณรัตนกุล':      '16731581825679pic.jpg',
    // สุขศึกษา
    'จิตรกร|เด่นกลาง':           '16718752464715pic.jpg',
    'ประภา|บุญนิธิ':             '16719603211174pic.jpg',
    'กริษณุ|ลิ้มศิริอังกูร':     '16719662201427pic.jpg',
    'ณัฐนันท์|งามลาภ':           '16719662478540pic.jpg',
    'ผกามาศ|ลิ้มศิริอังกูร':     '16719662375533pic.jpg',
    // แนะแนว
    'สรัญญา|ยางนอก':             '16716154663815pic.jpg',
    'ปิยะนันท์|เหมือนเหลา':      '16719606692728pic.jpg',
    'กัญญาวีร์|ประสพผล':         '16719662609094pic.jpg',
    'พรสุดา|ห่วงจริง':           '16719606775075pic.jpg',
  };

  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบ Sheet Teachers');
  const data = sheet.getDataRange().getValues();
  const folders = DriveApp.getFoldersByName('รูปครู');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('รูปครู');

  // รวบรวมรายการที่ต้องดาวน์โหลด
  const toFetch = []; // { sheetRow, firstName, lastName, photoFile, url }
  let skipped = 0, noMatch = 0;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const firstName = String(r[3] || '').trim();
    const lastName  = String(r[4] || '').trim();
    const existing  = String(r[7] || '').trim();
    if (!firstName) continue;
    if (existing)   { skipped++; continue; }
    const key = firstName + '|' + lastName;
    const photoFile = webMap[key];
    if (!photoFile) { noMatch++; continue; }
    toFetch.push({ sheetRow: i + 1, firstName, lastName, photoFile });
  }

  if (toFetch.length === 0) {
    return ok(`ไม่มีครูที่ต้องจับคู่ · ข้าม ${skipped} คน (มีรูปแล้ว) · ไม่พบข้อมูล ${noMatch} คน`, { matched: 0, skipped, noMatch, errors: 0, results: [] });
  }

  // ดาวน์โหลดพร้อมกันทั้งหมด
  const requests = toFetch.map(function(t) {
    return { url: WEB_BASE + t.photoFile, muteHttpExceptions: true };
  });
  const responses = UrlFetchApp.fetchAll(requests);

  let matched = 0, errors = 0;
  const results = [];
  const sheetUpdates = []; // { row, url }

  for (let j = 0; j < toFetch.length; j++) {
    const t = toFetch[j];
    const resp = responses[j];
    try {
      if (resp.getResponseCode() !== 200) {
        errors++;
        results.push({ name: t.firstName, err: 'HTTP ' + resp.getResponseCode() });
        continue;
      }
      const ext = t.photoFile.match(/\.[^.]+$/)[0];
      const fname = t.firstName + '_' + t.lastName + ext;
      const blob = resp.getBlob().setName(fname);
      const existing_ = folder.getFilesByName(fname);
      while (existing_.hasNext()) existing_.next().setTrashed(true);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w200';
      sheetUpdates.push({ row: t.sheetRow, url });
      matched++;
      results.push({ name: t.firstName + ' ' + t.lastName, url });
    } catch(e) {
      errors++;
      results.push({ name: t.firstName, err: e.message });
    }
  }

  // เขียนชีตครั้งเดียว
  for (let k = 0; k < sheetUpdates.length; k++) {
    sheet.getRange(sheetUpdates[k].row, 8).setValue(sheetUpdates[k].url);
  }

  log_(u.email, 'จับคู่รูปครูอัตโนมัติ', `matched:${matched} skipped:${skipped} noMatch:${noMatch} errors:${errors}`);
  return ok(`จับคู่สำเร็จ ${matched} คน · ข้าม ${skipped} คน (มีรูปแล้ว) · ไม่พบข้อมูล ${noMatch} คน · ผิดพลาด ${errors} คน`, { matched, skipped, noMatch, errors, results });
}

function uploadTeacherPhoto(row, base64Data, filename, mimeType) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  if (!row || !base64Data) return err('ข้อมูลไม่ครบ');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบ Sheet Teachers');
  try {
    const folders = DriveApp.getFoldersByName('รูปครู');
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('รูปครู');
    const existing = folder.getFilesByName(filename);
    while (existing.hasNext()) existing.next().setTrashed(true);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg', filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w200';
    sheet.getRange(row, 8).setValue(url);
    log_(u.email, 'อัพโหลดรูปครู', `แถว: ${row} | ${file.getName()}`);
    return ok('สำเร็จ', { url });
  } catch(e) { return err('อัพโหลดไม่สำเร็จ: ' + e.message); }
}

function deleteTeacherPhoto(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบ Sheet Teachers');
  sheet.getRange(row, 8).setValue('');
  log_(u.email, 'ลบรูปครู', `แถว: ${row}`);
  return ok('ลบรูปสำเร็จ');
}

function deleteTeacher(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const name = sh(CONFIG.SHEETS.TEACHERS).getRange(row, 4).getValue();
  sh(CONFIG.SHEETS.TEACHERS).deleteRow(row);
  renumberTeachers_();
  log_(u.email, 'ลบครู', name);
  return ok('ลบสำเร็จ');
}

function batchAddTeachers(teachers) {
  try {
    const u = getCurrentUser(arguments[arguments.length - 1]);
    if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
    if (!Array.isArray(teachers) || !teachers.length) return err('ไม่มีข้อมูลครู');
    let sheet = sh(CONFIG.SHEETS.TEACHERS);
    if (!sheet) {
      // sheet ไม่มี — สร้างใหม่พร้อม header
      sheet = ss().insertSheet(CONFIG.SHEETS.TEACHERS);
      sheet.appendRow(['ลำดับ','รหัส','คำนำหน้า','ชื่อ','นามสกุล','วิชา','ห้องประจำ','รูปภาพ']);
    }
    let added = 0;
    teachers.forEach(d => {
      if (!d.firstName) return;
      sheet.appendRow(['', d.code || '', d.prefix || '', d.firstName, d.lastName || '', d.subject || '', d.room || '', '']);
      added++;
    });
    renumberTeachers_();
    log_(u.email, 'นำเข้าครู CSV', `${added} คน`);
    return ok(`นำเข้าครูสำเร็จ ${added} คน`);
  } catch(e) {
    return err('เกิดข้อผิดพลาด: ' + e.message);
  }
}

function renumberTeachers_() {
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;
  const nums = Array.from({length: last - 1}, (_, i) => [i + 1]);
  sheet.getRange(2, 1, last - 1, 1).setValues(nums);
}

// ── Issues (Problem Reports) ──────────────────────────────────────────────────

function lookupStudentByCode(code) {
  if (!code) return { found: false };
  const students = sh(CONFIG.SHEETS.STUDENTS);
  if (!students) return { found: false };
  const sData = students.getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    if (String(sData[i][1]) === String(code).trim()) {
      const s = sData[i];
      const ipads = getAllIPads();
      const ipad = ipads.find(r => String(r.personCode) === String(code).trim());
      return {
        found: true,
        student: { code: String(s[1]), prefix: s[2], firstName: String(s[3]), lastName: String(s[4]), grade: String(s[5]), room: String(s[6]) },
        ipad: ipad ? { serial: ipad.serial, assetCode: ipad.assetCode, status: ipad.status, borrowDate: ipad.borrowDate, row: ipad.row } : null
      };
    }
  }
  return { found: false };
}

function searchStudentsForIssue(query) {
  query = String(query || '').trim();
  if (!query || query.length < 2) return [];
  const q = query.toLowerCase();
  const students = sh(CONFIG.SHEETS.STUDENTS);
  if (!students) return [];
  const sData = students.getDataRange().getValues();
  const ipads = getAllIPads();
  const results = [];
  for (let i = 1; i < sData.length && results.length < 10; i++) {
    const code = String(sData[i][1] || '');
    const firstName = String(sData[i][3] || '');
    const lastName = String(sData[i][4] || '');
    const grade = String(sData[i][5] || '');
    const room = String(sData[i][6] || '');
    const ipad = ipads.find(r => String(r.personCode) === code);
    const serial = ipad ? ipad.serial : '';
    if (firstName.toLowerCase().includes(q) || lastName.toLowerCase().includes(q) ||
        serial.toLowerCase().includes(q) || code.includes(q)) {
      results.push({ code, prefix: String(sData[i][2] || ''), firstName, lastName, grade, room, serial });
    }
  }
  return results;
}

function uploadIssueFile(base64Data, filename, mimeType) {
  try {
    const folderName = 'แจ้งปัญหาไอแพด';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/octet-stream', filename);
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return ok('อัพโหลดสำเร็จ', { url: file.getUrl(), name: filename });
  } catch(e) { return err('อัพโหลดไม่สำเร็จ: ' + e.message); }
}

function issueDeviceFromType_(issueType) {
  const text = String(issueType || '').trim();
  const devices = ['ไอแพด', 'ปากกา', 'เคส', 'สายชาร์จ'];
  for (let i = 0; i < devices.length; i++) {
    if (text === devices[i] || text.indexOf(devices[i] + ' - ') === 0) return devices[i];
  }
  return '';
}

function submitIssue(data) {
  // data: { studentCode, issueType, description, fileUrls: ['url1',...] }
  if (!data || !data.studentCode) return err('กรุณาระบุรหัสนักเรียน');
  const lookup = lookupStudentByCode(data.studentCode);
  if (!lookup.found) return err('ไม่พบข้อมูลนักเรียนรหัส ' + data.studentCode);
  if (!lookup.ipad) return err('ไม่พบข้อมูลไอแพดของนักเรียน กรุณาติดต่อผู้ดูแลระบบ');

  let sheet = sh(CONFIG.SHEETS.ISSUES);
  if (!sheet) {
    // auto-create Issues sheet
    sheet = ss().insertSheet(CONFIG.SHEETS.ISSUES);
    sheet.appendRow(['ID','รหัสนักเรียน','ชื่อ','Serial','ประเภทปัญหา','รายละเอียด','ไฟล์แนบ','สถานะเดิม','สถานะปัญหา','วันที่แจ้ง','หมายเหตุAdmin','อัพเดทล่าสุด','กำหนดส่ง']);
  }

  const fileUrls = (data.fileUrls || []).slice(0, 10);
  const incomingDevice = issueDeviceFromType_(data.issueType);

  // Prevent duplicate open reports for the same serial and same device/accessory.
  let existingRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 4, sheet.getLastRow() - 1, 6).getValues()
    : [];
  existingRows = existingRows.filter(row => {
    const existingDevice = issueDeviceFromType_(row[1]);
    return !incomingDevice || !existingDevice || incomingDevice === existingDevice;
  });
  const alreadyOpen = existingRows.some(row =>
    String(row[0]).trim() === lookup.ipad.serial &&
    [CONFIG.STATUS.PENDING_ISSUE, 'รับเรื่องแล้ว', 'อยู่ระหว่างการตรวจสอบ', 'อยู่ระหว่างการเคลม'].includes(String(row[5]).trim())
  );
  if (alreadyOpen) return err('มีการแจ้งปัญหาสำหรับไอแพดเครื่องนี้อยู่แล้ว กรุณารอให้เจ้าหน้าที่ดำเนินการก่อน');

  const lastRow = sheet.getLastRow();
  const newId   = lastRow > 1 ? (parseInt(sheet.getRange(lastRow, 1).getValue(), 10) || 0) + 1 : 1;
  const studentName = `${lookup.student.prefix}${lookup.student.firstName} ${lookup.student.lastName}`;

  sheet.appendRow([newId, data.studentCode, studentName, lookup.ipad.serial,
    data.issueType || 'ไม่ระบุประเภท', data.description || '',
    fileUrls.join('\n'), lookup.ipad.status, CONFIG.STATUS.PENDING_ISSUE,
    new Date(), '', new Date(), '']);

  sh(CONFIG.SHEETS.IPAD).getRange(lookup.ipad.row, 11).setValue(CONFIG.STATUS.PENDING_ISSUE);
  invalidateCache_();

  const ts = fmtTs(new Date());
  const fileInfo = fileUrls.length ? `\n\nไฟล์แนบ (${fileUrls.length} ไฟล์):\n${fileUrls.join('\n')}` : '';
  notifyAdmins_(`[แจ้งปัญหาไอแพด] ${studentName} — ${data.issueType || 'ปัญหาไอแพด'}`,
    `มีการแจ้งปัญหาไอแพดใหม่\n\nนักเรียน: ${studentName} (${data.studentCode})\n` +
    `ชั้น/ห้อง: ${lookup.student.grade}/${lookup.student.room}\nSerial: ${lookup.ipad.serial}\n` +
    `ประเภทปัญหา: ${data.issueType || 'ไม่ระบุ'}\nรายละเอียด: ${data.description || '-'}\nเวลา: ${ts}` +
    fileInfo + `\n\nกรุณาเข้าระบบเพื่อดำเนินการ`);
  notifyAdminPush_(
    'แจ้งปัญหา iPad ใหม่',
    `${studentName} ${lookup.student.grade}/${lookup.student.room} | Serial ${lookup.ipad.serial} | ${data.issueType || 'ไม่ระบุประเภท'}`,
    'issues',
    { issueId: newId, serial: lookup.ipad.serial, studentCode: data.studentCode }
  );
  log_(data.studentCode, 'แจ้งปัญหาไอแพด', `Serial: ${lookup.ipad.serial} | ${data.issueType}`);
  return ok('แจ้งปัญหาสำเร็จ เจ้าหน้าที่จะดำเนินการโดยเร็ว');
}

function analyzeIssue(description) {
  if (!description || !description.trim()) return keywordClassify_('');
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')
                || PropertiesService.getUserProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return keywordClassify_(description);
    const prompt =
      'คุณเป็นผู้ช่วยวิเคราะห์ปัญหาไอแพดในโรงเรียน ' +
      'วิเคราะห์คำอธิบายปัญหาต่อไปนี้แล้วตอบกลับเป็น JSON เท่านั้น (ไม่มีข้อความอื่น)\n\n' +
      'รูปแบบ JSON ที่ต้องการ:\n' +
      '{"issueType":"<ประเภท>","suggestions":["<คำแนะนำ1>","<คำแนะนำ2>","<คำแนะนำ3>"]}\n\n' +
      'ประเภทที่เลือกได้เท่านั้น:\n' +
      'หน้าจอแตก/แตกร้าว, แบตเตอรี่เสีย, ปุ่มเสีย, ลำโพงเสีย, กล้องเสีย, ชาร์จไม่ได้, ซอฟต์แวร์มีปัญหา, อื่นๆ\n\n' +
      'คำแนะนำต้องเป็นภาษาไทย สั้นกระชับ และทำได้จริงด้วยตนเอง\n\n' +
      'คำอธิบายปัญหา: ' + description;
    const resp = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 300 }
        })
      }
    );
    const json = JSON.parse(resp.getContentText());
    const text = json.candidates[0].content.parts[0].text;
    const parsed = JSON.parse(text);
    const validTypes = ['หน้าจอแตก/แตกร้าว','แบตเตอรี่เสีย','ปุ่มเสีย','ลำโพงเสีย','กล้องเสีย','ชาร์จไม่ได้','ซอฟต์แวร์มีปัญหา','อื่นๆ'];
    const issueType = validTypes.includes(parsed.issueType) ? parsed.issueType : 'อื่นๆ';
    const suggestions = Array.isArray(parsed.suggestions) ? parsed.suggestions.slice(0, 4) : getSuggestionsForType_(issueType);
    return { success: true, issueType: issueType, suggestions: suggestions, source: 'ai' };
  } catch(e) {
    return keywordClassify_(description);
  }
}

function keywordClassify_(description) {
  const d = (description || '').toLowerCase();
  let type;
  if      (d.match(/จอ|แตก|ร้าว|กระจก|screen/))                      type = 'หน้าจอแตก/แตกร้าว';
  else if (d.match(/แบต|battery|หมดเร็ว|ไม่ชาร์จ|ชาร์จแล้วหมดเร็ว/)) type = 'แบตเตอรี่เสีย';
  else if (d.match(/ปุ่ม|โฮม|home|power|volume|เปิดไม่ติด/))           type = 'ปุ่มเสีย';
  else if (d.match(/ลำโพง|เสียง|speaker|audio|ได้ยิน/))               type = 'ลำโพงเสีย';
  else if (d.match(/กล้อง|ถ่าย|camera/))                               type = 'กล้องเสีย';
  else if (d.match(/ชาร์จ|สายชาร์จ|หัวชาร์จ|charge/))                 type = 'ชาร์จไม่ได้';
  else if (d.match(/แอพ|แอป|app|ค้าง|crash|อัพเดท|update|ระบบ|ios/))  type = 'ซอฟต์แวร์มีปัญหา';
  else                                                                    type = 'อื่นๆ';
  return { success: true, issueType: type, suggestions: getSuggestionsForType_(type), source: 'keyword' };
}

function getSuggestionsForType_(type) {
  const map = {
    'หน้าจอแตก/แตกร้าว': ['ระวังอย่าให้กระจกสัมผัสมือโดยตรง','หลีกเลี่ยงการกดบนหน้าจอ','งดใช้ไอแพดจนกว่าแอดมินจะตรวจสอบ'],
    'แบตเตอรี่เสีย': ['ลองชาร์จด้วยสายและหัวชาร์จอื่น','ตรวจสอบว่าพอร์ตชาร์จสะอาด','หากแบตหมดเร็วผิดปกติให้แจ้งแอดมิน'],
    'ปุ่มเสีย': ['ลองรีสตาร์ทไอแพดโดยกดปุ่มข้างเครื่องค้างไว้','ตรวจดูว่าปุ่มไม่ติดขัดหรือมีสิ่งสกปรก','หากไม่ดีขึ้นให้แจ้งแอดมิน'],
    'ลำโพงเสีย': ['ตรวจสอบว่าไม่ได้เปิด Silent Mode','ลองปรับระดับเสียงขึ้น','ลองรีสตาร์ทเครื่อง','หากยังไม่มีเสียงให้แจ้งแอดมิน'],
    'กล้องเสีย': ['ปิดแอพกล้องแล้วเปิดใหม่','รีสตาร์ทเครื่อง','ตรวจสอบว่าเลนส์กล้องสะอาด','หากยังมีปัญหาแจ้งแอดมิน'],
    'ชาร์จไม่ได้': ['ลองเปลี่ยนสายชาร์จ','ทำความสะอาดพอร์ตชาร์จด้วยลมเป่า','ตรวจสอบว่าหัวชาร์จทำงานปกติ','หากยังไม่ชาร์จแจ้งแอดมิน'],
    'ซอฟต์แวร์มีปัญหา': ['ปิดแอพที่มีปัญหาแล้วเปิดใหม่','รีสตาร์ทเครื่อง','ตรวจสอบว่า iOS อัพเดทล่าสุด','หากยังมีปัญหาแจ้งแอดมิน'],
    'อื่นๆ': ['ลองรีสตาร์ทเครื่องก่อน','บันทึกอาการปัญหาให้ละเอียด','แจ้งแอดมินเพื่อตรวจสอบเพิ่มเติม']
  };
  return map[type] || map['อื่นๆ'];
}

function getIssues(filters) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.ISSUES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    rows.push({
      row: i + 1, id: r[0], studentCode: String(r[1]), studentName: r[2],
      serial: String(r[3]), issueType: r[4], description: r[5],
      fileUrls: String(r[6] || '').split('\n').filter(Boolean),
      prevStatus: r[7], status: r[8],
      reportedAt: r[9] ? fmtTs(r[9]) : '',
      adminNotes: r[10] || '',
      updatedAt: r[11] ? fmtTs(r[11]) : '',
      roomDeadline: r[12] || ''
    });
  }
  if (filters && filters.status && filters.status !== 'all') {
    rows = rows.filter(r => r.status === filters.status);
  }
  return rows.reverse();
}

function updateIssueStatus(issueRow, newIssueStatus, adminNotes, roomDeadline) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.ISSUES);
  const S = CONFIG.STATUS;

  const prevStatus = sheet.getRange(issueRow, 8).getValue();
  const serial     = sheet.getRange(issueRow, 4).getValue();
  const studentCode = sheet.getRange(issueRow, 2).getValue();
  const studentName = sheet.getRange(issueRow, 3).getValue();
  const issueId    = sheet.getRange(issueRow, 1).getValue();

  sheet.getRange(issueRow, 9).setValue(newIssueStatus);
  sheet.getRange(issueRow, 11).setValue(adminNotes || '');
  sheet.getRange(issueRow, 12).setValue(new Date());
  if (roomDeadline) sheet.getRange(issueRow, 13).setValue(roomDeadline);

  let newIPadStatus = null;
  if (newIssueStatus === 'รับเรื่องแล้ว' || newIssueStatus === 'อยู่ระหว่างการตรวจสอบ') newIPadStatus = S.UNDER_INSPECTION;
  else if (newIssueStatus === 'แก้ไขแล้ว') {
    // restore to previous "normal" status; never restore to a problem status
    const safeStatuses = [S.BORROWED, S.RETURNED, S.AVAILABLE];
    newIPadStatus = safeStatuses.includes(prevStatus) ? prevStatus : S.BORROWED;
  }
  else if (newIssueStatus === 'ใช้งานไม่ได้') newIPadStatus = S.DEFECTIVE;
  else if (newIssueStatus === 'อยู่ระหว่างการเคลม')          newIPadStatus = S.UNDER_CLAIM;

  if (newIPadStatus && serial) {
    const ipads = getAllIPads();
    const ipad  = ipads.find(r => r.serial === String(serial));
    if (ipad) { sh(CONFIG.SHEETS.IPAD).getRange(ipad.row, 11).setValue(newIPadStatus); invalidateCache_(); }
  }

  log_(u.email, `อัพเดทปัญหา #${issueId}`, `${studentName} → ${newIssueStatus}`);
  return ok('อัพเดทสำเร็จ');
}

function getMyIssues(studentCode) {
  if (!studentCode) return { success: false, issues: [] };
  const sheet = sh(CONFIG.SHEETS.ISSUES);
  if (!sheet) return { success: true, issues: [] };
  const data = sheet.getDataRange().getValues();
  const issues = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (String(r[1]).trim() !== String(studentCode).trim()) continue;
    issues.push({
      row: i + 1, id: r[0], issueType: r[4], description: r[5],
      status: r[8], reportedAt: r[9] ? fmtTs(r[9]) : '',
      adminNotes: r[10] || '', roomDeadline: r[12] || ''
    });
  }
  return { success: true, issues: issues.reverse() };
}

function cancelStudentIssue(issueId, studentCode) {
  if (!issueId || !studentCode) return err('ข้อมูลไม่ครบ');
  const sheet = sh(CONFIG.SHEETS.ISSUES);
  if (!sheet) return err('ไม่พบข้อมูล');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(issueId)) continue;
    if (String(data[i][1]).trim() !== String(studentCode).trim()) return err('ไม่มีสิทธิ์ยกเลิก');
    const status = String(data[i][8] || '');
    if (status !== 'อยู่ระหว่างการส่งเรื่อง') return err('ไม่สามารถยกเลิกได้ เนื่องจากเรื่องอยู่ระหว่างดำเนินการแล้ว');
    sheet.getRange(i + 1, 9).setValue('ยกเลิกโดยนักเรียน');
    sheet.getRange(i + 1, 12).setValue(new Date());
    return ok('ยกเลิกเรื่องแล้ว');
  }
  return err('ไม่พบรายการนี้');
}

function getIssueSummary() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return { pending: 0, inspect: 0, total: 0 };
  const sheet = sh(CONFIG.SHEETS.ISSUES);
  if (!sheet || sheet.getLastRow() < 2) return { pending: 0, inspect: 0, total: 0 };
  const data = sheet.getDataRange().getValues();
  let pending = 0, inspect = 0, total = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const status = String(data[i][8] || '');
    if (status === 'แก้ไขแล้ว' || status === 'ยกเลิกโดยนักเรียน') continue;
    total++;
    if (status === 'อยู่ระหว่างการส่งเรื่อง') pending++;
    else if (status === 'อยู่ระหว่างการตรวจสอบ') inspect++;
  }
  return { pending, inspect, total };
}

function submitAppRequest(data) {
  if (!data || !data.studentCode) return err('กรุณาระบุรหัสนักเรียน');
  if (!data.appName || !data.appName.trim()) return err('กรุณาระบุชื่อแอพ');
  if (!data.reason || !data.reason.trim()) return err('กรุณาระบุเหตุผล');
  const lookup = lookupStudentByCode(data.studentCode);
  if (!lookup.found) return err('ไม่พบข้อมูลนักเรียนรหัส ' + data.studentCode);
  let sheet = sh(CONFIG.SHEETS.APP_REQUESTS);
  if (!sheet) {
    sheet = ss().insertSheet(CONFIG.SHEETS.APP_REQUESTS);
    sheet.appendRow(['ID','รหัสนักเรียน','ชื่อ','ชั้น','ห้อง','ชื่อแอพ','เหตุผล','สถานะ','วันที่ยื่น','หมายเหตุAdmin']);
    sheet.getRange(1,1,1,10).setBackground('#3F51B5').setFontColor('#fff').setFontWeight('bold');
  }
  const newId = 'AR' + new Date().getTime();
  const studentName = lookup.student.prefix + lookup.student.firstName + ' ' + lookup.student.lastName;
  sheet.appendRow([
    newId, data.studentCode, studentName,
    lookup.student.grade || '', lookup.student.room || '',
    data.appName.trim(), data.reason.trim(),
    'รอดำเนินการ', new Date(), ''
  ]);
  notifyAdmins_(`[คำร้องขอแอพ] ${studentName} — ${data.appName}`,
    `นักเรียน: ${studentName} (${data.studentCode})\n` +
    `ชั้น/ห้อง: ${lookup.student.grade}/${lookup.student.room}\n` +
    `แอพที่ขอ: ${data.appName}\nเหตุผล: ${data.reason}\n\nกรุณาเข้าระบบเพื่อดำเนินการ`);
  log_(data.studentCode, 'ส่งคำร้องขอแอพ', `แอพ: ${data.appName}`);
  return ok('ส่งคำร้องสำเร็จ ผู้ดูแลจะตรวจสอบและแจ้งผลโดยเร็ว');
}

// ── Accounts (Student/Teacher Login) ─────────────────────────────────────────

function lookupAccount_(type, code) {
  const sheet = sh(CONFIG.SHEETS.ACCOUNTS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === type && String(data[i][1]) === String(code)) {
      return { row: i + 1, type: data[i][0], code: String(data[i][1]), username: data[i][2], password: data[i][3], email: data[i][4] };
    }
  }
  return null;
}

function loginByStudentCode(code) {
  if (!code) return err('กรุณาระบุรหัสนักเรียน');
  const lookup = lookupStudentByCode(code);
  if (!lookup.found) return err('ไม่พบรหัสนักเรียนนี้ในระบบ');
  const account = lookupAccount_('student', code);
  log_(code, 'Student Lookup', `รหัส: ${code}`);
  return ok('พบข้อมูล', { student: lookup.student, ipad: lookup.ipad, hasAccount: !!account });
}

function loginTeacherByName(firstName, lastName) {
  if (!firstName || !lastName) return err('กรุณาระบุชื่อและนามสกุล');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบข้อมูลครูในระบบ');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[3]).trim() === String(firstName).trim() && String(r[4]).trim() === String(lastName).trim()) {
      const teacherCode = String(r[1]) || String(i);
      const teacher = { code: teacherCode, prefix: r[2], firstName: String(r[3]), lastName: String(r[4]), subject: r[5] || '', room: r[6] || '' };
      const account = lookupAccount_('teacher', teacherCode);
      const ipads = getAllIPads();
      const ipad = ipads.find(p => p.position === 'ครู' && p.firstName.trim() === firstName.trim() && p.lastName.trim() === lastName.trim());
      log_(firstName + ' ' + lastName, 'Teacher Lookup', '');
      return ok('พบข้อมูล', { teacher, ipad: ipad ? { serial: ipad.serial, assetCode: ipad.assetCode, status: ipad.status, row: ipad.row } : null, hasAccount: !!account });
    }
  }
  return err('ไม่พบข้อมูลครู กรุณาตรวจสอบชื่อ-นามสกุลอีกครั้ง');
}

function loginByUsername(username, password) {
  if (!username || !password) return err('กรุณากรอก username และ password');
  const sheet = sh(CONFIG.SHEETS.ACCOUNTS);
  if (!sheet) return err('ยังไม่มีบัญชีในระบบ');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[2]).trim() === String(username).trim() && String(r[3]).trim() === String(password).trim()) {
      sheet.getRange(i + 1, 7).setValue(new Date());
      const code = String(r[1]), type = String(r[0]);
      if (type === 'student') {
        const lookup = lookupStudentByCode(code);
        if (!lookup.found) return err('บัญชีนี้ถูกลบออกจากระบบนักเรียนแล้ว กรุณาติดต่อ Admin');
        log_(username, 'Student Login', `code: ${code}`);
        return ok('เข้าสู่ระบบสำเร็จ', { type, student: lookup.student, ipad: lookup.ipad });
      } else {
        log_(username, 'Teacher Login', `code: ${code}`);
        return ok('เข้าสู่ระบบสำเร็จ', { type, code, email: r[4] });
      }
    }
  }
  return err('username หรือ password ไม่ถูกต้อง');
}

function createAccount(type, code, username, password, email) {
  if (!type || !code || !username || !password) return err('ข้อมูลไม่ครบ');
  if (username.length < 4) return err('username ต้องมีอย่างน้อย 4 ตัวอักษร');
  if (password.length < 6) return err('password ต้องมีอย่างน้อย 6 ตัวอักษร');
  const sheet = sh(CONFIG.SHEETS.ACCOUNTS);
  if (!sheet) return err('ไม่พบ Sheet Accounts');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === username.trim().toLowerCase() && String(data[i][1]) !== String(code)) {
      return err('username นี้มีผู้ใช้แล้ว กรุณาเลือก username อื่น');
    }
    if (String(data[i][0]) === type && String(data[i][1]) === String(code)) {
      sheet.getRange(i + 1, 3, 1, 3).setValues([[username, password, email || '']]);
      log_(code, 'อัพเดทบัญชี', `type: ${type}`);
      return ok('อัพเดทบัญชีสำเร็จ');
    }
  }
  sheet.appendRow([type, code, username, password, email || '', new Date(), new Date()]);
  log_(code, 'สร้างบัญชีใหม่', `type: ${type}, username: ${username}`);
  return ok('สร้างบัญชีสำเร็จ');
}

function requestPasswordReset(identifier, type, requestEmail) {
  if (!identifier) return err('กรุณาระบุรหัส/ชื่อ');
  const ts = fmtTs(new Date());
  const emailLine = requestEmail ? `\nอีเมลที่ขอรับ: ${requestEmail}` : '';
  notifyAdmins_(`[ลืมรหัสผ่าน] ${type === 'student' ? 'นักเรียน' : 'ครู'}: ${identifier}`,
    `มีการขอรีเซ็ตรหัสผ่าน\n\nประเภท: ${type === 'student' ? 'นักเรียน' : 'ครู'}\n` +
    `รหัส/ชื่อ: ${identifier}${emailLine}\nเวลา: ${ts}\n\n` +
    `กรุณา Reset password ในระบบแล้วแจ้งผู้ใช้`);
  log_(identifier, 'ขอรีเซ็ตรหัสผ่าน', `email: ${requestEmail || 'ไม่ระบุ'}`);
  const msg = requestEmail
    ? 'ส่งคำขอสำเร็จ — Admin จะส่ง password ชั่วคราวไปที่ ' + requestEmail
    : 'ส่งคำขอสำเร็จ — กรุณาติดต่อ Admin โดยตรงเพื่อขอ password ใหม่';
  return ok(msg);
}

function adminResetPassword(type, code, newPassword) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.ACCOUNTS);
  if (!sheet) return err('ไม่พบ Sheet Accounts');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === type && String(data[i][1]) === String(code)) {
      sheet.getRange(i + 1, 4).setValue(newPassword);
      log_(u.email, 'รีเซ็ตรหัสผ่าน', `${type}: ${code}`);
      return ok('รีเซ็ตรหัสผ่านสำเร็จ');
    }
  }
  return err('ไม่พบบัญชีนี้ในระบบ');
}

// ── Student Photos ────────────────────────────────────────────────────────────

function getStudentsByGradeForPhoto(grade) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[1] && !r[3]) continue;
    if (grade && String(r[5]) !== String(grade)) continue;
    result.push({ row: i + 1, code: String(r[1]), prefix: r[2], firstName: String(r[3]),
      lastName: String(r[4]), grade: String(r[5]), room: String(r[6]),
      photoUrl: String(r[7] || '') });
  }
  return ok('', { students: result });
}

function uploadStudentPhoto(code, base64Data, filename, mimeType) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  if (!code || !base64Data) return err('ข้อมูลไม่ครบ');
  const sheet = sh(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return err('ไม่พบ Sheet Students');
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  let targetPhotoUrl = '';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(code)) {
      targetRow = i + 1;
      targetPhotoUrl = String(data[i][7] || '').trim();
      break;
    }
  }
  if (targetRow < 0) return err('ไม่พบรหัสนักเรียน ' + code);
  if (targetPhotoUrl) return ok('ข้าม: นักเรียนมีรูปแล้ว', { code, url: targetPhotoUrl, skipped: true });
  try {
    const folderName = 'รูปนักเรียน';
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    const existing = folder.getFilesByName(filename);
    while (existing.hasNext()) existing.next().setTrashed(true);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg', filename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w120';
    sheet.getRange(targetRow, 8).setValue(url);
    log_(u.email, 'อัพโหลดรูปนักเรียน', `รหัส: ${code} | ${file.getName()}`);
    return ok('สำเร็จ', { url, code });
  } catch(e) { return err('อัพโหลดล้มเหลว: ' + e.message); }
}

// ── Deduplication ─────────────────────────────────────────────────────────────

function findDuplicates(sheetName, keyCol) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(sheetName);
  if (!sheet) return err('ไม่พบ Sheet: ' + sheetName);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ok('', { groups: [], total: 0 });

  if (keyCol < 1 || keyCol > data[0].length) return err('keyCol ไม่ถูกต้อง');
  const keyIdx = keyCol - 1;
  const seen = {};   // key → first rowNum (1-based)
  const groups = {}; // key → [{rowNum, rowData}]

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][keyIdx]).trim();
    if (!key || key === '' || key === 'undefined') continue;
    const rowNum = i + 1;
    if (!seen[key]) {
      seen[key] = rowNum;
      groups[key] = [{ rowNum, rowData: data[i] }];
    } else {
      groups[key].push({ rowNum, rowData: data[i] });
    }
  }

  const dupGroups = Object.entries(groups)
    .filter(([, rows]) => rows.length > 1)
    .map(([key, rows]) => ({ key, rows }));

  const total = dupGroups.reduce((s, g) => s + g.rows.length - 1, 0);
  return ok('', { groups: dupGroups, total, header: data[0] });
}

function removeDuplicates(sheetName, keyCol) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(sheetName);
  if (!sheet) return err('ไม่พบ Sheet: ' + sheetName);

  const data = sheet.getDataRange().getValues();
  if (keyCol < 1 || keyCol > data[0].length) return err('keyCol ไม่ถูกต้อง');
  const keyIdx = keyCol - 1;
  const seen = new Set();
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][keyIdx]).trim();
    if (!key || key === '' || key === 'undefined') continue;
    if (seen.has(key)) rowsToDelete.push(i + 1);
    else seen.add(key);
  }

  // Delete bottom-up to preserve row numbers
  rowsToDelete.slice().reverse().forEach(r => sheet.deleteRow(r));

  invalidateCache_();
  log_(u.email, 'ลบรายซ้ำ', `Sheet: ${sheetName} | ลบ ${rowsToDelete.length} แถว`);
  return ok(`ลบแถวซ้ำ ${rowsToDelete.length} แถวจาก ${sheetName}`, { removed: rowsToDelete.length });
}

// ── OTP Login ─────────────────────────────────────────────────────────────────

function requestOTP(emailOrName) {
  if (!emailOrName) return err('กรุณากรอกชื่อหรือ email');
  let email = emailOrName.trim();

  // If input has no @, treat as name — look up email in Users or Accounts sheets
  if (!email.includes('@')) {
    const nameLow = email.toLowerCase();
    const usersSheet_ = sh(CONFIG.SHEETS.USERS);
    if (!usersSheet_) return err('ระบบยังไม่พร้อม กรุณาติดต่อ Admin');
    const usersData = usersSheet_.getDataRange().getValues().slice(1);
    const adminRow = usersData.find(r => String(r[1]).trim().toLowerCase() === nameLow);
    if (adminRow && String(adminRow[0]).includes('@')) {
      email = String(adminRow[0]).trim();
    } else {
      // Search Students sheet by firstName+lastName → look up email in Accounts
      const stuSheet = sh(CONFIG.SHEETS.STUDENTS);
      if (stuSheet) {
        const stuData = stuSheet.getDataRange().getValues().slice(1);
        const stuRow = stuData.find(r => (String(r[3]) + ' ' + String(r[4])).trim().toLowerCase() === nameLow);
        if (stuRow) {
          const acc = lookupAccount_('student', String(stuRow[1]).trim());
          if (acc && String(acc.email).includes('@')) email = String(acc.email).trim();
        }
      }
      // Search Teachers sheet by firstName+lastName → look up email in Accounts
      if (!email.includes('@')) {
        const tSheet = sh(CONFIG.SHEETS.TEACHERS);
        if (tSheet) {
          const tData = tSheet.getDataRange().getValues().slice(1);
          const tRow = tData.find(r => (String(r[3]) + ' ' + String(r[4])).trim().toLowerCase() === nameLow);
          if (tRow) {
            const acc = lookupAccount_('teacher', String(tRow[1]).trim());
            if (acc && String(acc.email).includes('@')) email = String(acc.email).trim();
          }
        }
      }
    }
    if (!email.includes('@')) return err('ไม่พบชื่อ "' + emailOrName.trim() + '" ในระบบ กรุณาตรวจสอบชื่อ-นามสกุล หรือใช้ email แทน');
  }

  const emailLow = email.toLowerCase();

  // Verify email is registered
  let found = false;
  const usersSheet__ = sh(CONFIG.SHEETS.USERS);
  const usersData = usersSheet__ ? usersSheet__.getDataRange().getValues().slice(1) : [];
  if (usersData.find(r => String(r[0]).trim().toLowerCase() === emailLow)) found = true;

  if (!found) {
    const accSheet = sh(CONFIG.SHEETS.ACCOUNTS);
    if (accSheet) {
      const accData = accSheet.getDataRange().getValues().slice(1);
      if (accData.find(r => String(r[4]).trim().toLowerCase() === emailLow)) found = true;
    }
  }

  if (!found) return err('ไม่พบ email นี้ในระบบ กรุณาตรวจสอบว่าได้ลงทะเบียน email ไว้ในระบบหรือยัง');

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('otp_' + emailLow, otp, 300);
  try {
    MailApp.sendEmail({
      to: email.trim(),
      subject: 'รหัส OTP เข้าสู่ระบบไอแพด — โรงเรียนจักราชวิทยา',
      htmlBody: `<div style="font-family:sans-serif;max-width:420px;margin:auto;padding:24px">
        <h2 style="color:#4F7EF7;margin-bottom:8px">🔑 รหัส OTP ของท่าน</h2>
        <p style="color:#64748B;font-size:13px;margin-bottom:16px">ระบบบริหารจัดการไอแพด โรงเรียนจักราชวิทยา</p>
        <div style="font-size:44px;font-weight:900;letter-spacing:10px;color:#0F172A;background:#F0F4FF;padding:22px;border-radius:12px;text-align:center;border:2px solid #C7D7FD">${otp}</div>
        <p style="color:#64748B;margin-top:16px;font-size:13px">⏱ รหัสนี้หมดอายุใน <b>5 นาที</b><br>🔒 ห้ามเปิดเผยรหัสนี้แก่ผู้อื่น</p>
        <div style="background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;padding:10px 14px;margin-top:14px;font-size:12px;color:#92400E">
          📬 <b>ถ้าไม่เจออีเมล:</b> กรุณาตรวจสอบในโฟลเดอร์ <b>จดหมายขยะ (Spam / Junk)</b> ด้วย
        </div>
        <p style="color:#94A3B8;font-size:11px;margin-top:16px">ระบบบริหารจัดการไอแพด — โรงเรียนจักราชวิทยา</p>
      </div>`
    });
  } catch(e) { return err('ส่ง email ไม่สำเร็จ: ' + e.message); }
  log_('system', 'ขอ OTP', emailLow);
  return ok('ส่ง OTP ไปที่ ' + email.trim() + ' แล้ว', { resolvedEmail: email.trim() });
}

function verifyOTP(email, otp) {
  if (!email || !otp) return err('ข้อมูลไม่ครบ');
  const emailLow = email.trim().toLowerCase();
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'otp_' + emailLow;
  const attKey   = 'otp_att_' + emailLow;
  const stored   = cache.get(cacheKey);
  if (!stored) return err('OTP หมดอายุ กรุณาขอใหม่');
  const attempts = parseInt(cache.get(attKey) || '0', 10);
  if (attempts >= 5) {
    cache.remove(cacheKey);
    cache.remove(attKey);
    return err('ป้อน OTP ผิดเกินกำหนด OTP ถูกยกเลิก กรุณาขอใหม่');
  }
  if (stored !== String(otp).trim()) {
    cache.put(attKey, String(attempts + 1), 300);
    return err('OTP ไม่ถูกต้อง (เหลือ ' + (4 - attempts) + ' ครั้ง)');
  }
  cache.remove(cacheKey);
  cache.remove(attKey);

  // Check admin Users sheet first
  const usersSheetV_ = sh(CONFIG.SHEETS.USERS);
  const usersData = usersSheetV_ ? usersSheetV_.getDataRange().getValues().slice(1) : [];
  const adminUser = usersData.find(r => String(r[0]).trim().toLowerCase() === emailLow);
  if (adminUser && adminUser[3] !== CONFIG.ROLES.USER) {
    const userObj = { email: adminUser[0], name: adminUser[1], code: adminUser[2], role: adminUser[3], position: adminUser[4] };
    const tok = createSession_(userObj);
    log_(email, 'OTP Login (admin)', 'สำเร็จ');
    return ok('เข้าสู่ระบบสำเร็จ', Object.assign({ loginType: 'admin' }, userObj, { token: tok }));
  }

  // Check Accounts sheet (teachers/students)
  const accSheet = sh(CONFIG.SHEETS.ACCOUNTS);
  if (accSheet) {
    const accData = accSheet.getDataRange().getValues().slice(1);
    const acc = accData.find(r => String(r[4]).trim().toLowerCase() === emailLow);
    if (acc) {
      const type = String(acc[0]), code = String(acc[1]);
      accSheet.getRange(accData.indexOf(acc) + 2, 7).setValue(new Date());
      if (type === 'student') {
        const lookup = lookupStudentByCode(code);
        if (!lookup.found) return err('ไม่พบข้อมูลนักเรียนในระบบ');
        log_(email, 'OTP Login (student)', 'สำเร็จ');
        return ok('เข้าสู่ระบบสำเร็จ', { loginType:'student', student: lookup.student, ipad: lookup.ipad });
      } else {
        const tSheet_ = sh(CONFIG.SHEETS.TEACHERS);
        let teacher = null;
        if (tSheet_ && tSheet_.getLastRow() >= 2) {
          const tData_ = tSheet_.getDataRange().getValues().slice(1);
          const tRow_ = tData_.find(r => String(r[1]).trim() === code);
          if (tRow_) teacher = { code, prefix: String(tRow_[2]), firstName: String(tRow_[3]), lastName: String(tRow_[4]), subject: String(tRow_[5] || ''), room: String(tRow_[6] || '') };
        }
        const accEmail = String(acc[4] || '').trim();
        log_(email, 'OTP Login (teacher)', 'สำเร็จ');
        return ok('เข้าสู่ระบบสำเร็จ', { loginType:'teacher', code, email: accEmail, teacher });
      }
    }
  }

  return err('ไม่พบบัญชีสำหรับ email นี้');
}

// ── Export Reports ─────────────────────────────────────────────────────────────

function exportReport(type) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');

  let rows = [], headers = [];

  if (type === 'borrowed') {
    headers = ['Serial','รหัสครุภัณฑ์','คำนำหน้า','ชื่อ','นามสกุล','รหัส','ตำแหน่ง','ชั้น','ห้อง','วันที่ยืม'];
    rows = getAllIPads().filter(r => r.status === CONFIG.STATUS.BORROWED).map(r =>
      [r.serial, r.assetCode, r.prefix, r.firstName, r.lastName, r.personCode, r.position, r.grade, r.room, r.borrowDate]);
  } else if (type === 'available') {
    headers = ['Serial','รหัสครุภัณฑ์','หมายเหตุ'];
    rows = getAllIPads().filter(r => r.status === CONFIG.STATUS.AVAILABLE).map(r =>
      [r.serial, r.assetCode, r.notes]);
  } else if (type === 'students_no_ipad') {
    headers = ['รหัส','คำนำหน้า','ชื่อ','นามสกุล','ชั้น','ห้อง'];
    const ipads = getAllIPads();
    const assignedCodes = new Set(ipads.map(r => String(r.personCode).trim()).filter(Boolean));
    const assignedNames = new Set(ipads.map(r => String(r.firstName).trim()+'|'+String(r.lastName).trim()));
    const stuData = sh(CONFIG.SHEETS.STUDENTS).getDataRange().getValues().slice(1);
    stuData.forEach(r => {
      const code = String(r[1]).trim(), key = String(r[3]).trim()+'|'+String(r[4]).trim();
      if (!assignedCodes.has(code) && !assignedNames.has(key))
        rows.push([r[1], r[2], r[3], r[4], r[5], r[6]]);
    });
  } else if (type === 'issues') {
    headers = ['ID','รหัส','ชื่อ','Serial','ประเภท','รายละเอียด','สถานะ','วันที่แจ้ง','กำหนดส่ง'];
    const data = sh(CONFIG.SHEETS.ISSUES).getDataRange().getValues().slice(1);
    rows = data.map(r => [r[0],r[1],r[2],r[3],r[4],r[5],r[8],r[9],r[12]]);
  } else if (type === 'all_ipads') {
    headers = ['Serial','รหัสครุภัณฑ์','คำนำหน้า','ชื่อ','นามสกุล','รหัส','ตำแหน่ง','ชั้น','ห้อง','สถานะ','วันที่ยืม','วันที่คืน'];
    rows = getAllIPads().map(r =>
      [r.serial, r.assetCode, r.prefix, r.firstName, r.lastName, r.personCode, r.position, r.grade, r.room, r.status, r.borrowDate, r.returnDate]);
  } else {
    return err('ประเภทรายงานไม่ถูกต้อง');
  }

  const csvRows = [headers, ...rows].map(row =>
    row.map(c => {
      const s = String(c == null ? '' : c).replace(/"/g, '""');
      return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s}"` : s;
    }).join(',')
  );
  log_(u.email, 'Export รายงาน', type + ' | ' + rows.length + ' รายการ');
  return ok('', { csv: '﻿' + csvRows.join('\n'), count: rows.length, type });
}

// ── Batch Return ───────────────────────────────────────────────────────────────

function batchReturn(grade, room) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');

  const ipadSheet = sh(CONFIG.SHEETS.IPAD);
  const data = ipadSheet.getDataRange().getValues();
  const S = CONFIG.STATUS;
  const now = new Date();
  const problemStatuses = new Set([S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM]);
  let count = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const rowGrade = String(data[i][8]).trim();
    const rowRoom  = String(data[i][9]).trim();
    const status   = String(data[i][10]).trim();
    if (grade && rowGrade !== String(grade).trim()) continue;
    if (room  && rowRoom  !== String(room).trim())  continue;
    if (problemStatuses.has(status)) { skipped++; continue; }
    if (status !== S.BORROWED) continue;
    ipadSheet.getRange(i + 1, 11).setValue(S.RETURNED);
    ipadSheet.getRange(i + 1, 13).setValue(now);
    count++;
  }

  if (count > 0) invalidateCache_();
  const label = grade ? `${grade}${room ? '/'+room : ''}` : 'ทั้งหมด';
  log_(u.email, 'Batch คืน', `${label} จำนวน ${count} เครื่อง`);
  const skippedMsg = skipped > 0 ? ` (ข้าม ${skipped} เครื่องที่มีปัญหาอยู่)` : '';
  return ok(`คืนไอแพด ${count} เครื่องสำเร็จ${skippedMsg}`, { count, skipped });
}

// ── iPad Borrowing History ─────────────────────────────────────────────────────

function getIPadHistory(serial) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const logSheet = sh(CONFIG.SHEETS.LOG);
  if (!logSheet) return ok('', { history: [] });
  const data = logSheet.getDataRange().getValues().slice(1);
  const keyword = String(serial).trim().toLowerCase();
  const history = data
    .filter(r => String(r[3]).toLowerCase().includes(keyword))
    .map(r => ({ date: r[0] ? fmt(r[0]) : '', email: r[1], action: r[2], detail: r[3] }))
    .reverse()
    .slice(0, 50);
  return ok('', { history, serial });
}

// ── Dashboard Chart Data ───────────────────────────────────────────────────────

function getDashboardChartData() {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const ipads = getAllIPads();
  const S = CONFIG.STATUS;

  const statusCount = {};
  ipads.forEach(r => { statusCount[r.status] = (statusCount[r.status] || 0) + 1; });

  const gradeCount = {};
  ipads.filter(r => r.status === S.BORROWED && r.grade).forEach(r => {
    const g = String(r.grade).replace(/^ม\.\s*/,'');
    gradeCount[g] = (gradeCount[g] || 0) + 1;
  });

  // Daily borrow trend from Log (last 14 days)
  const logSheet = sh(CONFIG.SHEETS.LOG);
  const trend = {};
  if (logSheet) {
    const logData = logSheet.getDataRange().getValues().slice(1);
    const cutoff  = new Date(); cutoff.setDate(cutoff.getDate() - 14);
    logData.forEach(r => {
      if (!r[0] || new Date(r[0]) < cutoff) return;
      if (!String(r[2]).includes('ยืม') && !String(r[2]).includes('Borrow')) return;
      const d = Utilities.formatDate(new Date(r[0]), 'Asia/Bangkok', 'dd/MM');
      trend[d] = (trend[d] || 0) + 1;
    });
  }

  return ok('', { statusCount, gradeCount, trend });
}

// ── Setup ─────────────────────────────────────────────────────────────────────

function initializeSpreadsheet() {
  const adminEmail = Session.getActiveUser().getEmail();

  function makeSheet(name, headers, color) {
    if (ss().getSheetByName(name)) return;
    const s = ss().insertSheet(name);
    s.appendRow(headers);
    s.getRange('1:1').setFontWeight('bold').setBackground(color).setFontColor('white');
    s.setFrozenRows(1);
  }

  makeSheet(CONFIG.SHEETS.IPAD,
    ['ID','Serial','รหัสครุภัณฑ์','คำนำหน้า','ชื่อ','นามสกุล','รหัส','ตำแหน่ง','ระดับชั้น','ห้อง','สถานะ','วันที่ยืม','วันที่คืน','หมายเหตุ'], '#1a73e8');
  makeSheet(CONFIG.SHEETS.USERS,    ['Email','ชื่อ-นามสกุล','รหัส','Role','ตำแหน่ง'], '#0f9d58');
  makeSheet(CONFIG.SHEETS.LOG,      ['วันที่-เวลา','Email','การกระทำ','รายละเอียด'], '#db4437');
  makeSheet(CONFIG.SHEETS.STUDENTS, ['ลำดับ','รหัสนักเรียน','คำนำหน้า','ชื่อ','นามสกุล','ระดับชั้น','ห้อง','รูปนักเรียน'], '#f57c00');
  makeSheet(CONFIG.SHEETS.TEACHERS, ['ลำดับ','รหัสครู','คำนำหน้า','ชื่อ','นามสกุล','วิชา','ห้องประจำ'], '#e91e63');
  makeSheet(CONFIG.SHEETS.ADMIN,    ['user','รหัส','role'], '#6a1b9a');
  makeSheet(CONFIG.SHEETS.ISSUES,   ['ID','รหัสนักเรียน','ชื่อ','Serial','ประเภทปัญหา','รายละเอียด','ไฟล์แนบ','สถานะเดิม','สถานะปัญหา','วันที่แจ้ง','หมายเหตุAdmin','อัพเดทล่าสุด','กำหนดส่ง'], '#795548');
  makeSheet(CONFIG.SHEETS.ACCOUNTS, ['ประเภท','รหัส','Username','Password','Email','สร้างเมื่อ','Loginล่าสุด'], '#607d8b');

  const adminSheet = ss().getSheetByName(CONFIG.SHEETS.ADMIN);
  if (adminSheet.getLastRow() < 2) adminSheet.appendRow(['admin', '112233', 'super admin']);
  else ensureDefaultAdminAccount_();

  const usersSheet = ss().getSheetByName(CONFIG.SHEETS.USERS);
  const existing   = usersSheet.getDataRange().getValues();
  if (!existing.slice(1).some(r => r[0] === adminEmail)) {
    usersSheet.appendRow([adminEmail, 'Super Administrator', 'ADMIN001', CONFIG.ROLES.SUPER, 'ผู้ดูแลระบบ']);
  }

  SpreadsheetApp.getUi().alert('✅ เริ่มต้นระบบสำเร็จ\n\nสร้าง Sheet ครบแล้ว พร้อมใช้งาน');
}

// ── Unregistered Devices Report ───────────────────────────────────────────────
// Run this function from Apps Script editor → gets you a Google Sheet you can
// File > Download > Excel (.xlsx)

function runUnregisteredReport() {
  const UNREGISTERED_SERIALS = [
    'M2C217TNQ7','C36444YW7X','DY9MPG70PY','CPVYL97MF2','H400FJ409C','KG404H6270',
    'LXCY9Q942G','K6122XP3KW','HGH605G7CN','G94T6HXY20','LXK2D7J30J','FHYTR4LQ0L',
    'MWL5Q6264X','G7XCH2MVKD','MMFHFM07LW','F3XM9XPM61','H0F33H3JFC','G5VHWTJW6V',
    'LV3GCP7WGW','KFKVTDQ3GX','L30N2HFVV3','F61PQNH54R','JCXMLH073V','GWXL65PWPG',
    'D930XXG7L5','FD657R4QW4','JC4TQ7R7FK','K16J2CV69W','HCFC4R7WDN','FJ4GVKL74R',
    'J217KJW3TL','F57DY4251W','F9HWLH91VF','KTH6P493H4','MHWVQD796R','JQX4RY5W5H',
    'KMGF43Y7GX','MCYTG9691D','K27YGP6GP7','KXN3642VYJ','L1CYT6X71F','CTW9XJMWP1',
    'J27L403VKP','LL4RK7M0MP','CVYVJ2FMQP','G9HXGW265R','MLXVV7GR75','DG7R06XRP0',
    'L097THKKX4','FW2G65QJ6Y','C6WX6MCWRN','JM92RJJ4C9','G2CJP4K45G','CL4QYVGXF0',
    'CVMHG2QFNQ','KGK9P6PXJ0','G4JGK6R14T','JTH2R24M0M','KX37N36H36','C6XCHXL2KR',
    'K4XH74GW4C','J3DT2W4LXH','K5WYDFPHR2','M2RW4FV407','CQ2V123FHT','J46MVW4393',
    'KKWWLYWG4X','M7VY9W7C40','KTV7FPPVHD','GHHDF9P1V4','L504JNY66W','H2G2QDK9LQ',
    'H6Q59972JC','H66K30J92P','KQWP42LQ2Y','FJWQYMH20T','C7ML9KD9LM','CGGYKY07Q4',
    'G6933T2002','F490XV2547','JL4LDQYJWW','C6J49HGXNG','MNXRNWND7V','M73FX4J7P7',
    'HFW1R4JQPK','L17WP641H7','K767M3WQ2J','GV14LV771R','HT44X6XTG6','FD2YWN7K6W',
    'LR7F23R75R','K46DQYP9PG','J67YQ2DFH6','H9H90R96CM','CKFF7MTGQ9','KJJ9GCWYXW',
    'FQPFWT26QD','H2JN7T426H','FYXJ3XNHV6','GNVW6VJX9N','C2YC9VMP4X','FJFVRWW0XJ',
    'L794GJG9VW','GVH321M9V7','KW96QYTVPQ','JY2G5QL734','F0CFDWYYDF','LF73HTK4CV',
    'GHV9KLF7DH','HD47VWC04L','K7G4QTFPYQ','CN7PCWY4Y1','GFNY9XK40N','MQ6N16C2V2',
    'K619MFWM96','D9VYVP2VK2','L7TGW4PX7G','DJVHC070NK','F9RR9RFK67','M54THR42TH',
    'GDXL3R9JJV','KYPHYPPWQY','HGV9RRQ9J2','FCVJ2P52MX','CTW54DFVCG','FPVD97TTXY',
    'HMVVDGGK6F','M6435HHCYM','LKF7QM3H42','L92WH549JW','KHQN6VC6KK','CL9G0P2DQK',
    'HTWV6GYWCC','DC4W23GRX0','J2HGHLHP6M','LPFDFP6TWV','FV7QVD4QP5','LV1WRMQ9QN',
    'M12K40J411','FYJ4KGXKLW','GVYGRXD9W4','GJG2YWG02F','KTY9DC9XC2','H253J4JLY2',
    'HHCWP2742V','KJ7H029D20','J1QGYL6LQR','DHXN7FVQXP','MHG7MGYC41','M7X0MHH2JM',
    'C0M0HJ9JLY','KQJ9K7W50Y','L6JVVM0JGK','F774LVVD6D','H6RYVJPXL4','F6P39JP79N',
    'H64WN22VYY','JW9JV61MV2','G1Q96P0WYY','HFM77CT4D5','JJPGKMQDLP','DL2W7012YC',
    'MGQWWQT9XM','KP91L655W9','H33WX6X3HW','L1PW4CGWY9','G34DVXY9M5','JQD64X49P4',
    'K9X3W33473','KYFFVVD7FQ','G70QPYMJ6J','HX7MVK42R7','D6WYF6D6HT','JT6CF7LHVC',
    'GJC46KJTG2','MR9FJW21F3','CWHYH7T5F4','D349JWH21Y','LDQ379FDJ9','GQGPV4GW75',
    'CRWFF22DJP','F0CFC4G00W','C1FPWJK7M1','GG9PC764FC','FC7PN6QYFC','KWLD6HPQTV',
    'DQTHGPM6KF','L25917W4N3','LXMP4V4LF4','MQ4YV9G4MQ','GJFT9G5XF2','K7P4R39YKV',
    'L4WYKRQXL2','F5VH373V0R','FDYDC9HTPH','C4RXV471X9','FVMG0PKWNC','GN6LXFXJ3X',
    'J7RX4HFRX6','FHYD9NYQ59','H3F9HFWGK9','H5HC06G671','LX2NC4GWX7','JQDC65HX5H',
    'JL72V4CGJ9','KD2PRGC4D0','KCWJ41RY77','HWH9X4FVJD','FCY7VN46CF','D75C36JXCP',
    'KY0RM26X3P','M1VYXQ5F02','JXJWYG5V23','LG770YXRF0','J7T7GCWQWQ','CTPQQYN3L7',
    'GTPF03GHVV','GJYPQK3G00','GKM0P7D0H9','GL4VMKQK50','LJYR2YGVMR','M7T75WDFHQ',
    'J22PK949LY','KTCQ9HVWGM','DVH07R2743','F03HFXC99M','K07L45NHDQ','K4FDJ61QNY',
    'FGWQ02HNFJ','J0KKXR392T','D19VP6YTGQ','MGFPJCWRMJ','GNW7JG9LH4','DHXV2F17QK',
    'J5QPYX5FPG','FQMQ9W219H','H9L646LN97','H0CMWW1GK4','HPHW50M4RL','MH6C0FNVXL',
    'FVQ03W79F5','HQ5960V9Q6','CNG3PD2JK0','JG9HX7N29M','J274W661V9','C9K9MW64C4',
    'FW3TVW3K5V','KX4TQR7NQH','LLK9PF09MY','JVPQDM9HQF','GKCF0XJ020','DDXW77R37L',
    'GJ6NFPFHXF','CFJYW7XXKR','FDWKKWVW5M','GXX23LTXWX','F965WNWPFP','DL7CT9799H',
    'JLJWXXFR61','D590K00C4D','J9F7X5C3G7','F42MN24H56','H2WLT2LNP4','LVFT2YWVNM',
    'DNQFC64QM9','C7XJXJX0WP','KW36CQT29Q','JW9JHG7K2P','HD66JWG6C9','MT4Y3G5P4F',
    'KD72K50QT2','H3DY2C7QC9','D9HN2XYK66','L2LTV176TJ','JGT4NM2P3J','H4CHY29R7J',
    'J6JRVP2WT2','CPQQX0P6JN','JN02360XKK','H6P6PVC9T9','G32LQ41QNJ','KGKY60YF4V',
    'CD6LHF9F06','L00F7D57L2','CKXYKWQCYP','CK6LHH5HJN','GXM44QKWNY','LF4GJ33KF4',
    'CVVJ62K9YH','G99377NQH4','L07M2XVKDG','JXV7F9HP77','L4DD4527YL','KF6NKJ91Q7',
    'HQ91P9G2Y5','MNPC2RKX4Y','CJ3C2C4XDR','H5KQV3344D','G7H3Y7Q9XF','K00M706HL0',
    'GY42L0YVPK','MNQVV2LRP2','F7G3W9RWNW','MQ2MGXQQJQ','HXQV4R96VN','L26GNPH69R',
    'CVNC434JXJ','F07Y26VGKX','FQ7C3HJDFY','F4522PJW1J','LR7L0MHKQ6','HR2FNYJVX0',
    'FV2JWG7XKP','L56PGF65X9','LC2XQDNF3F','GHQFY27N3T','KGPDYPVVR2','G74RR44MDJ',
    'LVY7VX7J5W','M44LH29QYW','CCXPMY43TC','D6Q1C322QR','FWHKYX5609','JX0NV6DX01',
    'DQYQ09H4V2','GMVHGWYLVQ','F7QY6TL9V6','F99Q32FXC9','HYGVX9YYF7','GXQQ779K22',
    'FHNW77HH40','JX7LVJLCM2','G37M69FW95','CQJ43PPGHJ','GQFHC9RD9C','D4J94JQJC5',
    'J6G46P79JQ','KT76909THH','MJ6JKQKJ64','JYD76GFP95','K60QV0MHVW','K66726145X',
    'FN0Q552J5C','D926X4YGGK','DVP37LXWHC','K9DD7D4DG7','MP2WR759RJ','FHT7GJ6WTQ',
    'HVJ96QK9CH','CG7L4LP671','LYYXH4V4HP','HT9Y73YVXH','JR2507N9CP','FM7TJY6KQL',
    'G765NY9CTX','MVHH0Y04YC','CY4MY65G2M','G0FF6NF27M','DXJ4HR67GR','J59Q269KK2',
    'FY2FG3XYM3','D3NXM59YJD','CV643GVGFC','FPYF75H4MC','M3HPL7WF49','KW4K0GQ056',
    'HHJXJJMFMC','JP9334L509','C27F96PFQW','L93CMCQ7PG','G3JFJQ0MV3','MF7W6TXFVT',
    'H1G2DQKKWD','K33R7102VN','K4N02W7X6P','D57JGKLQ2R','CYKPXV27V0','LQ6FG402NV',
    'D1WQCV7K14','KV914LFTQ7','HQYQCHYXX3','JQ2R7RQRQX','HXGV3HT2JK','LQ45QCFNR2',
    'FXFHR72KL0','HQ9XT46TX4','KH0096RQQG','DK4HTK609V','JVXL17R622','JNQFGM627N',
    'C590TLW7NY','HY60P7129H','H24L02VXDR','L2FQ6Q7VXD'
  ];

  // Build lookup map from iPad_Data: serial → row data
  const ipadSheet = sh(CONFIG.SHEETS.IPAD);
  const ipadData  = ipadSheet.getDataRange().getValues();
  const ipadMap   = {};
  for (let i = 1; i < ipadData.length; i++) {
    const s = String(ipadData[i][1]).trim().toUpperCase();
    if (s) ipadMap[s] = ipadData[i];
  }

  // Build result rows
  const headers = [
    'ลำดับ','Serial Number','พบในระบบ',
    'รหัสครุภัณฑ์','คำนำหน้า','ชื่อ','นามสกุล',
    'รหัส','ตำแหน่ง','ระดับชั้น','ห้อง',
    'สถานะ','วันที่ยืม','วันที่คืน','หมายเหตุ'
  ];
  const resultRows = [headers];
  UNREGISTERED_SERIALS.forEach((serial, idx) => {
    const r = ipadMap[serial.toUpperCase()];
    if (r) {
      resultRows.push([
        idx + 1, serial, 'พบ',
        r[2], r[3], r[4], r[5],
        r[6], r[7], r[8], r[9],
        r[10],
        r[11] ? Utilities.formatDate(new Date(r[11]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
        r[12] ? Utilities.formatDate(new Date(r[12]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
        r[13] || ''
      ]);
    } else {
      resultRows.push([idx + 1, serial, 'ไม่พบ', '', '', '', '', '', '', '', '', '', '', '', '']);
    }
  });

  // Create result sheet inside this spreadsheet (overwrite if exists)
  const RESULT_SHEET_NAME = 'Unregistered_Report';
  let outSheet = ss().getSheetByName(RESULT_SHEET_NAME);
  if (outSheet) {
    outSheet.clearContents();
  } else {
    outSheet = ss().insertSheet(RESULT_SHEET_NAME);
  }

  outSheet.getRange(1, 1, resultRows.length, headers.length).setValues(resultRows);

  // Style header row
  const headerRange = outSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
  outSheet.setFrozenRows(1);
  outSheet.autoResizeColumns(1, headers.length);

  // Color-code: ไม่พบ → light red, พบ → light green
  for (let i = 2; i <= resultRows.length; i++) {
    const found = String(outSheet.getRange(i, 3).getValue());
    outSheet.getRange(i, 1, 1, headers.length)
      .setBackground(found === 'พบ' ? '#e6f4ea' : '#fce8e6');
  }

  const foundCount    = resultRows.slice(1).filter(r => r[2] === 'พบ').length;
  const notFoundCount = resultRows.slice(1).filter(r => r[2] === 'ไม่พบ').length;

  const url = ss().getUrl() + '#gid=' + outSheet.getSheetId();
  Logger.log(`✅ รายงานพร้อม: ${url}`);
  Logger.log(`พบในระบบ: ${foundCount} | ไม่พบ: ${notFoundCount} | รวม: ${UNREGISTERED_SERIALS.length}`);

  SpreadsheetApp.getUi().alert(
    `✅ รายงานพร้อมแล้ว\n\n` +
    `พบในระบบ: ${foundCount} เครื่อง\n` +
    `ไม่พบในระบบ: ${notFoundCount} เครื่อง\n` +
    `รวมทั้งหมด: ${UNREGISTERED_SERIALS.length} เครื่อง\n\n` +
    `ดูที่ Sheet: "${RESULT_SHEET_NAME}"\n` +
    `แล้ว File > Download > Microsoft Excel เพื่อรับไฟล์`
  );

  return { foundCount, notFoundCount, total: UNREGISTERED_SERIALS.length, sheetUrl: url };
}

// ── Web-callable: match uploaded serial list against iPad_Data ────────────────
function matchSerialsWithSystem(serials) {
  try {
    const ipadSheet = sh(CONFIG.SHEETS.IPAD);
    const ipadData  = ipadSheet.getDataRange().getValues();
    const ipadMap   = {};
    for (let i = 1; i < ipadData.length; i++) {
      const s = String(ipadData[i][1]).trim().toUpperCase();
      if (s) ipadMap[s] = ipadData[i];
    }

    const rows = serials.map((serial, idx) => {
      const key = String(serial).trim().toUpperCase();
      const r   = ipadMap[key];
      if (r) {
        return {
          no: idx + 1, serial: key, found: true,
          assetCode: r[2] || '', prefix: r[3] || '', firstName: r[4] || '', lastName: r[5] || '',
          personCode: r[6] || '', position: r[7] || '', grade: r[8] || '', room: r[9] || '',
          status: r[10] || '',
          borrowDate: r[11] ? Utilities.formatDate(new Date(r[11]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
          returnDate: r[12] ? Utilities.formatDate(new Date(r[12]), 'Asia/Bangkok', 'dd/MM/yyyy') : '',
          notes: r[13] || ''
        };
      }
      return { no: idx + 1, serial: key, found: false, assetCode:'', prefix:'', firstName:'', lastName:'', personCode:'', position:'', grade:'', room:'', status:'', borrowDate:'', returnDate:'', notes:'' };
    });

    const foundCount    = rows.filter(r => r.found).length;
    const notFoundCount = rows.filter(r => !r.found).length;
    return { ok: true, rows, foundCount, notFoundCount, total: serials.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── LINE Messaging API Webhook ────────────────────────────────────────────────

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var events = body.events || [];
    for (var i = 0; i < events.length; i++) {
      var ev = events[i];
      var userId = (ev.source || {}).userId || '';
      if (ev.type === 'follow') {
        lineReplyWelcome_(ev.replyToken);
      } else if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
        lineHandleMessage_(ev.replyToken, userId, ev.message.text);
      }
    }
  } catch (err) {
    log_('system', 'LINE webhook error', err.message);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lineReplyWelcome_(replyToken) {
  var props = PropertiesService.getScriptProperties();
  var token = String(props.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '').trim();
  if (!token) return;

  var gasUrl = 'https://script.google.com/macros/s/AKfycby079h5QsFuJtoOWk9E2-jJK1uhxnjlpw1Jrg-HthTxa5CeN15CRPyZCVxvHRYYdstmUQ/exec?authuser=0';
  var installUrl = 'https://wisanu15.github.io/IPAD_JV/';
  var manualUrl  = gasUrl + '&section=manual';

  var flex = {
    type: 'flex',
    altText: 'ยินดีต้อนรับสู่ระบบ iPad JV',
    contents: {
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#4F7EF7',
        paddingAll: '20px',
        contents: [
          {
            type: 'text',
            text: '📱 ระบบบริหารจัดการไอแพด',
            color: '#FFFFFF',
            size: 'xl',
            weight: 'bold',
            wrap: true
          },
          {
            type: 'text',
            text: 'โรงเรียนจักราชวิทยา',
            color: '#FFFFFFCC',
            size: 'sm',
            margin: 'xs'
          }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '20px',
        spacing: 'md',
        contents: [
          {
            type: 'text',
            text: 'ยินดีต้อนรับ! 👋',
            size: 'lg',
            weight: 'bold',
            color: '#0F172A'
          },
          {
            type: 'text',
            text: 'ระบบนี้ใช้สำหรับบริหารจัดการไอแพดของโรงเรียน สามารถใช้งานได้ดังนี้',
            size: 'sm',
            color: '#475569',
            wrap: true
          },
          {
            type: 'separator',
            margin: 'md'
          },
          {
            type: 'box',
            layout: 'vertical',
            spacing: 'sm',
            margin: 'md',
            contents: [
              lineFeatureRow_('🔎', 'ตรวจสถานะ', 'ดูว่าไอแพดของตัวเองอยู่ที่ไหน'),
              lineFeatureRow_('🔔', 'แจ้งปัญหา', 'แจ้งเมื่อไอแพดหาย เสีย หรือมีปัญหา'),
              lineFeatureRow_('🔄', 'ยืม / คืน', 'บันทึกการยืม-คืนไอแพด'),
              lineFeatureRow_('📲', 'ขอแอพ', 'ยื่นคำร้องขอติดตั้งแอพพลิเคชัน'),
              lineFeatureRow_('📥', 'ติดตั้งแอป', 'ติดตั้งไว้ที่หน้าจอหลักเพื่อใช้งานง่าย')
            ]
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        spacing: 'sm',
        paddingAll: '16px',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: '#4F7EF7',
            height: 'sm',
            action: { type: 'uri', label: '📘 เปิดคู่มือการใช้งาน', uri: manualUrl }
          },
          {
            type: 'button',
            style: 'secondary',
            height: 'sm',
            action: { type: 'uri', label: '📥 ติดตั้งแอป (Add to Home Screen)', uri: installUrl }
          }
        ]
      }
    }
  };

  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: replyToken, messages: [flex] }),
      muteHttpExceptions: true
    });
  } catch (err) {
    log_('system', 'LINE reply welcome error', err.message);
  }
}

function cleanOldSessionTokens() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var deleted = [];
  Object.keys(all).forEach(function(k) {
    if (k.startsWith('_remember_') || k.startsWith('_dev_')) {
      props.deleteProperty(k);
      deleted.push(k);
    }
  });
  Logger.log('Deleted ' + deleted.length + ' tokens: ' + deleted.join(', '));
  return 'Deleted: ' + deleted.length;
}

function lineFeatureRow_(emoji, title, desc) {
  return {
    type: 'box',
    layout: 'horizontal',
    spacing: 'md',
    contents: [
      { type: 'text', text: emoji, size: 'md', flex: 0 },
      {
        type: 'box',
        layout: 'vertical',
        flex: 1,
        contents: [
          { type: 'text', text: title, size: 'sm', weight: 'bold', color: '#0F172A' },
          { type: 'text', text: desc,  size: 'xs', color: '#64748B', wrap: true }
        ]
      }
    ]
  };
}

// ── น้องแพด AI Chat ───────────────────────────────────────────────────────────

var LINE_GAS_BASE_ = 'https://script.google.com/macros/s/AKfycby079h5QsFuJtoOWk9E2-jJK1uhxnjlpw1Jrg-HthTxa5CeN15CRPyZCVxvHRYYdstmUQ/exec?authuser=0';
var LINE_INSTALL_URL_ = 'https://wisanu15.github.io/IPAD_JV/';

function lineReply_(token, replyToken, messages) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: replyToken, messages: messages.slice(0, 5) }),
      muteHttpExceptions: true
    });
  } catch (e) { log_('system', 'LINE reply error', e.message); }
}

function fetchStudentRecord_(query) {
  try {
    var code = String(query || '').trim();
    if (!code) return null;
    // Use the existing lookupStudentByCode (searches Students sheet + iPad data)
    var r = lookupStudentByCode(code);
    if (r && r.found) {
      return {
        prefix: r.student.prefix, firstName: r.student.firstName,
        lastName: r.student.lastName, grade: r.student.grade, room: r.student.room,
        serial: r.ipad ? r.ipad.serial : '(ยังไม่ได้รับมอบหมาย)',
        status: r.ipad ? r.ipad.status : '-'
      };
    }
    // Fallback: search by name / serial using searchStudentsForIssue
    var hits = searchStudentsForIssue(code);
    if (hits && hits.length > 0) {
      var s = hits[0];
      return {
        prefix: s.prefix, firstName: s.firstName, lastName: s.lastName,
        grade: s.grade, room: s.room,
        serial: s.serial || '(ยังไม่ได้รับมอบหมาย)', status: '-'
      };
    }
    return null;
  } catch (e) { return null; }
}


function geminiChat_(text) {
  try {
    var props = PropertiesService.getScriptProperties();
    var groqKey   = String(props.getProperty('GROQ_API_KEY')   || '').trim();
    var geminiKey = String(props.getProperty('GEMINI_API_KEY') || '').trim();
    if (!groqKey && !geminiKey) return null;

    var sysPrompt =
      'คุณคือ "น้องแพด" ผู้ช่วย AI ระบบบริหารจัดการไอแพด โรงเรียนจักราชวิทยา\n' +
      'บุคลิก: เป็นกันเอง ใจดี ตอบสั้น ภาษาไทยวัยรุ่น ใส่อิโมจิพอเหมาะ\n' +
      'ตอบเป็น JSON เท่านั้น:\n' +
      '{"intent":"<intent>","reply":"<ข้อความตอบไทย>","studentCode":"<รหัส/ชื่อที่ถาม หรือ null>"}\n\n' +
      'intent: chat|lookup|status|issue|borrow|return|appreq|manual|install|admin|other\n' +
      'lookup = ถามข้อมูลนักเรียน/serial → ใส่รหัสหรือชื่อใน studentCode\n' +
      'ตอบ 1-2 ประโยค';

    var parsed = null;

    if (groqKey) {
      var gr = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + groqKey },
        payload: JSON.stringify({
          model: 'llama-3.1-8b-instant',
          messages: [{ role: 'system', content: sysPrompt }, { role: 'user', content: text }],
          response_format: { type: 'json_object' },
          max_tokens: 300, temperature: 0.4
        })
      });
      if (gr.getResponseCode() === 200) {
        var gd = JSON.parse(gr.getContentText());
        parsed = JSON.parse(((gd.choices || [])[0] || {}).message.content || '{}');
      } else {
        log_('system', 'Groq HTTP ' + gr.getResponseCode(), gr.getContentText().substring(0, 200));
      }
    }

    if (!parsed && geminiKey) {
      var gm = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + geminiKey,
        {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({
            contents: [{ parts: [{ text: sysPrompt + '\n\nข้อความผู้ใช้: ' + text }] }],
            generationConfig: { responseMimeType: 'application/json', maxOutputTokens: 300, temperature: 0.4 }
          })
        }
      );
      if (gm.getResponseCode() === 200) {
        var md = JSON.parse(gm.getContentText());
        var mparts = (((md.candidates || [])[0] || {}).content || {}).parts || [];
        parsed = JSON.parse((mparts[0] || {}).text || '{}');
      } else {
        log_('system', 'Gemini HTTP ' + gm.getResponseCode(), gm.getContentText().substring(0, 200));
      }
    }

    return parsed;
  } catch (e) {
    log_('system', 'AI exception', e.message);
    return null;
  }
}

function lineHandleMessage_(replyToken, userId, text) {
  var props = PropertiesService.getScriptProperties();
  var token = String(props.getProperty('LINE_CHANNEL_ACCESS_TOKEN') || '').trim();
  if (!token) return;

  var modeKey = '_linemode_' + userId;
  var mode = props.getProperty(modeKey) || 'ai';

  // User wants AI back
  if (/^(น้องแพด|\/ai|bot|ai)$/i.test(text.trim())) {
    props.deleteProperty(modeKey);
    lineReply_(token, replyToken, [{
      type: 'text',
      text: 'กลับมาแล้วครับ! 😊 น้องแพด AI พร้อมช่วยอีกครั้งแล้ว\n\n🤖 น้องแพด'
    }]);
    return;
  }

  // Human mode — AI stays silent, let admin handle
  if (mode === 'human') return;

  // Keyword matching first — saves Gemini quota
  var t = text.trim().toLowerCase();
  var fastIntent = null;
  if (/แจ้งปัญหา|ไอแพด.*หาย|หาย.*ไอแพด|ไอแพด.*เสีย|เสีย.*ไอแพด|ชำรุด|มีรอย/.test(t)) fastIntent = 'issue';
  else if (/^ยืม|ยืม.*ไอแพด|ไอแพด.*ยืม/.test(t))                                       fastIntent = 'borrow';
  else if (/^คืน|คืน.*ไอแพด|ไอแพด.*คืน/.test(t))                                       fastIntent = 'return';
  else if (/ขอแอพ|ขอ.*แอพ|ขอ.*app/i.test(t))                                            fastIntent = 'appreq';
  else if (/คู่มือ|วิธีใช้งาน/.test(t))                                                 fastIntent = 'manual';
  else if (/ติดตั้งแอป|add to home|pwa/.test(t))                                        fastIntent = 'install';
  else if (/คุยกับแอดมิน|ขอแอดมิน|ติดต่อแอดมิน/.test(t))                              fastIntent = 'admin';

  var intent = fastIntent || 'other';
  var aiReply = '';
  var ai = null;

  // Only call Gemini when keyword matching can't decide
  if (!fastIntent) {
    ai = geminiChat_(text);
    intent = (ai && ai.intent) || 'other';
    aiReply = (ai && ai.reply) || '';
  }

  var msgs = [];

  if (intent === 'lookup') {
    // Get search query from Gemini, fallback to extracting digits/text from original message
    var q = (ai && ai.studentCode && String(ai.studentCode) !== 'null') ? String(ai.studentCode).trim() : '';
    if (!q) {
      var numMatch = text.match(/\d{3,}/);
      if (numMatch) q = numMatch[0];
    }
    if (!q) q = text.replace(/รหัส|นักเรียน|คือใคร|ไอแพด|serial|ชั้น|ห้อง|ของใคร|อยู่ที่ไหน/g, '').trim();

    if (q) {
      var found = fetchStudentRecord_(q);
      if (found) {
        msgs.push({ type: 'text',
          text: '📋 ผลค้นหา "' + q + '"\n' +
                '👤 ' + found.prefix + found.firstName + ' ' + found.lastName + '\n' +
                '📚 ชั้น ' + found.grade + '/' + found.room + '\n' +
                '📱 Serial: ' + found.serial + '\n' +
                '✅ สถานะ: ' + found.status });
      } else {
        msgs.push({ type: 'text', text: 'ขอโทษนะครับ 😅 ไม่พบข้อมูล "' + q + '" ในระบบเลยครับ' });
      }
    } else {
      msgs.push({ type: 'text', text: 'บอกรหัสนักเรียนหรือชื่อที่ต้องการค้นหามาด้วยนะครับ 😊' });
    }

  } else if (intent === 'admin') {
    props.setProperty(modeKey, 'human');
    lineReply_(token, replyToken, [{
      type: 'text',
      text: 'ได้เลยครับ! 👨‍💼 ส่งต่อให้แอดมินแล้ว\nรอสักครู่นะครับ แอดมินจะมาตอบเอง\n\n(พิมพ์ "น้องแพด" เมื่อต้องการกลับมาคุยกับ AI อีกครั้ง)'
    }]);
    return;

  } else if (intent === 'chat' || intent === 'other') {
    msgs.push({ type: 'text', text: aiReply || 'สวัสดีครับ 😊 กดเมนูด้านล่างเพื่อใช้งานระบบได้เลยนะครับ' });

  } else if (intent === 'status') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    if (!studentData) {
      msgs.push(lineQuickFlex_('🔎 ตรวจสถานะไอแพด',
        'ดูสถานะไอแพดของตัวเอง ว่าอยู่ที่ไหน ใครถือครองอยู่', '#4F7EF7',
        [{ label: '🔎 ตรวจสถานะ', uri: LINE_GAS_BASE_ + '&section=status' }]));
    }

  } else if (intent === 'issue') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineQuickFlex_('🔔 แจ้งปัญหาไอแพด',
      'Admin จะได้รับแจ้งทันทีครับ', '#F59E0B',
      [{ label: '🔔 แจ้งปัญหา', uri: LINE_GAS_BASE_ + '&section=issue' }]));

  } else if (intent === 'borrow') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineQuickFlex_('📱 ยืมไอแพด', 'บันทึกการยืมไอแพด', '#8B5CF6',
      [{ label: '📱 ยืมไอแพด', uri: LINE_GAS_BASE_ + '&section=register' }]));

  } else if (intent === 'return') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineQuickFlex_('🔄 คืนไอแพด', 'บันทึกการคืนไอแพด', '#8B5CF6',
      [{ label: '🔄 คืนไอแพด', uri: LINE_GAS_BASE_ + '&section=return' }]));

  } else if (intent === 'appreq') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineQuickFlex_('📲 ขอแอพพลิเคชัน', 'ยื่นคำร้องขอให้ Admin ติดตั้งแอพ', '#06B6D4',
      [{ label: '📲 ขอแอพ', uri: LINE_GAS_BASE_ + '&section=appreq' }]));

  } else if (intent === 'manual') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineMenuFlex_(LINE_GAS_BASE_, LINE_INSTALL_URL_));

  } else if (intent === 'install') {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    msgs.push(lineQuickFlex_('📥 ติดตั้งแอป', 'เพิ่ม iPad JV ที่หน้าจอหลัก', '#EF4444',
      [{ label: '📥 ติดตั้งแอป', uri: LINE_INSTALL_URL_ },
       { label: '📘 คู่มือ', uri: LINE_GAS_BASE_ + '&section=manual' }]));

  } else {
    if (aiReply) msgs.push({ type: 'text', text: aiReply });
    else msgs.push(lineDefaultFlex_(LINE_GAS_BASE_, LINE_INSTALL_URL_));
  }

  if (msgs.length === 0) return;

  // Append AI indicator to the last text message
  var last = msgs[msgs.length - 1];
  if (last.type === 'text') last.text += '\n\n🤖 น้องแพด';

  lineReply_(token, replyToken, msgs);
}

function lineQuickFlex_(title, desc, color, buttons) {
  var btns = buttons.map(function(b) {
    return {
      type: 'button', style: 'primary', color: color, height: 'sm',
      action: { type: 'uri', label: b.label, uri: b.uri }
    };
  });
  return {
    type: 'flex', altText: title,
    contents: {
      type: 'bubble',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: color, paddingAll: '16px',
        contents: [{ type: 'text', text: title, color: '#FFFFFF', size: 'lg', weight: 'bold', wrap: true }]
      },
      body: {
        type: 'box', layout: 'vertical', paddingAll: '16px',
        contents: [{ type: 'text', text: desc, size: 'sm', color: '#475569', wrap: true }]
      },
      footer: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: btns
      }
    }
  };
}

function lineMenuFlex_(gasBase, installUrl) {
  return {
    type: 'flex', altText: 'เมนูระบบ iPad JV',
    contents: {
      type: 'bubble', size: 'giga',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#4F7EF7', paddingAll: '16px',
        contents: [
          { type: 'text', text: '📱 ระบบบริหารจัดการไอแพด', color: '#FFFFFF', size: 'lg', weight: 'bold', wrap: true },
          { type: 'text', text: 'เลือกเมนูที่ต้องการได้เลย', color: '#FFFFFFCC', size: 'xs', margin: 'xs' }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: [
          lineMenuBtn_('🔎 ตรวจสถานะ',  'ดูข้อมูลไอแพด',      '#4F7EF7', gasBase + '&section=status'),
          lineMenuBtn_('🔔 แจ้งปัญหา',   'ไอแพดหาย/เสีย',      '#F59E0B', gasBase + '&section=issue'),
          lineMenuBtn_('🔄 ยืม / คืน',   'บันทึกยืม-คืนไอแพด', '#8B5CF6', gasBase + '&section=register'),
          lineMenuBtn_('📲 ขอแอพ',       'ขอติดตั้งแอพพลิเคชัน','#06B6D4', gasBase + '&section=appreq'),
          lineMenuBtn_('📘 คู่มือ',       'วิธีใช้งานระบบ',      '#10B981', gasBase + '&section=manual'),
          lineMenuBtn_('📥 ติดตั้งแอป',  'Add to Home Screen',  '#EF4444', installUrl)
        ]
      }
    }
  };
}

function lineMenuBtn_(label, desc, color, uri) {
  return {
    type: 'box', layout: 'horizontal', spacing: 'md',
    backgroundColor: color + '18', borderWidth: '1px', borderColor: color + '44',
    cornerRadius: '10px', paddingAll: '12px',
    action: { type: 'uri', uri: uri },
    contents: [
      { type: 'box', layout: 'vertical', flex: 0, justifyContent: 'center',
        contents: [{ type: 'text', text: label.split(' ')[0], size: 'xl' }] },
      { type: 'box', layout: 'vertical', flex: 1, spacing: 'xs',
        contents: [
          { type: 'text', text: label.replace(/^.\s/, ''), size: 'sm', weight: 'bold', color: '#0F172A' },
          { type: 'text', text: desc, size: 'xs', color: '#64748B' }
        ]
      },
      { type: 'text', text: '›', size: 'xl', color: color, flex: 0, gravity: 'center' }
    ]
  };
}

function lineDefaultFlex_(gasBase, installUrl) {
  return {
    type: 'flex', altText: 'กรุณาใช้ปุ่มเมนูด้านล่างเพื่อเข้าสู่ระบบ',
    contents: {
      type: 'bubble',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#4F7EF7', paddingAll: '16px',
        contents: [
          { type: 'text', text: '📱 ระบบบริหารจัดการไอแพด', color: '#FFFFFF', size: 'md', weight: 'bold', wrap: true }
        ]
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'md', paddingAll: '18px',
        contents: [
          {
            type: 'box', layout: 'horizontal', spacing: 'sm', backgroundColor: '#EFF6FF',
            cornerRadius: '10px', paddingAll: '12px',
            contents: [
              { type: 'text', text: '👇', size: 'xl', flex: 0 },
              { type: 'text', text: 'กรุณากด ปุ่มเมนูด้านล่าง เพื่อใช้งานระบบได้เลย ไม่ต้องพิมพ์', size: 'sm', color: '#1D4ED8', wrap: true, weight: 'bold' }
            ]
          },
          {
            type: 'box', layout: 'vertical', spacing: 'xs',
            contents: [
              { type: 'text', text: 'เมนูที่ใช้บ่อย:', size: 'xs', color: '#94A3B8', weight: 'bold' },
              { type: 'text', text: '🔎 ตรวจสถานะ  •  🔔 แจ้งปัญหา  •  🔄 ยืม/คืน', size: 'xs', color: '#475569', wrap: true },
              { type: 'text', text: '📲 ขอแอพ  •  📘 คู่มือ  •  📥 ติดตั้งแอป', size: 'xs', color: '#475569', wrap: true }
            ]
          },
          {
            type: 'box', layout: 'horizontal', spacing: 'sm', backgroundColor: '#FEF9EC',
            cornerRadius: '8px', paddingAll: '10px', margin: 'md',
            contents: [
              { type: 'text', text: '💬', size: 'md', flex: 0 },
              { type: 'text', text: 'ถ้าเมนูใช้ไม่ได้หรือต้องการความช่วยเหลือ ค่อยพิมพ์คุยกับแอดมินได้เลย', size: 'xs', color: '#92400E', wrap: true }
            ]
          }
        ]
      },
      footer: {
        type: 'box', layout: 'vertical', spacing: 'sm', paddingAll: '12px',
        contents: [
          { type: 'button', style: 'primary', color: '#4F7EF7', height: 'sm',
            action: { type: 'uri', label: '📱 เปิดระบบ', uri: gasBase } },
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'uri', label: '📘 ดูคู่มือการใช้งาน', uri: gasBase + '&section=manual' } },
          { type: 'button', style: 'secondary', height: 'sm',
            action: { type: 'message', label: '👨‍💼 คุยกับแอดมิน', text: 'ขอคุยกับแอดมินครับ' } }
        ]
      }
    }
  };
}
