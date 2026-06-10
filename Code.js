// ============================================================
// Code.gs — iPad Management System v2
// ============================================================

const CONFIG = {
  SHEETS: {
    IPAD: 'iPad_Data', USERS: 'Users', LOG: 'Log', STUDENTS: 'Students',
    ADMIN: 'Admin', DATABASE: 'Database', ISSUES: 'Issues', ACCOUNTS: 'Accounts',
    TEACHERS: 'Teachers'
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
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
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

// ── iPad CRUD ─────────────────────────────────────────────────────────────────

function getAllIPads() {
  const cache  = getCache_();
  const cached = cache.get('all_ipads');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  const data = sh(CONFIG.SHEETS.IPAD).getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[0] === '' && r[1] === '') continue;
    rows.push({
      row: i + 1, id: r[0], serial: String(r[1]), assetCode: r[2],
      prefix: r[3], firstName: r[4], lastName: r[5], personCode: r[6],
      position: r[7], grade: r[8], room: r[9], status: r[10],
      borrowDate: r[11] ? fmt(r[11]) : '',
      returnDate: r[12] ? fmt(r[12]) : '',
      notes: r[13] || ''
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

  const stats = {
    total:            totalReal,
    available:        totalReal - borrowed,
    borrowed:         borrowed,
    returned:         ipads.filter(r => r.status === S.RETURNED).length,
    claimed:          ipads.filter(r => r.status === S.CLAIMED).length,
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
    if (!stats.gradeMap[g][rm]) stats.gradeMap[g][rm] = { total:0, available:0, borrowed:0, returned:0, claimed:0, pending:0 };
    stats.gradeMap[g][rm].total++;
    if      (r.status === S.AVAILABLE)         stats.gradeMap[g][rm].available++;
    else if (r.status === S.BORROWED)          stats.gradeMap[g][rm].borrowed++;
    else if (r.status === S.RETURNED)          stats.gradeMap[g][rm].returned++;
    else if (r.status === S.CLAIMED)           stats.gradeMap[g][rm].claimed++;
    else                                       stats.gradeMap[g][rm].pending++;
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
  getAllIPads().forEach(r => rows.push([
    r.id, r.serial, r.assetCode, r.prefix, r.firstName, r.lastName,
    r.personCode, r.position, r.grade, r.room, r.status,
    r.borrowDate, r.returnDate, `"${String(r.notes).replace(/"/g,'""')}"`
  ].join(',')));
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
  notifyAdminPush_('🔔 ทดสอบ', 'ทดสอบการแจ้งเตือนจาก iPad JV', '', {});
}

function notifyAdminPush_(title, message, url, data) {
  try {
    const props = PropertiesService.getScriptProperties();
    const appId = String(props.getProperty('ONESIGNAL_APP_ID') || '').trim();
    const apiKey = String(props.getProperty('ONESIGNAL_REST_API_KEY') || '').trim();
    if (!appId || !apiKey) return;

    const payload = {
      app_id: appId,
      target_channel: 'push',
      filters: [{ field: 'tag', key: 'role', relation: '=', value: 'admin' }],
      headings: { en: title || 'iPad JV', th: title || 'iPad JV' },
      contents: { en: message || '', th: message || '' },
      url: url || String(props.getProperty('IPAD_JV_PWA_URL') || 'https://wisanu15.github.io/IPAD_JV/'),
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
      return ok('เข้าสู่ระบบสำเร็จ', Object.assign({}, current, { token: tok }));
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
  const assigned = {}, assetBySerial = {};
  ipads.forEach(r => {
    if (r.personCode) assigned[r.serial] = true;
    if (r.assetCode)  assetBySerial[r.serial] = r.assetCode;
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
    return ipads.filter(r => r.status === CONFIG.STATUS.AVAILABLE && !r.personCode)
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
  const u = Session.getActiveUser().getEmail() || 'visitor';
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
  if (count > 0) { invalidateCache_(); log_(u, 'ลงทะเบียนไอแพด', `${grade}/${room} จำนวน ${count} เครื่อง`); }
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
    log_(u || 'visitor', 'แจ้งเตือน admin', `Manual entry: ${lines.join('; ')}`);
  }
  if (errors.length && !count) return err(errors.join('\n'));
  return ok(`ลงทะเบียนสำเร็จ ${count} เครื่อง` + (errors.length ? `\n⚠️ ${errors.join(', ')}` : ''));
}

function registerIPadBatchMultiple(classes) {
  let totalCount = 0;
  const results = [], allErrors = [];
  for (const cls of classes) {
    if (!cls.items || !cls.items.length) continue;
    const r = registerIPadBatch(cls.items, cls.grade, cls.room);
    results.push(`${cls.grade}/${cls.room}: ${r.message}`);
    if (!r.success) { allErrors.push(`${cls.grade}/${cls.room}: ${r.message}`); }
    else { const m = r.message.match(/(\d+)/); if (m) totalCount += Number(m[1]); }
  }
  if (allErrors.length && !totalCount) return err(allErrors.join('\n'));
  return ok(`ลงทะเบียนสำเร็จรวม ${totalCount} เครื่อง จาก ${classes.length} ห้อง` +
    (allErrors.length ? `\n⚠️ บางห้องมีข้อผิดพลาด` : ''), { results });
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
  const stats = {
    total:          stuTotal,
    borrowed:       stuBorrowed,
    borrowedFromDB: ipads.filter(r => r.status === S.BORROWED).length,
    returned:       stuReturned,
    available:      ipads.filter(r => r.position === 'นักเรียน' && r.status === S.AVAILABLE).length,
    claimed:        ipads.filter(r => r.position === 'นักเรียน' && r.status === S.CLAIMED).length,
    availableFromDB: availCount,
    problem:        ipads.filter(r => [S.PENDING_ISSUE, S.UNDER_INSPECTION, S.DEFECTIVE, S.UNDER_CLAIM, S.CLAIMED].includes(r.status)).length,
    totalIPads:     ipads.length,
    totalFromDB:    totalFromDB || (availCount + ipads.filter(r => r.personCode).length)
  };

  const teacherList = ipads.filter(r => r.position === 'ครู').map(r => ({
    prefix: r.prefix, firstName: r.firstName, lastName: r.lastName,
    personCode: r.personCode, serial: r.serial, assetCode: r.assetCode,
    status: r.status, borrowDate: r.borrowDate
  }));
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
    rows.push({ row: i + 1, seq: r[0], code: String(r[1]), prefix: r[2], firstName: r[3], lastName: r[4], grade: r[5], room: r[6] });
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
  log_(u.email, 'เพิ่มนักเรียน', `${d.prefix}${d.firstName} ${d.lastName}`);
  return ok('เพิ่มนักเรียนสำเร็จ');
}

function updateStudent(row, d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  sh(CONFIG.SHEETS.STUDENTS).getRange(row, 2, 1, 6).setValues([[d.code, d.prefix, d.firstName, d.lastName, d.grade, d.room]]);
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
    rows.push({ row: i + 1, seq: r[0], code: String(r[1]), prefix: r[2], firstName: r[3], lastName: r[4], subject: r[5] || '', room: r[6] || '' });
  }
  if (filters && filters.query) {
    const q = filters.query.toLowerCase();
    rows = rows.filter(r => r.code.includes(q) || r.firstName.toLowerCase().includes(q) || r.lastName.toLowerCase().includes(q) || (r.subject || '').toLowerCase().includes(q));
  }
  return rows;
}

function addTeacher(d) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role === CONFIG.ROLES.USER) return err('ไม่มีสิทธิ์');
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return err('ไม่พบ Sheet Teachers');
  sheet.appendRow(['', d.code || '', d.prefix || 'นาย', d.firstName, d.lastName, d.subject || '', d.room || '']);
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

function deleteTeacher(row) {
  const u = getCurrentUser(arguments[arguments.length - 1]);
  if (u.role !== CONFIG.ROLES.SUPER) return err('เฉพาะ Super Admin เท่านั้น');
  const name = sh(CONFIG.SHEETS.TEACHERS).getRange(row, 4).getValue();
  sh(CONFIG.SHEETS.TEACHERS).deleteRow(row);
  renumberTeachers_();
  log_(u.email, 'ลบครู', name);
  return ok('ลบสำเร็จ');
}

function renumberTeachers_() {
  const sheet = sh(CONFIG.SHEETS.TEACHERS);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;
  for (let i = 2; i <= last; i++) sheet.getRange(i, 1).setValue(i - 1);
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

  // Prevent double-submit: reject if an open issue already exists for this serial
  const existingRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 4, sheet.getLastRow() - 1, 6).getValues()
    : [];
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
    '',
    { issueId: newId, serial: lookup.ipad.serial, studentCode: data.studentCode }
  );
  log_(data.studentCode, 'แจ้งปัญหาไอแพด', `Serial: ${lookup.ipad.serial} | ${data.issueType}`);
  return ok('แจ้งปัญหาสำเร็จ เจ้าหน้าที่จะดำเนินการโดยเร็ว');
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

function requestOTP(email) {
  if (!email) return err('กรุณากรอก email');
  const emailLow = email.trim().toLowerCase();

  // Check admin Users sheet
  let found = false;
  const usersData = sh(CONFIG.SHEETS.USERS).getDataRange().getValues().slice(1);
  if (usersData.find(r => String(r[0]).trim().toLowerCase() === emailLow)) found = true;

  // Check Accounts sheet (teachers/students who registered an email)
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
  return ok('ส่ง OTP ไปที่ ' + email.trim() + ' แล้ว');
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
  const usersData = sh(CONFIG.SHEETS.USERS).getDataRange().getValues().slice(1);
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
        log_(email, 'OTP Login (teacher)', 'สำเร็จ');
        return ok('เข้าสู่ระบบสำเร็จ', { loginType:'teacher', code });
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
    .filter(r => String(r[3]).toLowerCase().includes(keyword) || String(r[2]).toLowerCase().includes('serial') && String(r[3]).toLowerCase().includes(keyword))
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
      const d = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM');
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
