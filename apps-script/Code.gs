// ============================================================
// HRM NGHỈ PHÉP — Google Apps Script v3.2
// Vai trò: employee, lecturer, manager, manager_lecturer, hr, admin
// admin = Quản trị hệ thống: toàn quyền (mọi chức năng của Nhân sự + thêm/sửa/xóa mọi dữ liệu)
// Năm phép tính theo NĂM HỌC: 01/07 → 30/06 năm sau
// CCCD và ngày tháng luôn được lưu dạng VĂN BẢN trong Google Sheet
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1hRx_unC2Hx5Afg1ZoE5FUoQhbXwjSTWA6I-0ekS25dc',
  // Địa chỉ giao diện trên Cloudflare Pages (dùng cho link trong email) — SỬA THÀNH ĐỊA CHỈ THẬT
  APP_URL: 'https://hrm-nghiphep.pages.dev/',
  VERSION: 'HRM_V320',

  SHEETS: {
    EMPLOYEES:        'EMPLOYEES',
    LEAVE_REQUESTS:   'LEAVE_REQUESTS',
    LOGS:             'LOGS',
    HOLIDAY_SETTINGS: 'HOLIDAY_SETTINGS'
  },
  HEADERS: {
    EMPLOYEES:        ['cccd','name','department','role','email','start_date'],
    LEAVE_REQUESTS:   ['id','employee_id','employee_name','department','from_date','to_date','days',
                       'reason','status','approver_id','approver_note','hr_confirm','hr_note',
                       'created_at','updated_at','attachment','history','half_day','return_info'],
    LOGS:             ['timestamp','action','user_id','detail'],
    HOLIDAY_SETTINGS: ['year','type','from_date','to_date','note','created_by','created_at']
  },
  HEADER_COLORS: {
    EMPLOYEES: '#1A3260', LEAVE_REQUESTS: '#34A853', LOGS: '#EA4335', HOLIDAY_SETTINGS: '#F39C12'
  },
  // Cột (tính từ 1) luôn định dạng văn bản: giữ số 0 đầu CCCD, tránh Sheet tự đổi ngày
  TEXT_COLUMNS: {
    EMPLOYEES:        [1, 6],               // cccd, start_date
    LEAVE_REQUESTS:   [1, 2, 5, 6, 10, 12], // id, employee_id, from_date, to_date, approver_id, hr_confirm
    LOGS:             [3],                  // user_id
    HOLIDAY_SETTINGS: [3, 4, 6]             // from_date, to_date, created_by
  },

  ANNUAL_LEAVE_DEFAULT:   12,
  LEAVE_YEAR_START_MONTH: 7,                // Năm học bắt đầu 01/07
  TOKEN_TTL_MS:           8 * 60 * 60 * 1000,

  ROLES: {
    EMPLOYEE: 'employee', LECTURER: 'lecturer', MANAGER: 'manager',
    MANAGER_LECTURER: 'manager_lecturer', HR: 'hr', ADMIN: 'admin'
  },
  ADMIN_PW_MIN_LENGTH:  8,
  ADMIN_PW_MAX_FAILS:   5,                  // Sai quá 5 lần → khóa tạm
  ADMIN_PW_LOCK_SEC:    15 * 60,            // Khóa 15 phút
  STATUS: {
    PENDING:          'pending',
    APPROVED:         'approved',
    REJECTED:         'rejected',
    HR_CONFIRMED:     'hr_confirmed',
    HR_RETURNED:      'hr_returned',
    RETURNED:         'returned',          // Đã trả phép toàn bộ
    RETURN_PENDING:   'return_pending',
    RETURN_CONFIRMED: 'return_confirmed',
    RETURN_REJECTED:  'return_rejected'
  },
  PARTIAL_SUFFIX: '_partial_returned',
  HOLIDAY_TYPES: ['summer', 'tet']
};

// Chỉ số cột (0-based)
const LR = { ID:0, EMP_ID:1, EMP_NAME:2, DEPT:3, FROM:4, TO:5, DAYS:6, REASON:7, STATUS:8,
             APPROVER_ID:9, APPROVER_NOTE:10, HR_ID:11, HR_NOTE:12, CREATED:13, UPDATED:14,
             ATTACH:15, HISTORY:16, HALF_DAY:17, RETURN_INFO:18 };
const EC = { CCCD:0, NAME:1, DEPT:2, ROLE:3, EMAIL:4, START:5 };
const HC = { YEAR:0, TYPE:1, FROM:2, TO:3, NOTE:4, BY:5, AT:6 };

// ============================================================
// VAI TRÒ & TRẠNG THÁI
// ============================================================

function isLecturer(role) { return role === CONFIG.ROLES.LECTURER || role === CONFIG.ROLES.MANAGER_LECTURER; }
function isManager(role)  { return role === CONFIG.ROLES.MANAGER  || role === CONFIG.ROLES.MANAGER_LECTURER; }
function isAdmin(role)    { return role === CONFIG.ROLES.ADMIN; }
/** Quyền Nhân sự — Quản trị hệ thống cũng có toàn bộ quyền này */
function isHR(role)       { return role === CONFIG.ROLES.HR || isAdmin(role); }
function canApprove(role) { return isManager(role) || isHR(role); }

/** Bỏ hậu tố _partial_returned (kể cả khi bị lặp) để lấy trạng thái gốc */
function baseStatus(st) { return String(st || '').replace(/(_partial_returned)+$/, ''); }
/** Đơn đã được tính vào số ngày phép đã dùng */
function isUsedStatus(st) { var b = baseStatus(st); return b === CONFIG.STATUS.APPROVED || b === CONFIG.STATUS.HR_CONFIRMED; }
/** Đơn đang giữ chỗ quỹ phép (chưa duyệt xong) */
function isReservedStatus(st) { return st === CONFIG.STATUS.PENDING || st === CONFIG.STATUS.HR_RETURNED; }
/** Đơn đã đóng, không còn hiệu lực */
function isClosedStatus(st) { return st === CONFIG.STATUS.REJECTED || st === CONFIG.STATUS.RETURNED; }

// ============================================================
// TIỆN ÍCH CHUNG
// ============================================================

function ok_(obj)    { return Object.assign({ success: true }, obj || {}); }
function fail_(msg, extra) { return Object.assign({ success: false, message: msg }, extra || {}); }
function authFail_() { return { success: false, auth_error: true, message: 'Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại' }; }

function parseJSON_(raw, def) {
  if (!raw) return def;
  try { return JSON.parse(raw); } catch (_) { return def; }
}

function normalizeCCCD(cccd) {
  if (cccd === null || cccd === undefined) return '';
  var s = String(cccd).trim().replace(/\D/g, '');
  if (!s) return '';
  while (s.length < 12) s = '0' + s;
  return s;
}

function roundHalf_(n) { return Math.round((parseFloat(n) || 0) * 2) / 2; }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return fail_('Hệ thống đang bận, vui lòng thử lại sau vài giây');
  try { return fn(); } finally { lock.releaseLock(); }
}

// ============================================================
// NGÀY THÁNG
// ============================================================

var _tzCache = null;
function tz_() {
  if (!_tzCache) _tzCache = ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  return _tzCache;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function ymd_(y, m, d) {
  var dt = new Date(y, m, d);
  return (dt.getFullYear() === y && dt.getMonth() === m && dt.getDate() === d) ? dt : null;
}

/** Chuyển giá trị ô (Date, 'dd/MM/yyyy', 'yyyy-MM-dd') thành Date lúc 00:00 */
function toDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v)) return null;
    v = Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var s = String(v).trim(), m;
  if ((m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/))) return ymd_(+m[3], +m[2] - 1, +m[1]);
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)))          return ymd_(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function fmtDMY(d) { return d ? pad2_(d.getDate()) + '/' + pad2_(d.getMonth() + 1) + '/' + d.getFullYear() : ''; }
function cellDMY_(v) { var d = toDate(v); return d ? fmtDMY(d) : String(v || ''); }
function isoOf_(v)   { return v instanceof Date ? v.toISOString() : String(v || ''); }

function today_() { var n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); }

function calcWorkDays(start, end) {
  var c = 0, d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  while (d <= end) { var w = d.getDay(); if (w !== 0 && w !== 6) c++; d.setDate(d.getDate() + 1); }
  return c;
}

// ============================================================
// NĂM PHÉP = NĂM HỌC (01/07 → 30/06)
// ============================================================

/** Năm bắt đầu của năm học chứa ngày d. VD: 15/03/2027 → 2026 (năm học 2026-2027) */
function leaveYearStartOf(d) {
  d = d || today_();
  return d.getMonth() >= CONFIG.LEAVE_YEAR_START_MONTH - 1 ? d.getFullYear() : d.getFullYear() - 1;
}

function leaveYearRange(startYear) {
  var m0 = CONFIG.LEAVE_YEAR_START_MONTH - 1;
  var from = new Date(startYear, m0, 1);
  var to   = new Date(startYear + 1, m0, 0); // ngày cuối tháng trước tháng bắt đầu → 30/06
  return { start_year: startYear, from: from, to: to, label: startYear + '-' + (startYear + 1),
           from_str: fmtDMY(from), to_str: fmtDMY(to) };
}

// ============================================================
// THÂM NIÊN & SỐ NGÀY PHÉP (Điều 113 BLLĐ 2019)
// 12 ngày + 1 ngày cho mỗi 5 năm làm việc, tối đa 18 ngày
// ============================================================

function calcSeniorityYears(startDate) {
  var start = toDate(startDate);
  if (!start) return 0;
  var now = today_();
  var years = now.getFullYear() - start.getFullYear();
  if (now.getMonth() < start.getMonth() ||
     (now.getMonth() === start.getMonth() && now.getDate() < start.getDate())) years--;
  return Math.max(0, years);
}

function calcAnnualLeaveDays(startDate) {
  if (!toDate(startDate)) return CONFIG.ANNUAL_LEAVE_DEFAULT;
  return Math.min(12 + Math.floor(calcSeniorityYears(startDate) / 5), 18);
}

// ============================================================
// TRUY CẬP SHEET (có cache trong 1 lần thực thi)
// ============================================================

var _ssInst = null, _rowsCache = {};

function ss_() { if (!_ssInst) _ssInst = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); return _ssInst; }
function sheet_(name) { return ss_().getSheetByName(name); }

function rows_(name) {
  if (!_rowsCache[name]) {
    var sh = sheet_(name);
    _rowsCache[name] = sh ? sh.getDataRange().getValues() : [];
  }
  return _rowsCache[name];
}

function invalidate_(name) { delete _rowsCache[name]; }

function isTextCol_(name, col) { return (CONFIG.TEXT_COLUMNS[name] || []).indexOf(col) !== -1; }

/** Thêm dòng mới, định dạng TEXT cho các cột cần thiết TRƯỚC khi ghi giá trị */
function appendRow_(name, values) {
  var sh = sheet_(name);
  var r  = sh.getLastRow() + 1;
  (CONFIG.TEXT_COLUMNS[name] || []).forEach(function (c) {
    if (c <= values.length) sh.getRange(r, c).setNumberFormat('@');
  });
  sh.getRange(r, 1, 1, values.length).setValues([values]);
  invalidate_(name);
  return r;
}

/** Ghi nhiều ô trên 1 dòng. cells = { soCot: giaTri } */
function setCells_(name, row, cells) {
  var sh = sheet_(name);
  Object.keys(cells).forEach(function (k) {
    var col = parseInt(k, 10);
    var rg  = sh.getRange(row, col);
    if (isTextCol_(name, col)) rg.setNumberFormat('@');
    rg.setValue(cells[k]);
  });
  invalidate_(name);
}

function logActivity(action, userId, detail) {
  try {
    if (!sheet_(CONFIG.SHEETS.LOGS)) return;
    appendRow_(CONFIG.SHEETS.LOGS, [new Date().toISOString(), action, normalizeCCCD(userId) || String(userId || ''), detail || '']);
  } catch (_) {}
}

// ============================================================
// NHÂN VIÊN
// ============================================================

/** Đọc sheet EMPLOYEES, gộp các dòng cùng CCCD (nhiều đơn vị) */
function loadEmployeeMap_() {
  var data = rows_(CONFIG.SHEETS.EMPLOYEES);
  var map = {}, order = [];
  for (var i = 1; i < data.length; i++) {
    var cccd = normalizeCCCD(data[i][EC.CCCD]);
    if (!cccd) continue;
    var dept = String(data[i][EC.DEPT] || '').trim();
    if (!map[cccd]) {
      map[cccd] = {
        cccd:        cccd,
        name:        String(data[i][EC.NAME] || ''),
        department:  dept,
        departments: dept ? [dept] : [],
        role:        String(data[i][EC.ROLE] || CONFIG.ROLES.EMPLOYEE).trim(),
        email:       String(data[i][EC.EMAIL] || '').trim(),
        start_date:  cellDMY_(data[i][EC.START])
      };
      order.push(cccd);
    } else if (dept && map[cccd].departments.indexOf(dept) === -1) {
      map[cccd].departments.push(dept);
    }
  }
  return { map: map, order: order };
}

/** Tổng ngày đã dùng / đang chờ theo CCCD trong 1 năm học (đọc sheet 1 lần) */
function leaveUsageMap_(startYear) {
  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  var range = leaveYearRange(startYear);
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][LR.ID]) continue;
    var from = toDate(data[i][LR.FROM]);
    if (!from || from < range.from || from > range.to) continue;
    var cccd = normalizeCCCD(data[i][LR.EMP_ID]);
    var st   = String(data[i][LR.STATUS] || '');
    var days = parseFloat(data[i][LR.DAYS]) || 0;
    var u = map[cccd] || (map[cccd] = { used: 0, pending: 0 });
    if (isUsedStatus(st))          u.used    += days;
    else if (isReservedStatus(st)) u.pending += days;
  }
  return map;
}

function withLeave_(e, usage, startYear) {
  var range  = leaveYearRange(startYear);
  var annual = calcAnnualLeaveDays(e.start_date);
  var u      = usage[e.cccd] || { used: 0, pending: 0 };
  var used   = roundHalf_(u.used);
  return Object.assign({}, e, {
    annual_leave_days:    annual,
    used_leave_days:      used,
    pending_leave_days:   roundHalf_(u.pending),
    remaining_leave_days: annual - used,
    seniority_years:      calcSeniorityYears(e.start_date),
    leave_year_start:     startYear,
    leave_year:           range.label,
    leave_year_from:      range.from_str,
    leave_year_to:        range.to_str
  });
}

function findEmployeeByCCCD(cccd, startYear) {
  var norm = normalizeCCCD(cccd);
  if (!norm) return null;
  var e = loadEmployeeMap_().map[norm];
  if (!e) return null;
  var ly = startYear || leaveYearStartOf();
  return withLeave_(e, leaveUsageMap_(ly), ly);
}

function sanitizeEmp(e) {
  return {
    cccd: e.cccd, name: e.name, email: e.email,
    department: e.department, departments: e.departments,
    role: e.role, start_date: e.start_date, seniority_years: e.seniority_years,
    annual_leave_days: e.annual_leave_days, used_leave_days: e.used_leave_days,
    pending_leave_days: e.pending_leave_days, remaining_leave_days: e.remaining_leave_days,
    leave_year_start: e.leave_year_start, leave_year: e.leave_year,
    leave_year_from: e.leave_year_from, leave_year_to: e.leave_year_to,
    is_lecturer: isLecturer(e.role), is_manager: isManager(e.role), is_hr: isHR(e.role), is_admin: isAdmin(e.role)
  };
}

// ============================================================
// ENTRY POINT
// ============================================================

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = String(p.action || p.path || '').trim().toLowerCase();
  if (!action) {
    // Giao diện đã chuyển sang Cloudflare Pages → mở URL Web App sẽ dẫn người dùng sang đó
    var url = CONFIG.APP_URL;
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<div style="font-family:Arial,sans-serif;padding:40px;text-align:center">' +
      '<p>Hệ thống quản lý nghỉ phép đã chuyển sang địa chỉ mới:</p>' +
      '<p><a href="' + url + '" target="_top" style="font-size:18px">' + url + '</a></p></div>')
      .setTitle('HRM · Sổ tay Quản lý nghỉ phép');
  }
  return jsonOut_(route(action, p, p.token || ''));
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  var body = p;
  if (e && e.postData && e.postData.contents) body = parseJSON_(e.postData.contents, p);
  var action = String(p.action || body.action || '').trim().toLowerCase();
  return jsonOut_(route(action, body, p.token || body.token || ''));
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Cầu nối cho google.script.run từ index.html */
function handleApiCall(payloadJson) {
  try {
    var data = JSON.parse(payloadJson);
    return route(String(data.action || '').trim().toLowerCase(), data, data.token || '');
  } catch (err) {
    Logger.log('[handleApiCall] ' + err);
    return fail_('Lỗi xử lý yêu cầu: ' + err);
  }
}

// ============================================================
// ROUTER
// ============================================================

function route(action, data, token) {
  data = data || {};
  try {
    switch (action) {
      case 'login':               return login(data.cccd);
      case 'login-with-dept':     return loginWithDept(data.cccd, data.department);
      case 'login-admin':         return loginAdmin(data.cccd, data.password, data.new_password);
      case 'me':                  return getMe(token);
      case 'logout':              return doLogout(token);

      case 'leave-request':       return createLeaveRequest(token, data);
      case 'leave-requests':      return getLeaveRequests(token, data);
      case 'approve':             return approveRequest(token, data);
      case 'reject':              return rejectRequest(token, data);
      case 'hr-confirm':          return hrConfirm(token, data);
      case 'hr-return':           return hrReturn(token, data);

      case 'create-return':       return createReturnRequest(token, data);
      case 'hr-confirm-return':   return hrConfirmReturn(token, data);
      case 'hr-reject-return':    return hrRejectReturn(token, data);

      case 'stats':               return getStats(token);
      case 'employee-leave-info': return getEmployeeLeaveInfo(token, data);
      case 'employees':           return (data.cccd && data.name) ? createEmployee(token, data) : getEmployees(token);
      case 'create-employee':     return createEmployee(token, data);
      case 'update-employee':     return updateEmployee(token, data);

      case 'holiday-settings':    return ok_({ holidays: getHolidaySettings(data.year) });
      case 'save-holiday':        return saveHolidaySetting(token, data);
      case 'delete-holiday':      return deleteHolidaySetting(token, data);

      case 'ping':                return ok_({ version: CONFIG.VERSION, message: 'pong', ts: new Date().toISOString() });
      case 'setup':               return requireAdmin_(token, setupSheets);
      case 'fix-cccd':
      case 'fix-data':            return requireAdmin_(token, fixData);

      // ── Quản trị hệ thống (chỉ admin) ──
      case 'admin-create-request': return requireAdmin_(token, function (me) { return adminCreateRequest(me, data); });
      case 'admin-update-request': return requireAdmin_(token, function (me) { return adminUpdateRequest(me, data); });
      case 'admin-delete-request': return requireAdmin_(token, function (me) { return adminDeleteRequest(me, data); });
      case 'delete-employee':      return requireAdmin_(token, function (me) { return deleteEmployee(me, data); });
      case 'logs':                 return requireAdmin_(token, function () { return getLogs(data); });
      case 'system-info':          return requireAdmin_(token, getSystemInfo);
      case 'logout-all':           return requireAdmin_(token, function (me) { return logoutAll(me, token); });
      case 'admin-reset-password': return requireAdmin_(token, function (me) { return adminResetPassword(me, data); });
      case 'change-password':      return requireAdmin_(token, function (me) { return changeOwnPassword(me, data); });
      default:                    return fail_('Unknown action: "' + action + '"', { version: CONFIG.VERSION });
    }
  } catch (err) {
    Logger.log('[route] ' + action + ' ERROR: ' + (err && err.stack || err));
    return fail_('Lỗi hệ thống: ' + err);
  }
}

function requireAdmin_(token, fn) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isAdmin(emp.role)) return fail_('Chỉ Quản trị hệ thống mới có quyền');
  return fn(emp);
}

// ============================================================
// ĐĂNG NHẬP / TOKEN
// ============================================================

function login(cccd) {
  var norm = normalizeCCCD(cccd);
  if (!norm || String(cccd).replace(/\D/g, '').length !== 12) return fail_('Vui lòng nhập đủ 12 số CCCD');
  var emp = findEmployeeByCCCD(norm);
  if (!emp) return fail_('CCCD không tồn tại. Liên hệ phòng Nhân sự.');

  // Quản trị hệ thống: bắt buộc thêm mật khẩu
  if (isAdmin(emp.role)) {
    return ok_({ needPassword: true, hasPassword: !!getAdminPw_(emp.cccd), name: emp.name, cccd: emp.cccd,
                 min_length: CONFIG.ADMIN_PW_MIN_LENGTH });
  }

  if (isManager(emp.role) && emp.departments.length > 1) {
    return ok_({ needSelectDept: true, departments: emp.departments, name: emp.name, cccd: emp.cccd });
  }
  return issueToken_(emp, emp.department);
}

function loginWithDept(cccd, department) {
  if (!cccd || !department) return fail_('Thiếu thông tin');
  var emp = findEmployeeByCCCD(cccd);
  if (!emp) return fail_('CCCD không tồn tại');
  if (isAdmin(emp.role)) return fail_('Tài khoản Quản trị phải đăng nhập bằng mật khẩu');
  if (emp.departments.indexOf(department) === -1) return fail_('Đơn vị không hợp lệ');
  emp.department = department;
  return issueToken_(emp, department);
}

// ============================================================
// MẬT KHẨU QUẢN TRỊ (lưu dạng băm SHA-256 + salt trong Script Properties)
// Quên mật khẩu: admin khác đặt lại trong trang Hệ thống, hoặc xóa thuộc tính
// ADMINPW_<cccd> trong Apps Script → Cài đặt dự án → Thuộc tính tập lệnh.
// ============================================================

function hashPw_(salt, pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function getAdminPw_(cccd) { return parseJSON_(PropertiesService.getScriptProperties().getProperty('ADMINPW_' + cccd), null); }
function setAdminPw_(cccd, pw) {
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('ADMINPW_' + cccd, JSON.stringify({ salt: salt, hash: hashPw_(salt, pw), at: new Date().toISOString() }));
}
function validatePw_(pw) {
  pw = String(pw || '');
  if (pw.length < CONFIG.ADMIN_PW_MIN_LENGTH) return 'Mật khẩu phải có ít nhất ' + CONFIG.ADMIN_PW_MIN_LENGTH + ' ký tự';
  if (!/[A-Za-z]/.test(pw) || !/\d/.test(pw)) return 'Mật khẩu phải có cả chữ và số';
  return '';
}

function loginAdmin(cccd, password, newPassword) {
  var emp = findEmployeeByCCCD(cccd);
  if (!emp || !isAdmin(emp.role)) return fail_('Tài khoản không phải Quản trị hệ thống');

  var cache = CacheService.getScriptCache();
  var failKey = 'ADMINFAIL_' + emp.cccd;
  var fails = parseInt(cache.get(failKey) || '0', 10);
  if (fails >= CONFIG.ADMIN_PW_MAX_FAILS) return fail_('Sai mật khẩu quá nhiều lần. Vui lòng thử lại sau 15 phút.');

  var stored = getAdminPw_(emp.cccd);
  if (!stored) {
    // Lần đầu: tạo mật khẩu
    var err = validatePw_(newPassword);
    if (err) return fail_(err);
    setAdminPw_(emp.cccd, newPassword);
    logActivity('ADMIN_SET_PASSWORD', emp.cccd, 'Tạo mật khẩu lần đầu');
  } else if (hashPw_(stored.salt, String(password || '')) !== stored.hash) {
    cache.put(failKey, String(fails + 1), CONFIG.ADMIN_PW_LOCK_SEC);
    logActivity('ADMIN_LOGIN_FAIL', emp.cccd, 'Sai mật khẩu lần ' + (fails + 1));
    var left = CONFIG.ADMIN_PW_MAX_FAILS - fails - 1;
    return fail_(left > 0 ? 'Mật khẩu không đúng. Còn ' + left + ' lần thử.' : 'Sai mật khẩu quá nhiều lần. Vui lòng thử lại sau 15 phút.');
  }
  cache.remove(failKey);
  return issueToken_(emp, emp.department);
}

function changeOwnPassword(me, data) {
  var stored = getAdminPw_(me.cccd);
  if (stored && hashPw_(stored.salt, String(data.old_password || '')) !== stored.hash) return fail_('Mật khẩu hiện tại không đúng');
  var err = validatePw_(data.new_password);
  if (err) return fail_(err);
  setAdminPw_(me.cccd, data.new_password);
  logActivity('ADMIN_CHANGE_PASSWORD', me.cccd, '');
  return ok_({ message: 'Đã đổi mật khẩu' });
}

/** Xóa mật khẩu của một admin khác → lần đăng nhập sau họ sẽ tạo mật khẩu mới */
function adminResetPassword(me, data) {
  var target = findEmployeeByCCCD(data.cccd);
  if (!target || !isAdmin(target.role)) return fail_('Không tìm thấy tài khoản Quản trị');
  if (target.cccd === me.cccd) return fail_('Dùng chức năng "Đổi mật khẩu" cho tài khoản của bạn');
  PropertiesService.getScriptProperties().deleteProperty('ADMINPW_' + target.cccd);
  CacheService.getScriptCache().remove('ADMINFAIL_' + target.cccd);
  logActivity('ADMIN_RESET_PASSWORD', me.cccd, target.cccd);
  return ok_({ message: 'Đã đặt lại mật khẩu cho ' + target.name + '. Người này sẽ tạo mật khẩu mới ở lần đăng nhập tới.' });
}

function issueToken_(emp, department) {
  purgeExpiredTokens_();
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('TOKEN_' + token, JSON.stringify({
    cccd: emp.cccd, dept: department || '', expiry: Date.now() + CONFIG.TOKEN_TTL_MS
  }));
  logActivity('LOGIN', emp.cccd, 'dept=' + department + ' role=' + emp.role);
  return ok_({ token: token, user: sanitizeEmp(emp) });
}

function verifyToken(token) {
  if (!token) return null;
  try {
    var props = PropertiesService.getScriptProperties();
    var d = parseJSON_(props.getProperty('TOKEN_' + token), null);
    if (!d) return null;
    if (Date.now() > d.expiry) { props.deleteProperty('TOKEN_' + token); return null; }
    var emp = findEmployeeByCCCD(d.cccd);
    if (emp && d.dept && emp.departments.indexOf(d.dept) !== -1) emp.department = d.dept;
    return emp;
  } catch (_) { return null; }
}

function purgeExpiredTokens_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties(), now = Date.now();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('TOKEN_') !== 0) return;
      var d = parseJSON_(all[k], null);
      if (!d || now > d.expiry) props.deleteProperty(k);
    });
  } catch (_) {}
}

function getMe(token) {
  var emp = verifyToken(token);
  return emp ? ok_({ user: sanitizeEmp(emp) }) : authFail_();
}

function doLogout(token) {
  if (token) try { PropertiesService.getScriptProperties().deleteProperty('TOKEN_' + token); } catch (_) {}
  return ok_({ message: 'Đã đăng xuất' });
}

// ============================================================
// ĐƠN NGHỈ PHÉP
// ============================================================

function mapRequest_(row) {
  var from = toDate(row[LR.FROM]);
  return {
    id:            String(row[LR.ID]),
    employee_id:   normalizeCCCD(row[LR.EMP_ID]),
    employee_name: String(row[LR.EMP_NAME] || ''),
    department:    String(row[LR.DEPT] || ''),
    from_date:     cellDMY_(row[LR.FROM]),
    to_date:       cellDMY_(row[LR.TO]),
    days:          parseFloat(row[LR.DAYS]) || 0,
    reason:        String(row[LR.REASON] || ''),
    status:        String(row[LR.STATUS] || CONFIG.STATUS.PENDING),
    approver_id:   normalizeCCCD(row[LR.APPROVER_ID]),
    approver_note: String(row[LR.APPROVER_NOTE] || ''),
    hr_confirm:    normalizeCCCD(row[LR.HR_ID]),
    hr_note:       String(row[LR.HR_NOTE] || ''),
    created_at:    isoOf_(row[LR.CREATED]),
    updated_at:    isoOf_(row[LR.UPDATED]),
    attachment:    String(row[LR.ATTACH] || ''),
    history:       parseJSON_(row[LR.HISTORY], []),
    half_day:      String(row[LR.HALF_DAY] || 'none'),
    return_info:   parseJSON_(row[LR.RETURN_INFO], null),
    leave_year:    from ? leaveYearRange(leaveYearStartOf(from)).label : ''
  };
}

function findRequestRow_(requestId) {
  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][LR.ID]) === String(requestId)) return { row: i + 1, data: data[i] };
  }
  return null;
}

/** Mã đơn duy nhất: LR + thời điểm (ms). Gọi bên trong khóa ghi. */
function newRequestId_() {
  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS), used = {};
  for (var i = 1; i < data.length; i++) used[String(data[i][LR.ID])] = 1;
  var n = Date.now();
  while (used['LR' + n]) n++;
  return 'LR' + n;
}

function getRequestById(requestId) {
  var f = findRequestRow_(requestId);
  return f ? mapRequest_(f.data) : null;
}

function historyEntry_(emp, status, note, roleOverride) {
  return {
    time: new Date().toISOString(), status: status,
    actor_id: emp.cccd, actor_name: emp.name, actor_dept: emp.department,
    role: roleOverride || emp.role, note: note || ''
  };
}

/** Trùng lịch với đơn còn hiệu lực. Cho phép sáng + chiều cùng một ngày. */
function hasConflict_(cccd, from, to, halfDay, excludeId) {
  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  var norm = normalizeCCCD(cccd);
  for (var i = 1; i < data.length; i++) {
    if (!data[i][LR.ID] || normalizeCCCD(data[i][LR.EMP_ID]) !== norm) continue;
    if (excludeId && String(data[i][LR.ID]) === String(excludeId)) continue;
    if (isClosedStatus(String(data[i][LR.STATUS]))) continue;
    var eF = toDate(data[i][LR.FROM]), eT = toDate(data[i][LR.TO]);
    if (!eF || !eT || from > eT || to < eF) continue;
    var otherHalf = String(data[i][LR.HALF_DAY] || 'none');
    if (halfDay !== 'none' && otherHalf !== 'none' && otherHalf !== halfDay) continue;
    return true;
  }
  return false;
}

function createLeaveRequest(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();

  var reason = String(data.reason || '').trim();
  if (!data.from_date || !data.to_date) return fail_('Thiếu ngày bắt đầu/kết thúc');
  if (!reason) return fail_('Vui lòng nhập lý do');

  var from = toDate(data.from_date), to = toDate(data.to_date);
  if (!from || !to) return fail_('Ngày không hợp lệ. Dùng định dạng dd/mm/yyyy.');
  if (from > to)    return fail_('Ngày bắt đầu sau ngày kết thúc');

  // Không cho một đơn vắt qua 2 năm học
  var ly = leaveYearStartOf(from);
  var range = leaveYearRange(ly);
  if (leaveYearStartOf(to) !== ly) {
    return fail_('Đơn nghỉ không được vượt qua ngày ' + range.to_str + ' (hết năm học ' + range.label +
                 '). Vui lòng tách thành 2 đơn.');
  }

  // Kỳ nghỉ chính thức của giảng viên
  if (isLecturer(emp.role)) {
    var holiday = findHolidayOverlap(from, to);
    if (holiday) {
      var typeLabel = holiday.type === 'summer' ? 'nghỉ hè' : 'nghỉ Tết';
      return fail_('Khoảng thời gian này thuộc kỳ ' + typeLabel + ' (' + holiday.from_date + ' đến ' +
                   holiday.to_date + '). Giảng viên không cần tạo đơn trong kỳ nghỉ này.',
                   { holiday_type: holiday.type, is_holiday_block: true });
    }
  }

  // Số ngày (hỗ trợ nửa ngày)
  var halfDay = ['morning', 'afternoon'].indexOf(data.half_day) !== -1 ? data.half_day : 'none';
  var days;
  if (halfDay !== 'none') {
    if (from.getTime() !== to.getTime()) return fail_('Nghỉ nửa ngày chỉ áp dụng khi từ ngày và đến ngày là cùng 1 ngày');
    if (from.getDay() === 0 || from.getDay() === 6) return fail_('Không thể nghỉ nửa ngày vào cuối tuần');
    days = 0.5;
  } else {
    days = calcWorkDays(from, to);
    if (days <= 0) return fail_('Không có ngày làm việc (cuối tuần không tính)');
  }

  var id;
  var res = withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var me = findEmployeeByCCCD(emp.cccd, ly);
    var avail = me.annual_leave_days - me.used_leave_days - me.pending_leave_days;
    if (days > avail) {
      return fail_('Vượt số ngày phép còn lại của năm học ' + range.label + '. Bạn còn ' + Math.max(avail, 0) + ' ngày' +
                   (me.pending_leave_days > 0 ? ' (đã trừ ' + me.pending_leave_days + ' ngày đang chờ duyệt)' : '') + '.');
    }
    if (hasConflict_(emp.cccd, from, to, halfDay)) return fail_('Trùng lịch với đơn đã có trong khoảng thời gian này');

    id = newRequestId_();
    var now = new Date().toISOString();
    var halfLabel = { none: '', morning: ' (buổi sáng)', afternoon: ' (buổi chiều)' }[halfDay];
    var history = [historyEntry_(emp, CONFIG.STATUS.PENDING, 'Tạo đơn xin nghỉ phép' + halfLabel)];

    appendRow_(CONFIG.SHEETS.LEAVE_REQUESTS, [
      id, emp.cccd, emp.name, emp.department, fmtDMY(from), fmtDMY(to), days, reason,
      CONFIG.STATUS.PENDING, '', '', '', '', now, now,
      String(data.attachment || '').trim(), JSON.stringify(history), halfDay, ''
    ]);
    return ok_();
  });
  if (!res.success) return res;

  logActivity('CREATE', emp.cccd, id + ' ' + days + 'd half=' + halfDay);
  try { notifyManagerNewRequest(emp, id, fmtDMY(from), fmtDMY(to), days, reason, halfDay); }
  catch (e) { Logger.log('[Email] ' + e); }

  var msg = halfDay !== 'none'
    ? 'Tạo đơn thành công (' + (halfDay === 'morning' ? 'buổi sáng' : 'buổi chiều') + ' ngày ' + fmtDMY(from) + '), đang chờ phê duyệt.'
    : 'Tạo đơn thành công (' + days + ' ngày làm việc), đang chờ phê duyệt.';
  return ok_({ message: msg, request_id: id, days: days, half_day: halfDay });
}

function getLeaveRequests(token, params) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  params = params || {};

  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  var rows = [];
  for (var i = 1; i < data.length; i++) if (data[i][LR.ID]) rows.push(mapRequest_(data[i]));

  // Phạm vi xem
  var scope   = String(params.scope || '').trim();
  var myCCCD  = emp.cccd;
  var myDepts = emp.departments;
  var mine    = function (r) { return r.employee_id === myCCCD; };

  if (scope === 'mine') {
    rows = rows.filter(mine);
  } else if (scope === 'team') {
    if (isManager(emp.role))  rows = rows.filter(function (r) { return myDepts.indexOf(r.department) !== -1 && !mine(r); });
    else if (isHR(emp.role))  rows = rows.filter(function (r) { return !mine(r); });
    else                      rows = [];
  } else if (isHR(emp.role)) {
    // HR: xem tất cả
  } else if (isManager(emp.role)) {
    rows = rows.filter(function (r) { return myDepts.indexOf(r.department) !== -1; });
  } else {
    rows = rows.filter(mine);
  }

  // Lọc trạng thái
  var st = String(params.status || '').trim();
  if (st && st !== 'all') {
    rows = rows.filter(function (r) {
      if (st === CONFIG.STATUS.RETURN_PENDING) return r.return_info && r.return_info.status === CONFIG.STATUS.RETURN_PENDING;
      return r.status === st || baseStatus(r.status) === st;
    });
  }

  // Lọc năm học (tùy chọn)
  if (params.leave_year && params.leave_year !== 'all') {
    var label = leaveYearRange(parseInt(params.leave_year, 10)).label;
    rows = rows.filter(function (r) { return r.leave_year === label; });
  }

  rows.sort(function (a, b) { return new Date(b.created_at || 0) - new Date(a.created_at || 0); });
  return ok_({ requests: rows });
}

// ============================================================
// PHÊ DUYỆT
// ============================================================

/**
 * Đổi trạng thái đơn, chỉ khi trạng thái hiện tại nằm trong allowedFrom
 * role: 'approver' → ghi cột người duyệt; 'hr' → ghi cột nhân sự
 */
function updateStatus_(emp, reqId, newStatus, note, role, allowedFrom) {
  return withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(reqId);
    if (!f) return fail_('Không tìm thấy đơn: ' + reqId);
    var cur = String(f.data[LR.STATUS] || '');
    if (allowedFrom.indexOf(cur) === -1) return fail_('Đơn đã được xử lý trước đó (trạng thái hiện tại: ' + cur + ')');

    var history = parseJSON_(f.data[LR.HISTORY], []);
    history.push(historyEntry_(emp, newStatus, note, role));

    var cells = {};
    cells[LR.STATUS + 1]  = newStatus;
    cells[LR.UPDATED + 1] = new Date().toISOString();
    cells[LR.HISTORY + 1] = JSON.stringify(history);
    if (role === 'hr') { cells[LR.HR_ID + 1] = emp.cccd;       cells[LR.HR_NOTE + 1] = note || ''; }
    else               { cells[LR.APPROVER_ID + 1] = emp.cccd; cells[LR.APPROVER_NOTE + 1] = note || ''; }
    setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, f.row, cells);
    return ok_({ message: 'OK', request: mapRequest_(f.data) });
  });
}

/** Trưởng ĐV chỉ duyệt đơn trong đơn vị mình và không tự duyệt đơn của mình */
function checkApproveScope_(emp, reqId) {
  var req = getRequestById(reqId);
  if (!req) return 'Không tìm thấy đơn';
  if (isHR(emp.role)) return '';
  if (req.employee_id === emp.cccd) return 'Không thể tự duyệt đơn của chính mình';
  if (emp.departments.indexOf(req.department) === -1) return 'Đơn không thuộc đơn vị bạn quản lý';
  return '';
}

function approveRequest(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!canApprove(emp.role)) return fail_('Không có quyền phê duyệt');
  var err = checkApproveScope_(emp, data.request_id);
  if (err) return fail_(err);

  var r = updateStatus_(emp, data.request_id, CONFIG.STATUS.APPROVED, String(data.note || '').trim() || 'Đã duyệt',
                        'approver', [CONFIG.STATUS.PENDING]);
  if (!r.success) return r;
  logActivity('APPROVE', emp.cccd, data.request_id);
  try {
    var q = r.request;
    notifyHRPendingConfirm(q.employee_name, q.department, q.id, q.from_date, q.to_date, q.days);
  } catch (e) { Logger.log('[Email] approve: ' + e); }
  return ok_({ message: 'Đã đồng ý đơn nghỉ phép' });
}

function rejectRequest(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!canApprove(emp.role)) return fail_('Không có quyền');
  var note = String(data.note || '').trim();
  if (!note) return fail_('Vui lòng nhập lý do từ chối');
  var err = checkApproveScope_(emp, data.request_id);
  if (err) return fail_(err);

  var r = updateStatus_(emp, data.request_id, CONFIG.STATUS.REJECTED, note, 'approver', [CONFIG.STATUS.PENDING]);
  if (!r.success) return r;
  logActivity('REJECT', emp.cccd, data.request_id);
  try {
    var e2 = findEmployeeByCCCD(r.request.employee_id);
    if (e2) notifyEmployeeManagerRejected(e2.email, data.request_id, note);
  } catch (e) { Logger.log('[Email] reject: ' + e); }
  return ok_({ message: 'Đã từ chối đơn nghỉ phép' });
}

function hrConfirm(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền');
  var note = String(data.note || '').trim() || 'Nhân sự đã ghi nhận';

  var r = updateStatus_(emp, data.request_id, CONFIG.STATUS.HR_CONFIRMED, note, 'hr',
                        [CONFIG.STATUS.APPROVED, CONFIG.STATUS.HR_RETURNED]);
  if (!r.success) return r;
  logActivity('HR_CONFIRM', emp.cccd, data.request_id);
  try {
    var e2 = findEmployeeByCCCD(r.request.employee_id);
    if (e2) notifyEmployeeHRResult(e2.email, data.request_id, CONFIG.STATUS.HR_CONFIRMED, note);
  } catch (e) { Logger.log('[Email] hrConfirm: ' + e); }
  return ok_({ message: 'Đã xác nhận ghi nhận đơn' });
}

function hrReturn(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền trả lại');
  var note = String(data.note || '').trim();
  if (!data.request_id) return fail_('Thiếu mã đơn');
  if (!note) return fail_('Vui lòng nhập lý do trả lại');

  var r = updateStatus_(emp, data.request_id, CONFIG.STATUS.HR_RETURNED, note, 'hr', [CONFIG.STATUS.APPROVED]);
  if (!r.success) return r;
  logActivity('HR_RETURN', emp.cccd, data.request_id);
  try {
    var e2 = findEmployeeByCCCD(r.request.employee_id);
    if (e2) notifyEmployeeHRResult(e2.email, data.request_id, CONFIG.STATUS.HR_RETURNED, note);
  } catch (e) { Logger.log('[Email] hrReturn: ' + e); }
  return ok_({ message: 'Đã trả lại đơn' });
}

// ============================================================
// TRẢ PHÉP
// return_type: 'full' | 'partial' | 'half'
// ============================================================

function createReturnRequest(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  var reason = String(data.reason || '').trim();
  if (!data.request_id) return fail_('Thiếu mã đơn gốc');
  if (!data.return_type) return fail_('Vui lòng chọn loại trả phép');
  if (!reason)           return fail_('Vui lòng nhập lý do trả phép');

  var res = withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(data.request_id);
    if (!f) return fail_('Không tìm thấy đơn gốc');
    var orig = mapRequest_(f.data);

    if (orig.employee_id !== emp.cccd) return fail_('Bạn không có quyền trả phép đơn này');
    if (!isUsedStatus(orig.status))   return fail_('Chỉ có thể trả phép với đơn đã được duyệt');
    if (orig.return_info && orig.return_info.status === CONFIG.STATUS.RETURN_PENDING) {
      return fail_('Đơn này đang có yêu cầu trả phép chờ Nhân sự xác nhận');
    }
    var fromDate = toDate(orig.from_date);
    if (fromDate && fromDate <= today_()) return fail_('Chỉ có thể trả phép trước ngày bắt đầu nghỉ (' + orig.from_date + ')');

    var origDays = orig.days, returnDays;
    if (data.return_type === 'full') {
      returnDays = origDays;
    } else if (data.return_type === 'half') {
      if (origDays < 1) return fail_('Đơn nghỉ nửa ngày không thể trả nửa ngày, hãy chọn trả toàn bộ');
      if (['morning', 'afternoon'].indexOf(data.return_session) === -1) return fail_('Vui lòng chọn buổi muốn trả (sáng/chiều)');
      returnDays = 0.5;
    } else if (data.return_type === 'partial') {
      returnDays = parseFloat(data.return_days) || 0;
      if (returnDays <= 0 || returnDays >= origDays) {
        return fail_('Số ngày trả phải lớn hơn 0 và nhỏ hơn tổng ngày đã đăng ký (' + origDays + ' ngày)');
      }
      if ((returnDays * 2) % 1 !== 0) return fail_('Số ngày trả phải là số nguyên hoặc X.5 (nửa ngày)');
    } else {
      return fail_('Loại trả phép không hợp lệ');
    }

    var now = new Date().toISOString();
    var returnInfo = {
      return_type: data.return_type, return_days: returnDays,
      return_session: data.return_type === 'half' ? data.return_session : '',
      reason: reason, created_at: now, created_by: emp.cccd,
      status: CONFIG.STATUS.RETURN_PENDING, hr_note: '', confirmed_at: '', confirmed_by: ''
    };
    var typeLabel = { full: 'toàn bộ', partial: 'một phần (' + returnDays + ' ngày)', half: 'nửa ngày' }[data.return_type];
    var history = orig.history;
    history.push(historyEntry_(emp, CONFIG.STATUS.RETURN_PENDING, 'Yêu cầu trả phép ' + typeLabel + ': ' + reason));

    var cells = {};
    cells[LR.RETURN_INFO + 1] = JSON.stringify(returnInfo);
    cells[LR.HISTORY + 1]     = JSON.stringify(history);
    cells[LR.UPDATED + 1]     = now;
    setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, f.row, cells);
    return ok_({ return_days: returnDays, orig_days: origDays });
  });
  if (!res.success) return res;

  logActivity('CREATE_RETURN', emp.cccd, data.request_id + ' type=' + data.return_type + ' days=' + res.return_days);
  try { notifyHRReturnRequest(emp, data.request_id, res.return_days, data.return_type, data.return_session, reason); }
  catch (e) { Logger.log('[Email] notifyHRReturn: ' + e); }

  res.message = data.return_type === 'full'
    ? 'Đã gửi yêu cầu trả toàn bộ ' + res.orig_days + ' ngày phép. Chờ Nhân sự xác nhận.'
    : 'Đã gửi yêu cầu trả ' + res.return_days + ' ngày phép. Chờ Nhân sự xác nhận.';
  return res;
}

function hrConfirmReturn(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền xác nhận trả phép');
  if (!data.request_id) return fail_('Thiếu mã đơn');
  var note = String(data.note || '').trim();

  var res = withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(data.request_id);
    if (!f) return fail_('Không tìm thấy đơn: ' + data.request_id);
    var req = mapRequest_(f.data);
    var ri  = req.return_info;
    if (!ri) return fail_('Không tìm thấy yêu cầu trả phép');
    if (ri.status !== CONFIG.STATUS.RETURN_PENDING) return fail_('Yêu cầu trả phép này đã được xử lý rồi');

    var now = new Date().toISOString();
    var returnDays = parseFloat(ri.return_days) || 0;
    var isFull = ri.return_type === 'full';
    var newDays = isFull ? req.days : roundHalf_(req.days - returnDays);

    ri.status = CONFIG.STATUS.RETURN_CONFIRMED;
    ri.hr_note = note; ri.confirmed_at = now; ri.confirmed_by = emp.cccd;

    var sessionLabel = ri.return_session === 'morning' ? ' (buổi sáng)' : ri.return_session === 'afternoon' ? ' (buổi chiều)' : '';
    var history = req.history;
    history.push(historyEntry_(emp, CONFIG.STATUS.RETURN_CONFIRMED,
      'Nhân sự xác nhận trả ' + returnDays + ' ngày phép' + sessionLabel + (note ? '. ' + note : ''), 'hr'));

    var cells = {};
    cells[LR.RETURN_INFO + 1] = JSON.stringify(ri);
    cells[LR.HISTORY + 1]     = JSON.stringify(history);
    cells[LR.UPDATED + 1]     = now;
    if (isFull) {
      cells[LR.STATUS + 1] = CONFIG.STATUS.RETURNED;
    } else {
      cells[LR.DAYS + 1]   = newDays;
      cells[LR.STATUS + 1] = baseStatus(req.status) + CONFIG.PARTIAL_SUFFIX;
    }
    setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, f.row, cells);
    return ok_({ return_days: returnDays, new_days: isFull ? 0 : newDays, employee_id: req.employee_id });
  });
  if (!res.success) return res;

  logActivity('CONFIRM_RETURN', emp.cccd, data.request_id + ' returned=' + res.return_days + 'd');
  try {
    var e2 = findEmployeeByCCCD(res.employee_id);
    if (e2) notifyEmployeeReturnResult(e2.email, data.request_id, true, res.return_days, note);
  } catch (e) { Logger.log('[Email] confirmReturn: ' + e); }

  res.message = 'Đã xác nhận trả ' + res.return_days + ' ngày phép. Phép đã được hoàn lại.';
  return res;
}

function hrRejectReturn(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền');
  var note = String(data.note || '').trim();
  if (!data.request_id) return fail_('Thiếu mã đơn');
  if (!note) return fail_('Vui lòng nhập lý do từ chối');

  var res = withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(data.request_id);
    if (!f) return fail_('Không tìm thấy đơn: ' + data.request_id);
    var req = mapRequest_(f.data);
    var ri  = req.return_info;
    if (!ri) return fail_('Không tìm thấy yêu cầu trả phép');
    if (ri.status !== CONFIG.STATUS.RETURN_PENDING) return fail_('Yêu cầu trả phép này đã được xử lý rồi');

    var now = new Date().toISOString();
    ri.status = CONFIG.STATUS.RETURN_REJECTED;
    ri.hr_note = note; ri.confirmed_at = now; ri.confirmed_by = emp.cccd;
    var history = req.history;
    history.push(historyEntry_(emp, CONFIG.STATUS.RETURN_REJECTED, 'Nhân sự từ chối trả phép: ' + note, 'hr'));

    var cells = {};
    cells[LR.RETURN_INFO + 1] = JSON.stringify(ri);
    cells[LR.HISTORY + 1]     = JSON.stringify(history);
    cells[LR.UPDATED + 1]     = now;
    setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, f.row, cells);
    return ok_({ employee_id: req.employee_id });
  });
  if (!res.success) return res;

  logActivity('REJECT_RETURN', emp.cccd, data.request_id);
  try {
    var e2 = findEmployeeByCCCD(res.employee_id);
    if (e2) notifyEmployeeReturnResult(e2.email, data.request_id, false, 0, note);
  } catch (e) { Logger.log('[Email] rejectReturn: ' + e); }
  return ok_({ message: 'Đã từ chối yêu cầu trả phép' });
}

// ============================================================
// THỐNG KÊ & THÔNG TIN PHÉP
// ============================================================

/** Thông tin phép của 1 nhân viên (Trưởng ĐV / Nhân sự xem khi duyệt).
 *  data.date (tùy chọn): xem theo năm học chứa ngày này, mặc định năm học hiện tại */
function getEmployeeLeaveInfo(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!canApprove(emp.role)) return fail_('Không có quyền');
  if (!data.cccd) return fail_('Thiếu CCCD nhân viên');

  var ly = leaveYearStartOf(toDate(data.date) || today_());
  var target = findEmployeeByCCCD(data.cccd, ly);
  if (!target) return fail_('Không tìm thấy nhân viên');

  var range = leaveYearRange(ly);
  var rows = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  var requests = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][LR.ID] || normalizeCCCD(rows[i][LR.EMP_ID]) !== target.cccd) continue;
    var st = String(rows[i][LR.STATUS] || '');
    if (st === CONFIG.STATUS.REJECTED) continue;
    var from = toDate(rows[i][LR.FROM]);
    if (!from || from < range.from || from > range.to) continue;
    var r = mapRequest_(rows[i]);
    requests.push({ id: r.id, from_date: r.from_date, to_date: r.to_date, days: r.days,
                    status: r.status, reason: r.reason, return_info: r.return_info });
  }

  return ok_({
    cccd: target.cccd, name: target.name, department: target.department,
    annual_days: target.annual_leave_days, used_days: target.used_leave_days,
    pending_days: target.pending_leave_days, remaining: target.remaining_leave_days,
    seniority: target.seniority_years, start_date: target.start_date,
    leave_year: range.label, requests: requests
  });
}

function getStats(token) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!canApprove(emp.role)) return fail_('Không có quyền');

  var data = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
  var range = leaveYearRange(leaveYearStartOf());
  var myDepts = emp.departments;
  var count = { pending: 0, approved: 0, rejected: 0, hr_confirmed: 0, hr_returned: 0, returned: 0, return_pending: 0 };
  var bp = {}, bd = {}, total = 0;

  for (var i = 1; i < data.length; i++) {
    if (!data[i][LR.ID]) continue;
    var r = mapRequest_(data[i]);
    if (!isHR(emp.role) && myDepts.indexOf(r.department) === -1) continue;
    total++;
    var b = baseStatus(r.status);
    if (count.hasOwnProperty(b)) count[b]++;
    if (r.return_info && r.return_info.status === CONFIG.STATUS.RETURN_PENDING) count.return_pending++;

    var from = toDate(r.from_date);
    if (isUsedStatus(r.status) && from && from >= range.from && from <= range.to) {
      bp[r.employee_id] = bp[r.employee_id] || { name: r.employee_name, department: r.department, total_days: 0 };
      bp[r.employee_id].total_days += r.days;
      bd[r.department] = bd[r.department] || { total_days: 0, count: 0 };
      bd[r.department].total_days += r.days;
      bd[r.department].count++;
    }
  }

  return ok_({ stats: {
    leave_year:     range.label,
    by_person:      Object.keys(bp).map(function (k) { return Object.assign({ id: k }, bp[k]); }),
    by_department:  Object.keys(bd).map(function (k) { return Object.assign({ department: k }, bd[k]); }),
    status_count:   count,
    total_requests: total
  }});
}

// ============================================================
// QUẢN LÝ NHÂN VIÊN
// ============================================================

function getEmployees(token) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền');

  var ly = leaveYearStartOf();
  var usage = leaveUsageMap_(ly);
  var em = loadEmployeeMap_();
  var list = em.order.map(function (c) { return withLeave_(em.map[c], usage, ly); });
  return ok_({ employees: list, leave_year: leaveYearRange(ly).label });
}

var VALID_ROLES = ['employee', 'lecturer', 'manager', 'manager_lecturer', 'hr', 'admin'];

/** Nhân sự không được tạo/sửa tài khoản Quản trị; chỉ Quản trị mới làm được */
function guardAdminRole_(actor, newRole, target) {
  if (isAdmin(actor.role)) return '';
  if (newRole === CONFIG.ROLES.ADMIN) return 'Chỉ Quản trị hệ thống mới được gán vai trò Quản trị';
  if (target && isAdmin(target.role)) return 'Chỉ Quản trị hệ thống mới được sửa tài khoản Quản trị';
  return '';
}

function createEmployee(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Không có quyền');

  var rawCCCD = String(data.cccd || '').replace(/\D/g, '');
  var name = String(data.name || '').trim();
  var dept = String(data.department || '').trim();
  if (!rawCCCD || !name || !dept || !data.role) return fail_('Thiếu thông tin bắt buộc');
  if (rawCCCD.length !== 12) return fail_('CCCD phải đủ 12 số (kể cả số 0 ở đầu)');
  if (VALID_ROLES.indexOf(data.role) === -1) return fail_('Vai trò không hợp lệ');
  var start = toDate(data.start_date);
  if (!start) return fail_('Vui lòng nhập ngày vào làm hợp lệ');

  var cccd = normalizeCCCD(rawCCCD);
  var existing = loadEmployeeMap_().map[cccd];
  var g = guardAdminRole_(emp, data.role, existing);
  if (g) return fail_(g);
  if (existing && existing.departments.indexOf(dept) !== -1) return fail_('CCCD đã tồn tại trong đơn vị này');

  appendRow_(CONFIG.SHEETS.EMPLOYEES, [cccd, name, dept, data.role, String(data.email || '').trim(), fmtDMY(start)]);
  var leaveDays = calcAnnualLeaveDays(fmtDMY(start));
  logActivity('CREATE_EMP', emp.cccd, cccd + ' ' + name + ' dept=' + dept + ' role=' + data.role);
  return ok_({ message: 'Thêm nhân viên thành công (' + leaveDays + ' ngày phép/năm học)' });
}

/**
 * Cập nhật tên/vai trò/email/ngày vào làm cho mọi dòng của CCCD; đơn vị chỉ đổi ở dòng original_department.
 * Quản trị có thể đổi số CCCD (new_cccd): mọi đơn nghỉ của người này được chuyển sang CCCD mới.
 */
function updateEmployee(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Không có quyền');

  var cccd = normalizeCCCD(data.cccd);
  if (!cccd) return fail_('Thiếu CCCD');
  if (data.role && VALID_ROLES.indexOf(data.role) === -1) return fail_('Vai trò không hợp lệ');
  var start = data.start_date ? toDate(data.start_date) : null;
  if (data.start_date && !start) return fail_('Ngày vào làm không hợp lệ');

  var empMap = loadEmployeeMap_().map;
  var target = empMap[cccd];
  if (!target) return fail_('Không tìm thấy nhân viên');
  var g = guardAdminRole_(emp, data.role, target);
  if (g) return fail_(g);
  if (target.cccd === emp.cccd && isAdmin(emp.role) && data.role && !isAdmin(data.role)) {
    return fail_('Không thể tự bỏ quyền Quản trị của chính mình. Nhờ một Quản trị khác thực hiện.');
  }

  // Đổi số CCCD (chỉ Quản trị)
  var newCCCD = '';
  var rawNew = String(data.new_cccd || '').replace(/\D/g, '');
  if (rawNew && normalizeCCCD(rawNew) !== cccd) {
    if (!isAdmin(emp.role)) return fail_('Chỉ Quản trị hệ thống mới được đổi số CCCD');
    if (rawNew.length !== 12) return fail_('CCCD mới phải đủ 12 số');
    newCCCD = normalizeCCCD(rawNew);
    if (empMap[newCCCD]) return fail_('CCCD mới đã tồn tại trong hệ thống');
  }

  return withLock_(function () {
    invalidate_(CONFIG.SHEETS.EMPLOYEES);
    var rows = rows_(CONFIG.SHEETS.EMPLOYEES);
    var origDept = String(data.original_department || '').trim();
    var newDept  = String(data.department || '').trim();
    var deptChanged = false;

    for (var i = 1; i < rows.length; i++) {
      if (normalizeCCCD(rows[i][EC.CCCD]) !== cccd) continue;
      var cells = {};
      cells[EC.CCCD + 1] = newCCCD || cccd; // luôn ghi lại dạng văn bản đủ 12 số
      if (data.name)                cells[EC.NAME + 1]  = String(data.name).trim();
      if (data.role)                cells[EC.ROLE + 1]  = data.role;
      if (data.email !== undefined) cells[EC.EMAIL + 1] = String(data.email || '').trim();
      if (start)                    cells[EC.START + 1] = fmtDMY(start);
      var rowDept = String(rows[i][EC.DEPT] || '').trim();
      if (newDept && !deptChanged && (!origDept || rowDept === origDept)) {
        cells[EC.DEPT + 1] = newDept; deptChanged = true;
      }
      setCells_(CONFIG.SHEETS.EMPLOYEES, i + 1, cells);
    }

    var moved = 0;
    if (newCCCD) {
      var lr = rows_(CONFIG.SHEETS.LEAVE_REQUESTS);
      for (var j = 1; j < lr.length; j++) {
        if (normalizeCCCD(lr[j][LR.EMP_ID]) !== cccd) continue;
        var c2 = {}; c2[LR.EMP_ID + 1] = newCCCD;
        if (data.name) c2[LR.EMP_NAME + 1] = String(data.name).trim();
        setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, j + 1, c2);
        moved++;
      }
      var props = PropertiesService.getScriptProperties();
      var pw = props.getProperty('ADMINPW_' + cccd);
      if (pw) { props.setProperty('ADMINPW_' + newCCCD, pw); props.deleteProperty('ADMINPW_' + cccd); }
    }
    logActivity('UPDATE_EMP', emp.cccd, cccd + (newCCCD ? ' → ' + newCCCD + ' (chuyển ' + moved + ' đơn)' : '') +
                (data.role ? ' role=' + data.role : ''));
    return ok_({ message: newCCCD ? 'Đã cập nhật và đổi CCCD (chuyển ' + moved + ' đơn nghỉ sang CCCD mới)' : 'Cập nhật thành công' });
  });
}

/** Xóa nhân viên (chỉ Quản trị). department = tên đơn vị hoặc 'all'. Đơn nghỉ cũ vẫn được giữ lại. */
function deleteEmployee(me, data) {
  var cccd = normalizeCCCD(data.cccd);
  if (!cccd) return fail_('Thiếu CCCD');
  if (cccd === me.cccd) return fail_('Không thể tự xóa tài khoản của chính mình');
  var dept = String(data.department || 'all').trim();

  return withLock_(function () {
    invalidate_(CONFIG.SHEETS.EMPLOYEES);
    var sh = sheet_(CONFIG.SHEETS.EMPLOYEES);
    var rows = rows_(CONFIG.SHEETS.EMPLOYEES);
    var removed = [], total = 0;
    for (var i = 1; i < rows.length; i++) if (normalizeCCCD(rows[i][EC.CCCD]) === cccd) total++;
    if (!total) return fail_('Không tìm thấy nhân viên');

    for (var k = rows.length - 1; k >= 1; k--) {
      if (normalizeCCCD(rows[k][EC.CCCD]) !== cccd) continue;
      if (dept !== 'all' && String(rows[k][EC.DEPT] || '').trim() !== dept) continue;
      removed.push(rows[k].map(String));
      sh.deleteRow(k + 1);
    }
    invalidate_(CONFIG.SHEETS.EMPLOYEES);
    if (!removed.length) return fail_('Nhân viên không thuộc đơn vị "' + dept + '"');
    if (removed.length === total) PropertiesService.getScriptProperties().deleteProperty('ADMINPW_' + cccd);
    // Lưu bản sao dòng đã xóa vào nhật ký để có thể khôi phục thủ công
    logActivity('DELETE_EMP', me.cccd, cccd + ' dept=' + dept + ' data=' + JSON.stringify(removed));
    var name = removed[0][EC.NAME];
    return ok_({ message: removed.length === total ? 'Đã xóa nhân viên ' + name : 'Đã xóa ' + name + ' khỏi đơn vị ' + dept });
  });
}

// ============================================================
// QUẢN TRỊ ĐƠN NGHỈ (chỉ Quản trị): tạo hộ, sửa mọi trường, xóa
// ============================================================

var ALL_REQUEST_STATUSES = ['pending', 'approved', 'rejected', 'hr_confirmed', 'hr_returned', 'returned',
                            'approved_partial_returned', 'hr_confirmed_partial_returned'];

/** Kiểm tra & tính số ngày cho dữ liệu đơn do Quản trị nhập. Trả về {error} hoặc {from,to,halfDay,days} */
function adminNormalizeDates_(data, fallback) {
  fallback = fallback || {};
  var from = toDate(data.from_date || fallback.from_date);
  var to   = toDate(data.to_date   || fallback.to_date);
  if (!from || !to) return { error: 'Ngày không hợp lệ. Dùng định dạng dd/mm/yyyy.' };
  if (from > to)    return { error: 'Ngày bắt đầu sau ngày kết thúc' };
  if (leaveYearStartOf(to) !== leaveYearStartOf(from)) {
    return { error: 'Đơn không được vượt qua ngày 30/06 (hết năm học). Vui lòng tách thành 2 đơn.' };
  }
  var halfDay = data.half_day !== undefined ? data.half_day : (fallback.half_day || 'none');
  if (['none', 'morning', 'afternoon'].indexOf(halfDay) === -1) halfDay = 'none';
  if (halfDay !== 'none' && from.getTime() !== to.getTime()) return { error: 'Nghỉ nửa ngày chỉ áp dụng cho 1 ngày' };

  var days;
  if (data.days !== undefined && data.days !== '' && data.days !== null) {
    days = parseFloat(data.days);
    if (!(days >= 0) || (days * 2) % 1 !== 0) return { error: 'Số ngày phải là số nguyên hoặc X.5' };
  } else {
    days = halfDay !== 'none' ? 0.5 : calcWorkDays(from, to);
  }
  return { from: from, to: to, halfDay: halfDay, days: days };
}

function adminCreateRequest(me, data) {
  var target = findEmployeeByCCCD(data.employee_id);
  if (!target) return fail_('Không tìm thấy nhân viên');
  var dept = String(data.department || target.department).trim();
  if (target.departments.indexOf(dept) === -1) return fail_('Nhân viên không thuộc đơn vị ' + dept);
  var reason = String(data.reason || '').trim();
  if (!reason) return fail_('Vui lòng nhập lý do');
  var status = data.status || CONFIG.STATUS.PENDING;
  if (ALL_REQUEST_STATUSES.indexOf(status) === -1) return fail_('Trạng thái không hợp lệ');
  var d = adminNormalizeDates_(data);
  if (d.error) return fail_(d.error);

  var id;
  var res = withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    if (!data.force && hasConflict_(target.cccd, d.from, d.to, d.halfDay)) {
      return fail_('Trùng lịch với đơn khác của nhân viên này', { conflict: true });
    }
    id = newRequestId_();
    var now = new Date().toISOString();
    var history = [historyEntry_(me, status, 'Quản trị viên tạo đơn hộ cho ' + target.name + (data.admin_note ? ': ' + data.admin_note : ''), 'admin')];
    var approved = isUsedStatus(status) || status === CONFIG.STATUS.REJECTED;
    var hrDone = baseStatus(status) === CONFIG.STATUS.HR_CONFIRMED || status === CONFIG.STATUS.HR_RETURNED;
    appendRow_(CONFIG.SHEETS.LEAVE_REQUESTS, [
      id, target.cccd, target.name, dept, fmtDMY(d.from), fmtDMY(d.to), d.days, reason, status,
      approved ? me.cccd : '', String(data.approver_note || ''), hrDone ? me.cccd : '', String(data.hr_note || ''),
      now, now, String(data.attachment || '').trim(), JSON.stringify(history), d.halfDay, ''
    ]);
    return ok_();
  });
  if (!res.success) return res;
  logActivity('ADMIN_CREATE_REQ', me.cccd, id + ' for=' + target.cccd + ' ' + fmtDMY(d.from) + '→' + fmtDMY(d.to) + ' ' + d.days + 'd status=' + status);
  return ok_({ message: 'Đã tạo đơn ' + id + ' cho ' + target.name + ' (' + d.days + ' ngày)', request_id: id });
}

function adminUpdateRequest(me, data) {
  if (!data.request_id) return fail_('Thiếu mã đơn');
  if (data.status && ALL_REQUEST_STATUSES.indexOf(data.status) === -1) return fail_('Trạng thái không hợp lệ');

  return withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(data.request_id);
    if (!f) return fail_('Không tìm thấy đơn: ' + data.request_id);
    var old = mapRequest_(f.data);

    var d = adminNormalizeDates_(data, old);
    if (d.error) return fail_(d.error);
    var target = findEmployeeByCCCD(old.employee_id);
    var dept = String(data.department || old.department).trim();
    if (target && dept !== old.department && target.departments.indexOf(dept) === -1) {
      return fail_('Nhân viên không thuộc đơn vị ' + dept);
    }
    if (!data.force && hasConflict_(old.employee_id, d.from, d.to, d.halfDay, old.id)) {
      return fail_('Trùng lịch với đơn khác của nhân viên này', { conflict: true });
    }

    var next = {
      department: dept, from_date: fmtDMY(d.from), to_date: fmtDMY(d.to), days: d.days, half_day: d.halfDay,
      reason: data.reason !== undefined ? String(data.reason).trim() : old.reason,
      status: data.status || old.status,
      approver_note: data.approver_note !== undefined ? String(data.approver_note) : old.approver_note,
      hr_note: data.hr_note !== undefined ? String(data.hr_note) : old.hr_note,
      attachment: data.attachment !== undefined ? String(data.attachment).trim() : old.attachment
    };
    if (!next.reason) return fail_('Lý do không được để trống');

    var labels = { department: 'Đơn vị', from_date: 'Từ ngày', to_date: 'Đến ngày', days: 'Số ngày', half_day: 'Buổi',
                   reason: 'Lý do', status: 'Trạng thái', approver_note: 'Ý kiến trưởng ĐV', hr_note: 'Ý kiến NS', attachment: 'Đính kèm' };
    var changes = [];
    Object.keys(labels).forEach(function (k) {
      if (String(next[k]) !== String(old[k])) changes.push(labels[k] + ': ' + (old[k] === '' ? '∅' : old[k]) + ' → ' + (next[k] === '' ? '∅' : next[k]));
    });
    var clearReturn = !!data.clear_return && !!old.return_info;
    if (clearReturn) changes.push('Xóa thông tin trả phép');
    if (!changes.length) return fail_('Không có thay đổi nào');

    var now = new Date().toISOString();
    var history = old.history;
    history.push(historyEntry_(me, next.status, 'Quản trị viên chỉnh sửa — ' + changes.join('; ') +
                 (data.admin_note ? ' | Ghi chú: ' + data.admin_note : ''), 'admin'));

    var cells = {};
    cells[LR.DEPT + 1] = next.department;   cells[LR.FROM + 1] = next.from_date; cells[LR.TO + 1] = next.to_date;
    cells[LR.DAYS + 1] = next.days;         cells[LR.HALF_DAY + 1] = next.half_day;
    cells[LR.REASON + 1] = next.reason;     cells[LR.STATUS + 1] = next.status;
    cells[LR.APPROVER_NOTE + 1] = next.approver_note; cells[LR.HR_NOTE + 1] = next.hr_note;
    cells[LR.ATTACH + 1] = next.attachment; cells[LR.UPDATED + 1] = now;
    cells[LR.HISTORY + 1] = JSON.stringify(history);
    if (clearReturn) cells[LR.RETURN_INFO + 1] = '';
    setCells_(CONFIG.SHEETS.LEAVE_REQUESTS, f.row, cells);

    logActivity('ADMIN_UPDATE_REQ', me.cccd, old.id + ' ' + changes.join('; '));
    return ok_({ message: 'Đã cập nhật đơn ' + old.id + ' (' + changes.length + ' thay đổi)' });
  });
}

function adminDeleteRequest(me, data) {
  if (!data.request_id) return fail_('Thiếu mã đơn');
  return withLock_(function () {
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    var f = findRequestRow_(data.request_id);
    if (!f) return fail_('Không tìm thấy đơn: ' + data.request_id);
    sheet_(CONFIG.SHEETS.LEAVE_REQUESTS).deleteRow(f.row);
    invalidate_(CONFIG.SHEETS.LEAVE_REQUESTS);
    // Lưu bản sao dòng đã xóa vào nhật ký để có thể khôi phục thủ công
    logActivity('ADMIN_DELETE_REQ', me.cccd, data.request_id + (data.reason ? ' lý do=' + data.reason : '') +
                ' data=' + JSON.stringify(f.data.map(function (v) { return v instanceof Date ? fmtDMY(v) : String(v); })));
    return ok_({ message: 'Đã xóa đơn ' + data.request_id });
  });
}

// ============================================================
// NHẬT KÝ & HỆ THỐNG (chỉ Quản trị)
// ============================================================

function getLogs(data) {
  var rows = rows_(CONFIG.SHEETS.LOGS);
  var names = loadEmployeeMap_().map;
  var q = String(data.q || '').trim().toLowerCase();
  var action = String(data.log_action || '').trim();
  var from = data.from ? toDate(data.from) : null;
  var to = data.to ? toDate(data.to) : null;
  if (to) to = new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1);
  var limit = Math.min(parseInt(data.limit, 10) || 300, 2000);
  var actions = {}, out = [];

  for (var i = rows.length - 1; i >= 1; i--) {
    var r = rows[i];
    if (!r[0]) continue;
    var act = String(r[1] || '');
    actions[act] = 1;
    if (out.length >= limit) continue;
    if (action && act !== action) continue;
    var ts = r[0] instanceof Date ? r[0] : new Date(String(r[0]));
    if (from && ts < from) continue;
    if (to && ts >= to) continue;
    var uid = normalizeCCCD(r[2]) || String(r[2] || '');
    var name = names[uid] ? names[uid].name : '';
    var detail = String(r[3] || '');
    if (q && (act + ' ' + uid + ' ' + name + ' ' + detail).toLowerCase().indexOf(q) === -1) continue;
    out.push({ time: isNaN(ts) ? String(r[0]) : ts.toISOString(), action: act, user_id: uid, user_name: name, detail: detail });
  }
  return ok_({ logs: out, actions: Object.keys(actions).sort(), total: Math.max(rows.length - 1, 0) });
}

function getSystemInfo() {
  var em = loadEmployeeMap_();
  var byRole = {}, admins = [];
  em.order.forEach(function (c) {
    var e = em.map[c];
    byRole[e.role] = (byRole[e.role] || 0) + 1;
    if (isAdmin(e.role)) admins.push({ cccd: e.cccd, name: e.name, has_password: !!getAdminPw_(e.cccd) });
  });
  var lr = rows_(CONFIG.SHEETS.LEAVE_REQUESTS), reqCount = 0;
  for (var i = 1; i < lr.length; i++) if (lr[i][LR.ID]) reqCount++;
  var props = PropertiesService.getScriptProperties().getProperties(), now = Date.now(), sessions = 0;
  Object.keys(props).forEach(function (k) {
    if (k.indexOf('TOKEN_') !== 0) return;
    var d = parseJSON_(props[k], null);
    if (d && now <= d.expiry) sessions++;
  });
  var ly = leaveYearRange(leaveYearStartOf());
  return ok_({ info: {
    version: CONFIG.VERSION, timezone: tz_(), leave_year: ly.label,
    spreadsheet_url: ss_().getUrl(), app_url: CONFIG.APP_URL,
    employees: em.order.length, employees_by_role: byRole, requests: reqCount,
    holidays: Math.max(rows_(CONFIG.SHEETS.HOLIDAY_SETTINGS).length - 1, 0),
    logs: Math.max(rows_(CONFIG.SHEETS.LOGS).length - 1, 0),
    active_sessions: sessions, admins: admins
  }});
}

/** Đăng xuất mọi phiên trừ phiên hiện tại */
function logoutAll(me, currentToken) {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties(), n = 0;
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('TOKEN_') === 0 && k !== 'TOKEN_' + currentToken) { props.deleteProperty(k); n++; }
  });
  logActivity('LOGOUT_ALL', me.cccd, n + ' phiên');
  return ok_({ message: 'Đã đăng xuất ' + n + ' phiên đăng nhập khác' });
}

// ============================================================
// KỲ NGHỈ GIẢNG VIÊN (nghỉ hè / nghỉ Tết)
// Cột "year" = năm bắt đầu năm học. VD 2026 = năm học 2026-2027
// ============================================================

/** year: năm học (số) hoặc 'all'. Mặc định năm học hiện tại */
function getHolidaySettings(year) {
  var data = rows_(CONFIG.SHEETS.HOLIDAY_SETTINGS);
  var all = year === 'all';
  var target = all ? null : (parseInt(year, 10) || leaveYearStartOf());
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][HC.YEAR]) continue;
    var y = parseInt(data[i][HC.YEAR], 10);
    if (!all && y !== target) continue;
    result.push({
      id: y + '_' + data[i][HC.TYPE], year: y, year_label: leaveYearRange(y).label,
      type: String(data[i][HC.TYPE]),
      from_date: cellDMY_(data[i][HC.FROM]), to_date: cellDMY_(data[i][HC.TO]),
      note: String(data[i][HC.NOTE] || ''), created_by: normalizeCCCD(data[i][HC.BY]),
      created_at: isoOf_(data[i][HC.AT])
    });
  }
  result.sort(function (a, b) { return (toDate(a.from_date) || 0) - (toDate(b.from_date) || 0); });
  return result;
}

/** Tìm kỳ nghỉ bị trùng với khoảng [from, to] — xét mọi năm học */
function findHolidayOverlap(from, to) {
  var list = getHolidaySettings('all');
  for (var i = 0; i < list.length; i++) {
    var hF = toDate(list[i].from_date), hT = toDate(list[i].to_date);
    if (hF && hT && from <= hT && to >= hF) return list[i];
  }
  return null;
}

function saveHolidaySetting(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền cài đặt kỳ nghỉ');

  var year = parseInt(data.year, 10);
  if (!year || !data.type || !data.from_date || !data.to_date) {
    return fail_('Thiếu thông tin: năm học, loại, ngày bắt đầu, ngày kết thúc');
  }
  if (CONFIG.HOLIDAY_TYPES.indexOf(data.type) === -1) return fail_('Loại kỳ nghỉ không hợp lệ (summer / tet)');
  var from = toDate(data.from_date), to = toDate(data.to_date);
  if (!from || !to) return fail_('Ngày không hợp lệ');
  if (from > to)    return fail_('Ngày bắt đầu sau ngày kết thúc');
  var range = leaveYearRange(year);
  if (from < range.from || to > range.to) {
    return fail_('Kỳ nghỉ phải nằm trong năm học ' + range.label + ' (' + range.from_str + ' – ' + range.to_str + ')');
  }

  ensureSheet_(CONFIG.SHEETS.HOLIDAY_SETTINGS);
  var now = new Date().toISOString();
  var note = String(data.note || '').trim();
  var rows = rows_(CONFIG.SHEETS.HOLIDAY_SETTINGS);

  for (var i = 1; i < rows.length; i++) {
    if (parseInt(rows[i][HC.YEAR], 10) === year && String(rows[i][HC.TYPE]) === data.type) {
      var cells = {};
      cells[HC.FROM + 1] = fmtDMY(from); cells[HC.TO + 1] = fmtDMY(to);
      cells[HC.NOTE + 1] = note; cells[HC.BY + 1] = emp.cccd; cells[HC.AT + 1] = now;
      setCells_(CONFIG.SHEETS.HOLIDAY_SETTINGS, i + 1, cells);
      logActivity('UPDATE_HOLIDAY', emp.cccd, year + '_' + data.type);
      return ok_({ message: 'Cập nhật kỳ nghỉ thành công' });
    }
  }
  appendRow_(CONFIG.SHEETS.HOLIDAY_SETTINGS, [year, data.type, fmtDMY(from), fmtDMY(to), note, emp.cccd, now]);
  logActivity('ADD_HOLIDAY', emp.cccd, year + '_' + data.type);
  return ok_({ message: 'Thêm kỳ nghỉ thành công' });
}

function deleteHolidaySetting(token, data) {
  var emp = verifyToken(token);
  if (!emp) return authFail_();
  if (!isHR(emp.role)) return fail_('Chỉ Nhân sự mới có quyền');
  var year = parseInt(data.year, 10);
  if (!year || !data.type) return fail_('Thiếu thông tin');

  var sh = sheet_(CONFIG.SHEETS.HOLIDAY_SETTINGS);
  if (!sh) return fail_('Sheet không tồn tại');
  var rows = rows_(CONFIG.SHEETS.HOLIDAY_SETTINGS);
  for (var i = rows.length - 1; i >= 1; i--) {
    if (parseInt(rows[i][HC.YEAR], 10) === year && String(rows[i][HC.TYPE]) === String(data.type)) {
      sh.deleteRow(i + 1);
      invalidate_(CONFIG.SHEETS.HOLIDAY_SETTINGS);
      logActivity('DELETE_HOLIDAY', emp.cccd, year + '_' + data.type);
      return ok_({ message: 'Đã xóa kỳ nghỉ' });
    }
  }
  return fail_('Không tìm thấy kỳ nghỉ');
}

// ============================================================
// THIẾT LẬP & CHUẨN HÓA DỮ LIỆU
// Chạy trực tiếp trong trình soạn thảo Apps Script: setupSheets, fixData
// ============================================================

function ensureSheet_(name) {
  var sh = sheet_(name);
  if (!sh) sh = ss_().insertSheet(name);
  var headers = CONFIG.HEADERS[name];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground(CONFIG.HEADER_COLORS[name]).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  (CONFIG.TEXT_COLUMNS[name] || []).forEach(function (c) {
    sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  return sh;
}

/** Tạo sheet còn thiếu + đặt định dạng văn bản. KHÔNG xóa dữ liệu. */
function setupSheets() {
  Object.keys(CONFIG.SHEETS).forEach(function (k) { ensureSheet_(CONFIG.SHEETS[k]); });
  return ok_({ message: 'Đã kiểm tra/tạo đủ các sheet. Version: ' + CONFIG.VERSION });
}

/**
 * Chuẩn hóa dữ liệu cũ:
 *  - CCCD đủ 12 số, lưu dạng văn bản (khôi phục số 0 đầu bị mất)
 *  - Ngày tháng chuyển về văn bản dd/MM/yyyy
 */
function fixData() {
  var plan = {
    EMPLOYEES:        { cccd: [1],          date: [6] },
    LEAVE_REQUESTS:   { cccd: [2, 10, 12],  date: [5, 6] },
    LOGS:             { cccd: [3],          date: [] },
    HOLIDAY_SETTINGS: { cccd: [6],          date: [3, 4] }
  };
  var report = [];
  Object.keys(plan).forEach(function (name) {
    var sh = sheet_(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) { ensureSheet_(name); return; }
    var n = last - 1;

    plan[name].cccd.forEach(function (c) {
      var rg = sh.getRange(2, c, n, 1);
      var vals = rg.getValues().map(function (r) {
        var v = r[0];
        if (v === '' || v === null) return [''];
        var digits = String(v).replace(/\D/g, '');
        return [digits ? normalizeCCCD(digits) : String(v)];
      });
      rg.setNumberFormat('@').setValues(vals);
    });
    plan[name].date.forEach(function (c) {
      var rg = sh.getRange(2, c, n, 1);
      var vals = rg.getValues().map(function (r) { return [r[0] === '' ? '' : cellDMY_(r[0])]; });
      rg.setNumberFormat('@').setValues(vals);
    });
    ensureSheet_(name);
    invalidate_(name);
    report.push(name + ': ' + n + ' dòng');
  });
  Logger.log('fixData: ' + report.join(' | '));
  return ok_({ message: 'Đã chuẩn hóa dữ liệu — ' + report.join(', ') });
}

// ============================================================
// EMAIL
// ============================================================

function escHtml_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function findManagerContacts_(department, excludeCCCD) {
  var data = rows_(CONFIG.SHEETS.EMPLOYEES), list = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][EC.DEPT]).trim() !== department || !isManager(String(data[i][EC.ROLE]).trim())) continue;
    if (normalizeCCCD(data[i][EC.CCCD]) === excludeCCCD || !data[i][EC.EMAIL]) continue;
    list.push({ email: String(data[i][EC.EMAIL]).trim(), name: data[i][EC.NAME] });
  }
  return list;
}

function findHREmails_() {
  var data = rows_(CONFIG.SHEETS.EMPLOYEES), seen = {}, list = [];
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][EC.EMAIL] || '').trim();
    if (String(data[i][EC.ROLE]).trim() === CONFIG.ROLES.HR && email && !seen[email]) { seen[email] = 1; list.push(email); }
  }
  return list;
}

function sendNotification(toEmail, subject, htmlBody) {
  if (!toEmail || !String(toEmail).trim()) return false;
  try {
    GmailApp.sendEmail(toEmail, subject, '', { htmlBody: htmlBody, name: 'HRM · Sổ tay Quản lý nghỉ phép' });
    return true;
  } catch (err) { Logger.log('[Email] ERROR: ' + err); return false; }
}

function buildEmailHTML(title, rows, footer) {
  var rowsHtml = rows.map(function (r) {
    return '<tr><td style="padding:8px 16px;color:#8BA8CC;font-size:13px;width:40%">' + escHtml_(r[0]) +
           '</td><td style="padding:8px 16px;color:#EEF3FA;font-size:13px;font-weight:600">' + escHtml_(r[1]) + '</td></tr>';
  }).join('');
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#0F2044;font-family:Arial,sans-serif">' +
    '<div style="max-width:520px;margin:32px auto;background:#1B3A6B;border-radius:12px;overflow:hidden;border:1px solid rgba(255,255,255,0.1)">' +
    '<div style="height:5px;background:linear-gradient(90deg,#E8A020,#F4B942,#E8A020)"></div>' +
    '<div style="padding:28px 32px 20px">' +
    '<div style="font-size:11px;font-weight:700;letter-spacing:1.5px;color:#E8A020;text-transform:uppercase;margin-bottom:8px">HRM · Sổ tay Quản lý nghỉ phép</div>' +
    '<h2 style="margin:0 0 20px;color:#EEF3FA;font-size:18px">' + escHtml_(title) + '</h2>' +
    '<table style="width:100%;border-collapse:collapse;background:rgba(0,0,0,0.2);border-radius:8px">' + rowsHtml + '</table>' +
    '</div><div style="padding:16px 32px 28px">' +
    '<div style="background:rgba(232,160,32,0.12);border:1px solid rgba(232,160,32,0.3);border-radius:8px;padding:12px 16px;font-size:13px;color:#F4B942">' + footer + '</div>' +
    '</div><div style="padding:14px 32px;border-top:1px solid rgba(255,255,255,0.07);font-size:11px;color:#4A6FA5;text-align:center">' +
    'Đại học Y Dược TP. Hồ Chí Minh · Sổ tay Quản lý nghỉ phép nội bộ</div></div></body></html>';
}

function loginLink_() { return '<a href="' + CONFIG.APP_URL + '" style="color:#F4B942">' + CONFIG.APP_URL + '</a>'; }

function notifyManagerNewRequest(emp, requestId, fromDate, toDate, days, reason, halfDay) {
  var daysText = halfDay && halfDay !== 'none'
    ? '0.5 ngày (' + (halfDay === 'morning' ? 'Buổi sáng' : 'Buổi chiều') + ')'
    : days + ' ngày làm việc';
  findManagerContacts_(emp.department, emp.cccd).forEach(function (m) {
    sendNotification(m.email, '[Cần duyệt] Đơn xin nghỉ phép của ' + emp.name,
      buildEmailHTML('Đơn xin nghỉ phép mới',
        [['VC-NLĐ', emp.name], ['Đơn vị', emp.department], ['Mã đơn', requestId],
         ['Từ ngày', fromDate], ['Đến ngày', toDate], ['Số ngày', daysText], ['Lý do', reason]],
        'Đơn đang chờ bạn phê duyệt. Vui lòng đăng nhập ' + loginLink_() + ' để xem xét.'));
  });
}

function notifyHRPendingConfirm(empName, dept, reqId, fromDate, toDate, days) {
  findHREmails_().forEach(function (email) {
    sendNotification(email, '[Chờ xác nhận] Đơn nghỉ phép của ' + empName,
      buildEmailHTML('Đơn nghỉ phép chờ Nhân sự xác nhận',
        [['VC-NLĐ', empName], ['Đơn vị', dept], ['Mã đơn', reqId],
         ['Từ ngày', fromDate], ['Đến ngày', toDate], ['Số ngày', days + ' ngày']],
        'Trưởng đơn vị đã đồng ý. Vui lòng đăng nhập ' + loginLink_() + ' để xác nhận.'));
  });
}

function notifyEmployeeManagerRejected(empEmail, reqId, note) {
  if (!empEmail) return;
  sendNotification(empEmail, '[Không đồng ý] Đơn nghỉ phép ' + reqId,
    buildEmailHTML('Trưởng đơn vị không đồng ý đơn nghỉ phép',
      [['Mã đơn', reqId], ['Lý do', note || '--']],
      'Liên hệ Trưởng đơn vị để biết thêm thông tin.'));
}

function notifyEmployeeHRResult(empEmail, reqId, status, note) {
  if (!empEmail) return;
  var okk = status === CONFIG.STATUS.HR_CONFIRMED;
  sendNotification(empEmail, (okk ? '[Hoàn tất] ' : '[Trả lại] ') + 'Đơn nghỉ phép ' + reqId,
    buildEmailHTML(okk ? 'Nhân sự đã xác nhận' : 'Nhân sự trả lại đơn',
      [['Mã đơn', reqId], ['Kết quả', okk ? 'Đã ghi nhận' : 'Trả lại'], ['Ghi chú', note || '--']],
      okk ? 'Đơn nghỉ phép của bạn đã hoàn tất.' : 'Liên hệ Nhân sự để biết thêm thông tin.'));
}

function notifyHRReturnRequest(emp, requestId, returnDays, returnType, returnSession, reason) {
  var typeLabel = { full: 'toàn bộ', partial: 'một phần (' + returnDays + ' ngày)', half: 'nửa ngày' };
  var sessionLabel = returnSession === 'morning' ? ' (buổi sáng)' : returnSession === 'afternoon' ? ' (buổi chiều)' : '';
  findHREmails_().forEach(function (email) {
    sendNotification(email, '[Trả phép] ' + emp.name + ' muốn trả phép',
      buildEmailHTML('Yêu cầu trả phép',
        [['Nhân viên', emp.name], ['Đơn vị', emp.department], ['Mã đơn gốc', requestId],
         ['Loại trả', typeLabel[returnType] || returnType], ['Số ngày trả', returnDays + ' ngày' + sessionLabel],
         ['Lý do', reason]],
        'Vui lòng đăng nhập ' + loginLink_() + ' để xác nhận hoặc từ chối yêu cầu trả phép này.'));
  });
}

function notifyEmployeeReturnResult(empEmail, reqId, isConfirmed, returnDays, note) {
  if (!empEmail) return;
  sendNotification(empEmail,
    (isConfirmed ? '[Hoàn phép] ' : '[Từ chối trả phép] ') + 'Đơn ' + reqId,
    buildEmailHTML(isConfirmed ? 'Nhân sự đã xác nhận trả phép' : 'Nhân sự từ chối trả phép',
      [['Mã đơn', reqId], ['Kết quả', isConfirmed ? 'Đã hoàn ' + returnDays + ' ngày phép' : 'Từ chối'], ['Ghi chú', note || '--']],
      isConfirmed ? 'Ngày phép đã được hoàn lại vào tài khoản của bạn.' : 'Liên hệ Nhân sự để biết thêm thông tin.'));
}
