/**
 * WM_03_Auth_Manager.gs
 * Apex Wealth - registration, login (with attempt limits), sessions, password reset.
 * Users_Registry columns: Timestamp | Email | PasswordHash | Balance
 * Secrets: the pepper lives in Script Properties (PEPPER). Never delete it, or every password stops working.
 * Limits come from Platform_Rules: MAX_LOGIN_ATTEMPTS, LOCKOUT_MINUTES, RESET_ENABLED, RESET_CODE_MINUTES.
 */

var AUTH_DB_NAME = 'ApexWealth_Database';
var AUTH_TAB = 'Users_Registry';
var AUTH_HEADERS = ['Timestamp', 'Email', 'PasswordHash', 'Balance'];
var SESSION_DAYS = 30;
var SESSION_PREFIX = 'S_';
var RESET_PREFIX = 'R_';
var HASH_PREFIX = 'v2$';
// Migration only: lets old accounts sign in once so their hash can be upgraded. Safe to delete later.
var LEGACY_PEPPER_ = 'apex-wealth-v2-pepper';

var COMMON_PASSWORDS_ = [
  'password', 'password1', 'password123', '123456789012', 'qwertyuiop12',
  'iloveyou1234', 'admin1234567', 'welcome12345', 'letmein12345', 'changeme1234'
];

/* ---------- Setup (run once from the editor) ---------- */
function setupSecurity() {
  getPepper_();
  var sheet = getRegistrySheet_();
  setupRules();
  var quota = MailApp.getRemainingDailyQuota();
  Logger.log('Security setup OK. Emails left today: ' + quota + '. Sheet: ' + sheet.getParent().getUrl());
  return 'Security setup OK';
}

/** Run to confirm reset emails work. Sends a test to your own Google account. */
function testEmail() {
  var to = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: 'Apex Wealth test email',
    body: 'If you can read this, password reset emails will work.',
    name: 'Apex Wealth'
  });
  Logger.log('Test email sent to ' + to);
  return 'Test email sent to ' + to;
}

function setupAuth() { return setupSecurity(); }

/* ---------- Secrets and hashing ---------- */
function getPepper_() {
  var props = PropertiesService.getScriptProperties();
  var p = props.getProperty('PEPPER');
  if (!p) {
    p = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('PEPPER', p);
  }
  return p;
}

function toHex_(bytes) {
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function sha_(text) {
  return toHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8));
}

function hashPassword_(email, password) {
  return HASH_PREFIX + sha_(email + '|' + password + '|' + getPepper_());
}

function legacyHash_(email, password) {
  return sha_(email + '|' + password + '|' + LEGACY_PEPPER_);
}

function hashToken_(token) {
  return sha_(token + '|' + getPepper_());
}

/** Returns true if the password matches. Sets user.upgrade when an old-format hash matched. */
function verifyPassword_(user, email, password) {
  if (String(user.hash).indexOf(HASH_PREFIX) === 0) {
    return user.hash === hashPassword_(email, password);
  }
  if (user.hash === legacyHash_(email, password)) {
    user.upgrade = true;
    return true;
  }
  return false;
}

/* ---------- Sheet access ---------- */
function getRegistrySheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DB_ID');
  var ss = null;

  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    var files = DriveApp.getFilesByName(AUTH_DB_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(AUTH_DB_NAME);
    }
    props.setProperty('DB_ID', ss.getId());
  }

  var sheet = ss.getSheetByName(AUTH_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(AUTH_TAB);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(AUTH_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  return sheet;
}

/* ---------- Helpers ---------- */
function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) && email.length <= 120;
}

function checkPasswordStrength_(password, email) {
  var p = String(password || '');
  if (p.length < 12) return 'Password must be at least 12 characters.';
  if (p.length > 64) return 'Password must be 64 characters or fewer.';
  if (/\s/.test(p)) return 'Password must not contain spaces.';
  if (!/[a-z]/.test(p)) return 'Password needs a lowercase letter.';
  if (!/[A-Z]/.test(p)) return 'Password needs an uppercase letter.';
  if (!/[0-9]/.test(p)) return 'Password needs a number.';
  if (!/[^A-Za-z0-9]/.test(p)) return 'Password needs a symbol (e.g. ! @ # $ %).';
  if (COMMON_PASSWORDS_.indexOf(p.toLowerCase()) !== -1) return 'That password is too common.';
  var name = String(email || '').split('@')[0];
  if (name.length >= 4 && p.toLowerCase().indexOf(name) !== -1) {
    return 'Password must not contain your email name.';
  }
  return '';
}

function findUserRow_(sheet, email) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim().toLowerCase() === email) {
      return { row: i + 1, hash: String(values[i][2]), balance: Number(values[i][3]) || 0 };
    }
  }
  return null;
}

function emailId_(email) {
  return hashToken_('email:' + email).substring(0, 32);
}

function getSecurityRules_() {
  var out = { maxAttempts: 5, lockMinutes: 15, resetEnabled: true, resetMinutes: 15 };
  try {
    var r = readRules_(getRegistrySheet_().getParent());
    out.maxAttempts = Math.max(1, Math.min(20, num_(r.MAX_LOGIN_ATTEMPTS, 5)));
    out.lockMinutes = Math.max(1, Math.min(300, num_(r.LOCKOUT_MINUTES, 15)));
    out.resetEnabled = r.RESET_ENABLED === undefined ? true : bool_(r.RESET_ENABLED);
    out.resetMinutes = Math.max(5, Math.min(120, num_(r.RESET_CODE_MINUTES, 15)));
  } catch (e) {}
  return out;
}

/* ---------- Sessions ---------- */
function createSession_(email) {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();

  var all = props.getProperties();
  for (var k in all) {
    if (k.indexOf(SESSION_PREFIX) === 0) {
      try {
        var s = JSON.parse(all[k]);
        if (!s.exp || s.exp < now) props.deleteProperty(k);
      } catch (e) {
        props.deleteProperty(k);
      }
    }
  }

  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var exp = now + SESSION_DAYS * 24 * 60 * 60 * 1000;
  props.setProperty(SESSION_PREFIX + hashToken_(token), JSON.stringify({ email: email, exp: exp }));
  return token;
}

function invalidateSessions_(email) {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  for (var k in all) {
    if (k.indexOf(SESSION_PREFIX) === 0) {
      try {
        if (JSON.parse(all[k]).email === email) props.deleteProperty(k);
      } catch (e) {
        props.deleteProperty(k);
      }
    }
  }
}

function restoreSession(token) {
  try {
    token = String(token || '');
    if (token.length < 32) return { success: false, message: 'No session.' };

    var props = PropertiesService.getScriptProperties();
    var key = SESSION_PREFIX + hashToken_(token);
    var raw = props.getProperty(key);
    if (!raw) return { success: false, message: 'Session expired.' };

    var s = JSON.parse(raw);
    if (!s.exp || s.exp < Date.now()) {
      props.deleteProperty(key);
      return { success: false, message: 'Session expired.' };
    }

    var sheet = getRegistrySheet_();
    var user = findUserRow_(sheet, s.email);
    if (!user) {
      props.deleteProperty(key);
      return { success: false, message: 'Account not found.' };
    }
    return { success: true, email: s.email, balance: user.balance };
  } catch (err) {
    Logger.log('restoreSession error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

function logoutUser(token) {
  try {
    token = String(token || '');
    if (token.length >= 32) {
      PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + hashToken_(token));
    }
    return { success: true };
  } catch (err) {
    Logger.log('logoutUser error: ' + err);
    return { success: true };
  }
}

/* ---------- Registration ---------- */
function registerUser(data) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(15000);
    locked = true;

    var email = normalizeEmail_(data && data.email);
    var password = String((data && data.password) || '');

    if (!isValidEmail_(email)) {
      return { success: false, message: 'Enter a valid email address.' };
    }
    var weak = checkPasswordStrength_(password, email);
    if (weak) {
      return { success: false, message: weak };
    }

    var sheet = getRegistrySheet_();
    if (findUserRow_(sheet, email)) {
      return { success: false, message: 'An account with this email already exists.' };
    }

    sheet.appendRow([new Date(), email, hashPassword_(email, password), 0]);

    return {
      success: true,
      message: 'Account created successfully.',
      email: email,
      balance: 0,
      token: createSession_(email)
    };
  } catch (err) {
    Logger.log('registerUser error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ---------- Login with attempt limits ---------- */
function clearLoginLimits_(id) {
  var cache = CacheService.getScriptCache();
  cache.remove('LK_' + id);
  cache.remove('LF_' + id);
}

/** Records a failed attempt. Returns attempts left (0 means now locked). */
function recordLoginFailure_(id, rules) {
  var cache = CacheService.getScriptCache();
  var count = Number(cache.get('LF_' + id) || 0) + 1;
  if (count >= rules.maxAttempts) {
    cache.put('LK_' + id, String(Date.now() + rules.lockMinutes * 60000), Math.min(21600, rules.lockMinutes * 60));
    cache.remove('LF_' + id);
    return 0;
  }
  cache.put('LF_' + id, String(count), Math.min(21600, Math.max(rules.lockMinutes, 15) * 60));
  return rules.maxAttempts - count;
}

function loginUser(data) {
  try {
    var email = normalizeEmail_(data && data.email);
    var password = String((data && data.password) || '');

    if (!isValidEmail_(email) || !password) {
      return { success: false, message: 'Invalid email or password.' };
    }

    var rules = getSecurityRules_();
    var id = emailId_(email);
    var cache = CacheService.getScriptCache();

    var until = Number(cache.get('LK_' + id) || 0);
    if (until > Date.now()) {
      var mins = Math.ceil((until - Date.now()) / 60000);
      return { success: false, message: 'Too many failed attempts. Try again in ' + mins + ' minute' + (mins === 1 ? '' : 's') + '.' };
    }

    var sheet = getRegistrySheet_();
    var user = findUserRow_(sheet, email);
    var ok = user ? verifyPassword_(user, email, password) : false;

    if (!ok) {
      var left = recordLoginFailure_(id, rules);
      if (left === 0) {
        return { success: false, message: 'Too many failed attempts. Try again in ' + rules.lockMinutes + ' minutes.' };
      }
      return { success: false, message: 'Invalid email or password. ' + left + ' attempt' + (left === 1 ? '' : 's') + ' left.' };
    }

    clearLoginLimits_(id);

    if (user.upgrade) {
      sheet.getRange(user.row, 3).setValue(hashPassword_(email, password));
    }

    return {
      success: true,
      message: 'Welcome back.',
      email: email,
      balance: user.balance,
      token: createSession_(email)
    };
  } catch (err) {
    Logger.log('loginUser error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

/* ---------- Forgot password ---------- */
function requestPasswordReset(data) {
  try {
    var email = normalizeEmail_(data && data.email);
    if (!isValidEmail_(email)) return { success: false, message: 'Enter a valid email address.' };

    var rules = getSecurityRules_();
    if (!rules.resetEnabled) {
      return { success: false, message: 'Password reset is currently unavailable. Please contact support.' };
    }

    var id = emailId_(email);
    var cache = CacheService.getScriptCache();
    var reqCount = Number(cache.get('RQ_' + id) || 0);
    if (reqCount >= 3) {
      return { success: false, message: 'Too many reset requests. Please try again in an hour.' };
    }
    cache.put('RQ_' + id, String(reqCount + 1), 3600);

    var generic = { success: true, message: 'If an account exists for that email, a 6-digit code has been sent.' };

    var user = findUserRow_(getRegistrySheet_(), email);
    if (!user) return generic;

    var props = PropertiesService.getScriptProperties();
    var now = Date.now();
    var all = props.getProperties();
    for (var k in all) {
      if (k.indexOf(RESET_PREFIX) === 0) {
        try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); }
        catch (e) { props.deleteProperty(k); }
      }
    }

    var code = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').substring(0, 8), 16) % 1000000)).slice(-6);
    props.setProperty(RESET_PREFIX + id, JSON.stringify({
      h: hashToken_(code + '|' + email),
      exp: now + rules.resetMinutes * 60000,
      tries: 0
    }));

    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Your Apex Wealth password reset code',
        body: 'Your Apex Wealth password reset code is: ' + code + '\n\n' +
              'It expires in ' + rules.resetMinutes + ' minutes and can be used once.\n' +
              'If you did not request this, ignore this email. Your password has not changed.\n' +
              'Never share this code with anyone.',
        name: 'Apex Wealth'
      });
    } catch (mailErr) {
      Logger.log('Reset email failed: ' + mailErr);
    }
    return generic;
  } catch (err) {
    Logger.log('requestPasswordReset error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

function resetPassword(data) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var email = normalizeEmail_(data && data.email);
    var code = String((data && data.code) || '').replace(/\s/g, '');
    var password = String((data && data.password) || '');
    var bad = { success: false, message: 'Invalid or expired code.' };

    if (!isValidEmail_(email) || !/^\d{6}$/.test(code)) return bad;
    var weak = checkPasswordStrength_(password, email);
    if (weak) return { success: false, message: weak };

    lock.waitLock(15000);
    locked = true;

    var id = emailId_(email);
    var props = PropertiesService.getScriptProperties();
    var key = RESET_PREFIX + id;
    var raw = props.getProperty(key);
    if (!raw) return bad;

    var r = JSON.parse(raw);
    if (!r.exp || r.exp < Date.now() || r.tries >= 5) {
      props.deleteProperty(key);
      return bad;
    }
    if (hashToken_(code + '|' + email) !== r.h) {
      r.tries = (r.tries || 0) + 1;
      if (r.tries >= 5) props.deleteProperty(key);
      else props.setProperty(key, JSON.stringify(r));
      return bad;
    }

    var sheet = getRegistrySheet_();
    var user = findUserRow_(sheet, email);
    if (!user) { props.deleteProperty(key); return bad; }

    sheet.getRange(user.row, 3).setValue(hashPassword_(email, password));
    props.deleteProperty(key);
    invalidateSessions_(email);
    clearLoginLimits_(id);

    return { success: true, message: 'Password updated. You can now sign in.' };
  } catch (err) {
    Logger.log('resetPassword error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

