/**
 * WM_06_Consent.gs
 * Apex Wealth - consent text and records, editable from the Sheet.
 * Tabs: Consent_Text (content), Consents (records)
 * Rules in Platform_Rules:
 *   CONSENT_VERSION: bump (v1 -> v2) to force everyone to consent again.
 *   CONSENT_MODE: EVERY_INVESTMENT (default) or ONCE_PER_VERSION.
 */

var CONSENT_TAB = 'Consents';
var CONSENT_TEXT_TAB = 'Consent_Text';
var CONSENT_VERSION_KEY = 'CONSENT_VERSION';
var CONSENT_MODE_KEY = 'CONSENT_MODE';

var DEFAULT_CONSENT_ = [
  [1, 'SECTION', 'Eligibility', 'You must be at least {MIN_AGE} years old and legally able to enter into binding agreements where you live.', 'TRUE'],
  [2, 'SECTION', 'Your responsibility', 'You are solely responsible for your investment decisions, the accuracy of the information you provide, and the security of your account and password.', 'TRUE'],
  [3, 'SECTION', 'Risk', 'All investing involves risk, including loss of some or all of your principal. Projected or illustrative returns are hypothetical, are not guarantees, and past performance does not predict future results.', 'TRUE'],
  [4, 'SECTION', 'No advice', 'Nothing on this platform is financial, investment, tax or legal advice. Consider speaking with a qualified professional before investing.', 'TRUE'],
  [5, 'SECTION', 'Withdrawals', "Withdrawing before a plan's maturity date returns your principal only. Returns are paid only at maturity.", 'TRUE'],
  [6, 'SECTION', 'Funds you use', 'You confirm the funds you use are lawfully yours and not from any unlawful source.', 'TRUE'],
  [7, 'SECTION', 'Terms and Privacy', 'By continuing you agree to the Terms of Service and Privacy Policy.', 'TRUE'],
  [8, 'CHECKBOX', '', 'I am at least {MIN_AGE} years old.', 'TRUE'],
  [9, 'CHECKBOX', '', 'I am fully responsible for my own investment decisions and my account.', 'TRUE'],
  [10, 'CHECKBOX', '', 'I understand I can lose money, and returns are not guaranteed.', 'TRUE'],
  [11, 'CHECKBOX', '', 'I understand early withdrawal returns principal only, and I agree to the Terms and Privacy Policy.', 'TRUE']
];

function dd_(list) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a value from the list.')
    .build();
}

function getConsentSheet_(ss) {
  var sh = ss.getSheetByName(CONSENT_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONSENT_TAB);
    sh.appendRow(['Timestamp', 'Email', 'Version', 'MinAge', 'Note', 'Statements', 'Used', 'UsedFor']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!String(sh.getRange(1, 6).getValue())) {
    sh.getRange(1, 6).setValue('Statements').setFontWeight('bold');
  }
  if (!String(sh.getRange(1, 7).getValue())) {
    sh.getRange(1, 7).setValue('Used').setFontWeight('bold');
    sh.getRange(1, 8).setValue('UsedFor').setFontWeight('bold');
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 7, last - 1, 1).setValue('LEGACY');
  }
  return sh;
}

function getConsentTextSheet_(ss) {
  var sh = ss.getSheetByName(CONSENT_TEXT_TAB);
  if (!sh) {
    sh = ss.insertSheet(CONSENT_TEXT_TAB);
    sh.appendRow(['No', 'Type', 'Title', 'Text', 'Active']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 60);
    sh.setColumnWidth(2, 110);
    sh.setColumnWidth(3, 170);
    sh.setColumnWidth(4, 520);
    sh.setColumnWidth(5, 80);
    sh.getRange(2, 3, 60, 2).setNumberFormat('@');
    sh.getRange(2, 5, 60, 1).setNumberFormat('@');
    sh.getRange(2, 1, DEFAULT_CONSENT_.length, 5).setValues(DEFAULT_CONSENT_);
    sh.getRange(2, 4, 60, 1).setWrap(true);
  }
  sh.getRange(2, 1, 59, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(1, 99).setAllowInvalid(false).build());
  sh.getRange(2, 2, 59, 1).setDataValidation(dd_(['SECTION', 'CHECKBOX']));
  sh.getRange(2, 5, 59, 1).setDataValidation(dd_(['TRUE', 'FALSE']));
  return sh;
}

function ensureRule_(rules, key, value, note, list) {
  var vals = rules.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === key) { row = i + 1; break; }
  }
  if (row < 0) {
    rules.appendRow([key, value, note]);
    row = rules.getLastRow();
  }
  rules.getRange(row, 2).setNumberFormat('@').setDataValidation(dd_(list));
}

function ensureVersionRule_(ss) {
  var rules = ss.getSheetByName('Platform_Rules');
  if (!rules) return;
  var list = [];
  for (var v = 1; v <= 30; v++) list.push('v' + v);
  ensureRule_(rules, CONSENT_VERSION_KEY, 'v1',
    'Change to the next version after editing consent text. Everyone must consent again.', list);
  ensureRule_(rules, CONSENT_MODE_KEY, 'EVERY_INVESTMENT',
    'EVERY_INVESTMENT = consent before each new plan. ONCE_PER_VERSION = once until the version changes.',
    ['EVERY_INVESTMENT', 'ONCE_PER_VERSION']);
}

function setupConsent() {
  var ss = getRegistrySheet_().getParent();
  getConsentSheet_(ss);
  getConsentTextSheet_(ss);
  ensureVersionRule_(ss);
  Logger.log('Consent setup OK');
  return 'Consent setup OK';
}

function getConsentVersion_(ss) {
  try { return readRules_(ss)[CONSENT_VERSION_KEY] || 'v1'; }
  catch (e) { return 'v1'; }
}

function getConsentMode_(ss) {
  try {
    return readRules_(ss)[CONSENT_MODE_KEY] === 'ONCE_PER_VERSION' ? 'ONCE_PER_VERSION' : 'EVERY_INVESTMENT';
  } catch (e) {
    return 'EVERY_INVESTMENT';
  }
}

/** ONCE_PER_VERSION check: any consent row for this email and version. */
function hasConsented_(ss, email) {
  var ver = getConsentVersion_(ss);
  var vals = getConsentSheet_(ss).getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1]).toLowerCase() === email && String(vals[i][2]) === ver) return true;
  }
  return false;
}

/** EVERY_INVESTMENT check: newest unused consent row for this email and version. Returns row number or -1. */
function findFreshConsentRow_(ss, email) {
  var ver = getConsentVersion_(ss);
  var vals = getConsentSheet_(ss).getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][1]).toLowerCase() === email && String(vals[i][2]) === ver &&
        String(vals[i][6]).trim() === '') return i + 1;
  }
  return -1;
}

/** Called when a plan request is submitted. Returns true if consent is valid (and uses it up in EVERY mode). */
function consumeConsent_(ss, email, detail) {
  if (getConsentMode_(ss) === 'ONCE_PER_VERSION') return hasConsented_(ss, email);
  var row = findFreshConsentRow_(ss, email);
  if (row < 0) return false;
  var sh = getConsentSheet_(ss);
  sh.getRange(row, 7).setValue('USED');
  sh.getRange(row, 8).setValue(detail);
  return true;
}

function readConsentItems_(ss, minAge) {
  var vals = getConsentTextSheet_(ss).getDataRange().getValues();
  var sections = [], checks = [];
  for (var i = 1; i < vals.length; i++) {
    var no = Number(vals[i][0]);
    var type = String(vals[i][1]).trim().toUpperCase();
    var title = String(vals[i][2]).trim();
    var text = String(vals[i][3]).trim();
    var active = String(vals[i][4]).trim().toUpperCase() === 'TRUE';
    if (!active || !text || !isFinite(no)) continue;
    text = text.replace(/\{MIN_AGE\}/g, String(minAge));
    title = title.replace(/\{MIN_AGE\}/g, String(minAge));
    var item = { no: no, title: title, text: text };
    if (type === 'SECTION') sections.push(item);
    else if (type === 'CHECKBOX') checks.push(item);
  }
  sections.sort(function (a, b) { return a.no - b.no; });
  checks.sort(function (a, b) { return a.no - b.no; });
  return { sections: sections, checks: checks };
}

function getConsentContent() {
  try {
    var ss = getRegistrySheet_().getParent();
    ensureVersionRule_(ss);
    var cfg = getPlatformConfig();
    var minAge = cfg.success ? cfg.minAge : 18;
    var items = readConsentItems_(ss, minAge);
    if (!items.checks.length) {
      return { success: false, message: 'No active consent checkboxes are set up.' };
    }
    return {
      success: true,
      version: getConsentVersion_(ss),
      sections: items.sections,
      checks: items.checks
    };
  } catch (err) {
    Logger.log('getConsentContent error: ' + err);
    return { success: false, message: 'Could not load the consent. Try again.' };
  }
}

function getConsentStatus(token) {
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    var ss = getRegistrySheet_().getParent();
    if (getConsentMode_(ss) === 'EVERY_INVESTMENT') return { success: true, consented: false };
    return { success: true, consented: hasConsented_(ss, s.email) };
  } catch (err) {
    Logger.log('getConsentStatus error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

function recordConsent(token) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    lock.waitLock(15000);
    locked = true;
    var ss = getRegistrySheet_().getParent();
    var mode = getConsentMode_(ss);
    if (mode === 'EVERY_INVESTMENT' || !hasConsented_(ss, s.email)) {
      var cfg = getPlatformConfig();
      var minAge = cfg.success ? cfg.minAge : 18;
      var items = readConsentItems_(ss, minAge);
      var stmts = items.checks.map(function (c) { return c.no + '. ' + c.text; }).join(' | ');
      getConsentSheet_(ss).appendRow([new Date(), s.email, getConsentVersion_(ss), minAge,
        'All statements ticked and accepted. Mode: ' + mode, stmts, '', '']);
    }
    return { success: true };
  } catch (err) {
    Logger.log('recordConsent error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

