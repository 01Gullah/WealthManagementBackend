/**
 * WM_04_Rules_Manager.gs
 * Apex Wealth - platform rules, products, banners and dropdown choices, stored in the database.
 * Every setting is a dropdown in the Sheet or in admin. setupRules is safe to re-run.
 * Banner: ACTIVE_BANNER (OFF or 1-5) in Platform_Rules picks a message from the Banners tab.
 * Choices tab: editable word lists (FLAG_WORDS, BANNER_COLORS) used by dropdowns.
 */

var RULES_TAB = 'Platform_Rules';
var PRODUCTS_TAB = 'Products';
var BANNERS_TAB = 'Banners';
var CHOICES_TAB = 'Choices';
var BANNER_FALLBACK_COLOR_ = '#fcd34d';

var DEFAULT_RULES_ = [
  ['JURISDICTION', 'US', 'Target jurisdiction.'],
  ['ACTIVE_BANNER', 'OFF', 'OFF hides the top banner. 1 to 5 shows that message from the Banners tab.'],
  ['RISK_DISCLOSURE', 'Projections are hypothetical illustrations, not guarantees. All investing involves risk, including loss of principal. Past performance does not predict future results. This is not financial advice.', 'Typed text. Shown near the calculator.'],
  ['MIN_DEPOSIT', '500', 'USD'],
  ['MAX_DEPOSIT', '100000', 'USD'],
  ['MIN_TERM_MONTHS', '3', ''],
  ['MAX_TERM_MONTHS', '36', ''],
  ['MIN_AGE', '18', 'Years.'],
  ['ACCREDITED_ONLY', 'FALSE', 'TRUE if counsel restricts the offering to accredited investors.'],
  ['KYC_REQUIRED', 'TRUE', 'Identity verification before accepting payments (not built yet).'],
  ['SHOW_PROJECTED_RETURNS', 'TRUE', 'FALSE hides calculator results if counsel advises.'],
  ['MAX_LOGIN_ATTEMPTS', '5', 'Wrong passwords allowed before a temporary lockout.'],
  ['LOCKOUT_MINUTES', '15', 'How long a locked account must wait.'],
  ['RESET_ENABLED', 'TRUE', 'FALSE turns off the forgot-password feature.'],
  ['RESET_CODE_MINUTES', '15', 'How long an emailed reset code stays valid.'],
  ['FOOTER_TEXT', 'Digital assets carry risk, including loss of principal. Nothing on this site is financial advice.', 'Typed text. Shown in the footer. Leave blank to hide.']
];

var DEFAULT_PRODUCTS_ = [
  ['conservative', 'Conservative', 8, 'Low', 'TRUE', 'Illustrative rate.'],
  ['balanced', 'Balanced', 18, 'Medium', 'TRUE', 'Illustrative rate.'],
  ['aggressive', 'Aggressive', 35, 'High', 'TRUE', 'Illustrative rate.']
];

var DEFAULT_BANNERS_ = [
  ['1', 'Payments not open', 'Payments are not open yet. You can register and submit a plan request.', 'Amber'],
  ['2', 'Under maintenance', 'We are performing scheduled maintenance. Some features may be unavailable. Please check back soon.', 'Amber'],
  ['3', 'Welcome', 'Welcome to Apex Wealth. New features are being added regularly.', 'Mint'],
  ['4', 'Important notice', 'Important notice: please read the consent terms carefully before investing.', 'Blue'],
  ['5', 'Custom', 'Custom message: edit this text in the Banners tab.', 'Gray']
];

var DEFAULT_CHOICES_ = [
  ['FLAG_WORDS', 'Bonus', 'TRUE', ''],
  ['FLAG_WORDS', 'Gift', 'TRUE', ''],
  ['FLAG_WORDS', 'Reward', 'TRUE', ''],
  ['FLAG_WORDS', 'Welcome gift', 'TRUE', ''],
  ['BANNER_COLORS', 'Mint', 'TRUE', '#34d399'],
  ['BANNER_COLORS', 'Blue', 'TRUE', '#60a5fa'],
  ['BANNER_COLORS', 'Amber', 'TRUE', '#fcd34d'],
  ['BANNER_COLORS', 'Red', 'TRUE', '#fb7185'],
  ['BANNER_COLORS', 'Gray', 'TRUE', '#9ca3af']
];

function range_(from, to, step) {
  var out = [];
  for (var v = from; v <= to; v += step) out.push(String(v));
  return out;
}

var RULE_CHOICES_ = {
  JURISDICTION: ['US'],
  ACTIVE_BANNER: ['OFF', '1', '2', '3', '4', '5'],
  MIN_DEPOSIT: ['100', '250', '500', '1000', '2500', '5000'],
  MAX_DEPOSIT: ['10000', '25000', '50000', '100000', '250000', '500000', '1000000'],
  MIN_TERM_MONTHS: ['1', '3', '6', '12'],
  MAX_TERM_MONTHS: ['12', '24', '36', '48', '60'],
  MIN_AGE: ['18', '21'],
  ACCREDITED_ONLY: ['TRUE', 'FALSE'],
  KYC_REQUIRED: ['TRUE', 'FALSE'],
  SHOW_PROJECTED_RETURNS: ['TRUE', 'FALSE'],
  MAX_LOGIN_ATTEMPTS: ['3', '5', '8', '10'],
  LOCKOUT_MINUTES: ['5', '15', '30', '60'],
  RESET_ENABLED: ['TRUE', 'FALSE'],
  RESET_CODE_MINUTES: ['10', '15', '30', '60']
};

function setupRules() {
  var ss = getRegistrySheet_().getParent();
  ensureRulesTabs_(ss);
  Logger.log('Rules setup OK. Sheet: ' + ss.getUrl());
  return 'Rules setup OK';
}

/* ---------- Choices (editable dropdown word lists) ---------- */
function ensureChoicesTab_(ss) {
  var sh = ss.getSheetByName(CHOICES_TAB);
  if (!sh) {
    sh = ss.insertSheet(CHOICES_TAB);
    sh.appendRow(['List', 'Word', 'Active', 'Value']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(3, 80);
    sh.setColumnWidth(4, 120);
    sh.getRange(2, 1, 80, 4).setNumberFormat('@');
    sh.getRange(2, 1, DEFAULT_CHOICES_.length, 4).setValues(DEFAULT_CHOICES_);
  }
  return sh;
}

/** Words of one list. activeOnly true returns only Active TRUE words. */
function choiceWords_(ss, list, activeOnly) {
  var sh = ensureChoicesTab_(ss);
  var vals = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() !== list) continue;
    var w = String(vals[i][1]).trim();
    if (!w) continue;
    if (activeOnly && String(vals[i][2]).trim().toUpperCase() !== 'TRUE') continue;
    out.push(w);
  }
  return out;
}

/** Value column (for example a color code) of one word in a list, active or not. */
function choiceValue_(ss, list, word) {
  var sh = ensureChoicesTab_(ss);
  var vals = sh.getDataRange().getValues();
  var w = String(word).trim().toLowerCase();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === list && String(vals[i][1]).trim().toLowerCase() === w) {
      return String(vals[i][3] || '').trim();
    }
  }
  return '';
}

/* ---------- Banners ---------- */
function ensureBannersTab_(ss) {
  var sh = ss.getSheetByName(BANNERS_TAB);
  if (!sh) {
    sh = ss.insertSheet(BANNERS_TAB);
    sh.appendRow(['Slot', 'Label', 'Message', 'Color']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 60);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(3, 560);
    sh.setColumnWidth(4, 110);
    sh.getRange(2, 1, 20, 4).setNumberFormat('@');
    sh.getRange(2, 1, DEFAULT_BANNERS_.length, 4).setValues(DEFAULT_BANNERS_);
    sh.getRange(2, 3, 20, 1).setWrap(true);
  } else if (String(sh.getRange(1, 4).getValue()).trim() !== 'Color') {
    sh.getRange(1, 4).setValue('Color').setFontWeight('bold');
    sh.setColumnWidth(4, 110);
    sh.getRange(2, 4, 20, 1).setNumberFormat('@');
    var last = sh.getLastRow();
    for (var r = 2; r <= last; r++) {
      if (String(sh.getRange(r, 1).getValue()).trim() !== '') sh.getRange(r, 4).setValue('Amber');
    }
  }
  return sh;
}

function ensureRulesTabs_(ss) {
  var rules = ss.getSheetByName(RULES_TAB);
  if (!rules) {
    rules = ss.insertSheet(RULES_TAB);
    rules.appendRow(['Key', 'Value', 'Notes']);
    rules.getRange(1, 1, 1, 3).setFontWeight('bold');
    rules.setFrozenRows(1);
    rules.setColumnWidth(1, 210);
    rules.setColumnWidth(2, 380);
    rules.setColumnWidth(3, 380);
  }
  rules.getRange(1, 2, 60, 1).setNumberFormat('@');

  // Remove the old single-banner row (replaced by ACTIVE_BANNER + Banners tab)
  var old = rules.getDataRange().getValues();
  for (var o = old.length - 1; o >= 1; o--) {
    if (String(old[o][0]).trim() === 'BANNER_TEXT') rules.deleteRow(o + 1);
  }

  var existing = {};
  var vals = rules.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) existing[String(vals[i][0]).trim()] = true;
  for (var d = 0; d < DEFAULT_RULES_.length; d++) {
    if (!existing[DEFAULT_RULES_[d][0]]) rules.appendRow(DEFAULT_RULES_[d]);
  }

  var prods = ss.getSheetByName(PRODUCTS_TAB);
  if (!prods) {
    prods = ss.insertSheet(PRODUCTS_TAB);
    prods.appendRow(['ProductId', 'Name', 'APY_Percent', 'RiskLevel', 'Active', 'Description']);
    prods.getRange(1, 1, 1, 6).setFontWeight('bold');
    prods.setFrozenRows(1);
    prods.getRange(2, 5, 30, 1).setNumberFormat('@');
    prods.getRange(2, 1, DEFAULT_PRODUCTS_.length, 6).setValues(DEFAULT_PRODUCTS_);
  }
  prods.getRange(2, 5, 30, 1).setNumberFormat('@');
  prods.getRange(2, 3, 30, 1).setNumberFormat('0');

  ensureChoicesTab_(ss);
  ensureBannersTab_(ss);
  applyDropdowns_(rules, prods);
}

function dropdown_(list) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a value from the list.')
    .build();
}

function applyDropdowns_(rules, prods) {
  var vals = rules.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    var key = String(vals[i][0]).trim();
    if (RULE_CHOICES_[key]) rules.getRange(i + 1, 2).setDataValidation(dropdown_(RULE_CHOICES_[key]));
  }
  prods.getRange(2, 3, 29, 1).setDataValidation(dropdown_(range_(1, 60, 1)));
  prods.getRange(2, 4, 29, 1).setDataValidation(dropdown_(['Low', 'Medium', 'High']));
  prods.getRange(2, 5, 29, 1).setDataValidation(dropdown_(['TRUE', 'FALSE']));
}

function readRules_(ss) {
  var sheet = ss.getSheetByName(RULES_TAB);
  var values = sheet.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0]).trim();
    if (key) out[key] = String(values[i][1]).trim();
  }
  return out;
}

/** Returns {text, color} for the active banner. color is a hex code. */
function getBanner_(ss, active) {
  var out = { text: '', color: BANNER_FALLBACK_COLOR_ };
  var a = String(active || 'OFF').trim().toUpperCase();
  if (a === 'OFF' || a === '') return out;
  var sh = ss.getSheetByName(BANNERS_TAB);
  if (!sh) return out;
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === a) {
      out.text = String(vals[i][2]).trim();
      var word = String(vals[i][3] || '').trim();
      var hex = word ? choiceValue_(ss, 'BANNER_COLORS', word) : '';
      if (/^#[0-9a-fA-F]{6}$/.test(hex)) out.color = hex;
      return out;
    }
  }
  return out;
}

function getBannerText_(ss, active) {
  return getBanner_(ss, active).text;
}

function num_(v, fallback) {
  var n = Number(v);
  return (v !== '' && v !== undefined && isFinite(n)) ? n : fallback;
}

function bool_(v) {
  return String(v).trim().toUpperCase() === 'TRUE';
}

function getPlatformConfig() {
  try {
    var ss = getRegistrySheet_().getParent();
    ensureRulesTabs_(ss);
    var r = readRules_(ss);

    var minDep = num_(r.MIN_DEPOSIT, 500);
    var maxDep = num_(r.MAX_DEPOSIT, 100000);
    var minTerm = num_(r.MIN_TERM_MONTHS, 3);
    var maxTerm = num_(r.MAX_TERM_MONTHS, 36);
    if (maxDep <= minDep) maxDep = minDep * 10;
    if (maxTerm <= minTerm) maxTerm = minTerm + 12;

    var prodVals = ss.getSheetByName(PRODUCTS_TAB).getDataRange().getValues();
    var products = [];
    for (var i = 1; i < prodVals.length; i++) {
      var row = prodVals[i];
      if (!row[0] || !bool_(row[4])) continue;
      var pct = num_(row[2], -1);
      if (pct < 0 || pct > 200) continue;
      products.push({
        id: String(row[0]).trim(),
        name: String(row[1]).trim(),
        apy: pct / 100,
        risk: String(row[3]).trim(),
        description: String(row[5]).trim()
      });
    }

    var banner = getBanner_(ss, r.ACTIVE_BANNER);

    return {
      success: true,
      jurisdiction: r.JURISDICTION || 'US',
      bannerText: banner.text,
      bannerColor: banner.color,
      footerText: r.FOOTER_TEXT === undefined ? '' : r.FOOTER_TEXT,
      riskDisclosure: r.RISK_DISCLOSURE || '',
      minDeposit: minDep,
      maxDeposit: maxDep,
      minTerm: minTerm,
      maxTerm: maxTerm,
      minAge: num_(r.MIN_AGE, 18),
      accreditedOnly: bool_(r.ACCREDITED_ONLY),
      kycRequired: bool_(r.KYC_REQUIRED),
      showProjectedReturns: r.SHOW_PROJECTED_RETURNS === undefined ? true : bool_(r.SHOW_PROJECTED_RETURNS),
      products: products
    };
  } catch (err) {
    Logger.log('getPlatformConfig error: ' + err);
    return { success: false, message: 'Config unavailable.' };
  }
}

function validateAgainstRules_(amount, months) {
  var c = getPlatformConfig();
  if (!c.success) return 'Rules unavailable.';
  if (amount < c.minDeposit) return 'Minimum is ' + c.minDeposit + ' USD.';
  if (amount > c.maxDeposit) return 'Maximum is ' + c.maxDeposit + ' USD.';
  if (months < c.minTerm || months > c.maxTerm) return 'Term must be ' + c.minTerm + ' to ' + c.maxTerm + ' months.';
  return '';
}

function testRules() {
  var c = getPlatformConfig();
  Logger.log(JSON.stringify(c, null, 2));
  return c.success ? 'OK: ' + c.products.length + ' products, banner: ' + (c.bannerText || 'OFF') : 'FAILED';
}