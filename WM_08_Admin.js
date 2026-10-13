/**
 * WM_08_Admin.gs
 * Apex Wealth - admin-only functions. Authorization is checked on the SERVER every call.
 * Admin list: edit ADMIN_LIST_ below, save, run setupAdmin. Log tab: Admin_Log
 * Content editor: adminGetContent / adminSaveRow / adminAddRow for Rules, Products, Banners, Choices, Consent, Legal.
 * Audit viewer: adminGetAudit. Gift credit (earns returns): adminGiftCredit.
 * Withdrawal decisions: adminSetWithdrawal (ACCEPTED closes the plan; no money moves).
 * Payment test events: adminSimulatePayment (SANDBOX mode only).
 */

var ADMIN_LOG_TAB = 'Admin_Log';

var ADMIN_LIST_ = [
  'bashirmohammednajib@gmail.com',
  'gullahlukmanaminu@gmail.com'
];

var TF_ = ['TRUE', 'FALSE'];

var EXTRA_RULE_CHOICES_ = {
  CONSENT_VERSION: (function () { var a = []; for (var v = 1; v <= 30; v++) a.push('v' + v); return a; })(),
  CONSENT_MODE: ['EVERY_INVESTMENT', 'ONCE_PER_VERSION'],
  DASHBOARD_ACCESS: ['TEST_AFTER_REQUEST', 'AFTER_PAYMENT_CONFIRMED'],
  WITHDRAW_ENABLED: TF_,
  WITHDRAW_FORFEIT_RETURNS: TF_,
  PAYMENT_MODE: ['SANDBOX', 'LIVE'],
  ORDER_EXPIRY_MINUTES: ['15', '30', '60', '120', '1440']
};

var CHOICE_LISTS_ = ['FLAG_WORDS', 'BANNER_COLORS', 'GIFT_AMOUNTS', 'GIFT_TERMS', 'WITHDRAW_FEES', 'PAY_METHODS'];

var ADMIN_TABS_ = {
  RULES: {
    sheet: 'Platform_Rules', label: 'Rules', addable: false,
    note: 'Change a value and tap Save. Live after the next page refresh. Dropdown settings only accept valid choices. PAYMENT_MODE LIVE stays closed until a verified payment provider is connected in code.',
    cols: [{ n: 'Setting', k: 'text', ro: true }, { n: 'Value', k: 'text', req: true }, { n: 'Notes', k: 'text', ro: true }]
  },
  PRODUCTS: {
    sheet: 'Products', label: 'Products', addable: true,
    note: 'Rates are illustrative whole percents. Set Active to FALSE to hide a plan. Product ID cannot change after creation.',
    cols: [
      { n: 'ProductId', k: 'id', roExisting: true },
      { n: 'Name', k: 'text', req: true },
      { n: 'APY_Percent', k: 'number', min: 1, max: 60 },
      { n: 'RiskLevel', k: 'select', opts: ['Low', 'Medium', 'High'] },
      { n: 'Active', k: 'select', opts: TF_ },
      { n: 'Description', k: 'text' }
    ]
  },
  BANNERS: {
    sheet: 'Banners', label: 'Banners', addable: false,
    note: 'Edit the 5 messages and pick a color for each. Colors come from the Choices tab. Choose which banner shows with ACTIVE_BANNER in the Rules tab (OFF hides it).',
    cols: [
      { n: 'Slot', k: 'text', ro: true },
      { n: 'Label', k: 'text', req: true },
      { n: 'Message', k: 'area', req: true },
      { n: 'Color', k: 'select', opts: ['Amber'] }
    ]
  },
  CHOICES: {
    sheet: 'Choices', label: 'Choices', addable: true,
    note: 'These are the words inside dropdowns. FLAG_WORDS = labels on investments and gifts. GIFT_AMOUNTS = amount steps for gifts (whole dollars). GIFT_TERMS = term steps for gifts (whole months). WITHDRAW_FEES = early withdrawal fee percents (0 to 50). PAY_METHODS = payment methods customers can pick (for example USDT on TRON). BANNER_COLORS = banner colors; the color code is # followed by 6 letters or numbers, like #34d399. Set Active to FALSE to hide a word. Keep at least one Active word in each list.',
    cols: [
      { n: 'List', k: 'select', opts: CHOICE_LISTS_, roExisting: true },
      { n: 'Word', k: 'text', req: true },
      { n: 'Active', k: 'select', opts: TF_ },
      { n: 'Color code', k: 'text' }
    ]
  },
  CONSENT: {
    sheet: 'Consent_Text', label: 'Consent', addable: true,
    note: 'SECTION = paragraph in the scroll box. CHECKBOX = tick box (keep at least one Active). Use {MIN_AGE} for the age rule. After changing wording, raise CONSENT_VERSION in Rules so everyone must consent again.',
    cols: [
      { n: 'No', k: 'number', min: 1, max: 99 },
      { n: 'Type', k: 'select', opts: ['SECTION', 'CHECKBOX'] },
      { n: 'Title', k: 'text' },
      { n: 'Text', k: 'area', req: true },
      { n: 'Active', k: 'select', opts: TF_ }
    ]
  },
  LEGAL: {
    sheet: 'Legal_Pages', label: 'Legal', addable: true,
    note: 'Terms of Service and Privacy Policy. HEADING = bold title, PARAGRAPH = text. No sets the order within each page. Starter text must be replaced by your attorney before launch.',
    cols: [
      { n: 'Page', k: 'select', opts: ['TERMS', 'PRIVACY'] },
      { n: 'No', k: 'number', min: 1, max: 99 },
      { n: 'Type', k: 'select', opts: ['HEADING', 'PARAGRAPH'] },
      { n: 'Text', k: 'area', req: true },
      { n: 'Active', k: 'select', opts: TF_ }
    ]
  }
};

function getAdminLogSheet_(ss) {
  var sh = ss.getSheetByName(ADMIN_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(ADMIN_LOG_TAB);
    sh.appendRow(['Timestamp', 'Admin', 'Action', 'Target', 'Detail']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Run after any change to ADMIN_LIST_. Overwrites the stored admin list. */
function setupAdmin() {
  var list = ADMIN_LIST_.map(function (e) { return String(e).trim().toLowerCase(); });
  PropertiesService.getScriptProperties().setProperty('ADMIN_EMAILS', list.join(','));
  getAdminLogSheet_(getRegistrySheet_().getParent());
  Logger.log('Admins: ' + list.join(', '));
  return 'Admin setup OK: ' + list.join(', ');
}

function getAdmins_() {
  var raw = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS') || '';
  return raw.split(',').map(function (e) { return e.trim().toLowerCase(); }).filter(function (e) { return e; });
}

function requireAdmin_(token) {
  var s = restoreSession(token);
  if (!s.success) return null;
  return getAdmins_().indexOf(s.email) >= 0 ? s : null;
}

function logAdmin_(admin, action, target, detail) {
  try {
    getAdminLogSheet_(getRegistrySheet_().getParent()).appendRow([new Date(), admin, action, target, String(detail).substring(0, 400)]);
  } catch (e) {}
}

/* ---------- Overview (requests and users) ---------- */
function adminOverview(token) {
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var ss = getRegistrySheet_().getParent();
    ensureGiftSetup_(ss);
    ensureWithdrawSetup_(ss);
    ensurePaymentSetup_(ss);
    expireOrders_(ss);

    var uv = getRegistrySheet_().getDataRange().getValues();
    var users = [];
    for (var i = uv.length - 1; i >= 1 && users.length < 200; i--) {
      users.push({ email: String(uv[i][1]), created: fmtDate_(uv[i][0]), balance: Number(uv[i][3]) || 0 });
    }

    var rv = getLiveSheet_(ss).getDataRange().getValues();
    var reqs = [], pending = 0, confirmed = 0;
    for (var j = rv.length - 1; j >= 1; j--) {
      var st = String(rv[j][6] || 'PENDING').trim().toUpperCase();
      var kind = String(rv[j][7] || '').trim().toUpperCase() === 'GIFT' ? 'GIFT' : 'INVESTMENT';
      if (kind !== 'GIFT') {
        if (st === 'PENDING') pending++;
        if (st === 'CONFIRMED') confirmed++;
      }
      if (reqs.length < 200) {
        reqs.push({
          row: j + 1, email: String(rv[j][1]), product: String(rv[j][2]),
          amount: Number(rv[j][3]) || 0, months: Number(rv[j][4]) || 0,
          date: fmtDate_(rv[j][0]), status: st,
          kind: kind, flag: String(rv[j][8] || '')
        });
      }
    }

    var wv = getWithdrawSheet_(ss).getDataRange().getValues();
    var wds = [], wdProcessing = 0;
    for (var w = wv.length - 1; w >= 1; w--) {
      var wst = String(wv[w][11] || 'PROCESSING').trim().toUpperCase();
      if (wst === 'PROCESSING') wdProcessing++;
      if (wds.length < 100) {
        wds.push({
          row: w + 1, email: String(wv[w][1]), plan: String(wv[w][3]),
          amount: Number(wv[w][4]) || 0, value: Number(wv[w][5]) || 0,
          early: String(wv[w][6]).toUpperCase() === 'TRUE',
          forfeit: String(wv[w][7]).toUpperCase() === 'TRUE',
          feePct: Number(wv[w][8]) || 0, fee: Number(wv[w][9]) || 0,
          payout: Number(wv[w][10]) || 0, status: wst, date: fmtDate_(wv[w][0])
        });
      }
    }

    var pv = getPaymentSheet_(ss).getDataRange().getValues();
    var pays = [], payOpen = 0;
    for (var q = pv.length - 1; q >= 1; q--) {
      var pst = String(pv[q][7] || 'PENDING').trim().toUpperCase();
      if (pst === 'PENDING' || pst === 'CONFIRMING') payOpen++;
      if (pays.length < 100) {
        pays.push({
          row: q + 1, ref: String(pv[q][1]), email: String(pv[q][2]), method: String(pv[q][4]),
          expected: Number(pv[q][5]) || 0, received: Number(pv[q][6]) || 0,
          status: pst, date: fmtDate_(pv[q][0]), note: String(pv[q][11] || '')
        });
      }
    }

    var cfg = getPlatformConfig();
    var giftProducts = [];
    if (cfg.success) {
      for (var p = 0; p < cfg.products.length; p++) {
        giftProducts.push({ id: cfg.products[p].id, name: cfg.products[p].name, apy: cfg.products[p].apy });
      }
    }
    var ps = paymentSettings_(ss);

    return {
      success: true, admin: adm.email,
      counts: { users: uv.length - 1, requests: rv.length - 1, pending: pending, confirmed: confirmed },
      users: users, requests: reqs,
      withdrawals: wds, wdProcessing: wdProcessing,
      payments: pays, payOpen: payOpen, payMode: ps.mode, payOutcomes: SANDBOX_OUTCOMES_,
      giftAmounts: choiceWords_(ss, 'GIFT_AMOUNTS', true),
      giftFlags: choiceWords_(ss, 'FLAG_WORDS', true),
      giftTerms: choiceWords_(ss, 'GIFT_TERMS', true),
      giftProducts: giftProducts
    };
  } catch (err) {
    Logger.log('adminOverview error: ' + err);
    return { success: false, message: 'Server error.' };
  }
}

function adminSetStatus(token, row, email, status) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    status = String(status || '').toUpperCase();
    if (['PENDING', 'CONFIRMED', 'REJECTED'].indexOf(status) < 0) return { success: false, message: 'Invalid status.' };
    row = Math.round(Number(row));
    email = String(email || '').toLowerCase();

    lock.waitLock(15000);
    locked = true;
    var sh = getLiveSheet_(getRegistrySheet_().getParent());
    if (!isFinite(row) || row < 2 || row > sh.getLastRow()) return { success: false, message: 'Row not found.' };
    if (String(sh.getRange(row, 2).getValue()).toLowerCase() !== email) {
      return { success: false, message: 'Row changed. Refresh and try again.' };
    }
    var old = String(sh.getRange(row, 7).getValue() || 'PENDING');
    sh.getRange(row, 7).setValue(status);
    logAdmin_(adm.email, 'SET_STATUS', email, 'row ' + row + ': ' + old + ' -> ' + status);
    return { success: true, message: 'Status updated.' };
  } catch (err) {
    Logger.log('adminSetStatus error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ---------- Payment test events (SANDBOX only; goes through the same processing as a real event) ---------- */
function adminSimulatePayment(token, ref, outcome) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var ss = getRegistrySheet_().getParent();
    ensurePaymentSetup_(ss);
    if (paymentSettings_(ss).mode !== 'SANDBOX') {
      return { success: false, message: 'Simulation works only in SANDBOX mode.' };
    }
    ref = String(ref || '').trim();
    outcome = String(outcome || '').trim();
    if (!/^AW-[A-Z0-9]{10}$/.test(ref)) return { success: false, message: 'Invalid order reference.' };
    if (SANDBOX_OUTCOMES_.indexOf(outcome) < 0) return { success: false, message: 'Pick a valid outcome.' };

    lock.waitLock(15000);
    locked = true;
    var res = simulateProviderEvent_(ss, ref, outcome);
    logAdmin_(adm.email, 'SIMULATE_PAYMENT', ref, outcome + ' -> ' + (res.message || ''));
    return res;
  } catch (err) {
    Logger.log('adminSimulatePayment error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ---------- Withdrawal decision (records the decision only; no money moves) ---------- */
function adminSetWithdrawal(token, row, email, status) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    status = String(status || '').toUpperCase();
    if (['PROCESSING', 'ACCEPTED', 'REJECTED'].indexOf(status) < 0) return { success: false, message: 'Invalid status.' };
    row = Math.round(Number(row));
    email = String(email || '').toLowerCase();

    lock.waitLock(15000);
    locked = true;
    var sh = getWithdrawSheet_(getRegistrySheet_().getParent());
    if (!isFinite(row) || row < 2 || row > sh.getLastRow()) return { success: false, message: 'Row not found.' };
    if (String(sh.getRange(row, 2).getValue()).toLowerCase() !== email) {
      return { success: false, message: 'Row changed. Refresh and try again.' };
    }
    var old = String(sh.getRange(row, 12).getValue() || 'PROCESSING');
    sh.getRange(row, 12).setValue(status);
    logAdmin_(adm.email, 'SET_WITHDRAWAL', email, 'row ' + row + ': ' + old + ' -> ' + status +
      ' | payout ' + sh.getRange(row, 11).getValue());
    return { success: true, message: 'Withdrawal ' + status.toLowerCase() + '. No money was sent by this system.' };
  } catch (err) {
    Logger.log('adminSetWithdrawal error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

function adminUnlockUser(token, email) {
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    email = normalizeEmail_(email);
    if (!isValidEmail_(email)) return { success: false, message: 'Invalid email.' };
    clearLoginLimits_(emailId_(email));
    logAdmin_(adm.email, 'UNLOCK_USER', email, 'login limits cleared');
    return { success: true, message: 'Login limits cleared.' };
  } catch (err) {
    Logger.log('adminUnlockUser error: ' + err);
    return { success: false, message: 'Server error.' };
  }
}

/* ---------- Gift credit (earns returns on the chosen plan and term) ---------- */
function adminGiftCredit(token, email, amountWord, flagWord, productId, termWord) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    email = normalizeEmail_(email);
    if (!isValidEmail_(email)) return { success: false, message: 'Pick a customer.' };

    var ss = getRegistrySheet_().getParent();
    ensureGiftSetup_(ss);

    var uv = getRegistrySheet_().getDataRange().getValues();
    var found = false;
    for (var i = 1; i < uv.length; i++) {
      if (String(uv[i][1]).trim().toLowerCase() === email) { found = true; break; }
    }
    if (!found) return { success: false, message: 'Customer not found.' };

    amountWord = String(amountWord || '').trim();
    flagWord = String(flagWord || '').trim();
    termWord = String(termWord || '').trim();
    productId = String(productId || '').trim();
    if (choiceWords_(ss, 'GIFT_AMOUNTS', true).indexOf(amountWord) < 0) return { success: false, message: 'Pick a valid amount.' };
    if (choiceWords_(ss, 'FLAG_WORDS', true).indexOf(flagWord) < 0) return { success: false, message: 'Pick a valid label.' };
    if (choiceWords_(ss, 'GIFT_TERMS', true).indexOf(termWord) < 0) return { success: false, message: 'Pick a valid term.' };
    var amount = Number(amountWord), months = Number(termWord);
    if (!isFinite(amount) || amount <= 0) return { success: false, message: 'Pick a valid amount.' };
    if (!isFinite(months) || months < 1) return { success: false, message: 'Pick a valid term.' };

    var cfg = getPlatformConfig();
    var okProd = false;
    if (cfg.success) {
      for (var p = 0; p < cfg.products.length; p++) {
        if (cfg.products[p].id === productId) okProd = true;
      }
    }
    if (!okProd) return { success: false, message: 'Pick an available plan.' };

    lock.waitLock(15000);
    locked = true;
    getLiveSheet_(ss).appendRow([new Date(), email, productId, amount, months,
      'Gift credit by admin. No payment.', 'CONFIRMED', 'GIFT', flagWord]);
    logAdmin_(adm.email, 'GIFT_CREDIT', email, amount + ' USD | plan ' + productId + ' | ' + months + ' months | label ' + flagWord);
    return { success: true, message: 'Gift of $' + amount + ' added for ' + email + ' (' + productId + ', ' + months + ' months).' };
  } catch (err) {
    Logger.log('adminGiftCredit error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ---------- Audit log viewer (read only) ---------- */
function auditTs_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  return String(v || '');
}

function adminGetAudit(token) {
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var ss = getRegistrySheet_().getParent();
    var vals = getAdminLogSheet_(ss).getDataRange().getValues();
    var rows = [];
    for (var i = vals.length - 1; i >= 1 && rows.length < 100; i--) {
      if (!vals[i][2] && !vals[i][1]) continue;
      rows.push({
        ts: auditTs_(vals[i][0]),
        admin: String(vals[i][1] || ''),
        action: String(vals[i][2] || ''),
        target: String(vals[i][3] || ''),
        detail: String(vals[i][4] || '')
      });
    }
    return { success: true, rows: rows };
  } catch (err) {
    Logger.log('adminGetAudit error: ' + err);
    return { success: false, message: 'Server error.' };
  }
}

/* ---------- Content editor ---------- */
function cellStr_(v) {
  if (v instanceof Date) return '';
  return (v === null || v === undefined) ? '' : String(v);
}

function specsFor_(ss, tabKey, rowVals) {
  var def = ADMIN_TABS_[tabKey];

  if (tabKey === 'BANNERS') {
    var words = choiceWords_(ss, 'BANNER_COLORS', true);
    var curColor = String(rowVals[3] || '').trim();
    if (curColor && words.indexOf(curColor) < 0) words.push(curColor);
    if (!words.length) words = ['Amber'];
    return [def.cols[0], def.cols[1], def.cols[2], { n: 'Color', k: 'select', opts: words }];
  }

  if (tabKey === 'CHOICES') {
    var isColor = String(rowVals[0]).trim() === 'BANNER_COLORS';
    var valSpec = isColor ? { n: 'Color code', k: 'text', req: true } : { n: 'Color code', k: 'text', ro: true };
    return [def.cols[0], def.cols[1], def.cols[2], valSpec];
  }

  if (tabKey !== 'RULES') return def.cols;
  var key = String(rowVals[0]).trim();
  var choices = RULE_CHOICES_[key] || EXTRA_RULE_CHOICES_[key];
  var valueSpec;
  if (choices) {
    valueSpec = { n: 'Value', k: 'select', opts: choices };
  } else if (key === 'FLAG_FIRST_INVESTMENT' || key === 'FLAG_INVESTMENT') {
    var fw = choiceWords_(ss, 'FLAG_WORDS', true);
    var cv = String(rowVals[1] || '').trim();
    if (cv && fw.indexOf(cv) < 0) fw.push(cv);
    if (!fw.length) fw = [cv || 'Investment'];
    valueSpec = { n: 'Value', k: 'select', opts: fw };
  } else if (key === 'WITHDRAW_FEE_PERCENT') {
    var fe = choiceWords_(ss, 'WITHDRAW_FEES', true);
    var cf = String(rowVals[1] || '').trim();
    if (cf && fe.indexOf(cf) < 0) fe.push(cf);
    if (!fe.length) fe = [cf || '5'];
    valueSpec = { n: 'Value', k: 'select', opts: fe };
  } else if (key === 'RISK_DISCLOSURE' || key === 'WITHDRAW_WARNING_TEXT' || key === 'WITHDRAW_CHECK_TEXT') {
    valueSpec = { n: 'Value', k: 'area', req: true };
  } else if (key === 'FOOTER_TEXT') {
    valueSpec = { n: 'Value', k: 'area' };
  } else {
    valueSpec = { n: 'Value', k: 'text', req: true };
  }
  return [def.cols[0], valueSpec, def.cols[2]];
}

function prepareTab_(ss, tabKey) {
  if (tabKey === 'RULES') { ensureRulesTabs_(ss); ensureVersionRule_(ss); ensureAccessRule_(ss); ensureGiftSetup_(ss); ensureWithdrawSetup_(ss); ensurePaymentSetup_(ss); }
  else if (tabKey === 'PRODUCTS' || tabKey === 'BANNERS') { ensureRulesTabs_(ss); }
  else if (tabKey === 'CHOICES') { ensureChoicesTab_(ss); ensureGiftSetup_(ss); ensureWithdrawSetup_(ss); ensurePaymentSetup_(ss); }
  else if (tabKey === 'CONSENT') { getConsentTextSheet_(ss); }
  else if (tabKey === 'LEGAL') { getLegalSheet_(ss); }
}

/** Validates input against specs. Returns {error} or {out}. */
function validateVals_(specs, vals, existing, isNew) {
  var out = [];
  for (var i = 0; i < specs.length; i++) {
    var sp = specs[i];
    if (sp.ro || (sp.roExisting && !isNew)) { out.push(existing[i]); continue; }
    var v = String(vals && vals[i] !== undefined ? vals[i] : '').trim();
    if (sp.k === 'select') {
      if (sp.opts.indexOf(v) < 0) return { error: 'Pick a valid choice for ' + sp.n + '.' };
      out.push(v);
    } else if (sp.k === 'number') {
      var n = Number(v);
      if (v === '' || !isFinite(n) || Math.round(n) !== n || n < sp.min || n > sp.max) {
        return { error: sp.n + ' must be a whole number from ' + sp.min + ' to ' + sp.max + '.' };
      }
      out.push(n);
    } else if (sp.k === 'id') {
      if (!/^[a-z0-9_-]{2,30}$/.test(v)) return { error: 'Product ID: 2-30 characters, lowercase letters, numbers, - or _.' };
      out.push(v);
    } else {
      if (v.length > 2000) return { error: sp.n + ' is too long (max 2000 characters).' };
      if (v.charAt(0) === '=') return { error: sp.n + ' cannot start with "=".' };
      if (sp.req && !v) return { error: 'Fill in ' + sp.n + '.' };
      out.push(v);
    }
  }
  return { out: out };
}

/** Extra checks for the Choices tab. Returns an error message or ''. May fix out[3]. */
function checkChoice_(ss, out, skipRow) {
  var list = String(out[0]), word = String(out[1]), active = String(out[2]), val = String(out[3] || '');
  if (list === 'GIFT_AMOUNTS') {
    if (!/^[0-9]{1,6}$/.test(word) || Number(word) < 1 || Number(word) > 100000) {
      return 'Amount must be a whole number from 1 to 100000.';
    }
  } else if (list === 'GIFT_TERMS') {
    if (!/^[0-9]{1,2}$/.test(word) || Number(word) < 1 || Number(word) > 60) {
      return 'Term must be a whole number of months from 1 to 60.';
    }
  } else if (list === 'WITHDRAW_FEES') {
    if (!/^[0-9]{1,2}$/.test(word) || Number(word) > 50) {
      return 'Fee must be a whole percent from 0 to 50.';
    }
  } else if (!/^[A-Za-z0-9][A-Za-z0-9 '&-]{0,29}$/.test(word)) {
    return 'Word: 1 to 30 characters. Letters, numbers, spaces, - & or apostrophe only.';
  }
  if (list === 'BANNER_COLORS') {
    if (!/^#[0-9a-fA-F]{6}$/.test(val)) return 'Color code must look like #34d399.';
    out[3] = val.toLowerCase();
  } else {
    out[3] = '';
  }
  var vals = ensureChoicesTab_(ss).getDataRange().getValues();
  var activeCount = (active === 'TRUE') ? 1 : 0;
  for (var i = 1; i < vals.length; i++) {
    if ((i + 1) === skipRow) continue;
    if (String(vals[i][0]).trim() !== list) continue;
    if (String(vals[i][1]).trim().toLowerCase() === word.toLowerCase()) return 'That word already exists in this list.';
    if (String(vals[i][2]).trim().toUpperCase() === 'TRUE') activeCount++;
  }
  if (activeCount < 1) return 'Keep at least one Active word in this list.';
  return '';
}

function adminGetContent(token, tabKey) {
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var def = ADMIN_TABS_[tabKey];
    if (!def) return { success: false, message: 'Unknown section.' };
    var ss = getRegistrySheet_().getParent();
    prepareTab_(ss, tabKey);
    var vals = ss.getSheetByName(def.sheet).getDataRange().getValues();
    var n = def.cols.length, rows = [];
    for (var i = 1; i < vals.length && rows.length < 300; i++) {
      var cells = vals[i].slice(0, n).map(cellStr_);
      while (cells.length < n) cells.push('');
      var blank = cells.every(function (c) { return c === ''; });
      if (blank) continue;
      var item = { row: i + 1, vals: cells };
      if (tabKey === 'RULES' || tabKey === 'BANNERS' || tabKey === 'CHOICES') item.spec = specsFor_(ss, tabKey, cells);
      rows.push(item);
    }
    return { success: true, key: tabKey, label: def.label, note: def.note, addable: def.addable, cols: def.cols, rows: rows };
  } catch (err) {
    Logger.log('adminGetContent error: ' + err);
    return { success: false, message: 'Server error.' };
  }
}

function adminSaveRow(token, tabKey, row, check, vals) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var def = ADMIN_TABS_[tabKey];
    if (!def) return { success: false, message: 'Unknown section.' };
    row = Math.round(Number(row));

    lock.waitLock(15000);
    locked = true;
    var ss = getRegistrySheet_().getParent();
    var sh = ss.getSheetByName(def.sheet);
    var n = def.cols.length;
    if (!sh || !isFinite(row) || row < 2 || row > sh.getLastRow()) return { success: false, message: 'Row not found. Refresh.' };

    var current = sh.getRange(row, 1, 1, n).getValues()[0].map(cellStr_);
    if (current[0] !== String(check)) return { success: false, message: 'This row changed. Refresh and try again.' };

    var specs = specsFor_(ss, tabKey, current);
    var res = validateVals_(specs, vals, current, false);
    if (res.error) return { success: false, message: res.error };

    if (tabKey === 'CHOICES') {
      var cerr = checkChoice_(ss, res.out, row);
      if (cerr) return { success: false, message: cerr };
    }

    var changes = [];
    for (var i = 0; i < n; i++) {
      var sp = specs[i];
      if (sp.ro || sp.roExisting) continue;
      if (String(res.out[i]) !== current[i]) {
        sh.getRange(row, i + 1).setValue(res.out[i]);
        changes.push(sp.n + ': "' + current[i].substring(0, 60) + '" -> "' + String(res.out[i]).substring(0, 60) + '"');
      }
    }
    if (!changes.length) return { success: true, message: 'No changes.', check: current[0] };
    logAdmin_(adm.email, 'EDIT_' + tabKey, current[0], 'row ' + row + ' | ' + changes.join(' | '));
    return { success: true, message: 'Saved. Live after refresh.', check: String(res.out[0]) };
  } catch (err) {
    Logger.log('adminSaveRow error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

function adminAddRow(token, tabKey, vals) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var def = ADMIN_TABS_[tabKey];
    if (!def || !def.addable) return { success: false, message: 'Adding is not allowed here.' };

    lock.waitLock(15000);
    locked = true;
    var ss = getRegistrySheet_().getParent();
    prepareTab_(ss, tabKey);
    var sh = ss.getSheetByName(def.sheet);

    var blank = def.cols.map(function () { return ''; });
    var res = validateVals_(def.cols, vals, blank, true);
    if (res.error) return { success: false, message: res.error };

    if (tabKey === 'PRODUCTS') {
      var ex = sh.getDataRange().getValues();
      for (var i = 1; i < ex.length; i++) {
        if (String(ex[i][0]).trim().toLowerCase() === res.out[0]) return { success: false, message: 'That Product ID already exists.' };
      }
    }
    if (tabKey === 'CHOICES') {
      var cerr = checkChoice_(ss, res.out, -1);
      if (cerr) return { success: false, message: cerr };
    }
    sh.appendRow(res.out);
    logAdmin_(adm.email, 'ADD_' + tabKey, String(res.out[0]), res.out.join(' | '));
    return { success: true, message: 'Added.' };
  } catch (err) {
    Logger.log('adminAddRow error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}