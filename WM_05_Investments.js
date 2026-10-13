/**
 * WM_05_Investments.gs
 * Apex Wealth - account view, dashboard data, plan requests, gift credits, withdrawal requests.
 * Tab: Live_Interest
 * Columns: Timestamp | Email | ProductId | Amount | Months | Note | Status | Kind | Flag
 * Status (dropdown): PENDING / CONFIRMED / REJECTED
 * Kind: blank or INVESTMENT = customer plan request. GIFT = admin credit (Flag holds its label word).
 * Gifts carry a ProductId and Months and earn returns like an investment.
 * Payments: a plan request creates a payment order (see WM_10_Payments.gs). CONFIRMED comes from
 * processPaymentEvent_ only (manual CONFIRMED in admin is for testing).
 * Tab: Withdrawals (one row per request, linked to a Live_Interest row)
 * Columns: Timestamp | Email | SourceRow | Plan | Amount | Value | Early | Forfeit | FeePercent | Fee | Payout | Status | Ack | Note
 * Status: PROCESSING / ACCEPTED / REJECTED. ACCEPTED closes the plan. No money moves in this system.
 * Rule: DASHBOARD_ACCESS in Platform_Rules
 *   TEST_AFTER_REQUEST      = any non-rejected request unlocks the dashboard
 *   AFTER_PAYMENT_CONFIRMED = only CONFIRMED requests unlock it
 * Flag words for first and later investments: FLAG_FIRST_INVESTMENT, FLAG_INVESTMENT in Platform_Rules.
 * Withdrawal rules: WITHDRAW_ENABLED, WITHDRAW_FEE_PERCENT, WITHDRAW_FORFEIT_RETURNS,
 *   WITHDRAW_WARNING_TEXT, WITHDRAW_CHECK_TEXT in Platform_Rules.
 */

var LIVE_TAB = 'Live_Interest';
var WD_TAB = 'Withdrawals';
var ACCESS_KEY = 'DASHBOARD_ACCESS';
var MONTH_MS_ = 30.4375 * 24 * 60 * 60 * 1000;
var DEFAULT_FLAG_FIRST_ = 'First investment';
var DEFAULT_FLAG_INV_ = 'Investment';
var DEFAULT_WD_WARN_ = 'Withdrawing before this plan matures means you lose the returns earned on it, and an early withdrawal fee of {FEE}% applies. Your request is reviewed by our team before anything is paid.';
var DEFAULT_WD_CHECK_ = 'I understand and accept the early withdrawal terms above.';

function ddl_(list) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a value from the list.')
    .build();
}

function getLiveSheet_(ss) {
  var sh = ss.getSheetByName(LIVE_TAB);
  if (!sh) {
    sh = ss.insertSheet(LIVE_TAB);
    sh.appendRow(['Timestamp', 'Email', 'ProductId', 'Amount', 'Months', 'Note', 'Status', 'Kind', 'Flag']);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (!String(sh.getRange(1, 7).getValue())) {
    sh.getRange(1, 7).setValue('Status').setFontWeight('bold');
  }
  if (!String(sh.getRange(1, 8).getValue())) {
    sh.getRange(1, 8).setValue('Kind').setFontWeight('bold');
  }
  if (!String(sh.getRange(1, 9).getValue())) {
    sh.getRange(1, 9).setValue('Flag').setFontWeight('bold');
  }
  sh.getRange(2, 7, 500, 1).setNumberFormat('@')
    .setDataValidation(ddl_(['PENDING', 'CONFIRMED', 'REJECTED']));
  return sh;
}

function getWithdrawSheet_(ss) {
  var sh = ss.getSheetByName(WD_TAB);
  if (!sh) {
    sh = ss.insertSheet(WD_TAB);
    sh.appendRow(['Timestamp', 'Email', 'SourceRow', 'Plan', 'Amount', 'Value', 'Early', 'Forfeit',
      'FeePercent', 'Fee', 'Payout', 'Status', 'Ack', 'Note']);
    sh.getRange(1, 1, 1, 14).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 12, 500, 1).setNumberFormat('@')
      .setDataValidation(ddl_(['PROCESSING', 'ACCEPTED', 'REJECTED']));
  }
  return sh;
}

function ensureAccessRule_(ss) {
  var rules = ss.getSheetByName('Platform_Rules');
  if (!rules) return;
  var vals = rules.getDataRange().getValues();
  var row = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === ACCESS_KEY) { row = i + 1; break; }
  }
  if (row < 0) {
    rules.appendRow([ACCESS_KEY, 'TEST_AFTER_REQUEST',
      'TEST_AFTER_REQUEST = dashboard opens after any request (testing). AFTER_PAYMENT_CONFIRMED = only after Status is CONFIRMED in Live_Interest.']);
    row = rules.getLastRow();
  }
  rules.getRange(row, 2).setNumberFormat('@')
    .setDataValidation(ddl_(['TEST_AFTER_REQUEST', 'AFTER_PAYMENT_CONFIRMED']));
}

function getAccess_(ss) {
  try {
    var v = readRules_(ss)[ACCESS_KEY];
    return v === 'AFTER_PAYMENT_CONFIRMED' ? 'AFTER_PAYMENT_CONFIRMED' : 'TEST_AFTER_REQUEST';
  } catch (e) {
    return 'TEST_AFTER_REQUEST';
  }
}

/** Adds flag words, gift amount steps, gift terms and flag rules once. Safe to call often. */
function ensureGiftSetup_(ss) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('GIFT_SETUP_V2') === '1') return;
  ensureRulesTabs_(ss);
  var sh = ensureChoicesTab_(ss);
  var vals = sh.getDataRange().getValues();
  var have = {}, amountRows = 0, termRows = 0;
  for (var i = 1; i < vals.length; i++) {
    var l = String(vals[i][0]).trim();
    have[l + '|' + String(vals[i][1]).trim().toLowerCase()] = true;
    if (l === 'GIFT_AMOUNTS') amountRows++;
    if (l === 'GIFT_TERMS') termRows++;
  }
  var add = [];
  [DEFAULT_FLAG_FIRST_, DEFAULT_FLAG_INV_].forEach(function (w) {
    if (!have['FLAG_WORDS|' + w.toLowerCase()]) add.push(['FLAG_WORDS', w, 'TRUE', '']);
  });
  if (!amountRows) {
    ['25', '50', '100', '250', '500', '1000'].forEach(function (a) { add.push(['GIFT_AMOUNTS', a, 'TRUE', '']); });
  }
  if (!termRows) {
    ['3', '6', '12', '18', '24', '36'].forEach(function (a) { add.push(['GIFT_TERMS', a, 'TRUE', '']); });
  }
  if (add.length) {
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setNumberFormat('@').setValues(add);
  }

  var rules = ss.getSheetByName('Platform_Rules');
  var rv = rules.getDataRange().getValues();
  var seen = {};
  for (var r = 1; r < rv.length; r++) seen[String(rv[r][0]).trim()] = true;
  if (!seen.FLAG_FIRST_INVESTMENT) {
    rules.appendRow(['FLAG_FIRST_INVESTMENT', DEFAULT_FLAG_FIRST_, 'Label on a customer\'s first investment. Pick from the Choices list (FLAG_WORDS).']);
  }
  if (!seen.FLAG_INVESTMENT) {
    rules.appendRow(['FLAG_INVESTMENT', DEFAULT_FLAG_INV_, 'Label on later investments. Pick from the Choices list (FLAG_WORDS).']);
  }
  props.setProperty('GIFT_SETUP_V2', '1');
}

/** Adds withdrawal rules, fee choices and the Withdrawals tab once. Safe to call often. */
function ensureWithdrawSetup_(ss) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('WITHDRAW_SETUP_V1') === '1') return;
  ensureRulesTabs_(ss);
  var sh = ensureChoicesTab_(ss);
  var vals = sh.getDataRange().getValues();
  var feeRows = 0;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'WITHDRAW_FEES') feeRows++;
  }
  if (!feeRows) {
    var add = [];
    ['0', '1', '2', '3', '5', '10', '15', '20', '25'].forEach(function (a) { add.push(['WITHDRAW_FEES', a, 'TRUE', '']); });
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setNumberFormat('@').setValues(add);
  }

  var rules = ss.getSheetByName('Platform_Rules');
  var rv = rules.getDataRange().getValues();
  var seen = {};
  for (var r = 1; r < rv.length; r++) seen[String(rv[r][0]).trim()] = true;
  var defs = [
    ['WITHDRAW_ENABLED', 'TRUE', 'TRUE lets customers send withdrawal requests. FALSE closes the Withdraw button.'],
    ['WITHDRAW_FEE_PERCENT', '5', 'Fee on EARLY withdrawals, in percent. Pick from the Choices list (WITHDRAW_FEES).'],
    ['WITHDRAW_FORFEIT_RETURNS', 'TRUE', 'TRUE = an early withdrawal loses the returns earned so far. FALSE = returns are kept.'],
    ['WITHDRAW_WARNING_TEXT', DEFAULT_WD_WARN_, 'Typed text. Pop-up warning for early withdrawals. Use {FEE} for the fee percent. Keep it consistent with the fee and forfeit settings. Attorney should review.'],
    ['WITHDRAW_CHECK_TEXT', DEFAULT_WD_CHECK_, 'Typed text. Wording next to the tick box in the early withdrawal pop-up.']
  ];
  defs.forEach(function (d) { if (!seen[d[0]]) rules.appendRow(d); });
  getWithdrawSheet_(ss);
  props.setProperty('WITHDRAW_SETUP_V1', '1');
}

function setupInvestments() {
  var ss = getRegistrySheet_().getParent();
  getLiveSheet_(ss);
  setupRules();
  ensureAccessRule_(ss);
  ensureGiftSetup_(ss);
  ensureWithdrawSetup_(ss);
  ensurePaymentSetup_(ss);
  Logger.log('Setup OK');
  return 'Setup OK';
}

/** Run ONCE if old demo tabs still exist. */
function cleanupDemo() {
  var ss = getRegistrySheet_().getParent();
  ['Demo_Status', 'Investments'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s) ss.deleteSheet(s);
  });
  Logger.log('Cleanup done.');
  return 'Cleanup done';
}

function fmtDate_(d) {
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMM d, yyyy'); }
  catch (e) { return ''; }
}

function round2_(n) { return Math.round(n * 100) / 100; }

/** Maths for one Live_Interest row. prod = map of active products by id. */
function computePlan_(r, prod, now, isGift) {
  var pid = String(r[2]);
  var pinfo = prod[pid];
  var amount = Number(r[3]) || 0;
  var months = Number(r[4]) || 0;
  var apy = pinfo ? pinfo.apy : 0;
  var start = new Date(r[0]).getTime();
  var maturityTs = start + months * MONTH_MS_;
  var noTerm = (isGift && months === 0);
  var el = Math.max(0, Math.min(months, (now - start) / MONTH_MS_));
  var accrued = amount * (Math.pow(1 + apy / 12, el) - 1);
  return {
    name: pinfo ? pinfo.name : (isGift ? 'Account credit' : pid),
    amount: amount, months: months, apy: apy,
    maturityTs: maturityTs, noTerm: noTerm,
    matured: noTerm ? true : (now >= maturityTs),
    accrued: accrued
  };
}

/** Map of Live_Interest row number -> latest withdrawal status (PROCESSING/ACCEPTED/REJECTED). */
function withdrawMap_(ss, email) {
  var map = {};
  var sh = ss.getSheetByName(WD_TAB);
  if (!sh || sh.getLastRow() < 2) return map;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (email && String(v[i][1]).trim().toLowerCase() !== email) continue;
    var src = Number(v[i][2]);
    if (!isFinite(src)) continue;
    map[src] = String(v[i][11] || '').trim().toUpperCase();
  }
  return map;
}

function getAccount(token) {
  try {
    var s = restoreSession(token);
    if (!s.success) return s;
    var ss = getRegistrySheet_().getParent();
    ensureAccessRule_(ss);
    ensureWithdrawSetup_(ss);
    ensurePaymentSetup_(ss);
    expireOrders_(ss);
    var access = getAccess_(ss);
    var cfg = getPlatformConfig();
    var rl = readRules_(ss);
    var flagFirst = rl.FLAG_FIRST_INVESTMENT || DEFAULT_FLAG_FIRST_;
    var flagInv = rl.FLAG_INVESTMENT || DEFAULT_FLAG_INV_;
    var wdEnabled = (rl.WITHDRAW_ENABLED === undefined) ? true : bool_(rl.WITHDRAW_ENABLED);
    var wdMap = withdrawMap_(ss, s.email);
    var pm = payMap_(ss, s.email);

    var prod = {};
    if (cfg.success) {
      for (var p = 0; p < cfg.products.length; p++) prod[cfg.products[p].id] = cfg.products[p];
    }

    var vals = getLiveSheet_(ss).getDataRange().getValues();
    var now = Date.now();
    var requests = [], plans = [];
    var principal = 0, profit = 0, activeCount = 0, nextMat = null;
    var alloc = {};

    // Row of this customer's first real investment (gifts and rejected rows do not count)
    var firstIdx = -1;
    for (var f = 1; f < vals.length; f++) {
      if (String(vals[f][1]).toLowerCase() !== s.email) continue;
      if (String(vals[f][7] || '').trim().toUpperCase() === 'GIFT') continue;
      if (String(vals[f][6] || 'PENDING').trim().toUpperCase() === 'REJECTED') continue;
      firstIdx = f; break;
    }

    for (var i = vals.length - 1; i >= 1; i--) {
      var r = vals[i];
      if (String(r[1]).toLowerCase() !== s.email) continue;
      var status = String(r[6] || 'PENDING').trim().toUpperCase();
      var isGift = String(r[7] || '').trim().toUpperCase() === 'GIFT';
      var cp = computePlan_(r, prod, now, isGift);
      var flag = isGift ? String(r[8] || '') : (i === firstIdx ? flagFirst : flagInv);
      var wd = wdMap[i + 1] || '';
      if (wd === 'REJECTED') wd = '';
      var payInfo = isGift ? null : (pm[i + 1] || null);

      requests.push({
        row: i + 1, name: cp.name, amount: cp.amount, months: cp.months,
        date: fmtDate_(r[0]), status: status,
        kind: isGift ? 'GIFT' : 'INVESTMENT', flag: flag, wd: wd, pay: payInfo
      });

      var counted = (access === 'AFTER_PAYMENT_CONFIRMED') ? (status === 'CONFIRMED') : (status !== 'REJECTED');
      if (!counted) continue;

      var closed = (wd === 'ACCEPTED');
      if (!closed) {
        principal += cp.amount;
        profit += cp.accrued;
        if (!cp.matured) {
          activeCount++;
          if (nextMat === null || cp.maturityTs < nextMat) nextMat = cp.maturityTs;
        }
        alloc[cp.name] = (alloc[cp.name] || 0) + cp.amount;
      }
      plans.push({
        row: i + 1, name: cp.name, amount: cp.amount, months: cp.months, apy: cp.apy,
        date: fmtDate_(r[0]), maturity: cp.noTerm ? '' : fmtDate_(cp.maturityTs),
        matured: cp.matured, status: status,
        kind: isGift ? 'GIFT' : 'INVESTMENT', flag: flag,
        wd: wd, closed: closed, canWithdraw: (!wd && wdEnabled),
        accrued: round2_(cp.accrued), value: round2_(cp.amount + cp.accrued)
      });
    }

    var allocation = [];
    for (var k in alloc) {
      allocation.push({ name: k, pct: principal > 0 ? Math.round(alloc[k] / principal * 100) : 0 });
    }
    allocation.sort(function (a, b) { return b.pct - a.pct; });

    return {
      success: true,
      email: s.email,
      access: access,
      withdrawEnabled: wdEnabled,
      unlocked: plans.length > 0,
      balance: round2_(principal + profit),
      principal: round2_(principal),
      profit: round2_(profit),
      profitPct: principal > 0 ? Math.round(profit / principal * 1000) / 10 : 0,
      activePlans: activeCount,
      totalPlans: plans.length,
      nextMaturity: nextMat ? fmtDate_(nextMat) : '',
      plans: plans,
      allocation: allocation,
      requests: requests
    };
  } catch (err) {
    Logger.log('getAccount error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

/**
 * Saves a plan request as PENDING and, when payments are open, creates a payment order.
 * No payment is taken here. Requires valid consent (see CONSENT_MODE).
 */
function submitLiveInterest(token, productId, amount, months, method) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };

    var ss = getRegistrySheet_().getParent();

    amount = Number(amount); months = Math.round(Number(months));
    if (!isFinite(amount) || !isFinite(months)) return { success: false, message: 'Invalid amount.' };
    var bad = validateAgainstRules_(amount, months);
    if (bad) return { success: false, message: bad };

    var cfg = getPlatformConfig();
    var ok = false;
    for (var i = 0; i < cfg.products.length; i++) {
      if (cfg.products[i].id === String(productId)) ok = true;
    }
    if (!ok) return { success: false, message: 'That plan is not available.' };

    ensurePaymentSetup_(ss);
    var pay = paymentSettings_(ss);
    method = String(method || '').trim();
    if (pay.open && pay.methods.indexOf(method) < 0) {
      return { success: false, message: 'Please pick a payment method.' };
    }

    lock.waitLock(15000);
    locked = true;

    var detail = String(productId) + ' | ' + amount + ' USD | ' + months + ' months';
    if (!consumeConsent_(ss, s.email, detail)) {
      return { success: false, message: 'Please accept the consent first.' };
    }

    var sh = getLiveSheet_(ss);
    sh.appendRow([new Date(), s.email, String(productId), amount, months,
      'Request only. No payment taken.', 'PENDING', 'INVESTMENT', '']);

    var msg = 'Request saved. Payments are not open yet.';
    var ref = '';
    if (pay.open) {
      ref = createPaymentOrder_(ss, s.email, sh.getLastRow(), method, amount);
      msg = 'Order ' + ref + ' created.' + (pay.mode === 'SANDBOX' ? ' TEST MODE: no real payment is taken.' : '');
    }
    return { success: true, message: msg, ref: ref };
  } catch (err) {
    Logger.log('submitLiveInterest error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/* ---------- Withdrawal requests (no money moves here) ---------- */

/** Checks one plan row is withdrawable and works out the figures on the SERVER. Returns {error} or {ok,...}. */
function withdrawContext_(ss, email, row) {
  var rl = readRules_(ss);
  var enabled = (rl.WITHDRAW_ENABLED === undefined) ? true : bool_(rl.WITHDRAW_ENABLED);
  if (!enabled) return { error: 'Withdrawals are not open right now.' };

  row = Math.round(Number(row));
  var sh = getLiveSheet_(ss);
  if (!isFinite(row) || row < 2 || row > sh.getLastRow()) return { error: 'Plan not found.' };
  var r = sh.getRange(row, 1, 1, 9).getValues()[0];
  if (String(r[1]).trim().toLowerCase() !== email) return { error: 'Plan not found.' };

  var status = String(r[6] || 'PENDING').trim().toUpperCase();
  if (status === 'REJECTED') return { error: 'This plan is not available.' };
  if (getAccess_(ss) === 'AFTER_PAYMENT_CONFIRMED' && status !== 'CONFIRMED') {
    return { error: 'This plan is not active yet.' };
  }

  var wd = withdrawMap_(ss, email)[row] || '';
  if (wd === 'PROCESSING') return { error: 'A withdrawal request is already being processed for this plan.' };
  if (wd === 'ACCEPTED') return { error: 'This plan has already been withdrawn.' };

  var cfg = getPlatformConfig();
  var prod = {};
  if (cfg.success) {
    for (var p = 0; p < cfg.products.length; p++) prod[cfg.products[p].id] = cfg.products[p];
  }
  var isGift = String(r[7] || '').trim().toUpperCase() === 'GIFT';
  var plan = computePlan_(r, prod, Date.now(), isGift);

  var feePct = Math.max(0, Math.min(50, num_(rl.WITHDRAW_FEE_PERCENT, 5)));
  var forfeit = (rl.WITHDRAW_FORFEIT_RETURNS === undefined) ? true : bool_(rl.WITHDRAW_FORFEIT_RETURNS);
  var early = !plan.matured;
  var value = round2_(plan.amount + plan.accrued);
  var base = (early && forfeit) ? plan.amount : value;
  var fee = early ? round2_(base * feePct / 100) : 0;
  var payout = round2_(base - fee);
  var lost = (early && forfeit) ? round2_(value - plan.amount) : 0;

  return {
    ok: true, row: row, planName: plan.name, amount: plan.amount, value: value,
    early: early, forfeit: (early && forfeit), feePct: early ? feePct : 0,
    fee: fee, payout: payout, lost: lost,
    warning: early ? String(rl.WITHDRAW_WARNING_TEXT || DEFAULT_WD_WARN_).replace(/\{FEE\}/g, String(feePct)) : '',
    checkText: early ? String(rl.WITHDRAW_CHECK_TEXT || DEFAULT_WD_CHECK_) : ''
  };
}

function getWithdrawalPreview(token, row) {
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    var ss = getRegistrySheet_().getParent();
    ensureWithdrawSetup_(ss);
    var c = withdrawContext_(ss, s.email, row);
    if (c.error) return { success: false, message: c.error };
    return {
      success: true, planName: c.planName, amount: c.amount, value: c.value,
      early: c.early, forfeit: c.forfeit, lost: c.lost, feePct: c.feePct,
      fee: c.fee, payout: c.payout, warning: c.warning, checkText: c.checkText
    };
  } catch (err) {
    Logger.log('getWithdrawalPreview error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

/** Saves a withdrawal REQUEST as PROCESSING. An admin must decide. No money moves. */
function requestWithdrawal(token, row, ack) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    var ss = getRegistrySheet_().getParent();
    ensureWithdrawSetup_(ss);

    lock.waitLock(15000);
    locked = true;

    var c = withdrawContext_(ss, s.email, row);
    if (c.error) return { success: false, message: c.error };
    if (c.early && ack !== true) return { success: false, message: 'Please tick the box to confirm.' };

    getWithdrawSheet_(ss).appendRow([new Date(), s.email, c.row, c.planName, c.amount, c.value,
      c.early ? 'TRUE' : 'FALSE', c.forfeit ? 'TRUE' : 'FALSE', c.feePct, c.fee, c.payout,
      'PROCESSING', c.early ? 'YES' : 'N/A', 'Request only. No money moves automatically.']);
    return { success: true, message: 'Request received. Status: Processing request.' };
  } catch (err) {
    Logger.log('requestWithdrawal error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}