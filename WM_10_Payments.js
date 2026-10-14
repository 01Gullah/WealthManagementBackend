/**
 * WM_10_Payments.gs
 * Apex Wealth - payment orders (provider-neutral). Tab: Payments
 * Columns: Timestamp | OrderRef | Email | SourceRow | Method | ExpectedUSD | ReceivedUSD |
 *          Status | ProviderRef | Source | Updated | Note
 * Status: PENDING / CONFIRMING / CONFIRMED / OVERPAID / UNDERPAID / EXPIRED / FAILED
 * Rules: PAYMENT_MODE (SANDBOX / LIVE), ORDER_EXPIRY_MINUTES. Methods: Choices list PAY_METHODS.
 *
 * SAFETY: A plan is confirmed ONLY inside processPaymentEvent_, which must be fed by a VERIFIED
 * provider event. No customer click, tx hash, redirect or editable status may call it.
 * SANDBOX: the admin simulator feeds it test events. LIVE stays closed until PROVIDER_CONNECTED_
 * is true AND a real adapter exists.
 *
 * REAL PROVIDER LATER (design): for each non-final order, call the provider's status API with the
 * order reference (server side, key in Script Properties), map the answer to an event
 * {kind, amount, providerRef} and pass it to processPaymentEvent_ under the script lock.
 * Apps Script doPost cannot read headers, so do not rely on header signatures.
 */

var PAY_TAB = 'Payments';
var PROVIDER_CONNECTED_ = false;
var PAY_FINAL_ = ['CONFIRMED', 'OVERPAID', 'UNDERPAID', 'EXPIRED', 'FAILED'];
var SANDBOX_OUTCOMES_ = ['Payment detected', 'Paid in full', 'Paid too little', 'Paid too much', 'Failed', 'Expired'];
var DEFAULT_PAY_METHODS_ = ['USDT on TRON', 'USDT on Ethereum', 'USDC on Ethereum', 'Bitcoin'];

function getPaymentSheet_(ss) {
  var sh = ss.getSheetByName(PAY_TAB);
  if (!sh) {
    sh = ss.insertSheet(PAY_TAB);
    sh.appendRow(['Timestamp', 'OrderRef', 'Email', 'SourceRow', 'Method', 'ExpectedUSD', 'ReceivedUSD',
      'Status', 'ProviderRef', 'Source', 'Updated', 'Note']);
    sh.getRange(1, 1, 1, 12).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 8, 500, 1).setNumberFormat('@')
      .setDataValidation(ddl_(['PENDING', 'CONFIRMING', 'CONFIRMED', 'OVERPAID', 'UNDERPAID', 'EXPIRED', 'FAILED']));
  }
  return sh;
}

/** Adds payment rules, method choices and the Payments tab once. Safe to call often. */
function ensurePaymentSetup_(ss) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('PAYMENT_SETUP_V1') === '1') return;
  ensureRulesTabs_(ss);
  var sh = ensureChoicesTab_(ss);
  var vals = sh.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'PAY_METHODS') n++;
  }
  if (!n) {
    var add = DEFAULT_PAY_METHODS_.map(function (m) { return ['PAY_METHODS', m, 'TRUE', '']; });
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setNumberFormat('@').setValues(add);
  }
  var rules = ss.getSheetByName('Platform_Rules');
  var rv = rules.getDataRange().getValues();
  var seen = {};
  for (var r = 1; r < rv.length; r++) seen[String(rv[r][0]).trim()] = true;
  if (!seen.PAYMENT_MODE) {
    rules.appendRow(['PAYMENT_MODE', 'SANDBOX',
      'SANDBOX = test orders only, no real money. LIVE stays closed until a verified payment provider is connected in code.']);
  }
  if (!seen.ORDER_EXPIRY_MINUTES) {
    rules.appendRow(['ORDER_EXPIRY_MINUTES', '60', 'Minutes before an unpaid order expires.']);
  }
  getPaymentSheet_(ss);
  props.setProperty('PAYMENT_SETUP_V1', '1');
}

/** {mode, open, methods}. open = customers can create orders. */
function paymentSettings_(ss) {
  var rl = readRules_(ss);
  var mode = (String(rl.PAYMENT_MODE || 'SANDBOX').toUpperCase() === 'LIVE') ? 'LIVE' : 'SANDBOX';
  var methods = choiceWords_(ss, 'PAY_METHODS', true);
  var open = (mode === 'SANDBOX' || PROVIDER_CONNECTED_) && methods.length > 0;
  return { mode: mode, open: open, methods: methods };
}

/** Public, non-sensitive. The portal uses it to show the payment method dropdown. */
function getPaymentOptions() {
  try {
    var ss = getRegistrySheet_().getParent();
    ensurePaymentSetup_(ss);
    var p = paymentSettings_(ss);
    return { success: true, mode: p.mode, open: p.open, methods: p.methods };
  } catch (err) {
    Logger.log('getPaymentOptions error: ' + err);
    return { success: false };
  }
}

/** Creates a PENDING order for a Live_Interest row. Call inside the script lock. Returns the order reference. */
function createPaymentOrder_(ss, email, sourceRow, method, amount) {
  var ref = 'AW-' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
  getPaymentSheet_(ss).appendRow([new Date(), ref, email, sourceRow, method, round2_(amount), 0,
    'PENDING', '', 'CUSTOMER', new Date(), 'Order created. Awaiting payment.']);
  return ref;
}

/** Map of Live_Interest row number -> latest {ref, status} for one customer. */
function payMap_(ss, email) {
  var map = {};
  var sh = ss.getSheetByName(PAY_TAB);
  if (!sh || sh.getLastRow() < 2) return map;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][2]).trim().toLowerCase() !== email) continue;
    var src = Number(v[i][3]);
    if (!isFinite(src)) continue;
    map[src] = { ref: String(v[i][1]), status: String(v[i][7] || 'PENDING').trim().toUpperCase() };
  }
  return map;
}

/** Expires unpaid PENDING orders older than ORDER_EXPIRY_MINUTES. Skips quietly if busy. */
function expireOrders_(ss) {
  var sh = ss.getSheetByName(PAY_TAB);
  if (!sh || sh.getLastRow() < 2) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) return;
  try {
    var mins = num_(readRules_(ss).ORDER_EXPIRY_MINUTES, 60);
    var cutoff = Date.now() - mins * 60000;
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][7]).trim().toUpperCase() !== 'PENDING') continue;
      if (new Date(v[i][0]).getTime() >= cutoff) continue;
      sh.getRange(i + 1, 8).setValue('EXPIRED');
      sh.getRange(i + 1, 11).setValue(new Date());
      sh.getRange(i + 1, 12).setValue('Order expired with no payment.');
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * THE ONLY PLACE a payment can confirm a plan. Call under the script lock, with a VERIFIED event:
 * ev = { kind: 'confirming'|'paid'|'failed'|'expired', amount, providerRef, source }
 * Idempotent: orders in a final state are never changed again.
 */
function processPaymentEvent_(ss, ref, ev) {
  var sh = getPaymentSheet_(ss);
  var v = sh.getDataRange().getValues();
  var idx = -1;
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1]) === ref) { idx = i; break; }
  }
  if (idx < 0) return { success: false, message: 'Order not found.' };
  var row = idx + 1;
  var status = String(v[idx][7] || 'PENDING').trim().toUpperCase();
  var expected = Number(v[idx][5]) || 0;
  var src = String(ev.source || 'PROVIDER');

  if (PAY_FINAL_.indexOf(status) >= 0) {
    return { success: true, duplicate: true, message: 'Already processed (' + status + '). No change.' };
  }

  var newStatus = status, received = Number(v[idx][6]) || 0, note = '';
  if (ev.kind === 'confirming') {
    if (status === 'CONFIRMING') return { success: true, duplicate: true, message: 'Already marked as detected. No change.' };
    newStatus = 'CONFIRMING';
    note = 'Payment detected. Waiting for confirmations.';
  } else if (ev.kind === 'paid') {
    received = round2_(Number(ev.amount) || 0);
    if (received < expected * 0.995) {
      newStatus = 'UNDERPAID';
      note = 'Received ' + received + ' of ' + expected + ' USD. Not confirmed. Handle manually (top-up or refund).';
    } else if (received > expected * 1.005) {
      newStatus = 'OVERPAID';
      note = 'Excess of ' + round2_(received - expected) + ' USD received. Plan confirmed. Refund the excess manually.';
    } else {
      newStatus = 'CONFIRMED';
      note = 'Paid in full.';
    }
  } else if (ev.kind === 'failed') {
    newStatus = 'FAILED';
    note = 'Payment failed.';
  } else if (ev.kind === 'expired') {
    newStatus = 'EXPIRED';
    note = 'Order expired.';
  } else {
    return { success: false, message: 'Unknown event.' };
  }

  sh.getRange(row, 7, 1, 6).setValues([[received, newStatus, String(ev.providerRef || ''), src, new Date(), note]]);

  if (newStatus === 'CONFIRMED' || newStatus === 'OVERPAID') {
    var live = getLiveSheet_(ss);
    var srcRow = Number(v[idx][3]);
    if (srcRow >= 2 && srcRow <= live.getLastRow() &&
        String(live.getRange(srcRow, 2).getValue()).trim().toLowerCase() === String(v[idx][2]).trim().toLowerCase() &&
        String(live.getRange(srcRow, 7).getValue() || 'PENDING').trim().toUpperCase() === 'PENDING') {
      live.getRange(srcRow, 7).setValue('CONFIRMED');
    }
  }
  logAdmin_('SYSTEM', 'PAYMENT_' + newStatus, ref, src + ' | expected ' + expected + ' | received ' + received);
  return { success: true, status: newStatus, message: 'Order ' + ref + ' is now ' + newStatus + '.' };
}

/** SANDBOX ONLY: builds a test event and feeds it through processPaymentEvent_. */
function simulateProviderEvent_(ss, ref, outcome) {
  var sh = getPaymentSheet_(ss);
  var v = sh.getDataRange().getValues();
  var expected = null;
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][1]) === ref) { expected = Number(v[i][5]) || 0; break; }
  }
  if (expected === null) return { success: false, message: 'Order not found.' };
  var ev = { source: 'SANDBOX', providerRef: 'SANDBOX-' + ref + '-' + outcome.replace(/\s+/g, '_').toUpperCase() };
  if (outcome === 'Payment detected') ev.kind = 'confirming';
  else if (outcome === 'Paid in full') { ev.kind = 'paid'; ev.amount = expected; }
  else if (outcome === 'Paid too little') { ev.kind = 'paid'; ev.amount = expected * 0.5; }
  else if (outcome === 'Paid too much') { ev.kind = 'paid'; ev.amount = expected * 1.2; }
  else if (outcome === 'Failed') ev.kind = 'failed';
  else if (outcome === 'Expired') ev.kind = 'expired';
  else return { success: false, message: 'Unknown outcome.' };
  return processPaymentEvent_(ss, ref, ev);
}
