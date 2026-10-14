/**
 * WM_12_Wallet_Flow.gs
 * Apex Wealth - direct-wallet payment flow, TEST MODE.
 * Tab: Wallet_Submissions
 * Columns: Timestamp | Email | SourceRow | Asset | Amount | Txid | Status | Reviewer | ReviewedAt | Note
 * Status: SUBMITTED / APPROVED / REJECTED
 * SAFETY: real wallet addresses are never shown, and approvals only work while PAYMENT_MODE is SANDBOX,
 * unless WALLET_LIVE_ENABLED_ is changed to true in this file AND PAYMENT_MODE is LIVE.
 * A customer-supplied txid is never proof of payment. Approval is a manual, logged TEST decision and
 * goes through processPaymentEvent_ (WM_10), the only place a plan can be confirmed.
 */

var WALLET_TAB = 'Wallet_Submissions';
var WALLET_LIVE_ENABLED_ = false;
var TEST_ADDRESS_ = 'TEST-ADDRESS-NOT-REAL-DO-NOT-SEND';

function walletDd_(list) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a value from the list.')
    .build();
}

function getWalletSheet_(ss) {
  var sh = ss.getSheetByName(WALLET_TAB);
  if (!sh) {
    sh = ss.insertSheet(WALLET_TAB);
    sh.appendRow(['Timestamp', 'Email', 'SourceRow', 'Asset', 'Amount', 'Txid', 'Status', 'Reviewer', 'ReviewedAt', 'Note']);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 6, 500, 1).setNumberFormat('@');
    sh.getRange(2, 7, 500, 1).setNumberFormat('@')
      .setDataValidation(walletDd_(['SUBMITTED', 'APPROVED', 'REJECTED']));
  }
  return sh;
}

function ensureWalletRules_(ss) {
  var rules = ss.getSheetByName('Platform_Rules');
  if (!rules) return;
  var vals = rules.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'WALLET_ADDRESS') return;
  }
  rules.appendRow(['WALLET_ADDRESS', TEST_ADDRESS_,
    'Typed text. Used only if WALLET_LIVE_ENABLED_ is true in WM_12 and PAYMENT_MODE is LIVE. Leave as the test value for now.']);
}

function setupWallet() {
  var ss = getRegistrySheet_().getParent();
  getWalletSheet_(ss);
  ensureWalletRules_(ss);
  Logger.log('Wallet setup OK (test mode)');
  return 'Wallet setup OK (test mode)';
}

function walletIsLive_(ss) {
  var mode = '';
  try { mode = String(readRules_(ss).PAYMENT_MODE || '').trim().toUpperCase(); } catch (e) {}
  return (WALLET_LIVE_ENABLED_ === true && mode === 'LIVE');
}

function walletIsSandbox_(ss) {
  var mode = 'SANDBOX';
  try { mode = String(readRules_(ss).PAYMENT_MODE || 'SANDBOX').trim().toUpperCase(); } catch (e) {}
  return mode !== 'LIVE';
}

/** Instructions for one of the customer's own pending requests. */
function getWalletInstructions(token, row) {
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    var ss = getRegistrySheet_().getParent();
    ensureWalletRules_(ss);
    row = Math.round(Number(row));
    var sh = getLiveSheet_(ss);
    if (!isFinite(row) || row < 2 || row > sh.getLastRow()) return { success: false, message: 'Request not found.' };
    var r = sh.getRange(row, 1, 1, 9).getValues()[0];
    if (String(r[1]).trim().toLowerCase() !== s.email) return { success: false, message: 'Request not found.' };
    if (String(r[7] || '').trim().toUpperCase() === 'GIFT') return { success: false, message: 'Not a payable request.' };
    if (String(r[6] || 'PENDING').trim().toUpperCase() !== 'PENDING') return { success: false, message: 'This request is not awaiting payment.' };

    var pm = payMap_(ss, s.email)[row];
    if (!pm) return { success: false, message: 'No payment order exists for this request.' };
    if (PAY_FINAL_.indexOf(pm.status) >= 0) {
      return { success: false, message: 'This order is ' + pm.status.toLowerCase() + '. Please start a new plan.' };
    }

    var live = walletIsLive_(ss);
    var address = TEST_ADDRESS_;
    if (live) {
      address = String(readRules_(ss).WALLET_ADDRESS || '').trim();
      if (!/^0x[a-fA-F0-9]{40}$/.test(address)) return { success: false, message: 'Payment details are not available right now.' };
    }
    return {
      success: true, test: !live, address: address, ref: pm.ref,
      asset: 'USDT (ERC-20)', amount: (Number(r[3]) || 0).toFixed(2),
      note: live ? 'Send the exact amount, then submit your transaction hash.'
                 : 'TEST MODE: this is not a real address. Do not send any funds.'
    };
  } catch (err) {
    Logger.log('getWalletInstructions error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  }
}

/** Customer submits a transaction hash. A claim for review, not proof of payment. */
function submitTxid(token, row, txid) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var s = restoreSession(token);
    if (!s.success) return { success: false, message: 'Please sign in again.' };
    txid = String(txid || '').trim();
    if (!/^0x[a-fA-F0-9]{64}$/.test(txid)) return { success: false, message: 'Enter a valid transaction hash (0x followed by 64 characters).' };

    var cache = CacheService.getScriptCache();
    var key = 'TX_' + emailId_(s.email);
    var n = Number(cache.get(key) || 0);
    if (n >= 5) return { success: false, message: 'Too many submissions. Please try again in an hour.' };

    lock.waitLock(15000);
    locked = true;
    var ss = getRegistrySheet_().getParent();
    var ins = getWalletInstructions(token, row);
    if (!ins.success) return ins;
    row = Math.round(Number(row));

    var ws = getWalletSheet_(ss);
    var v = ws.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][5]).trim().toLowerCase() === txid.toLowerCase()) {
        return { success: false, message: 'That transaction hash has already been submitted.' };
      }
      if (Number(v[i][2]) === row && String(v[i][6]).trim().toUpperCase() !== 'REJECTED') {
        return { success: false, message: 'A hash is already under review for this request.' };
      }
    }
    cache.put(key, String(n + 1), 3600);
    ws.appendRow([new Date(), s.email, row, ins.asset, ins.amount, txid, 'SUBMITTED', '', '', ins.test ? 'TEST MODE' : '']);
    return { success: true, message: 'Hash received. Status: under review.' };
  } catch (err) {
    Logger.log('submitTxid error: ' + err);
    return { success: false, message: 'Server error. Please try again.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Admin: list submissions, newest first. */
function adminListWallet(token) {
  try {
    var adm = requireAdmin_(token);
    if (!adm) return { success: false, message: 'Not authorized.' };
    var ss = getRegistrySheet_().getParent();
    var v = getWalletSheet_(ss).getDataRange().getValues();
    var out = [];
    for (var i = v.length - 1; i >= 1 && out.length < 200; i--) {
      out.push({
        row: i + 1, email: String(v[i][1]), source: Number(v[i][2]), asset: String(v[i][3]),
        amount: String(v[i][4]), txid: String(v[i][5]), status: String(v[i][6]),
        date: fmtDate_(v[i][0]), note: String(v[i][9] || '')
      });
    }
    return { success: true, live: walletIsLive_(ss), items: out };
  } catch (err) {
    Logger.log('adminListWallet error: ' + err);
    return { success: false, message: 'Server error.' };
  }
}

/** Core decision. APPROVED feeds processPaymentEvent_ (SANDBOX only). Not exposed to the browser. */
function decideWallet_(adminEmail, subRow, decision, note) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    decision = String(decision || '').toUpperCase();
    if (['APPROVED', 'REJECTED'].indexOf(decision) < 0) return { success: false, message: 'Invalid decision.' };
    note = String(note || '').substring(0, 200);
    if (note.charAt(0) === '=') return { success: false, message: 'Note cannot start with "=".' };

    lock.waitLock(15000);
    locked = true;
    var ss = getRegistrySheet_().getParent();
    var ws = getWalletSheet_(ss);
    subRow = Math.round(Number(subRow));
    if (!isFinite(subRow) || subRow < 2 || subRow > ws.getLastRow()) return { success: false, message: 'Submission not found.' };
    var sub = ws.getRange(subRow, 1, 1, 10).getValues()[0];
    if (String(sub[6]).trim().toUpperCase() !== 'SUBMITTED') return { success: false, message: 'Already decided.' };

    var resultMsg = 'Rejected.';
    if (decision === 'APPROVED') {
      if (walletIsLive_(ss) || !walletIsSandbox_(ss)) {
        return { success: false, message: 'Approval is test-mode only. Live payments need a verified provider event, which is not built.' };
      }
      var email = String(sub[1]).trim().toLowerCase();
      var src = Number(sub[2]);
      var live = getLiveSheet_(ss);
      if (!isFinite(src) || src < 2 || src > live.getLastRow()) return { success: false, message: 'Plan row not found.' };
      if (String(live.getRange(src, 2).getValue()).trim().toLowerCase() !== email) {
        return { success: false, message: 'Plan row does not match the customer.' };
      }
      var pm = payMap_(ss, email)[src];
      if (!pm) return { success: false, message: 'No payment order exists for this plan.' };
      var amount = Number(live.getRange(src, 4).getValue()) || 0;
      var res = processPaymentEvent_(ss, pm.ref, { kind: 'paid', amount: amount, providerRef: String(sub[5]), source: 'ADMIN_TEST_WALLET' });
      if (!res.success || res.duplicate) return { success: false, message: res.message || 'Order could not be processed.' };
      if (res.status !== 'CONFIRMED' && res.status !== 'OVERPAID') return { success: false, message: 'Order ended as ' + res.status + '.' };
      resultMsg = 'Approved (test mode). Plan confirmed.';
    }
    ws.getRange(subRow, 7, 1, 4).setValues([[decision, adminEmail, new Date(), note || String(sub[9] || '')]]);
    logAdmin_(adminEmail, 'WALLET_' + decision, String(sub[1]), 'submission row ' + subRow + ' plan row ' + sub[2]);
    return { success: true, message: resultMsg };
  } catch (err) {
    Logger.log('decideWallet_ error: ' + err);
    return { success: false, message: 'Server error.' };
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Admin (browser): approve or reject a submission. */
function adminDecideWallet(token, subRow, decision, note) {
  var adm = requireAdmin_(token);
  if (!adm) return { success: false, message: 'Not authorized.' };
  return decideWallet_(adm.email, subRow, decision, note);
}

/** Editor test: files a fake submission for the newest non-gift request that has a pending order. */
function walletTestSubmission() {
  var ss = getRegistrySheet_().getParent();
  var live = getLiveSheet_(ss);
  for (var r = live.getLastRow(); r >= 2; r--) {
    var row = live.getRange(r, 1, 1, 9).getValues()[0];
    if (String(row[7] || '').trim().toUpperCase() === 'GIFT') continue;
    if (String(row[6] || 'PENDING').trim().toUpperCase() !== 'PENDING') continue;
    var fake = ('0x' + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '')).substring(0, 66);
    getWalletSheet_(ss).appendRow([new Date(), String(row[1]), r, 'USDT (ERC-20)', (Number(row[3]) || 0).toFixed(2),
      fake, 'SUBMITTED', '', '', 'TEST MODE (editor test)']);
    Logger.log('Test submission added for row ' + r);
    return 'Test submission added';
  }
  return 'No pending plan requests found';
}

/** Editor test: approves the newest SUBMITTED hash through the safe path. */
function walletTestApprove() {
  var ss = getRegistrySheet_().getParent();
  var ws = getWalletSheet_(ss);
  for (var r = ws.getLastRow(); r >= 2; r--) {
    if (String(ws.getRange(r, 7).getValue()).trim().toUpperCase() === 'SUBMITTED') {
      var res = decideWallet_('EDITOR_TEST', r, 'APPROVED', 'Editor test approval');
      Logger.log(JSON.stringify(res));
      return res.message;
    }
  }
  return 'No SUBMITTED hash found';
}
