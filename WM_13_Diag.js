/**
 * WM_13_Diag.gs
 * One-time diagnostic. Safe to delete afterwards. Does not change any data (it opens and closes a session).
 */
function diag() {
  var out = [];
  function step(name, fn) {
    try {
      var r = fn();
      var s = (typeof r === 'string') ? r : JSON.stringify(r);
      out.push(name + ': ' + String(s).substring(0, 250));
    } catch (e) {
      out.push(name + ': ERROR ' + e);
    }
  }
  step('getPlatformConfig', function () { var c = getPlatformConfig(); return { success: c.success, products: c.products ? c.products.length : 0 }; });
  step('getPaymentOptions', function () { return getPaymentOptions(); });
  step('getTermTokens', function () { return getTermTokens(); });
  step('getConsentContent', function () { var c = getConsentContent(); return { success: c.success, msg: c.message, checks: c.checks ? c.checks.length : 0 }; });

  var email = '';
  var sheet = getRegistrySheet_();
  for (var i = 0; i < ADMIN_LIST_.length; i++) {
    if (findUserRow_(sheet, ADMIN_LIST_[i])) { email = ADMIN_LIST_[i]; break; }
  }
  if (!email) { out.push('No admin email is registered as a user.'); }
  else {
    var tok = createSession_(email);
    step('getAccount for ' + email, function () {
      var a = getAccount(tok);
      return { success: a.success, msg: a.message, unlocked: a.unlocked, plans: a.plans ? a.plans.length : 0, requests: a.requests ? a.requests.length : 0 };
    });
    logoutUser(tok);
  }
  Logger.log(out.join('\n'));
  return out.join('\n');
}
