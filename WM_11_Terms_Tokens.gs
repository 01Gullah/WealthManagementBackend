/**
 * WM_11_Terms_Tokens.gs
 * Apex Wealth - live values for {FEE} and {FORFEIT} in consent and Terms text.
 * Reads WITHDRAW_FEE_PERCENT and WITHDRAW_FORFEIT_RETURNS from Platform_Rules.
 */

function getTermTokens() {
  try {
    var ss = getRegistrySheet_().getParent();
    var rl = readRules_(ss);
    var fee = Math.max(0, Math.min(50, num_(rl.WITHDRAW_FEE_PERCENT, 5)));
    var forfeit = (rl.WITHDRAW_FORFEIT_RETURNS === undefined) ? true : bool_(rl.WITHDRAW_FORFEIT_RETURNS);
    return {
      success: true,
      fee: String(fee),
      forfeit: forfeit ? 'you lose the returns earned so far' : 'you keep the returns earned so far'
    };
  } catch (err) {
    Logger.log('getTermTokens error: ' + err);
    return { success: false };
  }
}