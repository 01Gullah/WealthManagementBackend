/**
 * WM_07_Legal.gs
 * Apex Wealth - Terms of Service and Privacy Policy, editable from the Sheet.
 * Tab: Legal_Pages (Page | No | Type | Text | Active)
 * STARTER TEXT ONLY. Must be reviewed and replaced by a qualified attorney before launch.
 */

var LEGAL_TAB = 'Legal_Pages';

var DEFAULT_LEGAL_ = [
  ['TERMS', 1, 'HEADING', 'Acceptance of terms', 'TRUE'],
  ['TERMS', 2, 'PARAGRAPH', 'By creating an account or using Apex Wealth you agree to these Terms of Service. If you do not agree, do not use the platform.', 'TRUE'],
  ['TERMS', 3, 'HEADING', 'Eligibility', 'TRUE'],
  ['TERMS', 4, 'PARAGRAPH', 'You must meet the minimum age shown on the platform and be legally able to enter into binding agreements where you live. Availability may be limited by location.', 'TRUE'],
  ['TERMS', 5, 'HEADING', 'Your account', 'TRUE'],
  ['TERMS', 6, 'PARAGRAPH', 'You are responsible for keeping your password secure and for all activity on your account. Tell us immediately if you suspect unauthorized use.', 'TRUE'],
  ['TERMS', 7, 'HEADING', 'Risk and no guarantees', 'TRUE'],
  ['TERMS', 8, 'PARAGRAPH', 'All investing involves risk, including loss of some or all of your principal. Projections and illustrative rates are hypothetical, are not guarantees, and past performance does not predict future results.', 'TRUE'],
  ['TERMS', 9, 'HEADING', 'No advice', 'TRUE'],
  ['TERMS', 10, 'PARAGRAPH', 'Nothing on this platform is financial, investment, tax or legal advice. Consider speaking with a qualified professional before making decisions.', 'TRUE'],
  ['TERMS', 11, 'HEADING', 'Plans and withdrawals', 'TRUE'],
  ['TERMS', 12, 'PARAGRAPH', "Withdrawing before a plan's maturity date returns your principal only. Returns, if any, are paid only at maturity. Terms for each plan are shown before you confirm it.", 'TRUE'],
  ['TERMS', 13, 'HEADING', 'Lawful use', 'TRUE'],
  ['TERMS', 14, 'PARAGRAPH', 'You confirm that any funds you use are lawfully yours. We may request identity verification and may suspend or close accounts to meet legal obligations.', 'TRUE'],
  ['TERMS', 15, 'HEADING', 'Changes and contact', 'TRUE'],
  ['TERMS', 16, 'PARAGRAPH', 'We may update these terms. Continued use after an update means you accept the new version. Contact support through the details provided on the platform.', 'TRUE'],

  ['PRIVACY', 1, 'HEADING', 'What we collect', 'TRUE'],
  ['PRIVACY', 2, 'PARAGRAPH', 'We collect your email address, a hashed form of your password (we never store your password in readable form), your plan requests, and records of the consents you give.', 'TRUE'],
  ['PRIVACY', 3, 'HEADING', 'How we use it', 'TRUE'],
  ['PRIVACY', 4, 'PARAGRAPH', 'We use this information to operate your account, show your plans and dashboard, keep records of your consents, secure the platform, and meet legal obligations.', 'TRUE'],
  ['PRIVACY', 5, 'HEADING', 'Where it is stored', 'TRUE'],
  ['PRIVACY', 6, 'PARAGRAPH', 'Data is stored using Google services. Access is limited to authorized administrators.', 'TRUE'],
  ['PRIVACY', 7, 'HEADING', 'Sharing', 'TRUE'],
  ['PRIVACY', 8, 'PARAGRAPH', 'We do not sell your personal information. We may share information with service providers who help us operate the platform, or where the law requires it.', 'TRUE'],
  ['PRIVACY', 9, 'HEADING', 'Your choices', 'TRUE'],
  ['PRIVACY', 10, 'PARAGRAPH', 'You may ask to access, correct or delete your information, subject to records we must keep by law. Contact support to make a request.', 'TRUE'],
  ['PRIVACY', 11, 'HEADING', 'Security and changes', 'TRUE'],
  ['PRIVACY', 12, 'PARAGRAPH', 'We use reasonable safeguards but no system is completely secure. We may update this policy and will show the current version here.', 'TRUE']
];

function ddlg_(list) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .setHelpText('Pick a value from the list.')
    .build();
}

function getLegalSheet_(ss) {
  var sh = ss.getSheetByName(LEGAL_TAB);
  if (!sh) {
    sh = ss.insertSheet(LEGAL_TAB);
    sh.appendRow(['Page', 'No', 'Type', 'Text', 'Active']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 100);
    sh.setColumnWidth(2, 60);
    sh.setColumnWidth(3, 120);
    sh.setColumnWidth(4, 560);
    sh.setColumnWidth(5, 80);
    sh.getRange(2, 4, 200, 1).setNumberFormat('@').setWrap(true);
    sh.getRange(2, 5, 200, 1).setNumberFormat('@');
    var rows = DEFAULT_LEGAL_.map(function (r) { return [r[0], r[1], r[2], r[3], r[4]]; });
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  sh.getRange(2, 1, 199, 1).setDataValidation(ddlg_(['TERMS', 'PRIVACY']));
  sh.getRange(2, 2, 199, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(1, 99).setAllowInvalid(false).build());
  sh.getRange(2, 3, 199, 1).setDataValidation(ddlg_(['HEADING', 'PARAGRAPH']));
  sh.getRange(2, 5, 199, 1).setDataValidation(ddlg_(['TRUE', 'FALSE']));
  return sh;
}

function setupLegal() {
  var ss = getRegistrySheet_().getParent();
  getLegalSheet_(ss);
  Logger.log('Legal setup OK');
  return 'Legal setup OK';
}

/** Public: returns the active items of one page (TERMS or PRIVACY), in order. */
function getLegalPage(page) {
  try {
    page = String(page || '').toUpperCase();
    if (page !== 'TERMS' && page !== 'PRIVACY') return { success: false, message: 'Page not found.' };
    var ss = getRegistrySheet_().getParent();
    var vals = getLegalSheet_(ss).getDataRange().getValues();
    var items = [];
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toUpperCase() !== page) continue;
      if (String(vals[i][4]).trim().toUpperCase() !== 'TRUE') continue;
      var no = Number(vals[i][1]);
      var text = String(vals[i][3]).trim();
      if (!text || !isFinite(no)) continue;
      items.push({ no: no, type: String(vals[i][2]).trim().toUpperCase() === 'HEADING' ? 'HEADING' : 'PARAGRAPH', text: text });
    }
    items.sort(function (a, b) { return a.no - b.no; });
    if (!items.length) return { success: false, message: 'This page has no content yet.' };
    return { success: true, title: page === 'TERMS' ? 'Terms of Service' : 'Privacy Policy', items: items };
  } catch (err) {
    Logger.log('getLegalPage error: ' + err);
    return { success: false, message: 'Could not load this page. Try again.' };
  }
}
