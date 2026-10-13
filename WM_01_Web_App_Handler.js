/**
 * WM_01_Web_App_Handler.gs
 * Apex Wealth - Web App entry point
 * Normal:  .../exec          -> serves the portal
 * Admin:   .../exec?admin=1  -> serves the admin page (server checks admin rights on every call)
 * Debug:   .../exec?debug=1  -> serves a visible test page
 */

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.debug === '1') {
    return buildDebugPage_();
  }

  try {
    if (params.admin === '1') {
      return HtmlService
        .createHtmlOutputFromFile('WM_09_Admin')
        .setTitle('Apex Admin')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService
      .createHtmlOutputFromFile('WM_02_Portal')
      .setTitle(CONFIG.PROJECT_NAME + ' | Wealth & Asset Management')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return buildErrorPage_(err);
  }
}

function buildDebugPage_() {
  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>body{margin:0;padding:24px;background:#0b0f19;color:#fff;' +
    'font-family:-apple-system,Helvetica,Arial,sans-serif}' +
    'h1{color:#10b981}.box{background:#111827;border:1px solid #10b981;' +
    'border-radius:12px;padding:16px;margin-top:16px}</style></head><body>' +
    '<h1>Apex Wealth: Debug OK</h1>' +
    '<div class="box">Deployment is working.<br>Version: ' + CONFIG.VERSION + '</div>' +
    '<div class="box" id="js">JavaScript: checking...</div>' +
    '<script>document.getElementById("js").innerHTML="JavaScript: running OK";</script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Apex Wealth Debug')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function buildErrorPage_(err) {
  var msg = String(err && err.message ? err.message : err)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;');
  var html =
    '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>body{margin:0;padding:24px;background:#0b0f19;color:#fff;' +
    'font-family:-apple-system,Helvetica,Arial,sans-serif}' +
    'h1{color:#ef4444}pre{white-space:pre-wrap;background:#111827;' +
    'padding:16px;border-radius:12px;color:#fca5a5}</style></head><body>' +
    '<h1>Apex Wealth: Load Error</h1><pre>' + msg + '</pre></body></html>';
  return HtmlService.createHtmlOutput(html)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

