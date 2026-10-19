/**
 * login-log.gs — Google Apps Script backend for the dashboard's login log.
 *
 * Setup (once):
 *   1. Create a new Google Sheet (keep it private), e.g. "Dashboard Login Log".
 *   2. Extensions → Apps Script, paste this file, set CLIENT_ID below, save.
 *   3. Deploy → New deployment → Web app
 *        Execute as: Me      Who has access: Anyone
 *   4. Copy the /exec URL into config.js → AUTH.LOGIN_LOG_URL.
 *
 * Every request must carry a fresh Google ID token from the dashboard. The token is
 * verified with Google, so entries can't be forged and the log can't be read by
 * anyone outside ALLOWED_DOMAIN even though the web app URL is public.
 */

const CLIENT_ID = '7266197012-iog9olr657n7p8b9v30vo285lnlprebg.apps.googleusercontent.com';
const ALLOWED_DOMAIN = 'whatap.io';
const SHEET_NAME = 'LOG';
const LIST_LIMIT = 500;   // most recent rows returned to the dashboard

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const user = verifyIdToken_(body.idToken);
    if (!user) return json_({ ok: false, error: 'unauthorized' });

    const sheet = getSheet_();
    if (body.action === 'log') {
      sheet.appendRow([new Date(), user.email, user.name, String(body.userAgent || '').slice(0, 300)]);
      return json_({ ok: true });
    }
    if (body.action === 'list') {
      const last = sheet.getLastRow();
      if (last < 2) return json_({ ok: true, entries: [] });
      const start = Math.max(2, last - LIST_LIMIT + 1);
      const rows = sheet.getRange(start, 1, last - start + 1, 3).getValues();
      const entries = rows.reverse().map(r => ({
        time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
        email: r[1],
        name: r[2],
      }));
      return json_({ ok: true, entries });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Ask Google to validate the ID token; returns {email, name} only for allowed accounts. */
function verifyIdToken_(idToken) {
  if (!idToken) return null;
  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const c = JSON.parse(res.getContentText());
  const email = String(c.email || '').toLowerCase();
  const ok = c.aud === CLIENT_ID
    && String(c.email_verified) === 'true'
    && c.hd === ALLOWED_DOMAIN
    && email.endsWith('@' + ALLOWED_DOMAIN);
  return ok ? { email: email, name: c.name || '' } : null;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Time', 'Email', 'Name', 'User Agent']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
