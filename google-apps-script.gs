/**
 * MET-Ai Solutions — Contact Form Backend (Google Apps Script)
 * ──────────────────────────────────────────────────────────────
 * This script receives form submissions from met-ai-solutions.html
 * and appends them as a new row in this Google Sheet.
 * Optionally, it also emails Eeswar a notification for every new lead.
 *
 * SETUP
 * 1. Open your Google Sheet.
 * 2. Row 1 should have these exact headers (in this order):
 *    Timestamp | First Name | Last Name | Email | Phone | Organization | Interest | Message
 * 3. Extensions -> Apps Script. Delete any starter code, paste this file's
 *    contents in, and save.
 * 4. Deploy -> New deployment -> type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    Deploy, authorize, and copy the Web app URL (it ends in /exec).
 * 5. Paste that URL into GOOGLE_SHEETS_URL in met-ai-solutions.html
 *    (search for that variable name in the file).
 *
 * NOTES
 * - Every redeploy after editing this script requires "Manage deployments"
 *   -> edit (pencil icon) -> New version -> Deploy, or the old /exec URL
 *   will keep running the old code.
 * - EMAIL_NOTIFY below is optional. Leave it blank ('') to skip emailing
 *   and only log to the Sheet.
 */

var EMAIL_NOTIFY = 'admin@metaisolutions.com'; // set to '' to disable email alerts

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.ts       || new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      data.fname    || '',
      data.lname    || '',
      data.email    || '',
      data.phone    || '',
      data.org      || '',
      data.interest || '',
      data.msg      || ''
    ]);

    if (EMAIL_NOTIFY) {
      var subject = 'New MET-Ai enquiry - ' + (data.fname || '') + ' ' + (data.lname || '');
      var body =
        'A new contact form submission has been received on the MET-Ai website.\n\n' +
        'Name: '          + (data.fname || '') + ' ' + (data.lname || '') + '\n' +
        'Email: '         + (data.email || '') + '\n' +
        'Phone: '         + (data.phone || '-') + '\n' +
        'Organisation: '  + (data.org || '-') + '\n' +
        'Interested in: ' + (data.interest || '-') + '\n' +
        'Submitted: '     + (data.ts || '') + '\n\n' +
        'Message:\n' + (data.msg || '') + '\n\n' +
        '(Logged automatically to the MET-Ai leads spreadsheet.)';
      MailApp.sendEmail({
        to: EMAIL_NOTIFY,
        replyTo: data.email || EMAIL_NOTIFY,
        subject: subject,
        body: body
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Optional: run this once manually from the Apps Script editor to
 *  confirm the script has edit access to the sheet before deploying. */
function testWrite() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(['TEST', 'Test', 'User', 'test@example.com', '', '', '', 'Setup check']);
}
