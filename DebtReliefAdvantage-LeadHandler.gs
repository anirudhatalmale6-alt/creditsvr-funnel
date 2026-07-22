/**
 * Debt Relief Advantage — Lead Handler
 * Saves every lead from debtreliefadvantage.com into this Google Sheet
 * AND emails a notification to you.
 *
 * SETUP (one time):
 *  1) In your Google Sheet: Extensions → Apps Script
 *  2) Delete anything there, paste ALL of this code
 *  3) Change NOTIFY_EMAIL below to the email where you want lead alerts
 *  4) Click Deploy → New deployment → type "Web app"
 *       Execute as:  Me
 *       Who has access:  Anyone
 *  5) Authorize when prompted, then copy the Web app URL (ends in /exec)
 *  6) Send me that /exec URL and I'll connect the form
 */

var NOTIFY_EMAIL = 'PUT-YOUR-EMAIL-HERE@example.com'; // <-- CHANGE THIS

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Leads') || ss.insertSheet('Leads');

    if (sh.getLastRow() === 0) {
      sh.appendRow(['Timestamp', 'First Name', 'Last Name', 'Email', 'Phone', 'State',
                    'Debt Amount', 'Debt Type', 'Monthly Income', 'Behind on Payments',
                    'Employment', 'Source']);
    }

    sh.appendRow([
      new Date(),
      p.first_name || '', p.last_name || '', p.email || '', p.phone || '', p.state || '',
      p.debt_amount || '', p.debt_type || '', p.monthly_income || '', p.behind || '',
      p.employment || '', p.page || 'Debt Relief Advantage'
    ]);

    var subject = p._subject || 'New Debt Relief Advantage Lead';
    var body =
      'New lead from debtreliefadvantage.com\n\n' +
      'Name: ' + (p.first_name || '') + ' ' + (p.last_name || '') + '\n' +
      'Email: ' + (p.email || '') + '\n' +
      'Phone: ' + (p.phone || '') + '\n' +
      'State: ' + (p.state || '') + '\n\n' +
      'Debt Amount: ' + (p.debt_amount || '') + '\n' +
      'Debt Type: ' + (p.debt_type || '') + '\n' +
      'Monthly Income: ' + (p.monthly_income || '') + '\n' +
      'Behind on Payments: ' + (p.behind || '') + '\n' +
      'Employment: ' + (p.employment || '') + '\n\n' +
      'Submitted: ' + new Date();

    if (NOTIFY_EMAIL && NOTIFY_EMAIL.indexOf('@') > 0 && NOTIFY_EMAIL.indexOf('PUT-YOUR') < 0) {
      MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Debt Relief Advantage lead endpoint is running.');
}
