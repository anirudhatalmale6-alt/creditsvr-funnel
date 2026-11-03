/**
 * Debt Relief Advantage — lead handler + TurboDebt "Host & Post"
 *
 * The web page sends every completed form here. This script then, in order:
 *
 *   1. SAVES the lead to the "Leads" tab, qualified or not. TurboDebt's IO
 *      requires consent records, IP logs, timestamps, source URLs and
 *      TrustedForm certificates to be kept for 5 years and produced within
 *      5 business days of a request. This sheet is that record.
 *   2. DECIDES whether the lead meets TurboDebt's paid-lead terms (IO,
 *      "What Constitutes Gross Lead"). The page does not decide this; it
 *      could be tampered with.
 *   3. POSTS qualifying leads to TurboDebt and returns their live-chat link
 *      to the page, which sends the visitor straight there. Their spec says
 *      the visitor MUST land on that exact link, unaltered.
 *
 * ---------------------------------------------------------------- SETUP
 *   a) Project Settings (gear icon) > Script properties > Add script property
 *        Property:  TURBODEBT_TOKEN
 *        Value:     the partner token TurboDebt gives you
 *      Keep the token here and nowhere else. A token on the public web page
 *      could be copied by anyone and used to post junk leads in your name,
 *      which TurboDebt would charge back.
 *      Until it is set, leads are still saved but nothing goes to TurboDebt.
 *   b) Optional: put your email in NOTIFY_EMAIL below for a copy of each lead.
 *   c) Deploy > New deployment > type "Web app"
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Copy the /exec link and send it over.
 *   After any change to this file: Ctrl+S, then Deploy > Manage deployments >
 *   edit > Version: New version. Saving alone does not update the live app.
 */

var SHEET_NAME    = 'Leads';
var NOTIFY_EMAIL  = '';   // optional, e.g. 'you@example.com'
var LEAD_SECRET   = 'dra-6Qm9VbT2xKp4LwZ8';   // must match the web page
var TURBODEBT_URL = 'https://www.acquisitionbrands.com/atc/lead/';

/* ------------------------------------------------ TurboDebt's lead terms */

// IO Special Instructions: 44 states, DC, and 3 territories. Excluded:
// CT, WI, MN, OR, VT, WV. TurboDebt lists the Virgin Islands as "USVI", but
// their state field takes 2-letter codes, so it is sent as VI.
var ELIGIBLE_STATES = ['AL','AK','AZ','AR','CA','CO','DC','DE','FL','GA','GU','HI',
  'ID','IL','IN','IA','KS','KY','LA','MD','MA','ME','MI','MS','MO','MT','NE','NH',
  'NV','NJ','NM','NY','NC','ND','OH','OK','PA','PR','RI','SC','SD','TN','TX','VI',
  'UT','VA','WA','WY'];

// IO: "more than $10K in credit card or personal loan debt".
var QUALIFYING_DEBT_TYPES = ['Credit Cards', 'Personal Loans'];
var MIN_DEBT_EXCLUSIVE    = 10000;

// IO: individual (not household) income "more than $1,500 per month". The page's
// bands are split exactly at that line so every answer is unambiguous.
var QUALIFYING_INCOME = ['$5,000+', '$4,000 - $5,000', '$3,000 - $4,000',
                         '$2,000 - $3,000', '$1,501 - $2,000'];

var COLUMNS = ['Timestamp', 'Lead ID', 'First Name', 'Last Name', 'Email', 'Phone',
  'State', 'Debt Amount', 'Debt Type', 'Monthly Income', 'Behind on Payments',
  'Employment', 'Qualified', 'Not Qualified Because', 'TurboDebt Status',
  'TurboDebt Message', 'TurboDebt Chat Link', 'IP Address', 'TrustedForm Certificate',
  'Consent Text', 'sub1', 'Page URL', 'Browser'];

/* ------------------------------------------------------------- web app */

function doPost(e) {
  var raw = (e && e.parameter) || {};
  if (String(raw.secret || '') !== LEAD_SECRET) return json_({ ok: false, error: 'rejected' });

  var p = normalise_(raw);
  var reasons = assess(p);
  var qualified = reasons.length === 0;
  var leadId = Utilities.getUuid();
  var token = PropertiesService.getScriptProperties().getProperty('TURBODEBT_TOKEN') || '';

  // Save FIRST, so the record exists even if the post to TurboDebt hangs or fails.
  var sheet = sheet_();
  var row;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet.appendRow([new Date(), leadId, p.firstname, p.lastname, p.email, p.phone,
      p.state, p.debt_amount, p.debt_type, p.monthly_income, p.behind, p.employment,
      qualified ? 'Yes' : 'No', reasons.join(', '),
      qualified ? (token ? 'sending' : 'not sent - no token yet') : 'not sent',
      '', '', p.ip, p.trustedform, p.tcpa, p.sub1, p.page_url, p.user_agent]);
    row = sheet.getLastRow();
  } finally {
    lock.releaseLock();
  }

  var td = { status: '', message: '', redirect_url: '' };
  if (qualified && token) {
    td = postToTurboDebt(p, leadId, token);
    // Columns O, P, Q: status, message, chat link.
    sheet.getRange(row, 15, 1, 3).setValues([[td.status, td.message, td.redirect_url]]);
  }

  notify_(p, qualified, reasons, td);

  return json_({
    ok: true,
    qualified: qualified,
    status: td.status,
    // Passed through exactly as TurboDebt sent it. Their spec: "PLEASE DO NOT
    // ALTER OUR POST RESPONSE REDIRECT URLS VALUE IN ANY WAY".
    redirect_url: td.redirect_url
  });
}

/* Returns the list of reasons a lead does NOT meet TurboDebt's terms. Empty = qualified. */
function assess(p) {
  var reasons = [];
  if (QUALIFYING_DEBT_TYPES.indexOf(p.debt_type) < 0) reasons.push('debt type');
  if (!(p.debt_amount > MIN_DEBT_EXCLUSIVE)) reasons.push('debt under $10k');
  if (QUALIFYING_INCOME.indexOf(p.monthly_income) < 0) reasons.push('income');
  if (ELIGIBLE_STATES.indexOf(p.state) < 0) reasons.push('state');
  if (!p.firstname || !p.lastname || !p.email) reasons.push('name or email missing');
  if (!/^\d{10}$/.test(p.phone)) reasons.push('phone');
  // Required by TurboDebt's spec, so a lead without them cannot be posted.
  if (!p.ip) reasons.push('no IP');
  if (!p.trustedform) reasons.push('no TrustedForm');
  return reasons;
}

function postToTurboDebt(p, leadId, token) {
  var payload = {
    token: token,
    firstname: p.firstname,
    lastname: p.lastname,
    email: p.email,
    phone: p.phone,
    debt_amount: String(p.debt_amount),
    state: p.state,
    ip: p.ip,
    unique_id: leadId,
    tcpa: p.tcpa,
    trustedform: p.trustedform,
    sub1: p.sub1,
    sub2: 'posts',   // their spec: hard-coded delivery channel for host & post
    sourceid: p.sub1.split(/[_-]/)[0]
  };
  try {
    var r = UrlFetchApp.fetch(TURBODEBT_URL, {
      method: 'post', payload: payload, muteHttpExceptions: true, followRedirects: true
    });
    var code = r.getResponseCode();
    var body = r.getContentText();
    var j = {};
    try { j = JSON.parse(body); } catch (ignore) {}
    return {
      status: String(j.status || ('http ' + code)),
      message: String(j.message || (j.status ? '' : body.slice(0, 300))),
      redirect_url: String(j.redirect_url || '')
    };
  } catch (err) {
    return { status: 'post failed', message: String(err).slice(0, 300), redirect_url: '' };
  }
}

/* --------------------------------------------------------------- helpers */

function normalise_(r) {
  var phone = String(r.phone || '').replace(/\D/g, '');
  if (phone.length === 11 && phone.charAt(0) === '1') phone = phone.slice(1);
  var sub1 = String(r.sub1 || '').replace(/[^A-Za-z0-9._-]/g, '').slice(0, 80) || 'dra_direct';
  return {
    firstname: clip_(r.firstname, 60),
    lastname: clip_(r.lastname, 60),
    email: clip_(r.email, 120).toLowerCase(),
    phone: phone,
    state: String(r.state || '').trim().toUpperCase(),
    debt_amount: parseInt(String(r.debt_amount || '').replace(/\D/g, ''), 10) || 0,
    debt_type: clip_(r.debt_type, 40),
    monthly_income: clip_(r.monthly_income, 40),
    behind: clip_(r.behind, 60),
    employment: clip_(r.employment, 60),
    ip: clip_(r.ip, 45),
    tcpa: clip_(r.tcpa, 3000),
    trustedform: clip_(r.trustedform, 300),
    sub1: sub1,
    page_url: clip_(r.page_url, 500),
    user_agent: clip_(r.user_agent, 300)
  };
}

function clip_(v, n) { return String(v || '').trim().slice(0, n); }

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function notify_(p, qualified, reasons, td) {
  if (!NOTIFY_EMAIL) return;
  var lines = [
    'Name: ' + p.firstname + ' ' + p.lastname,
    'Email: ' + p.email,
    'Phone: ' + p.phone,
    'State: ' + p.state,
    'Debt: $' + p.debt_amount + ' (' + p.debt_type + ')',
    'Income: ' + p.monthly_income,
    'Behind: ' + p.behind,
    'Employment: ' + p.employment,
    '',
    'Qualified for TurboDebt: ' + (qualified ? 'YES' : 'NO - ' + reasons.join(', ')),
    'TurboDebt result: ' + (td.status || 'not sent') + (td.message ? ' - ' + td.message : ''),
    'Source: ' + p.sub1
  ];
  MailApp.sendEmail(NOTIFY_EMAIL,
    (qualified ? 'Qualified lead' : 'Lead (not qualified)') + ' - Debt Relief Advantage',
    lines.join('\n'));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
