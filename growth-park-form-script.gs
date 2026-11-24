// Growth Park forms
// Saves every submission to this Google Sheet (one tab per form) and emails you.

const NOTIFY_EMAIL = 'info.growthpark@gmail.com';

const TABS = {
  enquiry:  ['Date', 'Name', 'Email', 'Business / brand', 'Website or social link', 'Service', 'Biggest challenge', 'Hoping to achieve', 'Monthly budget', 'Anything else'],
  playbook: ['Date', 'First name', 'Last name', 'Email', 'Looking for support', 'Annual revenue', 'Social media situation'],
  contact:  ['Date', 'Name', 'Email', 'Message']
};

const FIELDS = {
  enquiry:  ['name', 'email', 'business', 'link', 'service', 'challenge', 'goal', 'budget', 'more'],
  playbook: ['first', 'last', 'email', 'q1', 'q2', 'q3'],
  contact:  ['name', 'email', 'message']
};

const TITLES = {
  enquiry:  'New enquiry (Work with us)',
  playbook: 'New free playbook signup',
  contact:  'New contact message'
};

function clean_(v) {
  v = String(v == null ? '' : v).slice(0, 3000);
  return /^[=+\-@]/.test(v) ? "'" + v : v; // stops spreadsheet formulas
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const src = d.source;
    if (!FIELDS[src]) throw new Error('Unknown form');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(src);
    if (!sheet) {
      sheet = ss.insertSheet(src);
      sheet.appendRow(TABS[src]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date()].concat(FIELDS[src].map(function (k) { return clean_(d[k]); })));

    const body = FIELDS[src].map(function (k, i) {
      return TABS[src][i + 1] + ': ' + String(d[k] || '');
    }).join('\n');
    const mail = {
      to: NOTIFY_EMAIL,
      subject: TITLES[src] + ': ' + String(d.name || d.first || d.email || ''),
      body: body
    };
    if (d.email) mail.replyTo = String(d.email);
    MailApp.sendEmail(mail);

    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return out_({ ok: true, status: 'Growth Park form script is running' });
}
