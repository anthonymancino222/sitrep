// ============================================================
//  Moquin QA SITREP — Google Apps Script backend
//  Paste this into Code.gs in your Apps Script project
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Moquin QA · SITREP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Entry point for the GitHub-Pages-hosted index.html, which calls this
// via fetch() instead of google.script.run (that API only works when
// Apps Script itself serves the page).
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result = sendSitrep(data);
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendSitrep(data) {
  var to       = 'qualitygroup@moquinpress.com,anthony.mancino@moquinpress.com';
  var subject  = data.subject;
  var textBody = data.body;
  var htmlBody = data.htmlBody;

  GmailApp.sendEmail(to, subject, textBody, {
    htmlBody: htmlBody,
    name: data.submitter + ' · Moquin QA'
  });

  return 'sent';
}
