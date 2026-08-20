// ============================================================
//  Moquin QA SITREP — Google Apps Script backend
//  v4 — with Google Sheets logging + Drive photo saving
// ============================================================

var SHEET_NAME       = 'SITREP Log';
var PHOTO_FOLDER_NAME = 'SITREP Saved Photos'; // Must match exactly in Drive

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

// ---- Find the "SITREP Saved Photos" folder in Drive --------
function getPhotoFolder() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  // If it doesn't exist yet, create it
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

// ---- Get or create the log spreadsheet ---------------------
function getLogSheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId  = props.getProperty('LOG_SPREADSHEET_ID');
  var ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); }
    catch(e) { ssId = null; }
  }

  if (!ssId) {
    ss = SpreadsheetApp.create('Moquin QA — SITREP Log');
    props.setProperty('LOG_SPREADSHEET_ID', ss.getId());
    Logger.log('SITREP Log spreadsheet created: ' + ss.getUrl());
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet().setName(SHEET_NAME);
    var headers = [
      'Timestamp', 'Date', 'Shift', 'Submitted By',
      'Kama', 'New Die Cutter', 'Old Die Cutter',
      'Labels', 'Long Gluer', 'Short Gluer',
      'Priority Message', 'Photos'
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a1a1a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 110);
    sheet.setColumnWidth(3, 80);
    sheet.setColumnWidth(4, 130);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(8, 200);
    sheet.setColumnWidth(9, 200);
    sheet.setColumnWidth(10, 200);
    sheet.setColumnWidth(11, 280);
    sheet.setColumnWidth(12, 300);
  }

  return sheet;
}

// ---- Save photos to Drive and return their links -----------
function savePhotosToDrive(photos, dateStr, shift, submitter) {
  if (!photos || photos.length === 0) return [];

  var folder    = getPhotoFolder();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  // Create a sub-folder per SITREP: e.g. "2026-05-27 1st Anthony M"
  var subFolderName = dateStr.replace(/\//g, '-') + ' ' + shift + ' ' + submitter;
  var subFolder = folder.createFolder(subFolderName);

  var links = [];
  var photoCount = 0;

  for (var i = 0; i < photos.length; i++) {
    var photo = photos[i];
    if (!photo || !photo.base64) continue;

    photoCount++;
    var imageData = Utilities.base64Decode(photo.base64);
    var ext = photo.mimeType === 'image/png' ? '.png' : '.jpg';
    var fileName = 'photo_' + photoCount + '_' + timestamp + ext;
    var blob = Utilities.newBlob(imageData, photo.mimeType, fileName);
    var file = subFolder.createFile(blob);

    // Make the file viewable by anyone with the link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links.push({ name: fileName, url: file.getUrl() });
  }

  return { links: links, folderUrl: subFolder.getUrl() };
}

// ---- Log one SITREP row to the sheet -----------------------
function logToSheet(data, photoResult) {
  var sheet = getLogSheet();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');

  var stationOrder = ['kama', 'newdc', 'olddc', 'labels', 'longgluer', 'shortgluer'];
  var stationCells = [];
  for (var i = 0; i < stationOrder.length; i++) {
    var sid   = stationOrder[i];
    var entry = data.stationData ? data.stationData[sid] : null;
    stationCells.push(entry ? entry : 'Nothing to report');
  }

  // Build photos cell value
  var photosCell = '';
  if (photoResult && photoResult.links && photoResult.links.length > 0) {
    photosCell = photoResult.links.length + ' photo(s) — ' + photoResult.folderUrl;
  }

  var row = [
    timestamp,
    data.reportDate  || '',
    data.shift       || '',
    data.submitter   || '',
    stationCells[0],
    stationCells[1],
    stationCells[2],
    stationCells[3],
    stationCells[4],
    stationCells[5],
    data.priority    || '',
    photosCell
  ];

  sheet.appendRow(row);

  var lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f5f4f0');
  }

  var stationStartCol = 5;
  for (var j = 0; j < stationCells.length; j++) {
    var cell = sheet.getRange(lastRow, stationStartCol + j);
    if (stationCells[j] && stationCells[j] !== 'Nothing to report') {
      cell.setFontColor('#c0392b').setFontWeight('bold');
    } else {
      cell.setFontColor('#555555').setFontWeight('normal');
    }
  }

  // Make photos cell a clickable hyperlink if folder exists
  if (photoResult && photoResult.folderUrl) {
    var photoCell = sheet.getRange(lastRow, 12);
    var linkText  = photoResult.links.length + ' photo(s) — view folder';
    var formula   = '=HYPERLINK("' + photoResult.folderUrl + '","' + linkText + '")';
    photoCell.setFormula(formula).setFontColor('#1a4fa0');
  }
}

// ---- Main send function ------------------------------------
function sendSitrep(data) {
  var to       = 'qualitygroup@moquinpress.com,anthony.mancino@moquinpress.com';
  var subject  = data.subject;
  var textBody = data.body;
  var htmlBody = data.htmlBody;

  // Save photos to Drive first
  var photoResult = null;
  if (data.photos && data.photos.length > 0) {
    try {
      photoResult = savePhotosToDrive(data.photos, data.reportDate, data.shift, data.submitter);
    } catch(e) {
      Logger.log('Photo save error: ' + e.toString());
    }
  }

  // Build email options
  var options = {
    htmlBody: htmlBody,
    name: data.submitter + ' · Moquin QA'
  };

  // Still attach photos to email as well
  if (data.photos && data.photos.length > 0) {
    var attachments = [];
    for (var i = 0; i < data.photos.length; i++) {
      var photo = data.photos[i];
      if (!photo) continue;
      var imageData = Utilities.base64Decode(photo.base64);
      var blob = Utilities.newBlob(imageData, photo.mimeType, photo.name);
      attachments.push(blob);
    }
    if (attachments.length > 0) options.attachments = attachments;
  }

  // Add Drive folder link to email if photos were saved
  if (photoResult && photoResult.folderUrl) {
    options.htmlBody = htmlBody.replace(
      '</div>',
      '<p style="margin:16px 0 0 0;font-family:Arial,sans-serif;font-size:13px;color:#1a4fa0;">'
      + '&#128247; <a href="' + photoResult.folderUrl + '">View saved photos in Google Drive</a>'
      + '</p></div>'
    );
  }

  // Send email
  GmailApp.sendEmail(to, subject, textBody, options);

  // Log to sheet
  try {
    logToSheet(data, photoResult);
  } catch(e) {
    Logger.log('Sheet logging error: ' + e.toString());
  }

  return 'sent';
}

// ---- Utility: get the log sheet URL ------------------------
function getLogUrl() {
  var props = PropertiesService.getScriptProperties();
  var ssId  = props.getProperty('LOG_SPREADSHEET_ID');
  if (!ssId) return 'Log not created yet — send a SITREP first.';
  return SpreadsheetApp.openById(ssId).getUrl();
}

// ─── SITREP REMINDER SYSTEM ───────────────────

var SITREP_FORM_URL = "https://script.google.com/macros/s/AKfycbx1evCwmp-0CacMkkOTulSoVwFFg4wsSoDxhoyZSGnA2xG_iDGBpQRjmcSdvylJOzk/exec";

function sendSitrepReminder() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('LOG_SPREADSHEET_ID'));
  var sheet = ss.getSheetByName('SITREP Log');
  var data = sheet.getDataRange().getValues();

  var today = new Date();
  var dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return; // Skip weekends

  var todayStr = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();

  var firstShiftSubmitted = false;
  var secondShiftSubmitted = false;

  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][1]); // Column B = Date
    var rowDateStr = (rowDate.getMonth() + 1) + "/" + rowDate.getDate() + "/" + rowDate.getFullYear();
    var rowShift = data[i][2]; // Column C = Shift

    if (rowDateStr === todayStr) {
      if (rowShift === '1st') firstShiftSubmitted = true;
      if (rowShift === '2nd') secondShiftSubmitted = true;
    }
  }

  var missingShifts = [];
  if (!firstShiftSubmitted) missingShifts.push('1st Shift');
  if (!secondShiftSubmitted) missingShifts.push('2nd Shift');

  if (missingShifts.length > 0) {
    MailApp.sendEmail({
      to: "qualitygroup@moquinpress.com, anthony.mancino@moquinpress.com",
      subject: "⚠️ SITREP Missing – " + todayStr,
      body: "Hello Team,\n\nThe following shift(s) have not submitted a SITREP today (" + todayStr + "):\n\n• " + missingShifts.join('\n• ') + "\n\nPlease submit your report.\n\nThank you."
    });
  }
}
