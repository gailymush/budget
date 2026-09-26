/**
 * Expense Passbook — Google Apps Script Backend
 *
 * How to use:
 * 1. Create a new Google Sheet (a blank spreadsheet is fine).
 * 2. From the top menu, go to "Extensions" → "Apps Script".
 * 3. Paste the entire contents of this file, replacing the default code,
 *    then click Save (the disk icon).
 * 4. In the top right, click "Deploy" → "New deployment":
 *      - Select type: "Web app"
 *      - Description: e.g. "Expense Passbook"
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click "Deploy". The first time, you'll be asked to authorize it —
 *    choose your Google account.
 *    If you see a "Google hasn't verified this app" warning,
 *    click "Advanced" → "Go to (unsafe)". This is fine because
 *    it's a script you wrote and deployed yourself.
 * 5. After deployment, copy the "Web app URL" (it ends in /exec)
 *    and paste it into the Settings field of the Expense Passbook app.
 *
 * This script automatically creates a sheet named "Transactions"
 * in the spreadsheet to store your data, so you don't need to set up
 * any columns manually.
 */

var SHEET_NAME = 'Transactions';
var HEADERS = ['Created At', 'Date', 'Type', 'Category', 'Item', 'Amount', 'Note', 'ID'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  var sheet = getSheet_();

  if (action === 'list') {
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      // Skip completely empty rows
      if (row.join('') === '') continue;
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        var val = row[c];
        if (val instanceof Date) {
          // Convert date values to a YYYY-MM-DD string
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[headers[c]] = val;
      }
      rows.push(obj);
    }
    return jsonOutput_({ ok: true, data: rows });
  }

  return jsonOutput_({ ok: true, message: 'Expense Passbook API is running' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      body.date || '',
      body.type || '',
      body.category || '',
      body.item || '',
      Number(body.amount) || 0,
      body.note || '',
      body.id || ''
    ]);
    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
