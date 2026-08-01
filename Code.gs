/**
 * 記帳存摺 — Google Apps Script 後端
 *
 * 使用方式：
 * 1. 建立一個新的 Google 試算表（空白活頁簿即可）。
 * 2. 上方選單「擴充功能」→「Apps Script」。
 * 3. 把這個檔案的全部內容貼進去，取代預設的程式碼，按下儲存（磁片圖示）。
 * 4. 右上角「部署」→「新增部署作業」：
 *      - 類型選「網頁應用程式」
 *      - 說明可填「記帳存摺」
 *      - 執行身分：我
 *      - 誰可以存取：任何人
 *    按「部署」，第一次會要求授權，選擇你的 Google 帳號，
 *    如果出現「Google 尚未驗證這個應用程式」的警告，
 *    點「進階」→「前往（不安全）」即可（因為這是你自己寫、自己部署的程式）。
 * 5. 部署完成後複製「網頁應用程式網址」（結尾是 /exec），
 *    貼到記帳存摺 App 的「設定」欄位裡。
 *
 * 這支程式會自動在試算表裡建立一個叫「記帳明細」的工作表來存放資料，
 * 不需要手動建立欄位。
 */

var SHEET_NAME = '記帳明細';
var HEADERS = ['建立時間', '日期', '收支', '類別', '項目', '金額', '備註', 'ID'];

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
      // 略過完全空白的列
      if (row.join('') === '') continue;
      var obj = {};
      for (var c = 0; c < headers.length; c++) {
        var val = row[c];
        if (val instanceof Date) {
          // 日期欄位轉成 YYYY-MM-DD 字串
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[headers[c]] = val;
      }
      rows.push(obj);
    }
    return jsonOutput_({ ok: true, data: rows });
  }

  return jsonOutput_({ ok: true, message: '記帳存摺 API 運作中' });
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
