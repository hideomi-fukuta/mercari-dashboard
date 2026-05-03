function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  var targets = ['売上','仕入','粗利','目標粗利','目標との差','目標達成率','合計出品数','販売数','仕入数','出品中','平均売価','平均利益','平均仕入れ値','回転率','利益率'];
  var result = { months: months, data: {} };
  data.forEach(function(row) {
    var label = String(row[1]).trim();
    if (targets.indexOf(label) !== -1) {
      result.data[label] = months.map(function(_, i) { return row[i + 2]; });
    }
  });
  var output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    var targetRow = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1]).trim() === params.metric) {
        targetRow = i;
        break;
      }
    }

    if (targetRow === -1) {
      return ContentService.createTextOutput(JSON.stringify({ error: '項目が見つかりません' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 月列: col C(index 2) = 1月, monthIdx=0 → column 3 (1-based)
    var colNum = params.monthIdx + 3;
    sheet.getRange(targetRow + 1, colNum).setValue(Number(params.value));

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
