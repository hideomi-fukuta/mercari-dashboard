function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  const months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  const targets = ['売上','仕入','粗利','目標粗利','目標との差','目標達成率','合計出品数','販売数','仕入数','出品中','平均売価','平均利益','平均仕入れ値','回転率','利益率'];

  const result = { months: months, data: {} };

  data.forEach(function(row) {
    const label = String(row[1]).trim();
    if (targets.indexOf(label) !== -1) {
      result.data[label] = months.map(function(_, i) { return row[i + 2]; });
    }
  });

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
