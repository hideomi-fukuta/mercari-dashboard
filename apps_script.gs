function doGet(e) {
  var type = e && e.parameter && e.parameter.type ? e.parameter.type : 'summary';
  if (type === 'products') { return getProducts(); }
  return getSummary();
}

function getSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('全体数値表') || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
  var targets = ['売上','仕入','粗利','目標粗利','目標との差','目標達成率','合計出品数','販売数','仕入数','出品中','平均売価','平均利益','平均仕入れ値','回転率','利益率'];
  var result = { months: months, data: {} };
  for (var r = 0; r < data.length; r++) {
    var label = String(data[r][1]).trim();
    if (targets.indexOf(label) !== -1) {
      result.data[label] = months.map(function(_, i) { return data[r][i + 2]; });
    }
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function findHeaderRow(data) {
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).trim() === '商品名' || String(data[i][j]).trim() === '管理番号') return i;
    }
  }
  return -1;
}

function getProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('商品管理表');
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ error: '商品管理表が見つかりません' })).setMimeType(ContentService.MimeType.JSON);
  var data = sheet.getDataRange().getValues();
  var headerRowIdx = findHeaderRow(data);
  if (headerRowIdx === -1) return ContentService.createTextOutput(JSON.stringify({ error: 'ヘッダー行が見つかりません' })).setMimeType(ContentService.MimeType.JSON);
  var headers = data[headerRowIdx].map(function(h) { return String(h).trim(); });
  var nameIdx = headers.indexOf('商品名');
  var numIdx  = headers.indexOf('管理番号');
  var products = [];
  for (var i = headerRowIdx + 1; i < data.length; i++) {
    if (!data[i][nameIdx] && !data[i][numIdx]) continue;
    var product = { _row: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        var v = data[i][j];
        product[headers[j]] = (v instanceof Date) ? v.getFullYear()+'/'+(v.getMonth()+1)+'/'+v.getDate() : v;
      }
    }
    products.push(product);
  }
  return ContentService.createTextOutput(JSON.stringify({ headers: headers, products: products })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (params.action === 'addProduct') {
      var addSheet = ss.getSheetByName('商品管理表');
      var addData  = addSheet.getDataRange().getValues();
      var addHdrIdx = findHeaderRow(addData);
      var addHeaders = addData[addHdrIdx].map(function(h) { return String(h).trim(); });
      var addNumIdx = addHeaders.indexOf('管理番号');
      var maxNum = 0;
      for (var i = addHdrIdx + 1; i < addData.length; i++) {
        var n = Number(addData[i][addNumIdx]);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      }
      var newRow = new Array(addHeaders.length).fill('');
      var fieldMap = params.fields;
      Object.keys(fieldMap).forEach(function(key) {
        var idx = addHeaders.indexOf(key);
        if (idx !== -1) newRow[idx] = fieldMap[key];
      });
      if (addNumIdx !== -1) newRow[addNumIdx] = maxNum + 1;
      addSheet.appendRow(newRow);

    } else if (params.action === 'updateProduct') {
      var updSheet  = ss.getSheetByName('商品管理表');
      var updData   = updSheet.getDataRange().getValues();
      var updHeaders = updData[params.headerRow - 1].map(function(h) { return String(h).trim(); });
      var colIdx = updHeaders.indexOf(params.field);
      if (colIdx === -1) return ContentService.createTextOutput(JSON.stringify({ error: '列が見つかりません: ' + params.field })).setMimeType(ContentService.MimeType.JSON);
      updSheet.getRange(params.row, colIdx + 1).setValue(params.value);

    } else {
      var sumSheet = ss.getSheetByName('全体数値表') || ss.getSheets()[0];
      var sumData  = sumSheet.getDataRange().getValues();
      var targetRow = -1;
      for (var i = 0; i < sumData.length; i++) {
        if (String(sumData[i][1]).trim() === params.metric) { targetRow = i; break; }
      }
      if (targetRow === -1) return ContentService.createTextOutput(JSON.stringify({ error: '項目が見つかりません' })).setMimeType(ContentService.MimeType.JSON);
      sumSheet.getRange(targetRow + 1, params.monthIdx + 3).setValue(Number(params.value));
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
