function doGet(e) {
  var params = e ? e.parameter : {};
  var action = params.action || '';

  if (action === 'updateProduct') return handleUpdateProduct(params);
  if (action === 'addProduct')    return handleAddProduct(params);
  if (action === 'deleteProduct') return handleDeleteProduct(params);
  if (action === 'updateSummary') return handleUpdateSummary(params);
  if (params.type === 'products') return getProducts();
  return getSummary();
}

var SSID = '1og8AZ3rdm8f36lsE4TfI8ZsB9hhZGCi3jxWbqvp7deM';

function handleUpdateProduct(params) {
  var ss = SpreadsheetApp.openById(SSID);
  var sheet = ss.getSheetByName('商品管理表');
  var row    = parseInt(params.row);
  var colIdx = parseInt(params.colIdx);
  var value  = params.value;
  if (!isNaN(parseFloat(value)) && isFinite(value)) value = parseFloat(value);
  sheet.getRange(row, colIdx + 1).setValue(value);
  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
}

function handleAddProduct(params) {
  try {
    var ss = SpreadsheetApp.openById(SSID);
    var sheet = ss.getSheetByName('商品管理表');
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({ error: '商品管理表が見つかりません', sheets: ss.getSheets().map(function(s){return s.getName();}) })).setMimeType(ContentService.MimeType.JSON);
    var data  = sheet.getDataRange().getValues();
    var hdrIdx = findHeaderRow(data);
    if (hdrIdx === -1) return ContentService.createTextOutput(JSON.stringify({ error: 'ヘッダー行が見つかりません' })).setMimeType(ContentService.MimeType.JSON);
    var headers = data[hdrIdx].map(function(h) { return String(h).trim(); });
    var nameIdx = headers.indexOf('商品名');
    var numIdx = headers.indexOf('管理番号');
    if (numIdx === -1) numIdx = 6;
    var maxNum = 0;
    for (var i = hdrIdx + 1; i < data.length; i++) {
      var n = Number(data[i][numIdx]);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
    var fields = JSON.parse(params.fields);
    var newRow = new Array(headers.length).fill('');
    Object.keys(fields).forEach(function(key) {
      var idx = isNaN(Number(key)) ? headers.indexOf(key) : Number(key);
      if (idx !== -1 && idx < newRow.length) newRow[idx] = fields[key];
    });
    newRow[numIdx] = maxNum + 1;
    // ヘッダー行以降で最後のデータ行を探す（空行や誤データを無視）
    var lastDataRow = hdrIdx + 1;
    for (var i = hdrIdx + 1; i < data.length; i++) {
      var num = Number(data[i][numIdx]);
      if (!isNaN(num) && num > 0) lastDataRow = i + 1;
    }
    sheet.getRange(lastDataRow + 1, 1, 1, newRow.length).setValues([newRow]);
    return ContentService.createTextOutput(JSON.stringify({ success: true, appendedRow: lastDataRow + 1, newNum: maxNum + 1 })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleDeleteProduct(params) {
  try {
    var ss = SpreadsheetApp.openById(SSID);
    var sheet = ss.getSheetByName('商品管理表');
    var row = parseInt(params.row);
    sheet.deleteRow(row);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUpdateSummary(params) {
  var ss = SpreadsheetApp.openById(SSID);
  var sheet = ss.getSheetByName('全体数値表') || ss.getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim() === params.metric) {
      sheet.getRange(i + 1, parseInt(params.monthIdx) + 3).setValue(Number(params.value));
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ error: '項目が見つかりません' })).setMimeType(ContentService.MimeType.JSON);
}

function getSummary() {
  var ss = SpreadsheetApp.openById(SSID);
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
  var ss = SpreadsheetApp.openById(SSID);
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
