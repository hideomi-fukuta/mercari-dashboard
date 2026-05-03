function doGet(e) {
  var type = e && e.parameter && e.parameter.type ? e.parameter.type : 'summary';

  if (type === 'products') {
    return getProducts();
  } else {
    return getSummary();
  }
}

function getSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('全体数値表') || ss.getSheets()[0];
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
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('商品管理表');
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ error: '商品管理表が見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();

  // ヘッダー行を探す（商品名 or 管理番号がある行）
  var headerRowIdx = -1;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      if (String(row[j]).trim() === '商品名' || String(row[j]).trim() === '管理番号') {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx !== -1) break;
  }

  if (headerRowIdx === -1) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'ヘッダー行が見つかりません' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var headers = data[headerRowIdx].map(function(h) { return String(h).trim(); });
  var products = [];

  for (var i = headerRowIdx + 1; i < data.length; i++) {
    var row = data[i];
    // 商品名か管理番号がある行だけ取り込む
    var nameIdx = headers.indexOf('商品名');
    var numIdx  = headers.indexOf('管理番号');
    if (!row[nameIdx] && !row[numIdx]) continue;

    var product = { _row: i + 1 }; // 1-based行番号(更新時に使用)
    headers.forEach(function(h, j) {
      if (h) {
        var val = row[j];
        // 日付はISO文字列に変換
        if (val instanceof Date) {
          product[h] = val.getFullYear() + '/' + (val.getMonth()+1) + '/' + val.getDate();
        } else {
          product[h] = val;
        }
      }
    });
    products.push(product);
  }

  return ContentService.createTextOutput(JSON.stringify({ headers: headers, products: products }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    if (params.type === 'products') { return getProducts(); }
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (params.action === 'updateProduct') {
      // 商品管理表のセルを更新
      var sheet = ss.getSheetByName('商品管理表');
      var data = sheet.getDataRange().getValues();
      var headers = data[params.headerRow - 1].map(function(h) { return String(h).trim(); });
      var colIdx = headers.indexOf(params.field);
      if (colIdx === -1) {
        return ContentService.createTextOutput(JSON.stringify({ error: '列が見つかりません: ' + params.field }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.getRange(params.row, colIdx + 1).setValue(params.value);

    } else {
      // 全体数値表のセルを更新（既存機能）
      var sheet2 = ss.getSheetByName('全体数値表') || ss.getSheets()[0];
      var data2 = sheet2.getDataRange().getValues();
      var targetRow = -1;
      for (var i = 0; i < data2.length; i++) {
        if (String(data2[i][1]).trim() === params.metric) { targetRow = i; break; }
      }
      if (targetRow === -1) {
        return ContentService.createTextOutput(JSON.stringify({ error: '項目が見つかりません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet2.getRange(targetRow + 1, params.monthIdx + 3).setValue(Number(params.value));
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
