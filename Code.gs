var SPREADSHEET_ID = '1eOg7cd1V6BlvsbiJgeoHXAty1TwLOwRXUEpvyhIBtII';
var INVENTORY_SHEET_NAME = 'Inventory';
var CATEGORIES_SHEET_NAME = 'Categories';
var LOG_SHEET_NAME = 'ConsumptionLog';

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';

  if (action === 'data') {
    var auth = authorizeWebAppRequest_(params.idToken);
    if (!auth.ok) return jsonpOrJson_({ status: 'error', message: auth.message }, params.callback);
    var payload = buildApiPayload_();
    return jsonpOrJson_(payload, params.callback);
  }

  if (action === 'rawdata') {
    var rawAuth = authorizeWebAppRequest_(params.idToken);
    if (!rawAuth.ok) return respondJson_({ status: 'error', message: rawAuth.message });
    return respondJson_(buildApiPayload_());
  }

  if (action === 'categories') {
    var categoryAuth = authorizeWebAppRequest_(params.idToken);
    if (!categoryAuth.ok) return jsonpOrJson_({ status: 'error', message: categoryAuth.message }, params.callback);
    return jsonpOrJson_({
      status: 'ok',
      categories: readCategories_()
    }, params.callback);
  }

  if (action === 'ping') {
    return jsonpOrJson_({
      status: 'ok',
      timestamp: new Date().toISOString(),
      summary: buildSummary_()
    }, params.callback);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Inventory API is running. Use ?action=data or POST action=saveInventory.',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = parseRequestBody_(e);

  if (body.payload) {
    try {
      var payloadData = JSON.parse(body.payload);
      for (var payloadKey in payloadData) {
        body[payloadKey] = payloadData[payloadKey];
      }
    } catch (err) {}
  }

  if (body.events && Array.isArray(body.events)) {
    return handleLineWebhook_(body);
  }

  var action = body.action || (e && e.parameter && e.parameter.action) || '';

  var auth = authorizeWebAppRequest_(body.idToken || (e && e.parameter && e.parameter.idToken));
  if (!auth.ok) {
    return respondJson_({
      status: 'error',
      message: auth.message
    });
  }

  if (action === 'saveInventory') {
    var items = Array.isArray(body.items) ? body.items : [];
    saveInventoryItems_(items);
    syncCategories_(items);
    return respondJson_({
      status: 'ok',
      message: 'inventory saved',
      summary: buildSummary_()
    });
  }

  if (action === 'addItem') {
    upsertInventoryItem_(body.item || body);
    return respondJson_({
      status: 'ok',
      message: 'item saved',
      summary: buildSummary_()
    });
  }

  if (action === 'adjustItem') {
    adjustInventoryItem_(body.name, Number(body.delta || 0), body.memo || '');
    return respondJson_({
      status: 'ok',
      message: 'item adjusted',
      summary: buildSummary_()
    });
  }

  if (action === 'deleteItem') {
    deleteInventoryItem_(body.name);
    return respondJson_({
      status: 'ok',
      message: 'item deleted',
      summary: buildSummary_()
    });
  }

  return respondJson_({
    status: 'ok',
    message: 'no action',
    summary: buildSummary_()
  });
}

function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  var contents = e.postData.contents;
  try {
    return JSON.parse(contents);
  } catch (err) {
    var out = {};
    var pairs = contents.split('&');
    for (var i = 0; i < pairs.length; i++) {
      var pair = pairs[i];
      if (!pair) continue;
      var idx = pair.indexOf('=');
      var key = idx >= 0 ? decodeURIComponent(pair.slice(0, idx).replace(/\+/g, ' ')) : decodeURIComponent(pair.replace(/\+/g, ' '));
      var value = idx >= 0 ? decodeURIComponent(pair.slice(idx + 1).replace(/\+/g, ' ')) : '';
      out[key] = value;
    }
    return out;
  }
}

function respondJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJson_(obj, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return respondJson_(obj);
}

function authorizeWebAppRequest_(idToken) {
  var props = PropertiesService.getScriptProperties();
  var channelId = String(props.getProperty('LIFF_CHANNEL_ID') || '').trim();
  var allowed = String(props.getProperty('ALLOWED_LINE_USER_IDS') || '')
    .split(',')
    .map(function(value) { return value.trim(); })
    .filter(function(value) { return value; });

  if (!channelId || allowed.length === 0) {
    return { ok: true, userId: '' };
  }

  if (!idToken) {
    return { ok: false, message: 'LINE login required' };
  }

  try {
    var res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
      method: 'post',
      payload: {
        id_token: idToken,
        client_id: channelId
      },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      return { ok: false, message: 'LINE token verification failed' };
    }

    var claims = JSON.parse(res.getContentText());
    var userId = String(claims.sub || '');
    if (allowed.indexOf(userId) === -1) {
      return { ok: false, message: 'This LINE account is not allowed' };
    }
    return { ok: true, userId: userId };
  } catch (err) {
    return { ok: false, message: 'LINE authorization failed' };
  }
}

function getSheet_(sheetName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
}

function ensureSheetHeaders_() {
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + INVENTORY_SHEET_NAME);
  }

  var headers = ['分類', '品目名', '在庫数', '最低管理在庫', '単位'];
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeader = true;
  for (var i = 0; i < headers.length; i++) {
    if (String(firstRow[i] || '') !== headers[i]) {
      needsHeader = true;
      break;
    }
    needsHeader = false;
  }

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function readInventoryItems_() {
  ensureSheetHeaders_();
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var items = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue;
    items.push({
      category: String(row[0] || ''),
      name: String(row[1] || ''),
      stock: Number(row[2] || 0),
      minStock: Number(row[3] || 0),
      unit: String(row[4] || '')
    });
  }
  return items;
}

function saveInventoryItems_(items) {
  ensureSheetHeaders_();
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  var cleaned = [];
  for (var i = 0; i < items.length; i++) {
    var item = normalizeInventoryItem_(items[i]);
    if (!item.name) continue;
    cleaned.push([item.category, item.name, item.stock, item.minStock, item.unit]);
  }
  cleaned.sort(function(a, b) {
    var cat = String(a[0]).localeCompare(String(b[0]), 'ja');
    if (cat !== 0) return cat;
    return String(a[1]).localeCompare(String(b[1]), 'ja');
  });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
  }
  if (cleaned.length) {
    sheet.getRange(2, 1, cleaned.length, 5).setValues(cleaned);
  }
}

function normalizeInventoryItem_(input) {
  var item = input || {};
  return {
    category: String(item.category || item['分類'] || ''),
    name: String(item.name || item['品目名'] || ''),
    stock: Math.max(0, Number(item.stock != null ? item.stock : item['在庫数']) || 0),
    minStock: Math.max(0, Number(item.minStock != null ? item.minStock : item['最低管理在庫']) || 0),
    unit: String(item.unit || item['単位'] || '')
  };
}

function upsertInventoryItem_(input) {
  ensureSheetHeaders_();
  var item = normalizeInventoryItem_(input);
  if (!item.name) throw new Error('name is required');

  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1] || '') === item.name) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    sheet.appendRow([item.category, item.name, item.stock, item.minStock, item.unit]);
  } else {
    sheet.getRange(foundRow, 1, 1, 5).setValues([[item.category, item.name, item.stock, item.minStock, item.unit]]);
  }
}

function deleteInventoryItem_(name) {
  if (!name) return false;
  ensureSheetHeaders_();
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1] || '') === String(name)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function adjustInventoryItem_(name, delta, memo) {
  if (!name) throw new Error('name is required');
  ensureSheetHeaders_();
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][1] || '') === String(name)) {
      var stock = Math.max(0, Number(values[i][2] || 0) + Number(delta || 0));
      sheet.getRange(i + 1, 3).setValue(stock);
      logConsumption_(name, delta, memo);
      return true;
    }
  }
  throw new Error('item not found: ' + name);
}

function logConsumption_(name, delta, memo) {
  var sheet = getSheet_(LOG_SHEET_NAME);
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['品目', '日付', '数量']);
  }
  if (Number(delta || 0) < 0) {
    sheet.appendRow([name, formatDate_(new Date()), Math.abs(Number(delta || 0))]);
  }
}

function buildApiPayload_() {
  var items = readInventoryItems_();
  return {
    status: 'ok',
    items: items,
    summary: buildSummary_(items),
    history: readHistory_(),
    categories: readCategories_().concat(readInventoryCategories_(items))
  };
}

function buildSummary_(items) {
  items = items || readInventoryItems_();
  var low = 0;
  var zero = 0;
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    categories[item.category || '未分類'] = true;
    if (Number(item.stock || 0) <= 0) {
      zero++;
    } else if (Number(item.stock || 0) < Number(item.minStock || 0)) {
      low++;
    }
  }
  return {
    total: items.length,
    low: low,
    zero: zero,
    categories: Object.keys(categories).length
  };
}

function readHistory_() {
  var sheet = getSheet_(LOG_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, Math.min(lastRow - 1, 50), 3).getValues();
  var history = [];
  for (var i = 0; i < values.length; i++) {
    history.push({
      item: String(values[i][0] || ''),
      date: String(values[i][1] || ''),
      quantity: Number(values[i][2] || 0)
    });
  }
  return history.reverse();
}

function readCategories_() {
  var sheet = getSheet_(CATEGORIES_SHEET_NAME);
  if (!sheet) {
    return [];
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return [];
  }
  var values = sheet.getRange(1, 1, lastRow, Math.min(2, sheet.getLastColumn())).getValues();
  var categories = [];
  for (var i = 0; i < values.length; i++) {
    var category = String(values[i][0] || '').trim();
    if (!category || category === '分類') continue;
    categories.push(category);
  }
  return categories;
}

function readInventoryCategories_(items) {
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var category = String(items[i].category || '').trim();
    if (category) categories[category] = true;
  }
  return Object.keys(categories);
}

function syncCategories_(items) {
  var sheet = getSheet_(CATEGORIES_SHEET_NAME);
  if (!sheet) return;
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var category = String(items[i].category || '').trim();
    if (category) categories[category] = true;
  }
  var existing = readCategories_();
  for (var j = 0; j < existing.length; j++) {
    if (existing[j]) categories[existing[j]] = true;
  }
  var list = Object.keys(categories).sort(function(a, b) {
    return a.localeCompare(b, 'ja');
  });
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  if (!list.length) {
    sheet.getRange(1, 1).setValue('分類');
    return;
  }
  var rows = [['分類']];
  for (var j = 0; j < list.length; j++) {
    rows.push([list[j]]);
  }
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

function handleLineWebhook_(body) {
  var events = body.events || [];
  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    if (!event || !event.replyToken) continue;
    if (event.type === 'message' && event.message && event.message.type === 'text') {
      handleLineMessage_(event.replyToken, String(event.message.text || '').trim());
    }
  }
  return respondJson_({ status: 'ok' });
}

function handleLineMessage_(replyToken, message) {
  if (message === '使い方') {
    sendReply_(replyToken, getInstructions_());
    return;
  }
  if (message === '一覧') {
    sendReply_(replyToken, listCategorizedProducts_());
    return;
  }
  if (message === '低在庫') {
    sendReply_(replyToken, listLowStockProducts_());
    return;
  }
  if (message === '在庫切れ') {
    sendReply_(replyToken, listZeroStockProducts_());
    return;
  }
  if (message.indexOf('入庫 ') === 0) {
    var inParts = message.split(/\s+/);
    var inName = inParts[1];
    var inDelta = Number(inParts[2] || 1);
    adjustInventoryItem_(inName, Math.abs(inDelta), 'LINE入庫');
    sendReply_(replyToken, inName + ' を ' + Math.abs(inDelta) + ' 増やしました。');
    return;
  }
  if (message.indexOf('出庫 ') === 0) {
    var outParts = message.split(/\s+/);
    var outName = outParts[1];
    var outDelta = Number(outParts[2] || 1);
    adjustInventoryItem_(outName, -Math.abs(outDelta), 'LINE出庫');
    sendReply_(replyToken, outName + ' を ' + Math.abs(outDelta) + ' 減らしました。');
    return;
  }
  if (message.indexOf('削除 ') === 0) {
    var delName = message.replace(/^削除\s+/, '');
    if (deleteInventoryItem_(delName)) {
      sendReply_(replyToken, delName + ' を削除しました。');
    } else {
      sendReply_(replyToken, delName + ' は見つかりませんでした。');
    }
    return;
  }

  sendReply_(replyToken, '「一覧」「低在庫」「在庫切れ」「入庫 品目名 数量」「出庫 品目名 数量」「削除 品目名」が使えます。');
}

function getInstructions_() {
  return [
    '【使い方】',
    '・一覧 : 在庫をカテゴリ別に表示',
    '・低在庫 : 最低在庫未満を表示',
    '・在庫切れ : 在庫0を表示',
    '・入庫 品目名 数量 : 在庫を増やす',
    '・出庫 品目名 数量 : 在庫を減らす',
    '・削除 品目名 : 品目を削除',
    '',
    'GitHub Pages の画面からも操作できます。'
  ].join('\n');
}

function listCategorizedProducts_() {
  var items = readInventoryItems_();
  if (!items.length) return '登録されている品目はありません。';
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (!categories[item.category]) categories[item.category] = [];
    categories[item.category].push(item);
  }
  var msg = '【在庫一覧】\n';
  Object.keys(categories).sort().forEach(function(category) {
    msg += '\n[' + category + ']\n';
    categories[category].forEach(function(item) {
      msg += item.name + ' : ' + item.stock + item.unit + ' / 最低 ' + item.minStock + item.unit + '\n';
    });
  });
  return msg;
}

function listLowStockProducts_() {
  var items = readInventoryItems_().filter(function(item) {
    return Number(item.stock) > 0 && Number(item.stock) < Number(item.minStock);
  });
  if (!items.length) return '低在庫の品目はありません。';
  return '【低在庫】\n' + items.map(function(item) {
    return item.name + ' : ' + item.stock + item.unit + ' / 最低 ' + item.minStock + item.unit;
  }).join('\n');
}

function listZeroStockProducts_() {
  var items = readInventoryItems_().filter(function(item) {
    return Number(item.stock) <= 0;
  });
  if (!items.length) return '在庫切れの品目はありません。';
  return '【在庫切れ】\n' + items.map(function(item) {
    return item.name;
  }).join('\n');
}

function sendReply_(replyToken, message) {
  var token = PropertiesService.getScriptProperties().getProperty('CHANNEL_ACCESS_TOKEN');
  if (!token) {
    throw new Error('CHANNEL_ACCESS_TOKEN is not set in Script Properties.');
  }
  var url = 'https://api.line.me/v2/bot/message/reply';
  var payload = {
    replyToken: replyToken,
    messages: [{
      type: 'text',
      text: message
    }]
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function formatDate_(date) {
  var d = new Date(date);
  var year = d.getFullYear();
  var month = String(d.getMonth() + 1).padStart(2, '0');
  var day = String(d.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}
