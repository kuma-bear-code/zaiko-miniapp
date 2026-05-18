var SPREADSHEET_ID = '1Vj3MYKP6Uc4acfL0m-Bo4ansNP2rhk8FRf1W1DXcRTw';
var INVENTORY_SHEET_NAME = 'Inventory';
var CATEGORIES_SHEET_NAME = 'Categories';
var LOG_SHEET_NAME = 'ConsumptionLog';

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';

  if (action === 'data') {
    var payload = buildApiPayload_();
    return jsonpOrJson_(payload, params.callback);
  }

  if (action === 'categories') {
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

  if (action === 'saveInventory') {
    var items = Array.isArray(body.items) ? body.items : [];
    saveInventoryItems_(items);
    syncCategories_(items);
    if (Array.isArray(body.history)) {
      saveHistorySnapshot_(body.history);
    }
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

function getSheet_(sheetName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
}

function ensureSheetHeaders_() {
  var sheet = getSheet_(INVENTORY_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + INVENTORY_SHEET_NAME);
  }

  var headers = ['蛻・｡・, '蜩∫岼蜷・, '蝨ｨ蠎ｫ謨ｰ', '譛菴守ｮ｡逅・惠蠎ｫ', '蜊倅ｽ・];
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

function saveHistorySnapshot_(history) {
  if (!history.length) return;
  var sheet = getSheet_(LOG_SHEET_NAME);
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['譌･譎・, '繧ｿ繧､繝医Ν', '隧ｳ邏ｰ']);
  }
  var latest = history[0];
  if (latest && latest.time) {
    sheet.appendRow([
      latest.time instanceof Date ? latest.time : new Date(latest.time),
      latest.title || '蜷梧悄',
      latest.detail || ''
    ]);
  }
}

function normalizeInventoryItem_(input) {
  var item = input || {};
  return {
    category: String(item.category || item['蛻・｡・] || ''),
    name: String(item.name || item['蜩∫岼蜷・] || ''),
    stock: Math.max(0, Number(item.stock != null ? item.stock : item['蝨ｨ蠎ｫ謨ｰ']) || 0),
    minStock: Math.max(0, Number(item.minStock != null ? item.minStock : item['譛菴守ｮ｡逅・惠蠎ｫ']) || 0),
    unit: String(item.unit || item['蜊倅ｽ・] || '')
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
    sheet.appendRow(['譌･譎・, '蜩∫岼蜷・, '蠅玲ｸ・, '繝｡繝｢']);
  }
  sheet.appendRow([new Date(), name, delta, memo || '']);
}

function buildApiPayload_() {
  var items = readInventoryItems_();
  var forecasts = buildPurchaseForecasts_(items);
  return {
    status: 'ok',
    items: items,
    summary: buildSummary_(items, forecasts),
    history: readHistory_(),
    categories: readCategories_(),
    forecasts: forecasts
  };
}

function buildSummary_(items, forecasts) {
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
  var forecastSoon = 0;
  if (Array.isArray(forecasts)) {
    for (var j = 0; j < forecasts.length; j++) {
      if (typeof forecasts[j].daysToMin === 'number' && forecasts[j].daysToMin <= 7) {
        forecastSoon++;
      }
    }
  }
  return {
    total: items.length,
    low: low,
    zero: zero,
    categories: Object.keys(categories).length,
    forecastSoon: forecastSoon
  };
}

function buildPurchaseForecasts_(items) {
  var statsByItem = readConsumptionStats_();
  var forecasts = [];
  var now = new Date();
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var name = String(item.name || '').trim();
    var stats = statsByItem[name];
    var forecast = buildForecastForItem_(item, stats, now);
    if (forecast) {
      forecasts.push(forecast);
    }
  }
  forecasts.sort(function(a, b) {
    return a.daysToMin - b.daysToMin;
  });
  return forecasts;
}

function readConsumptionStats_() {
  var sheet = getSheet_(LOG_SHEET_NAME);
  var stats = {};
  if (!sheet) return stats;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return stats;

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  var now = new Date();
  var windowStart = new Date(now.getTime() - 180 * 24 * 60 * 60 * 1000);

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var date = row[0] instanceof Date ? row[0] : new Date(row[0]);
    var name = String(row[1] || '').trim();
    var delta = Number(row[2]);
    if (!name || isNaN(date.getTime()) || isNaN(delta) || delta >= 0) continue;
    if (date < windowStart) continue;

    var consumed = Math.abs(delta);
    if (!stats[name]) {
      stats[name] = {
        name: name,
        totalConsumed: 0,
        events: 0,
        firstDate: date,
        lastDate: date
      };
    }
    var entry = stats[name];
    entry.totalConsumed += consumed;
    entry.events += 1;
    if (date < entry.firstDate) entry.firstDate = date;
    if (date > entry.lastDate) entry.lastDate = date;
  }

  return stats;
}

function buildForecastForItem_(item, stats, now) {
  if (!item || !stats || stats.events < 2 || !stats.totalConsumed) return null;

  var stock = Number(item.stock || 0);
  var minStock = Number(item.minStock || 0);
  var unit = String(item.unit || '個').trim() || '個';
  var spanDays = Math.max(1, (stats.lastDate - stats.firstDate) / (1000 * 60 * 60 * 24));
  var dailyRate = stats.totalConsumed / spanDays;
  if (!(dailyRate > 0)) return null;

  var available = Math.max(0, stock - minStock);
  var daysToMin = available <= 0 ? 0 : Math.ceil(available / dailyRate);
  var nextPurchaseDate = new Date(now.getTime() + daysToMin * 24 * 60 * 60 * 1000);

  return {
    name: String(item.name || ''),
    category: String(item.category || ''),
    stock: stock,
    minStock: minStock,
    unit: unit,
    totalConsumed: stats.totalConsumed,
    sampleCount: stats.events,
    daysPerUnit: spanDays / stats.totalConsumed,
    dailyRate: dailyRate,
    daysToMin: daysToMin,
    nextPurchaseDate: nextPurchaseDate.toISOString(),
    lastConsumptionDate: stats.lastDate.toISOString()
  };
}

function readHistory_() {
  var sheet = getSheet_(LOG_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, Math.min(lastRow - 1, 12), 4).getValues();
  var history = [];
  for (var i = 0; i < values.length; i++) {
    history.push({
      time: values[i][0] instanceof Date ? values[i][0].toISOString() : String(values[i][0] || ''),
      title: String(values[i][1] || ''),
      detail: String(values[i][3] || values[i][2] || '')
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
    if (!category || category === '蛻・｡・) continue;
    categories.push(category);
  }
  return categories;
}

function syncCategories_(items) {
  var sheet = getSheet_(CATEGORIES_SHEET_NAME);
  if (!sheet) return;
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var category = String(items[i].category || '').trim();
    if (category) categories[category] = true;
  }
  var list = Object.keys(categories).sort(function(a, b) {
    return a.localeCompare(b, 'ja');
  });
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  if (!list.length) {
    sheet.getRange(1, 1).setValue('蛻・｡・);
    return;
  }
  var rows = [['蛻・｡・]];
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
  if (message === '菴ｿ縺・婿') {
    sendReply_(replyToken, getInstructions_());
    return;
  }
  if (message === '荳隕ｧ') {
    sendReply_(replyToken, listCategorizedProducts_());
    return;
  }
  if (message === '菴主惠蠎ｫ') {
    sendReply_(replyToken, listLowStockProducts_());
    return;
  }
  if (message === '蝨ｨ蠎ｫ蛻・ｌ') {
    sendReply_(replyToken, listZeroStockProducts_());
    return;
  }
  if (message.indexOf('蜈･蠎ｫ ') === 0) {
    var inParts = message.split(/\s+/);
    var inName = inParts[1];
    var inDelta = Number(inParts[2] || 1);
    adjustInventoryItem_(inName, Math.abs(inDelta), 'LINE蜈･蠎ｫ');
    sendReply_(replyToken, inName + ' 繧・' + Math.abs(inDelta) + ' 蠅励ｄ縺励∪縺励◆縲・);
    return;
  }
  if (message.indexOf('蜃ｺ蠎ｫ ') === 0) {
    var outParts = message.split(/\s+/);
    var outName = outParts[1];
    var outDelta = Number(outParts[2] || 1);
    adjustInventoryItem_(outName, -Math.abs(outDelta), 'LINE蜃ｺ蠎ｫ');
    sendReply_(replyToken, outName + ' 繧・' + Math.abs(outDelta) + ' 貂帙ｉ縺励∪縺励◆縲・);
    return;
  }
  if (message.indexOf('蜑企勁 ') === 0) {
    var delName = message.replace(/^蜑企勁\s+/, '');
    if (deleteInventoryItem_(delName)) {
      sendReply_(replyToken, delName + ' 繧貞炎髯､縺励∪縺励◆縲・);
    } else {
      sendReply_(replyToken, delName + ' 縺ｯ隕九▽縺九ｊ縺ｾ縺帙ｓ縺ｧ縺励◆縲・);
    }
    return;
  }

  sendReply_(replyToken, '縲御ｸ隕ｧ縲阪御ｽ主惠蠎ｫ縲阪悟惠蠎ｫ蛻・ｌ縲阪悟・蠎ｫ 蜩∫岼蜷・謨ｰ驥上阪悟・蠎ｫ 蜩∫岼蜷・謨ｰ驥上阪悟炎髯､ 蜩∫岼蜷阪阪′菴ｿ縺医∪縺吶・);
}

function getInstructions_() {
  return [
    '縲蝉ｽｿ縺・婿縲・,
    '繝ｻ荳隕ｧ : 蝨ｨ蠎ｫ繧偵き繝・ざ繝ｪ蛻･縺ｫ陦ｨ遉ｺ',
    '繝ｻ菴主惠蠎ｫ : 譛菴主惠蠎ｫ譛ｪ貅繧定｡ｨ遉ｺ',
    '繝ｻ蝨ｨ蠎ｫ蛻・ｌ : 蝨ｨ蠎ｫ0繧定｡ｨ遉ｺ',
    '繝ｻ蜈･蠎ｫ 蜩∫岼蜷・謨ｰ驥・: 蝨ｨ蠎ｫ繧貞｢励ｄ縺・,
    '繝ｻ蜃ｺ蠎ｫ 蜩∫岼蜷・謨ｰ驥・: 蝨ｨ蠎ｫ繧呈ｸ帙ｉ縺・,
    '繝ｻ蜑企勁 蜩∫岼蜷・: 蜩∫岼繧貞炎髯､',
    '',
    'GitHub Pages 縺ｮ逕ｻ髱｢縺九ｉ繧よ桃菴懊〒縺阪∪縺吶・
  ].join('\n');
}

function listCategorizedProducts_() {
  var items = readInventoryItems_();
  if (!items.length) return '逋ｻ骭ｲ縺輔ｌ縺ｦ縺・ｋ蜩∫岼縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・;
  var categories = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (!categories[item.category]) categories[item.category] = [];
    categories[item.category].push(item);
  }
  var msg = '縲仙惠蠎ｫ荳隕ｧ縲曾n';
  Object.keys(categories).sort().forEach(function(category) {
    msg += '\n[' + category + ']\n';
    categories[category].forEach(function(item) {
      msg += item.name + ' : ' + item.stock + item.unit + ' / 譛菴・' + item.minStock + item.unit + '\n';
    });
  });
  return msg;
}

function listLowStockProducts_() {
  var items = readInventoryItems_().filter(function(item) {
    return Number(item.stock) > 0 && Number(item.stock) < Number(item.minStock);
  });
  if (!items.length) return '菴主惠蠎ｫ縺ｮ蜩∫岼縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・;
  return '縲蝉ｽ主惠蠎ｫ縲曾n' + items.map(function(item) {
    return item.name + ' : ' + item.stock + item.unit + ' / 譛菴・' + item.minStock + item.unit;
  }).join('\n');
}

function listZeroStockProducts_() {
  var items = readInventoryItems_().filter(function(item) {
    return Number(item.stock) <= 0;
  });
  if (!items.length) return '蝨ｨ蠎ｫ蛻・ｌ縺ｮ蜩∫岼縺ｯ縺ゅｊ縺ｾ縺帙ｓ縲・;
  return '縲仙惠蠎ｫ蛻・ｌ縲曾n' + items.map(function(item) {
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
