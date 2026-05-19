// =======================================================
// LINE x Spreadsheet Inventory BOT + Web App API
// - LINE group operation flow stays as the primary behavior
// - GitHub Pages / LIFF app can read and write through doGet/doPost
// - Optional LINE Login guard for the web app only
// =======================================================

// ---------- Fallback defaults ----------
const DEFAULTS = {
  SPREADSHEET_ID: '1eOg7cd1V6BlvsbiJgeoHXAty1TwLOwRXUEpvyhIBtII',
  CHANNEL_ACCESS_TOKEN: '',
  GROUP_ID: ''
};

// ---------- Config ----------
function CONF() {
  const p = PropertiesService.getScriptProperties();
  return {
    SPREADSHEET_ID: p.getProperty('SPREADSHEET_ID') || DEFAULTS.SPREADSHEET_ID,
    CHANNEL_ACCESS_TOKEN: p.getProperty('CHANNEL_ACCESS_TOKEN') || DEFAULTS.CHANNEL_ACCESS_TOKEN,
    GROUP_ID: p.getProperty('GROUP_ID') || DEFAULTS.GROUP_ID
  };
}

function setupSecrets() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('SPREADSHEET_ID', '<<YOUR_SPREADSHEET_ID>>');
  p.setProperty('CHANNEL_ACCESS_TOKEN', '<<YOUR_CHANNEL_ACCESS_TOKEN>>');
  p.setProperty('GROUP_ID', '<<YOUR_GROUP_OR_USER_ID>>');
}

// ---------- UI layout ----------
const COLS = { product: 7, stock: 3, min: 3, unit: 2, buttons: 11 };

// ---------- Constants ----------
const SHEETS = {
  inventory: 'Inventory',
  categories: 'Categories',
  log: 'ConsumptionLog'
};

const STATE = {
  awaitingOperation: 'awaitingOperation',
  awaitingQuantity: 'awaitingQuantity',
  awaitingMinimum: 'awaitingMinimum',
  awaitingCategory: 'awaitingCategory',
  awaitingUnit: 'awaitingUnit',
  awaitingInitialStock: 'awaitingInitialStock',
  awaitingPurchaseDecision: 'awaitingPurchaseDecision',
  awaitingPurchaseQuantity: 'awaitingPurchaseQuantity',
  awaitingReclassificationSelection: 'awaitingReclassificationSelection',
  awaitingReclassificationTarget: 'awaitingReclassificationTarget',
  awaitingDeletionByNumber: 'awaitingDeletionByNumber',
  awaitingChangeSelection: 'awaitingChangeSelection',
  awaitingChangeOption: 'awaitingChangeOption',
  awaitingChangeInventory: 'awaitingChangeInventory',
  awaitingChangeMinInventory: 'awaitingChangeMinInventory',
  awaitingChangeBoth: 'awaitingChangeBoth',
  awaitingChangeBothMin: 'awaitingChangeBothMin',
  awaitingListSelection: 'awaitingListSelection',
  awaitingListSelectionByCategory: 'awaitingListSelectionByCategory',
  awaitingListOperation: 'awaitingListOperation',
  awaitingListQuantity: 'awaitingListQuantity'
};

const MSG_LIMITS = { replyMaxParts: 5, chunkSize: 1800 };
const ROWS_PER_BUBBLE = 5;
const BUBBLES_PER_PAGE = 3;
const PAGE_SIZE = ROWS_PER_BUBBLE * BUBBLES_PER_PAGE;
const QR_CAT_PAGE_SIZE = 10;

// =======================================================
// Webhook / Web App
// =======================================================
function doPost(e) {
  try {
    const body = parseRequestBody_(e);
    const ss = getSpreadsheetOrThrow_();
    ensureSheets_(ss);

    if (body.payload) {
      try {
        const payloadData = JSON.parse(body.payload);
        Object.keys(payloadData).forEach(function(key) {
          body[key] = payloadData[key];
        });
      } catch (err) {}
    }

    if (body.events && Array.isArray(body.events)) {
      return handleLineWebhook_(body, ss);
    }

    const action = body.action || (e && e.parameter && e.parameter.action) || '';
    const auth = authorizeWebAppRequest_(body.idToken || (e && e.parameter && e.parameter.idToken));
    if (!auth.ok) {
      return respondJson_({ status: 'error', message: auth.message });
    }

    if (action === 'saveInventory') {
      const items = Array.isArray(body.items) ? body.items : [];
      saveInventoryItems_(ss, items);
      syncCategories_(ss, items);
      return respondJson_({ status: 'ok', message: 'inventory saved', summary: buildSummary_(ss) });
    }

    if (action === 'addItem') {
      upsertInventoryItem_(ss, body.item || body);
      return respondJson_({ status: 'ok', message: 'item saved', summary: buildSummary_(ss) });
    }

    if (action === 'adjustItem') {
      adjustInventoryItem_(ss, body.name, Number(body.delta || 0), body.memo || '');
      return respondJson_({ status: 'ok', message: 'item adjusted', summary: buildSummary_(ss) });
    }

    if (action === 'deleteItem') {
      deleteProduct(ss, body.name);
      return respondJson_({ status: 'ok', message: 'item deleted', summary: buildSummary_(ss) });
    }

    return respondJson_({ status: 'ok', message: 'no action', summary: buildSummary_(ss) });
  } catch (err) {
    console.error('doPost error', err);
    return respondJson_({ status: 'error', message: shortErr_(err) });
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';
    const ss = getSpreadsheetOrThrow_();
    ensureSheets_(ss);

    if (action === 'data') {
      const auth = authorizeWebAppRequest_(params.idToken);
      if (!auth.ok) return jsonpOrJson_({ status: 'error', message: auth.message }, params.callback);
      return jsonpOrJson_(buildApiPayload_(ss), params.callback);
    }

    if (action === 'rawdata') {
      const rawAuth = authorizeWebAppRequest_(params.idToken);
      if (!rawAuth.ok) return respondJson_({ status: 'error', message: rawAuth.message });
      return respondJson_(buildApiPayload_(ss));
    }

    if (action === 'categories') {
      const categoryAuth = authorizeWebAppRequest_(params.idToken);
      if (!categoryAuth.ok) return jsonpOrJson_({ status: 'error', message: categoryAuth.message }, params.callback);
      return jsonpOrJson_({ status: 'ok', categories: getCategoriesList(ss) }, params.callback);
    }

    if (action === 'ping') {
      return jsonpOrJson_({
        status: 'ok',
        timestamp: new Date().toISOString(),
        summary: buildSummary_(ss)
      }, params.callback);
    }

    return respondJson_({
      status: 'ok',
      message: 'Inventory API is running. Use ?action=data or POST action=saveInventory.',
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    console.error('doGet error', err);
    return respondJson_({ status: 'error', message: shortErr_(err) });
  }
}

// =======================================================
// Startup / sheets
// =======================================================
function getSpreadsheetOrThrow_() {
  const conf = CONF();
  if (!conf.SPREADSHEET_ID || conf.SPREADSHEET_ID.indexOf('PUT_') === 0) {
    throw new Error('SPREADSHEET_ID is not set.');
  }
  try {
    return SpreadsheetApp.openById(conf.SPREADSHEET_ID);
  } catch (e) {
    throw new Error('SPREADSHEET_ID is invalid or inaccessible.');
  }
}

function ensureSheets_(ss) {
  let inv = ss.getSheetByName(SHEETS.inventory);
  if (!inv) {
    inv = ss.insertSheet(SHEETS.inventory);
    inv.appendRow(['分類', '品目', '在庫', '最低在庫', '単位']);
  } else {
    const cols = inv.getLastColumn();
    if (cols < 5) inv.insertColumnsAfter(cols, 5 - cols);
    const header = inv.getRange(1, 1, 1, 5).getValues()[0];
    if (!header[0]) inv.getRange(1, 1).setValue('分類');
    if (!header[1]) inv.getRange(1, 2).setValue('品目');
    if (!header[2]) inv.getRange(1, 3).setValue('在庫');
    if (!header[3]) inv.getRange(1, 4).setValue('最低在庫');
    if (!header[4]) inv.getRange(1, 5).setValue('単位');
  }

  let cat = ss.getSheetByName(SHEETS.categories);
  if (!cat) {
    cat = ss.insertSheet(SHEETS.categories);
    cat.appendRow(['未分類']);
  }

  let log = ss.getSheetByName(SHEETS.log);
  if (!log) {
    log = ss.insertSheet(SHEETS.log);
    log.appendRow(['品目', '日付', '数量']);
  }
}

// =======================================================
// Web auth
// =======================================================
function authorizeWebAppRequest_(idToken) {
  const props = PropertiesService.getScriptProperties();
  const channelId = String(props.getProperty('LIFF_CHANNEL_ID') || '').trim();
  const allowed = String(props.getProperty('ALLOWED_LINE_USER_IDS') || '')
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
    const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
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

    const claims = JSON.parse(res.getContentText());
    const userId = String(claims.sub || '');
    if (allowed.indexOf(userId) === -1) {
      return { ok: false, message: 'This LINE account is not allowed' };
    }
    return { ok: true, userId: userId };
  } catch (err) {
    return { ok: false, message: 'LINE authorization failed' };
  }
}

// =======================================================
// API payload / parsing
// =======================================================
function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const contents = e.postData.contents;
  try {
    return JSON.parse(contents);
  } catch (err) {
    const out = {};
    contents.split('&').forEach(function(pair) {
      if (!pair) return;
      const idx = pair.indexOf('=');
      const key = idx >= 0 ? decodeURIComponent(pair.slice(0, idx).replace(/\+/g, ' ')) : decodeURIComponent(pair.replace(/\+/g, ' '));
      const value = idx >= 0 ? decodeURIComponent(pair.slice(idx + 1).replace(/\+/g, ' ')) : '';
      out[key] = value;
    });
    return out;
  }
}

function respondJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJson_(obj, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return respondJson_(obj);
}

function buildApiPayload_(ss) {
  const items = readInventoryItems_(ss);
  return {
    status: 'ok',
    items: items,
    history: readHistory_(ss),
    summary: buildSummary_(ss, items),
    categories: getCategoriesList(ss).concat(readInventoryCategories_(items)).filter(uniqueOnly_)
  };
}

function buildSummary_(ss, items) {
  items = items || readInventoryItems_(ss);
  let low = 0;
  let zero = 0;
  const categories = {};
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (item.category) categories[item.category] = true;
    if (Number(item.stock || 0) <= 0) zero++;
    else if (Number(item.stock || 0) < Number(item.minStock || 0)) low++;
  }
  return {
    total: items.length,
    low: low,
    zero: zero,
    categories: Object.keys(categories).length
  };
}

function uniqueOnly_(value, index, array) {
  return array.indexOf(value) === index;
}

// =======================================================
// LINE event handling
// =======================================================
function handleLineWebhook_(body, ss) {
  const events = body.events || [];
  events.forEach(function(ev) {
    try {
      if (!ev || !ev.replyToken) return;
      if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
        handleMessageEvent(ev, ss);
      } else if (ev.type === 'postback') {
        handlePostbackEvent(ev, ss);
      }
    } catch (innerErr) {
      console.error('handleEvent error', innerErr);
      if (ev.replyToken) safeReply(ev.replyToken, 'エラー: ' + shortErr_(innerErr));
    }
  });
  return respondJson_({ status: 'ok' });
}

function handleMessageEvent(event, ss) {
  const text = (event.message.text || '').trim();
  const replyToken = event.replyToken;
  const src = event.source;

  if (isLineIdCommand_(text)) {
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, buildLineIdMessage_(event));
  }

  if (text === '使い方') {
    clearPendingStateBySource(src);
    return multiReplyQ(replyToken, ss, chunkify(getUsageInstructions()));
  }
  if (text === '使用頻度') {
    clearPendingStateBySource(src);
    return multiReplyQ(replyToken, ss, chunkify(getUsageFrequencyInstructions(ss)));
  }
  if (text === '不足一覧') {
    clearPendingStateBySource(src);
    return multiReplyQ(replyToken, ss, chunkify(getShortageListByCategory(ss)));
  }
  if (text === '未分類一覧') {
    const un = getUnclassifiedItems(ss);
    if (un.length === 0) {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '未分類の品目はありません。');
    }
    let msg = '【未分類一覧】\n';
    un.forEach(function(r, i) { msg += (i + 1) + '. ' + r.product + '\n'; });
    msg += '\n再分類したい品目の番号を、カンマまたは改行区切りで入力してください。\n（0または「再分類なし」でキャンセル）';
    setPendingStateBySource(src, { state: STATE.awaitingReclassificationSelection, unclassified: un });
    return multiReplyQ(replyToken, ss, chunkify(msg));
  }
  if (text === '変更') {
    const numbered = getNumberedInventoryList(ss);
    setPendingStateBySource(src, { state: STATE.awaitingChangeSelection, mapping: numbered.mapping });
    const msg = numbered.message + '\n\nどの品目を修正しますか？番号を入力してください。（修正不要なら0 または「不要」）';
    return multiReplyQ(replyToken, ss, chunkify(msg));
  }
  if (text === '削除') {
    const numbered = getNumberedInventoryList(ss);
    setPendingStateBySource(src, { state: STATE.awaitingDeletionByNumber, mapping: numbered.mapping });
    const msg = numbered.message + '\n\nどの品目を削除しますか？番号を入力してください。（キャンセルは0 または「不要」）';
    return multiReplyQ(replyToken, ss, chunkify(msg));
  }

  const mCat = text.match(/^(.+?)一覧(?:\s+(\d+))?$/);
  if (mCat) {
    const category = mCat[1];
    const page = mCat[2] ? Math.max(1, parseInt(mCat[2], 10)) : 1;
    const cats = getCategoriesList(ss);
    if (cats.indexOf(category) !== -1) {
      return replyCategoryFlexList(replyToken, src, ss, category, page);
    }
  }

  const mAll = text.match(/^一覧(?:\s+(\d+))?$/);
  if (mAll) {
    const page = mAll[1] ? Math.max(1, parseInt(mAll[1], 10)) : 1;
    return replyAllFlexList(replyToken, src, ss, page);
  }

  const pending = getPendingStateBySource(src);
  if (pending) {
    return handlePendingFlow(ss, event, pending);
  }

  const itemName = text;
  if (isProductExists(ss, itemName)) {
    const count = getInventoryCount(ss, itemName);
    const unit = getUnit(ss, itemName);
    setPendingStateBySource(src, { state: STATE.awaitingOperation, item: itemName });
    return replyWithQuick(replyToken, itemName + 'の現在の在庫は ' + count + ' ' + unit + ' です。\n『購（購入）』か『消（消費）』を選んでください。', ['購', '消']);
  }

  addProduct(ss, itemName, 0, '', '個');
  setPendingStateBySource(src, { state: STATE.awaitingMinimum, item: itemName });
  return safeReplyQ(replyToken, ss, '新しい品目「' + itemName + '」を追加しました。現在の在庫は 0 個です。\nまず、最低管理在庫数を数字で入力してください。');
}

function handlePostbackEvent(event, ss) {
  const replyToken = event.replyToken;
  const src = event.source;
  const data = parseQueryString((event.postback && event.postback.data) || '');

  if (data.nav === 'all') {
    const page = Math.max(1, parseInt(data.page || '1', 10));
    return replyAllFlexList(replyToken, src, ss, page);
  }
  if (data.nav === 'cat') {
    const page = Math.max(1, parseInt(data.page || '1', 10));
    const category = data.category ? decodeURIComponent(data.category) : '';
    if (!category) return safeReplyQ(replyToken, ss, 'カテゴリ情報が不足しています。');
    return replyCategoryFlexList(replyToken, src, ss, category, page);
  }
  if (data.nav === 'qr' && data.qtype === 'cat') {
    const page = Math.max(1, parseInt(data.page || '1', 10));
    return replyQuickNavHint(replyToken, ss, page);
  }

  const op = data.op === 'buy' ? '購入' : (data.op === 'consume' ? '消費' : null);
  const product = data.product ? decodeURIComponent(data.product) : null;
  if (!op || !product) return safeReplyQ(replyToken, ss, '操作の解釈に失敗しました。');

  if (!isProductExists(ss, product)) {
    return safeReplyQ(replyToken, ss, '「' + product + '」は未登録です。先に登録してください。');
  }
  const unit = getUnit(ss, product);
  setPendingStateBySource(src, { state: STATE.awaitingListQuantity, item: { product: product, unit: unit }, operation: op });
  return safeReplyQ(replyToken, ss, '何' + unit + op + 'しましたか？（例：3）');
}

function isLineIdCommand_(message) {
  const normalized = String(message || '').trim().replace(/\s+/g, '').toLowerCase();
  return normalized === 'id' ||
    normalized === 'lineid' ||
    normalized === 'userid' ||
    normalized === 'user_id' ||
    normalized === 'id確認' ||
    normalized === 'ユーザーid';
}

function buildLineIdMessage_(event) {
  const source = (event && event.source) ? event.source : {};
  const lines = [
    '【LINE ID確認】',
    'userId: ' + (source.userId || '取得できませんでした'),
    'type: ' + (source.type || 'unknown')
  ];
  if (source.groupId) lines.push('groupId: ' + source.groupId);
  if (source.roomId) lines.push('roomId: ' + source.roomId);
  lines.push('');
  lines.push('Webアプリの許可には userId を使います。');
  return lines.join('\n');
}

// =======================================================
// List replies
// =======================================================
function replyCategoryFlexList(replyToken, src, ss, category, page) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const data = sh.getDataRange().getValues();
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const cat = data[i][0] || '未分類';
    if (cat === category) {
      rows.push({
        product: data[i][1],
        stock: Number(data[i][2]) || 0,
        min: data[i][3],
        unit: defaultUnit(data[i][4]),
        category: cat
      });
    }
  }

  if (rows.length === 0) return safeReplyQ(replyToken, ss, '分類「' + category + '」の在庫情報はありません。');
  rows.sort(function(a, b) { return (a.product || '').localeCompare(b.product || ''); });

  const lastPage = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  const p = Math.min(Math.max(1, page || 1), lastPage);
  const start = (p - 1) * PAGE_SIZE;
  const end = Math.min(start + PAGE_SIZE, rows.length);
  const pageRows = rows.slice(start, end).map(function(r) {
    return {
      product: r.product,
      productWrapped: wrapByFullWidth(r.product, 6, 2),
      stock: r.stock,
      min: r.min,
      unit: r.unit,
      category: r.category
    };
  });

  const mapping = pageRows.map(function(r, idx) {
    return { product: r.product, category: r.category, stock: r.stock, min: r.min, unit: r.unit, number: idx + 1 };
  });
  setPendingStateBySource(src, { state: STATE.awaitingListSelectionByCategory, mapping: mapping, page: p, category: category });

  const pageInfo = p + '/' + lastPage + 'ページ（全' + rows.length + '件）';
  const contents = buildCarouselBubbles(pageRows, category, pageInfo, 'cat', p, lastPage, category);
  return replyFlexQ(replyToken, ss, '【' + category + '】 ' + pageInfo, { type: 'carousel', contents: contents });
}

function replyAllFlexList(replyToken, src, ss, page) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return safeReplyQ(replyToken, ss, '在庫情報はありません。');

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({
      category: data[i][0] || '未分類',
      product: data[i][1],
      stock: Number(data[i][2]) || 0,
      min: data[i][3],
      unit: defaultUnit(data[i][4])
    });
  }
  rows.sort(function(a, b) {
    return (a.category || '').localeCompare(b.category || '') || (a.product || '').localeCompare(b.product || '');
  });

  const lastPage = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
  const p = Math.min(Math.max(1, page || 1), lastPage);
  const start = (p - 1) * PAGE_SIZE;
  const end = Math.min(start + PAGE_SIZE, rows.length);
  const pageRows = rows.slice(start, end).map(function(r) {
    return {
      product: r.product,
      productWrapped: wrapByFullWidth(r.product, 6, 2),
      stock: r.stock,
      min: r.min,
      unit: r.unit,
      category: r.category
    };
  });

  const mapping = pageRows.map(function(r, idx) {
    return { product: r.product, category: r.category, stock: r.stock, min: r.min, unit: r.unit, number: idx + 1 };
  });
  setPendingStateBySource(src, { state: STATE.awaitingListSelection, mapping: mapping, page: p });

  const pageInfo = p + '/' + lastPage + 'ページ（全' + rows.length + '件）';
  const contents = buildCarouselBubbles(pageRows, '在庫一覧（全体）', pageInfo, 'all', p, lastPage, null);
  return replyFlexQ(replyToken, ss, '在庫一覧（全体） ' + pageInfo, { type: 'carousel', contents: contents });
}

function buildCarouselBubbles(rowsPerPage, title, pageInfo, navType, page, lastPage, category) {
  const bubbles = [];
  for (let i = 0; i < rowsPerPage.length; i += ROWS_PER_BUBBLE) {
    const chunk = rowsPerPage.slice(i, i + ROWS_PER_BUBBLE);
    const bodyContents = [];
    bodyContents.push(makeHeaderRow());
    bodyContents.push({ type: 'separator', margin: 'sm' });
    chunk.forEach(function(r) { bodyContents.push(makeRowItem(r)); });

    bubbles.push({
      type: 'bubble',
      size: 'giga',
      header: {
        type: 'box',
        layout: 'vertical',
        contents: [
          { type: 'text', text: '【' + title + '】', weight: 'bold', size: 'md', wrap: true },
          { type: 'text', text: pageInfo, size: 'xs', color: '#888888', wrap: true }
        ],
        paddingAll: '12px'
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'xs',
        contents: bodyContents,
        paddingAll: '10px'
      },
      footer: makePagerFooter(navType, page, lastPage, category)
    });
  }
  return bubbles;
}

function makeHeaderRow() {
  return {
    type: 'box',
    layout: 'horizontal',
    spacing: 'sm',
    contents: [
      { type: 'text', text: '品目', size: 'xs', weight: 'bold', flex: COLS.product, align: 'start', wrap: false, gravity: 'center' },
      { type: 'text', text: '在庫', size: 'xs', weight: 'bold', flex: COLS.stock, align: 'end', wrap: false, gravity: 'center' },
      { type: 'text', text: '最低', size: 'xs', weight: 'bold', flex: COLS.min, align: 'end', wrap: false, gravity: 'center' },
      { type: 'text', text: '単位', size: 'xs', weight: 'bold', flex: COLS.unit, align: 'center', wrap: false, gravity: 'center' },
      {
        type: 'box',
        layout: 'vertical',
        flex: COLS.buttons,
        contents: [{ type: 'text', text: '操作', size: 'xs', weight: 'bold', align: 'center' }]
      }
    ]
  };
}

function makeRowItem(r) {
  return {
    type: 'box',
    layout: 'horizontal',
    spacing: 'sm',
    contents: [
      textCell(r.productWrapped, COLS.product, 'start', false, 'sm', true, 2, 'center'),
      textCell(String(r.stock), COLS.stock, 'end', true, 'sm', true, 1, 'center'),
      textCell(r.min != null ? String(r.min) : '', COLS.min, 'end', true, 'sm', false, 1, 'center'),
      textCell(r.unit || '個', COLS.unit, 'center', true, 'sm', false, 1, 'center'),
      {
        type: 'box',
        layout: 'horizontal',
        spacing: 'sm',
        flex: COLS.buttons,
        contents: [
          {
            type: 'button',
            style: 'primary',
            height: 'md',
            action: { type: 'postback', label: '購', data: 'op=buy&product=' + encodeURIComponent(r.product), displayText: '購入: ' + r.product },
            flex: 1
          },
          {
            type: 'button',
            style: 'secondary',
            height: 'md',
            action: { type: 'postback', label: '消', data: 'op=consume&product=' + encodeURIComponent(r.product), displayText: '消費: ' + r.product },
            flex: 1
          }
        ]
      }
    ]
  };
}

function textCell(text, flex, align, noWrap, size, bold, maxLines, gravity) {
  return {
    type: 'text',
    text: text || '',
    size: size || 'sm',
    flex: flex,
    align: align,
    weight: bold ? 'bold' : 'regular',
    wrap: !noWrap,
    maxLines: Math.max(1, maxLines || 1),
    gravity: gravity || 'center'
  };
}

function makePagerFooter(navType, page, lastPage, category) {
  const prevPage = page > 1 ? page - 1 : 1;
  const nextPage = page < lastPage ? page + 1 : lastPage;
  const dataPrev = navType === 'cat'
    ? 'nav=cat&category=' + encodeURIComponent(category) + '&page=' + prevPage
    : 'nav=all&page=' + prevPage;
  const dataNext = navType === 'cat'
    ? 'nav=cat&category=' + encodeURIComponent(category) + '&page=' + nextPage
    : 'nav=all&page=' + nextPage;

  return {
    type: 'box',
    layout: 'horizontal',
    spacing: 'md',
    contents: [
      { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '前へ', data: dataPrev, displayText: '前へ' }, flex: 1 },
      { type: 'button', style: 'primary', height: 'sm', action: { type: 'postback', label: '次へ', data: dataNext, displayText: '次へ' }, flex: 1 }
    ],
    paddingAll: '10px'
  };
}

// =======================================================
// Pending flow
// =======================================================
function handlePendingFlow(ss, event, pending) {
  const replyToken = event.replyToken;
  const text = (event.message.text || '').trim();
  const src = event.source;

  if (pending.state === STATE.awaitingDeletionByNumber) {
    if (text === '0' || text === '不要') {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '削除をキャンセルしました。');
    }
    const sel = parseInt(text, 10);
    if (isNaN(sel) || sel < 1 || sel > pending.mapping.length) return safeReplyQ(replyToken, ss, '有効な番号を入力してください。');
    const selectedItem = pending.mapping[sel - 1];
    const res = deleteProduct(ss, selectedItem.product);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, res ? selectedItem.product + 'を削除しました。' : selectedItem.product + 'は存在しません。');
  }

  if (pending.state === STATE.awaitingOperation) {
    const norm = normalizeOp(text);
    if (norm) {
      setPendingStateBySource(src, { state: STATE.awaitingQuantity, item: pending.item, operation: norm });
      const unit = getUnit(ss, pending.item);
      return safeReplyQ(replyToken, ss, '何' + unit + norm + 'しましたか？（例：3）');
    }
    return replyWithQuick(replyToken, '「購（購入）」か「消（消費）」を選んでください。', ['購', '消']);
  }

  if (pending.state === STATE.awaitingQuantity) {
    const qty = parseQuantity(text);
    if (qty == null) {
      const unit = getUnit(ss, pending.item);
      return safeReplyQ(replyToken, ss, '数量は1以上の整数で入力してください。（例：3）\n単位：' + unit);
    }
    const newCount = updateInventoryQuantitySafe(ss, pending.item, pending.operation, qty);
    const unit = getUnit(ss, pending.item);
    clearPendingStateBySource(src);
    if (newCount == null) return safeReplyQ(replyToken, ss, '「' + pending.item + '」は未登録です。先に登録してください。');

    let extra = '';
    if (pending.operation === '消費') {
      const minInv = getMinimumInventory(ss, pending.item);
      if (minInv != null && isShort(newCount, minInv)) {
        extra = '\n【注意】在庫が最低管理在庫数以下です。現在 ' + newCount + ' ' + unit + ' / 最低 ' + minInv + ' ' + unit + '。';
      }
    }
    return safeReplyQ(replyToken, ss, pending.item + 'の在庫は現在 ' + newCount + ' ' + unit + 'です。' + extra);
  }

  if (pending.state === STATE.awaitingMinimum) {
    const minValue = parseQuantity(text, true);
    if (minValue == null) return safeReplyQ(replyToken, ss, '最低管理在庫数は0以上の整数で入力してください。');
    updateMinimumInventory(ss, pending.item, minValue);
    const cats = getCategoriesList(ss);
    setPendingStateBySource(src, { state: STATE.awaitingCategory, item: pending.item });
    return multiReplyQ(replyToken, ss, chunkify(pending.item + 'の最低管理在庫数を ' + minValue + ' に設定しました。\n次に、以下の分類から該当するものを入力してください：\n' + cats.join('\n')));
  }

  if (pending.state === STATE.awaitingCategory) {
    const cats = getCategoriesList(ss);
    if (cats.indexOf(text) === -1) {
      return multiReplyQ(replyToken, ss, chunkify('有効な分類を入力してください。以下から選んでください：\n' + cats.join('\n')));
    }
    updateCategory(ss, pending.item, text);
    sortInventory(ss);
    setPendingStateBySource(src, { state: STATE.awaitingUnit, item: pending.item });
    return safeReplyQ(replyToken, ss, pending.item + 'の分類を「' + text + '」に設定しました。\n次に、単位を入力してください（例：個 / 本 / セット）。未入力なら「個」になります。');
  }

  if (pending.state === STATE.awaitingUnit) {
    const unit = text && text !== '未入力' ? text.trim() : '個';
    updateUnit(ss, pending.item, unit);
    setPendingStateBySource(src, { state: STATE.awaitingInitialStock, item: pending.item });
    return safeReplyQ(replyToken, ss, 'いま在庫はありますか？あれば在庫数を教えてください（例：12）。ない場合は「ない」と入力してください。');
  }

  if (pending.state === STATE.awaitingInitialStock) {
    const n = parseQuantity(text, true);
    if (n != null) {
      updateInitialStock(ss, pending.item, n);
      const unit = getUnit(ss, pending.item);
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, pending.item + 'の在庫が ' + n + ' ' + unit + 'に設定されました。登録完了です。');
    }
    if (text === 'ない' || text === '無い' || text === '無し') {
      setPendingStateBySource(src, { state: STATE.awaitingPurchaseDecision, item: pending.item });
      return safeReplyQ(replyToken, ss, '購入しましたか？（「購入しました」または「購入していない」）');
    }
    return safeReplyQ(replyToken, ss, '在庫数を数字で入力するか、在庫がない場合は「ない」と入力してください。');
  }

  if (pending.state === STATE.awaitingPurchaseDecision) {
    if (text === '購入しました') {
      setPendingStateBySource(src, { state: STATE.awaitingPurchaseQuantity, item: pending.item });
      const unit = getUnit(ss, pending.item);
      return safeReplyQ(replyToken, ss, '何' + unit + '購入しましたか？（例：3）');
    }
    if (text === '購入していない') {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '今回は登録のみにします。');
    }
    return safeReplyQ(replyToken, ss, '「購入しました」または「購入していない」と入力してください。');
  }

  if (pending.state === STATE.awaitingPurchaseQuantity) {
    const q = parseQuantity(text);
    if (q == null) {
      const unit = getUnit(ss, pending.item);
      return safeReplyQ(replyToken, ss, '数量は1以上の整数で入力してください。（例：3）\n単位：' + unit);
    }
    updateInitialStock(ss, pending.item, q);
    const unit = getUnit(ss, pending.item);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, pending.item + 'の在庫が ' + q + ' ' + unit + 'に設定されました。登録完了です。');
  }

  if (pending.state === STATE.awaitingReclassificationSelection) {
    if (text === '0' || text === '再分類なし') {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '再分類をキャンセルしました。');
    }
    const selections = text.split(/[\s,]+/);
    const idx = [];
    selections.forEach(function(s) {
      const n = parseInt(s, 10);
      if (!isNaN(n) && n > 0 && n <= pending.unclassified.length) idx.push(n - 1);
    });
    if (idx.length === 0) return safeReplyQ(replyToken, ss, '有効な番号が入力されませんでした。');
    const selectedProducts = idx.map(function(i) { return pending.unclassified[i].product; });
    const cats = getCategoriesList(ss);
    let catMsg = '再分類先を以下から選んでください：\n';
    cats.forEach(function(c, i) { catMsg += (i + 1) + '. ' + c + '\n'; });
    catMsg += '\n番号を入力してください。';
    setPendingStateBySource(src, { state: STATE.awaitingReclassificationTarget, selected: selectedProducts });
    return multiReplyQ(replyToken, ss, chunkify(catMsg));
  }

  if (pending.state === STATE.awaitingReclassificationTarget) {
    const cats = getCategoriesList(ss);
    const n = parseInt(text, 10);
    if (isNaN(n) || n < 1 || n > cats.length) return safeReplyQ(replyToken, ss, '有効な番号を入力してください。');
    const target = cats[n - 1];
    pending.selected.forEach(function(p) { updateCategory(ss, p, target); });
    sortInventory(ss);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, '選択された品目を「' + target + '」に再分類しました。');
  }

  if (pending.state === STATE.awaitingChangeSelection) {
    if (text === '0' || text === '不要') {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '修正をキャンセルしました。');
    }
    const sel = parseInt(text, 10);
    if (isNaN(sel) || sel < 1 || sel > pending.mapping.length) return safeReplyQ(replyToken, ss, '有効な番号を入力してください。');
    const selectedItem = pending.mapping[sel - 1];
    setPendingStateBySource(src, { state: STATE.awaitingChangeOption, item: selectedItem });
    return safeReplyQ(replyToken, ss, selectedItem.product + 'の修正項目を選んでください。\n1: 在庫数\n2: 最低管理在庫数\n3: 両方');
  }

  if (pending.state === STATE.awaitingChangeOption) {
    if (text === '1') {
      setPendingStateBySource(src, { state: STATE.awaitingChangeInventory, item: pending.item });
      return safeReplyQ(replyToken, ss, pending.item.product + 'の新しい在庫数を入力してください。');
    }
    if (text === '2') {
      setPendingStateBySource(src, { state: STATE.awaitingChangeMinInventory, item: pending.item });
      return safeReplyQ(replyToken, ss, pending.item.product + 'の新しい最低管理在庫数を入力してください。');
    }
    if (text === '3') {
      setPendingStateBySource(src, { state: STATE.awaitingChangeBoth, item: pending.item });
      return safeReplyQ(replyToken, ss, pending.item.product + 'の新しい在庫数を入力してください。');
    }
    return safeReplyQ(replyToken, ss, '1, 2, 3のいずれかを入力してください。');
  }

  if (pending.state === STATE.awaitingChangeInventory) {
    const newStock = parseQuantity(text, true);
    if (newStock == null) return safeReplyQ(replyToken, ss, '数字を入力してください。');
    changeInventoryValue(ss, pending.item.product, newStock);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, pending.item.product + 'の在庫数を ' + newStock + ' に変更しました。');
  }

  if (pending.state === STATE.awaitingChangeMinInventory) {
    const newMin = parseQuantity(text, true);
    if (newMin == null) return safeReplyQ(replyToken, ss, '数字を入力してください。');
    changeMinimumValue(ss, pending.item.product, newMin);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, pending.item.product + 'の最低管理在庫数を ' + newMin + ' に変更しました。');
  }

  if (pending.state === STATE.awaitingChangeBoth) {
    const newStockBoth = parseQuantity(text, true);
    if (newStockBoth == null) return safeReplyQ(replyToken, ss, '数字を入力してください。');
    changeInventoryValue(ss, pending.item.product, newStockBoth);
    setPendingStateBySource(src, { state: STATE.awaitingChangeBothMin, item: pending.item, newStock: newStockBoth });
    return safeReplyQ(replyToken, ss, '次に、最低管理在庫数を入力してください。');
  }

  if (pending.state === STATE.awaitingChangeBothMin) {
    const newMinBoth = parseQuantity(text, true);
    if (newMinBoth == null) return safeReplyQ(replyToken, ss, '数字を入力してください。');
    changeMinimumValue(ss, pending.item.product, newMinBoth);
    clearPendingStateBySource(src);
    return safeReplyQ(replyToken, ss, '修正完了です。');
  }

  if (pending.state === STATE.awaitingListSelection || pending.state === STATE.awaitingListSelectionByCategory) {
    if (text === '0' || text === '不要') {
      clearPendingStateBySource(src);
      return safeReplyQ(replyToken, ss, '更新をキャンセルしました。');
    }
    const n = parseInt(text, 10);
    if (isNaN(n) || n < 1 || n > pending.mapping.length) return safeReplyQ(replyToken, ss, '有効な番号を入力してください。');
    const selectedItem = pending.mapping[n - 1];
    setPendingStateBySource(src, { state: STATE.awaitingListOperation, item: selectedItem });
    return replyWithQuick(replyToken, selectedItem.product + 'の在庫更新。購（購入）ですか、消（消費）ですか？', ['購', '消']);
  }

  if (pending.state === STATE.awaitingListOperation) {
    const norm = normalizeOp(text);
    if (norm) {
      setPendingStateBySource(src, { state: STATE.awaitingListQuantity, item: pending.item, operation: norm });
      return safeReplyQ(replyToken, ss, '何' + (pending.item.unit || '個') + norm + 'しましたか？（例：3）');
    }
    return replyWithQuick(replyToken, '「購（購入）」か「消（消費）」を選んでください。', ['購', '消']);
  }

  if (pending.state === STATE.awaitingListQuantity) {
    const q = parseQuantity(text);
    if (q == null) return safeReplyQ(replyToken, ss, '数量は1以上の整数で入力してください。（例：3）\n単位：' + (pending.item.unit || '個'));
    const newCount = updateInventoryQuantitySafe(ss, pending.item.product, pending.operation, q);
    clearPendingStateBySource(src);
    if (newCount == null) return safeReplyQ(replyToken, ss, '「' + pending.item.product + '」は未登録です。');
    let extra = '';
    if (pending.operation === '消費') {
      const minInv = getMinimumInventory(ss, pending.item.product);
      if (minInv != null && isShort(newCount, minInv)) {
        extra = '\n【注意】在庫が最低管理在庫数以下です。現在 ' + newCount + ' ' + (pending.item.unit || '個') + ' / 最低 ' + minInv + ' ' + (pending.item.unit || '個');
      }
    }
    return safeReplyQ(replyToken, ss, pending.item.product + 'の在庫は現在 ' + newCount + ' ' + (pending.item.unit || '個') + 'です。' + extra);
  }

  clearPendingStateBySource(src);
  return safeReplyQ(replyToken, ss, '状態がリセットされました。もう一度お試しください。');
}

// =======================================================
// State storage
// =======================================================
function sourceKey(src) {
  if (src.userId) return 'user_' + src.userId;
  if (src.groupId) return 'group_' + src.groupId;
  if (src.roomId) return 'room_' + src.roomId;
  return 'anon';
}

function setPendingStateBySource(src, obj) {
  PropertiesService.getScriptProperties().setProperty('pending_' + sourceKey(src), JSON.stringify(obj));
}

function getPendingStateBySource(src) {
  const s = PropertiesService.getScriptProperties().getProperty('pending_' + sourceKey(src));
  return s ? JSON.parse(s) : null;
}

function clearPendingStateBySource(src) {
  PropertiesService.getScriptProperties().deleteProperty('pending_' + sourceKey(src));
}

// =======================================================
// Spreadsheet I/O
// =======================================================
function readInventoryItems_(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][1]) continue;
    items.push({
      category: String(v[i][0] || ''),
      name: String(v[i][1] || ''),
      stock: Number(v[i][2] || 0),
      minStock: Number(v[i][3] || 0),
      unit: defaultUnit(v[i][4])
    });
  }
  return items;
}

function saveInventoryItems_(ss, items) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const rows = [];
  (items || []).forEach(function(input) {
    const item = normalizeInventoryItem_(input);
    if (!item.name) return;
    rows.push([item.category, item.name, item.stock, item.minStock, item.unit]);
  });
  rows.sort(function(a, b) {
    return String(a[0]).localeCompare(String(b[0]), 'ja') || String(a[1]).localeCompare(String(b[1]), 'ja');
  });
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  }
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function normalizeInventoryItem_(input) {
  const item = input || {};
  return {
    category: String(item.category || item['分類'] || ''),
    name: String(item.name || item['品目'] || ''),
    stock: Math.max(0, Number(item.stock != null ? item.stock : item['在庫']) || 0),
    minStock: Math.max(0, Number(item.minStock != null ? item.minStock : item['最低在庫']) || 0),
    unit: defaultUnit(item.unit || item['単位'])
  };
}

function upsertInventoryItem_(ss, input) {
  const item = normalizeInventoryItem_(input);
  if (!item.name) return false;
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]) === item.name) {
      sh.getRange(i + 1, 1, 1, 5).setValues([[item.category, item.name, item.stock, item.minStock, item.unit]]);
      return true;
    }
  }
  sh.appendRow([item.category, item.name, item.stock, item.minStock, item.unit]);
  return true;
}

function adjustInventoryItem_(ss, name, delta, memo) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]) === String(name)) {
      const current = Number(v[i][2] || 0);
      const next = Math.max(0, current + Number(delta || 0));
      sh.getRange(i + 1, 3).setValue(next);
      if (Number(delta || 0) < 0) {
        logConsumption(ss, name, Math.abs(Number(delta || 0)));
      }
      return true;
    }
  }
  if (Number(delta || 0) > 0) {
    sh.appendRow(['', name, Number(delta || 0), 0, '個']);
    return true;
  }
  return false;
}

function isProductExists(ss, name) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) return true;
  return false;
}

function getInventoryCount(ss, name) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) return Number(v[i][2]) || 0;
  return 0;
}

function defaultUnit(u) {
  return (u && String(u).trim()) ? String(u).trim() : '個';
}

function getUnit(ss, name) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) return defaultUnit(v[i][4]);
  return '個';
}

function addProduct(ss, name, initial, minInv, unit) {
  ss.getSheetByName(SHEETS.inventory).appendRow(['', name, initial, minInv, defaultUnit(unit)]);
}

function updateMinimumInventory(ss, name, minCount) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 4).setValue(minCount); break; }
}

function updateCategory(ss, name, cat) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 1).setValue(cat); break; }
}

function updateUnit(ss, name, unit) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 5).setValue(defaultUnit(unit)); break; }
}

function updateInitialStock(ss, name, stock) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 3).setValue(stock); break; }
}

function sortInventory(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 5).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
}

function changeInventoryValue(ss, name, newStock) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 3).setValue(newStock); break; }
}

function changeMinimumValue(ss, name, newMin) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][1] === name) { sh.getRange(i + 1, 4).setValue(newMin); break; }
}

function getNumberedInventoryList(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  if (v.length <= 1) return { message: '在庫情報はありません。', mapping: [] };

  const grouped = {};
  for (let i = 1; i < v.length; i++) {
    const cat = v[i][0] || '未分類';
    if (!grouped[cat]) grouped[cat] = [];
    grouped[cat].push({ product: v[i][1], stock: v[i][2], min: v[i][3], unit: v[i][4] });
  }
  const cats = Object.keys(grouped).sort();
  const mapping = [];
  let out = '';
  let n = 1;
  cats.forEach(function(cat) {
    out += '【' + cat + '】\n';
    grouped[cat].forEach(function(item) {
      const u = defaultUnit(item.unit);
      out += n + '. ' + item.product + '、在庫: ' + item.stock + ' ' + u + '、最低: ' + item.min + ' ' + u + '\n';
      mapping.push({ product: item.product, category: cat, stock: item.stock, min: item.min, unit: u, number: n });
      n++;
    });
    out += '\n';
  });
  return { message: out.trim(), mapping: mapping };
}

function readHistory_(ss) {
  const sh = ss.getSheetByName(SHEETS.log);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const limit = Math.min(lastRow - 1, 120);
  const startRow = Math.max(2, lastRow - limit + 1);
  const values = sh.getRange(startRow, 1, limit, 3).getValues();
  const history = [];
  for (let i = 0; i < values.length; i++) {
    const dateValue = values[i][1];
    const date = dateValue instanceof Date
      ? Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dateValue || '');
    history.push({
      item: String(values[i][0] || ''),
      date: date,
      quantity: Number(values[i][2] || 0)
    });
  }
  return history.reverse();
}

function readInventoryCategories_(items) {
  const categories = {};
  (items || []).forEach(function(item) {
    const category = String(item.category || '').trim();
    if (category) categories[category] = true;
  });
  return Object.keys(categories);
}

function syncCategories_(ss, items) {
  const sheet = ss.getSheetByName(SHEETS.categories);
  if (!sheet) return;

  const categories = {};
  (items || []).forEach(function(item) {
    const category = String(item.category || '').trim();
    if (category) categories[category] = true;
  });
  getCategoriesList(ss).forEach(function(category) {
    if (category) categories[category] = true;
  });

  const list = Object.keys(categories).sort(function(a, b) {
    return a.localeCompare(b, 'ja');
  });
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  if (!list.length) {
    sheet.getRange(1, 1).setValue('未分類');
    return;
  }
  const rows = [['未分類']].concat(list.map(function(category) { return [category]; }));
  sheet.getRange(1, 1, rows.length, 1).setValues(rows);
}

function getCategoriesList(ss) {
  const sh = ss.getSheetByName(SHEETS.categories);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  const cats = [];
  for (let i = 0; i < v.length; i++) {
    const x = (v[i][0] || '').toString().trim();
    if (x) cats.push(x);
  }
  return cats;
}

function getUnclassifiedItems(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < v.length; i++) {
    const cat = (v[i][0] || '').toString().trim();
    const product = v[i][1];
    if (!cat || cat === '未分類') list.push({ product: product });
  }
  return list;
}

function updateInventoryQuantitySafe(ss, itemName, operation, quantity) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const sh = ss.getSheetByName(SHEETS.inventory);
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (v[i][1] === itemName) {
        let count = Number(v[i][2]) || 0;
        if (operation === '購入') {
          count += quantity;
        } else if (operation === '消費') {
          const newVal = Math.max(0, count - quantity);
          const consumed = Math.min(count, quantity);
          count = newVal;
          if (consumed > 0) logConsumption(ss, itemName, consumed);
        }
        sh.getRange(i + 1, 3).setValue(count);
        return count;
      }
    }
    if (operation === '購入') {
      addProduct(ss, itemName, quantity, '', '個');
      return quantity;
    }
    return null;
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(ss, name) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] === name) {
      sh.deleteRow(i + 1);
      deleteConsumptionLogs(ss, name);
      return true;
    }
  }
  return false;
}

// =======================================================
// Consumption log / shortage
// =======================================================
function getConsumptionLogSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.log);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.log);
    sh.appendRow(['品目', '日付', '数量']);
  }
  return sh;
}

function logConsumption(ss, name, qty) {
  const sh = getConsumptionLogSheet_(ss);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sh.appendRow([name, today, qty]);
}

function deleteConsumptionLogs(ss, name) {
  const sh = getConsumptionLogSheet_(ss);
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === name) sh.deleteRow(i + 1);
}

function getConsumptionRate(ss, name) {
  const sh = getConsumptionLogSheet_(ss);
  const v = sh.getDataRange().getValues();
  const rec = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === name) {
      const d = new Date(v[i][1]);
      const q = parseInt(v[i][2], 10);
      if (!isNaN(d.getTime()) && !isNaN(q)) rec.push({ date: d, qty: q });
    }
  }
  if (rec.length < 2) return null;
  rec.sort(function(a, b) { return a.date - b.date; });
  const first = rec[0].date;
  const last = rec[rec.length - 1].date;
  const days = (last - first) / (1000 * 60 * 60 * 24);
  const total = rec.reduce(function(sum, r) { return sum + r.qty; }, 0);
  if (total === 0) return null;
  return days / total;
}

function isShort(stock, minInv) {
  if (minInv === '' || minInv == null || isNaN(minInv)) return false;
  return Number(stock) <= Number(minInv);
}

function notifyLowInventory() {
  const conf = CONF();
  const ss = SpreadsheetApp.openById(conf.SPREADSHEET_ID);
  ensureSheets_(ss);
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  const grouped = {};
  for (let i = 1; i < v.length; i++) {
    const cat = v[i][0] || '未分類';
    const product = v[i][1];
    const stock = Number(v[i][2]) || 0;
    const minInv = v[i][3];
    const u = defaultUnit(v[i][4]);
    if (isShort(stock, minInv)) {
      const shortage = Math.max(0, Number(minInv) - stock);
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(product + '：在庫 ' + stock + ' ' + u + ' / 最低 ' + minInv + ' ' + u + ' (不足: ' + shortage + ' ' + u + ')');
    }
  }
  let msg = '';
  Object.keys(grouped).forEach(function(cat) {
    msg += '【' + cat + '】\n' + grouped[cat].join('\n') + '\n';
  });
  if (msg && conf.GROUP_ID) {
    msg = '【在庫アラート】\n以下の品目の在庫が最低管理在庫以下です。\n' + msg.trim();
    postJson_('https://api.line.me/v2/bot/message/push', { to: conf.GROUP_ID, messages: [{ type: 'text', text: msg }] });
  }
}

function getShortageListByCategory(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  const grouped = {};
  for (let i = 1; i < v.length; i++) {
    const cat = v[i][0] || '未分類';
    const product = v[i][1];
    const stock = Number(v[i][2]) || 0;
    const minInv = v[i][3];
    const u = defaultUnit(v[i][4]);
    if (isShort(stock, minInv)) {
      const shortage = Math.max(0, Number(minInv) - stock);
      if (!grouped[cat]) grouped[cat] = [];
      grouped[cat].push(product + '：不足 ' + shortage + ' ' + u);
    }
  }
  if (Object.keys(grouped).length === 0) return '不足情報はありません。';
  let msg = '【不足一覧】\n';
  Object.keys(grouped).forEach(function(cat) {
    msg += '【' + cat + '】\n' + grouped[cat].join('\n') + '\n';
  });
  return msg.trim();
}

// =======================================================
// Usage text
// =======================================================
function getUsageInstructions() {
  return (
`【在庫管理アプリの使い方】

＜基本＞
・品目名を送信 → 在庫表示 → 『購（購入）』『消（消費）』で更新
・新規登録は「最低在庫 → 分類 → 単位 → 初期在庫」の順

＜一覧（Flex）＞
・「一覧」「一覧 2」… 全体一覧：1ページ15行
・「○○一覧」「○○一覧 2」… 分類一覧
・列：品目 / 在庫 / 最低 / 単位 / 操作

＜その他＞
・「変更」：在庫数/最低在庫を修正
・「削除」：番号指定で削除
・「不足一覧」：不足数を分類ごとに表示
・「未分類一覧」：未分類品目の再分類
・「使用頻度」：平均使用間隔（日/個）
・「id」：LINE userId の確認`
  );
}

function getUsageFrequencyInstructions(ss) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  if (v.length <= 1) return '在庫情報はありません。';
  const lines = ['【使用頻度】'];
  for (let i = 1; i < v.length; i++) {
    const name = v[i][1];
    if (!name) continue;
    const rate = getConsumptionRate(ss, name);
    if (rate == null) continue;
    lines.push(name + '：平均 ' + rate.toFixed(1) + ' 日 / 個');
  }
  if (lines.length === 1) return '使用頻度データはまだ十分にありません。';
  return lines.join('\n');
}

// =======================================================
// LINE send helpers
// =======================================================
function postJson_(url, payload) {
  const conf = CONF();
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + conf.CHANNEL_ACCESS_TOKEN
  };
  const opt = { method: 'post', headers: headers, payload: JSON.stringify(payload), muteHttpExceptions: true };
  const res = UrlFetchApp.fetch(url, opt);
  const code = res.getResponseCode();
  if (code >= 300) console.error('LINE API error', code, res.getContentText());
  return res;
}

function safeReplyQ(replyToken, ss, text, qrPage) {
  return postJson_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [{
      type: 'text',
      text: text,
      quickReply: { items: buildQuickReplyNavItems(ss, qrPage || 1) }
    }]
  });
}

function safeReply(replyToken, text) {
  return postJson_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  });
}

function multiReplyQ(replyToken, ss, parts, qrPage) {
  const arr = parts.slice(0, MSG_LIMITS.replyMaxParts);
  if (arr.length > 0) {
    arr[0] = Object.assign({}, arr[0], { quickReply: { items: buildQuickReplyNavItems(ss, qrPage || 1) } });
  }
  return postJson_('https://api.line.me/v2/bot/message/reply', { replyToken: replyToken, messages: arr });
}

function replyFlexQ(replyToken, ss, altText, flexContents, qrPage) {
  return postJson_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [{
      type: 'flex',
      altText: altText || '在庫一覧',
      contents: flexContents,
      quickReply: { items: buildQuickReplyNavItems(ss, qrPage || 1) }
    }]
  });
}

function replyWithQuick(replyToken, text, labels) {
  const items = (labels || []).slice(0, 12).map(function(lbl) {
    return { type: 'action', action: { type: 'message', label: lbl, text: lbl } };
  });
  return postJson_('https://api.line.me/v2/bot/message/reply', {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text, quickReply: { items: items } }]
  });
}

function buildQuickReplyNavItems(ss, page) {
  const cats = getCategoriesList(ss).slice().sort(function(a, b) { return a.localeCompare(b); });
  const items = [];
  items.push({ type: 'action', action: { type: 'message', label: '全体一覧', text: '一覧' } });

  if (cats.length <= 12) {
    cats.forEach(function(cat) {
      items.push({ type: 'action', action: { type: 'message', label: cat, text: cat + '一覧' } });
    });
    return items;
  }

  const lastPage = Math.max(1, Math.ceil(cats.length / QR_CAT_PAGE_SIZE));
  const p = Math.min(Math.max(1, page || 1), lastPage);
  const start = (p - 1) * QR_CAT_PAGE_SIZE;
  const end = Math.min(start + QR_CAT_PAGE_SIZE, cats.length);
  const pageCats = cats.slice(start, end);

  if (p > 1) {
    items.push({ type: 'action', action: { type: 'postback', label: '◀ 分類 前', data: 'nav=qr&qtype=cat&page=' + (p - 1), displayText: '分類（前へ）' } });
  }
  items.push({ type: 'action', action: { type: 'postback', label: '分類 次 ▶', data: 'nav=qr&qtype=cat&page=' + (p < lastPage ? p + 1 : lastPage), displayText: '分類（次へ）' } });
  pageCats.forEach(function(cat) {
    items.push({ type: 'action', action: { type: 'message', label: cat, text: cat + '一覧' } });
  });
  return items.slice(0, 13);
}

function replyQuickNavHint(replyToken, ss, page) {
  return safeReplyQ(replyToken, ss, '分類を選んでください。', page);
}

// =======================================================
// Utilities
// =======================================================
function chunkify(text) {
  const chunks = [];
  let rest = text;
  while (rest.length > MSG_LIMITS.chunkSize) {
    let cut = rest.lastIndexOf('\n', MSG_LIMITS.chunkSize);
    if (cut < 0) cut = MSG_LIMITS.chunkSize;
    chunks.push({ type: 'text', text: rest.slice(0, cut) });
    rest = rest.slice(cut);
  }
  if (rest) chunks.push({ type: 'text', text: rest });
  return chunks;
}

function parseQuantity(text, allowZero) {
  const m = text.match(/^(\d+)/);
  if (!m) return null;
  const n = parseInt(m[1], 10);
  if (isNaN(n)) return null;
  if (allowZero) return n >= 0 ? n : null;
  return n >= 1 ? n : null;
}

function parseQueryString(qs) {
  const obj = {};
  (qs || '').split('&').forEach(function(kv) {
    if (!kv) return;
    const p = kv.split('=');
    const k = decodeURIComponent(p[0] || '');
    const v = decodeURIComponent(p[1] || '');
    if (k) obj[k] = v;
  });
  return obj;
}

function normalizeOp(s) {
  if (!s) return null;
  if (s === '購' || s === '購入') return '購入';
  if (s === '消' || s === '消費') return '消費';
  return null;
}

function wrapByFullWidth(str, width, maxLines) {
  if (!str) return '';
  let lineW = 0;
  let lines = 1;
  let out = '';
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    const w = isHalfWidth(ch) ? 0.5 : 1;
    if (lineW + w > width) {
      out += '\n';
      lines++;
      lineW = 0;
      if (lines > maxLines) {
        out = out.slice(0, out.lastIndexOf('\n'));
        out += '…';
        break;
      }
    }
    out += ch;
    lineW += w;
  }
  return out;
}

function isHalfWidth(ch) {
  const code = ch.charCodeAt(0);
  return (code >= 0x20 && code <= 0x7E) || (code >= 0xFF61 && code <= 0xFF9F);
}

function getMinimumInventory(ss, itemName) {
  const sh = ss.getSheetByName(SHEETS.inventory);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][1] === itemName) return (v[i][3] !== '' && !isNaN(v[i][3])) ? Number(v[i][3]) : null;
  }
  return null;
}

function shortErr_(err) {
  const m = (err && (err.message || err.toString())) || 'unknown';
  return m.length > 140 ? (m.slice(0, 140) + '…') : m;
}
