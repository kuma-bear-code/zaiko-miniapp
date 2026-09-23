const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync('Code.gs', 'utf8');
const context = vm.createContext({
  Utilities: { formatDate: (date) => date.toISOString().slice(0, 10) },
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }
});
vm.runInContext(source, context);
assert.equal(context.authorizeWebAppRequest_('').ok, false);

const settings = { analysisDays: 180, fallbackAnalysisDays: 365, purchaseHorizonDays: 30, minSamples: 2 };
const now = new Date('2026-09-23T00:00:00Z');
const item = { name: 'test', category: 'food', unit: '個', stock: 5, minStock: 2 };
const bucket = (count, quantity, age) => ({ count, quantity, firstDate: new Date(now.getTime() - age * 86400000), lastDate: now });
const stats = (primary, fallback, all) => ({ primary, fallback, all });
const forecast = (product, data, pack = 1) => context.buildConsumptionForecastForItem_(product, data, now, settings, pack);

assert.equal(forecast(item, undefined).basis, 'insufficient');
assert.equal(forecast({ ...item, stock: 0 }, undefined).suggestedPurchase, 2);
assert.equal(forecast(item, stats(bucket(1, 2, 1), bucket(1, 2, 1), bucket(1, 2, 1))).daysLeft, null);
const recent = forecast(item, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100)));
assert.equal(recent.basis, 'recent180');
assert.equal(recent.targetStock, 7);
assert.equal(recent.suggestedPurchase, 2);
assert.equal(forecast({ ...item, stock: 0 }, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100))).daysLeft, 0);
assert.equal(forecast({ ...item, stock: 2 }, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100))).shortage, 5);
assert.equal(forecast({ ...item, stock: 8 }, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100))).suggestedPurchase, 0);
assert.equal(forecast(item, stats(bucket(1, 1, 50), bucket(2, 20, 250), bucket(2, 20, 250))).basis, 'recent365');
assert.equal(forecast(item, stats(bucket(0, 0, 1), bucket(1, 1, 250), bucket(2, 20, 500))).basis, 'all');
assert.equal(forecast(item, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100)), 12).suggestedPurchase, 12);

const logRows = [
  ['soap', '2026-09-01', 1],
  ['soap', '2026-09-01', 1],
  ['soap', '2026-09-02', 1],
  ['soap', '2026-09-03', 8]
];
const logSheet = {
  getLastRow: () => logRows.length + 1,
  getRange: () => ({ getValues: () => logRows })
};
const anomalies = context.buildConsumptionAnomalies_({ getSheetByName: () => logSheet }, { ...settings, anomalyMultiplier: 4 });
assert.equal(anomalies.length, 3);
assert.equal(anomalies.filter((row) => row.reasons.includes('同日同量の利用記録（予測に含む）')).length, 2);
assert.equal(anomalies.find((row) => row.quantity === 8).row, 5);
const recentDate = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
const usageRows = [['soap', recentDate, 1], ['soap', recentDate, 1]];
const usageSheet = {
  getLastRow: () => usageRows.length + 1,
  getRange: () => ({ getValues: () => usageRows })
};
const usageStats = context.readConsumptionStats_({ getSheetByName: () => usageSheet }, settings);
assert.equal(usageStats.soap.primary.count, 2);
assert.equal(usageStats.soap.primary.quantity, 2);

function sheet(initialRows) {
  const rows = initialRows.map((row) => [...row]);
  return {
    rows,
    getLastRow: () => rows.length,
    getMaxRows: () => 1000,
    getDataRange: () => ({ getValues: () => rows.map((row) => [...row]) }),
    getRange: (row, column, height = 1, width = 1) => ({
      getValue: () => rows[row - 1]?.[column - 1],
      getValues: () => rows.slice(row - 1, row - 1 + height).map((values) => values.slice(column - 1, column - 1 + width)),
      setValue: (value) => { rows[row - 1][column - 1] = value; },
      setFormula: (value) => { rows[row - 1] ||= []; rows[row - 1][column - 1] = value; },
      clearContent: () => {
        for (let offset = 0; offset < height; offset++) {
          for (let cell = 0; cell < width; cell++) rows[row - 1 + offset][column - 1 + cell] = '';
        }
      },
      setValues: (values) => values.forEach((entry, offset) => entry.forEach((value, cell) => { rows[row - 1 + offset][column - 1 + cell] = value; }))
    }),
    appendRow: (row) => rows.push([...row]),
    deleteRow: (row) => rows.splice(row - 1, 1)
  };
}

const inventory = sheet([
  ['分類', '品目', '在庫', '最低在庫', '単位', '写真URL', 'メモ', '保管場所'],
  ['日用品', 'soap', 2, 1, '個', '', '', '洗面所']
]);
const consumption = sheet([['品目', '日付', '数量']]);
const packs = sheet([['品目名', '入数'], ['soap', 12]]);
const sheets = { Inventory: inventory, ConsumptionLog: consumption, 商品設定: packs };
const ss = { getSheetByName: (name) => sheets[name] };

assert.equal(context.adjustInventoryItem_(ss, 'soap', -5, ''), true);
assert.equal(context.readInventoryItems_(ss)[0].stock, 0);
assert.equal(consumption.rows[1][2], 2);
assert.equal(context.adjustInventoryItem_(ss, 'soap', 3, ''), true);
assert.equal(context.readInventoryItems_(ss)[0].stock, 3);
assert.equal(context.bulkRestockItems_(ss, [{ name: 'soap', quantity: 12 }]).updated, 1);
assert.equal(context.readInventoryItems_(ss)[0].stock, 15);
const guardedInventory = sheet([
  ['分類', '品目', '在庫', '最低在庫', '単位'],
  ['日用品', 'guarded soap', 3, 1, '個']
]);
const guardedLog = sheet([['品目', '日付', '数量']]);
const guardedSs = { getSheetByName: (name) => ({ Inventory: guardedInventory, ConsumptionLog: guardedLog })[name] };
assert.throws(() => context.adjustInventoryItem_(guardedSs, 'guarded soap', -1, '', 2), /Inventory changed/);
assert.equal(guardedInventory.rows[1][2], 3);
assert.equal(guardedLog.rows.length, 1);
assert.equal(context.adjustInventoryItem_(guardedSs, 'guarded soap', -1, '', 3), true);
assert.equal(guardedInventory.rows[1][2], 2);
assert.equal(guardedLog.rows[1][2], 1);
assert.throws(() => context.updateInventoryItemFromParams_(ss, {
  name: 'soap', stock: 10, expectedStock: 14, minStock: 1
}), /Inventory changed/);
assert.throws(() => context.saveInventoryItems_(ss, []), /existing items must not be omitted/);
assert.equal(context.updateInventoryItemFromParams_(ss, {
  originalName: 'soap', name: 'new soap', category: '日用品', stock: 15,
  expectedStock: 15, minStock: 1, unit: '個', location: '棚'
}), true);
assert.equal(consumption.rows[1][0], 'new soap');
assert.equal(packs.rows[1][0], 'new soap');
assert.equal(context.readInventoryItems_(ss)[0].location, '棚');
assert.equal(context.deleteProduct(ss, 'new soap'), true);
assert.equal(context.readInventoryItems_(ss).length, 0);
assert.equal(consumption.rows.length, 2);

const analysis = sheet([
  Array(15).fill(''),
  ['', 'soap', '', '', '', '', '', '', '', '', '', '', 1, '', '']
]);
const dashboard = sheet(Array.from({ length: 37 }, () => Array(9).fill('')));
const analysisSs = { getSheetByName: (name) => ({ 商品分析: analysis, Dashboard: dashboard })[name] };
const sampleForecast = forecast({ ...item, name: 'soap', stock: 0 }, stats(bucket(2, 30, 100), bucket(2, 30, 100), bucket(2, 30, 100)), 12);
context.syncForecastAnalysis_(analysisSs, [{ ...item, stock: 0 }], [sampleForecast]);
assert.equal(analysis.rows[1][14], 12);
assert.equal(analysis.rows[1][9], '今すぐ購入');
assert.match(dashboard.rows[11][0], /商品分析/);

const purchaseInventory = sheet([
  ['分類', '品目', '在庫', '最低在庫', '単位', '写真URL', 'メモ', '保管場所'],
  ['日用品', 'soap', 2, 1, '個', '', '', '洗面所']
]);
const purchaseLog = sheet([
  ['品目', '日付', '数量'],
  ['soap', new Date('2026-08-01T00:00:00Z'), 15],
  ['soap', new Date('2026-09-01T00:00:00Z'), 15]
]);
const purchasePacks = sheet([['品目名', '入数'], ['soap', 12]]);
const purchaseSs = {
  getSheetByName: (name) => ({ Inventory: purchaseInventory, ConsumptionLog: purchaseLog, 商品設定: purchasePacks }[name])
};
const shortageMessage = context.getShortageListByCategory(purchaseSs);
assert.match(shortageMessage, /次回購入/);
assert.match(shortageMessage, /購入目安 12 個/);

const shoppingHtml = fs.readFileSync('shopping.html', 'utf8');
const shoppingLogic = shoppingHtml.slice(
  shoppingHtml.indexOf('    function normalize(value) {'),
  shoppingHtml.indexOf('    function render() {')
);
const shoppingContext = vm.createContext({
  state: {
    items: [
      { name: 'upcoming', stock: 3, minStock: 1, unit: '個' },
      { name: 'enough', stock: 5, minStock: 1, unit: '個' },
      { name: 'empty', stock: 0, minStock: 1, unit: '個' }
    ],
    forecasts: [
      { name: 'upcoming', daysLeft: 20, suggestedPurchase: 4 },
      { name: 'enough', daysLeft: 40, suggestedPurchase: 0 }
    ]
  },
  filter: 'all',
  DONE_KEY: 'done',
  localStorage: { getItem: () => '{}' }
});
vm.runInContext(shoppingLogic, shoppingContext);
const shoppingRows = shoppingContext.getShoppingItems();
assert.equal(shoppingRows.length, 2);
assert.equal(shoppingRows.find((row) => row.name === 'upcoming').suggested, 4);
assert.equal(shoppingRows.find((row) => row.name === 'upcoming').purchaseStatus, '次回購入');
assert.equal(shoppingRows.find((row) => row.name === 'empty').purchaseStatus, '今すぐ購入');
shoppingContext.filter = 'next';
assert.equal(shoppingContext.getShoppingItems().length, 1);

assert.equal(context.formatDigestDate_('2026-09-22'), '2026-09-22');
const originalDate = context.Date;
const originalSpreadsheetApp = context.SpreadsheetApp;
const originalPropertiesService = context.PropertiesService;
const originalUrlFetchApp = context.UrlFetchApp;
context.Date = class FixedDate extends Date {
  constructor(value) {
    if (arguments.length === 0) super('2026-09-23T00:15:00.000Z');
    else super(value);
  }
  static UTC(...args) { return Date.UTC(...args); }
};
const digest = sheet([
  ['対象日', '配信文', '配信状態', '配信日時', 'エラー'],
  ['2026-09-22', '昨日の利用記録と買い物候補です。', '未送信', '', '']
]);
const digestSs = { getSheetByName: (name) => name === 'LINE配信' ? digest : null };
context.SpreadsheetApp = { openById: () => digestSs };
context.PropertiesService = { getScriptProperties: () => ({
  getProperty: (name) => ({
    SPREADSHEET_ID: 'spreadsheet-id',
    CHANNEL_ACCESS_TOKEN: 'test-token',
    GROUP_ID: 'test-group'
  })[name] || null
}) };
let pushCount = 0;
context.UrlFetchApp = { fetch: (url, options) => {
  assert.equal(url, 'https://api.line.me/v2/bot/message/push');
  assert.equal(options.headers.Authorization, 'Bearer test-token');
  const payload = JSON.parse(options.payload);
  assert.equal(payload.to, 'test-group');
  assert.equal(payload.messages[0].text, '昨日の利用記録と買い物候補です。');
  pushCount++;
  return { getResponseCode: () => 200, getContentText: () => '' };
} };
assert.deepEqual(JSON.parse(JSON.stringify(context.sendPendingLineDigest())), { status: 'sent', date: '2026-09-22' });
assert.equal(digest.rows[1][2], '配信済み');
assert.equal(digest.rows[1][3], '2026-09-23');
assert.equal(pushCount, 1);
assert.equal(context.sendPendingLineDigest().status, 'skipped');
assert.equal(pushCount, 1);
digest.rows[1][2] = '未送信';
context.UrlFetchApp = { fetch: () => {
  pushCount++;
  return { getResponseCode: () => 500, getContentText: () => '{"message":"rejected"}' };
} };
const failedDigest = context.sendPendingLineDigest();
assert.equal(failedDigest.status, 'failed');
assert.equal(digest.rows[1][2], '配信失敗');
assert.match(digest.rows[1][4], /LINE API 500/);
assert.equal(context.sendPendingLineDigest().status, 'skipped');
assert.equal(pushCount, 2);
context.Date = originalDate;
context.SpreadsheetApp = originalSpreadsheetApp;
context.PropertiesService = originalPropertiesService;
context.UrlFetchApp = originalUrlFetchApp;

for (const file of ['index.html', 'shopping.html']) {
  const html = fs.readFileSync(file, 'utf8');
  const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map((match) => match[1]).filter(Boolean);
  scripts.forEach((script) => new vm.Script(script, { filename: file }));
}
console.log('Forecast and script syntax tests passed');
