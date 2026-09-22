const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const source = fs.readFileSync('Code.gs', 'utf8');
const context = vm.createContext({
  Utilities: { formatDate: (date) => date.toISOString().slice(0, 10) },
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' }
});
vm.runInContext(source, context);

const settings = { analysisDays: 180, fallbackAnalysisDays: 365, purchaseHorizonDays: 30, minSamples: 2 };
const now = new Date('2026-09-23T00:00:00Z');
const item = { name: 'test', category: 'food', unit: '個', stock: 5, minStock: 2 };
const bucket = (count, quantity, age) => ({ count, quantity, firstDate: new Date(now.getTime() - age * 86400000), lastDate: now });
const stats = (primary, fallback, all) => ({ primary, fallback, all });
const forecast = (product, data, pack = 1) => context.buildConsumptionForecastForItem_(product, data, now, settings, pack);

assert.equal(forecast(item, undefined), null);
assert.equal(forecast(item, stats(bucket(1, 2, 1), bucket(1, 2, 1), bucket(1, 2, 1))), null);
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

for (const file of ['index.html', 'shopping.html']) {
  const html = fs.readFileSync(file, 'utf8');
  const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)].map((match) => match[1]).filter(Boolean);
  scripts.forEach((script) => new vm.Script(script, { filename: file }));
}
console.log('Forecast and script syntax tests passed');
