// 集計ロジックのテスト(Node.js 18 以上)。実行: node --test tests/
const test = require('node:test');
const assert = require('node:assert/strict');
const r = require('../src/Report.js');

const rows = [
  { date: '2026-11-05', product: 'ブレンド', category: 'コーヒー豆', quantity: 2, amount: 2400 },
  { date: '2026-11-05', product: 'ミル', category: '器具', quantity: 1, amount: 8800 },
  { date: '2026-11-05', product: 'ブレンド', category: 'コーヒー豆', quantity: 1, amount: 1200 },
  { date: '2026-11-04', product: 'ブレンド', category: 'コーヒー豆', quantity: 5, amount: 6000 },
  { date: '2026-11-04', product: 'ケーキ', category: 'スイーツ', quantity: 1, amount: 2200 },
  { date: '2026-11-03', product: 'ミル', category: '器具', quantity: 9, amount: 99999 },
];

test('当日合計・前日比を計算できる', () => {
  const s = r.buildSummary(rows, '2026-11-05', '2026-11-04');
  assert.equal(s.total, 12400);
  assert.equal(s.prevTotal, 8200);
  assert.equal(s.diff, 4200);
  assert.equal(s.rate, 51.2);
  assert.equal(s.count, 3);
  assert.equal(s.quantity, 4);
});

test('カテゴリ別は売上順で、前日のみのカテゴリも含む', () => {
  const s = r.buildSummary(rows, '2026-11-05', '2026-11-04');
  assert.deepEqual(s.byCategory.map((c) => c.name), ['器具', 'コーヒー豆', 'スイーツ']);
  const sweets = s.byCategory.find((c) => c.name === 'スイーツ');
  assert.equal(sweets.total, 0);
  assert.equal(sweets.rate, -100);
  const tools = s.byCategory.find((c) => c.name === '器具');
  assert.equal(tools.rate, null); // 前日0円 → 比較不能
});

test('商品別トップは同じ商品を合算する', () => {
  const s = r.buildSummary(rows, '2026-11-05', '2026-11-04');
  assert.deepEqual(s.topProducts, [
    { name: 'ミル', total: 8800 },
    { name: 'ブレンド', total: 3600 },
  ]);
});

test('データがない日は0件・前日比なし', () => {
  const s = r.buildSummary(rows, '2026-12-01', '2026-11-30');
  assert.equal(s.total, 0);
  assert.equal(s.count, 0);
  assert.equal(s.rate, null);
});

test('金額・前日比の表示形式', () => {
  assert.equal(r.formatYen(1234567), '¥1,234,567');
  assert.equal(r.formatYen(-500), '-¥500');
  assert.equal(r.formatDiff(4200), '+¥4,200');
  assert.equal(r.formatRate(51.2), '+51.2%');
  assert.equal(r.formatRate(-3), '-3.0%');
  assert.equal(r.formatRate(null), '—');
});

test('メール本文に商品名がエスケープされて入る', () => {
  const s = r.buildSummary(
    [{ date: '2026-11-05', product: '<script>', category: 'A&B', quantity: 1, amount: 100 }],
    '2026-11-05', '2026-11-04');
  const html = r.buildHtmlBody(s, 'https://example.com');
  assert.ok(html.includes('&lt;script&gt;'));
  assert.ok(html.includes('A&amp;B'));
  assert.ok(!html.includes('<script>'));
  assert.match(r.buildSubject(s), /2026-11-05 売上 ¥100/);
});
